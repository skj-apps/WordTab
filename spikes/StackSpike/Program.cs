using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

namespace StackSpike
{
    /// <summary>
    /// WordTab spike 2 — out-of-process proof of the tabbed *illusion*: N real Word
    /// windows stacked at one rect, kept in geometric lockstep, switched by z-order.
    ///
    /// Spike 1 proved we can carve a 32px strip out of one Word window's layout and keep
    /// it there. That is a stripe, not a product. The product claim is "several documents
    /// look like one window with tabs", and this spike tests exactly that:
    ///
    ///   * every visible OpusApp is forced to the same outer rect and held there
    ///   * each gets its own strip child, and each strip paints the *same* tab bar, so the
    ///     strip appears continuous across a switch (this is how Office Tab does it)
    ///   * clicking a tab raises that document's window — no reparenting, no content move
    ///
    /// Still deliberately out-of-process, for the same reason as spike 1: it needs no COM
    /// add-in and no C++ toolchain, which keeps the spike about the stacking question alone.
    /// </summary>
    internal static class Program
    {
        const int STRIP_H = 32;
        const uint POLL_MS = 30;
        const int RESCAN_EVERY = 16;   // ~500ms at POLL_MS

        sealed class WordWin
        {
            public IntPtr Opus;
            public IntPtr Wwf;
            public IntPtr Strip;
            public int Seq;               // stable discovery order; tab order never reshuffles
            public string Title = "";

            public RECT OriginalOuter;    // for restore on exit
            public RECT NaturalWwf;       // where Word wants the document frame
            public RECT AppliedWwf;       // where we put it
            public bool HasApplied;
            public int NaturalBottomGap;  // client height minus NaturalWwf.bottom, for restore
            public RECT StripAt;
            public bool StripPlaced;
            public bool OnTaskbar = true;  // Word gives every window a button to start with
            public bool HiddenByUs;        // parked out of sight for the duration of a drag
        }

        static readonly List<WordWin> _wins = new List<WordWin>();
        static readonly Dictionary<IntPtr, WordWin> _byStrip = new Dictionary<IntPtr, WordWin>();
        static int _activeIdx;
        static int _nextSeq;

        static RECT _master;
        static bool _hasMaster;

        static uint _wordPid, _mainTid;
        static IntPtr _hook, _moveHook, _inst;
        static int _tick;

        // True while the user is inside Word's modal move/size loop.
        static bool _inMoveSize;
        static int _moveSizeTick;

        // Delegates handed to unmanaged code must outlive the call — keep them rooted.
        static WndProcDelegate _wndProc;
        static TimerProcDelegate _timerProc;
        static WinEventProcDelegate _winEventProc;
        static TimerProcDelegate _quitProc;

        static IntPtr _brBg, _brActive, _brInactive, _brEdge, _font;
        static ITaskbarList _taskbar;
        static StreamWriter _log;

        // STA because ITaskbarList is an apartment-threaded shell object, and because this
        // thread runs a Win32 message loop anyway.
        [STAThread]
        static int Main(string[] args)
        {
            // `--seconds N` exits cleanly after N seconds, which is what lets this be driven
            // from a script: Ctrl+C is not sendable cross-process, and killing the process
            // would leave every Word window shifted and stacked.
            int runSeconds = 0;
            for (int i = 0; i < args.Length - 1; i++)
                if (args[i] == "--seconds") int.TryParse(args[i + 1], out runSeconds);

            _mainTid = Native.GetCurrentThreadId();
            var logPath = Path.Combine(AppContext.BaseDirectory, "stack-spike.log");
            _log = new StreamWriter(logPath, false) { AutoFlush = true };
            Log("=== WordTab stack spike (spike 2) ===");
            Log("log file: " + logPath);

            _inst = Native.GetModuleHandle(null);
            if (!RegisterStripClass()) return 2;

            Rescan();
            if (_wins.Count == 0)
            {
                Log("FAIL: no usable Word window found. Open Word with a document, then rerun.");
                return 1;
            }
            if (_wins.Count == 1)
                Log("NOTE: only one Word window. Open a second document to see the point of this spike.");

            InitTaskbar();
            SyncAll();
            SyncTaskbar();

            _winEventProc = OnWinEvent;
            _hook = Native.SetWinEventHook(
                Native.EVENT_OBJECT_SHOW, Native.EVENT_OBJECT_LOCATIONCHANGE,
                IntPtr.Zero, _winEventProc, _wordPid, 0,
                Native.WINEVENT_OUTOFCONTEXT | Native.WINEVENT_SKIPOWNPROCESS);
            Log(_hook == IntPtr.Zero ? "WARN: WinEvent hook failed; relying on the poll timer"
                                     : "WinEvent hook installed");

            // Separate range: the move/size events sit far below the object events, and one
            // hook covering both would also deliver every event in between.
            _moveHook = Native.SetWinEventHook(
                Native.EVENT_SYSTEM_MOVESIZESTART, Native.EVENT_SYSTEM_MOVESIZEEND,
                IntPtr.Zero, _winEventProc, _wordPid, 0,
                Native.WINEVENT_OUTOFCONTEXT | Native.WINEVENT_SKIPOWNPROCESS);
            if (_moveHook == IntPtr.Zero)
                Log("WARN: move/size hook failed; dragging will glitch");

            _timerProc = OnTimer;
            Native.SetTimer(IntPtr.Zero, UIntPtr.Zero, POLL_MS, _timerProc);

            Console.CancelKeyPress += (s, e) =>
            {
                e.Cancel = true;
                Native.PostThreadMessage(_mainTid, Native.WM_QUIT, IntPtr.Zero, IntPtr.Zero);
            };
            StartKeyReader();

            if (runSeconds > 0)
            {
                _quitProc = (h, m, id, t) =>
                {
                    Native.KillTimer(IntPtr.Zero, id);
                    Native.PostThreadMessage(_mainTid, Native.WM_QUIT, IntPtr.Zero, IntPtr.Zero);
                };
                Native.SetTimer(IntPtr.Zero, UIntPtr.Zero, (uint)(runSeconds * 1000), _quitProc);
                Log("will exit automatically after " + runSeconds + "s");
            }

            Log("");
            Log("RUNNING with " + _wins.Count + " Word window(s), all snapped to one rect.");
            Log("Try this:");
            Log("  1. click a tab in the strip -> that document comes forward");
            Log("  2. or press 1..9 in THIS console to switch");
            Log("  3. drag or resize the Word window -> every stacked window must follow");
            Log("  4. open a new document -> it should join the strip within half a second");
            Log("  5. File (Backstage), then Back -> the strip must come back correctly");
            Log("  6. look at the taskbar -> one Word button, not one per document");
            Log("Press Ctrl+C here to stop and put every window back where it was.");
            Log("");

            MSG msg;
            while (Native.GetMessage(out msg, IntPtr.Zero, 0, 0) > 0)
            {
                if (msg.hwnd == IntPtr.Zero && msg.message == Native.WM_SWITCH)
                {
                    Activate((int)msg.wParam.ToInt64());
                    continue;
                }
                Native.TranslateMessage(ref msg);
                Native.DispatchMessage(ref msg);
            }

            Cleanup();
            return 0;
        }

        // ---------------------------------------------------------------- discovery

        /// <summary>
        /// Refresh the window set. New documents join at the end; closed ones drop out.
        /// Order is by discovery sequence, never by z-order — EnumWindows returns z-order
        /// and we reorder z-order on every switch, so using it would shuffle the tabs.
        /// </summary>
        static void Rescan()
        {
            var live = new List<IntPtr>();
            Native.EnumWindows((h, l) =>
            {
                if (Native.ClassOf(h) == "OpusApp") live.Add(h);
                return true;
            }, IntPtr.Zero);

            var usable = new HashSet<IntPtr>();
            foreach (var h in live)
            {
                IntPtr wwf = FindChild(h, "_WwF");
                RECT r; Native.GetWindowRect(h, out r);

                // Word keeps a permanent hidden background OpusApp with a full _Ww* tree,
                // and minimized windows park off-screen near -32000 (scaled by DPI, so
                // -21333 at 150%) at a stub size. Neither is something we can lay out into.
                // A window we hid ourselves for a drag is still part of the stack — without
                // this it would look closed, lose its strip, and never come back.
                bool visible = Native.IsWindowVisible(h)
                               || _wins.Exists(x => x.Opus == h && x.HiddenByUs);

                if (visible && !Native.IsCloaked(h) && !Native.IsIconic(h)
                    && wwf != IntPtr.Zero
                    && (r.right - r.left) >= 200 && (r.bottom - r.top) >= 200)
                {
                    usable.Add(h);
                }
            }

            // Drop windows that closed or were minimized away.
            for (int i = _wins.Count - 1; i >= 0; i--)
            {
                var w = _wins[i];
                if (usable.Contains(w.Opus) && Native.IsWindow(w.Opus)) continue;

                Log(string.Format("window left the stack: 0x{0:X} \"{1}\"", w.Opus.ToInt64(), w.Title));

                // It may have only been minimized rather than closed, in which case it needs
                // its taskbar button back or the user cannot get to it again.
                if (_taskbar != null && !w.OnTaskbar && Native.IsWindow(w.Opus))
                {
                    try { _taskbar.AddTab(w.Opus); }
                    catch (Exception ex) { Log("WARN: taskbar restore failed: " + ex.Message); }
                }

                if (w.Strip != IntPtr.Zero)
                {
                    _byStrip.Remove(w.Strip);
                    Native.DestroyWindow(w.Strip);
                }
                _wins.RemoveAt(i);
                if (_activeIdx >= i && _activeIdx > 0) _activeIdx--;
            }

            // Adopt anything new.
            foreach (var h in live)
            {
                if (!usable.Contains(h)) continue;
                if (_wins.Exists(w => w.Opus == h)) continue;

                var w = new WordWin
                {
                    Opus = h,
                    Wwf = FindChild(h, "_WwF"),
                    Seq = _nextSeq++,
                    Title = ShortTitle(h),
                };
                Native.GetWindowRect(h, out w.OriginalOuter);

                if (!CreateStrip(w))
                {
                    Log("WARN: could not create a strip for 0x" + h.ToString("X") + "; skipping it");
                    continue;
                }

                _wins.Add(w);
                _byStrip[w.Strip] = w;
                if (_wordPid == 0) Native.GetWindowThreadProcessId(h, out _wordPid);

                Log(string.Format("joined the stack [{0}]: OpusApp=0x{1:X} _WwF=0x{2:X} \"{3}\"",
                                  _wins.Count - 1, h.ToInt64(), w.Wwf.ToInt64(), w.Title));
            }

            _wins.Sort((a, b) => a.Seq.CompareTo(b.Seq));

            // Prefer whatever the user was actually looking at as the first active tab.
            if (!_hasMaster)
            {
                IntPtr fg = Native.GetForegroundWindow();
                int idx = _wins.FindIndex(w => w.Opus == fg);
                if (idx >= 0) _activeIdx = idx;
            }
            if (_activeIdx >= _wins.Count) _activeIdx = Math.Max(0, _wins.Count - 1);

            // Titles change as documents are saved/renamed.
            bool titlesChanged = false;
            foreach (var w in _wins)
            {
                string t = ShortTitle(w.Opus);
                if (t == w.Title) continue;
                w.Title = t;
                titlesChanged = true;
            }
            if (titlesChanged) InvalidateStrips();
        }

        static IntPtr FindChild(IntPtr parent, string cls)
        {
            IntPtr found = IntPtr.Zero;
            Native.EnumChildWindows(parent, (h, l) =>
            {
                if (Native.ClassOf(h) == cls) { found = h; return false; }
                return true;
            }, IntPtr.Zero);
            return found;
        }

        static string ShortTitle(IntPtr hWnd)
        {
            string t = Native.TitleOf(hWnd);
            const string suffix = " - Word";
            if (t.EndsWith(suffix, StringComparison.Ordinal))
                t = t.Substring(0, t.Length - suffix.Length);
            return t.Length == 0 ? "(untitled)" : t;
        }

        // ------------------------------------------------------------------ our windows

        static bool RegisterStripClass()
        {
            _wndProc = StripWndProc;
            _brBg = Native.CreateSolidBrush(Native.Rgb(238, 238, 238));
            _brActive = Native.CreateSolidBrush(Native.Rgb(255, 255, 255));
            _brInactive = Native.CreateSolidBrush(Native.Rgb(206, 206, 206));
            _brEdge = Native.CreateSolidBrush(Native.Rgb(150, 150, 150));
            _font = Native.GetStockObject(Native.DEFAULT_GUI_FONT);

            var wc = new WNDCLASSEX
            {
                cbSize = (uint)Marshal.SizeOf<WNDCLASSEX>(),
                style = 0,
                lpfnWndProc = Marshal.GetFunctionPointerForDelegate(_wndProc),
                hInstance = _inst,
                hbrBackground = IntPtr.Zero,   // we paint every pixel in WM_PAINT
                lpszClassName = "WordTabStackSpike",
            };

            if (Native.RegisterClassEx(ref wc) != 0) return true;
            Log("RegisterClassEx failed, err=" + Marshal.GetLastWin32Error());
            return false;
        }

        static bool CreateStrip(WordWin w)
        {
            // Creating a child under a window owned by another process is legal Win32;
            // it implicitly attaches our input queue to Word's for the lifetime of the child.
            // WS_EX_NOACTIVATE is what lets a click on our tab bar not steal activation.
            w.Strip = Native.CreateWindowEx(
                Native.WS_EX_NOACTIVATE, "WordTabStackSpike", "WordTab stack spike",
                Native.WS_CHILD | Native.WS_VISIBLE | Native.WS_CLIPSIBLINGS,
                0, 0, 10, STRIP_H, w.Opus, IntPtr.Zero, _inst, IntPtr.Zero);

            if (w.Strip != IntPtr.Zero) return true;
            Log("CreateWindowEx failed, err=" + Marshal.GetLastWin32Error());
            return false;
        }

        static IntPtr StripWndProc(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam)
        {
            switch (msg)
            {
                case Native.WM_ERASEBKGND:
                    return (IntPtr)1;      // WM_PAINT covers the whole client; skip the flicker

                case Native.WM_PAINT:
                    PaintStrip(hWnd);
                    return IntPtr.Zero;

                case Native.WM_LBUTTONDOWN:
                    {
                        int idx = HitTest(hWnd, Native.LoWord(lParam));
                        if (idx >= 0) Activate(idx);
                        return IntPtr.Zero;
                    }
            }
            return Native.DefWindowProc(hWnd, msg, wParam, lParam);
        }

        // ------------------------------------------------------------------ the tab bar

        // Every window paints an identical bar, so switching windows looks like the same
        // strip staying put — which is precisely the trick Office Tab uses.
        const int TAB_MAX_W = 220;
        const int TAB_MIN_W = 70;
        const int TAB_GAP = 2;
        const int BAR_PAD = 4;

        static int TabWidth(int clientW)
        {
            int n = Math.Max(1, _wins.Count);
            int avail = Math.Max(1, clientW - BAR_PAD * 2);
            return Math.Max(TAB_MIN_W, Math.Min(TAB_MAX_W, avail / n));
        }

        static int HitTest(IntPtr strip, int x)
        {
            RECT c;
            if (!Native.GetClientRect(strip, out c)) return -1;
            int tw = TabWidth(c.right);
            int rel = x - BAR_PAD;
            if (rel < 0) return -1;
            int idx = rel / tw;
            if (idx >= _wins.Count) return -1;
            if (rel - idx * tw >= tw - TAB_GAP) return -1;   // in the gap between tabs
            return idx;
        }

        static void PaintStrip(IntPtr strip)
        {
            PAINTSTRUCT ps;
            IntPtr hdc = Native.BeginPaint(strip, out ps);
            try
            {
                RECT c;
                if (!Native.GetClientRect(strip, out c)) return;

                Native.FillRect(hdc, ref c, _brBg);
                Native.SelectObject(hdc, _font);
                Native.SetBkMode(hdc, Native.TRANSPARENT);

                int tw = TabWidth(c.right);
                for (int i = 0; i < _wins.Count; i++)
                {
                    int left = BAR_PAD + i * tw;
                    if (left >= c.right) break;
                    bool active = i == _activeIdx;

                    var tab = new RECT
                    {
                        left = left,
                        top = 3,
                        right = Math.Min(c.right, left + tw - TAB_GAP),
                        bottom = c.bottom,
                    };
                    Native.FillRect(hdc, ref tab, _brEdge);

                    var inner = new RECT
                    {
                        left = tab.left + 1,
                        top = tab.top + 1,
                        right = Math.Max(tab.left + 1, tab.right - 1),
                        bottom = tab.bottom,
                    };
                    Native.FillRect(hdc, ref inner, active ? _brActive : _brInactive);

                    var text = new RECT
                    {
                        left = inner.left + 8,
                        top = inner.top,
                        right = Math.Max(inner.left + 8, inner.right - 8),
                        bottom = inner.bottom,
                    };
                    Native.SetTextColor(hdc, active ? Native.Rgb(16, 16, 16) : Native.Rgb(80, 80, 80));
                    string label = (i + 1) + "  " + _wins[i].Title;
                    Native.DrawText(hdc, label, label.Length, ref text,
                                    Native.DT_SINGLELINE | Native.DT_VCENTER |
                                    Native.DT_END_ELLIPSIS | Native.DT_NOPREFIX);
                }
            }
            finally
            {
                Native.EndPaint(strip, ref ps);
            }
        }

        static void InvalidateStrips()
        {
            foreach (var w in _wins)
                if (w.Strip != IntPtr.Zero) Native.InvalidateRect(w.Strip, IntPtr.Zero, false);
        }

        // ------------------------------------------------------------------- switching

        /// <summary>Raise one document's window. No reparenting, no content move — z-order only.</summary>
        static void Activate(int idx)
        {
            if (idx < 0 || idx >= _wins.Count) return;
            var w = _wins[idx];
            if (!Native.IsWindow(w.Opus)) return;

            _activeIdx = idx;

            // Raise first without activating, so the visual switch is immediate even if the
            // foreground handoff below is refused by the foreground lock.
            Native.SetWindowPos(w.Opus, Native.HWND_TOP, 0, 0, 0, 0,
                                Native.SWP_NOMOVE | Native.SWP_NOSIZE | Native.SWP_NOACTIVATE);

            // Keyboard focus has to follow, or the user types into the document underneath.
            // SetForegroundWindow is refused unless we share an input queue with the current
            // foreground thread, hence the AttachThreadInput dance.
            uint fgPid;
            uint fgTid = Native.GetWindowThreadProcessId(Native.GetForegroundWindow(), out fgPid);
            uint myTid = Native.GetCurrentThreadId();
            bool attached = fgTid != 0 && fgTid != myTid && Native.AttachThreadInput(myTid, fgTid, true);
            Native.BringWindowToTop(w.Opus);
            bool fg = Native.SetForegroundWindow(w.Opus);
            if (attached) Native.AttachThreadInput(myTid, fgTid, false);

            Log(string.Format("switch -> [{0}] \"{1}\"  foreground={2}", idx, w.Title, fg));

            InvalidateStrips();
            SyncAll();
            SyncTaskbar();
        }

        // ------------------------------------------------------------------- the taskbar

        static void InitTaskbar()
        {
            try
            {
                _taskbar = (ITaskbarList)new TaskbarListClass();
                _taskbar.HrInit();
                Log("ITaskbarList ready");
            }
            catch (Exception ex)
            {
                _taskbar = null;
                Log("WARN: ITaskbarList unavailable, taskbar buttons will not be suppressed: "
                    + ex.Message);
            }
        }

        /// <summary>
        /// N stacked windows otherwise produce N taskbar buttons, which gives the illusion
        /// away immediately. Only the window the user can actually see keeps a button.
        /// </summary>
        static void SyncTaskbar()
        {
            if (_taskbar == null) return;

            for (int i = 0; i < _wins.Count; i++)
            {
                var w = _wins[i];
                bool want = i == _activeIdx;
                if (w.OnTaskbar == want || !Native.IsWindow(w.Opus)) continue;

                try
                {
                    if (want) _taskbar.AddTab(w.Opus);
                    else _taskbar.DeleteTab(w.Opus);
                    w.OnTaskbar = want;
                    Log(string.Format("taskbar: {0} \"{1}\"", want ? "show" : "hide", w.Title));
                }
                catch (Exception ex)
                {
                    Log(string.Format("WARN: taskbar {0} on \"{1}\" failed: {2}",
                                      want ? "AddTab" : "DeleteTab", w.Title, ex.Message));
                }
            }
        }

        /// <summary>Give every window its taskbar button back, whatever state we left it in.</summary>
        static void RestoreTaskbar()
        {
            if (_taskbar == null) return;
            foreach (var w in _wins)
            {
                if (w.OnTaskbar || !Native.IsWindow(w.Opus)) continue;
                try { _taskbar.AddTab(w.Opus); w.OnTaskbar = true; }
                catch (Exception ex) { Log("WARN: taskbar restore failed: " + ex.Message); }
            }
        }

        static void StartKeyReader()
        {
            var t = new Thread(() =>
            {
                while (true)
                {
                    ConsoleKeyInfo k;
                    try { k = Console.ReadKey(true); }
                    catch (InvalidOperationException) { return; }   // no console (redirected)
                    if (k.KeyChar >= '1' && k.KeyChar <= '9')
                        Native.PostThreadMessage(_mainTid, Native.WM_SWITCH,
                                                 (IntPtr)(k.KeyChar - '1'), IntPtr.Zero);
                    else if (k.KeyChar == 'q' || k.KeyChar == 'Q')
                        Native.PostThreadMessage(_mainTid, Native.WM_QUIT, IntPtr.Zero, IntPtr.Zero);
                }
            });
            t.IsBackground = true;
            t.Start();
        }

        // ------------------------------------------------------------------- the geometry

        static void OnTimer(IntPtr h, uint m, UIntPtr id, uint t)
        {
            _tick++;

            // Watchdog: if MOVESIZEEND never arrives (the drag was interrupted, the hook
            // dropped an event), the followers would stay hidden forever. Roughly 15s.
            if (_inMoveSize && _tick - _moveSizeTick > 500)
            {
                Log("WARN: move/size never ended — restoring followers anyway");
                EndMoveSize();
            }

            if (_tick % RESCAN_EVERY == 0) { Rescan(); SyncTaskbar(); }
            SyncAll();
        }

        static void OnWinEvent(IntPtr hook, uint ev, IntPtr hWnd,
                               int idObject, int idChild, uint thread, uint time)
        {
            // EVENT_OBJECT_LOCATIONCHANGE also fires for the text caret on every keystroke.
            if (idObject != Native.OBJID_WINDOW) return;
            bool ours = _wins.Exists(w => w.Opus == hWnd || w.Wwf == hWnd);
            if (!ours) return;

            if (ev == Native.EVENT_SYSTEM_MOVESIZESTART) { BeginMoveSize(); return; }
            if (ev == Native.EVENT_SYSTEM_MOVESIZEEND) { EndMoveSize(); return; }

            SyncAll();
        }

        /// <summary>
        /// Dragging a stack is where out-of-process shows its seams. Word moves the dragged
        /// window smoothly inside a modal loop; we only get to reposition the others once per
        /// 30ms poll, so they visibly lag behind and peek out from under the one being dragged.
        ///
        /// There is no way to win that race from another process — an in-process add-in would
        /// handle WM_WINDOWPOSCHANGING and move all of them in the same frame. So instead of
        /// showing a bad approximation, hide the followers for the duration of the drag and put
        /// them back in place when it ends.
        /// </summary>
        static void BeginMoveSize()
        {
            if (_inMoveSize) return;
            _inMoveSize = true;
            _moveSizeTick = _tick;

            var active = _wins[Math.Min(_activeIdx, _wins.Count - 1)];
            foreach (var w in _wins)
            {
                if (w == active || w.HiddenByUs || !Native.IsWindow(w.Opus)) continue;
                Native.ShowWindow(w.Opus, Native.SW_HIDE);
                w.HiddenByUs = true;
            }
            Log("move/size started — followers hidden");
        }

        static void EndMoveSize()
        {
            if (!_inMoveSize) return;
            _inMoveSize = false;

            // Put them where the drag ended *before* revealing them, or they flash at the old
            // position for a frame.
            SyncAll();

            foreach (var w in _wins)
            {
                if (!w.HiddenByUs || !Native.IsWindow(w.Opus)) continue;
                Native.ShowWindow(w.Opus, Native.SW_SHOWNA);
                w.HiddenByUs = false;
            }
            SyncAll();
            Log("move/size ended — followers restored");
        }

        static void SyncAll()
        {
            if (_wins.Count == 0)
            {
                Log("no Word windows left — exiting");
                Native.PostThreadMessage(_mainTid, Native.WM_QUIT, IntPtr.Zero, IntPtr.Zero);
                return;
            }

            SyncOuterRects();

            // Only the active window gets laid out by Word. Force-resizing a background
            // OpusApp changes its frame but Word does not relayout the interior — measured:
            // after resizing the stack from 1095 to 900 wide, the background windows kept a
            // 1081px-wide `_WwF` inside an 885px client, and activating them did not fix it.
            // So the active window is the layout oracle and everyone else copies it. They are
            // all the same size by construction, so the correct interior is the same too.
            var active = _wins[Math.Min(_activeIdx, _wins.Count - 1)];
            SyncActive(active);

            foreach (var w in _wins)
            {
                if (w == active) continue;
                if (Sane(active.NaturalWwf)) SyncFollower(w, active.NaturalWwf);
                else SyncActive(w);        // no usable oracle yet — let it find its own layout
            }
        }

        static bool Sane(RECT natural)
        {
            return natural.right - natural.left > 200 && natural.bottom - natural.top > 100;
        }

        /// <summary>
        /// Hold every stacked window at one rect.
        ///
        /// The active window is the sole authority: whatever the user does to it becomes the
        /// master rect, and everyone else is pushed to match. Letting a background window
        /// claim authority would be a feedback loop, since our own SetWindowPos on it fires
        /// the same LOCATIONCHANGE event that would make us adopt its rect.
        /// </summary>
        static void SyncOuterRects()
        {
            var active = _wins[Math.Min(_activeIdx, _wins.Count - 1)];
            RECT cur;
            if (!Native.IsWindow(active.Opus) || !Native.GetWindowRect(active.Opus, out cur)) return;

            if (!_hasMaster || !Same(cur, _master))
            {
                // Not during a drag: that fires every poll and buries the log in noise.
                if (_hasMaster && !_inMoveSize)
                    Log(string.Format("master rect -> ({0},{1} {2}x{3}) from \"{4}\"",
                                      cur.left, cur.top, cur.right - cur.left, cur.bottom - cur.top,
                                      active.Title));
                _master = cur;
                _hasMaster = true;
            }

            int mw = _master.right - _master.left, mh = _master.bottom - _master.top;
            if (mw < 1 || mh < 1) return;

            // Mid-drag the followers are hidden; moving them every poll is what caused the
            // glitching in the first place. They get placed once, when the drag ends.
            if (_inMoveSize) return;

            foreach (var w in _wins)
            {
                if (w == active || !Native.IsWindow(w.Opus)) continue;
                RECT r;
                if (!Native.GetWindowRect(w.Opus, out r) || Same(r, _master)) continue;

                Native.SetWindowPos(w.Opus, IntPtr.Zero, _master.left, _master.top, mw, mh,
                                    Native.SWP_NOZORDER | Native.SWP_NOACTIVATE |
                                    Native.SWP_NOOWNERZORDER);
            }
        }

        /// <summary>
        /// Spike 1's finding, now per window: Word repeatedly resets `_WwF` to its own natural
        /// rect, so the sync must be idempotent against Word's layout rather than a repeated
        /// delta — re-applying -32 every pass walks the document frame down the window.
        ///
        /// That rule is idempotent against *Word*, but not against another process doing the
        /// same thing. Spike 2's first run had a stray spike 1 process still live on one of the
        /// windows: each saw the other's +32 as a fresh natural rect and added another 32, and
        /// that window's document frame walked down to 25px tall in seconds. Office Tab, which
        /// shifts the same `_WwF`, would do this to us on any machine where both are installed.
        ///
        /// So the top edge is tracked separately: if the frame's top is still exactly where we
        /// put it, nobody moved it, and only the other edges get re-derived. That converges
        /// instead of racing, and the height tripwire below refuses the collapse outright.
        /// </summary>
        static void SyncActive(WordWin w)
        {
            if (!Native.IsWindow(w.Opus) || !Native.IsWindow(w.Wwf)) return;

            RECT cur;
            if (!TryGetClientRect(w.Opus, w.Wwf, out cur)) return;

            if (w.HasApplied && Same(cur, w.AppliedWwf))
            {
                PlaceStrip(w);
                return;
            }

            bool topIsOurs = w.HasApplied && cur.top == w.AppliedWwf.top;
            var natural = new RECT
            {
                left = cur.left,
                top = topIsOurs ? w.NaturalWwf.top : cur.top,
                right = cur.right,
                bottom = cur.bottom,
            };

            Impose(w, natural, topIsOurs ? "top already ours" : "from Word");
        }

        /// <summary>Hold a background window's interior at the layout the active window proved.</summary>
        static void SyncFollower(WordWin w, RECT natural)
        {
            if (!Native.IsWindow(w.Opus) || !Native.IsWindow(w.Wwf)) return;

            RECT cur;
            if (!TryGetClientRect(w.Opus, w.Wwf, out cur)) return;

            if (w.HasApplied && Same(w.NaturalWwf, natural) && Same(cur, w.AppliedWwf))
            {
                PlaceStrip(w);
                return;
            }

            Impose(w, natural, "copied from active");
        }

        static void Impose(WordWin w, RECT natural, string why)
        {
            var applied = new RECT
            {
                left = natural.left,
                top = natural.top + STRIP_H,
                right = natural.right,
                bottom = natural.bottom,
            };

            int h = applied.bottom - applied.top;
            if (h < 1) return;
            if (h < STRIP_H * 2)
            {
                // A tripwire for the collapse described above: better to stop shifting and
                // leave the strip wrong than to squeeze the document frame out of existence.
                Log(string.Format("WARN: refusing to squeeze \"{0}\" to h={1}", w.Title, h));
                return;
            }

            Log(string.Format("layout \"{0}\": natural ({1},{2} {3}x{4}) -> ({5},{6} {7}x{8})  [{9}]",
                              w.Title, natural.left, natural.top,
                              natural.right - natural.left, natural.bottom - natural.top,
                              applied.left, applied.top, applied.right - applied.left, h, why));

            w.NaturalWwf = natural;
            w.AppliedWwf = applied;
            w.HasApplied = true;

            RECT pc;
            if (Native.GetClientRect(w.Opus, out pc)) w.NaturalBottomGap = pc.bottom - natural.bottom;

            Native.SetWindowPos(w.Wwf, IntPtr.Zero,
                                applied.left, applied.top,
                                applied.right - applied.left, h,
                                Native.SWP_NOZORDER | Native.SWP_NOACTIVATE);

            w.StripPlaced = false;
            PlaceStrip(w);
        }

        static void PlaceStrip(WordWin w)
        {
            var want = new RECT
            {
                left = w.NaturalWwf.left,
                top = w.NaturalWwf.top,
                right = w.NaturalWwf.right,
                bottom = w.NaturalWwf.top + STRIP_H,
            };

            if (w.StripPlaced && Same(want, w.StripAt)) return;

            Native.SetWindowPos(w.Strip, Native.HWND_TOP,
                                want.left, want.top,
                                want.right - want.left, STRIP_H,
                                Native.SWP_NOACTIVATE);
            w.StripAt = want;
            w.StripPlaced = true;
            Native.InvalidateRect(w.Strip, IntPtr.Zero, false);
        }

        /// <summary>Rect of <paramref name="hWnd"/> in <paramref name="parent"/>'s client coordinates.</summary>
        static bool TryGetClientRect(IntPtr parent, IntPtr hWnd, out RECT outRect)
        {
            outRect = default;
            RECT r;
            if (!Native.GetWindowRect(hWnd, out r)) return false;

            POINT tl; tl.x = r.left; tl.y = r.top;
            if (!Native.ScreenToClient(parent, ref tl)) return false;

            outRect = new RECT
            {
                left = tl.x,
                top = tl.y,
                right = tl.x + (r.right - r.left),
                bottom = tl.y + (r.bottom - r.top),
            };
            return true;
        }

        static bool Same(RECT a, RECT b)
        {
            return a.left == b.left && a.top == b.top && a.right == b.right && a.bottom == b.bottom;
        }

        // --------------------------------------------------------------------- teardown

        static void Cleanup()
        {
            Log("restoring every window...");

            if (_hook != IntPtr.Zero) Native.UnhookWinEvent(_hook);
            if (_moveHook != IntPtr.Zero) Native.UnhookWinEvent(_moveHook);
            RestoreTaskbar();

            // Anything we hid for a drag must come back, or the user is left with documents
            // they cannot reach by any means.
            foreach (var w in _wins)
            {
                if (!w.HiddenByUs || !Native.IsWindow(w.Opus)) continue;
                Native.ShowWindow(w.Opus, Native.SW_SHOWNA);
                w.HiddenByUs = false;
            }

            foreach (var w in _wins)
            {
                if (w.Strip != IntPtr.Zero) Native.DestroyWindow(w.Strip);

                // Put the window back to its own size first, then refit the document frame to
                // it. Restoring the recorded natural rect verbatim would leave the frame sized
                // for the stack, and Word will not relayout it for us — the same finding that
                // makes SyncFollower necessary now applies in reverse.
                if (Native.IsWindow(w.Opus))
                {
                    Native.SetWindowPos(w.Opus, IntPtr.Zero,
                                        w.OriginalOuter.left, w.OriginalOuter.top,
                                        w.OriginalOuter.right - w.OriginalOuter.left,
                                        w.OriginalOuter.bottom - w.OriginalOuter.top,
                                        Native.SWP_NOZORDER | Native.SWP_NOACTIVATE);
                }

                RECT pc;
                if (w.HasApplied && Native.IsWindow(w.Wwf) && Native.GetClientRect(w.Opus, out pc))
                {
                    int top = w.NaturalWwf.top;
                    int bottom = pc.bottom - w.NaturalBottomGap;
                    if (bottom - top > STRIP_H)
                    {
                        Native.SetWindowPos(w.Wwf, IntPtr.Zero,
                                            w.NaturalWwf.left, top,
                                            pc.right - w.NaturalWwf.left, bottom - top,
                                            Native.SWP_NOZORDER | Native.SWP_NOACTIVATE);
                    }
                }
            }

            foreach (var br in new[] { _brBg, _brActive, _brInactive, _brEdge })
                if (br != IntPtr.Zero) Native.DeleteObject(br);

            Log("done.");
            _log.Dispose();
        }

        static void Log(string s)
        {
            Console.WriteLine(s);
            _log.WriteLine(s);
        }
    }
}
