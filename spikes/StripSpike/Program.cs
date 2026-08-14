using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace StripSpike
{
    /// <summary>
    /// WordTab spike 1 — out-of-process proof that we can carve a strip out of Word's layout.
    ///
    /// What it does: finds a visible Word `OpusApp` window and its `_WwF` document frame,
    /// pushes `_WwF` down by STRIP_H pixels, and parks our own coloured child window in the
    /// gap that opens up between the ribbon and the document.
    ///
    /// What it deliberately does NOT do: run inside WINWORD. The real add-in will subclass
    /// `OpusApp` from in-process and react to WM_SIZE synchronously. Here we watch from
    /// outside via a WinEvent hook plus a polling timer, because that needs no COM add-in
    /// and no C++ toolchain — which keeps this spike about the geometry question alone:
    /// does the strip survive resize, maximize/restore, and Backstage?
    /// </summary>
    internal static class Program
    {
        const int STRIP_H = 32;
        const uint POLL_MS = 30;

        static IntPtr _opus, _wwf, _strip;
        static uint _wordPid, _mainTid;
        static IntPtr _hook;

        // Delegates handed to unmanaged code must outlive the call — keep them rooted.
        static WndProcDelegate _wndProc;
        static TimerProcDelegate _timerProc;
        static WinEventProcDelegate _winEventProc;

        // `_natural` is where Word wants `_WwF`; `_applied` is where we moved it to.
        static RECT _natural, _applied;
        static bool _hasApplied;
        static RECT _stripAt;
        static bool _stripPlaced;

        static StreamWriter _log;

        static int Main()
        {
            _mainTid = Native.GetCurrentThreadId();
            var logPath = Path.Combine(AppContext.BaseDirectory, "spike.log");
            _log = new StreamWriter(logPath, false) { AutoFlush = true };
            Log("=== WordTab strip spike ===");
            Log("log file: " + logPath);

            if (!FindTarget())
            {
                Log("FAIL: no usable Word window found. Open Word with a document, then rerun.");
                return 1;
            }

            if (!CreateStrip())
            {
                Log("FAIL: could not create the strip window.");
                return 2;
            }

            Sync();

            _winEventProc = OnWinEvent;
            _hook = Native.SetWinEventHook(
                Native.EVENT_OBJECT_SHOW, Native.EVENT_OBJECT_LOCATIONCHANGE,
                IntPtr.Zero, _winEventProc, _wordPid, 0,
                Native.WINEVENT_OUTOFCONTEXT | Native.WINEVENT_SKIPOWNPROCESS);
            Log(_hook == IntPtr.Zero ? "WARN: WinEvent hook failed; relying on the poll timer"
                                     : "WinEvent hook installed");

            _timerProc = OnTimer;
            Native.SetTimer(IntPtr.Zero, UIntPtr.Zero, POLL_MS, _timerProc);

            Console.CancelKeyPress += (s, e) =>
            {
                e.Cancel = true;
                Native.PostThreadMessage(_mainTid, Native.WM_QUIT, IntPtr.Zero, IntPtr.Zero);
            };

            Log("");
            Log("RUNNING. Now try, in Word:");
            Log("  1. resize the window by dragging an edge");
            Log("  2. maximize, then restore");
            Log("  3. open File (Backstage), then go Back");
            Log("The magenta strip should stay pinned between the ribbon and the page.");
            Log("Press Ctrl+C here to stop and put Word's layout back.");
            Log("");

            MSG msg;
            while (Native.GetMessage(out msg, IntPtr.Zero, 0, 0) > 0)
            {
                Native.TranslateMessage(ref msg);
                Native.DispatchMessage(ref msg);
            }

            Cleanup();
            return 0;
        }

        // ---------------------------------------------------------------- target lookup

        static bool FindTarget()
        {
            var candidates = new List<IntPtr>();
            Native.EnumWindows((h, l) =>
            {
                if (Native.ClassOf(h) == "OpusApp") candidates.Add(h);
                return true;
            }, IntPtr.Zero);

            Log("OpusApp windows found: " + candidates.Count);

            IntPtr fg = Native.GetForegroundWindow();
            var usable = new List<IntPtr>();

            foreach (var h in candidates)
            {
                bool visible = Native.IsWindowVisible(h);
                bool cloaked = Native.IsCloaked(h);
                bool iconic = Native.IsIconic(h);
                IntPtr wwf = FindChild(h, "_WwF");
                RECT r; Native.GetWindowRect(h, out r);

                // Word keeps a permanent hidden background OpusApp with a full _Ww* tree,
                // and minimized windows park off-screen near -32000 (scaled by DPI, so
                // -21333 at 150%) at a stub size. Neither is something we can lay out into.
                bool ok = visible && !cloaked && !iconic && wwf != IntPtr.Zero
                          && (r.right - r.left) >= 200 && (r.bottom - r.top) >= 200;

                Log(string.Format(
                    "  hwnd=0x{0:X} visible={1} cloaked={2} minimized={3} _WwF={4} " +
                    "rect=({5},{6} {7}x{8}) usable={9} title=\"{10}\"",
                    h.ToInt64(), visible, cloaked, iconic, wwf == IntPtr.Zero ? "no" : "yes",
                    r.left, r.top, r.right - r.left, r.bottom - r.top, ok, Native.TitleOf(h)));

                if (!ok) continue;
                usable.Add(h);

                // Prefer whatever the user is actually looking at.
                if (h == fg || _opus == IntPtr.Zero)
                {
                    _opus = h;
                    _wwf = wwf;
                }
            }

            if (_opus == IntPtr.Zero)
            {
                Log("no usable OpusApp: every Word window is hidden, cloaked or minimized.");
                return false;
            }

            if (usable.Count > 1)
                Log("note: " + usable.Count + " usable Word windows; this spike drives one only.");

            Native.GetWindowThreadProcessId(_opus, out _wordPid);
            Log(string.Format("target: OpusApp=0x{0:X} _WwF=0x{1:X} pid={2}",
                              _opus.ToInt64(), _wwf.ToInt64(), _wordPid));
            return true;
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

        // ------------------------------------------------------------------ our window

        static bool CreateStrip()
        {
            IntPtr inst = Native.GetModuleHandle(null);
            _wndProc = StripWndProc;

            var wc = new WNDCLASSEX
            {
                cbSize = (uint)Marshal.SizeOf<WNDCLASSEX>(),
                style = 0,
                lpfnWndProc = Marshal.GetFunctionPointerForDelegate(_wndProc),
                hInstance = inst,
                hbrBackground = Native.CreateSolidBrush(Native.Rgb(200, 30, 120)),
                lpszClassName = "WordTabStripSpike",
            };

            if (Native.RegisterClassEx(ref wc) == 0)
            {
                Log("RegisterClassEx failed, err=" + Marshal.GetLastWin32Error());
                return false;
            }

            // Creating a child under a window owned by another process is legal Win32;
            // it implicitly attaches our input queue to Word's for the lifetime of the child.
            _strip = Native.CreateWindowEx(
                Native.WS_EX_NOACTIVATE, "WordTabStripSpike", "WordTab strip spike",
                Native.WS_CHILD | Native.WS_VISIBLE | Native.WS_CLIPSIBLINGS,
                0, 0, 10, STRIP_H, _opus, IntPtr.Zero, inst, IntPtr.Zero);

            if (_strip == IntPtr.Zero)
            {
                Log("CreateWindowEx failed, err=" + Marshal.GetLastWin32Error());
                return false;
            }

            Log(string.Format("strip window created: 0x{0:X}", _strip.ToInt64()));
            return true;
        }

        static IntPtr StripWndProc(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam)
        {
            return Native.DefWindowProc(hWnd, msg, wParam, lParam);
        }

        // ----------------------------------------------------------------- the geometry

        static void OnTimer(IntPtr h, uint m, UIntPtr id, uint t) { Sync(); }

        static void OnWinEvent(IntPtr hook, uint ev, IntPtr hWnd,
                               int idObject, int idChild, uint thread, uint time)
        {
            // EVENT_OBJECT_LOCATIONCHANGE also fires for the text caret on every keystroke.
            if (idObject != Native.OBJID_WINDOW) return;
            if (hWnd != _opus && hWnd != _wwf) return;
            Sync();
        }

        static void Sync()
        {
            if (!Native.IsWindow(_opus) || !Native.IsWindow(_wwf))
            {
                Log("target window closed — exiting");
                Native.PostThreadMessage(_mainTid, Native.WM_QUIT, IntPtr.Zero, IntPtr.Zero);
                return;
            }

            RECT cur;
            if (!TryGetClientRect(_wwf, out cur)) return;

            // Already sitting exactly where we last put it => Word has not relaid out.
            if (_hasApplied && Same(cur, _applied))
            {
                PlaceStrip();
                return;
            }

            // Anything else means Word rewrote the layout, so treat what we see as natural.
            _natural = cur;
            _applied = new RECT
            {
                left = cur.left,
                top = cur.top + STRIP_H,
                right = cur.right,
                bottom = cur.bottom,
            };

            int h = _applied.bottom - _applied.top;
            if (h < 1) return;

            _hasApplied = true;
            Native.SetWindowPos(_wwf, IntPtr.Zero,
                                _applied.left, _applied.top,
                                _applied.right - _applied.left, h,
                                Native.SWP_NOZORDER | Native.SWP_NOACTIVATE);

            Log(string.Format("relayout: _WwF natural ({0},{1} {2}x{3}) -> ({4},{5} {6}x{7})",
                              _natural.left, _natural.top,
                              _natural.right - _natural.left, _natural.bottom - _natural.top,
                              _applied.left, _applied.top,
                              _applied.right - _applied.left, h));

            _stripPlaced = false;
            PlaceStrip();
        }

        static void PlaceStrip()
        {
            var want = new RECT
            {
                left = _natural.left,
                top = _natural.top,
                right = _natural.right,
                bottom = _natural.top + STRIP_H,
            };

            if (_stripPlaced && Same(want, _stripAt)) return;

            Native.SetWindowPos(_strip, Native.HWND_TOP,
                                want.left, want.top,
                                want.right - want.left, STRIP_H,
                                Native.SWP_NOACTIVATE);
            _stripAt = want;
            _stripPlaced = true;
        }

        /// <summary>Rect of <paramref name="hWnd"/> in its parent's client coordinates.</summary>
        static bool TryGetClientRect(IntPtr hWnd, out RECT outRect)
        {
            outRect = default;
            RECT r;
            if (!Native.GetWindowRect(hWnd, out r)) return false;

            POINT tl; tl.x = r.left; tl.y = r.top;
            if (!Native.ScreenToClient(_opus, ref tl)) return false;

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
            Log("restoring Word's layout...");

            if (_hook != IntPtr.Zero) Native.UnhookWinEvent(_hook);
            if (_strip != IntPtr.Zero) Native.DestroyWindow(_strip);

            if (_hasApplied && Native.IsWindow(_wwf))
            {
                Native.SetWindowPos(_wwf, IntPtr.Zero,
                                    _natural.left, _natural.top,
                                    _natural.right - _natural.left,
                                    _natural.bottom - _natural.top,
                                    Native.SWP_NOZORDER | Native.SWP_NOACTIVATE);
            }

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
