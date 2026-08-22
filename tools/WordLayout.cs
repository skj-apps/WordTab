using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

public static class WordLayout
{
    [DllImport("user32.dll")] static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc proc, IntPtr param);
    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr parent, EnumProc proc, IntPtr param);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr hwnd, StringBuilder text, int max);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int max);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll")] static extern bool IsIconic(IntPtr hwnd);
    [DllImport("user32.dll")] static extern bool IsZoomed(IntPtr hwnd);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr hwnd, out RECT rect);
    [DllImport("user32.dll")] static extern bool GetClientRect(IntPtr hwnd, out RECT rect);
    [DllImport("user32.dll")] static extern bool ScreenToClient(IntPtr hwnd, ref POINT p);
    [DllImport("user32.dll")] static extern bool SetWindowPos(IntPtr hwnd, IntPtr after, int x, int y, int cx, int cy, uint flags);
    [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr hwnd, int cmd);
    [DllImport("user32.dll")] static extern bool AttachThreadInput(uint from, uint to, bool attach);
    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr hwnd);
    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] static extern bool BringWindowToTop(IntPtr hwnd);
    [DllImport("user32.dll", SetLastError = true)] static extern uint SendInput(uint count, INPUT[] inputs, int size);
    [DllImport("user32.dll")] static extern int GetSystemMetrics(int index);
    [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();
    [DllImport("user32.dll")] static extern bool PrintWindow(IntPtr hwnd, IntPtr dc, uint flags);

    delegate bool EnumProc(IntPtr hwnd, IntPtr param);

    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
    [StructLayout(LayoutKind.Sequential)] public struct POINT { public int X, Y; }

    [StructLayout(LayoutKind.Sequential)] struct INPUT { public uint type; public InputUnion u; }
    [StructLayout(LayoutKind.Explicit)] struct InputUnion
    {
        [FieldOffset(0)] public MOUSEINPUT mi;
        [FieldOffset(0)] public KEYBDINPUT ki;
    }
    [StructLayout(LayoutKind.Sequential)] struct MOUSEINPUT { public int dx, dy; public uint mouseData, dwFlags, time; public IntPtr extra; }
    [StructLayout(LayoutKind.Sequential)] struct KEYBDINPUT { public ushort vk, scan; public uint flags, time; public IntPtr extra; }

    const uint INPUT_MOUSE = 0, INPUT_KEYBOARD = 1, KEYEVENTF_KEYUP = 2;
    const uint MOUSEEVENTF_MOVE = 0x0001, MOUSEEVENTF_LEFTDOWN = 0x0002, MOUSEEVENTF_LEFTUP = 0x0004;
    const uint MOUSEEVENTF_MIDDLEDOWN = 0x0020, MOUSEEVENTF_MIDDLEUP = 0x0040;
    const uint MOUSEEVENTF_RIGHTDOWN = 0x0008, MOUSEEVENTF_RIGHTUP = 0x0010;
    const uint MOUSEEVENTF_ABSOLUTE = 0x8000, MOUSEEVENTF_VIRTUALDESK = 0x4000;
    public const int SW_MAXIMIZE = 3, SW_RESTORE = 9;
    const uint SWP_NOZORDER = 0x0004, SWP_NOACTIVATE = 0x0010, SWP_NOMOVE = 0x0002, SWP_NOSIZE = 0x0001;

    public class Child
    {
        public IntPtr Hwnd;
        public string Class;
        public string Title;
        public bool Visible;
        public int Left, Top, Right, Bottom;   // in the frame's client coordinates
        public int Width  { get { return Right - Left; } }
        public int Height { get { return Bottom - Top; } }
    }

    public static void MakeDpiAware() { SetProcessDPIAware(); }

    public static List<IntPtr> Frames(int pid)
    {
        List<IntPtr> found = new List<IntPtr>();
        EnumWindows(delegate(IntPtr hwnd, IntPtr param)
        {
            uint owner;
            GetWindowThreadProcessId(hwnd, out owner);
            if (pid != 0 && owner != (uint)pid) return true;

            StringBuilder cls = new StringBuilder(64);
            GetClassName(hwnd, cls, 64);
            if (cls.ToString() == "OpusApp" && IsWindowVisible(hwnd) && !IsIconic(hwnd))
            {
                RECT r;
                GetWindowRect(hwnd, out r);
                if (r.Right - r.Left > 200 && r.Bottom - r.Top > 200) found.Add(hwnd);
            }
            return true;
        }, IntPtr.Zero);
        return found;
    }

    // Every top-level window the process owns, visible or not. Backstage, dialogs and Word's
    // hidden background windows all show up here and nowhere else.
    public static List<Child> TopLevel(int pid)
    {
        List<Child> found = new List<Child>();
        EnumWindows(delegate(IntPtr hwnd, IntPtr param)
        {
            uint owner;
            GetWindowThreadProcessId(hwnd, out owner);
            if (pid != 0 && owner != (uint)pid) return true;

            StringBuilder cls = new StringBuilder(96);
            GetClassName(hwnd, cls, 96);
            RECT r;
            GetWindowRect(hwnd, out r);

            Child c = new Child();
            c.Hwnd = hwnd;
            c.Class = cls.ToString();
            c.Title = TitleOf(hwnd);
            c.Visible = IsWindowVisible(hwnd);
            c.Left = r.Left; c.Top = r.Top; c.Right = r.Right; c.Bottom = r.Bottom;
            found.Add(c);
            return true;
        }, IntPtr.Zero);
        return found;
    }

    // Every direct child of the frame, in the frame's client coordinates - the space the strip and
    // `_WwF` are actually positioned in, so the numbers compare without conversion.
    public static List<Child> Children(IntPtr frame)
    {
        List<Child> found = new List<Child>();
        EnumChildWindows(frame, delegate(IntPtr hwnd, IntPtr param)
        {
            StringBuilder cls = new StringBuilder(96);
            GetClassName(hwnd, cls, 96);

            RECT r;
            if (!GetWindowRect(hwnd, out r)) return true;

            POINT tl; tl.X = r.Left; tl.Y = r.Top;
            if (!ScreenToClient(frame, ref tl)) return true;

            Child c = new Child();
            c.Hwnd = hwnd;
            c.Class = cls.ToString();
            c.Visible = IsWindowVisible(hwnd);
            c.Left = tl.X; c.Top = tl.Y;
            c.Right = tl.X + (r.Right - r.Left);
            c.Bottom = tl.Y + (r.Bottom - r.Top);
            found.Add(c);
            return true;
        }, IntPtr.Zero);
        return found;
    }

    public static RECT ClientOf(IntPtr hwnd) { RECT r; GetClientRect(hwnd, out r); return r; }

    // Windows 10 1607 and later. The rig runs at 150%, so anything that mirrors a DPI-scaled
    // layout computed inside the add-in needs this rather than an assumed 96.
    //
    // **It must consult HKCU\Software\WordTab\TabDpi for the same reason it calls GetDpiForWindow at
    // all.** The code below this line is a second copy of the add-in's ComputeLayout - same pad, gap,
    // minimum, desired, plus, close and chevron constants - and its whole job is to predict where the
    // add-in put a tab so a suite can click it. The moment the add-in is told to build itself at a
    // DPI the window is not running at, a harness that asks Windows instead is computing slots for a
    // strip that does not exist, and every click lands somewhere else. Which is silent: the click
    // still hits the row, just the wrong tab.
    //
    // So this reads the override first, exactly as DpiOf does in strip.cpp, and falls through to the
    // system in the ordinary case where the value is absent.
    [DllImport("user32.dll")] static extern uint GetDpiForWindow(IntPtr hwnd);
    public static int Dpi(IntPtr hwnd)
    {
        int forced = ForcedDpi();
        if (forced != 0) return forced;
        uint d = GetDpiForWindow(hwnd);
        return d >= 72 ? (int)d : 96;
    }

    // Read on every call rather than cached, so a suite that steps the override mid-run gets the new
    // answer without reloading the type - which it could not do anyway, since Add-Type is once per
    // process.
    //
    // RegGetValueW rather than Microsoft.Win32.Registry: naming any assembly in Add-Type replaces
    // PowerShell's default reference set, so every suite passes its own list, and reaching for the
    // Registry class would mean editing all of them to add one more. This is the same call the
    // add-in makes.
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode)]
    static extern int RegGetValueW(IntPtr hkey, string subKey, string value, uint flags,
                                   IntPtr type, ref int data, ref uint size);

    public static int ForcedDpi()
    {
        int data = 0;
        uint size = 4;
        IntPtr HKEY_CURRENT_USER = new IntPtr(unchecked((int)0x80000001));
        const uint RRF_RT_REG_DWORD = 0x00000010;
        int rc = RegGetValueW(HKEY_CURRENT_USER, @"Software\WordTab", "TabDpi",
                              RRF_RT_REG_DWORD, IntPtr.Zero, ref data, ref size);
        if (rc != 0) return 0;
        return (data >= 72 && data <= 480) ? data : 0;
    }

    public static void Close(IntPtr hwnd) { PostMessage(hwnd, 0x0010, IntPtr.Zero, IntPtr.Zero); }  // WM_CLOSE

    // The window's own close, as the title bar's x raises it: WM_SYSCOMMAND with SC_CLOSE.
    //
    // Not the same message as Close above, and the difference is the point. WM_CLOSE is what the
    // add-in itself posts to close one tab; SC_CLOSE is only ever the user. The add-in tells them
    // apart to decide whether closing takes the whole stack with it, so a check of that behaviour
    // has to send the one the user sends.
    public static void SysClose(IntPtr hwnd) { PostMessage(hwnd, 0x0112, (IntPtr)0xF060, IntPtr.Zero); }  // WM_SYSCOMMAND, SC_CLOSE
    [DllImport("user32.dll")] static extern bool PostMessage(IntPtr hwnd, uint msg, IntPtr w, IntPtr l);
    public static RECT RectOf(IntPtr hwnd) { RECT r; GetWindowRect(hwnd, out r); return r; }
    public static bool Maximized(IntPtr hwnd) { return IsZoomed(hwnd); }
    // 1024 rather than 256, and the extra room is not for window titles. WordTab's tooltip carries
    // its whole text on the window itself so a suite can read what was drawn rather than infer it,
    // and that is a document name and a full path with a newline between them - past 256 for any
    // document more than a few folders deep. A widened buffer is a strict superset: no caller can
    // tell the difference on a string that already fitted.
    public static string TitleOf(IntPtr hwnd) { StringBuilder t = new StringBuilder(1024); GetWindowText(hwnd, t, 1024); return t.ToString(); }

    public static void Resize(IntPtr hwnd, int cx, int cy)
    {
        SetWindowPos(hwnd, IntPtr.Zero, 0, 0, cx, cy, SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOMOVE);
    }

    // The usable area of the monitor a window is on - the screen less the taskbar. Needed because the
    // add-in places a torn-off window relative to it, so a check that compared against the whole
    // screen would be measuring a different rectangle from the one the product used.
    [StructLayout(LayoutKind.Sequential)]
    struct MONITORINFO { public int cbSize; public RECT rcMonitor; public RECT rcWork; public int dwFlags; }
    [DllImport("user32.dll")] static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint flags);
    [DllImport("user32.dll")] static extern bool GetMonitorInfo(IntPtr monitor, ref MONITORINFO info);

    public static RECT WorkArea(IntPtr hwnd)
    {
        MONITORINFO info = new MONITORINFO();
        info.cbSize = Marshal.SizeOf(typeof(MONITORINFO));
        IntPtr monitor = MonitorFromWindow(hwnd, 2 /* MONITOR_DEFAULTTONEAREST */);
        if (monitor != IntPtr.Zero && GetMonitorInfo(monitor, ref info))
            return info.rcWork;
        return RectOf(hwnd);
    }

    public static void MoveTo(IntPtr hwnd, int x, int y)
    {
        SetWindowPos(hwnd, IntPtr.Zero, x, y, 0, 0, SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOSIZE);
    }

    public static IntPtr GetForeground() { return GetForegroundWindow(); }

    // Put a window in front of the others WITHOUT giving it the keyboard: HWND_TOP, SWP_NOACTIVATE.
    //
    // This is not a synthetic gesture. It is what Word does to a document it has just opened on the
    // work rig - the window arrives on top, the focus does not follow it, and the row went on
    // highlighting the tab of the document that was no longer on screen. Focus and z-order are two
    // different answers to "which window is the user looking at", and this is how a suite on a rig
    // where they always agree can ask the question anyway.
    public static void RaiseWithoutFocus(IntPtr hwnd)
    {
        SetWindowPos(hwnd, IntPtr.Zero, 0, 0, 0, 0, SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE);
    }

    // The tooltip's appearance delay IS this number, and its auto-hide is ten times it: the add-in
    // asks Windows rather than carrying a constant of its own, so that WordTab's tooltip arrives when
    // the user's other tooltips arrive. A suite that waited a hardcoded 500ms would be asserting
    // against the default of a setting instead of against the product.
    [DllImport("user32.dll")] static extern int GetDoubleClickTime();
    public static int DoubleClickTime() { return GetDoubleClickTime(); }

    [DllImport("user32.dll")] static extern IntPtr GetWindowLongPtr(IntPtr hwnd, int index);
    const int GWL_EXSTYLE = -20;
    public const long WS_EX_TOOLWINDOW = 0x00000080;

    public static long ExStyle(IntPtr hwnd) { return (long)GetWindowLongPtr(hwnd, GWL_EXSTYLE); }
    public static bool IsToolWindow(IntPtr hwnd) { return (ExStyle(hwnd) & WS_EX_TOOLWINDOW) != 0; }
    public static bool Minimized(IntPtr hwnd) { return IsIconic(hwnd); }

    [DllImport("user32.dll")] static extern bool IsWindow(IntPtr hwnd);
    // Named to avoid colliding with the Win32 import above, which has to stay private.
    public static bool IsWindow2(IntPtr hwnd) { return IsWindow(hwnd); }

    // The virtual desktop, for full-screen captures. Alt+Tab and the taskbar cannot be enumerated
    // through any API, so the only honest way to check them is to photograph the screen.
    public static RECT ScreenRect()
    {
        RECT r;
        r.Left = GetSystemMetrics(76); r.Top = GetSystemMetrics(77);
        r.Right = r.Left + GetSystemMetrics(78); r.Bottom = r.Top + GetSystemMetrics(79);
        return r;
    }

    // Hold Alt and tap Tab, leaving the switcher on screen for a caller to photograph. AltRelease
    // must always be called afterwards, including on failure: a stuck Alt key is a wrecked desktop.
    public static void AltTabHold()
    {
        Key(0x12, false);           // VK_MENU down, held
        Thread.Sleep(120);
        Key(0x09, false);           // VK_TAB
        Thread.Sleep(60);
        Key(0x09, true);
    }

    public static void AltRelease()
    {
        Key(0x12, true);
        Thread.Sleep(200);
    }

    public static void Show(IntPtr hwnd, int cmd) { ShowWindow(hwnd, cmd); }

    public static bool Focus(IntPtr hwnd)
    {
        uint dummy;
        uint targetThread = GetWindowThreadProcessId(hwnd, out dummy);
        uint ourThread = GetCurrentThreadId();
        uint foreThread = GetWindowThreadProcessId(GetForegroundWindow(), out dummy);

        AttachThreadInput(ourThread, targetThread, true);
        AttachThreadInput(ourThread, foreThread, true);
        BringWindowToTop(hwnd);
        SetForegroundWindow(hwnd);
        AttachThreadInput(ourThread, foreThread, false);
        AttachThreadInput(ourThread, targetThread, false);

        Thread.Sleep(400);
        return GetForegroundWindow() == hwnd;
    }

    // A real click, in screen coordinates. Being the foreground window is not the same as having
    // been clicked in: Word's ribbon ignores injected keystrokes after the window was activated
    // programmatically, and one click into the document fixes it.
    public static void Click(int x, int y)
    {
        Move(x, y, 0);
        Thread.Sleep(150);
        Move(x, y, MOUSEEVENTF_LEFTDOWN);
        Thread.Sleep(80);
        Move(x, y, MOUSEEVENTF_LEFTUP);
        Thread.Sleep(200);
    }

    // Press, move, release - a real drag, entering the modal move/size loop the way a user does.
    // A SetWindowPos call produces one relayout; this produces one every frame, which is the case
    // the strip has to survive without making Word stutter.
    public static int DragBy(int startX, int startY, int steps, int dx, int dy, int delayMs)
    {
        Move(startX, startY, 0);
        Thread.Sleep(150);
        Move(startX, startY, MOUSEEVENTF_LEFTDOWN);
        Thread.Sleep(150);

        int x = startX, y = startY;
        for (int i = 0; i < steps; i++)
        {
            x += dx; y += dy;
            Move(x, y, 0);
            Thread.Sleep(delayMs);
        }

        Move(x, y, MOUSEEVENTF_LEFTUP);
        Thread.Sleep(400);
        return steps;
    }

    // ---------------------------------------------------------------------------------------------
    // Dragging inside one of *our* windows, as opposed to DragBy above, which is about driving
    // Word's own modal move loop by the caption.
    //
    // The intermediate movements are the point. A drag implemented as press, teleport, release
    // crosses no threshold, produces one WM_MOUSEMOVE at the destination, and would let a strip that
    // reorders nothing at all pass a reorder check. These deliver a real path.
    //
    // Split into hold and release so a check can photograph the strip while the button is still
    // down. Every caller of DragHold must call DragRelease from a finally block: a left button left
    // down is a wrecked desktop, exactly like a held Alt key.
    // ---------------------------------------------------------------------------------------------

    public static void DragHold(int x1, int y1, int x2, int y2, int steps, int delayMs)
    {
        if (steps < 1) steps = 1;
        Move(x1, y1, 0);
        Thread.Sleep(200);
        Move(x1, y1, MOUSEEVENTF_LEFTDOWN);
        Thread.Sleep(250);

        for (int i = 1; i <= steps; i++)
        {
            int x = x1 + (int)((long)(x2 - x1) * i / steps);
            int y = y1 + (int)((long)(y2 - y1) * i / steps);
            Move(x, y, 0);
            Thread.Sleep(delayMs);
        }
        Thread.Sleep(200);
    }

    // Carry on from where DragHold left off, button still down. Lets one gesture be asserted in
    // stages - a short move that must not reorder anything, then a long one that must.
    public static void DragMoveTo(int x1, int y1, int x2, int y2, int steps, int delayMs)
    {
        if (steps < 1) steps = 1;
        for (int i = 1; i <= steps; i++)
        {
            int x = x1 + (int)((long)(x2 - x1) * i / steps);
            int y = y1 + (int)((long)(y2 - y1) * i / steps);
            Move(x, y, 0);
            Thread.Sleep(delayMs);
        }
        Thread.Sleep(200);
    }

    public static void DragRelease(int x, int y)
    {
        Move(x, y, MOUSEEVENTF_LEFTUP);
        Thread.Sleep(500);
    }

    public static void DragTo(int x1, int y1, int x2, int y2, int steps, int delayMs)
    {
        try { DragHold(x1, y1, x2, y2, steps, delayMs); }
        finally { DragRelease(x2, y2); }
    }

    // A tap of the right button without letting the left one go - the cancel gesture. The left
    // button is released afterwards where it is, which is what a user who changed their mind does.
    public static void RightTap(int x, int y)
    {
        Move(x, y, MOUSEEVENTF_RIGHTDOWN);
        Thread.Sleep(120);
        Move(x, y, MOUSEEVENTF_RIGHTUP);
        Thread.Sleep(250);
    }

    static void Move(int x, int y, uint extraFlags)
    {
        int vx = GetSystemMetrics(76), vy = GetSystemMetrics(77);
        int vw = GetSystemMetrics(78), vh = GetSystemMetrics(79);

        INPUT[] input = new INPUT[1];
        input[0].type = INPUT_MOUSE;
        input[0].u.mi.dx = (int)(((long)(x - vx) * 65535) / (vw - 1));
        input[0].u.mi.dy = (int)(((long)(y - vy) * 65535) / (vh - 1));
        input[0].u.mi.dwFlags = MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_VIRTUALDESK | extraFlags;
        SendInput(1, input, Marshal.SizeOf(typeof(INPUT)));
    }

    static void Key(ushort vk, bool up)
    {
        INPUT[] input = new INPUT[1];
        input[0].type = INPUT_KEYBOARD;
        input[0].u.ki.vk = vk;
        input[0].u.ki.flags = up ? KEYEVENTF_KEYUP : 0;
        SendInput(1, input, Marshal.SizeOf(typeof(INPUT)));
    }

    // Alt+F opens Backstage; Escape leaves it. Injected rather than sent as a message because
    // Backstage is Word's own UI and only responds to real input.
    //
    // The pauses are load-bearing. Sending Alt-down and F-down back to back does nothing at all -
    // measured, and silently: Word stays exactly as it was, and a check written around it passes
    // while testing nothing. Pressing and releasing Alt first (the KeyTips route) also does
    // nothing. Alt held down, a pause, then F, is the sequence that works.
    public static void OpenBackstage()
    {
        Key(0x12, false);           // VK_MENU, held
        Thread.Sleep(60);
        Key(0x46, false);           // 'F'
        Thread.Sleep(60);
        Key(0x46, true);
        Key(0x12, true);
    }

    // Backstage is a child of the frame - class FullpageUIHost - covering the whole client area,
    // and Word hides every other child while it is up, ours included. So this is both how you know
    // it opened and why the strip correctly disappears while it is open.
    public static bool BackstageOpen(IntPtr frame)
    {
        foreach (Child c in Children(frame))
            if (c.Visible && c.Class == "FullpageUIHost") return true;
        return false;
    }

    public static void CloseBackstage()
    {
        Key(0x1B, false);           // VK_ESCAPE
        Thread.Sleep(40);
        Key(0x1B, true);
    }

    // ---------------------------------------------------------------------------------------------
    // The tab strip's layout, mirroring ComputeLayout in src\native\strip.cpp.
    //
    // There is no way to share these numbers across the process boundary, so there are two copies
    // and this is the second. If they ever drift the injected clicks land somewhere other than what
    // was drawn and the assertions fail loudly - which is the intended failure. The alternative, a
    // test that computes its own idea of where a tab is and clicks confidently into the gap between
    // two of them, passes while testing nothing.
    //
    // Everything comes back in **screen** coordinates, ready to click.
    // ---------------------------------------------------------------------------------------------

    public class TabLayout
    {
        public RECT[] Tabs;      // one per document, left to right
        public RECT[] Close;     // the X on each; empty when the tab is too narrow to carry one
        public RECT Plus;        // the new-document button
        public bool HasPlus;

        // The scrolling row. Track is the band the tabs live in and are clipped to; everything to
        // the right of it is the fixed cluster the row can never reach.
        public RECT Track;
        public RECT Prev, Next;  // the scroll buttons, empty unless the row overflows
        public bool HasNav;
        public bool CanPrev, CanNext;
        public int Scroll;       // as clamped, which may be less than what was asked for
        public int MaxScroll;
        public int Width;        // one tab's pitch, which is also one click of a scroll button
    }

    // MulDiv's rounding, which is to nearest with ties away from zero - not Math.Round's, which is
    // to even and would put a tab boundary one pixel out at some DPIs and not others.
    static int Sc(int logical, int dpi) { return (logical * dpi + 48) / 96; }

    // The same scaling, for a script that has to work out how narrow a window must be before a
    // given number of tabs stops fitting. Hard-coding a pixel count there would make the suite
    // silently measure nothing on a rig at a different DPI.
    public static int Scale(int logical, int dpi) { return Sc(logical, dpi); }

    static bool Empty(RECT r) { return r.Right <= r.Left || r.Bottom <= r.Top; }
    public static bool IsEmptyRect(RECT r) { return Empty(r); }
    public static POINT Center(RECT r) { POINT p; p.X = (r.Left + r.Right) / 2; p.Y = (r.Top + r.Bottom) / 2; return p; }

    // The row as it looks with the scroll at zero, which is where it always is unless the tabs
    // overflow - so this is the whole of it for every suite that does not deliberately overflow one.
    public static TabLayout Tabs(IntPtr strip, int count) { return Tabs(strip, count, 0, true); }

    public static TabLayout Tabs(IntPtr strip, int count, int scroll) { return Tabs(strip, count, scroll, true); }

    // `scroll` has to be supplied because there is no way to ask: it is state inside Word's process,
    // and the add-in does not publish it. What makes that workable is that it is knowable at the two
    // positions worth testing - a row that has just overflowed is at 0, and a row scrolled harder
    // than it can move is at MaxScroll - and that everywhere else it is exactly 0, because a row
    // that fits cannot be scrolled at all.
    //
    // `scrollEnabled` mirrors HKCU\Software\WordTab\TabScroll: false is the squeeze layout, where
    // the tabs divide the track between them with no minimum width.
    //
    // Both this and ComputeLayout in src\native\strip.cpp assume the buttons are switched on.
    public static TabLayout Tabs(IntPtr strip, int count, int scroll, bool scrollEnabled)
    {
        RECT s = RectOf(strip);
        int dpi = Dpi(strip);
        int w = s.Right - s.Left;            // a borderless child: client size == window size
        int h = s.Bottom - s.Top;

        int pad = Sc(6, dpi), gap = Sc(4, dpi);
        int minimum = Sc(70, dpi), desired = Sc(220, dpi);
        int plusW = Sc(26, dpi), closeW = Sc(16, dpi), chevW = Sc(20, dpi);
        int hair = Sc(2, dpi);

        int available = w - pad * 2 - plusW - gap;
        if (available < minimum) available = minimum;

        int width = desired;
        if (count > 0 && width * count > available) width = available / count;
        if (width < minimum) width = minimum;

        bool overflow = (count > 0 && width * count > available);

        TabLayout layout = new TabLayout();
        layout.Tabs = new RECT[count];
        layout.Close = new RECT[count];

        RECT track;
        track.Left = pad; track.Right = w - pad; track.Top = 0; track.Bottom = h;

        RECT prev, next, p;
        prev.Left = prev.Top = prev.Right = prev.Bottom = 0;
        next.Left = next.Top = next.Right = next.Bottom = 0;
        p.Left = p.Top = p.Right = p.Bottom = 0;
        bool hasNav = false;

        if (overflow)
        {
            p.Right = w - pad; p.Left = p.Right - plusW;
            p.Top = Sc(6, dpi); p.Bottom = h - Sc(6, dpi);

            next.Right = p.Left - gap; next.Left = next.Right - chevW;
            prev.Right = next.Left; prev.Left = prev.Right - chevW;
            next.Top = prev.Top = Sc(6, dpi);
            next.Bottom = prev.Bottom = h - Sc(6, dpi);

            track.Right = prev.Left - gap;
            hasNav = scrollEnabled;

            if (!scrollEnabled)
            {
                track.Right = p.Left - gap;
                prev.Left = prev.Top = prev.Right = prev.Bottom = 0;
                next.Left = next.Top = next.Right = next.Bottom = 0;
            }
        }

        int trackW = track.Right - track.Left;
        if (trackW < 0) trackW = 0;

        if (overflow && !scrollEnabled)
        {
            width = (count > 0) ? (trackW / count) : desired;
            if (width < 1) width = 1;
            overflow = false;
        }

        int maxScroll = 0;
        if (overflow)
        {
            maxScroll = count * width - trackW;
            if (maxScroll < 0) maxScroll = 0;
        }

        if (scroll > maxScroll) scroll = maxScroll;
        if (scroll < 0) scroll = 0;

        for (int i = 0; i < count; i++)
        {
            RECT t;
            t.Left = track.Left + i * width - scroll;
            t.Right = t.Left + width - hair;
            if (t.Right <= t.Left) t.Right = t.Left + 1;
            t.Top = Sc(3, dpi);
            t.Bottom = h;

            RECT c;
            c.Left = 0; c.Top = 0; c.Right = 0; c.Bottom = 0;
            if ((t.Right - t.Left) >= closeW * 3)
            {
                int middle = (t.Top + t.Bottom) / 2;
                int cr = t.Right - Sc(6, dpi);
                int cl = cr - closeW;
                if (cl >= track.Left && cr <= track.Right)
                {
                    c.Right = cr;
                    c.Left = cl;
                    c.Top = middle - closeW / 2;
                    c.Bottom = c.Top + closeW;
                    c.Left += s.Left; c.Right += s.Left; c.Top += s.Top; c.Bottom += s.Top;
                }
            }

            t.Left += s.Left; t.Right += s.Left; t.Top += s.Top; t.Bottom += s.Top;
            layout.Tabs[i] = t;
            layout.Close[i] = c;
        }

        if (!overflow)
        {
            int after = (count > 0) ? (layout.Tabs[count - 1].Right - s.Left + gap) : pad;
            int limit = w - pad - plusW;
            if (after > limit) after = limit;
            if (after < pad) after = pad;

            p.Left = after; p.Right = after + plusW;
            p.Top = Sc(6, dpi); p.Bottom = h - Sc(6, dpi);
        }

        layout.HasPlus = p.Right <= w && p.Bottom > p.Top;
        layout.HasNav = hasNav;
        layout.CanPrev = hasNav && scroll > 0;
        layout.CanNext = hasNav && scroll < maxScroll;
        layout.Scroll = scroll;
        layout.MaxScroll = maxScroll;
        layout.Width = width;

        p.Left += s.Left; p.Right += s.Left; p.Top += s.Top; p.Bottom += s.Top;
        layout.Plus = p;

        if (hasNav)
        {
            prev.Left += s.Left; prev.Right += s.Left; prev.Top += s.Top; prev.Bottom += s.Top;
            next.Left += s.Left; next.Right += s.Left; next.Top += s.Top; next.Bottom += s.Top;
        }
        layout.Prev = prev;
        layout.Next = next;

        track.Left += s.Left; track.Right += s.Left; track.Top += s.Top; track.Bottom += s.Top;
        layout.Track = track;

        return layout;
    }

    // Park the pointer somewhere without clicking - how a hover is produced for a camera.
    public static void MouseTo(int x, int y) { Move(x, y, 0); Thread.Sleep(120); }

    // What the system thinks is under a point, which is the only authority on where a click will
    // land. Worth asking directly: a click that produces nothing looks identical whether it missed
    // the window or the window ignored it, and these two answer which.
    [DllImport("user32.dll")] static extern IntPtr WindowFromPoint(POINT p);
    [DllImport("user32.dll")] static extern bool GetCursorPos(out POINT p);

    public static IntPtr WindowAt(int x, int y) { POINT p; p.X = x; p.Y = y; return WindowFromPoint(p); }
    public static POINT Cursor() { POINT p; GetCursorPos(out p); return p; }

    // What SHAPE the pointer currently is, as opposed to where it is.
    //
    // The only way a check script can read feedback the add-in gives outside the strip's own pixels: a
    // tab dragged clear of the row is over Word's document, where nothing of ours may paint, so the
    // cursor is what says the gesture has changed meaning. GetCursorInfo reports the system-wide
    // cursor, so this reads what the add-in set from inside Word.
    //
    // Comparable against SystemCursor() because the standard cursors are shared objects: LoadCursorW
    // with a NULL instance hands every process the same handle for the same one, so "is this
    // IDC_SIZEALL" is an equality test rather than a bitmap comparison. Measured, not assumed - see
    // the tear-off section of check-reorder.ps1.
    [StructLayout(LayoutKind.Sequential)]
    struct CURSORINFO { public int cbSize; public int flags; public IntPtr hCursor; public POINT pt; }
    [DllImport("user32.dll")] static extern bool GetCursorInfo(ref CURSORINFO info);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern IntPtr LoadCursorW(IntPtr inst, IntPtr name);

    public static IntPtr CursorShape()
    {
        CURSORINFO info = new CURSORINFO();
        info.cbSize = Marshal.SizeOf(typeof(CURSORINFO));
        if (!GetCursorInfo(ref info)) return IntPtr.Zero;
        return info.hCursor;
    }

    // IDC_ARROW is 32512 and IDC_SIZEALL is 32646; passed as the resource ordinals they are.
    public static IntPtr SystemCursor(int id) { return LoadCursorW(IntPtr.Zero, new IntPtr(id)); }

    // Which process owns a window. Word puts a Protected View document in a sandboxed WINWORD of its
    // own, and an add-in's Application object speaks only for the process it is loaded in - so "same
    // process?" is the difference between a document the object model can see and one it cannot,
    // however ordinary the window looks on screen.
    public static int PidOf(IntPtr hwnd) { uint pid; GetWindowThreadProcessId(hwnd, out pid); return (int)pid; }

    // The Word object model for ONE named window, rather than for whichever instance the running
    // object table feels like handing out. A probe that does New-Object -ComObject Word.Application
    // while Word is already up got an Application reporting Documents.Count = 0 and a Windows
    // collection with no Count on it - measured - because the class object it bound to was not the
    // Word the user is looking at. This asks a window directly instead: OBJID_NATIVEOM on Word's
    // _WwG document pane returns that window's own Word.Window object, in that window's own process.
    //
    // The add-in never needs this - it is handed Application by Word in OnConnection and so is
    // already inside the right process. This exists so a test script can be a SECOND oracle for what
    // the add-in reports, and so a Protected View document, which lives in a sandboxed WINWORD of
    // its own, can be asked the same questions as a normal one.
    [DllImport("oleacc.dll")] static extern int AccessibleObjectFromWindow(
        IntPtr hwnd, uint objectId, ref Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object obj);

    const uint OBJID_NATIVEOM = 0xFFFFFFF0;

    // The _WwG pane, which is the window that answers OBJID_NATIVEOM. IntPtr.Zero if the frame has
    // no document pane - which is the Start screen and an emptied window, not a failure.
    public static IntPtr DocumentPane(IntPtr frame)
    {
        foreach (Child c in Children(frame)) if (c.Class == "_WwG") return c.Hwnd;
        return IntPtr.Zero;
    }

    // Null rather than an exception when the window will not answer, because "Word would not answer"
    // is a result this measurement is looking for, not an error in taking it.
    public static object NativeOm(IntPtr frame)
    {
        IntPtr pane = DocumentPane(frame);
        if (pane == IntPtr.Zero) return null;
        Guid iid = new Guid("00020400-0000-0000-C000-000000000046");   // IID_IDispatch
        object obj;
        int hr = AccessibleObjectFromWindow(pane, OBJID_NATIVEOM, ref iid, out obj);
        return hr == 0 ? obj : null;
    }

    // A keystroke, for making a document dirty so the save prompt can be provoked on purpose.
    public static void Press(ushort vk) { Key(vk, false); Thread.Sleep(40); Key(vk, true); Thread.Sleep(60); }

    // Ctrl+key. Needed for Ctrl+W - Word's *document* close, which is the only way to reach the
    // state where a Word window is left on screen with no document in it. WM_CLOSE is not the same
    // thing: it closes the window, and on the last window it closes Word.
    public static void CtrlPress(ushort vk)
    {
        Key(0x11, false);                       // VK_CONTROL
        Thread.Sleep(60);
        Key(vk, false); Thread.Sleep(60); Key(vk, true);
        Thread.Sleep(60);
        Key(0x11, true);
        Thread.Sleep(150);
    }

    // The first child of a window, or IntPtr.Zero if it has none. This is the add-in's own document
    // test, mirrored: `_WwF` is Word's document *frame* and outlives the document, so "is there a
    // document" is "is there anything inside the document frame" - see StripHasDocument in
    // src\native\strip.cpp. EnumChildWindows would answer the same question far more expensively.
    [DllImport("user32.dll", EntryPoint = "GetWindow")] static extern IntPtr GetWindowRel(IntPtr hwnd, uint cmd);
    const uint GW_CHILD = 5;
    public static IntPtr FirstChild(IntPtr hwnd) { return GetWindowRel(hwnd, GW_CHILD); }
    public static string ClassOf(IntPtr hwnd)
    {
        StringBuilder cls = new StringBuilder(96);
        GetClassName(hwnd, cls, 96);
        return cls.ToString();
    }

    public static void MiddleClick(int x, int y)
    {
        Move(x, y, 0);
        Thread.Sleep(150);
        Move(x, y, MOUSEEVENTF_MIDDLEDOWN);
        Thread.Sleep(80);
        Move(x, y, MOUSEEVENTF_MIDDLEUP);
        Thread.Sleep(200);
    }

    // Press on one point and release on another - the gesture that must *not* fire a button.
    public static void PressAndSlideOff(int x1, int y1, int x2, int y2)
    {
        Move(x1, y1, 0);
        Thread.Sleep(150);
        Move(x1, y1, MOUSEEVENTF_LEFTDOWN);
        Thread.Sleep(120);
        Move(x2, y2, 0);
        Thread.Sleep(120);
        Move(x2, y2, MOUSEEVENTF_LEFTUP);
        Thread.Sleep(250);
    }

    // ---------------------------------------------------------------------------------------------
    // Popup menus, read from outside the process that owns them.
    //
    // A menu is not a window with children that can be enumerated - the items live in an HMENU, and
    // the window on screen (class #32768) only draws them. MN_GETHMENU is the documented way across:
    // send it to that window and it answers with the HMENU, and menu handles live in the shared user
    // handle table, so GetMenuString and friends read them from any process. This is how a menu can
    // be asserted item by item rather than photographed and eyeballed.
    // ---------------------------------------------------------------------------------------------

    [DllImport("user32.dll")] static extern IntPtr SendMessageTimeout(IntPtr hwnd, uint msg, IntPtr w, IntPtr l, uint flags, uint timeout, out IntPtr result);
    [DllImport("user32.dll")] static extern int GetMenuItemCount(IntPtr menu);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetMenuString(IntPtr menu, uint item, StringBuilder text, int max, uint flags);
    [DllImport("user32.dll")] static extern uint GetMenuState(IntPtr menu, uint item, uint flags);
    [DllImport("user32.dll")] static extern bool GetMenuItemRect(IntPtr hwnd, IntPtr menu, uint item, out RECT rect);
    [DllImport("user32.dll")] static extern bool IsWindowEnabled(IntPtr hwnd);

    const uint MN_GETHMENU = 0x01E1, SMTO_ABORTIFHUNG = 0x0002;
    const uint MF_BYPOSITION = 0x0400, MF_GRAYED = 0x0001, MF_DISABLED = 0x0002, MF_SEPARATOR = 0x0800;

    public class MenuItem
    {
        public int Index;
        public uint Id;
        public string Text;          // "-" for a separator
        public bool Enabled;
        public bool Separator;
        public RECT Rect;            // screen coordinates, empty if the menu is not on screen
        public bool HasRect;
    }

    // The popup menu currently on screen for a process, or IntPtr.Zero. Visible-only: Windows keeps
    // menu windows around after they close, and a hidden one is a menu that is not being shown.
    public static IntPtr PopupMenuWindow(int pid)
    {
        IntPtr found = IntPtr.Zero;
        EnumWindows(delegate(IntPtr hwnd, IntPtr param)
        {
            uint owner;
            GetWindowThreadProcessId(hwnd, out owner);
            if (pid != 0 && owner != (uint)pid) return true;

            StringBuilder cls = new StringBuilder(64);
            GetClassName(hwnd, cls, 64);
            if (cls.ToString() == "#32768" && IsWindowVisible(hwnd)) { found = hwnd; return false; }
            return true;
        }, IntPtr.Zero);
        return found;
    }

    public static IntPtr MenuOf(IntPtr menuWindow)
    {
        IntPtr result;
        if (SendMessageTimeout(menuWindow, MN_GETHMENU, IntPtr.Zero, IntPtr.Zero,
                               SMTO_ABORTIFHUNG, 2000, out result) == IntPtr.Zero)
            return IntPtr.Zero;
        return result;
    }

    // Every item, in order, with the separators kept in: a menu is as much about how it is grouped
    // as about what is on it, and a check that silently dropped the separators could not tell the
    // difference between "Close All" next to "New Document" and a rule between them.
    public static List<MenuItem> MenuItems(IntPtr menuWindow, IntPtr owner)
    {
        List<MenuItem> items = new List<MenuItem>();
        IntPtr menu = MenuOf(menuWindow);
        if (menu == IntPtr.Zero) return items;

        int count = GetMenuItemCount(menu);
        for (int i = 0; i < count; i++)
        {
            uint state = GetMenuState(menu, (uint)i, MF_BYPOSITION);
            MenuItem item = new MenuItem();
            item.Index = i;
            item.Separator = (state & MF_SEPARATOR) != 0;
            item.Enabled = (state & (MF_GRAYED | MF_DISABLED)) == 0;

            StringBuilder text = new StringBuilder(128);
            GetMenuString(menu, (uint)i, text, 128, MF_BYPOSITION);
            item.Text = item.Separator ? "-" : text.ToString();
            item.Id = item.Separator ? 0 : (uint)GetMenuItemID(menu, i);

            RECT r;
            item.HasRect = GetMenuItemRect(owner, menu, (uint)i, out r);
            item.Rect = r;
            items.Add(item);
        }
        return items;
    }

    [DllImport("user32.dll")] static extern uint GetMenuItemID(IntPtr menu, int pos);

    // Whether a window is accepting input. A modal dialog disables the window that owns it, which is
    // how the add-in tells "Word is asking the user something" from "the user said no".
    public static bool Enabled(IntPtr hwnd) { return IsWindowEnabled(hwnd); }

    public static void RightClick(int x, int y)
    {
        Move(x, y, 0);
        Thread.Sleep(150);
        Move(x, y, MOUSEEVENTF_RIGHTDOWN);
        Thread.Sleep(80);
        Move(x, y, MOUSEEVENTF_RIGHTUP);
        Thread.Sleep(400);
    }

    // Right-press on one point and release on another - the gesture that must produce no menu.
    public static void RightPressAndSlideOff(int x1, int y1, int x2, int y2)
    {
        Move(x1, y1, 0);
        Thread.Sleep(150);
        Move(x1, y1, MOUSEEVENTF_RIGHTDOWN);
        Thread.Sleep(120);
        Move(x2, y2, 0);
        Thread.Sleep(120);
        Move(x2, y2, MOUSEEVENTF_RIGHTUP);
        Thread.Sleep(350);
    }

    // PW_RENDERFULLCONTENT (2) so a window that is partly off-screen or occluded still prints.
    // The bitmap is built on the PowerShell side: System.Drawing's types are loadable at runtime
    // but dragging them through Add-Type's reference resolution is a chase with no end.
    public static bool Print(IntPtr hwnd, IntPtr dc) { return PrintWindow(hwnd, dc, 2); }

    [DllImport("shell32.dll")] static extern void SHChangeNotify(int eventId, uint flags, IntPtr item1, IntPtr item2);
    const int SHCNE_ASSOCCHANGED = 0x08000000;

    // Tell the shell a file-association setting changed. Needed by probe-titles.ps1 when it flips
    // HideFileExt: shell32 caches that per process, so a registry write alone is not something a
    // running program can be expected to see. Processes started afterwards read it fresh either way,
    // which is why this is belt and braces rather than the mechanism.
    public static void BroadcastSettingChange()
    {
        SHChangeNotify(SHCNE_ASSOCCHANGED, 0, IntPtr.Zero, IntPtr.Zero);
    }

    // ---- where a keystroke actually goes, and the two-modifier chord ---------------------------
    //
    // Both of these are additions only; nothing above changed. All eleven suites compile against
    // this file, so it may only ever grow.

    [StructLayout(LayoutKind.Sequential)]
    public struct GUITHREADINFO
    {
        public int cbSize;
        public int flags;
        public IntPtr hwndActive;
        public IntPtr hwndFocus;
        public IntPtr hwndCapture;
        public IntPtr hwndMenuOwner;
        public IntPtr hwndMoveSize;
        public IntPtr hwndCaret;
        public RECT rcCaret;
    }

    [DllImport("user32.dll")] static extern bool GetGUIThreadInfo(uint thread, ref GUITHREADINFO info);

    // The KEYBOARD focus inside another process, which GetForegroundWindow cannot tell you.
    //
    // This matters because a WM_KEYDOWN is delivered to the focus window and to nothing else -
    // keyboard messages do not travel up to parents the way a WM_CONTEXTMENU does. So "which window
    // does Word give the keyboard to" decides whether a subclass can ever see a keystroke, and it is
    // a measurement rather than a thing to reason about. hwndFocus comes back IntPtr.Zero when the
    // asked-about thread does not own the foreground, which is itself the answer to a different
    // question and must not be read as "no focus window".
    public static GUITHREADINFO ThreadGui(IntPtr hwnd)
    {
        uint pid;
        uint thread = GetWindowThreadProcessId(hwnd, out pid);
        GUITHREADINFO gui = new GUITHREADINFO();
        gui.cbSize = Marshal.SizeOf(typeof(GUITHREADINFO));
        if (!GetGUIThreadInfo(thread, ref gui)) { gui.hwndActive = IntPtr.Zero; gui.hwndFocus = IntPtr.Zero; }
        return gui;
    }

    // Ctrl+Shift+key. Word's own "previous window" is Ctrl+Shift+F6, and a chord with two modifiers
    // cannot be built by calling CtrlPress with a shifted key: both modifiers have to be down across
    // the whole of the target key's down-and-up, or Word sees a different chord.
    public static void CtrlShiftPress(ushort vk)
    {
        Key(0x11, false);                       // VK_CONTROL
        Key(0x10, false);                       // VK_SHIFT
        Thread.Sleep(60);
        Key(vk, false); Thread.Sleep(60); Key(vk, true);
        Thread.Sleep(60);
        Key(0x10, true);
        Key(0x11, true);
        Thread.Sleep(150);
    }
}
