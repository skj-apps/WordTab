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
    [DllImport("user32.dll")] static extern uint GetDpiForWindow(IntPtr hwnd);
    public static int Dpi(IntPtr hwnd) { uint d = GetDpiForWindow(hwnd); return d >= 72 ? (int)d : 96; }

    public static void Close(IntPtr hwnd) { PostMessage(hwnd, 0x0010, IntPtr.Zero, IntPtr.Zero); }  // WM_CLOSE
    [DllImport("user32.dll")] static extern bool PostMessage(IntPtr hwnd, uint msg, IntPtr w, IntPtr l);
    public static RECT RectOf(IntPtr hwnd) { RECT r; GetWindowRect(hwnd, out r); return r; }
    public static bool Maximized(IntPtr hwnd) { return IsZoomed(hwnd); }
    public static string TitleOf(IntPtr hwnd) { StringBuilder t = new StringBuilder(256); GetWindowText(hwnd, t, 256); return t.ToString(); }

    public static void Resize(IntPtr hwnd, int cx, int cy)
    {
        SetWindowPos(hwnd, IntPtr.Zero, 0, 0, cx, cy, SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOMOVE);
    }

    public static void MoveTo(IntPtr hwnd, int x, int y)
    {
        SetWindowPos(hwnd, IntPtr.Zero, x, y, 0, 0, SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOSIZE);
    }

    public static IntPtr GetForeground() { return GetForegroundWindow(); }

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
    }

    // MulDiv's rounding, which is to nearest with ties away from zero - not Math.Round's, which is
    // to even and would put a tab boundary one pixel out at some DPIs and not others.
    static int Sc(int logical, int dpi) { return (logical * dpi + 48) / 96; }

    static bool Empty(RECT r) { return r.Right <= r.Left || r.Bottom <= r.Top; }
    public static bool IsEmptyRect(RECT r) { return Empty(r); }
    public static POINT Center(RECT r) { POINT p; p.X = (r.Left + r.Right) / 2; p.Y = (r.Top + r.Bottom) / 2; return p; }

    public static TabLayout Tabs(IntPtr strip, int count)
    {
        RECT s = RectOf(strip);
        int dpi = Dpi(strip);
        int w = s.Right - s.Left;            // a borderless child: client size == window size
        int h = s.Bottom - s.Top;

        int pad = Sc(6, dpi), gap = Sc(4, dpi);
        int minimum = Sc(70, dpi), desired = Sc(220, dpi);
        int plusW = Sc(26, dpi), closeW = Sc(16, dpi);

        int available = w - pad * 2 - plusW - gap;
        if (available < minimum) available = minimum;

        int width = desired;
        if (count > 0 && width * count > available) width = available / count;
        if (width < minimum) width = minimum;

        TabLayout layout = new TabLayout();
        layout.Tabs = new RECT[count];
        layout.Close = new RECT[count];

        for (int i = 0; i < count; i++)
        {
            RECT t;
            t.Left = pad + i * width;
            t.Right = t.Left + width - Sc(2, dpi);
            t.Top = Sc(3, dpi);
            t.Bottom = h;

            RECT c;
            c.Left = 0; c.Top = 0; c.Right = 0; c.Bottom = 0;
            if ((t.Right - t.Left) >= closeW * 3)
            {
                int middle = (t.Top + t.Bottom) / 2;
                c.Right = t.Right - Sc(6, dpi);
                c.Left = c.Right - closeW;
                c.Top = middle - closeW / 2;
                c.Bottom = c.Top + closeW;
                c.Left += s.Left; c.Right += s.Left; c.Top += s.Top; c.Bottom += s.Top;
            }

            t.Left += s.Left; t.Right += s.Left; t.Top += s.Top; t.Bottom += s.Top;
            layout.Tabs[i] = t;
            layout.Close[i] = c;
        }

        int after = (count > 0) ? (layout.Tabs[count - 1].Right - s.Left + gap) : pad;
        int limit = w - pad - plusW;
        if (after > limit) after = limit;
        if (after < pad) after = pad;

        RECT p;
        p.Left = after; p.Right = after + plusW;
        p.Top = Sc(6, dpi); p.Bottom = h - Sc(6, dpi);
        layout.HasPlus = p.Right <= w && p.Bottom > p.Top;
        p.Left += s.Left; p.Right += s.Left; p.Top += s.Top; p.Bottom += s.Top;
        layout.Plus = p;

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
}
