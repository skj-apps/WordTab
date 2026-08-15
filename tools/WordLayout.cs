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

    // PW_RENDERFULLCONTENT (2) so a window that is partly off-screen or occluded still prints.
    // The bitmap is built on the PowerShell side: System.Drawing's types are loadable at runtime
    // but dragging them through Add-Type's reference resolution is a chase with no end.
    public static bool Print(IntPtr hwnd, IntPtr dc) { return PrintWindow(hwnd, dc, 2); }
}
