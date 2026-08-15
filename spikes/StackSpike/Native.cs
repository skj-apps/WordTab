using System;
using System.Runtime.InteropServices;
using System.Text;

namespace StackSpike
{
    [StructLayout(LayoutKind.Sequential)]
    internal struct RECT { public int left, top, right, bottom; }

    [StructLayout(LayoutKind.Sequential)]
    internal struct POINT { public int x, y; }

    [StructLayout(LayoutKind.Sequential)]
    internal struct MSG
    {
        public IntPtr hwnd;
        public uint message;
        public IntPtr wParam;
        public IntPtr lParam;
        public uint time;
        public POINT pt;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    internal struct WNDCLASSEX
    {
        public uint cbSize;
        public uint style;
        public IntPtr lpfnWndProc;
        public int cbClsExtra;
        public int cbWndExtra;
        public IntPtr hInstance;
        public IntPtr hIcon;
        public IntPtr hCursor;
        public IntPtr hbrBackground;
        [MarshalAs(UnmanagedType.LPWStr)] public string lpszMenuName;
        [MarshalAs(UnmanagedType.LPWStr)] public string lpszClassName;
        public IntPtr hIconSm;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct PAINTSTRUCT
    {
        public IntPtr hdc;
        public int fErase;
        public RECT rcPaint;
        public int fRestore;
        public int fIncUpdate;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 32)] public byte[] rgbReserved;
    }

    internal delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    internal delegate IntPtr WndProcDelegate(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);
    internal delegate void TimerProcDelegate(IntPtr hWnd, uint msg, UIntPtr id, uint time);
    internal delegate void WinEventProcDelegate(IntPtr hook, uint ev, IntPtr hWnd,
                                                int idObject, int idChild, uint thread, uint time);

    /// <summary>
    /// The shell's taskbar list. `DeleteTab` is the documented way to take a window off the
    /// taskbar without touching its styles, and it works cross-process — which matters here,
    /// because the windows belong to WINWORD, not to us.
    /// </summary>
    [ComImport, Guid("56FDF342-FD6D-11d0-958A-006097C9A090"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface ITaskbarList
    {
        void HrInit();
        void AddTab(IntPtr hWnd);
        void DeleteTab(IntPtr hWnd);
        void ActivateTab(IntPtr hWnd);
        void SetActiveAlt(IntPtr hWnd);
    }

    [ComImport, Guid("56FDF344-FD6D-11d0-958A-006097C9A090"), ClassInterface(ClassInterfaceType.None)]
    internal class TaskbarListClass { }

    internal static class Native
    {
        // ---- window styles / flags -------------------------------------------------
        public const int WS_CHILD = unchecked((int)0x40000000);
        public const int WS_VISIBLE = unchecked((int)0x10000000);
        public const int WS_CLIPSIBLINGS = unchecked((int)0x04000000);
        public const int WS_EX_NOACTIVATE = unchecked((int)0x08000000);

        public const uint SWP_NOSIZE = 0x0001;
        public const uint SWP_NOMOVE = 0x0002;
        public const uint SWP_NOZORDER = 0x0004;
        public const uint SWP_NOACTIVATE = 0x0010;
        public const uint SWP_NOOWNERZORDER = 0x0200;

        public static readonly IntPtr HWND_TOP = IntPtr.Zero;

        public const uint WM_DESTROY = 0x0002;
        public const uint WM_PAINT = 0x000F;
        public const uint WM_ERASEBKGND = 0x0014;
        public const uint WM_QUIT = 0x0012;
        public const uint WM_LBUTTONDOWN = 0x0201;
        public const uint WM_APP = 0x8000;

        /// <summary>Posted to the main thread by the console reader: wParam = tab index.</summary>
        public const uint WM_SWITCH = WM_APP + 1;

        public const uint EVENT_OBJECT_SHOW = 0x8002;
        public const uint EVENT_OBJECT_LOCATIONCHANGE = 0x800B;
        public const uint WINEVENT_OUTOFCONTEXT = 0x0000;
        public const uint WINEVENT_SKIPOWNPROCESS = 0x0002;

        public const int OBJID_WINDOW = 0;
        public const int DWMWA_CLOAKED = 14;

        public const int TRANSPARENT = 1;
        public const int DEFAULT_GUI_FONT = 17;

        public const uint DT_SINGLELINE = 0x0020;
        public const uint DT_CENTER = 0x0001;
        public const uint DT_VCENTER = 0x0004;
        public const uint DT_END_ELLIPSIS = 0x8000;
        public const uint DT_NOPREFIX = 0x0800;

        // ---- enumeration ------------------------------------------------------------
        [DllImport("user32.dll")]
        public static extern bool EnumWindows(EnumWindowsProc cb, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern bool EnumChildWindows(IntPtr parent, EnumWindowsProc cb, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder buf, int max);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder buf, int max);

        [DllImport("user32.dll")]
        public static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool IsZoomed(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);

        [DllImport("dwmapi.dll")]
        public static extern int DwmGetWindowAttribute(IntPtr hWnd, int attr, out int value, int size);

        // ---- geometry / z-order -----------------------------------------------------
        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT r);

        [DllImport("user32.dll")]
        public static extern bool GetClientRect(IntPtr hWnd, out RECT r);

        [DllImport("user32.dll")]
        public static extern bool ScreenToClient(IntPtr hWnd, ref POINT p);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr after,
                                               int x, int y, int cx, int cy, uint flags);

        [DllImport("user32.dll")]
        public static extern bool BringWindowToTop(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool AttachThreadInput(uint from, uint to, bool attach);

        // ---- window creation --------------------------------------------------------
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern ushort RegisterClassEx(ref WNDCLASSEX wc);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr CreateWindowEx(int exStyle, string cls, string name, int style,
                                                   int x, int y, int w, int h,
                                                   IntPtr parent, IntPtr menu, IntPtr inst, IntPtr param);

        [DllImport("user32.dll")]
        public static extern bool DestroyWindow(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr DefWindowProc(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr GetModuleHandle(string name);

        // ---- painting ---------------------------------------------------------------
        [DllImport("user32.dll")]
        public static extern IntPtr BeginPaint(IntPtr hWnd, out PAINTSTRUCT ps);

        [DllImport("user32.dll")]
        public static extern bool EndPaint(IntPtr hWnd, ref PAINTSTRUCT ps);

        [DllImport("user32.dll")]
        public static extern int FillRect(IntPtr hdc, ref RECT r, IntPtr brush);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int DrawText(IntPtr hdc, string text, int len, ref RECT r, uint fmt);

        [DllImport("user32.dll")]
        public static extern bool InvalidateRect(IntPtr hWnd, IntPtr r, bool erase);

        [DllImport("gdi32.dll")]
        public static extern IntPtr CreateSolidBrush(uint color);

        [DllImport("gdi32.dll")]
        public static extern int SetBkMode(IntPtr hdc, int mode);

        [DllImport("gdi32.dll")]
        public static extern uint SetTextColor(IntPtr hdc, uint color);

        [DllImport("gdi32.dll")]
        public static extern IntPtr SelectObject(IntPtr hdc, IntPtr obj);

        [DllImport("gdi32.dll")]
        public static extern IntPtr GetStockObject(int index);

        [DllImport("gdi32.dll")]
        public static extern bool DeleteObject(IntPtr obj);

        // ---- message loop -----------------------------------------------------------
        [DllImport("user32.dll")]
        public static extern int GetMessage(out MSG msg, IntPtr hWnd, uint min, uint max);

        [DllImport("user32.dll")]
        public static extern bool TranslateMessage(ref MSG msg);

        [DllImport("user32.dll")]
        public static extern IntPtr DispatchMessage(ref MSG msg);

        [DllImport("user32.dll")]
        public static extern bool PostThreadMessage(uint tid, uint msg, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        public static extern uint GetCurrentThreadId();

        [DllImport("user32.dll")]
        public static extern UIntPtr SetTimer(IntPtr hWnd, UIntPtr id, uint ms, TimerProcDelegate proc);

        [DllImport("user32.dll")]
        public static extern bool KillTimer(IntPtr hWnd, UIntPtr id);

        // ---- winevent hook ----------------------------------------------------------
        [DllImport("user32.dll")]
        public static extern IntPtr SetWinEventHook(uint min, uint max, IntPtr mod,
                                                    WinEventProcDelegate proc,
                                                    uint pid, uint tid, uint flags);

        [DllImport("user32.dll")]
        public static extern bool UnhookWinEvent(IntPtr hook);

        // ---- helpers ----------------------------------------------------------------
        public static string ClassOf(IntPtr hWnd)
        {
            var sb = new StringBuilder(256);
            GetClassName(hWnd, sb, sb.Capacity);
            return sb.ToString();
        }

        public static string TitleOf(IntPtr hWnd)
        {
            var sb = new StringBuilder(512);
            GetWindowText(hWnd, sb, sb.Capacity);
            return sb.ToString();
        }

        public static bool IsCloaked(IntPtr hWnd)
        {
            int v;
            return DwmGetWindowAttribute(hWnd, DWMWA_CLOAKED, out v, sizeof(int)) == 0 && v != 0;
        }

        public static uint Rgb(byte r, byte g, byte b)
        {
            return (uint)(r | (g << 8) | (b << 16));
        }

        public static int LoWord(IntPtr v) { return (short)(v.ToInt64() & 0xFFFF); }
        public static int HiWord(IntPtr v) { return (short)((v.ToInt64() >> 16) & 0xFFFF); }
    }
}
