<#
.SYNOPSIS
  Drag a running Word window with injected mouse input, to exercise the in-process frame subclass.

.DESCRIPTION
  The claim the in-process subclass exists to settle is that Word moves a dragged window inside a
  modal move/size loop, faster than an out-of-process poller can follow. Proving that needs a real
  drag - the modal loop entered the way a user enters it, driven by real input - not a SetWindowPos
  call, which produces one position change and no modal loop at all.

  The drag is a genuine one: press the left button on the title bar, move, release. The awkward
  part is finding the title bar, because Word's caption row is mostly not caption - it holds the
  Quick Access Toolbar, the search box and the window buttons. So rather than guessing a pixel,
  this asks Word itself: WM_NCHITTEST is sent across the caption row until a point answers
  HTCAPTION. That is the same question Windows asks before starting a drag.

  Word must be in the foreground first, since injected input is delivered to the foreground
  thread's queue and a modal loop in a background window would never see it. A plain
  SetForegroundWindow from another process is refused, so the AttachThreadInput handshake is used -
  the same one spike 2 needed for tab switching.

  The measurement is written by the add-in to %LOCALAPPDATA%\WordTab\wordtab.log; this script only
  causes it. The cursor is put back where it was.

.PARAMETER Steps
  Number of injected mouse movements during the drag. Each one is a chance for Word to move.

.PARAMETER StepPixels
  Distance per movement, in physical pixels.

.PARAMETER DelayMs
  Pause between movements. Keep this below the 30ms the out-of-process spikes polled at, or the
  drag is slower than the thing it is being compared against and proves nothing.

.PARAMETER OneWay
  Drag in one direction and leave the window there, instead of going out and back. Use this to see
  whether the other frames followed: after a there-and-back drag every window ends where it began,
  so the final positions cannot tell a perfect follow from no follow at all.

.EXAMPLE
  pwsh -File tools\drive-drag.ps1
  pwsh -File tools\drive-drag.ps1 -Steps 200 -DelayMs 4
#>
[CmdletBinding()]
param(
    [int]$Steps = 120,
    [int]$StepPixels = 5,
    [int]$DelayMs = 6,
    [switch]$OneWay
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$source = @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

public static class WordDrag
{
    [DllImport("user32.dll")] static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc proc, IntPtr param);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr hwnd, StringBuilder text, int max);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll")] static extern bool IsIconic(IntPtr hwnd);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr hwnd, out RECT rect);
    [DllImport("user32.dll")] static extern bool AttachThreadInput(uint from, uint to, bool attach);
    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr hwnd);
    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] static extern bool BringWindowToTop(IntPtr hwnd);
    [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr hwnd, uint msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")] static extern bool GetCursorPos(out POINT p);
    [DllImport("user32.dll")] static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] static extern int GetSystemMetrics(int index);
    [DllImport("user32.dll", SetLastError = true)] static extern uint SendInput(uint count, INPUT[] inputs, int size);
    [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();

    delegate bool EnumProc(IntPtr hwnd, IntPtr param);

    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
    [StructLayout(LayoutKind.Sequential)] struct POINT { public int X, Y; }

    // sizeof(INPUT) is 40 on x64: the union sits at offset 8. SendInput silently returns 0 if the
    // size does not match exactly, which looks identical to "input was blocked".
    [StructLayout(LayoutKind.Sequential)] struct INPUT { public uint type; public InputUnion u; }

    [StructLayout(LayoutKind.Explicit)]
    struct InputUnion
    {
        [FieldOffset(0)] public MOUSEINPUT mi;
        [FieldOffset(0)] public KEYBDINPUT ki;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct MOUSEINPUT { public int dx, dy; public uint mouseData, dwFlags, time; public IntPtr extra; }

    [StructLayout(LayoutKind.Sequential)]
    struct KEYBDINPUT { public ushort vk, scan; public uint flags, time; public IntPtr extra; }

    const uint INPUT_MOUSE = 0;
    const uint MOUSEEVENTF_MOVE = 0x0001, MOUSEEVENTF_LEFTDOWN = 0x0002, MOUSEEVENTF_LEFTUP = 0x0004;
    const uint MOUSEEVENTF_ABSOLUTE = 0x8000, MOUSEEVENTF_VIRTUALDESK = 0x4000;
    const uint WM_NCHITTEST = 0x0084;
    const int  HTCAPTION = 2;
    const int  SM_XVIRTUALSCREEN = 76, SM_YVIRTUALSCREEN = 77, SM_CXVIRTUALSCREEN = 78, SM_CYVIRTUALSCREEN = 79;

    public static void MakeDpiAware() { SetProcessDPIAware(); }

    public static List<IntPtr> FindFrames(int pid)
    {
        List<IntPtr> found = new List<IntPtr>();
        EnumWindows(delegate(IntPtr hwnd, IntPtr param)
        {
            uint owner;
            GetWindowThreadProcessId(hwnd, out owner);
            if (pid != 0 && owner != (uint)pid) return true;

            StringBuilder cls = new StringBuilder(64);
            GetClassName(hwnd, cls, 64);
            if (cls.ToString() == "OpusApp" && IsWindowVisible(hwnd) && !IsIconic(hwnd)) found.Add(hwnd);
            return true;
        }, IntPtr.Zero);
        return found;
    }

    public static RECT RectOf(IntPtr hwnd) { RECT r; GetWindowRect(hwnd, out r); return r; }

    public static string ForegroundClass()
    {
        StringBuilder cls = new StringBuilder(64);
        GetClassName(GetForegroundWindow(), cls, 64);
        return cls.ToString();
    }

    public static bool IsForeground(IntPtr hwnd) { return GetForegroundWindow() == hwnd; }

    // A background process cannot simply steal the foreground. Attaching our input queue to the
    // target's makes the system treat the call as coming from the active thread, which it allows.
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

    // Find a point on the title row that actually drags the window, by trying.
    //
    // Asking Word does not work. WM_NCHITTEST answers HTCAPTION for every point across the title
    // row, but the row is covered edge to edge by a NetUIHWND child - Word draws its own caption -
    // so a click at most of those points goes to the Quick Access Toolbar or the search box and
    // the window never moves. Both were measured: a grab at top+4 came out as SC_SIZE (Windows'
    // resize border wins there), and grabs at 20%, 60%, 68%, 74% and 80% of the width did nothing
    // at all.
    //
    // So this calibrates: a short test drag at each candidate, keeping the one that moves the
    // window, and putting it back. Slower than a formula and immune to where the toolbars happen
    // to end on this build, at this window width, for this user's Quick Access Toolbar.
    public static bool FindDragPoint(IntPtr hwnd, out int px, out int py)
    {
        px = 0; py = 0;
        RECT r = RectOf(hwnd);
        int width = r.Right - r.Left;

        // The gap between the Quick Access Toolbar and the search box first: that is where the
        // empty caption is on a default Word window.
        double[] fractions = { 0.30, 0.36, 0.24, 0.42, 0.62, 0.72, 0.80, 0.16 };
        int[] rows = { r.Top + 30, r.Top + 22, r.Top + 40 };

        foreach (int y in rows)
        {
            foreach (double fraction in fractions)
            {
                int x = r.Left + (int)(width * fraction);

                IntPtr hit = SendMessage(hwnd, WM_NCHITTEST, IntPtr.Zero,
                                         (IntPtr)((y << 16) | (x & 0xFFFF)));
                if ((int)hit != HTCAPTION)
                    continue;

                RECT before = RectOf(hwnd);
                TinyDrag(x, y, 8, 4);
                RECT after = RectOf(hwnd);

                if (after.Left != before.Left || after.Top != before.Top)
                {
                    // Undo the probe, then hand back the point that worked.
                    TinyDrag(x + 32, y + 16, -8, -4);
                    px = x; py = y;
                    return true;
                }
            }
        }
        return false;
    }

    // Four small steps with the button down - enough to commit a move, small enough to undo.
    static void TinyDrag(int x, int y, int stepX, int stepY)
    {
        Inject(x, y, 0);
        Thread.Sleep(80);
        Inject(x, y, MOUSEEVENTF_LEFTDOWN);
        Thread.Sleep(120);
        for (int i = 1; i <= 4; i++)
        {
            Inject(x + i * stepX, y + i * stepY, 0);
            Thread.Sleep(25);
        }
        Inject(x + 4 * stepX, y + 4 * stepY, MOUSEEVENTF_LEFTUP);
        Thread.Sleep(220);
    }

    static void Inject(int x, int y, uint extraFlags)
    {
        int vx = GetSystemMetrics(SM_XVIRTUALSCREEN);
        int vy = GetSystemMetrics(SM_YVIRTUALSCREEN);
        int vw = GetSystemMetrics(SM_CXVIRTUALSCREEN);
        int vh = GetSystemMetrics(SM_CYVIRTUALSCREEN);

        INPUT[] input = new INPUT[1];
        input[0].type = INPUT_MOUSE;
        input[0].u.mi.dx = (int)(((long)(x - vx) * 65535) / (vw - 1));
        input[0].u.mi.dy = (int)(((long)(y - vy) * 65535) / (vh - 1));
        input[0].u.mi.dwFlags = MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_VIRTUALDESK | extraFlags;
        SendInput(1, input, Marshal.SizeOf(typeof(INPUT)));
    }

    // A real press-move-release drag. Returns the number of movements injected while the button
    // was down.
    //
    // oneWay leaves the window displaced, so the follower windows' final positions are themselves
    // evidence. The default goes out and back, which keeps the window on screen but ends where it
    // started - and an end-state comparison then proves nothing either way.
    public static int Drag(int grabX, int grabY, int steps, int stepPixels, int delayMs, bool oneWay)
    {
        POINT restore;
        GetCursorPos(out restore);

        Inject(grabX, grabY, 0);
        Thread.Sleep(120);
        Inject(grabX, grabY, MOUSEEVENTF_LEFTDOWN);
        Thread.Sleep(120);

        int x = grabX, y = grabY;
        int half = steps / 2;
        for (int i = 0; i < steps; i++)
        {
            int direction = (oneWay || i < half) ? 1 : -1;
            x += direction * stepPixels;
            y += direction * (stepPixels / 2);
            Inject(x, y, 0);
            Thread.Sleep(delayMs);
        }

        Inject(x, y, MOUSEEVENTF_LEFTUP);
        Thread.Sleep(400);
        SetCursorPos(restore.X, restore.Y);
        return steps;
    }
}
'@

Add-Type -TypeDefinition $source -Language CSharp

# Physical pixels throughout. This rig runs at 150%, and without this every rectangle read below
# would be silently scaled.
[WordDrag]::MakeDpiAware() | Out-Null

$word = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue)
if ($word.Count -eq 0) { throw 'Word is not running. Start it first.' }

$frames = @()
foreach ($process in $word) { $frames += [WordDrag]::FindFrames($process.Id) }
if ($frames.Count -eq 0) { throw 'No visible OpusApp frames found. (A Word started by COM automation is hidden until Visible is set.)' }

Write-Host "==> $($frames.Count) visible Word frame(s)" -ForegroundColor Cyan
foreach ($frame in $frames) {
    $r = [WordDrag]::RectOf($frame)
    Write-Host ("    0x{0:X}  ({1},{2})  {3}x{4}" -f [int64]$frame, $r.Left, $r.Top, ($r.Right - $r.Left), ($r.Bottom - $r.Top)) -ForegroundColor DarkGray
}

$target = $frames[0]
$before = @{}
foreach ($frame in $frames) { $before[[int64]$frame] = [WordDrag]::RectOf($frame) }

Write-Host "==> Foreground before: $([WordDrag]::ForegroundClass())" -ForegroundColor DarkGray
if ([WordDrag]::Focus($target)) {
    Write-Host "    Word is foreground (injected input will reach its modal loop)" -ForegroundColor Green
} else {
    Write-Warning "Could not bring Word to the foreground (still $([WordDrag]::ForegroundClass())). The modal loop will not see the injected input."
}

Write-Host '==> Calibrating: looking for a title-row point that actually drags' -ForegroundColor Cyan
$px = 0; $py = 0
if ([WordDrag]::FindDragPoint($target, [ref]$px, [ref]$py)) {
    Write-Host "    grab point $px,$py moves the window" -ForegroundColor Green
} else {
    # The resize border along the top edge always works and enters the same modal loop, so the
    # frame-rate measurement survives even when no draggable caption pixel can be found.
    $r = [WordDrag]::RectOf($target)
    $px = $r.Left + [int](($r.Right - $r.Left) * 0.5)
    $py = $r.Top + 4
    Write-Warning "No draggable caption point found. Falling back to the top resize border ($px,$py) - this measures the same modal loop, but as a resize rather than a move, so the follow demonstration will not trigger."
}

Write-Host "==> Dragging: $Steps moves of ${StepPixels}px, ${DelayMs}ms apart" -ForegroundColor Cyan
$injected = [WordDrag]::Drag($px, $py, $Steps, $StepPixels, $DelayMs, [bool]$OneWay)
Write-Host "    injected $injected movements with the button down" -ForegroundColor Green

Write-Host '==> Frames after the drag' -ForegroundColor Cyan
foreach ($frame in $frames) {
    $r = [WordDrag]::RectOf($frame)
    $b = $before[[int64]$frame]
    $tag = if ($frame -eq $target) { 'dragged ' } else { 'follower' }
    Write-Host ("    $tag 0x{0:X}  ({1},{2}) -> ({3},{4})   delta ({5},{6})" -f `
        [int64]$frame, $b.Left, $b.Top, $r.Left, $r.Top, ($r.Left - $b.Left), ($r.Top - $b.Top)) -ForegroundColor DarkGray
}

Write-Host ''
Write-Host "Measurement: $env:LOCALAPPDATA\WordTab\wordtab.log (lines starting 'drag')" -ForegroundColor Gray
