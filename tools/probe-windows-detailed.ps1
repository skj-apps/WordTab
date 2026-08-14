# Wider probe: every top-level window in every WINWORD process, with styles, cloaking, and foreground.
Add-Type @'
using System;
using System.Text;
using System.Runtime.InteropServices;
public class P2 {
  public delegate bool EnumProc(IntPtr hWnd, IntPtr lParam);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr p);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr h, EnumProc cb, IntPtr p);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowTextW(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr GetParent(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr GetWindow(IntPtr h, uint cmd);
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll", EntryPoint="GetWindowLongPtrW")] public static extern IntPtr GetWindowLongPtr(IntPtr h, int i);
  [DllImport("dwmapi.dll")] public static extern int DwmGetWindowAttribute(IntPtr h, int attr, out int val, int sz);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
  public static string Cls(IntPtr h){ var sb=new StringBuilder(256); GetClassName(h,sb,256); return sb.ToString(); }
  public static string Txt(IntPtr h){ var sb=new StringBuilder(512); GetWindowTextW(h,sb,512); return sb.ToString(); }
  public static uint Pid(IntPtr h){ uint p; GetWindowThreadProcessId(h, out p); return p; }
  public static int Cloaked(IntPtr h){ int v=0; try { DwmGetWindowAttribute(h, 14, out v, 4); } catch {} return v; }
}
'@ -ErrorAction Stop

$procs = @{}
Get-Process -EA SilentlyContinue | ForEach-Object { $procs[[uint32]$_.Id] = $_.ProcessName }

$top = New-Object System.Collections.ArrayList
$cb = [P2+EnumProc]{ param($h,$l) [void]$top.Add($h); return $true }
[void][P2]::EnumWindows($cb,[IntPtr]::Zero)

$fg = [P2]::GetForegroundWindow()
Write-Output "FOREGROUND: 0x$('{0:X}' -f [int64]$fg)  cls=$([P2]::Cls($fg))  pid=$([P2]::Pid($fg)) ($($procs[[P2]::Pid($fg)]))  title='$([P2]::Txt($fg))'"
Write-Output ""
Write-Output "=========== ALL TOP-LEVEL WINDOWS IN WINWORD / OFFICETAB PROCESSES (no filtering) ==========="

$rows = @()
foreach ($h in $top) {
  $pid_ = [P2]::Pid($h)
  $pn = $procs[$pid_]
  if ($pn -notmatch 'WINWORD|OfficeTab|ExtendOffice') { continue }
  $r = New-Object P2+RECT; [void][P2]::GetWindowRect($h,[ref]$r)
  $style   = [int64][P2]::GetWindowLongPtr($h,-16)   # GWL_STYLE
  $exstyle = [int64][P2]::GetWindowLongPtr($h,-20)   # GWL_EXSTYLE
  $rows += [pscustomobject]@{
    HWND    = "0x{0:X}" -f [int64]$h
    Proc    = "$pn/$pid_"
    Class   = [P2]::Cls($h)
    Vis     = [P2]::IsWindowVisible($h)
    Icon    = [P2]::IsIconic($h)
    Cloak   = [P2]::Cloaked($h)
    Layered = [bool]($exstyle -band 0x80000)
    ToolWin = [bool]($exstyle -band 0x80)
    TopMost = [bool]($exstyle -band 0x8)
    Rect    = "$($r.L),$($r.T) $($r.R-$r.L)x$($r.B-$r.T)"
    Title   = [P2]::Txt($h)
  }
}
$rows | Format-Table -AutoSize

Write-Output ""
Write-Output "=========== CHILD TREE OF EACH VISIBLE, ON-SCREEN OpusApp ==========="
foreach ($row in ($rows | Where-Object { $_.Class -eq 'OpusApp' })) {
  $h = [IntPtr][Convert]::ToInt64($row.HWND.Substring(2),16)
  Write-Output "--- $($row.HWND) vis=$($row.Vis) cloak=$($row.Cloak) rect=$($row.Rect) '$($row.Title)' ---"
  $kids = New-Object System.Collections.ArrayList
  $cb2 = [P2+EnumProc]{ param($k,$l) [void]$kids.Add($k); return $true }
  [void][P2]::EnumChildWindows($h,$cb2,[IntPtr]::Zero)
  foreach ($k in $kids) {
    $kc = [P2]::Cls($k)
    if ($kc -match '_Ww|NetUI|Ribbon|OfficeTab|Tabs|Afx|ATL|Static|SysTabControl') {
      $kr = New-Object P2+RECT; [void][P2]::GetWindowRect($k,[ref]$kr)
      $par = [P2]::GetParent($k)
      "    0x{0,-9:X} {1,-26} par=0x{2,-9:X} vis={3,-5} {4},{5} {6}x{7} '{8}'" -f `
        [int64]$k, $kc, [int64]$par, [P2]::IsWindowVisible($k), $kr.L, $kr.T, ($kr.R-$kr.L), ($kr.B-$kr.T), [P2]::Txt($k)
    }
  }
}
