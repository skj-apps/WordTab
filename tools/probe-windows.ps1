# Probe Word's window topology to settle: reparented _WwG (arch A) vs hidden OpusApp (arch B)
Add-Type @'
using System;
using System.Text;
using System.Runtime.InteropServices;
public class WinProbe {
  public delegate bool EnumProc(IntPtr hWnd, IntPtr lParam);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr p);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr h, EnumProc cb, IntPtr p);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowTextW(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr GetParent(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr GetAncestor(IntPtr h, uint flags);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
  public static string Cls(IntPtr h) { var sb = new StringBuilder(256); GetClassName(h, sb, 256); return sb.ToString(); }
  public static string Txt(IntPtr h) { var sb = new StringBuilder(512); GetWindowTextW(h, sb, 512); return sb.ToString(); }
  public static uint Pid(IntPtr h) { uint p; GetWindowThreadProcessId(h, out p); return p; }
}
'@ -ErrorAction Stop

function Get-TopLevel {
  $list = New-Object System.Collections.ArrayList
  $cb = [WinProbe+EnumProc]{ param($h, $l) [void]$list.Add($h); return $true }
  [void][WinProbe]::EnumWindows($cb, [IntPtr]::Zero)
  return $list
}
function Get-Children($parent) {
  $list = New-Object System.Collections.ArrayList
  $cb = [WinProbe+EnumProc]{ param($h, $l) [void]$list.Add($h); return $true }
  [void][WinProbe]::EnumChildWindows($parent, $cb, [IntPtr]::Zero)
  return $list
}

$procs = @{}
Get-Process -EA SilentlyContinue | ForEach-Object { $procs[[uint32]$_.Id] = $_.ProcessName }

$top = Get-TopLevel
Write-Output "=================== TOP-LEVEL OpusApp WINDOWS ==================="
$opus = @()
foreach ($h in $top) {
  if ([WinProbe]::Cls($h) -eq 'OpusApp') {
    $pid_ = [WinProbe]::Pid($h)
    $vis  = [WinProbe]::IsWindowVisible($h)
    $r = New-Object WinProbe+RECT; [void][WinProbe]::GetWindowRect($h, [ref]$r)
    $opus += [pscustomobject]@{ HWND = "0x{0:X}" -f [int64]$h; Visible = $vis; PID = $pid_; Proc = $procs[$pid_]; Rect = "$($r.L),$($r.T) $($r.R - $r.L)x$($r.B - $r.T)"; Title = [WinProbe]::Txt($h) }
  }
}
$opus | Format-Table -AutoSize
Write-Output "OpusApp total: $($opus.Count)  visible: $(($opus | Where-Object Visible).Count)  hidden: $(($opus | Where-Object {-not $_.Visible}).Count)"

Write-Output ""
Write-Output "=================== ALL _WwG DOCUMENT VIEWS ==================="
$wwg = @()
foreach ($h in $top) {
  foreach ($c in (Get-Children $h)) {
    if ([WinProbe]::Cls($c) -eq '_WwG') {
      $par  = [WinProbe]::GetParent($c)
      $root = [WinProbe]::GetAncestor($c, 2)   # GA_ROOT
      $pid_ = [WinProbe]::Pid($c)
      $wwg += [pscustomobject]@{
        WwG        = "0x{0:X}" -f [int64]$c
        Parent     = "0x{0:X}" -f [int64]$par
        ParentCls  = [WinProbe]::Cls($par)
        RootHWND   = "0x{0:X}" -f [int64]$root
        RootCls    = [WinProbe]::Cls($root)
        RootTitle  = [WinProbe]::Txt($root)
        Proc       = $procs[$pid_]
      }
    }
  }
}
$wwg | Format-Table -AutoSize
Write-Output "_WwG total: $($wwg.Count)   distinct roots: $(($wwg | Select-Object -Unique RootHWND).Count)"

Write-Output ""
Write-Output "=================== NON-WORD TOP-LEVEL WINDOWS OWNED BY OFFICE TAB / WORD PROCS ==================="
foreach ($h in $top) {
  $pid_ = [WinProbe]::Pid($h); $pn = $procs[$pid_]
  if ($pn -match 'OfficeTab|ExtendOffice|WINWORD') {
    $cls = [WinProbe]::Cls($h)
    if ($cls -ne 'OpusApp') {
      $r = New-Object WinProbe+RECT; [void][WinProbe]::GetWindowRect($h, [ref]$r)
      $w = $r.R - $r.L; $ht = $r.B - $r.T
      if ([WinProbe]::IsWindowVisible($h) -and $w -gt 0 -and $ht -gt 0) {
        "{0,-14} {1,-28} vis={2} {3},{4} {5}x{6}  '{7}'" -f $pn, $cls, [WinProbe]::IsWindowVisible($h), $r.L, $r.T, $w, $ht, [WinProbe]::Txt($h)
      }
    }
  }
}

Write-Output ""
Write-Output "=================== MODULES LOADED INTO WINWORD ==================="
Get-Process WINWORD -EA SilentlyContinue | ForEach-Object {
  "--- WINWORD PID $($_.Id) ---"
  $_.Modules | Where-Object { $_.FileName -match 'ExtendOffice|OfficeTab|TabsforOffice' } |
    Select-Object ModuleName, FileName | Format-Table -AutoSize
}
