<#
.SYNOPSIS
  The ROW'S OWN SIZE: that it is remembered when the user settles on one, applied over whatever
  Word restores, and refused when it would land somewhere the user cannot reach.

.DESCRIPTION
  From the work rig's report: "still opening up 4 pages wide ... very annoying along w/ multiple
  pages opening". Their log said why. Every fresh Word start begins

      attach  hwnd=0x...  (existing)  ...  rect=(418,418 3840x1536)

  which is WORD restoring a near-full-width frame on a 5120x2160 screen, before the add-in has
  touched anything - and at that width Word's own layout puts about four pages side by side. The
  add-in then spreads it: `stack  hwnd=0x...  joined, snapped to 0x...` snaps every joining window
  to the first one, so the size of window #1 becomes the size of the whole row. Sizing the row fixed
  it until the next restart, when Word's rectangle came back.

  So the row remembers. What is checked here is the mechanism rather than the original complaint -
  the complaint needs a 5120px screen and a Word that restores wide, and this machine has neither,
  which is the same position tools\probe-view.ps1 is in. That is not a reason to check nothing: the
  fix does not depend on knowing WHY Word restores what it does, because it writes over the answer.
  What has to be true is that a size the user settles on is stored, comes back, and is refused when
  it is unreachable.

    1. **Nothing remembered: Word's rectangle stands.** The add-in must not invent a row size on a
       machine that has never been given one - it says nothing and leaves the window alone.

    2. **A size the user settles on is remembered.** Stored at WM_EXITSIZEMOVE, which is the moment
       a person stops dragging, and not from the intermediate rectangles a drag passes through.

    3. **It comes back on the next start**, applied to the window that defines the row.

    4. **An unreachable row is refused.** A remembered rectangle is only good while the screen it
       was on still exists; monitors get unplugged, and a row restored onto one that is gone is a
       row that cannot be reached. This is the rule MatchTo already keeps for a minimised master.

    5. **RowSize=0 turns the whole thing off**, back to Word's rectangle.

  The instrument is the add-in's own log plus GetWindowPlacement read from outside the process. The
  stored rectangle is compared against what GetWindowPlacement reports rather than against what this
  script asked for: rcNormalPosition is in workspace coordinates and SetWindowPos takes screen
  coordinates, and asserting across that boundary is how a check measures the wrong thing.

.PARAMETER KeepOpen
  Leave Word running afterwards.

.EXAMPLE
  pwsh -File tools\check-rowsize.ps1
#>
[CmdletBinding()]
param(
    [switch]$KeepOpen
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$source = Get-Content -Raw -Path (Join-Path $PSScriptRoot 'WordLayout.cs')
Add-Type -TypeDefinition $source -Language CSharp -ReferencedAssemblies @(
    'System.Runtime', 'System.Collections', 'System.Threading.Thread', 'netstandard'
)
[WordLayout]::MakeDpiAware() | Out-Null

# GetWindowPlacement and SendMessage are not in WordLayout.cs and are wanted only here, so they are
# declared here rather than widening a type eleven other suites depend on.
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public static class RowProbe
{
    [StructLayout(LayoutKind.Sequential)] public struct PT { public int X, Y; }
    [StructLayout(LayoutKind.Sequential)] public struct RC { public int Left, Top, Right, Bottom; }
    [StructLayout(LayoutKind.Sequential)]
    public struct PLACEMENT
    {
        public int length, flags, showCmd;
        public PT ptMinPosition, ptMaxPosition;
        public RC rcNormalPosition;
    }

    [DllImport("user32.dll")] static extern bool GetWindowPlacement(IntPtr hwnd, ref PLACEMENT p);
    [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr hwnd, uint msg, IntPtr w, IntPtr l);

    public static PLACEMENT Placement(IntPtr hwnd)
    {
        PLACEMENT p = new PLACEMENT();
        p.length = Marshal.SizeOf(typeof(PLACEMENT));
        GetWindowPlacement(hwnd, ref p);
        return p;
    }

    // WM_EXITSIZEMOVE. The message the add-in keys on, sent rather than posted so that it has been
    // handled by the time this returns and the registry can be read straight after.
    public static void SettleSize(IntPtr hwnd) { SendMessage(hwnd, 0x0232, IntPtr.Zero, IntPtr.Zero); }
}
'@ -Language CSharp -ReferencedAssemblies @('System.Runtime', 'netstandard')

. (Join-Path $PSScriptRoot 'WordTabHarness.ps1')

$script:Failures = 0
$script:Checks   = 0

function Write-Step($text) { Write-Host "==> $text" -ForegroundColor Cyan }
function Write-Note($text) { Write-Host "    $text" -ForegroundColor DarkGray }
function Assert($condition, $text) {
    $script:Checks++
    if ($condition) { Write-Host "    PASS  $text" -ForegroundColor Green }
    else { $script:Failures++; Write-Host "    FAIL  $text" -ForegroundColor Red }
}

$scratch = Join-Path $env:TEMP 'wordtab-check'
New-Item -ItemType Directory -Path $scratch -Force | Out-Null

$SettingsKey = 'HKCU:\Software\WordTab'
$RowNames    = @('RowLeft', 'RowTop', 'RowRight', 'RowBottom', 'RowMaximized')

function Open-Document($n) {
    $path = Join-Path $scratch "wordtab-rowsize-$n.rtf"
    "{\rtf1\ansi WordTab row-size check - document $n.\par}" | Set-Content -Path $path -Encoding Ascii
    Start-Process -FilePath 'winword.exe' -ArgumentList "`"$path`""
    return $path
}

function Clear-Row {
    foreach ($n in $RowNames) { Remove-ItemProperty -Path $SettingsKey -Name $n -ErrorAction SilentlyContinue }
}

function Get-StoredRow {
    $out = @{}
    foreach ($n in $RowNames) {
        try { $out[$n] = [int](Get-ItemProperty -Path $SettingsKey -Name $n -ErrorAction Stop).$n }
        catch { return $null }
    }
    return $out
}

function Set-StoredRow($left, $top, $right, $bottom, $maximized = 0) {
    Set-ItemProperty -Path $SettingsKey -Name RowLeft       -Value $left       -Type DWord
    Set-ItemProperty -Path $SettingsKey -Name RowTop        -Value $top        -Type DWord
    Set-ItemProperty -Path $SettingsKey -Name RowRight      -Value $right      -Type DWord
    Set-ItemProperty -Path $SettingsKey -Name RowBottom     -Value $bottom     -Type DWord
    Set-ItemProperty -Path $SettingsKey -Name RowMaximized  -Value $maximized  -Type DWord
}

function Get-FirstFrame {
    $frames = @(Get-WordFrameList)
    if ($frames.Count -eq 0) { return [IntPtr]::Zero }
    return [IntPtr]$frames[0]
}

# Wait-WordReady returns when the strip is bound. The stack JOIN is a separate, later event -
# measured about half a second later on this machine - and the join is what this suite is about.
# Asserting straight after Wait-WordReady is how the first run of this suite reported that no
# window had joined while the log plainly showed one half a second afterwards.
function Wait-Joined($seconds = 25) {
    Wait-Until { (Get-LogCount 'joined as the first window') -ge 1 } $seconds 200 | Out-Null
}

# Close-AllWord asks; it does not wait for the process to be gone. Opening the next document
# while the last WINWORD is still dying gets the new one handed off and exited ~250ms after it
# starts, which showed up as three steps whose window never joined anything.
function Close-AllWordAndWait {
    Close-AllWord | Out-Null
    Wait-WordGone 30 | Out-Null
}

function Format-Rc($rc) { return "($($rc.Left),$($rc.Top) $($rc.Right - $rc.Left)x$($rc.Bottom - $rc.Top))" }

# The value the add-in read is a DWORD; a negative coordinate comes back as its two's complement, so
# both sides are compared as 32-bit unsigned rather than one side being widened to a signed long.
function As-Dword([int]$v) { return [uint32]([uint32]::MaxValue -band $v) }

# ---- what was there before, so the machine is left as it was found -----------------------------

$origRowSize = try { (Get-ItemProperty -Path $SettingsKey -Name RowSize -ErrorAction Stop).RowSize } catch { $null }
$origRow     = Get-StoredRow

try {

# ---- 1. nothing remembered: Word's rectangle stands --------------------------------------------

Write-Step 'Nothing remembered yet'
Close-AllWordAndWait
Clear-Row
Remove-ItemProperty -Path $SettingsKey -Name RowSize -ErrorAction SilentlyContinue
Set-LogMark

Open-Document 1 | Out-Null
if (-not (Wait-WordReady 1 45)) { Write-Note 'document 1 did not arrive with a strip within 45s' }
Wait-Joined

Assert ((Get-LogCount 'joined as the first window') -ge 1) 'a window joined and defined the row'
Assert ((Get-LogCount 'row size=remembered') -ge 1) 'the row-size switch is on by default'
Assert ((Get-LogCount 'the row takes the size it was left at') -eq 0) `
       'nothing was remembered, so the add-in did not move the window'
Assert ($null -eq (Get-StoredRow)) 'and it invented no row rectangle of its own'

# ---- 2. a size the user settles on is remembered ------------------------------------------------

Write-Step 'A size the user settles on'
$frame = Get-FirstFrame
Assert ($frame -ne [IntPtr]::Zero) 'the frame is reachable from outside'

$work = [WordLayout]::WorkArea($frame)
$targetW = 900
$targetH = 900
$targetX = $work.Left + 120
$targetY = $work.Top + 90

[WordLayout]::MoveTo($frame, $targetX, $targetY)
[WordLayout]::Resize($frame, $targetW, $targetH)
Start-Sleep -Milliseconds 400

Set-LogMark
[RowProbe]::SettleSize($frame)
Start-Sleep -Milliseconds 400

# What the add-in should have stored is what GetWindowPlacement reports, not what was asked for.
$settled = [RowProbe]::Placement($frame).rcNormalPosition
$stored  = Get-StoredRow

Assert ((Get-LogCount 'the row remembers its size') -ge 1) 'the add-in recorded the size on WM_EXITSIZEMOVE'
Assert ($null -ne $stored) 'all five row values are present'
if ($null -ne $stored) {
    Assert ($stored['RowLeft']   -eq (As-Dword $settled.Left))   "stored left matches  $(Format-Rc $settled)"
    Assert ($stored['RowTop']    -eq (As-Dword $settled.Top))    'stored top matches'
    Assert ($stored['RowRight']  -eq (As-Dword $settled.Right))  'stored right matches'
    Assert ($stored['RowBottom'] -eq (As-Dword $settled.Bottom)) 'stored bottom matches'
    Assert ($stored['RowMaximized'] -eq 0) 'and it was not recorded as maximized'
}

# ---- 3. it comes back on the next start ---------------------------------------------------------

Write-Step 'It comes back on the next start'
Close-AllWordAndWait
Set-LogMark
Open-Document 2 | Out-Null
if (-not (Wait-WordReady 1 45)) { Write-Note 'document 2 did not arrive with a strip within 45s' }
Wait-Joined

Assert ((Get-LogCount 'the row takes the size it was left at') -ge 1) `
       'the add-in applied the remembered rectangle'

$after = [RowProbe]::Placement((Get-FirstFrame)).rcNormalPosition
Assert ($after.Left   -eq $settled.Left)   "the row came back at $(Format-Rc $after)"
Assert ($after.Top    -eq $settled.Top)    'top matches what was settled on'
Assert ($after.Right  -eq $settled.Right)  'right matches what was settled on'
Assert ($after.Bottom -eq $settled.Bottom) 'bottom matches what was settled on'

# ---- 4. an unreachable row is refused ------------------------------------------------------------

Write-Step 'A remembered row on a monitor that is gone'
Close-AllWordAndWait
# Far outside any monitor this machine has, standing in for the screen that was unplugged.
Set-StoredRow -30000 -30000 -28800 -28800
Set-LogMark
Open-Document 3 | Out-Null
if (-not (Wait-WordReady 1 45)) { Write-Note 'document 3 did not arrive with a strip within 45s' }
Wait-Joined

Assert ((Get-LogCount 'is on no monitor that exists now') -ge 1) 'the add-in refused the rectangle'
Assert ((Get-LogCount 'the row takes the size it was left at') -eq 0) 'and did not apply it'

$refused = [RowProbe]::Placement((Get-FirstFrame)).rcNormalPosition
Assert ($refused.Left -gt -30000) "the window is somewhere reachable $(Format-Rc $refused)"
Assert ($null -ne [WordLayout]::WorkArea((Get-FirstFrame))) 'and it is on a monitor'

# ---- 5. RowSize=0 turns it off --------------------------------------------------------------------

Write-Step 'RowSize=0'
Close-AllWordAndWait
Set-StoredRow $settled.Left $settled.Top $settled.Right $settled.Bottom 0
Set-ItemProperty -Path $SettingsKey -Name RowSize -Value 0 -Type DWord
Set-LogMark
Open-Document 4 | Out-Null
if (-not (Wait-WordReady 1 45)) { Write-Note 'document 4 did not arrive with a strip within 45s' }
Wait-Joined

Assert ((Get-LogCount 'row size=off') -ge 1) 'the add-in said the switch is off'
Assert ((Get-LogCount 'the row takes the size it was left at') -eq 0) 'and left the window to Word'
Assert ((Get-LogCount 'the row remembers its size') -eq 0) 'and recorded nothing either'

}
finally {
    # Put the machine back: these values are the user's, and a suite that leaves a row behind changes
    # what the next Word start does.
    Clear-Row
    Remove-ItemProperty -Path $SettingsKey -Name RowSize -ErrorAction SilentlyContinue
    if ($null -ne $origRowSize) { Set-ItemProperty -Path $SettingsKey -Name RowSize -Value $origRowSize -Type DWord }
    if ($null -ne $origRow) {
        Set-StoredRow $origRow['RowLeft'] $origRow['RowTop'] $origRow['RowRight'] $origRow['RowBottom'] $origRow['RowMaximized']
    }
}

# ---- summary --------------------------------------------------------------------------------------

Write-Host ''
if (-not $KeepOpen) { Close-AllWord | Out-Null }

if ($script:Failures -eq 0) {
    Write-Host "PASS  $($script:Checks) checks, 0 failures" -ForegroundColor Green
    exit 0
}
Write-Host "FAIL  $($script:Checks) checks, $($script:Failures) failure(s)" -ForegroundColor Red
exit 1
