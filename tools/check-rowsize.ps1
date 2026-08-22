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

    // WM_SYSCOMMAND / SC_MAXIMIZE - what the title bar's maximize button, a double-click on the
    // title bar and Win+Up all send. ShowWindow(SW_MAXIMIZE) maximizes the same window without it,
    // and that case must NOT be recorded, which is what step 8 checks either side of this.
    public static void UserMaximize(IntPtr hwnd) { SendMessage(hwnd, 0x0112, (IntPtr)0xF030, IntPtr.Zero); }
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

# The complaint this suite exists for is not something the add-in does - it is something WORD does,
# which the add-in then spreads across the row. Word remembers its own window size between runs, so
# whatever the last Word window was sized to is what the next one opens at. That is the whole
# mechanism, and priming it here is how a machine with one ordinary screen can be made to reproduce
# "Word opened it far too wide" on demand rather than waiting for a rig that has a 5120px monitor.
#
# RowSize is off while this runs, so the add-in leaves the priming window alone and does not record
# it either - what is being set up is Word's memory, not WordTab's.
function Prime-WordWidth($width, $height) {
    Close-AllWordAndWait
    Set-ItemProperty -Path $SettingsKey -Name RowSize -Value 0 -Type DWord
    Open-Document 9 | Out-Null
    Wait-WordReady 1 45 | Out-Null
    $f = Get-FirstFrame
    if ($f -eq [IntPtr]::Zero) { Write-Note 'no frame to prime Word width with'; return 0 }
    $work = [WordLayout]::WorkArea($f)
    [WordLayout]::Show($f, 9)                       # SW_RESTORE, in case Word came back maximized
    $w = [Math]::Min($width,  ($work.Right - $work.Left) - 80)
    $h = [Math]::Min($height, ($work.Bottom - $work.Top) - 80)
    [WordLayout]::MoveTo($f, $work.Left + 40, $work.Top + 40)
    [WordLayout]::Resize($f, $w, $h)
    Start-Sleep -Milliseconds 600
    Close-AllWordAndWait
    Remove-ItemProperty -Path $SettingsKey -Name RowSize -ErrorAction SilentlyContinue
    return $w
}

# TabDpi is what makes the page-width arithmetic reachable here. A page is 8.5in wide, so at the
# dev rig's real 192dpi three pages is 4896px - wider than any monitor here - and the default width
# could never be made to fire. Forcing the add-in's DPI shrinks a page to something this screen can
# hold three of, which is the only way this step can be driven at all.
#
# 72 is not a free choice and must not be lowered to suit a small screen: PageWidthPx refuses a
# forced dpi outside 72..480 and falls back to 96, so 72 is the smallest page that can be asked
# for. 612px, and three of them 1836px.
$ForcedDpi     = 72
$ForcedPage    = [int]((85 * $ForcedDpi) / 10)      # 612
$ForcedTrigger = [int](($ForcedPage * 30) / 10)     # 1836
$ForcedTarget  = [int](($ForcedPage * 14) / 10)     # 856

# And a screen has to be able to hold a Word window that wide, which is not a given: this rig
# measured 2856px across one morning and 1740px the same evening, and three checks went red on
# that alone - "Word was primed 1660px wide, which is over the 1836px that is three pages here"
# is a precondition failing, not the product. A check that cannot be run is a check that cannot
# fail, so it is skipped by name instead, with the reason. Prime-WordWidth can reach the work area
# less the 80px margin it keeps.
$WorkWide      = ([WordLayout]::WorkArea([IntPtr]::Zero)).Right - ([WordLayout]::WorkArea([IntPtr]::Zero)).Left
$CanForceWidth = (($WorkWide - 80) -gt $ForcedTrigger)

# ---- what was there before, so the machine is left as it was found -----------------------------

$origRowSize = try { (Get-ItemProperty -Path $SettingsKey -Name RowSize -ErrorAction Stop).RowSize } catch { $null }
$origTabDpi  = try { (Get-ItemProperty -Path $SettingsKey -Name TabDpi  -ErrorAction Stop).TabDpi  } catch { $null }
$origRow     = Get-StoredRow

try {

# ---- 1. nothing remembered: Word's rectangle stands --------------------------------------------

Write-Step 'Nothing remembered yet'
# Primed narrow on purpose. "Word's rectangle stands" is only the right answer while Word's
# rectangle is one a person would recognise as a Word window, and step 6 is the other half of this:
# left alone means left alone because there was nothing wrong with it, not because nothing looked.
Remove-ItemProperty -Path $SettingsKey -Name TabDpi -ErrorAction SilentlyContinue
Prime-WordWidth 1200 900 | Out-Null
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
Assert ((Get-LogCount 'no row size remembered and Word opened') -eq 0) `
       'and it did not narrow a window that was not too wide'

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

# ---- 6. nothing remembered, and Word opens it three pages wide ------------------------------------
#
# The complaint itself, reproduced by arithmetic rather than by hardware. Their rig is a 5120x2160
# screen at 150% where Word restores a 3840px frame; a page there is 8.5in x 144dpi = 1224px, so
# 3840px is three pages across. Here TabDpi forces 72dpi, which makes a page 612px and three pages
# 1836px - the same sum with numbers a screen this size can hold, when it can.

if (-not $CanForceWidth) {
    Write-Note "the work area is ${WorkWide}px across and this needs more than $($ForcedTrigger + 80)px to prime Word three pages wide - steps 6 and 7 skipped"
} else {
    Write-Step 'Nothing remembered, and Word opens it three pages wide'
    Close-AllWordAndWait
    Set-ItemProperty -Path $SettingsKey -Name TabDpi -Value $ForcedDpi -Type DWord
    $primed = Prime-WordWidth 2600 1200
    Clear-Row
    Remove-ItemProperty -Path $SettingsKey -Name RowSize -ErrorAction SilentlyContinue
    Set-LogMark
    Open-Document 5 | Out-Null
    if (-not (Wait-WordReady 1 45)) { Write-Note 'document 5 did not arrive with a strip within 45s' }
    Wait-Joined

    $narrowed = Get-LogLast 'no row size remembered and Word opened'
    Assert ($primed -ge $ForcedTrigger) `
           "Word was primed $($primed)px wide, which is over the ${ForcedTrigger}px that is three pages here"
    Assert ($null -ne $narrowed) 'the add-in said it was too wide and narrowed it'
    if ($null -ne $narrowed) { Write-Note $narrowed }

    $capped = [WordLayout]::RectOf((Get-FirstFrame))
    $cappedW = $capped.Right - $capped.Left
    Assert ([Math]::Abs($cappedW - $ForcedTarget) -le 4) `
           "the window came up ${cappedW}px wide, about one page ($ForcedTarget)"
    Assert (-not [WordLayout]::Maximized((Get-FirstFrame))) 'and not maximized'

    # The one that keeps the two apart. A default that wrote itself into the remembered slot would be
    # indistinguishable from a size the user chose, and would then be applied on every machine forever.
    Assert ($null -eq (Get-StoredRow)) 'and it recorded nothing - a default is not a size you chose'

    # ---- 7. a remembered size still beats the default --------------------------------------------------

    Write-Step 'A remembered size beats the default'
    Close-AllWordAndWait
    $wantW = $ForcedTarget * 2      # wider than the default, and still over the three-page threshold
    Set-StoredRow ($capped.Left) ($capped.Top) ($capped.Left + $wantW) ($capped.Top + 900) 0
    Set-LogMark
    Open-Document 6 | Out-Null
    if (-not (Wait-WordReady 1 45)) { Write-Note 'document 6 did not arrive with a strip within 45s' }
    Wait-Joined

    Assert ((Get-LogCount 'the row takes the size it was left at') -ge 1) 'the remembered rectangle was applied'
    Assert ((Get-LogCount 'no row size remembered and Word opened') -eq 0) 'and the default never ran'
    $kept = [RowProbe]::Placement((Get-FirstFrame)).rcNormalPosition
    Assert (($kept.Right - $kept.Left) -eq $wantW) `
           "the row is the ${wantW}px it was told to be, not the default $(Format-Rc $kept)"
}

# ---- 8. a size settled on WITHOUT a drag is remembered too ------------------------------------------
#
# WM_EXITSIZEMOVE ends the modal move/size loop, so it catches a drag and nothing else. Double-
# clicking the title bar and Win+Up never enter that loop, and until this was fixed a row maximized
# either way was never recorded - which is one of the ways a user can do everything they were asked
# to do and still get Word's own rectangle back on the next start.

Write-Step 'A size settled on without dragging'
Close-AllWordAndWait
Remove-ItemProperty -Path $SettingsKey -Name TabDpi -ErrorAction SilentlyContinue
Clear-Row
Remove-ItemProperty -Path $SettingsKey -Name RowSize -ErrorAction SilentlyContinue
# Marked BEFORE the document is opened, or Wait-Joined is satisfied by the PREVIOUS step's join
# line and returns at once - which is how this step first maximized a window half a second before
# it was a member of anything, and then reported that maximizing is not recorded. The suite has
# made this mistake before; see the note on Wait-Joined.
Set-LogMark
Open-Document 7 | Out-Null
if (-not (Wait-WordReady 1 45)) { Write-Note 'document 7 did not arrive with a strip within 45s' }
Wait-Joined
Assert ((Get-LogCount 'joined as the first window') -ge 1) 'the window joined the row first'

$frame8 = Get-FirstFrame
Assert ($frame8 -ne [IntPtr]::Zero) 'the frame is reachable from outside'

# First half, and it is the half that matters more. A resize NOBODY ASKED FOR must not be recorded.
# Keying this on WM_SIZE instead of the system command was measured to take two suites down in one
# battery - they ran against windows snapped to a rectangle an earlier suite had resized to - and on
# a wide screen it would let WORD's own startup rectangle install itself as the size the user chose,
# which is the one thing the remembered size exists to overrule.
Set-LogMark
[WordLayout]::Show($frame8, 9)          # SW_RESTORE, programmatic
[WordLayout]::Resize($frame8, 940, 820) # and a plain SetWindowPos - a resize, but not a person
Start-Sleep -Milliseconds 600
Assert ((Get-LogCount 'the row remembers its size') -eq 0) `
       'a resize the user did not ask for is not recorded'
Assert ($null -eq (Get-StoredRow)) 'and nothing was written'

# Second half: the same window, maximized the way a person maximizes it.
Set-LogMark
[RowProbe]::UserMaximize($frame8)       # WM_SYSCOMMAND / SC_MAXIMIZE - no modal loop, no drag
Start-Sleep -Milliseconds 600

$maxed = Get-StoredRow
Assert ((Get-LogCount 'the row remembers its size') -ge 1) `
       'maximizing without a drag IS recorded'
Assert ($null -ne $maxed) 'all five row values are present'
if ($null -ne $maxed) {
    Assert ($maxed['RowMaximized'] -eq 1) 'and it was recorded as maximized'
}
Assert ([WordLayout]::Maximized($frame8)) 'and the window really is maximized'

}
finally {
    # Put the machine back: these values are the user's, and a suite that leaves a row behind changes
    # what the next Word start does.
    Clear-Row
    Remove-ItemProperty -Path $SettingsKey -Name RowSize -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path $SettingsKey -Name TabDpi  -ErrorAction SilentlyContinue
    if ($null -ne $origRowSize) { Set-ItemProperty -Path $SettingsKey -Name RowSize -Value $origRowSize -Type DWord }
    if ($null -ne $origTabDpi)  { Set-ItemProperty -Path $SettingsKey -Name TabDpi  -Value $origTabDpi  -Type DWord }
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
