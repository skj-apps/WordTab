<#
.SYNOPSIS
  Drive and assert the scrolling tab row: what happens once there are more documents than fit.

.DESCRIPTION
  Until this slice, a row with more documents than fitted at the 70-logical-pixel minimum did this:
  the tabs kept their minimum width and ran off the end of the strip, and the new-document button,
  having nowhere left to go, was **clamped back on top of the last tab**. Hit-testing tries tabs
  first and drawing paints the + last, so the user saw a + and clicked a tab - and at 200% on a
  1280-pixel window, nine documents was enough for the + glyph's centre to land inside that tab's
  *close button*. Clicking the drawn + closed a document.

  What replaces it is a row that scrolls. The tabs live in a **track** that stops short of a fixed
  cluster - two scroll chevrons and the + - pinned to the right-hand end of the strip. Tabs are
  clipped to the track and hit-tested against the part of them inside it, so a tab and a button
  cannot occupy the same pixel by construction rather than by a clamp that gives up.

  **How this is proved without reading pixels.** Every assertion here is a click at a position
  `tools\WordLayout.cs` computed, followed by asking which document came forward. That is decisive
  in a way a photograph is not: the row's scroll position is state inside Word's process which the
  add-in does not publish, but it is *knowable* at the two ends - a row that has just overflowed is
  at zero, and a row scrolled harder than it can move is at its maximum - and everywhere else it is
  exactly zero, because a row that fits cannot scroll at all. So the script drives the row to a known
  position, computes where a given tab must therefore be, clicks there, and reads the title. If the
  mirror and the add-in disagree by so much as a tab width, the wrong document comes forward and the
  assertion fails loudly.

  **The window is narrowed rather than the documents multiplied.** Overflow needs
  `count * minimum > available`, and the width is the cheaper half of that product to change: six
  documents in a window sized for three tabs overflows exactly as hard as twenty documents in a
  maximised one, and takes a minute rather than five to set up. The target width is computed from the
  measured DPI, never hard-coded.

.PARAMETER KeepOpen
  Leave Word running afterwards, to look at it by hand.

.PARAMETER Screenshot
  Save photographs of the row at each scroll position.

.EXAMPLE
  pwsh -File tools\check-scroll.ps1
  pwsh -File tools\check-scroll.ps1 -KeepOpen -Screenshot
#>
[CmdletBinding()]
param(
    [switch]$KeepOpen,
    [switch]$Screenshot,
    [string]$ShotDir = $env:TEMP,
    [int]$Docs = 6,
    [int]$Visible = 3
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

trap {
    Write-Host ''
    Write-Host "UNHANDLED: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkRed
    break
}

$source = Get-Content -Raw -Path (Join-Path $PSScriptRoot 'WordLayout.cs')
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition $source -Language CSharp -ReferencedAssemblies @(
    'System.Runtime', 'System.Collections', 'System.Threading.Thread', 'netstandard'
)
[WordLayout]::MakeDpiAware() | Out-Null

# Confirmed input, one shared idea of what a Word dialog is, and the bounded waits. See the header of
# tools\WordTabHarness.ps1 for why these are not per-suite copies any more.
. (Join-Path $PSScriptRoot 'WordTabHarness.ps1')

# The wheel, which is the one piece of input WordLayout has no reason to carry: nothing else in this
# add-in responds to it.
Add-Type -Namespace WordTabCheck -Name Wheel -MemberDefinition @'
[DllImport("user32.dll")] public static extern void mouse_event(uint f, int dx, int dy, int data, System.UIntPtr extra);
'@

$SwitchKey = 'HKCU:\Software\WordTab'

$script:Failures = 0
$script:Checks   = 0

function Write-Step($text) { Write-Host "==> $text" -ForegroundColor Cyan }
function Write-Note($text) { Write-Host "    $text" -ForegroundColor DarkGray }
function Assert($condition, $text) {
    $script:Checks++
    if ($condition) { Write-Host "    PASS  $text" -ForegroundColor Green }
    else { $script:Failures++; Write-Host "    FAIL  $text" -ForegroundColor Red }
}

# Every one of these re-wraps in @() before asking for .Count: a PowerShell function that returns an
# empty array returns *nothing*, and .Count on nothing is a hard error under Set-StrictMode. The same
# trap costs a run on this project about once a slice.
function Get-WordPids { return Get-WordPidList }
function Get-WordPidCount { return Get-WordPidTally }
function Get-Frames { return Get-WordFrameList }
function Get-FrameCount { return Get-WordFrameTally }

function Get-Child($frame, $class) {
    foreach ($kid in [WordLayout]::Children($frame)) { if ($kid.Class -eq $class) { return $kid } }
    return $null
}

function Short($hwnd) { "0x{0:X}" -f [int64]$hwnd }
function Name($hwnd) {
    if (-not [WordLayout]::IsWindow2($hwnd)) { return '(gone)' }
    $t = [WordLayout]::TitleOf($hwnd) -replace ' - Word$', ''
    if ($t -eq '') { return (Short $hwnd) }
    return $t
}

# Set-WordForeground comes from WordTabHarness.ps1 now. SetForegroundWindow is refused from a process
# that is not already foreground, and a click on a tab then raises that document *inside* Word without
# Word coming forward - so a reading of the row comes back naming some other application.

function Get-TopStrip {
    $frames = @(Get-Frames)
    if ($frames.Count -eq 0) { throw 'No Word windows.' }
    $top = [WordLayout]::GetForeground()
    if (-not ($frames -contains $top)) { $top = $frames[0] }
    $strip = Get-Child $top 'WordTabStrip'
    if (-not $strip) { throw ("Window {0} has no WordTab strip." -f (Short $top)) }
    return $strip.Hwnd
}

# The row as the add-in must have laid it out, at a scroll position this script has just driven it
# to. Measured immediately before every use: the strip moves - 46 pixels between two clicks a second
# apart has been recorded on this rig - and a rectangle taken earlier aims at where it used to be.
function Get-Row($scroll, $scrollEnabled = $true) {
    $strip = Get-TopStrip
    $count = Get-FrameCount
    return [pscustomobject]@{
        Strip  = $strip
        Count  = $count
        Layout = [WordLayout]::Tabs($strip, $count, $scroll, $scrollEnabled)
    }
}

function Wait-For($predicate, $seconds = 20) {
    $deadline = (Get-Date).AddSeconds($seconds)
    while ((Get-Date) -lt $deadline) { if (& $predicate) { return $true }; Start-Sleep -Milliseconds 400 }
    return $false
}

# Set-LogMark, Get-LogSince and Get-LogCount come from WordTabHarness.ps1 now. The copy that used to
# live here silently read from offset 0 after the add-in rolled the log, so the negative check below
# - that the + did not close anything - could pass having read a file the evidence was deleted from.

function Save-StripShot($name) {
    if (-not $Screenshot) { return }
    New-Item -ItemType Directory -Path $ShotDir -Force | Out-Null
    $r = [WordLayout]::RectOf((Get-TopStrip))
    $bmp = New-Object System.Drawing.Bitmap(($r.Right - $r.Left), ($r.Bottom - $r.Top))
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($r.Left, $r.Top, 0, 0, $bmp.Size)
    $g.Dispose()
    $path = Join-Path $ShotDir "wordtab-scroll-$name.png"
    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    Write-Note $path
}

# ---- the universal probe ---------------------------------------------------------------------
#
# Click where slot `index` must be if the row is scrolled to `scroll`, and answer with whatever came
# forward. Everything in this suite is built out of this: it is the only way to ask "which document
# is at this position" from outside Word, and because it goes through the same arithmetic the add-in
# uses, an answer that is off by a tab means the two copies have drifted.
function Get-SlotDocument($index, $scroll, $scrollEnabled = $true) {
    Set-WordForeground | Out-Null
    $row = Get-Row $scroll $scrollEnabled
    if ($index -ge $row.Layout.Tabs.Length) { throw "Slot $index does not exist ($($row.Layout.Tabs.Length) tabs)." }

    $t = $row.Layout.Tabs[$index]
    $track = $row.Layout.Track

    # The visible part, because in a scrolled row the slot may be half under the edge of the track
    # and its centre may be off it entirely.
    $left  = [Math]::Max($t.Left, $track.Left)
    $right = [Math]::Min($t.Right, $track.Right)
    if ($right - $left -lt 8) { throw "Slot $index is not visible enough to click ($($right - $left) px)." }

    $x = [int](($left + $right) / 2)
    $y = [int](($t.Top + $t.Bottom) / 2)

    # Confirmed onto the strip, re-measuring on every attempt. This function IS the suite's oracle -
    # every scroll assertion is "click the computed slot, see which document came forward" - so a
    # click that landed on the document instead reports itself as the row being scrolled wrongly.
    # Windows are resized several times in this suite, so a rectangle is stale almost immediately.
    if (-not (Invoke-ConfirmedClick -What "probing slot $index" -Point {
                  $r = Get-Row $scroll $scrollEnabled
                  $tt = $r.Layout.Tabs[$index]
                  $l = [Math]::Max($tt.Left, $r.Layout.Track.Left)
                  $rt = [Math]::Min($tt.Right, $r.Layout.Track.Right)
                  [pscustomobject]@{ X = [int](($l + $rt) / 2); Y = [int](($tt.Top + $tt.Bottom) / 2) }
              })) {
        Write-Note "slot ${index}: the probe click never landed on the strip"
    }
    Start-Sleep -Milliseconds 800

    $now = [WordLayout]::GetForeground()
    if (-not (@(Get-Frames) -contains $now)) {
        Write-Note ("the foreground was `"{0}`" - taking it back and clicking again" -f [WordLayout]::TitleOf($now))
        Set-WordForeground | Out-Null
        Invoke-ConfirmedClick -What "re-probing slot $index" -Point {
            $r = Get-Row $scroll $scrollEnabled
            $tt = $r.Layout.Tabs[$index]
            $l = [Math]::Max($tt.Left, $r.Layout.Track.Left)
            $rt = [Math]::Min($tt.Right, $r.Layout.Track.Right)
            [pscustomobject]@{ X = [int](($l + $rt) / 2); Y = [int](($tt.Top + $tt.Bottom) / 2) }
        } | Out-Null
        Start-Sleep -Milliseconds 900
        $now = [WordLayout]::GetForeground()
    }
    return (Name $now)
}

# Drive the row to a known end with the chevrons. It asserts nothing - what it is for is putting the
# row somewhere this script can compute from.
#
# The click count is worked out rather than looped until something changes, because there is nothing
# to watch: the scroll position is not observable from here, which is the whole reason this suite
# drives the row to its ends. Two clicks more than the row is long is enough, and the extra ones land
# on a chevron that has gone dead and is deliberately not a hit target, so they do nothing.
function Move-RowTo($where) {
    Set-WordForeground | Out-Null
    $row = Get-Row 0
    if (-not $row.Layout.HasNav -or $row.Layout.Width -le 0) { return 0 }

    $steps = [int][Math]::Ceiling($row.Layout.MaxScroll / [double]$row.Layout.Width) + 2
    $clicks = 0
    for ($i = 0; $i -lt $steps; $i++) {
        $row = Get-Row 0
        if ($where -eq 'end') { $button = $row.Layout.Next } else { $button = $row.Layout.Prev }
        # Confirmed onto the strip - and the last two clicks deliberately land on a chevron that has
        # gone dead, which is still the strip, so the confirmation is right for those too.
        Invoke-ConfirmedClick -What "chevron click $($i + 1) towards the $where" -Point {
            $r = Get-Row 0
            [WordLayout]::Center($(if ($where -eq 'end') { $r.Layout.Next } else { $r.Layout.Prev }))
        } | Out-Null
        $clicks++
        Start-Sleep -Milliseconds 150
    }
    Start-Sleep -Milliseconds 400
    return $clicks
}

function Use-Wheel($notches) {
    # Negative wheel data is a scroll down, which this strip treats as a scroll right.
    $row = Get-Row 0
    $p = [WordLayout]::Center($row.Layout.Track)
    # The wheel goes to whatever is under the POINTER, not to the foreground window - so if this move
    # silently does not take, every notch below is delivered to some other window and the row simply
    # does not move, which reads exactly like a strip that ignores the wheel.
    if (-not (Set-Pointer $p.X $p.Y 'putting the pointer over the tab row')) {
        Write-Note 'the pointer would not go over the row - the wheel notches below will not reach the strip'
    }
    $overWhat = Get-ClassAt $p.X $p.Y
    if ($overWhat -ne 'WordTabStrip') { Write-Note "the pointer is over `"$overWhat`", not the strip - the wheel will go there" }
    Start-Sleep -Milliseconds 250
    for ($i = 0; $i -lt [Math]::Abs($notches); $i++) {
        [WordTabCheck.Wheel]::mouse_event(0x0800, 0, 0, $(if ($notches -lt 0) { 120 } else { -120 }), [UIntPtr]::Zero)
        Start-Sleep -Milliseconds 130
    }
    Start-Sleep -Milliseconds 500
}

# ---- a clean Word ------------------------------------------------------------------------------
#
# Gracefully, never Kill: a killed Word offers to recover these documents on the next launch, and
# this suite counts windows. And it starts from zero on purpose - a run that begins with another
# suite's leftovers measures something else, and has.

if ((Get-WordPidCount) -gt 0) {
    Write-Step "Closing $(Get-WordPidCount) Word process(es) already running"
    # Nothing in this suite used to look for a dialog, so "Word would not close" was the only
    # diagnosis it could give - whether the cause was a save prompt, a gallery, or Word wedged.
    $start = Close-AllWord
    if (-not $start.Closed) {
        throw ("Word would not close ({0}: {1}) - close it by hand, saving or discarding as you like, then re-run." -f
               $start.Reason, (Format-WordWindow $start.Dialog))
    }
    Start-Sleep -Seconds 2
}

$scratch = Join-Path $env:TEMP 'wordtab-check'
New-Item -ItemType Directory -Path $scratch -Force | Out-Null

function Open-Documents($n) {
    for ($i = 1; $i -le $n; $i++) {
        $path = Join-Path $scratch ("scroll-{0}.rtf" -f $i)
        "{\rtf1\ansi WordTab scrolling row, document $i.\par}" | Set-Content -Path $path -Encoding Ascii
        Start-Process -FilePath 'winword.exe' -ArgumentList "`"$path`""
        if (-not (Wait-For { (Get-FrameCount) -ge $i } 60)) { throw "Document $i never appeared." }
        Start-Sleep -Seconds 2
    }
    Start-Sleep -Seconds 3
}

Write-Step "Opening $Docs documents"
Open-Documents $Docs
Assert ((Get-FrameCount) -eq $Docs) "$Docs Word windows, one per document ($(Get-FrameCount))"

Set-WordForeground | Out-Null
Start-Sleep -Milliseconds 600

# ---- 1. the row that fits: nothing about this slice is visible -----------------------------------

Write-Step 'A row that fits - the layout this slice did not touch'
#
# The window is widened first, and that is worth stating rather than doing quietly: Word's own
# default window on this rig is 874 strip pixels, which at 200% is room for **five** tabs. Six
# documents in a freshly opened Word already overflow. So the defect this slice removes was not an
# edge case reachable only by opening twenty documents - it was three columns of a normal desktop
# away, and this section had to widen the window to find a row that fits at all.

$strip0 = Get-TopStrip
$dpi0   = [WordLayout]::Dpi($strip0)
$screen0 = [WordLayout]::ScreenRect()
$frame0 = [WordLayout]::GetForeground()
$rect0  = [WordLayout]::RectOf($frame0)
$strip0Rect = [WordLayout]::RectOf($strip0)
$nonClient0 = ($rect0.Right - $rect0.Left) - ($strip0Rect.Right - $strip0Rect.Left)
$fitWidth = $Docs * [WordLayout]::Scale(70, $dpi0) + [WordLayout]::Scale(6, $dpi0) * 2 +
            [WordLayout]::Scale(26, $dpi0) + [WordLayout]::Scale(4, $dpi0) + $nonClient0 + 40
$fitWidth = [Math]::Min($fitWidth, ($screen0.Right - $screen0.Left))
Write-Note ("Word opened at {0}px, which is {1} strip pixels - widening to {2} so {3} tabs fit" -f `
            ($rect0.Right - $rect0.Left), ($strip0Rect.Right - $strip0Rect.Left), $fitWidth, $Docs)
[WordLayout]::Resize($frame0, $fitWidth, ($rect0.Bottom - $rect0.Top))
Start-Sleep -Seconds 3
Set-WordForeground | Out-Null

$row = Get-Row 0
$L = $row.Layout
Write-Note ("strip {0} px, dpi {1}, tab width {2}" -f `
            ([WordLayout]::RectOf($row.Strip).Right - [WordLayout]::RectOf($row.Strip).Left), `
            [WordLayout]::Dpi($row.Strip), $L.Width)

Assert (-not $L.HasNav)        'no scroll buttons, because there is nothing to scroll'
Assert ($L.MaxScroll -eq 0)    "nothing to scroll ($($L.MaxScroll))"
Assert ($L.Scroll -eq 0)       'and the row is at the start'
Assert ($L.HasPlus)            'the + is there'
Assert ($L.Plus.Left -ge $L.Tabs[$Docs - 1].Right) 'and it sits after the last tab, not on top of it'

$overlap = $false
for ($i = 0; $i -lt $Docs; $i++) {
    if ($L.Tabs[$i].Right -gt $L.Plus.Left -and $L.Tabs[$i].Left -lt $L.Plus.Right) { $overlap = $true }
}
Assert (-not $overlap) 'no tab overlaps the + at all'
Assert ((Get-SlotDocument 0 0) -like 'scroll-1*') 'clicking the first tab brings up the first document'
Save-StripShot 'fits'

# ---- 2. narrow the window until the row overflows -------------------------------------------------

Write-Step "Narrowing the window until only about $Visible tabs fit"
#
# The width is computed, never hard-coded: the same suite has to overflow a row on a 100% rig and on
# this 200% one, and the two answers differ by a factor of two.

$strip = Get-TopStrip
$dpi   = [WordLayout]::Dpi($strip)
$minimum = [WordLayout]::Scale(70, $dpi)
$pad     = [WordLayout]::Scale(6, $dpi)
$gap     = [WordLayout]::Scale(4, $dpi)
$plusW   = [WordLayout]::Scale(26, $dpi)
$chevW   = [WordLayout]::Scale(20, $dpi)

$frame = [WordLayout]::GetForeground()
$stripRect  = [WordLayout]::RectOf($strip)
$frameRect  = [WordLayout]::RectOf($frame)
$nonClient  = ($frameRect.Right - $frameRect.Left) - ($stripRect.Right - $stripRect.Left)

$wantTrack  = $Visible * $minimum
$wantStrip  = $wantTrack + $pad * 2 + $gap * 2 + $chevW * 2 + $plusW
$wantWindow = $wantStrip + $nonClient
Write-Note ("minimum tab {0}px, so a {1}-tab track wants a {2}px window" -f $minimum, $Visible, $wantWindow)

[WordLayout]::Resize($frame, $wantWindow, ($frameRect.Bottom - $frameRect.Top))
Start-Sleep -Seconds 3
Set-WordForeground | Out-Null

$row = Get-Row 0
$L = $row.Layout
Assert ($L.HasNav)         'the row overflows, so the scroll buttons appear'
Assert ($L.MaxScroll -gt 0) "and there is somewhere to scroll to ($($L.MaxScroll) px)"
Assert ($L.Width -eq $minimum) "the tabs are at their minimum width and stop shrinking ($($L.Width) vs $minimum)"

$stripRect = [WordLayout]::RectOf($row.Strip)
Assert ($L.Plus.Right -eq ($stripRect.Right - $pad)) 'the + is pinned to the right-hand end of the strip'
Assert ($L.Track.Right -le $L.Prev.Left) 'the track stops before the scroll buttons'
Assert ($L.Prev.Right -le $L.Next.Left)  'and the two chevrons are side by side, left one first'
Assert ($L.Next.Right -le $L.Plus.Left)  'with the + beyond them both'

# ---- 3. THE DEFECT: the + is a +, at every scroll position ----------------------------------------

Write-Step 'The + never sits on a tab - the bug this slice is about'
#
# Stated as geometry first and then driven, because the geometry is the property and the click is the
# proof that the add-in agrees with it. Before this slice the click below closed a document.

foreach ($at in @(0, [int]($L.MaxScroll / 2), $L.MaxScroll)) {
    $probe = (Get-Row $at).Layout
    $hit = $false
    for ($i = 0; $i -lt $Docs; $i++) {
        $t = $probe.Tabs[$i]
        $left  = [Math]::Max($t.Left, $probe.Track.Left)
        $right = [Math]::Min($t.Right, $probe.Track.Right)
        if ($right -le $left) { continue }          # scrolled out of sight entirely
        if ($right -gt $probe.Plus.Left -and $left -lt $probe.Plus.Right) { $hit = $true }
        if ($right -gt $probe.Prev.Left -and $left -lt $probe.Next.Right -and $probe.HasNav) { $hit = $true }
    }
    Assert (-not $hit) "at scroll $at, no visible part of any tab reaches the buttons"
}

Write-Step 'Clicking the drawn + adds a document rather than closing one'
Set-LogMark
$before = Get-FrameCount
Move-RowTo 'end' | Out-Null            # the worst case: the row is as far right as it goes
Set-WordForeground | Out-Null
$row = Get-Row $L.MaxScroll
$p = [WordLayout]::Center($row.Layout.Plus)
Write-Note ("clicking the + at ({0},{1})" -f $p.X, $p.Y)
# This is the click the whole slice exists for: it used to land on a tab's close button. Confirming
# it lands on the strip at all is the floor - the two assertions below say what it hit once there.
Assert (Invoke-ConfirmedClick -What 'clicking the + on an overflowing row' `
                              -Point { [WordLayout]::Center((Get-Row $L.MaxScroll).Layout.Plus) }) `
       'the click on + landed on the strip'
Assert (Wait-For { (Get-FrameCount) -eq ($before + 1) } 45) `
       "a document appeared: $before -> $(Get-FrameCount)"
Assert ((Get-LogCount 'new-document button clicked') -ge 1) 'the add-in logged it as the + being clicked'
Assert ((Get-LogCount 'close clicked') -eq 0) 'and nothing was closed - which is what used to happen here'
Start-Sleep -Seconds 2
Save-StripShot 'overflowing'

$Docs = Get-FrameCount
Write-Note "the row is now $Docs tabs"

# ---- 4. the reveal: a new tab at the far end brings itself into view -------------------------------

Write-Step 'The tab that just appeared is the one on screen'
#
# The + appended a tab at the far end and Word made it active. A row that left it off screen would be
# a tab row that cannot say which document you are looking at. So the scroll must now be at its
# maximum, and the last slot must hold the new document.

$row = Get-Row 1000000
$max = $row.Layout.MaxScroll
$last = Get-SlotDocument ($Docs - 1) $max
Write-Note "the last slot holds: $last"
Assert ($last -notlike 'scroll-*') 'the last slot is the new blank document, which the row scrolled to reveal'

# ---- 5. the scroll buttons step one tab at a time -------------------------------------------------

Write-Step 'The scroll buttons'

$clicks = Move-RowTo 'start'
Write-Note "$clicks clicks of the left chevron to reach the start"
$row = Get-Row 0
Assert (-not $row.Layout.CanPrev) 'at the start there is nowhere further left to go'
Assert ($row.Layout.CanNext)      'and plenty of room to the right'
Assert ((Get-SlotDocument 0 0) -like 'scroll-1*') 'the first slot holds the first document, so the row really is at the start'

# One click of the right chevron moves the row by exactly one tab, which means slot 1 lands where
# slot 0 was. That is a stronger statement than "it moved": it is the step size, measured.
Set-WordForeground | Out-Null
$row = Get-Row 0
Invoke-ConfirmedClick -What 'one click of the right chevron' -Point { [WordLayout]::Center((Get-Row 0).Layout.Next) } | Out-Null
Start-Sleep -Milliseconds 600
$width = $row.Layout.Width
Assert ((Get-SlotDocument 1 $width) -like 'scroll-2*') "one click scrolled the row by exactly one tab ($width px)"

$clicks = Move-RowTo 'end'
Write-Note "$clicks more clicks to reach the end"
$row = Get-Row 1000000
$max = $row.Layout.MaxScroll
Assert ($row.Layout.CanPrev)       'at the end there is a long way back'
Assert (-not $row.Layout.CanNext)  'and nowhere further right'

# The button with nowhere to go is not a target at all, so clicking it must leave the row exactly
# where it is - proved by the row still answering the same way afterwards.
$atEnd = Get-SlotDocument ($Docs - 1) $max
Set-WordForeground | Out-Null
$row = Get-Row $max
# Confirmed to land, deliberately. This is the negative test: "the row did not move" is exactly what
# a click that never arrived also produces, so without this the assertion below can pass for the
# wrong reason. The dead chevron is still drawn on the strip, so WordTabStrip is the right expectation.
Assert (Invoke-ConfirmedClick -What 'clicking the dead right chevron' `
                              -Point { [WordLayout]::Center((Get-Row $max).Layout.Next) }) `
       'the click on the dead chevron really did land on the strip'
Start-Sleep -Milliseconds 500
Assert ((Get-SlotDocument ($Docs - 1) $max) -eq $atEnd) 'clicking the dead chevron does nothing at all'

# ---- 5b. holding a chevron keeps scrolling ---------------------------------------------------------
#
# One click is one tab, which is asserted above and is deliberate - a page per click overshoots the tab
# you were looking for every time. The cost of that is a long row taking a lot of clicks, and holding
# the button is the answer, the same way it is on a scrollbar arrow.
#
# Proved the way everything else in this suite is proved: the scroll position cannot be read from
# outside Word, but it is exactly knowable at both ends. A hold that started at the start and finishes
# with the row against the far end has scrolled by more than one tab, whatever the repeat rate is.
#
# The click semantics are unchanged and that is asserted too - a hold that also fired the release's
# step would overshoot every time by exactly one tab, which reads as the row being imprecise rather
# than as an extra step.

Write-Step 'Holding a scroll chevron down'

Move-RowTo 'start' | Out-Null
Assert ((Get-SlotDocument 0 0) -like 'scroll-1*') 'back at the start before the hold'

# First: an ordinary click must NOT look like a hold. The delay before repeating is what separates
# them, so this is the assertion that the delay exists at all.
Set-LogMark
Set-WordForeground | Out-Null
Invoke-ConfirmedClick -What 'a quick click of the right chevron' `
                      -Point { [WordLayout]::Center((Get-Row 0).Layout.Next) } | Out-Null
Start-Sleep -Milliseconds 700
Assert (@(Get-LogSince 'chevron held down').Count -eq 0) 'a click is not a hold - it scrolls once and stops'
Assert ((Get-SlotDocument 1 $width) -like 'scroll-2*') 'and it still moves the row by exactly one tab'

# Now the hold, from the start.
Move-RowTo 'start' | Out-Null
Assert ((Get-SlotDocument 0 0) -like 'scroll-1*') 'back at the start again for the hold'

Set-LogMark
Set-WordForeground | Out-Null
$row  = Get-Row 0
$hold = [WordLayout]::Center($row.Layout.Next)
$maxScroll = $row.Layout.MaxScroll

# Confirmed once, before the button goes down, and never again: this is one continuous gesture and a
# retry that re-pressed would be a different one.
$onWhat = Get-ClassAt $hold.X $hold.Y
Assert ($onWhat -eq 'WordTabStrip') "the hold starts on the strip, not on `"$onWhat`""

try {
    [WordLayout]::DragHold($hold.X, $hold.Y, $hold.X, $hold.Y, 1, 50)
    # Long enough to clear the 400ms delay and take several 90ms steps. The row is three tabs wide
    # here, so this is far more than enough to reach the end.
    Start-Sleep -Milliseconds 1500
}
finally {
    [WordLayout]::DragRelease($hold.X, $hold.Y)
}
Start-Sleep -Milliseconds 700

Assert (@(Get-LogSince 'chevron held down').Count -ge 1) 'the add-in reports the chevron being held'
Assert (@(Get-LogSince 'chevron released after a hold - no extra step').Count -eq 1) `
       'and the release adds no extra step on top of what the hold already did'

$row = Get-Row $maxScroll
Assert (-not $row.Layout.CanNext) 'holding it carried the row all the way to the far end'

# Against where CLICKING to the end put it, not against a guess at which document that is. The row
# order is whatever the sections above left it as, so naming a document here would be asserting the
# tab order rather than the scroll position - and the scroll position is what a hold is about.
$heldTo = Get-SlotDocument ($Docs - 1) $maxScroll
Write-Note "holding reached `"$heldTo`"; clicking to the end reached `"$atEnd`""
Assert ($heldTo -eq $atEnd) 'and it reached exactly where clicking all the way to the end reached'

# And back, which drives the other direction rather than assuming it is symmetrical.
Set-LogMark
Set-WordForeground | Out-Null
$row  = Get-Row $maxScroll
$hold = [WordLayout]::Center($row.Layout.Prev)
$onWhat = Get-ClassAt $hold.X $hold.Y
Assert ($onWhat -eq 'WordTabStrip') "the hold back starts on the strip, not on `"$onWhat`""
try {
    [WordLayout]::DragHold($hold.X, $hold.Y, $hold.X, $hold.Y, 1, 50)
    Start-Sleep -Milliseconds 1500
}
finally {
    [WordLayout]::DragRelease($hold.X, $hold.Y)
}
Start-Sleep -Milliseconds 700

Assert (@(Get-LogSince 'chevron held down').Count -ge 1) 'the left chevron holds too'
$row = Get-Row 0
Assert (-not $row.Layout.CanPrev) 'and it carried the row back to the start'
Assert ((Get-SlotDocument 0 0) -like 'scroll-1*') 'with the first document in the first slot again'

# ---- 6. the wheel ---------------------------------------------------------------------------------

Write-Step 'The wheel over the row'

Move-RowTo 'start' | Out-Null
Assert ((Get-SlotDocument 0 0) -like 'scroll-1*') 'back at the start'

Use-Wheel 2
$row = Get-Row (2 * $width)
Assert ((Get-SlotDocument 2 (2 * $width)) -like 'scroll-3*') 'two notches down scrolled the row two tabs to the right'

Use-Wheel -40
Assert ((Get-SlotDocument 0 0) -like 'scroll-1*') 'and wheeling back up returns it to the start, clamped'
Save-StripShot 'wheeled-back'

# ---- 7. the row keeps the active tab when the window changes shape --------------------------------

Write-Step 'Narrowing the window does not lose the active tab'
#
# The active document here is the first one - the click above made it so - and it sits in slot 0. The
# row is then scrolled to the far end, which puts that tab off screen, and the window is narrowed
# again. Narrowing changes how much of the row fits, so the add-in re-reveals the active tab and the
# row must come back to the start.
#
# Asserted on slot 1 rather than slot 0. Slot 0 holds the document that is already active, so
# clicking it would prove nothing whether the reveal happened or not; slot 1 holds a document whose
# identity is different at every other scroll position.

Move-RowTo 'end' | Out-Null
$frame = Set-WordForeground
$frameRect = [WordLayout]::RectOf($frame)
[WordLayout]::Resize($frame, ($frameRect.Right - $frameRect.Left - $minimum), ($frameRect.Bottom - $frameRect.Top))
Start-Sleep -Seconds 3
Set-WordForeground | Out-Null

Assert ((Get-SlotDocument 1 0) -like 'scroll-2*') 'the row came back to the active tab, so slot 1 is the second document'

Write-Step 'Widening it until everything fits again'
$frame = Set-WordForeground
$frameRect = [WordLayout]::RectOf($frame)
$screen = [WordLayout]::ScreenRect()
# Enough for every tab at its minimum plus the cluster, but never wider than the screen: a window
# sized past the desktop is one whose right-hand end - where every button in this slice lives -
# cannot be clicked.
$wide = $Docs * $minimum + $pad * 2 + $plusW + $gap + $nonClient + 40
$wide = [Math]::Min($wide, ($screen.Right - $screen.Left))
Write-Note "widening to $wide px"
[WordLayout]::Resize($frame, $wide, ($frameRect.Bottom - $frameRect.Top))
Start-Sleep -Seconds 3
Set-WordForeground | Out-Null

$row = Get-Row 0
Assert (-not $row.Layout.HasNav)     'the scroll buttons are gone again'
Assert ($row.Layout.MaxScroll -eq 0) 'and there is nothing left to scroll'
Assert ($row.Layout.Plus.Left -ge $row.Layout.Tabs[$Docs - 1].Right) 'the + is back after the last tab'
Assert ((Get-SlotDocument ($Docs - 1) 0) -notlike 'scroll-1*') 'and the last tab is where the layout says it is'
Save-StripShot 'fits-again'

# ---- 8. carrying a tab past the end of a scrolled row ---------------------------------------------

Write-Step 'Dragging a tab against the end of the row scrolls it'
#
# Without this, a drag in an overflowing row can only rearrange the tabs that happen to be on screen.
# The gesture here is a real one: press the first tab, carry it to the right-hand end of the track,
# and then *hold still*. The row has to move underneath it, which is the whole point - and it is why
# the auto-scroll runs on a timer rather than on mouse movement.

$frame = Set-WordForeground
$frameRect = [WordLayout]::RectOf($frame)
[WordLayout]::Resize($frame, $wantWindow, ($frameRect.Bottom - $frameRect.Top))
Start-Sleep -Seconds 3
Move-RowTo 'start' | Out-Null

$firstName = Get-SlotDocument 0 0
Write-Note "carrying '$firstName' from slot 0 to the far end"
Move-RowTo 'start' | Out-Null

Set-WordForeground | Out-Null
$row = Get-Row 0
$t = $row.Layout.Tabs[0]
$y = [int](($t.Top + $t.Bottom) / 2)
$grabX = [int]($t.Left + ($t.Right - $t.Left) * 0.4)
$edgeX = $row.Layout.Track.Right - 6

Set-LogMark
# Confirmed before the button goes down, and not again until it is up. If the grab misses, "the drag
# was recognised" fails for a reason that has nothing to do with edge-scrolling.
$onWhat = Get-ClassAt $grabX $y
Assert ($onWhat -eq 'WordTabStrip') "the drag starts on the strip (`"$onWhat`" is under the grab point)"
[WordLayout]::DragHold($grabX, $y, $edgeX, $y, 14, 55)
Start-Sleep -Milliseconds 2500          # holding still against the edge: the timer's job
[WordLayout]::DragRelease($edgeX, $y)
Start-Sleep -Seconds 1

Assert ((Get-LogCount 'drag started') -ge 1) 'the drag was recognised'
Assert ((Get-LogCount 'row scrolled') -ge 1) 'and holding at the edge scrolled the row'

$row = Get-Row 1000000
$max = $row.Layout.MaxScroll
Assert ((Get-SlotDocument ($Docs - 1) $max) -eq $firstName) `
       "the carried tab ended up last, which it could not reach without the row moving"

Save-StripShot 'after-drag'

# ---- 9. TabScroll=0: the escape hatch that shows every tab at once ---------------------------------

Write-Step 'HKCU\Software\WordTab\TabScroll=0 - every tab on screen, however narrow'
#
# Not "off gives back what it did before": what it did before was the defect. What it gives instead is
# the other honest answer - no minimum width at all, so the tabs divide the track between them
# however many there are. That is a worse row to read and a better one to be sure of, and it is the
# thing to reach for if the scrolling row ever strands a document off the end of a strip on a machine
# this has not run on.

Write-Note 'closing Word - the switch is read once, when the add-in starts'
$between = Close-AllWord
if (-not $between.Closed) {
    throw ("Word would not close before the TabScroll=0 run ({0}: {1})." -f
           $between.Reason, (Format-WordWindow $between.Dialog))
}
Start-Sleep -Seconds 2

$restore = $null
if (Test-Path $SwitchKey) {
    $existing = Get-ItemProperty -Path $SwitchKey -Name 'TabScroll' -ErrorAction SilentlyContinue
    if ($existing) { $restore = $existing.TabScroll }
} else {
    New-Item -Path $SwitchKey -Force | Out-Null
}
Set-ItemProperty -Path $SwitchKey -Name 'TabScroll' -Value 0 -Type DWord

# Whatever happens below, the switch goes back. A check script that leaves a machine configured
# differently from how it found it has changed the thing it was measuring.
try {
    $squeezeDocs = 5
    Write-Step "Opening $squeezeDocs documents with TabScroll=0"
    Open-Documents $squeezeDocs
    Assert ((Get-FrameCount) -eq $squeezeDocs) "$squeezeDocs windows ($(Get-FrameCount))"

    Set-WordForeground | Out-Null
    $frame = [WordLayout]::GetForeground()
    $frameRect = [WordLayout]::RectOf($frame)
    [WordLayout]::Resize($frame, $wantWindow, ($frameRect.Bottom - $frameRect.Top))
    Start-Sleep -Seconds 3
    Set-WordForeground | Out-Null

    $row = Get-Row 0 $false
    $L = $row.Layout
    Write-Note ("tab width {0}, minimum would have been {1}" -f $L.Width, $minimum)

    Assert (-not $L.HasNav)         'no scroll buttons - there is nothing to scroll'
    Assert ($L.MaxScroll -eq 0)     'and nothing to scroll to'
    Assert ($L.Width -lt $minimum)  "the tabs went below the minimum width rather than overflowing ($($L.Width))"

    $stripRect = [WordLayout]::RectOf($row.Strip)
    $allInside = $true
    for ($i = 0; $i -lt $squeezeDocs; $i++) {
        if ($L.Tabs[$i].Left -lt $L.Track.Left -or $L.Tabs[$i].Right -gt $L.Track.Right) { $allInside = $false }
    }
    Assert $allInside 'every tab is inside the track, so every document is on screen at once'
    Assert ((Get-SlotDocument ($squeezeDocs - 1) 0 $false) -like "scroll-$squeezeDocs*") `
           'and the last one is reachable with a single click, which is the whole point of this mode'
    Save-StripShot 'squeeze'
}
finally {
    if ($null -eq $restore) { Remove-ItemProperty -Path $SwitchKey -Name 'TabScroll' -ErrorAction SilentlyContinue }
    else { Set-ItemProperty -Path $SwitchKey -Name 'TabScroll' -Value $restore -Type DWord }
    Write-Note 'TabScroll put back'
}

# ---- done ------------------------------------------------------------------------------------------

if (-not $KeepOpen) {
    Write-Step 'Closing Word'
    $end = Close-AllWord
    if (-not $end.Closed) {
        Write-Note ("Word is still up ({0}): {1}" -f $end.Reason, (Format-WordWindow $end.Dialog))
        Write-Note 'Left running rather than killed. Answer it by hand before the next suite.'
    }
}

Write-Host ''
if ($script:Failures -eq 0) {
    Write-Host "PASS  $($script:Checks) checks, 0 failures." -ForegroundColor Green
} else {
    Write-Host "FAIL  $($script:Failures) of $($script:Checks) checks failed." -ForegroundColor Red
    exit 1
}
