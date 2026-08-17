<#
.SYNOPSIS
  Measure WordTab's tab reordering: dragging a tab moves it, and everything else about the row
  survives having been rearranged.

.DESCRIPTION
  The tab row used to be the order documents joined the stack, with nothing able to change it. This
  checks the order model and the gesture that drives it:

    - the row has an order that can be read from outside, and reading it twice gives the same answer
    - **a press that travels less than the drag threshold is a click, not a reorder.** This is the
      safety property of the whole slice: every switch between documents is a press that moves a
      pixel or two while the button is down, and if that rearranged the row, clicking a tab would be
      dangerous
    - a tab picked up and carried is *drawn* being carried. Isolated on purpose: the photograph at
      rest is taken with the same tab selected and the same tab hovered, and the tab is moved less
      than half a slot so the row does not reorder - so the only thing that can differ between the
      two photographs is the tab having been lifted out of its place
    - the row rearranges *while* the tab is being carried rather than on the drop, asserted
      mid-gesture with the button still held
    - dragging a tab two places to the right puts it two places to the right, and leaves every other
      tab in the order it was in. A reorder that moved the right tab and shuffled the others would
      pass the first half of that and still be wrong
    - the right button cancels a drag in progress and the row goes back exactly as it was - without
      leaving a context menu behind, ours or Word's
    - **a tab dragged clear of the row comes out of the stack when it is let go of, and one dragged
      below the strip but not clear of it still just reorders.** One press, two meanings, told apart
      by a distance measured off the strip's own height - and the pointer changes shape at the
      boundary, which is the only feedback available once the tab is over Word's document
    - a tab carried past the end of the row lands at the end rather than somewhere off it
    - the order is the same row in every window's strip, which is what makes the stack read as one
      window
    - a new document still arrives at the end after the row has been rearranged, and closing a tab
      out of the middle leaves the remaining order intact

  **How the order is read.** There is no API for "what does the strip say". A tab's identity is
  established the only honest way: click it and see which document comes forward. So one reading of
  the row is N clicks, and it is re-taken after every gesture rather than assumed. It doubles as the
  check that every window agrees on the row, because the click for tab i+1 is aimed at the strip
  belonging to whichever window tab i brought to the front.

  **Every click recomputes where it is aiming, immediately before it clicks** - the strip moves, and
  a rectangle measured three clicks ago points into the document. See check-tabs.ps1.

  Word is started on scratch documents written by this script and nothing is ever modified, so
  nothing prompts to save.

.PARAMETER Documents
  How many documents to open. Four is the minimum that can tell "moved two places" from "swapped
  with a neighbour" and still leave a tab beyond the move that must not have been touched. The fifth
  belongs to the tear-off section, which drags a tab out of the stack and then closes it - so every
  section below it runs on the four the rest of this suite was written against, and runs downstream
  of a window having left the row and come back.

.PARAMETER KeepOpen
  Leave Word running afterwards, to look at it by hand.

.PARAMETER Screenshot
  Save the at-rest and mid-drag photographs of the strip.

.EXAMPLE
  pwsh -File tools\check-reorder.ps1
  pwsh -File tools\check-reorder.ps1 -KeepOpen -Screenshot
#>
[CmdletBinding()]
param(
    [int]$Documents = 5,
    [switch]$KeepOpen,
    [switch]$Screenshot,
    [string]$ShotDir = $env:TEMP
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$source = Get-Content -Raw -Path (Join-Path $PSScriptRoot 'WordLayout.cs')
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition $source -Language CSharp -ReferencedAssemblies @(
    'System.Runtime', 'System.Collections', 'System.Threading.Thread', 'netstandard'
)
[WordLayout]::MakeDpiAware() | Out-Null

# Confirmed input, one shared idea of what a Word dialog is, and the bounded waits. See the header of
# tools\WordTabHarness.ps1 for why these are not per-suite copies any more.
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

# Virtual-key codes, the same one-line table check-menu and check-dot keep. A Windows constant rather
# than a decision about anything, so the copies cannot disagree about behaviour the way three
# Get-WordDialog implementations could - but it is still three copies, and it belongs in the harness
# the next time one of them needs a key the others do not have.
$VK = @{ W = 0x57 }

# Wait for the add-in to say something, and report how long it took.
#
# The alternative - sleep, then read - was in this suite and it is what a mid-gesture assertion must
# not be. A fixed wait long enough to be safe hides the difference between an add-in that reacts to
# the pointer instantly and one that is a second behind it, and a wait that is too short fails while
# naming the wrong thing: "the tab did not come back into the row" when what happened is that the
# strip had not processed the move yet. This asserts the event AND prints its latency, so a
# regression from instant to sluggish shows up in the notes instead of being absorbed by a sleep.
function Wait-Logged($needle, $count = 1, $seconds = 5) {
    $t0 = Get-Date
    $ok = Wait-Until { @(Get-LogSince $needle).Count -ge $count } $seconds 100
    return [pscustomobject]@{ Ok = $ok; Ms = [int]((Get-Date) - $t0).TotalMilliseconds }
}

function Get-Frames { return Get-WordFrameList }

# A PowerShell function returning an empty array returns *nothing*, and .Count on nothing throws -
# which bites exactly when the last document closes. Re-wrapping is what makes 0 come back as 0.
function Get-FrameCount { return Get-WordFrameTally }

function Get-WordPids { return Get-WordPidList }

function Get-Parts($frame) {
    $kids = [WordLayout]::Children($frame)
    [pscustomobject]@{
        Frame = $frame
        Rect  = [WordLayout]::RectOf($frame)
        Strip = @($kids | Where-Object { $_.Class -eq 'WordTabStrip' }) | Select-Object -First 1
        Wwf   = @($kids | Where-Object { $_.Class -eq '_WwF' }) | Select-Object -First 1
        Title = [WordLayout]::TitleOf($frame)
    }
}

# The strip the user is actually looking at: the one belonging to the window on top.
function Get-TopStrip {
    $frames = @(Get-Frames)
    if ($frames.Count -eq 0) { throw 'No Word windows.' }
    $top = [WordLayout]::GetForeground()
    if (-not ($frames -contains $top)) { $top = $frames[0] }
    $parts = Get-Parts $top
    if (-not $parts.Strip) { throw 'The foreground Word window has no WordTab strip.' }
    return $parts
}

# Where to aim, measured now - never held from an earlier call. `grip` is how far across the tab to
# take hold of it, as a fraction, so the arithmetic about where a carried tab lands stays readable.
function Get-TabSpot($index, $grip = 0.33) {
    $top = Get-TopStrip
    $count = (Get-FrameCount)
    $layout = [WordLayout]::Tabs($top.Strip.Hwnd, $count)
    if ($index -ge $layout.Tabs.Length) { throw "Tab $index does not exist ($($layout.Tabs.Length) tabs)." }

    $t = $layout.Tabs[$index]
    $strip = [WordLayout]::RectOf($top.Strip.Hwnd)
    $width = $t.Right - $t.Left
    [pscustomobject]@{
        X      = $t.Left + [int]($width * $grip)
        Y      = [int](($t.Top + $t.Bottom) / 2)
        Rect   = $t
        Width  = $width
        Strip  = $top.Strip.Hwnd
        Band   = $strip
        Count  = $count
    }
}

function Get-PlusSpot {
    $top = Get-TopStrip
    $layout = [WordLayout]::Tabs($top.Strip.Hwnd, (Get-FrameCount))
    $p = [WordLayout]::Center($layout.Plus)
    [pscustomobject]@{ X = $p.X; Y = $p.Y }
}

function Short($hwnd) { "0x{0:X}" -f [int64]$hwnd }
function Name($hwnd) {
    if (-not [WordLayout]::IsWindow2($hwnd)) { return '(gone)' }
    $t = [WordLayout]::TitleOf($hwnd) -replace ' - Word$', ''
    if ($t -eq '') { return (Short $hwnd) }
    return $t
}
function Format-Order($order) { return (@($order | ForEach-Object { Name $_ }) -join ' | ') }
function Format-Rect($r) { "({0},{1} {2}x{3})" -f $r.Left, $r.Top, ($r.Right - $r.Left), ($r.Bottom - $r.Top) }

# Make sure Word is the foreground application before a reading, and say so if it will not be.
#
# This is load-bearing, and it cost a run to find out. Switching tabs ends in SetForegroundWindow,
# and Windows refuses that call from a process that is not already the foreground application - so
# when something else owns the desktop, a click on a tab raises that document *inside* Word without
# making Word foreground. GetForegroundWindow then answers with the other application, and a reading
# of the row comes back naming Remote Desktop Connection three times, which is what happened.
#
# Focus() does the AttachThreadInput handshake, which is how a process that is not foreground asks
# for it. Only the check script does this - the add-in deliberately does not, because a click on a
# tab is user input into Word and Word can take the foreground on its own from there.
# Set-WordForeground comes from WordTabHarness.ps1 now. The copy that used to live here gave up
# after one attempt and returned IntPtr.Zero; the shared one retries, and as a last resort minimises
# the application sitting over Word - which is exactly the case this suite lost a drag to.

# Read the tab row, left to right, as a list of window handles.
function Get-Order {
    Set-WordForeground | Out-Null

    $count = (Get-FrameCount)
    $order = @()
    for ($i = 0; $i -lt $count; $i++) {
        # Confirmed onto the strip, and re-measured on every attempt. Reading the row is the whole
        # oracle of this suite - every reorder assertion is "the order before" against "the order
        # after" - so a click that missed does not report itself as a missed click, it reports
        # itself as a tab that did not move.
        if (-not (Invoke-ConfirmedClick -What "reading tab $i" -Point { Get-TabSpot $i })) {
            Write-Note "tab ${i}: the click never landed on the strip - the row reading below is unreliable"
        }
        Start-Sleep -Milliseconds 700

        $now = [WordLayout]::GetForeground()
        if (-not (@(Get-Frames) -contains $now)) {
            # The click switched the document without Word coming forward - see Set-WordForeground.
            # Take the foreground back and ask again rather than recording another application's
            # window as a tab.
            Write-Note ("tab {0}: the foreground was `"{1}`" - taking it back and re-reading" -f `
                        $i, [WordLayout]::TitleOf($now))
            Set-WordForeground | Out-Null
            Invoke-ConfirmedClick -What "re-reading tab $i" -Point { Get-TabSpot $i } | Out-Null
            Start-Sleep -Milliseconds 900
            $now = [WordLayout]::GetForeground()
        }
        $order += $now
    }
    return @($order)
}

function Test-SameOrder($a, $b) {
    if (@($a).Count -ne @($b).Count) { return $false }
    for ($i = 0; $i -lt @($a).Count; $i++) { if ($a[$i] -ne $b[$i]) { return $false } }
    return $true
}

# ---- the add-in's own log ------------------------------------------------------------------------
#
# Set-LogMark and Get-LogSince come from WordTabHarness.ps1 now. Read from a mark taken just before
# the gesture, never from the whole file: a search over everything would find the *previous* run's
# reorder and report a gesture that did nothing as having worked.
#
# The copy that used to live here read from offset 0 when the mark was past the end of the file -
# which is what a log that rolled looks like - so "nothing was logged since the mark" could pass
# having lost the evidence. It is now a stopped run with an explanation. See the log section of
# WordTabHarness.ps1.

# ---- photographs ---------------------------------------------------------------------------------

function Get-StripShot($strip) {
    $r = [WordLayout]::RectOf($strip)
    $bmp = New-Object System.Drawing.Bitmap(($r.Right - $r.Left), ($r.Bottom - $r.Top))
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($r.Left, $r.Top, 0, 0, $bmp.Size)
    $g.Dispose()
    return [pscustomobject]@{ Bitmap = $bmp; Origin = $r }
}

# How many sampled pixels differ between two shots, inside a rectangle given in screen coordinates.
function Measure-Changed($a, $b, $origin, $rect) {
    $changed = 0
    for ($y = $rect.Top; $y -lt $rect.Bottom; $y += 2) {
        for ($x = $rect.Left; $x -lt $rect.Right; $x += 2) {
            $px = $x - $origin.Left
            $py = $y - $origin.Top
            if ($px -lt 0 -or $py -lt 0 -or $px -ge $a.Width -or $py -ge $a.Height) { continue }
            if ($a.GetPixel($px, $py).ToArgb() -ne $b.GetPixel($px, $py).ToArgb()) { $changed++ }
        }
    }
    return $changed
}

function Test-OneRectangle($label) {
    $parts = @(Get-Frames | ForEach-Object { Get-Parts $_ })
    $rects = @($parts | ForEach-Object { Format-Rect $_.Rect } | Sort-Object -Unique)
    Assert ($rects.Count -eq 1) "$label - all $($parts.Count) windows still at one rectangle ($($rects.Count) distinct)"
}

function Wait-Frames($expected, $seconds = 20) {
    $deadline = (Get-Date).AddSeconds($seconds)
    while ((Get-Date) -lt $deadline) {
        if ((Get-FrameCount) -eq $expected) {
            Start-Sleep -Milliseconds 1200
            if ((Get-FrameCount) -eq $expected) { return $true }
        }
        Start-Sleep -Milliseconds 400
    }
    return $false
}

# ---- get Word up with N documents ----------------------------------------------------------------
#
# Any Word already running is closed first, and gracefully. Killing Word makes it offer to recover
# those documents on the next launch, which adds a window nobody asked for and a recovery pane over
# the document - and a check that starts in that state is measuring something other than what it says.

$running = @(Get-WordPids)
if ($running.Count -gt 0) {
    Write-Step "Closing $($running.Count) Word process(es) already running"
    $start = Close-AllWord
    if (-not $start.Closed) {
        throw ("Word would not close ({0}: {1}) - close it by hand, then re-run." -f
               $start.Reason, (Format-WordWindow $start.Dialog))
    }
    Start-Sleep -Seconds 2
}

$scratch = Join-Path $env:TEMP 'wordtab-check'
New-Item -ItemType Directory -Path $scratch -Force | Out-Null

# Names that read clearly in a printed order, so a failure says which document went where rather
# than printing four window handles.
$letters = @('Alpha', 'Bravo', 'Charlie', 'Delta', 'Echo', 'Foxtrot')
for ($i = 0; $i -lt $Documents; $i++) {
    $path = Join-Path $scratch ("wordtab-order-{0}.rtf" -f $letters[$i])
    "{\rtf1\ansi WordTab reorder check - document $($letters[$i]).\par}" | Set-Content -Path $path -Encoding Ascii
    Start-Process -FilePath 'winword.exe' -ArgumentList "`"$path`""
    # Waiting for Word to get into position, not for anything this suite claims about the product.
    if (-not (Wait-WordReady ($i + 1) 45)) { Write-Note "document $($letters[$i]) did not arrive with a strip on it within 45s" }
}

$deadline = (Get-Date).AddSeconds(45)
while ((Get-Date) -lt $deadline -and (Get-FrameCount) -lt $Documents) { Start-Sleep -Milliseconds 500 }

$frames = @(Get-Frames)
Write-Note "$($frames.Count) visible Word frame(s)"
if ($frames.Count -lt 3) { throw "Need at least 3 Word frames to check reordering; found $($frames.Count)." }

$settle = (Get-Date).AddSeconds(15)
while ((Get-Date) -lt $settle) {
    $parts = @(Get-Frames | ForEach-Object { Get-Parts $_ })
    $ready = @($parts | Where-Object { $_.Strip -and $_.Wwf -and $_.Strip.Bottom -eq $_.Wwf.Top })
    if ($ready.Count -eq $parts.Count) { break }
    Start-Sleep -Milliseconds 500
}

# ---- the row, before anything is dragged ----------------------------------------------------------

Write-Step 'Reading the tab row'
$baseline = Get-Order
Write-Note "order: $(Format-Order $baseline)"
Assert (@($baseline).Count -eq (Get-FrameCount)) "every window has a place in the row ($(@($baseline).Count) of $(Get-FrameCount))"
Assert (@($baseline | Sort-Object -Unique).Count -eq @($baseline).Count) 'each place in the row is a different window'

$again = Get-Order
Assert (Test-SameOrder $baseline $again) 'reading the row twice gives the same order - and every window''s strip agrees'

# ---- a press that does not travel far enough is a click, not a drag ---------------------------------
#
# The threshold is 4 logical pixels, so 2 physical pixels is under it at every DPI this can run at.

Write-Step 'Pressing a tab and moving two pixels'
Set-LogMark
Set-WordForeground | Out-Null
$spot = Get-TabSpot 0
# Confirmed BEFORE the button goes down, never during. This is one continuous button-down gesture and
# a retry that re-presses would break it. Without the check, a press that landed off the strip also
# produces "no drag started" and "nothing moved" - both assertions pass while nothing was measured.
$onWhat = Get-ClassAt $spot.X $spot.Y
Assert ($onWhat -eq 'WordTabStrip') "the press starts on the strip, not on `"$onWhat`""
[WordLayout]::DragTo($spot.X, $spot.Y, ($spot.X + 2), $spot.Y, 2, 150)
Start-Sleep -Milliseconds 800

Assert (@(Get-LogSince 'drag started').Count -eq 0) 'no drag started - it was a click'
Assert (@(Get-LogSince 'tab moved').Count -eq 0) 'nothing moved in the row'
$after = Get-Order
Assert (Test-SameOrder $baseline $after) "the order is untouched: $(Format-Order $after)"

# ---- picking a tab up ------------------------------------------------------------------------------
#
# One gesture, asserted in two stages. The first stage moves the tab a third of a slot: past the drag
# threshold, so it is being carried, but nowhere near half way past its neighbour, so the row must
# not reorder. That is what isolates the drawing: the photograph before it was taken with this same
# tab selected and this same tab hovered, so if the two differ, the difference is the tab having been
# lifted out of its place and nothing else.

Write-Step 'Carrying a tab a third of a slot'
Set-WordForeground | Out-Null
Set-LogMark
$mover = $baseline[0]

# Select it and hover it, so both photographs have the same tab selected and the same one hot. Both
# confirmed: if the hover silently does not take, the "resting" and "lifted" photographs are of the
# same pointer position and the drawing difference being measured is not the one named.
$spot = Get-TabSpot 0
Invoke-ConfirmedClick -What 'selecting the tab to be carried' -Point { Get-TabSpot 0 } | Out-Null
Start-Sleep -Milliseconds 800
$spot = Get-TabSpot 0
Assert (Set-Pointer $spot.X $spot.Y 'hovering the tab to be carried') 'the pointer could be put on the tab about to be carried'
Start-Sleep -Milliseconds 900

$top = Get-TopStrip
$rest = Get-StripShot $top.Strip.Hwnd
$lifted = $null
$carriedTo = 0

# Somewhere well beyond the reach of a third-of-a-slot shift: whatever changes there is not this tab.
$far = (Get-TabSpot ((Get-FrameCount) - 1)).Rect

try {
    $step1 = $spot.X + [int]($spot.Width / 3)
    # Confirmed here, and nowhere after: everything from DragHold to DragRelease is one gesture with
    # the button down, and re-aiming inside it would produce a different gesture from the one named.
    $onWhat = Get-ClassAt $spot.X $spot.Y
    Assert ($onWhat -eq 'WordTabStrip') "the pick-up starts on the strip, not on `"$onWhat`""
    [WordLayout]::DragHold($spot.X, $spot.Y, $step1, $spot.Y, 8, 70)
    $lifted = Get-StripShot $top.Strip.Hwnd

    Assert (@(Get-LogSince 'drag started').Count -ge 1) 'the add-in reports a drag in progress'
    Assert (@(Get-LogSince 'tab moved').Count -eq 0) 'a third of a slot is not far enough to reorder anything'

    # Stage two: on to the third slot, still without letting go. The row must rearrange now, while
    # the button is down - a design that waited for the drop would log nothing here.
    $target = Get-TabSpot 2
    $carriedTo = $target.X
    [WordLayout]::DragMoveTo($step1, $spot.Y, $target.X, $target.Y, 12, 60)

    $moves = @(Get-LogSince 'tab moved')
    Assert ($moves.Count -ge 1) "the row rearranged while the tab was being carried ($($moves.Count) move(s))"
    foreach ($line in $moves) { Write-Note $line.Trim() }
}
finally {
    if ($carriedTo -eq 0) { $carriedTo = $spot.X }
    [WordLayout]::DragRelease($carriedTo, $spot.Y)
}
Start-Sleep -Milliseconds 900

# If the strip moved between the two photographs the comparison means nothing, so say that rather
# than reporting a drawing failure that is really a relayout.
$still = ($rest.Origin.Left -eq $lifted.Origin.Left) -and ($rest.Origin.Top -eq $lifted.Origin.Top)
Assert $still 'the strip held still between the two photographs'

$onTab   = Measure-Changed $rest.Bitmap $lifted.Bitmap $rest.Origin $spot.Rect
$farAway = Measure-Changed $rest.Bitmap $lifted.Bitmap $rest.Origin $far
Write-Note "pixels changed where the tab was: $onTab   on the last tab: $farAway"
Assert ($onTab -gt 0) "the tab is visibly drawn out of its place while it is carried ($onTab pixels)"
Assert ($farAway -eq 0) "and nothing else in the row moved with it ($farAway pixels)"

if ($Screenshot) {
    New-Item -ItemType Directory -Path $ShotDir -Force | Out-Null
    $rest.Bitmap.Save((Join-Path $ShotDir 'wordtab-strip-rest.png'), [System.Drawing.Imaging.ImageFormat]::Png)
    $lifted.Bitmap.Save((Join-Path $ShotDir 'wordtab-strip-drag.png'), [System.Drawing.Imaging.ImageFormat]::Png)
    Write-Note (Join-Path $ShotDir 'wordtab-strip-rest.png')
    Write-Note (Join-Path $ShotDir 'wordtab-strip-drag.png')
}
$rest.Bitmap.Dispose(); $lifted.Bitmap.Dispose()

Assert (@(Get-LogSince 'drag ended').Count -ge 1) 'the drop was recorded'

# ---- where it landed --------------------------------------------------------------------------------

Write-Step 'Where the tab landed'
$moved = Get-Order
Write-Note "order: $(Format-Order $moved)"
Assert ($moved[2] -eq $mover) "`"$(Name $mover)`" is now the third tab"

# The rest of the row keeps its order. A reorder that moved the right tab and shuffled the others
# would pass the assertion above and still be wrong.
$others = @($baseline | Where-Object { $_ -ne $mover })
$expected = @($others[0], $others[1], $mover) + @($others | Select-Object -Skip 2)
Assert (Test-SameOrder $moved $expected) "every other tab kept its order: $(Format-Order $expected)"

Test-OneRectangle 'After a reorder'

# ---- the right button cancels a drag in progress --------------------------------------------------

Write-Step 'Cancelling a drag with the right button'
$before = Get-Order
Set-LogMark

Set-WordForeground | Out-Null
$spot = Get-TabSpot 0
$target = Get-TabSpot 2
try {
    $onWhat = Get-ClassAt $spot.X $spot.Y
    Assert ($onWhat -eq 'WordTabStrip') "the cancelled drag starts on the strip, not on `"$onWhat`""
    [WordLayout]::DragHold($spot.X, $spot.Y, $target.X, $target.Y, 12, 60)
    Assert (@(Get-LogSince 'tab moved').Count -ge 1) 'the row had rearranged before the cancel'
    [WordLayout]::RightTap($target.X, $target.Y)
}
finally {
    [WordLayout]::DragRelease($target.X, $target.Y)
}
Start-Sleep -Seconds 1

Assert (@(Get-LogSince 'drag cancelled').Count -ge 1) 'the add-in reports the drag was cancelled'

# No menu, ours or Word's. The right button during a drag means cancel, and a cancel that also opened
# a context menu would be two answers to one gesture.
Assert ((Get-WordMenu) -eq [IntPtr]::Zero) 'no context menu was left on screen'

$cancelled = Get-Order
Write-Note "order: $(Format-Order $cancelled)"
Assert (Test-SameOrder $before $cancelled) 'the row went back exactly as it was'

# ---- a drag that leaves the strip but not the row ---------------------------------------------------
#
# The control for everything below. Dragging a tab out of the row is a different gesture from dragging
# it along the row, and the only thing separating them is a distance - so before asserting that the
# distance works, assert that it is a distance at all rather than "the pointer left the strip".
#
# The threshold is measured off the strip rather than copied from the source. TEAROFF_LOGICAL_SLOP is
# defined as STRIP_LOGICAL_H - one row clear of the row - so the strip's own height on screen *is* the
# threshold in physical pixels, at whatever DPI this is running at, and a number mirrored from a
# #define could go stale without the suite noticing.

Write-Step 'Carrying a tab below the strip but inside the threshold'
$before = Get-Order
Set-LogMark
Set-WordForeground | Out-Null

$spot   = Get-TabSpot 0
$stripH = $spot.Band.Bottom - $spot.Band.Top
$mover  = $before[0]
Write-Note "the strip is ${stripH}px tall, so a tab has to go ${stripH}px clear of it to come out of the row"

$justBelow = $spot.Band.Bottom + [int]($stripH * 0.4)
$slot2     = Get-TabSpot 2
$onWhat = Get-ClassAt $spot.X $spot.Y
Assert ($onWhat -eq 'WordTabStrip') "the drag starts on the strip, not on `"$onWhat`""
[WordLayout]::DragTo($spot.X, $spot.Y, $slot2.X, $justBelow, 14, 60)
Start-Sleep -Milliseconds 900

Assert (@(Get-LogSince 'drag left the row').Count -eq 0) "0.4 of a row below the strip is still in the row"
Assert (@(Get-LogSince 'torn off').Count -eq 0) 'nothing came out of the stack'
$below = Get-Order
Write-Note "order: $(Format-Order $below)"
Assert ($below[2] -eq $mover) "`"$(Name $mover)`" reordered as usual, from below the strip"
Test-OneRectangle 'After a drag below the strip'

# ---- dragging a tab out of the row ------------------------------------------------------------------
#
# The gesture half of tear-off. What happens to the window afterwards - where it goes, what stops the
# janitor putting it straight back, when it stops being torn off - belongs to the stack, was built
# first and is driven by check-menu's Move to New Window section. This one is about the *input*: that
# one press can mean two things, that which one it means can be told apart before letting go, and that
# the two never happen at once.
#
# It is one continuous button-down gesture on purpose, and it goes out, back in, and out again. The
# route is the assertion: each crossing has to change what the drop would do, and the way back in is
# what drives the hysteresis - the return threshold is half the outward one, so a gesture that came
# back in at 0.3 of a row and out again at 1.6 has crossed both.
#
# The confirmation of where the pointer is happens ONCE, before the button goes down. Everything after
# it is inside the gesture and re-aiming there would be a different gesture from the one named.

Write-Step 'Dragging a tab out of the row'
$before = Get-Order
Write-Note "order: $(Format-Order $before)"
$victim     = $before[0]
$stackCount = Get-FrameCount
$stackRect  = [WordLayout]::RectOf($victim)
Write-Note ("tearing off `"{0}`" (0x{1:X}) out of {2} windows at {3}" -f `
            (Name $victim), [int64]$victim, $stackCount, (Format-Rect $stackRect))

Set-WordForeground | Out-Null
Set-LogMark
$spot      = Get-TabSpot 0
$slot2     = Get-TabSpot 2
$justBelow = $spot.Band.Bottom + [int]($stripH * 0.3)   # inside the return threshold
$wayBelow  = $spot.Band.Bottom + [int]($stripH * 1.6)   # past the outward one

$cursorInRow = [IntPtr]::Zero
$cursorTorn  = [IntPtr]::Zero
$cursorBack  = [IntPtr]::Zero
$dropAt      = $spot.X

try {
    $onWhat = Get-ClassAt $spot.X $spot.Y
    Assert ($onWhat -eq 'WordTabStrip') "the tear-off drag starts on the strip, not on `"$onWhat`""

    # In the row first, and moved far enough to actually rearrange it. That matters for the assertion
    # below: "the row went back" is only a claim if the row had gone somewhere.
    [WordLayout]::DragHold($spot.X, $spot.Y, $slot2.X, $spot.Y, 10, 70)
    $dropAt = $slot2.X
    Assert (@(Get-LogSince 'drag started').Count -ge 1) 'the add-in reports a drag in progress'
    # A slot at a time, not one jump: the row swaps as the carried tab's centre crosses each
    # neighbour, so this is "0 -> 1" and then "1 -> 2". The count is the claim; the return below is
    # the one worth naming exactly.
    $carriedMoves = @(Get-LogSince 'tab moved')
    foreach ($line in $carriedMoves) { Write-Note $line.Trim() }
    Assert ($carriedMoves.Count -ge 1) "and the row rearranged while it was being carried ($($carriedMoves.Count) move(s))"
    $cursorInRow = [WordLayout]::CursorShape()

    # Straight down and out. No horizontal component at all, which is the case a threshold measured
    # across the row could never see.
    [WordLayout]::DragMoveTo($slot2.X, $spot.Y, $slot2.X, $wayBelow, 10, 60)
    $out = Wait-Logged 'drag left the row'
    Write-Note "the add-in noticed the tab leave the row $($out.Ms)ms after the pointer got there"
    $cursorTorn = [WordLayout]::CursorShape()

    Assert $out.Ok 'a row-height clear of the strip takes the tab out of the row'
    Assert (@(Get-LogSince 'tab moved 2 -> 0').Count -ge 1) 'and the row lets go of it - the order it was picked up from comes back'

    # The pointer is the only feedback available out here: the tab is over Word's document, where
    # nothing of ours may paint. Asserted as a CHANGE first - that is the claim, and it cannot pass by
    # accident - and then named, because "some other cursor" would satisfy the change on its own.
    Write-Note ("cursor in the row 0x{0:X}, out of it 0x{1:X}, IDC_SIZEALL is 0x{2:X}" -f `
                [int64]$cursorInRow, [int64]$cursorTorn, [int64][WordLayout]::SystemCursor(32646))
    Assert ($cursorTorn -ne $cursorInRow) 'the pointer changes shape when the tab comes out of the row'
    Assert ($cursorTorn -eq [WordLayout]::SystemCursor(32646)) 'and it is IDC_SIZEALL - Windows'' own "you are relocating this"'

    # Back into the row, at less than the outward threshold. If the two thresholds were the same
    # number this would still be out, and the gesture would be back to flickering on the boundary.
    #
    # The pointer is read back after the move, and that is not belt-and-braces: a SendInput mouse move
    # can silently not take, and "the add-in did not notice the pointer come back" and "the pointer
    # never came back" are the same three failing assertions. This one was caught by running the same
    # script twice - it did the whole round trip the first time and never re-entered the row the
    # second - so the readback is what makes the difference reportable rather than mysterious.
    [WordLayout]::DragMoveTo($slot2.X, $wayBelow, $slot2.X, $justBelow, 8, 60)
    $at = [WordLayout]::Cursor()
    Write-Note ("aimed the pointer back to y={0} ({1}px below the strip); it is at y={2}" -f `
                $justBelow, ($justBelow - $spot.Band.Bottom), $at.Y)
    Assert ([Math]::Abs($at.Y - $justBelow) -le 2) 'the pointer really came back to within a third of a row of the strip'
    $back = Wait-Logged 'drag back in the row'
    Write-Note "the add-in noticed it come back $($back.Ms)ms later"
    $cursorBack = [WordLayout]::CursorShape()
    Assert $back.Ok 'coming back inside 0.3 of a row picks the tab up again'
    Assert ($cursorBack -eq [WordLayout]::SystemCursor(32512)) 'and the pointer is an ordinary arrow again'

    # And out again, which is where it is let go of.
    [WordLayout]::DragMoveTo($slot2.X, $justBelow, $slot2.X, $wayBelow, 8, 60)
    $at = [WordLayout]::Cursor()
    Write-Note ("aimed the pointer back out to y={0} ({1}px below the strip); it is at y={2}" -f `
                $wayBelow, ($wayBelow - $spot.Band.Bottom), $at.Y)
    Assert ([Math]::Abs($at.Y - $wayBelow) -le 2) 'and it went back out past the threshold'
    $out2 = Wait-Logged 'drag left the row' 2
    Write-Note "and leave again $($out2.Ms)ms later"
    Assert $out2.Ok 'and it can leave the row a second time'
    $dropAt = $slot2.X
}
finally {
    [WordLayout]::DragRelease($dropAt, $wayBelow)
}

Wait-Until { ([WordLayout]::RectOf($victim)).Left -ne $stackRect.Left } 8 | Out-Null
Start-Sleep -Milliseconds 1200

Assert (@(Get-LogSince 'drag dropped out of the row').Count -eq 1) 'letting go outside the row is recorded as a tear-off, once'
$said = @(Get-LogSince 'torn off, now its own window')
foreach ($line in $said) { Write-Note $line.Trim() }
Assert ($said.Count -eq 1) "the add-in took exactly one window out of the stack ($($said.Count))"

# Where everything ended up. Deliberately thinner than check-menu's section on the same mechanism -
# this is here to prove the drop reached it, not to re-check what it does.
$torn = @(Get-Frames | Where-Object { ([WordLayout]::RectOf($_)).Left -ne $stackRect.Left -or
                                      ([WordLayout]::RectOf($_)).Top  -ne $stackRect.Top })
foreach ($f in @(Get-Frames)) {
    Write-Note ("0x{0:X}  {1}  inAltTab={2}{3}" -f [int64]$f, (Format-Rect ([WordLayout]::RectOf($f))),
                (-not [WordLayout]::IsToolWindow($f)), $(if ($f -eq $victim) { '  <- dragged out' } else { '' }))
}
Assert ((Get-FrameCount) -eq $stackCount) "nothing closed - still $stackCount windows ($((Get-FrameCount)))"
Assert ($torn.Count -eq 1) "exactly one window left the stack ($($torn.Count))"
Assert ($torn.Count -eq 1 -and $torn[0] -eq $victim) 'and it is the one whose tab was dragged out'

# The safety property, and the reason a tear-off is not just a SetWindowPos: the stack takes a
# window's taskbar button and its Alt+Tab entry away, and a window with neither, underneath another at
# the same rectangle, cannot be reached by any means the user has.
$shown = @(Get-Frames | Where-Object { -not [WordLayout]::IsToolWindow($_) })
Assert ($shown.Count -eq 2) "two windows are in Alt+Tab and the taskbar - the one dragged out and the stack's ($($shown.Count))"

$place = Get-StripPlacement (@(Get-Frames))
Assert ($place.Measured -eq $stackCount) "all $stackCount windows could be measured for strip placement (measured $($place.Measured))"
Assert $place.Ok ('every strip still sits between the chrome and the document' + $place.Text)

# ---- and it tore off INSTEAD of reordering ----------------------------------------------------------
#
# The two things one press can mean, asserted not to have both happened. The tab was carried two
# places along the row before it was taken out of it, so a build that committed the reorder as well
# would leave the remaining tabs in a different order from the one they started in - and every
# geometry assertion above would still pass.
#
# Read after the torn-off window has gone, not before: the row is now shorter than the number of Word
# windows, and Get-TabSpot computes its slots from the window count. That mismatch does not throw - the
# slots are merely narrower than the real ones and the clicks still land - so it would be a wrong
# answer arrived at silently.

Write-Step 'Closing the window that was dragged out'
Invoke-ConfirmedKeyOn -Hwnd $victim -Vk $VK.W -Ctrl -What 'Ctrl+W on the window that was dragged out' | Out-Null
if (-not (Wait-Frames ($stackCount - 1) 30)) {
    Write-Note "it did not close; $((Get-FrameCount)) windows left"
}
Assert ((Get-FrameCount) -eq $stackCount - 1) "back to $($stackCount - 1) windows for the sections below ($((Get-FrameCount)))"

$kept = Get-Order
Write-Note "order: $(Format-Order $kept)"
$expected = @($before | Where-Object { $_ -ne $victim })
Assert (Test-SameOrder $kept $expected) 'the tabs left behind are in the order they started in - the drop tore off instead of reordering'
Test-OneRectangle 'After a tab was dragged out of the row'

# ---- carrying a tab past the end of the row ---------------------------------------------------------

Write-Step 'Dragging a tab past the end of the row'
$before = Get-Order
$count = (Get-FrameCount)
$mover = $before[0]

$spot = Get-TabSpot 0
$last = Get-TabSpot ($count - 1)
# Hard against the right-hand end of the strip, so it is the add-in's clamp that decides where the
# tab lands rather than the arithmetic happening to stop in the right place.
$beyond = [Math]::Min($last.Rect.Right + $last.Width, $last.Band.Right - 2)
$onWhat = Get-ClassAt $spot.X $spot.Y
Assert ($onWhat -eq 'WordTabStrip') "the off-the-end drag starts on the strip, not on `"$onWhat`""
[WordLayout]::DragTo($spot.X, $spot.Y, $beyond, $last.Y, 14, 60)
Start-Sleep -Milliseconds 900

$end = Get-Order
Write-Note "order: $(Format-Order $end)"
Assert ($end[$count - 1] -eq $mover) "`"$(Name $mover)`" landed on the last tab rather than off the end"
Assert (@($end | Sort-Object -Unique).Count -eq $count) 'every document still has exactly one tab'
Test-OneRectangle 'After carrying a tab off the end'

# ---- and back to the front ----------------------------------------------------------------------

Write-Step 'Dragging it back to the front'
$spot = Get-TabSpot ($count - 1)
$first = Get-TabSpot 0
$backTo = [Math]::Max($first.Rect.Left + 2, $first.Band.Left + 2)
$onWhat = Get-ClassAt $spot.X $spot.Y
Assert ($onWhat -eq 'WordTabStrip') "the drag back to the front starts on the strip, not on `"$onWhat`""
[WordLayout]::DragTo($spot.X, $spot.Y, $backTo, $first.Y, 14, 60)
Start-Sleep -Milliseconds 900

$front = Get-Order
Write-Note "order: $(Format-Order $front)"
Assert ($front[0] -eq $mover) "`"$(Name $mover)`" is the first tab again"

# ---- a new document still arrives at the end --------------------------------------------------------

# This section is downstream of a window having left the row and been closed, and that is what makes
# it load-bearing rather than a formality. Word does not destroy a frame when its last document
# closes - it hides it and puts the next document straight into it - so the + above may well be
# answered with the very frame the torn-off document was in. The row order is the member array, and a
# recycled frame still has its old place in it, so without a rule the new document silently inherits
# the position of the one that was dragged out: measured, and it arrived as tab 1 of 5.
#
# Which frame Word picks is Word's business, so the log line is reported rather than asserted. What is
# asserted is the part the user can see, and it is true either way: a new document arrives at the end.

Write-Step 'Opening a document after the row has been rearranged'
$before = Get-Order
Set-LogMark
Assert (Invoke-ConfirmedClick -What 'clicking the new-document button' -Point { Get-PlusSpot }) `
       'the click on + landed on the strip'
$arrived = Wait-Frames (@($before).Count + 1) 25
Assert $arrived "clicking + made a document ($((Get-FrameCount)) window(s), expected $(@($before).Count + 1))"

if ($arrived) {
    $withNew = Get-Order
    Write-Note "order: $(Format-Order $withNew)"
    $head = @($withNew | Select-Object -First @($before).Count)
    Assert (Test-SameOrder $head $before) 'the existing tabs kept the order they had been dragged into'
    Assert (-not ($before -contains $withNew[-1])) 'the new document is the last tab'
    foreach ($line in @(Get-LogSince 'rejoining with a new document')) { Write-Note $line.Trim() }
    Test-OneRectangle 'After a new document'
}

# ---- closing a tab leaves the rest in order ----------------------------------------------------------

Write-Step 'Closing a tab out of the middle'
$before = Get-Order
if (@($before).Count -ge 3) {
    $victim = $before[1]
    $expected = @($before | Where-Object { $_ -ne $victim })
    Write-Note ("closing `"{0}`"" -f (Name $victim))

    [WordLayout]::Close($victim)
    $closed = Wait-Frames (@($before).Count - 1) 20
    Assert $closed "that document closed ($((Get-FrameCount)) window(s), expected $(@($before).Count - 1))"

    if ($closed) {
        $left = Get-Order
        Write-Note "order: $(Format-Order $left)"
        Assert (Test-SameOrder $left $expected) 'the tabs either side of it kept their order'
    }
} else {
    Write-Note "only $(@($before).Count) tab(s) - skipping the close-from-the-middle check"
}

# ---- TabTearOff=0 leaves the drag as a plain reorder ------------------------------------------------
#
# The escape hatch, driven rather than assumed - and it was the one switch in this project that had
# never been proven off. A torn-off window is the only thing WordTab does that the user cannot undo
# from inside WordTab, so a machine where the gesture misfires needs a way to stop offering it, and a
# dead switch would leave them nothing.
#
# Checked POSITIVELY, which here means asserting what the gesture DOES rather than what it does not:
# the same drag, along the same route, out past the same threshold, still reorders the row. "Nothing
# happened" is also exactly what a drag that never landed produces, and this project has had three
# negative tests pass for that reason.
#
# It costs one Word restart, because the switch is read once in StackStart. That is the price of a
# switch that refuses the command outright rather than making it inert.

if (-not $KeepOpen) {
    Write-Step 'TabTearOff=0 leaves a drag out of the row as a plain reorder'

    $switchKey = 'HKCU:\Software\WordTab'
    $hadTear   = $false
    $oldTear   = $null
    if (Test-Path $switchKey) {
        $existing = Get-ItemProperty -Path $switchKey -Name 'TabTearOff' -ErrorAction SilentlyContinue
        if ($existing -and ($existing.PSObject.Properties.Name -contains 'TabTearOff')) {
            $hadTear = $true
            $oldTear = $existing.TabTearOff
        }
    }

    try {
        $end = Close-AllWord
        if (-not $end.Closed) {
            Write-Note ("Word would not close ({0}) - skipping the TabTearOff=0 check" -f $end.Reason)
        } else {
            if (-not (Test-Path $switchKey)) { New-Item -Path $switchKey | Out-Null }
            Set-ItemProperty -Path $switchKey -Name 'TabTearOff' -Value 0 -Type DWord
            Start-Sleep -Seconds 2

            Set-LogMark
            foreach ($n in @('Yankee', 'Zulu')) {
                $path = Join-Path $scratch ("wordtab-notear-{0}.rtf" -f $n)
                "{\rtf1\ansi WordTab TabTearOff=0 check - $n.\par}" | Set-Content -Path $path -Encoding Ascii
                Start-Process -FilePath 'winword.exe' -ArgumentList "`"$path`""
                Wait-WordReady (@('Yankee', 'Zulu').IndexOf($n) + 1) 45 | Out-Null
            }
            Start-Sleep -Seconds 2

            # The add-in's own account of the switch, read from the line StackStart writes rather than
            # inferred from the gesture doing nothing.
            $start = @(Get-LogSince 'StackStart  stacking=')
            foreach ($line in $start) { Write-Note "  log: $($line.Trim())" }
            Assert (@($start | Where-Object { $_ -like '*detach=off*' }).Count -ge 1) `
                'the add-in started with tearing off refused'

            $noTearFrames = @(Get-Frames)
            if ($noTearFrames.Count -ne 2) {
                Write-Note "expected two windows for the TabTearOff=0 check, found $($noTearFrames.Count) - skipping"
            } else {
                $before = Get-Order
                Write-Note "order: $(Format-Order $before)"
                $mover = $before[0]

                Set-WordForeground | Out-Null
                Set-LogMark
                $spot   = Get-TabSpot 0
                $slot1  = Get-TabSpot 1
                $stripH = $spot.Band.Bottom - $spot.Band.Top
                $wayBelow = $spot.Band.Bottom + [int]($stripH * 1.6)

                $onWhat = Get-ClassAt $spot.X $spot.Y
                Assert ($onWhat -eq 'WordTabStrip') "the TabTearOff=0 drag starts on the strip, not on `"$onWhat`""

                # Held, so the cursor can be read while the pointer is out past the threshold with the
                # button still down. Read after the release instead and the answer is Word's I-beam:
                # the pointer is over a document by then, and the shape only belongs to us for as long
                # as we hold the mouse capture.
                $cursorOut = [IntPtr]::Zero
                try {
                    [WordLayout]::DragHold($spot.X, $spot.Y, $slot1.X, $spot.Y, 10, 70)
                    [WordLayout]::DragMoveTo($slot1.X, $spot.Y, $slot1.X, $wayBelow, 10, 60)
                    Start-Sleep -Milliseconds 250
                    $cursorOut = [WordLayout]::CursorShape()
                }
                finally {
                    [WordLayout]::DragRelease($slot1.X, $wayBelow)
                }
                Start-Sleep -Milliseconds 1200

                Assert (@(Get-LogSince 'drag left the row').Count -eq 0) 'the tab never comes out of the row with the switch off'
                Assert (@(Get-LogSince 'torn off, now its own window').Count -eq 0) 'and nothing left the stack'
                Write-Note ("cursor out past the threshold 0x{0:X}, IDC_ARROW is 0x{1:X}" -f `
                            [int64]$cursorOut, [int64][WordLayout]::SystemCursor(32512))
                Assert ($cursorOut -eq [WordLayout]::SystemCursor(32512)) `
                    'the pointer stayed an ordinary arrow - nothing was offered that would then be refused'

                # The positive half. Same gesture, same route, and it still does the thing it did
                # before tear-off existed.
                $after = Get-Order
                Write-Note "order: $(Format-Order $after)"
                Assert ($after[1] -eq $mover) "`"$(Name $mover)`" reordered instead: the drag landed, it just meant something else"
                Test-OneRectangle 'With TabTearOff=0'
            }
        }
    } finally {
        # Put the switch back exactly as it was found, including absent. A suite that left TabTearOff=0
        # behind would turn the feature off on this rig and every suite after it would agree that
        # nothing is wrong.
        if ($hadTear) {
            Set-ItemProperty -Path $switchKey -Name 'TabTearOff' -Value $oldTear -Type DWord
        } elseif (Test-Path $switchKey) {
            Remove-ItemProperty -Path $switchKey -Name 'TabTearOff' -ErrorAction SilentlyContinue
        }
        $restored = Get-ItemProperty -Path $switchKey -Name 'TabTearOff' -ErrorAction SilentlyContinue
        $nowIs = if ($restored -and ($restored.PSObject.Properties.Name -contains 'TabTearOff')) { $restored.TabTearOff } else { '(absent)' }
        Write-Note "TabTearOff restored to $nowIs"
    }
}

# ---- done -----------------------------------------------------------------------------------------

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
    Write-Host "PASS  $($script:Checks) checks, 0 failures" -ForegroundColor Green
} else {
    Write-Host "FAIL  $($script:Failures) of $($script:Checks) checks failed" -ForegroundColor Red
}
Write-Host "Add-in log: $env:LOCALAPPDATA\WordTab\wordtab.log (lines starting 'strip' and 'stack')" -ForegroundColor Gray
exit $script:Failures
