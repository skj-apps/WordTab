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
  with a neighbour" and still leave a tab beyond the move that must not have been touched.

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
    [int]$Documents = 4,
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
# Read from a byte offset taken just before the gesture, never from the whole file: a search over
# everything would find the *previous* run's reorder and report a gesture that did nothing as having
# worked. The add-in appends UTF-8 and closes the file after each line, sharing it for read and
# write, so this can read it while Word is running.

$LogPath = Join-Path $env:LOCALAPPDATA 'WordTab\wordtab.log'
function Get-LogMark { if (Test-Path $LogPath) { return (Get-Item $LogPath).Length } else { return 0 } }

function Get-LogSince($offset, $pattern) {
    if (-not (Test-Path $LogPath)) { return @() }
    $stream = [System.IO.File]::Open($LogPath, 'Open', 'Read', 'ReadWrite')
    try {
        # The add-in deletes the log when it passes half a megabyte. If that happened mid-run the
        # offset points past the end of a smaller file, and reading from the start is the honest
        # answer rather than an exception.
        if ($offset -gt $stream.Length) { $offset = 0 }
        $stream.Seek([int64]$offset, 'Begin') | Out-Null
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
        $text = $reader.ReadToEnd()
    } finally { $stream.Dispose() }
    return @(($text -split "`r?`n") | Where-Object { $_ -like "*$pattern*" })
}

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
$mark = Get-LogMark
Set-WordForeground | Out-Null
$spot = Get-TabSpot 0
# Confirmed BEFORE the button goes down, never during. This is one continuous button-down gesture and
# a retry that re-presses would break it. Without the check, a press that landed off the strip also
# produces "no drag started" and "nothing moved" - both assertions pass while nothing was measured.
$onWhat = Get-ClassAt $spot.X $spot.Y
Assert ($onWhat -eq 'WordTabStrip') "the press starts on the strip, not on `"$onWhat`""
[WordLayout]::DragTo($spot.X, $spot.Y, ($spot.X + 2), $spot.Y, 2, 150)
Start-Sleep -Milliseconds 800

Assert (@(Get-LogSince $mark 'drag started').Count -eq 0) 'no drag started - it was a click'
Assert (@(Get-LogSince $mark 'tab moved').Count -eq 0) 'nothing moved in the row'
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
$mark = Get-LogMark
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

    Assert (@(Get-LogSince $mark 'drag started').Count -ge 1) 'the add-in reports a drag in progress'
    Assert (@(Get-LogSince $mark 'tab moved').Count -eq 0) 'a third of a slot is not far enough to reorder anything'

    # Stage two: on to the third slot, still without letting go. The row must rearrange now, while
    # the button is down - a design that waited for the drop would log nothing here.
    $target = Get-TabSpot 2
    $carriedTo = $target.X
    [WordLayout]::DragMoveTo($step1, $spot.Y, $target.X, $target.Y, 12, 60)

    $moves = @(Get-LogSince $mark 'tab moved')
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

Assert (@(Get-LogSince $mark 'drag ended').Count -ge 1) 'the drop was recorded'

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
$mark = Get-LogMark

Set-WordForeground | Out-Null
$spot = Get-TabSpot 0
$target = Get-TabSpot 2
try {
    $onWhat = Get-ClassAt $spot.X $spot.Y
    Assert ($onWhat -eq 'WordTabStrip') "the cancelled drag starts on the strip, not on `"$onWhat`""
    [WordLayout]::DragHold($spot.X, $spot.Y, $target.X, $target.Y, 12, 60)
    Assert (@(Get-LogSince $mark 'tab moved').Count -ge 1) 'the row had rearranged before the cancel'
    [WordLayout]::RightTap($target.X, $target.Y)
}
finally {
    [WordLayout]::DragRelease($target.X, $target.Y)
}
Start-Sleep -Seconds 1

Assert (@(Get-LogSince $mark 'drag cancelled').Count -ge 1) 'the add-in reports the drag was cancelled'

# No menu, ours or Word's. The right button during a drag means cancel, and a cancel that also opened
# a context menu would be two answers to one gesture.
Assert ((Get-WordMenu) -eq [IntPtr]::Zero) 'no context menu was left on screen'

$cancelled = Get-Order
Write-Note "order: $(Format-Order $cancelled)"
Assert (Test-SameOrder $before $cancelled) 'the row went back exactly as it was'

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

Write-Step 'Opening a document after the row has been rearranged'
$before = Get-Order
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
