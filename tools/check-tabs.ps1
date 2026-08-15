<#
.SYNOPSIS
  Measure WordTab's tab affordances: hover, the close button, middle-click, and the new-document
  button.

.DESCRIPTION
  The stack made several Word windows look like one window with tabs. This checks the things you
  can do *to* a tab, which is what separates a tab row from a row of coloured rectangles:

    - the layout has a close button on every tab and one new-document button, and the script's idea
      of where they are comes from the same arithmetic the add-in uses (see WordLayout.Tabs)
    - hovering a tab visibly changes that tab and nothing else. Asserted by photographing the strip
      with the pointer parked off it and again with the pointer on a close button, then comparing
      pixels inside the button against pixels on a different tab - a hover that lit the whole strip
      would pass a "something changed" check and fail this one
    - the new-document button makes a document, and it arrives as a tab in the same stack without
      anything placing it there
    - pressing a close button and releasing somewhere else closes nothing. This is the one that
      protects against a mis-aimed click destroying work
    - the close button closes its own document, from the active tab and from a background one
    - closing a *background* tab leaves the user on the document they were reading. Closing it
      requires activating it first (a modal save prompt behind another window is invisible), so
      without a way back, closing a tab you were not looking at would move you somewhere arbitrary
    - middle-clicking a tab closes it
    - closing down to the last window leaves that window reachable - strip intact, back in Alt+Tab

  **Every click recomputes where it is aiming, immediately before it clicks.** The strip moves: any
  relayout - a document opening, a message bar appearing, the ribbon changing height - shifts it,
  and this was measured moving 46 pixels between two clicks a second apart. A rectangle taken at the
  top of a test and reused three clicks later points into the document, and the click that lands
  there does nothing at all, which reads exactly like a broken feature.

  Word is started on scratch documents written by this script. Word's documents on the dev rig are
  disposable test fixtures. Everything opened here is clean and unmodified, so nothing prompts to
  save - deliberately: the save prompt is real behaviour and belongs in a hand test, not in a script
  that would hang waiting for it.

.PARAMETER Documents
  How many documents to open before the new-document button adds one more.

.PARAMETER KeepOpen
  Leave Word running afterwards, to look at it by hand.

.PARAMETER Screenshot
  Save the hover comparison.

.EXAMPLE
  pwsh -File tools\check-tabs.ps1
  pwsh -File tools\check-tabs.ps1 -KeepOpen -Screenshot
#>
[CmdletBinding()]
param(
    [int]$Documents = 3,
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

$script:Failures = 0
$script:Checks   = 0

function Write-Step($text) { Write-Host "==> $text" -ForegroundColor Cyan }
function Write-Note($text) { Write-Host "    $text" -ForegroundColor DarkGray }
function Assert($condition, $text) {
    $script:Checks++
    if ($condition) { Write-Host "    PASS  $text" -ForegroundColor Green }
    else { $script:Failures++; Write-Host "    FAIL  $text" -ForegroundColor Red }
}

function Get-Frames {
    $found = @()
    foreach ($process in @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue)) {
        $found += [WordLayout]::Frames($process.Id)
    }
    return @($found)
}

function Get-Parts($frame) {
    $kids = [WordLayout]::Children($frame)
    [pscustomobject]@{
        Frame = $frame
        Rect  = [WordLayout]::RectOf($frame)
        Strip = @($kids | Where-Object { $_.Class -eq 'WordTabStrip' })[0]
        Wwf   = @($kids | Where-Object { $_.Class -eq '_WwF' })[0]
        Title = [WordLayout]::TitleOf($frame)
    }
}

# The strip the user is actually looking at: the one belonging to the window on top. Every window in
# the stack has a strip and they all draw the same row, but only one of them is on screen.
function Get-TopStrip {
    $frames = Get-Frames
    if ($frames.Count -eq 0) { throw 'No Word windows.' }
    $top = [WordLayout]::GetForeground()
    if (-not ($frames -contains $top)) { $top = $frames[0] }
    $parts = Get-Parts $top
    if (-not $parts.Strip) { throw "The foreground Word window has no WordTab strip." }
    return $parts
}

# Where to aim, measured now. Every caller goes through this rather than holding a rectangle: see
# the note in the description about the strip moving between one click and the next.
#   'label' - the left third of a tab, clear of its close button
#   'close' - the close button on a tab
#   'plus'  - the new-document button
function Get-Spot($kind, $index) {
    $top = Get-TopStrip
    $count = (Get-Frames).Count
    $layout = [WordLayout]::Tabs($top.Strip.Hwnd, $count)

    if ($kind -eq 'plus') {
        $p = [WordLayout]::Center($layout.Plus)
        return [pscustomobject]@{ X = $p.X; Y = $p.Y; Rect = $layout.Plus; Strip = $top.Strip.Hwnd; Count = $count }
    }
    if ($index -ge $layout.Tabs.Length) { throw "Tab $index does not exist ($($layout.Tabs.Length) tabs)." }

    if ($kind -eq 'close') {
        $c = $layout.Close[$index]
        if ([WordLayout]::IsEmptyRect($c)) { throw "Tab $index has no close button." }
        $p = [WordLayout]::Center($c)
        return [pscustomobject]@{ X = $p.X; Y = $p.Y; Rect = $c; Strip = $top.Strip.Hwnd; Count = $count }
    }

    $t = $layout.Tabs[$index]
    return [pscustomobject]@{
        X = $t.Left + [int](($t.Right - $t.Left) / 3)
        Y = [int](($t.Top + $t.Bottom) / 2)
        Rect = $t; Strip = $top.Strip.Hwnd; Count = $count
    }
}

function Invoke-Spot($kind, $index) {
    $spot = Get-Spot $kind $index
    [WordLayout]::Click($spot.X, $spot.Y)
    return $spot
}

function Format-Rect($r) { "({0},{1} {2}x{3})" -f $r.Left, $r.Top, ($r.Right - $r.Left), ($r.Bottom - $r.Top) }

# Wait for the window count to settle on an expected number rather than sleeping a fixed time:
# Documents.Add and WM_CLOSE are both asynchronous, and how long Word takes varies with what else it
# is doing. Bounded - on timeout the assertion runs anyway and reports what it found.
function Wait-Frames($expected, $seconds = 20) {
    $deadline = (Get-Date).AddSeconds($seconds)
    while ((Get-Date) -lt $deadline) {
        if ((Get-Frames).Count -eq $expected) {
            Start-Sleep -Milliseconds 1200      # let the janitor join or drop it and repaint
            if ((Get-Frames).Count -eq $expected) { return $true }
        }
        Start-Sleep -Milliseconds 400
    }
    return $false
}

function Test-OneRectangle($label) {
    $parts = @(Get-Frames | ForEach-Object { Get-Parts $_ })
    $rects = @($parts | ForEach-Object { Format-Rect $_.Rect } | Sort-Object -Unique)
    Assert ($rects.Count -eq 1) "$label - all $($parts.Count) windows still at one rectangle ($($rects.Count) distinct)"
}

# A photograph of the strip as it is on screen, cursor excluded. PrintWindow would ask the window to
# draw itself, which is a different claim; this is what the user can see.
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

# ---- get Word up with N documents --------------------------------------------------------------
#
# Any Word already running is closed first, and gracefully. Killing Word makes it offer to recover
# those documents on the next launch, which adds a window nobody asked for and a recovery pane over
# the document - and a check that starts in that state is measuring something other than what it
# says it is. This has already happened once.

$running = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue)
if ($running.Count -gt 0) {
    Write-Step "Closing $($running.Count) Word process(es) already running"
    foreach ($process in $running) { $process.CloseMainWindow() | Out-Null }
    $deadline = (Get-Date).AddSeconds(20)
    while ((Get-Date) -lt $deadline -and (Get-Process -Name WINWORD -ErrorAction SilentlyContinue)) {
        Start-Sleep -Milliseconds 500
    }
    $left = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue)
    if ($left.Count -gt 0) { throw "Word would not close ($($left.Count) left) - close it by hand, saving or discarding as you like, then re-run." }
    Start-Sleep -Seconds 2
}

$scratch = Join-Path $env:TEMP 'wordtab-check'
New-Item -ItemType Directory -Path $scratch -Force | Out-Null

for ($i = 1; $i -le $Documents; $i++) {
    $path = Join-Path $scratch "wordtab-tabs-$i.rtf"
    "{\rtf1\ansi WordTab tab check - document $i.\par}" | Set-Content -Path $path -Encoding Ascii
    Start-Process -FilePath 'winword.exe' -ArgumentList "`"$path`""
    Start-Sleep -Seconds $(if ($i -eq 1) { 14 } else { 7 })
}

$deadline = (Get-Date).AddSeconds(45)
while ((Get-Date) -lt $deadline -and (Get-Frames).Count -lt $Documents) { Start-Sleep -Milliseconds 500 }

$frames = Get-Frames
Write-Note "$($frames.Count) visible Word frame(s)"
if ($frames.Count -lt 2) { throw "Need at least 2 Word frames to check tabs; found $($frames.Count)." }

$settle = (Get-Date).AddSeconds(15)
while ((Get-Date) -lt $settle) {
    $parts = @(Get-Frames | ForEach-Object { Get-Parts $_ })
    $ready = @($parts | Where-Object { $_.Strip -and $_.Wwf -and $_.Strip.Bottom -eq $_.Wwf.Top })
    if ($ready.Count -eq $parts.Count) { break }
    Start-Sleep -Milliseconds 500
}

# ---- the layout ---------------------------------------------------------------------------------

Write-Step 'The tab row'
$top = Get-TopStrip
$count = (Get-Frames).Count
$layout = [WordLayout]::Tabs($top.Strip.Hwnd, $count)

foreach ($i in 0..($count - 1)) {
    $c = $layout.Close[$i]
    Write-Note ("tab {0} {1}  close {2}" -f $i, (Format-Rect $layout.Tabs[$i]),
                $(if ([WordLayout]::IsEmptyRect($c)) { 'none' } else { Format-Rect $c }))
}
Write-Note ("plus {0}  present={1}" -f (Format-Rect $layout.Plus), $layout.HasPlus)

$withClose = @(0..($count - 1) | Where-Object { -not [WordLayout]::IsEmptyRect($layout.Close[$_]) })
Assert ($withClose.Count -eq $count) "every one of the $count tabs has a close button ($($withClose.Count))"
Assert ($layout.HasPlus) 'the new-document button is present and inside the strip'

$stripRect = [WordLayout]::RectOf($top.Strip.Hwnd)
$inside = ($layout.Plus.Right -le $stripRect.Right) -and ($layout.Plus.Left -ge $stripRect.Left)
Assert $inside 'the new-document button is reachable (within the strip, not pushed off the end)'

# Every tab is a different window. Established by clicking, not assumed: the tab row is in the order
# windows joined the stack and Windows enumerates them in z-order, which are different orders.
Write-Step 'Clicking each tab'
$activated = @()
for ($i = 0; $i -lt $count; $i++) {
    Invoke-Spot 'label' $i | Out-Null
    Start-Sleep -Milliseconds 900
    $now = [WordLayout]::GetForeground()
    $activated += $now
    Write-Note ("tab {0} -> 0x{1:X}  `"{2}`"" -f $i, [int64]$now, [WordLayout]::TitleOf($now))
}
Assert (@($activated | Sort-Object -Unique).Count -eq $count) "each of the $count tabs is a different window"

# ---- hover --------------------------------------------------------------------------------------
#
# The claim is not "the strip repaints" but "the tab under the pointer changes and the others do
# not". Both halves are asserted, because a hover that lit everything would satisfy the first.

Write-Step 'Hovering a tab'
[WordLayout]::MouseTo(4, 4)                      # off the strip entirely, so nothing is hot
Start-Sleep -Milliseconds 900
$top = Get-TopStrip
$cold = Get-StripShot $top.Strip.Hwnd

$target = if ($count -ge 2) { 1 } else { 0 }     # a tab that is not the active one, if there is one
$other  = if ($target -eq 0) { 1 } else { 0 }
$spot = Get-Spot 'close' $target
$control = (Get-Spot 'label' $other).Rect
$control.Right = $control.Left + [int](($control.Right - $control.Left) / 3)

[WordLayout]::MouseTo($spot.X, $spot.Y)
Start-Sleep -Milliseconds 900
$hot = Get-StripShot $top.Strip.Hwnd

# If the strip moved between the two photographs the comparison means nothing, so say so rather
# than reporting a hover failure that is really a relayout.
$still = ($cold.Origin.Left -eq $hot.Origin.Left) -and ($cold.Origin.Top -eq $hot.Origin.Top)
Assert $still 'the strip held still between the two photographs'

$onButton  = Measure-Changed $cold.Bitmap $hot.Bitmap $cold.Origin $spot.Rect
$elsewhere = Measure-Changed $cold.Bitmap $hot.Bitmap $cold.Origin $control

Write-Note "pixels changed under the pointer: $onButton   on another tab: $elsewhere"
Assert ($onButton -gt 0) "hovering tab $target visibly changes it ($onButton pixels)"
Assert ($elsewhere -eq 0) "hovering tab $target changes nothing on tab $other ($elsewhere pixels)"

if ($Screenshot) {
    New-Item -ItemType Directory -Path $ShotDir -Force | Out-Null
    $cold.Bitmap.Save((Join-Path $ShotDir 'wordtab-strip-cold.png'), [System.Drawing.Imaging.ImageFormat]::Png)
    $hot.Bitmap.Save((Join-Path $ShotDir 'wordtab-strip-hover.png'), [System.Drawing.Imaging.ImageFormat]::Png)
    Write-Note (Join-Path $ShotDir 'wordtab-strip-cold.png')
    Write-Note (Join-Path $ShotDir 'wordtab-strip-hover.png')
}
$cold.Bitmap.Dispose(); $hot.Bitmap.Dispose()
[WordLayout]::MouseTo(4, 4)

# ---- the new-document button ---------------------------------------------------------------------

Write-Step 'The new-document button'
$before = (Get-Frames).Count
Invoke-Spot 'plus' 0 | Out-Null

$arrived = Wait-Frames ($before + 1) 25
$now = (Get-Frames).Count
Assert $arrived "clicking + made a document: $now window(s), expected $($before + 1)"

if ($arrived) {
    Test-OneRectangle 'After +'
    $top = Get-TopStrip
    $layout = [WordLayout]::Tabs($top.Strip.Hwnd, $now)
    Assert ($layout.Tabs.Length -eq $now) "the new document has a tab of its own ($($layout.Tabs.Length) tabs for $now windows)"
    Write-Note ("foreground is now `"{0}`"" -f [WordLayout]::TitleOf([WordLayout]::GetForeground()))
}

# ---- pressing a close button and letting go somewhere else ----------------------------------------
#
# The safety property. A close button that fires on the press, or on a release anywhere, means one
# mis-aimed click destroys a document. This asserts that the gesture everyone uses to change their
# mind actually changes it.

Write-Step 'Pressing a close button and sliding off'
$count = (Get-Frames).Count
$button = Get-Spot 'close' 0
$away = Get-Spot 'label' 0
[WordLayout]::PressAndSlideOff($button.X, $button.Y, $away.X, $away.Y)
Start-Sleep -Seconds 3

Assert ((Get-Frames).Count -eq $count) "nothing closed ($((Get-Frames).Count) window(s), still $count)"

# ---- closing the active tab ------------------------------------------------------------------------

Write-Step 'Closing the active tab with its close button'
$count = (Get-Frames).Count
$victim = $count - 1

# Select that tab first, so "the active tab" is a known index rather than something to deduce -
# clicking a tab to find out which one is active would change the answer.
Invoke-Spot 'label' $victim | Out-Null
Start-Sleep -Milliseconds 1000
$activeFrame = [WordLayout]::GetForeground()
Write-Note ("closing tab {0}, `"{1}`"" -f $victim, [WordLayout]::TitleOf($activeFrame))

Invoke-Spot 'close' $victim | Out-Null
$closed = Wait-Frames ($count - 1) 20
Assert $closed "the close button closed that document ($((Get-Frames).Count) window(s), expected $($count - 1))"

# Gone from the tab row, which is the claim. Not "the window handle is destroyed": frame lifetime is
# not document lifetime in Word, and closing a document was measured to hide one frame and destroy a
# different one. The user's document is gone either way; which HWND Word recycled is its business.
Assert (-not ((Get-Frames) -contains $activeFrame)) 'that document is out of the tab row'
if ((Get-Frames).Count -gt 1) { Test-OneRectangle 'After closing the active tab' }

# ---- closing a background tab, and staying where you were --------------------------------------------
#
# Closing a tab activates it first: a stacked window sits underneath another one at the same
# rectangle, and a modal save prompt belonging to the window underneath can end up invisible. The
# price of that is being moved off the document you were reading, and this is the refund.

Write-Step 'Closing a background tab'
$count = (Get-Frames).Count
if ($count -ge 3) {
    Invoke-Spot 'label' 0 | Out-Null
    Start-Sleep -Milliseconds 1000
    $wasOn = [WordLayout]::GetForeground()
    Write-Note ("sitting on `"{0}`"" -f [WordLayout]::TitleOf($wasOn))

    Invoke-Spot 'close' 2 | Out-Null
    $closed = Wait-Frames ($count - 1) 20
    Assert $closed "the background tab's document closed ($((Get-Frames).Count) window(s), expected $($count - 1))"

    Start-Sleep -Milliseconds 1500
    $endedOn = [WordLayout]::GetForeground()
    Write-Note ("ended up on `"{0}`"" -f [WordLayout]::TitleOf($endedOn))
    Assert ($endedOn -eq $wasOn) 'still on the document the user was reading, not the one that closed'
    if ((Get-Frames).Count -gt 1) { Test-OneRectangle 'After closing a background tab' }
} else {
    Write-Note "only $count window(s) left - skipping the background-tab check"
}

# ---- middle-click ------------------------------------------------------------------------------------

Write-Step 'Middle-clicking a tab'
$count = (Get-Frames).Count
if ($count -ge 2) {
    $spot = Get-Spot 'label' ($count - 1)
    [WordLayout]::MiddleClick($spot.X, $spot.Y)
    $closed = Wait-Frames ($count - 1) 20
    Assert $closed "middle-click closed it ($((Get-Frames).Count) window(s), expected $($count - 1))"
} else {
    Write-Note "only $count window(s) left - skipping the middle-click check"
}

# ---- down to the last window -------------------------------------------------------------------------
#
# The property that outranks every feature here: the last window must be reachable. Everything the
# add-in does to hide a window has to come undone, or "the add-in ate my document" is the first
# impression - and it is not a recoverable one.

Write-Step 'Closing down to the last window'
for ($guard = 0; $guard -lt 8; $guard++) {
    $open = @(Get-Frames)
    if ($open.Count -le 1) { break }
    [WordLayout]::Close($open[0])
    Start-Sleep -Seconds 4
}
$last = @(Get-Frames)
if ($last.Count -eq 1) {
    Assert (-not [WordLayout]::IsToolWindow($last[0])) 'the last window is back in Alt+Tab'
    $parts = Get-Parts $last[0]
    Assert ($parts.Strip -and $parts.Wwf -and $parts.Strip.Bottom -eq $parts.Wwf.Top) `
        'the last window still has its strip, correctly placed'

    $layout = [WordLayout]::Tabs($parts.Strip.Hwnd, 1)
    Assert ($layout.HasPlus) 'the new-document button is still there with one document open'
    Assert (-not [WordLayout]::IsEmptyRect($layout.Close[0])) 'the last tab still has a close button'
} else {
    Write-Note "expected one window left, found $($last.Count) - skipping the last-window checks"
}

# ---- done ---------------------------------------------------------------------------------------
#
# CloseMainWindow and a long wait, and no Kill. Killing Word leaves it offering to recover these
# documents on the next launch, which quietly changes what the *next* run measures.

if (-not $KeepOpen) {
    Write-Step 'Closing Word'

    # One window at a time, not CloseMainWindow. A Word process holding a stack has N top-level
    # windows and CloseMainWindow closes exactly one of them - measured, four documents left open
    # after two rounds of it. Each of these is the same WM_CLOSE the user's own close button sends.
    for ($guard = 0; $guard -lt 12; $guard++) {
        $open = @(Get-Frames)
        if ($open.Count -eq 0) { break }
        [WordLayout]::Close($open[0])
        Start-Sleep -Seconds 3
    }
    foreach ($process in @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue)) { $process.CloseMainWindow() | Out-Null }
    $deadline = (Get-Date).AddSeconds(20)
    while ((Get-Date) -lt $deadline -and (Get-Process -Name WINWORD -ErrorAction SilentlyContinue)) {
        Start-Sleep -Milliseconds 500
    }
    $stuck = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue)
    if ($stuck.Count -gt 0) { Write-Note "$($stuck.Count) Word process(es) would not close - left running rather than killed" }
}

Write-Host ''
if ($script:Failures -eq 0) {
    Write-Host "PASS  $($script:Checks) checks, 0 failures" -ForegroundColor Green
} else {
    Write-Host "FAIL  $($script:Failures) of $($script:Checks) checks failed" -ForegroundColor Red
}
Write-Host "Add-in log: $env:LOCALAPPDATA\WordTab\wordtab.log (lines starting 'strip' and 'stack')" -ForegroundColor Gray
exit $script:Failures
