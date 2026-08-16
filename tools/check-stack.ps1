<#
.SYNOPSIS
  Measure WordTab's stack: N Word windows held at one rectangle, with one tab row, switchable.

.DESCRIPTION
  The product claim is "several documents look like one window with tabs". This checks the parts
  of that claim which are measurable from outside the process:

    - every stacked window sits at exactly one rectangle
    - every stacked window's document frame (`_WwF`) is identical, including the ones that are not
      focused. This is the one that matters: Word only lays out the window that has focus, so a
      background window keeps whatever interior it had unless the add-in copies the focused
      window's layout onto it. Getting this wrong is invisible until you switch tabs.
    - moving or resizing the active window takes the others with it
    - clicking each tab brings a different window to the front

  Then it closes a document and checks the stack closes up behind it, because frame lifetime is not
  document lifetime in Word - closing one of two documents was measured to hide one frame and
  destroy a different one.

  Word is started on scratch documents written by this script. Word's documents on the dev rig are
  disposable test fixtures.

.PARAMETER Documents
  How many documents to open. Three is enough to tell "the stack works" from "two windows happen to
  be at the same place".

.PARAMETER KeepOpen
  Leave Word running afterwards, to look at it by hand.

.PARAMETER Screenshot
  Save a PNG of each frame at the end.

.EXAMPLE
  pwsh -File tools\check-stack.ps1
  pwsh -File tools\check-stack.ps1 -Documents 3 -KeepOpen -Screenshot
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

# Always ask for the count through this, never `(Get-Frames).Count`: a PowerShell function that
# returns an empty array returns *nothing*, and .Count on nothing is a hard error under
# Set-StrictMode rather than 0 - which bites exactly when the last document has closed.
function Get-FrameCount { return Get-WordFrameTally }

function Get-Parts($frame) {
    $kids = [WordLayout]::Children($frame)
    [pscustomobject]@{
        Frame  = $frame
        Rect   = [WordLayout]::RectOf($frame)
        Strip  = @($kids | Where-Object { $_.Class -eq 'WordTabStrip' }) | Select-Object -First 1
        Wwf    = @($kids | Where-Object { $_.Class -eq '_WwF' }) | Select-Object -First 1
        Title  = [WordLayout]::TitleOf($frame)
    }
}

function Format-Rect($r) { "({0},{1} {2}x{3})" -f $r.Left, $r.Top, ($r.Right - $r.Left), ($r.Bottom - $r.Top) }

# The window on top, or the first one if the foreground belongs to another application entirely.
#
# Written as a search with a fallback rather than as `@(... | Where-Object ...)[0]`, because under
# Set-StrictMode indexing an empty result is a hard error, not $null - the same trap as
# `(Get-Frames).Count` on no frames. And "nothing matched" is a state this genuinely reaches:
# switching tabs ends in SetForegroundWindow, which Windows refuses from a process that is not
# already the foreground application, so anything that steals the desktop mid-run leaves the
# foreground outside Word. That has to read as a failed assertion, not as a crashed script.
function Get-TopParts($all) {
    $fg = [WordLayout]::GetForeground()
    foreach ($p in @($all)) { if ($p.Frame -eq $fg) { return $p } }
    if (@($all).Count -gt 0) { return @($all)[0] }
    return $null
}

# Set-WordForeground and Test-WordForeground now come from WordTabHarness.ps1. The local copy here
# was one of six that had drifted apart; the shared one also minimises an application that will not
# give the foreground up, and knows the three windows it must never minimise.

# One rectangle, or several? Compared as strings because that is exactly the question - identical
# or not - and it prints usefully when the answer is "not".
function Test-Stacked($label) {
    $frames = Get-Frames
    $parts = @($frames | ForEach-Object { Get-Parts $_ })

    $rects = @($parts | ForEach-Object { Format-Rect $_.Rect } | Sort-Object -Unique)
    Write-Step $label
    foreach ($p in $parts) {
        $wwf = if ($p.Wwf) { "y $($p.Wwf.Top)..$($p.Wwf.Bottom) x $($p.Wwf.Left)..$($p.Wwf.Right)" } else { 'none' }
        $strip = if ($p.Strip) { "y $($p.Strip.Top)..$($p.Strip.Bottom)" } else { 'none' }
        Write-Note ("0x{0:X}  {1}  strip {2}  _WwF {3}  `"{4}`"" -f [int64]$p.Frame, (Format-Rect $p.Rect), $strip, $wwf, $p.Title)
    }

    Assert ($rects.Count -eq 1) "$label - all $($parts.Count) windows at one rectangle ($($rects.Count) distinct: $($rects -join ' '))"

    $strips = @($parts | Where-Object { $_.Strip })
    Assert ($strips.Count -eq $parts.Count) "$label - every window has a strip ($($strips.Count) of $($parts.Count))"

    # The layout-oracle check. Word lays out only the focused window; if the add-in does not copy
    # that layout onto the others, the background windows keep a stale document frame and the
    # mismatch only becomes visible when you switch to one.
    $docs = @($parts | Where-Object { $_.Wwf } | ForEach-Object { "$($_.Wwf.Left),$($_.Wwf.Top),$($_.Wwf.Right),$($_.Wwf.Bottom)" } | Sort-Object -Unique)
    Assert ($docs.Count -eq 1) "$label - every document frame identical, focused or not ($($docs.Count) distinct)"

    # Consistency is not correctness. Every window agreeing on a wrong layout passes every check
    # above - which is exactly what happened once: closing a document put the strip at y=0 on top
    # of the ribbon in all of them, identically. So each window is also checked in absolute terms.
    $wrong = @()
    foreach ($p in $parts) {
        if (-not $p.Strip -or -not $p.Wwf) { continue }
        $kids = [WordLayout]::Children($p.Frame)
        $client = [WordLayout]::ClientOf($p.Frame)
        $above = $kids |
                 Where-Object { $_.Visible -and $_.Hwnd -ne $p.Strip.Hwnd -and $_.Hwnd -ne $p.Wwf.Hwnd -and
                                $_.Bottom -le $p.Strip.Top -and $_.Width -gt ($client.Right * 0.6) } |
                 Sort-Object Bottom -Descending | Select-Object -First 1
        if ($p.Strip.Bottom -ne $p.Wwf.Top) { $wrong += "0x$('{0:X}' -f [int64]$p.Frame) strip bottom $($p.Strip.Bottom) != document top $($p.Wwf.Top)" }
        elseif (-not $above)                { $wrong += "0x$('{0:X}' -f [int64]$p.Frame) nothing above the strip - it is sitting at the top of the window" }
        elseif ($above.Bottom -ne $p.Strip.Top) { $wrong += "0x$('{0:X}' -f [int64]$p.Frame) $($above.Class) ends at $($above.Bottom), strip starts at $($p.Strip.Top)" }
    }
    Assert ($wrong.Count -eq 0) ("$label - every strip sits between the chrome and the document" + $(if ($wrong.Count) { ": " + ($wrong -join '; ') } else { '' }))

    return $parts
}

# ---- get Word up with N documents --------------------------------------------------------------

$scratch = Join-Path $env:TEMP 'wordtab-check'
New-Item -ItemType Directory -Path $scratch -Force | Out-Null

$startedWord = -not (Get-Process -Name WINWORD -ErrorAction SilentlyContinue)

# Wait for each document to arrive rather than sleeping 14 seconds for the first and 7 for each one
# after. Those numbers were a guess about how long this machine takes, and on a three-document run
# they were 28 seconds of the suite on their own. Waiting for the window to exist AND to have a strip
# on it is both faster when Word is quick and more correct when it is slow: a frame appears before
# the add-in has drawn anything, and a suite that measures the strip in that gap measures nothing.
#
# This is a precondition, never an assertion. Nothing below asserts the document count - it throws if
# there are too few - so polling for it cannot make a check pass by definition. That distinction is
# what makes a wait honest, and it is why several fixed sleeps further down are deliberately left
# alone.
for ($i = 1; $i -le $Documents; $i++) {
    $path = Join-Path $scratch "wordtab-stack-$i.rtf"
    "{\rtf1\ansi WordTab stack check - document $i.\par}" | Set-Content -Path $path -Encoding Ascii
    Start-Process -FilePath 'winword.exe' -ArgumentList "`"$path`""
    if (-not (Wait-WordReady $i 45)) {
        Write-Note "document $i did not arrive with a strip on it within 45s - carrying on and letting the assertions report it"
    }
}

$frames = Get-Frames
Write-Note "$($frames.Count) visible Word frame(s)"
if ($frames.Count -lt 2) { throw "Need at least 2 Word frames to check a stack; found $($frames.Count)." }

# Wait for the layout to settle rather than measuring the settle itself: Word spends a second or
# two rearranging its chrome after a window appears, and the add-in joins windows to the stack and
# corrects itself on a half-second cadence. Bounded - if it never settles, the assertions run
# anyway and report what they found.
$settle = (Get-Date).AddSeconds(15)
while ((Get-Date) -lt $settle) {
    $parts = @(Get-Frames | ForEach-Object { Get-Parts $_ })
    $ready = @($parts | Where-Object { $_.Strip -and $_.Wwf -and $_.Strip.Bottom -eq $_.Wwf.Top })
    $rects = @($parts | ForEach-Object { Format-Rect $_.Rect } | Sort-Object -Unique)
    if ($ready.Count -eq $parts.Count -and $rects.Count -eq 1) { break }
    Start-Sleep -Milliseconds 500
}

$parts = Test-Stacked 'Stacked'

# ---- move and resize the active window ---------------------------------------------------------

$active = [WordLayout]::GetForeground()
if (-not ($frames -contains $active)) { $active = $frames[0]; [WordLayout]::Focus($active) | Out-Null }

Write-Step 'Moving the active window'
$r = [WordLayout]::RectOf($active)
[WordLayout]::MoveTo($active, $r.Left - 60, $r.Top + 40)
Start-Sleep -Milliseconds 900
Test-Stacked 'After moving the active window' | Out-Null

Write-Step 'Resizing the active window'
[WordLayout]::Resize($active, 1100, 820)
Start-Sleep -Milliseconds 1200
Test-Stacked 'After resizing the active window' | Out-Null

# ---- clicking the tabs -------------------------------------------------------------------------
#
# The tab geometry comes from [WordLayout]::Tabs, which mirrors ComputeLayout in
# src\native\strip.cpp. If the two ever diverge, the clicks land between tabs and the "distinct
# window per tab" assertion below fails loudly.
#
# The click aims at the middle of the tab's *label*, not the middle of the tab: the right-hand end
# now carries a close button, and a click there closes the document rather than selecting it.

Write-Step 'Clicking each tab'
[WordLayout]::Focus($active) | Out-Null
Start-Sleep -Milliseconds 500

Set-WordForeground | Out-Null
Assert (Test-WordForeground) 'Word is the foreground application, so a tab click can switch to it'

$parts = @(Get-Frames | ForEach-Object { Get-Parts $_ })
$count = $parts.Count
$top   = Get-TopParts $parts

$activated = @()
$landed = 0
for ($i = 0; $i -lt $count; $i++) {
    # Measured inside the loop, not once before it. The strip moves whenever Word relays a window
    # out, and it was measured shifting 46 pixels between two clicks a second apart - after which a
    # rectangle taken before the loop points into the document and the click does nothing at all.
    #
    # The scriptblock is the point: Invoke-ConfirmedClick re-runs it on every attempt, so a retry
    # after the foreground was taken back re-measures rather than re-using a rectangle that has
    # since moved. A click that never reached the strip and a tab that does not switch windows are
    # otherwise the same evidence.
    $aim = {
        $live = Get-TopParts @(Get-Frames | ForEach-Object { Get-Parts $_ })
        if (-not $live) { $live = $top }
        $tab = ([WordLayout]::Tabs($live.Strip.Hwnd, $count)).Tabs[$i]
        [pscustomobject]@{
            X = $tab.Left + [int](($tab.Right - $tab.Left) / 3)
            Y = [int](($tab.Top + $tab.Bottom) / 2)
        }
    }
    $x = (& $aim).X
    if (Invoke-ConfirmedClick -What "clicking tab $i" -Point $aim) { $landed++ }
    Start-Sleep -Milliseconds 900

    $now = [WordLayout]::GetForeground()
    $title = [WordLayout]::TitleOf($now)
    Write-Note ("tab {0} at x={1} -> foreground 0x{2:X} `"{3}`"" -f $i, $x, [int64]$now, $title)
    $activated += $now
}

Assert ($landed -eq $count) "every one of the $count tab clicks was confirmed to land on the strip ($landed of $count)"

$distinct = @($activated | Sort-Object -Unique)
Assert ($distinct.Count -eq $count) "each of the $count tabs brought a different window forward ($($distinct.Count) distinct)"
Assert (@($activated | Where-Object { $frames -contains $_ }).Count -eq $count) 'every tab activated a Word frame'

Test-Stacked 'After switching tabs' | Out-Null

# ---- one window to the rest of Windows ----------------------------------------------------------
#
# The taskbar and Alt+Tab cannot be enumerated through any API, so what is asserted here is the
# mechanism - exactly one window presented, and it is the active one - and the result itself is
# photographed for a human to confirm. WS_EX_TOOLWINDOW is what Alt+Tab reads; ITaskbarList is what
# the taskbar is told, and its calls are logged by the add-in.

Write-Step 'One window presented to the taskbar and Alt+Tab'
$parts = @(Get-Frames | ForEach-Object { Get-Parts $_ })
$foreground = [WordLayout]::GetForeground()
$tools = @($parts | Where-Object { [WordLayout]::IsToolWindow($_.Frame) })
$activeIsTool = [WordLayout]::IsToolWindow($foreground)

foreach ($p in $parts) {
    Write-Note ("0x{0:X}  toolwindow={1}  {2}" -f [int64]$p.Frame, [WordLayout]::IsToolWindow($p.Frame),
                $(if ($p.Frame -eq $foreground) { '<- active' } else { '' }))
}
Assert ($tools.Count -eq $parts.Count - 1) "exactly one window is left in Alt+Tab ($($tools.Count) of $($parts.Count) hidden)"
Assert (-not $activeIsTool) 'the one left is the active one'

if ($Screenshot) {
    New-Item -ItemType Directory -Path $ShotDir -Force | Out-Null
    $screen = [WordLayout]::ScreenRect()
    [WordLayout]::AltTabHold()
    Start-Sleep -Milliseconds 900
    try {
        $bmp = New-Object System.Drawing.Bitmap(($screen.Right - $screen.Left), ($screen.Bottom - $screen.Top))
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.CopyFromScreen($screen.Left, $screen.Top, 0, 0, $bmp.Size)
        $bmp.Save((Join-Path $ShotDir 'wordtab-alttab.png'), [System.Drawing.Imaging.ImageFormat]::Png)
        $g.Dispose(); $bmp.Dispose()
        Write-Note (Join-Path $ShotDir 'wordtab-alttab.png')
    } finally {
        [WordLayout]::AltRelease()
    }

    # And the taskbar, which is the bottom strip of the primary screen.
    Start-Sleep -Milliseconds 600
    $h = 120
    $bmp = New-Object System.Drawing.Bitmap(($screen.Right - $screen.Left), $h)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($screen.Left, ($screen.Bottom - $h), 0, 0, $bmp.Size)
    $bmp.Save((Join-Path $ShotDir 'wordtab-taskbar.png'), [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $bmp.Dispose()
    Write-Note (Join-Path $ShotDir 'wordtab-taskbar.png')
}

# ---- minimising ----------------------------------------------------------------------------------
#
# The stack is one window to the user, so it goes down and comes back as one. Only the active window
# has a taskbar button, so if the others did not come back with it they would be stranded: minimised,
# no button, no Alt+Tab entry.

Write-Step 'Minimising the stack'
$before = @(Get-Frames)
[WordLayout]::Focus($before[0]) | Out-Null
$activeNow = [WordLayout]::GetForeground()
if (-not ($before -contains $activeNow)) { $activeNow = $before[0] }

[WordLayout]::Show($activeNow, 6)          # SW_MINIMIZE
Start-Sleep -Milliseconds 1500
$down = @($before | Where-Object { [WordLayout]::Minimized($_) })
Assert ($down.Count -eq $before.Count) "every window went down with it ($($down.Count) of $($before.Count))"

[WordLayout]::Show($activeNow, 9)          # SW_RESTORE
Start-Sleep -Seconds 2
$up = @($before | Where-Object { -not [WordLayout]::Minimized($_) })
Assert ($up.Count -eq $before.Count) "every window came back with it ($($up.Count) of $($before.Count))"
Test-Stacked 'After minimise and restore' | Out-Null

# ---- closing a document ------------------------------------------------------------------------

Write-Step 'Closing one document'
$victim = (Get-Frames)[0]
[WordLayout]::Close($victim)
Start-Sleep -Seconds 4

$after = Get-Frames
Assert ($after.Count -eq $count - 1) "the stack closed up: $($after.Count) window(s) left, expected $($count - 1)"
if ($after.Count -gt 1) { Test-Stacked 'After closing a document' | Out-Null }

# Down to one window, which is the case that has to be safe above all others: the last window must
# be reachable. A window with no taskbar button and no Alt+Tab entry is gone as far as the user is
# concerned, and "the add-in ate my document" is not a recoverable first impression.
Write-Step 'Closing down to the last window'
for ($guard = 0; $guard -lt 6; $guard++) {
    $open = @(Get-Frames)
    if ($open.Count -le 1) { break }
    [WordLayout]::Close($open[0])
    Wait-Until { @(Get-Frames).Count -lt $open.Count } 10 300 | Out-Null
}

# A FIXED settle, and it is deliberate. The first version of this loop replaced the old
# `Start-Sleep -Seconds 4` after each close with the count-wait above and nothing else, reasoning
# that the assertions below are about the surviving window rather than about how many windows there
# are. That reasoning was wrong in a way worth recording: the count drops the instant Word destroys
# the window, but "the last window is back in Alt+Tab" is about the add-in taking WS_EX_TOOLWINDOW
# back OFF it, which its janitor does on its own half-second cadence AFTER the close. The wait
# returned before the add-in had done it, and the very next battery went red here.
#
# So the sleep the assertion actually depended on is restored, once, after the loop instead of after
# every close. What it buys is the add-in's reconcile, and how long that takes is part of the claim -
# waiting for `-not IsToolWindow` instead would make the assertion unable to fail.
Start-Sleep -Seconds 4
$last = @(Get-Frames)
if ($last.Count -eq 1) {
    Assert (-not [WordLayout]::IsToolWindow($last[0])) 'the last window is back in Alt+Tab'
    $parts = @(Get-Parts $last[0])
    Assert ($parts[0].Strip -and $parts[0].Wwf -and $parts[0].Strip.Bottom -eq $parts[0].Wwf.Top) `
        'the last window still has its strip, correctly placed'
} else {
    Write-Note "expected one window left, found $($last.Count) - skipping the last-window checks"
}

# ---- pictures ----------------------------------------------------------------------------------

if ($Screenshot) {
    Write-Step 'Screenshots'
    New-Item -ItemType Directory -Path $ShotDir -Force | Out-Null
    $index = 0
    foreach ($frame in (Get-Frames)) {
        $path = Join-Path $ShotDir ("wordtab-stack-{0}.png" -f $index++)
        $r = [WordLayout]::RectOf($frame)
        $bitmap = New-Object System.Drawing.Bitmap(($r.Right - $r.Left), ($r.Bottom - $r.Top))
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        $dc = $graphics.GetHdc()
        [WordLayout]::Print($frame, $dc) | Out-Null
        $graphics.ReleaseHdc($dc)
        $bitmap.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
        $graphics.Dispose(); $bitmap.Dispose()
        Write-Note $path
    }
}

# ---- done ---------------------------------------------------------------------------------------

# NEVER Kill, and that is not fastidiousness. A killed Word offers to recover those documents on the
# next launch, and the recovered documents then make the NEXT suite measure seven windows where it
# asserts three. This suite runs first in check-all, so its cleanup is the one with the most to
# poison - and it was doing exactly that: CloseMainWindow closes one of N frames, then Kill took the
# rest. The same violation was found and fixed inside soak-stack in the scrolling-row slice; this
# copy survived because nothing had gone looking for the second one.
#
# Close-AllWord posts WM_CLOSE per frame and stops if Word asks a question, because a test script may
# not answer a save prompt it did not raise.
if (-not $KeepOpen -and $startedWord) {
    Write-Step 'Closing Word'
    $end = Close-AllWord
    if (-not $end.Closed) {
        Write-Note ("Word is still up ({0}): {1}" -f $end.Reason, (Format-WordWindow $end.Dialog))
        Write-Note 'Answer it by hand before running the next suite - a leftover Word poisons whatever runs next.'
    }
}

Write-Host ''
if ($script:Failures -eq 0) {
    Write-Host "PASS  $($script:Checks) checks, 0 failures" -ForegroundColor Green
} else {
    Write-Host "FAIL  $($script:Failures) of $($script:Checks) checks failed" -ForegroundColor Red
}
Write-Host "Add-in log: $env:LOCALAPPDATA\WordTab\wordtab.log (lines starting 'stack')" -ForegroundColor Gray
exit $script:Failures
