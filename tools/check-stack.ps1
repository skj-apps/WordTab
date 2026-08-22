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

# Word's own report of one window's document, through that window's own object model. $null when Word
# will not answer, which must never be folded in with "Word answered and the document is clean". Used
# by the keyboard sections to answer "did a tab character get inserted" without reading pixels.
function Get-Saved($frame) {
    $om = [WordLayout]::NativeOm($frame)
    if ($null -eq $om) { return $null }
    try { return [bool]$om.Document.Saved } catch { return $null }
}

# Discard a change this suite made itself, to a file it authored itself, in %TEMP%.
#
# `Document.Saved` is settable and the standing rule in src\native is that the add-in only ever READS
# it - that rule protects the USER'S documents from the ADD-IN, which must never tell Word that
# something a person typed has been dealt with. A suite marking its own scratch .rtf saved is a
# different act, and the alternative is worse: the TabKeys=0 section deliberately provokes an edit,
# and Close-AllWord correctly refuses to answer a question it did not raise, so without this the run
# ends with a modal save prompt sitting on the user's screen. Same reasoning as probe-keyboard.ps1's
# Close-Word, which is the other place in this project that does it. FALSE when Word would not
# answer, which the caller must treat as "the fixture is not in a known state", never as success.
function Set-Saved($frame) {
    $om = [WordLayout]::NativeOm($frame)
    if ($null -eq $om) { return $false }
    try { $om.Document.Saved = $true; return $true } catch { return $false }
}

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
    #
    # This block used to live here and ONLY here, which is why the 38px Protected View bug reported
    # itself three steps from its cause twice: this suite has no mixed-chrome fixture, and the two
    # that do never ran the check. It is now Get-StripPlacement in tools\WordTabHarness.ps1 and
    # check-title and check-dot run it against a real Protected View window.
    $place = Get-StripPlacement (@($parts) | ForEach-Object { $_.Frame })
    Assert ($place.Measured -eq $parts.Count) "$label - all $($parts.Count) windows could be measured for strip placement (measured $($place.Measured))"
    Assert $place.Ok ("$label - every strip sits between the chrome and the document" + $place.Text)

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

# ---- Ctrl+Tab between documents ------------------------------------------------------------------
#
# The keyboard route to the row. Word gives the keyboard to _WwG, which this add-in does not
# subclass and could not usefully subclass - a WM_KEYDOWN goes to the focus window and nowhere else -
# so the chord is served by a WH_GETMESSAGE hook on Word's UI thread which rewrites it to WM_NULL and
# posts the switch to the coordinator. The measurement behind all of that is RESULT-keyboard.md.
#
# This rides on the click loop above rather than opening a fixture of its own, and that is not only
# economy: $activated IS the tab order. It was established by clicking each computed slot and reading
# which window came forward, which is the only way the tab order can be known from outside Word.
#
# THE ASSERTION THAT MATTERS IS THE DIRECTION, NOT "SOMETHING SWITCHED". Word orders its own window
# list most-recently-used; the tab row is ordered by position. The two agree exactly until the user
# switches documents once - which is the whole reason this feature could not be built by forwarding
# Word's Ctrl+F6, and the reason a test asserting only "a different document came forward" would pass
# against an implementation that is wrong in the way this one was specifically designed not to be.
#
# The clicks above visited the tabs left to right, so the run arrives here on the LAST tab with the
# MRU order running right to left - and there the two models disagree about every press below.

Write-Step 'Ctrl+Tab between documents'

if (($landed -ne $count) -or ($distinct.Count -ne $count)) {
    Write-Note 'the clicks above did not establish the tab order - skipping the keyboard checks rather than measuring against a guess'
} else {
    $order = @($activated)
    $fg    = [WordLayout]::GetForeground()
    $at    = -1
    for ($i = 0; $i -lt $order.Count; $i++) { if ($order[$i] -eq $fg) { $at = $i } }
    Assert ($at -ge 0) "the run is sitting on a tab whose position is known before the first chord (tab $at of $count)"

    # A precondition reported as a check. If Word has the keyboard anywhere but the document pane
    # then the chord is not being pressed where a user would press it, and "the hook works" would be
    # a claim about the ribbon. The hook itself does not care - it roots msg->hwnd to the frame - but
    # the *cost* being avoided is a tab character in the document, and that needs a caret.
    $gui = [WordLayout]::ThreadGui($fg)
    $focusClass = if ($gui.hwndFocus -ne [IntPtr]::Zero) { [WordLayout]::ClassOf($gui.hwndFocus) } else { '(none)' }
    Write-Note ("keyboard focus 0x{0:X} ({1})" -f [int64]$gui.hwndFocus, $focusClass)
    Assert ($focusClass -eq '_WwG') "Word has the keyboard in the document pane, where a user would press the chord (focus is $focusClass)"

    # The baseline for "the chord never reached the document". Recorded per window, because the
    # resize further up dirties a document all by itself - Word treats the view as part of the
    # document - so the honest test is that nothing went from saved to modified, not that everything
    # is clean. The count of genuinely clean documents is asserted too: without one, that test cannot
    # fail and would be passing for the wrong reason.
    $savedBefore = @{}
    foreach ($f in $order) { $savedBefore[[int64]$f] = Get-Saved $f }
    $clean = @($savedBefore.Values | Where-Object { $_ -eq $true }).Count
    Write-Note ("documents reporting saved=True before the chords: $clean of $count")

    if ($at -ge 0) {
        # +1, +1, -1, -1 from the last tab is 0, 1, 0, last - which crosses the right-hand end on the
        # first press and the left-hand end on the last, so both wraps are driven without a step that
        # exists only to set them up.
        $steps = @(
            @{ Shift = $false; Delta =  1 },
            @{ Shift = $false; Delta =  1 },
            @{ Shift = $true;  Delta = -1 },
            @{ Shift = $true;  Delta = -1 }
        )

        # Word's own answer, tracked properly rather than guessed at.
        #
        # The first version of this computed it as "the tab to the left of the one we are on", which
        # is right for the FIRST press and wrong for every one after it - Word's list reorders on
        # every switch, and two of the four steps then printed "and not to Word's MRU tab 0" about a
        # press whose correct answer was also tab 0. A comparison against a number that is only
        # sometimes the other model's answer is not a discriminator, and it read like one.
        #
        # The list is knowable exactly because this suite caused every activation in it: the clicks
        # above visited tab 0, then 1, then 2, so most-recently-used is the reverse of that. Word's
        # Ctrl+F6 goes to the next entry after the current one, and whatever is switched to moves to
        # the front. Only the forward direction was measured (RESULT-keyboard.md section 3); Word's
        # reverse chord is Ctrl+Shift+F6 and its ordering was not, so this is stated as the forward
        # answer and the discriminating steps are counted rather than assumed.
        $mruList = @()
        for ($i = $count - 1; $i -ge 0; $i--) { $mruList += $i }

        $expect  = $at
        $wrapped = 0
        $split   = 0
        Set-LogMark
        foreach ($step in $steps) {
            $from    = $expect
            $expect  = ((($expect + $step.Delta) % $count) + $count) % $count
            $chord   = if ($step.Shift) { 'Ctrl+Shift+Tab' } else { 'Ctrl+Tab' }
            $crosses = ([Math]::Abs($expect - $from) -ne 1)
            if ($crosses) { $wrapped++ }

            $mru = $mruList[1]
            $differs = ($mru -ne $expect)
            if ($differs) { $split++ }
            Write-Note ("$chord from tab $from - the row says tab $expect, Word's MRU order [$($mruList -join ',')] says tab $mru" +
                        $(if ($differs) { ' - they disagree, so this press tells the two models apart' } else { ' - they agree here' }))

            if (-not (Invoke-ConfirmedKey -Vk 0x09 -Ctrl -Shift:$step.Shift -What $chord)) {
                Assert $false "$chord could be delivered to Word"
                continue
            }
            Start-Sleep -Milliseconds 900

            $now   = [WordLayout]::GetForeground()
            $title = [WordLayout]::TitleOf($now)
            $landedAt = -1
            for ($i = 0; $i -lt $order.Count; $i++) { if ($order[$i] -eq $now) { $landedAt = $i } }
            Write-Note ("  -> foreground 0x{0:X} tab {1} `"{2}`"" -f [int64]$now, $landedAt, $title)

            Assert ($landedAt -eq $expect) `
                ("$chord from tab $from moved to tab $expect" + $(if ($crosses) { ' (wrapping)' } else { '' }) +
                 $(if ($differs) { ", not to Word's MRU tab $mru" } else { '' }) + " (landed on $landedAt)")

            # Whatever came forward goes to the front of Word's list, whether we put it there or Word
            # did. Driven off what actually landed rather than off what was expected, so a failed
            # press does not leave the model describing a machine it has diverged from.
            if ($landedAt -ge 0) { $mruList = @($landedAt) + @($mruList | Where-Object { $_ -ne $landedAt }) }
        }

        Assert ($wrapped -eq 2) "the four chords crossed both ends of the row ($wrapped of 2 wraps driven)"
        Assert ($split -ge 2) `
            ("the row order and Word's MRU order disagreed on $split of $($steps.Count) presses, so this " +
             "section can tell a tab-order implementation from one that forwards Word's own command")

        # The add-in's own view, which is the second oracle. The foreground window says which document
        # Word brought forward; the log says which tab WordTab decided on and which one it started
        # from. The two disagreeing would mean the row and the stack had come apart.
        $keyLines = @(Get-LogSince 'keys  ctrl')
        foreach ($line in $keyLines) { Write-Note "  log: $($line.Trim())" }
        Assert ($keyLines.Count -eq $steps.Count) `
            "the add-in logged one switch per chord ($($keyLines.Count) of $($steps.Count))"
        Assert (@($keyLines | Where-Object { $_ -like '*nowhere*' }).Count -eq 0) `
            'no chord found the row empty of anywhere to go'
        Assert (@($keyLines | Where-Object { $_ -like '*ctrl+shift+tab*' }).Count -eq 2) `
            "two of the four were logged as the reverse chord ($(@($keyLines | Where-Object { $_ -like '*ctrl+shift+tab*' }).Count))"

        # And the cost that was avoided. Word owns Ctrl+Tab and types a literal tab with it; the hook
        # has to swallow the chord rather than watch it go past. This is the check that says it did.
        $typed = @()
        foreach ($f in $order) {
            $after = Get-Saved $f
            if (($savedBefore[[int64]$f] -eq $true) -and ($after -eq $false)) { $typed += $f }
        }
        Assert ($clean -gt 0) "at least one document was unmodified before the chords, so the check below can fail ($clean)"
        Assert ($typed.Count -eq 0) "no chord put a tab character into a document ($($typed.Count) went from saved to modified)"
    }
}

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

# ---- the tab that is highlighted is the document on screen ---------------------------------------
#
# The report this exists for, their words: "when opening a second doc it switch to that but initial
# tab still highlighted". Their log has it three times in one day - `joined, snapped to 0x...` with
# no `active ->` after it - and it stays wrong until the user clicks something.
#
# The cause is not the opening. It is that the row read "which window is the user looking at" off
# the KEYBOARD: GetForegroundWindow, plus WM_ACTIVATE. Word puts a newly opened document's window in
# front without always giving it focus, and then those two answers differ. The stack holds every
# window at one rectangle, so the document on screen is the front-most one and nothing else - focus
# was only ever a proxy for that, and this is the case where the proxy is wrong.
#
# It cannot be provoked by opening a document HERE: on this rig Word raises the new window and hands
# it the keyboard in the same breath, so both answers agree and the check could never fail. What is
# reproduced instead is the mechanism - a window put in front with SWP_NOACTIVATE, which is exactly
# the state their Word leaves behind - and it is asserted from outside the add-in, through the one
# window it presents to Alt+Tab. That is g_active read back through the product rather than through
# its log.
if ((Get-FrameCount) -ge 2) {
    Write-Step 'A window raised without focus becomes the selected tab'

    $before   = @(Get-Frames)
    $showing  = @($before | Where-Object { -not [WordLayout]::IsToolWindow($_) })
    $hidden   = @($before | Where-Object { [WordLayout]::IsToolWindow($_) })

    if ($showing.Count -ne 1 -or $hidden.Count -lt 1) {
        Assert $false "the row presents exactly one window before the raise (presented $($showing.Count) of $($before.Count))"
    } else {
        $wasActive = $showing[0]
        $raised    = $hidden[0]
        $focusWas  = [WordLayout]::GetForeground()

        Set-LogMark
        [WordLayout]::RaiseWithoutFocus($raised)

        # Long enough for two janitor ticks, and a fixed wait rather than a poll: how long the row
        # takes to notice is part of the claim, and waiting for the answer would make this unable
        # to fail.
        Start-Sleep -Milliseconds 1500

        $focusNow = [WordLayout]::GetForeground()
        Write-Note ("raised 0x{0:X} in front of 0x{1:X}; foreground 0x{2:X} -> 0x{3:X}" -f `
                    [int64]$raised, [int64]$wasActive, [int64]$focusWas, [int64]$focusNow)
        foreach ($line in @(Get-LogSince 'active ->')) { Write-Note "  log: $($line.Trim())" }

        # The precondition that makes the rest mean anything: the keyboard did NOT move. Without
        # this, a pass could be the old focus rule answering a question it was never asked.
        Assert ($focusNow -eq $focusWas) 'the raise left the keyboard where it was'
        Assert (-not [WordLayout]::IsToolWindow($raised)) 'the window in front is the one presented to Alt+Tab'
        Assert ([WordLayout]::IsToolWindow($wasActive)) 'the window behind it gave the presentation up'
        Assert (@(Get-LogSince 'active ->').Count -ge 1) 'the row said the active tab changed'

        # Put the row back the way the suite found it, through the product's own path rather than by
        # raising the old window the same way: what follows asserts on the active window, and leaving
        # it decided by a z-order the user never touched would be this check bleeding into the next.
        [WordLayout]::Focus($wasActive) | Out-Null
        Start-Sleep -Milliseconds 900
    }
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

    # One document, and the chord still belongs to the tab row. This is the product decision the
    # keyboard slice made explicitly rather than by accident: the chord is swallowed on the basis of
    # "this key went to a window that is a tab in our row", NOT "there is somewhere to go". The
    # alternative types a tab character on a one-document Word and switches on a two-document one,
    # which makes what the key does depend on how many documents happen to be open.
    #
    # It costs nothing to check here: the window is already up and it is our own fixture. Marking it
    # saved first is what makes the assertion able to fail - the resize earlier in this run dirties a
    # document on its own, and "still modified" would then be true whatever the hook did. See the
    # note on the TabKeys=0 section below for why a suite may set Saved on a file it authored.
    if (Set-Saved $last[0]) {
        Set-LogMark
        if (Invoke-ConfirmedKey -Vk 0x09 -Ctrl -What 'Ctrl+Tab with one document open') {
            Start-Sleep -Milliseconds 700
            $lone = @(Get-LogSince 'keys  ctrl')
            foreach ($line in $lone) { Write-Note "  log: $($line.Trim())" }
            Assert (@($lone | Where-Object { $_ -like '*nowhere*' }).Count -eq 1) `
                "the chord was swallowed and reported having nowhere to go ($($lone.Count) line(s) logged)"
            Assert ((Get-Saved $last[0]) -eq $true) `
                'and no tab character reached the one open document'
        } else {
            Assert $false 'Ctrl+Tab could be delivered to the last window'
        }
    } else {
        Write-Note 'Word would not answer for the last document - skipping the one-document chord check'
    }
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

# ---- TabKeys=0 gives the chord back to Word -------------------------------------------------------
#
# The escape hatch, driven rather than assumed. Word owns Ctrl+Tab, and inside a table it is the only
# way to type a literal tab into a cell - measured, RESULT-keyboard.md section 4 - so this switch is
# the answer for somebody who works in tables all day, and a dead one leaves them nothing.
#
# It is checked POSITIVELY: the assertion is that the keystroke reaches Word and types, not that
# nothing happened. "Nothing happened" is also what a keystroke that never landed produces, and this
# project has had three negative tests pass for exactly that reason. Pairs with the one-document
# check above, which is the same chord in the same shape with the switch on and nothing typed.
#
# It costs one Word restart, because the switch is read once at FramesStart. That is the price of a
# switch that turns a hook OFF rather than making it inert, and turning it off properly is the point:
# with TabKeys=0 there is no WH_GETMESSAGE hook on Word's UI thread at all.

if (-not $KeepOpen) {
    Write-Step 'TabKeys=0 gives Ctrl+Tab back to Word'

    $switchKey  = 'HKCU:\Software\WordTab'
    $hadTabKeys = $false
    $oldTabKeys = $null
    if (Test-Path $switchKey) {
        $existing = Get-ItemProperty -Path $switchKey -Name 'TabKeys' -ErrorAction SilentlyContinue
        if ($existing -and ($existing.PSObject.Properties.Name -contains 'TabKeys')) {
            $hadTabKeys = $true
            $oldTabKeys = $existing.TabKeys
        }
    }

    try {
        $end = Close-AllWord
        if (-not $end.Closed) {
            Write-Note ("Word would not close ({0}) - skipping the TabKeys=0 check" -f $end.Reason)
        } else {
            if (-not (Test-Path $switchKey)) { New-Item -Path $switchKey | Out-Null }
            Set-ItemProperty -Path $switchKey -Name 'TabKeys' -Value 0 -Type DWord

            $path = Join-Path $scratch 'wordtab-stack-keys.rtf'
            "{\rtf1\ansi WordTab TabKeys=0 check.\par}" | Set-Content -Path $path -Encoding Ascii
            Set-LogMark
            Start-Process -FilePath 'winword.exe' -ArgumentList "`"$path`""
            if (-not (Wait-WordReady 1 45)) { Write-Note 'the document did not arrive with a strip on it within 45s' }
            Start-Sleep -Seconds 2

            # The add-in's own account of the switch, reported from what SetWindowsHookEx returned
            # rather than from the variable that asked for it - which is how TabThemeSample stayed
            # dead for two slices while the log said it was on.
            $start = @(Get-LogSince 'FramesStart  uiThread')
            foreach ($line in $start) { Write-Note "  log: $($line.Trim())" }
            Assert (@($start | Where-Object { $_ -like '*msgHook=off*' }).Count -ge 1) `
                'the add-in started with no keyboard hook installed'

            $keyFrames = @(Get-Frames)
            if ($keyFrames.Count -ne 1) {
                Write-Note "expected one window for the TabKeys=0 check, found $($keyFrames.Count) - skipping"
            } elseif (-not (Set-Saved $keyFrames[0])) {
                Write-Note 'Word would not answer for the fixture - skipping the TabKeys=0 keystroke'
            } else {
                Assert ((Get-Saved $keyFrames[0]) -eq $true) 'the fixture starts unmodified, so a typed tab will show'

                Set-LogMark
                if (Invoke-ConfirmedKey -Vk 0x09 -Ctrl -What 'Ctrl+Tab with TabKeys=0') {
                    Start-Sleep -Milliseconds 900
                    Assert ((Get-Saved $keyFrames[0]) -eq $false) `
                        'Ctrl+Tab reached Word and put a tab character in the document - the switch really is off'
                    Assert (@(Get-LogSince 'keys  ctrl').Count -eq 0) `
                        'and the add-in did not see the chord at all'
                    Set-Saved $keyFrames[0] | Out-Null
                } else {
                    Assert $false 'Ctrl+Tab could be delivered with TabKeys=0'
                }
            }
        }
    } finally {
        # Put the switch back exactly as it was found, including absent. A suite that leaves TabKeys=0
        # behind would disable the feature on the dev rig and every suite after it would agree that
        # nothing is wrong.
        if ($hadTabKeys) {
            Set-ItemProperty -Path $switchKey -Name 'TabKeys' -Value $oldTabKeys -Type DWord
        } elseif (Test-Path $switchKey) {
            Remove-ItemProperty -Path $switchKey -Name 'TabKeys' -ErrorAction SilentlyContinue
        }
        $restored = Get-ItemProperty -Path $switchKey -Name 'TabKeys' -ErrorAction SilentlyContinue
        $nowIs = if ($restored -and ($restored.PSObject.Properties.Name -contains 'TabKeys')) { $restored.TabKeys } else { '(absent)' }
        Write-Note "TabKeys restored to $nowIs"
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
