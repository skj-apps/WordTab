<#
.SYNOPSIS
  The parts of the check suites that must not disagree with each other: confirmed input, and
  "is Word asking the user something".

.DESCRIPTION
  Every suite grew its own copy of these, and the copies drifted. Measured across the eleven suites
  before this file existed:

    - 134 injected inputs, of which 4 confirmed they landed. The other 130 asked the desktop for a
      click or a keystroke and assumed it arrived
    - THREE different versions of Get-WordDialog, and none of them was right. check-dot excluded
      `Net UI Tool Window` but not `#32768`, so an open context menu read as a save prompt.
      check-menu and check-title excluded `#32768` but not `Net UI Tool Window` - which is the exact
      window that cost a whole battery run
    - check-menu's Set-WordForeground would minimise ANY window that would not give the foreground
      up, including `Progman`. That is the desktop, and minimising it is Show Desktop, which takes
      Word down with everything else. The same file states that rule in a comment 260 lines further
      down and obeys it there

  So the rule this file exists for: **a primitive that decides whether the harness is telling the
  truth lives in exactly one place.** Suites keep their own fixtures, their own assertions and their
  own Get-Frames; they do not keep their own idea of what a dialog is.

  Dot-source it AFTER Add-Type has defined [WordLayout], and BEFORE the suite's own functions, so a
  suite that still needs a local override gets one by defining it later:

      $source = Get-Content -Raw -Path (Join-Path $PSScriptRoot 'WordLayout.cs')
      Add-Type -TypeDefinition $source -Language CSharp -ReferencedAssemblies @(...)
      [WordLayout]::MakeDpiAware() | Out-Null
      . (Join-Path $PSScriptRoot 'WordTabHarness.ps1')

  Nothing in here asserts. These are measurements and actions; the suite decides what a failure
  means. Every confirming function returns $true/$false so the suite can Assert the arrival as a
  check in its own right - which is the point. A hover that never happened and a hover that is not
  drawn produce byte-identical evidence, and only one of them is a bug in the product.

.NOTES
  Depends on [WordLayout] only. It deliberately does NOT call back into suite-local helpers
  (Write-Note, Get-Frames, Set-WordForeground): those differ between suites, and a shared file whose
  behaviour depends on which suite loaded it is not shared.
#>

# ---- diagnosis --------------------------------------------------------------------------------
#
# Its own note writer rather than the suites' Write-Note. Same shape on screen; no late binding into
# a caller that may not have defined it yet.

function Write-HarnessNote($text) { Write-Host "    $text" -ForegroundColor DarkGray }

# ---- Word's windows ---------------------------------------------------------------------------

# EVERY caller of these re-wraps in @() before touching .Count, and so does every function below.
# A PowerShell function that returns an empty array returns *nothing*, so `(Get-WordPidList).Count`
# on no Word at all is a hard error under Set-StrictMode rather than 0 - and "no Word at all" is
# exactly the state Wait-WordGone and Set-WordForeground exist to handle. This trap has cost this
# project two runs already, under the name `(Get-Frames).Count`.
function Get-WordPidList {
    return @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
}

function Get-WordPidTally { return @(Get-WordPidList).Count }

# The document frames, the same filter [WordLayout]::Frames applies: visible, not minimised, bigger
# than 200x200 OpusApp windows. Suites keep their own Get-Frames - this is the copy the primitives
# below use, so their behaviour cannot depend on which suite dot-sourced them.
function Get-WordFrameList {
    $found = @()
    foreach ($id in Get-WordPidList) { $found += [WordLayout]::Frames($id) }
    return @($found)
}

# Always ask for the count through a re-wrapped @(): a PowerShell function that returns an empty
# array returns *nothing*, and .Count on nothing is a hard error under Set-StrictMode rather than 0.
function Get-WordFrameTally { return @(Get-WordFrameList).Count }

<#
  What a top-level window of Word's actually is. Four kinds, and the whole battery-losing bug was
  two of them being confused for each other.

    frame     an OpusApp document window. Never a question.

    menu      class #32768, the system popup menu - our own tab context menu is one. Big enough and
              visible enough to pass a naive size filter, and check-dot's "narrowed" version let it
              through.

    chrome    Office's floating UI: `Net UI Tool Window` hosts galleries, task-pane popouts and
              notification toasts. MEASURED at 622x298 while check-title was closing seven windows,
              where it was read as "Word is asking something", so the suite refused to close Word and
              left it running. check-dot then died on the same window in its own cleanup, and
              soak-stack opened six documents and measured ELEVEN. Three suites red, one cause.
              `MsoCommandBar` and the tooltip/shadow classes are chrome for the same reason.

    question  everything else visible and big enough. Measured by tools\probe-dialogs.ps1: the save
              prompt is `NUIDialog` 920x713, and the earlier battery measured a second NUIDialog at
              920x641. Save As raised with F12 is something else entirely - `#32770` 1441x960, the
              classic shell common dialog, titled "Save As".

  Deliberately a DENYLIST, not an allowlist of measured question classes, and the probe proved that
  choice rather than argued it. An allowlist would have been written as "NUIDialog is a question",
  because that is the only prompt class this project had ever seen - and it would have been blind to
  the `#32770` Save As, which is a real modal that really does stop Word.

  The two mistakes are not symmetrical either: calling chrome a question wastes a run, but calling a
  question chrome means a suite drives clicks and Escapes past a modal it cannot see, and can walk
  away leaving a real save prompt up on the user's screen. An unmeasured dialog class must land on
  the side that stops the run.

  And NOT identified by its text, which the probe also settled: the save prompt has no window title
  and NO child control with any text in it at all. There is nothing to match on even if matching on
  English were acceptable.

  Also deliberately NOT identified by the text in it. Word's prompts carry no window title and their
  static children's text is localised; the tab-name slice already learned not to parse Word's
  English. The text is reported for a human to read, never matched on.
#>
#
# Everything in this list was seen on screen by tools\probe-dialogs.ps1, or measured in the battery
# run it cost to learn the first one. Re-run that probe before changing any of it.
$script:WordChromeClasses = @(
    'Net UI Tool Window',        # galleries, task-pane popouts, toasts - measured at 622x298
    'SysShadow',                 # measured at 416x339 beside our own 411x334 context menu, so
                                 # every previous version of this test called a menu's SHADOW a
                                 # question: bigger than 200x100, and not OpusApp or #32768
    'MSO_KEYTIP_WINDOW_CLASS',   # the letters Word floats over the ribbon after Alt - measured at
                                 # 38x36 and 30x36, seven of them at once, while Backstage opened
    'MsoCommandBar',             # floating toolbars
    'MsoCommandBarPopup',
    'tooltips_class32',
    'MSO_BORDEREFFECT_WINDOW'    # the drop shadow Word draws around its own popups
)

function Get-WordWindowKind($class, $width, $height) {
    if ($class -eq 'OpusApp')  { return 'frame' }
    if ($class -eq '#32768')   { return 'menu' }
    if ($script:WordChromeClasses -contains $class) { return 'chrome' }

    # A floor, as a second filter and not as the first one. Both measured prompts are 920 wide, so
    # this rejects only slivers - Word's zero-size hidden helpers and the 1x1 windows it parks
    # off-screen. Kept low on purpose: a question that is smaller than expected must still stop a run.
    if ($width -lt 150 -or $height -lt 60) { return 'chrome' }

    return 'question'
}

# Every visible top-level window of every Word process, classified and ready to print. This is the
# thing to dump when a run goes wrong: "a dialog was up" is not a measurement, "class=Net UI Tool
# Window 622x298 |...|, kind=chrome" is.
function Get-WordWindows {
    $found = @()
    foreach ($id in @(Get-WordPidList)) {
        foreach ($w in [WordLayout]::TopLevel($id)) {
            if (-not $w.Visible) { continue }
            $width  = $w.Right - $w.Left
            $height = $w.Bottom - $w.Top
            # Left/Top/Right/Bottom and Visible are carried deliberately, so this object is a strict
            # SUPERSET of the WordLayout.Child the suites' own Get-WordDialog used to hand back.
            # Without them a caller that did `Format-Rect $dialog` gets "the property 'Left' cannot be
            # found on this object" - which is what happened, in the one branch of check-menu that
            # only runs when Word actually asks where to save a new document.
            $found += [pscustomobject]@{
                Hwnd = $w.Hwnd; Class = $w.Class; Title = $w.Title; Pid = $id; Visible = $true
                Left = $w.Left; Top = $w.Top; Right = $w.Right; Bottom = $w.Bottom
                Width = $width; Height = $height
                Kind = (Get-WordWindowKind $w.Class $width $height)
            }
        }
    }
    return @($found)
}

# Word asking the user something, or $null. The one answer the whole harness agrees on.
function Get-WordDialog {
    foreach ($w in @(Get-WordWindows)) { if ($w.Kind -eq 'question') { return $w } }
    return $null
}

# A popup menu of Word's, or IntPtr.Zero. Taking the foreground dismisses one, so several things
# below have to ask first.
function Get-WordMenu {
    foreach ($id in @(Get-WordPidList)) {
        $w = [WordLayout]::PopupMenuWindow($id)
        if ($w -ne [IntPtr]::Zero) { return $w }
    }
    return [IntPtr]::Zero
}

function Format-WordWindow($w) {
    if (-not $w) { return '(nothing)' }
    return "class={0} {1}x{2} kind={3} |{4}|" -f $w.Class, $w.Width, $w.Height, $w.Kind, $w.Title
}

# Kept under the old name too: the suites print dialogs with this.
function Format-Dialog($d) { return Format-WordWindow $d }

# Everything Word has on screen that is not a document window, for a failure message. Says what was
# actually there rather than that something was.
function Format-WordChrome {
    $others = @(@(Get-WordWindows) | Where-Object { $_.Kind -ne 'frame' })
    if ($others.Count -eq 0) { return '(no dialogs, menus or floating chrome)' }
    return (@($others | ForEach-Object { Format-WordWindow $_ }) -join '; ')
}

<#
  Word asking the user something, as opposed to Word's own chrome going past on its way somewhere.
  Confirmed twice, a second and a half apart, because a single sample reads shutdown chrome as a
  prompt. Both samples must be non-null; the second is the answer, since the window may have been
  replaced between them.
#>
function Get-WordSavePrompt {
    $first = Get-WordDialog
    if (-not $first) { return $null }
    Start-Sleep -Milliseconds 1500
    return Get-WordDialog
}

# ---- bounded waits ----------------------------------------------------------------------------
#
# A fixed Start-Sleep is a guess about someone else's machine. Where the thing being waited for can
# be observed, poll for it: faster when it is quick, still correct when it is slow, and it says so
# on timeout instead of sailing past. The eleven suites sleep 8.7 minutes of literal Start-Sleep per
# battery, measured, and most of it is spent waiting for Word to start.
#
# THE RULE THAT KEEPS THIS HONEST, and it is not "never poll for what you assert":
#
#   Convert a sleep that is waiting for WORD to do something - start, open a document, shut down.
#   That is the harness getting into position, and how long it takes says nothing about the product.
#
#   Do NOT convert a sleep that sits between an action under test and the assertion about it. That
#   gap is part of the claim. `Start-Sleep 1500` then "every window went down with it" asserts the
#   stack minimises *promptly*; `Wait-Until { all minimised } 10` then the same assertion still goes
#   red if it never happens, but now stays green if the add-in regresses from instant to nine
#   seconds. Nothing reports that. A wait in that position trades a timing check for an eventually
#   check, silently, and the suite still prints PASS.
#
# Every wait here is bounded and returns $false rather than throwing, so the assertion after it runs
# either way and reports what it found.

function Wait-Until([scriptblock]$condition, $seconds = 20, $pollMs = 200) {
    $deadline = (Get-Date).AddSeconds($seconds)
    while ((Get-Date) -lt $deadline) {
        if (& $condition) { return $true }
        Start-Sleep -Milliseconds $pollMs
    }
    return $false
}

# Word up with $expected document frames, settled. The second read after a pause is what tells a
# count on its way past from a count that has arrived: closing four documents goes 4-3-2 and a
# single sample can catch 2 on the way to 1.
function Wait-WordFrames($expected, $seconds = 25, $settleMs = 800) {
    $deadline = (Get-Date).AddSeconds($seconds)
    while ((Get-Date) -lt $deadline) {
        if ((Get-WordFrameTally) -eq $expected) {
            Start-Sleep -Milliseconds $settleMs
            if ((Get-WordFrameTally) -eq $expected) { return $true }
        }
        Start-Sleep -Milliseconds 300
    }
    return $false
}

# A frame with a strip on it, which is the real "Word is ready" - a frame exists before the add-in
# has drawn anything, and a suite that measures the strip in that gap measures nothing.
function Wait-WordReady($frames = 1, $seconds = 40) {
    return Wait-Until {
        $list = @(Get-WordFrameList)
        if ($list.Count -lt $frames) { return $false }
        foreach ($f in $list) {
            $strip = @([WordLayout]::Children($f) | Where-Object { $_.Class -eq 'WordTabStrip' })
            if ($strip.Count -eq 0) { return $false }
        }
        return $true
    } $seconds
}

function Wait-WordGone($seconds = 25) {
    return Wait-Until { (Get-WordPidTally) -eq 0 } $seconds 400
}

# Returns the dialog when waiting for one to appear and it did, and $null when it did not; the other
# way round when waiting for one to go. Either way it is a fresh read at the end, never a cached one -
# the caller is about to print it.
function Wait-WordDialog($present, $seconds = 10) {
    Wait-Until {
        $d = Get-WordDialog
        if ($present) { return ($null -ne $d) }
        return ($null -eq $d)
    } $seconds 300 | Out-Null
    return (Get-WordDialog)
}

# ---- where a strip is allowed to sit ----------------------------------------------------------

<#
  Every strip sits between the chrome above it and the document below it, checked per window in
  absolute terms.

  **Consistency is not correctness, and that is the whole reason this exists.** check-stack already
  asserts that every window in the stack agrees on one document-frame rectangle - and every window
  agreeing on a WRONG layout passes that. It has happened twice. Once, closing a document put the
  strip at y=0 on top of the ribbon in all of them, identically. And once, a Protected View window's
  chrome - which is 38px shorter than a normal window's, measured - was broadcast onto every other
  window in the stack, so each one carved its band 38px too high, inside the ribbon's NetUIHWND.

  **This assertion lived in check-stack only, and check-stack is the one suite with no mixed-chrome
  fixture.** The two suites that DO open a Protected View document, check-title and check-dot, never
  ran it. So the 38px bug reported itself as "the + click did nothing" - three steps from the cause,
  twice, in two different slices. Hoisting it here is the whole point: the suites that can reach the
  bug are now the suites that test for it.

  Returns a result object rather than asserting, like everything else in this file - and an OBJECT
  rather than an array of faults, deliberately: a PowerShell function that returns an empty array
  returns *nothing*, so `(Get-...Faults $f).Count` would be a hard error under StrictMode at exactly
  the moment there is nothing wrong. That trap has cost this project two runs under two other names.
  A pscustomobject is a scalar and survives the pipeline whatever is in it.

  `Measured` is part of the answer for the same reason the previous slice made a test say when it did
  not measure anything: a window with no strip or no document frame is skipped, not failed - whether
  it should HAVE a strip is check-stack's separate question - so `Ok` on its own cannot tell "every
  strip is placed correctly" apart from "there were no strips". Assert both.
#>
function Get-StripPlacement($frames) {
    $faults   = @()
    $measured = 0

    foreach ($frame in @($frames)) {
        # **An empty pipeline arrives here as a literal $null, and @($null) iterates ONCE.** Measured,
        # because the obvious test gets it wrong: in the CALLER's scope an empty pipeline is
        # AutomationNull and @(it) is empty, but parameter binding converts that to $null on the way
        # in, so `@($frames)` inside the function is a one-element array holding $null and
        # [WordLayout]::Children($null) throws "Cannot convert null to type System.IntPtr".
        #
        # check-stack calls this as `Get-StripPlacement (@($parts) | ForEach-Object { $_.Frame })`,
        # and $parts is empty exactly when a restore did not take - which is the moment the assertion
        # above it exists to report. Without this guard the suite would die with a conversion error
        # instead of reporting that failure. Fourth shape of this project's empty-collection trap.
        if ($null -eq $frame -or $frame -eq [IntPtr]::Zero) { continue }

        $kids  = [WordLayout]::Children($frame)
        $strip = @($kids | Where-Object { $_.Class -eq 'WordTabStrip' }) | Select-Object -First 1
        $wwf   = @($kids | Where-Object { $_.Class -eq '_WwF' })         | Select-Object -First 1
        if (-not $strip -or -not $wwf) { continue }
        $measured++

        # What is directly above the strip: the lowest visible child that ends at or above our top
        # edge and is most of the window wide. In a normal window that is the ribbon's NetUIHWND
        # chain; in a Protected View window it is the message bar's. Either is correct - what is
        # never correct is a gap or an overlap between it and us.
        $client = [WordLayout]::ClientOf($frame)
        $above = $kids |
                 Where-Object { $_.Visible -and $_.Hwnd -ne $strip.Hwnd -and $_.Hwnd -ne $wwf.Hwnd -and
                                $_.Bottom -le $strip.Top -and $_.Width -gt ($client.Right * 0.6) } |
                 Sort-Object Bottom -Descending | Select-Object -First 1

        $tag = "0x$('{0:X}' -f [int64]$frame)"
        if ($strip.Bottom -ne $wwf.Top) {
            $faults += "$tag strip bottom $($strip.Bottom) != document top $($wwf.Top)"
        } elseif (-not $above) {
            # **This is the shape the 38px bug makes, so it must not be described as something else.**
            # A strip placed 38px too high sits INSIDE the ribbon's NetUIHWND: the whole chrome chain
            # reports 0..356 while the strip claims 318, so nothing ends at or above the strip's top
            # and $above is empty. "It is sitting at the top of the window" would be a wrong story
            # about the right failure. Name the numbers and the chrome that overlaps instead.
            $nearest = $kids |
                       Where-Object { $_.Visible -and $_.Hwnd -ne $strip.Hwnd -and $_.Hwnd -ne $wwf.Hwnd -and
                                      $_.Width -gt ($client.Right * 0.6) } |
                       Sort-Object Top | Select-Object -First 1
            $what = 'there is no full-width chrome in the window at all'
            if ($nearest) { $what = "nearest full-width chrome is $($nearest.Class) y $($nearest.Top)..$($nearest.Bottom)" }
            $faults += "$tag nothing ends at or above the strip's top ($($strip.Top)) - $what"
        } elseif ($above.Bottom -ne $strip.Top) {
            $faults += "$tag $($above.Class) ends at $($above.Bottom), strip starts at $($strip.Top)"
        }
    }

    return [pscustomobject]@{
        Ok       = ($faults.Count -eq 0)
        Faults   = @($faults)
        Count    = $faults.Count
        Measured = $measured
        # The message too, so the call sites cannot word the same failure differently.
        Text     = $(if ($faults.Count) { ': ' + ($faults -join '; ') } else { '' })
    }
}

# ---- the foreground ---------------------------------------------------------------------------

<#
  Put Word in front, and say so if it would not go.

  SetForegroundWindow is refused from a process that is not already foreground, so this is not
  optional politeness. A suite that drives real clicks at screen coordinates against a window that
  is not on top does not merely lose the input - the window that IS on top *receives the clicks*.
  A reading of the tab row has come back naming "Remote Desktop Connection", and later "Terminal".

  Three guards, each of which was a bug in at least one suite before this file existed:

  1. Never while a menu is up. Taking the foreground dismisses the very thing being measured.

  2. Never minimise `Progman` or `WorkerW`. That is the desktop, and minimising it is Show Desktop,
     which takes Word down with everything else. check-menu's copy had no such guard.

  3. Never minimise a window belonging to a Word process. If Word has a modal save prompt up, the
     prompt holds the foreground and is not an OpusApp - so "not one of our frames, minimise it"
     minimises Word's own question, leaving Word disabled behind an invisible modal. That is the
     "Word beeping with nothing on screen" failure, reached from the other direction.
#>
function Set-WordForeground($seconds = 4, [switch]$Quiet) {
    if ((Get-WordMenu) -ne [IntPtr]::Zero) { return [WordLayout]::GetForeground() }

    $frames = @(Get-WordFrameList)
    if ($frames.Count -eq 0) { return [IntPtr]::Zero }

    for ($round = 1; $round -le 3; $round++) {
        $now = [WordLayout]::GetForeground()
        if ($frames -contains $now) { return $now }

        if (-not $Quiet) {
            Write-HarnessNote ("the foreground was `"{0}`" ({1}) - taking it back" -f
                               [WordLayout]::TitleOf($now), [WordLayout]::ClassOf($now))
        }
        [WordLayout]::Focus($frames[0]) | Out-Null
        if (Wait-Until { @(Get-WordFrameList) -contains ([WordLayout]::GetForeground()) } $seconds 250) {
            return [WordLayout]::GetForeground()
        }

        # It would not give the foreground up. Last resort, and a deliberate one: minimising an
        # application is recoverable from the taskbar, and is far better than a green run that
        # measured the wrong window. See the three guards above for what is never minimised.
        $now = [WordLayout]::GetForeground()
        if ($now -eq [IntPtr]::Zero) { continue }
        $cls = [WordLayout]::ClassOf($now)
        if (@(Get-WordFrameList) -contains $now) { return $now }
        if ($cls -eq 'Progman' -or $cls -eq 'WorkerW') {
            if (-not $Quiet) { Write-HarnessNote 'the desktop has the foreground - not minimising that' }
            continue
        }
        if (@(Get-WordPidList) -contains ([WordLayout]::PidOf($now))) {
            if (-not $Quiet) {
                $what = @(Get-WordWindows | Where-Object { $_.Hwnd -eq $now })
                $desc = if ($what.Count -gt 0) { Format-WordWindow $what[0] } else { "class=$cls" }
                Write-HarnessNote ("Word itself has the foreground with {0} - leaving it alone" -f $desc)
            }
            return $now
        }

        if (-not $Quiet) {
            Write-HarnessNote ("minimising `"{0}`" ({1}) - it is sitting over Word and taking the input meant for it" -f
                               [WordLayout]::TitleOf($now), $cls)
        }
        [WordLayout]::Show($now, 6)      # SW_MINIMIZE
        Start-Sleep -Milliseconds 600
    }
    return [WordLayout]::GetForeground()
}

# Is one of Word's DOCUMENT windows the foreground window? The right question before a click at
# screen coordinates: anything on top of Word receives those clicks itself.
function Test-WordForeground {
    return (@(Get-WordFrameList) -contains ([WordLayout]::GetForeground()))
}

# Is Word in a position to receive a keystroke at all? True for a frame, and also for Word's own
# dialogs, Backstage and menus - all of which are Word, and all of which are legitimate keystroke
# targets. This is the right question for a key press; "is the foreground one of our frames" is not.
function Test-WordHasFocus {
    $now = [WordLayout]::GetForeground()
    if ($now -eq [IntPtr]::Zero) { return $false }
    return (@(Get-WordPidList) -contains ([WordLayout]::PidOf($now)))
}

# ---- confirmed input --------------------------------------------------------------------------
#
# AN INJECTED INPUT YOU DID NOT CONFIRM LANDED IS NOT AN INPUT.
#
# SendInput's absolute mouse move silently does not take when another process holds the foreground
# or a hand is on the mouse. Measured: a run asked for (258,540) and left the pointer at (1894,459),
# over the desktop. check-look failed two hover assertions *reproducibly* - "0 pixels changed",
# hovered tab sampling the well colour - against hover code that was working perfectly.

# Move the pointer, and prove it arrived.
function Set-Pointer($x, $y, $what) {
    for ($try = 1; $try -le 5; $try++) {
        [WordLayout]::MouseTo($x, $y)
        Start-Sleep -Milliseconds 250
        $at = [WordLayout]::Cursor()
        if (([Math]::Abs($at.X - $x) -le 2) -and ([Math]::Abs($at.Y - $y) -le 2)) { return $true }
        Write-HarnessNote ("{0}: asked for ({1},{2}), the pointer is at ({3},{4}) - trying again" -f
                           $what, $x, $y, $at.X, $at.Y)
        Start-Sleep -Milliseconds 500
    }
    return $false
}

# What the system says is under a point. The only authority on where a click will land.
function Get-ClassAt($x, $y) {
    $hwnd = [WordLayout]::WindowAt($x, $y)
    if ($hwnd -eq [IntPtr]::Zero) { return '' }
    return [WordLayout]::ClassOf($hwnd)
}

<#
  Click a point, having first checked that the thing under it is the thing being aimed at.

  `point` is a SCRIPTBLOCK and is re-run on every attempt, deliberately. The strip moves - 46px
  between two clicks a second apart, measured - and taking the foreground back can move it again, so
  a rectangle captured before the retry loop is the wrong rectangle by the time the retry happens.

  Returns $true only if the click was delivered onto $Expect. The caller asserts that as a check of
  its own: a click that missed and a feature that did nothing are otherwise the same evidence.
#>
function Invoke-ConfirmedClick {
    param(
        [string]$What,
        [scriptblock]$Point,
        [string]$Expect = 'WordTabStrip',
        [ValidateSet('left', 'right', 'middle')][string]$Button = 'left',
        [int]$Tries = 3
    )
    for ($try = 1; $try -le $Tries; $try++) {
        Set-WordForeground | Out-Null
        $p = & $Point
        $cls = Get-ClassAt $p.X $p.Y
        if ($cls -eq $Expect) {
            switch ($Button) {
                'right'  { [WordLayout]::RightClick($p.X, $p.Y) }
                'middle' { [WordLayout]::MiddleClick($p.X, $p.Y) }
                default  { [WordLayout]::Click($p.X, $p.Y) }
            }
            return $true
        }
        Write-HarnessNote ("{0}: ({1},{2}) is over `"{3}`", not {4} - taking the foreground back and re-measuring" -f
                           $What, $p.X, $p.Y, $cls, $Expect)

        # Say where the strip actually IS, not just what is under the point. "over NetUIHWND" on its
        # own does not distinguish the two causes: a point computed from the wrong window's strip, and
        # a point computed correctly onto a strip that Word's ribbon is drawn on top of. The rectangle
        # and the foreground window's title tell them apart on the first failing run instead of the
        # second. The icons slice lost a whole diagnosis to exactly this ambiguity.
        $fg = [WordLayout]::GetForeground()
        $strips = @([WordLayout]::Children($fg) | Where-Object { $_.Class -eq 'WordTabStrip' })
        if ($strips.Count -gt 0) {
            $sr = [WordLayout]::RectOf($strips[0].Hwnd)
            Write-HarnessNote ("  the foreground is `"{0}`" and its strip is at ({1},{2} {3}x{4}); the point is {5}px below its top" -f
                               [WordLayout]::TitleOf($fg), $sr.Left, $sr.Top, ($sr.Right - $sr.Left),
                               ($sr.Bottom - $sr.Top), ($p.Y - $sr.Top))
        } else {
            Write-HarnessNote ("  the foreground is `"{0}`" ({1}) and has NO strip on it" -f
                               [WordLayout]::TitleOf($fg), [WordLayout]::ClassOf($fg))
        }
        Write-HarnessNote ("  Word has up: {0}" -f (Format-WordChrome))
        Start-Sleep -Milliseconds 800
    }
    return $false
}

# The strip-click form, kept under the name the suites already call.
function Invoke-StripClick($what, [scriptblock]$point) {
    return Invoke-ConfirmedClick -What $what -Point $point -Expect 'WordTabStrip'
}

<#
  A keystroke, sent only once Word is in a position to receive it.

  A silently-missed keystroke looks exactly like a broken feature, three steps away. check-menu spent
  a whole slice failing "Save wrote the document to disk" because a click aimed at a hard-coded
  offset landed on the desktop: nothing typed, document not modified, Word's Save correctly wrote
  nothing, and the assertion blamed Save.

  Confirms that a WORD window has the foreground - not specifically a frame. Word's own dialogs,
  Backstage and popup menus are all legitimate targets for Escape, and demanding a frame would make
  this refuse to dismiss the very things it is most needed for.
#>
function Invoke-ConfirmedKey {
    param(
        [uint16]$Vk,
        [string]$What,
        [switch]$Ctrl,
        [int]$Tries = 3
    )
    for ($try = 1; $try -le $Tries; $try++) {
        if (Test-WordHasFocus) {
            if ($Ctrl) { [WordLayout]::CtrlPress($Vk) } else { [WordLayout]::Press($Vk) }
            return $true
        }
        Write-HarnessNote ("{0}: the foreground is `"{1}`" ({2}), not Word - taking it back" -f
                           $What, [WordLayout]::TitleOf([WordLayout]::GetForeground()),
                           [WordLayout]::ClassOf([WordLayout]::GetForeground()))
        Set-WordForeground | Out-Null
        Start-Sleep -Milliseconds 400
    }
    return $false
}

<#
  A keystroke aimed at ONE named window, not merely at Word.

  Invoke-ConfirmedKey above is right for Escape - anything of Word's may take it. It is wrong for
  Ctrl+W, which closes whichever document is in front: Set-WordForeground puts *some* frame forward,
  and if it is not the one the caller meant, the keystroke closes the wrong document and the suite
  goes on to measure a window that should not exist. So this one focuses the specific window and
  refuses to press until GetForegroundWindow agrees.
#>
function Invoke-ConfirmedKeyOn {
    param(
        [IntPtr]$Hwnd,
        [uint16]$Vk,
        [string]$What,
        [switch]$Ctrl,
        [int]$Tries = 4
    )
    for ($try = 1; $try -le $Tries; $try++) {
        if ([WordLayout]::GetForeground() -eq $Hwnd) {
            if ($Ctrl) { [WordLayout]::CtrlPress($Vk) } else { [WordLayout]::Press($Vk) }
            return $true
        }
        [WordLayout]::Focus($Hwnd) | Out-Null
        Wait-Until { [WordLayout]::GetForeground() -eq $Hwnd } 3 200 | Out-Null
    }
    Write-HarnessNote ("{0}: 0x{1:X} would not come to the front - the foreground is `"{2}`" ({3}). Not sending it." -f
                       $What, [int64]$Hwnd, [WordLayout]::TitleOf([WordLayout]::GetForeground()),
                       [WordLayout]::ClassOf([WordLayout]::GetForeground()))
    return $false
}

<#
  Take hold of something with the left button down, having confirmed the grab point is on it.

  DragHold and DragRelease are split so a check can photograph the strip mid-gesture. EVERY caller
  of this must call [WordLayout]::DragRelease from a finally block, including on failure: a left
  button left down is a wrecked desktop, exactly like a held Alt key. This returns $false without
  pressing anything when the grab point is wrong, so a failed confirmation leaves no button down.
#>
function Start-ConfirmedDrag {
    param(
        [string]$What,
        [scriptblock]$From,
        [int]$ToX, [int]$ToY,
        [string]$Expect = 'WordTabStrip',
        [int]$Steps = 12,
        [int]$DelayMs = 25,
        [int]$Tries = 3
    )
    for ($try = 1; $try -le $Tries; $try++) {
        Set-WordForeground | Out-Null
        $p = & $From
        $cls = Get-ClassAt $p.X $p.Y
        if ($cls -eq $Expect) {
            [WordLayout]::DragHold($p.X, $p.Y, $ToX, $ToY, $Steps, $DelayMs)
            return [pscustomobject]@{ Ok = $true; X = $p.X; Y = $p.Y }
        }
        Write-HarnessNote ("{0}: grab point ({1},{2}) is over `"{3}`", not {4} - re-measuring" -f
                           $What, $p.X, $p.Y, $cls, $Expect)
        Start-Sleep -Milliseconds 800
    }
    return [pscustomobject]@{ Ok = $false; X = 0; Y = 0 }
}

# ---- shutting Word down -----------------------------------------------------------------------

<#
  Close every Word window, and refuse to leave a question unanswered.

  NEVER Kill. A killed Word offers to recover those documents on the next launch, and the recovered
  documents then make the next suite measure seven windows where it asserts three. That rule had a
  live violation inside soak-stack for six slices.

  A test script may not answer a save prompt it did not raise - Escape cancels, it does not discard -
  so a surviving question is reported and the caller decides. Returns an object rather than throwing,
  because "Word would not close" and "Word is asking something" are different results and merging
  them lets a suite that could not clean up pass for one that did.
#>
function Close-AllWord {
    param([int]$Seconds = 25, [switch]$Quiet)

    for ($round = 1; $round -le 12; $round++) {
        $frames = @(Get-WordFrameList)
        if ($frames.Count -eq 0) { break }
        [WordLayout]::Close($frames[0])

        # A prompt raised by that close has to be seen BEFORE the next close is posted, or the second
        # WM_CLOSE queues up behind a modal nobody looked at.
        #
        # WM_CLOSE is posted, so neither outcome is instant, and reading for a dialog straight after
        # posting it reliably reads "no dialog" - which is the bug the per-suite loops this replaced
        # avoided by sleeping 1200-2000ms first. Sleeping is the wrong fix twice over: too short and
        # the prompt is still missed, too long and it is paid on every close in every suite.
        #
        # So: wait for whichever of the two outcomes happens. The window goes (the common case,
        # answered in a poll or two) or Word puts a question up. Only then is the double read taken.
        Wait-Until { (@(Get-WordFrameList).Count -lt $frames.Count) -or ($null -ne (Get-WordDialog)) } 8 250 | Out-Null

        $prompt = Get-WordSavePrompt
        if ($prompt) {
            if (-not $Quiet) { Write-HarnessNote ("Word is asking something: {0}" -f (Format-WordWindow $prompt)) }
            return [pscustomobject]@{ Closed = $false; Reason = 'question'; Dialog = $prompt
                                      Frames = (Get-WordFrameTally) }
        }
    }

    if (Wait-WordGone $Seconds) { return [pscustomobject]@{ Closed = $true; Reason = 'gone'; Dialog = $null; Frames = 0 } }

    return [pscustomobject]@{ Closed = $false; Reason = 'still running'; Dialog = (Get-WordDialog)
                              Frames = (Get-WordFrameTally) }
}
