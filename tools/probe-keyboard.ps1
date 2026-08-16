<#
.SYNOPSIS
  Measure what it would take for WordTab to hear a keyboard request to switch documents.

.DESCRIPTION
  This is the research half of queue item 5, "Ctrl+Tab between documents". It writes no product
  code and asserts no feature, because there is no feature yet - it exists to replace three guesses
  with three measurements, so that the slice that follows is a build rather than an investigation.

  The project's note on this item says the obstacle is that "the strip is WS_EX_NOACTIVATE and never
  has focus, so keystrokes never arrive". That is true of the STRIP. It is not obviously true of the
  ADD-IN, which lives inside Word's own process and already installs a thread-local WH_CBT hook on
  Word's UI thread (src\native\frames.cpp:656) that has been reported `installed` on every run this
  project has ever logged. So the questions worth measuring are not "can a keystroke be heard" but:

    1. WHICH window does Word give the keyboard to? A WM_KEYDOWN goes to the focus window and
       nowhere else - keyboard messages do not travel up to parents the way WM_CONTEXTMENU does. If
       the focus window is not one of the two we subclass, then a subclass can never see a keystroke
       and a thread hook is the only route. That is a fact about Word, so it is measured, not argued.

    2. Does Word's OWN Ctrl+F6 already switch between the stacked documents, and does the tab row
       follow it? frames.cpp:344 turns WM_ACTIVATE into StackOnFrameActivate, so it might already
       work end to end - in which case this feature is much smaller than the queue thinks, and
       becomes "make Ctrl+Tab do what Ctrl+F6 already does". This project has twice had a product
       decision dissolve into a measurement; this is the third candidate.

    3. Does Word already OWN Ctrl+Tab? If Word puts a tab character in the document when it is
       pressed, then binding it costs the user something real, and that cost has to be known before
       it is spent rather than discovered afterwards by whoever uses it.

  HOW THE ANSWERS ARE READ. Three oracles, deliberately, because each one alone has a way of being
  wrong:
    - the FOREGROUND WINDOW and its title says which document Word brought forward;
    - the ADD-IN'S LOG (`stack  active -> `, stack.cpp:451) says which tab WordTab thinks is active,
      and the two disagreeing is itself a finding rather than an error in the probe;
    - WORD'S OBJECT MODEL through [WordLayout]::NativeOm, which asks one NAMED window for its own
      Word.Window and so cannot bind to the wrong instance the way New-Object -ComObject does.
      `Document.Saved` is how "did a tab character get inserted" is answered without reading pixels.

  EVERY KEYSTROKE IS CONFIRMED. An injected key that never landed and a feature that does nothing
  produce byte-identical evidence, and this project has already lost a diagnosis to exactly that.
  Invoke-ConfirmedKey checks a Word window holds the focus before it presses anything.

  AND EVERY KEYSTROKE IS PRECEDED BY A REAL CLICK INTO THE PAGE. A window that was only activated
  programmatically does not take keystrokes in Word - measured in check-dot - so the click is what
  makes the caret real, and it is aimed at the _WwF rectangle rather than at a fixed offset, because
  Word remembers its window size and a hard-coded offset has landed on the desktop before now.

  Fixtures are three scratch .rtf files in %TEMP%\wordtab-keyboard. They are disposable. Section 4
  may put a tab character into one of them on purpose; it is undone, and the cleanup marks them
  saved so that closing Word cannot raise a question - see Close-Word below for why that is
  acceptable here and would not be anywhere near the add-in.

.PARAMETER KeepOpen
  Leave Word running afterwards.

.EXAMPLE
  pwsh -File tools\probe-keyboard.ps1
#>
[CmdletBinding()]
param(
    [switch]$KeepOpen
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$source = Get-Content -Raw -Path (Join-Path $PSScriptRoot 'WordLayout.cs')
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition $source -Language CSharp -ReferencedAssemblies @(
    'System.Runtime', 'System.Collections', 'System.Threading.Thread', 'netstandard'
)
[WordLayout]::MakeDpiAware() | Out-Null

. (Join-Path $PSScriptRoot 'WordTabHarness.ps1')

trap {
    Write-Host ''
    Write-Host "ERROR  $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
    break
}

function Write-Step($text) { Write-Host ''; Write-Host "==> $text" -ForegroundColor Cyan }
function Write-Note($text) { Write-Host "    $text" -ForegroundColor DarkGray }
function Write-Found($text) { Write-Host "    FOUND  $text" -ForegroundColor Yellow }

$VK = @{ F6 = 0x75; TAB = 0x09; Z = 0x5A }

$script:Findings = @()
function Add-Finding($text) { $script:Findings += $text; Write-Found $text }

function Get-Frames { return Get-WordFrameList }

function Get-Parts($frame) {
    $kids = [WordLayout]::Children($frame)
    [pscustomobject]@{
        Frame = $frame
        Strip = @($kids | Where-Object { $_.Class -eq 'WordTabStrip' }) | Select-Object -First 1
        Wwf   = @($kids | Where-Object { $_.Class -eq '_WwF' }) | Select-Object -First 1
        Title = [WordLayout]::TitleOf($frame)
    }
}

# The document name Word itself reports for one window, through that window's own object model.
# $null when Word will not answer - which is a result, not a failure, and must never be folded in
# with "Word answered and the document is unnamed".
function Get-DocState($frame) {
    $om = [WordLayout]::NativeOm($frame)
    if ($null -eq $om) { return $null }
    try {
        $doc = $om.Document
        return [pscustomobject]@{ Name = [string]$doc.Name; Saved = [bool]$doc.Saved }
    } catch {
        return $null
    }
}

# Which document Word has in front, right now, by all three oracles at once.
function Get-Active($label) {
    $fg    = [WordLayout]::GetForeground()
    $cls   = [WordLayout]::ClassOf($fg)
    $isOur = ($cls -eq 'OpusApp') -and ([WordLayout]::PidOf($fg) -in @(Get-WordPidList))
    $doc   = if ($isOur) { Get-DocState $fg } else { $null }
    $state = [pscustomobject]@{
        Label    = $label
        Hwnd     = $fg
        Class    = $cls
        Title    = [WordLayout]::TitleOf($fg)
        IsFrame  = $isOur
        Document = if ($null -ne $doc) { $doc.Name } else { '(no answer)' }
        Saved    = if ($null -ne $doc) { $doc.Saved } else { $null }
    }
    Write-Note ("{0,-22} foreground 0x{1:X} ({2})  title |{3}|  document |{4}|  saved={5}" -f
                $label, [int64]$state.Hwnd, $state.Class, $state.Title, $state.Document,
                $(if ($null -ne $state.Saved) { $state.Saved } else { '?' }))
    return $state
}

# Put a real caret in the body of a Word window, and prove the click landed on Word's document frame
# rather than on the desktop.
#
# -Frame names WHICH window, for the sections that need a particular document to be the current one.
# Without it Set-WordForeground brings *some* frame forward, which is right when any will do and
# quietly wrong when the measurement is about which one. Same distinction the harness draws between
# Invoke-ConfirmedKey and Invoke-ConfirmedKeyOn.
function Set-Caret($what, $Frame = $null) {
    if ($null -ne $Frame) {
        [WordLayout]::Focus($Frame) | Out-Null
        Wait-Until { [WordLayout]::GetForeground() -eq $Frame } 4 200 | Out-Null
        if ([WordLayout]::GetForeground() -ne $Frame) {
            Write-Note "$what : 0x$('{0:X}' -f [int64]$Frame) would not come forward"
            return $false
        }
    } else {
        Set-WordForeground | Out-Null
    }
    Start-Sleep -Milliseconds 300
    $fg = [WordLayout]::GetForeground()
    if ([WordLayout]::ClassOf($fg) -ne 'OpusApp') {
        Write-Note "$what : the foreground is $([WordLayout]::ClassOf($fg)), not a Word frame"
        return $false
    }
    $parts = Get-Parts $fg
    if (-not $parts.Wwf) { Write-Note "$what : that frame has no document frame"; return $false }

    $view = [WordLayout]::RectOf($parts.Wwf.Hwnd)
    $vw = $view.Right - $view.Left
    $vh = $view.Bottom - $view.Top
    if ($vw -lt 200 -or $vh -lt 200) { throw "The document frame is only ${vw}x${vh} - too small to click into." }

    $x = $view.Left + [int]($vw / 3)
    $y = $view.Top  + [int]($vh / 2)

    # Passively, not through Invoke-ConfirmedClick: that takes the foreground first, and this click
    # has to be the thing that gives Word the foreground as genuine user input. Confirmed by what is
    # under the point instead, re-measured rather than assumed.
    $under = [WordLayout]::ClassOf([WordLayout]::WindowAt($x, $y))
    if ($under -notin @('_WwG', '_WwF', 'OpusApp')) {
        Write-Note "$what : (${x},${y}) is over |$under|, not the page"
        return $false
    }
    [WordLayout]::Click($x, $y)
    Start-Sleep -Milliseconds 500
    return $true
}

# ---- Word up and down --------------------------------------------------------------------------

$fixtureDir = Join-Path $env:TEMP 'wordtab-keyboard'

function New-Fixtures {
    if (Test-Path $fixtureDir) { Remove-Item $fixtureDir -Recurse -Force }
    New-Item -ItemType Directory -Path $fixtureDir | Out-Null
    $names = @('Alpha', 'Bravo', 'Charlie')
    $paths = @()
    foreach ($n in $names) {
        $p = Join-Path $fixtureDir "$n.rtf"
        # Minimal RTF, authored here rather than by Word: nothing in this probe depends on the file
        # format, only on there being three tellable-apart documents with a body to put a caret in.
        Set-Content -Path $p -Value "{\rtf1\ansi\deff0 $n document body text.\par}" -Encoding ASCII
        $paths += $p
    }
    return $paths
}

function Open-Document($path) {
    $before = @(Get-Frames)
    Start-Process -FilePath 'winword.exe' -ArgumentList "`"$path`""
    $deadline = (Get-Date).AddSeconds(45)
    while ((Get-Date) -lt $deadline) {
        $new = @(Get-Frames | Where-Object { $before -notcontains $_ })
        if ($new.Count -gt 0) {
            Wait-Until { @([WordLayout]::Children($new[0]) | Where-Object { $_.Class -eq 'WordTabStrip' }).Count -gt 0 } 20 300 | Out-Null
            Start-Sleep -Seconds 2
            return $new[0]
        }
        Start-Sleep -Milliseconds 400
    }
    throw "Word never opened a window for $path"
}

<#
  Close Word without letting it ask anything.

  This probe deliberately provokes an edit - that is the whole of section 4 - so one of its own
  scratch files can be dirty at the end, and Close-AllWord correctly refuses to answer a question it
  did not raise. So the fixtures are marked saved first.

  `Document.Saved` is settable and the standing rule in this project is that it is only ever READ.
  That rule protects the USER'S documents from the ADD-IN, which must never tell Word that something
  the user typed has been dealt with. It is not the same act as a probe discarding a change it made
  itself, sixty seconds earlier, to a file it authored itself, in %TEMP%, in a directory it is about
  to delete. The alternative is worse: leaving a modal question on the user's screen after the run.
  Nothing in src\native does this or may.
#>
function Close-Word {
    if (@(Get-WordPidList).Count -eq 0) { return }
    foreach ($f in @(Get-Frames)) {
        try {
            $om = [WordLayout]::NativeOm($f)
            if ($null -ne $om) { $om.Document.Saved = $true }
        } catch { }
    }
    $end = Close-AllWord
    if (-not $end.Closed) {
        if ($end.Reason -eq 'question') {
            throw ("Word is asking something: {0}  Answer it by hand and re-run." -f (Format-WordWindow $end.Dialog))
        }
        throw ("Word would not close ({0}, {1} frame(s) left)." -f $end.Reason, $end.Frames)
    }
    Start-Sleep -Seconds 2
}

# ================================================================================================

Write-Host ''
Write-Host 'probe-keyboard - what it takes for WordTab to hear a document-switch keystroke' -ForegroundColor White
Write-Host ''

try {
    Write-Step 'Starting from a Word that is actually closed'
    Close-Word
    $paths = New-Fixtures
    Write-Note "fixtures in $fixtureDir"

    foreach ($p in $paths) { Open-Document $p | Out-Null }
    $frames = @(Get-Frames)
    Write-Note "$($frames.Count) frame(s) up"
    if ($frames.Count -ne 3) { throw "Expected three Word frames, measured $($frames.Count)." }

    # ---- 1. Where Word puts the keyboard -------------------------------------------------------
    Write-Step '1. Which window does Word give the keyboard to?'

    Set-WordForeground | Out-Null
    Start-Sleep -Milliseconds 400
    if (-not (Set-Caret 'putting a real caret in the page')) {
        throw 'Could not put a caret in the document - nothing after this would be measuring a keystroke.'
    }

    $fg    = [WordLayout]::GetForeground()
    $gui   = [WordLayout]::ThreadGui($fg)
    $parts = Get-Parts $fg

    $focusCls  = if ($gui.hwndFocus -ne [IntPtr]::Zero) { [WordLayout]::ClassOf($gui.hwndFocus) } else { '(none)' }
    $activeCls = if ($gui.hwndActive -ne [IntPtr]::Zero) { [WordLayout]::ClassOf($gui.hwndActive) } else { '(none)' }

    Write-Note ("hwndActive 0x{0:X} |{1}|" -f [int64]$gui.hwndActive, $activeCls)
    Write-Note ("hwndFocus  0x{0:X} |{1}|" -f [int64]$gui.hwndFocus, $focusCls)
    Write-Note ("frame      0x{0:X} |OpusApp|" -f [int64]$fg)
    if ($parts.Wwf)   { Write-Note ("_WwF       0x{0:X}  (subclassed by the add-in)" -f [int64]$parts.Wwf.Hwnd) }
    if ($parts.Strip) { Write-Note ("strip      0x{0:X}  (WS_EX_NOACTIVATE)" -f [int64]$parts.Strip.Hwnd) }

    $focusIsFrame = ($gui.hwndFocus -eq $fg)
    $focusIsWwf   = ($parts.Wwf -and $gui.hwndFocus -eq $parts.Wwf.Hwnd)
    $focusIsStrip = ($parts.Strip -and $gui.hwndFocus -eq $parts.Strip.Hwnd)

    if ($focusIsStrip) {
        Add-Finding 'The STRIP holds the keyboard focus - the recorded obstacle is wrong outright.'
    } elseif ($focusIsFrame -or $focusIsWwf) {
        Add-Finding ("The focus window IS one the add-in already subclasses ({0}) - a subclass can see WM_KEYDOWN, no hook needed." -f $focusCls)
    } else {
        Add-Finding ("Word gives the keyboard to |{0}|, which the add-in does NOT subclass. Keyboard messages do not reach parents, so a subclass of OpusApp or _WwF can never see the keystroke: a thread-level hook is the route." -f $focusCls)
    }

    # ---- 2. Word's own Ctrl+F6 -----------------------------------------------------------------
    Write-Step "2. Word's own Ctrl+F6 - does it switch documents, and does the tab row follow?"

    $order = @()
    $before = Get-Active 'before Ctrl+F6'
    $order += $before.Document

    $followed = 0
    $moved    = 0
    for ($i = 1; $i -le 3; $i++) {
        Set-LogMark
        if (-not (Invoke-ConfirmedKey -Vk $VK.F6 -Ctrl -What "Ctrl+F6 (press $i)")) {
            throw 'Ctrl+F6 could not be delivered - Word would not hold the focus.'
        }
        Start-Sleep -Milliseconds 1400

        $after = Get-Active "after Ctrl+F6 #$i"
        $said  = Get-LogLast 'stack  active ->'
        if ($said) { Write-Note "add-in: $($said.Substring($said.IndexOf('stack')))" }
        else       { Write-Note 'add-in: said nothing about the active tab' }

        if ($after.Hwnd -ne $before.Hwnd) { $moved++ }
        if ($said -and $after.IsFrame -and ($said -match '0x0*([0-9A-Fa-f]+)\s*$')) {
            $claimed = [int64]('0x' + $Matches[1])
            if ($claimed -eq [int64]$after.Hwnd) { $followed++ }
            else { Write-Note ("MISMATCH: the tab row says 0x{0:X} is active, Word has 0x{1:X} in front" -f $claimed, [int64]$after.Hwnd) }
        }
        $order += $after.Document
        $before = $after
    }

    Write-Note ("cycle: {0}" -f ($order -join '  ->  '))

    if ($moved -eq 0) {
        Add-Finding 'Ctrl+F6 does NOT change which document is in front while the stack is on.'
    } elseif ($moved -eq 3 -and $followed -eq 3) {
        # Deliberately says nothing about what Ctrl+Tab should be implemented as. Whether Word's own
        # command can be borrowed depends on the order model, which section 6 measures and which this
        # sequence cannot tell apart on its own.
        Add-Finding 'Ctrl+F6 ALREADY switches documents, and the tab row follows it every time - WordTab tracks a switch it did not initiate, through WM_ACTIVATE. So keyboard document switching is not absent today; it is unbound and undiscoverable.'
    } else {
        Add-Finding ("Ctrl+F6 changed the front document {0} of 3 times, and the tab row agreed {1} of 3." -f $moved, $followed)
    }

    # ---- 3. Ctrl+Shift+F6 ----------------------------------------------------------------------
    Write-Step '3. Ctrl+Shift+F6 - the reverse direction'

    $back = @()
    $before = Get-Active 'before Ctrl+Shift+F6'
    $back += $before.Document
    $backMoved = 0
    for ($i = 1; $i -le 2; $i++) {
        Set-LogMark
        if (-not (Invoke-ConfirmedKey -Vk $VK.F6 -Ctrl -Shift -What "Ctrl+Shift+F6 (press $i)")) {
            throw 'Ctrl+Shift+F6 could not be delivered.'
        }
        Start-Sleep -Milliseconds 1400
        $after = Get-Active "after Ctrl+Shift+F6 #$i"
        if ($after.Hwnd -ne $before.Hwnd) { $backMoved++ }
        $back += $after.Document
        $before = $after
    }
    Write-Note ("cycle: {0}" -f ($back -join '  ->  '))
    if ($backMoved -eq 2) { Add-Finding 'Ctrl+Shift+F6 cycles the other way, so both directions already exist in Word.' }
    else { Add-Finding ("Ctrl+Shift+F6 moved the front document {0} of 2 times." -f $backMoved) }

    # ---- 4. Does Word own Ctrl+Tab? ------------------------------------------------------------
    Write-Step '4. Ctrl+Tab with the caret in the document body - does Word already own it?'

    if (-not (Set-Caret 'caret into the body before Ctrl+Tab')) {
        throw 'Could not put a caret in the body - Ctrl+Tab would be measuring nothing.'
    }
    $before = Get-Active 'before Ctrl+Tab'
    if ($null -eq $before.Saved) { throw 'Word would not report Document.Saved - there is no baseline to compare against.' }
    if (-not $before.Saved) { Write-Note 'NOTE: that document was already modified before the keystroke.' }

    Set-LogMark
    if (-not (Invoke-ConfirmedKey -Vk $VK.TAB -Ctrl -What 'Ctrl+Tab')) { throw 'Ctrl+Tab could not be delivered.' }
    Start-Sleep -Milliseconds 1400

    $after = Get-Active 'after Ctrl+Tab'
    $switched = ($after.Hwnd -ne $before.Hwnd)
    $dirtied  = ($before.Saved -and ($null -ne $after.Saved) -and (-not $after.Saved))

    if ($switched) {
        Add-Finding 'Ctrl+Tab ALREADY switches document in Word - it does not need binding at all.'
    } elseif ($dirtied) {
        # What that costs is section 7's question, not this one - in the body plain Tab may do the
        # same thing, in which case Ctrl+Tab is a second route to something the user already has.
        Add-Finding 'Word OWNS Ctrl+Tab in the document body: it put a tab character in the document. So it is not a free chord - it would have to be swallowed before Word sees it.'
    } else {
        Add-Finding 'Ctrl+Tab in the body does nothing Word cares about - it neither switches document nor edits it. It is free to take.'
    }

    if ($dirtied) {
        Write-Note 'undoing the inserted character'
        Invoke-ConfirmedKey -Vk $VK.Z -Ctrl -What 'Ctrl+Z' | Out-Null
        Start-Sleep -Milliseconds 700
        $undone = Get-Active 'after Ctrl+Z'
        Write-Note ("document saved-state is now {0}" -f $undone.Saved)
    }

    # ---- 5. Ctrl+Shift+Tab ---------------------------------------------------------------------
    Write-Step '5. Ctrl+Shift+Tab in the body'

    if (-not (Set-Caret 'caret into the body before Ctrl+Shift+Tab')) {
        throw 'Could not put a caret in the body.'
    }
    $before = Get-Active 'before Ctrl+Shift+Tab'
    Set-LogMark
    if (-not (Invoke-ConfirmedKey -Vk $VK.TAB -Ctrl -Shift -What 'Ctrl+Shift+Tab')) { throw 'Ctrl+Shift+Tab could not be delivered.' }
    Start-Sleep -Milliseconds 1400
    $after = Get-Active 'after Ctrl+Shift+Tab'

    if ($after.Hwnd -ne $before.Hwnd) { Add-Finding 'Ctrl+Shift+Tab already switches document in Word.' }
    elseif ($before.Saved -and ($null -ne $after.Saved) -and (-not $after.Saved)) { Add-Finding 'Word owns Ctrl+Shift+Tab in the body too - it edits the document.' }
    else { Add-Finding 'Ctrl+Shift+Tab in the body does nothing Word cares about.' }

    # ---- 6. Is Word's window order the TAB order, or most-recently-used? ------------------------
    #
    # Sections 2 and 3 cannot tell these apart and it is not a detail. Ctrl+F6 went
    # Charlie -> Bravo -> Alpha, which is "one tab to the left, wrapping" if Word keeps its windows
    # in creation order - and is ALSO exactly what "the next most recently used window" produces,
    # because the documents were opened in that order and so were last touched in that order. The
    # two hypotheses fit the same data.
    #
    # They need different code. If the order is the tab order, Ctrl+Tab can be implemented by handing
    # Word its own command and letting Word move. If it is MRU, that is impossible - "the tab to the
    # right of this one" is not a thing Word's command can express - and the add-in has to work out
    # the target itself and activate it, which is StackActivate and a different slice.
    #
    # The discriminator is to make the two orders disagree before pressing anything: activate Bravo,
    # then Alpha. Creation order is still Alpha, Bravo, Charlie; MRU is now Alpha, Bravo, Charlie
    # with Alpha current, so "one left of Alpha, wrapping" is CHARLIE and "next most recent" is
    # BRAVO. Whichever comes forward names the model.
    Write-Step "6. Is Word's window order the tab order, or most-recently-used?"

    $byOpen = @(Get-Frames)
    if ($byOpen.Count -ne 3) { throw "Expected three frames for the order test, measured $($byOpen.Count)." }
    # Get-WordFrameList is enumeration order, which is z-order and not the tab order. The tab order is
    # the order the documents were opened, and the only thing here that knows it is the title.
    $byName = @{}
    foreach ($f in $byOpen) {
        $d = Get-DocState $f
        if ($null -ne $d) { $byName[$d.Name] = $f }
    }
    foreach ($n in @('Alpha.rtf', 'Bravo.rtf', 'Charlie.rtf')) {
        if (-not $byName.ContainsKey($n)) { throw "Could not find the window holding $n - the order test has nothing to aim at." }
    }

    if (-not (Set-Caret 'making Bravo current' $byName['Bravo.rtf'])) { throw 'Could not make Bravo current.' }
    Start-Sleep -Milliseconds 600
    if (-not (Set-Caret 'making Alpha current' $byName['Alpha.rtf'])) { throw 'Could not make Alpha current.' }
    Start-Sleep -Milliseconds 600

    $before = Get-Active 'before the order test (Alpha current, Bravo second most recent)'
    if ($before.Document -ne 'Alpha.rtf') { throw "Alpha is not current - the order test would measure nothing. Front document is $($before.Document)." }

    Set-LogMark
    if (-not (Invoke-ConfirmedKey -Vk $VK.F6 -Ctrl -What 'Ctrl+F6 (order test)')) { throw 'Ctrl+F6 could not be delivered.' }
    Start-Sleep -Milliseconds 1400
    $after = Get-Active 'after Ctrl+F6 from Alpha'

    if ($after.Document -eq 'Charlie.rtf') {
        Add-Finding "Word's order is the TAB ORDER, not MRU: from Alpha with Bravo most-recent, Ctrl+F6 went to Charlie. So Ctrl+F6 is reliably one tab LEFT and Ctrl+Shift+F6 one tab RIGHT - inverted from every tabbed application, and Ctrl+Tab must be bound to the Ctrl+Shift+F6 direction, not the Ctrl+F6 one."
    } elseif ($after.Document -eq 'Bravo.rtf') {
        Add-Finding "Word's order is MOST-RECENTLY-USED: from Alpha with Bravo most-recent, Ctrl+F6 went to Bravo. So Word's own command cannot express `"the next tab`" and forwarding it is not an implementation of Ctrl+Tab - the add-in has to compute the target from g_members and call StackActivate itself."
    } else {
        Add-Finding ("The order test is inconclusive: from Alpha, Ctrl+F6 went to {0}." -f $after.Document)
    }

    # ---- 7. What binding Ctrl+Tab would actually cost -------------------------------------------
    #
    # Section 4 measured that Ctrl+Tab edits the document, which is the cost - but not the whole
    # shape of it. In ordinary body text plain Tab inserts a tab as well, so there Ctrl+Tab is a
    # SECOND way to do something the user already has a first way to do, and taking it costs nothing
    # anyone would notice. Inside a TABLE they diverge: Tab is "next cell" and Ctrl+Tab is the only
    # way to get a literal tab into a cell. So the real question is not "does Ctrl+Tab do something"
    # but "is there anywhere it is the ONLY way to do something", and that is one measurement.
    Write-Step '7. What binding Ctrl+Tab would actually cost'

    Write-Note '7a. plain Tab in ordinary body text'
    if (-not (Set-Caret 'caret into Charlie for the Tab test' $byName['Charlie.rtf'])) { throw 'Could not put a caret in Charlie.' }
    $before = Get-Active 'before plain Tab'
    if (-not $before.Saved) { throw 'Charlie was already modified - there is no baseline for the Tab test.' }
    if (-not (Invoke-ConfirmedKey -Vk $VK.TAB -What 'plain Tab in the body')) { throw 'Tab could not be delivered.' }
    Start-Sleep -Milliseconds 900
    $after = Get-Active 'after plain Tab'
    $tabEdits = ($null -ne $after.Saved) -and (-not $after.Saved)
    if ($tabEdits) {
        Add-Finding 'In ordinary body text plain Tab inserts a tab too, so Ctrl+Tab is a second route to something the user already has - taking it there costs nothing.'
        Invoke-ConfirmedKey -Vk $VK.Z -Ctrl -What 'Ctrl+Z after the Tab test' | Out-Null
        Start-Sleep -Milliseconds 700
    } else {
        Add-Finding 'Plain Tab did NOT edit the body - so Ctrl+Tab is not redundant with it and the cost of binding it is higher than assumed.'
    }

    Write-Note '7b. inside a table, where Tab and Ctrl+Tab are documented to differ'
    $tableFrame = $byName['Charlie.rtf']
    $om = [WordLayout]::NativeOm($tableFrame)
    if ($null -eq $om) { throw 'Charlie would not answer through the object model - the table test cannot be set up.' }

    # The table has to go at the START of the document, not at wherever the caret happens to be.
    # The first attempt added it at Selection.Range - which was mid-paragraph, left over from 7a -
    # and then Ctrl+Home landed in the text ABOVE the table. The guard below caught it and stopped
    # the run rather than measuring the body and reporting it as a cell, which is the whole reason
    # it is there. So: put the caret at the document start with a real keystroke first, and build
    # the table there.
    if (-not (Set-Caret 'caret into Charlie before building the table' $tableFrame)) { throw 'Could not put a caret in Charlie.' }
    if (-not (Invoke-ConfirmedKey -Vk 0x24 -Ctrl -What 'Ctrl+Home before building the table')) { throw 'Ctrl+Home could not be delivered.' }
    Start-Sleep -Milliseconds 500

    $om.Document.Tables.Add($om.Selection.Range, 2, 3) | Out-Null
    Start-Sleep -Milliseconds 800
    Write-Note 'authored a 2x3 table at the start of Charlie'

    # Back to the very start, which is now inside cell (1,1) - reached by a real keystroke rather
    # than by a programmatic Select(), so that what follows is measuring the same kind of caret a
    # user would have.
    if (-not (Invoke-ConfirmedKey -Vk 0x24 -Ctrl -What 'Ctrl+Home into the first cell')) { throw 'Ctrl+Home could not be delivered.' }
    Start-Sleep -Milliseconds 700

    # wdWithInTable = 12. Confirmed BEFORE anything is measured, because "the keystroke did nothing"
    # and "the caret was never in a cell" are the same evidence otherwise.
    $inTable = [bool]$om.Selection.Information(12)
    Write-Note "selection is inside a table: $inTable"
    if (-not $inTable) { throw 'The caret is not in a table cell - 7b would be measuring the body again.' }

    $om.Document.Saved = $true
    $startBefore = [int]$om.Selection.Range.Start
    if (-not (Invoke-ConfirmedKey -Vk $VK.TAB -What 'Tab inside a cell')) { throw 'Tab could not be delivered.' }
    Start-Sleep -Milliseconds 800
    $startAfterTab = [int]$om.Selection.Range.Start
    $savedAfterTab = [bool]$om.Document.Saved
    Write-Note ("Tab in a cell: caret {0} -> {1}, document saved={2}" -f $startBefore, $startAfterTab, $savedAfterTab)

    $om.Document.Saved = $true
    if (-not (Invoke-ConfirmedKey -Vk $VK.TAB -Ctrl -What 'Ctrl+Tab inside a cell')) { throw 'Ctrl+Tab could not be delivered.' }
    Start-Sleep -Milliseconds 800
    $savedAfterCtrlTab = [bool]$om.Document.Saved
    Write-Note ("Ctrl+Tab in a cell: document saved={0}" -f $savedAfterCtrlTab)

    if ((-not $savedAfterCtrlTab) -and $savedAfterTab -and ($startAfterTab -ne $startBefore)) {
        Add-Finding 'Inside a table the two really are different: Tab moves to the next cell without editing, and Ctrl+Tab is the ONLY way to put a literal tab in a cell. That is the entire real cost of binding Ctrl+Tab, and it is confined to tables.'
    } else {
        Add-Finding ("Inside a table: Tab moved the caret {0} and edited={1}; Ctrl+Tab edited={2}. They do not divide the way the documentation says on this build." -f
                     ($startAfterTab -ne $startBefore), (-not $savedAfterTab), (-not $savedAfterCtrlTab))
    }

    # ---- what the add-in said about any of this ------------------------------------------------
    Write-Step 'What the add-in logged for the whole run'
    Set-LogMark
}
finally {
    Write-Host ''
    if (-not $KeepOpen) {
        Write-Step 'Cleanup'
        try { Close-Word } catch { Write-Host "    cleanup: $($_.Exception.Message)" -ForegroundColor Yellow }
        if (Test-Path $fixtureDir) { Remove-Item $fixtureDir -Recurse -Force -ErrorAction SilentlyContinue }
    } else {
        Write-Note 'Word left running (-KeepOpen)'
    }

    Write-Host ''
    Write-Host '---- FINDINGS -------------------------------------------------------------------' -ForegroundColor White
    if ($script:Findings.Count -eq 0) { Write-Host '  (none - the probe did not get far enough to measure anything)' -ForegroundColor Red }
    foreach ($f in $script:Findings) { Write-Host "  * $f" -ForegroundColor Yellow }
    Write-Host ''
}
