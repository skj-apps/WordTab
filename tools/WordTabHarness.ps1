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

.PARAMETER SelfTest
  Run the log primitives against a scratch file and print one line per check. Dot-sourcing passes no
  arguments, so this is inert in every suite; `pwsh -File tools\WordTabHarness.ps1 -SelfTest` runs it
  in about a second with no Word and no desktop.

  It exists because the failure it guards against cannot be reached any other way inside a run: the
  log rolls at 512KB, which takes two whole batteries to provoke, and the failure it produces is
  silence. A scratch file can be rolled in a millisecond.

.NOTES
  Depends on [WordLayout] only - and the log primitives depend on nothing at all, which is what lets
  -SelfTest run them without Word. It deliberately does NOT call back into suite-local helpers
  (Write-Note, Get-Frames, Set-WordForeground): those differ between suites, and a shared file whose
  behaviour depends on which suite loaded it is not shared.
#>
param([switch]$SelfTest)

# **Dot-sourcing a script with a param block leaves its parameters behind in the CALLER's scope, and
# they keep their type constraint.** So every suite that dot-sources this file was getting a variable
# called $SelfTest that only accepts a [switch] - and check-all.ps1's own `$selfTest = & pwsh ...`
# then died with "Cannot convert System.Object[] to SwitchParameter", because PowerShell variable
# names are case-insensitive and it was assigning to this one. Found by a run, not by reading it.
#
# So the flag is copied somewhere harmless and the name is given back. A shared file may not leave
# names in the scope of everything that loads it, least of all names with types attached.
$script:HarnessSelfTest = [bool]$SelfTest
Remove-Variable -Name SelfTest -Scope 0 -ErrorAction SilentlyContinue

# ---- diagnosis --------------------------------------------------------------------------------
#
# Its own note writer rather than the suites' Write-Note. Same shape on screen; no late binding into
# a caller that may not have defined it yet.

function Write-HarnessNote($text) { Write-Host "    $text" -ForegroundColor DarkGray }

# ---- the add-in's own log ---------------------------------------------------------------------

<#
  The add-in's log is the only instrument that can see inside Word, and reading it correctly is not
  as simple as it looks. Seven copies of this had grown across the suites and every one of them was
  wrong in the same place.

  **THE LOG ROLLS BY DELETING ITSELF.** src\native\log.cpp:42 calls DeleteFileW once the file passes
  512KB, before every write. A full battery writes about 250KB, so two batteries in a session roll
  it - and it rolled in the middle of check-title during the run that found this.

  A mark was a byte offset, which after that deletion means nothing. The six copies then split into
  two ways of being wrong, and BOTH of them are silent:

    - four suites did `if ($offset -gt $stream.Length) { $offset = 0 }` and read the whole of the
      fresh file. Everything written between the mark and the roll is gone, so "the log says nothing
      since the mark" goes GREEN having lost the evidence that would have failed it
    - check-title clamped with `[Math]::Min($mark, $stream.Length)` and so read from the END of the
      fresh file. That returns nothing, always: every "says nothing" assertion passes and every
      "says X" assertion fails for a reason that is not the product

  The seventh copy, check-menu's Get-LogTail, had no mark at all - it took the last match in the last
  600 lines - so it could answer with the PREVIOUS run's line. That one sits under the most
  safety-critical assertion in the project ("the batch stopped because the user declined").

  The dangerous shape in all of them is the same: **a read that cannot say it read nothing.**

  So a mark here is not an offset. It is an offset plus the bytes that were at it - the tail of the
  last line the add-in had written when the mark was taken - and a read verifies those bytes are
  still there before it believes the offset. A rolled log is then a hard error with a full
  explanation, not an empty result: once the evidence has been deleted, no assertion downstream of it
  is answering the question it claims to.

  Three detectors, because a roll can be reached three ways:

    the file is shorter than the mark      the roll happened and the file has not caught up
    the anchor bytes have changed          it rolled AND regrew past the mark. An offset alone can
                                           never see this, and it is why the note that queued this
                                           work asked for a marker line rather than a number
    the file is shorter than the HIGH      covers a mark taken at 0 bytes, which has no anchor to
    WATER of what this mark has already    check. Every read updates it, and the suites read in
    seen                                   Wait-Until loops, so a roll between two reads is caught

  Marking the log by APPENDING a line to it was the other candidate and was rejected: .NET's
  FileMode.Append is a seek-to-end and then a write, not the atomic FILE_APPEND_DATA the add-in uses,
  so it can land on top of a line Word's thread wrote in between. A test instrument may not corrupt
  the evidence it is reading. Reading the bytes that are already there costs nothing and cannot.

  Reset-LogFile is the other half, and the cheaper one: a suite that starts from an empty log cannot
  reach 512KB on its own - the biggest single suite writes about 40KB - so the roll never fires
  mid-run. The detectors above stay as the backstop for when it does anyway.
#>

$LogPath = Join-Path $env:LOCALAPPDATA 'WordTab\wordtab.log'

# Where Reset-LogFile puts each suite's log. Inside the add-in's own data directory on purpose, so
# uninstall.ps1 still takes everything WordTab ever created with it in one Remove-Item -Recurse.
$script:LogHistoryDir = Join-Path $env:LOCALAPPDATA 'WordTab\history'
$script:LogHistoryKeep = 40

# How much of the tail of the file a mark remembers. Measured over one battery's log rather than
# guessed: 2030 lines, 62 to 315 bytes, mean 121. So 256 spans a line or two and always covers at
# least one of the add-in's millisecond timestamps, which is what makes the anchor IDENTIFY the file
# rather than merely describe it. It deliberately does not have to be a whole line - the longest
# lines are longer than this - because what is being checked is that a byte range is still where it
# was, not that it parses.
$script:LogAnchorBytes = 256

$script:LogMark = $null

# Exactly $count bytes at $at, as hex, or '' if they could not all be read. Hex rather than text
# because $at can land inside a UTF-8 sequence, and two different byte strings must never decode to
# the same comparison string.
function Get-LogBytesHex($stream, [int64]$at, [int64]$count) {
    if ($count -le 0) { return '' }
    if (($at + $count) -gt $stream.Length) { return '' }
    $buffer = New-Object byte[] $count
    $stream.Seek($at, 'Begin') | Out-Null
    $got = 0
    while ($got -lt $count) {
        $n = $stream.Read($buffer, $got, ($count - $got))
        if ($n -le 0) { break }
        $got += $n
    }
    if ($got -ne $count) { return '' }
    return [System.BitConverter]::ToString($buffer)
}

# Take a mark on the log as it stands now. Not exported under a Get- name deliberately: the old
# `$mark = Get-LogMark` returned a number, and a function of that name that returned anything else
# would give every un-migrated call site a plausible wrong answer instead of an error.
function New-LogMarkNow {
    $length = [int64]0
    $anchor = ''
    $at     = [int64]0
    if (Test-Path $LogPath) {
        $stream = [System.IO.File]::Open($LogPath, 'Open', 'Read', 'ReadWrite')
        try {
            $length = [int64]$stream.Length
            $want   = [int64][Math]::Min([int64]$script:LogAnchorBytes, $length)
            $at     = $length - $want
            $anchor = Get-LogBytesHex $stream $at $want
        } finally { $stream.Dispose() }
    }
    return [pscustomobject]@{
        Length      = $length
        Anchor      = $anchor
        AnchorAt    = $at
        AnchorCount = [int64]($length - $at)
        HighWater   = $length
        TakenAt     = (Get-Date)
    }
}

function Format-LogRolled($mark, $how) {
    return ("The add-in log rolled during this run - $how. The mark was taken at {0:HH:mm:ss} with " -f $mark.TakenAt) +
           ("the file at {0} bytes. src\native\log.cpp:42 DELETES the log once it passes 512KB, so " -f $mark.Length) +
           'everything this suite proves from the log was destroyed mid-run and nothing after this ' +
           'point can be trusted either way. Re-run the suite; tools\check-all.ps1 empties the log ' +
           'before each one, which is what stops this happening.'
}

# Everything the add-in wrote after the mark, as one string. Throws rather than returning '' when the
# evidence has been deleted: "nothing was logged" and "what was logged is gone" are different
# answers, and a suite that merges them passes for having measured something.
function Get-LogTextSince($mark) {
    if ($null -eq $mark) {
        throw 'No log mark has been taken. Call Set-LogMark before reading the log, or the read has no idea which run it is looking at.'
    }
    if (-not (Test-Path $LogPath)) {
        # Never written to at all is a real state and not a failure - but only if this mark has never
        # seen bytes in it either.
        if ($mark.Length -eq 0 -and $mark.HighWater -eq 0) { return '' }
        throw (Format-LogRolled $mark 'the file is not there at all')
    }

    $stream = [System.IO.File]::Open($LogPath, 'Open', 'Read', 'ReadWrite')
    try {
        $now = [int64]$stream.Length
        if ($now -lt $mark.Length) {
            throw (Format-LogRolled $mark ("the file is now {0} bytes, shorter than the mark" -f $now))
        }
        if ($now -lt $mark.HighWater) {
            throw (Format-LogRolled $mark ("the file reached {0} bytes while this mark was in use and is now {1}" -f $mark.HighWater, $now))
        }
        if ($mark.Anchor -ne '') {
            $here = Get-LogBytesHex $stream $mark.AnchorAt $mark.AnchorCount
            if ($here -ne $mark.Anchor) {
                throw (Format-LogRolled $mark ("the {0} bytes the mark was taken after are no longer at offset {1}, so the file grew back past it" -f
                                               $mark.AnchorCount, $mark.AnchorAt))
            }
        }
        $mark.HighWater = $now

        $stream.Seek([int64]$mark.Length, 'Begin') | Out-Null
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
        $text = $reader.ReadToEnd()
    } finally { $stream.Dispose() }
    return $text
}

<#
  Take a mark, or put a stashed one back.

  Returns NOTHING, on purpose. The suites call this as a bare statement, and a function that returned
  the mark would dump a pscustomobject into the middle of their output - which check-all.ps1 then
  scrapes for the summary line. Stash one with Get-CurrentLogMark instead.
#>
function Set-LogMark {
    param([object]$Mark)
    if ($null -ne $Mark) {
        if ($Mark -isnot [psobject] -or -not ($Mark.PSObject.Properties.Name -contains 'Anchor')) {
            throw "Set-LogMark -Mark takes a mark from Get-CurrentLogMark, not a $($Mark.GetType().Name)."
        }
        $script:LogMark = $Mark
        return
    }
    $script:LogMark = New-LogMarkNow
}

# The mark in force, to stash across a section that takes its own. check-dot does this around the
# Protected View window, where the palette assertions need a mark of their own and the section after
# them needs the original back.
function Get-CurrentLogMark { return $script:LogMark }

<#
  Every line logged since the mark that matches $Pattern. `-like "*$Pattern*"`, which is a WILDCARD
  match and not a substring one - `[`, `]`, `*` and `?` are special in it. That is what six of the
  seven copies did and every pattern in the suites is plain text, so nothing changes; it is written
  down because check-menu's copy used Select-String -SimpleMatch, which is a genuine substring, and
  the two only agree while the patterns stay free of those four characters.

  The type check on $Pattern is not defensive dressing. Four suites used to call this as
  `Get-LogSince $mark 'tab moved'`, mark FIRST, and with a [string]$Pattern PowerShell would happily
  coerce the mark object into a string and match nothing - the exact silent-empty-result failure this
  whole primitive exists to remove. Any call site missed by the migration stops the suite instead.
#>
function Get-LogSince {
    [CmdletBinding()]
    param([Parameter(Position = 0)][object]$Pattern)
    if ($null -eq $Pattern) { throw 'Get-LogSince needs a pattern.' }
    if ($Pattern -isnot [string]) {
        throw ("Get-LogSince takes the pattern and nothing else - the mark is the one set by Set-LogMark. " +
               "Got a $($Pattern.GetType().Name), which is the old `Get-LogSince `$mark 'pattern'` order.")
    }
    $text = Get-LogTextSince $script:LogMark
    return @(($text -split "`r?`n") | Where-Object { $_ -like "*$Pattern*" })
}

function Get-LogCount {
    [CmdletBinding()]
    param([Parameter(Position = 0)][object]$Pattern)
    return @(Get-LogSince $Pattern).Count
}

# The LAST line since the mark that matches, or $null. check-menu wants exactly one line - the
# add-in's account of why a batch close stopped - and $null is a real answer there: it means the
# add-in never said anything, which is a different failure from it saying the wrong thing.
function Get-LogLast {
    [CmdletBinding()]
    param([Parameter(Position = 0)][object]$Pattern)
    $lines = @(Get-LogSince $Pattern)
    if ($lines.Count -eq 0) { return $null }
    return $lines[$lines.Count - 1].Trim()
}

<#
  Move the current log aside so what follows starts with an empty one.

  Two things at once: a suite that starts at 0 bytes cannot reach the 512KB roll on its own, and the
  previous suite's log survives instead of being deleted by it. A battery used to leave one file
  holding whatever the last 512KB happened to be; it now leaves one file per suite, which is the
  evidence "the suite that reports a failure is not necessarily the suite that caused it" needs.

  Best-effort by design and it never throws. The add-in opens the log per line and closes it again
  (log.cpp:102), so nothing holds it between writes - but if the move fails anyway, the detectors in
  Get-LogTextSince are still there and the run is still honest. Losing the archive is not worth
  losing the run.
#>
function Reset-LogFile {
    param([string]$Label = 'run', [switch]$Quiet)

    $script:LogMark = $null
    try {
        if (-not (Test-Path $LogPath)) { return $true }
        if (-not (Test-Path $script:LogHistoryDir)) {
            New-Item -ItemType Directory -Path $script:LogHistoryDir -Force | Out-Null
        }
        $safe = ($Label -replace '[^A-Za-z0-9._-]', '-')
        # Milliseconds in the name, not just seconds: two archives in the same second would otherwise
        # collide and Move-Item -Force would silently overwrite the first one's evidence.
        $name = '{0}-{1}.log' -f (Get-Date -Format 'yyyyMMdd-HHmmss-fff'), $safe
        Move-Item -Path $LogPath -Destination (Join-Path $script:LogHistoryDir $name) -Force -ErrorAction Stop

        $old = @(Get-ChildItem -Path $script:LogHistoryDir -Filter '*.log' -ErrorAction SilentlyContinue |
                 Sort-Object LastWriteTime -Descending | Select-Object -Skip $script:LogHistoryKeep)
        foreach ($f in $old) { Remove-Item -Path $f.FullName -Force -ErrorAction SilentlyContinue }
        if (-not $Quiet) { Write-HarnessNote "the add-in log was moved to history\$name - this run starts from an empty one" }
        return $true
    } catch {
        if (-not $Quiet) { Write-HarnessNote "could not move the add-in log aside ($($_.Exception.Message)) - carrying on, the mark still checks for a roll" }
        return $false
    }
}

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

# The add-in's OWN top-level windows, and they are a third category rather than more chrome.
#
# "Word drew this and it is not asking anything" and "WordTab drew this" are different facts, and
# folding ours into the list above would put a window this project controls into a list whose comment
# says every entry was measured coming out of Word.
#
# It is not hypothetical tidiness. The tooltip is the add-in's first top-level window - the strip is a
# child and so was never enumerated here - and it is 917x92 on this rig, comfortably past the size
# floor. Before this existed the classifier called it a `question`, so the first hover in any suite
# made every dialog guard in the harness report that Word was asking something, including the one
# Close-AllWord uses to decide whether it may keep closing windows.
$script:WordTabClasses = @(
    'WordTabTip',      # the hover tooltip: a document's full name and the folder it is in
    'WordTabStrip'     # a child today, but it is ours and the answer must not depend on that
)

function Get-WordWindowKind($class, $width, $height) {
    if ($class -eq 'OpusApp')  { return 'frame' }
    if ($class -eq '#32768')   { return 'menu' }
    if ($script:WordTabClasses -contains $class)    { return 'ours' }
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

<#
  WordTab's own hover tooltip, or $null if none is on screen.

  **This is the only text the add-in draws that can be read from outside the process.** A tab name is
  painted and stored nowhere, which is why check-title has to assert a log line about what was
  *computed* and then photograph the pixels separately to show it was *drawn*. The tooltip puts its
  whole text on the window itself with SetWindowTextW for exactly that reason, so here the drawn
  string can simply be asked for.

  The two lines - the document's name, then the folder it is in - are separated by a newline, and are
  split out here so no caller has to know that.

  **Returns the VISIBLE one, and that distinction matters.** A tooltip window outlives the tooltip:
  one is created on the first hover of a frame and then hidden and re-shown for the rest of the
  session. A test that asked whether the window existed would pass forever after the first hover.
#>
function Get-WordTabTip {
    foreach ($id in @(Get-WordPidList)) {
        foreach ($w in [WordLayout]::TopLevel($id)) {
            if ($w.Class -ne 'WordTabTip') { continue }
            if (-not $w.Visible) { continue }

            $text  = [WordLayout]::TitleOf($w.Hwnd)
            $lines = @($text -split "`n")
            return [pscustomobject]@{
                Hwnd   = $w.Hwnd
                Pid    = $id
                Text   = $text
                Lines  = $lines
                # NOT @(... | Select-Object -First 1). The array wrapper survives the assignment and
                # the property comes back as the string "System.Object[]", which compares unequal to
                # every name there is and reads in a transcript like a tooltip drawing garbage. The
                # `@()` habit that guards .Count elsewhere in this file is wrong on a scalar.
                Name   = $lines | Select-Object -First 1
                Folder = $(if ($lines.Count -gt 1) { $lines[1] } else { '' })
                Left   = $w.Left;  Top    = $w.Top
                Right  = $w.Right; Bottom = $w.Bottom
                Width  = $w.Right - $w.Left
                Height = $w.Bottom - $w.Top
            }
        }
    }
    return $null
}

<#
  Wait for the tooltip to be on screen, or to be gone, and report how long it took.

  Bounded, and it prints the latency, which is the shape the tear-off slice settled on after a sleep
  cost it three false failures: a fixed wait long enough to be safe cannot tell a tooltip that
  arrives on time from one arriving a second late, and "how long did it take" is a measurement worth
  having in the transcript rather than a number nobody sees.

  The default allows for the appearance delay, which is the system's double-click time and is a
  setting the user can change - so it is read rather than assumed.
#>
function Wait-WordTabTip {
    param([bool]$Present, [int]$Seconds = 0, [string]$What = 'the tooltip')

    if ($Seconds -le 0) {
        $Seconds = [int]([math]::Ceiling(([WordLayout]::DoubleClickTime() / 1000.0) + 3))
    }

    # A result object rather than the tooltip itself, because half the calls here are waiting for it
    # to be ABSENT and "$null" would then mean both "it went, as asked" and "it never went". Those are
    # a pass and a failure, and a function that returns the same value for both is the shape of test
    # that reports the wrong thing.
    $started  = Get-Date
    $deadline = $started.AddSeconds($Seconds)
    while ((Get-Date) -lt $deadline) {
        $tip = Get-WordTabTip
        if (($null -ne $tip) -eq $Present) {
            $ms = [int]((Get-Date) - $started).TotalMilliseconds
            Write-HarnessNote ("{0}: {1} after {2}ms" -f $What, $(if ($Present) { 'up' } else { 'gone' }), $ms)
            return [pscustomobject]@{ Ok = $true; Tip = $tip; Ms = $ms }
        }
        Start-Sleep -Milliseconds 100
    }

    $ms = [int]((Get-Date) - $started).TotalMilliseconds
    Write-HarnessNote ("{0}: still {1} after {2}ms" -f $What, $(if ($Present) { 'absent' } else { 'up' }), $ms)
    return [pscustomobject]@{ Ok = $false; Tip = (Get-WordTabTip); Ms = $ms }
}

<#
  What the tab context menu holds, top to bottom, with '-' for a separator.

  Shared because three suites assert the whole list and they assert it about different things:
  check-menu that the menu is built correctly, check-look that owner-drawing did not cost the labels
  their GetMenuString readback. Two copies agree right up until a menu item is added, and then one
  suite goes red for a reason that has nothing to do with what it measures - which is exactly what
  happened when "Move to New Window" went in.

  The list does not depend on how many tabs the row holds, because items are greyed rather than
  hidden. It DOES depend on one thing: a window standing on its own outside the stack is offered a way
  back in, and no other window is. That item is hidden rather than greyed, and deliberately so - on
  every ordinary tab it would be a permanently grey line describing a state that tab is not in - so
  the two shapes are two lists here rather than one list and a footnote.

  -OnItsOwn is the menu on a window that has been torn off.
#>
function Get-TabMenuItems {
    param([switch]$OnItsOwn)

    if ($OnItsOwn) {
        return @('&Save', '&Move to New Window', 'Move &Back to the Tab Row', '-', '&Close',
                 'Close &Others', 'Close Tabs to the &Right', 'Close &All', '-', '&New Document')
    }
    return @('&Save', '&Move to New Window', '-', '&Close', 'Close &Others',
             'Close Tabs to the &Right', 'Close &All', '-', '&New Document')
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
        [switch]$Shift,
        [int]$Tries = 3
    )
    # -Shift is only meaningful with -Ctrl: it exists for Word's own Ctrl+Shift+F6, and building that
    # chord by hand in a caller would be a second copy of the confirmation above it, which is the one
    # thing this module exists to prevent. A bare -Shift is a caller mistake and says so rather than
    # quietly sending an unshifted key.
    if ($Shift -and -not $Ctrl) { throw "Invoke-ConfirmedKey: -Shift needs -Ctrl ($What)." }
    for ($try = 1; $try -le $Tries; $try++) {
        if (Test-WordHasFocus) {
            if ($Ctrl -and $Shift) { [WordLayout]::CtrlShiftPress($Vk) }
            elseif ($Ctrl)         { [WordLayout]::CtrlPress($Vk) }
            else                   { [WordLayout]::Press($Vk) }
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

# ---- the log primitives, proved against a scratch file -----------------------------------------
#
# Run with `pwsh -File tools\WordTabHarness.ps1 -SelfTest`. No Word, no desktop, about a second.
#
# The point of doing it this way: the fault being guarded against needs a 512KB log to appear in a
# real run, which is two whole batteries, and what it produces is SILENCE - assertions that pass
# having read nothing. A scratch file rolls in a millisecond, so every branch that used to be
# unreachable is reachable here, including the one an offset can never see at all: a log that rolled
# and then grew back past the mark.
#
# Under Set-StrictMode -Version Latest, like the suites, because that is where this project's
# empty-collection traps live.

if ($script:HarnessSelfTest) {
    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    $scratch = Join-Path ([System.IO.Path]::GetTempPath()) 'WordTabHarnessSelfTest'
    if (Test-Path $scratch) { Remove-Item -Path $scratch -Recurse -Force }
    New-Item -ItemType Directory -Path $scratch -Force | Out-Null

    # Point the primitives at the scratch file. An `if` block does not make a scope, so these are the
    # same script-scope variables the functions above close over.
    $LogPath = Join-Path $scratch 'wordtab.log'
    $script:LogHistoryDir = Join-Path $scratch 'history'

    $script:selfPass = 0
    $script:selfFail = 0

    function Check($ok, $what) {
        if ($ok) { $script:selfPass++; Write-Host "  PASS  $what" -ForegroundColor Green }
        else     { $script:selfFail++; Write-Host "  FAIL  $what" -ForegroundColor Red }
    }

    # A check that a read REFUSES. Half of these primitives exist to turn a silent empty result into
    # a stopped run, so "it threw" is the assertion, and the message is printed because a throw whose
    # text does not explain itself is not much better than the silence it replaced.
    function CheckThrows([scriptblock]$block, $what) {
        $threw = $false
        $message = ''
        try { & $block | Out-Null } catch { $threw = $true; $message = $_.Exception.Message }
        Check $threw $what
        if ($threw) {
            $short = $message -replace '\s+', ' '
            if ($short.Length -gt 120) { $short = $short.Substring(0, 120) + '...' }
            Write-Host "          $short" -ForegroundColor DarkGray
        }
    }

    function Add-Line($text) {
        [System.IO.File]::AppendAllText($LogPath, "$text`r`n", (New-Object System.Text.UTF8Encoding($false)))
    }
    # What log.cpp:42 does when the file passes 512KB. Not a truncate - a delete.
    function Remove-Log { if (Test-Path $LogPath) { Remove-Item -Path $LogPath -Force } }

    Write-Host ''
    Write-Host 'WordTabHarness self-test: the add-in log, read from a mark' -ForegroundColor Cyan
    Write-Host ''

    # 1. Never written to at all. A real state - the add-in may not have logged yet - and it must not
    #    look like a roll.
    Remove-Log
    Set-LogMark
    Check ((Get-LogCount 'anything') -eq 0) 'a log that has never been written reads as nothing, and does not throw'

    # 2. The ordinary case, both directions. "Nothing since the mark" is the answer that used to be
    #    produced by accident, so it is checked as hard as the positive one.
    Add-Line 'alpha one'
    Add-Line 'alpha two'
    Set-LogMark
    Add-Line 'beta one'
    Add-Line 'beta two'
    Check ((Get-LogCount 'beta') -eq 2)  'both lines written after the mark are found'
    Check ((Get-LogCount 'alpha') -eq 0) 'and neither of the two written before it is'
    Check ((Get-LogLast 'beta') -like '*beta two*') 'Get-LogLast returns the last matching line'
    Check ($null -eq (Get-LogLast 'gamma')) 'and $null when nothing matched, which is not the same as an empty line'

    # 3. Rolled to a shorter file. Four suites used to read this from offset 0 and answer with what
    #    survived; check-title clamped to the end and answered with nothing at all.
    Set-LogMark
    Remove-Log
    Add-Line 'fresh after the roll'
    CheckThrows { Get-LogCount 'alpha' } 'a log that rolled to a SHORTER file stops the run instead of answering'

    # 4. Rolled and grew back PAST the mark. No offset can see this - the file is longer than the
    #    mark, so every length test passes - and it is the whole reason a mark carries bytes.
    Remove-Log
    1..10 | ForEach-Object { Add-Line ("first generation line $_ " + ('x' * 40)) }
    Set-LogMark
    Remove-Log
    1..40 | ForEach-Object { Add-Line ("second generation line $_ " + ('y' * 40)) }
    Check ((Get-Item $LogPath).Length -gt $script:LogMark.Length) 'the regrown file is longer than the mark, so no length test can catch it'
    CheckThrows { Get-LogCount 'first generation' } 'and the anchor bytes catch it'

    # 5. A mark taken at 0 bytes has no anchor to check, so the roll is caught by the high-water mark
    #    of what this mark has already seen. The suites read in Wait-Until loops, so there is almost
    #    always an earlier read to have seen it.
    Remove-Log
    Set-LogMark
    1..40 | ForEach-Object { Add-Line ("zulu line $_ " + ('z' * 40)) }
    Check ((Get-LogCount 'zulu') -eq 40) 'a mark taken at 0 bytes reads what came after it'
    Remove-Log
    Add-Line 'tiny'
    CheckThrows { Get-LogCount 'zulu' } 'and a roll under it is caught by the high-water mark, with no anchor to help'

    # 6. The migration guard. Four suites called this as `Get-LogSince $mark 'pattern'`, and a
    #    [string]$Pattern would have stringified the mark and matched nothing - silently.
    Remove-Log
    Add-Line 'alpha one'
    Set-LogMark
    $stashed = Get-CurrentLogMark
    CheckThrows { Get-LogSince $stashed 'alpha' } 'the old mark-first argument order is an error, not an empty result'
    CheckThrows { Get-LogSince 'alpha' $stashed } 'and so is passing a mark as a second argument'
    # The two above are caught by CmdletBinding before the function body runs. This one is the type
    # check inside it - a migration that dropped the pattern and left the mark - and it is here so
    # that check is not a message nobody has ever seen.
    CheckThrows { Get-LogSince $stashed } 'a mark where the pattern belongs is an error too, rather than matching nothing'

    # 7. Stashing a mark across a section that takes its own. check-dot does this around the
    #    Protected View window.
    Remove-Log
    Add-Line 'before the outer mark'
    Set-LogMark
    $outer = Get-CurrentLogMark
    Add-Line 'between the two marks'
    Set-LogMark
    Add-Line 'after the inner mark'
    Check ((Get-LogCount 'between the two marks') -eq 0) 'the inner mark cannot see what came before it'
    Set-LogMark -Mark $outer
    Check ((Get-LogCount 'between the two marks') -eq 1) 'and putting the outer one back sees it again'
    CheckThrows { Set-LogMark -Mark 244138 } 'a byte offset is not a mark and cannot be restored as one'

    # 8/9. Reset-LogFile: the log moves to history, the live file starts empty, and the mark it
    #      cleared refuses to be read from rather than reading the whole of the new file.
    Remove-Log
    Add-Line 'this suite is finished with the machine'
    $moved = Reset-LogFile -Label 'selftest' -Quiet
    Check $moved 'Reset-LogFile moved the log aside'
    Check (-not (Test-Path $LogPath)) 'the live log is gone, so the next run starts from an empty one'
    $archived = @(Get-ChildItem -Path $script:LogHistoryDir -Filter '*selftest*.log')
    Check ($archived.Count -eq 1) 'and it is in history\, where a failure in the previous suite can still be read'
    if ($archived.Count -eq 1) {
        Check ((Get-Content -Raw -Path $archived[0].FullName) -like '*finished with the machine*') 'with its contents intact'
    }
    CheckThrows { Get-LogCount 'anything' } 'reading with no mark taken is an error, not a read of the whole file'

    # 10. Pruning keeps the newest and not an arbitrary set.
    $script:LogHistoryKeep = 2
    foreach ($n in 1..4) {
        Add-Line "archive $n"
        Reset-LogFile -Label "prune$n" -Quiet | Out-Null
    }
    $left = @(Get-ChildItem -Path $script:LogHistoryDir -Filter '*prune*.log' | Sort-Object Name)
    Check ($left.Count -eq 2) "pruning keeps the newest $script:LogHistoryKeep archives and drops the rest ($($left.Count) left of 4)"
    if ($left.Count -eq 2) {
        Check (($left[0].Name -like '*prune3*') -and ($left[1].Name -like '*prune4*')) 'and it keeps the NEWEST two'
    }

    Remove-Item -Path $scratch -Recurse -Force -ErrorAction SilentlyContinue

    # ---- what a window is, which is the other thing in here that can be wrong silently -----------
    #
    # Get-WordWindowKind is a pure function of a class name and a size, so it can be driven with no
    # Word at all - and it is worth driving, because the way it fails is not a red suite. A window
    # wrongly classified as a `question` makes every dialog guard in the harness report that Word is
    # asking something, and the guards are what stop a run: the whole battery goes red naming a
    # save prompt that does not exist. That is exactly what the tooltip did on its first drive.
    Write-Host ''
    Write-Host 'WordTabHarness self-test: what a top-level window is' -ForegroundColor Cyan
    Write-Host ''

    Check ((Get-WordWindowKind 'OpusApp' 900 700) -eq 'frame')  'a Word frame is a frame'
    Check ((Get-WordWindowKind '#32768' 411 334) -eq 'menu')    'a popup menu is a menu'
    Check ((Get-WordWindowKind 'SysShadow' 416 339) -eq 'chrome') "a menu's shadow is chrome, not a question"
    Check ((Get-WordWindowKind 'NUIDialog' 920 713) -eq 'question') 'the save prompt is still a question'
    Check ((Get-WordWindowKind '#32770' 1440 960) -eq 'question')   'Save As is still a question'

    # The sizes are the ones measured on this rig: the tooltip is well past the 150x60 floor, so
    # nothing but the class name keeps it out of `question`.
    Check ((Get-WordWindowKind 'WordTabTip' 917 92) -eq 'ours')  "WordTab's tooltip is ours, not a question"
    Check ((Get-WordWindowKind 'WordTabStrip' 1400 48) -eq 'ours') "WordTab's strip is ours, not a question"

    $total = $script:selfPass + $script:selfFail
    Write-Host ''
    if ($script:selfFail -eq 0) {
        Write-Host "PASS  $total checks, 0 failures" -ForegroundColor Green
        exit 0
    }
    Write-Host "FAIL  $total checks, $script:selfFail failures" -ForegroundColor Red
    exit 1
}
