<#
.SYNOPSIS
  Measure the name WordTab puts on a tab, and the new "Close Tabs to the Right" command.

.DESCRIPTION
  A tab name is Word's window title with the parts that are not the document removed. Word's title
  carries three things - the document, any state annotations, and the application - and until this
  slice only the application suffix came off, and it came off at the *first* match rather than the
  last. So this suite checks:

    - the document name survives intact, including a document whose own name contains " - Word".
      `Meeting - Wordsmith notes.docx` used to draw as `Meeting`; that is the regression this
      suite exists for, and it lost more of the name than any annotation ever did
    - `Compatibility Mode`, `Read-Only` and `Protected View` come off. All three provoked for real:
      a .doc and an .rtf for the first, the read-only file attribute for the second, and a
      mark-of-the-web for the third
    - a plain .docx is left completely alone - the control, and the assertion that catches a rule
      that fires when it should not
    - a never-saved document still reads `Document1`
    - **the extension is left exactly as Word gives it.** Not a preference: Word follows Explorer's
      "hide extensions for known file types", measured in tools\probe-titles.ps1, so the user has
      already answered this question and Word already listens to the answer
    - `TabTitleTrim=0` gives back the untrimmed name - and, importantly, still fixes the truncation
    - Close Tabs to the Right is greyed on the last tab, closes exactly the tabs to the right of
      the one it was invoked on, and leaves that one and everything left of it alone

  **How a drawn string is measured from outside.** A tab name is painted, never stored in a control,
  so there is nothing to read. The add-in logs the computed name, and this suite asserts against
  that line - but a log line saying what was *computed* is not evidence of what was *drawn*, which
  is exactly the failure this project has already had twice. So it is paired with a photograph: the
  rightmost column of text pixels inside a tab is measured with the trim on and again with it off,
  and the untrimmed name has to reach measurably further right. The log says which string; the
  pixels say the string reached the screen.

  Word is started on fixtures authored by Word itself in %TEMP%\wordtab-titles, plus scratch .rtf
  files for the close-to-the-right section. Everything here is disposable.

  **Every click recomputes where it is aiming immediately before it clicks.** The strip moves.

.PARAMETER KeepOpen
  Leave Word running afterwards.

.EXAMPLE
  pwsh -File tools\check-title.ps1
  pwsh -File tools\check-title.ps1 -Screenshot
#>
[CmdletBinding()]
param(
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
# tools\WordTabHarness.ps1 for why these are not per-suite copies any more. Invoke-StripClick, which
# this suite invented, lives there now and is used by every suite.
. (Join-Path $PSScriptRoot 'WordTabHarness.ps1')

$script:Failures = 0
$script:Checks   = 0

trap {
    Write-Host ''
    Write-Host "ERROR  $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
    break
}

function Write-Step($text) { Write-Host "==> $text" -ForegroundColor Cyan }
function Write-Note($text) { Write-Host "    $text" -ForegroundColor DarkGray }
function Assert($condition, $text) {
    $script:Checks++
    if ($condition) { Write-Host "    PASS  $text" -ForegroundColor Green }
    else { $script:Failures++; Write-Host "    FAIL  $text" -ForegroundColor Red }
}

function Get-WordPids { return Get-WordPidList }
function Get-Frames { return Get-WordFrameList }
# Always through this - a PowerShell function returning an empty array returns *nothing*, and
# `(Get-Frames).Count` is then a hard error under StrictMode at the two moments that matter most.
function Get-FrameCount { return Get-WordFrameTally }

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

# ---- the log, read from a mark ----------------------------------------------------------------
#
# Window handles are reused across Word restarts, and this suite restarts Word twice. So the log is
# never searched from the beginning: a mark is taken, and only what was written after it counts.
# Without that, the TabTitleTrim=0 section can match the line the TabTitleTrim=1 section wrote for a
# handle that has since been recycled, and it would pass by reading its own predecessor's evidence.

# Set-LogMark and Get-LogSince come from WordTabHarness.ps1 now, and the copy that used to live here
# is the reason this whole primitive was hoisted. **It clamped with [Math]::Min($mark, $length)**,
# so once the add-in rolled the log - which it did, in the middle of this very suite, during the
# battery that found it - every read came from the END of the fresh file and returned nothing. Every
# "the log says nothing since the mark" check passed, and every "the log says X" check went red for
# a reason that was not the product.

# The add-in's own account of the name it computed, for one window, since the mark. Returns the
# computed name and the raw window title it came from, so an assertion can show both.
function Get-TabName($frame, $seconds = 20) {
    $key = 'hwnd=0x{0:X16}' -f [int64]$frame
    $deadline = (Get-Date).AddSeconds($seconds)
    while ((Get-Date) -lt $deadline) {
        $lines = @(Get-LogSince 'tab name |' | Where-Object { $_ -like "*$key*" })
        if ($lines.Count -gt 0) {
            $last = $lines[$lines.Count - 1]
            if ($last -match 'tab name \|(?<n>.*)\|  from window title \|(?<r>.*)\|\s*$') {
                return [pscustomobject]@{ Name = $Matches['n']; Raw = $Matches['r'] }
            }
        }
        Start-Sleep -Milliseconds 500
    }
    return $null
}

# Set-WordForeground comes from WordTabHarness.ps1 now, with one guard this copy did not have: it
# never minimises a window belonging to Word itself. This suite provokes Protected View and can meet a
# modal of Word's, and minimising that leaves Word disabled behind a question nobody can see.

function Get-TopStrip {
    Set-WordForeground | Out-Null
    $frames = @(Get-Frames)
    if ($frames.Count -eq 0) { throw 'No Word windows.' }
    $top = [WordLayout]::GetForeground()
    if (-not ($frames -contains $top)) { $top = $frames[0] }
    $parts = Get-Parts $top
    if (-not $parts.Strip) { throw 'The foreground Word window has no WordTab strip.' }
    return $parts
}

function Get-Spot($kind, $index) {
    $top = Get-TopStrip
    $count = (Get-FrameCount)
    $layout = [WordLayout]::Tabs($top.Strip.Hwnd, $count)
    $stripRect = [WordLayout]::RectOf($top.Strip.Hwnd)

    if ($kind -eq 'empty') {
        $x = [int](($layout.Plus.Right + $stripRect.Right) / 2)
        if ($x -ge $stripRect.Right - 2) { $x = $stripRect.Right - 6 }
        return [pscustomobject]@{
            X = $x; Y = [int](($stripRect.Top + $stripRect.Bottom) / 2)
            Rect = $stripRect; Strip = $top.Strip.Hwnd; Count = $count
        }
    }

    if ($index -ge $layout.Tabs.Length) { throw "Tab $index does not exist ($($layout.Tabs.Length) tabs)." }
    $t = $layout.Tabs[$index]
    return [pscustomobject]@{
        X = $t.Left + [int](($t.Right - $t.Left) / 3)
        Y = [int](($t.Top + $t.Bottom) / 2)
        Rect = $t; Strip = $top.Strip.Hwnd; Count = $count
    }
}

function Format-Rect($r) { "({0},{1} {2}x{3})" -f $r.Left, $r.Top, ($r.Right - $r.Left), ($r.Bottom - $r.Top) }

# Click a point on the strip, having first checked that the strip is what is under it.
#
# `point` is a scriptblock that works the coordinate out, and it is re-run on every attempt - both
# because the strip moves between two clicks a second apart, and because taking the foreground back
# can move it again. Never a rectangle measured earlier.
#
# **This exists because a run lost the + to it.** SetForegroundWindow is refused from a process that
# is not already foreground, so a window that comes over Word does not merely steal the keystrokes,
# it receives the clicks - at screen coordinates that were perfectly correct for Word. The failure
# reads as "the + does not work": no new document, and nothing in the add-in's log at all, because
# the message was never posted. Checking the class under the point turns that into a retry.
# Invoke-StripClick now comes from WordTabHarness.ps1, unchanged in behaviour: aim, ask what is under
# the point, click only if it is the strip, and re-run the scriptblock on every attempt because the
# strip moves. It reports what Word has up when a click cannot be placed, which is the usual reason -
# and the class name on its own does not lead anywhere, since a banner over Word reports as "Static",
# a label inside it rather than anything that identifies it.

function Get-MenuWindow { return Get-WordMenu }

function Wait-Menu($present, $seconds = 6) {
    return (Wait-Until { (((Get-WordMenu) -ne [IntPtr]::Zero) -eq $present) } $seconds 200)
}

function Close-Menu {
    if ((Get-MenuWindow) -ne [IntPtr]::Zero) {
        Invoke-ConfirmedKey -Vk 0x1B -What 'Escape to close the menu' | Out-Null
        Wait-Menu $false | Out-Null
    }
}

function Open-TabMenu($index) {
    $owner = (Get-TopStrip).Frame
    # Confirmed onto the strip: a right-click that lands on the document opens WORD's context menu,
    # which is also a visible #32768, so "a menu appeared" would pass while measuring the wrong one.
    $spot = Get-Spot 'label' $index
    Invoke-ConfirmedClick -What "right-clicking tab $index" -Button right -Point { Get-Spot 'label' $index } | Out-Null
    $up = Wait-Menu $true
    if (-not $up) { return [pscustomobject]@{ Window = [IntPtr]::Zero; Items = @(); Spot = $spot } }
    $window = Get-MenuWindow
    return [pscustomobject]@{
        Window = $window
        Items  = @([WordLayout]::MenuItems($window, $owner))
        Spot   = $spot
    }
}

function Wait-Frames($expected, $seconds = 30) {
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

# Get-WordDialog comes from WordTabHarness.ps1 now. The copy that used to live here took anything
# over 200x100 that was not an OpusApp or a menu - and `Net UI Tool Window`, Office`s floating-UI
# host, is exactly that. One appeared at 622x298 while THIS suite was closing seven windows: it
# refused to close Word, left it running, and the two suites after it failed on the leftovers.

# ---- photographs ------------------------------------------------------------------------------
#
# CopyFromScreen, not PrintWindow: this is about what the user can see.

function Get-RectShot($r) {
    $bmp = New-Object System.Drawing.Bitmap(($r.Right - $r.Left), ($r.Bottom - $r.Top))
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($r.Left, $r.Top, 0, 0, $bmp.Size)
    $g.Dispose()
    return [pscustomobject]@{ Bitmap = $bmp; Origin = $r }
}

function Get-Pixel($shot, $x, $y) {
    $px = $x - $shot.Origin.Left
    $py = $y - $shot.Origin.Top
    if ($px -lt 0 -or $py -lt 0 -or $px -ge $shot.Bitmap.Width -or $py -ge $shot.Bitmap.Height) { return $null }
    return $shot.Bitmap.GetPixel($px, $py)
}

function Test-Near($a, $b, $tolerance) {
    if ($null -eq $a -or $null -eq $b) { return $false }
    return ([Math]::Abs($a.R - $b.R) -le $tolerance) -and
           ([Math]::Abs($a.G - $b.G) -le $tolerance) -and
           ([Math]::Abs($a.B - $b.B) -le $tolerance)
}

# The rightmost column of text inside a tab.
#
# The card is one flat colour, so a column that differs from it anywhere down the text band has a
# glyph in it. The card colour is sampled from just inside the left edge, before the text begins -
# the text inset is 12 logical px and this samples at 4, so the sample cannot land on a glyph.
#
# **The scan stops where the renderer stops drawing**, which is `close.left - 4` logical px when the
# tab carries a close button and `tab.right - 10` when it does not - DrawOneTab, same numbers. The
# first version of this scanned to the tab's right edge, ran onto the close button, and answered
# with the x of the *button* for both a trimmed and an untrimmed name: two different strings, one
# identical measurement, and an assertion that could never have failed for the right reason.
#
# **It returns the sampled card colour too, and the caller prints it.** A sampler that answers with a
# bare number cannot say what it measured, and this project has already lost a slice to one that
# could not: a fallback that happens to agree with the truth is not a measurement.
function Measure-TextExtent($shot, $tab, $close, $dpi) {
    $inset = [WordLayout]::Scale(4, $dpi)
    $midY  = [int](($tab.Top + $tab.Bottom) / 2)
    $card  = Get-Pixel $shot ($tab.Left + $inset) $midY
    if ($null -eq $card) { return $null }

    # A band around the vertical centre: tall enough to catch ascenders and descenders, short
    # enough never to reach the card's rounded corners or its border.
    $half = [WordLayout]::Scale(6, $dpi)
    $from = $tab.Left + [WordLayout]::Scale(8, $dpi)
    $to   = if ([WordLayout]::IsEmptyRect($close)) { $tab.Right - [WordLayout]::Scale(10, $dpi) }
            else { $close.Left - [WordLayout]::Scale(4, $dpi) }

    $right = -1
    $columns = 0
    for ($x = $from; $x -lt $to; $x++) {
        $ink = $false
        for ($y = $midY - $half; $y -le $midY + $half; $y++) {
            $c = Get-Pixel $shot $x $y
            if ($null -eq $c) { continue }
            if (-not (Test-Near $c $card 24)) { $ink = $true; break }
        }
        if ($ink) { $right = $x; $columns++ }
    }

    return [pscustomobject]@{
        Right   = $right
        Columns = $columns
        Width   = if ($right -lt 0) { 0 } else { $right - $from }
        Card    = $card
        From    = $from
        To      = $to
    }
}

# ---- Word up and down --------------------------------------------------------------------------

# Word asking the user something, as opposed to Word's own chrome going past.
#
# **Confirmed twice, a second and a half apart.** A single sample got this wrong: a window Word puts
# up while it is shutting down was read as a save prompt, and a run in which every one of the checks
# had passed aborted on its last step claiming there was unsaved work. A question being asked of the
# user is still there when you look again; shutdown chrome is not. Same lesson as
# transient-state-must-be-an-event, from the other side - a guard that can only sample must at least
# refuse to believe one frame of it.
function Get-SavePrompt { return Get-WordSavePrompt }

# What it was, in the message. A guard that aborts a run saying only "Word is asking something" is
# the same mistake as a colour sampler that answers with a bare FALSE: it cannot say what it
# measured, so the next person has to reproduce it before they can even start.
# Format-Dialog comes from WordTabHarness.ps1 now, and says the window`s KIND as well as its class.

# This is the Close-Word whose loose dialog test cost a whole battery run: it read a 622x298
# `Net UI Tool Window` as a question, refused to close Word, and left it running for the next two
# suites. The classification now comes from Get-WordWindowKind, and the "may not answer a question it
# did not raise" rule is enforced by Close-AllWord rather than restated here.
function Close-Word {
    if (@(Get-WordPids).Count -eq 0) { return }
    Write-Note "closing Word: $(Get-FrameCount) window(s)"
    $end = Close-AllWord
    if (-not $end.Closed) {
        if ($end.Reason -eq 'question') {
            throw ("Word is asking something: {0}  Answer it by hand and re-run." -f (Format-WordWindow $end.Dialog))
        }
        throw ("Word would not close ({0}, {1} frame(s) left) - close it by hand and re-run." -f $end.Reason, $end.Frames)
    }
    Start-Sleep -Seconds 2
}

function Open-Document($path, $first) {
    $before = @(Get-Frames)
    Start-Process -FilePath 'winword.exe' -ArgumentList "`"$path`""
    # Waiting for the new window rather than sleeping 16 seconds for the first and 7 for the rest.
    # This is the harness getting into position; nothing in this suite asserts how long Word takes to
    # open a file. It waits for a window that was NOT there before, so a stale frame cannot satisfy it.
    $deadline = (Get-Date).AddSeconds(45)
    while ((Get-Date) -lt $deadline) {
        $new = @(Get-Frames | Where-Object { $before -notcontains $_ })
        if ($new.Count -gt 0) {
            # And for the add-in to have carved its band, because everything here measures the strip.
            #
            # NOT `(Get-Parts $new[0]).Strip`. Get-Parts builds Strip as `@(... | Where-Object ...)[0]`
            # and under Set-StrictMode -Version Latest indexing an empty match is a HARD ERROR, not
            # $null - so a wait for "the strip has appeared" would throw at exactly the moment the
            # strip has not appeared, which is the only moment it is ever called.
            Wait-Until { @([WordLayout]::Children($new[0]) | Where-Object { $_.Class -eq 'WordTabStrip' }).Count -gt 0 } 20 300 | Out-Null
            Start-Sleep -Seconds 2
            return $new[0]
        }
        Start-Sleep -Milliseconds 400
    }
    return [IntPtr]::Zero
}

# ---- fixtures ----------------------------------------------------------------------------------

$Settings = 'HKCU:\Software\WordTab'
$dir = Join-Path $env:TEMP 'wordtab-titles'

Write-Step 'Word must be closed to author the fixtures'
Close-Word

if (Test-Path $dir) { Remove-Item -Path $dir -Recurse -Force }
New-Item -ItemType Directory -Path $dir -Force | Out-Null

# wdFormatXMLDocument = 12, wdFormatDocument97 = 0, wdFormatRTF = 6
#
# These names are SHORT on purpose, and the shortness is load-bearing in two later sections: one draws
# a trimmed name and asserts the ink stops short of the tab's right edge, and its partner turns the trim
# off and asserts the restored annotation pushes that same name out to an ellipsis. Both need a name a
# tab can hold, with the annotation as the only thing that overflows it.
#
# Lengthening them is therefore not a way to make a name be cut, and it was tried: when the label font
# went from 12 to 10 logical px these names started FITTING, which broke the tooltip section below, and
# names near 45 characters fixed that section by breaking five assertions in these two. The cut-name
# path is driven by the annotation and by the single-document section at the end of this file, not by
# the length of a filename.
$fixtures = @(
    @{ Key = 'docx';      File = 'Quarterly report.docx';        Format = 12
       Expect = 'Quarterly report.docx';        Why = 'a plain .docx is left completely alone' }
    @{ Key = 'doc';       File = 'Quarterly report.doc';         Format = 0
       Expect = 'Quarterly report.doc';         Why = 'Compatibility Mode comes off a .doc' }
    @{ Key = 'rtf';       File = 'Quarterly report.rtf';         Format = 6
       Expect = 'Quarterly report.rtf';         Why = 'Compatibility Mode comes off an .rtf' }
    @{ Key = 'readonly';  File = 'Locked report.docx';           Format = 12
       Expect = 'Locked report.docx';           Why = 'Read-Only comes off' }
    @{ Key = 'protected'; File = 'Downloaded report.docx';       Format = 12
       Expect = 'Downloaded report.docx';       Why = 'Protected View comes off' }
    @{ Key = 'dashword';  File = 'Meeting - Wordsmith notes.docx'; Format = 12
       Expect = 'Meeting - Wordsmith notes.docx'
       Why = 'a document whose own name contains " - Word" keeps all of it' }
)

Write-Step 'Authoring the fixtures with Word itself'
$word = New-Object -ComObject Word.Application
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    foreach ($f in $fixtures) {
        $path = Join-Path $dir $f.File
        $doc = $word.Documents.Add()
        $doc.Content.Text = "WordTab title check - $($f.File)"
        # SaveAs2 with plain arguments: PowerShell 7's late-bound COM does not accept [ref] here.
        $doc.SaveAs2($path, [int]$f.Format)
        $doc.Close(0)
        $f.Path = $path
    }
} finally {
    $word.Quit(0)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
Close-Word

$locked = @($fixtures | Where-Object { $_.Key -eq 'readonly' })[0]
Set-ItemProperty -Path $locked.Path -Name IsReadOnly -Value $true

$web = @($fixtures | Where-Object { $_.Key -eq 'protected' })[0]
Set-Content -Path $web.Path -Stream 'Zone.Identifier' -Value "[ZoneTransfer]`r`nZoneId=3"

Write-Note "$($fixtures.Count) fixtures in $dir; read-only attribute and mark-of-the-web applied"

# ---- what each one is called on its tab ---------------------------------------------------------

Write-Step 'Opening each fixture and reading the name the add-in computed for its tab'
Set-LogMark

$first = $true
foreach ($f in $fixtures) {
    $frame = Open-Document $f.Path $first
    $first = $false

    if ($frame -eq [IntPtr]::Zero) {
        Assert $false "$($f.File): a Word window appeared for it"
        continue
    }

    $seen = Get-TabName $frame
    if ($null -eq $seen) {
        Assert $false "$($f.File): the add-in reported a tab name for it"
        continue
    }

    Write-Note ("window title  |{0}|" -f $seen.Raw)
    Write-Note ("tab name      |{0}|" -f $seen.Name)
    Assert ($seen.Name -eq $f.Expect) "$($f.File) -> |$($f.Expect)| - $($f.Why)"
}

# The control that catches a rule firing when it should not: for the plain .docx the tab name and
# the window title must differ by exactly the application suffix and nothing else.
$plain = @($fixtures | Where-Object { $_.Key -eq 'docx' })[0]
$plainFrames = @(Get-Frames | Where-Object { [WordLayout]::TitleOf($_) -like "$($plain.Expect)*" })
if ($plainFrames.Count -gt 0) {
    $seen = Get-TabName $plainFrames[0] 5
    if ($seen) {
        Assert ($seen.Raw -eq ($seen.Name + ' - Word')) `
            'and on the plain .docx the only difference from Word''s title is " - Word"'
    } else { Assert $false 'the plain .docx still reports a tab name' }
} else { Assert $false 'the plain .docx window is still open' }

# ---- every strip still sits in its own window's chrome -------------------------------------------
#
# **This suite has the fixture that reaches the bug and never had the assertion, and that is the
# whole reason this block exists.** A Protected View window's chrome is 38px SHORTER than a normal
# window's - its reduced ribbon and message bar put the document frame at 318 where a normal
# window's is 356, measured - and the stack used to broadcast one window's whole interior onto every
# other. So every other window carved its 32px band 38px too high, inside the ribbon's NetUIHWND.
#
# check-stack owns this check but has no mixed-chrome fixture; the two suites that do - this one and
# check-dot - never ran it. The bug therefore surfaced HERE, twice, as "the + click did nothing":
# three steps from the cause, because the + was being clicked at a computed centre that had landed
# inside the ribbon. It is Get-StripPlacement in tools\WordTabHarness.ps1 now.
#
# Six windows are open at this point and one of them is in Protected View, which is exactly the mix
# that used to break. `Measured` is asserted as well as `Ok`, because "every strip is placed
# correctly" and "there were no strips to look at" are the same answer otherwise.

Write-Step 'Every strip sits between its own window''s chrome and its own document'
$openFrames = @(Get-Frames)
$placement  = Get-StripPlacement $openFrames
Write-Note "$($placement.Measured) of $($openFrames.Count) windows measured, $($placement.Count) fault(s)"
Assert ($openFrames.Count -eq $fixtures.Count) "all $($fixtures.Count) fixtures still have a window ($($openFrames.Count))"
Assert ($placement.Measured -eq $openFrames.Count) `
    "every one of the $($openFrames.Count) windows could be measured (measured $($placement.Measured))"
Assert $placement.Ok `
    ('every strip sits between the chrome and the document, with a Protected View window in the row' + $placement.Text)

# ---- a never-saved document ---------------------------------------------------------------------

Write-Step 'A document that has never been saved'
$before = @(Get-Frames)
$count = (Get-FrameCount)
$landed = Invoke-StripClick 'the +' {
    $top = Get-TopStrip
    $layout = [WordLayout]::Tabs($top.Strip.Hwnd, (Get-FrameCount))
    [WordLayout]::Center($layout.Plus)
}
Assert $landed 'the + could be clicked with the strip actually under the pointer'
Assert (Wait-Frames ($count + 1) 30) "the + opened a new document ($((Get-FrameCount)) windows)"

$new = @(Get-Frames | Where-Object { $before -notcontains $_ })
if ($new.Count -gt 0) {
    $seen = Get-TabName $new[0]
    if ($seen) {
        Write-Note ("window title  |{0}|" -f $seen.Raw)
        Write-Note ("tab name      |{0}|" -f $seen.Name)
        Assert ($seen.Name -like 'Document*') "an unsaved document keeps Word's own name (|$($seen.Name)|)"
        Assert ($seen.Name -notlike '*Word*') 'and nothing of the application suffix is left on it'
    } else { Assert $false 'the add-in reported a tab name for the new document' }
} else { Assert $false 'the new document has a window of its own' }

if ($Screenshot) {
    $shot = Get-RectShot ([WordLayout]::RectOf((Get-TopStrip).Strip.Hwnd))
    $shot.Bitmap.Save((Join-Path $ShotDir 'wordtab-title-row.png'), [System.Drawing.Imaging.ImageFormat]::Png)
    $shot.Bitmap.Dispose()
    Write-Note "row photographed to $(Join-Path $ShotDir 'wordtab-title-row.png')"
}

# ---- the tooltip -----------------------------------------------------------------------------
#
# **This suite owns the tooltip because the tooltip is about the name**, and everything that knows
# what a tab is called is already here. It is also the only suite that has the three cases in one
# room at once, without authoring a single new fixture:
#
#   - six saved documents in a folder, in a row narrow enough that their names are cut. That is the
#     whole reason the feature exists, and it is not a contrivance: this is Word's own default window
#     with seven documents in it
#   - a document in **Protected View**, which lives in no collection this add-in can reach, so Word
#     answers "no folder" for it. The tooltip is then one line rather than two, and the difference is
#     measurable as a height
#   - a document that has **never been saved**, whose name fits and which has no folder - so there is
#     nothing the tooltip could say that is not already on the tab, and it does not appear at all
#
# **Placed here, upstream of Close Tabs to the Right, rather than appended at the end.** Two slices
# running have now found a real defect that way and none has been found by a section that ran last.
# Everything below this point keeps using the same Word.
#
# **Every tab is hovered and the answers are matched by name afterwards**, rather than assuming which
# tab is which. The row order is the open order today; an assertion that depends on that is an
# assertion about the stack's member array, which is check-reorder's subject and not this one's.

Write-Step 'The tooltip: hovering every tab in the row'

# **The row must not be scrolled, and this suite does not otherwise care how wide Word is - which is
# exactly why it has to say so here.**
#
# Seven documents overflow a 900px window: the tabs go to their 70 logical px minimum and the row
# starts scrolling, so only the last five are on screen and the computed slots for the other two land
# on the button cluster. Standalone this suite passes with the window Word happens to have; inside the
# battery it inherits whatever the previous suite left, and **Word persists window placement on
# exit**. The first run of this section found that difference the hard way: five tooltips instead of
# seven, and two assertions failing about documents that were simply not on screen.
#
# So the width is made a fixture rather than an inheritance, and put back afterwards. 760 logical px
# is comfortably more than seven tabs at their minimum plus the chevrons and the +.
$tipFrame = (Get-TopStrip).Frame
$tipDpi   = [WordLayout]::Dpi($tipFrame)
$tipGeom  = [WordLayout]::RectOf($tipFrame)
$tipWide  = [WordLayout]::Scale(760, $tipDpi)
$tipTall  = [WordLayout]::Scale(500, $tipDpi)

try {

if (($tipGeom.Right - $tipGeom.Left) -lt $tipWide) {
    Write-Note ("widening Word from {0}px to {1}px so seven tabs do not scroll" -f
                ($tipGeom.Right - $tipGeom.Left), $tipWide)
    [WordLayout]::Resize($tipFrame, $tipWide, [Math]::Max($tipTall, $tipGeom.Bottom - $tipGeom.Top))
    Start-Sleep -Seconds 2
}

$stripRect = [WordLayout]::RectOf((Get-TopStrip).Strip.Hwnd)

# Park the pointer clear of the row between hovers, so each one is a fresh arrival at a tab rather
# than a panel that was already up. Below the strip, which is the document - never on the + or a
# chevron, which are also "not a tab" but are places a press would do something.
function Move-OffTheRow {
    $r = [WordLayout]::RectOf((Get-TopStrip).Strip.Hwnd)
    $ok = Set-Pointer ([int](($r.Left + $r.Right) / 2)) ([int]($r.Bottom + 120)) 'off the row'
    Start-Sleep -Milliseconds 300
    return $ok
}

Set-LogMark
$tabCount = (Get-FrameCount)
Assert ($tabCount -eq ($fixtures.Count + 1)) `
    "the row holds all $($fixtures.Count) fixtures plus the unsaved document ($tabCount tabs)"

$hovered = @()
for ($i = 0; $i -lt $tabCount; $i++) {
    Assert (Move-OffTheRow) "the pointer could be parked off the row before tab $i"

    $spot = Get-Spot 'tab' $i
    $onTab = Set-Pointer $spot.X $spot.Y "tab $i"
    Assert $onTab "the pointer reached tab $i at ($($spot.X),$($spot.Y))"
    if (-not $onTab) { continue }

    # A tooltip that is absent is a real answer here - one of the tabs is supposed to produce none -
    # so this waits for either outcome and records which, rather than failing on the wait.
    $r = Wait-WordTabTip -Present $true -What "tab $i"
    $hovered += [pscustomobject]@{
        Index = $i; Tip = $r.Tip; Ms = $r.Ms
        Rect  = $spot.Rect
    }
    if ($r.Ok) {
        Write-Note ("tab {0}: |{1}| / |{2}|  {3}x{4} at {5},{6}  after {7}ms" -f
                    $i, $r.Tip.Name, $r.Tip.Folder, $r.Tip.Width, $r.Tip.Height,
                    $r.Tip.Left, $r.Tip.Top, $r.Ms)
    } else {
        Write-Note ("tab {0}: no tooltip after {1}ms" -f $i, $r.Ms)
    }
}

$shown = @($hovered | Where-Object { $null -ne $_.Tip })
Write-Note "$($shown.Count) of $tabCount tabs produced a tooltip"

# Every SAVED fixture has something to add - a name the row had to cut, and a folder - so every one of
# them must produce a panel. A fixture tab that produced nothing means a slot that was not a tab, which
# is what a scrolled row looks like from outside, and is the failure this section's width fixture exists
# to prevent.
#
# The unsaved document is not in that count, and this used to say $tabCount - all seven. It passed only
# because `Document1` was being cut too, which was true at the width of the day and stopped being true
# when the label font got smaller. Nine characters is not a name a tab cannot fit, the document has no
# folder to name, so there is nothing for a panel to say and none appears. The loop above already
# expects that ("a tooltip that is absent is a real answer here"); this line was the one place that
# disagreed with it, and the count made the disagreement look like a product defect.
Assert ($shown.Count -eq $fixtures.Count) `
    "each of the $($fixtures.Count) saved fixtures produced a panel, and the unsaved document did not ($($shown.Count))"

# ---- what it says --------------------------------------------------------------------------------
#
# The name on the panel is asserted against the SAME expected string the tab-name section above
# asserts against the log. That is the point of the feature: what the tab could not fit, in full.
#
# Every saved fixture, not one of them: they are all in the row, they all have the same folder, and
# checking one would pass on a build that could only ever describe the tab it was pointed at.

$savedFixtures = @($fixtures | Where-Object { $_.Key -ne 'protected' })
foreach ($f in $savedFixtures) {
    $t = @($shown | Where-Object { $_.Tip.Name -eq $f.Expect }) | Select-Object -First 1
    Assert ($null -ne $t) "a tooltip named it in full: |$($f.Expect)|"
    if ($t) {
        Assert ($t.Tip.Folder -eq $dir) "  and gave the folder it is in (|$($t.Tip.Folder)|)"
    }
}

$plainFix = @($fixtures | Where-Object { $_.Key -eq 'docx' }) | Select-Object -First 1
$plainTip = @($shown | Where-Object { $_.Tip.Name -eq $plainFix.Expect }) | Select-Object -First 1

if ($plainTip) {
    Assert ($plainTip.Tip.Name.Length -gt 0 -and $plainTip.Tip.Name -notlike '*...*') `
        'and the name on the panel carries no ellipsis of its own'

    # The panel hangs below the strip, and it is left-aligned with the tab it describes unless the
    # monitor's edge moved it. Asserted as "not above the strip" plus "not further left than the
    # tab", which is what survives the clamp - a panel pushed left by the screen edge is correct
    # behaviour and an assertion on an exact x would call it a failure.
    Assert ($plainTip.Tip.Top -ge $stripRect.Bottom) `
        "it hangs below the tab row (panel top $($plainTip.Tip.Top), strip bottom $($stripRect.Bottom))"
    Assert ($plainTip.Tip.Left -le $plainTip.Rect.Left + 4) `
        "and it starts no further right than the tab it is about (panel $($plainTip.Tip.Left), tab $($plainTip.Rect.Left))"

    # No taskbar button and no Alt+Tab entry. This project spent a slice making Word one taskbar
    # button; a panel that appears on a hover is not allowed to add a second.
    Assert ([WordLayout]::IsToolWindow($plainTip.Tip.Hwnd)) `
        'the panel is a tool window, so it can never reach the taskbar or Alt+Tab'
}

# ---- Protected View: Word will not say where it is -------------------------------------------------
#
# A Protected View document is in none of this Application's collections, so the folder comes back
# empty - a determinate answer about a document this add-in cannot reach, not a failure to read one.
# The tooltip is then the name alone, and the honest test of "one line rather than two" is that the
# panel is SHORTER: a folder line that was drawn empty would leave the height unchanged.

$pvFix = @($fixtures | Where-Object { $_.Key -eq 'protected' }) | Select-Object -First 1
$pvTip = @($shown | Where-Object { $_.Tip.Name -eq $pvFix.Expect }) | Select-Object -First 1

Assert ($null -ne $pvTip) "the Protected View document still gets a tooltip with its full name: |$($pvFix.Expect)|"
if ($pvTip) {
    Assert ($pvTip.Tip.Folder -eq '') 'and no folder line, because Word will not place that document'
    Assert ($pvTip.Tip.Lines.Count -eq 1) 'so the panel carries one line of text, not two'
    if ($plainTip) {
        Assert ($pvTip.Tip.Height -lt $plainTip.Tip.Height) `
            "and it is measurably shorter than the two-line one ($($pvTip.Tip.Height) vs $($plainTip.Tip.Height))"
    }
}

# Every tooltip that DID appear was on a tab whose name had been cut, and the add-in says so. That is
# one of the two reasons a panel appears; the other one - a folder the row could not have shown at any
# width - is driven at the end of this suite, where a single document has a tab to itself.
#
# **The case where the panel stays away is driven there too, and deliberately not here.** Whether a
# name fits is a fact about the width of the row, and the width of the row is a fact about the window
# - so an assertion here that `Document1` produces nothing would be an assertion about how wide Word
# happened to be. It failed for exactly that reason on its first run inside the battery: at 900px the
# name was cut, the panel was right to appear, and the suite was wrong to say it should not.
#
# It used to assert the counts were EQUAL, and that held only while every name in the row happened to be
# cut. The label font going from 12 to 10 made these names fit, the panels were still right - they name a
# folder the row never shows - and the equality turned a correct product into five red lines. What is
# actually true is that a panel must have a reason, so that is what is asserted: every one of them names
# a folder, and no tab was logged as cut without a panel to show for it.
$cut = @(Get-LogSince 'name was cut')
#
# Every one of them except the Protected View document, which is the one fixture Word will not place -
# asserted directly above, where its missing folder line is the point rather than an exception to be
# worked around. Its panel earns its place by naming the document at all, which is the only thing that
# can be said about a file the object model will not admit to holding.
$noReason = @($shown | Where-Object { $_.Tip.Folder -eq '' -and $_.Tip.Name -ne $pvFix.Expect })
Assert ($noReason.Count -eq 0) `
    "each panel had something to add - the folder, or in Protected View's case the name ($($noReason.Count) with nothing)"
Assert ($cut.Count -le $shown.Count) `
    "no tab was cut without a panel to say so ($($cut.Count) cut, $($shown.Count) panels)"

# ---- and it is not a dialog ------------------------------------------------------------------------
#
# **The regression test for the thing this feature broke on its first drive.** The tooltip is the
# add-in's first top-level window, it is bigger than the harness's size floor, and the dialog
# classifier is a denylist of Word's own chrome classes - so it called the panel a `question`. Every
# guard in the harness stops a run when Word is asking something, including the one Close-AllWord
# uses, so the first hover in any suite would have reported a save prompt that did not exist.

Write-Step 'The tooltip is not a question'
Assert (Move-OffTheRow) 'the pointer could be parked off the row'
$spot = Get-Spot 'tab' 0
Assert (Set-Pointer $spot.X $spot.Y 'tab 0') 'the pointer reached tab 0 again'
$again = Wait-WordTabTip -Present $true -What 'tab 0'
Assert $again.Ok 'the tooltip came back on a second hover of the same tab'
if ($again.Ok) {
    $dialog = Get-WordDialog
    Assert ($null -eq $dialog) `
        "with the panel on screen the harness sees no question ($(Format-Dialog $dialog))"
    $kinds = @(Get-WordWindows | Where-Object { $_.Class -eq 'WordTabTip' })
    Assert ($kinds.Count -ge 1) 'the panel is enumerated as a top-level window of Word''s process'
    if ($kinds.Count -ge 1) {
        Assert ($kinds[0].Kind -eq 'ours') "and it is classified as ours, not as a question ($($kinds[0].Kind))"
    }
}

# ---- and it goes away ------------------------------------------------------------------------------
#
# Moving off the row takes it down, and that is a WM_MOUSELEAVE rather than a timer: the auto-hide is
# ten times the appearance delay, so a panel that only went on the timer would still be on screen
# here and this would pass for the wrong reason. The latency printed is the evidence for which of the
# two it was.

Write-Step 'The tooltip goes when the pointer does'
Assert (Move-OffTheRow) 'the pointer left the row'
$gone = Wait-WordTabTip -Present $false -What 'the tooltip'
Assert $gone.Ok 'the panel went when the pointer left the tab'
if ($gone.Ok) {
    Assert ($gone.Ms -lt ([WordLayout]::DoubleClickTime() * 5)) `
        "and it went on the pointer leaving, not on the auto-hide ($($gone.Ms)ms, auto-hide is $([WordLayout]::DoubleClickTime() * 10)ms)"
}

} finally {
    # Word persists its window placement on exit, so a section that widened the window and left it
    # widened would hand the next suite in the battery a Word it was not written against. The same
    # rule, and the same `finally`, as the maximized tear-off section in check-reorder.
    if ([WordLayout]::IsWindow2($tipFrame)) {
        [WordLayout]::Resize($tipFrame, ($tipGeom.Right - $tipGeom.Left), ($tipGeom.Bottom - $tipGeom.Top))
        Write-Note ("Word put back to {0}x{1}" -f ($tipGeom.Right - $tipGeom.Left), ($tipGeom.Bottom - $tipGeom.Top))
    }
}

# ---- Close Tabs to the Right ---------------------------------------------------------------------
#
# Its own documents, because the section above has left a mixture of formats and one unsaved
# document, and a batch close that meets a save prompt is check-menu.ps1's subject rather than this
# one's. Four .rtf files, opened in order, so the row reads 1 2 3 4 left to right.

Write-Step 'Close Tabs to the Right: four fresh documents'
Close-Word

$scratch = Join-Path $env:TEMP 'wordtab-check'
New-Item -ItemType Directory -Path $scratch -Force | Out-Null
$right = @()
for ($i = 1; $i -le 4; $i++) {
    $path = Join-Path $scratch "wordtab-title-$i.rtf"
    "{\rtf1\ansi WordTab title check - document $i.\par}" | Set-Content -Path $path -Encoding Ascii
    $right += $path
}

Set-LogMark
$first = $true
foreach ($path in $right) {
    Open-Document $path $first | Out-Null
    $first = $false
}
Assert (Wait-Frames 4 40) "four documents are open ($((Get-FrameCount)) windows)"

# The row order is the open order, so tab 3 is document 4 - the last one.
Write-Step 'The menu on the last tab'
$menu = Open-TabMenu 3
Assert ($menu.Window -ne [IntPtr]::Zero) 'the menu opened on the last tab'
$item = @($menu.Items | Where-Object { $_.Text -eq 'Close Tabs to the &Right' })
Assert ($item.Count -eq 1) 'Close Tabs to the Right is on the menu'
if ($item.Count -eq 1) {
    Assert (-not $item[0].Enabled) 'and it is greyed on the last tab - there is nothing to its right'
}
$others = @($menu.Items | Where-Object { $_.Text -eq 'Close &Others' })
Assert ($others.Count -eq 1 -and $others[0].Enabled) 'Close Others is still available beside it'
Close-Menu

Write-Step 'The menu on the second tab'
$menu = Open-TabMenu 1
Assert ($menu.Window -ne [IntPtr]::Zero) 'the menu opened on the second tab'
$item = @($menu.Items | Where-Object { $_.Text -eq 'Close Tabs to the &Right' })
Assert ($item.Count -eq 1 -and $item[0].Enabled) 'Close Tabs to the Right is available with two tabs to the right'

if ($item.Count -eq 1 -and $item[0].HasRect) {
    $centre = [WordLayout]::Center($item[0].Rect)
    Write-Note ("clicking it at {0}" -f (Format-Rect $item[0].Rect))
    # Expecting #32768, the popup menu itself. Set-WordForeground inside the confirmed click returns
    # immediately while a menu is up, so this cannot dismiss the menu it is aiming at.
    Assert (Invoke-ConfirmedClick -What 'clicking Close Tabs to the Right' -Expect '#32768' `
                                  -Point { [pscustomobject]@{ X = $centre.X; Y = $centre.Y } }) `
           'the click landed on the menu'
} else {
    Write-Note 'no rectangle to click - using the access key'
    Invoke-ConfirmedKey -Vk 0x52 -What "the menu's Close-to-the-Right mnemonic" | Out-Null
}
Wait-Menu $false | Out-Null

$went = Wait-Frames 2 45
Assert $went "it closed the two tabs to the right and stopped ($((Get-FrameCount)) windows)"

$left = @(Get-Frames | ForEach-Object { [WordLayout]::TitleOf($_) } | Sort-Object)
Write-Note ("still open: {0}" -f ($left -join ', '))
Assert (@($left | Where-Object { $_ -like 'wordtab-title-1*' }).Count -eq 1) 'document 1 is still open'
Assert (@($left | Where-Object { $_ -like 'wordtab-title-2*' }).Count -eq 1) 'document 2 - the one it was invoked on - is still open'
Assert (@($left | Where-Object { $_ -like 'wordtab-title-3*' }).Count -eq 0) 'document 3 was closed'
Assert (@($left | Where-Object { $_ -like 'wordtab-title-4*' }).Count -eq 0) 'document 4 was closed'
Assert ((Get-WordDialog) -eq $null) 'and nothing was left asking the user anything'

$batch = @(Get-LogSince 'close to the right')
foreach ($line in $batch) { Write-Note $line }
Assert (@($batch | Where-Object { $_ -like '*2 tab(s) queued*' }).Count -ge 1) `
    'the add-in queued exactly two - the count it logged agrees with the row'

# ---- the tab name reaches the screen -------------------------------------------------------------
#
# Everything above is the add-in's own account of what it computed. This is the part that photographs
# what was drawn. One document, so the tab is at its full 220 logical px and the trimmed name has
# room to spare - which is what makes the two measurements distinguishable rather than both being an
# ellipsis at the close button.

Write-Step 'What was actually painted, with the trim on'
Close-Word
Set-LogMark
$doc = @($fixtures | Where-Object { $_.Key -eq 'doc' })[0]
$frame = Open-Document $doc.Path $true
Assert ($frame -ne [IntPtr]::Zero) 'the .doc opened on its own'

Set-WordForeground | Out-Null
Start-Sleep -Seconds 2
$top = Get-TopStrip
$dpi = [WordLayout]::Dpi($top.Strip.Hwnd)
$layout = [WordLayout]::Tabs($top.Strip.Hwnd, 1)
$shot = Get-RectShot ([WordLayout]::RectOf($top.Strip.Hwnd))
$trimmed = Measure-TextExtent $shot $layout.Tabs[0] $layout.Close[0] $dpi
if ($Screenshot) {
    $shot.Bitmap.Save((Join-Path $ShotDir 'wordtab-title-trimmed.png'), [System.Drawing.Imaging.ImageFormat]::Png)
}
$shot.Bitmap.Dispose()

Assert ($null -ne $trimmed) 'the tab was photographed'
if ($trimmed) {
    Write-Note ("card RGB({0},{1},{2}); text runs to x={3}, {4} columns of it, tab is {5}" -f
                $trimmed.Card.R, $trimmed.Card.G, $trimmed.Card.B, $trimmed.Right, $trimmed.Columns,
                (Format-Rect $layout.Tabs[0]))
    Assert ($trimmed.Columns -gt 20) 'there is text painted on the tab at all'
    Assert ($trimmed.Right -lt ($trimmed.To - [WordLayout]::Scale(20, $dpi))) `
        'and it stops well short of the right-hand end - the trimmed name is not an ellipsis'
}

Write-Step 'The same tab with TabTitleTrim=0'
$restore = $null
try {
    $existing = Get-ItemProperty -Path $Settings -Name 'TabTitleTrim' -ErrorAction SilentlyContinue
    if ($existing -and $existing.PSObject.Properties.Name -contains 'TabTitleTrim') {
        $restore = [int]$existing.TabTitleTrim
    }

    Close-Word

    # **Never `New-Item -Force` here.** On a registry key that already exists, -Force *recreates* it
    # and every value under it is gone - the registry twin of `New-Item -ItemType File -Force`
    # truncating a file. This wiped the whole switch key on a run of this suite, and the visible
    # symptom was three sections later and nothing to do with the registry: `ShowLoadBanner` went
    # back to its default, so the add-in's own banner came up over Word, and the clicks aimed at the
    # + landed on a Static control belonging to that dialog. The suite reported "the + does not
    # work". Create the key only when it is not there.
    if (-not (Test-Path $Settings)) { New-Item -Path $Settings | Out-Null }
    Set-ItemProperty -Path $Settings -Name 'TabTitleTrim' -Value 0 -Type DWord
    Write-Note 'TabTitleTrim=0 - the tab should now read what Word''s title bar reads, minus " - Word"'

    Set-LogMark
    $frame = Open-Document $doc.Path $true
    Assert ($frame -ne [IntPtr]::Zero) 'the .doc opened again'

    $seen = Get-TabName $frame
    if ($seen) {
        Write-Note ("tab name      |{0}|" -f $seen.Name)
        Assert ($seen.Name -eq 'Quarterly report.doc  -  Compatibility Mode') `
            'the annotation is back on the name'
        # The half of the rule that is a defect fix rather than a policy stays fixed either way.
        Assert ($seen.Name -notlike '* - Word') 'but the application suffix still comes off'
    } else { Assert $false 'the add-in reported a tab name with the trim off' }

    Set-WordForeground | Out-Null
    Start-Sleep -Seconds 2
    $top = Get-TopStrip
    $layout = [WordLayout]::Tabs($top.Strip.Hwnd, 1)
    $shot = Get-RectShot ([WordLayout]::RectOf($top.Strip.Hwnd))
    $untrimmed = Measure-TextExtent $shot $layout.Tabs[0] $layout.Close[0] $dpi
    if ($Screenshot) {
        $shot.Bitmap.Save((Join-Path $ShotDir 'wordtab-title-untrimmed.png'), [System.Drawing.Imaging.ImageFormat]::Png)
    }
    $shot.Bitmap.Dispose()

    if ($untrimmed -and $trimmed) {
        Write-Note ("text now runs to x={0} of {1}, {2} columns (was x={3}, {4})" -f
                    $untrimmed.Right, $untrimmed.To, $untrimmed.Columns, $trimmed.Right, $trimmed.Columns)
        Assert ($untrimmed.Right -gt $trimmed.Right + [WordLayout]::Scale(10, $dpi)) `
            'the untrimmed name reaches measurably further right - the log line describes real pixels'
        Assert ($untrimmed.Right -ge ($untrimmed.To - [WordLayout]::Scale(12, $dpi))) `
            'and it now runs the full width to an ellipsis, which is what the trim was buying back'
        Assert ($untrimmed.Columns -gt $trimmed.Columns) 'and there is more ink on the tab'
    } else {
        Assert $false 'both tabs were photographed'
    }
} finally {
    if ($null -eq $restore) {
        Remove-ItemProperty -Path $Settings -Name 'TabTitleTrim' -ErrorAction SilentlyContinue
        Write-Note 'TabTitleTrim removed - back to the default'
    } else {
        Set-ItemProperty -Path $Settings -Name 'TabTitleTrim' -Value $restore -Type DWord
        Write-Note "TabTitleTrim put back to $restore"
    }
}

# ---- the tooltip on its own switch -----------------------------------------------------------------
#
# One document, so the tab is at its full 220 logical px and the name fits with room to spare. That
# makes this a different branch from the section above, where every panel was raised because a name
# had been cut: here there is nothing wrong with the tab at all, and the panel appears because the
# FOLDER is something the row cannot say at any width. Both halves of "is there anything to add" are
# therefore driven, and by fixtures that were going to exist anyway.
#
# Then the switch off, positively: not "nothing appeared" - a hover that never landed produces that
# too - but the pointer confirmed onto the same tab, the startup line asserted to say the switch is
# off, and the add-in's log asserted to contain nothing about a tooltip at all. TabTip is read in
# StripStart, so it costs a Word restart.

Write-Step 'The tooltip with TabTip at its default'
Close-Word
Set-LogMark
$frame = Open-Document $doc.Path $true
Assert ($frame -ne [IntPtr]::Zero) 'the .doc opened on its own again'
Start-Sleep -Seconds 1

$tipSpot = Get-Spot 'tab' 0
Assert (Set-Pointer $tipSpot.X $tipSpot.Y 'the only tab') 'the pointer reached the only tab'
$one = Wait-WordTabTip -Present $true -What 'the only tab'
Assert $one.Ok 'a tooltip appeared on a tab whose name fits perfectly well'
if ($one.Ok) {
    Write-Note ("|{0}| / |{1}|" -f $one.Tip.Name, $one.Tip.Folder)
    Assert ($one.Tip.Folder -eq $dir) 'and what it adds is the folder - the thing no tab width could show'
    $fits = @(Get-LogSince 'name fits')
    Assert ($fits.Count -ge 1) 'the add-in agrees the name was not cut, so the folder is the whole reason it appeared'
}

# ---- and the one that says nothing ------------------------------------------------------------------
#
# The other half of the same rule, and it is driven HERE rather than up in the seven-tab row because
# it is the half that depends on the tab being wide. Two tabs in one window is the widest arrangement
# this suite can produce that still has a never-saved document in it, and `Document1` fits in a tab
# that size on any window Word will open. Up in the crowded row its name was cut, the panel was right
# to appear, and an assertion that it should not have was an assertion about the window rather than
# about the product.
#
# Both halves come from the add-in's own account as well as from the screen, because "no tooltip" is
# also what a hover that never landed produces - and the hover itself is asserted separately, which
# is the other half of that guard.

Write-Step 'A never-saved document beside it, on a tab with room to spare'
Set-LogMark
$before = @(Get-Frames)
Assert (Invoke-StripClick 'the +' {
    $top = Get-TopStrip
    $layout = [WordLayout]::Tabs($top.Strip.Hwnd, (Get-FrameCount))
    [WordLayout]::Center($layout.Plus)
}) 'the + could be clicked'
Assert (Wait-Frames 2 30) "a second document opened ($((Get-FrameCount)) windows)"

$fresh = @(Get-Frames | Where-Object { $before -notcontains $_ }) | Select-Object -First 1
if ($fresh) {
    $freshIndex = -1
    $top = Get-TopStrip
    $layout = [WordLayout]::Tabs($top.Strip.Hwnd, (Get-FrameCount))
    # Which slot the new document is in: the row order is the open order, so it is the last one -
    # but that is asserted by hovering and reading the panel rather than assumed, and the never-saved
    # document is the one tab that produces no panel at all, so it is found by elimination.
    for ($i = 0; $i -lt $layout.Tabs.Length; $i++) {
        $spot = Get-Spot 'tab' $i
        Assert (Set-Pointer $spot.X $spot.Y "tab $i") "the pointer reached tab $i"
        $r = Wait-WordTabTip -Present $true -What "tab $i"
        if (-not $r.Ok) { $freshIndex = $i; break }
        Write-Note ("tab {0}: |{1}| / |{2}|" -f $i, $r.Tip.Name, $r.Tip.Folder)
        Assert (Move-OffTheRow) "the pointer left the row after tab $i"
    }

    Assert ($freshIndex -ge 0) 'one of the two tabs produces no panel at all'
    $quiet = @(Get-LogSince 'nothing to add')
    Assert ($quiet.Count -ge 1) 'and the add-in says why it stayed away'
    if ($quiet.Count -ge 1) {
        $last = $quiet[$quiet.Count - 1]
        Write-Note $last
        Assert ($last -like '*Document*') 'the tab it stayed away from is the never-saved document'
        Assert ($last -like '*name fits*') 'because the name fitted'
        Assert ($last -like '*never saved*') 'and there is no folder to add - Word answered, it has none'
    }
} else {
    Assert $false 'the new document has a window of its own'
}

Write-Step 'The same hover with TabTip=0'
$restoreTip = $null
try {
    $existing = Get-ItemProperty -Path $Settings -Name 'TabTip' -ErrorAction SilentlyContinue
    if ($existing -and $existing.PSObject.Properties.Name -contains 'TabTip') {
        $restoreTip = [int]$existing.TabTip
    }

    Close-Word
    if (-not (Test-Path $Settings)) { New-Item -Path $Settings | Out-Null }
    Set-ItemProperty -Path $Settings -Name 'TabTip' -Value 0 -Type DWord
    Write-Note 'TabTip=0 - no timer armed, no panel created, and Word never asked where a document lives'

    Set-LogMark
    $frame = Open-Document $doc.Path $true
    Assert ($frame -ne [IntPtr]::Zero) 'the .doc opened with the switch off'
    Start-Sleep -Seconds 1

    # The switch is only observable from outside through this line, and a switch that reports its own
    # default is how TabThemeSample stayed dead for two slices.
    $start = @(Get-LogSince 'StripStart')
    Assert ($start.Count -ge 1) 'the add-in logged its startup line'
    if ($start.Count -ge 1) {
        Write-Note $start[$start.Count - 1]
        Assert ($start[$start.Count - 1] -like '*tip=off*') 'and it reads tip=off'
    }

    $offSpot = Get-Spot 'tab' 0
    Assert (Set-Pointer $offSpot.X $offSpot.Y 'the only tab') 'the pointer reached the same tab with the switch off'

    # A generous wait: with the switch on this took the double-click time, so anything up to three
    # times that and still nothing is a real absence rather than an impatient one.
    $none = Wait-WordTabTip -Present $false -Seconds ([int]([math]::Ceiling(([WordLayout]::DoubleClickTime() * 3) / 1000.0)) + 2) -What 'the tooltip'
    Assert $none.Ok 'no panel appeared'
    Assert ((Get-LogCount 'tip:') -eq 0) 'and the add-in wrote nothing about a tooltip at all - the mechanism is out of the way, not merely silent'
} finally {
    if ($null -eq $restoreTip) {
        Remove-ItemProperty -Path $Settings -Name 'TabTip' -ErrorAction SilentlyContinue
        Write-Note 'TabTip removed - back to the default'
    } else {
        Set-ItemProperty -Path $Settings -Name 'TabTip' -Value $restoreTip -Type DWord
        Write-Note "TabTip put back to $restoreTip"
    }
}

# ---- done ----------------------------------------------------------------------------------------

if (-not $KeepOpen) {
    Write-Step 'Closing Word'
    Close-Word
}

Write-Host ''
if ($script:Failures -eq 0) {
    Write-Host "$($script:Checks) checks, all passed." -ForegroundColor Green
} else {
    Write-Host "$($script:Checks) checks, $($script:Failures) FAILED." -ForegroundColor Red
    exit 1
}
