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

function Get-WordPids { @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }) }
function Get-Frames {
    $found = @()
    foreach ($id in Get-WordPids) { $found += [WordLayout]::Frames($id) }
    return @($found)
}
# Always through this - a PowerShell function returning an empty array returns *nothing*, and
# `(Get-Frames).Count` is then a hard error under StrictMode at the two moments that matter most.
function Get-FrameCount { return @(Get-Frames).Count }

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

# ---- the log, read from a mark ----------------------------------------------------------------
#
# Window handles are reused across Word restarts, and this suite restarts Word twice. So the log is
# never searched from the beginning: a mark is taken, and only what was written after it counts.
# Without that, the TabTitleTrim=0 section can match the line the TabTitleTrim=1 section wrote for a
# handle that has since been recycled, and it would pass by reading its own predecessor's evidence.

$LogPath = Join-Path $env:LOCALAPPDATA 'WordTab\wordtab.log'
$script:LogMark = 0

function Set-LogMark {
    $script:LogMark = if (Test-Path $LogPath) { (Get-Item $LogPath).Length } else { 0 }
}

function Get-LogSince($pattern) {
    if (-not (Test-Path $LogPath)) { return @() }
    $stream = [System.IO.File]::Open($LogPath, 'Open', 'Read', 'ReadWrite')
    try {
        $offset = [Math]::Min([int64]$script:LogMark, $stream.Length)
        $stream.Seek([int64]$offset, 'Begin') | Out-Null
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
        $text = $reader.ReadToEnd()
    } finally { $stream.Dispose() }
    return @(($text -split "`r?`n") | Where-Object { $_ -like "*$pattern*" })
}

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

function Set-WordForeground($seconds = 4) {
    if ((Get-MenuWindow) -ne [IntPtr]::Zero) { return [WordLayout]::GetForeground() }
    $frames = @(Get-Frames)
    if ($frames.Count -eq 0) { return [IntPtr]::Zero }

    for ($round = 1; $round -le 3; $round++) {
        $now = [WordLayout]::GetForeground()
        if ($frames -contains $now) { return $now }

        Write-Note ("the foreground was `"{0}`" - taking it back" -f [WordLayout]::TitleOf($now))
        [WordLayout]::Focus($frames[0]) | Out-Null
        $deadline = (Get-Date).AddSeconds($seconds)
        while ((Get-Date) -lt $deadline) {
            $now = [WordLayout]::GetForeground()
            if (@(Get-Frames) -contains $now) { return $now }
            Start-Sleep -Milliseconds 250
        }

        # This suite drives real clicks at screen coordinates, so a window over Word receives them.
        # Never Progman or WorkerW: that is the desktop, and minimising it is "Show Desktop".
        $now = [WordLayout]::GetForeground()
        $cls = if ($now -ne [IntPtr]::Zero) { [WordLayout]::ClassOf($now) } else { '' }
        if ($now -ne [IntPtr]::Zero -and -not (@(Get-Frames) -contains $now) -and
            $cls -ne 'Progman' -and $cls -ne 'WorkerW') {
            Write-Note ("minimising `"{0}`" - it is sitting over Word" -f [WordLayout]::TitleOf($now))
            [WordLayout]::Show($now, 6)      # SW_MINIMIZE
            Start-Sleep -Milliseconds 800
        }
    }
    return [WordLayout]::GetForeground()
}

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
function Invoke-StripClick($what, [scriptblock]$point) {
    for ($try = 1; $try -le 3; $try++) {
        Set-WordForeground | Out-Null
        $p = & $point
        $cls = [WordLayout]::ClassOf([WordLayout]::WindowAt($p.X, $p.Y))
        if ($cls -eq 'WordTabStrip') {
            [WordLayout]::Click($p.X, $p.Y)
            return $true
        }
        Write-Note ("{0}: ({1},{2}) is over `"{3}`", not the strip - taking the foreground back and re-measuring" -f
                    $what, $p.X, $p.Y, $cls)

        # And say whether Word has a dialog up, because that is the usual reason and the class name
        # on its own does not lead anywhere: a banner over Word reports as "Static", which is a
        # label inside it rather than anything that identifies it.
        $blocking = Get-WordDialog
        if ($blocking) { Write-Note ("  Word has a dialog up: {0}" -f (Format-Dialog $blocking)) }

        Start-Sleep -Milliseconds 800
    }
    return $false
}

function Get-MenuWindow {
    foreach ($id in Get-WordPids) {
        $w = [WordLayout]::PopupMenuWindow($id)
        if ($w -ne [IntPtr]::Zero) { return $w }
    }
    return [IntPtr]::Zero
}

function Wait-Menu($present, $seconds = 6) {
    $deadline = (Get-Date).AddSeconds($seconds)
    while ((Get-Date) -lt $deadline) {
        $up = (Get-MenuWindow) -ne [IntPtr]::Zero
        if ($up -eq $present) { return $true }
        Start-Sleep -Milliseconds 200
    }
    return $false
}

function Close-Menu {
    if ((Get-MenuWindow) -ne [IntPtr]::Zero) {
        [WordLayout]::Press(0x1B)
        Wait-Menu $false | Out-Null
    }
}

function Open-TabMenu($index) {
    $spot = Get-Spot 'label' $index
    $owner = (Get-TopStrip).Frame
    [WordLayout]::RightClick($spot.X, $spot.Y)
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

function Get-WordDialog {
    foreach ($id in Get-WordPids) {
        foreach ($w in [WordLayout]::TopLevel($id)) {
            if (-not $w.Visible) { continue }
            if ($w.Class -eq 'OpusApp' -or $w.Class -eq '#32768') { continue }
            if (($w.Right - $w.Left) -lt 200 -or ($w.Bottom - $w.Top) -lt 100) { continue }
            return $w
        }
    }
    return $null
}

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
function Get-SavePrompt {
    if (-not (Get-WordDialog)) { return $null }
    Start-Sleep -Milliseconds 1500
    return Get-WordDialog
}

# What it was, in the message. A guard that aborts a run saying only "Word is asking something" is
# the same mistake as a colour sampler that answers with a bare FALSE: it cannot say what it
# measured, so the next person has to reproduce it before they can even start.
function Format-Dialog($d) {
    if (-not $d) { return '(nothing)' }
    return "class={0} {1}x{2} |{3}|" -f $d.Class, $d.Width, $d.Height, $d.Title
}

function Close-Word {
    if (@(Get-WordPids).Count -eq 0) { return }
    Write-Note "closing Word: $(Get-FrameCount) window(s)"
    for ($guard = 0; $guard -lt 24; $guard++) {
        $open = @(Get-Frames)
        if ($open.Count -eq 0) { break }
        [WordLayout]::Close($open[0])
        Start-Sleep -Milliseconds 1200
        # A leftover dirty document. Deliberately not answered: a test script may not discard work
        # it did not create.
        $prompt = Get-SavePrompt
        if ($prompt) {
            Write-Note ("Word put up: {0}" -f (Format-Dialog $prompt))
            [WordLayout]::Press(0x1B)
            Start-Sleep -Seconds 2
            throw ("Word is asking something and it is still there after Escape: {0}  Answer it by hand and re-run." -f
                   (Format-Dialog $prompt))
        }
    }
    foreach ($p in @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue)) { $p.CloseMainWindow() | Out-Null }
    $deadline = (Get-Date).AddSeconds(20)
    while ((Get-Date) -lt $deadline -and @(Get-WordPids).Count -gt 0) { Start-Sleep -Milliseconds 500 }
    if (@(Get-WordPids).Count -gt 0) { throw 'Word would not close - close it by hand and re-run.' }
    Start-Sleep -Seconds 2
}

function Open-Document($path, $first) {
    $before = @(Get-Frames)
    Start-Process -FilePath 'winword.exe' -ArgumentList "`"$path`""
    Start-Sleep -Seconds $(if ($first) { 16 } else { 7 })
    $deadline = (Get-Date).AddSeconds(30)
    while ((Get-Date) -lt $deadline) {
        $new = @(Get-Frames | Where-Object { $before -notcontains $_ })
        if ($new.Count -gt 0) { Start-Sleep -Seconds 2; return $new[0] }
        Start-Sleep -Milliseconds 500
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
    [WordLayout]::Click($centre.X, $centre.Y)
} else {
    Write-Note 'no rectangle to click - using the access key'
    [WordLayout]::Press(0x52)   # R
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
