<#
.SYNOPSIS
  Measure WordTab's strip inside a running Word, and put Word through the layout changes that are
  supposed to break it.

.DESCRIPTION
  The strip's claim is a geometric one: WordTab owns a horizontal band between Word's ribbon and
  its document, and that band stays exactly there through everything Word does to its own layout.
  A claim like that should not be checked by looking at it, so this reads the actual rectangles
  from outside the process and asserts three things after every change:

    - the strip sits immediately above the document frame  (strip.bottom == _WwF.top)
    - it spans the document frame exactly                  (same left and right)
    - nothing is left uncovered above it                   (no gap to the chrome above)

  Then it drives the four cases that matter, in order of how likely each was to break it:
  resize, maximize, restore, and Backstage - the full-screen File menu, which is what reportedly
  broke the old "Doc Tabs" add-in, because leaving it makes Word rebuild its layout from scratch.

  Word is started if it is not already running, on a scratch document written by this script.
  Word's documents on the dev rig are disposable test fixtures.

.PARAMETER KeepOpen
  Leave Word running afterwards instead of closing it. Use this to look at the result by hand.

.PARAMETER Screenshot
  Save a PNG of each Word frame at the end, next to the script's output.

.PARAMETER SecondDocument
  Open a second document, so the per-frame strip can be checked on more than one frame.

.EXAMPLE
  pwsh -File tools\check-strip.ps1
  pwsh -File tools\check-strip.ps1 -KeepOpen -Screenshot -SecondDocument
#>
[CmdletBinding()]
param(
    [switch]$KeepOpen,
    [switch]$Screenshot,
    [switch]$SecondDocument,
    [string]$ShotDir = $env:TEMP
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# The Win32 side lives in tools\WordLayout.cs so probes and later slices can load the same one.
$source = Get-Content -Raw -Path (Join-Path $PSScriptRoot 'WordLayout.cs')

# The explicit reference list is not decoration: naming any assembly at all replaces PowerShell's
# default set, so the ones the compiler would otherwise have had for free have to come back too.
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition $source -Language CSharp -ReferencedAssemblies @(
    'System.Runtime', 'System.Collections', 'System.Threading.Thread', 'netstandard'
)

# Physical pixels. This rig is at 150%, and without this every rectangle below is silently scaled.
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
    if ($condition) {
        Write-Host "    PASS  $text" -ForegroundColor Green
    } else {
        $script:Failures++
        Write-Host "    FAIL  $text" -ForegroundColor Red
    }
}

# ---- get Word up, with a document -------------------------------------------------------------

$scratch = Join-Path $env:TEMP 'wordtab-check'
New-Item -ItemType Directory -Path $scratch -Force | Out-Null
$doc1 = Join-Path $scratch 'wordtab-check-1.rtf'
$doc2 = Join-Path $scratch 'wordtab-check-2.rtf'
'{\rtf1\ansi WordTab layout check - document one.\par}' | Set-Content -Path $doc1 -Encoding Ascii
'{\rtf1\ansi WordTab layout check - document two.\par}' | Set-Content -Path $doc2 -Encoding Ascii

$startedWord = $false
if (-not (Get-Process -Name WINWORD -ErrorAction SilentlyContinue)) {
    Write-Step 'Starting Word'
    Start-Process -FilePath 'winword.exe' -ArgumentList "`"$doc1`""
    $startedWord = $true
}

$deadline = (Get-Date).AddSeconds(60)
$frames = @()
while ((Get-Date) -lt $deadline) {
    $frames = @()
    foreach ($process in @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue)) {
        $frames += [WordLayout]::Frames($process.Id)
    }
    # A frame with no `_WwF` yet is Word still building its window, not a frame to measure.
    if ($frames.Count -gt 0) {
        $kids = [WordLayout]::Children($frames[0])
        if (@($kids | Where-Object { $_.Class -eq '_WwF' }).Count -gt 0) { break }
    }
    Start-Sleep -Milliseconds 500
}
if ($frames.Count -eq 0) { throw 'No usable Word frame appeared within 60s.' }

if ($SecondDocument) {
    Write-Step 'Opening a second document'
    Start-Process -FilePath 'winword.exe' -ArgumentList "`"$doc2`""
    # Waiting for Word to get into position, not for anything this suite claims about the product.
    if (-not (Wait-WordReady 2 45)) { Write-Note 'the second document did not arrive with a strip on it within 45s' }
    $frames = @()
    foreach ($process in @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue)) {
        $frames += [WordLayout]::Frames($process.Id)
    }
}

# Drive the window that is in front, and put it in front first. This is not tidiness: Word lays out
# only the focused window, and when several documents are open the add-in stacks them and drives the
# geometry from the active one. Resizing a window behind the stack is not something a user can do,
# and measuring one is measuring a window Word has stopped maintaining.
$target = [WordLayout]::GetForeground()
if (-not ($frames -contains $target)) { $target = $frames[0] }
[WordLayout]::Focus($target) | Out-Null
Write-Note ("target frame 0x{0:X}  `"{1}`"  ({2} visible frame(s))" -f [int64]$target, [WordLayout]::TitleOf($target), $frames.Count)

# ---- the measurement ---------------------------------------------------------------------------

function Get-Layout($frame) {
    $kids = [WordLayout]::Children($frame)
    $strip = @($kids | Where-Object { $_.Class -eq 'WordTabStrip' }) | Select-Object -First 1
    $wwf   = @($kids | Where-Object { $_.Class -eq '_WwF' }) | Select-Object -First 1
    [pscustomobject]@{
        Frame    = $frame
        Children = $kids
        Strip    = $strip
        Wwf      = $wwf
        Client   = [WordLayout]::ClientOf($frame)
    }
}

function Show-Layout($layout) {
    # Only the children that lay out the window: the ones spanning most of its width. Word hangs
    # dozens of small helper windows off the frame and listing them all buries the four that matter.
    $wide = $layout.Children | Where-Object { $_.Visible -and $_.Width -gt ($layout.Client.Right * 0.6) } |
            Sort-Object Top
    foreach ($c in $wide) {
        $mark = if ($c.Class -eq 'WordTabStrip') { '  <- ours' } else { '' }
        Write-Note ("{0,-16} y {1,5} .. {2,-5} h={3,-5} x {4,4}..{5,-5}{6}" -f `
            $c.Class, $c.Top, $c.Bottom, $c.Height, $c.Left, $c.Right, $mark)
    }
}

function Test-Layout($label, [switch]$Quiet) {
    $layout = Get-Layout $target
    Write-Step $label
    if (-not $Quiet) { Show-Layout $layout }

    if (-not $layout.Wwf) {
        Assert $false "$label - `_WwF` not found (Word has no document frame)"
        return
    }
    if (-not $layout.Strip) {
        Assert $false "$label - no WordTabStrip child (the add-in is not running, or the strip is off)"
        return
    }

    $strip = $layout.Strip
    $wwf   = $layout.Wwf

    Assert ($strip.Bottom -eq $wwf.Top) `
        ("$label - strip bottom {0} meets document top {1}" -f $strip.Bottom, $wwf.Top)
    Assert ($strip.Left -eq $wwf.Left -and $strip.Right -eq $wwf.Right) `
        ("$label - strip spans the document exactly ({0}..{1} vs {2}..{3})" -f $strip.Left, $strip.Right, $wwf.Left, $wwf.Right)

    # Whatever Word draws directly above us - the ribbon - should end exactly where we begin. A gap
    # means we took a band Word still thinks is its own; an overlap means we are covering chrome.
    $above = $layout.Children |
             Where-Object { $_.Visible -and $_.Hwnd -ne $strip.Hwnd -and $_.Hwnd -ne $wwf.Hwnd -and
                            $_.Bottom -le $strip.Top -and $_.Width -gt ($layout.Client.Right * 0.6) } |
             Sort-Object Bottom -Descending | Select-Object -First 1
    if ($above) {
        Assert ($strip.Top -eq $above.Bottom) `
            ("$label - no gap above: {0} ends at {1}, strip starts at {2}" -f $above.Class, $above.Bottom, $strip.Top)
    } else {
        Write-Note "(nothing above the strip to compare against)"
    }

    Assert ($wwf.Height -gt 0 -and $wwf.Bottom -le $layout.Client.Bottom) `
        ("$label - document frame still inside the window (h={0}, bottom {1} <= {2})" -f $wwf.Height, $wwf.Bottom, $layout.Client.Bottom)
}

function Save-FrameShot($frame, $name) {
    if (-not $Screenshot) { return }
    New-Item -ItemType Directory -Path $ShotDir -Force | Out-Null
    $path = Join-Path $ShotDir "wordtab-$name.png"
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

# ---- baseline --------------------------------------------------------------------------------
#
# Wait for the layout to settle before asserting anything about it. Word spends a second or two
# after a window appears rearranging its own chrome, and the add-in corrects itself on a
# half-second cadence, so a measurement taken the instant a `_WwF` exists is measuring the settle
# rather than the result. The wait is bounded: if it never settles, the assertions run anyway and
# report what they found.

$settle = (Get-Date).AddSeconds(12)
while ((Get-Date) -lt $settle) {
    $layout = Get-Layout $target
    if ($layout.Strip -and $layout.Wwf -and $layout.Strip.Bottom -eq $layout.Wwf.Top) { break }
    Start-Sleep -Milliseconds 500
}

Test-Layout 'Baseline'

# ---- resize ----------------------------------------------------------------------------------

# Focused again before each geometry step: Word only lays out the window that has focus, so a step
# driven against a window that has lost it measures nothing.
[WordLayout]::Focus($target) | Out-Null
if ([WordLayout]::Maximized($target)) { [WordLayout]::Show($target, [WordLayout]::SW_RESTORE); Start-Sleep -Milliseconds 800 }

foreach ($size in @(@(1000, 800), @(1360, 900), @(900, 700))) {
    [WordLayout]::Resize($target, $size[0], $size[1])
    Start-Sleep -Milliseconds 900
    Test-Layout ("Resized to {0}x{1}" -f $size[0], $size[1]) -Quiet
}

# ---- maximize / restore -------------------------------------------------------------------------

[WordLayout]::Focus($target) | Out-Null
[WordLayout]::Show($target, [WordLayout]::SW_MAXIMIZE)
Start-Sleep -Milliseconds 1200
Test-Layout 'Maximized'

[WordLayout]::Show($target, [WordLayout]::SW_RESTORE)
Start-Sleep -Milliseconds 1200
Test-Layout 'Restored'

# ---- resize by dragging the border -----------------------------------------------------------
#
# The stress case. Everything above changes the window once; a drag changes it every frame, inside
# Word's modal loop, with our code between Word and its own window procedure for each one. Two
# things are being watched: that the strip still lands correctly at the end, and that the add-in
# does not write a log line per frame - a file write inside that loop is a stutter the user feels.

$logFile = Join-Path $env:LOCALAPPDATA 'WordTab\wordtab.log'
function Get-StripLineCount {
    if (-not (Test-Path $logFile)) { return 0 }
    # Only this window's own lines. frames.cpp also writes a dozen lines per drag, but those are the
    # previous slice's timing report, buffered in memory and flushed once at WM_EXITSIZEMOVE -
    # deliberately not a per-frame cost, and not what this is watching for. And every other window
    # in the stack keeps its own log, so counting all of them would just measure how many documents
    # happen to be open.
    $mine = 'hwnd=0x{0:X16}' -f [int64]$target
    return @(Get-Content $logFile | Where-Object { $_ -match '\bstrip\s+hwnd=' -and $_ -match $mine }).Count
}
$linesBefore = Get-StripLineCount

if ([WordLayout]::Focus($target)) {
    $r = [WordLayout]::RectOf($target)
    $grabX = $r.Right - 2
    $grabY = $r.Top + [int](($r.Bottom - $r.Top) / 2)

    Write-Step 'Resize drag (Word''s modal loop, one relayout per frame)'
    [WordLayout]::DragBy($grabX, $grabY, 40, -6, 0, 8) | Out-Null
    Start-Sleep -Milliseconds 600
    Test-Layout 'During-drag end state (narrower)' -Quiet

    $r = [WordLayout]::RectOf($target)
    [WordLayout]::DragBy($r.Right - 2, $grabY, 40, 6, 0, 8) | Out-Null
    Start-Sleep -Milliseconds 600
    Test-Layout 'Back to the original width' -Quiet

    $written = (Get-StripLineCount) - $linesBefore
    Assert ($written -lt 15) "the strip stayed quiet during the drag ($written log lines for 80 frames)"
} else {
    Write-Warning 'Could not bring Word to the foreground; skipping the resize drag.'
}

# ---- Backstage ----------------------------------------------------------------------------------
#
# The one that reportedly broke the old Doc Tabs add-in: leaving Backstage makes Word rebuild its
# layout from scratch. Inside Backstage the document frame is gone, so there is nothing to assert
# except that we did not leave a strip floating over the File menu; the real check is what comes
# back afterwards.

if ([WordLayout]::Focus($target)) {
    Write-Step 'Backstage (File menu)'

    # Retried, and confirmed rather than assumed. Injected input into another process's UI fails
    # quietly, and a Backstage check that silently never opened Backstage is worse than no check.
    # One click into the document first. Word's ribbon ignores Alt+F when the window was activated
    # programmatically rather than clicked - measured, and it fails silently.
    #
    # Aimed at the document frame's own SCREEN rectangle. The previous version added client-relative
    # child coordinates to the frame's window rectangle and dropped Wwf.Left entirely, so it was off
    # by the window border in x and by the caption in y - which on a narrow window lands outside the
    # document, and a click that misses here makes Alt+F fail three lines later for a reason that
    # looks nothing like a mis-aimed click.
    $layout = Get-Layout $target
    if ($layout.Wwf) {
        $aim = {
            $v = [WordLayout]::RectOf($layout.Wwf.Hwnd)
            [pscustomobject]@{ X = $v.Left + [int](($v.Right - $v.Left) / 2)
                               Y = $v.Top  + [int](($v.Bottom - $v.Top) / 2) }
        }
        $where = & $aim
        $under = Get-ClassAt $where.X $where.Y
        # _WwG is the document pane, _WwN the one inside a split view. Either is the document.
        $expect = if ($under -eq '_WwN') { '_WwN' } else { '_WwG' }
        if (-not (Invoke-ConfirmedClick -What 'clicking into the document before Alt+F' -Point $aim -Expect $expect)) {
            Write-Note ("the click into the document did not land - ({0},{1}) is over `"{2}`"" -f $where.X, $where.Y, $under)
        }
    }

    $opened = $false
    for ($attempt = 1; $attempt -le 3 -and -not $opened; $attempt++) {
        # A failed Alt+F can leave Word showing KeyTips, where the next one means something else.
        # Confirmed before sending: a keystroke goes to whatever has the foreground, and a silently
        # missed one is indistinguishable from Backstage refusing to open.
        if (-not (Test-WordHasFocus)) {
            Write-Note ("the foreground is `"{0}`", not Word - taking it back before Alt+F" -f
                        [WordLayout]::TitleOf([WordLayout]::GetForeground()))
            Set-WordForeground | Out-Null
        }
        [WordLayout]::CloseBackstage()
        Start-Sleep -Milliseconds 300
        [WordLayout]::OpenBackstage()
        Start-Sleep -Milliseconds 2000
        $opened = [WordLayout]::BackstageOpen($target)
    }
    Assert $opened 'Backstage opened (FullpageUIHost is up)'

    if ($opened) {
        # A picture answers what the rectangles cannot: whether our strip is sitting on top of the
        # File menu. It is a child of the frame, so it could be.
        Save-FrameShot $target 'backstage'

        $inside = Get-Layout $target
        $hidden = (-not $inside.Strip) -or (-not $inside.Strip.Visible)
        Assert $hidden 'the strip is hidden while Backstage is up (Word hides its children, ours included)'

        [WordLayout]::CloseBackstage()
        Start-Sleep -Milliseconds 2500
        Assert (-not [WordLayout]::BackstageOpen($target)) 'Backstage closed again'
        Test-Layout 'After leaving Backstage'
    }
} else {
    Write-Warning 'Could not bring Word to the foreground; skipping the Backstage check.'
}

# ---- every frame, not just the target -----------------------------------------------------------

$frames = @()
foreach ($process in @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue)) {
    $frames += [WordLayout]::Frames($process.Id)
}
if ($frames.Count -gt 1) {
    Write-Step "All $($frames.Count) frames have their own strip"
    foreach ($frame in $frames) {
        $layout = Get-Layout $frame
        $ok = $layout.Strip -and $layout.Wwf -and $layout.Strip.Bottom -eq $layout.Wwf.Top
        Assert $ok ("frame 0x{0:X} `"{1}`"" -f [int64]$frame, [WordLayout]::TitleOf($frame))
    }
}

# ---- pictures -------------------------------------------------------------------------------------

if ($Screenshot) {
    Write-Step 'Screenshots'
    $index = 0
    foreach ($frame in $frames) { Save-FrameShot $frame ("frame-{0}" -f $index++) }
}

# ---- done -------------------------------------------------------------------------------------

# NEVER Kill. The comment this replaces argued that these are scratch files so a Kill is harmless -
# and the harm is not to the file. A killed Word offers to RECOVER those documents on the next
# launch, and the recovered windows then make the next suite measure more windows than it opened.
# That exact failure cost a run in the scrolling-row slice, traced back to soak-stack's cleanup;
# this was the second copy, and check-stack's was the third.
#
# Close-AllWord stops and reports if Word asks a question rather than pressing anything: a test
# script may not answer a save prompt it did not raise.
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
Write-Host "Add-in log: $env:LOCALAPPDATA\WordTab\wordtab.log (lines starting 'strip')" -ForegroundColor Gray
exit $script:Failures
