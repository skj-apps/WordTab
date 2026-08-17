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

# ---- the same strip at four different scales -----------------------------------------------------
#
# Until this section existed, every scaled size in the add-in had only ever been computed at ONE
# number. The dev rig runs at 192 and Windows offers its display exactly one scale factor - measured
# with the private DisplayConfig scaling packet, which answers min == cur == max - so 96, 120 and 144
# were unreachable here. The work rig is a laptop plus a 40" second monitor, which is where a second
# DPI first happens, and it is the machine nobody here can reach.
#
# What this drives that nothing else did: `StripOnFrameDpiChanged`. That function had never executed,
# on any rig, in the life of the project. It re-derives the height, the font, the corner radius, the
# glyph stroke and the back buffer, takes down a tooltip measured at the old scale, and un-applies
# the shift it made to `_WwF` at the old scale so the next layout can re-apply at the new one. All of
# it was reasoned and none of it was run.
#
# WHAT THIS DOES NOT TEST, stated rather than discovered later: the trigger. In production the
# trigger is WM_DPICHANGED arriving from Windows, and that message cannot be synthesised from another
# process - it carries a suggested RECT by pointer and the frame proc chains it on to Word, which
# would dereference an address from the wrong address space. So the override is the way in, and what
# is covered is everything downstream of the two lines in frames.cpp that receive the message.
#
# It owns its own Word because the override has to be in the registry BEFORE StripStart reads it, and
# it puts the machine back the same way it found it - the value removed, Word restarted, the strip
# asserted to be at the scale Windows actually says. That last restart is not tidying; it is the
# check that the restore worked, which is the only reason to run this before the closing section
# rather than after it.

Write-Step 'The strip at scales this machine does not run at'

$dpiKey     = 'HKCU:\Software\WordTab'
$realDpi    = [WordLayout]::Dpi($target)     # asked BEFORE any override is written
$dpiDoc     = Join-Path $scratch 'wordtab-check-dpi.rtf'
'{\rtf1\ansi WordTab DPI check.\par}' | Set-Content -Path $dpiDoc -Encoding Ascii

Write-Note ("this display reports $realDpi dpi, and offers no other scale factor")

try {
    # Word has to go down and come back with the value already set. Close-AllWord rather than Kill,
    # for the reason the closing section gives.
    $down = Close-AllWord
    if (-not $down.Closed) {
        Assert $false ("DPI section - Word would not close ({0}), so this section cannot establish its own Word" -f $down.Reason)
    } else {
        # Start at the DPI the machine really is. Nothing about the strip differs from an ordinary
        # run at this point, which is what makes the first measurement a control rather than a
        # restatement of the override.
        New-ItemProperty -Path $dpiKey -Name TabDpi -Value $realDpi -PropertyType DWord -Force | Out-Null

        Set-LogMark
        Start-Process -FilePath 'winword.exe' -ArgumentList "`"$dpiDoc`""
        if (-not (Wait-WordReady 1 45)) {
            Assert $false 'DPI section - Word did not come back up with a strip within 45s'
        } else {
            $dpiTarget = @(Get-WordFrameList)[0]
            [WordLayout]::Focus($dpiTarget) | Out-Null
            Start-Sleep -Milliseconds 1500

            # The add-in must SAY it is overridden. Without this the whole section could pass by
            # measuring a strip that was never overridden at all - every assertion below compares a
            # height against a number computed from the same DPI, so an override that silently did
            # nothing at $realDpi would look identical to one that worked.
            Assert (@(Get-LogSince 'DPI FORCED to').Count -ge 1) `
                ("DPI section - the add-in reports the override in force (TabDpi={0})" -f $realDpi)

            # 32 logical px is STRIP_LOGICAL_H, and MulDiv rounds to nearest - the same rounding
            # WordLayout::Scale uses, which is why the harness and the add-in can be compared at all.
            function Get-ExpectedStripHeight($dpi) { return [int][Math]::Round(32.0 * $dpi / 96.0) }

            $steps = @($realDpi, 96, 120, 144, $realDpi)
            $previous = $null

            foreach ($dpi in $steps) {
                Set-ItemProperty -Path $dpiKey -Name TabDpi -Value $dpi

                # The janitor is on a 500ms cadence and re-reads the override only when one was set
                # at startup. Bounded wait on the thing itself rather than a sleep long enough to be
                # safe: a fixed sleep hides the difference between reacting and being seconds behind.
                $want = Get-ExpectedStripHeight $dpi
                $arrived = Wait-Until {
                    $l = Get-Layout $dpiTarget
                    $l.Strip -and ($l.Strip.Bottom - $l.Strip.Top) -eq $want
                } 8 200

                $layout = Get-Layout $dpiTarget
                $height = if ($layout.Strip) { $layout.Strip.Bottom - $layout.Strip.Top } else { -1 }

                Assert $arrived ("{0} dpi - the strip re-scaled to {1}px (STRIP_LOGICAL_H 32 x {0}/96), measured {2}" -f `
                                 $dpi, $want, $height)

                # The harness has to agree with the add-in about the scale, and this is the ONLY
                # place in the battery that can find out. WordLayout::Dpi carries its own copy of the
                # override read - it must, because every slot it computes for a click is derived from
                # it - and two copies of "what DPI is this" agree right up until one of them is
                # wrong. Nothing else ever runs with TabDpi set, so without this line that second
                # copy is never executed at all.
                #
                # It fails SILENTLY if it is ever broken, which is why it gets an assertion rather
                # than trust: a bad RegGetValueW p/invoke returns non-zero, ForcedDpi answers 0, and
                # Dpi falls through to asking Windows - which on this rig returns 192, the same
                # number the override is set to for two of these five steps. Only the 96, 120 and 144
                # steps can tell a working read from a dead one.
                Assert ([WordLayout]::Dpi($dpiTarget) -eq $dpi) `
                    ("{0} dpi - the harness reads the same override the add-in did (WordLayout::Dpi says {1})" -f `
                     $dpi, [WordLayout]::Dpi($dpiTarget))

                # The transition had to be REPORTED, and by the function under test. A height that
                # happens to be right because Word rebuilt the window is a different thing from
                # StripOnFrameDpiChanged having run.
                if ($null -ne $previous -and $previous -ne $dpi) {
                    $line = "DPI $previous -> $dpi, re-scaling the strip"
                    Assert (@(Get-LogSince $line).Count -ge 1) `
                        ("{0} -> {1} dpi - StripOnFrameDpiChanged ran and said so" -f $previous, $dpi)
                }

                # The seams. This is the assertion that a wrong scale would actually show up in: the
                # strip has to still meet the document frame exactly and still sit directly under
                # Word's ribbon, with no gap and no overlap, at every one of these sizes.
                if ($layout.Strip -and $layout.Wwf) {
                    Assert ($layout.Strip.Bottom -eq $layout.Wwf.Top) `
                        ("{0} dpi - strip bottom {1} still meets document top {2}" -f $dpi, $layout.Strip.Bottom, $layout.Wwf.Top)
                }
                $place = Get-StripPlacement (@(Get-WordFrameList))
                Assert ($place.Ok -and $place.Measured -ge 1) `
                    ("{0} dpi - every strip still sits under Word's chrome ({1} measured){2}" -f `
                     $dpi, $place.Measured, $place.Text)

                $previous = $dpi
            }
        }
    }
}
finally {
    # The value goes whatever happened above, including a throw. A TabDpi left behind would make
    # every suite after this one in the battery run against a strip built at a scale the machine is
    # not at - and it would do it silently, because the add-in only mentions the override on the
    # startup line nobody downstream reads.
    Remove-ItemProperty -Path $dpiKey -Name TabDpi -ErrorAction SilentlyContinue
}

# And the restore, asserted rather than assumed. Word comes back with no override at all: the strip
# must be at the scale Windows reports, and the add-in must have stopped claiming otherwise.
$down = Close-AllWord
if ($down.Closed) {
    Set-LogMark
    Start-Process -FilePath 'winword.exe' -ArgumentList "`"$dpiDoc`""
    if (Wait-WordReady 1 45) {
        $restored = @(Get-WordFrameList)[0]
        [WordLayout]::Focus($restored) | Out-Null
        Start-Sleep -Milliseconds 1500

        $layout = Get-Layout $restored
        $height = if ($layout.Strip) { $layout.Strip.Bottom - $layout.Strip.Top } else { -1 }
        $want   = [int][Math]::Round(32.0 * $realDpi / 96.0)

        Assert ($height -eq $want) `
            ("restored - no override, strip back to {0}px at the display's own {1} dpi (measured {2})" -f $want, $realDpi, $height)
        Assert (@(Get-LogSince 'DPI FORCED to').Count -eq 0) `
            'restored - the add-in no longer reports a forced DPI'

        $frames = @(Get-WordFrameList)
    } else {
        Assert $false 'restored - Word did not come back up after the override was removed'
    }
} else {
    Assert $false ("restored - Word would not close to drop the override ({0})" -f $down.Reason)
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
