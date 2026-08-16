<#
.SYNOPSIS
  Measure what WordTab does with a Word window that has no document open.

.DESCRIPTION
  Word has two states where a window is on screen with nothing in it: the Start screen it opens on
  when launched with no file, and the empty frame it leaves standing when you close its last
  document with Ctrl+W. Both are real, both are visible, and neither is a document.

  The premise this add-in was built on was that such a window has no `_WwF` and so falls out of the
  stack by itself. **That is false.** `_WwF` is Word's document *frame*, and Word keeps it for the
  life of the window: closing the last document destroys the `_WwB` and `_WwG` inside it and leaves
  `_WwF` behind, empty. So the window passed the membership test, joined the stack, and was given a
  tab labelled "Word" - a tab for no document, with a close button and everything.

  The policy this checks is: **the tab row lists documents.** A window with no document is not a
  tab. It is not a stack member, so it keeps its own taskbar button and Alt+Tab entry and the stack
  neither moves it nor hides it. Its strip stays - the band is already carved, so nothing jumps when
  a document arrives - and draws an empty row with only the + on it. The moment a document appears
  in the window it joins the row normally.

  **How "the row is empty" is proved.** Not by looking at pixels. The + button's position is a
  function of the tab count: with no tabs it sits at the left padding, with one tab it sits past
  that tab. So the script computes both positions, asserts they differ, and clicks the *no-tabs*
  one. If the row really is empty that is the +, and a document appears in this very window. If the
  row still had the old "Word" tab, that same click would land on the tab and nothing would happen.
  The two outcomes cannot be confused. The reverse check is done too: clicking where the "Word" tab
  used to be must do nothing at all.

  **Two things measured here that the sequence depends on, and that are easy to get wrong:**

    - On the *cold* Start screen the strip exists but is **behind Word's own `FullpageUIHost`**,
      which covers the whole client area exactly as Backstage does. So the row is not clickable
      there, and this suite does not pretend otherwise: it asserts the occlusion instead, and gets
      out of the Start screen with Escape the way a user would.
    - Opening a document while a document-less window is up makes Word **reuse that window**. So the
      second document has to be opened only once the first window has a document of its own, or
      there is never a second window to stack.

  Word's documents on the dev rig are disposable test fixtures. Everything opened here is clean and
  unmodified, so nothing prompts to save.

.PARAMETER KeepOpen
  Leave Word running afterwards, to look at it by hand.

.PARAMETER Screenshot
  Save photographs of the empty row, the two-tab row, and the row it comes back to.

.EXAMPLE
  pwsh -File tools\check-startscreen.ps1
  pwsh -File tools\check-startscreen.ps1 -KeepOpen -Screenshot
#>
[CmdletBinding()]
param(
    [switch]$KeepOpen,
    [switch]$Screenshot,
    [string]$ShotDir = $env:TEMP
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# A run that drives Word for minutes and then dies with one line and no line number costs the whole
# run to diagnose. Borrowed from check-menu.ps1.
trap {
    Write-Host ''
    Write-Host "UNHANDLED: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkRed
    break
}

$source = Get-Content -Raw -Path (Join-Path $PSScriptRoot 'WordLayout.cs')
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition $source -Language CSharp -ReferencedAssemblies @(
    'System.Runtime', 'System.Collections', 'System.Threading.Thread', 'netstandard'
)
[WordLayout]::MakeDpiAware() | Out-Null

# Confirmed input, one shared idea of what a Word dialog is, and the bounded waits. See the header of
# tools\WordTabHarness.ps1 for why these are not per-suite copies any more.
. (Join-Path $PSScriptRoot 'WordTabHarness.ps1')

$LogPath = Join-Path $env:LOCALAPPDATA 'WordTab\wordtab.log'
$VK_W        = 0x57
$SW_MINIMIZE = 6

$script:Failures = 0
$script:Checks   = 0

function Write-Step($text) { Write-Host "==> $text" -ForegroundColor Cyan }
function Write-Note($text) { Write-Host "    $text" -ForegroundColor DarkGray }
function Assert($condition, $text) {
    $script:Checks++
    if ($condition) { Write-Host "    PASS  $text" -ForegroundColor Green }
    else { $script:Failures++; Write-Host "    FAIL  $text" -ForegroundColor Red }
}

# Every one of these re-wraps in @() before asking for .Count. A PowerShell function that returns an
# empty array returns *nothing*, and `.Count` on nothing is a hard error under Set-StrictMode rather
# than 0. This suite spends most of its time at zero or one of everything, so the trap that has
# already cost two runs on this project is the normal case here.
function Get-WordPids { return Get-WordPidList }
function Get-WordPidCount { return Get-WordPidTally }

function Get-Frames { return Get-WordFrameList }
function Get-FrameCount { return Get-WordFrameTally }

# Deliberately not the `@(... | Where-Object ...)[0]` shape the older suites use: under
# Set-StrictMode that *throws* on an empty match rather than yielding $null, and every window this
# suite looks at is one where the match legitimately comes back empty.
function Get-Child($frame, $class) {
    foreach ($kid in [WordLayout]::Children($frame)) {
        if ($kid.Class -eq $class) { return $kid }
    }
    return $null
}

function Get-Parts($frame) {
    [pscustomobject]@{
        Frame = $frame
        Rect  = [WordLayout]::RectOf($frame)
        Strip = (Get-Child $frame 'WordTabStrip')
        Wwf   = (Get-Child $frame '_WwF')
        Title = [WordLayout]::TitleOf($frame)
    }
}

# The add-in's own document test, mirrored: is there anything inside the document frame?
# See StripHasDocument in src\native\strip.cpp.
function Test-HasDocument($frame) {
    $wwf = Get-Child $frame '_WwF'
    if (-not $wwf) { return $false }
    return ([WordLayout]::FirstChild($wwf.Hwnd) -ne [IntPtr]::Zero)
}

function Get-LogMark { if (Test-Path $LogPath) { return (Get-Item $LogPath).Length } else { return 0 } }

function Get-LogSince($offset, $pattern) {
    if (-not (Test-Path $LogPath)) { return @() }
    $stream = [System.IO.File]::Open($LogPath, 'Open', 'Read', 'ReadWrite')
    try {
        # The add-in deletes the log when it passes half a megabyte. If that happened mid-run the
        # offset points past the end of a smaller file, and reading from the start is the honest
        # answer rather than an exception.
        if ($offset -gt $stream.Length) { $offset = 0 }
        $stream.Seek([int64]$offset, 'Begin') | Out-Null
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
        $text = $reader.ReadToEnd()
    } finally { $stream.Dispose() }
    return @(($text -split "`r?`n") | Where-Object { $_ -like "*$pattern*" })
}

function Get-LogCount($offset, $pattern) { return @(Get-LogSince $offset $pattern).Count }

# Where the + is when the row has $count tabs. Recomputed at every use, never held: the strip moves
# whenever Word relays the window out, and this suite changes the row on purpose.
function Get-PlusAt($strip, $count) {
    $layout = [WordLayout]::Tabs($strip, $count)
    if (-not $layout.HasPlus) { throw "No + button in a $count-tab layout." }
    return [WordLayout]::Center($layout.Plus)
}

function Get-TheFrame {
    $frames = @(Get-Frames)
    if ($frames.Count -eq 0) { throw 'No visible Word window.' }
    $top = [WordLayout]::GetForeground()
    if ($frames -contains $top) { return $top }
    return $frames[0]
}

function Get-StripOf($frame) {
    $parts = Get-Parts $frame
    if (-not $parts.Strip) { throw ("Window 0x{0:X} has no WordTab strip." -f [int64]$frame) }
    return $parts.Strip.Hwnd
}

function Wait-For($predicate, $seconds = 20) {
    $deadline = (Get-Date).AddSeconds($seconds)
    while ((Get-Date) -lt $deadline) {
        if (& $predicate) { return $true }
        Start-Sleep -Milliseconds 400
    }
    return $false
}

function Save-StripShot($strip, $name) {
    if (-not $Screenshot) { return }
    New-Item -ItemType Directory -Path $ShotDir -Force | Out-Null
    $r = [WordLayout]::RectOf($strip)
    $bmp = New-Object System.Drawing.Bitmap(($r.Right - $r.Left), ($r.Bottom - $r.Top))
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($r.Left, $r.Top, 0, 0, $bmp.Size)
    $g.Dispose()
    $path = Join-Path $ShotDir "wordtab-$name.png"
    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    Write-Note $path
}

# ---- a clean Word ---------------------------------------------------------------------------------
#
# Gracefully, never Kill. A killed Word offers to recover these documents on the next launch, which
# adds a window nobody asked for - and this suite counts windows.

if ((Get-WordPidCount) -gt 0) {
    Write-Step "Closing $(Get-WordPidCount) Word process(es) already running"
    $start = Close-AllWord
    if (-not $start.Closed) {
        throw ("Word would not close ({0}: {1}) - close it by hand, saving or discarding as you like, then re-run." -f
               $start.Reason, (Format-WordWindow $start.Dialog))
    }
    Start-Sleep -Seconds 2
}

$scratch = Join-Path $env:TEMP 'wordtab-check'
New-Item -ItemType Directory -Path $scratch -Force | Out-Null

# ---- 1. Word launched with no document at all -------------------------------------------------------

Write-Step 'Word launched with no document at all'
$mark = Get-LogMark
Start-Process -FilePath 'winword.exe'
if (-not (Wait-For { (Get-FrameCount) -ge 1 } 45)) { throw 'Word did not put a window up.' }
Start-Sleep -Seconds 8

Assert ((Get-FrameCount) -eq 1) "exactly one Word window ($(Get-FrameCount))"
$startFrame = Get-TheFrame
$parts = Get-Parts $startFrame
Write-Note ("frame 0x{0:X} '{1}'" -f [int64]$startFrame, $parts.Title)

Assert ($null -ne $parts.Wwf) 'it has a _WwF - the old test would have called this a document'
Assert (-not (Test-HasDocument $startFrame)) 'and the _WwF is empty, which is what "no document" actually looks like'
Assert ($null -ne $parts.Strip) 'the strip is bound anyway, so the band is carved before any document arrives'
Assert (-not [WordLayout]::IsToolWindow($startFrame)) 'it is still in Alt+Tab - nothing has made it harder to reach'
Assert ((Get-LogCount $mark 'joined') -eq 0) "it did not join the stack ($(Get-LogCount $mark 'joined') join lines)"

# What the user is actually looking at here is Word's own Start screen, which covers the whole client
# area - our strip included, exactly as Backstage does. Worth asserting rather than assuming: it is
# the reason this suite does not try to click the row until later.
[WordLayout]::Focus($startFrame) | Out-Null
Start-Sleep -Milliseconds 800
$strip = Get-StripOf $startFrame
$plusEmpty = Get-PlusAt $strip 0
$under = [WordLayout]::WindowAt($plusEmpty.X, $plusEmpty.Y)
Write-Note ("at the + position sits 0x{0:X} ({1})" -f [int64]$under, [WordLayout]::ClassOf($under))
Assert ($under -ne $strip) "on the Start screen the row is behind Word's own full-page UI, so nothing of ours is clickable"
Save-StripShot $strip 'startscreen-cold'

# ---- 2. Escape out of the Start screen: the same window becomes a document ---------------------------

Write-Step 'Opening a document while the Start screen is up'
#
# Deliberately not Escape. Escape does dismiss the Start screen and does produce a blank document,
# but injected keystrokes only land if Word has the foreground at that instant, and a run where it
# does not looks exactly like the add-in failing to notice a document. Opening a file needs no
# focus and no injected input at all, and it is the path a user takes far more often anyway.

$mark = Get-LogMark
$path = Join-Path $scratch 'wordtab-startscreen-1.rtf'
"{\rtf1\ansi WordTab start-screen check.\par}" | Set-Content -Path $path -Encoding Ascii
Start-Process -FilePath 'winword.exe' -ArgumentList "`"$path`""
Assert (Wait-For { Test-HasDocument (Get-TheFrame) } 45) 'a document appeared'
Start-Sleep -Seconds 3

Assert ((Get-FrameCount) -eq 1) "Word used the Start-screen window rather than opening another ($(Get-FrameCount))"
Assert ((Get-TheFrame) -eq $startFrame) 'literally the same window handle - so no empty tab is left behind'
Assert (Wait-For { (Get-LogCount $mark 'joined') -ge 1 } 10) 'and now that it has a document it joins the stack by itself'

# ---- 3. a second document, so the row is a real row --------------------------------------------------

Write-Step 'A second document, from the + button'
#
# Deliberately the + and not `winword.exe <file>`. Word *replaces* a pristine, unmodified blank
# document rather than opening a second window beside it - so opening a file here would leave one
# window and quietly measure nothing. Measured, and it cost a run of this suite to find.

Assert (Invoke-ConfirmedClick -What 'clicking the + with one tab in the row' -Point {
            Get-PlusAt (Get-StripOf (Get-TheFrame)) 1
        }) 'the click on + landed on the strip'
Assert (Wait-For { (Get-FrameCount) -eq 2 } 45) "two windows ($(Get-FrameCount))"
Start-Sleep -Seconds 4

$tools = @(Get-Frames | Where-Object { [WordLayout]::IsToolWindow($_) })
Assert ($tools.Count -eq 1) "the background one is out of Alt+Tab, as a stacked tab should be ($($tools.Count))"

$top = Get-TheFrame
$topStrip = Get-StripOf $top
$twoTabs = [WordLayout]::Tabs($topStrip, 2)
Assert ($twoTabs.Tabs.Length -eq 2) 'the row is laid out for two tabs'
Save-StripShot $topStrip 'startscreen-two-tabs'

# ---- 4. Word's own Ctrl+W, down to the last document -------------------------------------------------
#
# Ctrl+W is Word's *document* close, and it is the only way to reach the state this suite is about.
# WM_CLOSE is a different thing: it closes the window, and on the last window it closes Word.

Write-Step "Word's own Ctrl+W with another document still open"
# Aimed at ONE window and confirmed before it is sent. Ctrl+W closes whichever document is in front,
# so a Focus that quietly did not take does not lose the keystroke - it closes the wrong document,
# and everything after this measures a window that should not be there.
Assert (Invoke-ConfirmedKeyOn -Hwnd $top -Vk $VK_W -What 'Ctrl+W on the chosen window' -Ctrl) `
       'Ctrl+W went to the window it was aimed at'
Assert (Wait-For { (Get-FrameCount) -eq 1 } 25) "one window left ($(Get-FrameCount))"
Start-Sleep -Seconds 3

$survivor = Get-TheFrame
Assert (Test-HasDocument $survivor) 'the surviving window still has its document'
Assert (-not [WordLayout]::IsToolWindow($survivor)) 'and the last window is back in Alt+Tab'

Write-Step "Word's own Ctrl+W on the LAST document"
$mark = Get-LogMark
Assert (Invoke-ConfirmedKeyOn -Hwnd $survivor -Vk $VK_W -What 'Ctrl+W on the last document' -Ctrl) `
       'Ctrl+W went to the last document, not to whatever else was in front'
Start-Sleep -Seconds 8

Assert ((Get-WordPidCount) -gt 0) 'Word is still running - closing a document is not closing Word'
Assert ((Get-FrameCount) -eq 1) "and the window is still on screen ($(Get-FrameCount))"

$empty = Get-TheFrame
Assert ($empty -eq $survivor) 'the same window, still standing'
$emptyParts = Get-Parts $empty
Write-Note "title now '$($emptyParts.Title)'"

Assert ($null -ne $emptyParts.Wwf) 'it still has a _WwF - which is exactly why the old test got this wrong'
Assert (-not (Test-HasDocument $empty)) 'but the _WwF is empty again: no document'
Assert ((Get-LogCount $mark 'no document open') -ge 1) `
    "it left the stack, and the log says why: 'no document open' ($(Get-LogCount $mark 'no document open') lines)"
Assert (-not [WordLayout]::IsToolWindow($empty)) 'it is back in Alt+Tab - a window nobody can reach is the one thing forbidden'
Assert ($null -ne $emptyParts.Strip) 'its strip is still there, so nothing will jump when a document comes back'

# ---- 5. the row is empty: the + moved, and the tab is gone -------------------------------------------

Write-Step 'The row after the last document closed'
[WordLayout]::Focus($empty) | Out-Null
Start-Sleep -Milliseconds 800
$strip = Get-StripOf $empty
Write-Note ("strip at {0}" -f ([WordLayout]::RectOf($strip) | ForEach-Object { "$($_.Left),$($_.Top) $($_.Right - $_.Left)x$($_.Bottom - $_.Top)" }))
Save-StripShot $strip 'startscreen-emptied-row'

$plusEmpty = Get-PlusAt $strip 0
$plusOne   = Get-PlusAt $strip 1
Write-Note "+ with no tabs at $($plusEmpty.X),$($plusEmpty.Y);  + with one tab at $($plusOne.X),$($plusOne.Y)"
Assert ($plusEmpty.X -ne $plusOne.X) `
    "the no-tabs and one-tab layouts put the + $([Math]::Abs($plusOne.X - $plusEmpty.X))px apart, so a click tells them apart"

# Unlike the Start screen, nothing of Word's covers the row here - so these clicks are real.
$under = [WordLayout]::WindowAt($plusEmpty.X, $plusEmpty.Y)
Assert ($under -eq $strip) 'the row is on top and clickable now that the Start screen has gone'

# The negative, first: where the tab used to be there is now nothing.
$oneTab = [WordLayout]::Tabs($strip, 1)
$tabSpot = [WordLayout]::Center($oneTab.Tabs[0])
$overTab = [WordLayout]::WindowAt($tabSpot.X, $tabSpot.Y)
Assert ($overTab -eq $strip) 'the place the old "Word" tab occupied is still inside the strip, so the next click is a fair test'
$beforePids = Get-WordPidCount
# Confirmed to LAND, deliberately - this is the negative test, and "nothing happened" is exactly what
# a click that never arrived also produces. The whole point of this section is that the click was
# real and the row still did nothing.
Assert (Invoke-ConfirmedClick -What 'clicking where the old "Word" tab used to be' -Point {
            [WordLayout]::Center(([WordLayout]::Tabs((Get-StripOf (Get-TheFrame)), 1)).Tabs[0])
        }) 'the click really did land on the strip, so "nothing happened" means the row ignored it'
Start-Sleep -Seconds 3
Assert ((Get-WordPidCount) -eq $beforePids) 'clicking where the "Word" tab used to be does not close Word'
Assert ((Get-FrameCount) -eq 1) "nor open anything ($(Get-FrameCount) windows)"
Assert (-not (Test-HasDocument (Get-TheFrame))) 'and it certainly does not conjure a document'

# ---- 6. + brings a document back, into this same window ----------------------------------------------

Write-Step 'Pressing + to come back from empty'
$mark = Get-LogMark
Assert (Invoke-ConfirmedClick -What 'clicking the + on an empty row' -Point {
            Get-PlusAt (Get-StripOf (Get-TheFrame)) 0
        }) 'the click on + landed on the strip'
Assert (Wait-For { Test-HasDocument (Get-TheFrame) } 25) 'a document is back - so that really was the +, and the row really was empty'
Start-Sleep -Seconds 3
Assert ((Get-FrameCount) -eq 1) "Word put it in this same window rather than opening another ($(Get-FrameCount) windows)"
Assert (Wait-For { (Get-LogCount $mark 'joined') -ge 1 } 10) 'and it rejoins the stack'
Save-StripShot (Get-StripOf (Get-TheFrame)) 'startscreen-back-to-one-tab'

# ---- 6b. and the same round trip through the path a user is most likely to take ----------------------
#
# Close your last document, then double-click a file in Explorer. Word puts it in the window that is
# already standing there, which is the outcome that leaves no stray tab behind. If Word ever stopped
# doing that, the empty window would still be sitting in the row alongside the file that just opened.

Write-Step 'Emptying it again, then opening a real file from outside'
Assert (Invoke-ConfirmedKeyOn -Hwnd (Get-TheFrame) -Vk $VK_W -What 'Ctrl+W to empty the window again' -Ctrl) `
       'Ctrl+W went to the window it was aimed at'
Assert (Wait-For { -not (Test-HasDocument (Get-TheFrame)) } 20) 'empty again'
Start-Sleep -Seconds 2

$mark = Get-LogMark
$path2 = Join-Path $scratch 'wordtab-startscreen-2.rtf'
"{\rtf1\ansi WordTab start-screen check, second file.\par}" | Set-Content -Path $path2 -Encoding Ascii
Start-Process -FilePath 'winword.exe' -ArgumentList "`"$path2`""
Assert (Wait-For { Test-HasDocument (Get-TheFrame) } 45) 'the file opened'
Start-Sleep -Seconds 3
Assert ((Get-FrameCount) -eq 1) "into the window that was already there, leaving no stray tab ($(Get-FrameCount) windows)"
Assert ((Get-TheFrame) -eq $empty) 'the very window whose document had been closed'
Assert (Wait-For { (Get-LogCount $mark 'joined') -ge 1 } 10) 'and it joins the row'

# ---- 7. the predicate must survive a minimise ---------------------------------------------------------
#
# The stack keeps minimised windows as members on purpose: when the whole stack goes down together,
# dropping every window would leave nothing to bring back and only one would return. A document test
# that failed while minimised would break that, silently.

Write-Step 'A minimised document is still a document'
$frame = Get-TheFrame
[WordLayout]::Focus($frame) | Out-Null
Start-Sleep -Milliseconds 700
[WordLayout]::Show($frame, $SW_MINIMIZE)
Start-Sleep -Seconds 4

Assert ([WordLayout]::Minimized($frame)) 'the window is minimised'
Assert (Test-HasDocument $frame) 'and its _WwF still holds the document, so it stays a member'

[WordLayout]::Show($frame, [WordLayout]::SW_RESTORE)
Start-Sleep -Seconds 4
Assert (-not [WordLayout]::Minimized($frame)) 'restored'
Assert (Test-HasDocument $frame) 'still a document'
Assert (-not [WordLayout]::IsToolWindow($frame)) 'and reachable'

# ---- close down ----------------------------------------------------------------------------------------

if (-not $KeepOpen) {
    Write-Step 'Closing Word'
    $end = Close-AllWord
    if (-not $end.Closed) {
        Write-Note ("Word is still up ({0}): {1}" -f $end.Reason, (Format-WordWindow $end.Dialog))
        Write-Note 'Left running rather than killed. Answer it by hand before the next suite.'
    }
}

Write-Host ''
if ($script:Failures -eq 0) {
    Write-Host "PASS  $($script:Checks) checks, 0 failures" -ForegroundColor Green
} else {
    Write-Host "FAIL  $($script:Failures) of $($script:Checks) checks failed" -ForegroundColor Red
}
Write-Host "Add-in log: $env:LOCALAPPDATA\WordTab\wordtab.log (lines starting 'strip' and 'stack')" -ForegroundColor Gray
exit $script:Failures
