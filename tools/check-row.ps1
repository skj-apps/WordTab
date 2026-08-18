<#
.SYNOPSIS
  The tab ROW: what order tabs are in, what keeps a window in it, and what closing the window does.

.DESCRIPTION
  Three claims, all of them from the user's own report after a first afternoon of real use, and all
  three about membership and order rather than about geometry - which is what every other suite here
  measures.

    1. **A new document's tab goes to the END of the row.** Their words: "now newly opened docs
       popping infront lets make them pop to end... i like to keep same 1 or 2 in fromt." Checked
       both ways round: opening a document into a fresh Word, and opening one after a tab in the
       middle of the row has been closed, which is the case that put a new document in the middle -
       Word does not destroy a frame when its last document closes, it hides it and hands the next
       document to it, and that frame holds a slot in the row that is not the end.

    2. **A window is not dropped from the row for one bad reading.** Their words: "changing setting
       like view>page width closes tabs", and the document survives it - the window is evicted and
       comes back as a plain Word window with no strip. The trigger has not been reproduced on this
       machine (tools\probe-view.ps1 drives every View command Word's object model offers and the
       row never moves), so what is checked here is the GUARD rather than the trigger: a window
       hidden and shown again inside one janitor tick must keep its tab.

    3. **Closing the window closes the stack.** Their words: "you have to close all tabs
       individually - that needs fixing next time we edit."

  The row order is read from the add-in's own log line - `stack  row, left to right:` - which is
  written whenever the row changes and never otherwise. There is no way to read tab order from
  outside the process: the tabs are painted, not controls. The line is therefore the instrument, and
  it is checked against something independent wherever that is possible - the number of tabs against
  the number of frames, and the name on each against the window title.

.PARAMETER KeepOpen
  Leave Word running afterwards.

.EXAMPLE
  pwsh -File tools\check-row.ps1
#>
[CmdletBinding()]
param(
    [switch]$KeepOpen
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$source = Get-Content -Raw -Path (Join-Path $PSScriptRoot 'WordLayout.cs')
Add-Type -TypeDefinition $source -Language CSharp -ReferencedAssemblies @(
    'System.Runtime', 'System.Collections', 'System.Threading.Thread', 'netstandard'
)
[WordLayout]::MakeDpiAware() | Out-Null

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

$scratch = Join-Path $env:TEMP 'wordtab-check'
New-Item -ItemType Directory -Path $scratch -Force | Out-Null

function Open-Document($n) {
    $path = Join-Path $scratch "wordtab-row-$n.rtf"
    "{\rtf1\ansi WordTab row check - document $n.\par}" | Set-Content -Path $path -Encoding Ascii
    Start-Process -FilePath 'winword.exe' -ArgumentList "`"$path`""
    return $path
}

# The row the add-in last wrote, as an array of tab names, or $null when it has not written one since
# the mark. Names rather than handles: a name can be checked against a window title, and a handle in
# a log line can only be checked against another log line.
#
# **$null and an empty row are different answers and are kept apart.** "The add-in has not said" is
# not "the row is empty", and folding the two together is exactly how an assertion passes having
# measured nothing - the failure this project keeps meeting.
function Get-Row {
    $line = Get-LogLast 'stack  row'
    if ($null -eq $line) { return $null }
    if ($line -like '*row: empty*') { return @() }

    $tail = ($line -split 'left to right: ', 2)[-1]
    return @([regex]::Matches($tail, '\|([^|]*)\|') | ForEach-Object { $_.Groups[1].Value })
}

function Wait-Row($count, $seconds = 20) {
    Wait-Until { $r = Get-Row; ($null -ne $r) -and (@($r).Count -eq $count) } $seconds 250 | Out-Null
    return Get-Row
}

function Format-Row($row) {
    if ($null -eq $row) { return '(the add-in has not written a row line)' }
    if (@($row).Count -eq 0) { return '(empty)' }
    return ($row -join ' | ')
}

Close-AllWord | Out-Null
Reset-LogFile -Label 'row' -Quiet | Out-Null

# ---- 1. three documents, opened in order -------------------------------------------------------

Write-Step 'Three documents, opened one at a time'
Set-LogMark
foreach ($n in 1..3) {
    Open-Document $n | Out-Null
    if (-not (Wait-WordReady $n 45)) { Write-Note "document $n did not arrive with a strip within 45s" }
}
Start-Sleep -Seconds 2

$row = Wait-Row 3
Write-Note ("row: " + (Format-Row $row))
Assert ($null -ne $row) 'the add-in wrote a row line'
Assert (@($row).Count -eq (Get-WordFrameTally)) "the row has one tab per Word window ($(@($row).Count) tabs, $(Get-WordFrameTally) windows)"
Assert (@($row).Count -eq 3 -and $row[0] -like '*row-1*' -and $row[1] -like '*row-2*' -and $row[2] -like '*row-3*') `
       'they are in the order they were opened, left to right'

# ---- 2. a fourth document goes to the END ------------------------------------------------------

Write-Step 'A fourth document, with three already open'
$stackRect = [WordLayout]::RectOf((Get-WordFrameList)[0])
$stackText = "({0},{1} {2}x{3})" -f $stackRect.Left, $stackRect.Top,
                                    ($stackRect.Right - $stackRect.Left), ($stackRect.Bottom - $stackRect.Top)
Set-LogMark
Open-Document 4 | Out-Null
if (-not (Wait-WordReady 4 45)) { Write-Note 'document 4 did not arrive with a strip within 45s' }
Start-Sleep -Seconds 2

$row = Wait-Row 4
Write-Note ("row: " + (Format-Row $row))
Assert (@($row).Count -eq 4) "the row has four tabs (has $(@($row).Count))"
Assert (@($row).Count -eq 4 -and $row[3] -like '*row-4*') 'the new document is the LAST tab, not the first'
Assert (@($row).Count -eq 4 -and $row[0] -like '*row-1*') 'and the tab that was on the left is still on the left'

# ...and it was BORN there, which is the other half of the queue's last item: opening a document
# used to show its window a few inches above the stack before it settled into it. The add-in writes
# the stack's rectangle into the CREATESTRUCT inside the CBT hook, so the window is created on the
# stack and has nowhere to flash from.
#
# The evidence is the rect on the `attach` line, which is read after the window exists and before
# anything has moved it. A flash cannot be photographed reliably - it is one frame - but a window
# created at the right coordinates cannot produce one.
$born = @(Get-LogSince 'attach  hwnd' | Where-Object { $_ -like '*(new frame)*' })
Write-Note ("the stack is at $stackText; new frames were born at: " +
            (($born | ForEach-Object { ($_ -split 'rect=')[-1] -replace '\s+frames=.*', '' }) -join ' '))
Assert ($born.Count -ge 1) "at least one new frame was created and attached ($($born.Count))"
Assert (($born.Count -ge 1) -and (@($born | Where-Object { $_ -like "*rect=$stackText*" }).Count -eq $born.Count)) `
       'every window created while the stack was up was created ON the stack, not beside it'

# ---- 3. close a MIDDLE tab, then open another --------------------------------------------------
#
# This is the case the user reported. Word hides the frame whose last document closed and hands the
# next document straight to it, so the new document arrives holding a slot in the middle of the row
# unless something puts it at the end.

Write-Step 'Close the middle tab, then open a fifth document'
$before = @(Get-WordFrameList)
$middle = $null
foreach ($f in $before) {
    if ([WordLayout]::TitleOf($f) -like '*row-2*') { $middle = $f }
}
Assert ($null -ne $middle) 'the second document was found to close'

if ($null -ne $middle) {
    Set-LogMark
    [WordLayout]::Close($middle)
    Wait-Until { (Get-WordFrameTally) -lt $before.Count } 15 250 | Out-Null
    Start-Sleep -Seconds 2

    $row = Wait-Row 3
    Write-Note ("row after the close: " + (Format-Row $row))
    Assert (@($row).Count -eq 3) "closing one tab left three ($(@($row).Count))"

    Set-LogMark
    Open-Document 5 | Out-Null
    Wait-Until { (Get-WordFrameTally) -ge 4 } 45 500 | Out-Null
    Start-Sleep -Seconds 3

    $row = Wait-Row 4
    Write-Note ("row after opening the fifth: " + (Format-Row $row))
    Assert (@($row).Count -eq 4) "the row has four tabs again (has $(@($row).Count))"
    Assert (@($row).Count -eq 4 -and $row[3] -like '*row-5*') `
           'the document opened into a recycled frame still gets the LAST tab'
    Assert (@($row).Count -ge 1 -and $row[0] -like '*row-1*') 'and the leftmost tab has not moved'
}

# ---- 4. a window hidden and shown again keeps its tab ------------------------------------------
#
# The guard for "a View change closes tabs". Membership is re-derived twice a second from what is
# true about the window, and Word hides and re-shows its frames of its own accord - so a window that
# is invisible for a moment must not lose its place. Driven by hiding one, because that is a state
# the add-in reads exactly as it reads the real thing, and one this machine can produce on demand.

Write-Step 'A window hidden for a fraction of a tick keeps its tab'
$frames = @(Get-WordFrameList)
if ($frames.Count -lt 2) {
    Write-Note 'fewer than two windows - skipping'
} else {
    $victim = $frames[0]
    foreach ($f in $frames) { if ([WordLayout]::TitleOf($f) -like '*row-1*') { $victim = $f } }
    $rowBefore = Get-Row
    Write-Note ("row before: " + (Format-Row $rowBefore))

    Set-LogMark
    [WordLayout]::Show($victim, 0)      # SW_HIDE
    Start-Sleep -Milliseconds 250
    [WordLayout]::Show($victim, 5)      # SW_SHOW
    Start-Sleep -Seconds 3

    # Two spaces, and they are load-bearing. The add-in writes BOTH `hwnd=0x..  left the stack (...)`
    # when a window really goes and `hwnd=0x..  would have left the stack (...) - waiting` when it
    # declines to give up on one - and the second contains the first as a substring. The first draft
    # of this check counted the guard's own message as the failure it was guarding against.
    $left = Get-LogCount '  left the stack'
    Assert ($left -eq 0) "no window left the stack over a 250ms hide (the log said so $left time(s))"

    $waited = Get-LogCount 'would have left the stack'
    Assert ($waited -ge 1) "and the add-in said out loud that it was waiting rather than acting ($waited)"

    # The row line is written only when the row CHANGES, so silence here is the assertion: not one
    # was written, therefore nothing about the row moved. Stronger than comparing two reads, which
    # would also pass if the row had changed and changed back.
    $rowLines = Get-LogCount 'stack  row'
    Assert ($rowLines -eq 0) "and the row never changed at all - no new row line ($rowLines)"
    Write-Note ("row is still: " + (Format-Row $rowBefore))
}

# ---- 5. the window's own close takes the whole stack with it -----------------------------------
#
# SC_CLOSE, which is the message the title bar's x raises. Posted rather than clicked: the x is
# painted by Word inside its own caption and hitting it by coordinate is a guess, whereas the message
# it produces is not. The add-in logs SC_CLOSE when it arrives, so a run that wanted to check the
# real button can compare that line against this one.

Write-Step "The window's own close button closes every tab"
$frames = @(Get-WordFrameList)
Assert ($frames.Count -ge 2) "there are at least two windows to close together ($($frames.Count))"

if ($frames.Count -ge 2) {
    Set-LogMark
    [WordLayout]::SysClose($frames[0])

    # Generously bounded. Each close is one WM_CLOSE stepped by the janitor at half a second, and any
    # of them may raise a save prompt - these are scratch files nobody has typed into, so none should,
    # and a prompt appearing here is a finding rather than something to wait out.
    $gone = Wait-Until { (Get-WordFrameTally) -eq 0 } 40 500

    $prompt = Get-WordSavePrompt
    Assert ($null -eq $prompt) 'Word asked nothing (no document had unsaved changes)'
    Assert ($gone) "every window closed, not just the one ($(Get-WordFrameTally) left)"
    Assert ((Get-LogCount 'SC_CLOSE') -ge 1) 'the add-in saw the close as SC_CLOSE'
    Assert ((Get-LogCount 'the window''s own close') -ge 1) 'and took it over as a whole-stack close'
}

# ---- summary -----------------------------------------------------------------------------------

Write-Host ''
if (-not $KeepOpen) { Close-AllWord | Out-Null }

if ($script:Failures -eq 0) {
    Write-Host "PASS  $($script:Checks) checks, 0 failures" -ForegroundColor Green
    exit 0
}
Write-Host "FAIL  $($script:Checks) checks, $($script:Failures) failures" -ForegroundColor Red
exit 1
