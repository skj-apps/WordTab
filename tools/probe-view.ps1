<#
.SYNOPSIS
  What a View change does to a stacked Word window - the document frame, the strip, and membership.

.DESCRIPTION
  The user's report: "changing setting like view>page width closes tabs". The document survives -
  "it just opened a new word w/out it" - so the window is EVICTED from the row and comes back as a
  plain Word window with no strip on it.

  This asks the machine what actually happens, rather than reasoning about it. For each view change
  it records, per frame:

    - every `_WwF` child of the frame, whether it is visible, and whether anything is inside it.
      **More than one is the finding to look for.** StripHasDocument caches the `_WwF` it bound to
      and only rebinds when that handle is DESTROYED; if Word hides the old document frame and
      builds a second one beside it, the add-in keeps reading the empty one forever - which reads as
      "no document", drops the window out of the stack, and hides the strip with it.
    - whether the strip exists and is visible
    - what the add-in logged while the change was happening

  Word's documents on this rig are disposable test fixtures - see the memory note.

.PARAMETER Documents
  How many documents to open. Three, so that "one window left the row" is visible as a row of two.

.PARAMETER KeepOpen
  Leave Word running afterwards.

.EXAMPLE
  pwsh -File tools\probe-view.ps1
#>
[CmdletBinding()]
param(
    [int]$Documents = 3,
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

function Write-Step($text) { Write-Host "==> $text" -ForegroundColor Cyan }
function Write-Note($text) { Write-Host "    $text" -ForegroundColor DarkGray }

# One frame's document frames and strip, as a line of text. Everything the hypothesis turns on is in
# here: how many `_WwF` there are, which of them are visible, and which of them hold anything.
function Format-Frame($frame) {
    $kids = [WordLayout]::Children($frame)
    $wwfs = @($kids | Where-Object { $_.Class -eq '_WwF' })
    $strip = @($kids | Where-Object { $_.Class -eq 'WordTabStrip' }) | Select-Object -First 1

    $parts = foreach ($w in $wwfs) {
        $inside = [WordLayout]::FirstChild($w.Hwnd)
        $what = if ($inside -eq [IntPtr]::Zero) { 'EMPTY' } else { [WordLayout]::ClassOf($inside) }
        "_WwF=0x{0:X} vis={1} holds={2} ({3},{4} {5}x{6})" -f `
            [int64]$w.Hwnd, [int]$w.Visible, $what, $w.Left, $w.Top, $w.Width, $w.Height
    }

    $stripText = if ($null -eq $strip) { 'NO STRIP' }
                 else { "strip=0x{0:X} vis={1} y={2}..{3}" -f [int64]$strip.Hwnd, [int]$strip.Visible, $strip.Top, $strip.Bottom }

    "0x{0:X}  {1}  |  {2}  |  {3}" -f [int64]$frame, $stripText, ($parts -join '  '), [WordLayout]::TitleOf($frame)
}

function Show-Snapshot($label) {
    Write-Host "    --- $label" -ForegroundColor Yellow
    foreach ($f in Get-WordFrameList) { Write-Note (Format-Frame $f) }
}

# ---- fixtures ----------------------------------------------------------------------------------

Close-AllWord | Out-Null
Reset-LogFile -Label 'probe-view' -Quiet | Out-Null

$scratch = Join-Path $env:TEMP 'wordtab-check'
New-Item -ItemType Directory -Path $scratch -Force | Out-Null

for ($i = 1; $i -le $Documents; $i++) {
    $path = Join-Path $scratch "wordtab-view-$i.rtf"
    "{\rtf1\ansi WordTab view probe - document $i.\par}" | Set-Content -Path $path -Encoding Ascii
    Start-Process -FilePath 'winword.exe' -ArgumentList "`"$path`""
    if (-not (Wait-WordReady $i 45)) { Write-Note "document $i did not arrive with a strip within 45s" }
}

Start-Sleep -Seconds 3
Write-Step "Baseline: $((Get-WordFrameList).Count) frame(s)"
Show-Snapshot 'before any view change'

# ---- the changes, one at a time, on whichever window is in front --------------------------------
#
# Driven through Word's own object model rather than through the ribbon: the ribbon button and the
# property set the same thing, and a menu path that has to be found by keytips is a second thing that
# can fail. If none of these reproduces it, the ribbon is the next thing to try.

$changes = @(
    @{ What = 'View > Page Width (Zoom.PageFit = wdPageFitBestFit)'; Do = { param($om) $om.ActiveWindow.View.Zoom.PageFit = 2 } }
    @{ What = 'Zoom back to 100%';                                   Do = { param($om) $om.ActiveWindow.View.Zoom.Percentage = 100 } }
    @{ What = 'View > Web Layout (View.Type = wdWebView)';           Do = { param($om) $om.ActiveWindow.View.Type = 6 } }
    @{ What = 'View > Print Layout (View.Type = wdPrintView)';       Do = { param($om) $om.ActiveWindow.View.Type = 3 } }
    @{ What = 'View > Draft (View.Type = wdNormalView)';             Do = { param($om) $om.ActiveWindow.View.Type = 1 } }
    @{ What = 'View > Print Layout again';                           Do = { param($om) $om.ActiveWindow.View.Type = 3 } }

    # ...and the same commands the RIBBON runs, which is not the same act as setting the property.
    # ExecuteMso raises Word's own command by the id the button carries, so anything the button does
    # around the property change happens too.
    @{ What = 'ribbon: ZoomPageWidth';          Do = { param($om) $om.CommandBars.ExecuteMso('ZoomPageWidth') } }
    @{ What = 'ribbon: Zoom100';                Do = { param($om) $om.CommandBars.ExecuteMso('Zoom100') } }
    @{ What = 'ribbon: ViewWebLayoutView';      Do = { param($om) $om.CommandBars.ExecuteMso('ViewWebLayoutView') } }
    @{ What = 'ribbon: ViewPrintLayoutView';    Do = { param($om) $om.CommandBars.ExecuteMso('ViewPrintLayoutView') } }
    @{ What = 'ribbon: ViewDraftView';          Do = { param($om) $om.CommandBars.ExecuteMso('ViewDraftView') } }
    @{ What = 'ribbon: ViewPrintLayoutView';    Do = { param($om) $om.CommandBars.ExecuteMso('ViewPrintLayoutView') } }
    @{ What = 'ribbon: ViewRulerShowHide';      Do = { param($om) $om.CommandBars.ExecuteMso('ViewRulerShowHide') } }
    @{ What = 'ribbon: ViewNavigationPane';     Do = { param($om) $om.CommandBars.ExecuteMso('NavigationPane') } }
    @{ What = 'ribbon: WindowSplit';            Do = { param($om) $om.CommandBars.ExecuteMso('WindowSplit') } }
    @{ What = 'ribbon: WindowRemoveSplit';      Do = { param($om) $om.CommandBars.ExecuteMso('WindowRemoveSplit') } }
    @{ What = 'ribbon: ViewSideBySide';         Do = { param($om) $om.CommandBars.ExecuteMso('ViewSideBySide') } }
    @{ What = 'ribbon: WindowArrangeAll';       Do = { param($om) $om.CommandBars.ExecuteMso('WindowArrangeAll') } }
    @{ What = 'ribbon: ViewReadingView (Read Mode)'; Do = { param($om) $om.CommandBars.ExecuteMso('ViewReadingView') } }
    @{ What = 'back out of Read Mode';          Do = { param($om) $om.ActiveWindow.View.Type = 3 } }
)

foreach ($change in $changes) {
    $frames = @(Get-WordFrameList)
    if ($frames.Count -eq 0) { Write-Note 'no frames left - stopping'; break }

    $front = [WordLayout]::GetForeground()
    if (-not ($frames -contains $front)) { $front = $frames[0] }

    $om = [WordLayout]::NativeOm($front)
    if ($null -eq $om) { Write-Note "no object model on 0x$('{0:X}' -f [int64]$front) - skipping"; continue }

    Set-LogMark
    Write-Step $change.What
    try { & $change.Do $om.Application } catch { Write-Note "the call failed: $($_.Exception.Message)"; continue }

    Start-Sleep -Seconds 3         # six janitor ticks

    # What the command actually did, so that a command Word ignored is not read as "no defect".
    try {
        $view = $om.Application.ActiveWindow.View
        Write-Note ("view type={0} zoom={1}% pagefit={2}" -f $view.Type, $view.Zoom.Percentage, $view.Zoom.PageFit)
    } catch { Write-Note "could not read the view back: $($_.Exception.Message)" }

    Show-Snapshot 'after'

    $text  = Get-LogTextSince (Get-CurrentLogMark)
    $lines = @(($text -split "`r?`n") | Where-Object { $_ -match 'stack |bound _WwF|tab name|WM_SHOWWINDOW|NCDESTROY|attach ' })
    foreach ($line in $lines) { Write-Host "      $line" -ForegroundColor DarkCyan }
    if ($lines.Count -eq 0) { Write-Note '(the add-in logged nothing about the stack)' }
}

Write-Host ''
Write-Step 'Done'
if (-not $KeepOpen) { Close-AllWord | Out-Null }
