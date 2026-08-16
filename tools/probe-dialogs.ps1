<#
.SYNOPSIS
  Measure what Word actually puts on screen, so the harness can tell a question from a gallery.

.DESCRIPTION
  The check suites have to answer one question from outside Word: "is Word asking the user
  something?" Three suites grew three different answers and none of them was right, because the
  answer was reasoned about rather than measured. This is the measurement.

  It photographs nothing. It enumerates every top-level window Word owns - class, size, title, and
  the text of every child control - in each state that matters:

    idle        a document open and nothing else going on
    menu        our own tab context menu up
    backstage   the File menu open
    prompt      a modified document being closed, so Word asks whether to save
    saveas      the Save As question
    shutdown    sampled every 250ms while Word closes, because the transient chrome Word shows on
                its way down is exactly what a single sample mistakes for a question

  Everything it prints is evidence for Get-WordWindowKind in tools\WordTabHarness.ps1. Re-run it
  before ever changing that classifier, and on any machine where the harness starts seeing dialogs
  that are not there.

  What it must NOT be used for: matching on the text. Word's prompts carry no window title and their
  static children's text is localised. The text is dumped so a human can read it; the classifier
  reads class and size only. The tab-name slice already learned not to parse Word's English.

.PARAMETER KeepOpen
  Leave Word running at the end.

.EXAMPLE
  pwsh -File tools\probe-dialogs.ps1
#>
[CmdletBinding()]
param(
    [switch]$KeepOpen,
    [string]$ScratchDir = (Join-Path $env:TEMP 'WordTabProbeDialogs')
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

$VK = @{ ESC = 0x1B; W = 0x57; S = 0x53; X = 0x58; F12 = 0x7B }

# Every window, visible or not, with what the classifier currently calls it and - for anything that
# is not a document frame - the text of its child controls. GetWindowText on a static control is how
# a dialog says what it is asking; there is no other handle on it from outside the process.
function Show-State($label) {
    Write-Step "STATE: $label"
    $any = $false
    foreach ($id in @(Get-WordPidList)) {
        foreach ($w in [WordLayout]::TopLevel($id)) {
            $width = $w.Right - $w.Left
            $height = $w.Bottom - $w.Top
            if (-not $w.Visible) { continue }
            if ($width -le 1 -or $height -le 1) { continue }
            $any = $true
            $kind = Get-WordWindowKind $w.Class $width $height
            $colour = switch ($kind) { 'question' { 'Yellow' } 'frame' { 'DarkGray' } default { 'Gray' } }
            Write-Host ("    pid={0,-6} {1,-24} {2,5}x{3,-5} kind={4,-9} |{5}|" -f
                        $id, $w.Class, $width, $height, $kind, $w.Title) -ForegroundColor $colour

            if ($w.Class -eq 'OpusApp') { continue }
            $texts = @()
            foreach ($c in [WordLayout]::Children($w.Hwnd)) {
                $t = [WordLayout]::TitleOf($c.Hwnd)
                if ($t -and $t.Trim()) { $texts += ("{0}=|{1}|" -f $c.Class, $t.Trim()) }
            }
            if ($texts.Count -gt 0) {
                foreach ($t in ($texts | Select-Object -First 12)) { Write-Host "        $t" -ForegroundColor DarkCyan }
                if ($texts.Count -gt 12) { Write-Host ("        ... and {0} more" -f ($texts.Count - 12)) -ForegroundColor DarkCyan }
            } else {
                Write-Host '        (no child control has any text)' -ForegroundColor DarkCyan
            }
        }
    }
    if (-not $any) { Write-Note '(no visible Word windows)' }
    Write-Host ''
}

# What Get-WordDialog would answer right now, and whether that is the right answer.
function Show-Verdict($label, $expected) {
    $d = Get-WordDialog
    $got = if ($d) { Format-WordWindow $d } else { '(nothing)' }
    $ok  = if ($d) { $expected -eq 'question' } else { $expected -eq 'nothing' }
    $mark = if ($ok) { 'AGREES ' } else { 'DISAGREES' }
    $colour = if ($ok) { 'Green' } else { 'Red' }
    Write-Host ("    Get-WordDialog on `"{0}`": {1}  expected {2}, got {3}" -f
                $label, $mark, $expected, $got) -ForegroundColor $colour
    Write-Host ''
}

# The @() is not decoration. A PowerShell function returning ONE item returns that item, not a
# one-element array - so `$frames.Count` on a single Word window is "the property 'Count' cannot be
# found", which is the same trap as the empty case from the other direction. Caught by this probe on
# its first run, which is the only state it ever runs in: exactly one Word window.
function Get-TheFrame {
    $frames = @(Get-WordFrameList)
    if ($frames.Count -eq 0) { throw 'No Word window.' }
    return $frames[0]
}

function Get-Strip($frame) {
    return @([WordLayout]::Children($frame) | Where-Object { $_.Class -eq 'WordTabStrip' }) | Select-Object -First 1
}

# ---- run ----------------------------------------------------------------------------------------

Write-Host ''
Write-Host 'probe-dialogs - what Word puts on screen, measured' -ForegroundColor White
Write-Host ''

$closing = Close-AllWord
if (-not $closing.Closed) {
    throw ("Word was already up and would not close ({0}: {1}). Close it by hand and re-run." -f
           $closing.Reason, (Format-WordWindow $closing.Dialog))
}

if (-not (Test-Path $ScratchDir)) { New-Item -ItemType Directory -Path $ScratchDir | Out-Null }
$doc = Join-Path $ScratchDir 'probe-dialogs.rtf'
# A real RTF, written as text. Word opens it as a genuine document, which is what the states below
# need - a fixture Word will not treat as read-only or sandboxed.
Set-Content -Path $doc -Value '{\rtf1\ansi Probe fixture for probe-dialogs.ps1.\par}' -Encoding ASCII

try {
    Write-Step 'Starting Word on the fixture'
    Start-Process -FilePath 'winword.exe' -ArgumentList "`"$doc`""
    if (-not (Wait-WordReady 1 60)) { throw 'Word did not come up with a strip on it.' }
    Start-Sleep -Seconds 2
    Set-WordForeground | Out-Null

    Show-State 'idle'
    Show-Verdict 'idle' 'nothing'

    # ---- the tab context menu ------------------------------------------------------------------
    Write-Step 'Opening the tab context menu'
    $frame = Get-TheFrame
    $strip = Get-Strip $frame
    if ($strip) {
        $layout = [WordLayout]::Tabs($strip.Hwnd, (Get-WordFrameTally))
        if ($layout.Tabs.Length -gt 0) {
            $t = $layout.Tabs[0]
            $x = $t.Left + [int](($t.Right - $t.Left) / 3)
            $y = [int](($t.Top + $t.Bottom) / 2)
            $cls = Get-ClassAt $x $y
            if ($cls -eq 'WordTabStrip') {
                [WordLayout]::RightClick($x, $y)
                Start-Sleep -Milliseconds 900
                Show-State 'menu'
                Show-Verdict 'menu up' 'nothing'
                Invoke-ConfirmedKey -Vk $VK.ESC -What 'closing the menu' | Out-Null
                Start-Sleep -Milliseconds 700
            } else {
                Write-Note "the tab point is over `"$cls`", not the strip - skipping the menu state"
            }
        }
    } else {
        Write-Note 'no WordTabStrip on the frame - is the add-in installed?'
    }

    # ---- Backstage -----------------------------------------------------------------------------
    Write-Step 'Opening Backstage'
    Set-WordForeground | Out-Null
    $wwf = @([WordLayout]::Children((Get-TheFrame)) | Where-Object { $_.Class -eq '_WwF' }) | Select-Object -First 1
    if ($wwf) {
        $view = [WordLayout]::RectOf($wwf.Hwnd)
        [WordLayout]::Click(($view.Left + [int](($view.Right - $view.Left) / 3)),
                            ($view.Top + [int](($view.Bottom - $view.Top) / 2)))
        Start-Sleep -Milliseconds 600
    }
    [WordLayout]::OpenBackstage()
    Start-Sleep -Seconds 2
    Show-State 'backstage'
    Show-Verdict 'backstage' 'nothing'
    [WordLayout]::CloseBackstage()
    Start-Sleep -Seconds 2

    # ---- the save prompt -----------------------------------------------------------------------
    Write-Step 'Modifying the document and asking Word to close it'
    Set-WordForeground | Out-Null
    $wwf = @([WordLayout]::Children((Get-TheFrame)) | Where-Object { $_.Class -eq '_WwF' }) | Select-Object -First 1
    $view = [WordLayout]::RectOf($wwf.Hwnd)
    $tx = $view.Left + [int](($view.Right - $view.Left) / 3)
    $ty = $view.Top  + [int](($view.Bottom - $view.Top) / 2)
    if ((Get-ClassAt $tx $ty) -notin @('_WwG', '_WwF', '_WwB')) {
        Write-Note ("the typing point is over `"{0}`" - typing anyway, the class list may be incomplete" -f (Get-ClassAt $tx $ty))
    }
    [WordLayout]::Click($tx, $ty)
    Start-Sleep -Milliseconds 700
    for ($k = 0; $k -lt 6; $k++) { [WordLayout]::Press($VK.X) }
    Start-Sleep -Milliseconds 800

    Invoke-ConfirmedKey -Vk $VK.W -What 'Ctrl+W to close the document' -Ctrl | Out-Null
    if (Wait-Until { $null -ne (Get-WordDialog) } 12 300) {
        Show-State 'prompt'
        Show-Verdict 'save prompt' 'question'
    } else {
        Write-Note 'no prompt appeared - dumping anyway'
        Show-State 'prompt (none appeared)'
    }
    Invoke-ConfirmedKey -Vk $VK.ESC -What 'cancelling the save prompt' | Out-Null
    Start-Sleep -Seconds 2
    Show-Verdict 'after Escape' 'nothing'

    # ---- Save As -------------------------------------------------------------------------------
    Write-Step 'Opening Save As with F12'
    Set-WordForeground | Out-Null
    Invoke-ConfirmedKey -Vk $VK.F12 -What 'F12 for Save As' | Out-Null
    if (Wait-Until { $null -ne (Get-WordDialog) } 12 300) {
        Show-State 'saveas'
        Show-Verdict 'Save As' 'question'
    } else {
        Write-Note 'no Save As window appeared - dumping anyway'
        Show-State 'saveas (none appeared)'
    }
    Invoke-ConfirmedKey -Vk $VK.ESC -What 'cancelling Save As' | Out-Null
    Start-Sleep -Seconds 2

    # ---- shutdown ------------------------------------------------------------------------------
    #
    # The state the double sample in Get-WordSavePrompt exists for. Word shows chrome on its way down
    # and a single read of it says "Word is asking something". Sampled fast enough to catch it.
    Write-Step 'Saving the fixture, then sampling every 250ms while Word closes'
    Set-WordForeground | Out-Null
    Invoke-ConfirmedKey -Vk $VK.S -What 'Ctrl+S to save the fixture' -Ctrl | Out-Null
    Start-Sleep -Seconds 3
    if (Get-WordDialog) {
        Write-Note ("Word asked something on save: {0} - Escaping" -f (Format-WordWindow (Get-WordDialog)))
        Invoke-ConfirmedKey -Vk $VK.ESC -What 'clearing the save question' | Out-Null
        Start-Sleep -Seconds 2
    }

    $seen = @{}
    foreach ($f in (Get-WordFrameList)) { [WordLayout]::Close($f) }
    $deadline = (Get-Date).AddSeconds(20)
    while ((Get-Date) -lt $deadline) {
        foreach ($w in (Get-WordWindows)) {
            $key = "{0}|{1}x{2}|{3}" -f $w.Class, $w.Width, $w.Height, $w.Kind
            if (-not $seen.ContainsKey($key)) {
                $seen[$key] = 1
                Write-Host ("    +{0,6:n1}s  {1,-24} {2,5}x{3,-5} kind={4,-9} |{5}|" -f
                            (20 - ($deadline - (Get-Date)).TotalSeconds), $w.Class, $w.Width, $w.Height,
                            $w.Kind, $w.Title) -ForegroundColor Gray
            } else { $seen[$key]++ }
        }
        if ((Get-WordPidTally) -eq 0) { break }
        Start-Sleep -Milliseconds 250
    }
    Write-Host ''
    Write-Step 'Distinct windows seen during shutdown, with how many samples each survived'
    foreach ($k in ($seen.Keys | Sort-Object)) {
        $samples = $seen[$k]
        # One sample is roughly 250ms. Anything that survives two samples would also survive the
        # 1500ms double-read in Get-WordSavePrompt, which is the number that matters.
        $note = if ($samples -le 1) { 'transient - a single read would have caught it, a double read would not' } else { "{0} samples (~{1:n1}s)" -f $samples, ($samples * 0.25) }
        Write-Host ("    {0,-56} {1}" -f $k, $note) -ForegroundColor DarkGray
    }
    Write-Host ''
}
finally {
    if (-not $KeepOpen) {
        $end = Close-AllWord
        if (-not $end.Closed) {
            Write-Host ("    Word is still up ({0}): {1}" -f $end.Reason, (Format-WordWindow $end.Dialog)) -ForegroundColor Yellow
        }
    }
    Remove-Item -Path $ScratchDir -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host 'Probe complete. Every line above is evidence for Get-WordWindowKind.' -ForegroundColor Green
