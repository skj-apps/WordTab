<#
.SYNOPSIS
  Push the stack through the combinations the acceptance checks do not cover.

.DESCRIPTION
  `check-stack.ps1` proves the stack works. This one tries to break it, on the cases that are
  combinations rather than features and so fall between the two check scripts:

    - Backstage opened while several windows are stacked. Each was tested alone; together, Word
      covers the active window's whole client area while N-1 other windows sit behind it.
    - A window launched with no document arriving on an established stack. It should not disturb
      what is already stacked. (The document-less states themselves - Word's Start screen, and the
      empty frame left when the last document is closed - are `check-startscreen.ps1`.)
    - Maximise as a *state* rather than a rectangle. Spike 2 could only copy rectangles, which left
      stacked windows looking maximized without being maximized; every window should now actually be
      maximized.
    - Rapid tab switching, which exercises the layout-oracle handover as fast as Word will take it.
    - Six documents, where tabs have to share a strip that is not wide enough for them.

  Everything here is a real Word window driven by real input. It is slow on purpose.

.EXAMPLE
  pwsh -File tools\soak-stack.ps1
#>
[CmdletBinding()]
param(
    [switch]$KeepOpen,
    [string]$ShotDir = $env:TEMP
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition (Get-Content -Raw -Path (Join-Path $PSScriptRoot 'WordLayout.cs')) `
    -Language CSharp -ReferencedAssemblies @('System.Runtime','System.Collections','System.Threading.Thread','netstandard')
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
    if ($condition) { Write-Host "    PASS  $text" -ForegroundColor Green }
    else { $script:Failures++; Write-Host "    FAIL  $text" -ForegroundColor Red }
}

function Get-Frames { return Get-WordFrameList }

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

# The whole invariant in one function: every stacked window at one rectangle, every strip flush
# between the chrome above it and the document below it.
function Test-Intact($label) {
    $parts = @(Get-Frames | ForEach-Object { Get-Parts $_ })
    if ($parts.Count -eq 0) { Assert $false "$label - no Word windows left"; return }

    $rects = @($parts | ForEach-Object { "($($_.Rect.Left),$($_.Rect.Top) $($_.Rect.Right - $_.Rect.Left)x$($_.Rect.Bottom - $_.Rect.Top))" } | Sort-Object -Unique)
    Assert ($rects.Count -eq 1) "$label - $($parts.Count) window(s) at one rectangle ($($rects -join ' '))"

    $bad = @()
    foreach ($p in $parts) {
        if (-not $p.Strip) { $bad += "0x$('{0:X}' -f [int64]$p.Frame) has no strip"; continue }
        if (-not $p.Wwf)   { $bad += "0x$('{0:X}' -f [int64]$p.Frame) has no document frame"; continue }
        if ($p.Strip.Bottom -ne $p.Wwf.Top) { $bad += "0x$('{0:X}' -f [int64]$p.Frame) strip bottom $($p.Strip.Bottom) != document top $($p.Wwf.Top)"; continue }

        $client = [WordLayout]::ClientOf($p.Frame)
        $above = [WordLayout]::Children($p.Frame) |
                 Where-Object { $_.Visible -and $_.Hwnd -ne $p.Strip.Hwnd -and $_.Hwnd -ne $p.Wwf.Hwnd -and
                                $_.Bottom -le $p.Strip.Top -and $_.Width -gt ($client.Right * 0.6) } |
                 Sort-Object Bottom -Descending | Select-Object -First 1
        if (-not $above) { $bad += "0x$('{0:X}' -f [int64]$p.Frame) nothing above the strip"; }
        elseif ($above.Bottom -ne $p.Strip.Top) { $bad += "0x$('{0:X}' -f [int64]$p.Frame) gap: $($above.Class) ends $($above.Bottom), strip starts $($p.Strip.Top)" }
    }
    Assert ($bad.Count -eq 0) ("$label - every strip correctly placed" + $(if ($bad.Count) { ": " + ($bad -join '; ') } else { '' }))
    return $parts
}

function Wait-Settled($expected, $seconds = 20) {
    $deadline = (Get-Date).AddSeconds($seconds)
    while ((Get-Date) -lt $deadline) {
        $parts = @(Get-Frames | ForEach-Object { Get-Parts $_ })
        $rects = @($parts | ForEach-Object { "$($_.Rect.Left),$($_.Rect.Top),$($_.Rect.Right),$($_.Rect.Bottom)" } | Sort-Object -Unique)
        $ok = @($parts | Where-Object { $_.Strip -and $_.Wwf -and $_.Strip.Bottom -eq $_.Wwf.Top })
        if ($parts.Count -ge $expected -and $ok.Count -eq $parts.Count -and $rects.Count -eq 1) { return }
        Start-Sleep -Milliseconds 500
    }
}

function Open-Document($index) {
    $scratch = Join-Path $env:TEMP 'wordtab-check'
    New-Item -ItemType Directory -Path $scratch -Force | Out-Null
    $path = Join-Path $scratch "wordtab-soak-$index.rtf"
    "{\rtf1\ansi WordTab soak - document $index.\par}" | Set-Content -Path $path -Encoding Ascii
    Start-Process -FilePath 'winword.exe' -ArgumentList "`"$path`""
}

# ---- open six documents --------------------------------------------------------------------------

Write-Step 'Opening six documents'
# Waiting for each document to arrive with a strip on it rather than sleeping 14 seconds for the
# first and 6 for each of the other five - 44 seconds of guessing on every run. This is the harness
# getting into position; the assertions below are about geometry, not about how fast Word opens files.
for ($i = 1; $i -le 6; $i++) {
    Open-Document $i
    if (-not (Wait-WordReady $i 45)) { Write-Note "document $i did not arrive with a strip on it within 45s" }
}
Wait-Settled 6 30
$parts = Test-Intact 'Six documents'
Write-Note "$((Get-Frames).Count) frames, strip width shared between them"

# ---- maximise as a state, not a rectangle --------------------------------------------------------

Write-Step 'Maximising the stack'
$active = [WordLayout]::GetForeground()
if (-not ((Get-Frames) -contains $active)) { $active = (Get-Frames)[0]; [WordLayout]::Focus($active) | Out-Null }
[WordLayout]::Show($active, 3)          # SW_MAXIMIZE
Start-Sleep -Seconds 3

$zoomed = @((Get-Frames) | Where-Object { [WordLayout]::Maximized($_) })
Assert ($zoomed.Count -eq (Get-Frames).Count) `
    "every window is actually maximized, not just the same size ($($zoomed.Count) of $((Get-Frames).Count))"
Test-Intact 'Maximized' | Out-Null

# Re-read which window is in front rather than reusing the one from before the maximize. Activation
# can move while the stack rearranges, and restoring a window that is *behind* the stack is not a
# thing a user can do - the add-in correctly puts it straight back, so asserting on it tests
# nothing but the test's own bookkeeping.
$active = [WordLayout]::GetForeground()
if (-not ((Get-Frames) -contains $active)) { $active = (Get-Frames)[0]; [WordLayout]::Focus($active) | Out-Null }

[WordLayout]::Show($active, 9)          # SW_RESTORE
Start-Sleep -Seconds 3
$zoomed = @((Get-Frames) | Where-Object { [WordLayout]::Maximized($_) })
Assert ($zoomed.Count -eq 0) "every window restored ($($zoomed.Count) still maximized)"
Test-Intact 'Restored' | Out-Null

# ---- Backstage on a stacked window ---------------------------------------------------------------

Write-Step 'Backstage while stacked'
$active = [WordLayout]::GetForeground()
if (-not ((Get-Frames) -contains $active)) { $active = (Get-Frames)[0] }
[WordLayout]::Focus($active) | Out-Null

# Aimed at the document frame's own SCREEN rectangle. The previous version added client-relative
# child coordinates to the frame's window rectangle and dropped Wwf.Left, so it was off by the border
# in x and by the caption in y - and a click that misses here makes Alt+F fail three lines later for
# a reason that looks nothing like a mis-aimed click. Same defect, same fix, as check-strip.
$parts = @(Get-Parts $active)
if ($parts[0].Wwf) {
    $view = [WordLayout]::RectOf($parts[0].Wwf.Hwnd)
    $cx = $view.Left + [int](($view.Right - $view.Left) / 2)
    $cy = $view.Top  + [int](($view.Bottom - $view.Top) / 2)
    $under = Get-ClassAt $cx $cy
    if ($under -in @('_WwG', '_WwN', '_WwB')) { [WordLayout]::Click($cx, $cy) }
    else { Write-Note "the click into the document would land on `"$under`" - skipping it" }
}

$opened = $false
for ($attempt = 1; $attempt -le 3 -and -not $opened; $attempt++) {
    # Confirmed before sending: a keystroke goes to whatever has the foreground, and a silently
    # missed Alt+F is indistinguishable from Backstage refusing to open.
    if (-not (Test-WordHasFocus)) { Set-WordForeground | Out-Null }
    [WordLayout]::CloseBackstage(); Start-Sleep -Milliseconds 300
    [WordLayout]::OpenBackstage(); Start-Sleep -Milliseconds 2000
    $opened = [WordLayout]::BackstageOpen($active)
}
Assert $opened 'Backstage opened on the active stacked window'
if ($opened) {
    $inside = Get-Parts $active
    Assert ((-not $inside.Strip) -or (-not $inside.Strip.Visible)) 'its strip is hidden while Backstage is up'
    [WordLayout]::CloseBackstage()
    Start-Sleep -Seconds 3
    Test-Intact 'After leaving Backstage, still stacked' | Out-Null
}

# ---- Word's Start screen -------------------------------------------------------------------------
#
# A window with no document is deliberately excluded from the stack - it has no document to be a tab
# for - so what matters here is that it neither joins nor disturbs what is already stacked.
#
# **The old note here said such a window "has no `_WwF`". That is false**, and believing it is what
# gave a document-less Word window a tab labelled "Word" for four slices. `_WwF` is the document
# *frame* and Word keeps it, empty, after the last document closes. The membership test looks inside
# it now - see StripHasDocument in src\native\strip.cpp and RESULT-startscreen.md.
#
# What this block actually exercises, on this rig, is a *new document* arriving on an established
# stack: launching WINWORD with no arguments while a document is already open makes Word create a
# blank document with its Start screen drawn over it. Worth having, and it is not the Start-screen
# case. The document-less states are covered properly by tools\check-startscreen.ps1.
Write-Step "A window launched with no document"
$before = @(Get-Frames)
Start-Process -FilePath 'winword.exe'
Start-Sleep -Seconds 10

$now = @(Get-Frames)
$new = @($now | Where-Object { $before -notcontains $_ })
Write-Note "$($new.Count) new frame(s) appeared"
foreach ($f in $new) {
    $kids = [WordLayout]::Children($f)
    $hasDoc = @($kids | Where-Object { $_.Class -eq '_WwF' }).Count -gt 0
    $hasStrip = @($kids | Where-Object { $_.Class -eq 'WordTabStrip' }).Count -gt 0
    Write-Note ("0x{0:X}  `"{1}`"  document frame={2}  strip={3}" -f [int64]$f, [WordLayout]::TitleOf($f), $hasDoc, $hasStrip)
}

# The documents that were already stacked must still be stacked, whatever the new window did.
$stillStacked = @($before | Where-Object { [WordLayout]::IsWindow2($_) } | ForEach-Object { Get-Parts $_ })
$rects = @($stillStacked | ForEach-Object { "$($_.Rect.Left),$($_.Rect.Top),$($_.Rect.Right),$($_.Rect.Bottom)" } | Sort-Object -Unique)
Assert ($rects.Count -eq 1) "the documents already stacked are still at one rectangle ($($rects.Count) distinct)"
$bad = @($stillStacked | Where-Object { -not $_.Strip -or -not $_.Wwf -or $_.Strip.Bottom -ne $_.Wwf.Top })
Assert ($bad.Count -eq 0) "their strips are still correctly placed ($($bad.Count) wrong)"

foreach ($f in $new) { [WordLayout]::Close($f) }
Start-Sleep -Seconds 4

# ---- rapid tab switching -------------------------------------------------------------------------

Write-Step 'Rapid tab switching'
$frames = @(Get-Frames)
$switched = 0
for ($round = 0; $round -lt 12; $round++) {
    $target = $frames[$round % $frames.Count]
    if (-not [WordLayout]::IsWindow2($target)) { continue }
    [WordLayout]::Focus($target) | Out-Null
    Start-Sleep -Milliseconds 250
    if ([WordLayout]::GetForeground() -eq $target) { $switched++ }
}
Write-Note "$switched of 12 switches landed on the intended window"
Start-Sleep -Seconds 2
Test-Intact 'After rapid switching' | Out-Null

# ---- closing from the middle ---------------------------------------------------------------------

Write-Step 'Closing documents from the middle of the stack'
$frames = @(Get-Frames)
if ($frames.Count -ge 3) {
    [WordLayout]::Close($frames[[int]($frames.Count / 2)])
    Start-Sleep -Seconds 5
    Test-Intact 'After closing a middle document' | Out-Null

    [WordLayout]::Close((Get-Frames)[0])
    Start-Sleep -Seconds 5
    Test-Intact 'After closing another' | Out-Null
}

Write-Step 'Opening one more onto the reduced stack'
Open-Document 7
Start-Sleep -Seconds 8
Wait-Settled ((Get-Frames).Count) 20
Test-Intact 'After opening onto a reduced stack' | Out-Null

# ---- done -----------------------------------------------------------------------------------------

if (-not $KeepOpen) {
    # Every window, one at a time, with WM_CLOSE - and never Kill.
    #
    # This used to be CloseMainWindow followed by Kill, and both halves were wrong. Word's "main
    # window" is one of seven frames here, so closing it left the other six standing; and the Kill
    # that followed is the thing this project's notes have forbidden since the first slice, because a
    # killed Word offers to recover those documents the next time it starts.
    #
    # That is not theoretical. It cost a run of check-look in this slice: soak-stack finished green,
    # killed Word, and the next suite opened three documents into a Word that had helpfully brought
    # back four of soak's - so a suite that asserts "3 Word windows" measured seven. The same trap is
    # written down twice in the project memory and it had a live source in this file the whole time.
    #
    # `$startedWord` used to guard this and no longer does: the suite opened seven documents, and
    # leaving them behind is exactly what contaminates whatever runs next. Every other suite already
    # closes Word unconditionally before it starts, so this is the same rule at the other end.
    #
    # And it now says WHY when it cannot: this loop had no idea what a dialog was, so a save prompt
    # and a wedged Word produced the same message.
    Write-Step 'Closing Word'
    $end = Close-AllWord
    if (-not $end.Closed) {
        Write-Note ("Word is still running ({0}): {1}" -f $end.Reason, (Format-WordWindow $end.Dialog))
        Write-Note 'Close it by hand before the next suite.'
    }
}

Write-Host ''
if ($script:Failures -eq 0) {
    Write-Host "PASS  $($script:Checks) checks, 0 failures" -ForegroundColor Green
} else {
    Write-Host "FAIL  $($script:Failures) of $($script:Checks) checks failed" -ForegroundColor Red
}
exit $script:Failures
