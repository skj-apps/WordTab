<#
.SYNOPSIS
  ONE PAGE ACROSS: that a document which opens showing several pages side by side is put back to
  one, that a document already on one page is not touched, and that the switch turns it off.

.DESCRIPTION
  From the work rig, in their words: "i have to chage view to single page on word open" - "every
  time i mean". Three answers had already been given to "it opens 3 pages wide", all of them about
  the WIDTH of the window, and this is not that.

  What Word actually does, measured on 16.0.20228 and reproduced by this suite:

    - A healthy Word reports `Zoom.PageColumns = 99` - its "as many as fit" - and how many pages sit
      side by side is then purely a function of the window's width.
    - Once anything sets a COLUMN COUNT, Word keeps it as a DEFAULT and crushes the zoom to fit that
      many pages. `PageColumns = 2` in a 1305px window took the zoom to 10%. It survives closing the
      document, closing Word, and opening a document Word has never seen - including a brand new
      blank one from the template.
    - The ribbon's One Page and 100% buttons fix the window in front of you and do NOT clear that
      default. That is the whole of "every time".

  So this suite BREAKS Word the way the work rig is broken - through the object model, which is the
  same door the ribbon uses - and then asserts that opening a document puts it right. The break is
  asserted first: a suite that could not break Word would pass this without the add-in doing
  anything at all.

  The instrument is Word's own object model for the state and the add-in's log for the decision. Both
  are needed: the log alone cannot prove the view changed, and the view alone cannot tell the add-in
  doing it from Word happening to open that way.

.PARAMETER KeepOpen
  Leave Word running afterwards.

.EXAMPLE
  pwsh -File tools\check-onepage.ps1
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

$SettingsKey = 'HKCU:\Software\WordTab'

# Word is driven through its own object model here rather than by launching winword.exe, because the
# view has to be read back as well as broken, and a Word this script created is a Word this script
# can still talk to. The add-in loads into it exactly as it would otherwise - these are ordinary
# OpusApp frames and the janitor joins them without knowing who opened them.
function New-Doc($name) {
    $path = Join-Path $scratch "wordtab-onepage-$name.rtf"
    "{\rtf1\ansi WordTab one-page check - $name.\par}" | Set-Content -Path $path -Encoding Ascii
    return $path
}

function Get-View($word) {
    $z = $word.ActiveWindow.View.Zoom
    return [pscustomobject]@{ Columns = [int]$z.PageColumns; Percent = [int]$z.Percentage }
}

function Format-View($v) { return "$($v.Columns) page(s) across at $($v.Percent)%" }

# A Word, a document, and the add-in given time to join it. The wait is fixed rather than a poll on
# the thing being measured: how soon the view is corrected is part of the claim.
function Open-In-Word($path) {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    $doc = $word.Documents.Open($path)
    Start-Sleep -Seconds 3
    return @{ Word = $word; Doc = $doc }
}

function Close-Word($session) {
    if ($session.Doc)  { try { $session.Doc.Close(0) } catch { } }
    if ($session.Word) { try { $session.Word.Quit(0) } catch { } }
    Start-Sleep -Seconds 3
    Close-AllWord | Out-Null
    Start-Sleep -Seconds 1
}

# ---- what was there before, so the machine is left as it was found -----------------------------

$origOnePage = try { (Get-ItemProperty -Path $SettingsKey -Name 'OnePage' -ErrorAction Stop).OnePage } catch { $null }

try {

# ---- 1. break Word the way the work rig is broken ----------------------------------------------

Write-Step 'Breaking Word: three pages side by side, which Word then remembers'
Close-AllWord | Out-Null
Start-Sleep -Seconds 2

# With the add-in's correction OFF, or it would undo the break as fast as it is made.
Set-ItemProperty -Path $SettingsKey -Name 'OnePage' -Value 0 -Type DWord

$break = Open-In-Word (New-Doc 'break')
$break.Word.ActiveWindow.View.Zoom.PageColumns = 3
Start-Sleep -Milliseconds 800
$broken = Get-View $break.Word
Write-Note "set to $(Format-View $broken)"
Close-Word $break

# The precondition, and it is asserted rather than assumed: Word has to still be broken in a NEW
# process, for a document it has never opened, or nothing below can fail.
Set-LogMark
$fresh = Open-In-Word (New-Doc 'fresh-with-switch-off')
$stuck = Get-View $fresh.Word
Write-Note "a new Word, a new document, switch off: $(Format-View $stuck)"
Assert ($stuck.Columns -gt 1) "Word carried the column count into a document it had never seen ($($stuck.Columns))"
Assert ((Get-LogCount 'view  hwnd=') -eq 0) 'and with OnePage=0 the add-in said nothing and changed nothing'
Assert ((Get-View $fresh.Word).Columns -eq $stuck.Columns) 'the view is still what Word opened it as'
Close-Word $fresh

# ---- 2. with the switch on, the add-in puts it right --------------------------------------------

Write-Step 'With OnePage on, a document that opens several pages across'
Remove-ItemProperty -Path $SettingsKey -Name 'OnePage' -ErrorAction SilentlyContinue   # default is on
Set-LogMark
$fixed = Open-In-Word (New-Doc 'fixed')
$after = Get-View $fixed.Word
Write-Note "after the add-in: $(Format-View $after)"
$said = Get-LogLast 'view  hwnd='
if ($said) { Write-Note "  log: $said" }

Assert ($null -ne $said) 'the add-in said what it found and what it did'
Assert ($after.Columns -eq 1) "the document is one page across (got $($after.Columns))"
Assert ($after.Percent -eq 100) "and at 100% rather than the crushed zoom (got $($after.Percent)%)"
Assert ($null -ne $said -and $said -match 'pages across') 'and the line names how many pages it found, not just that it acted'

# ---- 3. a window already on one page is not touched ---------------------------------------------
#
# The claim is "acts only when there is something wrong", and the way to fail it is to act always.
# A second document opened into the same, now-healthy Word must produce no line at all.

Write-Step 'A document that opens on one page already'
Set-LogMark
$second = $fixed.Word.Documents.Open((New-Doc 'already-fine'))
Start-Sleep -Seconds 3
$fine = Get-View $fixed.Word
Write-Note "second document: $(Format-View $fine)"
$saidAgain = Get-LogLast 'view  hwnd='
if ($saidAgain) { Write-Note "  log: $saidAgain" }

# This is the check that caught the half-fix. Correcting the column count does NOT make Word
# recompute the zoom it crushed to fit them, so the second document came up one page across at 10%
# - one page, and unreadable. Asserting the columns alone passed it.
Assert ($fine.Columns -eq 1) 'the second document is one page across'
Assert ($fine.Percent -ge 50) "and readable rather than the leftover fit-many-pages zoom (got $($fine.Percent)%)"
try { $second.Close(0) } catch { }
Close-Word $fixed

# ---- 4. and the correction sticks, because Word remembers what the add-in set --------------------

Write-Step 'The next Word, which is the one the user actually complains about'
Set-LogMark
$later = Open-In-Word (New-Doc 'later')
$now = Get-View $later.Word
Write-Note "a new Word afterwards: $(Format-View $now)"
$saidLater = Get-LogLast 'view  hwnd='
if ($saidLater) { Write-Note "  log: $saidLater" }
Assert ($now.Columns -le 1) 'Word opens on one page by itself now - the column count it kept is the corrected one'
Assert ($now.Percent -ge 50) "and the document is readable (got $($now.Percent)%)"
Close-Word $later

} finally {
    if ($null -eq $origOnePage) {
        Remove-ItemProperty -Path $SettingsKey -Name 'OnePage' -ErrorAction SilentlyContinue
    } else {
        Set-ItemProperty -Path $SettingsKey -Name 'OnePage' -Value $origOnePage -Type DWord
    }
    $back = try { (Get-ItemProperty -Path $SettingsKey -Name 'OnePage' -ErrorAction Stop).OnePage } catch { '(absent)' }
    Write-Note "OnePage restored to $back"

    if (-not $KeepOpen) {
        $end = Close-AllWord
        if (-not $end.Closed) {
            Write-Note ("Word is still up ({0}): {1}" -f $end.Reason, (Format-WordWindow $end.Dialog))
            Write-Note 'Answer it by hand before running the next suite - a leftover Word poisons whatever runs next.'
        }
    }
}

Write-Host ''
if ($script:Failures -eq 0) {
    Write-Host "PASS  $($script:Checks) checks, 0 failures" -ForegroundColor Green
} else {
    Write-Host "FAIL  $($script:Failures) of $($script:Checks) checks failed" -ForegroundColor Red
}
Write-Host "Add-in log: $env:LOCALAPPDATA\WordTab\wordtab.log (lines starting 'view')" -ForegroundColor Gray
exit $script:Failures
