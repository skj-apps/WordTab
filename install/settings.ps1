<#
.SYNOPSIS
  Show or change WordTab's settings, without regedit.

.DESCRIPTION
  Every part of WordTab can be switched off on its own. They live as DWORD values under
  HKCU\Software\WordTab - no admin rights, no files, nothing outside the current user - and until this
  script existed the only way to reach them was regedit, which is not a reasonable thing to ask of
  somebody whose tab row has gone wrong on a Monday morning.

  Absent means on. Every switch defaults to 1, so a value only ever needs to exist in order to turn
  something OFF, and -Reset removes them all rather than writing zeros.

  **This script decides nothing.** It writes registry values and shows you what the add-in reported
  the last time it started; the add-in is the authority on what those values did, and its startup
  lines are printed at the bottom for exactly that reason. If the table here and the log disagree,
  the log is right.

  **Word reads these once, when it starts.** Change one, close Word completely, start it again.

.PARAMETER Set
  One or more NAME=VALUE pairs, e.g. -Set TabDot=0 or -Set TabDrag=0,TabScroll=0

.PARAMETER Reset
  Remove every WordTab setting, putting all of them back to their defaults.

.PARAMETER Quiet
  Skip the explanations and print just the table.

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File install\settings.ps1
  powershell -ExecutionPolicy Bypass -File install\settings.ps1 -Set TabDrag=0
  powershell -ExecutionPolicy Bypass -File install\settings.ps1 -Reset
#>
[CmdletBinding()]
param(
    [string[]]$Set,
    [switch]$Reset,
    [switch]$Quiet
)

$ErrorActionPreference = 'Stop'
$Key = 'HKCU:\Software\WordTab'

# What each switch does, in the terms somebody turning it off would think in. The names and defaults
# are the add-in's, not this script's - see the note in the help above about which one is the
# authority. Order is roughly "biggest hammer first".
$Switches = @(
    @{ Name = 'Stack';          Does = 'Put several documents in one window at all. Off, WordTab does nothing visible.' }
    @{ Name = 'Strip';          Does = 'Draw the tab row. Off, no strip is created and Word looks untouched.' }
    @{ Name = 'Taskbar';        Does = 'Show one taskbar button for the stack instead of one per document.' }
    @{ Name = 'AltTab';         Does = 'Show one Alt+Tab entry for the stack instead of one per document.' }
    @{ Name = 'TabStyle';       Does = 'Draw tabs as rounded cards in Word''s colours. Off, plain rectangles.' }
    @{ Name = 'TabThemeSample'; Does = 'Take the colours off Word''s own ribbon. Off, use the Office theme setting instead. TRY THIS FIRST if the strip is the wrong colour.' }
    @{ Name = 'TabButtons';     Does = 'The x on each tab and the + at the end.' }
    @{ Name = 'TabMenu';        Does = 'The right-click menu on a tab.' }
    @{ Name = 'TabDrag';        Does = 'Dragging a tab - to reorder the row, to pull a document out into its own window, or to put one back.' }
    @{ Name = 'TabTearOff';     Does = 'Taking a document out of the stack into a window of its own. Off, neither the menu item nor the drag is offered.' }
    @{ Name = 'TabScroll';      Does = 'Scroll the row when there are more tabs than fit. Off, every tab is shown and they get narrower instead.' }
    @{ Name = 'TabTitleTrim';   Does = 'Trim "  -  Compatibility Mode" and the like off a tab''s name.' }
    @{ Name = 'TabDot';         Does = 'Show a dot instead of the x on a document with unsaved changes. Off, WordTab never asks Word about your documents.' }
    @{ Name = 'TabKeys';        Does = 'Ctrl+Tab and Ctrl+Shift+Tab step along the tab row. TURN THIS OFF if you need Ctrl+Tab to type a tab inside a table.' }
    @{ Name = 'TabTip';         Does = 'Resting the pointer on a tab shows the document''s full name and the folder it is in. Off, WordTab never asks Word where your documents live.' }
    @{ Name = 'ShowLoadBanner'; Does = 'A dialog at Word startup confirming WordTab loaded. Off by default from the installer.' }
)

function Get-Current($name) {
    if (-not (Test-Path $Key)) { return $null }
    $p = Get-ItemProperty -Path $Key -Name $name -ErrorAction SilentlyContinue
    if ($p -and ($p.PSObject.Properties.Name -contains $name)) { return $p.$name }
    return $null
}

# ---- -Reset -------------------------------------------------------------------------------------

if ($Reset) {
    if (Test-Path $Key) {
        foreach ($s in $Switches) {
            Remove-ItemProperty -Path $Key -Name $s.Name -ErrorAction SilentlyContinue
        }
        Write-Host 'Every WordTab setting removed - all of them are back to their defaults.' -ForegroundColor Green
    } else {
        Write-Host 'There were no WordTab settings to remove; everything is already at its default.' -ForegroundColor Green
    }
    Write-Host 'Close Word completely and start it again for this to take effect.' -ForegroundColor Yellow
    Write-Host ''
}

# ---- -Set ---------------------------------------------------------------------------------------

if ($Set) {
    if (-not (Test-Path $Key)) { New-Item -Path $Key -Force | Out-Null }

    foreach ($pair in $Set) {
        $bits = $pair -split '=', 2
        if ($bits.Count -ne 2) {
            Write-Host "Not a NAME=VALUE pair: `"$pair`"" -ForegroundColor Red
            continue
        }

        $name  = $bits[0].Trim()
        $value = $bits[1].Trim()

        # Matched against the known list rather than written blindly. A typo would otherwise become a
        # registry value that looks like a setting, is read by nothing, and is indistinguishable from
        # a switch that does not work.
        $known = @($Switches | Where-Object { $_.Name -ieq $name } | Select-Object -First 1)
        if (-not $known) {
            Write-Host "There is no WordTab setting called `"$name`". Run this script with no arguments to see the list." -ForegroundColor Red
            continue
        }

        $number = 0
        if (-not [int]::TryParse($value, [ref]$number)) {
            Write-Host "`"$value`" is not a number. These are 1 for on and 0 for off." -ForegroundColor Red
            continue
        }

        Set-ItemProperty -Path $Key -Name $known.Name -Value $number -Type DWord
        Write-Host ("{0} = {1}" -f $known.Name, $number) -ForegroundColor Green
    }
    Write-Host 'Close Word completely and start it again for this to take effect.' -ForegroundColor Yellow
    Write-Host ''
}

# ---- what things are now -------------------------------------------------------------------------

Write-Host "WordTab settings  ($Key)" -ForegroundColor Cyan
Write-Host ''
foreach ($s in $Switches) {
    $current = Get-Current $s.Name
    $shown   = if ($null -eq $current) { 'on (default)' } elseif ($current -eq 0) { 'OFF' } else { "on ($current)" }
    $colour  = if ($null -ne $current -and $current -eq 0) { 'Yellow' } else { 'Gray' }
    Write-Host ("  {0,-16} {1}" -f $s.Name, $shown) -ForegroundColor $colour
    if (-not $Quiet) { Write-Host ("  {0,-16}   {1}" -f '', $s.Does) -ForegroundColor DarkGray }
}

if (-not $Quiet) {
    Write-Host ''
    Write-Host 'To turn something off:  -Set TabDot=0        To put everything back:  -Reset' -ForegroundColor Gray
    Write-Host 'Word reads these once, at startup. Close it completely and start it again.' -ForegroundColor Yellow
}

# ---- and what the add-in itself last reported ------------------------------------------------------
#
# The table above is documentation and can go stale; these lines are the add-in saying what it
# actually did with these values the last time Word started. If the two disagree, believe these.

$log = Join-Path $env:LOCALAPPDATA 'WordTab\wordtab.log'
if (Test-Path $log) {
    $starts = @(Get-Content $log -ErrorAction SilentlyContinue |
                Where-Object { $_ -match 'StackStart|StripStart|FramesStart|TaskbarStart' })
    if ($starts.Count -gt 0) {
        Write-Host ''
        Write-Host 'What the add-in reported the last time Word started:' -ForegroundColor Cyan
        # The last four only. The log holds every start since it was last rolled, and there are four
        # of these lines per start - so printing all of them would show settings that have since been
        # changed, which is worse than showing none.
        foreach ($line in ($starts | Select-Object -Last 4)) {
            Write-Host ("  {0}" -f ($line -replace '^\d{4}-\d\d-\d\d ', '').Trim()) -ForegroundColor DarkGray
        }
    }
} else {
    Write-Host ''
    Write-Host "No add-in log yet ($log) - start Word once and run this again to see what it reports." -ForegroundColor DarkGray
}
