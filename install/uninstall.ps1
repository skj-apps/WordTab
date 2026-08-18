<#
.SYNOPSIS
  Remove WordTab: unregister the add-in and delete the installed files. No admin rights required.

.DESCRIPTION
  The exact inverse of install.ps1. Everything it touches is under HKCU and %LOCALAPPDATA%, so
  this leaves no machine-wide residue behind because there was never any to begin with.

  install.ps1 copies this file to %LOCALAPPDATA%\Programs\WordTab and points the Windows uninstall
  entry at that copy, so on an installed machine the usual way to reach this script is Settings >
  Apps > Installed apps > WordTab > Uninstall. It deletes the folder it is running from, which is
  fine: PowerShell has finished reading a -File script before its first line runs.

.PARAMETER KeepLog
  Leave %LOCALAPPDATA%\WordTab (the log and settings) in place. Useful when uninstalling to
  investigate a load failure - the log is the evidence.

.PARAMETER Pause
  Wait for Enter before closing. For the Settings > Apps route only, where Windows opens a console
  of its own that would otherwise vanish the instant this ends - taking the result with it, and
  "close Word first" along with it.

.EXAMPLE
  pwsh -File install\uninstall.ps1
  pwsh -File install\uninstall.ps1 -KeepLog
#>
[CmdletBinding()]
param(
    [switch]$KeepLog,
    [switch]$Pause
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$Clsid  = '4BF75ED9-10EE-4866-BF4A-3D663A4149A1'
$ProgId = 'WordTab.Connect'

$InstallDir = Join-Path $env:LOCALAPPDATA 'Programs\WordTab'
$DataDir    = Join-Path $env:LOCALAPPDATA 'WordTab'
# Windows' installed-programs list, matching install.ps1. Restated here rather than shared through
# install\common.ps1 for the same reason the CLSID and the two folders above are: install.ps1 copies
# THIS FILE ALONE next to the DLL, with no siblings, so anything it dot-sourced would be missing on
# the only machine that matters. Self-contained is the requirement, not an oversight.
$ArpKey     = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\WordTab'

function Write-Step($text) { Write-Host "==> $text" -ForegroundColor Cyan }
function Write-Ok  ($text) { Write-Host "    $text" -ForegroundColor Green }
function Write-Note($text) { Write-Host "    $text" -ForegroundColor DarkGray }

# Everything below runs inside a try so that the Settings > Apps route can report. Windows launches
# the uninstall string in a console of its own and closes it the moment the script ends, so an
# unhandled throw there is a window that blinks and a WordTab still installed - see -Pause.
$removed = $false
try {

if (-not [Environment]::Is64BitProcess) {
    throw 'Run this from 64-bit PowerShell, the same view install.ps1 wrote to.'
}

$word = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue)
if ($word.Count -gt 0) {
    throw "Close Word first ($($word.Count) WINWORD process(es) running) - it holds the installed DLL open."
}

Write-Step 'Unregistering'
$keys = @(
    "HKCU:\Software\Microsoft\Office\Word\Addins\$ProgId",
    "HKCU:\Software\Classes\CLSID\{$Clsid}",
    "HKCU:\Software\Classes\$ProgId"
)
foreach ($key in $keys) {
    if (Test-Path $key) {
        Remove-Item -Path $key -Recurse -Force
        Write-Ok "removed $key"
    } else {
        Write-Note "not present  $key"
    }
}

Write-Step 'Removing files'
if (Test-Path $InstallDir) {
    # This folder holds the copy of THIS script that Settings > Apps runs, and deleting it out from
    # under itself is fine: PowerShell reads a -File script in full and closes the handle before the
    # first line executes. Driven both ways rather than assumed.
    Remove-Item -Path $InstallDir -Recurse -Force
    Write-Ok "removed $InstallDir"
} else {
    Write-Note "not present  $InstallDir"
}

if ($KeepLog) {
    Write-Note "kept $DataDir (-KeepLog)"
} elseif (Test-Path $DataDir) {
    Remove-Item -Path $DataDir -Recurse -Force
    Write-Ok "removed $DataDir"
    # ShowLoadBanner lives here rather than with the files; drop it with the rest of the settings.
    if (Test-Path 'HKCU:\Software\WordTab') {
        Remove-Item -Path 'HKCU:\Software\WordTab' -Recurse -Force
        Write-Ok 'removed HKCU:\Software\WordTab'
    }
} else {
    Write-Note "not present  $DataDir"
}

# Last, deliberately. This entry is the only route to an uninstaller that most people will ever see,
# so it outlives every step that could fail: if the DLL could not be deleted because Word grabbed it
# between the check above and now, the listing is still there to try again from. Removing it first
# would leave files on the machine and nothing pointing at the way to remove them.
if (Test-Path $ArpKey) {
    Remove-Item -Path $ArpKey -Recurse -Force
    Write-Ok "removed $ArpKey"
} else {
    Write-Note "not present  $ArpKey"
}

Write-Host ''
Write-Host 'Uninstalled.' -ForegroundColor Green
$removed = $true

}
catch {
    Write-Host ''
    Write-Host 'WordTab was NOT removed.' -ForegroundColor Red
    Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($Pause) {
        Write-Host ''
        Write-Host 'Press Enter to close this window.' -ForegroundColor DarkGray
        [void](Read-Host)
    }
}

if (-not $removed) { exit 1 }
