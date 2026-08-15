<#
.SYNOPSIS
  Remove WordTab: unregister the add-in and delete the installed files. No admin rights required.

.DESCRIPTION
  The exact inverse of install.ps1. Everything it touches is under HKCU and %LOCALAPPDATA%, so
  this leaves no machine-wide residue behind because there was never any to begin with.

.PARAMETER KeepLog
  Leave %LOCALAPPDATA%\WordTab (the log and settings) in place. Useful when uninstalling to
  investigate a load failure - the log is the evidence.

.EXAMPLE
  pwsh -File install\uninstall.ps1
  pwsh -File install\uninstall.ps1 -KeepLog
#>
[CmdletBinding()]
param(
    [switch]$KeepLog
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$Clsid  = '4BF75ED9-10EE-4866-BF4A-3D663A4149A1'
$ProgId = 'WordTab.Connect'

$InstallDir = Join-Path $env:LOCALAPPDATA 'Programs\WordTab'
$DataDir    = Join-Path $env:LOCALAPPDATA 'WordTab'

function Write-Step($text) { Write-Host "==> $text" -ForegroundColor Cyan }
function Write-Ok  ($text) { Write-Host "    $text" -ForegroundColor Green }
function Write-Note($text) { Write-Host "    $text" -ForegroundColor DarkGray }

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

Write-Host ''
Write-Host 'Uninstalled.' -ForegroundColor Green
