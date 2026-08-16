<#
.SYNOPSIS
  Build, install and register WordTab as a per-user Word COM add-in. No admin rights required.

.DESCRIPTION
  Copies the exact deployment shape measured on the work rig's own Office Tab install:
  binaries under %LOCALAPPDATA%\Programs\<Product>\, registration entirely under HKCU.
  Nothing is written to HKLM, Program Files, or the machine certificate store, so this runs
  under a plain user token on a locked-down corporate machine.

  The server is a native DLL, and that is not a preference. A managed (.NET) COM server cannot be
  activated from a per-user registration at all - measured, see src\WordTab.Connect\RESULT.md.
  Native servers register per-user fine, which is also how Office Tab itself does it.

  This is the real installer in embryo, not a dev convenience script. Keep it admin-free.

.PARAMETER SkipBuild
  Register whatever is already in src\native\build instead of rebuilding.

.PARAMETER NoBanner
  Install with the "WordTab is loaded" dialog switched off (it writes the log either way).
  Flips HKCU\Software\WordTab\ShowLoadBanner, so it takes effect without a rebuild.

.PARAMETER SkipSmokeTest
  Skip the post-install activation check.

.EXAMPLE
  pwsh -File install\install.ps1
  pwsh -File install\install.ps1 -NoBanner
#>
[CmdletBinding()]
param(
    [switch]$SkipBuild,
    [switch]$NoBanner,
    [switch]$SkipSmokeTest
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# ---- identity. Must match src\native\wordtab.h; verified below, not trusted. ------------------
$Clsid       = '4BF75ED9-10EE-4866-BF4A-3D663A4149A1'
$ProgId      = 'WordTab.Connect'
$DllName     = 'WordTab.dll'
$FriendlyNm  = 'WordTab'
$Description = 'Tabbed document interface for Word'

$RepoRoot   = Split-Path -Parent $PSScriptRoot
$NativeDir  = Join-Path $RepoRoot 'src\native'
$HeaderFile = Join-Path $NativeDir 'wordtab.h'
$BuiltDll   = Join-Path $NativeDir "build\$DllName"
$InstallDir = Join-Path $env:LOCALAPPDATA 'Programs\WordTab'

$ClsidKey    = "HKCU:\Software\Classes\CLSID\{$Clsid}"
$ProgIdKey   = "HKCU:\Software\Classes\$ProgId"
$AddinKey    = "HKCU:\Software\Microsoft\Office\Word\Addins\$ProgId"
$SettingsKey = 'HKCU:\Software\WordTab'

function Write-Step($text) { Write-Host "==> $text" -ForegroundColor Cyan }
function Write-Ok  ($text) { Write-Host "    $text" -ForegroundColor Green }
function Write-Note($text) { Write-Host "    $text" -ForegroundColor DarkGray }

# ---- preconditions ----------------------------------------------------------------------------

Write-Step 'Checking preconditions'

# A 32-bit PowerShell would write HKCU\Software\Classes\Wow6432Node\CLSID, which 64-bit Word
# never reads. The failure is silent, so refuse rather than produce a dead registration.
if (-not [Environment]::Is64BitProcess) {
    throw 'Run this from 64-bit PowerShell. A 32-bit host registers into the WOW6432Node view, which x64 Word does not read.'
}
Write-Ok '64-bit PowerShell (registers into the view x64 Word reads)'

$word = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue)
if ($word.Count -gt 0) {
    throw "Close Word first ($($word.Count) WINWORD process(es) running) - it holds the installed DLL open."
}
Write-Ok 'Word is not running'

# The CLSID lives in two places by necessity: the C++ header and this script. Neither can be
# generated from the other without a build step, so instead they are compared every run.
if (-not (Test-Path $HeaderFile)) { throw "Cannot find $HeaderFile" }
$declared = Select-String -Path $HeaderFile -Pattern 'WORDTAB_CLSID_STRING\s+L"\{([0-9A-Fa-f-]{36})\}"' |
            ForEach-Object { $_.Matches[0].Groups[1].Value } | Select-Object -First 1
if (-not $declared) { throw "Could not find WORDTAB_CLSID_STRING in $HeaderFile" }
if ($declared -ne $Clsid) {
    throw "CLSID mismatch. wordtab.h says {$declared}, this script says {$Clsid}. Fix one before installing."
}
Write-Ok "CLSID matches wordtab.h: {$Clsid}"

# ---- build ------------------------------------------------------------------------------------

if ($SkipBuild) {
    Write-Step 'Skipping build (-SkipBuild)'
} else {
    Write-Step 'Building the native server'
    & (Join-Path $NativeDir 'build.ps1')
    if ($LASTEXITCODE -ne 0) { throw "Build failed (exit $LASTEXITCODE)." }
}

if (-not (Test-Path $BuiltDll)) { throw "Build output not found: $BuiltDll" }

# A stale DLL is the most expensive kind of wrong: everything installs cleanly, Word loads it
# happily, and the change under test simply is not in it. Compare against the sources rather than
# trusting that a build just happened - -SkipBuild bypasses it, and a failed build leaves the
# previous DLL sitting there looking perfectly valid.
$newestSource = Get-ChildItem -Path $NativeDir -Include '*.cpp', '*.h', '*.def' -File -Recurse |
                Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($newestSource -and (Get-Item $BuiltDll).LastWriteTime -lt $newestSource.LastWriteTime) {
    throw "$DllName is older than $($newestSource.Name) - it does not contain the current sources. Build it (drop -SkipBuild, and check the build actually succeeded)."
}

# ---- install files ------------------------------------------------------------------------------

Write-Step "Installing to $InstallDir"
New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
Copy-Item -Path $BuiltDll -Destination $InstallDir -Force
$installedDll = Join-Path $InstallDir $DllName
if (-not (Test-Path $installedDll)) { throw "Copy failed: $installedDll missing." }
Write-Ok ("{0}  ({1:N0} bytes)" -f $installedDll, (Get-Item $installedDll).Length)

# ---- register ------------------------------------------------------------------------------------

function Set-Default($path, $value) {
    if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
    New-ItemProperty -Path $path -Name '(default)' -Value $value -PropertyType String -Force | Out-Null
}
function Set-Str($path, $name, $value) {
    # -Force on New-ItemProperty creates the *value*, not the key, so make the key first.
    if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
    New-ItemProperty -Path $path -Name $name -Value $value -PropertyType String -Force | Out-Null
}

Write-Step 'Registering the COM class under HKCU'

# Delete before writing. An earlier managed registration of this same CLSID would have left
# Class / Assembly / RuntimeVersion / CodeBase values behind, and inheriting half of a different
# server's registration is a confusing way to fail.
try { [Microsoft.Win32.Registry]::CurrentUser.DeleteSubKeyTree("Software\Classes\CLSID\{$Clsid}", $false) } catch { }

Set-Default $ClsidKey $FriendlyNm

# For a native server this is the whole story: the path to the DLL, and how it may be called.
# COM will LoadLibrary it and call its exported DllGetClassObject. No shim, no runtime, no
# registry lookup that skips the per-user hive.
$inproc = "$ClsidKey\InprocServer32"
Set-Default $inproc $installedDll
Set-Str $inproc 'ThreadingModel' 'Both'

Set-Default "$ClsidKey\ProgId" $ProgId
Set-Default $ProgIdKey $FriendlyNm
Set-Default "$ProgIdKey\CLSID" "{$Clsid}"
Write-Ok "HKCU\Software\Classes\CLSID\{$Clsid}  ->  $DllName"

Write-Step 'Registering the add-in with Word'
New-Item -Path $AddinKey -Force | Out-Null
Set-Str $AddinKey 'FriendlyName' $FriendlyNm
Set-Str $AddinKey 'Description'  $Description
# LoadBehavior 3 = load at startup, every time. Word rewrites this to 2 if loading fails, so its
# value after a Word launch is a useful thing to read back.
New-ItemProperty -Path $AddinKey -Name 'LoadBehavior'    -Value 3 -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path $AddinKey -Name 'CommandLineSafe' -Value 0 -PropertyType DWord -Force | Out-Null
Write-Ok "$AddinKey  LoadBehavior=3"

# **Not `New-Item -Force`.** On a key that already exists, -Force recreates it and every value under
# it is gone - the registry twin of `New-Item -ItemType File -Force` truncating a file. This key is
# the user's, not the installer's: TabScroll, TabStyle, TabThemeSample, TabDrag, TabTitleTrim and
# the rest all live here, and re-installing over an existing copy silently reset every one of them
# to its default. The add-in registration above is recreated on purpose - every value in it is
# rewritten on the next two lines - but nothing here rewrites a switch the user set.
if (-not (Test-Path $SettingsKey)) { New-Item -Path $SettingsKey | Out-Null }
New-ItemProperty -Path $SettingsKey -Name 'ShowLoadBanner' -Value ([int](-not $NoBanner)) -PropertyType DWord -Force | Out-Null
Write-Ok ("Load banner: {0}" -f $(if ($NoBanner) { 'off' } else { 'on' }))

# ---- verify ------------------------------------------------------------------------------------

Write-Step 'Verifying'

# Word keeps a list of add-ins it killed after a crash or a failed load. An entry here beats
# LoadBehavior=3 and is the classic "I registered it and nothing happens" cause, so surface it.
$disabled = @(Get-ChildItem 'HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency\DisabledItems' -ErrorAction SilentlyContinue |
              ForEach-Object { Get-ItemProperty $_.PSPath } | ForEach-Object { $_.PSObject.Properties } |
              Where-Object { $_.Name -notlike 'PS*' })
if ($disabled.Count -gt 0) {
    Write-Warning "Word has $($disabled.Count) item(s) in Resiliency\DisabledItems. If WordTab does not load, clear that key: File > Options > Add-ins > Manage: Disabled Items."
} else {
    Write-Ok 'Word has no disabled items'
}

if (-not $SkipSmokeTest) {
    # Activate the class outside Word. This separates "the COM registration is correct" from
    # "Word chose to load us", which are different failures with different fixes - and it runs in
    # a second instead of a Word launch. It is also the check that caught the managed-server
    # problem before a single Word launch was wasted on it.
    Write-Step 'Smoke test: activating the class outside Word'
    $script = "try { `$o = New-Object -ComObject '$ProgId'; if (`$o) { 'OK'; [Runtime.InteropServices.Marshal]::ReleaseComObject(`$o) | Out-Null } } catch { 'FAIL ' + `$_.Exception.Message; exit 1 }"
    $result = & "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -Command $script 2>&1
    if ($LASTEXITCODE -eq 0 -and "$result" -like 'OK*') {
        Write-Ok "CoCreate('$ProgId') succeeded"
    } else {
        Write-Warning "CoCreate('$ProgId') failed: $result"
        Write-Warning 'The registration is wrong. Word will not load it either. Fix this before launching Word.'
    }
}

Write-Host ''
Write-Host 'Installed.' -ForegroundColor Green
Write-Host "  Start Word. Expect a 'WordTab is loaded inside Word' dialog." -ForegroundColor Gray
Write-Host "  Log: $env:LOCALAPPDATA\WordTab\wordtab.log" -ForegroundColor Gray
Write-Host "  Remove with: pwsh -File install\uninstall.ps1" -ForegroundColor Gray
