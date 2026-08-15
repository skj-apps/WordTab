<#
.SYNOPSIS
  Build, install and register WordTab as a per-user Word COM add-in. No admin rights required.

.DESCRIPTION
  Copies the exact deployment shape measured on the work rig's own Office Tab install:
  binaries under %LOCALAPPDATA%\Programs\<Product>\, registration entirely under HKCU.
  Nothing is written to HKLM, Program Files, or the machine certificate store, so this runs
  under a plain user token on a locked-down corporate machine.

  This is the real installer in embryo, not a dev convenience script. Keep it admin-free.

.PARAMETER SkipBuild
  Register whatever is already in the build output instead of rebuilding.

.PARAMETER NoBanner
  Install with the "WordTab is loaded" dialog switched off (it writes the log either way).
  Flips HKCU\Software\WordTab\ShowLoadBanner, so it takes effect without a rebuild.

.PARAMETER SkipSmokeTest
  Skip the post-install CoCreate check.

.EXAMPLE
  pwsh -File install\install.ps1
  pwsh -File install\install.ps1 -NoBanner
#>
[CmdletBinding()]
param(
    [switch]$SkipBuild,
    [switch]$NoBanner,
    [switch]$SkipSmokeTest,
    [ValidateSet('Debug','Release')] [string]$Configuration = 'Release'
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# ---- identity. Must match src\WordTab.Connect\Connect.cs; verified below, not trusted. --------
$Clsid       = '4BF75ED9-10EE-4866-BF4A-3D663A4149A1'
$ProgId      = 'WordTab.Connect'
$AssemblyNm  = 'WordTab.Connect'
$ClassName   = 'WordTab.Connect'          # namespace + type, what mscoree activates
$FriendlyNm  = 'WordTab'
$Description = 'Tabbed document interface for Word'

$RepoRoot   = Split-Path -Parent $PSScriptRoot
$ProjectDir = Join-Path $RepoRoot 'src\WordTab.Connect'
$SourceFile = Join-Path $ProjectDir 'Connect.cs'
$InstallDir = Join-Path $env:LOCALAPPDATA 'Programs\WordTab'

$ClsidKey    = "HKCU:\Software\Classes\CLSID\{$Clsid}"
$ProgIdKey   = "HKCU:\Software\Classes\$ProgId"
$AddinKey    = "HKCU:\Software\Microsoft\Office\Word\Addins\$ProgId"
$SettingsKey = 'HKCU:\Software\WordTab'
$ManagedCategory = '{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}'   # what regasm stamps on .NET classes

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

# The CLSID lives in two places by necessity: the C# attribute and this script. Neither can be
# generated from the other without a build step, so instead they are compared every run.
if (-not (Test-Path $SourceFile)) { throw "Cannot find $SourceFile" }
$declared = Select-String -Path $SourceFile -Pattern 'ClsidString\s*=\s*"([0-9A-Fa-f-]{36})"' |
            ForEach-Object { $_.Matches[0].Groups[1].Value } | Select-Object -First 1
if (-not $declared) { throw "Could not find ClsidString in $SourceFile" }
if ($declared -ne $Clsid) {
    throw "CLSID mismatch. Connect.cs says {$declared}, this script says {$Clsid}. Fix one before installing."
}
Write-Ok "CLSID matches Connect.cs: {$Clsid}"

# ---- build ------------------------------------------------------------------------------------

$dllOut = Join-Path $ProjectDir "bin\$Configuration\net481\$AssemblyNm.dll"

if ($SkipBuild) {
    Write-Step 'Skipping build (-SkipBuild)'
} else {
    Write-Step "Building ($Configuration, net481)"
    & dotnet build (Join-Path $ProjectDir "$AssemblyNm.csproj") -c $Configuration -v quiet --nologo
    if ($LASTEXITCODE -ne 0) { throw "Build failed (exit $LASTEXITCODE)." }
    Write-Ok 'Build succeeded'
}

if (-not (Test-Path $dllOut)) { throw "Build output not found: $dllOut" }

# ---- install files ------------------------------------------------------------------------------

Write-Step "Installing to $InstallDir"
New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
Copy-Item -Path (Join-Path (Split-Path $dllOut) '*') -Destination $InstallDir -Force -Recurse
$installedDll = Join-Path $InstallDir "$AssemblyNm.dll"
if (-not (Test-Path $installedDll)) { throw "Copy failed: $installedDll missing." }
Write-Ok ("{0}  ({1:N0} bytes)" -f $installedDll, (Get-Item $installedDll).Length)

# The registry records the assembly's *full* name. Read it off the built DLL so a version bump in
# the csproj can never leave a stale registration that silently fails to activate.
try {
    $fullName = [System.Reflection.AssemblyName]::GetAssemblyName($installedDll).FullName
} catch {
    $fullName = "$AssemblyNm, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
    Write-Note "Could not read assembly name from the DLL; using the expected default."
}
Write-Ok "Assembly: $fullName"

$codeBase = 'file:///' + ($installedDll -replace '\\', '/')
$mscoree  = Join-Path $env:SystemRoot 'System32\mscoree.dll'

# ---- register ------------------------------------------------------------------------------------

function Set-Default($path, $value) {
    New-Item -Path $path -Force | Out-Null
    New-ItemProperty -Path $path -Name '(default)' -Value $value -PropertyType String -Force | Out-Null
}
function Set-Str($path, $name, $value) {
    # -Force on New-ItemProperty creates the *value*, not the key, so make the key first.
    if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
    New-ItemProperty -Path $path -Name $name -Value $value -PropertyType String -Force | Out-Null
}

Write-Step 'Registering the COM class under HKCU'

Set-Default $ClsidKey $FriendlyNm

# InprocServer32 points at the .NET shim, which then reads Class/Assembly/RuntimeVersion/CodeBase
# to find and activate the managed type. CodeBase is what lets an unsigned assembly outside the
# GAC be found at all - it is not optional here.
$inproc = "$ClsidKey\InprocServer32"
Set-Default $inproc $mscoree
Set-Str $inproc 'ThreadingModel' 'Both'
Set-Str $inproc 'Class'          $ClassName
Set-Str $inproc 'Assembly'       $fullName
Set-Str $inproc 'RuntimeVersion' 'v4.0.30319'
Set-Str $inproc 'CodeBase'       $codeBase

# regasm also writes a version-stamped subkey; mirror it so our registration is shaped exactly
# like one produced by the supported tool.
$version = ($fullName -split 'Version=')[1].Split(',')[0]
$versioned = "$inproc\$version"
Set-Str $versioned 'Class'          $ClassName
Set-Str $versioned 'Assembly'       $fullName
Set-Str $versioned 'RuntimeVersion' 'v4.0.30319'
Set-Str $versioned 'CodeBase'       $codeBase

Set-Default "$ClsidKey\ProgId" $ProgId
New-Item -Path "$ClsidKey\Implemented Categories\$ManagedCategory" -Force | Out-Null

Set-Default $ProgIdKey $FriendlyNm
Set-Default "$ProgIdKey\CLSID" "{$Clsid}"
Write-Ok "HKCU\Software\Classes\CLSID\{$Clsid}  ->  $ClassName"

Write-Step 'Registering the add-in with Word'
New-Item -Path $AddinKey -Force | Out-Null
Set-Str $AddinKey 'FriendlyName' $FriendlyNm
Set-Str $AddinKey 'Description'  $Description
# LoadBehavior 3 = load at startup, every time. Word rewrites this to 2 if we throw during load.
New-ItemProperty -Path $AddinKey -Name 'LoadBehavior'    -Value 3 -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path $AddinKey -Name 'CommandLineSafe' -Value 0 -PropertyType DWord -Force | Out-Null
Write-Ok "$AddinKey  LoadBehavior=3"

New-Item -Path $SettingsKey -Force | Out-Null
New-ItemProperty -Path $SettingsKey -Name 'ShowLoadBanner' -Value ([int](-not $NoBanner)) -PropertyType DWord -Force | Out-Null
Write-Ok ("Load banner: {0}" -f $(if ($NoBanner) { 'off' } else { 'on' }))

# ---- verify ------------------------------------------------------------------------------------

Write-Step 'Verifying'

# Word keeps a list of add-ins it killed after a crash or a throw during load. An entry here beats
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
    # a second instead of a Word launch. Windows PowerShell is used on purpose: it is a native
    # .NET Framework host, the same runtime the add-in targets.
    Write-Step 'Smoke test: activating the class outside Word'
    $script = "try { `$o = New-Object -ComObject '$ProgId'; if (`$o) { 'OK ' + `$o.GetType().FullName; [Runtime.InteropServices.Marshal]::ReleaseComObject(`$o) | Out-Null } } catch { 'FAIL ' + `$_.Exception.Message; exit 1 }"
    $result = & "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -Command $script 2>&1
    if ($LASTEXITCODE -eq 0 -and "$result" -like 'OK*') {
        Write-Ok "CoCreate('$ProgId') succeeded - $result"
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
