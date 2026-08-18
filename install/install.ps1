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

  It also installs a PACKAGE. install\package.ps1 produces a folder holding this script, its
  uninstaller, the built DLL and a manifest, and this script recognises one by its PAYLOAD.txt and
  verifies the DLL against the hash recorded there instead of building. That is how WordTab reaches a
  machine with no C++ toolchain, which the target corporate rig is.

.PARAMETER SkipBuild
  Register whatever is already in src\native\build instead of rebuilding. Not needed for a package -
  a package has nothing to build and this script works that out for itself.

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

# Which build this is, for the Windows uninstall entry's version column. WordTab has no version
# number scheme and its DLL carries no version resource, so the commit is the version - the same
# identifier the zip name, PAYLOAD.txt and settings.ps1 -Report already use. Filled in below from
# whichever of the two is available; deliberately not a fallback chain, since a package has no git
# and a repo has no manifest.
$BuildId = 'unknown'

$RepoRoot   = Split-Path -Parent $PSScriptRoot
$NativeDir  = Join-Path $RepoRoot 'src\native'
$HeaderFile = Join-Path $NativeDir 'wordtab.h'
$BuiltDll   = Join-Path $NativeDir "build\$DllName"
$InstallDir = Join-Path $env:LOCALAPPDATA 'Programs\WordTab'

$ClsidKey    = "HKCU:\Software\Classes\CLSID\{$Clsid}"
$ProgIdKey   = "HKCU:\Software\Classes\$ProgId"
$AddinKey    = "HKCU:\Software\Microsoft\Office\Word\Addins\$ProgId"
$SettingsKey = 'HKCU:\Software\WordTab'

# Windows' own list of installed programs. Per-user under HKCU, which is the same no-admin rule
# everything else here follows, and Settings > Apps enumerates it alongside the machine-wide hives.
# It is where a corporate user - or whoever they ask - looks first to remove something, and until now
# WordTab was not anywhere Windows itself knew to enumerate.
$ArpKey      = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\WordTab'

function Write-Step($text) { Write-Host "==> $text" -ForegroundColor Cyan }
function Write-Ok  ($text) { Write-Host "    $text" -ForegroundColor Green }
function Write-Note($text) { Write-Host "    $text" -ForegroundColor DarkGray }

# The registry questions this script and settings.ps1 both ask. A package carries this file beside
# the two scripts that need it; a repo has it in the same place. Named explicitly rather than
# guarded with a Test-Path fallback, because a fallback would be a second implementation of the very
# thing this file exists to have one of - and install.ps1 already depends on a sibling file
# (src\native\wordtab.h, for the CLSID check below).
$CommonFile = Join-Path $PSScriptRoot 'common.ps1'
if (-not (Test-Path $CommonFile)) {
    throw "install\common.ps1 is missing. It sits beside this script in both the repo and a package; copy the whole folder rather than install.ps1 on its own."
}
. $CommonFile

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

# A packaged payload - see install\package.ps1 - is this same tree with the compiler-shaped hole
# filled in: the built DLL, the one header the CLSID check reads, and a manifest naming exactly which
# build it is. It exists because the target machine for this project has no C++ toolchain and
# downloading one onto a locked-down corporate box is a worse ask than carrying 200KB.
$PayloadFile = Join-Path $RepoRoot 'PAYLOAD.txt'
$IsPayload   = Test-Path $PayloadFile

if ($IsPayload) {
    Write-Step 'Installing from a package - there are no sources here and nothing to build'
} elseif ($SkipBuild) {
    Write-Step 'Skipping build (-SkipBuild)'
} else {
    Write-Step 'Building the native server'
    & (Join-Path $NativeDir 'build.ps1')
    if ($LASTEXITCODE -ne 0) { throw "Build failed (exit $LASTEXITCODE)." }
}

if (-not (Test-Path $BuiltDll)) { throw "Build output not found: $BuiltDll" }

# A stale DLL is the most expensive kind of wrong: everything installs cleanly, Word loads it
# happily, and the change under test simply is not in it. Two ways to catch that, and the package
# gets to use the better one.
if ($IsPayload) {
    # The hash the packager recorded, against the bytes that arrived. This REPLACES the timestamp
    # comparison rather than adding to it, and that is the point: a file's LastWriteTime survives a
    # zip, an email attachment and a USB stick unpredictably - a git clone rewrites every one of them
    # to the moment of the clone - while its bytes either arrived intact or did not. The proxy is
    # only ever standing in for this check; where this check is available, the proxy is noise that
    # can fail on a correct payload.
    $manifest = Get-Content -Path $PayloadFile
    $wantHash = ($manifest | Where-Object { $_ -like 'Sha256:*' } | Select-Object -First 1)
    $builtAt  = ($manifest | Where-Object { $_ -like 'Commit:*' } | Select-Object -First 1)
    if (-not $wantHash) { throw "PAYLOAD.txt carries no Sha256: line, so the DLL cannot be verified. Re-package it." }
    $wantHash = $wantHash.Substring(7).Trim()
    $gotHash  = (Get-FileHash -Path $BuiltDll -Algorithm SHA256).Hash
    if ($gotHash -ne $wantHash) {
        throw "$DllName does not match the package manifest. Expected SHA256 $wantHash, got $gotHash. The file was altered or truncated in transit - re-copy the package."
    }
    Write-Ok "$DllName matches the manifest (SHA256 $($gotHash.Substring(0,16))...)"
    if ($builtAt) { Write-Note $builtAt.Trim(); $BuildId = $builtAt.Substring(7).Trim() }
} else {
    $newestSource = Get-ChildItem -Path $NativeDir -Include '*.cpp', '*.h', '*.def' -File -Recurse |
                    Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($newestSource -and (Get-Item $BuiltDll).LastWriteTime -lt $newestSource.LastWriteTime) {
        throw "$DllName is older than $($newestSource.Name) - it does not contain the current sources. Build it (drop -SkipBuild, and check the build actually succeeded)."
    }
    # A repo install is a dev install, and git is the manifest here. Not being able to answer is a
    # cosmetic loss - the entry still installs and still uninstalls - so this never stops an install.
    try {
        $head = & git -C $RepoRoot rev-parse --short HEAD 2>$null
        if ($LASTEXITCODE -eq 0 -and $head) { $BuildId = "$head".Trim() }
    } catch { }
}

# ---- install files ------------------------------------------------------------------------------

Write-Step "Installing to $InstallDir"
New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
Copy-Item -Path $BuiltDll -Destination $InstallDir -Force
$installedDll = Join-Path $InstallDir $DllName
if (-not (Test-Path $installedDll)) { throw "Copy failed: $installedDll missing." }

# A package that travelled over the internet - OneDrive, an email attachment, a browser download -
# arrives carrying a mark-of-the-web alternate data stream, and the copy above carries it across.
# The same mark that puts a Word document into Protected View can make policy refuse to load a DLL,
# and the failure looks exactly like a bad registration. One call, and it costs nothing when the file
# never left the machine.
try { Unblock-File -Path $installedDll -ErrorAction Stop; Write-Note 'cleared the mark-of-the-web, if it had one' } catch { }

Write-Ok ("{0}  ({1:N0} bytes)" -f $installedDll, (Get-Item $installedDll).Length)

# The uninstaller travels WITH the install, not only with the package it arrived in. A one-file .cmd
# install unpacks the payload to a temp folder and deletes it again (install\make-onefile.ps1), so
# before this line the machine kept the DLL and nothing that could remove it. The Windows uninstall
# entry written further down has to name a file that is still there when someone clicks Uninstall
# months later, and this is that file.
$UninstallSrc = Join-Path $PSScriptRoot 'uninstall.ps1'
if (-not (Test-Path $UninstallSrc)) {
    throw "install\uninstall.ps1 is missing. It sits beside this script in both the repo and a package; copy the whole folder rather than install.ps1 on its own."
}
Copy-Item -Path $UninstallSrc -Destination $InstallDir -Force
$installedUninstaller = Join-Path $InstallDir 'uninstall.ps1'
# Same mark-of-the-web reasoning as the DLL above. A marked .ps1 is refused outright under the
# RemoteSigned policy a corporate machine is likely to be on, and the failure would only show up at
# uninstall time - long after anyone would connect it to how the file arrived.
try { Unblock-File -Path $installedUninstaller -ErrorAction Stop } catch { }
Write-Ok $installedUninstaller

# The diagnostic travels with the install for the same reason the uninstaller does, and for one more:
# it is the ONLY way to see a machine nobody here can reach. Asking for it after something has gone
# wrong is too late if it went out of existence with the temp folder the .cmd unpacked into.
#
# common.ps1 as well, because settings.ps1 dot-sources it and says so by name rather than guessing.
foreach ($tool in @('settings.ps1', 'common.ps1')) {
    $from = Join-Path $PSScriptRoot $tool
    if (-not (Test-Path $from)) {
        throw "install\$tool is missing. It sits beside this script in both the repo and a package; copy the whole folder rather than install.ps1 on its own."
    }
    Copy-Item -Path $from -Destination $InstallDir -Force
    try { Unblock-File -Path (Join-Path $InstallDir $tool) -ErrorAction Stop } catch { }
}

# And something to double-click. A .ps1 is no more runnable at diagnosis time than it was at install
# time - that is the whole reason install\make-onefile.ps1 exists - and the moment someone needs this
# is the moment they are least able to be walked through opening a shell in a hidden folder.
#
# It clears PSModulePath for the reason written out at length in make-onefile.ps1: inherited from a
# PowerShell 7 terminal it makes Windows PowerShell lose Get-FileHash, which this report needs, and
# the failure reads as a broken tool rather than a borrowed variable.
$ReportCmd = Join-Path $InstallDir 'WordTab Report.cmd'
@'
@echo off
setlocal
title WordTab report
echo.
echo   Collecting everything needed to work out what WordTab is doing on this PC.
echo   This only READS. Nothing is installed, changed or removed.
echo.
set "PSModulePath="
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0settings.ps1" -Report
echo.
echo   Send the file named above.
echo.
pause
'@ | Set-Content -Path $ReportCmd -Encoding ASCII
Write-Ok $ReportCmd

# And an off switch that is a thing you can double-click, for the same reason the report is one.
#
# Word already has an off switch - the tick box in File > Options > Add-ins > COM Add-ins - and these
# files write the value that dialog writes rather than a second one. Two reasons they exist anyway: a
# corporate policy can hide that dialog, and "press Start, type WordTab" is findable by somebody who
# does not know Word has an add-ins page at all. Neither file decides anything; both hand off to
# settings.ps1, which holds the one implementation.
$OffCmd = Join-Path $InstallDir 'WordTab Off.cmd'
@'
@echo off
setlocal
title WordTab off
echo.
echo   Turning WordTab OFF. Word will start with no tabs, exactly as it did before it was installed.
echo   WordTab stays installed and nothing else is changed - put it back with "WordTab On".
echo.
set "PSModulePath="
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0settings.ps1" -Off
pause
'@ | Set-Content -Path $OffCmd -Encoding ASCII
Write-Ok $OffCmd

$OnCmd = Join-Path $InstallDir 'WordTab On.cmd'
@'
@echo off
setlocal
title WordTab on
echo.
echo   Turning WordTab back ON.
echo.
set "PSModulePath="
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0settings.ps1" -On
pause
'@ | Set-Content -Path $OnCmd -Encoding ASCII
Write-Ok $OnCmd

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
# LoadBehavior=3 outright and is the classic "I registered it and nothing happens" cause.
#
# The readers live in install\common.ps1 and are shared with settings.ps1 -Report, which asks the
# same questions of the same keys. Two implementations of "has Word disabled this" would agree right
# up until one of them was fixed, and this project has a scar for exactly that.
$disabledPaths = @(Get-DisabledAddinPaths)
$ourDll        = Join-Path $InstallDir $DllName
$oursDisabled  = @($disabledPaths | Where-Object { $_ -and $_.ToLowerInvariant() -eq $ourDll.ToLowerInvariant() })

if ($oursDisabled.Count -gt 0) {
    # The one case in here that means this install will not work. Loud, and with the fix.
    Write-Warning 'WORD HAS DISABLED WORDTAB. It will not load until you clear it, however correct the registration is.'
    Write-Note '  Word > File > Options > Add-ins > Manage: Disabled Items > Go... > select WordTab > Enable'
} elseif ($disabledPaths.Count -gt 0) {
    # Not ours, so not a blocker - but it is worth naming, because an add-in Word has disabled is
    # also one that cannot conflict with us, and the rival check below depends on knowing that.
    Write-Ok "Word has not disabled WordTab ($($disabledPaths.Count) other entry/entries in Disabled Items)"
    foreach ($path in ($disabledPaths | Select-Object -Unique)) { Write-Note "  disabled: $path" }
} else {
    Write-Ok 'Word has no disabled items'
}

# Another tabbed-Word add-in drives the same `_WwF` document frame that WordTab carves the strip out
# of, so whether one is going to LOAD is worth reporting accurately. Two things decide that, and
# this used to get both wrong:
#
#   - **HKCU beats HKLM for the same ProgId.** The old code scanned both hives and reported a hit in
#     either, so the dev rig - HKLM 3, HKCU 2 - was told an add-in would load that Word had already
#     given up on.
#   - **Disabled Items beats LoadBehavior.** An add-in Word killed after a failed load stays dead at
#     LoadBehavior=3. The dev rig's Office Tab is in exactly that state.
#
# Both rules live in Get-AddinStatus in common.ps1 now, so the report reaches the same verdict.
#
# **And the claim that used to be here - that WordTab and Office Tab "have coexisted through every
# check suite" - was measured on 2026-08-17 and is FALSE.** Office Tab is registered on the dev rig
# with LoadBehavior=3, which is what that sentence was read off; it has never once loaded during a
# check. Word put it in Disabled Items on 2026-08-15 and no suite has run beside a live rival since.
# Re-enabling it (Disabled Items cleared, LoadBehavior 3 on all three ProgIds, its own trial
# unexpired and its own Word switch on, its launcher running) got its DLL into WINWORD and still drew
# no strip, so the two have never been in one Word together and the coexistence claim cannot be made
# from this rig at all. Say untested, because untested is what it is.
#
# Turning somebody else's software off on their machine is their call, so this prints the command
# rather than running it. One warning, not one per registration: the dev rig alone has three of these
# ProgIds in both hives, which came out as five warnings and buried the useful line.
$rivalStatus   = @(Get-RivalAddins)
$rivalsLoading = @($rivalStatus | Where-Object { $_.Registered } | ForEach-Object { $_.ProgId })
$rivalsBlocked = @($rivalStatus | Where-Object { $_.Blocked }    | ForEach-Object { $_.ProgId })
$rivalsPerUser = @($rivalStatus | Where-Object { $_.Registered -and $_.Hive -eq 'HKCU:' } | ForEach-Object { $_.ProgId })

if ($rivalsLoading.Count -gt 0) {
    # "registered to load", not "will load". The registry says what Word has been ASKED to do; the
    # dev rig has a ProgId at LoadBehavior=3, absent from Disabled Items, whose DLL does not turn up
    # in WINWORD's module list at all. Claiming the stronger thing would be an assertion about the
    # product made from a fact about a registry key.
    Write-Warning "Another tabbed-Word add-in is registered to load with Word: $(@($rivalsLoading | Select-Object -Unique) -join ', ')"
    Write-Note 'WordTab has never been tested beside a working one - they carve up the same document frame,'
    Write-Note 'and both of them stack Word windows. Turn the other one off before starting Word:'
    foreach ($r in @($rivalsPerUser | Select-Object -Unique)) {
        Write-Note "  Set-ItemProperty 'HKCU:\Software\Microsoft\Office\Word\Addins\$r' LoadBehavior 0"
    }
    if (@($rivalsLoading | Where-Object { $rivalsPerUser -notcontains $_ }).Count -gt 0) {
        Write-Note '  and for any registered machine-wide: Word > File > Options > Add-ins > Manage: COM Add-ins'
    }
    Write-Note 'Put it back the same way with LoadBehavior 3, or by re-ticking it in that dialog.'
} elseif ($rivalsBlocked.Count -gt 0) {
    Write-Ok "no other tabbed-Word add-in will load (Word has disabled: $(@($rivalsBlocked | Select-Object -Unique) -join ', '))"
} else {
    Write-Ok 'no other tabbed-Word add-in is set to load'
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

# ---- tell Windows it is installed ---------------------------------------------------------------

# Settings > Apps > Installed apps, and Control Panel's Programs and Features behind it, both
# enumerate this key under HKCU as well as the machine hives. Eight values, and "how do I get rid of
# this" stops being a question only this repo can answer.
#
# powershell.exe by absolute path rather than pwsh: Windows launches this string itself, with no
# shell to resolve a bare name, and PowerShell 7 is not on a stock corporate Windows. System32
# resolves to the 64-bit copy for the 64-bit Explorer that launches it - which is the view
# uninstall.ps1 refuses to run outside of, so a wrong one would say so rather than half-work.
Write-Step 'Listing in Settings > Apps'
$psExe  = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
$sizeKb = [int]([Math]::Ceiling(((Get-ChildItem -Path $InstallDir -File -Recurse | Measure-Object -Property Length -Sum).Sum) / 1KB))

if (-not (Test-Path $ArpKey)) { New-Item -Path $ArpKey -Force | Out-Null }

# -Pause on the visible one because Windows launches it from Explorer: the console it opens closes
# the instant the script ends, so "close Word first" would flash past unread and a user who clicked
# Uninstall would see a window blink and WordTab still installed. QuietUninstallString is the same
# command without it, which is what anything removing WordTab unattended is supposed to use.
$strings = [ordered]@{
    DisplayName          = $FriendlyNm
    DisplayVersion       = $BuildId
    Publisher            = $FriendlyNm
    Comments             = $Description
    InstallLocation      = $InstallDir
    InstallDate          = (Get-Date -Format 'yyyyMMdd')
    UninstallString      = '"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}" -Pause' -f $psExe, $installedUninstaller
    QuietUninstallString = '"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}"' -f $psExe, $installedUninstaller
}
foreach ($name in $strings.Keys) {
    New-ItemProperty -Path $ArpKey -Name $name -Value $strings[$name] -PropertyType String -Force | Out-Null
}
# There is nothing to modify and nothing to repair; without these Windows offers both and each one
# would do nothing. EstimatedSize is in KB and is what the size column reads.
New-ItemProperty -Path $ArpKey -Name 'NoModify'      -Value 1       -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path $ArpKey -Name 'NoRepair'      -Value 1       -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path $ArpKey -Name 'EstimatedSize' -Value $sizeKb -PropertyType DWord -Force | Out-Null
# No DisplayIcon. WordTab ships no icon resource - see src\native\NOTE-no-icons.md - and pointing the
# value at a file that has none makes Windows draw a blank where the generic app icon would be.
Write-Ok "'$FriendlyNm', version $BuildId, $sizeKb KB"

# One shortcut in the user's own Start menu, so the report is found the way everything else on
# Windows is found: press Start, type WordTab. The alternative was telling somebody to navigate to a
# folder under AppData, which is hidden by default - a diagnostic nobody can reach is not one.
#
# Per-user Programs, not All Users: the same no-admin rule as the rest of this script.
#
# Three of them now, and the off switch is the reason the list is worth being: somebody whose Word is
# behaving strangely should be able to find the way to stand this add-in down by typing its name into
# Start, without knowing that the answer lives in Word's own options or in a folder under AppData.
$Shortcuts = @(
    @{ Name = 'WordTab Report.lnk'; Target = $ReportCmd
       Desc = 'Write a WordTab diagnostic report to your Desktop. Reads only.' }
    @{ Name = 'WordTab Off.lnk';    Target = $OffCmd
       Desc = 'Stop WordTab loading into Word, leaving it installed.' }
    @{ Name = 'WordTab On.lnk';     Target = $OnCmd
       Desc = 'Let WordTab load into Word again.' }
)
$Programs = [Environment]::GetFolderPath('Programs')
foreach ($s in $Shortcuts) {
    $path = Join-Path $Programs $s.Name
    try {
        $shell = New-Object -ComObject WScript.Shell
        $link  = $shell.CreateShortcut($path)
        $link.TargetPath       = $s.Target
        $link.WorkingDirectory = $InstallDir
        $link.Description      = $s.Desc
        $link.Save()
        Write-Ok "Start menu: $($s.Name)"
    } catch {
        # Not fatal. Every one of these is a .cmd on disk that the installer has already printed the
        # path of; a missing shortcut costs discoverability, not the thing itself.
        Write-Note "could not add $($s.Name) ($($_.Exception.Message))"
    }
}
Write-Note 'Press Start and type WordTab to find them.'

Write-Host ''
Write-Host 'Installed.' -ForegroundColor Green
if ($NoBanner) {
    Write-Host '  Start Word and open two documents. Expect one window with a tab for each.' -ForegroundColor Gray
    Write-Host '  There is no load banner (-NoBanner); the log below is the proof it loaded.' -ForegroundColor Gray
} else {
    Write-Host "  Start Word. Expect a 'WordTab is loaded inside Word' dialog." -ForegroundColor Gray
}
Write-Host "  Log: $env:LOCALAPPDATA\WordTab\wordtab.log" -ForegroundColor Gray
Write-Host '  If anything looks wrong: press Start, type WordTab, open "WordTab Report".' -ForegroundColor Gray
Write-Host '  It writes a file to your Desktop. Send that file.' -ForegroundColor Gray
Write-Host "  Remove it from Settings > Apps > Installed apps, like any other program." -ForegroundColor Gray
