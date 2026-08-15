<#
.SYNOPSIS
  Build WordTab.dll, the native in-process COM server.

.DESCRIPTION
  Uses w64devkit, a portable MinGW-w64 toolchain that lives in %LOCALAPPDATA% and needs no admin
  rights, no installer and no registry entries. Deleting its folder uninstalls it.

  Native rather than .NET because a managed COM server cannot be activated from a per-user
  registration, and per-user is all we get without admin. See src\WordTab.Connect\RESULT.md.

  The result is a single self-contained DLL: statically linked, so it drags in no libstdc++ or
  libgcc runtime, and dependent only on DLLs that ship with Windows. Run with -Verify to have that
  checked rather than assumed.

.PARAMETER DebugBuild
  Build unoptimised with debug info instead of the default optimised build.
  Not named -Debug: that is one of PowerShell's own common parameters and collides.

.PARAMETER Verify
  After linking, print the export table and the DLL's import dependencies.

.EXAMPLE
  pwsh -File src\native\build.ps1
  pwsh -File src\native\build.ps1 -Verify
#>
[CmdletBinding()]
param(
    [switch]$DebugBuild,
    [switch]$Verify
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$ToolkitRoot = Join-Path $env:LOCALAPPDATA 'Programs\w64devkit\w64devkit'
$Gpp         = Join-Path $ToolkitRoot 'bin\g++.exe'
$ObjDump     = Join-Path $ToolkitRoot 'bin\objdump.exe'

$SourceDir = $PSScriptRoot
$OutDir    = Join-Path $SourceDir 'build'
$OutDll    = Join-Path $OutDir 'WordTab.dll'

function Write-Step($text) { Write-Host "==> $text" -ForegroundColor Cyan }
function Write-Ok  ($text) { Write-Host "    $text" -ForegroundColor Green }

if (-not (Test-Path $Gpp)) {
    throw @"
w64devkit not found at $ToolkitRoot

It is a portable toolchain - no admin rights and no installer. To set it up:
  1. download w64devkit-x64-<version>.7z.exe from https://github.com/skeeto/w64devkit/releases
  2. run it with:  <file> -o"`$env:LOCALAPPDATA\Programs\w64devkit" -y
"@
}

# g++ shells out to `as` and `ld` by bare name, so the toolchain's bin has to be on PATH.
# Scoped to this process: nothing about the machine's PATH is touched.
$env:PATH = (Join-Path $ToolkitRoot 'bin') + [IO.Path]::PathSeparator + $env:PATH

Write-Step 'Compiling'
Write-Ok  ((& $Gpp --version | Select-Object -First 1))

New-Item -ItemType Directory -Path $OutDir -Force | Out-Null

$sources = @(
    (Join-Path $SourceDir 'dllmain.cpp')
    (Join-Path $SourceDir 'connect.cpp')
    (Join-Path $SourceDir 'frames.cpp')
    (Join-Path $SourceDir 'log.cpp')
)

$flags = @(
    '-shared'                       # build a DLL
    '-municode'                     # wide-character entry points
    '-m64'
    '-Wall', '-Wextra'
    '-fno-exceptions', '-fno-rtti'  # nothing may unwind across a COM boundary anyway
    '-static'                       # no libstdc++/libgcc DLL alongside us
    '-Wl,--enable-stdcall-fixup'
)
$flags += if ($DebugBuild) { @('-O0', '-g') } else { @('-O2', '-s') }   # -s strips symbols

# Import libraries. ole32/oleaut32 for COM and BSTR, uuid for the standard IIDs, shell32 for
# the known-folder lookup the logger uses, user32 for the windows and hooks, comctl32 for
# SetWindowSubclass - which is used rather than swapping GWLP_WNDPROC by hand because it is the
# only way to leave a subclass chain safely when another add-in has subclassed the same window
# after us.
$libs = @('-lole32', '-loleaut32', '-luuid', '-lshell32', '-ladvapi32', '-luser32', '-lcomctl32')

$arguments = $flags + $sources + (Join-Path $SourceDir 'wordtab.def') + @('-o', $OutDll) + $libs

& $Gpp @arguments
if ($LASTEXITCODE -ne 0) { throw "Compile/link failed (exit $LASTEXITCODE)." }

if (-not (Test-Path $OutDll)) { throw "Linker reported success but $OutDll is missing." }
Write-Ok ("{0}  ({1:N0} bytes)" -f $OutDll, (Get-Item $OutDll).Length)

if ($Verify) {
    Write-Step 'Exports (COM needs exactly these two, undecorated)'
    # objdump prints the name table as:  [   0] +base[   1]  0000 DllCanUnloadNow
    $exports = & $ObjDump -p $OutDll | Select-String -Pattern '\+base\[\s*\d+\]\s+[0-9a-f]{4}\s+(\S+)'
    if (-not $exports) { throw 'No exports found in the DLL. COM cannot load a server without DllGetClassObject.' }
    foreach ($line in $exports) { Write-Ok ($line.Matches[0].Groups[1].Value) }

    Write-Step 'Imported DLLs (all of these must ship with Windows)'
    $imports = & $ObjDump -p $OutDll | Select-String -Pattern '^\s*DLL Name:\s*(.+)$'
    foreach ($line in $imports) { Write-Ok ($line.Matches[0].Groups[1].Value.Trim()) }
}

Write-Host ''
Write-Host 'Built.' -ForegroundColor Green
Write-Host "  Install with: pwsh -File install\install.ps1" -ForegroundColor Gray
