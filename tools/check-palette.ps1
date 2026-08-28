<#
.SYNOPSIS
  Assert the strip palette's adoption gate - the logic that decides whether a colour sampled off
  Word's ribbon is believed.

.DESCRIPTION
  **The second suite here that does not start Word, and for the same reason check-governor.ps1 does
  not: starting Word would not help.**

  Every other suite drives the real add-in and asserts what it does. That works because the thing
  being tested happens on this machine. This one is the exception. The faults it exists for are both
  properties of the work rig's Word and neither can be produced here on demand:

  - its ribbon reads `RGB(9,9,9)` while the window is active and `RGB(10,10,10)` while it is not, so
    every switch between two documents moved the palette by one unit in each channel;
  - and twice in one working day the sampler returned `RGB(0,0,0)` with `RGB(10,10,10)` either side
    of it, which the old rule believed, so the whole tab row changed colour for two seconds.

  `check-look.ps1` drives real Word and photographs the strip, and on this rig Word's ribbon holds
  one colour for the whole run - so it cannot tell the old rule from the new one. **That is exactly
  how the old rule shipped**: it passed every suite on this machine and then rebuilt the palette over
  a hundred times in one day on theirs, each rebuild destroying and remaking two brushes and
  invalidating every strip in the row.

  What this does instead is take the decision out of `AdoptSampledChrome` and replay it against
  colours copied from the work rig's own log, running the old rule and the new one side by side over
  identical input.

  `tools\palette-test.cpp` holds the two step functions and the assertions. It deliberately includes
  nothing from `src\native`: if the copy ever drifts from `strip.cpp`, the copy is wrong, and the
  copy is what is being asserted. **Keep them in step by hand when the gate changes.**

  Being pure arithmetic, this cannot be flaky, and it runs in about a second - which is why
  check-all.ps1 puts it beside check-governor at the front.

.PARAMETER Screenshot
  Accepted and ignored. There is nothing on screen to photograph; the parameter exists so
  check-all.ps1 can pass it down uniformly.

.EXAMPLE
  pwsh -File tools\check-palette.ps1
#>
[CmdletBinding()]
param(
    [switch]$Screenshot
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Write-Step($text) { Write-Host "==> $text" -ForegroundColor Cyan }

$source = Join-Path $PSScriptRoot 'palette-test.cpp'
if (-not (Test-Path $source)) {
    Write-Host "FAIL  the palette test's source is missing: $source" -ForegroundColor Red
    Write-Host '0 checks, 1 failures'
    exit 1
}

# The same toolchain build.ps1 uses, and located the same way, so there is one answer on this machine
# to "where is the compiler". g++ shells out to `as` and `ld` by bare name, so its bin has to go on
# PATH rather than merely being called by full path.
$toolkit = Join-Path $env:LOCALAPPDATA 'Programs\w64devkit\w64devkit'
$gpp     = Join-Path $toolkit 'bin\g++.exe'
if (-not (Test-Path $gpp)) {
    Write-Host "FAIL  w64devkit not found at $toolkit - see src\native\build.ps1 for how to install it" -ForegroundColor Red
    Write-Host '0 checks, 1 failures'
    exit 1
}
$env:PATH = (Join-Path $toolkit 'bin') + ';' + $env:PATH

Write-Step 'Compiling the palette test'
$exe = Join-Path ([System.IO.Path]::GetTempPath()) 'wordtab-palette-test.exe'
& $gpp -std=c++17 -O1 -Wall -Wextra -Werror -o $exe $source 2>&1 | ForEach-Object { Write-Host "    $_" }
if ($LASTEXITCODE -ne 0) {
    Write-Host '    FAIL  the palette test did not compile' -ForegroundColor Red
    Write-Host '0 checks, 1 failures'
    exit 1
}
Write-Host "    $exe" -ForegroundColor DarkGray

Write-Step 'Replaying the work rig''s sampled ribbon colours through both rules'
$output = & $exe 2>&1
$code   = $LASTEXITCODE
$output | ForEach-Object { Write-Host $_ }

# The exe prints its own PASS/FAIL lines in the house format; this only has to total them up and say
# it in the shape check-all.ps1 reads. Counting the lines rather than trusting the exit code alone,
# because a suite that exits 0 having asserted nothing is the failure this project keeps meeting.
$lines  = @($output | ForEach-Object { "$_" })
$passes = @($lines | Where-Object { $_ -match '^\s*PASS\s' }).Count
$fails  = @($lines | Where-Object { $_ -match '^\s*FAIL\s' }).Count
$total  = $passes + $fails

Write-Host ''
if ($total -eq 0) {
    Write-Host 'FAIL  the palette test ran but asserted nothing' -ForegroundColor Red
    Write-Host '0 checks, 1 failures'
    exit 1
}

if ($fails -eq 0 -and $code -eq 0) {
    Write-Host ("PASS  {0} checks, 0 failures" -f $total) -ForegroundColor Green
    exit 0
}

Write-Host ("FAIL  {0} checks, {1} failures (exit {2})" -f $total, $fails, $code) -ForegroundColor Red
exit 1
