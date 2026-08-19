<#
.SYNOPSIS
  Assert the dot poll's cadence governor - the logic that decides how often WordTab asks Word whether
  a document has unsaved changes.

.DESCRIPTION
  **The only suite here that does not start Word, and the reason it exists is that starting Word
  would not help.**

  Every other suite drives the real add-in and asserts what it does. That works because the thing
  being tested happens on this machine. The governor is the exception: it exists entirely to handle a
  Word that answers `Document.Saved` slowly, and on this rig Word never does. Local documents answer
  in a few hundred microseconds; the user's work PC, whose documents live on SharePoint, was measured
  at a mean of 96,000-155,000us and a worst of 518,726us for the same call.

  So the branch that matters is unreachable from here, and `check-dot.ps1` can only ever exercise the
  cheap side of it. **That is exactly how the first version of the governor shipped broken** - it
  passed all twelve suites on this machine and then oscillated for hours on theirs, at about 5.5% of
  Word's UI thread, in half-second freezes.

  What this does instead is take the arithmetic out of `PollModified` and replay it against cost
  sequences copied from the work rig's own log, running the old rule and the new one side by side
  over identical input. It is a test of the *decision*, with real measurements as the input, rather
  than a test of the add-in - much like replaying a recorded temperature log at a thermostat instead
  of waiting for winter.

  `tools\governor-test.cpp` holds the three step functions and the assertions. It deliberately
  includes nothing from `src\native`: if the copy ever drifts from `strip.cpp`, the copy is wrong,
  and the copy is what is being asserted. **Keep them in step by hand when PollModified changes.**

  Being pure arithmetic, this is the one suite that cannot be flaky, and it runs in about a second -
  which is why `check-all.ps1` puts it first.

.PARAMETER Screenshot
  Accepted and ignored. There is nothing on screen to photograph; the parameter exists so
  check-all.ps1 can pass it down uniformly.

.EXAMPLE
  pwsh -File tools\check-governor.ps1
#>
[CmdletBinding()]
param(
    [switch]$Screenshot
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Write-Step($text) { Write-Host "==> $text" -ForegroundColor Cyan }

$source = Join-Path $PSScriptRoot 'governor-test.cpp'
if (-not (Test-Path $source)) {
    Write-Host "FAIL  the governor test's source is missing: $source" -ForegroundColor Red
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

Write-Step 'Compiling the governor test'
$exe = Join-Path ([System.IO.Path]::GetTempPath()) 'wordtab-governor-test.exe'
& $gpp -std=c++17 -O1 -Wall -Wextra -Werror -o $exe $source 2>&1 | ForEach-Object { Write-Host "    $_" }
if ($LASTEXITCODE -ne 0) {
    Write-Host '    FAIL  the governor test did not compile' -ForegroundColor Red
    Write-Host '0 checks, 1 failures'
    exit 1
}
Write-Host "    $exe" -ForegroundColor DarkGray

Write-Step 'Replaying the work rig''s measured pass costs through both governors'
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
    Write-Host 'FAIL  the governor test ran but asserted nothing' -ForegroundColor Red
    Write-Host '0 checks, 1 failures'
    exit 1
}

if ($fails -eq 0 -and $code -eq 0) {
    Write-Host ("PASS  {0} checks, 0 failures" -f $total) -ForegroundColor Green
    exit 0
}

Write-Host ("FAIL  {0} checks, {1} failures (exit {2})" -f $total, $fails, $code) -ForegroundColor Red
exit 1
