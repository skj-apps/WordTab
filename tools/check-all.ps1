<#
.SYNOPSIS
  Run every check suite in turn and print one table at the end.

.DESCRIPTION
  Each suite found bugs the others could not, so after any change to geometry or drawing code they
  all have to run. This is the runner: it starts them one at a time, never in parallel - they all
  drive the same Word and the same desktop - and it reports each suite's own check count and exit
  status rather than merging them, because "311 of 311" hides which suite went quiet.

  **Word must be closed before the first one starts.** A suite that begins with the previous run's
  windows still up measures something else, and has.

.PARAMETER Only
  Run just these suites, by name (e.g. -Only title,menu).

.PARAMETER Screenshot
  Pass -Screenshot down to each suite.
#>
[CmdletBinding()]
param(
    [string[]]$Only,
    [switch]$Screenshot
)

$ErrorActionPreference = 'Continue'

# Ordered cheapest-first so a broken build fails fast, with soak last: it opens the most documents
# and is the one whose leftovers used to poison whatever ran next.
$suites = @('stack', 'strip', 'tabs', 'menu', 'reorder', 'startscreen', 'look', 'scroll', 'title', 'soak-stack')
if ($Only) {
    # `pwsh -File script.ps1 -Only reorder,look` passes ONE literal string, not two: with -File every
    # argument arrives as text and PowerShell does no array parsing, so a [string[]] parameter gets a
    # single element "reorder,look" that matches no suite. Split here so both forms work.
    $Only = @($Only | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })

    # Written out rather than nested Where-Objects. The clever one-liner version silently matched
    # nothing and the runner cheerfully reported "All 0 suites green" - which is the worst answer a
    # test runner can give, because it is the same shape as success.
    $wanted = @()
    foreach ($name in $suites) {
        foreach ($want in $Only) {
            if ($name -like "*$want*") { $wanted += $name; break }
        }
    }
    $suites = $wanted
}

if ($suites.Count -eq 0) { Write-Host 'No suites matched.' -ForegroundColor Red; exit 1 }
Write-Host ("Running: {0}" -f ($suites -join ', ')) -ForegroundColor Cyan

$results = @()
foreach ($name in $suites) {
    $script = Join-Path $PSScriptRoot "check-$name.ps1"
    if ($name -eq 'soak-stack') { $script = Join-Path $PSScriptRoot 'soak-stack.ps1' }
    if (-not (Test-Path $script)) { Write-Host "no such suite: $script" -ForegroundColor Yellow; continue }

    Write-Host ''
    Write-Host ('=' * 90) -ForegroundColor DarkGray
    Write-Host "RUNNING  $name" -ForegroundColor Cyan
    Write-Host ('=' * 90) -ForegroundColor DarkGray

    $started = Get-Date
    $args = @('-File', $script)
    if ($Screenshot) { $args += '-Screenshot' }
    $output = & pwsh @args 2>&1
    $code = $LASTEXITCODE
    $output | ForEach-Object { Write-Host $_ }

    # Not anchored: the suites do not agree on how they announce the total. Some print
    # "34 checks, all passed." at column 0, others "PASS  40 checks, 0 failures". Anchoring this
    # matched neither on the first run and every suite reported "(no summary line)" while passing.
    $tail = @($output | ForEach-Object { "$_" } | Where-Object { $_ -match '\d+ checks' })
    $summary = if ($tail.Count -gt 0) { $tail[-1].Trim() } else { '(no summary line)' }
    $fails = @($output | ForEach-Object { "$_" } | Where-Object { $_ -match '^\s*FAIL ' })

    $results += [pscustomobject]@{
        Suite   = $name
        Exit    = $code
        Summary = $summary
        Fails   = $fails.Count
        Minutes = [Math]::Round(((Get-Date) - $started).TotalMinutes, 1)
    }
}

Write-Host ''
Write-Host ('=' * 90) -ForegroundColor DarkGray
Write-Host 'ALL SUITES' -ForegroundColor Cyan
Write-Host ('=' * 90) -ForegroundColor DarkGray
foreach ($r in $results) {
    $colour = if ($r.Exit -eq 0 -and $r.Fails -eq 0) { 'Green' } else { 'Red' }
    Write-Host ("  {0,-12} exit={1} fails={2} {3,5} min   {4}" -f
                $r.Suite, $r.Exit, $r.Fails, $r.Minutes, $r.Summary) -ForegroundColor $colour
}

$bad = @($results | Where-Object { $_.Exit -ne 0 -or $_.Fails -gt 0 })
Write-Host ''
if ($bad.Count -eq 0) {
    Write-Host "All $($results.Count) suites green." -ForegroundColor Green
} else {
    Write-Host "$($bad.Count) of $($results.Count) suites need looking at: $(($bad.Suite) -join ', ')" -ForegroundColor Red
    exit 1
}
