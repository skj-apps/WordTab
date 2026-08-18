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
$suites = @('stack', 'row', 'strip', 'tabs', 'menu', 'reorder', 'startscreen', 'look', 'scroll', 'title', 'dot', 'soak-stack')
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

# The shared harness, for Reset-LogFile below. Nothing else in this runner needs it, and the log
# primitives are the one part of it that loads without [WordLayout].
. (Join-Path $PSScriptRoot 'WordTabHarness.ps1')

<#
  Prove the shared primitives before eleven suites are run on top of them.

  One second, no Word, no desktop - and it covers the one failure that a real run cannot reach on
  purpose: the add-in deletes its log at 512KB, which takes two whole batteries to provoke, and what
  it produces is assertions passing having read nothing. Running the battery on a harness that
  cannot detect that is running it on an instrument nobody checked.
#>
Write-Host ''
# NOT called $selfTest. The harness declares `param([switch]$SelfTest)`, dot-sourcing it above leaves
# that parameter in THIS scope with its [switch] type still attached, and PowerShell variable names
# are case-insensitive - so `$selfTest = <the output>` failed with "Cannot convert System.Object[] to
# SwitchParameter" and this gate silently reported "(no summary line)" for a self-test that passed.
# The harness gives the name back now; the distinct name here is the belt to that braces.
$selfOutput = & pwsh -File (Join-Path $PSScriptRoot 'WordTabHarness.ps1') -SelfTest 2>&1
$selfCode = $LASTEXITCODE
if ($selfCode -ne 0) {
    $selfOutput | ForEach-Object { Write-Host $_ }
    Write-Host ''
    Write-Host 'The harness self-test failed. Not running any suite on top of it.' -ForegroundColor Red
    exit 1
}
$selfSummary = @($selfOutput | ForEach-Object { "$_" } | Where-Object { $_ -match '\d+ checks' })
if ($selfSummary.Count -eq 0) {
    # A gate that cannot read its own result must not wave the run through. The first version of this
    # printed "(no summary line)" and carried on - which is the same shape as `check-all -Only` once
    # matching no suites and reporting "All 0 suites green."
    $selfOutput | ForEach-Object { Write-Host $_ }
    Write-Host ''
    Write-Host 'The harness self-test exited 0 but said nothing this could read. Not running any suite on top of it.' -ForegroundColor Red
    exit 1
}
Write-Host ("harness self-test: {0}" -f $selfSummary[$selfSummary.Count - 1].Trim()) -ForegroundColor DarkGray

# Whatever was in the log before the battery started keeps its own file, rather than being deleted by
# the first suite that pushes the total past 512KB.
Reset-LogFile -Label 'before-battery' | Out-Null

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

    # That suite's log, under that suite's name, and the next one starts from an empty file.
    #
    # Two things at once. **The add-in DELETES its log once it passes 512KB** (log.cpp:42) and a
    # battery writes about 250KB, so the second battery in a session used to roll it mid-suite - it
    # did, inside check-title - and every assertion reading the log from a mark was then reading a
    # file its evidence had been deleted from. From an empty start no single suite comes close:
    # measured across a whole battery, they run 9.0KB (startscreen) to 47.8KB (scroll).
    #
    # And it makes the battery's own evidence survive. "The suite that reports a failure is not
    # necessarily the suite that caused it" has cost this project two diagnoses, and until now the
    # only record left afterwards was whichever 512KB happened to be last.
    #
    # **AFTER the suite, not before it.** The first version reset before each run, which named every
    # archive after the suite that was about to start rather than the one that wrote it - a file
    # called title.log holding scroll's output, which is worse than no file at all.
    Reset-LogFile -Label $name | Out-Null

    # The suite's own console output, kept beside its log, and this is the other half of "the
    # battery's evidence has to survive the battery".
    #
    # The archived log says what the ADD-IN did. This says what the SUITE said about it - and it is
    # the only thing in the world that names *which assertion* went red. That mattered twice on
    # 2026-08-16: a single check in `reorder` failed inside a 14-minute battery whose console output
    # had been captured in a way that kept only the summary, and it took three more runs to establish
    # that it would not reproduce - and it still has no name. A battery that cannot say what failed
    # has to be run again to find out, which is the most expensive kind of missing evidence here.
    $historyDir = $script:LogHistoryDir
    if (-not $historyDir) { $historyDir = Join-Path $env:LOCALAPPDATA 'WordTab\history' }
    try {
        New-Item -ItemType Directory -Path $historyDir -Force | Out-Null
        $transcriptName = '{0}-{1}.out.txt' -f (Get-Date -Format 'yyyyMMdd-HHmmss-fff'), $name
        $transcriptPath = Join-Path $historyDir $transcriptName
        $output | ForEach-Object { "$_" } | Set-Content -Path $transcriptPath -Encoding UTF8
        Write-Host "    the suite's own output was kept in history\$transcriptName" -ForegroundColor DarkGray

        # Pruned on the same rule as the logs, and separately from them: they are two files per suite
        # per battery and letting one kind push the other out of the window would leave halves of
        # pairs behind.
        $oldOut = @(Get-ChildItem -Path $historyDir -Filter '*.out.txt' -ErrorAction SilentlyContinue |
                    Sort-Object LastWriteTime -Descending | Select-Object -Skip 40)
        foreach ($f in $oldOut) { Remove-Item -Path $f.FullName -Force -ErrorAction SilentlyContinue }
    } catch {
        Write-Host "    could not keep the suite's output ($($_.Exception.Message))" -ForegroundColor Yellow
        $transcriptPath = $null
    }

    # Not anchored: the suites do not agree on how they announce the total. Some print
    # "34 checks, all passed." at column 0, others "PASS  40 checks, 0 failures". Anchoring this
    # matched neither on the first run and every suite reported "(no summary line)" while passing.
    $tail = @($output | ForEach-Object { "$_" } | Where-Object { $_ -match '\d+ checks' })
    $summary = if ($tail.Count -gt 0) { $tail[-1].Trim() } else { '(no summary line)' }
    $fails = @($output | ForEach-Object { "$_" } | Where-Object { $_ -match '^\s*FAIL ' })

    $results += [pscustomobject]@{
        Suite      = $name
        Exit       = $code
        Summary    = $summary
        Fails      = $fails.Count
        Minutes    = [Math]::Round(((Get-Date) - $started).TotalMinutes, 1)
        Transcript = $transcriptPath
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

    # Named here rather than left to be found. A red suite in a long battery is read hours later by
    # somebody who no longer has the console, and the two files that answer "what actually happened"
    # are both on disk under names nobody would guess.
    Write-Host ''
    Write-Host 'What each of them said, and what the add-in was doing at the time:' -ForegroundColor DarkGray
    foreach ($r in $bad) {
        if ($r.Transcript) {
            Write-Host ("  {0,-12} {1}" -f $r.Suite, $r.Transcript) -ForegroundColor DarkGray
        } else {
            Write-Host ("  {0,-12} (its output could not be kept)" -f $r.Suite) -ForegroundColor Yellow
        }
    }
    Write-Host ("  add-in logs alongside them in {0}" -f
                $(if ($script:LogHistoryDir) { $script:LogHistoryDir } else { Join-Path $env:LOCALAPPDATA 'WordTab\history' })) -ForegroundColor DarkGray
    exit 1
}
