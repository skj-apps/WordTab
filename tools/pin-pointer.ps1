# Pin for Set-Pointer: does calling it TWICE at the same point deliver two moves, or one?
#
# move-delivery.exe watch <ms> puts up a window that counts the WM_MOUSEMOVEs it is sent and prints
# the total. This drives the REAL Set-Pointer from tools\WordTabHarness.ps1 at it - once, then again
# at the identical point - which is exactly the shape of a suite that parks the pointer, reads a
# button, and then hovers the same place.
#
# Expected with the fix:     2 or more   (the second call nudges off and comes back)
# Expected without it:       1           (the second call is a no-op that still returns $true)

param([switch]$Old)
$ErrorActionPreference = 'Stop'
$exe = Join-Path $env:TEMP 'wordtab-move-delivery.exe'
if (-not (Test-Path $exe)) {
    # Built here rather than assumed, so this is one command for whoever runs it next.
    $toolkit = Join-Path $env:LOCALAPPDATA 'Programs\w64devkit\w64devkit'
    $env:PATH = (Join-Path $toolkit 'bin') + [IO.Path]::PathSeparator + $env:PATH
    & (Join-Path $toolkit 'bin\g++.exe') -O2 -mconsole -DUNICODE -D_UNICODE `
        -o $exe (Join-Path $PSScriptRoot 'move-delivery.cpp') -luser32
    if (-not (Test-Path $exe)) { throw 'could not build move-delivery.exe' }
}

$source = Get-Content -Raw -Path (Join-Path $PSScriptRoot 'WordLayout.cs')
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition $source -Language CSharp -ReferencedAssemblies @(
    'System.Runtime', 'System.Collections', 'System.Threading.Thread', 'netstandard'
)
[WordLayout]::MakeDpiAware() | Out-Null

# The harness wants these; it is dot-sourced for Set-Pointer and Get-ClassAt only.
$script:Failures = 0
$script:Checks   = 0
function Write-Note($text) { Write-Host "    $text" -ForegroundColor DarkGray }
. (Join-Path $PSScriptRoot 'WordTabHarness.ps1')

# The control arm. This is the OLD Set-Pointer body, verbatim, redefined on top of the real one so the
# same rig measures the same thing twice. Without it, "the fix works" is a number with nothing to
# compare it to - and this project has shipped a check that passed with its own fix removed before.
if ($Old) {
    Write-Host 'CONTROL ARM: the pre-fix Set-Pointer' -ForegroundColor Yellow
    function Set-Pointer($x, $y, $what) {
        for ($try = 1; $try -le 5; $try++) {
            [WordLayout]::MouseTo($x, $y)
            Start-Sleep -Milliseconds 250
            $at = [WordLayout]::Cursor()
            if (([Math]::Abs($at.X - $x) -le 2) -and ([Math]::Abs($at.Y - $y) -le 2)) { return $true }
            Start-Sleep -Milliseconds 500
        }
        return $false
    }
}

$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName               = $exe
$psi.Arguments              = 'watch 4000'
$psi.RedirectStandardOutput = $true
$psi.UseShellExecute        = $false
$p = [System.Diagnostics.Process]::Start($psi)

$rectLine = $p.StandardOutput.ReadLine()      # RECT l t r b
if ($rectLine -notmatch '^RECT (-?\d+) (-?\d+) (-?\d+) (-?\d+)') { throw "unexpected: $rectLine" }
$l = [int]$Matches[1]; $t = [int]$Matches[2]; $r = [int]$Matches[3]; $b = [int]$Matches[4]
$x = [int](($l + $r) / 2)
$y = [int](($t + $b) / 2)
Write-Host ("the watcher is at ({0},{1} {2}x{3}); aiming at ({4},{5})" -f $l, $t, ($r-$l), ($b-$t), $x, $y)

# Refuse to report a number taken while somebody is using the mouse. Every WM_MOUSEMOVE counts the
# same whether this script injected it or a hand did, so a busy rig makes the CONTROL arm pass and the
# whole pin meaningless. Seen on the first run of this.
$quietLine = $p.StandardOutput.ReadLine()     # QUIET n
if ($quietLine -notmatch '^QUIET (\d+)') { throw "unexpected: $quietLine" }
$quiet = [int]$Matches[1]
if ($quiet -ne 0) {
    Write-Host ''
    Write-Host ("SKIPPED  the pointer is being moved by something else ({0} stray move(s) in 600ms of nothing)." -f $quiet) -ForegroundColor Yellow
    Write-Host '         Take your hand off the mouse and run it again - a count taken now proves nothing.' -ForegroundColor Yellow
    $null = $p.StandardOutput.ReadToEnd()
    $p.WaitForExit()
    exit 2
}

# Start from somewhere else, so call 1 is unambiguously a real move.
[WordLayout]::MouseTo(($l + 5), ($t + 5))
Start-Sleep -Milliseconds 300

$one = Set-Pointer $x $y 'pin: first call'
$two = Set-Pointer $x $y 'pin: second call at the SAME point'
Write-Host ("Set-Pointer returned {0} then {1}" -f $one, $two)

$movesLine = $p.StandardOutput.ReadLine()     # MOVES n
$p.WaitForExit()
if ($movesLine -notmatch '^MOVES (\d+)') { throw "unexpected: $movesLine" }
$moves = [int]$Matches[1]

Write-Host ''
Write-Host ("WM_MOUSEMOVE delivered to the watcher: {0}" -f $moves)
if ($moves -ge 2) {
    Write-Host 'PASS  the second call delivered a move of its own' -ForegroundColor Green
} else {
    Write-Host 'FAIL  the second call delivered nothing - Set-Pointer returned true for a move nobody got' -ForegroundColor Red
}
