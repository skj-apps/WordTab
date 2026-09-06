<#
.SYNOPSIS
  Build a VARIANT of the add-in, run a check suite against it, and put the real binary back.

.DESCRIPTION
  Written on 2026-09-05 to find a regression the working tree has and `ef49c79` does not, and kept
  because the question it was built for is still open. See `src\native\RESULT-battery-20260905.md`.

  Word loads `%LOCALAPPDATA%\Programs\WordTab\WordTab.dll`, **not** `src\native\build\`. So an A/B
  between two versions of the add-in means swapping that deployed file, which is the one operation in
  this repo that can leave the machine in a state nobody asked for. Everything below exists to make
  that safe:

    * the working tree is COPIED, never edited, and never stashed
    * the deployed DLL is hashed before anything moves and hash-checked after it goes back
    * Word is force-closed if it outstays the polite wait, because it holds the DLL open and an
      aborted suite leaves it running - the first version of this threw on the restore and left the
      PATCHED binary deployed, which is the one outcome this must never produce
    * the add-in log is emptied before the run, exactly as check-all.ps1 does before each suite

  Two modes:

    -FromTree <files>   start from ef49c79's sources and take only these from the working tree.
                        Bisects WHICH FILE carries a change. Note strip.cpp and frames.cpp each call
                        a function the other defines, so neither can be taken alone - use
                        -StubModalLoop to give ef49c79's frames.cpp a FramesInModalMoveLoop that
                        returns FALSE (= never stand down, which is what ef49c79 did).

    -Patch <name>       start from the working tree and neutralise ONE named change. Bisects WHICH
                        CHANGE inside a file. Add new ones to the switch below; every patch throws if
                        its anchor does not match, which is not optional - a silently unpatched build
                        "proves" whatever you hoped.

.EXAMPLE
  pwsh -File tools\ab-binary.ps1 -FromTree connect.cpp,stack.cpp,wordtab.h
  pwsh -File tools\ab-binary.ps1 -FromTree connect.cpp,stack.cpp,wordtab.h,strip.cpp -StubModalLoop
  pwsh -File tools\ab-binary.ps1 -Patch trybind -Suite reorder

.NOTES
  Three PowerShell traps cost a build each while this was written, and all three are load-bearing:

    1. `git show HEAD:path | Set-Content -NoNewline` joins the line ARRAY with NO separator and
       collapses the whole file onto one line. Extraction goes through cmd redirection instead, and
       every file's size is checked against `git cat-file -s`.
    2. `pwsh -File script.ps1 -FromTree a,b` passes ONE literal string. check-all.ps1:39-42 says this
       in its own words about -Only. Split on commas.
    3. `-like` reads `[` and `]` as a character class, so the anchor `raw[0] == ...` searched for
       `raw0` and matched nothing. Use .Contains().
#>
[CmdletBinding()]
param(
    [string[]]$FromTree,
    [ValidateSet('trybind','dragstand','dragstand2','hover','sampler','captureblt','cursorhold')][string]$Patch,
    [string]$Suite = 'reorder',
    [switch]$StubModalLoop,
    [switch]$KeepConsole
)

$ErrorActionPreference = 'Stop'

if (-not $FromTree -and -not $Patch) { throw 'Give -FromTree or -Patch. With neither, this would build the working tree and prove nothing.' }
if ($FromTree -and $Patch)           { throw '-FromTree and -Patch are different modes; use one.' }

$repo     = Split-Path $PSScriptRoot -Parent
$treeSrc  = Join-Path $repo 'src\native'
$root     = Join-Path $env:TEMP 'wordtab-ab'
$headSrc  = Join-Path $root 'head-src'
$tag      = if ($Patch) { "patch-$Patch" } else { 'fromtree' }
$work     = Join-Path $root $tag
$deployed = Join-Path $env:LOCALAPPDATA 'Programs\WordTab\WordTab.dll'
$backup   = Join-Path $root 'WordTab-deployed-before.dll'
$outFile  = Join-Path $root "$tag-$Suite.txt"

$sources = @('dllmain.cpp','connect.cpp','frames.cpp','strip.cpp','stack.cpp',
             'taskbar.cpp','log.cpp','wordtab.h','wordtab.def','build.ps1')

if (@(Get-Process WINWORD -ErrorAction SilentlyContinue).Count -gt 0) {
    throw 'Word is running and holds the DLL open. Close it first.'
}

New-Item -ItemType Directory -Path $root -Force | Out-Null
$before = (Get-FileHash $deployed -Algorithm SHA256).Hash
Copy-Item $deployed $backup -Force
Write-Host ("deployed DLL backed up: {0}" -f $before.Substring(0,16)) -ForegroundColor DarkGray

if (Test-Path $work) { Remove-Item $work -Recurse -Force }
New-Item -ItemType Directory -Path $work -Force | Out-Null

if ($FromTree) {
    # ef49c79's sources, straight out of git, through cmd redirection so the bytes survive - then
    # SIZE-CHECKED against git's own idea of the blob, because the failure mode of getting this wrong
    # is a file that still compiles-ish and a build that means nothing.
    New-Item -ItemType Directory -Path $headSrc -Force | Out-Null
    Push-Location $repo
    try {
        foreach ($f in $sources) {
            $dest = Join-Path $headSrc $f
            & cmd /c "git show HEAD:src/native/$f > `"$dest`"" 2>&1 | Out-Null
            $want = [int](& git cat-file -s "HEAD:src/native/$f")
            $got  = (Get-Item $dest).Length
            if ($got -ne $want) { throw "extracting $f gave $got bytes, git says $want - the extraction is wrong, not the source" }
        }
    } finally { Pop-Location }

    $FromTree = @($FromTree | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    foreach ($f in $FromTree) {
        if (-not (Test-Path (Join-Path $treeSrc $f))) { throw "no such source file: $f" }
    }

    Copy-Item (Join-Path $headSrc '*') $work -Force
    foreach ($f in $FromTree) { Copy-Item (Join-Path $treeSrc $f) (Join-Path $work $f) -Force }

    if ($StubModalLoop) {
        Add-Content -Path (Join-Path $work 'frames.cpp') -Encoding UTF8 -Value @'

// Added by tools\ab-binary.ps1, not by the project. ef49c79 has no modal-loop accessor and the
// working tree's strip.cpp calls one. FALSE = never stand down = what ef49c79 did.
BOOL FramesInModalMoveLoop(void) { return FALSE; }
'@
        Write-Host 'frames.cpp: FramesInModalMoveLoop stubbed FALSE' -ForegroundColor DarkGray
    }
    Write-Host ("from the working tree: {0}   everything else: ef49c79" -f ($FromTree -join ', ')) -ForegroundColor Cyan
}
else {
    foreach ($f in $sources) { Copy-Item (Join-Path $treeSrc $f) (Join-Path $work $f) -Force }

    $stripPath = Join-Path $work 'strip.cpp'
    $text = [IO.File]::ReadAllText($stripPath)
    # The FILE's newline, not this script's. strip.cpp is CRLF; a multi-line anchor built with the
    # wrong one matches nothing. See [[git-bash-lies-about-line-endings]].
    $nl = if ($text.Contains("`r`n")) { "`r`n" } else { "`n" }

    switch ($Patch) {
        # The unfinished-frame guard. Neutralising it FIXES the IDC_SIZEALL regression and COSTS 17
        # other checks, which is why "just remove it" is not a fix.
        'trybind' {
            $find    = "    if (raw[0] == L'\0' && (!IsWindowVisible(state->frame) || IsIconic(state->frame)))"
            $replace = "    if (0 && raw[0] == L'\0' && (!IsWindowVisible(state->frame) || IsIconic(state->frame)))"
        }
        # TryBind stands down for the whole gesture.
        #
        # **THE ANCHOR CARRIES ITS PRECEDING COMMENT LINE AND THAT IS NOT DECORATION.**
        # `wchar_t raw[256];` + `GetWindowTextW(state->frame, raw, 256);` appears TWICE in strip.cpp -
        # in ReadTitle (~5492) and in TryBind (~5889) - and .Replace() rewrites BOTH. The 2026-09-05
        # runs of this patch and of dragstand2 did exactly that and silently tested a second change
        # nobody asked for; their numbers are marked contaminated in RESULT-battery-20260905.md. The
        # comment line above `wchar_t raw[256];` belongs only to TryBind.
        'dragstand' {
            $find    = @('    // straight back here, exactly as it does for a frame whose `_WwF` is not built yet.',
                         '    wchar_t raw[256];', '    GetWindowTextW(state->frame, raw, 256);') -join $nl
            $replace = @('    // straight back here, exactly as it does for a frame whose `_WwF` is not built yet.',
                         '    if (g_dragging && g_dragFrame)', '        return;', '',
                         '    wchar_t raw[256];', '    GetWindowTextW(state->frame, raw, 256);') -join $nl
        }
        # Stands down only for a frame ALREADY known unfinished.
        #
        # **ITS 2026-09-05 RESULT IS VOID AND THE HYPOTHESIS IT "REFUTED" IS STILL LIVE.** That run
        # hit the two-site anchor bug above, so it also suppressed ReadTitle for the whole of every
        # drag - untitledLogged is set once at strip.cpp:5896 and never cleared. It was written up as
        # refuting "the per-tick GetWindowTextW re-enters Word's wndproc mid-drag" on the grounds that
        # a SUPERSET failing proves the subset fails. **That inference is invalid**: the extra change
        # can break it independently. Re-run this in its current single-site form before concluding
        # anything.
        'dragstand2' {
            $find    = @('    // straight back here, exactly as it does for a frame whose `_WwF` is not built yet.',
                         '    wchar_t raw[256];', '    GetWindowTextW(state->frame, raw, 256);') -join $nl
            $replace = @('    // straight back here, exactly as it does for a frame whose `_WwF` is not built yet.',
                         '    if (state->untitledLogged && g_dragging && g_dragFrame)', '        return;', '',
                         '    wchar_t raw[256];', '    GetWindowTextW(state->frame, raw, 256);') -join $nl
        }
        # THE ONLY CODE DIFFERENCE between the last build known to PASS this check (2026-09-04
        # 17:44) and the failing tree. Reconstructed from the recorded edit history, not guessed:
        # strip.cpp was written exactly three times after that build - 09-05 16:59:06, 16:59:10 and
        # 17:04:37 - and all three edits are the SampleChrome rewrite. connect.cpp, stack.cpp,
        # frames.cpp and wordtab.h were all last written on 09-04 BEFORE that build, so they are
        # byte-identical across the regression. That is what makes these two patches a bisect and
        # not a fishing trip.
        #
        # `0 &&` leaves the DC and bitmap made and destroyed but never blits, so `banded` stays FALSE
        # and every one of the 24 points falls through to `GetPixel(screen, ...)` - the same points,
        # the same y, the same order. That IS the old sampler, ~750,000us and all.
        'sampler' {
            $find    = '    if (band && bandBmp && BitBlt(band, 0, 0, bandW, 1, screen, bandX, y, SRCCOPY | CAPTUREBLT))'
            $replace = '    if (0 && band && bandBmp && BitBlt(band, 0, 0, bandW, 1, screen, bandX, y, SRCCOPY | CAPTUREBLT))'
        }
        # The blit stays - and stays fast - but without CAPTUREBLT. Splits "the sampler is slow again"
        # from "CAPTUREBLT touches the cursor layer". NOTE this is expected to FAIL on the evidence
        # already in hand: TabThemeSample=0 removes the blit ENTIRELY and the check still failed
        # (verified in the 2026-09-05 transcript, `FAIL 1 of 118`). It is run anyway because that
        # evidence is one run, and because a PASS here would overturn it.
        'captureblt' {
            $find    = 'BitBlt(band, 0, 0, bandW, 1, screen, bandX, y, SRCCOPY | CAPTUREBLT)'
            $replace = 'BitBlt(band, 0, 0, bandW, 1, screen, bandX, y, SRCCOPY)'
        }
        # The drag-cursor hold, off. This is the PIN for the two cursor checks: with it neutralised
        # they must BOTH go red, and if they do not then they are not testing the fix. The project has
        # been bitten by exactly that - a new test written to match its own fix, which passed something
        # an existing suite would have failed. See [[green-suites-are-not-a-clean-ab]].
        'cursorhold' {
            $find    = '    if (!g_dragging || !g_dragCursorSet)'
            $replace = '    if (1 || !g_dragging || !g_dragCursorSet)'
        }
        # The hover colour back to its shipped form. A colour cannot move a cursor, which is exactly
        # why it was worth one build to stop assuming so.
        'hover' {
            $find    = "    out->hover    = Mix(out->back, chrome, 50);"
            $replace = "    out->hover    = Step(chrome, -13);"
        }
    }

    # EXACTLY one, not "at least one". .Replace() rewrites every occurrence, so an anchor that matches
    # twice quietly builds a variant carrying a second change - which is how the 2026-09-05 dragstand
    # runs came to patch ReadTitle as well as TryBind and report numbers for something nobody meant to
    # test. Zero matches and two matches are both fatal, and for the same reason: the build would not
    # be the experiment.
    $hits = ([regex]::Matches($text, [regex]::Escape($find))).Count
    if ($hits -ne 1) {
        throw ("patch '{0}' matched {1} place(s) in strip.cpp, need exactly 1. " -f $Patch, $hits) +
              'Zero means the anchor moved; more than one means .Replace() would rewrite all of them. ' +
              'Do NOT run this build and read it as a result - lengthen the anchor until it is unique.'
    }
    [IO.File]::WriteAllText($stripPath, $text.Replace($find, $replace))
    Write-Host "working tree with '$Patch' neutralised" -ForegroundColor Cyan
}

Write-Host '==> building' -ForegroundColor Cyan
$buildLog = & pwsh -File (Join-Path $work 'build.ps1') 2>&1
$dll = Join-Path $work 'build\WordTab.dll'
if (-not (Test-Path $dll)) {
    $buildLog | Select-Object -Last 25 | ForEach-Object { Write-Host $_ }
    throw 'the variant produced no DLL'
}
$variant = (Get-FileHash $dll -Algorithm SHA256).Hash
Write-Host ("variant {0}   deployed-was {1}" -f $variant.Substring(0,16), $before.Substring(0,16)) -ForegroundColor DarkGray
if ($variant -eq $before) { throw 'the variant is byte-identical to what is deployed - nothing would be tested' }

Add-Type -Namespace WordTabAb -Name U -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int n);
'@
$term = $null
if (-not $KeepConsole) {
    $term = Get-Process WindowsTerminal -ErrorAction SilentlyContinue |
            Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
    # Said out loud when it cannot be found. Minimising the console is a MEASUREMENT CONTROL, not a
    # cosmetic - the console is a foreground thief and check-reorder reads tab titles through the
    # foreground window. A silent no-op here would produce a run that looks controlled and is not.
    if (-not $term) {
        Write-Host 'WARNING: no Windows Terminal window found to minimise. If this session is hosted' -ForegroundColor Yellow
        Write-Host '         in a different console, IT can still steal the foreground - check the run for' -ForegroundColor Yellow
        Write-Host '         "taking it back" and for a tab named after your console before believing it.' -ForegroundColor Yellow
    }
}

try {
    Copy-Item $dll $deployed -Force

    # check-all.ps1 empties the log before every suite; running one on its own does not. Seven
    # back-to-back runs took it to 524,176 bytes and log.cpp:42 deleted it MID-SUITE, which the
    # harness caught and threw on - so that run proved nothing in either direction.
    # Moved into %TEMP%\wordtab-ab, NOT into %LOCALAPPDATA%\WordTab\history. That directory is pruned
    # to the newest 40 of each kind (WordTabHarness.ps1:131, :328) and a single 16-suite battery writes
    # 17 logs into it, so A/B evidence parked there is evicted by the next battery or two - which is
    # exactly the evidence you most need to keep.
    $live = Join-Path $env:LOCALAPPDATA 'WordTab\wordtab.log'
    if (Test-Path $live) {
        Move-Item $live (Join-Path $root ('{0}-before-{1}.log' -f (Get-Date -Format 'yyyyMMdd-HHmmss-fff'), $tag)) -Force
    }

    # The console that runs the battery is itself a foreground thief: check-reorder reads a tab's
    # title through the foreground window and recorded `order: Terminal | ...`, five checks red.
    # Minimised, it cannot be handed the foreground when a Word window closes.
    if ($term) { [WordTabAb.U]::ShowWindow($term.MainWindowHandle, 6) | Out-Null; Start-Sleep -Milliseconds 800 }

    $script = if ($Suite -eq 'soak-stack') { Join-Path $PSScriptRoot 'soak-stack.ps1' }
              else { Join-Path $PSScriptRoot "check-$Suite.ps1" }
    if (-not (Test-Path $script)) { throw "no such suite: $script" }
    & pwsh -File $script *>&1 | Set-Content -Path $outFile -Encoding UTF8
}
finally {
    if ($term) { [WordTabAb.U]::ShowWindow($term.MainWindowHandle, 9) | Out-Null }

    for ($i = 0; $i -lt 15; $i++) {
        if (@(Get-Process WINWORD -ErrorAction SilentlyContinue).Count -eq 0) { break }
        Start-Sleep -Seconds 2
    }
    if (@(Get-Process WINWORD -ErrorAction SilentlyContinue).Count -gt 0) {
        Write-Host 'Word outstayed the wait and still holds the DLL - closing it so the restore can run' -ForegroundColor Yellow
        Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
        Start-Sleep -Seconds 3
    }
    Copy-Item $backup $deployed -Force
    $after = (Get-FileHash $deployed -Algorithm SHA256).Hash
    if ($after -eq $before) {
        Write-Host ("RESTORED: the deployed DLL is what it was ({0})" -f $after.Substring(0,16)) -ForegroundColor Green
    }
    else {
        # THROW, not just print. A leftover variant binary is the worst state this script can produce
        # and the person who needs to know may be a wrapper script or somebody who scrolled past a red
        # line. Printing alone exits 0, which reads as a clean run.
        Write-Host ("*** RESTORE FAILED *** deployed {0}, wanted {1}" -f $after.Substring(0,16), $before.Substring(0,16)) -ForegroundColor Red
        throw ("RESTORE FAILED: {0} is NOT the binary that was there before this run. " -f $deployed) +
              ("A copy of the original is at {0}. Close Word and put it back before doing anything else." -f $backup)
    }
}

Write-Host ''
Write-Host "the suite's output: $outFile" -ForegroundColor DarkGray
Get-Content $outFile | Where-Object { $_ -match '^\s*FAIL|cursor during|checks' } |
    Select-Object -Last 12 | ForEach-Object { Write-Host "  $_" }
