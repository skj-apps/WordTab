<#
.SYNOPSIS
  Drive and assert the unsaved-changes dot: the mark a tab carries when its document has been
  changed and not yet saved.

.DESCRIPTION
  This is the first thing WordTab draws that is about Word's *documents* rather than its windows, so
  the suite has to prove two separate things and never confuse them:

    - **the state**, which is Word's `Document.Saved` read through the object model on the janitor.
      The add-in logs it when it changes (`dot |on|` / `dot |off|`), and the assertions here are
      against those lines
    - **the drawing**, which is a filled dot standing in the close button's own square, replaced by
      the x the moment the pointer arrives on the tab

  A log line saying what was computed is not evidence of what was drawn - this project has had that
  failure twice - so every state assertion is paired with a photograph, and the photograph
  distinguishes the two shapes rather than merely noticing that something changed:

    - **on the axis, close to the centre**, a disc has ink and the x has none: the x is two diagonal
      strokes, so at the same height as the centre there is nothing beside it
    - **on the diagonal, out at the arm's length**, the x has ink and the disc has none: the disc's
      radius is 3 logical px and the arms reach 4 in each direction, which is 5.7 away

  Both probes carry a control taken from a corner of the same close button in the same photograph, so
  nothing here depends on knowing an absolute colour.

  What is deliberately NOT reproduced from measurement: a Protected View document. Word keeps such a
  document in a sandboxed process and it appears in neither `Application.Windows` nor `Documents` -
  measured in tools\probe-saved.ps1 - so it can never carry a dot. What matters is that it does not
  poison the pass for anything else, and that is what section 6 asserts: with one open, an ordinary
  document's dot still lights.

  **The close button is still the close button.** Section 4 clicks the dot on a modified document and
  requires Word's save prompt to appear, then cancels it with Escape and requires the document to
  still be open. That is a real proof that the rectangle still hit-tests as HIT_CLOSE, and it
  discards nothing: Escape on that prompt is Cancel.

  Fixtures are authored by Word in %TEMP%\wordtab-dot and are disposable. Word is closed at both ends.

.PARAMETER KeepOpen
  Leave Word running afterwards.

.PARAMETER Screenshot
  Write the close-button photographs to -ShotDir.

.EXAMPLE
  pwsh -File tools\check-dot.ps1
  pwsh -File tools\check-dot.ps1 -Screenshot
#>
[CmdletBinding()]
param(
    [switch]$KeepOpen,
    [switch]$Screenshot,
    [string]$ShotDir = $env:TEMP
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$source = Get-Content -Raw -Path (Join-Path $PSScriptRoot 'WordLayout.cs')
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition $source -Language CSharp -ReferencedAssemblies @(
    'System.Runtime', 'System.Collections', 'System.Threading.Thread', 'netstandard'
)
[WordLayout]::MakeDpiAware() | Out-Null

$script:Failures = 0
$script:Checks   = 0

trap {
    Write-Host ''
    Write-Host "ERROR  $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
    break
}

function Write-Step($text) { Write-Host "==> $text" -ForegroundColor Cyan }
function Write-Note($text) { Write-Host "    $text" -ForegroundColor DarkGray }
function Assert($condition, $text) {
    $script:Checks++
    if ($condition) { Write-Host "    PASS  $text" -ForegroundColor Green }
    else { $script:Failures++; Write-Host "    FAIL  $text" -ForegroundColor Red }
}

$VK = @{ S = 0x53; X = 0x58; ESC = 0x1B }

function Get-WordPids { @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }) }
function Get-Frames {
    $found = @()
    foreach ($id in Get-WordPids) { $found += [WordLayout]::Frames($id) }
    return @($found)
}
# Always through this - a PowerShell function returning an empty array returns *nothing*, and
# `(Get-Frames).Count` is then a hard error under StrictMode at the two moments that matter most.
function Get-FrameCount { return @(Get-Frames).Count }

function Get-Parts($frame) {
    $kids = [WordLayout]::Children($frame)
    [pscustomobject]@{
        Frame = $frame
        Rect  = [WordLayout]::RectOf($frame)
        Strip = @($kids | Where-Object { $_.Class -eq 'WordTabStrip' })[0]
        Wwf   = @($kids | Where-Object { $_.Class -eq '_WwF' })[0]
        Title = [WordLayout]::TitleOf($frame)
    }
}

# ---- the log, read from a mark ----------------------------------------------------------------
#
# Window handles are reused across Word restarts and this suite restarts Word twice, so the log is
# never searched from the beginning: a mark is taken and only what was written after it counts.
# Without that, the TabDot=0 section would happily match the line the TabDot=1 section wrote for a
# handle that has since been recycled, and pass by reading its predecessor's evidence.

$LogPath = Join-Path $env:LOCALAPPDATA 'WordTab\wordtab.log'
$script:LogMark = 0

function Set-LogMark {
    $script:LogMark = if (Test-Path $LogPath) { (Get-Item $LogPath).Length } else { 0 }
}

function Get-LogSince($pattern) {
    if (-not (Test-Path $LogPath)) { return @() }
    $stream = [System.IO.File]::Open($LogPath, 'Open', 'Read', 'ReadWrite')
    try {
        $offset = [Math]::Min([int64]$script:LogMark, $stream.Length)
        $stream.Seek([int64]$offset, 'Begin') | Out-Null
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
        $text = $reader.ReadToEnd()
    } finally { $stream.Dispose() }
    return @(($text -split "`r?`n") | Where-Object { $_ -like "*$pattern*" })
}

# The add-in's own account of the dot for one window, since the mark: the LAST thing it said, or
# $null if it has never said anything about that window. "Never said anything" is a real answer here
# and must not read as "off" - a poll that never ran and a poll that answered "clean" are exactly the
# pair this project keeps having to tell apart.
function Get-DotState($frame) {
    $key = 'hwnd=0x{0:X16}' -f [int64]$frame
    $lines = @(Get-LogSince 'dot |' | Where-Object { $_ -like "*$key*" })
    if ($lines.Count -eq 0) { return $null }
    $last = $lines[$lines.Count - 1]
    if ($last -match 'dot \|(?<s>on|off)\|  document \|(?<d>.*)\|\s*$') {
        return [pscustomobject]@{ On = ($Matches['s'] -eq 'on'); Document = $Matches['d'] }
    }
    return $null
}

# Wait for the add-in to report a particular state for a window. The janitor runs at 2Hz, so this is
# a wait rather than a read, and the timeout is generous enough that a slow tick is never the reason
# a check fails.
function Wait-Dot($frame, $on, $seconds = 12) {
    $deadline = (Get-Date).AddSeconds($seconds)
    while ((Get-Date) -lt $deadline) {
        $state = Get-DotState $frame
        if ($state -and $state.On -eq $on) { return $state }
        Start-Sleep -Milliseconds 400
    }
    return Get-DotState $frame
}

function Set-WordForeground($seconds = 4) {
    $frames = @(Get-Frames)
    if ($frames.Count -eq 0) { return [IntPtr]::Zero }

    for ($round = 1; $round -le 3; $round++) {
        $now = [WordLayout]::GetForeground()
        if ($frames -contains $now) { return $now }

        Write-Note ("the foreground was `"{0}`" - taking it back" -f [WordLayout]::TitleOf($now))
        [WordLayout]::Focus($frames[0]) | Out-Null
        $deadline = (Get-Date).AddSeconds($seconds)
        while ((Get-Date) -lt $deadline) {
            $now = [WordLayout]::GetForeground()
            if (@(Get-Frames) -contains $now) { return $now }
            Start-Sleep -Milliseconds 250
        }

        # This suite drives real clicks at screen coordinates, so a window over Word receives them.
        # Never Progman or WorkerW: that is the desktop, and minimising it is "Show Desktop".
        $now = [WordLayout]::GetForeground()
        $cls = if ($now -ne [IntPtr]::Zero) { [WordLayout]::ClassOf($now) } else { '' }
        if ($now -ne [IntPtr]::Zero -and -not (@(Get-Frames) -contains $now) -and
            $cls -ne 'Progman' -and $cls -ne 'WorkerW') {
            Write-Note ("minimising `"{0}`" - it is sitting over Word" -f [WordLayout]::TitleOf($now))
            [WordLayout]::Show($now, 6)      # SW_MINIMIZE
            Start-Sleep -Milliseconds 800
        }
    }
    return [WordLayout]::GetForeground()
}

function Get-TopStrip {
    Set-WordForeground | Out-Null
    $frames = @(Get-Frames)
    if ($frames.Count -eq 0) { throw 'No Word windows.' }
    $top = [WordLayout]::GetForeground()
    if (-not ($frames -contains $top)) { $top = $frames[0] }
    $parts = Get-Parts $top
    if (-not $parts.Strip) { throw 'The foreground Word window has no WordTab strip.' }
    return $parts
}

# The row as it is *right now*. Never a rectangle measured earlier: the strip moves, 46px between two
# clicks a second apart, measured.
function Get-Row {
    $top = Get-TopStrip
    $count = Get-FrameCount
    return [pscustomobject]@{
        Strip  = $top.Strip.Hwnd
        Dpi    = [WordLayout]::Dpi($top.Strip.Hwnd)
        Layout = [WordLayout]::Tabs($top.Strip.Hwnd, $count)
        Count  = $count
    }
}

function Get-Spot($index) {
    $row = Get-Row
    if ($index -ge $row.Layout.Tabs.Length) { throw "Tab $index does not exist ($($row.Layout.Tabs.Length) tabs)." }
    $t = $row.Layout.Tabs[$index]
    return [pscustomobject]@{
        X = $t.Left + [int](($t.Right - $t.Left) / 3)
        Y = [int](($t.Top + $t.Bottom) / 2)
        Rect = $t; Strip = $row.Strip; Count = $row.Count
    }
}

function Format-Rect($r) { "({0},{1} {2}x{3})" -f $r.Left, $r.Top, ($r.Right - $r.Left), ($r.Bottom - $r.Top) }

# Click a point on the strip, having first checked that the strip is what is under it. `point` is a
# scriptblock re-run on every attempt, because the strip moves and because taking the foreground back
# can move it again. AN INJECTED INPUT THAT WAS NOT CONFIRMED TO LAND IS NOT AN INPUT.
function Invoke-StripClick($what, [scriptblock]$point) {
    for ($try = 1; $try -le 3; $try++) {
        Set-WordForeground | Out-Null
        $p = & $point
        $cls = [WordLayout]::ClassOf([WordLayout]::WindowAt($p.X, $p.Y))
        if ($cls -eq 'WordTabStrip') {
            [WordLayout]::Click($p.X, $p.Y)
            return $true
        }
        Write-Note ("{0}: ({1},{2}) is over `"{3}`", not the strip - taking the foreground back and re-measuring" -f
                    $what, $p.X, $p.Y, $cls)
        $blocking = Get-WordDialog
        if ($blocking) { Write-Note ("  Word has a dialog up: {0}" -f (Format-Dialog $blocking)) }
        Start-Sleep -Milliseconds 800
    }
    return $false
}

# Move the pointer, and prove it arrived. SendInput's absolute move silently does not take when
# another process holds the foreground or a hand is on the mouse, and a hover that never happened
# looks exactly like a hover that was not drawn.
function Set-Pointer($x, $y, $what) {
    for ($try = 1; $try -le 5; $try++) {
        [WordLayout]::MouseTo($x, $y)
        Start-Sleep -Milliseconds 250
        $at = [WordLayout]::Cursor()
        if (([Math]::Abs($at.X - $x) -le 2) -and ([Math]::Abs($at.Y - $y) -le 2)) { return $true }
        Write-Note ("{0}: asked for ({1},{2}), the pointer is at ({3},{4}) - trying again" -f
                    $what, $x, $y, $at.X, $at.Y)
        Start-Sleep -Milliseconds 500
    }
    return $false
}

# Word asking the user something.
#
# `Net UI Tool Window` is excluded, and that exclusion cost a whole battery run to learn. It is
# Office's floating-UI host - galleries, task-pane popouts, notification toasts - not a question, and
# one of them appeared at 622x298 while check-title was closing seven windows. Because "any visible
# top-level window of Word's that is not an OpusApp and is bigger than 150x60" matched it, the suite
# refused to close Word, left it running, and the two suites after it failed on the leftovers: the
# dot suite hit the same window, and soak-stack opened six documents and measured ELEVEN.
#
# What a real prompt is, measured rather than assumed - twice in that same battery: the save prompt
# is `NUIDialog` at 920x713 and the Save As question is `NUIDialog` at 920x641. The size floor stays
# as a second filter.
#
# NOTE: the other nine suites still carry the unnarrowed version. Retrofitting them is deliberately
# not done here - check-menu's prompt detection is the most safety-critical assertion in the project
# and it is not this slice's to change - but it belongs on the same list as the Set-Pointer retrofit.
function Get-WordDialog {
    foreach ($id in Get-WordPids) {
        foreach ($w in [WordLayout]::TopLevel($id)) {
            if (-not $w.Visible) { continue }
            if ($w.Class -eq 'OpusApp') { continue }
            if ($w.Class -eq 'Net UI Tool Window') { continue }
            $width = $w.Right - $w.Left
            $height = $w.Bottom - $w.Top
            if ($width -lt 150 -or $height -lt 60) { continue }
            return [pscustomobject]@{ Hwnd = $w.Hwnd; Class = $w.Class; Title = $w.Title
                                      Width = $width; Height = $height }
        }
    }
    return $null
}

function Format-Dialog($d) {
    if (-not $d) { return '(nothing)' }
    return "class={0} {1}x{2} |{3}|" -f $d.Class, $d.Width, $d.Height, $d.Title
}

# Word asking the user something, as opposed to Word's own chrome going past. Confirmed twice, a
# second and a half apart: a single sample reads shutdown chrome as a save prompt.
function Get-SavePrompt {
    if (-not (Get-WordDialog)) { return $null }
    Start-Sleep -Milliseconds 1500
    return Get-WordDialog
}

function Wait-Prompt($present, $seconds = 10) {
    $deadline = (Get-Date).AddSeconds($seconds)
    while ((Get-Date) -lt $deadline) {
        $d = Get-WordDialog
        if ($present -and $d) { return $d }
        if (-not $present -and -not $d) { return $null }
        Start-Sleep -Milliseconds 400
    }
    return Get-WordDialog
}

# ---- photographs ------------------------------------------------------------------------------
#
# CopyFromScreen, not PrintWindow: this is about what the user can see.

function Get-RectShot($r) {
    $bmp = New-Object System.Drawing.Bitmap(($r.Right - $r.Left), ($r.Bottom - $r.Top))
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($r.Left, $r.Top, 0, 0, $bmp.Size)
    $g.Dispose()
    return [pscustomobject]@{ Bitmap = $bmp; Origin = $r }
}

function Get-Pixel($shot, $x, $y) {
    $px = $x - $shot.Origin.Left
    $py = $y - $shot.Origin.Top
    if ($px -lt 0 -or $py -lt 0 -or $px -ge $shot.Bitmap.Width -or $py -ge $shot.Bitmap.Height) { return $null }
    return $shot.Bitmap.GetPixel($px, $py)
}

function Test-Near($a, $b, $tolerance) {
    if ($null -eq $a -or $null -eq $b) { return $false }
    return ([Math]::Abs($a.R - $b.R) -le $tolerance) -and
           ([Math]::Abs($a.G - $b.G) -le $tolerance) -and
           ([Math]::Abs($a.B - $b.B) -le $tolerance)
}

function Measure-Changed($a, $b, $origin, $rect, $exclude) {
    $changed = 0
    for ($y = $rect.Top; $y -lt $rect.Bottom; $y += 2) {
        for ($x = $rect.Left; $x -lt $rect.Right; $x += 2) {
            if ($exclude -and $x -ge $exclude.Left -and $x -lt $exclude.Right -and
                $y -ge $exclude.Top -and $y -lt $exclude.Bottom) { continue }
            $px = $x - $origin.Left
            $py = $y - $origin.Top
            if ($px -lt 0 -or $py -lt 0 -or $px -ge $a.Width -or $py -ge $a.Height) { continue }
            if ($a.GetPixel($px, $py).ToArgb() -ne $b.GetPixel($px, $py).ToArgb()) { $changed++ }
        }
    }
    return $changed
}

# What shape is in a tab's close button, read from a photograph and named rather than merely counted.
#
# The two probes are what tell a disc from an x, and each is paired with a control taken from a
# corner of the same square in the same photograph - so this reads no absolute colour and cannot be
# fooled by a theme, a DPI or a repaint.
#
#   OnAxis    a point beside the centre at the same height. A disc covers it; the x's two diagonal
#             strokes are nowhere near it.
#   Diagonal  a point out where an arm ends. The x runs through it; the disc's radius is 3 logical
#             px against an arm reaching 4 in each direction, which is 5.7 away.
function Read-Button($close, $dpi, $label) {
    $shot = Get-RectShot $close
    try {
        if ($Screenshot) {
            $path = Join-Path $ShotDir ("wordtab-dot-{0}.png" -f ($label -replace '[^\w\-]', '-'))
            $shot.Bitmap.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
            Write-Note "photograph: $path"
        }

        $cx = [int](($close.Left + $close.Right) / 2)
        $cy = [int](($close.Top + $close.Bottom) / 2)
        $r  = [WordLayout]::Scale(3, $dpi)      # DOT_LOGICAL_RADIUS
        $arm = [WordLayout]::Scale(4, $dpi)     # the x's arm, from DrawGlyph

        # The control: the left edge of the square at the centre's height. Neither shape reaches it -
        # the x's arm is 4 logical px and the square is 8 from its centre to its edge - and it is
        # comfortably inside the hover chip, which is the square inflated by 2. So on a hovered tab
        # the ground is the chip's colour and on a resting one it is the card's, and either way it is
        # the right thing to call "not ink" because it came out of the same photograph.
        $ground = Get-Pixel $shot ($close.Left + 1) $cy

        $onAxis = @()
        foreach ($dx in @(-($r - 1), ($r - 1))) { $onAxis += ,(Get-Pixel $shot ($cx + $dx) $cy) }

        $diagonal = @()
        foreach ($d in @(-($arm - 1), ($arm - 1))) { $diagonal += ,(Get-Pixel $shot ($cx + $d) ($cy + $d)) }

        # "Ink" is simply "not the ground of this same square". Tolerance 24 because the shapes are
        # anti-aliased and a probe one pixel off the solid part is a blend rather than a full hit.
        $inkOnAxis   = @($onAxis   | Where-Object { -not (Test-Near $_ $ground 24) }).Count
        $inkDiagonal = @($diagonal | Where-Object { -not (Test-Near $_ $ground 24) }).Count

        Write-Note ("{0}: close {1} dpi={2} r={3} arm={4}  ink on-axis {5}/2  on the diagonal {6}/2" -f
                    $label, (Format-Rect $close), $dpi, $r, $arm, $inkOnAxis, $inkDiagonal)

        return [pscustomobject]@{
            OnAxis = $inkOnAxis; Diagonal = $inkDiagonal
            IsDot = ($inkOnAxis -eq 2 -and $inkDiagonal -eq 0)
            IsCross = ($inkDiagonal -eq 2)
        }
    } finally { $shot.Bitmap.Dispose() }
}

# The close button of one tab, right now, with the pointer somewhere it cannot affect the answer.
function Read-TabButton($index, $label) {
    Assert (Set-Pointer 4 4 "parking the pointer before reading $label") "the pointer parked clear of the strip for $label"
    Start-Sleep -Milliseconds 600
    $row = Get-Row
    if ($index -ge $row.Layout.Close.Length) { throw "Tab $index does not exist." }
    $close = $row.Layout.Close[$index]
    if ([WordLayout]::IsEmptyRect($close)) { throw "Tab $index has no close button rectangle - the window is too narrow." }
    return Read-Button $close $row.Dpi $label
}

# ---- Word up and down --------------------------------------------------------------------------

function Close-Word {
    if (@(Get-WordPids).Count -eq 0) { return }
    Write-Note "closing Word: $(Get-FrameCount) window(s)"
    for ($guard = 0; $guard -lt 24; $guard++) {
        $open = @(Get-Frames)
        if ($open.Count -eq 0) { break }
        [WordLayout]::Close($open[0])
        Start-Sleep -Milliseconds 1200
        # A leftover dirty document. Deliberately not answered: a test script may not discard work it
        # did not create. This suite makes documents dirty on purpose, so it saves them itself before
        # ever getting here.
        $prompt = Get-SavePrompt
        if ($prompt) {
            Write-Note ("Word put up: {0}" -f (Format-Dialog $prompt))
            [WordLayout]::Press($VK.ESC)
            Start-Sleep -Seconds 2
            throw ("Word is asking something and it is still there after Escape: {0}  Answer it by hand and re-run." -f
                   (Format-Dialog $prompt))
        }
    }
    foreach ($p in @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue)) { $p.CloseMainWindow() | Out-Null }
    $deadline = (Get-Date).AddSeconds(20)
    while ((Get-Date) -lt $deadline -and @(Get-WordPids).Count -gt 0) { Start-Sleep -Milliseconds 500 }
    if (@(Get-WordPids).Count -gt 0) { throw 'Word would not close - close it by hand and re-run.' }
    Start-Sleep -Seconds 2
}

function Open-Document($path, $first) {
    $before = @(Get-Frames)
    Start-Process -FilePath 'winword.exe' -ArgumentList "`"$path`""
    Start-Sleep -Seconds $(if ($first) { 16 } else { 7 })
    $deadline = (Get-Date).AddSeconds(30)
    while ((Get-Date) -lt $deadline) {
        $new = @(Get-Frames | Where-Object { $before -notcontains $_ })
        if ($new.Count -gt 0) { Start-Sleep -Seconds 2; return $new[0] }
        Start-Sleep -Milliseconds 500
    }
    return [IntPtr]::Zero
}

# Put the caret in the document behind tab $index and type, and *prove* it before returning. A
# keystroke sent to a window that was only activated programmatically goes nowhere in Word, which is
# why this clicks into the page first; and a missed keystroke is indistinguishable from a broken
# feature three steps away, which is why the edit is photographed.
function Set-Dirty($index) {
    for ($try = 1; $try -le 6; $try++) {
        Set-WordForeground | Out-Null
        $spot = Get-Spot $index
        [WordLayout]::Click($spot.X, $spot.Y)
        Start-Sleep -Milliseconds 900

        $top = Get-TopStrip
        if (-not $top.Wwf) { throw 'The foreground Word window has no document frame to type into.' }

        # Aimed at the document frame itself, never at a fixed offset from the strip. Word remembers
        # its window size between sessions, and a hard-coded offset has landed on the desktop.
        $view = [WordLayout]::RectOf($top.Wwf.Hwnd)
        $vw = $view.Right - $view.Left
        $vh = $view.Bottom - $view.Top
        if ($vw -lt 200 -or $vh -lt 200) { throw "The document frame is only ${vw}x${vh} - too small to type into." }

        $x = $view.Left + [int]($vw / 3)
        $y = $view.Top  + [int]($vh / 2)

        # Two bands: where the characters will appear, and a control band below them that five
        # characters cannot reach. Without the control, a window covering Word changes every sampled
        # pixel and that reads as a huge successful edit.
        $bandH = [Math]::Min(100, [int]($vh / 6))
        $bandW = [Math]::Min(600, [int]($vw / 2))

        $page = New-Object WordLayout+RECT
        $page.Left = $x - 40; $page.Top = $y - $bandH
        $page.Right = $x + $bandW; $page.Bottom = $y + $bandH

        $control = New-Object WordLayout+RECT
        $control.Left = $page.Left; $control.Top = $page.Bottom + $bandH
        $control.Right = $page.Right; $control.Bottom = $control.Top + $bandH

        $shot = New-Object WordLayout+RECT
        $shot.Left = $page.Left; $shot.Top = $page.Top
        $shot.Right = $page.Right; $shot.Bottom = $control.Bottom

        $before = Get-RectShot $shot

        # No Focus() here: the click into the page is what makes Word foreground, and it does it as
        # user input, which Windows honours unconditionally.
        [WordLayout]::Click($x, $y)
        Start-Sleep -Milliseconds 600
        for ($k = 0; $k -lt 5; $k++) { [WordLayout]::Press($VK.X) }
        Start-Sleep -Milliseconds 700
        $after = Get-RectShot $shot

        $changed = Measure-Changed $before.Bitmap $after.Bitmap $before.Origin $page $null
        $moved   = Measure-Changed $before.Bitmap $after.Bitmap $before.Origin $control $null
        $before.Bitmap.Dispose(); $after.Bitmap.Dispose()

        $samples = [int](($page.Right - $page.Left) / 2) * [int](($page.Bottom - $page.Top) / 2)
        $localised = ($changed -ge 100) -and ($changed -le [int]($samples / 2)) -and ($moved -lt 100)

        if ($localised) {
            Write-Note "the document was modified ($changed of $samples samples on the line, $moved below it)"
            return $true
        }

        $covered = ($moved -ge 100) -or ($changed -gt [int]($samples / 2))
        $why = if ($covered) {
            "the whole view changed ($changed of $samples on the line, $moved on the control band) - something covered Word, not a keystroke"
        } else {
            "the keystrokes did not reach the document ($changed of $samples samples changed)"
        }
        Write-Note ("attempt {0}: {1}. Foreground is `"{2}`"" -f $try, $why, [WordLayout]::TitleOf([WordLayout]::GetForeground()))

        if ($covered) {
            $front = [WordLayout]::GetForeground()
            $cls = if ($front -ne [IntPtr]::Zero) { [WordLayout]::ClassOf($front) } else { '' }
            if ($front -ne [IntPtr]::Zero -and -not (@(Get-Frames) -contains $front) -and
                $cls -notin @('Progman', 'WorkerW') -and [WordLayout]::TitleOf($front) -ne '') {
                Write-Note ("minimising `"{0}`" ({1}) - it is sitting over Word" -f [WordLayout]::TitleOf($front), $cls)
                [WordLayout]::Show($front, 6)      # SW_MINIMIZE
                Start-Sleep -Milliseconds 900
            }
        }
    }
    return $false
}

# Save the document behind tab $index with Ctrl+S, having activated it by clicking its tab. Every
# fixture here is a real file that Word has already written once, so this never opens Save As.
function Save-Tab($index) {
    Set-WordForeground | Out-Null
    $spot = Get-Spot $index
    [WordLayout]::Click($spot.X, $spot.Y)
    Start-Sleep -Milliseconds 900
    $top = Get-TopStrip
    if ($top.Wwf) {
        $view = [WordLayout]::RectOf($top.Wwf.Hwnd)
        [WordLayout]::Click(($view.Left + [int](($view.Right - $view.Left) / 3)),
                            ($view.Top + [int](($view.Bottom - $view.Top) / 2)))
        Start-Sleep -Milliseconds 600
    }
    [WordLayout]::CtrlPress($VK.S)
    Start-Sleep -Seconds 2
}

# ---- fixtures ----------------------------------------------------------------------------------

$Settings = 'HKCU:\Software\WordTab'
$dir = Join-Path $env:TEMP 'wordtab-dot'

Write-Step 'Word must be closed to author the fixtures'
Close-Word

if (Test-Path $dir) { Remove-Item -Path $dir -Recurse -Force }
New-Item -ItemType Directory -Path $dir -Force | Out-Null

$alpha      = Join-Path $dir 'Alpha.docx'
$beta       = Join-Path $dir 'Beta.docx'
$downloaded = Join-Path $dir 'Downloaded.docx'

Write-Step 'Authoring the fixtures with Word itself, so the formats are real'
$word = New-Object -ComObject Word.Application
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    foreach ($pair in @(@($alpha, 'Alpha fixture for the dot suite.'), @($beta, 'Beta fixture for the dot suite.'))) {
        $doc = $word.Documents.Add()
        $word.Selection.TypeText($pair[1])
        $doc.SaveAs2($pair[0], 12)
        $doc.Close(0)
    }
} finally {
    try { $word.Quit(0) } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
Close-Word

Copy-Item $alpha $downloaded
Set-Content -Path $downloaded -Stream 'Zone.Identifier' -Value "[ZoneTransfer]`r`nZoneId=3"
Write-Note "fixtures in $dir"

# ---- 1. two clean documents ---------------------------------------------------------------------

Write-Step 'Two documents, both saved and untouched'
Set-LogMark

$alphaFrame = Open-Document $alpha $true
$betaFrame  = Open-Document $beta  $false

Assert ($alphaFrame -ne [IntPtr]::Zero) 'Alpha.docx opened a window'
Assert ($betaFrame  -ne [IntPtr]::Zero) 'Beta.docx opened a window'
Assert ((Get-FrameCount) -eq 2) "exactly two Word windows are open (got $(Get-FrameCount))"

# The tab order is the order documents joined, so Alpha is tab 0 and Beta is tab 1. Verified rather
# than assumed - every assertion below indexes on it.
Assert (Invoke-StripClick 'tab 0' { Get-Spot 0 }) 'tab 0 could be clicked'
Start-Sleep -Seconds 2
$front = [WordLayout]::TitleOf([WordLayout]::GetForeground())
Assert ($front -like 'Alpha*') "tab 0 is Alpha (the foreground title is |$front|)"

Start-Sleep -Seconds 2
$a0 = Get-DotState $alphaFrame
$b0 = Get-DotState $betaFrame
Assert ($null -eq $a0 -or -not $a0.On) "Alpha has no dot while it is unmodified (log says: $(if ($a0) { $a0.On } else { 'nothing yet' }))"
Assert ($null -eq $b0 -or -not $b0.On) "Beta has no dot while it is unmodified (log says: $(if ($b0) { $b0.On } else { 'nothing yet' }))"

$poll = @(Get-LogSince 'dot poll:')
Assert ($poll.Count -gt 0) 'the add-in is polling Word for the modified state at all'
if ($poll.Count -gt 0) { Write-Note $poll[$poll.Count - 1].Trim() }

Write-Step 'And both tabs draw an x, not a dot'
$cleanAlpha = Read-TabButton 0 'Alpha-clean'
Assert $cleanAlpha.IsCross 'an unmodified Alpha draws the x (ink on the diagonal)'
Assert (-not $cleanAlpha.IsDot) 'an unmodified Alpha does not draw a dot'

$cleanBeta = Read-TabButton 1 'Beta-clean'
Assert $cleanBeta.IsCross 'an unmodified Beta draws the x'

# ---- 2. one of them is modified ------------------------------------------------------------------

Write-Step 'Typing into Beta'
Assert (Set-Dirty 1) 'Beta was actually modified (proved by photographing the page)'

$b1 = Wait-Dot $betaFrame $true
Assert ($null -ne $b1 -and $b1.On) 'the add-in reports the dot ON for Beta'
if ($b1) { Assert ($b1.Document -like 'Beta*') "...and names the document it is about (|$($b1.Document)|)" }

$a1 = Get-DotState $alphaFrame
Assert ($null -eq $a1 -or -not $a1.On) 'Alpha is still reported clean - one document was modified, not both'

Write-Step 'The dot reaches the screen'
$dotBeta = Read-TabButton 1 'Beta-dirty'
Assert $dotBeta.IsDot 'Beta draws a dot: ink beside the centre, none out on the diagonal'
Assert (-not $dotBeta.IsCross) '...and the x is gone from it'

$stillAlpha = Read-TabButton 0 'Alpha-still-clean'
Assert $stillAlpha.IsCross 'Alpha still draws its x - one tab changed, not the row'

# ---- 3. hover brings the x back -------------------------------------------------------------------

Write-Step 'Hovering the modified tab'
$row = Get-Row
$spot = $row.Layout.Tabs[1]
$hx = [int](($spot.Left + $spot.Right) / 2)
$hy = [int](($spot.Top + $spot.Bottom) / 2)

Assert (Set-Pointer $hx $hy 'hovering Beta') 'the pointer arrived on Beta'
Start-Sleep -Milliseconds 700

$row = Get-Row
$hovered = Read-Button $row.Layout.Close[1] $row.Dpi 'Beta-hovered'
Assert $hovered.IsCross 'hovering a modified tab brings the x back'
Assert (-not $hovered.IsDot) '...and takes the dot away while the pointer is there'

Assert (Set-Pointer 4 4 'parking the pointer again') 'the pointer left the strip'
Start-Sleep -Milliseconds 700
$backAgain = Read-TabButton 1 'Beta-unhovered'
Assert $backAgain.IsDot 'the dot comes back when the pointer leaves'

# ---- 4. the dot is still the close button ---------------------------------------------------------

Write-Step 'Clicking the dot still closes the document'
Assert (Invoke-StripClick 'the dot on Beta' { $r = Get-Row; $c = $r.Layout.Close[1]
                                              [pscustomobject]@{ X = [int](($c.Left + $c.Right) / 2)
                                                                 Y = [int](($c.Top + $c.Bottom) / 2) } }) 'the dot could be clicked'

$prompt = Wait-Prompt $true 12
Assert ($null -ne $prompt) "Word asked about the unsaved changes, so the click did reach the close button ($(Format-Dialog $prompt))"

# Escape is Cancel. A test script may not discard work, and this suite created that work.
[WordLayout]::Press($VK.ESC)
Start-Sleep -Seconds 2
Assert ($null -eq (Wait-Prompt $false 10)) 'the prompt went away on Escape'
Assert ((Get-FrameCount) -eq 2) "cancelling left both documents open (got $(Get-FrameCount))"

$b2 = Get-DotState $betaFrame
Assert ($null -ne $b2 -and $b2.On) 'Beta still has its dot after the cancelled close - nothing was saved'

# ---- 4b. an armed press is not the same thing as hover -----------------------------------------------
#
# Found by an adversarial audit of this slice, and it was a real defect. A press on the close button
# takes capture and sets pressKind; sliding the pointer off the tab while still holding clears
# hotFrame, because only WM_CAPTURECHANGED clears pressKind. So the button is still armed with `hot`
# FALSE - and the first version of the swap drew a resting dot there, while an unmodified tab under
# the identical gesture kept its held chip. One gesture, drawn two ways depending on whether the
# document happened to be saved. The press is live either way: coming back onto the button and
# releasing still closes the document.

Write-Step 'Pressing the dot and sliding off the tab, still holding'

$row = Get-Row
$armedClose = $row.Layout.Close[1]
$otherTab   = $row.Layout.Tabs[0]
$fromX = [int](($armedClose.Left + $armedClose.Right) / 2)
$fromY = [int](($armedClose.Top + $armedClose.Bottom) / 2)
$toX   = [int](($otherTab.Left + $otherTab.Right) / 2)
$toY   = [int](($otherTab.Top + $otherTab.Bottom) / 2)

Assert ([WordLayout]::ClassOf([WordLayout]::WindowAt($fromX, $fromY)) -eq 'WordTabStrip') `
       'the strip really is under the point the press will start from'

# Press on Beta's close button, then carry the pointer onto tab 0 without letting go. This is not a
# tab drag - g_dragStrip is only set by a press on the tab itself, never on its close button.
[WordLayout]::DragHold($fromX, $fromY, $toX, $toY, 8, 40)
Start-Sleep -Milliseconds 700

$at = [WordLayout]::Cursor()
Assert (([Math]::Abs($at.X - $toX) -le 3) -and ([Math]::Abs($at.Y - $toY) -le 3)) `
       "the pointer really did travel off the tab while holding (at $($at.X),$($at.Y), wanted $toX,$toY)"

# Read the rect measured before the press: parking the pointer would release the gesture.
$armed = Read-Button $armedClose $row.Dpi 'Beta-armed-pointer-away'
Assert $armed.IsCross 'a close button that is still armed keeps its x with the pointer off the tab'
Assert (-not $armed.IsDot) '...rather than falling back to a resting dot'

[WordLayout]::DragRelease($toX, $toY)
Start-Sleep -Seconds 2
Assert ((Get-FrameCount) -eq 2) "releasing away from the button closed nothing (got $(Get-FrameCount))"
Assert ($null -eq (Get-WordDialog)) 'and asked nothing - the press was cancelled, not completed'

$b2b = Get-DotState $betaFrame
Assert ($null -ne $b2b -and $b2b.On) 'Beta still reports its dot after the cancelled press'

Assert (Set-Pointer 4 4 'parking the pointer after the gesture') 'the pointer left the strip'
Start-Sleep -Milliseconds 700
$afterArmed = Read-TabButton 1 'Beta-after-armed-press'
Assert $afterArmed.IsDot 'and the dot is back once nothing is pressed or hovered'

# ---- 5. saving puts it out --------------------------------------------------------------------------

Write-Step 'Saving Beta'
Save-Tab 1
$b3 = Wait-Dot $betaFrame $false
Assert ($null -ne $b3 -and -not $b3.On) 'the dot goes out when the document is saved'

$savedBeta = Read-TabButton 1 'Beta-saved'
Assert $savedBeta.IsCross 'and the x is back on the tab'

# ---- 6. a Protected View document does not poison the pass -------------------------------------------

Write-Step 'A Protected View document alongside them'
$pvFrame = Open-Document $downloaded $false
Assert ($pvFrame -ne [IntPtr]::Zero) 'the downloaded document opened a window'
Assert ((Get-FrameCount) -eq 3) "three windows now (got $(Get-FrameCount))"

$pvTitle = [WordLayout]::TitleOf($pvFrame)
Assert ($pvTitle -like '*Protected View*') "Word really did open it in Protected View (|$pvTitle|)"

# It is in neither Application.Windows nor Documents - measured - so it can never carry a dot. What
# has to hold is that an unmatched window does not stop the pass answering for everything else.
Write-Step 'The other tabs still work with it open'
Assert (Set-Dirty 0) 'Alpha was modified with a Protected View document in the row'
$a2 = Wait-Dot $alphaFrame $true
Assert ($null -ne $a2 -and $a2.On) 'Alpha still gets its dot - the unmatched window did not poison the pass'

$pvState = Get-DotState $pvFrame
Assert ($null -eq $pvState -or -not $pvState.On) 'the Protected View tab never claims to be modified'

Write-Step 'Saving Alpha again before anything is closed'
Save-Tab 0
$a3 = Wait-Dot $alphaFrame $false
Assert ($null -ne $a3 -and -not $a3.On) 'Alpha is clean again'

Close-Word

# ---- 7. TabDot=0 ------------------------------------------------------------------------------------

Write-Step 'TabDot=0 gives back exactly what the tab did before'

# Never New-Item -Force on a registry key: that RECREATES it and destroys every value in it, which
# was live in install.ps1 for the whole project and silently reset every switch on every re-install.
if (-not (Test-Path $Settings)) { New-Item -Path $Settings | Out-Null }
$restore = $null
try { $restore = (Get-ItemProperty -Path $Settings -Name 'TabDot' -ErrorAction Stop).TabDot } catch { $restore = $null }

try {
    Set-ItemProperty -Path $Settings -Name 'TabDot' -Value 0 -Type DWord
    Write-Note 'TabDot=0'

    Set-LogMark
    $alphaFrame = Open-Document $alpha $true
    $betaFrame  = Open-Document $beta  $false
    Assert ((Get-FrameCount) -eq 2) "two windows with TabDot=0 (got $(Get-FrameCount))"

    $start = @(Get-LogSince 'StripStart')
    Assert ($start.Count -gt 0 -and $start[$start.Count - 1] -like '*dot=off*') 'the add-in reports the dot switched off at startup'

    Assert (Set-Dirty 1) 'Beta was modified with the dot switched off'
    Start-Sleep -Seconds 4

    Assert (@(Get-LogSince 'dot |').Count -eq 0) 'nothing is reported about any dot'
    Assert (@(Get-LogSince 'dot poll:').Count -eq 0) 'and Word is not being polled at all - the switch takes the mechanism out, not just the drawing'

    $off = Read-TabButton 1 'Beta-dirty-TabDot0'
    Assert $off.IsCross 'a modified tab draws the x, exactly as it did before this slice'
    Assert (-not $off.IsDot) 'and no dot'

    Write-Step 'Saving Beta before the switch goes back'
    Save-Tab 1
    Start-Sleep -Seconds 2
} finally {
    if ($null -eq $restore) {
        Remove-ItemProperty -Path $Settings -Name 'TabDot' -ErrorAction SilentlyContinue
        Write-Note 'TabDot removed (it was not set before this run)'
    } else {
        Set-ItemProperty -Path $Settings -Name 'TabDot' -Value $restore -Type DWord
        Write-Note "TabDot put back to $restore"
    }
}

# ---- done ----------------------------------------------------------------------------------------

if (-not $KeepOpen) {
    Write-Step 'Closing Word'
    Close-Word
}

Write-Host ''
if ($script:Failures -eq 0) {
    Write-Host "$($script:Checks) checks, all passed." -ForegroundColor Green
} else {
    Write-Host "$($script:Checks) checks, $($script:Failures) FAILED." -ForegroundColor Red
    exit 1
}
