<#
.SYNOPSIS
  Measure the tab strip's appearance: the palette, the shapes, and the context menu's colours.

.DESCRIPTION
  Every other suite in this project asserts behaviour and treats pixels as evidence of behaviour -
  `check-tabs` counts *changed* pixels to prove a hover happened, and never asks what colour anything
  is. This one is about the look itself, so it is the first suite that reads absolute colours, and
  that needs saying because it is the reason it can fail for reasons the others cannot: a window over
  Word photographs as a perfectly plausible restyle. Every colour check here therefore has a control
  that says what the sample is being compared *against*, taken from the same photograph.

  **The claim it tests, in one sentence: the strip's palette is derived from the colour Word is
  actually painting, not from a table, and it follows Word when Word changes.**

  Four things are proved:

    - **The palette is Word's.** The well the tabs sit in is one measured step (26 levels) down from
      the ribbon immediately above it, and the active tab's card is the ribbon colour exactly. Both
      are read off one photograph, so no assumption about the theme is involved - the ribbon is
      whatever it is that day, and the strip has to agree with it.

    - **It follows a live theme change.** Windows' app theme is flipped from dark to light with Word
      *running*, and the strip has to come back light within a few seconds without Word restarting.
      This is the one thing the old registry-table palette could not do, and it is the reason the
      palette is sampled at all. The theme is put back afterwards, including on a crash - see the
      trap. Skip it with -NoThemeSwitch if flipping the desktop mid-run is not welcome.

    - **The cards are shaped.** A tab's top-left corner is well-coloured where a square corner would
      be card-coloured, and the arc between the two contains intermediate colours - which is what
      anti-aliasing is and is not something a staircase can fake. The hairline along the bottom of
      the strip is present beside the active card and absent underneath it, which is the difference
      between a row of buttons and a row of tabs.

    - **The context menu is drawn by us.** It is owner-drawn so that it is dark on a dark Word rather
      than the system's white; the menu window is photographed and its background compared against
      the strip's, with the system's own menu colour as the control it must NOT be.

  And one switch: `TabStyle=0` restores the flat rectangles, checked by restarting Word with it set
  and asserting the corner that was carved is square again.

  Word's documents on the dev rig are disposable test fixtures. Everything opened here is clean and
  unmodified, so nothing prompts to save.

.PARAMETER Documents
  How many to open. Three is enough for an active tab, an inactive one and a neighbour to use as a
  control.

.PARAMETER NoThemeSwitch
  Skip the live theme-change section, which flips Windows' app theme to light and back.

.PARAMETER KeepOpen
  Leave Word running afterwards, to look at it by hand.

.PARAMETER Screenshot
  Save photographs of every state asserted about.

.EXAMPLE
  pwsh -File tools\check-look.ps1
  pwsh -File tools\check-look.ps1 -KeepOpen -Screenshot
#>
[CmdletBinding()]
param(
    [int]$Documents = 3,
    [switch]$NoThemeSwitch,
    [switch]$KeepOpen,
    [switch]$Screenshot,
    [string]$ShotDir = $env:TEMP
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$Personalize = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize'
$script:ThemeWasLight = $null

# This suite changes a machine-wide setting. It may not leave it changed, whatever happens - so the
# restore is in the trap as well as at the end, and it is the first thing the trap does.
trap {
    Write-Host ''
    if ($null -ne $script:ThemeWasLight) {
        try {
            Set-ItemProperty -Path $Personalize -Name 'AppsUseLightTheme' -Value $script:ThemeWasLight -Type DWord
            Write-Host "    Windows app theme put back to AppsUseLightTheme=$($script:ThemeWasLight)" -ForegroundColor DarkGray
        } catch { }
    }
    Write-Host "UNHANDLED: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkRed
    break
}

$source = Get-Content -Raw -Path (Join-Path $PSScriptRoot 'WordLayout.cs')
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition $source -Language CSharp -ReferencedAssemblies @(
    'System.Runtime', 'System.Collections', 'System.Threading.Thread', 'netstandard'
)
[WordLayout]::MakeDpiAware() | Out-Null

# Confirmed input, one shared idea of what a Word dialog is, and the bounded waits. See the header of
# tools\WordTabHarness.ps1 for why these are not per-suite copies any more. Set-Pointer, which this
# suite invented after losing two hover assertions to a pointer that never moved, lives there now and
# is used by every suite.
. (Join-Path $PSScriptRoot 'WordTabHarness.ps1')

$StyleKey = 'HKCU:\Software\WordTab'
$VK_W     = 0x57

$script:Failures = 0
$script:Checks   = 0

function Write-Step($text) { Write-Host "==> $text" -ForegroundColor Cyan }
function Write-Note($text) { Write-Host "    $text" -ForegroundColor DarkGray }
function Assert($condition, $text) {
    $script:Checks++
    if ($condition) { Write-Host "    PASS  $text" -ForegroundColor Green }
    else { $script:Failures++; Write-Host "    FAIL  $text" -ForegroundColor Red }
}

function Get-WordPids { return Get-WordPidList }
function Get-WordPidCount { return Get-WordPidTally }

function Get-Frames { return Get-WordFrameList }
function Get-FrameCount { return Get-WordFrameTally }

function Get-Child($frame, $class) {
    foreach ($kid in [WordLayout]::Children($frame)) {
        if ($kid.Class -eq $class) { return $kid }
    }
    return $null
}

function Wait-For($predicate, $seconds = 25) {
    $deadline = (Get-Date).AddSeconds($seconds)
    while ((Get-Date) -lt $deadline) {
        if (& $predicate) { return $true }
        Start-Sleep -Milliseconds 400
    }
    return $false
}

function Get-TheFrame {
    $frames = @(Get-Frames)
    if ($frames.Count -eq 0) { throw 'No visible Word window.' }
    $top = [WordLayout]::GetForeground()
    if ($frames -contains $top) { return $top }
    return $frames[0]
}

function Get-StripOf($frame) {
    $strip = Get-Child $frame 'WordTabStrip'
    if (-not $strip) { throw ("Window 0x{0:X} has no WordTab strip." -f [int64]$frame) }
    return $strip.Hwnd
}

# Set-LogMark and Get-LogSince come from WordTabHarness.ps1 now - the copy that used to live here
# read from offset 0 after a roll, which is a wrong answer wearing the clothes of a right one.

# ---- photographs and the colours in them ------------------------------------------------------
#
# CopyFromScreen, not PrintWindow: this suite is about what the user can see, and PrintWindow asks
# the window to draw itself, which is a different claim. Borrowed from check-tabs.ps1:178.

function Get-RectShot($r) {
    $bmp = New-Object System.Drawing.Bitmap(($r.Right - $r.Left), ($r.Bottom - $r.Top))
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($r.Left, $r.Top, 0, 0, $bmp.Size)
    $g.Dispose()
    return [pscustomobject]@{ Bitmap = $bmp; Origin = $r }
}
function Get-StripShot($strip) { return Get-RectShot ([WordLayout]::RectOf($strip)) }

# One pixel of a photograph, addressed in SCREEN coordinates like every rectangle in this project.
function Get-Pixel($shot, $x, $y) {
    $px = $x - $shot.Origin.Left
    $py = $y - $shot.Origin.Top
    if ($px -lt 0 -or $py -lt 0 -or $px -ge $shot.Bitmap.Width -or $py -ge $shot.Bitmap.Height) { return $null }
    return $shot.Bitmap.GetPixel($px, $py)
}

# The most common colour along a horizontal run. A single pixel lands on a glyph or a border often
# enough to be worthless; the mode of a scan is the band.
function Get-ModeColour($shot, $y, $x0, $x1) {
    $tally = @{}
    for ($x = $x0; $x -lt $x1; $x += 2) {
        $c = Get-Pixel $shot $x $y
        if ($null -eq $c) { continue }
        $k = '{0},{1},{2}' -f $c.R, $c.G, $c.B
        if ($tally.ContainsKey($k)) { $tally[$k]++ } else { $tally[$k] = 1 }
    }
    if ($tally.Count -eq 0) { return $null }
    $best = $null; $bestN = -1; $total = 0
    foreach ($k in $tally.Keys) { $total += $tally[$k]; if ($tally[$k] -gt $bestN) { $bestN = $tally[$k]; $best = $k } }
    $parts = $best -split ','
    return [pscustomobject]@{
        R = [int]$parts[0]; G = [int]$parts[1]; B = [int]$parts[2]
        Share = [int][Math]::Round(100.0 * $bestN / $total)
        Text  = "RGB($best)"
    }
}

function Test-Near($a, $b, $tolerance) {
    if ($null -eq $a -or $null -eq $b) { return $false }
    return ([Math]::Abs($a.R - $b.R) -le $tolerance) -and
           ([Math]::Abs($a.G - $b.G) -le $tolerance) -and
           ([Math]::Abs($a.B - $b.B) -le $tolerance)
}
function Get-Luma($c) { return [int](($c.R * 54 + $c.G * 183 + $c.B * 19) / 256) }
function Show-Colour($label, $c) {
    if ($null -eq $c) { Write-Note ("{0,-34} (nothing there)" -f $label); return }
    $share = if ($c.PSObject.Properties.Name -contains 'Share') { "  {0}% of the scan" -f $c.Share } else { '' }
    Write-Note ("{0,-34} RGB({1},{2},{3}){4}" -f $label, $c.R, $c.G, $c.B, $share)
}

function Measure-Changed($a, $b, $origin, $rect) {
    $changed = 0
    for ($y = $rect.Top; $y -lt $rect.Bottom; $y += 2) {
        for ($x = $rect.Left; $x -lt $rect.Right; $x += 2) {
            $px = $x - $origin.Left
            $py = $y - $origin.Top
            if ($px -lt 0 -or $py -lt 0 -or $px -ge $a.Width -or $py -ge $a.Height) { continue }
            if ($a.GetPixel($px, $py).ToArgb() -ne $b.GetPixel($px, $py).ToArgb()) { $changed++ }
        }
    }
    return $changed
}

function Save-StripShot($strip, $name) {
    if (-not $Screenshot) { return }
    New-Item -ItemType Directory -Path $ShotDir -Force | Out-Null
    $shot = Get-StripShot $strip
    $path = Join-Path $ShotDir "wordtab-$name.png"
    $shot.Bitmap.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $shot.Bitmap.Dispose()
    Write-Note $path
}

# ---- Word ---------------------------------------------------------------------------------------

# This suite had no idea what a Word dialog was - nothing in it ever looked for one - so "Word would
# not close" was the only diagnosis it could give, whether the cause was a save prompt, a gallery
# left open, or Word genuinely wedged. Close-AllWord names which.
function Stop-Word {
    if ((Get-WordPidCount) -eq 0) { return }
    $end = Close-AllWord
    if (-not $end.Closed) {
        throw ("Word would not close ({0}: {1}) - close it by hand, saving or discarding as you like, then re-run." -f
               $end.Reason, (Format-WordWindow $end.Dialog))
    }
    Start-Sleep -Seconds 2
}

function Open-Documents($count) {
    $scratch = Join-Path $env:TEMP 'wordtab-check'
    New-Item -ItemType Directory -Path $scratch -Force | Out-Null
    $names = @('Quarterly report', 'Meeting notes', 'Draft proposal', 'Budget', 'Appendix', 'Notes')
    for ($i = 1; $i -le $count; $i++) {
        $path = Join-Path $scratch ("wordtab-look-{0}.rtf" -f $names[($i - 1) % $names.Count])
        "{\rtf1\ansi WordTab look check $i.\par}" | Set-Content -Path $path -Encoding Ascii
        Start-Process -FilePath 'winword.exe' -ArgumentList "`"$path`""
        if (-not (Wait-For { (Get-FrameCount) -ge $i } 60)) { throw "Document $i did not open." }
        Start-Sleep -Seconds $(if ($i -eq 1) { 8 } else { 3 })
    }
    Start-Sleep -Seconds 2
}

function Focus-Word {
    $frame = Get-TheFrame
    [WordLayout]::Focus($frame) | Out-Null
    Start-Sleep -Milliseconds 1200
    return $frame
}

# Where the tabs are, recomputed at every use. The strip moves whenever Word relays the window out,
# and a rectangle taken earlier has been measured 46px wrong one second later.
function Get-Layout($strip) { return [WordLayout]::Tabs($strip, (Get-FrameCount)) }

# MulDiv's rounding, matching Scaled() in strip.cpp and Sc() in WordLayout.cs.
function Get-Scaled($logical, $dpi) { return [int](($logical * $dpi + 48) / 96) }

# The well, sampled where the well actually is: the empty run to the RIGHT of the + button, at
# mid-height.
#
# Not a scan across the top of the strip, which is what the first version of this did. Two things
# live up there and both of them lie: the tabs themselves start three logical pixels down, and -
# measured for this slice - **Word's own `DropShadow` child window sits over the top twelve physical
# pixels of our strip**, so a scan there reads Word's shadow blended over our colours. It came back
# RGB(61,61,61) for a well that is RGB(15,15,15).
#
# **The fallback samples inside a tab, and the caller has to say which one.** When the row is long
# enough to fill the strip there is no bare well left to read - the only well-coloured pixels are
# inside the *inactive* tabs, which are deliberately not drawn as cards at all. So the fallback needs
# a tab that is known not to be the active one, and `inactiveIndex` is the caller promising it.
#
# It used to assume tab zero, and the assumption was true only by luck: on a window wide enough the
# fallback never ran, and on a narrow one it happened to be a section where tab zero was idle. The
# theme-change section pressed tab zero four sections earlier, so on a Word that remembered a narrow
# window it sampled the *active card* and reported the well as Word's own chrome - a step of zero
# levels where the code's step is twenty-six. Measured on the build before the scrolling row as well
# as after it, so this is the suite being fragile rather than the strip being wrong.
function Get-WellColour($shot, $stripRect, $layout, $inactiveIndex = 0) {
    $from = $layout.Plus.Right + 40
    $to   = [Math]::Min($stripRect.Right - 20, $from + 400)
    if ($to -le $from) {
        if ($layout.Tabs.Length -gt $inactiveIndex) {
            $tab = $layout.Tabs[$inactiveIndex]
            $from = $tab.Left + 30
            $to   = [Math]::Min($tab.Right - 30, $from + 100)
        }
        if ($to -le $from) { $from = $stripRect.Left + 20; $to = $from + 100 }
    }
    return Get-ModeColour $shot ([int](($stripRect.Top + $stripRect.Bottom) / 2)) $from $to
}

Write-Host ''
Write-Host 'WordTab - the look' -ForegroundColor White
Write-Host ''

# ---- a clean Word --------------------------------------------------------------------------------

if ((Get-WordPidCount) -gt 0) {
    Write-Step "Closing $(Get-WordPidCount) Word process(es) already running"
    Stop-Word
}

Write-Step "Opening $Documents documents"
Open-Documents $Documents
Assert ((Get-FrameCount) -eq $Documents) "$Documents Word windows ($(Get-FrameCount))"

$frame = Focus-Word
$strip = Get-StripOf $frame
Set-Pointer 4 4 'parking the pointer clear of the strip' | Out-Null
Start-Sleep -Milliseconds 900

$stripRect = [WordLayout]::RectOf($strip)
Write-Note ("strip at {0},{1} {2}x{3}" -f $stripRect.Left, $stripRect.Top,
            ($stripRect.Right - $stripRect.Left), ($stripRect.Bottom - $stripRect.Top))
Save-StripShot $strip 'look-rest'

# ---- 1. the palette is Word's own ----------------------------------------------------------------
#
# One photograph, wide enough to take in the ribbon above the strip as well as the strip itself, so
# the two colours being compared cannot come from different moments.

Write-Step "The palette against Word's own chrome"

$band = New-Object WordLayout+RECT
$band.Left   = $stripRect.Left
$band.Right  = $stripRect.Right
$band.Top    = $stripRect.Top - 12
$band.Bottom = $stripRect.Bottom
$wide = Get-RectShot $band

$layout = Get-Layout $strip
$activeIndex = $layout.Tabs.Length - 1     # the last document opened is the one in front
$active = $layout.Tabs[$activeIndex]

$scanFrom = $stripRect.Left + 40
$scanTo   = [Math]::Min($stripRect.Right - 40, $scanFrom + 700)

$ribbon = Get-ModeColour $wide ($stripRect.Top - 4) $scanFrom $scanTo
$well   = Get-WellColour $wide $stripRect $layout
$card   = Get-ModeColour $wide (($active.Top + $active.Bottom) / 2) `
                          ($active.Left + 30) ($active.Left + 90)

Show-Colour "Word's ribbon, just above us" $ribbon
Show-Colour 'the well the tabs sit in'     $well
Show-Colour 'the active tab card'          $card

Assert ($null -ne $ribbon -and $ribbon.Share -ge 60) `
    "the ribbon above the strip is a flat band ($(if ($ribbon) { "$($ribbon.Share)%" } else { 'unreadable' }) of the scan)"

# The whole derivation in one number: the well is 26 levels below the ribbon. Tolerance of 6 because
# a screen capture of a composited desktop is not obliged to be exact.
$stepR = $ribbon.R - $well.R
Write-Note ("ribbon -> well is a step of {0} levels (the code's step is 26)" -f $stepR)
Assert ([Math]::Abs($stepR - 26) -le 6) `
    "the well is one measured step below Word's ribbon ($stepR levels, expected 26)"

Assert (Test-Near $card $ribbon 6) `
    "the active card is Word's ribbon colour ($($card.Text) against $($ribbon.Text))"

Assert ((Get-Luma $well) -lt (Get-Luma $card)) `
    'the active card is lighter than the well it sits in, so it reads as the one in front'

# The control: an inactive tab must NOT be the card colour, or "active" means nothing.
$idle = $layout.Tabs[0]
$idleColour = Get-ModeColour $wide (($idle.Top + $idle.Bottom) / 2) ($idle.Left + 30) ($idle.Left + 90)
Show-Colour 'an inactive tab' $idleColour
Assert (Test-Near $idleColour $well 4) `
    "an inactive tab is the well, not a card ($($idleColour.Text) against $($well.Text))"
Assert (-not (Test-Near $idleColour $card 6)) `
    'and it is not the active card colour, so the row has exactly one tab in front'

$wide.Bitmap.Dispose()

# ---- 2. the cards are shaped ---------------------------------------------------------------------

Write-Step 'The shape of a card'

$shot = Get-StripShot $strip
$layout = Get-Layout $strip
$active = $layout.Tabs[$activeIndex]

# The card is drawn one logical pixel inside its rectangle, with a 6-logical-pixel radius. At this
# rig's 144 dpi that is a 9px arc starting 1px in. So: 2px in from the left edge and 2px down from
# the top is outside the arc and must still be the well; well inside the card is the card.
$corner = Get-Pixel $shot ($active.Left + 2) ($active.Top + 2)
$middle = Get-Pixel $shot (($active.Left + $active.Right) / 2) ($active.Top + 14)
Show-Colour 'the card top-left corner pixel' $corner
Show-Colour 'well inside the card'           $middle

Assert (Test-Near $corner $well 8) `
    'the corner of the active card is cut away - a square card would have card colour there'
Assert (Test-Near $middle $card 8) `
    'and the middle of it is the card colour, so it is a rounded card rather than a missing one'

# Anti-aliasing: inside the corner square there must be colours that are neither the well nor the
# card. A staircase has exactly two colours and no third.
#
# The square is the radius, taken from the same constant the add-in scales - the card is drawn one
# logical pixel inside its rectangle with a 6-logical-pixel radius, so at this rig's 144 dpi the arc
# lives in the 9x9 box starting one pixel in.
$dpi    = [WordLayout]::Dpi($strip)
$radius = Get-Scaled 6 $dpi
$inset  = Get-Scaled 1 $dpi
Write-Note "corner radius is $radius physical pixels at $dpi dpi"

$blend = 0
$cutAway = 0
$sampled = 0
for ($dx = 0; $dx -lt $radius; $dx++) {
    for ($dy = 0; $dy -lt $radius; $dy++) {
        $c = Get-Pixel $shot ($active.Left + $inset + $dx) ($active.Top + $dy)
        if ($null -eq $c) { continue }
        $sampled++
        if (Test-Near $c $well 6) { $cutAway++ }
        elseif (-not (Test-Near $c $card 6)) { $blend++ }
    }
}
Write-Note "$sampled pixels in the corner box: $cutAway well, $blend blended, the rest card"
Assert ($cutAway -ge 10) "a quadrant of the corner box is cut away ($cutAway well-coloured pixels)"
Assert ($blend -ge 3) "and the cut is anti-aliased, not a staircase ($blend blended pixels)"

# The hairline: present beside the active card, absent under it. That gap is what makes the active
# tab belong to the document below rather than sit on top of it.
$lineY = $stripRect.Bottom - 1
$underCard = Get-Pixel $shot (($active.Left + $active.Right) / 2) $lineY
$besideCard = Get-Pixel $shot ($active.Left - 12) $lineY
Show-Colour 'the bottom row under the card' $underCard
Show-Colour 'the bottom row beside it'      $besideCard
Assert (-not (Test-Near $underCard $besideCard 6)) `
    'the hairline along the bottom stops under the active card, so the card meets the document'
Assert (Test-Near $underCard $card 10) `
    'and what is there instead is the card itself, running all the way down'

$shot.Bitmap.Dispose()

# ---- 3. hover is visible, and it is local ---------------------------------------------------------
#
# The same shape as check-tabs.ps1:286 - the point of repeating it here is that the restyle changed
# what hover looks like, and a hover that bled onto a neighbour would be a new defect.

# Move the pointer, and prove it arrived before anything is concluded from a photograph.
#
# `SendInput`'s absolute move is not guaranteed to take: it is refused while another process holds a
# capture or has just taken the foreground, and it is simply overridden by a hand on the mouse. When
# it does not take, the cold and hot photographs are of the *same* pointer position, every hover
# assertion fails with "0 pixels", and that reads exactly like a hover that was never drawn.
#
# Measured, twice: a run asked for (258,540) and left the pointer at (1894,459), over the desktop.
# Two suites' worth of diagnosis went into a hover highlight that was working perfectly.
#
# Set-Pointer itself now lives in tools\WordTabHarness.ps1, unchanged, so the other ten suites get it
# too. This is where it was invented; the comment stays here because this is where the evidence is.

Write-Step 'Hover'

Assert (Set-Pointer 4 4 'parking the pointer') 'the pointer could be parked clear of the strip'
Start-Sleep -Milliseconds 900
$cold = Get-StripShot $strip

$layout = Get-Layout $strip
$target  = $layout.Tabs[0]
$control = $layout.Tabs[1]
$controlThird = New-Object WordLayout+RECT
$controlThird.Left   = $control.Left
$controlThird.Right  = $control.Left + [int](($control.Right - $control.Left) / 3)
$controlThird.Top    = $control.Top
$controlThird.Bottom = $control.Bottom

$hoverX = $target.Left + [int](($target.Right - $target.Left) / 3)
$hoverY = [int](($target.Top + $target.Bottom) / 2)
Assert (Set-Pointer $hoverX $hoverY 'hovering tab 0') 'the pointer could be put on tab 0'
Start-Sleep -Milliseconds 900

# Say where the pointer actually ended up and what is under it. "0 pixels changed" has two very
# different causes - the hover is not drawn, or the pointer is not there - and the assertion alone
# cannot tell them apart. Both have happened in this project; the second one wasted a diagnosis.
$at = [WordLayout]::Cursor()
$under = [WordLayout]::ClassOf([WordLayout]::WindowAt($at.X, $at.Y))
Write-Note ("aimed at ({0},{1}); pointer is at ({2},{3}) over `"{4}`"; tab 0 is ({5},{6} {7}x{8})" -f
            $hoverX, $hoverY, $at.X, $at.Y, $under,
            $target.Left, $target.Top, ($target.Right - $target.Left), ($target.Bottom - $target.Top))

$hot = Get-StripShot $strip

Assert (($cold.Origin.Left -eq $hot.Origin.Left) -and ($cold.Origin.Top -eq $hot.Origin.Top)) `
    'the strip held still between the two photographs'

$onTab     = Measure-Changed $cold.Bitmap $hot.Bitmap $cold.Origin $target
$elsewhere = Measure-Changed $cold.Bitmap $hot.Bitmap $cold.Origin $controlThird
Assert ($onTab -gt 0) "hovering tab 0 visibly changes it ($onTab pixels)"
Assert ($elsewhere -eq 0) "and changes nothing on tab 1 ($elsewhere pixels)"

$hoverColour = Get-ModeColour $hot ($target.Top + 14) ($target.Left + 30) ($target.Left + 90)
Show-Colour 'a hovered inactive tab' $hoverColour
Assert ((Get-Luma $hoverColour) -gt (Get-Luma $well)) `
    'a hovered tab lifts out of the well rather than sinking into it'
Assert ((Get-Luma $hoverColour) -lt (Get-Luma $card)) `
    'but not as far as the active card, so hover cannot be mistaken for selection'

if ($Screenshot) {
    New-Item -ItemType Directory -Path $ShotDir -Force | Out-Null
    $cold.Bitmap.Save((Join-Path $ShotDir 'wordtab-look-cold.png'), [System.Drawing.Imaging.ImageFormat]::Png)
    $hot.Bitmap.Save((Join-Path $ShotDir 'wordtab-look-hover.png'), [System.Drawing.Imaging.ImageFormat]::Png)
    Write-Note (Join-Path $ShotDir 'wordtab-look-hover.png')
}
$cold.Bitmap.Dispose(); $hot.Bitmap.Dispose()

# ---- 4. a carried tab is lifted --------------------------------------------------------------------

Write-Step 'The lift under a carried tab'

# Everything except the lift is held constant between the two photographs, the way
# check-reorder.ps1:383 does it: the tab is ALREADY selected and ALREADY hovered when the at-rest
# picture is taken. Without that the comparison measures the selection moving to the pressed tab -
# which it does, on the press - and reports 6242 changed pixels on a tab the lift never touched.
$layout = Get-Layout $strip
$pick = $layout.Tabs[0]
$width = $pick.Right - $pick.Left
$x = $pick.Left + [int]($width / 3)
$y = [int](($pick.Top + $pick.Bottom) / 2)

Invoke-ConfirmedClick -What 'selecting the tab to be carried' -Point {
    $l = Get-Layout (Get-StripOf (Get-TheFrame))
    $p = $l.Tabs[0]
    [pscustomobject]@{ X = $p.Left + [int](($p.Right - $p.Left) / 3); Y = [int](($p.Top + $p.Bottom) / 2) }
} | Out-Null
Start-Sleep -Seconds 2
[WordLayout]::Focus((Get-TheFrame)) | Out-Null
Start-Sleep -Milliseconds 1200
$strip = Get-StripOf (Get-TheFrame)

$layout = Get-Layout $strip
$pick = $layout.Tabs[0]
$far  = $layout.Tabs[$layout.Tabs.Length - 1]
$width = $pick.Right - $pick.Left
$x = $pick.Left + [int]($width / 3)
$y = [int](($pick.Top + $pick.Bottom) / 2)

Assert (Set-Pointer $x $y 'hovering the tab to be carried') 'the pointer could be put on the tab about to be carried'
Start-Sleep -Milliseconds 900
$rest = Get-StripShot $strip

# Confirmed before the button goes down, and not again until it is up: everything from DragHold to
# DragRelease is one gesture, and re-aiming inside it would measure a different one.
$onWhat = Get-ClassAt $x $y
Assert ($onWhat -eq 'WordTabStrip') "the pick-up starts on the strip, not on `"$onWhat`""
[WordLayout]::DragHold($x, $y, ($x + [int]($width / 3)), $y, 8, 70)
Start-Sleep -Milliseconds 700
$lifted = Get-StripShot $strip
Save-StripShot $strip 'look-carried'
[WordLayout]::DragRelease(($x + [int]($width / 3)), $y)
Start-Sleep -Milliseconds 800

$onCarried = Measure-Changed $rest.Bitmap $lifted.Bitmap $rest.Origin $pick
$farAway   = Measure-Changed $rest.Bitmap $lifted.Bitmap $rest.Origin $far
Assert ($onCarried -gt 0) "the carried tab is drawn out of its place ($onCarried pixels)"
Assert ($farAway -eq 0) "and the lift does not reach the far tab ($farAway pixels)"

$rest.Bitmap.Dispose(); $lifted.Bitmap.Dispose()

# ---- 5. the context menu is ours, and it is not the system's colours ---------------------------------

Write-Step 'The context menu'

$frame = Focus-Word
$strip = Get-StripOf $frame
$layout = Get-Layout $strip
$spot = $layout.Tabs[0]
# Expecting the strip: a right-click that lands on the document opens WORD's context menu, which is
# also a visible #32768, and this section then photographs Word's menu and calls it ours.
Assert (Invoke-ConfirmedClick -What 'right-clicking tab 0' -Button right -Point {
            $l = Get-Layout (Get-StripOf (Get-TheFrame))
            $p = $l.Tabs[0]
            [pscustomobject]@{ X = $p.Left + [int](($p.Right - $p.Left) / 3); Y = [int](($p.Top + $p.Bottom) / 2) }
        }) 'the right-click landed on the strip, so any menu that appears is ours'
Start-Sleep -Seconds 2

$menuWindow = [IntPtr]::Zero
foreach ($id in Get-WordPids) {
    $w = [WordLayout]::PopupMenuWindow($id)
    if ($w -ne [IntPtr]::Zero) { $menuWindow = $w }
}
Assert ($menuWindow -ne [IntPtr]::Zero) 'the menu is a real popup menu window, not a window of our own'

if ($menuWindow -ne [IntPtr]::Zero) {
    $menuRect = [WordLayout]::RectOf($menuWindow)
    Write-Note ("menu at {0},{1} {2}x{3}" -f $menuRect.Left, $menuRect.Top,
                ($menuRect.Right - $menuRect.Left), ($menuRect.Bottom - $menuRect.Top))
    $menuShot = Get-RectShot $menuRect
    if ($Screenshot) {
        New-Item -ItemType Directory -Path $ShotDir -Force | Out-Null
        $menuShot.Bitmap.Save((Join-Path $ShotDir 'wordtab-look-menu.png'), [System.Drawing.Imaging.ImageFormat]::Png)
        Write-Note (Join-Path $ShotDir 'wordtab-look-menu.png')
    }

    # Sampled a third of the way down and across the left gutter, which is background in every item.
    $menuBg = Get-ModeColour $menuShot ($menuRect.Top + [int](($menuRect.Bottom - $menuRect.Top) / 3)) `
                             ($menuRect.Left + 6) ($menuRect.Left + 26)
    Show-Colour 'the menu background' $menuBg

    # The control is the system's own menu colour - what the menu WAS before this slice, and what it
    # must no longer be when Word is dark.
    $sys = [System.Drawing.SystemColors]::Menu
    Show-Colour 'the system menu colour (the control)' $sys
    $wordIsDark = ((Get-Luma $card) -lt 128)
    Write-Note ("Word is $(if ($wordIsDark) { 'dark' } else { 'light' }) right now")

    Assert ($null -ne $menuBg -and $menuBg.Share -ge 60) `
        "the menu's background is a flat colour we drew ($(if ($menuBg) { "$($menuBg.Share)%" } else { 'unreadable' }) of the scan)"

    if ($wordIsDark) {
        Assert ((Get-Luma $menuBg) -lt 128) `
            "the menu is dark on a dark Word ($($menuBg.Text))"
        Assert (-not (Test-Near $menuBg $sys 24)) `
            'and it is not the system menu colour it used to be'
    } else {
        Assert ((Get-Luma $menuBg) -ge 128) "the menu is light on a light Word ($($menuBg.Text))"
    }

    # Owner-drawing must not have cost the menu its contents: this is what check-menu.ps1 reads.
    $items = @([WordLayout]::MenuItems($menuWindow, $frame))
    $texts = @($items | ForEach-Object { if ($_.Separator) { '-' } else { $_.Text } })
    Write-Note ("items: {0}" -f ($texts -join ', '))
    # The list itself comes from the harness - see Get-TabMenuItems for why it is not written out
    # here. This suite's claim is about the readback surviving owner-drawing, not about the contents.
    $wanted = @(Get-TabMenuItems)
    Assert ($items.Count -eq $wanted.Count) "the menu still has $($wanted.Count) entries ($($items.Count))"
    Assert (($texts -join '|') -eq ($wanted -join '|')) `
        'and every label still reads back through GetMenuString, owner-drawn or not'

    $menuShot.Bitmap.Dispose()
}

Invoke-ConfirmedKey -Vk 0x1B -What 'Escape to close the menu' | Out-Null
Start-Sleep -Seconds 1
Assert ((Get-WordMenu) -eq [IntPtr]::Zero) 'and it closes on Escape, leaving nothing behind'

# ---- 6. the palette follows Word while Word is running -----------------------------------------------
#
# The claim the sampled palette exists for. Office's own `UI Theme` value cannot be used to drive
# this - measured for this slice, Word rewrites it at startup from the roaming account setting - so
# the lever is Windows' app theme, which Word follows when its own setting is "use system setting".

if ($NoThemeSwitch) {
    Write-Step 'Live theme change - SKIPPED (-NoThemeSwitch)'
} else {
    Write-Step 'A theme change with Word running'

    $current = Get-ItemProperty -Path $Personalize -Name 'AppsUseLightTheme' -ErrorAction SilentlyContinue
    $script:ThemeWasLight = if ($current -and $current.PSObject.Properties.Name -contains 'AppsUseLightTheme') {
        [int]$current.AppsUseLightTheme
    } else { 1 }
    Write-Note "Windows app theme was AppsUseLightTheme=$($script:ThemeWasLight)"

    $frame = Focus-Word
    $strip = Get-StripOf $frame

    # Put the selection on the *last* tab before reading anything, so that tab zero is idle and the
    # well sampler's fallback has somewhere honest to read. Four sections ago this suite pressed tab
    # zero to photograph a carried tab, and on a Word that has remembered a narrow window that leaves
    # the only readable well-coloured pixels inside the one tab that is not well-coloured. Stated
    # here rather than left to whatever the previous section happened to do.
    $layout = Get-Layout $strip
    $lastTab = $layout.Tabs[$layout.Tabs.Length - 1]
    Invoke-ConfirmedClick -What 'selecting the last tab so tab zero is idle' -Point {
        $l = Get-Layout (Get-StripOf (Get-TheFrame))
        $t = $l.Tabs[$l.Tabs.Length - 1]
        [pscustomobject]@{ X = [int](($t.Left + $t.Right) / 2); Y = [int](($t.Top + $t.Bottom) / 2) }
    } | Out-Null
    Start-Sleep -Seconds 2
    $frame = Focus-Word
    $strip = Get-StripOf $frame

    Set-Pointer 4 4 'parking the pointer clear of the strip' | Out-Null
    Start-Sleep -Milliseconds 800

    $stripRect = [WordLayout]::RectOf($strip)
    $beforeShot = Get-StripShot $strip
    $before = Get-WellColour $beforeShot $stripRect (Get-Layout $strip) 0
    $beforeShot.Bitmap.Dispose()
    Show-Colour 'the well before the change' $before

    $flipTo = if ($script:ThemeWasLight -eq 0) { 1 } else { 0 }
    Set-LogMark
    Set-ItemProperty -Path $Personalize -Name 'AppsUseLightTheme' -Value $flipTo -Type DWord
    Write-Note "flipped to AppsUseLightTheme=$flipTo; Word repaints, then the janitor has two seconds to notice"
    Start-Sleep -Seconds 9

    [WordLayout]::Focus($frame) | Out-Null
    Start-Sleep -Milliseconds 1200
    Set-Pointer 4 4 'parking the pointer clear of the strip' | Out-Null
    Start-Sleep -Milliseconds 800

    $stripRect = [WordLayout]::RectOf($strip)
    $afterShot = Get-StripShot $strip
    $after  = Get-WellColour $afterShot $stripRect (Get-Layout $strip) 0
    $ribbon2 = $null
    $band2 = New-Object WordLayout+RECT
    $band2.Left = $stripRect.Left; $band2.Right = $stripRect.Right
    $band2.Top  = $stripRect.Top - 12; $band2.Bottom = $stripRect.Top
    $ribbonShot = Get-RectShot $band2
    $ribbon2 = Get-ModeColour $ribbonShot ($stripRect.Top - 4) ($stripRect.Left + 40) ($stripRect.Left + 600)

    Show-Colour 'the well after the change'   $after
    Show-Colour "Word's ribbon after it"      $ribbon2
    Save-StripShot $strip 'look-retheme'

    Assert (-not (Test-Near $before $after 20)) `
        "the strip changed colour without Word restarting ($($before.Text) -> $($after.Text))"

    $sign = if ($flipTo -eq 1) { 1 } else { -1 }
    Assert ((($after.R - $before.R) * $sign) -gt 40) `
        "and it changed in the right direction - $(if ($flipTo -eq 1) { 'lighter' } else { 'darker' }), with Windows"

    $step2 = $ribbon2.R - $after.R
    Write-Note ("ribbon -> well is now a step of {0} levels" -f $step2)
    Assert ([Math]::Abs($step2 - 26) -le 8) `
        "the derivation still holds against the new chrome ($step2 levels, expected 26)"

    Assert (@(Get-LogSince 'palette sampled').Count -ge 1) `
        'and the add-in says in its log that it re-sampled rather than guessed'

    $afterShot.Bitmap.Dispose(); $ribbonShot.Bitmap.Dispose()

    Set-ItemProperty -Path $Personalize -Name 'AppsUseLightTheme' -Value $script:ThemeWasLight -Type DWord
    Write-Note "Windows app theme put back to AppsUseLightTheme=$($script:ThemeWasLight)"
    Start-Sleep -Seconds 6
    $script:ThemeWasLight = $null
}

# ---- 7. the empty row is a designed state ------------------------------------------------------------
#
# Reached by closing the last document in the whole application. Closing one of three closes that
# window; only the last one leaves a window standing with nothing in it.

Write-Step 'The empty row'

Stop-Word
Open-Documents 1
$frame = Focus-Word
$strip = Get-StripOf $frame
Start-Sleep -Milliseconds 800

# Aimed at one named window. Ctrl+W closes whichever document is in front, so a Focus that quietly
# did not take does not lose the keystroke - it closes the wrong document.
Assert (Invoke-ConfirmedKeyOn -Hwnd $frame -Vk $VK_W -What 'Ctrl+W on the last document' -Ctrl) `
       'Ctrl+W went to the window it was aimed at'
Assert (Wait-For { $w = Get-Child (Get-TheFrame) '_WwF'
                   $w -and ([WordLayout]::FirstChild($w.Hwnd) -eq [IntPtr]::Zero) } 25) `
    'the last document closed and left its window standing'
Start-Sleep -Seconds 3

$frame = Focus-Word
$strip = Get-StripOf $frame
Set-Pointer 4 4 'parking the pointer clear of the strip' | Out-Null
Start-Sleep -Milliseconds 900
$stripRect = [WordLayout]::RectOf($strip)
$emptyShot = Get-StripShot $strip
Save-StripShot $strip 'look-empty'

$emptyLayout = [WordLayout]::Tabs($strip, 0)
Assert $emptyLayout.HasPlus 'the + is still there with nothing open'

# The message: pixels that are not the background, in the band to the right of the +. The control is
# the band further right again, which must be empty - otherwise "there is something there" would be
# satisfied by the strip being any colour at all.
$emptyWell = Get-ModeColour $emptyShot ($stripRect.Top + 6) `
                            ($emptyLayout.Plus.Right + 200) ($emptyLayout.Plus.Right + 400)
$inked = 0
$controlInked = 0
for ($x = $emptyLayout.Plus.Right + 10; $x -lt $emptyLayout.Plus.Right + 190; $x += 2) {
    for ($y = $stripRect.Top + 6; $y -lt $stripRect.Bottom - 4; $y += 2) {
        $c = Get-Pixel $emptyShot $x $y
        if ($null -ne $c -and -not (Test-Near $c $emptyWell 12)) { $inked++ }
    }
}
for ($x = $emptyLayout.Plus.Right + 420; $x -lt $emptyLayout.Plus.Right + 600; $x += 2) {
    for ($y = $stripRect.Top + 6; $y -lt $stripRect.Bottom - 4; $y += 2) {
        $c = Get-Pixel $emptyShot $x $y
        if ($null -ne $c -and -not (Test-Near $c $emptyWell 12)) { $controlInked++ }
    }
}
Write-Note "$inked inked pixels beside the +, $controlInked further along the empty row"
Assert ($inked -gt 40) "the empty row says something rather than being a bare band ($inked pixels)"
Assert ($controlInked -eq 0) "and the rest of the row really is empty ($controlInked pixels)"

# The message is painted, not clickable. This is the property that keeps check-startscreen.ps1's
# geometry proof true: the +'s hit rectangle was deliberately not widened to cover the label.
$before = Get-FrameCount
$messageX = $emptyLayout.Plus.Right + 100
$messageY = [int](($stripRect.Top + $stripRect.Bottom) / 2)
Assert ([WordLayout]::WindowAt($messageX, $messageY) -eq $strip) `
    'the message is inside the strip, so the next click is a fair test'
# Confirmed to land, deliberately: this is the negative test, and a click that never arrived produces
# exactly the same "nothing happened" as a label correctly ignoring one.
Assert (Invoke-ConfirmedClick -What 'clicking the empty row''s message' `
                             -Point { [pscustomobject]@{ X = $messageX; Y = $messageY } }) `
       'the click really did land on the strip, so "nothing happened" means the label ignored it'
Start-Sleep -Seconds 3
Assert (-not ([WordLayout]::FirstChild((Get-Child (Get-TheFrame) '_WwF').Hwnd) -ne [IntPtr]::Zero)) `
    'clicking the message does nothing - it is a label, not a button'
Assert ((Get-FrameCount) -eq $before) "and opens no window ($(Get-FrameCount))"

$emptyShot.Bitmap.Dispose()

# ---- 8. the switch ------------------------------------------------------------------------------------

Write-Step 'TabStyle=0 gives back the flat rectangles'

Stop-Word
Set-ItemProperty -Path $StyleKey -Name 'TabStyle' -Value 0 -Type DWord
try {
    Open-Documents 2
    $frame = Focus-Word
    $strip = Get-StripOf $frame
    Set-Pointer 4 4 'parking the pointer clear of the strip' | Out-Null
    Start-Sleep -Milliseconds 900

    $stripRect = [WordLayout]::RectOf($strip)
    $flatShot = Get-StripShot $strip
    Save-StripShot $strip 'look-tabstyle-off'

    $layout = Get-Layout $strip
    $flatActive = $layout.Tabs[$layout.Tabs.Length - 1]
    $flatWell = Get-WellColour $flatShot $stripRect $layout
    $flatCard = Get-ModeColour $flatShot (($flatActive.Top + $flatActive.Bottom) / 2) `
                               ($flatActive.Left + 30) ($flatActive.Left + 90)
    $flatCorner = Get-Pixel $flatShot ($flatActive.Left + 3) ($flatActive.Top + 3)
    Show-Colour 'the flat well'    $flatWell
    Show-Colour 'the flat card'    $flatCard
    Show-Colour 'its corner pixel' $flatCorner

    Assert (Test-Near $flatCorner $flatCard 14) `
        'with the switch off the corner is square again - the card colour goes right into it'

    # The symmetric measurement to the one above: a rounded card leaves a quadrant of well showing in
    # its corner box, a square one leaves none. Counting *blended* pixels would not do here - the
    # flat renderer draws a one-pixel border in the edge colour, which is neither the well nor the
    # card and would be counted as anti-aliasing that is not there.
    $flatCutAway = 0
    for ($dx = 0; $dx -lt $radius; $dx++) {
        for ($dy = 0; $dy -lt $radius; $dy++) {
            $c = Get-Pixel $flatShot ($flatActive.Left + $inset + $dx) ($flatActive.Top + $dy)
            if ($null -eq $c) { continue }
            if (Test-Near $c $flatWell 6) { $flatCutAway++ }
        }
    }
    Write-Note "$flatCutAway well-coloured pixels in the same corner box (rounded had $cutAway)"
    Assert ($flatCutAway -lt 3) "and nothing is cut out of it ($flatCutAway well-coloured pixels)"

    # The palette is still derived even with the renderer off - that part is a correction, not a
    # style - so the tabs are the right colours, just square.
    Assert ((Get-Luma $flatCard) -ne (Get-Luma $flatWell)) `
        'the tabs are still coloured from Word, just not shaped'
    $flatShot.Bitmap.Dispose()
}
finally {
    Remove-ItemProperty -Path $StyleKey -Name 'TabStyle' -ErrorAction SilentlyContinue
    Write-Note 'TabStyle removed - back to the default'
}

# ---- done -----------------------------------------------------------------------------------------

if (-not $KeepOpen) {
    Write-Step 'Closing Word'
    Stop-Word
}

Write-Host ''
if ($script:Failures -eq 0) {
    Write-Host "PASSED  $($script:Checks) checks, 0 failures" -ForegroundColor Green
} else {
    Write-Host "FAILED  $($script:Failures) of $($script:Checks) checks" -ForegroundColor Red
}
Write-Host ''
exit $(if ($script:Failures -eq 0) { 0 } else { 1 })
