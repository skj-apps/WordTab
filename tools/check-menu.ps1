<#
.SYNOPSIS
  Measure WordTab's tab context menu: what is on it, and what each command actually does to a
  document.

.DESCRIPTION
  Right-clicking a tab is the first place WordTab offers a *choice* rather than a single action, and
  three of the five choices reach the user's documents. So this checks the menu itself and then each
  command against something outside Word that cannot be faked:

    - the menu appears on a right-click, and its items are exactly the ones intended, in the
      intended groups. Read item by item out of the live HMENU (MN_GETHMENU), not photographed:
      a picture of a menu cannot tell you that "Close Others" is greyed
    - the tab the menu belongs to stays lit while it is open, with the pointer parked off the strip
      so the highlight cannot be the ordinary hover
    - a right-press that slides off the tab produces no menu, the same rule the close button follows
    - Save writes the document: asserted against the file's timestamp on disk, which is the only
      claim that matters and the only one Word cannot appear to satisfy without doing it
    - Save on a document that has never been saved *asks* instead of writing, and the document
      survives the question being cancelled
    - Close closes that document, and the menu item can be reached with the mouse
    - **a save prompt the user cancels abandons the rest of the batch.** This is the reason
      "Close Others" is a queue rather than a loop: cancel has to mean cancel, not "ask me again
      about the next five"
    - Close Others leaves exactly the tab it was invoked on, and is greyed when there is only one
    - Close All closes the lot

  Word is started on scratch .rtf files written by this script. Word on the dev rig is a test rig:
  everything here is disposable, and this script deliberately *does* provoke the save prompt - with
  a bounded wait and an Escape, never an unbounded one.

  **Every click recomputes where it is aiming immediately before it clicks.** The strip moves; a
  rectangle measured three clicks ago points into the document. See check-tabs.ps1.

.PARAMETER Documents
  How many documents to open. Four gives enough tabs for Close Others to be interesting.

.PARAMETER KeepOpen
  Leave Word running afterwards. The last check closes every document, so this only matters when
  something failed earlier.

.EXAMPLE
  pwsh -File tools\check-menu.ps1
  pwsh -File tools\check-menu.ps1 -Screenshot
#>
[CmdletBinding()]
param(
    [int]$Documents = 4,
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

# This script drives Word for several minutes and can fail anywhere in it. Without this the host
# prints one line with no line number and the run has to be repeated to find out where.
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

function Get-WordPids { @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }) }

function Get-Frames {
    $found = @()
    foreach ($id in Get-WordPids) { $found += [WordLayout]::Frames($id) }
    return @($found)
}

# Always ask for the count through this. A PowerShell function that returns an empty array returns
# *nothing*, so `(Get-Frames).Count` is a hard error - "the property Count cannot be found on this
# object" - at exactly the two moments this script cares about most: the last document closing, and
# the instant during a batch close when Word momentarily has no window that qualifies. Re-wrapping
# at the call site is what makes 0 come back as 0.
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

function Get-TopStrip {
    $frames = @(Get-Frames)
    if ($frames.Count -eq 0) { throw 'No Word windows.' }
    $top = [WordLayout]::GetForeground()
    if (-not ($frames -contains $top)) { $top = $frames[0] }
    $parts = Get-Parts $top
    if (-not $parts.Strip) { throw 'The foreground Word window has no WordTab strip.' }
    return $parts
}

# Where to aim, measured now - see the note in the description.
#   'label' - the left third of a tab, clear of its close button
#   'empty' - the strip to the right of the new-document button, which belongs to no tab
function Get-Spot($kind, $index) {
    $top = Get-TopStrip
    $count = (Get-FrameCount)
    $layout = [WordLayout]::Tabs($top.Strip.Hwnd, $count)
    $stripRect = [WordLayout]::RectOf($top.Strip.Hwnd)

    if ($kind -eq 'empty') {
        $x = [int](($layout.Plus.Right + $stripRect.Right) / 2)
        if ($x -ge $stripRect.Right - 2) { $x = $stripRect.Right - 6 }
        return [pscustomobject]@{
            X = $x; Y = [int](($stripRect.Top + $stripRect.Bottom) / 2)
            Rect = $stripRect; Strip = $top.Strip.Hwnd; Count = $count
        }
    }

    if ($index -ge $layout.Tabs.Length) { throw "Tab $index does not exist ($($layout.Tabs.Length) tabs)." }
    $t = $layout.Tabs[$index]
    return [pscustomobject]@{
        X = $t.Left + [int](($t.Right - $t.Left) / 3)
        Y = [int](($t.Top + $t.Bottom) / 2)
        Rect = $t; Strip = $top.Strip.Hwnd; Count = $count
    }
}

function Format-Rect($r) { "({0},{1} {2}x{3})" -f $r.Left, $r.Top, ($r.Right - $r.Left), ($r.Bottom - $r.Top) }

# The add-in's own account of what it did. Used for one thing only, and it is worth the coupling:
# "the batch stopped" and "the batch stopped *because the user declined*" look identical from
# outside, and the first version of this code stopped for the wrong reason - a second after posting
# WM_CLOSE, before the prompt was even on screen - while every outside-visible assertion passed.
$LogPath = Join-Path $env:LOCALAPPDATA 'WordTab\wordtab.log'
function Get-LogTail($pattern, $count = 1) {
    if (-not (Test-Path $LogPath)) { return @() }
    return @(Get-Content -Path $LogPath -Tail 600 |
             Select-String -Pattern $pattern -SimpleMatch |
             Select-Object -Last $count |
             ForEach-Object { $_.Line.Trim() })
}

function Get-MenuWindow {
    foreach ($id in Get-WordPids) {
        $w = [WordLayout]::PopupMenuWindow($id)
        if ($w -ne [IntPtr]::Zero) { return $w }
    }
    return [IntPtr]::Zero
}

function Wait-Menu($present, $seconds = 6) {
    $deadline = (Get-Date).AddSeconds($seconds)
    while ((Get-Date) -lt $deadline) {
        $up = (Get-MenuWindow) -ne [IntPtr]::Zero
        if ($up -eq $present) { return $true }
        Start-Sleep -Milliseconds 200
    }
    return $false
}

# Anything Word has put up that is not a document window and not the menu: a save prompt, Save As,
# a compatibility warning. This is how "Word is asking the user something" is detected from outside.
function Get-WordDialog {
    foreach ($id in Get-WordPids) {
        foreach ($w in [WordLayout]::TopLevel($id)) {
            if (-not $w.Visible) { continue }
            if ($w.Class -eq 'OpusApp' -or $w.Class -eq '#32768') { continue }
            if (($w.Right - $w.Left) -lt 200 -or ($w.Bottom - $w.Top) -lt 100) { continue }
            return $w
        }
    }
    return $null
}

function Wait-Dialog($seconds = 10) {
    $deadline = (Get-Date).AddSeconds($seconds)
    while ((Get-Date) -lt $deadline) {
        $d = Get-WordDialog
        if ($d) { return $d }
        Start-Sleep -Milliseconds 250
    }
    return $null
}

# Right-click a tab and return what came up. The menu's items are read through its owner frame,
# which is what GetMenuItemRect needs to answer with screen rectangles.
function Open-TabMenu($kind, $index) {
    $spot = Get-Spot $kind $index
    $owner = (Get-TopStrip).Frame
    [WordLayout]::RightClick($spot.X, $spot.Y)
    $up = Wait-Menu $true
    if (-not $up) { return [pscustomobject]@{ Window = [IntPtr]::Zero; Items = @(); Spot = $spot } }
    $window = Get-MenuWindow
    return [pscustomobject]@{
        Window = $window
        Items  = @([WordLayout]::MenuItems($window, $owner))
        Spot   = $spot
    }
}

$VK = @{ S = 0x53; C = 0x43; O = 0x4F; A = 0x41; N = 0x4E; ESC = 0x1B; X = 0x58 }

function Close-Menu {
    if ((Get-MenuWindow) -ne [IntPtr]::Zero) {
        [WordLayout]::Press($VK.ESC)
        Wait-Menu $false | Out-Null
    }
}

function Wait-Frames($expected, $seconds = 25) {
    $deadline = (Get-Date).AddSeconds($seconds)
    while ((Get-Date) -lt $deadline) {
        if ((Get-FrameCount) -eq $expected) {
            Start-Sleep -Milliseconds 1200
            if ((Get-FrameCount) -eq $expected) { return $true }
        }
        Start-Sleep -Milliseconds 400
    }
    return $false
}

# Put the caret in the document and type, so the document is genuinely modified. A keystroke sent to
# a window that was activated programmatically goes nowhere in Word - one real click into the page
# fixes it, which is why this clicks first.
function Set-Dirty($index) {
    $spot = Get-Spot 'label' $index
    [WordLayout]::Click($spot.X, $spot.Y)
    Start-Sleep -Milliseconds 900

    $top = Get-TopStrip
    $strip = [WordLayout]::RectOf($top.Strip.Hwnd)
    $x = $strip.Left + [int](($strip.Right - $strip.Left) / 3)
    $y = $strip.Bottom + 300
    [WordLayout]::Click($x, $y)
    Start-Sleep -Milliseconds 500
    [WordLayout]::Press($VK.X)
    Start-Sleep -Milliseconds 500
    return [WordLayout]::GetForeground()
}

function Get-StripShot($strip) {
    $r = [WordLayout]::RectOf($strip)
    $bmp = New-Object System.Drawing.Bitmap(($r.Right - $r.Left), ($r.Bottom - $r.Top))
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($r.Left, $r.Top, 0, 0, $bmp.Size)
    $g.Dispose()
    return [pscustomobject]@{ Bitmap = $bmp; Origin = $r }
}

# How many sampled pixels differ between two shots inside a rectangle, ignoring anything inside
# `exclude`. The exclusion is not fastidiousness: an open menu overlaps the bottom of the strip it
# was raised from, and its drop shadow reaches further still, so a naive comparison reports the menu
# disappearing as the strip changing. That is what it reported the first time this ran.
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

# ---- get Word up with N documents ----------------------------------------------------------------

$running = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue)
if ($running.Count -gt 0) {
    Write-Step "Closing Word first: $(Get-FrameCount) window(s) in $($running.Count) process(es)"

    # One window at a time, not CloseMainWindow. A Word process holding a stack has N top-level
    # windows and CloseMainWindow closes exactly one of them, so a previous run's four documents go
    # to three and the wait times out looking like Word refusing to close. Measured, twice.
    for ($guard = 0; $guard -lt 16; $guard++) {
        $open = @(Get-Frames)
        if ($open.Count -eq 0) { break }
        [WordLayout]::Close($open[0])
        Start-Sleep -Seconds 2
        # A leftover document with unsaved changes. Deliberately not answered here: this script
        # discards nothing it did not create, and "the cleanup step threw away my work" is not a
        # thing a test script may ever do. Cancel the question so Word is left exactly as it was,
        # and say what to do.
        if (Get-WordDialog) {
            [WordLayout]::Press(0x1B)
            Start-Sleep -Seconds 2
            throw 'Word is asking about unsaved changes left over from an earlier run. Answer it by hand (on this rig the check documents are disposable, so Don''t Save is fine) and re-run.'
        }
    }
    foreach ($process in @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue)) { $process.CloseMainWindow() | Out-Null }
    $deadline = (Get-Date).AddSeconds(20)
    while ((Get-Date) -lt $deadline -and (Get-Process -Name WINWORD -ErrorAction SilentlyContinue)) {
        Start-Sleep -Milliseconds 500
    }
    $left = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue)
    if ($left.Count -gt 0) { throw "Word would not close ($($left.Count) left) - close it by hand, then re-run." }
    Start-Sleep -Seconds 2
}

$scratch = Join-Path $env:TEMP 'wordtab-check'
New-Item -ItemType Directory -Path $scratch -Force | Out-Null

$paths = @()
for ($i = 1; $i -le $Documents; $i++) {
    $path = Join-Path $scratch "wordtab-menu-$i.rtf"
    "{\rtf1\ansi WordTab menu check - document $i.\par}" | Set-Content -Path $path -Encoding Ascii
    $paths += $path
    Start-Process -FilePath 'winword.exe' -ArgumentList "`"$path`""
    Start-Sleep -Seconds $(if ($i -eq 1) { 14 } else { 7 })
}

$deadline = (Get-Date).AddSeconds(45)
while ((Get-Date) -lt $deadline -and (Get-FrameCount) -lt $Documents) { Start-Sleep -Milliseconds 500 }

$frames = @(Get-Frames)
Write-Note "$($frames.Count) visible Word frame(s)"
if ($frames.Count -lt 3) { throw "Need at least 3 Word frames to check the menu; found $($frames.Count)." }

$settle = (Get-Date).AddSeconds(15)
while ((Get-Date) -lt $settle) {
    $parts = @(Get-Frames | ForEach-Object { Get-Parts $_ })
    $ready = @($parts | Where-Object { $_.Strip -and $_.Wwf -and $_.Strip.Bottom -eq $_.Wwf.Top })
    if ($ready.Count -eq $parts.Count) { break }
    Start-Sleep -Milliseconds 500
}

# ---- the menu itself -------------------------------------------------------------------------------

Write-Step 'Right-clicking a tab'
$menu = Open-TabMenu 'label' 1
Assert ($menu.Window -ne [IntPtr]::Zero) 'a context menu appeared'

$expected = @('&Save', '-', '&Close', 'Close &Others', 'Close &All', '-', '&New Document')
$actual = @($menu.Items | ForEach-Object { $_.Text })
foreach ($item in $menu.Items) {
    Write-Note ("[{0}] id={1} enabled={2} {3} `"{4}`"" -f $item.Index, $item.Id, $item.Enabled,
                $(if ($item.HasRect) { Format-Rect $item.Rect } else { 'no rect' }), $item.Text)
}
Assert ($actual.Count -eq $expected.Count) "the menu has $($expected.Count) entries ($($actual.Count))"
Assert (($actual -join '|') -eq ($expected -join '|')) "the entries are exactly: $($expected -join ', ')"

$others = @($menu.Items | Where-Object { $_.Text -eq 'Close &Others' })[0]
Assert ($others -and $others.Enabled) 'Close Others is available with more than one tab'

# The tab the menu belongs to stays lit. The pointer is parked off the strip first, so what is being
# measured is the menu's own highlight and not the hover that any pointer over a tab would produce.
Write-Step 'The tab the menu belongs to'
[WordLayout]::MouseTo(4, 4)
Start-Sleep -Milliseconds 700
$top = Get-TopStrip

# A menu opens at the pointer and grows down and to the right, so it covers the bottom of the strip
# below where it was raised - and draws a shadow past that again. None of that can be compared: the
# thing on top of it is the very thing being dismissed. The menu's rectangle, generously inflated,
# is excluded from the target measurement, and the control tab is chosen to the *left* of the click
# where the menu cannot reach at all - which is then asserted rather than assumed.
$menuRect = [WordLayout]::RectOf($menu.Window)
$blocked = [pscustomobject]@{
    Left = $menuRect.Left - 40; Top = $menuRect.Top - 40
    Right = $menuRect.Right + 40; Bottom = $menuRect.Bottom + 40
}
Write-Note ("the menu is at {0}; ignoring {1} of the strip" -f (Format-Rect $menuRect), (Format-Rect $blocked))

$lit = Get-StripShot $top.Strip.Hwnd
$targetRect = $menu.Spot.Rect
$controlRect = (Get-Spot 'label' 0).Rect

Close-Menu
Start-Sleep -Milliseconds 700
$cold = Get-StripShot $top.Strip.Hwnd

$still = ($lit.Origin.Left -eq $cold.Origin.Left) -and ($lit.Origin.Top -eq $cold.Origin.Top)
Assert $still 'the strip held still between the two photographs'
Assert ($controlRect.Right -le $blocked.Left) 'the control tab is clear of the menu and its shadow'

$onTab = Measure-Changed $lit.Bitmap $cold.Bitmap $lit.Origin $targetRect $blocked
$elsewhere = Measure-Changed $lit.Bitmap $cold.Bitmap $lit.Origin $controlRect $blocked
Write-Note "pixels that changed when the menu closed - on its own tab: $onTab   on another tab: $elsewhere"
Assert ($onTab -gt 0) "tab 1 is lit while its menu is open ($onTab pixels), pointer nowhere near it"
Assert ($elsewhere -eq 0) "no other tab is lit by it ($elsewhere pixels)"

if ($Screenshot) {
    New-Item -ItemType Directory -Path $ShotDir -Force | Out-Null
    $lit.Bitmap.Save((Join-Path $ShotDir 'wordtab-menu-lit.png'), [System.Drawing.Imaging.ImageFormat]::Png)
    $cold.Bitmap.Save((Join-Path $ShotDir 'wordtab-menu-cold.png'), [System.Drawing.Imaging.ImageFormat]::Png)
    Write-Note (Join-Path $ShotDir 'wordtab-menu-lit.png')
}
$lit.Bitmap.Dispose(); $cold.Bitmap.Dispose()

Assert ((Get-MenuWindow) -eq [IntPtr]::Zero) 'Escape dismissed the menu'
Assert ((Get-FrameCount) -eq $frames.Count) "dismissing it closed nothing ($((Get-FrameCount)) windows)"

# ---- a right-press that changes its mind -------------------------------------------------------------

Write-Step 'Right-pressing a tab and sliding off it'
$onTabSpot = Get-Spot 'label' 1
$offSpot = Get-Spot 'empty' 0
[WordLayout]::RightPressAndSlideOff($onTabSpot.X, $onTabSpot.Y, $offSpot.X, $offSpot.Y)
Start-Sleep -Milliseconds 600
Assert ((Get-MenuWindow) -eq [IntPtr]::Zero) 'no menu - the release did not land on the tab it started on'
Close-Menu

# ---- the empty part of the strip ------------------------------------------------------------------

Write-Step 'Right-clicking the strip where there is no tab'
$menu = Open-TabMenu 'empty' 0
Assert ($menu.Window -ne [IntPtr]::Zero) 'a menu appeared there too, rather than a dead area'
$actual = @($menu.Items | ForEach-Object { $_.Text })
Write-Note ("entries: {0}" -f ($actual -join ', '))
Assert (($actual -join '|') -eq '&New Document') 'it offers only New Document - nothing that needs a tab'
Close-Menu

# ---- Save writes the file ----------------------------------------------------------------------------

Write-Step 'Save, checked against the file on disk'
$before = (Get-Item $paths[0]).LastWriteTime
Write-Note "$($paths[0])  last written $($before.ToString('HH:mm:ss.fff'))"

Set-Dirty 0 | Out-Null
Start-Sleep -Milliseconds 500

$menu = Open-TabMenu 'label' 0
if ($menu.Window -ne [IntPtr]::Zero) {
    [WordLayout]::Press($VK.S)
    Wait-Menu $false | Out-Null
} else {
    Assert $false 'the menu opened on tab 0'
}

$saved = $false
$deadline = (Get-Date).AddSeconds(15)
while ((Get-Date) -lt $deadline) {
    if ((Get-Item $paths[0]).LastWriteTime -gt $before) { $saved = $true; break }
    Start-Sleep -Milliseconds 400
}
$after = (Get-Item $paths[0]).LastWriteTime
Write-Note "now last written $($after.ToString('HH:mm:ss.fff'))"
Assert $saved 'Save wrote the document to disk'
Assert ((Get-WordDialog) -eq $null) 'it saved without asking anything - the document had a name already'

# ---- Save on a document that has never been saved -------------------------------------------------------

Write-Step 'Save on a document that has never been saved'
$count = (Get-FrameCount)
$plusSpot = Get-Spot 'label' 0
$top = Get-TopStrip
$layout = [WordLayout]::Tabs($top.Strip.Hwnd, $count)
$p = [WordLayout]::Center($layout.Plus)
[WordLayout]::Click($p.X, $p.Y)
$arrived = Wait-Frames ($count + 1) 25
Assert $arrived "a new blank document arrived as a tab ($((Get-FrameCount)) windows)"

if ($arrived) {
    $newIndex = (Get-FrameCount) - 1
    $menu = Open-TabMenu 'label' $newIndex
    if ($menu.Window -ne [IntPtr]::Zero) {
        [WordLayout]::Press($VK.S)
        Wait-Menu $false | Out-Null
    }

    # "Word asked" has two shapes and either is right: the classic Save As dialog as its own
    # top-level window, or the Save As page of Backstage, which is a child of the frame. Asserting
    # only "some window appeared" would also be satisfied by a tooltip, so both are named and the
    # one that happened is recorded.
    $asked = $null
    $backstage = $false
    $deadline = (Get-Date).AddSeconds(15)
    while ((Get-Date) -lt $deadline) {
        $asked = Get-WordDialog
        $active = [WordLayout]::GetForeground()
        $backstage = ($active -ne [IntPtr]::Zero) -and [WordLayout]::BackstageOpen($active)
        if ($asked -or $backstage) { break }
        Start-Sleep -Milliseconds 300
    }

    if ($asked) { Write-Note ("Word put up a top-level window: class `"{0}`" {1}" -f $asked.Class, (Format-Rect $asked)) }
    if ($backstage) { Write-Note 'Word opened the Save As page of Backstage' }
    Assert ($asked -or $backstage) 'Word asked where to save it rather than writing a file of its own choosing'

    [WordLayout]::Press($VK.ESC)
    Start-Sleep -Seconds 3
    if (Get-WordDialog) { [WordLayout]::Press($VK.ESC); Start-Sleep -Seconds 3 }

    $stillAsking = (Get-WordDialog) -ne $null
    Assert (-not $stillAsking) 'Escape closed the question'
    Assert ((Get-FrameCount) -eq ($count + 1)) "the document survived being asked about ($((Get-FrameCount)) windows)"

    # Backstage hides every child of the frame, the strip included. Leaving it is when the strip has
    # to come back - and if it did not, every check after this one would be clicking into a document.
    Start-Sleep -Seconds 2
    $parts = Get-Parts ([WordLayout]::GetForeground())
    Assert ($parts.Strip -and $parts.Wwf -and $parts.Strip.Bottom -eq $parts.Wwf.Top) `
        'the strip is back in place after the question was dismissed'
}

# ---- Close, clicked with the mouse -------------------------------------------------------------------

Write-Step 'Close, from the menu'
$count = (Get-FrameCount)
$victim = $count - 1
$menu = Open-TabMenu 'label' $victim
$item = @($menu.Items | Where-Object { $_.Text -eq '&Close' })[0]

if ($item -and $item.HasRect) {
    $centre = [WordLayout]::Center($item.Rect)
    Write-Note ("clicking the Close item at {0}" -f (Format-Rect $item.Rect))
    [WordLayout]::Click($centre.X, $centre.Y)
} else {
    Write-Note 'the item has no rectangle to click - using its access key instead'
    [WordLayout]::Press($VK.C)
}
Wait-Menu $false | Out-Null

$closed = Wait-Frames ($count - 1) 25
Assert $closed "Close closed that document ($((Get-FrameCount)) windows, expected $($count - 1))"

# ---- a cancelled save prompt abandons the batch ------------------------------------------------------
#
# The reason the batch is a queue. Tab 0 is dirtied and left in the background, and the batch is
# started from the last tab: the queue is then every other tab in order, so tab 0 is the *first*
# close and the prompt arrives before anything has gone. Cancel it, and nothing at all should close.

Write-Step 'Cancelling a save prompt stops the rest of the batch'
$count = (Get-FrameCount)
if ($count -ge 3) {
    Set-Dirty 0 | Out-Null

    $keep = $count - 1
    $spot = Get-Spot 'label' $keep
    [WordLayout]::Click($spot.X, $spot.Y)
    Start-Sleep -Milliseconds 1200
    Write-Note ("keeping `"{0}`", {1} tabs in total" -f [WordLayout]::TitleOf([WordLayout]::GetForeground()), $count)

    $menu = Open-TabMenu 'label' $keep
    if ($menu.Window -ne [IntPtr]::Zero) {
        [WordLayout]::Press($VK.O)
        Wait-Menu $false | Out-Null
    }

    $prompt = Wait-Dialog 15
    if ($prompt) { Write-Note ("Word asked: `"{0}`" ({1})" -f $prompt.Title, $prompt.Class) }
    Assert ($prompt -ne $null) 'the dirty document raised Word''s save prompt before anything closed'

    if ($prompt) {
        $ownerDisabled = @(Get-Frames | Where-Object { -not [WordLayout]::Enabled($_) }).Count
        Write-Note "$ownerDisabled frame(s) disabled while the prompt is up"

        [WordLayout]::Press($VK.ESC)          # Cancel
        Start-Sleep -Seconds 10               # the grace after the question goes, and then some

        Assert ((Get-WordDialog) -eq $null) 'the prompt went away'
        Assert ((Get-FrameCount) -eq $count) `
            "cancelling stopped the whole batch - all $count documents are still open ($((Get-FrameCount)))"

        # ...and stopped for the right reason. Every assertion above is also satisfied by a batch
        # that gave up before Word had even asked - which is exactly what the first version of this
        # code did, one second after posting WM_CLOSE, while passing all of them.
        $reason = @(Get-LogTail 'close batch' 1)
        if ($reason.Count -eq 1) { Write-Note $reason[0] }
        Assert (($reason.Count -eq 1) -and ($reason[0] -match 'declined')) `
            'the add-in stopped it because the user declined, not because it ran out of patience'

        # And it noticed the question at all. A prompt that goes up and comes down between two
        # janitor ticks is invisible to anything that polls, which is why this is heard as an event.
        $sawIt = @(Get-LogTail 'Word is asking the user about this document' 1)
        if ($sawIt.Count -eq 1) { Write-Note $sawIt[0] }
        Assert ($sawIt.Count -eq 1) 'the add-in saw the question go up before it saw it come down'
    }
} else {
    Write-Note "only $count window(s) - skipping the cancelled-batch check"
}

# ---- Close Others -------------------------------------------------------------------------------------

Write-Step 'Close Others'
$count = (Get-FrameCount)
if ($count -ge 2) {
    # Tab 0 is still dirty from the check above. Save it through the menu rather than answering a
    # prompt: this is the batch running cleanly, which is a different claim from the one above.
    $menu = Open-TabMenu 'label' 0
    if ($menu.Window -ne [IntPtr]::Zero) { [WordLayout]::Press($VK.S); Wait-Menu $false | Out-Null }
    Start-Sleep -Seconds 3
    if (Get-WordDialog) { Write-Note 'a dialog is up after saving tab 0 - dismissing it'; [WordLayout]::Press($VK.ESC); Start-Sleep -Seconds 2 }

    $keepIndex = 1
    $spot = Get-Spot 'label' $keepIndex
    [WordLayout]::Click($spot.X, $spot.Y)
    Start-Sleep -Milliseconds 1200
    $keepTitle = [WordLayout]::TitleOf([WordLayout]::GetForeground())
    Write-Note "keeping `"$keepTitle`" out of $count tabs"

    $menu = Open-TabMenu 'label' $keepIndex
    if ($menu.Window -ne [IntPtr]::Zero) { [WordLayout]::Press($VK.O); Wait-Menu $false | Out-Null }

    $down = Wait-Frames 1 45
    Assert $down "Close Others left exactly one document ($((Get-FrameCount)))"
    if ((Get-FrameCount) -eq 1) {
        $leftTitle = [WordLayout]::TitleOf((Get-Frames)[0])
        Write-Note "left with `"$leftTitle`""
        Assert ($leftTitle -eq $keepTitle) 'the one left is the tab it was invoked on'
    }
} else {
    Write-Note "only $count window(s) - skipping Close Others"
}

# ---- Close Others is greyed when there is nothing else to close --------------------------------------

Write-Step 'Close Others with only one tab'
if ((Get-FrameCount) -eq 1) {
    $menu = Open-TabMenu 'label' 0
    Assert ($menu.Window -ne [IntPtr]::Zero) 'the menu still opens on the last tab'
    $others = @($menu.Items | Where-Object { $_.Text -eq 'Close &Others' })[0]
    Assert ($others -and -not $others.Enabled) 'Close Others is greyed - there are no others'
    $close = @($menu.Items | Where-Object { $_.Text -eq '&Close' })[0]
    Assert ($close -and $close.Enabled) 'Close is still available'
    Close-Menu
} else {
    Write-Note "$((Get-FrameCount)) window(s) - skipping the greyed check"
}

# ---- Close All -----------------------------------------------------------------------------------------

Write-Step 'Close All'
if ((Get-FrameCount) -ge 1) {
    $menu = Open-TabMenu 'label' 0
    if ($menu.Window -ne [IntPtr]::Zero) { [WordLayout]::Press($VK.A); Wait-Menu $false | Out-Null }

    $gone = Wait-Frames 0 45
    if (-not $gone -and (Get-WordDialog)) {
        Write-Note 'a dialog is up - a document was still dirty; cancelling and reporting'
        [WordLayout]::Press($VK.ESC)
        Start-Sleep -Seconds 3
    }
    Assert $gone "Close All closed every document ($((Get-FrameCount)) windows left)"
}

# ---- done ------------------------------------------------------------------------------------------------

if (-not $KeepOpen) {
    Write-Step 'Making sure Word is closed'
    for ($guard = 0; $guard -lt 12; $guard++) {
        $open = @(Get-Frames)
        if ($open.Count -eq 0) { break }
        [WordLayout]::Close($open[0])
        Start-Sleep -Seconds 3
    }
    foreach ($process in @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue)) { $process.CloseMainWindow() | Out-Null }
    $deadline = (Get-Date).AddSeconds(20)
    while ((Get-Date) -lt $deadline -and (Get-Process -Name WINWORD -ErrorAction SilentlyContinue)) {
        Start-Sleep -Milliseconds 500
    }
    $stuck = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue)
    if ($stuck.Count -gt 0) { Write-Note "$($stuck.Count) Word process(es) would not close - left running rather than killed" }
}

Write-Host ''
if ($script:Failures -eq 0) {
    Write-Host "PASS  $($script:Checks) checks, 0 failures" -ForegroundColor Green
} else {
    Write-Host "FAIL  $($script:Failures) of $($script:Checks) checks failed" -ForegroundColor Red
}
Write-Host "Add-in log: $env:LOCALAPPDATA\WordTab\wordtab.log (lines starting 'strip', 'stack' and 'save')" -ForegroundColor Gray
exit $script:Failures
