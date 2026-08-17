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
    # Five, not four, and the fifth belongs to one section. "Move to New Window" takes a window out of
    # the stack and there is no way to put it back yet, so that section tears off the last tab and
    # then closes it - which leaves the four windows every section after it was written against.
    [int]$Documents = 5,
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

# Confirmed input, one shared idea of what a Word dialog is, and the bounded waits. See the header of
# tools\WordTabHarness.ps1 for why these are not per-suite copies any more.
#
# This suite has the most to gain from that. Its prompt detection is the most safety-critical
# assertion in the project - it is the only suite that deliberately provokes Word's save prompt - and
# the local copy it used to carry let `Net UI Tool Window` chrome read as a question. Its local
# Set-WordForeground would also minimise `Progman`, which is the desktop, 260 lines above a comment
# in this same file explaining why that must never happen.
. (Join-Path $PSScriptRoot 'WordTabHarness.ps1')

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

function Get-WordPids { return Get-WordPidList }

function Get-Frames { return Get-WordFrameList }

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
        Strip = @($kids | Where-Object { $_.Class -eq 'WordTabStrip' }) | Select-Object -First 1
        Wwf   = @($kids | Where-Object { $_.Class -eq '_WwF' }) | Select-Object -First 1
        Title = [WordLayout]::TitleOf($frame)
    }
}

# Set-WordForeground comes from WordTabHarness.ps1 now. The copy that used to live here had the
# right idea - a window over Word does not merely steal the keystrokes, it *receives the clicks*, and
# a reading of the tab row once came back naming "Terminal" as one of the documents - but it would
# minimise ANY window that would not give the foreground up. That includes `Progman`, which is the
# desktop: minimising it is Show Desktop, which takes Word down with everything else. The shared one
# refuses to minimise the desktop, and refuses to minimise a window belonging to Word itself, which
# is what a modal save prompt is.

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

# Where to aim, measured now - see the note in the description.
#   'label' - the left third of a tab, clear of its close button
#   'empty' - the strip to the right of the new-document button, which belongs to no tab
# `$tabs` overrides how many tabs the row is assumed to hold. It defaults to the number of Word
# windows, which is the same thing right up until one of them leaves the stack - after "Move to New
# Window" there are five windows and the row in front of you has four. Getting that wrong does not
# throw: the computed slots are simply narrower than the real ones and the click still lands, which is
# the kind of passing-for-the-wrong-reason this suite is full of guards against.
function Get-Spot($kind, $index, $tabs = 0) {
    $top = Get-TopStrip
    $count = if ($tabs -gt 0) { $tabs } else { Get-FrameCount }
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
#
# **That reading is now taken from a mark, and it was not before.** The version here searched the
# last 600 lines of the whole file and took the last match, so it could answer with a line from a
# PREVIOUS run - and whether it did depended on how much the add-in happened to have logged since.
# Under the most safety-critical assertion in this project, that is a coin toss dressed as evidence.
# Set-LogMark / Get-LogLast come from WordTabHarness.ps1; the mark is taken immediately before the
# gesture, so a line found after it can only have been written by this batch.

function Get-MenuWindow { return Get-WordMenu }

function Wait-Menu($present, $seconds = 6) {
    return (Wait-Until { (((Get-WordMenu) -ne [IntPtr]::Zero) -eq $present) } $seconds 200)
}

# "Word is asking the user something", the most safety-critical measurement in this suite, now comes
# from WordTabHarness.ps1.
#
# The version that used to live here excluded OpusApp and #32768 and took anything else over 200x100.
# It could not tell a save prompt from `Net UI Tool Window`, which is Office's floating-UI host -
# a gallery or a toast, not a question. One appeared at 622x298 during a battery run and three suites
# went red off it. See Get-WordWindowKind for what a question is, measured rather than assumed.
function Wait-Dialog($seconds = 10) { return (Wait-WordDialog $true $seconds) }

# Right-click a tab and return what came up. The menu's items are read through its owner frame,
# which is what GetMenuItemRect needs to answer with screen rectangles.
function Open-TabMenu($kind, $index, $tabs = 0) {
    $owner = (Get-TopStrip).Frame
    # Confirmed onto the strip before the right button goes down. A right-click that lands on the
    # document opens WORD's context menu, which is also a visible #32768 - so "a context menu
    # appeared" would pass while measuring Word's menu instead of ours.
    $spot = Get-Spot $kind $index $tabs
    $aim = { $script:menuSpot = Get-Spot $kind $index $tabs; return $script:menuSpot }
    if (Invoke-ConfirmedClick -What "right-clicking $kind $index" -Point $aim -Button right) {
        $spot = $script:menuSpot
    } else {
        Write-Note "the right-click never landed on the strip"
    }
    $up = Wait-Menu $true
    if (-not $up) { return [pscustomobject]@{ Window = [IntPtr]::Zero; Items = @(); Spot = $spot } }
    $window = Get-MenuWindow
    return [pscustomobject]@{
        Window = $window
        Items  = @([WordLayout]::MenuItems($window, $owner))
        Spot   = $spot
    }
}

$VK = @{ S = 0x53; C = 0x43; O = 0x4F; A = 0x41; N = 0x4E; M = 0x4D; W = 0x57; ESC = 0x1B; X = 0x58
         B = 0x42 }

function Close-Menu {
    if ((Get-MenuWindow) -ne [IntPtr]::Zero) {
        # Confirmed that Word can receive it. An Escape that goes to another application leaves the
        # menu up, and the next section then measures a menu it thinks it closed.
        Invoke-ConfirmedKey -Vk $VK.ESC -What 'Escape to close the menu' | Out-Null
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

# Put the caret in the document and type, so the document behind tab $index is genuinely modified -
# and *prove* it before returning.
#
# A keystroke sent to a window that was only activated programmatically goes nowhere in Word, which
# is why this clicks into the page first. Two further things make this the most fragile step in the
# suite, and both fail silently.
#
# Clicking a tab ends in SetForegroundWindow inside Word, and Windows refuses that call from a
# process that is not already the foreground application - so the document can switch without Word
# coming forward, and the keystroke then goes to whatever does own the desktop. That is not
# hypothetical: run from a terminal, this suite typed into the terminal and reported "Save wrote the
# document to disk" as a failure, twice, with the add-in behaving perfectly and its own log saying
# so ("Document.Save returned (Word confirmed the window handle)"). A Save on an *unmodified*
# document correctly writes nothing, so a missed keystroke is indistinguishable from a broken Save.
#
# So the keystrokes are checked by photographing the page before and after. A threshold rather than
# "any pixel changed", because the caret blinks: a caret is a couple of pixels wide and several
# characters are not, and the two are hundreds of sampled pixels apart.
function Set-Dirty($index) {
    for ($try = 1; $try -le 6; $try++) {
        Set-WordForeground | Out-Null
        # Confirmed onto the strip. If this click misses, the wrong document gets typed into and the
        # failure surfaces three steps later as "Save wrote the document to disk - FAIL".
        if (-not (Invoke-ConfirmedClick -What "selecting tab $index before typing" -Point { Get-Spot 'label' $index })) {
            Write-Note "attempt ${try}: could not put the click on tab $index"
            continue
        }
        Start-Sleep -Milliseconds 900

        $top = Get-TopStrip
        if (-not $top.Wwf) { throw 'The foreground Word window has no document frame to type into.' }

        # Aimed at the document frame itself, never at a fixed offset from the strip. Word remembers
        # its window size between sessions, and this suite has seen it come back small enough that
        # "the strip's bottom plus 300 pixels" landed on the *desktop* - measured, the click reported
        # SysListView32 under it and the foreground as "Program Manager". Nothing was typed, the
        # document was never modified, and the assertion that failed was "Save wrote the document to
        # disk", which is three steps away from the actual cause.
        $view = [WordLayout]::RectOf($top.Wwf.Hwnd)
        $vw = $view.Right - $view.Left
        $vh = $view.Bottom - $view.Top
        if ($vw -lt 200 -or $vh -lt 200) { throw "The document frame is only ${vw}x${vh} - too small to type into." }

        $x = $view.Left + [int]($vw / 3)
        $y = $view.Top  + [int]($vh / 2)

        # Two bands, both sized from the view so they stay inside it whatever the window size: where
        # the characters will appear, and a control band below them that five characters cannot
        # reach. The control is not fastidiousness - it is what stops this check passing for the
        # wrong reason. If another window covers Word between the two photographs, *every* sampled
        # pixel differs in both bands, which reads as a huge successful edit. Measured: exactly that,
        # 15000 of 15000 samples, on a contended desktop.
        $bandH = [Math]::Min(100, [int]($vh / 6))
        $bandW = [Math]::Min(600, [int]($vw / 2))

        $page = New-Object WordLayout+RECT
        $page.Left   = $x - 40
        $page.Top    = $y - $bandH
        $page.Right  = $x + $bandW
        $page.Bottom = $y + $bandH

        $control = New-Object WordLayout+RECT
        $control.Left   = $page.Left
        $control.Top    = $page.Bottom + $bandH
        $control.Right  = $page.Right
        $control.Bottom = $control.Top + $bandH

        $shot = New-Object WordLayout+RECT
        $shot.Left = $page.Left; $shot.Top = $page.Top
        $shot.Right = $page.Right; $shot.Bottom = $control.Bottom

        # The photograph is taken *before* the click, so that the gap between the click into the page
        # and the keystrokes is as short as it can be. That gap is the whole risk: a real click makes
        # Word foreground by user input, which Windows honours - and anything that takes the desktop
        # back before the keys are injected swallows them. Half a second of waiting in here was
        # measured losing them about half the time.
        $before = Get-RectShot $shot

        # Deliberately no Focus() here, and deliberately NOT Invoke-ConfirmedClick either: that takes
        # the foreground back before it clicks, and calling Focus() in between was measured to leave
        # the window activated but the caret unplaced, so the characters went nowhere and only the
        # click's own repaint showed up in the photograph. The click into the page is what makes Word
        # foreground, and it does it as *user input*, which Windows honours unconditionally.
        #
        # So the aim is confirmed the passive way - ask what is under the point, change nothing - and
        # the outcome is confirmed by the photograph below. Two different failures with two different
        # names: this one says the click was aimed off the page, that one says the keys never arrived.
        $under = Get-ClassAt $x $y
        if ($under -notin @('_WwG', '_WwN', '_WwB')) {
            Write-Note ("attempt {0}: ({1},{2}) is over `"{3}`", not the document page - re-measuring" -f $try, $x, $y, $under)
            continue
        }
        [WordLayout]::Click($x, $y)
        Start-Sleep -Milliseconds 600      # Word needs this to place the caret before it will take keys
        for ($k = 0; $k -lt 5; $k++) { [WordLayout]::Press($VK.X) }
        Start-Sleep -Milliseconds 700
        $after = Get-RectShot $shot

        $changed = Measure-Changed $before.Bitmap $after.Bitmap $before.Origin $page $null
        $moved   = Measure-Changed $before.Bitmap $after.Bitmap $before.Origin $control $null
        $before.Bitmap.Dispose(); $after.Bitmap.Dispose()

        # An edit is *localised*: five characters reflow one line inside a band a hundred pixels tall
        # and six hundred wide, which is a few percent of the samples. A change that fills the band -
        # or that reaches the control band below it - is a different window being photographed, not
        # text being typed. Both bounds are load-bearing; without the upper one this check reported a
        # fully occluded window as a large successful edit.
        $samples = [int](($page.Right - $page.Left) / 2) * [int](($page.Bottom - $page.Top) / 2)
        $localised = ($changed -ge 100) -and ($changed -le [int]($samples / 2)) -and ($moved -lt 100)

        if ($localised) {
            Write-Note "the document was modified ($changed of $samples samples on the line, $moved below it)"
            return [WordLayout]::GetForeground()
        }

        $covered = ($moved -ge 100) -or ($changed -gt [int]($samples / 2))
        $why = if ($covered) {
            "the whole view changed ($changed of $samples on the line, $moved on the control band) - something covered Word, not a keystroke"
        } else {
            "the keystrokes did not reach the document ($changed of $samples samples changed)"
        }
        $front = [WordLayout]::GetForeground()
        Write-Note ("attempt {0}: {1}. Foreground is `"{2}`"" -f $try, $why, [WordLayout]::TitleOf($front))

        # Something is *over* Word rather than merely in front of it in the input queue - an
        # always-on-top window survives the foreground handshake, which is why Set-WordForeground
        # cannot see this: it asks GetForegroundWindow, gets a Word frame, and returns happy while
        # the clicks and keystrokes still land on the window on top. Photographing the page is what
        # catches it, so this is the right place to get the offender out of the way. Gated on a
        # proven failure and recoverable from the taskbar.
        if ($covered) {
            $front = [WordLayout]::GetForeground()
            # Never the shell. `Progman` and `WorkerW` are the desktop itself, and minimising the
            # desktop is "Show Desktop" - it would take Word down with everything else, turning a
            # recoverable obstruction into a run that cannot continue at all.
            $cls = if ($front -ne [IntPtr]::Zero) { [WordLayout]::ClassOf($front) } else { '' }
            if ($front -ne [IntPtr]::Zero -and -not (@(Get-Frames) -contains $front) -and
                $cls -notin @('Progman', 'WorkerW') -and [WordLayout]::TitleOf($front) -ne '') {
                Write-Note ("minimising `"{0}`" ({1}) - it is sitting over Word and taking the input meant for it" -f `
                            [WordLayout]::TitleOf($front), $cls)
                [WordLayout]::Show($front, 6)      # SW_MINIMIZE
                Start-Sleep -Milliseconds 900
            }
        }
    }

    Write-Note 'WARNING: could not modify the document - the assertions that need a dirty document will fail'
    return [WordLayout]::GetForeground()
}

function Get-RectShot($r) {
    $bmp = New-Object System.Drawing.Bitmap(($r.Right - $r.Left), ($r.Bottom - $r.Top))
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($r.Left, $r.Top, 0, 0, $bmp.Size)
    $g.Dispose()
    return [pscustomobject]@{ Bitmap = $bmp; Origin = $r }
}

function Get-StripShot($strip) { return Get-RectShot ([WordLayout]::RectOf($strip)) }

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
    # A leftover document with unsaved changes is deliberately NOT answered: this script discards
    # nothing it did not create, and "the cleanup step threw away my work" is not a thing a test
    # script may ever do. Close-AllWord stops on a question and hands it back.
    $start = Close-AllWord
    if (-not $start.Closed) {
        if ($start.Reason -eq 'question') {
            throw ("Word is asking about unsaved changes left over from an earlier run: {0}  Answer it by hand (on this rig the check documents are disposable, so Don't Save is fine) and re-run." -f
                   (Format-WordWindow $start.Dialog))
        }
        throw ("Word would not close ({0}, {1} frame(s) left) - close it by hand, then re-run." -f $start.Reason, $start.Frames)
    }
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
    # Waiting for Word to get into position. Nothing below asserts how long Word takes to open a
    # file, so a wait here cannot weaken a check - and 14 + 7 + 7 was 28 seconds of every run.
    if (-not (Wait-WordReady $i 45)) { Write-Note "document $i did not arrive with a strip on it within 45s" }
}

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

$expected = @(Get-TabMenuItems)
$actual = @($menu.Items | ForEach-Object { $_.Text })
foreach ($item in $menu.Items) {
    Write-Note ("[{0}] id={1} enabled={2} {3} `"{4}`"" -f $item.Index, $item.Id, $item.Enabled,
                $(if ($item.HasRect) { Format-Rect $item.Rect } else { 'no rect' }), $item.Text)
}
Assert ($actual.Count -eq $expected.Count) "the menu has $($expected.Count) entries ($($actual.Count))"
Assert (($actual -join '|') -eq ($expected -join '|')) "the entries are exactly: $($expected -join ', ')"

$others = @($menu.Items | Where-Object { $_.Text -eq 'Close &Others' } | Select-Object -First 1)
Assert ($others -and $others.Enabled) 'Close Others is available with more than one tab'

# The tab the menu belongs to stays lit. The pointer is parked off the strip first, so what is being
# measured is the menu's own highlight and not the hover that any pointer over a tab would produce.
Write-Step 'The tab the menu belongs to'
# Confirmed by cursor readback, not by what is under the point - (4,4) is deliberately off the strip.
# This park is load-bearing: if it silently does not take, the pointer is still on the tab and the
# "menu highlight" being measured is the ordinary hover.
Assert (Set-Pointer 4 4 'parking the pointer clear of the strip') 'the pointer could be parked off the strip'
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
# The press point is confirmed before the gesture, never during it - this is one continuous
# button-down movement and re-aiming halfway through would break the thing being measured. Without
# this, a press that landed off the strip also produces "no menu", and the cancel goes unchecked.
$onWhat = Get-ClassAt $onTabSpot.X $onTabSpot.Y
Assert ($onWhat -eq 'WordTabStrip') "the right-press starts on the strip (`"$onWhat`" is under the press point)"
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

# ---- Move to New Window ------------------------------------------------------------------------------
#
# Taking a tab out of the stack, which is the first half of tear-off; the drag gesture is the second.
#
# The mechanism under test is not "a window moved". Every window in the stack sits at the same
# rectangle, so a window that leaves and keeps its position is indistinguishable from one that never
# left - and the two things this has to get right are both invisible in a screenshot. The first is
# reachability: the stack takes a window's taskbar button and its Alt+Tab entry away, and a window
# with neither, underneath another window at the same rectangle, cannot be got back to by any means
# the user has. The second is that it STAYS out - membership is otherwise re-derived from what is true
# about the window twice a second, and a torn-off window passes every one of those tests.
#
# The tab torn off is the LAST one, and it is identified by clicking it first. Which window a slot
# belongs to is knowable from outside Word only by activating it and reading the title, which is the
# same oracle check-stack's click loop uses - and it matters here, because "some window left the
# stack" is a much weaker claim than "the one whose tab I right-clicked left the stack".

Write-Step 'Move to New Window'
$stackCount = Get-FrameCount
$lastTab = $stackCount - 1
Write-Note "$stackCount windows in the stack; tearing off tab $lastTab"

# Click it first, so the window behind that slot is known before anything moves it.
Invoke-StripClick "clicking tab $lastTab to learn which window it is" { Get-Spot 'label' $lastTab } | Out-Null
Start-Sleep -Milliseconds 900
$victim = [WordLayout]::GetForeground()
$victimTitle = [WordLayout]::TitleOf($victim)
$stackRect = [WordLayout]::RectOf($victim)
Assert (@(Get-Frames) -contains $victim) "tab $lastTab activated a Word window (`"$victimTitle`")"
Write-Note ("it is 0x{0:X} at {1}" -f [int64]$victim, (Format-Rect $stackRect))

$menu = Open-TabMenu 'label' $lastTab
Assert ($menu.Window -ne [IntPtr]::Zero) 'the menu opened on that tab'
$move = @($menu.Items | Where-Object { $_.Text -eq '&Move to New Window' } | Select-Object -First 1)
Assert ($move -and $move.Enabled) 'Move to New Window is offered and available with more than one tab'

Set-LogMark
if ($menu.Window -ne [IntPtr]::Zero) {
    Invoke-ConfirmedKey -Vk $VK.M -What "the menu's Move to New Window mnemonic" | Out-Null
    Wait-Menu $false | Out-Null
}

# Waiting for the window to get where it is going, not for the assertion. The move is one
# SetWindowPos inside the command; what takes time is Word repainting the interior after StripRefit.
Wait-Until { ([WordLayout]::RectOf($victim)).Left -ne $stackRect.Left } 8 | Out-Null
Start-Sleep -Milliseconds 800

function Get-TearOffState($victim, $stackRect) {
    $frames = @(Get-Frames)
    $out = @($frames | Where-Object { ([WordLayout]::RectOf($_)).Left -ne $stackRect.Left -or
                                      ([WordLayout]::RectOf($_)).Top -ne $stackRect.Top })
    $inStack = @($frames | Where-Object { $out -notcontains $_ })
    return [pscustomobject]@{
        Frames  = $frames
        Out     = $out
        InStack = $inStack
        Shown   = @($frames | Where-Object { -not [WordLayout]::IsToolWindow($_) })
    }
}

$state = Get-TearOffState $victim $stackRect
foreach ($f in $state.Frames) {
    Write-Note ("0x{0:X}  {1}  inAltTab={2}{3}" -f [int64]$f, (Format-Rect ([WordLayout]::RectOf($f))),
                (-not [WordLayout]::IsToolWindow($f)), $(if ($f -eq $victim) { '  <- torn off' } else { '' }))
}

Assert (@(Get-Frames).Count -eq $stackCount) "nothing closed - still $stackCount windows ($((Get-FrameCount)))"
Assert ($state.Out.Count -eq 1) "exactly one window left the stack ($($state.Out.Count))"
Assert ($state.Out.Count -eq 1 -and $state.Out[0] -eq $victim) 'and it is the one whose tab was right-clicked'

# Reachability, which is the safety property. Two windows are now presented to the shell: the one
# torn off, and the active one in what is left of the stack. Anything less than two means a window
# the user cannot get to.
Assert ($state.Shown.Count -eq 2) "two windows are in Alt+Tab and the taskbar - the torn-off one and the stack's ($($state.Shown.Count))"
Assert (-not [WordLayout]::IsToolWindow($victim)) 'the torn-off window got its Alt+Tab entry back'
Assert ([WordLayout]::GetForeground() -eq $victim) 'and it is in front - the user asked for this window'

# It has to have gone somewhere the user can see. Same size, offset by one caption-and-border, which
# is the step the shell cascades new windows by.
$tornRect = [WordLayout]::RectOf($victim)
$dx = $tornRect.Left - $stackRect.Left
$dy = $tornRect.Top - $stackRect.Top
Write-Note "moved by ($dx,$dy)"
Assert ($dx -gt 0 -and $dy -gt 0) "it moved down and right of the stack ($dx,$dy)"
Assert ($dx -eq $dy) 'by the same step in both directions'
Assert ((($tornRect.Right - $tornRect.Left) -eq ($stackRect.Right - $stackRect.Left)) -and
        (($tornRect.Bottom - $tornRect.Top) -eq ($stackRect.Bottom - $stackRect.Top))) 'and it kept its size'

# The windows left behind are still one stack.
$stillOne = @($state.InStack | Where-Object {
    $r = [WordLayout]::RectOf($_)
    $r.Left -eq $stackRect.Left -and $r.Top -eq $stackRect.Top -and
    $r.Right -eq $stackRect.Right -and $r.Bottom -eq $stackRect.Bottom
})
Assert ($state.InStack.Count -eq $stackCount - 1) "the other $($stackCount - 1) are still stacked ($($state.InStack.Count))"
Assert ($stillOne.Count -eq $state.InStack.Count) 'and they are all still at one rectangle'

$place = Get-StripPlacement (@(Get-Frames))
Assert ($place.Measured -eq $stackCount) "all $stackCount windows could be measured for strip placement (measured $($place.Measured))"
Assert $place.Ok ('every strip still sits between the chrome and the document' + $place.Text)

# The add-in's own account. Asserted as a count rather than as "at least one", because a command that
# ran twice - the menu posts, the janitor ticks underneath it - would leave two windows out and one
# of the geometry assertions above would still pass.
$said = @(Get-LogSince 'torn off, now its own window')
foreach ($line in $said) { Write-Note $line }
Assert ($said.Count -eq 1) "the add-in tore off exactly one window ($($said.Count) log lines)"
Assert (@(Get-LogSince "($($stackCount - 1) left in the stack)").Count -eq 1) "and it says $($stackCount - 1) are left"

# **The assertion this section exists for.** The janitor re-tests membership twice a second and the
# torn-off window passes every test it applies - visible, sized, holding a document. Without the
# sticky flag it rejoins within half a second and every assertion above still passes, because they
# all ran inside the first tick. Three ticks, then ask again.
Write-Step 'It stays out'
Start-Sleep -Milliseconds 1800
$after = Get-TearOffState $victim $stackRect
Assert ($after.Out.Count -eq 1 -and $after.Out[0] -eq $victim) 'three janitor ticks later it is still out of the stack'
Assert ($after.Shown.Count -eq 2) 'still two windows presented to the shell'
Assert (([WordLayout]::RectOf($victim)).Left -eq $tornRect.Left) 'and nothing moved it back'

# The last tab in a stack has nowhere to go, and the torn-off window is now a stack of one. Free to
# check here and it is the case the guard in StackTearOffTab exists for.
Write-Step 'Move to New Window on a window that is already on its own'
# One tab, explicitly: the torn-off window is in front, its row holds a single tab, and there are
# still five Word windows. See the note on Get-Spot.
#
# Which window this menu belongs to is asserted rather than assumed. Get-TopStrip takes the
# foreground, and "the greyed item was on the torn-off window's own row" is the whole claim - reading
# it off the stack's row instead would be a pass measuring the wrong window.
Assert ((Get-TopStrip).Frame -eq $victim) 'the torn-off window is the one in front'
$menu = Open-TabMenu 'label' 0 1
Assert ($menu.Window -ne [IntPtr]::Zero) 'the menu opens on the torn-off window'
$move = @($menu.Items | Where-Object { $_.Text -eq '&Move to New Window' } | Select-Object -First 1)
Assert ($move -and -not $move.Enabled) 'Move to New Window is greyed - it is the only tab there is'

# ...and the way back is offered, on this window and nowhere else. The whole list is asserted rather
# than just the new item: an item that is HIDDEN on other menus rather than greyed changes the shape
# of this one, and the shape is what the rule about Close All never moving under the pointer is about.
$ownItems = @($menu.Items | ForEach-Object { if ($_.Separator) { '-' } else { $_.Text } })
$ownWanted = @(Get-TabMenuItems -OnItsOwn)
Write-Note ("items: {0}" -f ($ownItems -join ', '))
Assert (($ownItems -join '|') -eq ($ownWanted -join '|')) `
       'a window on its own is offered Move Back to the Tab Row, and nothing else has changed'
$back = @($menu.Items | Where-Object { $_.Text -eq 'Move &Back to the Tab Row' } | Select-Object -First 1)
Assert ($back -and $back.Enabled) 'and it is available - there is a stack for it to go back to'

# ---- and it works ------------------------------------------------------------------------------------
#
# Driven, not merely offered. An item asserted to be present and enabled and never invoked is a claim
# about a menu, not about a feature - and the command behind this one moves a window, snaps it onto
# the stack rectangle and rebuilds the row.
#
# The gesture that does the same thing is driven in check-reorder, which is where the drags live. Same
# split as tearing off: the menu route is proven here, the drag route there, and both raise the one
# command so neither is a second implementation.

Set-LogMark
if ($menu.Window -ne [IntPtr]::Zero) {
    Invoke-ConfirmedKey -Vk $VK.B -What "the menu's Move Back to the Tab Row mnemonic" | Out-Null
    Wait-Menu $false | Out-Null
}
Wait-Until { ([WordLayout]::RectOf($victim)).Left -eq $stackRect.Left } 10 | Out-Null
Start-Sleep -Milliseconds 1000

$wentBack = @(Get-LogSince 'put back into the stack')
foreach ($line in $wentBack) { Write-Note $line.Trim() }
Assert ($wentBack.Count -eq 1) "the add-in put exactly one window back ($($wentBack.Count))"
Assert ((Get-FrameCount) -eq $stackCount) "still $stackCount windows ($((Get-FrameCount)))"

$backState = Get-TearOffState $victim $stackRect
Assert ($backState.Out.Count -eq 0) "no window is standing outside the stack any more ($($backState.Out.Count))"
Assert ($backState.Shown.Count -eq 1) "and the stack is one taskbar button again ($($backState.Shown.Count))"

# Out again, so the sections below meet the fixture they were written against: this section ends by
# closing the torn-off document, and the assertion that the tear-off is forgotten with the DOCUMENT
# rather than with the window is the one that caught a real bug three sections downstream.
Write-Step 'Taking it back out to restore the fixture'
Set-LogMark
$menu = Open-TabMenu 'label' ($stackCount - 1)
Assert ($menu.Window -ne [IntPtr]::Zero) 'the menu opens on the rejoined tab, which is now the last one'
if ($menu.Window -ne [IntPtr]::Zero) {
    Invoke-ConfirmedKey -Vk $VK.M -What "Move to New Window on the rejoined tab" | Out-Null
    Wait-Menu $false | Out-Null
}
Wait-Until { ([WordLayout]::RectOf($victim)).Left -ne $stackRect.Left } 8 | Out-Null
Start-Sleep -Milliseconds 800
Assert (@(Get-LogSince 'torn off, now its own window').Count -eq 1) 'it is out of the stack again'

# Back to a stack of four, which is what every section below was written against.
Write-Step 'Closing the torn-off window'
# Aimed at the one window, not at Word. Ctrl+W closes whatever is in front, and taking the foreground
# generically here would close a document the sections below are about. Invoke-ConfirmedKeyOn focuses
# the named window itself and refuses to press until the system agrees it is the foreground.
Invoke-ConfirmedKeyOn -Hwnd $victim -Vk $VK.W -Ctrl -What 'Ctrl+W on the torn-off window' | Out-Null
if (-not (Wait-Frames ($stackCount - 1) 30)) {
    Write-Note "the torn-off window did not close; $((Get-FrameCount)) windows left"
}
Assert ((Get-FrameCount) -eq $stackCount - 1) "back to $($stackCount - 1) windows for the sections below ($((Get-FrameCount)))"

# **And the tear-off has to be forgotten with the document, not with the window.** Word does not
# destroy a frame when its last document closes - it hides it and puts the next document straight back
# into it - so a flag that lived as long as the HWND outlived the decision it recorded. Left that way,
# the next document opened anywhere in this Word lands in the recycled frame and silently refuses to
# join the row: five windows, four tabs, and nothing on screen to explain it. Every section below this
# one is what found that, three steps from its cause, so it is asserted here where it happens.
$freed = Wait-Until { @(Get-LogSince 'no longer torn off').Count -ge 1 } 8
Assert $freed 'the tear-off is forgotten once the document has gone - that frame is Word''s to reuse'

# ---- Save writes the file ----------------------------------------------------------------------------

Write-Step 'Save, checked against the file on disk'
$before = (Get-Item $paths[0]).LastWriteTime
Write-Note "$($paths[0])  last written $($before.ToString('HH:mm:ss.fff'))"

Set-Dirty 0 | Out-Null
Start-Sleep -Milliseconds 500

$menu = Open-TabMenu 'label' 0
if ($menu.Window -ne [IntPtr]::Zero) {
    # Clicked rather than typed, like the Close test below. A popup menu is a topmost window, so a
    # click at its item's rectangle lands on it whatever else owns the desktop; an injected access
    # key only lands if the menu has the keyboard at that instant, and this is the assertion that
    # suffers most when it does not - a Save that was never invoked and a Save that ran on an
    # unmodified document both leave the file's timestamp exactly where it was.
    $save = @($menu.Items | Where-Object { $_.Text -eq '&Save' })
    if ($save.Count -gt 0 -and $save[0].HasRect) {
        $centre = [WordLayout]::Center($save[0].Rect)
        Write-Note ("clicking the Save item at {0}" -f (Format-Rect $save[0].Rect))
        # Expecting #32768, the popup menu itself - not the strip. Set-WordForeground inside the
        # confirmed click returns immediately while a menu is up, so this cannot dismiss the menu it
        # is aiming at. A Save that was never invoked and a Save that ran on an unmodified document
        # leave the file's timestamp in exactly the same place, so an unconfirmed click here is the
        # single most misleading input in the suite.
        Assert (Invoke-ConfirmedClick -What 'clicking the Save menu item' -Expect '#32768' `
                                      -Point { [pscustomobject]@{ X = $centre.X; Y = $centre.Y } }) `
               'the click on Save landed on the menu'
    } else {
        Write-Note 'the Save item has no rectangle to click - using its access key instead'
        Invoke-ConfirmedKey -Vk $VK.S -What "the menu's Save mnemonic" | Out-Null
    }
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
Assert (Invoke-ConfirmedClick -What 'clicking the new-document button' -Point {
            $t = Get-TopStrip
            [WordLayout]::Center(([WordLayout]::Tabs($t.Strip.Hwnd, (Get-FrameCount))).Plus)
        }) 'the click on + landed on the strip'
$arrived = Wait-Frames ($count + 1) 25
Assert $arrived "a new blank document arrived as a tab ($((Get-FrameCount)) windows)"

if ($arrived) {
    $newIndex = (Get-FrameCount) - 1
    $menu = Open-TabMenu 'label' $newIndex
    if ($menu.Window -ne [IntPtr]::Zero) {
        Invoke-ConfirmedKey -Vk $VK.S -What "the menu's Save mnemonic" | Out-Null
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

    Invoke-ConfirmedKey -Vk $VK.ESC -What 'Escape' | Out-Null
    Start-Sleep -Seconds 3
    if (Get-WordDialog) { Invoke-ConfirmedKey -Vk $VK.ESC -What 'Escape' | Out-Null; Start-Sleep -Seconds 3 }

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
$item = @($menu.Items | Where-Object { $_.Text -eq '&Close' } | Select-Object -First 1)

if ($item -and $item.HasRect) {
    $centre = [WordLayout]::Center($item.Rect)
    Write-Note ("clicking the Close item at {0}" -f (Format-Rect $item.Rect))
    Assert (Invoke-ConfirmedClick -What 'clicking the Close menu item' -Expect '#32768' `
                                  -Point { [pscustomobject]@{ X = $centre.X; Y = $centre.Y } }) `
           'the click on Close landed on the menu'
} else {
    Write-Note 'the item has no rectangle to click - using its access key instead'
    Invoke-ConfirmedKey -Vk $VK.C -What "the menu's Close mnemonic" | Out-Null
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
    Invoke-ConfirmedClick -What "selecting tab $keep" -Point { Get-Spot 'label' $keep } | Out-Null
    Start-Sleep -Milliseconds 1200
    Write-Note ("keeping `"{0}`", {1} tabs in total" -f [WordLayout]::TitleOf([WordLayout]::GetForeground()), $count)

    # From here on, anything in the log was written by this batch and by nothing before it.
    Set-LogMark

    $menu = Open-TabMenu 'label' $keep
    if ($menu.Window -ne [IntPtr]::Zero) {
        Invoke-ConfirmedKey -Vk $VK.O -What "the menu's Close Others mnemonic" | Out-Null
        Wait-Menu $false | Out-Null
    }

    $prompt = Wait-Dialog 15
    if ($prompt) { Write-Note ("Word asked: `"{0}`" ({1})" -f $prompt.Title, $prompt.Class) }
    Assert ($prompt -ne $null) 'the dirty document raised Word''s save prompt before anything closed'

    if ($prompt) {
        $ownerDisabled = @(Get-Frames | Where-Object { -not [WordLayout]::Enabled($_) }).Count
        Write-Note "$ownerDisabled frame(s) disabled while the prompt is up"

        Invoke-ConfirmedKey -Vk $VK.ESC -What 'Escape' | Out-Null          # Cancel
        Start-Sleep -Seconds 10               # the grace after the question goes, and then some

        Assert ((Get-WordDialog) -eq $null) 'the prompt went away'
        Assert ((Get-FrameCount) -eq $count) `
            "cancelling stopped the whole batch - all $count documents are still open ($((Get-FrameCount)))"

        # ...and stopped for the right reason. Every assertion above is also satisfied by a batch
        # that gave up before Word had even asked - which is exactly what the first version of this
        # code did, one second after posting WM_CLOSE, while passing all of them.
        $reason = Get-LogLast 'close batch'
        if ($reason) { Write-Note $reason }
        Assert (($null -ne $reason) -and ($reason -match 'declined')) `
            'the add-in stopped it because the user declined, not because it ran out of patience'

        # And it noticed the question at all. A prompt that goes up and comes down between two
        # janitor ticks is invisible to anything that polls, which is why this is heard as an event.
        $sawIt = Get-LogLast 'Word is asking the user about this document'
        if ($sawIt) { Write-Note $sawIt }
        Assert ($null -ne $sawIt) 'the add-in saw the question go up before it saw it come down'
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
    if ($menu.Window -ne [IntPtr]::Zero) { Invoke-ConfirmedKey -Vk $VK.S -What "the menu's Save mnemonic" | Out-Null; Wait-Menu $false | Out-Null }
    Start-Sleep -Seconds 3
    if (Get-WordDialog) { Write-Note 'a dialog is up after saving tab 0 - dismissing it'; Invoke-ConfirmedKey -Vk $VK.ESC -What 'Escape' | Out-Null; Start-Sleep -Seconds 2 }

    $keepIndex = 1
    $spot = Get-Spot 'label' $keepIndex
    Invoke-ConfirmedClick -What "selecting tab $keepIndex" -Point { Get-Spot 'label' $keepIndex } | Out-Null
    Start-Sleep -Milliseconds 1200
    $keepTitle = [WordLayout]::TitleOf([WordLayout]::GetForeground())
    Write-Note "keeping `"$keepTitle`" out of $count tabs"

    $menu = Open-TabMenu 'label' $keepIndex
    if ($menu.Window -ne [IntPtr]::Zero) { Invoke-ConfirmedKey -Vk $VK.O -What "the menu's Close Others mnemonic" | Out-Null; Wait-Menu $false | Out-Null }

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
    $others = @($menu.Items | Where-Object { $_.Text -eq 'Close &Others' } | Select-Object -First 1)
    Assert ($others -and -not $others.Enabled) 'Close Others is greyed - there are no others'
    $close = @($menu.Items | Where-Object { $_.Text -eq '&Close' } | Select-Object -First 1)
    Assert ($close -and $close.Enabled) 'Close is still available'
    Close-Menu
} else {
    Write-Note "$((Get-FrameCount)) window(s) - skipping the greyed check"
}

# ---- Close All -----------------------------------------------------------------------------------------

Write-Step 'Close All'
if ((Get-FrameCount) -ge 1) {
    $menu = Open-TabMenu 'label' 0
    if ($menu.Window -ne [IntPtr]::Zero) { Invoke-ConfirmedKey -Vk $VK.A -What "the menu's Close All mnemonic" | Out-Null; Wait-Menu $false | Out-Null }

    $gone = Wait-Frames 0 45
    if (-not $gone -and (Get-WordDialog)) {
        Write-Note 'a dialog is up - a document was still dirty; cancelling and reporting'
        Invoke-ConfirmedKey -Vk $VK.ESC -What 'Escape' | Out-Null
        Start-Sleep -Seconds 3
    }
    Assert $gone "Close All closed every document ($((Get-FrameCount)) windows left)"
}

# ---- done ------------------------------------------------------------------------------------------------

if (-not $KeepOpen) {
    Write-Step 'Making sure Word is closed'
    $end = Close-AllWord
    if (-not $end.Closed) {
        Write-Note ("Word is still up ({0}): {1}" -f $end.Reason, (Format-WordWindow $end.Dialog))
        Write-Note 'Left running rather than killed. Answer it by hand before the next suite - a leftover Word poisons whatever runs next.'
    }
}

Write-Host ''
if ($script:Failures -eq 0) {
    Write-Host "PASS  $($script:Checks) checks, 0 failures" -ForegroundColor Green
} else {
    Write-Host "FAIL  $($script:Failures) of $($script:Checks) checks failed" -ForegroundColor Red
}
Write-Host "Add-in log: $env:LOCALAPPDATA\WordTab\wordtab.log (lines starting 'strip', 'stack' and 'save')" -ForegroundColor Gray
exit $script:Failures
