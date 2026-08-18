<#
.SYNOPSIS
  Hammer the hover tooltip and count how often it fails to appear - with, and without, an unrelated
  change to the tab row landing inside the half second it is being waited for.

.DESCRIPTION
  This exists for one unexplained result: in a battery on 2026-08-17, check-title's `title` suite
  had one tab of seven produce no tooltip in 4020ms and NO LOG LINE AT ALL, then never did it again
  in 208 further attempts. "No panel and nothing logged" was read as "TipShow was never reached",
  and that reading was not safe: TipShow had four exits that returned without writing anything, so
  a tooltip abandoned before it was shown looked exactly like a hover that never happened.

  The mechanism this probe was written to provoke: StripRefreshTabs called TipStop, and TipStop
  kills a PENDING arm as well as taking down a panel that is up. A pointer resting on a tab sends no
  further WM_MOUSEMOVE, so nothing re-arms it - which means any row change inside that half second
  (a dot coming on in another document, a title Word rewrote, a window activating) meant the tooltip
  never appeared at all, for as long as the hand stayed still.

  -Churn provokes exactly that, from a document that is NOT the one being hovered. Word's own object
  model is used to flip Document.Saved immediately before the pointer arrives, so the add-in's
  janitor - which polls at 500ms - sees the dot change and refreshes the row somewhere inside the
  window where the tooltip is being waited for. Nothing about the hovered tab changes.

  The A/B this was built for: run it -Churn against a build with the defect and again against one
  without. The numbers are the evidence, not the reasoning above.

.PARAMETER Rounds
  Hovers to perform. Each one parks the pointer off the row first, so each is a fresh arrival.

.PARAMETER Churn
  Change an unrelated document's dirty state immediately before each hover.

.PARAMETER KeepOpen
  Leave Word running afterwards.

.EXAMPLE
  pwsh -File tools\probe-tip.ps1 -Rounds 40
  pwsh -File tools\probe-tip.ps1 -Rounds 40 -Churn
#>
[CmdletBinding()]
param(
    [int]$Rounds = 30,
    [switch]$Churn,
    [switch]$KeepOpen
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$source = Get-Content -Raw -Path (Join-Path $PSScriptRoot 'WordLayout.cs')
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition $source -Language CSharp -ReferencedAssemblies @(
    'System.Runtime', 'System.Collections', 'System.Threading.Thread', 'netstandard'
)
[WordLayout]::MakeDpiAware() | Out-Null
. (Join-Path $PSScriptRoot 'WordTabHarness.ps1')

function Write-Step($text) { Write-Host "==> $text" -ForegroundColor Cyan }
function Write-Note($text) { Write-Host "    $text" -ForegroundColor DarkGray }

function Get-Frames { return Get-WordFrameList }
function Get-FrameCount { return Get-WordFrameTally }

function Get-Parts($frame) {
    $kids = [WordLayout]::Children($frame)
    [pscustomobject]@{
        Frame = $frame
        Rect  = [WordLayout]::RectOf($frame)
        Strip = @($kids | Where-Object { $_.Class -eq 'WordTabStrip' }) | Select-Object -First 1
    }
}

function Get-TopStrip {
    Set-WordForeground -Quiet | Out-Null
    $frames = @(Get-Frames)
    if ($frames.Count -eq 0) { throw 'No Word windows.' }
    $top = [WordLayout]::GetForeground()
    if (-not ($frames -contains $top)) { $top = $frames[0] }
    $parts = Get-Parts $top
    if (-not $parts.Strip) { throw 'The foreground Word window has no WordTab strip.' }
    return $parts
}

# A third of the way in, the same spot check-title hovers, so a miss here is a miss there.
function Get-TabSpot($index) {
    $top    = Get-TopStrip
    $layout = [WordLayout]::Tabs($top.Strip.Hwnd, (Get-FrameCount))
    if ($index -ge $layout.Tabs.Length) { throw "Tab $index does not exist ($($layout.Tabs.Length) tabs)." }
    $t = $layout.Tabs[$index]
    return [pscustomobject]@{
        X = $t.Left + [int](($t.Right - $t.Left) / 3)
        Y = [int](($t.Top + $t.Bottom) / 2)
    }
}

function Move-OffTheRow {
    $r = [WordLayout]::RectOf((Get-TopStrip).Strip.Hwnd)
    $ok = Set-Pointer ([int](($r.Left + $r.Right) / 2)) ([int]($r.Bottom + 120)) 'off the row'
    Start-Sleep -Milliseconds 300
    return $ok
}

# ---- fixtures ------------------------------------------------------------------------------------

$dir = Join-Path $env:TEMP 'wordtab-tip'
if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
$fixtures = @('Tip probe one.rtf', 'Tip probe two.rtf', 'Tip probe three.rtf') |
            ForEach-Object { Join-Path $dir $_ }
foreach ($f in $fixtures) {
    if (-not (Test-Path $f)) {
        Set-Content -Path $f -Encoding ASCII -Value '{\rtf1\ansi A document long enough to have a name that is cut.\par}'
    }
}

Write-Step 'Starting Word on three documents'
if (@(Get-WordPidList).Count -gt 0) {
    $end = Close-AllWord
    if (-not $end.Closed) { throw "Word would not close ($($end.Reason)) - close it by hand and re-run." }
    Start-Sleep -Seconds 2
}
foreach ($f in $fixtures) {
    $before = @(Get-Frames)
    Start-Process -FilePath 'winword.exe' -ArgumentList "`"$f`""
    $deadline = (Get-Date).AddSeconds(45)
    while ((Get-Date) -lt $deadline) {
        $new = @(Get-Frames | Where-Object { $before -notcontains $_ })
        if ($new.Count -gt 0) {
            Wait-Until { @([WordLayout]::Children($new[0]) | Where-Object { $_.Class -eq 'WordTabStrip' }).Count -gt 0 } 20 300 | Out-Null
            Start-Sleep -Seconds 2
            break
        }
        Start-Sleep -Milliseconds 400
    }
}
$tabs = Get-FrameCount
Write-Note "$tabs tabs in the row"
if ($tabs -lt 2) { throw "Need at least two tabs; the row has $tabs." }

# The window has to be wide enough that three tabs do not scroll, or a hover lands on a chevron.
$frame = (Get-TopStrip).Frame
$dpi   = [WordLayout]::Dpi($frame)
$geom  = [WordLayout]::RectOf($frame)
$wide  = [WordLayout]::Scale(760, $dpi)
if (($geom.Right - $geom.Left) -lt $wide) {
    [WordLayout]::Resize($frame, $wide, [Math]::Max([WordLayout]::Scale(500, $dpi), $geom.Bottom - $geom.Top))
    Start-Sleep -Seconds 2
}

# ---- the churn source ----------------------------------------------------------------------------

$churnDoc = $null
if ($Churn) {
    # GetActiveObject and NOT `New-Object -ComObject Word.Application`, which is what this was
    # written as first and which quietly does the wrong thing: it CoCreates, and Word answers by
    # starting a SECOND, hidden winword.exe with no documents in it. The symptom is
    # "The requested member of the collection does not exist" from Documents.Item(1), which reads
    # like a Word problem rather than like talking to the wrong Word. Measured: one winword.exe
    # before the call, two after.
    #
    # The CLSID is written out rather than resolved through CLSIDFromProgID because that call came
    # back with an all-zero GUID here, and a zero CLSID makes GetActiveObject fail with
    # MK_E_UNAVAILABLE - which looks exactly like "Word is not running".
    #
    # .NET Core dropped Marshal.GetActiveObject, so pwsh 7 has to reach the running-object table
    # itself. This is the whole of that.
    Add-Type -Namespace WordTabProbe -Name Ole -MemberDefinition @'
[DllImport("oleaut32.dll", PreserveSig=false)]
public static extern void GetActiveObject(ref Guid rclsid, IntPtr pvReserved,
                                          [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);
'@
    $clsid = [Guid]'000209FF-0000-0000-C000-000000000046'
    $app = $null
    [WordTabProbe.Ole]::GetActiveObject([ref]$clsid, [IntPtr]::Zero, [ref]$app)
    if ($app.Documents.Count -lt 1) { throw 'Attached to a Word with no documents in it.' }
    # Document 1 and never the hovered tab: the point is that a change to a document nobody is
    # pointing at takes down a tooltip about one they are.
    $churnDoc = $app.Documents.Item(1)
    Write-Note ("churning |{0}| - a document that is not being hovered" -f $churnDoc.Name)
}

# ---- the hovers ----------------------------------------------------------------------------------

Set-LogMark
$up = 0; $missed = 0; $unreached = 0

for ($round = 0; $round -lt $Rounds; $round++) {
    $tab = ($round % $tabs)
    if (-not (Move-OffTheRow)) { $unreached++; continue }

    if ($Churn) {
        # Flipped rather than set, so every round is a CHANGE - the janitor refreshes the row when
        # the dot changes, not when it is on. Saved=$true does not write the file; it only clears
        # the flag, which is also what stops this probe leaving Word with questions to ask at close.
        $churnDoc.Saved = -not $churnDoc.Saved
    }

    $spot = Get-TabSpot $tab
    if (-not (Set-Pointer $spot.X $spot.Y "tab $tab")) { $unreached++; continue }

    $r = Wait-WordTabTip -Present $true -Seconds 3 -What "round $round (tab $tab)"
    if ($r.Ok) { $up++ } else { $missed++ }
}

if ($Churn -and $churnDoc) { $churnDoc.Saved = $true }

# ---- what the add-in said --------------------------------------------------------------------------

Write-Step 'What the add-in logged'
$tipLines  = @(Get-LogSince 'tip:')
$abandoned = @($tipLines | Where-Object { $_ -match 'abandoned' })
$restarted = @($tipLines | Where-Object { $_ -match 'starting the wait again' })
$shown     = @($tipLines | Where-Object { $_ -match 'folder \|' }).Count
Write-Note "$($tipLines.Count) tip: lines - $shown shown, $($abandoned.Count) abandoned, $($restarted.Count) restarted"
foreach ($line in (@($abandoned) + @($restarted) | Select-Object -First 12)) { Write-Note ("  " + $line.Trim()) }

Write-Host ''
Write-Host ("{0} hovers: {1} tooltip, {2} NO TOOLTIP, {3} pointer never arrived  ({4})" -f
            $Rounds, $up, $missed, $unreached, $(if ($Churn) { 'with churn' } else { 'quiet' })) `
           -ForegroundColor $(if ($missed -gt 0) { 'Red' } else { 'Green' })

if (-not $KeepOpen) {
    $end = Close-AllWord
    if (-not $end.Closed) { Write-Note "Word would not close ($($end.Reason)) - left running." }
}
