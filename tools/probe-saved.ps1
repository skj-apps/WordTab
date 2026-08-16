<#
.SYNOPSIS
  Measure how Word reports "this document has unsaved changes", and how a document is reached from
  a window handle.

.DESCRIPTION
  The modified-document dot is the first thing WordTab draws that is about Word's *documents* rather
  than its windows, so two things have to be measured before a line of it is written:

    1. **What `Document.Saved` actually says**, in the states a user can put a document in: brand
       new and untouched, brand new and typed into, saved to disk, changed again, saved again,
       read-only, and in Protected View. The property is documented as "True if the document has not
       changed since it was last saved" - which is the opposite polarity from the thing being drawn,
       and worth confirming rather than assuming.

    2. **Whether a document can be found from an `OpusApp` window handle without activating
       anything.** `Application.Windows` is a collection of document windows and each one reports
       `Hwnd`; the menu slice measured that this matches our frame handle for the *active* window.
       This asks whether it holds for every window at once, which is what a per-tick poll needs.

  Also times a full pass, because the poll will run on the add-in's half-second janitor inside Word's
  own UI thread, and "how expensive is this" is not something to find out from a user.

  WHAT IT MEASURED (2026-08-15, this rig):
    - Saved is True for a brand new untouched document, False once typed into, True again after a
      save. So the dot is drawn on `Saved == FALSE`, and a fresh empty Document1 correctly has none.
    - Saved is SETTABLE. The add-in must only ever read it.
    - Every Window.Hwnd matches an OpusApp frame, for all windows at once - not only the active one.
    - 2.04 ms for a pass over two windows, cross-process, so a real upper bound.
    - A PROTECTED VIEW document is not in Application.Windows and not in Documents. It is alone in
      ProtectedViewWindows, and its Window.Hwnd is the sandbox's handle, not the frame's. A lookup by
      frame handle therefore finds nothing for that tab - which is the right answer anyway, since a
      Protected View document cannot be edited and so can never be dirty.
    - New-Object -ComObject Word.Application is the WRONG oracle from outside: it bound to an
      Application reporting zero documents while three were on screen. Section 3a keeps that.

  Everything here is disposable: fixtures go in %TEMP%\wordtab-saved and Word is closed at both ends.

.EXAMPLE
  pwsh -File tools\probe-saved.ps1
#>
[CmdletBinding()]
param([switch]$KeepOpen)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$source = Get-Content -Raw -Path (Join-Path $PSScriptRoot 'WordLayout.cs')
Add-Type -TypeDefinition $source -Language CSharp -ReferencedAssemblies @(
    'System.Runtime', 'System.Collections', 'System.Threading.Thread', 'netstandard'
)
[WordLayout]::MakeDpiAware() | Out-Null

function Write-Step($text) { Write-Host "==> $text" -ForegroundColor Cyan }
function Write-Note($text) { Write-Host "    $text" -ForegroundColor DarkGray }
function Write-Fact($text) { Write-Host "    $text" -ForegroundColor Yellow }

function Get-WordPids { @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }) }
function Get-Frames {
    $found = @()
    foreach ($id in Get-WordPids) { $found += [WordLayout]::Frames($id) }
    return @($found)
}
function Get-FrameCount { return @(Get-Frames).Count }

function Close-Word {
    if (@(Get-WordPids).Count -eq 0) { return }
    for ($guard = 0; $guard -lt 24; $guard++) {
        $open = @(Get-Frames)
        if ($open.Count -eq 0) { break }
        [WordLayout]::Close($open[0])
        Start-Sleep -Milliseconds 1000
    }
    foreach ($p in @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue)) { $p.CloseMainWindow() | Out-Null }
    $deadline = (Get-Date).AddSeconds(20)
    while ((Get-Date) -lt $deadline -and @(Get-WordPids).Count -gt 0) { Start-Sleep -Milliseconds 500 }
    Start-Sleep -Seconds 2

    # Opening a Protected View document leaves a WINWORD behind that has no window at all - measured,
    # one per run of this probe, and three of them accumulated before install.ps1 refused to run
    # because "3 WINWORD processes are running". Neither route above reaches it: there is no frame to
    # send WM_CLOSE to and CloseMainWindow has no main window to close.
    #
    # This is the one place in the project that ends a WINWORD by force, and the guard is what makes
    # it legitimate rather than a violation of "never Kill a Word": a process with no OpusApp frame
    # has no document open, so there is nothing for Word to offer to recover. A process WITH a frame
    # is left alone and reported, because that one may hold the user's work.
    foreach ($p in @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue)) {
        $frames = @([WordLayout]::Frames($p.Id))
        if ($frames.Count -eq 0) {
            Write-Note "  windowless WINWORD $($p.Id) left behind by Protected View - ending it"
            Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
        } else {
            Write-Note "  WINWORD $($p.Id) still has $($frames.Count) frame(s) - left alone"
        }
    }
    Start-Sleep -Seconds 1
}

# Every document window Word will admit to, with the two properties the poll would read.
# Errors are reported rather than swallowed: "Word would not answer" is the interesting answer.
function Get-WindowTable($app) {
    $rows = @()
    $count = $app.Windows.Count
    for ($i = 1; $i -le $count; $i++) {
        $row = [ordered]@{ Index = $i; Hwnd = $null; Doc = $null; Saved = $null; Path = $null; Error = '' }
        try {
            $w = $app.Windows.Item($i)
            try { $row.Hwnd = [int64]$w.Hwnd } catch { $row.Error += 'Hwnd:' + $_.Exception.Message + ' ' }
            try {
                $d = $w.Document
                $row.Doc   = $d.Name
                $row.Saved = [bool]$d.Saved
                $row.Path  = $d.Path
            } catch { $row.Error += 'Document:' + $_.Exception.Message }
        } catch { $row.Error += 'Windows.Item:' + $_.Exception.Message }
        $rows += [pscustomobject]$row
    }
    return @($rows)
}

function Show-WindowTable($app, $what) {
    $rows = Get-WindowTable $app
    Write-Note ("{0}: Application.Windows.Count = {1}, OpusApp frames = {2}" -f $what, $rows.Count, (Get-FrameCount))
    foreach ($r in $rows) {
        Write-Note ("  [{0}] hwnd=0x{1:X}  Saved={2}  |{3}|  path=|{4}| {5}" -f
                    $r.Index,
                    $(if ($null -eq $r.Hwnd) { 0 } else { $r.Hwnd }),
                    $r.Saved, $r.Doc, $r.Path, $r.Error)
    }
    return $rows
}

$dir = Join-Path $env:TEMP 'wordtab-saved'
Write-Step 'Word must be closed to start clean'
Close-Word
if (Test-Path $dir) { Remove-Item -Path $dir -Recurse -Force }
New-Item -ItemType Directory -Path $dir -Force | Out-Null

# ---- 1. Document.Saved through the states a user can produce -------------------------------------

Write-Step 'Document.Saved, state by state'

$word = New-Object -ComObject Word.Application
$plain = Join-Path $dir 'Saved probe.docx'
try {
    $word.Visible = $true
    $word.DisplayAlerts = 0
    Start-Sleep -Seconds 3

    $doc = $word.Documents.Add()
    Start-Sleep -Seconds 2
    Write-Fact ("a brand new, untouched document:            Saved = {0}" -f [bool]$doc.Saved)

    $word.Selection.TypeText("WordTab saved-state probe.")
    Write-Fact ("...after typing into it:                    Saved = {0}" -f [bool]$doc.Saved)

    $doc.SaveAs2($plain, 12)
    Start-Sleep -Seconds 1
    Write-Fact ("...after Save As:                           Saved = {0}" -f [bool]$doc.Saved)

    $word.Selection.TypeText(" More.")
    Write-Fact ("...after typing again:                      Saved = {0}" -f [bool]$doc.Saved)

    $doc.Save()
    Start-Sleep -Seconds 1
    Write-Fact ("...after Save:                              Saved = {0}" -f [bool]$doc.Saved)

    # And the write side, which is what an add-in must never do by accident: is Saved settable?
    # Worth knowing only so the code can be sure it never writes it.
    try { $doc.Saved = $true; Write-Fact ("Document.Saved is settable (set True -> {0})" -f [bool]$doc.Saved) }
    catch { Write-Fact "Document.Saved refused a write: $($_.Exception.Message)" }

    Write-Step 'Application.Windows against the OpusApp frames'
    $second = $word.Documents.Add()
    Start-Sleep -Seconds 2
    $word.Selection.TypeText("second document, dirty.")
    Start-Sleep -Milliseconds 500

    $rows = Show-WindowTable $word 'two documents, the second one dirty'

    $frames = @(Get-Frames | ForEach-Object { [int64]$_ })
    foreach ($r in $rows) {
        $match = ($null -ne $r.Hwnd) -and ($frames -contains $r.Hwnd)
        Write-Fact ("  window {0} Hwnd 0x{1:X} {2} an OpusApp frame this process owns" -f
                    $r.Index, $r.Hwnd, $(if ($match) { 'IS' } else { 'is NOT' }))
    }
    Write-Fact ("frames seen from outside: {0}" -f (($frames | ForEach-Object { '0x{0:X}' -f $_ }) -join ', '))

    # ---- 2. What a full pass costs ---------------------------------------------------------------
    #
    # The shape the add-in will use: Windows.Count once, then Hwnd + Document + Saved per window.
    # Timed out of process, so it is an upper bound on the in-process cost rather than the number
    # itself - the add-in is inside WINWORD and pays no marshalling at all.

    Write-Step 'What a pass costs (out of process, so an upper bound)'
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $passes = 50
    for ($n = 0; $n -lt $passes; $n++) {
        $c = $word.Windows.Count
        for ($i = 1; $i -le $c; $i++) {
            $w = $word.Windows.Item($i)
            $null = $w.Hwnd
            $null = $w.Document.Saved
        }
    }
    $sw.Stop()
    Write-Fact ("{0} passes over {1} window(s): {2:N2} ms each, cross-process" -f
                $passes, $word.Windows.Count, ($sw.Elapsed.TotalMilliseconds / $passes))

    $second.Saved = $true
    $second.Close(0)
    Start-Sleep -Seconds 1
} finally {
    try { $word.Quit(0) } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
Close-Word

# ---- 3. Read-only and Protected View -------------------------------------------------------------
#
# Both have to be opened the way a user opens them - Start-Process, not Documents.Open - because
# Protected View is a decision Word makes about a file arriving from the shell, and automation
# bypasses it. This is also the case where a document may not be in this Application's object model
# at all, which the poll has to survive.

Write-Step 'Read-only and Protected View, opened the way a user opens them'

$readonly  = Join-Path $dir 'Locked probe.docx'
$protected = Join-Path $dir 'Downloaded probe.docx'
Copy-Item $plain $readonly
Copy-Item $plain $protected
Set-ItemProperty -Path $readonly -Name IsReadOnly -Value $true
Set-Content -Path $protected -Stream 'Zone.Identifier' -Value "[ZoneTransfer]`r`nZoneId=3"

Start-Process -FilePath 'winword.exe' -ArgumentList "`"$plain`""
Start-Sleep -Seconds 16
Start-Process -FilePath 'winword.exe' -ArgumentList "`"$readonly`""
Start-Sleep -Seconds 8
Start-Process -FilePath 'winword.exe' -ArgumentList "`"$protected`""
Start-Sleep -Seconds 12

Write-Note ("titles now open: {0}" -f ((@(Get-Frames) | ForEach-Object { [WordLayout]::TitleOf($_) }) -join ' | '))

# Which process owns which frame. This is the whole question of this section: the add-in lives in a
# process, its Application object speaks only for that process, and a frame owned by another one is
# not in the object model it can see - however normal it looks on screen.
Write-Fact ("WINWORD processes: {0}" -f ((Get-WordPids) -join ', '))
foreach ($f in @(Get-Frames)) {
    Write-Fact ("  frame 0x{0:X}  pid={1}  |{2}|" -f
                [int64]$f, [WordLayout]::PidOf($f), [WordLayout]::TitleOf($f))
}

# ---- 3a. The attach that does not work, kept because it is a trap ------------------------------
#
# The obvious way to reach a running Word from outside is New-Object -ComObject Word.Application, on
# the reasoning that Word is a single-instance server. It is not reliable here: with a Protected View
# document up there is a second WINWORD, and this bound to something reporting Documents.Count = 0
# with a Windows collection that has no Count on it. Kept in the probe, and reported rather than
# swallowed, so nobody re-derives it - and so the number below can be compared against it.

Write-Step 'The running object table, which is the wrong oracle'
$live = $null
try {
    $live = New-Object -ComObject Word.Application
    $seen = '?'
    try { $seen = $live.Documents.Count } catch { $seen = "Documents.Count failed: $($_.Exception.Message)" }
    Write-Fact ("New-Object -ComObject Word.Application gave: Name=|{0}| Version={1} Documents.Count={2}" -f
                $live.Name, $live.Version, $seen)
    try {
        Write-Fact ("  ...and its Windows.Count = {0}" -f $live.Windows.Count)
    } catch {
        Write-Fact ("  ...and its Windows collection would not answer Count: {0}" -f $_.Exception.Message)
    }
    Write-Fact ("  {0} frames are on screen, so this object is NOT the Word the user is looking at" -f (Get-FrameCount))
} catch {
    Write-Fact "could not attach a Word.Application at all: $($_.Exception.Message)"
} finally {
    if ($live) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($live) | Out-Null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

# ---- 3b. Asking each window for its own object model -------------------------------------------
#
# OBJID_NATIVEOM on a frame's _WwG document pane returns that window's own Word.Window, in that
# window's own process - so there is no question of which instance answered. From it, the frame's
# document, AND the Application that owns it, which is the object an add-in loaded in that process
# would hold.

Write-Step 'Each frame asked for its own object model (OBJID_NATIVEOM)'

$frames = @(Get-Frames)
$apps   = @()          # one entry per distinct Application reached, for section 3c

foreach ($f in $frames) {
    $title = [WordLayout]::TitleOf($f)
    $om = $null
    try { $om = [WordLayout]::NativeOm($f) } catch { }

    if ($null -eq $om) {
        Write-Fact ("  frame 0x{0:X} pid={1} |{2}| -> NO native object model on its _WwG (pane present: {3})" -f
                    [int64]$f, [WordLayout]::PidOf($f), $title,
                    ([WordLayout]::DocumentPane($f) -ne [IntPtr]::Zero))
        continue
    }

    $line = "  frame 0x{0:X} pid={1} |{2}| ->" -f [int64]$f, [WordLayout]::PidOf($f), $title
    try {
        $hwnd = [int64]$om.Hwnd
        $line += " Window.Hwnd=0x{0:X} ({1})" -f $hwnd, $(if ($hwnd -eq [int64]$f) { 'MATCHES the frame' } else { 'DIFFERENT from the frame' })
    } catch { $line += " Hwnd failed: $($_.Exception.Message)" }
    try {
        $d = $om.Document
        $line += "  |{0}|  Saved={1}" -f $d.Name, [bool]$d.Saved
    } catch { $line += "  Document failed: $($_.Exception.Message)" }
    Write-Fact $line

    # Deduped by the Application's own COM identity, NOT by the frame's process id. Those are not
    # the same thing and assuming they were cost this probe a run: a Protected View frame is owned by
    # the main WINWORD, but the object model behind its pane belongs to the sandboxed one, so keying
    # on the frame's pid threw away the only Application that mattered.
    try {
        $app = $om.Application
        $iu  = [System.Runtime.InteropServices.Marshal]::GetIUnknownForObject($app)
        [System.Runtime.InteropServices.Marshal]::Release($iu) | Out-Null
        if (-not ($apps | Where-Object { $_.Identity -eq $iu })) {
            $apps += [pscustomobject]@{ Frame = $f; Identity = $iu; App = $app }
        }
    } catch { }
}

# ---- 3c. Would the add-in in the MAIN Word see the Protected View window? -----------------------
#
# This is the question the dot actually turns on. The add-in is in one process and holds one
# Application; a frame that Application cannot enumerate must produce no dot rather than a wrong one.
# Word also keeps Protected View windows in a collection of their OWN, so both are counted here.

Write-Step 'What one Application can enumerate'

Write-Note ("{0} distinct Application object(s) behind {1} frame(s)" -f $apps.Count, $frames.Count)

foreach ($entry in $apps) {
    $app = $entry.App
    $owner = [WordLayout]::PidOf($entry.Frame)
    try {
        $rows = @(Get-WindowTable $app)
        Write-Fact ("Application reached via frame 0x{0:X} (frame owned by pid {1}): Windows.Count={2}, Documents.Count={3}" -f
                    [int64]$entry.Frame, $owner, $rows.Count, $app.Documents.Count)
        foreach ($f in $frames) {
            $hit = @($rows | Where-Object { $_.Hwnd -eq [int64]$f })
            Write-Fact ("    frame 0x{0:X} |{1}| -> {2}" -f [int64]$f, [WordLayout]::TitleOf($f),
                        $(if ($hit.Count -eq 1) { "Saved=$($hit[0].Saved) |$($hit[0].Doc)|" }
                          elseif ($hit.Count -eq 0) { 'NOT in this Application.Windows' }
                          else { "$($hit.Count) windows claim this handle" }))
        }
    } catch {
        Write-Fact ("Application reached via frame 0x{0:X} would not enumerate Windows: {1}" -f [int64]$entry.Frame, $_.Exception.Message)
    }

    try {
        $pv = $app.ProtectedViewWindows
        Write-Fact ("    ProtectedViewWindows.Count = {0}" -f $pv.Count)
        for ($i = 1; $i -le $pv.Count; $i++) {
            $w = $pv.Item($i)
            $caption = ''; $saved = ''
            try { $caption = $w.Caption } catch { $caption = "Caption failed: $($_.Exception.Message)" }
            try { $saved = [bool]$w.Document.Saved } catch { $saved = "Document.Saved failed: $($_.Exception.Message)" }
            Write-Fact ("      [{0}] |{1}|  Saved={2}" -f $i, $caption, $saved)
        }
    } catch {
        Write-Fact ("    ProtectedViewWindows unreachable: {0}" -f $_.Exception.Message)
    }

    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch { }
}
[GC]::Collect(); [GC]::WaitForPendingFinalizers()

if (-not $KeepOpen) {
    Write-Step 'Closing Word'
    Close-Word
}

Write-Host ''
Write-Host 'Probe complete - the Yellow lines are the measurements.' -ForegroundColor Green
