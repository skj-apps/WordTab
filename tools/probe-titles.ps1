<#
.SYNOPSIS
  Measure what Word actually puts in an OpusApp title bar, for every document state that changes it.

.DESCRIPTION
  The tab name is the frame title with " - Word" removed, and that is the whole of it. This probe
  exists because the next slice rewrites that rule and the rule has to be written against measured
  strings rather than remembered ones:

    - what a plain .docx reads
    - whether the extension is shown, and for which formats (the whole "drop the extension?"
      decision rests on this, and on Explorer's HideFileExt, which is printed)
    - the exact shape of the compatibility-mode annotation: bracketed or dashed, and where
    - [Read-Only] and [Protected View], both provoked for real (file attribute, mark-of-the-web)
    - a never-saved document
    - **a document whose own name contains " - Word"**, which the current rule truncates at the
      first match. That fixture is here to fail.

  Every title is printed twice: once as text, once as codepoints, because a separator that is an
  en dash or a non-breaking space is invisible in the first form and fatal to a wcsstr.

  Fixtures are authored by Word itself (invisible, then quit) so the .doc and .docx are real files
  in real formats, not hand-rolled. Everything lands in %TEMP%\wordtab-titles and is disposable.

.PARAMETER KeepOpen
  Leave Word running with the fixtures open.
#>
[CmdletBinding()]
param(
    [switch]$KeepOpen,
    [switch]$ExtensionProbe
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$source = Get-Content -Raw -Path (Join-Path $PSScriptRoot 'WordLayout.cs')
Add-Type -TypeDefinition $source -Language CSharp -ReferencedAssemblies @(
    'System.Runtime', 'System.Collections', 'System.Threading.Thread', 'netstandard'
)
[WordLayout]::MakeDpiAware() | Out-Null

trap {
    Write-Host ''
    Write-Host "ERROR  $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
    break
}

function Write-Step($text) { Write-Host "==> $text" -ForegroundColor Cyan }
function Write-Note($text) { Write-Host "    $text" -ForegroundColor DarkGray }

function Get-WordPids { @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }) }
function Get-Frames {
    $found = @()
    foreach ($id in Get-WordPids) { $found += [WordLayout]::Frames($id) }
    return @($found)
}
function Get-FrameCount { return @(Get-Frames).Count }

# A title is only half-measured until you can see the bytes. Word's separator has been a plain
# hyphen-minus every time it has been looked at, but "every time it has been looked at" is a much
# weaker claim than a printed codepoint, and the rule about to be written is a string match.
function Show-Title($label, $text) {
    Write-Host ("    {0,-26} |{1}|" -f $label, $text) -ForegroundColor White
    $codes = ($text.ToCharArray() | ForEach-Object {
        $c = [int]$_
        if ($c -lt 32 -or $c -gt 126) { "U+{0:X4}" -f $c } else { [string]$_ }
    }) -join ''
    Write-Host ("    {0,-26}  {1}" -f '', $codes) -ForegroundColor DarkGray
}

function Close-Word {
    if (@(Get-WordPids).Count -eq 0) { return }
    for ($guard = 0; $guard -lt 24; $guard++) {
        $open = @(Get-Frames)
        if ($open.Count -eq 0) { break }
        [WordLayout]::Close($open[0])
        Start-Sleep -Milliseconds 1200
    }
    foreach ($p in @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue)) { $p.CloseMainWindow() | Out-Null }
    $deadline = (Get-Date).AddSeconds(20)
    while ((Get-Date) -lt $deadline -and @(Get-WordPids).Count -gt 0) { Start-Sleep -Milliseconds 500 }
    if (@(Get-WordPids).Count -gt 0) { throw 'Word would not close - close it by hand and re-run.' }
    Start-Sleep -Seconds 2
}

# ---- fixtures ----------------------------------------------------------------------------------

$dir = Join-Path $env:TEMP 'wordtab-titles'
if (Test-Path $dir) { Remove-Item -Path $dir -Recurse -Force }
New-Item -ItemType Directory -Path $dir -Force | Out-Null

Write-Step 'Word must be closed to author the fixtures'
Close-Word

# wdFormatXMLDocument = 12, wdFormatDocument97 = 0, wdFormatRTF = 6
$fixtures = @(
    @{ Name = 'Quarterly report';        Ext = 'docx'; Format = 12; Note = 'plain .docx - the control' }
    @{ Name = 'Quarterly report';        Ext = 'doc';  Format = 0;  Note = 'Word 97-2003 - expect compatibility mode' }
    @{ Name = 'Quarterly report';        Ext = 'rtf';  Format = 6;  Note = 'RTF - expect compatibility mode' }
    @{ Name = 'Locked report';           Ext = 'docx'; Format = 12; Note = 'read-only attribute - expect [Read-Only]' }
    @{ Name = 'Downloaded report';       Ext = 'docx'; Format = 12; Note = 'mark-of-the-web - expect [Protected View]' }
    @{ Name = 'Meeting - Wordsmith notes'; Ext = 'docx'; Format = 12; Note = 'name contains " - Word" - the current rule truncates this' }
)

Write-Step 'Authoring fixtures with Word itself (invisible)'
$word = New-Object -ComObject Word.Application
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    foreach ($f in $fixtures) {
        $path = Join-Path $dir "$($f.Name).$($f.Ext)"
        $doc = $word.Documents.Add()
        $doc.Content.Text = "WordTab title probe - $($f.Name).$($f.Ext)"
        # SaveAs2 with plain arguments: PowerShell 7's late-bound COM does not accept [ref] here,
        # and wraps the path in a PSObject that the marshaller then refuses.
        $doc.SaveAs2($path, [int]$f.Format)
        $doc.Close(0)
        $f.Path = $path
        Write-Note "$($f.Name).$($f.Ext)  -  $($f.Note)"
    }
} finally {
    $word.Quit(0)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
Close-Word

Write-Step 'Applying the two states that are not a file format'
$locked = @($fixtures | Where-Object { $_.Name -eq 'Locked report' })[0]
Set-ItemProperty -Path $locked.Path -Name IsReadOnly -Value $true
Write-Note "read-only attribute set on $($locked.Path)"

$web = @($fixtures | Where-Object { $_.Name -eq 'Downloaded report' })[0]
Set-Content -Path $web.Path -Stream 'Zone.Identifier' -Value "[ZoneTransfer]`r`nZoneId=3"
Write-Note "mark-of-the-web written to $($web.Path):Zone.Identifier"

$hideExt = (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name HideFileExt).HideFileExt
Write-Note "Explorer HideFileExt = $hideExt  (1 = extensions hidden in Explorer)"

# ---- read the titles ---------------------------------------------------------------------------

Write-Step 'Opening each fixture and reading its OpusApp title'
Write-Host ''

$seen = @{}
$results = @()

foreach ($f in $fixtures) {
    $before = @(Get-Frames)
    foreach ($h in $before) { $seen[$h] = $true }

    Start-Process -FilePath 'winword.exe' -ArgumentList "`"$($f.Path)`""
    $wait = if ($before.Count -eq 0) { 18 } else { 8 }
    Start-Sleep -Seconds $wait

    $deadline = (Get-Date).AddSeconds(25)
    $new = @()
    while ((Get-Date) -lt $deadline) {
        $new = @(Get-Frames | Where-Object { -not $seen.ContainsKey($_) })
        if ($new.Count -gt 0) { break }
        Start-Sleep -Milliseconds 500
    }

    if ($new.Count -eq 0) {
        Write-Host ("    {0,-26} <no new frame appeared>" -f "$($f.Name).$($f.Ext)") -ForegroundColor Yellow
        continue
    }

    Start-Sleep -Seconds 2
    $title = [WordLayout]::TitleOf($new[0])
    Show-Title "$($f.Name).$($f.Ext)" $title
    Write-Note $f.Note
    Write-Host ''

    $results += [pscustomobject]@{ File = "$($f.Name).$($f.Ext)"; Title = $title; Note = $f.Note }
}

Write-Step 'A document that has never been saved'
$before = @(Get-Frames)
foreach ($h in $before) { $seen[$h] = $true }
Start-Process -FilePath 'winword.exe' -ArgumentList '/w'
Start-Sleep -Seconds 8
$new = @(Get-Frames | Where-Object { -not $seen.ContainsKey($_) })
if ($new.Count -gt 0) {
    $title = [WordLayout]::TitleOf($new[0])
    Show-Title 'never saved' $title
    $results += [pscustomobject]@{ File = '(never saved)'; Title = $title; Note = 'Word''s own default name' }
} else {
    Write-Note 'no new frame - Word may have reused an existing window'
}

# ---- what the current rule makes of each one ----------------------------------------------------

Write-Host ''
Write-Step 'What the shipping rule (strip the first " - Word") makes of each title'
foreach ($r in $results) {
    $cut = $r.Title
    $at = $cut.IndexOf(' - Word')
    if ($at -ge 0) { $cut = $cut.Substring(0, $at) }
    if ($cut -eq '') { $cut = 'Word' }
    $flag = if ($cut -eq [System.IO.Path]::GetFileNameWithoutExtension($r.File) -or
                $cut -eq $r.File) { ' ' } else { '*' }
    Write-Host ("  {0} {1,-30} -> |{2}|" -f $flag, $r.File, $cut) -ForegroundColor $(if ($flag -eq '*') { 'Yellow' } else { 'Green' })
}
Write-Host ''
Write-Note '* = the tab does not read as the document name. Those are the cases the slice has to answer.'

# ---- does Word follow Explorer's "hide extensions for known file types"? -------------------------
#
# This decides whether the tab name can be relied on to carry an extension at all. If Word follows
# the shell setting then "keep the extension so a .doc and a .docx are told apart" is only true on
# machines where the user has extensions showing - which makes it a machine-dependent tab name, not
# a decision we get to take. Flipped for real and restored in the finally, the same way
# check-look.ps1 flips the app theme.

if ($ExtensionProbe) {
    Write-Host ''
    Write-Step 'Flipping Explorer HideFileExt to 1 and re-reading one title'
    Close-Word

    $advanced = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'
    $original = (Get-ItemProperty -Path $advanced -Name HideFileExt).HideFileExt
    try {
        Set-ItemProperty -Path $advanced -Name HideFileExt -Value 1
        # Word reads this through the shell, which caches it per process. Word was just closed, so a
        # fresh one picks it up; the broadcast is belt and braces for the shell's own cache.
        [WordLayout]::BroadcastSettingChange()
        Start-Sleep -Seconds 2

        $control = @($fixtures | Where-Object { $_.Ext -eq 'docx' -and $_.Name -eq 'Quarterly report' })[0]
        Start-Process -FilePath 'winword.exe' -ArgumentList "`"$($control.Path)`""
        Start-Sleep -Seconds 18
        $deadline = (Get-Date).AddSeconds(25)
        while ((Get-Date) -lt $deadline -and (Get-FrameCount) -lt 1) { Start-Sleep -Milliseconds 500 }

        $frames = @(Get-Frames)
        if ($frames.Count -gt 0) {
            Start-Sleep -Seconds 2
            Show-Title 'HideFileExt=1' ([WordLayout]::TitleOf($frames[0]))
            Write-Note 'Compare with |Quarterly report.docx - Word| above.'
        } else {
            Write-Note 'no frame appeared'
        }
        Close-Word
    } finally {
        Set-ItemProperty -Path $advanced -Name HideFileExt -Value $original
        [WordLayout]::BroadcastSettingChange()
        Write-Note "HideFileExt restored to $original"
    }
}

if (-not $KeepOpen) {
    Write-Host ''
    Write-Step 'Closing Word'
    Close-Word
    Write-Note "Fixtures left in $dir"
}
