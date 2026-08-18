<#
.SYNOPSIS
  Turn a built package into ONE double-clickable file.

.DESCRIPTION
  The problem, stated plainly: a .ps1 is not something Windows runs when you double-click it, and the
  way to run one is not guessable - you have to open a shell IN the extracted folder first. The
  person installing this is not a PowerShell user, is doing it once, after work, on a machine with no
  help available. "Unzip, then open a terminal, then paste this" is three steps and two of them are
  unfamiliar.

  So: one file. Double-click it. It installs.

  WordTab-Install.cmd is a .cmd file with the whole package base64'd onto the end of it. cmd.exe runs
  the few lines at the top and never reaches the payload, because every payload line starts with ::
  (a comment) and the header exits before them anyway.

  **It contains no installer of its own.** It unpacks to a temp folder and runs the packaged
  install.ps1 - the same verbatim copy the zip carries, which is the same file the repo uses. That is
  the whole point: a self-extracting installer that re-implemented registration would be a third
  implementation of "how WordTab is registered", and the three would agree until one was fixed. This
  one decides nothing. It carries bytes and hands off.

  The PowerShell bootstrap is passed as -EncodedCommand rather than inline, because the quoting rules
  of cmd.exe and PowerShell disagree about quotes, carets, percent signs and parentheses, and an
  installer that breaks on one machine's path is worse than no installer.

.PARAMETER Zip
  The package .zip to embed. Defaults to the newest in dist\.

.PARAMETER OutFile
  Where to write the .cmd. Defaults to dist\WordTab-Install-<commit>.cmd, taking the commit from the
  zip's own name so the build is legible without running it.
#>
[CmdletBinding()]
param(
    [string]$Zip,
    [string]$OutFile
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$RepoRoot = Split-Path -Parent $PSScriptRoot
$DistDir  = Join-Path $RepoRoot 'dist'

if (-not $Zip) {
    $newest = Get-ChildItem -Path $DistDir -Filter 'WordTab-*.zip' -File |
              Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $newest) { throw "No package zip in $DistDir. Run install\package.ps1 first." }
    $Zip = $newest.FullName
}
if (-not (Test-Path $Zip)) { throw "No such package: $Zip" }
if (-not $OutFile) {
    # Carry the commit into the file name. The zip is already named <date>-<commit>, and dropping that
    # on the way to the .cmd costs the one question that matters on a machine you cannot reach: is the
    # file sitting on it the current build? A fixed WordTab-Install.cmd cannot answer it - a stale copy
    # and a fresh one are the same name and the same icon - so the answer needed an install and a look
    # at Settings > Apps. Named this way it is answerable before running it, by reading it.
    $base = [IO.Path]::GetFileNameWithoutExtension($Zip)
    $OutFile = if ($base -match '^WordTab-\d{8}-(?<c>[0-9a-zA-Z]+)$') {
        Join-Path $DistDir "WordTab-Install-$($Matches.c).cmd"
    } else {
        Join-Path $DistDir 'WordTab-Install.cmd'
    }
}

Write-Host "==> Embedding $(Split-Path $Zip -Leaf)" -ForegroundColor Cyan

# ---- the bootstrap that runs on the far machine --------------------------------------------------
#
# Reads the .cmd it was launched from, takes everything after the marker, strips the :: that keeps
# cmd.exe from trying to execute it, and unpacks it. $env:WTSELF is set by the header.
$bootstrap = @'
$ErrorActionPreference = 'Stop'
try {
    $self = $env:WTSELF
    if (-not $self -or -not (Test-Path -LiteralPath $self)) { throw "Cannot find my own file." }

    $lines  = [System.IO.File]::ReadAllLines($self)
    $marker = [Array]::IndexOf($lines, '::WORDTAB-PAYLOAD::')
    if ($marker -lt 0) { throw "This file is damaged: the payload marker is missing. It was probably altered in transit - copy it again." }

    $sb = New-Object System.Text.StringBuilder
    for ($i = $marker + 1; $i -lt $lines.Length; $i++) {
        $line = $lines[$i]
        if ($line.Length -gt 2 -and $line.StartsWith('::')) { [void]$sb.Append($line.Substring(2)) }
    }
    if ($sb.Length -eq 0) { throw "This file is damaged: it carries no payload. Copy it again." }

    try { $bytes = [Convert]::FromBase64String($sb.ToString()) }
    catch { throw "This file is damaged: the payload will not decode. Some mail systems rewrite attachments - copy it again by USB or OneDrive." }

    $work = Join-Path $env:TEMP ("WordTab-setup-" + [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Force -Path $work | Out-Null
    try {
        $zip = Join-Path $work 'package.zip'
        [System.IO.File]::WriteAllBytes($zip, $bytes)

        # Expand-Archive is in Windows PowerShell 5.1, which is the shell this is guaranteed to find.
        Expand-Archive -LiteralPath $zip -DestinationPath $work -Force

        $installer = Join-Path $work 'install\install.ps1'
        if (-not (Test-Path -LiteralPath $installer)) { throw "The payload unpacked but install.ps1 is not in it." }

        # The real installer. Everything above this line only moved bytes around.
        & $installer
    }
    finally {
        Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
    }
}
catch {
    Write-Host ""
    Write-Host "WordTab could not be installed." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Write-Host "Nothing was changed on this machine." -ForegroundColor Gray
    exit 1
}
'@

$encoded = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($bootstrap))

# ---- the payload ---------------------------------------------------------------------------------

$b64 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($Zip))
$width = 240
$payloadLines = New-Object System.Collections.Generic.List[string]
for ($i = 0; $i -lt $b64.Length; $i += $width) {
    $payloadLines.Add('::' + $b64.Substring($i, [Math]::Min($width, $b64.Length - $i)))
}

# ---- the header ----------------------------------------------------------------------------------
#
# `exit /b` before the marker so cmd.exe never walks the payload at all. The :: prefix is belt and
# braces for the case where something reflows the file.
$header = @(
    '@echo off'
    'setlocal'
    'title WordTab setup'
    'echo.'
    'echo   WordTab - tabbed documents for Word'
    'echo.'
    'echo   Installing for your user account only.'
    'echo   No administrator rights are needed and nothing is written outside your profile.'
    'echo.'
    'set "WTSELF=%~f0"'
    ''
    ':: Hand Windows PowerShell a clean module path. If this file is launched from a PowerShell 7'
    ':: terminal rather than from Explorer, PSModulePath is INHERITED, and pwsh puts its own'
    ':: Microsoft.PowerShell.Utility ahead of the system one. Windows PowerShell 5.1 then loads the'
    ':: PS7 copy, cannot use it, and loses Get-FileHash - which install.ps1 needs to verify the'
    ':: payload. It fails with "Get-FileHash is not recognized", which reads as a broken installer'
    ':: rather than a borrowed variable. Measured, not guessed: MISSING inherited, FOUND cleared.'
    ':: Emptying it is right rather than setting a value - PowerShell rebuilds its own correct'
    ':: default when the variable is absent, and hardcoding paths here would be a third opinion'
    ':: about where modules live.'
    'set "PSModulePath="'
    ''
    ("powershell.exe -NoProfile -ExecutionPolicy Bypass -EncodedCommand {0}" -f $encoded)
    'echo.'
    'echo   ----------------------------------------------------------------'
    'echo   If the line above says  Installed.  then start Word and open two'
    'echo   documents - they should share one window with a tab each.'
    'echo.'
    'pause'
    'exit /b'
    '::WORDTAB-PAYLOAD::'
)

# ASCII with CRLF. cmd.exe is not a Unicode host, and a UTF-8 BOM on line 1 of a .cmd is executed as
# a command - it fails with a garbled first line that looks nothing like the cause.
$all = ($header + $payloadLines) -join "`r`n"
[System.IO.File]::WriteAllText($OutFile, $all + "`r`n", (New-Object System.Text.ASCIIEncoding))

$kb = [Math]::Round((Get-Item $OutFile).Length / 1KB, 1)
Write-Host "    $OutFile  ($kb KB)" -ForegroundColor Green
Write-Host "    payload: $(Split-Path $Zip -Leaf), $($payloadLines.Count) lines" -ForegroundColor DarkGray
Write-Host ""
Write-Host "Double-click it. That is the whole procedure." -ForegroundColor Cyan
