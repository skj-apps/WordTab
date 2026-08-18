<#
.SYNOPSIS
  Build the current commit, wrap it into the one-file installer, and publish that file as a GitHub
  release.

.DESCRIPTION
  The download link in the README has to point at something, and that something has to be the current
  build. This is the one command that keeps both true: push, package, wrap, release, with the commit
  as the tag.

  Why it is a script and not three commands plus a browser upload: the file that has to reach the far
  machine is the one-file .cmd, and it is easy - it has happened here - to upload the ZIP instead. The
  zip is a correct package and a wrong delivery: it puts the person on the other end back to unzip,
  find README.txt, open a PowerShell prompt in the right folder. Uploading by hand is the step where
  the wrong file gets chosen, so uploading by hand is the step to remove.

  It refuses to publish a dirty tree. A release says "these bytes are that commit", and a dirty tree
  cannot support the claim - the commit in the file name would name something that was never built.

.PARAMETER Repo
  owner/name to publish to. Defaults to skj-apps/WordTab.

.PARAMETER Draft
  Publish as a draft, to look at before it becomes a download link.
#>
[CmdletBinding()]
param(
    [string]$Repo = 'skj-apps/WordTab',
    [switch]$Draft
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$RepoRoot = Split-Path -Parent $PSScriptRoot
$DistDir  = Join-Path $RepoRoot 'dist'

function Write-Step($text) { Write-Host "==> $text" -ForegroundColor Cyan }
function Write-Ok  ($text) { Write-Host "    $text" -ForegroundColor Green }
function Write-Note($text) { Write-Host "    $text" -ForegroundColor DarkGray }

# ---- the account, before anything is built -------------------------------------------------------
#
# Checked first and by name, because the failure it catches is silent in the worst way: a token that
# can READ a public repo perfectly well and cannot write to it. gh's own error at push time talks
# about the repository, not about which of two accounts you are, and the two accounts here differ by
# a hyphen. Better to say it before spending a build.

if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
    throw "The GitHub CLI (gh) is not on PATH. Install it, or build with make-onefile.ps1 and upload by hand."
}

Write-Step "Checking write access to $Repo"
$who = (& gh api user --jq '.login' 2>$null)
if ($LASTEXITCODE -ne 0) { throw "gh is not signed in. Run: gh auth login" }
$canPush = (& gh api "repos/$Repo" --jq '.permissions.push' 2>$null)
if ($LASTEXITCODE -ne 0) { throw "Cannot see $Repo as '$who'. Check the name, or that the account can reach it." }
if ("$canPush".Trim() -ne 'true') {
    throw @"
Signed in as '$who', which has read but not write access to $Repo.

Either sign in as the account that owns it:
    gh auth login          (then: gh auth switch)
or add '$who' to $Repo as a collaborator with write access.

Nothing has been built or published.
"@
}
Write-Ok "'$who' can write to $Repo"

# ---- the tree ------------------------------------------------------------------------------------

$dirty = (& git -C $RepoRoot status --porcelain)
if ($dirty) {
    Write-Host ''
    Write-Host 'The working tree has uncommitted changes:' -ForegroundColor Red
    $dirty | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
    throw "Commit or stash them first. A release names a commit, so the commit has to contain what is in the file."
}

$commit = (& git -C $RepoRoot rev-parse --short HEAD).Trim()
$branch = (& git -C $RepoRoot rev-parse --abbrev-ref HEAD).Trim()
$tag    = "build-$commit"
Write-Ok "commit $commit on $branch"

# The tag has to resolve on the far side, so the commit has to be there. A release pointing at a
# commit the remote has never seen is a download link with no source next to it.
Write-Step "Pushing $branch to origin"
& git -C $RepoRoot push --set-upstream origin $branch
if ($LASTEXITCODE -ne 0) { throw "git push failed. Nothing has been published." }

# ---- build ---------------------------------------------------------------------------------------

Write-Step 'Packaging'
& (Join-Path $PSScriptRoot 'package.ps1')

Write-Step 'Wrapping into one file'
& (Join-Path $PSScriptRoot 'make-onefile.ps1')

$installer = Join-Path $DistDir "WordTab-Install-$commit.cmd"
if (-not (Test-Path -LiteralPath $installer)) {
    throw "Expected $installer and it is not there. make-onefile.ps1 named it something else - do not upload a file whose build you cannot read off its name."
}
$zip = Get-ChildItem -Path $DistDir -Filter "WordTab-*-$commit.zip" -File |
       Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $zip) { throw "No package zip for $commit in $DistDir." }

# The hash of the DLL itself, read from the package rather than recomputed here. install.ps1 checks
# the payload against this same line, so publishing any other number would be publishing a second
# opinion about what the file is.
$payloadTxt = Join-Path $DistDir ([IO.Path]::GetFileNameWithoutExtension($zip.Name) + '\PAYLOAD.txt')
if (-not (Test-Path -LiteralPath $payloadTxt)) { throw "No PAYLOAD.txt beside the package at $payloadTxt." }
$manifest = Get-Content -Path $payloadTxt
$sha = (($manifest | Where-Object { $_ -like 'Sha256:*' } | Select-Object -First 1) -replace '^Sha256:\s*', '').Trim()
if (-not $sha) { throw "PAYLOAD.txt carries no Sha256: line." }

$kb = [Math]::Round((Get-Item $installer).Length / 1KB, 1)
Write-Ok "$(Split-Path $installer -Leaf)  ($kb KB)"

# ---- release -------------------------------------------------------------------------------------

$leaf  = Split-Path $installer -Leaf
$notes = @"
Several Word documents in one window, with a tab for each.

## Install

Download **$leaf** below, close Word, and double-click it. That is the whole procedure - the installer
carries the package inside itself, so there is nothing to unzip and no shell to open.

It installs for your user account only: no administrator rights, nothing written outside your profile,
no registry key outside ``HKEY_CURRENT_USER``. It appears in **Settings > Apps > Installed apps** as
``WordTab``, version ``$commit``, and uninstalls from there like anything else.

Then start Word and open two documents. They should share one window with a tab each.

## What this build is

| | |
|---|---|
| Source commit | ``$commit`` |
| WordTab.dll SHA-256 | ``$sha`` |

The ``.zip`` is the same payload with its parts visible, for anyone who would rather unpack it and
read ``install\install.ps1`` before running it.

The installer is unsigned self-built software, so SmartScreen or endpoint protection may block it on
sight. ``dist\WordTab-for-IT.md`` in the repo is a one-page description written to be forwarded to
whoever has to approve it.
"@

& gh release view $tag -R $Repo --json tagName 2>$null | Out-Null
if ($LASTEXITCODE -eq 0) {
    Write-Step "Release $tag exists - replacing its files"
    & gh release upload $tag $installer $zip.FullName -R $Repo --clobber
    $notes | & gh release edit $tag -R $Repo --notes-file -
} else {
    Write-Step "Creating release $tag"
    $ghArgs = @($tag, $installer, $zip.FullName, '-R', $Repo, '--title', "WordTab $commit", '--notes-file', '-')
    if ($Draft) { $ghArgs += '--draft' }
    $notes | & gh release create @ghArgs
}
if ($LASTEXITCODE -ne 0) { throw 'gh failed while publishing the release.' }

Write-Host ''
Write-Ok 'Published.'
Write-Note "  https://github.com/$Repo/releases/latest"
Write-Note '  On the target machine: open that page, download the .cmd, double-click it.'
