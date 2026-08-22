<#
.SYNOPSIS
  Build WordTab and lay it out as a package that installs on a machine with no compiler.

.DESCRIPTION
  The problem this exists for, stated plainly: install.ps1 builds from source, the target corporate
  rig has no C++ toolchain, and src\native\build\ is gitignored - so cloning this repo over there
  produces a tree that cannot be installed and does not say why until it is tried.

  Downloading a toolchain onto a locked-down machine is a worse ask than carrying 200KB, so the
  answer is to carry the 200KB. The package is this repo with the compiler-shaped hole filled in:

      WordTab-<date>-<commit>\
        PAYLOAD.txt                     which build this is, and the hash to prove it
        README.txt                      what to do, in order
        install\install.ps1             VERBATIM COPY - not a second installer
        install\uninstall.ps1           VERBATIM COPY
        install\settings.ps1           VERBATIM COPY
        install\common.ps1             VERBATIM COPY - the other two dot-source it
        src\native\wordtab.h            the CLSID that install.ps1 checks itself against
        src\native\build\WordTab.dll    the build

  **The scripts are copied, never rewritten.** A package with its own installer would be a second
  implementation of "how WordTab is registered", and the two would agree right up until the day one
  of them was fixed. install.ps1 recognises a package by its PAYLOAD.txt and skips the build.

  Everything the package does on the far machine is per-user: HKCU and %LOCALAPPDATA%, no admin.

.PARAMETER SkipBuild
  Package whatever is already in src\native\build. The manifest still records its real hash.

.PARAMETER OutDir
  Where to put the package. Defaults to dist\ beside the repo root (gitignored).

.PARAMETER NoZip
  Leave the folder without also producing a .zip.

.EXAMPLE
  pwsh -File install\package.ps1
#>
[CmdletBinding()]
param(
    [switch]$SkipBuild,
    [string]$OutDir,
    [switch]$NoZip
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$RepoRoot  = Split-Path -Parent $PSScriptRoot
$NativeDir = Join-Path $RepoRoot 'src\native'
$BuiltDll  = Join-Path $NativeDir 'build\WordTab.dll'
if (-not $OutDir) { $OutDir = Join-Path $RepoRoot 'dist' }

function Write-Step($text) { Write-Host "==> $text" -ForegroundColor Cyan }
function Write-Ok  ($text) { Write-Host "    $text" -ForegroundColor Green }
function Write-Note($text) { Write-Host "    $text" -ForegroundColor DarkGray }

# ---- what exactly is being packaged --------------------------------------------------------------

Write-Step 'Identifying the build'

$commit = 'unknown'
$dirty  = $false
try {
    $commit = (& git -C $RepoRoot rev-parse --short HEAD 2>$null)
    $status = @(& git -C $RepoRoot status --porcelain 2>$null)
    $dirty  = ($status.Count -gt 0)
} catch { }

# A package is a thing that leaves the machine and comes back as a bug report. "Which build is that?"
# has to have an answer, and "the tip of main, probably" is not one.
if ($dirty) {
    Write-Warning "The working tree has uncommitted changes. The package will be labelled ${commit}+dirty, and nobody will be able to reconstruct it from that label."
    $commit = "$commit+dirty"
}

if (-not $SkipBuild) {
    Write-Step 'Building'
    & (Join-Path $NativeDir 'build.ps1')
    if ($LASTEXITCODE -ne 0) { throw "Build failed (exit $LASTEXITCODE)." }
} else {
    Write-Step 'Skipping build (-SkipBuild)'
}

if (-not (Test-Path $BuiltDll)) { throw "No build output at $BuiltDll" }

# The staleness check install.ps1 does in a source tree, done here instead - because the package is
# where it stops being checkable. Past this point the DLL travels without its sources.
$newestSource = Get-ChildItem -Path $NativeDir -Include '*.cpp', '*.h', '*.def' -File -Recurse |
                Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($newestSource -and (Get-Item $BuiltDll).LastWriteTime -lt $newestSource.LastWriteTime) {
    throw "WordTab.dll is older than $($newestSource.Name). Packaging that would ship a build nobody can identify as stale once it is on another machine. Build it first (drop -SkipBuild)."
}

$sha256 = (Get-FileHash -Path $BuiltDll -Algorithm SHA256).Hash
$md5    = (Get-FileHash -Path $BuiltDll -Algorithm MD5).Hash
$size   = (Get-Item $BuiltDll).Length
$stamp  = Get-Date -Format 'yyyyMMdd'

Write-Ok "commit $commit"
Write-Ok ("WordTab.dll  {0:N0} bytes  SHA256 {1}" -f $size, $sha256)
Write-Note "MD5 $md5"

# ---- lay it out -----------------------------------------------------------------------------------

$name    = "WordTab-$stamp-$($commit -replace '[^0-9a-zA-Z]', '')"
$payload = Join-Path $OutDir $name

Write-Step "Assembling $payload"
if (Test-Path $payload) { Remove-Item -Path $payload -Recurse -Force }
New-Item -ItemType Directory -Path (Join-Path $payload 'install')          -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $payload 'src\native\build') -Force | Out-Null

Copy-Item (Join-Path $PSScriptRoot 'install.ps1')   (Join-Path $payload 'install\install.ps1')
Copy-Item (Join-Path $PSScriptRoot 'uninstall.ps1') (Join-Path $payload 'install\uninstall.ps1')
# The escape hatch travels with the thing it is an escape from. A package whose tab row comes out the
# wrong colour on a machine nobody here has seen is exactly the case settings.ps1 exists for, and
# leaving it behind would mean the answer to "turn that bit off" was regedit.
Copy-Item (Join-Path $PSScriptRoot 'settings.ps1')  (Join-Path $payload 'install\settings.ps1')
# Both of the above dot-source this and refuse to run without it. It is not optional payload: it
# holds the one implementation of "has Word disabled this add-in" and "which rivals will load".
Copy-Item (Join-Path $PSScriptRoot 'common.ps1')    (Join-Path $payload 'install\common.ps1')
Copy-Item (Join-Path $NativeDir 'wordtab.h')        (Join-Path $payload 'src\native\wordtab.h')
Copy-Item $BuiltDll                                 (Join-Path $payload 'src\native\build\WordTab.dll')
Write-Ok 'install.ps1, uninstall.ps1, settings.ps1 and common.ps1 copied verbatim'

# The Word this was built and checked against. A package that turns up on a machine with a different
# Word is still worth trying, but the difference is the first thing to look at.
$wordVersion = 'unknown'
try {
    $exe = Join-Path $env:ProgramFiles 'Microsoft Office\root\Office16\WINWORD.EXE'
    if (Test-Path $exe) { $wordVersion = (Get-Item $exe).VersionInfo.ProductVersion }
} catch { }

@"
WordTab package
Commit: $commit
Built:  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Sha256: $sha256
Md5:    $md5
Bytes:  $size
BuiltOnWord: $wordVersion

install.ps1 reads the Sha256 line and refuses to install a DLL that does not match it. That check
replaces the source-timestamp check it does in a repo, which cannot work here: a package has no
sources, and a zip, an email or a git clone rewrites file times anyway.
"@ | Set-Content -Path (Join-Path $payload 'PAYLOAD.txt') -Encoding ASCII

# ---- the note for whoever runs it ------------------------------------------------------------------

@"
WordTab - tabbed documents for Word
Package $name  (commit $commit)

WHAT THIS IS
  Several Word documents in one window with a tab for each, instead of one window per document.
  It installs entirely under your own user account: no admin rights, no UAC prompt, nothing written
  to Program Files, HKLM, or the machine certificate store.

    binaries    %LOCALAPPDATA%\Programs\WordTab\
    settings    HKCU\Software\WordTab
    log         %LOCALAPPDATA%\WordTab\wordtab.log

WHAT TO DO, IN ORDER

  1. CLOSE WORD. All of it. The installer refuses to run while any WINWORD process is up, because
     Word holds the installed DLL open.

  2. Open PowerShell - 64-BIT, which is what you get by default. The installer refuses to run from a
     32-bit host, because that would register into a view 64-bit Word never reads.

  3. cd into this folder and run:

         powershell -ExecutionPolicy Bypass -File install\install.ps1

     (Windows PowerShell 5.1 is fine. So is pwsh 7 if you have it.)

     It will print what it did, then activate the class outside Word as a smoke test. If that smoke
     test fails, Word will not load it either, and the message says so.

  4. Start Word and open two documents.

     EXPECTED: one Word window with two tabs, each named after its document.
     EXPECTED also: no dialog. The installer's own smoke test and the log are what prove it
     loaded. If you want Word to say so on every start, install with -Banner, or turn it on
     later with:

         powershell -ExecutionPolicy Bypass -File install\settings.ps1 -Set ShowLoadBanner=1

WHAT YOU CAN DO WITH IT

  Click a tab to switch. Drag one along the row to reorder it. Drag one DOWN, clear of the row, and
  let go - that document comes out into a window of its own; drag its tab back onto the row to put it
  back. Right-click a tab for Save, Close, Close Others, Close Tabs to the Right, Close All, and the
  same two Move commands. The x closes a tab, middle-click does too, + makes a new document, and
  Ctrl+Tab / Ctrl+Shift+Tab step along the row. A document with unsaved changes shows a dot where its
  x is. Rest the pointer on a tab and it tells you the document's full name and the folder it is in -
  which is what you want when there are enough documents open that the names get cut short.

  Size the window once and the row comes back that size the next time Word starts, rather than
  whatever size Word restores its own window to - which on a wide screen is what leaves you
  looking at several pages side by side. Before you have sized it even once, a window Word opens
  three or more pages wide is narrowed to about one page, so the first start is not wide either.
  Whatever size you give it afterwards wins from then on. Turn the whole thing off with RowSize=0
  if you would rather Word decided.

IF YOU HAVE TEN MINUTES, TRY THESE IN THIS ORDER

  They are ordered by how likely each is to be wrong on a machine this has never run on, not by how
  interesting they are. Everything above the line has been driven hundreds of times on the machine
  it was built on; everything below it has not been possible to test there at all.

  1. Open two documents.            One window, two tabs, one taskbar button.
  2. Click between the tabs.        Instant, no flicker, the ribbon stays put.
  3. Open six or seven more.        The row starts scrolling; the chevrons beside + step along it,
                                    and holding one scrolls continuously.
  4. Drag a tab along the row.      It reorders, and the tab follows your pointer.
  5. Drag a tab downward, clear     It comes out into a window of its own, and the tab travels with
     of the row, and let go.        your pointer while it is out. Drag it back onto the row to
                                    put it back, or right-click it > Move Back to the Tab Row.
  6. Ctrl+Tab / Ctrl+Shift+Tab.     Steps along the row and wraps at the ends.
  7. Type in a document.            A dot replaces the x on that tab until you save.
  8. Rest the pointer on a tab.     The document's full name and the folder it is in.
  9. Open something from email      Word's Protected View. The tabs should stay the same colour as
     or a download.                 the others and the strip should be the same height.
  10. Close every document.         The row goes empty rather than showing a tab for nothing.

  ---- and these have NEVER been tested, because the machine this was built on cannot ----

  11. Move the Word window to your  Everything above should still be true, at that screen's own
      OTHER monitor.                scaling. THIS IS THE MOST LIKELY THING TO BE WRONG. If the tab
                                    row is the wrong height, the wrong size, or in the wrong place
                                    on one screen and right on the other, that is the reason.
  12. Maximize it on each screen.   The row should span the window and sit directly above the page.
  13. Drag a tab out while Word is  The new window should land on the SAME screen, not jump to the
      on the second screen.         other one.

  If any of these is wrong, run the report below and send it. Say which screen Word was on.

IF YOU WANT TO TURN SOMETHING OFF

  Every part of it is a separate switch, and none of this needs regedit:

         powershell -ExecutionPolicy Bypass -File install\settings.ps1

  That lists every setting, what it does, and what the add-in reported at the last Word startup.
  To change one:  -Set TabDot=0    To put everything back:  -Reset
  Close Word completely afterwards - it reads these once, when it starts.

  The two most likely to be wanted:
    TabThemeSample=0   if the tab row is the wrong colour on this machine. It stops taking colours
                       off Word's ribbon and uses the Office theme setting instead. TRY THIS FIRST.
    TabKeys=0          if you need Ctrl+Tab to type a tab character inside a table.

IF SOMETHING IS WRONG

  PRESS START, TYPE  WordTab , OPEN  "WordTab Report"  - AND SEND THE FILE IT WRITES.

  The installer puts it there. There is nothing to find and no shell to open. From this folder,
  if you still have it, the same thing is:

         powershell -ExecutionPolicy Bypass -File install\settings.ps1 -Report

  It writes one text file to your Desktop with everything needed to work out what happened: which
  build is installed, whether Word registered it, whether Word DISABLED it, whether another tabbed-
  Word add-in is set to load, your Word version, your screen, every setting, and the add-in's own
  log in full. It reads only - nothing is installed, changed or removed, and no document is opened.

  It does include the names of documents you have had open, because those appear in the log as tab
  names. Read it before you send it if that matters.

  The log on its own is at %LOCALAPPDATA%\WordTab\wordtab.log if you would rather send just that.

  Nothing in the log at all means Word never loaded the add-in. Check, in this order:
    - HKCU\Software\Microsoft\Office\Word\Addins\WordTab.Connect  - LoadBehavior should be 3.
      Word rewrites it to 2 when a load fails, so a 2 there means it tried and gave up.
    - File > Options > Add-ins > Manage: Disabled Items. Word parks add-ins there after a crash and
      that beats LoadBehavior=3.

IF YOU ALREADY RUN OFFICE TAB (OR ANOTHER TABBED-WORD ADD-IN)

  TURN IT OFF BEFORE YOU START WORD. WordTab has never been run beside a working one, and this is
  the honest version of a claim that used to sit here saying they coexist fine. They do not conflict
  in theory - they both carve the strip out of the same document frame and both stack Word's windows
  on top of each other, which is one job done twice.

  It is not known to be dangerous, and it is one command each way. The installer prints the exact
  line for whatever it finds registered. To put the other one back afterwards, set LoadBehavior to 3
  again, or re-tick it in File > Options > Add-ins > Manage: COM Add-ins.

  If you would rather try both at once: open two documents and look at the tab row. One row of tabs
  is fine. Two rows, a row that flickers, or a document that jumps up and down means turn one off.

TO REMOVE IT

  Settings > Apps > Installed apps > WordTab > Uninstall. It is listed there like every other
  program on the machine, and removing it needs no administrator rights either.

  Or, if you still have this folder:

         powershell -ExecutionPolicy Bypass -File install\uninstall.ps1

  Either way removes the registration, the files, the settings and the log directory. Word then
  starts exactly as it did before - no strip, and the document back flush under the ribbon. Add
  -KeepLog to the command above to keep the log for diagnosis.
"@ | Set-Content -Path (Join-Path $payload 'README.txt') -Encoding ASCII

Write-Ok 'PAYLOAD.txt and README.txt written'

# ---- zip -------------------------------------------------------------------------------------------

if (-not $NoZip) {
    $zip = "$payload.zip"
    if (Test-Path $zip) { Remove-Item $zip -Force }
    Compress-Archive -Path (Join-Path $payload '*') -DestinationPath $zip
    Write-Ok ("{0}  ({1:N0} bytes)" -f $zip, (Get-Item $zip).Length)
}

Write-Host ''
Write-Host 'Packaged.' -ForegroundColor Green
Write-Host "  Folder: $payload" -ForegroundColor Gray
if (-not $NoZip) { Write-Host "  Zip:    $payload.zip" -ForegroundColor Gray }
Write-Host '  Copy it to the target machine, then follow README.txt.' -ForegroundColor Gray
