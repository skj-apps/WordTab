# WordTab

**Several Word documents in one window, with a tab for each.**

Open two or three documents and Word becomes a single window with a row of tabs across the top —
each named after its document, in Word's own colours. Click a tab to switch. Drag one to reorder the
row, or drag it clear of the row to pull that document out into a window of its own; drag it back to
put it in again. Close a tab with its ×, or middle-click it. Press **+** for a new document.
Ctrl+Tab and Ctrl+Shift+Tab step along the row. A document with unsaved changes shows a dot where its
× is. Rest the pointer on a tab and it gives you the document's full name and the folder it is in —
which is the thing you want once there are enough documents open for the names to be cut short. The
whole stack drags, resizes, maximizes and minimises as one window, with one taskbar button and one
Alt+Tab entry.

## Why it exists

Office Tab does this and costs money per seat, and the machine this was written for is a **corporate
Windows PC with no administrator rights** — no MSI, nothing written outside the user's own profile,
no registry key outside `HKCU`. Everything here obeys that: the add-in is a single DLL under
`%LOCALAPPDATA%\Programs\WordTab`, registered entirely under `HKCU`, installed and removed by a
PowerShell script that needs no elevation.

That constraint also decided the language. A .NET COM server **cannot** be activated from `HKCU` —
measured, not assumed — so the add-in is native C++ with no framework dependency.

## Installing

On a machine with the toolchain (see *Building*):

```
pwsh -File install\install.ps1 -NoBanner
```

On a machine with **no compiler** — which is the point — build a package on a machine that has one:

```
pwsh -File install\package.ps1
```

That writes a ~113KB zip to `dist\`. Copy it to the target machine by any means, unzip it, close
Word, and from a PowerShell prompt in that folder:

```
powershell -ExecutionPolicy Bypass -File install\install.ps1
```

It verifies the DLL against a SHA256 recorded in the package, clears the mark-of-the-web that a
downloaded zip leaves on it, registers the add-in, and checks the class actually activates. Then
start Word and open two documents.

Removing it leaves nothing behind:

```
powershell -ExecutionPolicy Bypass -File install\uninstall.ps1
```

## Turning parts of it off

Every part of WordTab can be switched off on its own, and none of it needs regedit:

```
powershell -ExecutionPolicy Bypass -File install\settings.ps1              # show everything
powershell -ExecutionPolicy Bypass -File install\settings.ps1 -Set TabDot=0
powershell -ExecutionPolicy Bypass -File install\settings.ps1 -Reset       # back to defaults
```

The settings are DWORDs under `HKCU\Software\WordTab`; absent means on. Word reads them once at
startup, so close Word completely after changing one. The two most likely to be wanted:

- **`TabThemeSample=0`** — if the tab row is the wrong colour on a machine this has not run on. It
  stops sampling Word's ribbon and uses the Office theme setting instead.
- **`TabKeys=0`** — if you need Ctrl+Tab to type a tab character inside a table, which is the one
  thing taking that chord costs.

`settings.ps1` also prints what the add-in itself reported at the last Word startup. If that
disagrees with the table, believe the add-in.

## How it works

WordTab is an in-process COM add-in. Word gives every document its own top-level `OpusApp` window;
WordTab keeps them all at one rectangle with one in front, hides the others from the taskbar and
Alt+Tab, and carves a 32px band out of the top of each window's document frame to draw the tab row
in. The row is the same in every window, so switching tabs looks like a strip standing still while
the page behind it changes.

The band is taken by subclassing Word's `_WwF` document frame and rewriting the position Word is
*proposing* in `WM_WINDOWPOSCHANGING`, rather than moving it afterwards — so Word never puts it in
the wrong place and there is nothing to correct. Membership of the stack is re-derived twice a second
from what is true about each window, which is what makes it robust against Word hiding, showing and
recycling frames without telling anyone.

There is no configuration UI, no network access, no telemetry, and nothing is written outside
`%LOCALAPPDATA%\WordTab` and `HKCU\Software\WordTab`.

## Building

**w64devkit** (portable MinGW-w64, GCC 16) unzipped to `%LOCALAPPDATA%\Programs\w64devkit\` — no
installer, no admin rights. `src\native\build.ps1` prints setup instructions if it is missing.

```
pwsh -File src\native\build.ps1
```

The build is byte-for-byte reproducible: the same sources give the same DLL hash.

## Checking it

Eleven suites drive a real Word with real mouse and keyboard input and assert what happens — about
560 checks, roughly fifteen minutes.

```
pwsh -File tools\check-all.ps1 *>&1 | Tee-Object -FilePath "$env:TEMP\battery.txt"
pwsh -File tools\check-all.ps1 -Only reorder,menu
pwsh -File tools\WordTabHarness.ps1 -SelfTest        # 23 checks, one second, no Word
```

**They cannot run in parallel and they take over the desktop** — one Word, one mouse pointer — so a
hand on the mouse will break a run. Each suite keeps its own console output and the add-in's log in
`%LOCALAPPDATA%\WordTab\history\`, which is where to look when one goes red.

`tools\probe-*.ps1` are measurements rather than tests: what Word actually does in some situation,
kept as evidence for the decision that came out of it.

## Where the reasoning is

Every slice of work left a writeup in `src\native\RESULT-*.md` — what was built, what was measured,
which decisions were made and why, and the mistakes worth not repeating. They are the real
documentation of this project; the code comments point at them. `NOTE-no-icons.md` records a feature
that was built, measured and deliberately scrapped.

## Known limitations

- Dragging a tab out of the row has no live preview — the pointer changes shape and the row lets go
  of the tab, but no window follows the hand until you release.
- Two windows that are both standing on their own cannot be dropped onto each other to form a stack;
  the target has to be a window already in one.
- A Word that has only ever held Protected View documents keeps the fallback palette until an
  ordinary document is in front, because a Protected View window has no ribbon body to sample.
- Multi-monitor behaviour is unproven. Per-window DPI changes are handled, but this has only ever run
  on a single-monitor machine.
