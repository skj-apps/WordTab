# The row's own size

**The complaint.** From the work rig, on `build-00f68f2`: *"also still opening up 4 pages wide and
word tab dialog appears every open change that. very annoying along w/ multiple pages opening."*

Two complaints, not three - "4 pages wide" and "multiple pages opening" are the same thing said
twice, and the dialog is separate (see below).

## What the log said

Their `settings.ps1 -Report` dump settled it without needing a guess. Every fresh Word start in it
begins the same way:

```
attach  hwnd=0x...  (existing)  thread=...  visible=0  rect=(418,418 3840x1536)  frames=1
```

That is **Word** restoring a frame about 3840px wide, on a 5120x2160 primary at 150%, **before the
add-in has touched anything**. At that width Word's own layout puts about four pages side by side.
The multi-page view is Word's, not our drawing, and it is not a WordTab defect.

What *is* ours is that it spreads. `Join` in `stack.cpp` snaps every window that joins to the first
one:

```
stack  hwnd=0x...  joined, snapped to 0x...  (2 in the stack)
```

and the live section of their report showed all three frames at exactly the same rectangle,
`(-8,668 1292x1418)`. So **whatever size window #1 happens to open at becomes the size of the whole
row.** Sizing the row fixed it for that session; the next start took Word's rectangle again.

## What was built

`RowSize` (on by default). The row records the size the user settles on and applies it to the window
that defines the row, over the top of whatever Word restored.

- **Recorded at `WM_EXITSIZEMOVE`** - the moment a person stops dragging - and only for the window
  that currently defines the row. Not from the intermediate rectangles a drag passes through, of
  which the trace shows dozens per drag.
- **`GetWindowPlacement().rcNormalPosition`, not `GetWindowRect`.** The restored rectangle survives
  underneath a maximized window, so maximizing the row does not overwrite the size chosen for it;
  the maximized state is stored beside it as a state. Minimised is refused outright - Windows parks
  a minimised window near -32000, which is the same reason `MatchTo` will not copy a minimised
  master's rectangle.
- **Applied in exactly one place**: the `if (!master)` branch of `Join`, the branch whose comment
  already said *"First window in: it defines where the stack is."* That is the only window whose
  size matters, because every other one is snapped to it.
- **Refused when unreachable.** `MonitorFromRect(..., MONITOR_DEFAULTTONULL)`. Monitors get
  unplugged and laptops get undocked, and a row restored onto a screen that no longer exists is a
  row the user cannot reach. Same rule `MatchTo` keeps for a minimised master; the log says so when
  it happens.
- **Four DWORDs, not a `REG_BINARY` blob** - `RowLeft` / `RowTop` / `RowRight` / `RowBottom` /
  `RowMaximized`. Every other value under the key is a readable DWORD that `settings.ps1` and regedit
  can both show. Negative coordinates survive as two's complement, which is not a corner case: their
  own row sat at `left = -8`.

**This does not explain why Word restores the rectangle it does, and does not need to.** That could
not be reproduced here - this machine has neither a 5120px screen nor a Word that restores wide - and
the fix writes over the answer rather than correcting it. Same position `tools\probe-view.ps1` is in.

## What was checked

`tools\check-rowsize.ps1`, new, **24 checks, 0 failures**, wired into `check-all.ps1` after `row`.

1. Nothing remembered: the add-in says nothing, moves nothing, and invents no rectangle.
2. A settled size is stored, and what is stored matches what `GetWindowPlacement` reports.
3. It is applied on the next start, and the window comes back at it.
4. A rectangle at `(-30000,-30000)` is refused, logged, and the window stays reachable.
5. `RowSize=0` turns the whole thing off, back to Word's rectangle.

**Two failures in the suite's first run were the suite's fault, and both are worth keeping.**

- `Wait-WordReady` returns when the **strip binds**. The stack **join** is a separate, later event -
  measured about half a second later - so asserting straight after `Wait-WordReady` reported "no
  window joined" while the log plainly showed one 500ms afterwards. The suite now waits for the join
  itself.
- `Close-AllWord` *asks*; it does not wait for the process to be gone. Opening the next document
  while the last WINWORD was still dying got the new one handed off and exited ~250ms after startup,
  so three steps ran against a window that never joined anything. The suite now waits for the
  process.

And the shape worth remembering: in that first run the **geometry** assertions passed while the
**log** assertions failed, because Word's own window memory happened to produce the same rectangle.
A geometry check alone would have reported success for a feature that had not run at all.

## Also in this slice: the load banner

Separate complaint, same report. `install.ps1` wrote `ShowLoadBanner` with an unconditional
`-Force` on every install:

```powershell
New-ItemProperty ... -Name 'ShowLoadBanner' -Value ([int](-not $NoBanner)) -PropertyType DWord -Force
```

The comment **directly above that line** sets out the rule it breaks - that `-Force` must never
rewrite a switch the user set, because this key is the user's. `ShowLoadBanner` was the one value in
it the installer overwrote, so a user who turned the dialog off got it back on the next build, and
the shipped one-click `.cmd` never passed `-NoBanner`.

Now: seeded when absent, left alone when present, `-NoBanner` / `-Banner` to say it explicitly, and
**the seed is off**. The banner earned its place when nothing else proved the add-in had loaded; the
installer's smoke test and the log both say so now, and a dialog on every Word start is a cost paid
forever for a fact learned once. A/B against the real installer, all four cases as expected: absent
seeds off, an existing 1 is preserved, `-Banner` overrides a 0, `-NoBanner` writes 0.

## Still open

- **Whether this actually fixes the complaint on their rig.** It cannot be reproduced here. The
  proof will be a fresh `-Report` showing `Row remembered` and a row that stays the width they set.
- **Office Tab is half-live on that machine** - `TabsforOfficeHelper.Helper` at `LoadBehavior=3`,
  `blocked=False` - and the two add-ins carve up the same document frame. They are uninstalling it.
  Until then, no UX complaint from that rig is cleanly attributable.
