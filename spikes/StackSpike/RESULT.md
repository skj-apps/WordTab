# Spike 2 result: PASS

Run on the dev rig, 2026-08-14. M365 x64, Word 16.0.20228.20190, 150% DPI, three
documents open. Office Tab was installed but **not loaded** into the Word process
(checked by module list, not by the add-in UI) so nothing competed for `_WwF`.

## Question

Spike 1 proved we can carve a 32px strip out of *one* Word window. That is a stripe,
not a product. The product claim is "several documents look like one window with tabs".
So: can N real `OpusApp` windows be stacked at one rect, held in geometric lockstep,
and switched by z-order — with a tab bar that appears continuous across the switch?

## Answer

Yes.

Three Word windows that started at `0,0`, `0,0` and `612,71` were snapped to a single
rect and held there. Measured live with the spike running:

```
                     window 0x51EF8   window 0x9042A   window 0xB2E62
outer rect           80,40 1020x700   80,40 1020x700   80,40 1020x700
ribbon (NetUIHWND)   y  41 ..  219    y  41 ..  219    y  41 ..  219
WordTabStackSpike    y 219 ..  251    y 219 ..  251    y 219 ..  251   <- ours
_WwF document frame  y 251 ..  710    y 251 ..  710    y 251 ..  710
```

Identical in all three, contiguous, no gap or overlap. `distinct outer rects: 1`.

Switching works end to end. A click on tab 2 (delivered as `WM_LBUTTONDOWN` to the
strip; see "Reproducing" for why not a real click) produced:

```
switch -> [1] "Document1"  foreground=True
```

and the z-order flipped so that window became topmost *and* took keyboard focus. No
reparenting, no content move — `SetWindowPos(HWND_TOP)` plus an `AttachThreadInput`
handshake for the foreground handoff, which Windows would otherwise refuse.

The tab bar renders in every window at the same place, so a switch looks like one
strip staying put rather than three strips swapping. That is the Office Tab trick,
reproduced.

## The finding that actually matters

**Word only lays out a window's interior while that window has focus.**

Force-resize a stacked window that is not focused and Word moves its NetUI chrome —
ribbon, status bar — but leaves `_WwF` completely alone. Measured: after resizing the
stack from 1095 to 900 wide, the background windows still had a **1081px-wide `_WwF`
inside an 885px client**, and activating them did not repair it. Nudging with
`SWP_FRAMECHANGED` and with a 1px resize both failed to trigger a relayout.

That is fatal for a stack, because by construction only one window is ever focused.
Waiting for Word to lay out the other N-1 windows means waiting forever.

The fix is to stop waiting: **the focused window is the layout oracle, and every other
stacked window copies its `_WwF` rect verbatim.** This is only sound because stacking
already guarantees all the windows are the same size, so the correct interior is the
same too. In the log, one window reports `[from Word]` and the rest report
`[copied from active]`:

```
switch -> [1] "Document1"  foreground=True
master rect -> (80,40 1020x700) from "Document1"
layout "Document1": natural (0,178 1005x491) -> (0,210 1005x459)  [from Word]
layout "Document7": natural (0,178 1005x491) -> (0,210 1005x459)  [copied from active]
layout "Document5": natural (0,178 1005x491) -> (0,210 1005x459)  [copied from active]
```

This is the multi-window analogue of spike 1's idempotence finding, and any
reimplementation needs both properties.

Two consequences that fall out of it:

- **Followers cannot derive their own natural rect**, so they cannot walk. Only the
  active window re-derives from Word, and it is guarded (below).
- **Restore has the same problem in reverse.** On exit we put each window back to its
  own size *first* and then refit `_WwF` to it, because Word will not do that either.

## Spike 1's rule was not quite enough

Spike 1 kept the strip stable by being idempotent against Word's layout: read `_WwF`,
and if it is not exactly where we last put it, treat what we see as natural and shift
from there.

That is idempotent against *Word*. It is not idempotent against **another process
doing the same thing**. Spike 2's first run had a stray spike 1 process still live on
one of the three windows from an earlier session; each saw the other's +32 as a fresh
natural rect and added another 32, and that window's document frame walked down to
**25px tall** within seconds. Office Tab shifts the same `_WwF` and would do this to us
on any machine where both are installed.

Two changes close it:

- The top edge is tracked separately. If `_WwF`'s top is still exactly where we put it,
  nobody moved it, and only the other edges are re-derived. This converges instead of
  racing, and it is also what makes a switch to a stale window behave.
- A height tripwire refuses to squeeze the frame below `2 * STRIP_H` and logs a warning.
  Better to leave the strip wrong than to squeeze the document out of existence.

## What this does NOT prove

- **Still out-of-process**, for the same reason as spike 1 — no COM add-in, no C++
  toolchain needed, which keeps the spike about stacking alone. The 30ms poll means a
  sub-frame window where geometry is stale.
- **Taskbar.** Three stacked windows still produce three taskbar buttons.
  `ITaskbarList::DeleteTab` on the inactive ones is untouched — that is the next slice.
- **Maximize/restore state.** Only the *rect* is propagated. A stacked window given a
  maximized window's rect looks right but is not actually maximized, so double-clicking
  its title bar behaves oddly. Not addressed.
- **Real mouse clicks on the tab bar.** `RealChildWindowFromPoint` and
  `ChildWindowFromPointEx` both resolve the tab bar's centre to our window, and it sits
  at child z-order index 0 (topmost) — so the plumbing is right. But the clicks in this
  run were posted, not physical, because another application was above Word in the
  global z-order on the test desktop. Worth one manual confirmation by hand.
- **Closing a document mid-session** was not exercised; only opening one was.

## Reproducing

```
dotnet build -c Release
bin\Release\net9.0-windows\win-x64\StackSpike.exe
```

Open two or three Word documents first. Click a tab, or press `1`..`9` in the console.
`q` or Ctrl+C restores every window. `--seconds N` exits cleanly after N seconds, which
is what makes the spike scriptable — Ctrl+C cannot be sent cross-process, and killing it
would leave every Word window stacked and shifted.

Check Office Tab is not loaded before judging any of this:

```powershell
(Get-Process WINWORD).Modules | Where-Object { $_.ModuleName -match 'Tabs|OfficeTab' }
```

Empty means the field is clear. The registry `LoadBehavior` is *not* a reliable check —
on this rig `OfficeTab.TabsforWord2013` reads `LoadBehavior=3` under both HKCU and HKLM
while the DLL is not in the process at all.
