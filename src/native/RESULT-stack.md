# Slice: the stack - several Word windows as one window with tabs

**Date:** 2026-08-14
**Outcome: PASSED.** N real `OpusApp` windows are held at one rectangle, each drawing the same row
of tabs, and clicking a tab brings that document to the front. From the outside it is one Word
window with tabs.

Spike 2 proved this out of process on a 30ms poll and found the two things that make it hard
(`spikes\StackSpike\RESULT.md`). This slice does it from inside Word, where the drag problem it
could not solve simply does not arise.

## Proof

`tools\check-stack.ps1` opens three documents, reads the real rectangles from outside the process,
and asserts after every change. Three consecutive runs:

    PASS  23 checks, 0 failures

    0x90EE2  (892,180 1100x820)  strip y 267..315  _WwF y 315..774  "wordtab-stack-3.rtf ..."
    0x80E5E  (892,180 1100x820)  strip y 267..315  _WwF y 315..774  "wordtab-stack-2.rtf ..."
    0xB0ED6  (892,180 1100x820)  strip y 267..315  _WwF y 315..774  "wordtab-stack-1.rtf ..."
    0x70106  (892,180 1100x820)  strip y 267..315  _WwF y 315..774  "wordtab-check-1.rtf ..."

Four windows, one rectangle, one interior. Asserted after: stacking, moving the active window,
resizing it, clicking every tab, and closing a document. Plus a real caption drag through
`tools\drive-drag.ps1`: `dragged (858,146) -> (1058,226) delta (200,80)`, follower identical.

Clicking each of the four tabs brought a **different** window to the foreground - checked without
assuming any tab order, which is why it is evidence rather than a restatement of the code.

A screenshot of the result is the clearest statement of it: one Word window, two tabs between the
ribbon and the page, the active one lit, the document below it the one the active tab names.

## What it is

`src\native\stack.cpp`. The three files now split cleanly: `frames.cpp` is message plumbing,
`strip.cpp` is pixels and `_WwF` geometry, `stack.cpp` is policy - who is in the stack, where the
stack is, which tab is selected.

- **Membership is decided by what is true now**, on the strip's half-second janitor: visible, not
  minimised, and has a `_WwF`. Never keyed on window creation or destruction, because in Word those
  do not line up with documents - closing one of two documents was measured to hide one frame and
  destroy a *different* one. A Start-screen window has no `_WwF` and is deliberately left out.
  > **Correction, 2026-08-15.** The last sentence is false and was never measured. `_WwF` is the
  > document *frame* and Word keeps it, empty, after the last document closes — so a window with no
  > document *did* pass this test, joined the stack, and was given a tab labelled "Word". The test
  > now looks inside `_WwF`. See `RESULT-startscreen.md`.
- **Joining** snaps the window to the active window's rectangle, and to its maximized-or-not
  *state* rather than just its rectangle. Spike 2 could copy rectangles only, which left a stacked
  window looking maximized without being maximized.
- **Lockstep** happens in the active frame's `WM_WINDOWPOSCHANGING`: the others are moved to where
  it is *about* to be, in the same message, before it gets there. This is the whole reason for
  being in-process. Spike 2, polling at 30ms, could only ever be a frame behind, and its workaround
  was to hide the other windows for the duration of a drag.
- **The layout oracle.** Word lays out only the focused window, so the focused window's `_WwF` rect
  is copied onto every other window in the stack whenever Word changes it. Without this a
  background window keeps whatever interior it had, and you only find out when you switch to it.
- **Switching** is `SetWindowPos(HWND_TOP)` plus `SetForegroundWindow`. In-process, from a click in
  the foreground window, that is allowed - spike 2 needed an `AttachThreadInput` handshake to do
  the same thing from outside.
- **Every strip draws the stack's selected tab**, not its own. That is what makes a switch look
  like one strip standing still while the page changes, rather than two strips swapping.
- **Switchable** at `HKCU\Software\WordTab\Stack` (default on). Verified off: two windows stay where
  Word put them, each with its own single-tab strip.

The old `FollowDrag` demonstration in `frames.cpp` is gone - the stack does that properly now - and
with it the `FollowDrag` registry switch. The drag *measurement* stays, and still reports how many
windows moved in the same message.

## Two bugs this slice found, both worth keeping

**Closing a document put every strip on top of its ribbon.** The check said all four windows agreed
on their layout, and they did - they agreed on a wrong one. On the way out, Word dismantles a
frame's layout and leaves `_WwF` filling the whole client area, top edge at 0. We took that as a
natural rect, and because that frame was still the active one, the oracle broadcast it to every
other window in the stack. One document closing wrecked the layout of all of them.

Two changes: the strip **freezes while its window is hidden** - nobody can see a hidden window, so
there is nothing to gain by shifting it, and the janitor re-derives when it is shown again - and
the oracle refuses to broadcast from a window that is not visible.

The lesson for the checks matters as much: **consistency is not correctness.** Every window agreeing
passes every cross-window assertion. `check-stack.ps1` now also asserts each window in absolute
terms - strip flush under the chrome, document flush under the strip - which is what would have
caught this.

**Then the fix caused a double shift.** Re-deriving on hidden→shown read a document frame we had
already shifted and shifted it again, 32px lower each time. That is the accumulating-shift failure
the whole design exists to prevent, reintroduced by a path called "initial". `ApplyInitial` is now
idempotent like everything else - if the document frame is already where we put it, there is
nothing to derive - and `wasVisible` is seeded from the window's real state at bind time rather
than from zero, so a window that has been visible all along is not treated as having just appeared.

## Findings

**The strip can drift by a pixel.** Word and DWM occasionally shift the frame's client origin by
1px without moving `_WwF`, which leaves the strip a pixel out of line with the document frame it is
supposed to be flush against. Placement now compares against **where the strip actually is** rather
than where we last put it, and the janitor re-places it every half second, so any drift heals
within 500ms and is logged when it happens. Trusting our own bookkeeping was what let a visible
seam survive.

**Do not move a window that leaves the stack.** The first version put a departing window back where
it came from. Word hides and re-shows frames of its own accord, so documents jumped around the
screen for no reason a user could see. A hidden window is not being looked at, and when it comes
back the join snaps it to the stack again - so only `StackStop` puts windows back.

**Assert after the settle, not during it.** Word spends a second or two after a window appears
rearranging its own chrome, and the add-in corrects on a half-second cadence. Both check scripts
now wait, bounded, for the layout to settle before asserting - a measurement taken the instant a
`_WwF` exists is measuring the startup transient. Killing Word between test runs also makes it
restore documents on the next launch, which quietly changes how many windows a test is looking at.

**Word health across the whole session:** `LoadBehavior` still 3, `Resiliency\DisabledItems` empty,
no `Application Error` events for WINWORD, and the height tripwire never fired.

## Reproducing

    pwsh -File install\install.ps1 -NoBanner      # build, install, register (close Word first)
    pwsh -File tools\check-stack.ps1              # three documents, drives them, asserts, closes
    pwsh -File tools\check-stack.ps1 -Documents 3 -KeepOpen -Screenshot

By hand, which is the test that matters: open two or three documents. There is one Word window with
a tab for each. Click the tabs. Drag the window, resize it, maximize it - they stay one window.

## Not covered

- **The taskbar still shows one button per window.** `ITaskbarList::DeleteTab` on every window but
  the active one is what spike 2 used, and it worked cross-process; in-process it should be easier.
  That is the next slice.
- **Alt+Tab still lists every window**, a separate mechanism from the taskbar - `WS_EX_TOOLWINDOW`
  or DWM cloaking, neither tried.
- **Minimising** the active window does not take the stack with it.
- **No tab affordances**: no close button, no drag to reorder, no context menu, no new-tab button.
  Those are product questions.
- **Start-screen windows** (Word with no document) are excluded from the stack rather than given a
  policy.
- **Office Tab coexistence** and **a DPI change** are reasoned about in code and untested.
- Still untested on the work rig.
