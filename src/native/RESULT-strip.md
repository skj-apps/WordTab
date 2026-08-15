# Slice: the strip, in-process

**Date:** 2026-08-14
**Outcome: PASSED.** WordTab now owns a horizontal band of Word's own layout, carved out from
inside Word, and keeps it through every layout change Word makes - including Backstage.

Spike 1 proved this geometry was possible from *outside* Word, on a 30ms poll, which meant every
relayout Word made was briefly visible before we corrected it. This is the same geometry with the
poll removed, and it is a different mechanism rather than the same one moved: we no longer watch
`_WwF` and fix it afterwards, we subclass `_WwF` and rewrite the position Word is *proposing*, in
`WM_WINDOWPOSCHANGING`, before the move happens. There is no correction because there is never
anything to correct.

## Proof

`tools\check-strip.ps1` starts Word on a scratch document, reads the real rectangles from outside
the process, and asserts after every change that the strip sits immediately above the document
frame, spans it exactly, and leaves no gap to the chrome above.

    PASS  42 checks, 0 failures

The baseline stack, measured live at 150% DPI:

    MsoCommandBarDock  y    0 .. 267   h=267     the ribbon
    WordTabStrip       y  267 .. 315   h= 48     <- ours (32 logical px at 150%)
    _WwF               y  315 .. 1071  h=756     the document
    MsoCommandBarDock  y 1071 .. 1104  h= 33     the status bar

Contiguous, no gap, no overlap - the same stack Office Tab produces. Held through:

- three programmatic resizes, maximize, and restore
- **a real resize drag**, both directions, 40 injected mouse movements each: Word's modal loop,
  one `_WwF` relayout per frame, our code between Word and its own window procedure for every one
- **Backstage**, opened for real and confirmed open before anything was asserted about it

`spikes\StripSpike\bin\...\spike.log` and `%LOCALAPPDATA%\WordTab\wordtab.log` hold the raw lines.

## Backstage, which is the case that mattered

Backstage is a `FullpageUIHost` child of the frame covering the whole client area, and **Word hides
every other child while it is up, ours included**. So the strip disappears when the File menu opens
and comes back in the right place when it closes, without us doing anything - the screenshot taken
inside Backstage shows the File menu completely clean, no strip floating over it.

This was worth the trouble to test properly. The first version of the check reported the Backstage
case as passing when Backstage had **never opened** - injected `Alt+F` failed silently and the
script measured an unchanged window. Two findings came out of fixing it, both now encoded in
`tools\WordLayout.cs`:

- **`Alt`-down and `F`-down sent back to back does nothing.** Alt must be held, then a pause, then
  F. Pressing and releasing Alt first (the KeyTips route) does nothing either.
- **Being the foreground window is not the same as having been clicked in.** After the window is
  activated programmatically, Word's ribbon ignores injected keystrokes until a real click lands in
  the document.

The check now confirms `FullpageUIHost` is up before asserting anything, and fails loudly if it is
not. A test that silently tests nothing is worse than no test.

## What it is

`src\native\strip.cpp`, plus four call sites in `frames.cpp`: `StripStart` in `FramesStart`,
`StripAttachFrame` in `AttachFrame`, `StripDetachFrame` in `DetachFrame`, `StripStop` in
`FramesStop`. One strip per `OpusApp` frame; it lives and dies with the frame.

- **Binding.** Each subclassed frame is searched for its `_WwF` child. It does not always exist yet
  - a Start-screen window has none at all - so a 500ms janitor timer retries, and rebinds if Word
  replaces the window. The geometry itself is entirely event-driven; the timer only discovers.
- **The shift.** `_WwF` is subclassed with comctl32's `SetWindowSubclass`, and `WM_WINDOWPOSCHANGING`
  rewrites the proposed rectangle: top down by the strip height, height reduced to match, then
  `SWP_NOMOVE`/`SWP_NOSIZE` cleared so the new values are honoured. Every `SetWindowPos` and
  `MoveWindow` on that window, from anywhere in the process, passes through this first.
- **The strip window** is a `WS_EX_NOACTIVATE` child of the frame, painted rather than composed of
  controls, showing one placeholder tab with the document's name. `HTTRANSPARENT`: it swallows no
  clicks, because tab hit-testing belongs with the real tab strip.
- **Restore.** On detach and on shutdown, `_WwF` goes back to Word's own rectangle.
- **Switchable** at `HKCU\Software\WordTab\Strip` (default on), no rebuild needed. Verified with it
  off: no strip child, and `_WwF` back at y=267 flush against the ribbon - Word's layout untouched.

## The two rules from the spikes, and where they live now

**Idempotent against Word's layout.** Word rewrites `_WwF` to its own natural rect repeatedly - six
times in one logged spike 1 session - and a design that nudges by -32 each time walks the document
frame off the bottom of the window in seconds. That is the old Doc Tabs failure mode. What Word
proposes is therefore treated as natural, and the shift is computed from it, never accumulated onto
it. `AdjustProposed` returns immediately if the proposal is already exactly where we put it.

**Idempotent against Word is not idempotent against another add-in.** Office Tab shifts this same
`_WwF`; spike 2 measured two processes each reading the other's +32 as a fresh natural rect and
squeezing the document to 25px tall. Two defences carried over: if the top edge is exactly where we
left it, nobody re-laid it out, so the top is kept and only the other edges are adopted; and a
tripwire refuses any shift that would leave the document less than two strips tall, logging once
rather than per message. Neither has been exercised against a real Office Tab install - it is not
installed on this rig - so they are precautions, not measurements.

## Findings

**The measurement must not perturb what it measures - again.** Word relayouts `_WwF` on every frame
of a resize drag, so an unthrottled log line per relayout is a file write every ~15ms inside Word's
modal loop. Relayout logging is throttled to one line per 250ms per frame, and the suppressed count
is carried into the next line that does get written: `[+11 more not logged]`. Measured across an
80-frame resize drag: **6 log lines**. The alternative was either a stutter the user can feel or a
silence that hides how often Word relayouts.

**Word's drop shadow now overlaps us.** A `DropShadow` child sits at y 267..276, which was over the
document before and is over the top 9px of our strip now. Harmless at present, and worth
remembering when the strip is drawn properly.

**DPI is per-window and the strip must scale with it.** 32 hard-coded pixels is two thirds of the
right height on this 150% rig. `GetDpiForWindow` is reached through `GetProcAddress` rather than a
header, because w64devkit's headers are older than the API. `WM_DPICHANGED` unwinds the old shift
and re-applies at the new scale - **untested**: this rig has one monitor.

**The strip follows Word's theme, minimally.** A light grey band across a black Word window reads
as a failure even when the geometry under it is perfect. Office's `UI Theme` value picks a light or
dark palette, falling back to Windows' `AppsUseLightTheme` when it says "use system setting". This
is not a design decision - what the tabs actually look like is a product question for later.

**No regression in the previous slice.** With strips on both frames, `tools\drive-drag.ps1` still
reports `moved 1 other frame(s) on 40 of 41 updates`, the same as before. Across the whole session:
`LoadBehavior` still 3, `Resiliency\DisabledItems` still empty, no `Application Error` events for
WINWORD.

## Reproducing

    pwsh -File install\install.ps1 -NoBanner        # build, install, register
    pwsh -File tools\check-strip.ps1                # starts Word, drives it, asserts, closes it
    pwsh -File tools\check-strip.ps1 -KeepOpen -Screenshot -SecondDocument

By hand, which is the test that matters: open Word and look between the ribbon and the page. There
is a band with the document's name in it. Resize, maximize, open File and come back - it stays put.

## Not covered

- **One strip per frame.** The shared strip that shows a tab per stacked window, and does something
  when a tab is clicked, is the stacking slice.
- **Nothing is clickable.** The strip returns `HTTRANSPARENT`.
- **Start-screen windows** (Word launched with no document) have no `_WwF` and so get no strip.
  Correct for now, but the product still needs a policy for them.
- **Office Tab coexistence** and **a DPI change** are both reasoned about in code and untested.
- Still untested on the work rig.
