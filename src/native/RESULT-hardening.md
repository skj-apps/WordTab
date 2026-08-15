# Slice: hardening - the combinations nothing had tried

**Date:** 2026-08-15
**Outcome: PASSED.** 100 assertions across three suites, 0 failures, repeated runs.

    tools\check-strip.ps1   45 checks   one window: resize, maximize, Backstage, a real resize drag
    tools\check-stack.ps1   33 checks   three windows: stacking, tabs, taskbar, Alt+Tab, minimise
    tools\soak-stack.ps1    22 checks   six windows, and the combinations of the above

The strip, the stack and the taskbar each passed on their own. This slice put them together, which
is where the remaining bugs were.

## What the soak covers

`tools\soak-stack.ps1` opens six documents and then tries to break the stack: maximise and restore,
Backstage on a stacked window, a window launched with no document arriving on an established stack,
twelve rapid tab switches, closing documents out of the middle, and opening one more onto what is
left. After every step it asserts the same invariant - every window at one rectangle, every strip
flush between the chrome above it and the document below it.

## Three bugs it found

**Maximise propagated as a rectangle, not a state.** Only the window the user maximised was really
maximized; the other five had been resized to match it. And they were not even the same size: a
maximized window and a window resized to a maximized window's numbers differ by 72px on this rig,
because one of them covers the strip of screen the taskbar sits on. This was the gap spike 2
recorded and could not close from outside. Maximize and restore are now propagated as *states*,
through `SetWindowPlacement` rather than `ShowWindow` because the latter activates the window it
changes - and activating a window behind the one the user is looking at pulls it in front and takes
the keyboard with it.

**Once the stack came apart, nothing put it back.** Every divergence found so far had its own cause
and its own missed event, and fixing them one at a time is a losing game - the next cause is one
Word update away. The stack now **reconciles**: twice a second it compares every member against the
active window and repairs what does not match. Cheap when nothing is wrong (two rectangle
comparisons per window, no calls), and it gives up after ten consecutive attempts on the same window
rather than fighting something that will not stay put. The same principle as the strip comparing
itself against where it actually is rather than where it last put itself.

**The strip could sit a pixel off the document frame.** Word occasionally reports a child a pixel
from where it was put - the frame's client origin shifts by one when Windows redraws its border -
and a strip positioned from stored numbers then has a visible seam under it. The strip is now
positioned from **where the document frame actually is**, so the two cannot disagree.

That fix needed a bound, and finding out why is the more useful half of it. Derived from the live
rect with no limit, the strip followed the document frame into states it should have ignored: when
Word resets `_WwF` before we shift it back, the strip was moved up over the ribbon, and during a
teardown, 48px off the top of the window. The live rect is now trusted for pixel-level disagreement
only - less than one strip height - and our own intended rectangle is used for anything larger.
Measured after the bound: 32 corrections across a full three-suite run, **every one of them exactly
±1 pixel**, and none of the 48px kind.

## Two things the checks themselves were doing wrong

**A geometry test has to drive the window that has focus.** `check-strip.ps1` drove whichever window
it found first. Once several documents are open they are stacked, and a background window is one
Word has stopped laying out - so the test was resizing a window nothing maintains and asserting
about the result. It now brings its target to the front first and re-focuses before each step, which
is the product's own rule turned into a test rule.

**Restoring a background window is not something a user can do.** The soak captured "the active
window" before maximising and restored that handle afterwards, by which time activation had moved.
The add-in correctly put the window straight back, and the assertion failed - testing nothing but
the test's own bookkeeping. It now re-reads which window is in front before acting on it.

## Not covered

**Word's Start screen.** On this rig Word opens a blank document instead, so launching `WINWORD.EXE`
with no arguments exercises "a new document arrives on an established stack" - worth having, but not
the Start-screen case. That still needs the Start screen option turned back on, and the product still
needs a policy for such windows.

Everything listed as uncovered in `RESULT-stack.md` and `RESULT-taskbar.md` remains so: no tab
affordances, Office Tab coexistence untested, DPI change untested, work rig untested.
