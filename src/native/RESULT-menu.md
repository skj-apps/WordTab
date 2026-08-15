# The tab context menu: Save, Close, Close Others, Close All, New Document

**PASSED 2026-08-15.** `pwsh -File tools\check-menu.ps1` — **32 checks, 0 failures**, plus the four
existing suites re-run unchanged: `check-stack` 33, `check-strip` 40, `check-tabs` 22, `soak-stack`
22. **149 checks, 0 failures across the five.**

Right-click a tab and you get Save, Close, Close Others, Close All and New Document. Right-click the
empty part of the strip and you get New Document on its own, because a right-click that does nothing
reads as a broken area rather than a deliberate one.

Two of those commands are new kinds of thing for this add-in. **Save is the first time WordTab writes
the user's data.** **Close Others and Close All are the first time one click closes several
documents**, which means several save prompts, which means the user gets to say no half way through.
Almost all of the work in this slice is in those two facts.

---

## What was built

**The menu is a plain Win32 popup, built and destroyed on every right-click.** Everything on it
depends on what is true at that instant - how many tabs there are, whether the pointer was over one -
so there is no menu to keep in sync with the tab row if there is no menu between right-clicks.
`TPM_RETURNCMD` returns the chosen id as the function's return value, so there is no `WM_COMMAND` to
route, no id space to keep clear of Word's own thousands of command ids, and no window that has to
still exist by the time a command arrives.

**Nothing is done from inside the menu's modal loop.** `TrackPopupMenu` pumps messages: Word's
timers fire, our janitor runs, documents can open and close. By the time it returns, the strip may
have been detached and destroyed and the state array compacted, so the strip's state pointer is
re-derived from the window afterwards and the chosen command is *posted*, then handled in a fresh
message where everything is looked up again.

**Right-click follows the same press-and-release rule as the close button.** A right press claims a
tab; the menu appears only if the release lands on the same one. Sliding off cancels, and the check
asserts it.

**The tab whose menu is open stays lit.** It is the only thing on screen saying which document the
commands are about, and a tab can be narrow enough that its name is an ellipsis.

---

## The decisions worth keeping

### Save goes through the active window, and checks that Word agrees

`Document.Save`, late-bound, so everything about saving is Word's: an unchanged document is not
written, one that has never been saved gets the Save As dialog, and read-only cases, macros and
AutoRecover behave exactly as they do from Ctrl+S. This is the second and last use of Word's object
model in the whole add-in.

The tab is activated first, for the same reason closing one is: **a Save As dialog owned by a window
underneath another at the same rectangle cannot be seen.** Having activated it,
`Application.ActiveWindow` *is* that tab - and then `Window.Hwnd` is read and compared against the
frame before anything is saved. If they disagree, nothing is saved and the log says so.

That check costs one property read and it is the difference between "should be the right document"
and "is". This is the only path in WordTab that touches the user's data, and a silent save into the
wrong document is not a failure mode worth leaving open. Measured on this rig: **Word reports
`Window.Hwnd` and it matches the `OpusApp` frame handle every time** — logged on every save as
"Word confirmed the window handle".

The activation happens in one message and the save in the next, because Word updates which window is
active while it processes the activation. Asking in the same message asks before Word knows, and the
handle check would then refuse a save that was perfectly legitimate.

### Close Others and Close All are a queue, not a loop

Posting `WM_CLOSE` to six windows at once means the second save prompt arrives on top of the first,
for a document the user cannot see behind it - and "Cancel" on the first one closes the other five
anyway, which is the opposite of what cancel means.

So a batch is a queue with exactly one close in flight, stepped by the janitor. The queue is
computed once, in tab order, **with the active tab last**: the user keeps looking at the document
they were on for as long as the batch allows, and the last prompt they answer is about the document
they were actually reading rather than one they have never seen. Close Others activates the kept tab
first, so that is where they are left standing between one close and the next.

### Detecting "the user said no" — wrong twice before it was right

This is the interesting part of the slice, and both wrong answers passed every assertion that
watched from outside Word.

**Attempt 1: infer it from time.** If the tab is still there and still enabled two janitor ticks
after `WM_CLOSE`, the user must have cancelled. The add-in's own log caught it:

```
12:39:22.361  closing this tab (a background one)
12:39:23.392  close batch stopped - the document is still open, so the user declined
```

One second, and the prompt was not even on screen yet. Word is slower than that, so "still open" on
its own is evidence of nothing. The batch stopped for the wrong reason, and the check passed because
the *observable* outcome - nothing else closed - happened to be the one wanted.

**Attempt 2: require having seen the question, by polling for it.** A modal dialog disables the
window that owns it, so watch `IsWindowEnabled` on the janitor's half-second tick. This failed
against a script that answered the prompt in under a second: the prompt went up and came down
entirely between two ticks, so as far as the add-in was concerned nothing was ever asked, and it gave
up after twelve seconds with "WM_CLOSE produced neither a closed document nor a question". A script
does that routinely; an impatient user hammering Enter will do it eventually.

**Attempt 3, and the right one: hear the question as an event.** `EnableWindow` sends `WM_ENABLE`,
and the frame subclass is already there. A message cannot be missed by being quick.

```
12:57:32.999  Word is asking the user about this document (the window was disabled) - the batch waits
12:57:36.045  close batch stopped - the question was answered and the document is still open,
              so the user declined
```

The poll is kept as well, for a dialog that is modal without disabling anything, but nothing depends
on it alone. **The general lesson, which is the same one the strip and the reconciler already
learned: a state that can be transient must be heard, not sampled.**

The grace after the question disappears is for the other direction. "Save" dismisses the dialog and
*then* writes the file and closes the document; concluding "declined" in that gap would stop a batch
the user had just agreed to. Six ticks, and waiting costs nothing because a document that does close
is caught on an earlier tick by the state that looks for it.

### The states a close in flight can be in

| what is true | what the batch does |
| --- | --- |
| the tab is out of the row | it closed; start the next one |
| Word is asking about it | wait, with **no timeout** - someone typing a filename into Save As is not a hang |
| it was asked about, the question has gone, the document is still here | the user declined; abandon the rest |
| nothing has happened at all | wait 24 ticks, then give up and say exactly that |

---

## What the checking found

**Two of the three findings were in the add-in and one was in the measurement**, which is a change
from the last slice where all three were in the measurement.

**The menu covers the strip it was raised from.** The first hover comparison reported the *control*
tab changing by 714 pixels when the menu closed. Nothing was wrong: a menu opens at the pointer and
grows down and right, so it overlaps the bottom of the strip below the click and draws a shadow past
that again, and the photograph caught it. The comparison now excludes the menu's rectangle inflated
by 40px, and asserts that the control tab is clear of it rather than assuming so. With that:
**259 pixels changed on the tab whose menu was open, 0 on another tab, with the pointer parked in the
corner of the screen** - so the highlight is the menu's and not an ordinary hover.

**`(Get-Frames).Count` is a hard error when there are no frames.** A PowerShell function that returns
an empty array returns *nothing*, and `$null.Count` under `Set-StrictMode` throws rather than
answering 0. It fired at exactly the two moments this suite cares about most - the last document
closing, and the instant during a batch close when Word momentarily has no window that qualifies.
Both this suite and `check-tabs.ps1` now go through `Get-FrameCount`, which re-wraps at the call
site. The same latent bug was in `check-tabs.ps1` and had never been reached.

**`CloseMainWindow` closes one window, not one Word** - again, this time in the *startup* cleanup
rather than the teardown. A process holding a four-document stack went to three and the wait timed
out looking like Word refusing to close. The preamble now closes frames one at a time. It also
refuses to answer a save prompt left over from an earlier run: it cancels the question so Word is
left exactly as it was, and says what to do. A test script may not throw away work it did not create.

---

## Measured

- Menu at 144 DPI: 255x203px, seven entries, ids 1-5, two separators. Read out of the live `HMENU`
  through `MN_GETHMENU`, not photographed - a picture cannot tell you that Close Others is greyed.
- With one tab open, Close Others is greyed and Close is not.
- Save on a saved document: the `.rtf` on disk gains a new timestamp, and no dialog appears.
- Save on a document that has never been saved: Word raises its Save As `NUIDialog` (690x556, over
  the stack rect). Escape cancels it, Word answers `DISP_E_EXCEPTION` "Command failed" - which is an
  answer, not a fault - the document stays open, and the strip comes back in place.
- Cancelling one prompt in a Close Others of three: **all four documents still open**, and the log
  says it stopped because the user declined.
- Close Others on four tabs, all clean: three closed in 1.2 seconds, one at a time, "close batch
  finished (0 tab(s) left unclosed)". The tab left is the one it was invoked on, by title.
- Close All: every document closed and Word exited.

## Switches

`HKCU\Software\WordTab\TabMenu` (DWORD, default 1). Set to 0 and a right-click on the strip does
nothing at all, which is what it did before this slice. Bisectable like every other piece.

## What this slice does not do

**The menu is a system menu, so it is light even when Word is dark** - visible in the screenshot
taken for this writeup. Making it follow Office's theme means owner-drawing it, or the undocumented
uxtheme ordinals that would change menu rendering for the whole of Word, and neither belongs here:
the look is deliberately the last slice, done once against the finished set of elements.

No drag to reorder. No Save All, Rename, Open Containing Folder or Copy Path - the five commands are
the ones asked for, and everything beyond them is a product question rather than a mechanism one.
Close Others closes to the *right* and to the left equally; "Close Tabs to the Right" would need the
explicit tab-order model that the reorder slice has to build anyway.
