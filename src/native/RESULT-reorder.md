# Dragging a tab to reorder it

**PASSED 2026-08-15.** `pwsh -File tools\check-reorder.ps1` — **30 checks, 0 failures**, plus the
five existing suites re-run unchanged: `check-stack` 34, `check-strip` 40, `check-tabs` 22,
`check-menu` 32, `soak-stack` 22. **180 checks, 0 failures across the six.**

Before this slice the tab row was the order documents happened to join the stack, and nothing could
change it. Now a tab can be picked up and carried, the row rearranges under it as it goes, and where
it is let go of is where it stays.

---

## What was built

**The order model: the array *is* the order.** `g_members` in `stack.cpp` was already walked in array
order by `StackTabs`, so the position of an entry in that array is now, explicitly, the tab's
position in the row — and reordering moves the entry. There is no `order` field.

That is a decision rather than an omission. An `order` int has to be kept consistent with the array
by every path that adds a member, drops one, or compacts the array after a close, and any single one
of them getting it wrong produces two tabs claiming the same position. An array cannot disagree with
itself about what order it is in.

It also made three other things right for free:

- `FirstJoined` is now the leftmost tab rather than the earliest-joined one, so closing the active
  tab falls back to a neighbour the user can see rather than to a document that happens to be old.
- A batch close runs left to right, which is the order the user reads the row in.
- Joining still appends, so a new document arrives at the end — which is where every other tabbed
  application puts one, and which the check asserts still holds *after* the row has been rearranged.

Two calls carry it: `StackTabIndex(frame)` answers where a tab sits (or `-1` for a window that is not
a tab in a stack), and `StackMoveTab(frame, toIndex)` moves it and returns `TRUE` only if the row
really changed. The index space is the one `StackTabs` hands out, so the strip never has to translate
between "what I drew" and "what the stack thinks".

The translation that *is* needed is inside `StackMoveTab`: `g_members` also holds windows that are
not joined — one Word has hidden, one on its way in or out — and those have no place in the row. So a
tab index is resolved to the array slot the tab at that position actually occupies, and the entry is
moved there.

**The row rearranges live, not on the drop.** As the tab is carried, `StackMoveTab` runs on every
mouse movement that changes which slot it is over. By the time the button is released the order is
already what the user can see, so the drop itself has nothing to commit — it is bookkeeping and a log
line.

Live won over an insertion marker for a specific reason: every window in the stack draws the same
row, so a live reorder is something all of them already know how to show. An insertion marker would
have been new state the drawing code had to be taught, shown by only one strip, for a result that is
harder to read.

The consequence is that cancelling has to *undo*, which is why the position the tab was picked up
from is kept for the whole gesture.

**The carried tab is drawn by whichever strip is in front, not by the one holding the mouse.** This
is the part that is easy to get wrong. A press on a tab activates that document *before* the drag
begins, and activating raises a different window — whose strip is now the one on screen. So the strip
that owns the mouse capture is very often not the strip the user can see.

Per-strip drag state would draw the tab travelling on a window that is underneath another one, and
nothing at all on the window the user is looking at. So the drag is held as the add-in's state rather
than one strip's, and every strip draws the same carried tab. The one on top is by construction the
one that shows it. Sound because all the strips in a stack are the same width at the same position —
that is what the stack guarantees — so one x coordinate is meaningful in all of them.

Confirmed by photograph rather than by argument: `drag-1-lifted.png` in the session scratchpad shows
a *background* tab picked up and drawn out of its place, on a strip belonging to a different window
than the one with the capture.

**One tab-drawing function.** `DrawOneTab` takes the rectangle to draw at, so the carried tab goes
through exactly the same code as a tab sitting still. A dragged tab cannot end up looking like a
different kind of object. It is held out of the main loop and drawn last, so it is on top of the tabs
it is passing over rather than half under them.

**A press is a click until it travels.** Four logical pixels, which is what Windows itself uses as
its drag threshold. Horizontal distance only: the row has no vertical meaning — there is no tear-off
here, so dragging a tab downwards is the same gesture done untidily, not a different one — and a
threshold that counted vertical movement would start a reorder from a hand that slipped while
clicking.

**The swap is measured from the tab, not from the pointer.** The slot a carried tab belongs in is the
one containing the *tab's* centre. Swapping on the pointer's position makes a tab grabbed by its
right-hand edge jump a place the instant it is picked up.

New switch: `TabDrag` (default 1). Off, a press on a tab switches to it and nothing else, and the
mouse capture is never taken — so the mechanism is out of the way rather than merely inert.

---

## The decisions worth keeping

**There is no Escape, and that is deliberate.** The strip is `WS_EX_NOACTIVATE` and never holds the
keyboard focus, so a keypress never reaches it. The tempting fix is to read `GetAsyncKeyState` on
each mouse movement — and that is precisely the mistake the batch close was written twice to avoid:
**a state that can come and go inside a polling interval has to arrive as an event or it is not being
observed at all.** A tap of Escape between two mouse movements would simply be missed.

So the cancels are the ones that *are* events: the right button, and losing the mouse capture. Both
go through `DragUndo`, which puts the tab back where it was picked up from. Dragging it back by hand
is the third. A thread-local `WH_KEYBOARD` hook would make Escape into an event and remains available
if it is ever wanted, but a keyboard hook installed in Word's UI thread is a real risk surface to buy
a convenience with, and it was not bought here.

**The right-button cancel keeps the capture.** The left button is still down, and releasing the mouse
mid-gesture would deliver that release to whatever happens to be underneath the pointer. So
`g_dragStrip` stays set as "this strip still holds the capture for a gesture that is over", with
`g_dragFrame` cleared — and the eventual `WM_LBUTTONUP` does the cleanup.

**`WM_RBUTTONUP` during a drag has to be swallowed, not ignored.** `DefWindowProc` turns a right
release into `WM_CONTEXTMENU`, and a child window's `WM_CONTEXTMENU` goes to its *parent*. Letting it
through would have ended a cancelled drag by opening Word's own context menu. Found by reasoning
about the fall-through while writing the cancel, and asserted by the check: after a right-button
cancel there is no popup menu on screen, ours or Word's.

**A drop clears the drag state before releasing the capture.** `ReleaseCapture` sends
`WM_CAPTURECHANGED` to the same window, and that handler's job is to undo an interrupted drag — the
exact opposite of what a completed drop means. Ordering is the whole of the fix.

**A carried tab whose document goes away is not an error.** Word can hide the window, a close already
in flight can complete, the janitor can drop it from the stack. `DragMove` checks that the tab is
still in the row on every movement and simply ends the gesture if it is not: there is nothing to
carry and nowhere to put it back.

**Nothing outside `StackMoveTab` may hold a `Member*` across it** — the entries move. Nothing does:
every caller reaches the stack through an `HWND`, which was already the house rule because frame
lifetime is not document lifetime.

---

## What the check script had to learn

**Reading the row is N clicks, and Word has to be in front for them to count.** There is no API for
"what does the strip say", so a tab's identity is established the only honest way: click it and see
which document comes forward. That doubles as the check that every window agrees on the row, because
the click for tab *i+1* is aimed at the strip belonging to whichever window tab *i* brought to the
front.

It also produced the run that cost the most time. Switching tabs ends in `SetForegroundWindow`, which
**Windows refuses from a process that is not already the foreground application** — so when something
else owned the desktop, a click on a tab raised that document *inside* Word without making Word
foreground. `GetForegroundWindow` then answered with the other application, and the first reading of
the row came back naming Remote Desktop Connection three times. The add-in was behaving correctly
throughout; the log showed every activation. `Set-WordForeground` now takes the foreground through
the `AttachThreadInput` handshake before a reading and retries any click that lands outside Word.

**The same run found a latent trap in `check-stack.ps1`.** Under `Set-StrictMode -Version Latest`,
indexing an empty array is a hard error rather than `$null` — so `@($parts | Where-Object { ... })[0]`
*crashes the script* when the foreground is not a Word frame, instead of falling through to its own
fallback. This is the same shape as the `(Get-Frames).Count` trap already on record. Replaced with a
`Get-TopParts` search-with-fallback, and the suite now asserts up front that Word is foreground, so a
stolen desktop reads as a failed assertion rather than a stack trace. `check-stack` is 34 checks
because of that new assertion.

**The drawing is isolated, not merely observed.** Asserting "the strip changed during a drag" would
pass on a hover highlight. So the photograph at rest is taken with the *same tab selected and the
same tab hovered*, and the first stage of the gesture moves the tab a third of a slot — past the drag
threshold, so it is being carried, but nowhere near half way past its neighbour, so the row must not
reorder. The only thing that can differ between the two photographs is the tab having been lifted out
of its place. Measured: 957 pixels changed where the tab was, 0 anywhere else in the row.

**The live rearrangement is asserted mid-gesture, with the button still held**, from the add-in's own
log. A design that waited for the drop would log nothing at that moment. Every hold is paired with a
`DragRelease` in a `finally` block — a left button left down is a wrecked desktop, exactly like a held
Alt key.

**"It moved" is not the whole claim.** A reorder that moved the right tab and shuffled the others
would pass "it is now the third tab". The check builds the expected order explicitly and compares the
whole row.

---

## What is not here

- **No tear-off.** Dragging a tab out of the strip does nothing special; the insertion point is
  computed from the horizontal position only. Making a tab into its own window is a separate feature
  and a much larger one.
- **No keyboard reorder**, for the same reason there is no Escape — see above.
- **No scrolling tab row.** With enough documents open the tabs shrink to a 70-logical-pixel minimum
  and then overflow the strip, and a tab carried into the overflow is clamped to the last position
  that is drawable. That gap predates this slice.
- **The carried tab has no lift** — no shadow, no raised border. It is drawn on top and offset, which
  reads correctly, but the emphasis belongs with the rest of the look. That is the next slice, and
  this one deliberately left the placeholder styling alone rather than painting it twice.

---

## Running it

```
pwsh -File tools\check-reorder.ps1
pwsh -File tools\check-reorder.ps1 -KeepOpen -Screenshot
```

Four documents, a reading of the row between every gesture, and the mid-drag photographs saved with
`-Screenshot`. Like every other suite it closes Word gracefully and never kills it.
