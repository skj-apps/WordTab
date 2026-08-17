# A tab row that scrolls

**PASSED.** 46 checks in the new `tools\check-scroll.ps1`; the other eight suites re-run —
**311 checks, 0 failures across the nine, from a clean start.**

Open more documents than fit and the row now scrolls. The tabs live in a **track** that stops short
of a fixed cluster pinned to the right-hand end of the strip — two scroll chevrons and the **+** —
and they are clipped to it and hit-tested against the part of them inside it. Wheel over the row to
move it, click a chevron to step one tab, **hold a chevron to keep scrolling**, and carrying a tab
against either end scrolls the row under it. The active tab brings itself into view when it changes
and when the window changes shape.

---

## Holding a chevron — added 2026-08-16

One click is still one tab, which is what the rest of this writeup argues for. The cost of that is a
long row taking a lot of clicks, and holding the button is the answer, the same way it is on a
scrollbar arrow. **`check-scroll` 49 → 62 checks.**

- **Two intervals, because one would be wrong in both directions.** Nothing for `CHEVRON_HOLD_MS`
  (400ms, comfortably longer than a click), then a step every `CHEVRON_REPEAT_MS` (90ms). Long enough
  never to fire on an ordinary click; short enough to cross a long row.
- **The click still happens on RELEASE, unchanged.** Moving it to the press would have been the
  smaller change, and it is what a scrollbar arrow does — but it would have taken the
  press-and-slide-off cancel away from these two buttons and left the close button as the only one in
  the strip that still had it.
- **A hold that already scrolled suppresses the release's step.** Otherwise every held chevron
  overshoots by exactly one tab, which reads as the row being imprecise rather than as an extra step.
- **Each tick re-asks two questions rather than trusting what it was armed with**: is the pointer
  still on that button — so sliding off pauses the repeat and sliding back on resumes it, the same
  rule the release already follows — and is there anywhere left to scroll, which stops the timer
  spinning at 90ms against a row already at its end.
- **`SetTimer` with an id that already exists replaces that timer's interval**, so the hold delay
  becomes the repeat rate rather than a second timer.
- **Proved the way everything else here is proved.** The scroll position cannot be read from outside
  Word but is exactly knowable at both ends, so a hold that starts at the start and ends with the row
  against the far end has scrolled by more than one tab whatever the repeat rate turns out to be.
  Asserted against **where clicking to the end reached**, not against a named document — the row order
  is whatever the sections above left it, so naming one would assert the tab order rather than the
  scroll position. The first version did name one and was wrong: an earlier section adds a `Document1`
  that ends up last. Both directions driven, and **a quick click is asserted NOT to log a hold**,
  which is the assertion that the delay exists at all.

---

## The defect

Before this slice, a row with more documents than fitted at the 70-logical-pixel minimum did this:

```c
int after = (count > 0) ? (out->tab[count - 1].right + gap) : (client->left + pad);
int limit = client->right - pad - plusW;
if (after > limit)
    after = limit;                    // <- and now the + is on top of the last tab
```

The tabs stopped shrinking at the minimum and ran off the end of the strip. The **+** had nowhere
left to go, so it was clamped back to the right edge — *on top of the last tab*. Hit-testing tries
tabs first and drawing paints the **+** last, so **the user saw a + and clicked a tab**. At 200% on a
1280-pixel window, nine documents was enough for the **+** glyph's centre to fall inside that tab's
**close button**: clicking the drawn **+** closed a document.

**It was not an edge case.** Word's own default window on this rig is 874 strip pixels, which at 200%
is room for five tabs. **Six documents in a freshly opened Word already overflowed.** That number was
measured while writing the check suite, whose first section had to widen the window before it could
find a row that fits at all.

---

## The shape of the fix

There are three ways the row can be laid out, and `ComputeLayout` decides which is in force. Nothing
else in the file knows.

| | when | tabs | the + |
|---|---|---|---|
| **FIT** | they all fit at the minimum | at least `TAB_LOGICAL_MIN_W` wide | after the last one |
| **SCROLL** | they do not | at the minimum, scrolled in a track | pinned right, with two chevrons |
| **SQUEEZE** | `TabScroll=0` | no minimum at all — every tab on screen | after the last one |

The condition separating FIT from the other two is `width * count > available` **after** the minimum
has been applied — which is exactly, and only, the case in which the old code clamped. So:

> **Everything this slice changes is inside the condition that was broken.** FIT is the arithmetic
> that was there before, in the same order, to the pixel. That is what let a rewrite of the layout be
> checked against 265 assertions written before it — every one of them runs in FIT.

### The cluster and the track

```
[ tabs ......................... ] gap [ ‹ ][ › ] gap [ + ]
 track.left                track.right
```

`track.right` is `prev.left - gap`. A tab and a button cannot occupy the same pixel **by
construction** rather than by a clamp that gives up when it runs out of room — which is the general
form of the bug, and the reason the same mistake is fixed in a second place: a tab's **close button
is dropped entirely unless the whole of it is inside the track**, so there is never an × drawn half
under a chevron for the user to aim at.

The two chevrons touch each other on purpose. They are one control with two directions, and a gap
between them reads as two unrelated buttons that happen to be adjacent.

### Clipping

`Surface` grew a clip rectangle, honoured in `Blend`, and `SurfClip` sets the GDI clip region on the
surface's DC in the same call so the tab names are clipped with the pixels they sit on. The tab loop
runs inside `SurfClip(s, &layout.track)` and everything after it runs outside.

**A tab is drawn at its full width and the pixels past the end are discarded.** Clipping the
*rectangle* instead would have rounded the corners of the cut, and a tab that appears to end neatly
where it has in fact been truncated is a tab that hides the fact there is more of it. The flat
renderer does the same thing with `SaveDC` / `IntersectClipRect` / `RestoreDC` — applied only if
`SaveDC` succeeded, because `WM_PRINTCLIENT` arrives with somebody else's device context and a clip
left on it would truncate whatever they drew next.

---

## The decisions worth not re-deriving

### The scroll position is global, like the drag, and for the same reason

Every window in the stack draws the same row. A row that stood at a different scroll position in
each of them would stop being one row the moment there were enough documents for it to matter. One
number, `g_scroll`, every strip.

**`ComputeLayout` is its only writer, and it writes only when the strip it is laying out has tabs.**
That second half is not a detail. A window with no document keeps its strip and paints an empty row
([[wwf-outlives-the-document]] and `RESULT-startscreen.md`), and `StripRefreshTabs` invalidates
*every* strip — so letting the empty one clamp the shared number would reset the real row to the
start every time anything repainted. Caught by reasoning before it was written, not by a photograph
afterwards, which is the only bug in this slice that was.

Writing it from the layout rather than from whoever moved it is what makes it self-healing: widen the
window and the next paint discovers there is less to scroll and shortens it, with nothing having to
notice the resize.

### Revealing the active tab is an event, and there are two of them

A tab row whose selected tab is off screen has stopped answering the one question it exists to
answer. But revealing on *every* layout would drag the row back a frame after the wheel moved it, and
the wheel would appear not to work at all. So the reveal fires on the two things that can put the
active tab off screen without the user asking:

- **the active document changed** — a new document appends a tab at the far end and switches to it,
  Ctrl+F6 walks the windows, closing a tab moves the selection;
- **the shape of the row changed** — `maxScroll` is a function of the tab count and the width of the
  track, so it moves when a document opens or closes and on every step of a resize drag.

The second one is here because the first one was not enough, and the photograph is in this slice's
history: six documents opened on a wide window, then the window narrowed to three tabs, and the user
was left looking at tabs one to three with tab six selected and nothing on screen saying so.

Neither is a poll. Both are quantities `ComputeLayout` has already computed, compared with what they
were last time — and **a wheel scroll changes neither, which is exactly the property that makes the
wheel work.** This is [[transient-state-must-be-an-event]] applied to something that is not
transient: the rule is not "never compare against last time", it is "a state that comes and goes
inside a tick cannot be sampled at any cadence". A scroll position does not come and go.

### Clicking a tab claims it as already revealed

`WM_LBUTTONDOWN` sets `g_scrollShown = hit.frame` **before** `StackActivate`. A tab at the end of a
scrolled row can be half off the track, and the reveal that follows an activation would slide the row
under a pointer that is still holding the button down: the tab would move out from under the hand
between the press and the drag, and the carried tab would sit a scroll's width from the pointer for
the rest of the gesture. The user can see the tab — they just clicked it.

### The wheel arrives, and it is Windows' decision that it does

`WM_MOUSEWHEEL` goes to the window with keyboard focus, and the strip is `WS_EX_NOACTIVATE` and never
has focus. What delivers it is **"scroll inactive windows when I hover over them"**, which is on by
default and routes the wheel to the window under the pointer. Measured, not assumed: the first smoke
run photographed the row before and after three notches and it had moved.

That it is a *setting* is why the chevrons are not optional. `WM_MOUSEHWHEEL` is handled too, and in
the opposite sign — a wheel forward is up and up is left, but a horizontal wheel's positive direction
is already right.

### One tab per chevron click

A page per click gets to the far end of a long row sooner and overshoots the tab you were looking for
every time. The wheel is already the way to cross a row quickly.

**A chevron with nowhere to go is drawn dimmed and is not a hit target at all.** A dimmed button that
lights up under the pointer and then does nothing when pressed is worse than one that ignores the
pointer. Both states are photographed, in both themes.

### Carrying a tab past the end

Without it, a drag in an overflowing row could only rearrange the tabs that happened to be on screen.
The timer is armed when the drag passes the slop and killed on all three ways a gesture can end —
`DragForget`, `DragUndo`, and the "that tab is no longer in the row" path — which is one start and
three stops rather than four of each.

**It is a timer rather than mouse movement because the gesture that needs it is the one where the
pointer has stopped.** What it samples is the pointer's *position*, a continuous quantity, not the
kind of come-and-go state that has to be heard rather than polled. On a tick it scrolls and then
re-runs `DragMove` with the current cursor position, rather than keeping a second copy of the
"which slot is this over" arithmetic that could disagree with the first.

The drop slot is computed in the row's own coordinates now — `(centre - track.left + scroll) / width`
— rather than by scanning the rectangles on screen, because in a scrolled row the slots either side
of the visible ones are real places a tab can be put and there is no rectangle to find them by. At
`scroll == 0` it is the same answer the scan gave.

### `TabScroll=0` is not "give back what it did before"

Every other slice added a switch that restores the previous behaviour. What the previous behaviour
*was* here is the defect, and a switch whose only effect is to reproduce it would exist to reproduce
it — the same reasoning that gave the start-screen slice no switch at all.

So the switch gives the **other** honest answer to too many documents: **no minimum width**, so the
tabs divide the track between them however many there are and **every document is on screen at
once**. It is a worse row to read and a better one to be *sure* of, and it is the thing to reach for
if the scrolling row ever strands somebody's document off the end of a strip on a machine this has
not run on. Five documents in a window sized for three tabs gives 101-pixel tabs where the minimum is
140, all five inside the track, all five reachable with one click — measured.

---

## How the suite proves it without reading pixels

The row's scroll position is state inside Word's process and the add-in does not publish it. It is
nevertheless *knowable*: a row that has just overflowed is at zero, a row scrolled harder than it can
move is at its maximum, and everywhere else it is exactly zero because a row that fits cannot scroll.

So every assertion is the same shape: **drive the row to a known end, compute where a given slot must
therefore be, click there, and read the title of whatever came forward.** If `tools\WordLayout.cs`
and `ComputeLayout` disagree by so much as a tab width, the wrong document comes forward. That is
decisive in a way a photograph is not — and it is how "one click of a chevron moves the row by
exactly one tab" is measured rather than asserted as "it moved".

The mirror gained the same three modes and a `scroll` parameter. `Tabs(strip, count)` still means
`Tabs(strip, count, 0, true)`, so the eight existing suites needed no edit at all.

**The window is narrowed rather than the documents multiplied.** Overflow needs
`count * minimum > available`, and the width is the cheaper half of that product to change: six
documents in a window sized for three tabs overflows exactly as hard as twenty in a maximised one and
takes a minute rather than five to set up. The target width is computed from the measured DPI.

The end-to-end proof of the defect is one click: with the row scrolled as far right as it goes, click
the centre of the drawn **+** and assert that a document *appeared*, that the log says
`new-document button clicked`, and that it says nothing about a close. Before this slice that click
closed a document.

---

## Two things found in the harness, neither of them in the strip

Both were found by this slice's regression run and both had been there for slices.

### `soak-stack.ps1` was killing Word

Its cleanup was `CloseMainWindow()` followed by `Kill()`. Both halves were wrong. Word's "main
window" is one of seven frames there, so closing it left the other six standing — and the `Kill` that
followed is the thing this project's notes have forbidden since the first slice, because **a killed
Word offers to recover those documents the next time it starts.**

That is not theoretical. It cost a run of `check-look` in this slice: soak finished green, killed
Word, and the next suite opened three documents into a Word that had helpfully brought back four of
soak's — so a suite asserting "3 Word windows" measured seven. The trap is written down twice in the
project memory and had a live source in this file the whole time. It is now WM_CLOSE per frame, no
`Kill`, and unconditional: the guard that only cleaned up if the suite had started Word meant the run
that most needed cleaning up was the one that got none.

Verified by running `soak-stack` and `check-look` back to back with nothing in between: 22 and 41,
and `3 Word windows (3)`.

### `check-look.ps1`'s well sampler was reading the active card

It samples the well in the empty run to the **right of the +**, and falls back to a fixed box at the
left of the strip when the row fills the strip and there is no such run. That box is inside tab zero,
and the fallback silently assumed tab zero was idle.

It is not, in the theme-change section: four sections earlier the suite presses tab zero to
photograph a carried tab. So on a Word that had remembered a narrow window, the "well" it read was
the **active card** — Word's own chrome — and the ribbon-to-well step came out as 0 levels where the
code's step is 26.

**Measured on the build from before this slice as well as after it**, by stashing the changes and
re-running: identical failure. So the suite was fragile and my new suite only exposed it, by changing
the window size Word remembers. `Get-WellColour` now takes the index of a tab the caller *promises*
is inactive, and the theme section selects the last tab before reading so that the promise is true by
construction rather than by luck.

Both of these are the same lesson from a different angle: **a fallback that happens to agree with the
truth is not a measurement.** The look slice recorded it about a colour sampler that returned a bare
FALSE. Here it is a sampler that returned a plausible colour from the wrong pixel.

---

## What to preserve

- **`ComputeLayout` is the only writer of `g_scroll`, and only when the strip has tabs.** An empty
  row belongs to a window that is not in the stack and its layout must not speak for the row every
  other window is showing.
- **A tab's hit area is the part of it inside the track**, never the whole rectangle. The defect this
  slice removes is what happens when a button is drawn over something that is still hit-tested.
- **A close button that is not wholly inside the track does not exist.** Same rule, second place.
- **The reveal is two comparisons against last time, not a rule applied every layout.** If a third
  trigger is ever needed, add a comparison; do not make it unconditional, or the wheel stops working.
- **FIT must stay identical.** It is what 265 assertions from eight earlier suites measure.
- **`SurfClip(s, NULL)` before and after the drawing.** The surface's DC belongs to the strip and
  outlives the paint; a clip left on it truncates the next one.

## Switches

| | default | off |
|---|---|---|
| `HKCU\Software\WordTab\TabScroll` | 1 | the squeeze layout: no minimum tab width, no scrolling, every document on screen |

`TabButtons=0` removes the chevrons along with the × and the **+**; the row still scrolls, and the
wheel is the only way to move it, which is what that switch means.

## The suite

```
pwsh -File tools\check-scroll.ps1            # 46 checks
pwsh -File tools\check-scroll.ps1 -Screenshot -KeepOpen
```

Six documents, a window narrowed until three tabs fit, and the row driven to both ends with the
chevrons, the wheel and a drag. It sets `TabScroll=0` for its last section and puts it back in a
`finally` — a check script that leaves a machine configured differently from how it found it has
changed the thing it was measuring.
