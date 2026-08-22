# One page across

**The complaint.** From the work rig, after three answers to *"it opens 3 pages wide"* had already
been given and shipped:

> fyi i have to chage view to single page on word open if thats what your talking about
>
> every time i mean

And then, when I proposed a registry command and a second file to double-click:

> nope. why would i do that instead of clicking a cmd file?
>
> why make things harder and more complex? thats just not logical
>
> idk why i have to even say this tbh

All three of those are the same correction and all three were right. What follows is the measurement
that should have come first, and the fix in the only place that costs them nothing.

## What Word is actually doing

Measured on 16.0.20228, by breaking a healthy Word and then curing it.

| state | what Word reports |
|---|---|
| healthy | `Zoom.PageColumns = 99`, `Percentage = 100` |
| after anything sets a column count | `PageColumns = 3`, `Percentage = 10` |

`99` is Word's *as many as fit*, not a request for ninety-nine pages: in that state the pages side by
side are purely a function of the window's width, which is what the row's own width already answers.

The moment a **column count** is set, two things happen and both of them are sticky:

1. Word **crushes the zoom** to fit that many pages. Three columns in a 1305px window is 10%, which
   is Word's own floor.
2. Word **keeps it as a default**. It survives closing the document, closing Word, and opening a
   document Word has never seen - including a brand new blank one from the template. A registry diff
   across the change says where it lives: `HKCU\Software\Microsoft\Office\16.0\Word\Data\Settings`,
   an opaque blob. Nothing else in Word's whole key moved, and `Normal.dotm` was not touched.

**And the ribbon's One Page and 100% buttons do not clear it.** Measured: `Zoom100` takes the current
window to `100% / 1 column` and the next Word opens at `10% / 2 columns` again. That is the whole of
*"every time"* - the button they were pressing could never have stuck, so they were re-pressing it
for days.

## Why none of the earlier work touched this

Everything built for "3 pages wide" was about the **width of the window**: the row remembering the
size they settled on, and then a default width for a window Word opens three or more pages across.
That is a real cause of pages side by side and it was worth fixing - and it is a different cause from
this one. With a column count stuck at 2, narrowing the window does not help; it makes the pages
*smaller*, because Word shrinks the zoom further to keep fitting two of them.

Two causes, one symptom. The arithmetic in `RESULT-rowsize.md` was right and answered the other one.

## The fix, and where it had to go

Three deliveries were on the table: a registry command, a second `.cmd` file to double-click, and
this. Their answer to the first two was that they already download and double-click one file, so
anything else is worse - which is correct, and is also the standing lesson in
`RESULT-rowsize.md` about fixes that need a manual step.

So it is in the add-in, which is already there at every Word start:

**`WordTabOnePageView` (connect.cpp), called from `Join` (stack.cpp)** - the moment a document
arrives in the row. Two illnesses, one cure:

- opened **more than one page across** → one page, and 100%
- opened at **under 50%** with one page → 100%

A document that opens on one page at a readable size is not touched at all, whatever zoom it is at.
A zoom somebody chose is theirs; a zoom under half size is not a choice, it is what a fit-many-pages
calculation left behind.

**On a timer this would be indefensible; as an event it is not.** A document arriving in the row is
an event, and this asks Word two questions and writes two properties, once, measured at 3ms. The dot
poll is the standing warning about the other kind - see `RESULT-dot.md`, where one object-model call
measured 430us here and half a second on the work rig - and the difference is per-event against
per-window-twice-a-second.

`OnePage=0` turns it off. It is on by default because it is the setting they were correcting by hand
on every single Word start.

## What was checked

`tools\check-onepage.ps1`, new, 11 checks, in the battery after `rowsize`. It **breaks Word the way
the work rig is broken** - through the object model, which is the same door the ribbon uses - and
then asserts that opening a document puts it right.

The break is asserted first, with the switch off: *"Word carried the column count into a document it
had never seen (3)"*. Without that, a suite that could not break Word would pass while the add-in did
nothing.

### The suite caught the fix being half a fix

The first version corrected the column count only, and passed:

```
PASS  the document is one page across (got 1)
```

while printing, three lines further down and asserted by nothing:

```
second document: 1 page(s) across at 10%
```

One page, and unreadable. Word recomputes the zoom when it is asked for three columns and does
**not** recompute it when it is asked for one, so the crushed value survives into every document
afterwards - and the user would have gone on fixing the zoom by hand having been told it was fixed.
The suite now asserts what the person actually cares about, which is not the column count:

```
Assert ($fine.Percent -ge 50) 'and readable rather than the leftover fit-many-pages zoom'
```

**The check that matters was the one printing a number nothing was looking at.** Same shape as the
`check-rowsize` failure that read "geometry passed, log failed": if a suite prints a fact, something
should be asserting it.

## What is not covered

- **That this is what their machine has.** It has never been reproduced on their rig, because
  nothing in the log records the view - their report cannot say what `PageColumns` was. What is
  reproduced is a Word broken the same way, and the add-in's line now names the numbers it found, so
  the next report from them says outright whether this was it: `view hwnd=0x... this document opened
  3 pages across at 10%`.
- **A user who wants two pages side by side** gets one, until they set `OnePage=0`. That is a real
  cost and it is the reason the switch exists.
