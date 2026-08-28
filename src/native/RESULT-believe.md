# Nothing is believed on one reading

The first field report taken on `6d9c22f` — the work rig's DLL hash matches this repo's build byte
for byte, so for the first time in three reports the log is a verdict on the code that is checked in
rather than on the build before it.

The three fixes in that build hold. The row came back at the size it was left at across a reboot
(`the row takes the size it was left at, not the one Word restored: (1,603 1285x1473)`, 07:51:46),
no tab was highlighted for a document that was not on screen, and nothing in nine and a half hours
of log says a document opened several pages wide.

What the log did name is four things, and they turn out to be one thing four times: **the add-in
stating something it had only one reading for.** A colour it sampled once. A name it took while the
stack was still deciding. A tooltip for a window that had gone. And an instrument averaging two
different measurements together and reporting the mean as if it were one.

---

## 1. The palette believed every sample, including the ones that were wrong

### The measurement

Word's ribbon, sampled off the screen every two seconds, does not hold still on that rig:

```
16:32:03.336  palette sampled ... chrome=RGB(9,9,9)  dark  well=RGB(0,0,0) card=RGB(9,9,9)  edge=RGB(35,35,35)
16:32:07.346  palette sampled ... chrome=RGB(10,10,10) dark well=RGB(0,0,0) card=RGB(10,10,10) edge=RGB(36,36,36)
16:32:35.371  palette sampled ... chrome=RGB(9,9,9)   ...
16:32:53.389  palette sampled ... chrome=RGB(10,10,10) ...
```

It reads `RGB(9,9,9)` while its window is active and `RGB(10,10,10)` while it is not, so **every
switch between two documents moved the palette by one unit in each channel.** `ApplyPalette` compares
for equality, so each of those was a new theme: `DerivePalette`, two brushes destroyed and remade,
and `StripRefreshTabs` invalidating every strip in the row. Over a hundred times in one working day,
for a difference no display can show.

And twice — 14:05:24.848 and 14:09:19.310 — the sampler returned something that was not the ribbon
at all:

```
14:05:24.848  palette sampled ... chrome=RGB(0,0,0) dark  well=RGB(0,0,0) card=RGB(0,0,0) edge=RGB(26,26,26)
14:05:26.908  palette sampled ... chrome=RGB(10,10,10) dark well=RGB(0,0,0) card=RGB(10,10,10) edge=RGB(36,36,36)
```

Two point zero six seconds apart — one sample interval. A single reading, believed, and the whole tab
row drew in a different colour until the next one put it back.

### Why the sampler cannot catch that itself

Nothing about the `RGB(0,0,0)` sample is detectable from inside it. `SampleChrome` already asks *who
owns this pixel* before reading it, already refuses a band that is less than three quarters flat, and
already refuses a window with a message bar between its ribbon and the strip. That sample passed all
three, because it was flat, it was owned by the ribbon, and there was nothing docked. The only thing
that separates it from a real theme change is that **a real theme change is still true two seconds
later.**

### The fix

Two rules, and each catches what the other cannot.

- **A dead band.** `CHROME_SAME_WITHIN 4`: a colour within four units in every channel of the one
  already applied is the same colour. That is the 9/10 flap, and it is nowhere near the themes this
  has to tell apart — black samples at 9, dark grey at 41, and the two light ones in the 240s and at
  255. The smallest real gap is an order of magnitude above the band and the observed jitter an order
  of magnitude below it.
- **A second reading.** A colour outside the band becomes a *candidate*, and is adopted only when the
  next sample agrees with it. That is what the `RGB(0,0,0)` outlier fails, since ten units clears the
  dead band easily.

The first sample of a process is exempt from the second rule. Until then the palette is the theme
registry's guess rather than a measurement, and making the first measurement wait would leave a
possibly-wrong colour up for twice as long as it is now — the opposite of the point.

The cost is that a genuine theme change reaches the strip one sample — two seconds — later. A theme
change is a person in File > Account.

### Why the suite for it starts no Word

`check-look.ps1` drives real Word and photographs the strip, and on this rig **Word's ribbon holds one
colour for the whole run.** It cannot tell the old rule from the new one, which is exactly how the old
rule shipped: it passed every suite on this machine and then rebuilt the palette a hundred times a
day on theirs. This is the same shape as the dot poll's governor and it gets the same answer —
`tools\palette-test.cpp` lifts the decision out and replays it against the colours in the field log,
old rule beside new, and `tools\check-palette.ps1` compiles and runs it in about a second. 23 checks.

It includes nothing from `src\native`, deliberately. If the copy drifts from `strip.cpp` the copy is
wrong, and the copy is what is being asserted.

---

## 2. A tab called "Word" in the middle of the row

### The measurement

16:12:52. Enable Editing was clicked on a Protected View document. Word closes the sandbox frame and
opens a real one for the same file, and for a second and a half the row was:

```
16:12:52.639  stack  row, left to right: ... 0x...61ACA |075721309-00 backdated property, contents increase|
                                            0x...11CAA |Word|
                                            0x...F16B2 |075721309-00 - Wujick (FL) - TIV in Windzone|
16:12:54.032  stack  hwnd=0x...11CAA  left the stack (hidden or minimised)  (6 remain)
```

Seven tabs, one of them called "Word". The same thing at 11:01:43 the day before.

### Two causes, and the second one is the one that mattered

The first is ordering. `ReadTitle` runs in the strip janitor's per-window loop, and `StackJanitor` —
which starts the grace window that decides whether a missing document is a moment or a fact — runs
*after* that loop. So on the tick a document goes away, the title has already been read. The field
log has the gap: 11:01:43.420 the name changed, 11:01:43.431 the stack noticed. Eleven milliseconds,
and long enough to draw.

The second was found while fixing the first, and it is why the obvious repair would have done
nothing at all: **the painters never read `state->title`.** Both of them, and the tooltip, and the
row's own log line, called `WordTabFrameTitle(frame, ...)` and got the window title as it is this
instant. `state->title` was only ever a change-detector for logging and repainting. Holding a name
there would have held it precisely nowhere.

### The fix

One function answers "what is this tab called", and everything asks it. `StripTabName` returns
`state->title` — the name the janitor last decided on — and falls back to the live title when there
is no strip for that frame or before the first tick has recorded one. Both painters, the tooltip, the
close dialog's heading and the row's log line now go through it, so they cannot disagree.

`ReadTitle` then holds the placeholder back. When a frame's raw title collapses to `Word` or to
nothing while its tab has a real name:

- the first sight is never believed — that one tick is what puts the second sight *after* a
  `StackJanitor` that has seen the same thing;
- and after that it is believed only while `StackIsWaitingFor` says the stack is not still deciding.

A frame that really is sitting in the row with no document — which happens, and which once had this
row showing a tab called "Word" for four slices — still gets its "Word". Half a second later than
before, and only once the stack has agreed the tab belongs there.

### The near-miss, which the battery caught and no assertion did

`StripTabName` was written first as "return `state->title` whenever there is one", which reads like
the tidier rule. It is wrong, and wrong in this fix's own direction: **a strip binds to a frame
before Word has put a title on it**, so the name recorded at bind time is the placeholder, and
serving that for the half second until the janitor's next tick makes a document *opening* flash a tab
called "Word".

No suite asserts anything about it, and all sixteen stayed green. What found it was counting the
defect's own signature in the per-suite logs before and after — `row, left to right:` lines carrying
a `|Word|` tab. Across the battery that count went **29 to 4**: every other suite fell to zero and
`check-onepage` went **0 to 4**, which is the whole of the regression. With the rule as it now
stands it is **29 to 0**, twice over.

So the question is asked of the window first, and `state->title` is consulted only when the window has
stopped offering a name at all. That also gets every stage right with no flag to keep in step: before
`ReadTitle` has ruled, the last real name is still in `state->title` and is what is drawn; once
`ReadTitle` rules that the frame genuinely has no document, `state->title` *becomes* the placeholder
and the live answer and the held one are the same thing.

Tested on the **raw** title rather than the cleaned name, which is the difference between a rule and
a guess: a document actually called Word gives `Word - Word` here and is not this.

---

## 3. A tooltip for a window that was not there

```
16:19:14.483  tip: hwnd=0x0000000000000000  ||  nothing to add (name fits; Word would not say where it is)
```

Twice in two days, both times as a frame was destroyed under the pointer. `TipShow` tests
`hit.frame != g_tipFrame` to decide whether the pointer is still on the tab the wait was armed for,
and when the tab has gone **both sides are NULL**. Two nothings compare equal, so the whole body ran
against a window that does not exist — including a call into Word's object model asking for the
document path of `NULL`.

Guarded at the top: a wait that ends with nothing armed has nothing to show, and says so.

---

## 4. The row shrank and nobody said who did it

Not fixed, because it is not understood. What the log has, at 11:47:26:

```
11:47:26.461  drag  203 position updates (202 moved, 0 resized) over 3842.1ms
11:47:26.471  WM_EXITSIZEMOVE
11:47:26.471  stack  the row remembers its size: (-57,733 1285x1473)
11:47:26.487  WM_SIZE  hwnd=0x...210A2  restored  1239x1397
   ...four more, one per window in the row...
11:47:28.405  drag  ... +0.0ms  CHANGING  pos=(1,673) size=(1261x1410)
```

A move with **no resize in it** ended, and sixteen milliseconds later every window in the row was
24x63px smaller and somewhere else. Two seconds after that the user moved the row again, and that
gesture banked the smaller rectangle as the size they had chosen.

Three things resize these windows — WordTab's own `MatchTo`, Word, and Windows — and the log cannot
tell them apart, because the `WM_SIZE` line carries a client size and nothing else. So the line now
carries the outer rectangle, which is what the row's remembered size is measured in, and says whether
WordTab asked for it:

```
WM_SIZE  hwnd=0x...  restored  1239x1397  outer (7,671 1260x1410)  nobody here asked for it
```

`g_inSync` is already true or false at that instant for its own reasons, so it is free. The drag case
needs no words: this line is not written at all inside a modal move/size loop, so anything reaching it
is either ours or nobody's. The next time this happens the log names which.

---

## 5. The instrument was averaging two different things

Found while reading the report, and it is the reason the largest number in it could not be read:

```
11:47:18.690  strip  dot poll: 1 window(s), 480 passes, mean 232103 us, worst 516165 us
```

That reads as the **active window** costing a quarter of a second a pass — and `PollModified`'s own
comment tells the next reader that a line saying so is the one to look for, because it would mean the
whole design's assumption had failed. It had not. `passes`, `total`, `worst` and `worstEver` were one
set of statics shared by both kinds of pass; `1 window(s)` is only whichever kind happened to run last
before the counter came due, and the mean is over both.

Split, and both lines now name which pass they are describing. The cadence lines have always said
`the whole row` or `the active window`; now the summary and the record do too.

---

## What is deliberately not fixed

**The dot poll still freezes Word for about half a second at a time**, roughly once a minute while the
row sits idle. The governor is meeting the target it was drawn for, and the arithmetic in the report
says so: 480 passes between 07:54 and 11:47 at a mean of 232ms is 111 seconds of blocked UI thread in
three hours fifty-three, and 480 passes between 13:39 and 15:36 at a mean of 115ms is 55 seconds in
one hour fifty-six. Both are **0.8%** of Word's UI thread, against the 1% the ladder was drawn for and
the 5.5% that queued it.

What has not changed is that the cost is spent in half-second lumps rather than as a smear, because
that is what a cold SharePoint `Document.Saved` costs on that rig. Making those lumps smaller is a
different piece of work — not asking the object model from the UI thread at all — and it is not a
defect in what shipped.

---

## Files

| file | what changed |
| --- | --- |
| `src\native\strip.cpp` | `ChromeNear` and `AdoptSampledChrome`; `StripTabName`; `ReadTitle` holds the placeholder; `TipShow` guards an unarmed wait; the dot poll's statistics split by kind |
| `src\native\stack.cpp` | `StackIsWaitingFor`; the row's log line and the close dialog ask `StripTabName` |
| `src\native\frames.cpp` | `WM_SIZE` carries the outer rect and who asked for it |
| `src\native\wordtab.h` | `StackIsWaitingFor`, `StripTabName` |
| `tools\palette-test.cpp` | the adoption gate's decision, old rule beside new, on the field log's colours |
| `tools\check-palette.ps1` | compiles and runs it; 23 checks, no Word |
| `tools\check-all.ps1` | `palette` runs beside `governor` at the front |
| `README.md` | the theme change now takes about four seconds to reach the row, under Known limitations |
