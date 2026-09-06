# The 2026-09-04 queue, worked end to end

Input: `NOTE-queue-20260904.md`, six items, worst-first, "do all of it, no triage". Plus the two free
items in its Refuted section. All of it is done. `RESULT-unseen.md`'s two fixes are untouched and
still in the tree beside this work.

**Everything below was built on the working tree at `ef49c79` + the two unseen fixes.** The note's
line numbers are all **19 low for `strip.cpp` and 48 low for `stack.cpp`**, because it was written
against `HEAD` rather than the tree it describes. Every mechanism it names is real and was found
exactly as described; only the addressing had moved.

---

## The headline, and it was not on the queue

### The chrome sampler costs half a second to eight tenths of a second, every two seconds, on Word's UI thread

Item 2 asked for QPC brackets around `SampleChrome` as a side errand - *"the only janitor stage with
no instrument"*, a *"live candidate"* for the drag gaps. The instrument went in and answered
immediately:

```
strip  chrome sample: 20 passes, mean 512856 us, worst 752096 us
strip  chrome sample: 20 passes, mean 743963 us, worst 782110 us
strip  chrome sample: 796719 us (the most it has ever taken)
```

**Mean 512,856-743,963us. Worst 796,719us.** On this rig, where the dot poll that the entire cadence
governor was designed, measured, rewritten and re-argued for costs **about 430us**. Three orders of
magnitude, on the thread Word draws with, every fourth janitor tick - which is every two seconds for
as long as Word is open with a ribbon showing.

**The cause is established, off the add-in entirely.** `tools/pixel-cost.cpp` does nothing but the
three calls `SampleChrome` makes, on this machine, in a standalone process:

```
round 0:  GetDC(NULL) 34 us   24x GetPixel 929092 us   24x WindowFromPoint 87 us
round 1:  GetDC(NULL)  9 us   24x GetPixel 968250 us   24x WindowFromPoint 79 us
round 2:  GetDC(NULL) 10 us   24x GetPixel 1284716 us  24x WindowFromPoint 1336 us
```

`GetDC` is free. `WindowFromPoint` is free. **`GetPixel` against the screen DC is ~40ms per call**,
because each one forces a readback out of a DWM-composited surface, and `SampleChrome` makes
twenty-four of them per strip per sample.

This is not a regression. It has been there since the sampler was written and it has simply never
been timed - which is exactly what the note predicted about the one stage with no instrument, and
why "measure the thing nobody has measured" was worth doing for its own sake.

### And then it was fixed: 744ms -> 21ms, but NOT the obvious way

One `BitBlt` of the 1-pixel-tall band the samples lie on, into a memory DC, and `GetPixel` against
*that*. One readback instead of twenty-four. Same points, same `y`, and `WindowFromPoint` still asked
per point **against the screen** - that is the question that makes reading the screen honest, and it
was never the expensive half (79-1336us for all 24).

| | mean | worst |
|---|---|---|
| before | 743,963us | 796,719us |
| after | **20,749us** | **34,074us** |

**THE TRAP, AND IT IS THE PART WORTH KEEPING.** The first version used a plain `SRCCOPY` and it was
**wrong**. It looked perfect: the adopted palette was byte-identical on a static ribbon, both themes,
every derived colour. Then `check-look`'s live theme change failed **12 of 48** - Word's ribbon was
measurably `RGB(255,255,255)` on screen and the blit kept handing back the old dark pixels, so the
palette never re-derived. All twelve cascaded from that one stale read: a dark well under a white
ribbon paints dark-on-dark, which is why even *"the empty row says something"* went red.

**This is the same redirection-surface staleness that already rules out reading the ribbon's own
window DC** - written down in this project months ago and in `reading-a-colour-off-another-window`.
A plain `SRCCOPY` off the screen DC is entitled to come from a cached surface. **Per-pixel `GetPixel`
was ACCIDENTALLY immune** because it forces a readback every call, and that immunity was the thing
being traded away for the speed, invisibly.

`SRCCOPY | CAPTUREBLT` asks for what is on the screen now. With it, `check-look` is **48 of 48** and
the theme change reads exactly right:

```
the well before the change   RGB(15,15,15)
the well after the change    RGB(229,229,229)
PASS  the strip changed colour without Word restarting
PASS  and it changed in the right direction - lighter, with Windows
PASS  the derivation still holds against the new chrome (26 levels, expected 26)
```

**The lesson, which outlives the fix: a static A/B could not see this.** Byte-identical output on a
ribbon that is not changing is exactly what a stale cache produces. The only instrument that could
tell the difference was the one test that *changes the input while the add-in is running* - and it
already existed. The drag stand-down stays: it costs nothing, and 34ms inside a drag is still 34ms.

**And it may be why `check-reorder` is "the flaky one".** That suite drives an injected-input tab
drag and its one intermittent failure is always the same: *"the card follows the pointer, pixel for
pixel"* - on the final battery, `pointer moved 180px, the card moved 150px`, with the suite's own
guard confirming the pointer really did move. A 0.75s freeze of Word's UI thread arriving every two
seconds is exactly what loses 30px of pointer travel, and its log carries samples of 748,812us and
750,828us seconds either side of the tear-off. Not proved - but **falsifiable, and cheaply**: fix the
readback and `check-reorder`'s flake rate should fall. That is the test to run after the next slice,
and it is a better one than re-running it and hoping.

**It is also the best candidate yet for the work rig's drag gaps.** Their traces carry 260.8ms and
173.3ms gaps. A dot poll costing 179,188us was the standing explanation; a chrome sample costing
750,000us is a better one, and it rides the same timer into the same modal loop.

---

## What each item got

### 1. The row grows 21-27px on every drag - PROBE ONLY

As instructed: *"Do the PROBE first, and only the probe, until it has been read."* No guard was
written, and the note's trap - a guard keyed on `!g_inModalLoop` silently springing back Win+Left and
Win+Right, which `frames.cpp:485` already names as unguardable - is therefore not reachable.

`ProbeUnaskedResize` in `frames.cpp` fires on a `WM_WINDOWPOSCHANGING` outside the modal loop that
changes the size, and prints the current rect, the proposed rect, the growth, **the non-client margin
`AdjustWindowRectEx` itself reports**, **`rcNormalPosition` beside `GetWindowRect`**, the flags,
`hwndInsertAfter`, `IsZoomed`, the DPI context, and **ms since that frame's own `WM_EXITSIZEMOVE`**.

Those fields are chosen to separate the note's two leads and to be able to reject both:

| Lead | What settles it |
|---|---|
| something re-asserting a *stored* rectangle | the proposal printed against `rcNormalPosition`, which IS the stored restore rect |
| an outer rect read as a client rect | the growth printed against the frame's real non-client margin |
| neither | both comparisons come back negative, and `sinceDragMs` says whether it is even the post-drag message |

Rate-limited to one line per 250ms, chosen from the thing measured (the growth lands at +48ms, the
burst is three messages in one tick), and **it says how many it swallowed** rather than reading as
"this happened once".

It already fires correctly on the dev rig, on a maximize, which is a resize nobody in-process asked
for and is the shape it is looking for:

```
resize  hwnd=0x350560  NOBODY HERE ASKED: now (40,40 940x820) -> proposed (-13,-13 3286x2132)
        grew +2346x+1312  non-client margin 26x26  rcNormalPosition (40,40 940x820)
        flags=0x... zoomed=0  ... 0ms since this frame's drag ended
```

### 6. The strip at the other monitor's DPI - PROBE ONLY, folded into item 1's

*"Fold this probe and item 1's into ONE probe with two callers"* - done literally.
`StripDescribeDpiContext` is one helper in `strip.cpp`, declared in `wordtab.h`, called from the
strip-drift line and from item 1's resize probe. It reports the window's DPI awareness context, the
calling thread's, and the raw DPI, all via `GetProcAddress` off `user32` in the exact pattern
`DpiOf` already used - because w64devkit's headers predate every one of those APIs (verified: they
contain no `GetDpiForWindow` and no `DPI_AWARENESS_CONTEXT` at all).

One thing had to be split to make the probe honest: `DpiOf` applies the `TabDpi` override, so a probe
built on it would agree with `state->dpi` **by construction** and could never show the disagreement it
exists to look for. `DpiFromWindows` is now the raw read and `DpiOf` is the override in front of it.

The drafted "suspect read, don't correct" guard was **not** written, per the trap: it fires on the
restore, where the strip genuinely was 1462x58, and would turn a self-healing 1s flicker into a
permanent 244px overhang.

### 2. The janitor inside Word's modal drag loop - FIXED

`FramesInModalMoveLoop()` exported from `frames.cpp`. `PollModified` stands down on it exactly as it
already does for `StackCloseInFlight`, and says how many passes a gesture cost:

```
strip  dot poll: %d pass(es) skipped while Word was inside its own move/size loop
```

**Only that stage, per the note's constraint** - the per-strip loop above it (`PlaceStrip`) is what
keeps the strip on its frame while the frame moves, and stopping it would trade a stutter for a
visibly lagging strip.

`g_inModalLoop` **can stick**: it is set only by `TraceReset` on `WM_ENTERSIZEMOVE` and cleared only
by `TraceFlush` on `WM_EXITSIZEMOVE`, and a frame destroyed mid-drag never sends the second message.
The accessor unsticks it on `!IsWindow(g_modalWindow)` and says so; a stand-down that somehow
outlived that is made loud by a line every 120 skipped passes rather than silently stopping the dot
for the life of the process.

The comment that hid this for so long is corrected too: *"Word's message loop dispatches it"* was the
false half - during a drag Word does not return to its message loop at all, which `frames.cpp` says
in its own words thirty lines away.

**And the chrome sample now stands down on the same query**, which the note did not ask for and the
measurement above is the entire reason for.

### 3 + 4. The dot-poll governor - FIXED, landed together

**Item 4**, one line: `else if (RungFor(basis * 2) < *rung)`. A pass must now be cheap enough that
*twice* its cost would still sit lower before the cadence is trusted back down. `basis * 2` can only
map to a rung >= `RungFor(basis)`, so it is strictly harder to descend and cannot reintroduce
`f333349`. `want` is still used at the climb, so `-Werror` is satisfied. No overflow: `basis` is a
64-bit microsecond count and the worst pass ever measured is 520,183us.

**Item 3**, the cap: `*rung = (want > 1) ? 1 : want`. The active pass keeps its memorylessness - the
subject changes underneath it when the user switches tab - but is never asked less often than every
two seconds. Both misleading comments are corrected rather than cited as licence: the one claiming
the typed-in tab *"is still asked twice a second however far the row has backed off"* now says
plainly that this was weakened and why, and the one claiming the active cadence line *"has never yet
appeared after the first pass of a process"* now records that it has, seven times, worst 52056us ->
20000ms with 35.35s observed. A **third** copy of the same invariant in the section header 185 lines
above was found and corrected too - it still said "at the same governed cadence", which is now two
different rules.

The governor line names its window: `(hwnd=0x...)`, appended **after** `now asking every %d ms` so
`check-dot.ps1`'s regex is untouched.

**Mirrored by hand into `tools/governor-test.cpp`**, as `check-governor.ps1` demands in as many words.

### 5. OnePage's five meanings of FALSE - FIXED

`WordTabOnePageView` now decides a reason at each of its exits and hands it back; `stack.cpp` prints
what it was given. **The illness test is not re-derived anywhere** - the trap the note says two
agents already fell into, and the field data proves why:

```
view  hwnd=0x...  left alone at 99 page(s) across, 100% - already on one page at a readable zoom - nothing to do
```

**99 columns, 43 times in one battery.** 99 is Word's "as many as fit" and IS the healthy value.
Any version of this that dropped the `columns < 99` would have shouted "WORD REFUSED THE WRITE" on
every one of those.

The fourth reason exists because `columns == 0 && percent == 0` is genuinely ambiguous: the
out-params are written unconditionally, so it is also what a `Zoom` object that answered and then
refused both property reads leaves behind. Those are now different words.

The instrument found something on its first run, which is the point of it: **two Protected View
windows** logged `no Word window claims this frame`, a case that was previously indistinguishable
from a healthy join.

### The free items

**The report names a rolled log.** The collector is `install/settings.ps1 -Report`, **not**
`tools/workrig-recon.ps1` as the note says - that file is the pre-deployment AppLocker/WDAC recon and
never reads the log at all. The log section now prints the first line verbatim and says whether it is
a `---- DllGetClassObject ----` startup line; if it is not, the log rolled at 512KB and is not the
whole history. (40 of 40 preserved logs on this rig begin with that string.)

**The report reads the view.** Also `install/settings.ps1`, for the same reason, and it needed a
guarded object-model attach because that whole section is Win32 `EnumWindows`. It prints pages-across
and zoom per window, flags `1 < columns < 99` as the fault, and - free - prints the object model's
window count against the Win32 OpusApp count, because a gap between them is the Protected View
signature.

**One guard above both `ReadTitle` and `ApplyInitial`.** A frame Word has made and not titled binds
nothing until it is titled or on screen. The measured cost of not having it: **116 empty-title binds
across 40 battery logs**, each one asserting a tab name for a window with no tab and deriving a
`natural` rect from an unlaid-out frame - top edge 156 where every finished frame in the same log
reads 356.

**The predicate is the title, not visibility, and that correction came from measurement.** Reasoning
from the code alone gives `IsWindowVisible` - and it is wrong: *every* frame on this rig attaches
`visible=0`, including every good one, so a visibility-only guard defers every bind there is. Off
screen is the second half of the AND, to bound the cost of the title test being wrong. `"Word"` is
deliberately **not** in the predicate: that is the Start screen, the last-document-closed window, and
Protected View for its first 31ms.

---

## The A/B

Baseline is archived out of the 40-file prune window in
`%LOCALAPPDATA%\WordTab\history\baseline-ef49c79-wt\` - 16 suites, 851 checks, against a binary whose
SHA256 matched the built DLL byte for byte.

| | baseline | after |
|---|---|---|
| checks | 851 | 863 |
| `check-governor` | 16 | **21** |
| `check-onepage` | 11 | **18** |

Said exactly: the last full battery run against this binary was **14 of 16 green**, with `rowsize`
and `reorder` red. `rowsize` was the suite bug described below and is fixed and green three runs in a
row since; `reorder` is the documented injected-input flake. The binary has not changed since that
battery - only `tools\check-rowsize.ps1` and this file.

**The new assertions were tested against a reverted build, not assumed.** A copy of
`governor-test.cpp` with only the two governor fixes backed out fails exactly three of the new checks
- the 2s ceiling, the sweep across every cost on the ladder, and the ping-pong replay (7 cadence
changes reverted against 1 fixed). Checks that pass either way would have pinned nothing.

Two suites went red on the full runs and **neither was a regression**, but only one of them was
left alone.

**`rowsize` was a real suite bug, and the first fix only found half of it.** It failed
*"the add-in said it was too wide and narrowed it"* in two runs out of three while every behavioural
assertion beside it passed - the window really did come up 856px, about one page. The add-in was
right and the suite read too early. `Wait-Joined` waits for `joined as the first window`, and the
line the step asserts is written **about 20ms after** it (19ms, 21ms and 23ms measured across three
batteries).

Adding a wait for that one line made step 6 green three runs in a row - **and the next battery failed
step 7 instead**, on `the row takes the size it was left at`, with its behavioural partner
("the row is the 1714px it was told to be") passing beside it. Same bug, different line: the suite has
**five** positive log assertions that follow `Wait-Joined`, and `Wait-Joined` is a barrier for none of
them. All five now wait for the line they are actually about, via one `Wait-Logged` helper. Fixing
the one that happened to fail would have left four loaded.

**The lesson is the general one and it is worth more than the fix:** a test that waits for event A and
then asserts on event B is not flaky, it is wrong, and it will keep producing "known flaky" results
until somebody times the gap. Both failures had a passing behavioural assertion sitting next to them
saying the add-in had done the right thing, which is the tell.

**The last battery of the night was contaminated and is not evidence of anything.** A game
(`Mewgenics.exe`) was running on the dev rig and stole the foreground **17 times** during the run.
The chain is worth writing down because it produced eight failures that look like a geometry
regression and are not one: the game took the keyboard during `check-stack`, which then threw
`Index was outside the bounds of the array` and **aborted mid-suite**, which left three Word windows
up, so `check-row` started against them and reported `the row has four tabs (has 7)` eight ways.
`check-all.ps1`'s own header warns about exactly this - *"Word must be closed before the first one
starts. A suite that begins with the previous run's windows still up measures something else, and
has."*

**Before trusting any red suite here, check what else is running.** Every suite that does not depend
on the foreground was green in all three batteries. And `rowsize` passed 41/41 in that same
contaminated run, which is the context it had failed in twice - so the `Wait-Logged` fix is confirmed
where it matters.

**`reorder` is the documented flake and was left alone** - see the paragraph above on the chrome
sampler, which is now a concrete and testable explanation for it rather than a shrug.

`check-onepage` went red for a better reason: **the suite's own comment was wrong and my first
assertions believed it.** It claimed a second document into a corrected Word "must produce no line at
all"; the paragraph below it in the same file already knew better - correcting the columns does not
make Word recompute the crushed zoom, so the second document arrives one page across at 10% and is
the *zoom* half of the same illness. The assertions now match what the add-in actually does, and the
comment says what was wrong with the old claim.

---

## What is still open

1. **The chrome sampler's 24 `GetPixel` calls.** Measured, understood, unfixed. Next slice.
2. **Item 1 is a probe and nothing more.** The ratchet is still there; the next work-rig report is
   what the probe was built for. **The falsifiable prediction stands**: their next Word start should
   log `the row takes the size it was left at ... (-1,667 1261x1410)`, wider than the `1240x1397` it
   has started at twice.
3. **Item 6 is a probe and nothing more.** If the frame's and the strip's DPI contexts ever print
   differently, that is mixed-mode virtualisation of our own child and it is ours to fix.
4. **`stack` and `rowsize` read the log on a race.** ~20ms, pre-existing, now observed failing.
5. Nothing here is committed or published.
