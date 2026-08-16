# Every window's chrome is its own

**The defect: open a document you downloaded from the internet alongside any other document, and every
other tab's strip jumps 38 physical pixels up into Word's ribbon, with the `+` drawn underneath
`NetUIHWND` where it cannot be clicked. It never comes back.**

Diagnosed and fixed 2026-08-16. It was the queued slice, it was pre-existing, it had been met twice
before and never explained, and **the first question in the brief — "can a user reach this by hand?"
— is answered YES, in one move**, by the single most ordinary thing a corporate user does with Word:
opening an attachment.

## The cause, in one sentence

A Protected View window's document frame legitimately sits 38px higher than a normal window's, and
the stack broadcast copied that window's interior onto every other window as though every window in a
stack must have the same chrome.

## The measurement

From the add-in's own log, `pid=8604`, one Word process, seven frames, the moment it goes wrong:

```
15:18:36.534  hwnd=0x740662  tab name |Downloaded report.docx| from |… - Protected View - Word|
15:18:36.534  hwnd=0x740662  bound _WwF=0x6A1120  dpi=192 stripH=64  _WwF at (0,156 874x485)
15:18:36.534  hwnd=0x740662  initial: _WwF natural (0,156 874x485) -> (0,220 874x421)
15:18:36.936  hwnd=0x740662  initial: _WwF natural (0,318 874x323) -> (0,382 874x259)
15:18:36.942  stack  active -> 0x740662
15:18:36.944  hwnd=0x35068C  the active window changed and pushed its interior out: (0,318 874x323)
15:18:36.948  hwnd=0x4E1092  the active window changed and pushed its interior out: (0,318 874x323)
15:18:36.951  hwnd=0x4010CC  the active window changed and pushed its interior out: (0,318 874x323)
15:18:36.953  hwnd=0x5C1116  the active window changed and pushed its interior out: (0,318 874x323)
```

Nineteen milliseconds. Every one of those four windows had logged `initial: _WwF natural (0,356 …)`
when it was created, which is what a normal window on this rig reads, in every session in the file.

And then the seventh window, which was still to come, and which is the one `check-title` measures:

```
15:18:39.300  hwnd=0x5509DA  initial: _WwF natural (0,356 874x285) -> (0,420 874x221)
15:18:39.454  hwnd=0x5509DA  joined the stack, taking the master's interior: (0,318 874x323)
15:18:39.454  hwnd=0x5509DA  strip had drifted: at (1,357 874x64), expected (0,318 874x64) - correcting
```

**It read its own layout correctly and was overwritten 154ms later**, by `MatchTo`, with the interior
of the Protected View window it had joined behind.

**318 is not a transient and that had to be proved, because the fix depends on it.** If 318 were a
half-built ribbon then the Protected View window was wrong too and the answer would be to wait for it
to settle; if 318 is that window's real interior then the answer is that it must keep it and nobody
else may take it. The add-in's own chrome sampler settles it without a new probe, because it only
accepts a band whose bottom is within 4px of the strip's top:

```
15:18:39.180  strip  palette sampled from Word's ribbon: chrome=RGB(109,87,0) dark …
15:18:39.181  strip  chrome sample: ok                     <- on the Protected View window, strip at 318
15:18:40.437  strip  chrome sample: no NetUIHWND of the right width sits directly above the strip
                                                            <- on a normal window, strip at 318
```

A real chrome band ends at 318 in the Protected View window and there is nothing there in a normal
one. Word settled that window at 318 and left it there for the 2.5 seconds it was active, and Word
lays out the window that is active. **318 is correct for that window and wrong for every other.**

**The arithmetic, for anyone reading this on another rig:** `dpi=192`, so 200%. Word's settled top
chrome on this machine is 178 logical px — the same number at three different scalings across this
repo's history (178 DPI-unaware in spike 1, 267 at 144 dpi in `RESULT-strip.md`, 356 at 192 dpi here).
The Protected View frame is 159 logical. The difference is **19 logical px, 38 physical**, and that is
why no ribbon *state* explains it: a collapsed ribbon is 100 logical, the dismantle on closing a
document is the whole 178, maximise-versus-resize is 72. **There is no literal 38 anywhere in the
placement code, and there was never going to be.**

## The fix

`StripSetNatural` — the one writer of a window's `natural` rect that Word did not propose — now keeps
the receiving window's own top edge:

```c
RECT want = *natural;
if (state->hasApplied)
    want.top = state->natural.top;
```

**The stack holds every window at one rectangle, so the left, right and bottom edges of one window's
document frame are true of all of them. The top edge is not a property of the stack at all** — it is
the height of that window's own chrome, and Word does not give every window the same chrome. Three
call sites broadcast (`MatchTo` on join and on reconcile, `StackOnFrameActivate`, `StackOnActiveLayout`)
and all three were pushing an edge they had no business pushing.

**Why nothing healed it, which is the half that makes it permanent rather than a flicker:** Word lays
out only the focused window. That rule is the whole reason the broadcast exists — and it is also the
reason a wrong value can never be re-derived. `Reconcile` compares outer rectangles twice a second and
declares a stack whose interiors are all wrong to be in step, because from the outside it is.

**Accepted, and stated rather than discovered later:** a background window no longer gets a proactive
top correction from the active one. If Word ever changes every window's chrome height while they are
stacked — collapsing the ribbon, say — a background window keeps its old top until Word lays it out,
which happens when it comes forward. It is invisible while it is behind another window, so there is
nothing to see; and taking a number that is right for one window and wrong for another is what this
whole slice is about.

## The proof

**The full battery: 441 checks, 0 failures across all eleven suites, 13.1 minutes** — from
`check-title` 36 and `check-dot` 61, which are the two that were red. But a green run is the weaker
half of the evidence, because this defect was intermittent enough to survive a bisect. The stronger
half is the same battery's log:

- **Every `natural.top` recorded in thirteen minutes of driving Word: `318` twice — the two Protected
  View frames, one in `check-title` and one in `check-dot` — and `356` a hundred and sixty-one times.**
  Each window's own chrome, and nobody else's. (Also `156` ×21, a bind that catches Word mid-build and
  is corrected within the second, and `0` ×18, the layout Word leaves behind when a document closes,
  which the height tripwire refuses.)
- **The moment that used to do the damage now passes without an entry.** `stack active -> 0x43067A`
  at 15:33:59.828 is the Protected View window taking the foreground; no window's interior changes.
- **`stack joined, snapped to 0x43067A` at 15:34:02.359** is a window joining *behind* the Protected
  View window — the `MatchTo` path that overwrote the seventh window's correct 356 — and it keeps its
  own top.
- **Every broadcast that did fire in the battery carries `top=356`**, the receiver's own, including one
  `(0,356 874x274)` where a real height change was propagated. The mechanism still does its job; it
  has stopped doing the other one.

## What the instrumentation cost, and why it existed at all

Two changes to the log, both of which are why this took one run instead of another undiagnosed slice:

- **`StripSetNatural` logged nothing.** Two of the three writers of `natural` announced themselves and
  the third — the only one that can write a number Word never proposed for that window — moved `_WwF`
  in silence. It now logs, and `why` names the path that asked, which is how the four broadcast lines
  above exist to be read.
- **`LogRelayout`'s 250ms throttle was dropping the evidence.** The whole event above spans 19ms. The
  throttle now never suppresses a move of the **top** edge, only of the other three: the flood it was
  built for is a resize drag, which moves the other three every ~15ms and leaves the top alone, so
  exempting the one edge this file is about gives nothing back. **The previous failing run produced no
  usable log for exactly this reason, and its log was then deleted anyway** — `uninstall.ps1:70`
  removes `%LOCALAPPDATA%\WordTab` wholesale, and the clean-deployment round trip ran 25 minutes after
  the battery that found the bug. **The round trip destroys the evidence of whatever ran before it.**

## The icons slice's bisect was wrong, and the reason is worth keeping

`NOTE-no-icons.md` records this exact symptom at `956ab7c`, blames the icon drawing code, and says
"confirmed by bisect and NOT by argument": `check-title` went 34-with-3-failed with the icon code and
35-all-passed with `strip.cpp` stashed and rebuilt. **`git diff 956ab7c HEAD -- src/native` is two
`.md` files and no product code at all** — the same failure has now been reproduced on code that is
byte-identical to the version the bisect exonerated. The bisect measured **whether the Protected View
window happened to take the foreground that run**, twice, and read it as a property of the build.

**A two-sample bisect on an intermittent fault produces a confident wrong answer, and it feels like
the good kind of evidence while it does it.** The tell was there and was written down and not
followed: "cause never diagnosed."

## The repro, compacted

The trigger recorded in `RESULT-harness.md` was `look` → `scroll` → `title`. **`check-scroll` is not
part of it**: `look` → `title` reproduces, in **3.7 minutes** (48 checks + 35 checks) instead of three
suites. `check-look` is not obviously part of it either — the cause is "a Protected View window becomes
active in a stack", which `check-title` builds on its own — so the ordering is most likely raising the
odds of a race for the foreground rather than leaving state behind. That matches `check-title` alone
having passed 36 of 36 once.

**The structural fact worth more than the ordering: `check-title` and `check-dot` are the only two
suites that open a Protected View document** (`check-dot.ps1:523` writes the same `Zone.Identifier`),
and they are exactly the two suites that went red. Nothing else in the battery builds a stack whose
windows have different chrome, which is why the six-plain-documents probe in the harness slice could
not reproduce it from geometry alone. **The fixture is the reproducer; the ordering is noise.**

## Still open, found by this slice and deliberately not fixed in it

**A Protected View window turns every tab in every window yellow.** `g_palette` is global
(`strip.cpp:349`), the sampler takes the flat band directly above the strip, and above a Protected View
window's strip that band is the yellow message bar: `chrome=RGB(109,87,0)` in the log above, on a rig
whose Word is dark grey. It is the *same root cause as this slice* — a per-window truth used as a
stack-wide one — and it is untouched by this fix, because the geometry fix correctly leaves that strip
sitting under that bar. It is a cosmetic bug where this one was a functional bug, so it is queued
rather than folded in on the night before an install.
