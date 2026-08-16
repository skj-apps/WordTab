# A Protected View document no longer paints every tab yellow

**The defect: open one document you downloaded from the internet, and every tab in every Word window
turns mustard yellow — cards, well, hairline, context menu, all of it. It stays that way until
another window comes forward, and comes back every time the downloaded one does.**

Fixed 2026-08-16. Queued by the chrome slice the same day, which found it while fixing the 38px
defect and deliberately left it: cosmetic, and the night before the work-rig install was not the time
to touch the painter. **Same root cause as the 38px bug — a per-window truth used as a stack-wide
one — and untouched by that fix**, because that fix correctly leaves a Protected View strip under its
message bar. This slice is the other half of it, plus the assertion that would have caught both.

## The cause, in one sentence

The sampler found Word's ribbon by asking for *the flat band directly above the strip*, and above a
Protected View strip that band is the yellow message bar.

## The measurement

Two documents, one plain and one carrying a mark-of-the-web, opened side by side; every visible wide
child of each frame enumerated with its parent chain, and the colour read off each candidate the way
`SampleChrome` reads it — screen DC, 24 points, each point asked who owns it first.

**A normal window's dock holds one `MsoCommandBar` chain. A Protected View window's holds two.**

```
NORMAL                                     PROTECTED VIEW
client (0,0 874x685)  dpi 192              client (0,0 874x685)  dpi 192
strip  y 356..420                          strip  y 318..382
_WwF   y 420..641                          _WwF   y 382..641

MsoCommandBarDock   y   0..356             MsoCommandBarDock   y   0..318
  MsoCommandBar     y   0..356               MsoCommandBar     y 156..318   <- the message bar
    …NetUIHWND      y   0..356                 …NetUIHWND      y 156..318   RGB(109,87,0)
                                             MsoCommandBar     y   0..156   <- the ribbon
                                               …NetUIHWND      y   0..156   RGB(9,9,9)
```

Both chains are five deep and all five report the same rectangle —
`MsoCommandBarDock < MsoCommandBar < MsoWorkPane < NUIPane < NetUIHWND` — so the only thing that
distinguishes the ribbon from the message bar is *where it is*.

The old rule was "widest visible `NetUIHWND`, bottom edge within 4px of our top edge". In the
Protected View window that is the message bar, and the add-in's own log said so:

```
16:52:38.682  strip  palette sampled from Word's ribbon: chrome=RGB(109,87,0) dark
              well=RGB(83,61,0) card=RGB(109,87,0) edge=RGB(135,113,26)
```

`g_palette` is one object shared by every strip, and the poll stops at the first window that answers.
So whichever window happened to answer first spoke for the whole stack, and raising windows in turn
flipped the palette back and forth — `41 → 109,87,0 → 41 → 109,87,0` in one 90-second measurement.

## The finding that decided the fix

The obvious repair — "sample the *top* chrome instead of the thing above the strip" — is wrong, and
one more scan proved it before a line was written. Every row of both windows' top chrome, from the
top edge down:

```
                     NORMAL              PROTECTED VIEW
y   0.. 58           RGB(9,9,9)          RGB(9,9,9)     title bar
y  59..155           RGB(9,9,9) f0.54    RGB(9,9,9)     the ribbon's tab row (low flatness = text)
y 156..356           RGB(41,41,41)       — the chrome ends at 156 —
```

**Protected View replaces Word's ribbon body with the message bar.** The two windows are pixel-for-
pixel identical for the first 156 rows; the normal window then has 200 more rows of `RGB(41,41,41)`
and the Protected View window has none. **The colour the strip needs does not exist anywhere in that
window** — not at a different offset, not at a different row. Reading the top chrome's bottom band
there returns `RGB(9,9,9)`, which is Word's *workspace* colour, and would have been a second wrong
answer wearing the clothes of a fix.

So the honest outcome is that **the window declines**, and that is what it now does.

## The fix

`RibbonHuntProc` checks **both** edges. The ribbon is the widest visible `NetUIHWND` that runs from
the top of the window *down to* the strip's top edge; a candidate directly above the strip whose top
is further down is recorded as an `intruder` and rejected. `SampleChrome` then returns FALSE with a
reason that names what it saw, the poll moves on to the next window, and if none can answer the
palette keeps what it already has.

Nothing in it is specific to Protected View. Any bar Word puts between its ribbon and the document —
`Enable Content`, read-only, sign-in, recovery — has the same shape and gets the same answer.

**The window in front is now asked first, and its answer is the one reported.** Two reasons, neither
of them a preference:

- It is the only window whose pixels can be read at all. The stack holds every window at one
  rectangle, so every other strip's ribbon is *behind* this one and reads back as "something is
  covering Word's ribbon".
- The reason that reached the log used to be whichever window happened to be last in an array. So
  "the front window is showing a message bar" could be reported as "something is covering the
  ribbon" — true of a different window, three steps from the cause, which is the exact shape that has
  cost this subsystem two diagnoses.

And `ApplyPalette`'s log line now names the window it sampled
(`sampled from Word's ribbon in 0x…`), because a line that says only "sampled from Word's ribbon"
cannot tell the healthy case from this bug.

### Accepted cost, stated rather than discovered

**A Word that has only ever had Protected View documents in it keeps the fallback palette** — the one
derived from the Office theme setting in the registry, which the look slice already documents as
approximate because Word rewrites that value at startup. On this rig the fallback and the sample are
the same number. The moment any ordinary document is in front, the real colour is sampled and adopted.

The alternative was to guess the ribbon colour from the workspace colour, and a fallback that happens
to agree with the truth is not a measurement.

## Why the palette stays stack-wide

It was worth asking whether `g_palette` should become per-window, since "a per-window truth used as a
stack-wide one" is exactly what went wrong. **No — and the measurement is the reason.** The ribbon
colour is a property of *Word's theme*, not of a window: both windows above would report
`RGB(41,41,41)` if both had a ribbon body. What is per-window is whether a window is *able to answer*,
and that is what the fix models. Making the palette per-window would have modelled the wrong thing and
cost every strip its own brushes.

Compare the 38px fix, which went the other way and correctly so: a window's chrome *height* really is
its own, so the top edge is never broadcast.

## The second half: the assertion moves to where the fixtures are

`check-stack.ps1` had the per-window assertion **every strip sits between the chrome and the
document** — strip bottom is the document top, and the thing above the strip ends exactly where the
strip begins. It is the assertion the 38px bug violated, and it lived in the one suite with **no
mixed-chrome fixture**. `check-title` and `check-dot`, the two suites that do open a Protected View
document, never ran it.

That is why the 38px bug reported itself as **"the + click did nothing"**, twice, in two different
slices: the `+`'s computed centre had landed inside the ribbon's `NetUIHWND`, three steps from the
cause. It even survived a bisect that blamed the icons slice, on code that turned out to be two `.md`
files.

It is now `Get-StripPlacement` in `tools\WordTabHarness.ps1`, dot-sourced by all eleven suites and
called by three: `check-stack` (every stacked configuration it drives), `check-title` (six windows,
one of them in Protected View) and `check-dot` (three windows, one of them in Protected View).

Two decisions in it worth not re-deriving:

- **It returns an object, not an array of faults.** A PowerShell function that returns an empty array
  returns *nothing*, so `(Get-…Faults $f).Count` is a hard error under `StrictMode` at exactly the
  moment there is nothing wrong. That trap has now cost this project runs under three different
  names; a `pscustomobject` is a scalar and survives the pipeline whatever is in it.
- **`Measured` is part of the answer.** A window with no strip or no document frame is skipped —
  whether it *should* have a strip is check-stack's separate question — so `Ok` alone cannot tell
  "every strip is placed correctly" from "there were no strips to look at". Every call site asserts
  both. Same lesson as the previous slice's *a test that says when it did not measure anything*.

## What the new rule gives up, measured against what it gains

The rule is deliberately strict: the sampled band must run **from the top of the window down to the
strip**. Anything docked in between makes the window decline. That is exactly right for Protected
View, where there is no ribbon body to sample at all — but it is worth being honest that it is
*narrower* than the old rule in one case that was not measured.

**Quick Access Toolbar below the ribbon.** If a user turns that on, Word docks a second full-width
row between the ribbon and the document. Structurally that is the same shape as the message bar, so
the new rule declines — but unlike the message bar, that row is painted in the ribbon's own colour,
so the *old* rule sampled it and got the right answer by luck. On such a machine the strip now falls
back to the Office-theme palette instead of the sampled one.

This was raised by the audit and deliberately **not** fixed, for the reason this project keeps
relearning: **that configuration has not been measured.** Every rule written for it would be a guess,
and the last two defects in this subsystem both came from a plausible rule written ahead of a
measurement. The symptom is the documented `TabThemeSample=0` behaviour, which has an escape hatch,
and the queue carries "measure QAT-below-ribbon" as a small follow-up.

## The adversarial audit, and the two things it was right about

Four lenses over the diff, every claim then put to three skeptics. Fourteen claims. The two that
changed the code:

- **`Get-StripPlacement` would have KILLED check-stack rather than reported a failure.** check-stack
  calls it as `Get-StripPlacement (@($parts) | ForEach-Object { $_.Frame })`, and `$parts` is empty
  exactly when a restore did not take — which is the moment the assertion above it exists to report.
  **The trap is subtler than the usual one and my first test refuted the claim wrongly:** in the
  caller's own scope an empty pipeline is `AutomationNull` and `@(it).Count` is `0`, so it looks safe.
  Parameter binding converts it to a real `$null` on the way in, and `@($null)` is a ONE-element
  array, so the loop runs once with `$null` and `[WordLayout]::Children($null)` throws. Measured
  both ways round before believing either. Fourth shape of this project's empty-collection trap.
- **Section 6a proved nothing on its own.** Both of its claims — "a window reported the docked-band
  reason" and "the palette did not change" — are *also* exactly what a build whose rule rejected
  **every** ribbon would produce. The suite would have gone green over a sampler switched off by
  accident. Section 6c is the control: bring an ordinary window back to the front and require the
  sampler to succeed again.

Three more were fixed as written-down weaknesses rather than live bugs: the `Measured` assertion was
compared against the length of the list it was computed from (so `0 -eq 0` passed having measured
nothing — it is asserted against **three** now); the "nothing above the strip" fault described the
38px bug as "it is sitting at the top of the window", which is a wrong story about the right failure,
and now names the strip top and the chrome that overlaps it; and the decline message asserted a cause
("a bar sits between…") that the evidence only supports as "the band above the strip does not reach
the top of the window", so it now says what was measured.

The rest were refuted, several of them correctly pointing at behaviour that is deliberate — the
change-only logging, and `ApplyPalette`'s early-out.

## A hazard found off to the side: the log rolls, and the harness could not see it

`src\native\log.cpp:42` deletes the log once it passes 512KB. A full battery generates enough to trip
that about once, and **it happened mid-`check-title` during this slice's first battery.**

Every suite reads the log from a byte offset, and the shared shape was
`$offset = [Math]::Min($mark, $stream.Length)` — which after a roll silently reads from the *end* of a
fresh file. The two assertion shapes then fail in opposite directions: "the log says X since the mark"
goes red for a reason that is not the product, and **"the log says nothing since the mark" goes green
having read nothing at all.** Section 6a's palette assertion is exactly the second shape.

`check-dot`'s `Get-LogSince` now reports a shrunk file instead of clamping. **The general fix is not
in this slice** and is queued: `Set-LogMark`/`Get-LogSince` exist in six copies across the suites, so
hoisting them into `WordTabHarness.ps1` — with a marker line rather than a byte offset, which also
closes the rolled-then-regrew hole — touches all eleven and needs its own battery.
[[one-copy-of-what-decides-truth]] again, third time.

## How the palette fix is proved

`check-dot` section 6a, at the moment the Protected View window arrives, with the log read from a
mark taken immediately before it opens:

- **positive** — `chrome sample: a bar sits between Word's ribbon and this window's strip` appears.
  Without this, "no yellow" is also exactly what a sampler that had quietly stopped polling produces.
- **negative** — no `strip  palette` line at all: the palette did not move. `ApplyPalette` is its only
  writer and logs every change, so no line means no change.
- **control (6c)** — an ordinary window is brought back to the front and the sampler must succeed
  again, so neither of the above can be satisfied by a sampler that answers nothing at all.

Deliberately **no new suite and no new probe**, and no pixel assertion here: every colour drawn comes
from that one object, and `check-look` already photographs it against Word's own ribbon. The
measurement that designed the fix was scratch and is not in the repo — the durable instrument is the
log line in the product.

## The mistake that cost the most time, and it was mine

**I ran the adversarial audit as background agents while the suites were driving Word.** This file's
own notes say it, in capitals, from a previous slice: *they cannot run in parallel — one Word, one
desktop — and two overlapping runs will wreck each other.* I had read that this session and applied
it to suites, not to agents.

What it produced, in order, each looking like a real product defect:

1. `CO_E_SERVER_EXEC_FAILURE` creating `Word.Application`, leaving a half-started WINWORD.
2. A `check-dot` run where the Protected View window vanished mid-suite — the add-in's log showed a
   *second* Word process exiting one second after it appeared, and that frame's tab name changing
   from `Downloaded.docx` to `Word`.
3. Word coming back with a document frame 151px tall, too small for `Set-Dirty` to type into.
4. Three failures in the save section, from Office floating chrome parked over the strip.

None of them was the product. **And the workaround for (3) — resizing Word's window to 1400x1050 —
became the cause of (4)**, because the suite passes at the geometry it was written against and I had
changed a real, persisted Word setting to get past a symptom. Putting `AppWindowPos` back to the
user's own `1523,112 900x700` made `check-dot` go 68 of 68 with no code change at all.

Two rules out of it. **The no-parallel rule is about the desktop, not about suites** — anything that
can open Word, take a screenshot or move the pointer counts, including agents I launched myself. And
**when a symptom appears mid-slice, restore the environment before changing it**: the first move
should have been "what else is touching this machine", not "make the window bigger".

The one good thing: the assertion that caught (2) was the one this slice added — `all three windows
are still open to be measured (got 2)` — which reported the real state instead of measuring two
windows and passing. That is the whole point of `Measured`, and it earned its keep against a cause
nobody was looking for.

## Battery

**All eleven green against this build: 457 checks, 0 failures, 13.2 min.**

```
  stack        41   0.5 min      look         48   2.0 min
  strip        40   0.4 min      scroll       49   2.2 min
  tabs         26   0.6 min      title        39   1.2 min
  menu         37   1.1 min      dot          68   1.6 min
  reorder      37   1.6 min      soak-stack   22   1.1 min
  startscreen  50   0.9 min
```

441 → 457, and the 16 new checks are exactly the ones this slice added: `check-stack` +6 (the
`Measured` assertion, once per configuration `Test-Stacked` drives), `check-title` +3, `check-dot` +7.
Nothing else moved, which is the point of a slice that touches the painter's *input* rather than its
output — `ComputeLayout` is untouched and no rectangle moved, so every assertion written before this
still measures the same thing.

**The log is the stronger half of the proof, and it did not roll this run** (250KB of the 512KB
limit, one continuous file from 17:01:31 to 17:14:36 covering the whole battery), so every log-based
assertion read what it claimed:

- the new rule fired **exactly twice** — once in `check-title` and once in `check-dot`, which are
  exactly the two suites with a Protected View fixture, and nowhere else in thirteen minutes of
  driving Word;
- **the palette moved twice, both times correctly**, and both from `check-look` flipping Windows'
  theme: `chrome=RGB(255,255,255) light` then `chrome=RGB(41,41,41) dark`. That also exercises the
  new diagnostic, which the audit had correctly pointed out was unexercised because the fallback and
  the sample are the same number on this rig:

```
17:07:18  strip  palette sampled from Word's ribbon in 0x00000000008D10F0: chrome=RGB(255,255,255) light …
17:07:30  strip  palette sampled from Word's ribbon in 0x00000000008D10F0: chrome=RGB(41,41,41) dark …
17:11:03  strip  chrome sample: something is docked between Word's ribbon and this window's strip …
17:12:44  strip  chrome sample: something is docked between Word's ribbon and this window's strip …
```
