# A dot for a document you have not saved

**PASSED.** 61 checks in the new `tools\check-dot.ps1`; the other ten suites re-run — **409 checks,
0 failures across the eleven.** Writeup for the first slice that reads Word's *documents* rather than
its windows.

| suite | checks | | suite | checks |
|---|---|---|---|---|
| stack | 34 | | look | 43 |
| strip | 40 | | scroll | 46 |
| tabs | 22 | | title | 35 |
| menu | 32 | | **dot** | **61** |
| reorder | 30 | | soak | 22 |
| startscreen | 44 | | | |

A tab whose document has unsaved changes draws a filled dot where its close button is. Move the
pointer onto the tab and the **×** comes back, so nothing is lost and nothing is added: the dot lives
inside a rectangle that already existed.

---

## The width argument, which is why this is not a dot beside the name

The brief carried a decision — a mark beside the name, or the VS Code trick where the **×** *is* the
mark until you hover. It was settled on a measurement already in the repo rather than on taste, which
is the second time in this project a "product decision" has turned out to be a number.

The scrolling-row slice measured Word's own default window on this rig at **874 strip px, which is
five tabs at 200%**. Six documents in a freshly opened Word already overflow. A mark beside every
name costs width on every tab forever, and makes every row scroll sooner, in order to display
information that is usually *no*. Putting it in the close button's square costs nothing at all.

That has a second consequence worth more than the first: **`ComputeLayout` is untouched, so no
rectangle moved.** `tools\WordLayout.cs` — the hand-maintained mirror the check scripts click
through — needed no edit, and every assertion written before this slice is still measuring the same
thing afterwards. Same discipline as the look slice, and the same payoff.

The cost is stated plainly: **a tab with no close button has no dot.** Too narrow to carry one, or
partly outside the track, or `TabButtons=0` — no square, no mark. That is not a gap so much as the
same scarcity answered the same way twice, and the alternative is a rect of its own in
`ComputeLayout`, which is a much larger slice and puts 348 existing assertions back in play.

---

## The measurement came first, and it changed the failure model

`tools\probe-saved.ps1` drives Word through the states a user can put a document in and prints what
`Document.Saved` says. Five findings, none of which were safe to assume:

- **`Saved` is TRUE when the document has NOT changed.** The dot is its negation. A brand new,
  untouched document reports `Saved = True` — so an empty `Document1` correctly carries no dot, which
  is also what Word itself thinks: it does not prompt when you close one.
- **`Saved` is settable.** Nothing in the add-in ever writes it, and the comment on
  `WordTabReadModified` says so.
- **`Window.Hwnd` matches the `OpusApp` frame handle for every window at once**, not only the active
  one. The menu slice had measured it for the active window; this is what makes it possible to find a
  document without activating anything, and therefore safe to do on a timer at all.
- **A full pass costs 2.04 ms cross-process for two windows** — a real upper bound, since the add-in
  is inside `WINWORD` and pays no marshalling.
- **A Protected View document is in neither `Application.Windows` nor `Documents`.** It lives alone
  in `ProtectedViewWindows`, in a sandboxed `WINWORD` of its own, and its `Window.Hwnd` is the
  sandbox's handle rather than the frame's.

That last one is the one that changed the code. A lookup by frame handle simply finds nothing for
such a tab — which is also the right answer on its own merits, because a document in Protected View
cannot be edited and so can never be dirty. But it forced a distinction that would otherwise have
been glossed:

> **Word would not answer** and **Word answered, and no window claims this frame** are different
> results and must not share a return value.

`WordTabReadModified` returns FALSE for the first and does not touch the caller's array, so every
flag keeps what it had. It returns TRUE for the second with that entry set to FALSE, because "no
window claims this frame" is a real answer. A tick that could not take a measurement must not be able
to pass for a measurement that came back clean — the same rule that the colour sampler learned the
hard way.

### The probe's own trap, kept rather than deleted

The obvious way to reach a running Word from outside is `New-Object -ComObject Word.Application`, on
the reasoning that Word is a single-instance server. **It is not reliable here.** With a Protected
View document up there are three `WINWORD` processes, and that call bound to one reporting
`Documents.Count = 0` and a `Windows` collection with no `Count` on it — while three documents were
on screen. Section 3a of the probe keeps it, and reports it, so nobody re-derives it.

What replaced it is `OBJID_NATIVEOM` on a frame's `_WwG` document pane
(`WordLayout.NativeOm`), which returns *that window's own* `Word.Window` in *that window's own*
process. No question of which instance answered. The add-in never needs it — Word hands it
`Application` at `OnConnection`, so it is already in the right process — but a test script needs it
to be a second oracle, and it is the only way to ask a sandboxed Protected View document anything.

The first run of the fixed probe still measured the wrong thing, and the reason is worth keeping:
**it deduplicated the Applications it found by the frame's process id.** A Protected View frame is
owned by the *main* `WINWORD`, but the object model behind its pane belongs to the *sandboxed* one —
so keying on the frame's pid threw away the only Application that mattered and reported
`Windows.Count = 0` as though that were Word's answer. Deduplicating by the Application's own COM
identity fixed it, and the comment in the probe says why.

---

## Poll, not event — and the rule this does not break

The standing rule is that transient state must be heard rather than sampled. It was got wrong twice
over Word's save prompt, at real cost, and both wrong versions passed every assertion watching from
outside Word. So it is worth being exact about why polling is the honest mechanism here.

**The rule is about state that appears and disappears inside one tick**, which no sampler catches at
any cadence. "This document has unsaved changes" is not that. It goes true on a keystroke and stays
true until a save. A sampler one tick late is one tick late rather than wrong.

**And there is no event to hear.** Word's `Application` events cover documents opening, closing and
being saved. None of them is "the modified flag changed", and typing raises none of them. Hearing
this properly would mean an `IConnectionPoint` sink for events that still would not answer the
question.

The residue, named rather than hidden: type a character and press Ctrl+Z inside the same 500 ms and
the dot never appears. That is invisible to the eye as well.

---

## What it costs, measured on the rig rather than reasoned about

The first version of the timing log reported only new records, and that turned out to say almost
nothing:

```
strip  dot poll: 1 window(s) in 26508 us (the most it has ever taken)
strip  dot poll: 2 window(s), 20 passes, mean 1597 us, worst 26508 us
strip  dot poll: 3 window(s), 480 passes, mean 442 us, worst 3204 us
strip  dot poll: 3 window(s), 480 passes, mean 296 us, worst 2156 us
```

**The very first pass of a process costs about 26 ms**, against a steady state around **0.3 ms**.
That is Word building its automation machinery, because the janitor is usually the first thing in the
process ever to touch the object model — the smoke test measured 68 µs for the same call, because the
script driving it had already talked to Word. A worst case with no typical case beside it is not a
measurement of what something costs, so the log now reports a summary ten seconds in and every four
minutes after.

26 ms once, during Word's own startup, is worth 0.3 ms twice a second afterwards. If it ever is not,
`TabDot=0` removes the polling entirely rather than merely hiding what it draws.

---

## The shape of it

**One pass over `Application.Windows` for the whole row**, not a lookup per tab. The collection has to
be walked either way, so this is the difference between four `Invoke`s and four `Invoke`s *per
document*. `PollModified` gathers the frames, calls once, and compares each answer against the cached
flag — `ReadTitle`'s pattern exactly, including that it only logs and repaints when something changed.

`GetItemAt` is **the first call in this add-in that passes an argument through `IDispatch::Invoke`**.
`Item` is asked for by name rather than through `DISPID_VALUE`, because a collection's default member
is a convention and a named lookup that fails says so instead of invoking something else.

**The flag lives on `StripState`, not on the stack's `Member`,** and that is not arbitrary:
`StackTabs` answers with the frame as its own single tab when stacking is off or a member has not
joined, so a flag on `Member` would be correct everywhere except under `Stack=0`, where it would
simply never appear.

The poll stands down in two states, both already known without asking Word:

- **a frame is disabled** — which is what a modal dialog does to the window that owns it, and the same
  signal the batch close hears as an event. Word is asking the user something; it is not the moment to
  ask Word something. The whole pass stands down rather than skipping one window, because the object
  model is Word's rather than that window's.
- **a batch close is in flight** — `StackCloseInFlight()`, newly exported, is the test `CloseBatchEnd`
  already opened with. A batch is a run of `WM_CLOSE`s with Word's save prompt between them, and it is
  the one stretch where Word is repeatedly part-way through something we asked for.

---

## Two things the audit caught, and both were mine

Nineteen claims were raised against this code and all nineteen were refuted by a majority. Two of the
lone dissents were right anyway, which is the argument for reading the near-misses rather than
trusting the vote.

**1. An armed press is not a subset of hover.** Press the close button on a modified tab and slide the
pointer off it while still holding: `hotFrame` is cleared by the mouse move, but `pressKind` is only
cleared by `WM_CAPTURECHANGED`. So the button is still armed with `hot` FALSE, and the first version
of the swap dropped the held chip and drew a resting dot — while an unmodified tab under the identical
gesture kept its chip. The press is live either way; moving back onto the button still closes the
document. The condition is `modified && !hot && !downClose`, and **one gesture must not be drawn two
ways depending on whether the document happens to be saved.** The majority's counter-argument was that
"armed is a strict subset of hot", which is simply false, and two independent readers traced why.

Section 4b of the suite drives it: press the dot, carry the pointer onto another tab without letting
go, photograph the button, release away from it. The x stays, nothing closes, and the dot comes back
once nothing is pressed or hovered. A fix an audit found is not banked until something drives it.

**2. A restore that was not a restore.** `TabThemeSample` was found dead in this slice: the assignment
line went missing when the tab-name slice inserted the block below it, so `g_sampleEnabled` had been
permanently TRUE and the startup log had been reporting `theme sample=on` unconditionally — **a log
line naming a setting it could not observe.** The documented escape hatch, "the first thing to try if
the strip is ever the wrong colour on a machine this has not run on", did nothing at all.

The first attempt at putting it back wrote

```c
g_sampleEnabled = WordTabReadFlag(L"TabThemeSample", TRUE);
```

but `git show c6a0e11:src/native/strip.cpp` line 3295 says it was

```c
g_sampleEnabled = g_lookEnabled ? WordTabReadFlag(L"TabThemeSample", TRUE) : FALSE;
```

Dropping the ternary would have made `TabStyle=0` sample where it used to fall back to the registry —
**a change of behaviour wearing the clothes of a repair.** Restored exactly. Whether the sampler
should be tied to the look at all is a real question, and it is not this slice's to answer.

---

## `TabDot`, and what off means

`HKCU\Software\WordTab\TabDot` (DWORD, default 1). Off, a close button is always an **×** and *nothing
reads Word's object model on the janitor at all* — `PollModified` returns before it asks. The switch
takes the mechanism out of the way rather than merely hiding what it draws, and that is the point of
it: this is the first sustained use of the object model in the add-in, and if some Word somewhere
objects to being asked twice a second, this is the setting that stops the asking. The suite asserts
both halves — no `dot |` lines *and* no `dot poll:` lines.

**The dot ships on the flat renderer too**, as a GDI ellipse. `TabStyle=0` exists so a visual
regression can be bisected, and a fallback that quietly drops a feature is a fallback that cannot be
compared against.

---

## How you measure a shape that is only ever pixels

A dot is drawn and never stored, so there is nothing to read — the same problem as the tab name, and
the same answer: the add-in logs the state when it changes, pipe-wrapped, and the suite asserts
against that line **and then photographs the button.** A log line about what was computed is not
evidence of what was drawn.

The photograph does not merely notice that something changed; it names the shape, with two probes that
are opposites:

- **on the axis, near the centre** — a disc has ink there and the **×** has none, because the **×** is
  two *diagonal* strokes and there is nothing beside its centre at the same height
- **on the diagonal, out at the arm's length** — the **×** runs through it and the disc does not,
  because the radius is 3 logical px against arms reaching 4 in each direction, which is 5.7 away

Each probe carries a control taken from the left edge of the *same* close button in the *same*
photograph, so nothing here reads an absolute colour. It works hovered as well as at rest: with the
chip drawn the ground is the chip's colour, and the comparison still holds.

Measured on this rig at 200%: `dpi=192 r=6 arm=8`, and the two states read
`ink on-axis 0/2, diagonal 2/2` for the **×** against `2/2` and `0/2` for the dot.

**And the click still closes.** The suite clicks the dot on a modified document, requires Word's save
prompt to appear — that is a real proof the rectangle still hit-tests as `HIT_CLOSE` — then cancels it
with Escape and requires the document to still be open. Escape on that prompt is Cancel, so nothing is
discarded: a test script may not throw away work, including work it created itself.

---

## Also found

**One floating Office window aborted three suites, and the cascade is the lesson.** A battery run
died in `check-title`'s `Close-Word` on `class=Net UI Tool Window 622x298 ||`. That is not a
question — it is Office's floating-UI host, for galleries, task-pane popouts and notification
toasts — but `Get-WordDialog` matches *any* visible top-level window of Word's that is not an
`OpusApp` and is bigger than 150x60, so the suite refused to close Word and left it running. Then:

- `check-dot` ran next, hit the same window in *its* opening `Close-Word`, and died immediately;
- `soak-stack` ran after that, **opened six documents and measured eleven**, and five of its strip
  assertions failed against a leftover window whose ribbon was a different height.

Three red suites, one cause, and none of them the code under test. `check-title` re-run on its own
passed 35 of 35 including the identical close. The rule already in the notes is the one that
applies — *if a suite ever again sees more windows than it opened, suspect the previous suite's
cleanup before anything else* — and it is now worth stating the other half: **the suite that reports
the failure is not necessarily the suite that caused it.**

What a real prompt is, measured rather than assumed, twice in that same battery: the save prompt is
`NUIDialog` at 920x713 and the Save As question is `NUIDialog` at 920x641. `check-dot`'s
`Get-WordDialog` now excludes `Net UI Tool Window` and says why. **The other nine suites still carry
the unnarrowed version on purpose** — `check-menu`'s prompt detection is the most safety-critical
assertion in the project and it is not this slice's to change. It belongs on the same retrofit list
as `Set-Pointer` and `Invoke-StripClick`.

**`probe-saved.ps1` leaked a windowless `WINWORD` per run.** Opening a Protected View document leaves
behind a `WINWORD` with no window at all — no frame to send `WM_CLOSE` to, and `CloseMainWindow` has
no main window to close. Three accumulated before `install.ps1` refused to run at all with "3 WINWORD
processes running". The probe's cleanup now ends them, and the guard is what makes that legitimate
rather than a violation of *never `Kill` a Word*: **a process with no `OpusApp` frame has no document
open, so there is nothing for Word to offer to recover.** A process *with* a frame is left alone and
reported.

---

## What to preserve

- **A tick that could not measure must not look like a measurement.** The two failure shapes of
  `WordTabReadModified` are the whole of its contract.
- **`Window.Hwnd` is the join, and it is the frame handle** — for every window, not just the active
  one. Do not go back to `ActiveWindow`; it is provably stale during activation.
- **No rectangle moved.** If a future slice gives the dot a rect of its own, `tools\WordLayout.cs`
  must be mirrored in the same edit or every injected click in ten suites misses.
- **The poll asks about every strip, which is a superset of the tabs.** A frame Word made and never
  put a document in is polled and reads clean. That is correct and it is cheaper than coupling
  `PollModified` to `StackTabs`.
- **Read the audit's near-misses.** Two of the four lone dissents in this slice were right and the
  majority was wrong, on a point of fact each time.
- **The suite that reports a failure is not necessarily the suite that caused it.** Three red suites
  in one battery, one floating Office window, none of it the code under test.

---

## Files

- `src\native\connect.cpp` — `GetItemAt` (the first argument-carrying `Invoke`), `WordTabReadModified`
- `src\native\strip.cpp` — `StripState::modified`, `g_dotEnabled`, `SurfDisc`, `FrameModified`,
  `DrawOneTab`'s swap, `DrawStripFlat`'s swap, `PollModified`, the `TabDot` switch, and the restored
  `TabThemeSample` read
- `src\native\stack.cpp` — `StackCloseInFlight`
- `src\native\wordtab.h` — the two new declarations
- `tools\probe-saved.ps1` — the measurement, kept as evidence
- `tools\WordLayout.cs` — `PidOf`, `DocumentPane`, `NativeOm`
- `tools\check-dot.ps1` — 61 checks
- `tools\check-all.ps1` — the runner

---

## What it costs on a machine that is not this one

**Added 2026-08-18, from the user's own report off the work rig.** Everything above was measured
against local files on a development machine, and the number it produced — a two-window pass in about
200us — is not the number this poll has anywhere else.

| documents | where they live | mean | worst |
|---|---|---|---|
| 1 | local disk (this rig) | ~430us | 2.7ms cold |
| 3-5 | SharePoint (their rig) | **96,000-155,000us** | **518,726us** |

Half a second on Word's UI thread, twice a second. WordTab had never met a cloud-backed document, and
`Document.Saved` on one is three orders of magnitude slower to answer than on a local file.

There is nothing to fix in the *question*. The add-in asks Word for `Document.Saved`; how long Word
takes to answer is Word's business, and the alternative — an `IConnectionPoint` sink — covers
documents opening, closing and being saved but has no event for "the modified flag changed", which is
the only thing this needs. So what changes is **how often**, and **how much is asked for**.

### The governor

`RungFor` turns a pass's cost into a rung on a ladder, and the rungs are chosen so the poll holds to
about 1% of the UI thread whatever the cost turns out to be:

| a pass costs | rung | ask again in |
|---|---|---|
| under 4ms | 0 | 500ms — twice a second, exactly as before |
| 4-15ms | 1 | 2s |
| 15-50ms | 2 | 6s |
| 50-200ms | 3 | 20s |
| over 200ms | 4 | 60s |

**On this rig nothing changes at all.** A full pass is a few hundred microseconds, `RungFor` returns
rung 0, and the cadence is the one it has always had.

> **2026-08-19: the first cut of this shipped in `f333349` and did not work on the work rig.** The
> ladder was right; the way it was walked was not. See *The governor that could not settle* below,
> which is the version that is actually in the code.

### ...and in between, only the window with the keyboard

A document is edited in the window that has the focus and saved from the window that has the focus.
So between whole-row passes the poll asks about the foreground window alone — one `Document.Saved`
instead of five — on its own governed interval, and skips entirely when the foreground window belongs
to another application, where nothing of ours can be changing.

The periodic full pass is what covers the rest: a background document Word saves by itself, which is
what AutoSave on a SharePoint document does. Being a few seconds late on one of those is a dot that
appears late, not a dot that is wrong.

### The governor that could not settle

**`f333349` set the interval from the cheaper of the last two passes, and documented the fast
recovery as deliberate:** *"being too eager costs a few milliseconds, being too slow costs a dot that
is a minute late."* The work rig's next report showed that on a SharePoint machine the first half of
that sentence is false.

There, a pass costs **either ~300-600us with the answer warm or 90,000-518,000us with it cold, and
nothing in between**. So the cheaper of two was nearly always the warm one, and one lucky sample
undid every backoff:

```
10:20:18  took 268 us    (and 494927 us before)  -> now asking every 500 ms
10:20:27  took 95044 us  (and 486406 us before)  -> now asking every 20000 ms
10:20:46  took 344 us    (and 95044 us before)   -> now asking every 500 ms
10:20:49  took 92637 us  (and 501877 us before)  -> now asking every 20000 ms
```

That cycle ran about every 22 seconds for hours, at **mean 135,302us over 480 passes in twenty
minutes — about 5.5% of Word's UI thread against the 1% the ladder was drawn for**, and spent as
half-second freezes of the thread Word draws with rather than as a smear. It bought two dot
transitions in a 25-hour log.

**The reason it could not settle is not a tuning problem.** The governor's *input* is not independent
of its *output*: asking often keeps SharePoint's answer warm and therefore cheap, and backing off
lets it go cold and therefore expensive. A governor reading a quantity its own cadence controls has
two stable states and flips between them — which is exactly what the log shows, thrashing at 500ms or
pinned at 60s for two hours and twenty minutes with every single pass slow. Nothing in between was
reachable.

### So it climbs in one step and comes down one rung at a time

For the **whole-row** pass:

- the basis is the **dearer** of the last two passes, so one slow pass backs off at once
- recovery walks the ladder a rung per pass. That is five cheap passes from 60s back to 500ms, not
  four: because the basis is the dearer of two, the first cheap pass only clears the slow sample out
  and moves nothing. **A single lucky warm answer changes the cadence not at all.**
- the first pass of a process is now discarded **explicitly**. It was free under `min` —
  `min(anything, 0)` is 0 — but a 27,547us cold start is Word building its automation machinery, and
  under `max` it would put a healthy machine on a six-second cadence and cost four more passes to
  walk back off it.

For the **active-window** pass, none of that applies and it is governed on its own cost alone, with
nothing carried between passes. The two-sample rule is only meaningful when consecutive passes
measure the same thing; the row is the same set every time and the foreground window is not, so
"the dearer of the last two" would let one slow document set the cadence for every tab the user
afterwards moves to. It does not need the stickiness either, because the feedback that made the row
bistable is absent: Word keeps the document being looked at live. On the work rig the active pass
**never once exceeded 4ms across five and a half hours** while its row passes were running to half a
second. If that ever stops being true the log says so in its own words, and that line is the one to
look for.

### Proving it, when the rig cannot reproduce it

The dev rig cannot produce a slow pass, so `check-dot.ps1` exercises none of the above — which is
precisely how the first version got out. Both state machines were therefore replayed side by side
over the cost sequence in the work rig's own log:

```
old returned to 500ms 11 times in 15 passes; new 0 after pass 2
```

with the cold start discarded, one slow pass reaching 60s immediately, and recovery measured at five
cheap passes. The check that matters most on *this* rig is the opposite one: a 26,621us cold start
was discarded and **no cadence-change line fired for the whole run**.

The cadence is logged when it **changes** and never per pass, with both measurements on the line:

```
strip  dot poll: the whole row took 155000 us for 5 window(s) (and 149000 us the time before)
       - now asking every 20000 ms
```

`TabDot=0` still turns the whole thing off, and that remains the mitigation for a machine where even
this is too much.
