# The row: what order tabs are in, what keeps one, and what closing the window means

**PASSED.** 25 checks in the new `tools\check-row.ps1`; the other eleven suites re-run —
**719 checks, 0 failures across the twelve in one clean battery.** `check-row` has grown by four
since that run — the unsaved-work section below — and passes 25 of 25 standalone; nothing in the
add-in changed to add them, so the eleven other numbers are the ones the battery produced.

| suite | checks | | suite | checks |
|---|---|---|---|---|
| stack | 60 | | startscreen | 50 |
| **row** | **25** | | look | 48 |
| strip | 67 | | scroll | 62 |
| tabs | 26 | | title | 105 |
| menu | 71 | | dot | 69 |
| reorder | 118 | | soak | 22 |

The queue the user left after their first full afternoon of real use, on the work machine, done in
one pass. Five items, four of them in this file and one — the dot poll's cost — in
[RESULT-dot.md](RESULT-dot.md#what-it-costs-on-a-machine-that-is-not-this-one).

| what they said | what it was |
|---|---|
| "changing setting like view>page width closes tabs" | **not reproduced here.** Guarded, and made to explain itself next time |
| "now newly opened docs popping infront lets make them pop to end" | a frame Word had already made held an early slot in the row |
| "you have to close all tabs individually - that needs fixing" | nothing was listening for the window's own close |
| the document area lays out wrong on the work rig | still open — waiting on their `LoadBehavior 0` A/B |
| opening an existing document flashes above the stack | the placement happened after Word had shown the window |

New suite: `tools\check-row.ps1`. New probe: `tools\probe-view.ps1`, which found nothing, and saying
so is the point of it.

---

## 1. The one that could not be reproduced, and what was built instead

**The report.** A View change costs the user the row. Asked whether the document went with it:
*"oh no it didnt it just opened a new word w/out it"* — so the window is **evicted** from the stack
and comes back as a plain Word window with no strip on it. The document is never at risk. That is
the worst thing in the queue: a routine Word UI action, and the row is gone.

**What was tried.** `tools\probe-view.ps1` opens three documents and drives every View command Word
exposes, twice over — once by setting the property (`View.Type`, `Zoom.PageFit`, `Zoom.Percentage`)
and once through `CommandBars.ExecuteMso`, which raises the ribbon button's own command rather than
imitating it: `ZoomPageWidth`, `Zoom100`, `ViewWebLayoutView`, `ViewPrintLayoutView`, `ViewDraftView`,
`WindowSplit`. After each one it photographs, per window, every `_WwF` the frame has, whether each is
visible, what is inside each, and whether the strip is there.

**Nothing moved.** One document frame per window throughout, all visible, all holding a `_WwB`, the
strip in place, and the add-in logged not one line about the stack across the whole run. The view
really did change each time — the probe reads `View.Type` and the zoom back afterwards, so a command
Word ignored cannot pass for a defect that is not there.

So the trigger is on their machine and not on mine, and the two candidates are the ones already
known: **Office Tab's helper is loading over there**, and it carves up the same `_WwF` — WordTab has
never once run beside a working copy of it, because Office Tab has been in Word's
`Resiliency\DisabledItems` on this rig since 2026-08-15 and will not come back. And their documents
are SharePoint-backed `.docx` where mine are local `.rtf`.

The second suspect got its own line in the log this slice, and it is cheap: **whose child `_WwF`
is**. Every rectangle in `strip.cpp` is computed in the frame's client space and handed to
`SetWindowPos`, which reads it as the *parent's* — and those are the same space only while the
document frame hangs directly off the frame. If Office Tab has taken `_WwF` in under a container of
its own, every number the strip computes is off by that container's origin, which is also a candidate
for the still-open "the document area lays out wrong" item. It is not fixed, because it cannot be
measured here. It now says so on sight.

### The mechanism that fits, guarded rather than fixed

Not a guess about Word so much as an audit of what the add-in would do if Word did the obvious thing.
`TryBind` bound to the first `_WwF` it found and only ever rebound when that handle was **destroyed**.
So if Word retires a document frame by *hiding* it and building a second one beside it:

- the old handle is still a window, so nothing ever rebinds
- the old frame is empty, so `StripHasDocument` answers "no document" about a window that has one
- membership is decided from that answer, so the window is dropped from the row — **permanently**,
  because nothing re-tests the binding
- and the strip is hidden with it, because the janitor shows the strip when `_WwF` is visible

Which is the reported symptom exactly, down to the missing strip. Three changes, each narrow enough
to be free when Word is behaving:

**`PickWwf` chooses the document frame rather than taking the first one.** Score: something inside
beats nothing inside, and among equals a visible frame beats a hidden one. "Something inside"
outranks "visible" because Backstage hides a perfectly good document frame.

**`StripHasDocument` asks the window, not the handle.** The cheap test — one `GetWindow` on the bound
frame — is unchanged and answers nearly every call. Only when it is about to say *no document* does it
enumerate, which is precisely where a wrong answer is expensive.

**The janitor rebinds to a replacement.** Only a binding that has gone empty or invisible is
questioned at all, and it is only given up for a candidate that is both non-empty and visible. On a
window behaving normally this test never fires.

### And a tab is no longer given up on one reading

The fourth time this project has met a state that is true for a moment being read as a state that is
true ([[transient-state-must-be-an-event]]). There is no event to hear here, so the next best thing
is to refuse to act on a single look: a window already in the row has to fail the test to stay in it
for **700ms** before its tab goes. Not for a *count of ticks* — see below, that version was wrong.
Two exceptions, both immediate: a window that has actually been destroyed, and a close **we** posted
the `WM_CLOSE` for, because that one the user is watching for.

### The diagnostic, which is half the work

The log that reported this said a window left the stack and nothing whatsoever about the state that
decided it. A window with two document frames and a window with one empty one produce the same line,
and they are different bugs. So the first time a window looks like it should leave, before anything
is done about it, the add-in now writes:

```
stack  hwnd=0x...  would have left the stack (no document open) - waiting 700 ms in case it is a
moment rather than a fact.  window visible=1:  _WwF=0x...(bound) visible=0 holds=empty  _WwF=0x... visible=1 holds=_WwB
```

Two `_WwF` on that line names the mechanism above. One empty one names something else. Either way the
next report answers the question instead of raising it.

---

## 2. A new document's tab goes to the end

**The spec, in their own correction:** *"i dont mind it popping the doc i just opened but i want ti at
end of tabs. i like to keep same 1 or 2 in fromt."* Activation and position are separate and only
position was wrong — the new document **should** take the focus. Browser behaviour: append right,
take focus, leave the parked tabs on the left alone. Their reason is muscle memory, which makes a
stable left-hand end a requirement rather than a taste.

The array **is** the row, and joining appends, so a new document's tab was already at the end — for a
frame Word creates while the user watches. It is not at the end for a frame Word made **earlier** and
is only now filling: their log shows empty frames attaching with no title at all, and a frame that
attached before the documents around it holds an early slot and hands it to whatever lands in it.

`rejoin` already existed for the recycled-frame case, and it says exactly the right thing: *whatever
document turns up here next has never had a tab, so its tab goes at the end*. It was set only for a
frame observed empty **while out of the row** — which misses a frame that was filled between two
janitor ticks. So it is now set at `StackAttachFrame` as well: a window we have only just met has
never had a tab either.

Order-preserving when several frames arrive together, which is every cold start with more than one
document — they join in array order and each moves to the end in turn, which leaves them in the order
they were in. Driven rather than reasoned: `check-row.ps1` opens three and checks all three.

### The row is now in the log

Two of the five items are about **order**, and neither was answerable from the log that reported them:
there were lines for a window joining and a window leaving, and none at all for what the row then
*was*. A membership line that does not say what the membership is cannot settle an argument about
ordering.

```
stack  row, left to right: 0x...0840 |report.docx|  0x...0576 |notes.docx|  0x...08E0 |minutes.docx|
```

Written when the row changes and never otherwise — it builds the line and compares it against the last
one it wrote, so a row sitting still costs a `_snwprintf` twice a second and no file write at all.
This is also the instrument `check-row.ps1` reads: tab order cannot be read from outside the process,
because the tabs are painted rather than controls.

---

## 3. Closing the window closes the stack

*"you have to close all tabs individually - that needs fixing next time we edit."* Confirmed in their
log at 14:27-14:28, four frames going down one at a time.

The stack is one window to the user in every other respect — one taskbar button, one Alt+Tab entry, it
moves and resizes as one — so the button that closes a window has to close the window they can see.

**`SC_CLOSE`, not `WM_CLOSE`, and the difference is the whole of the safety.** `WM_CLOSE` is what the
add-in itself posts to close one tab, and what the batch posts to close them in turn; intercepting
that would be the stack answering its own question, forever. `SC_CLOSE` is only ever the user: the
title bar's ×, Alt+F4, the window menu.

Nothing new had to be built underneath it. `StackCloseAll` and its queue already existed for the
right-click menu, and the queue is what makes this safe: one close at a time, each with its own save
prompt, and a prompt the user cancels abandons the rest. Closing a stack of five documents with
unsaved work asks five questions in turn, and Cancel on the second leaves the remaining three open.

Everything that is not unambiguously *the user pressed close on a stack of several documents* answers
FALSE and lets Word do exactly what it did before: not stacked, a row of one, a batch already running,
or `TabCloseStack=0`. That switch is there because this is the one thing WordTab does that ends with
several of the user's documents closed.

**And that is why the new route is checked against unsaved work rather than only against closing.**
The question that matters about this command is not whether it closes them, it is whether it can
close one somebody had not finished with. Driven: dirty a document, press the window's close, and
Word's own prompt goes up **before anything has closed**; Cancel, and all four documents are still
open, with the add-in's log saying it stopped *because the user declined* rather than because it ran
out of patience. That last distinction is not pedantry — an earlier version of the batch concluded
"declined" one second after posting `WM_CLOSE`, before Word had even asked, and passed every
assertion that did not look for the reason.

---

## 4. A window is born on the stack

The queue's last item, which the user called minor: opening an existing document shows its window a
few inches above the stack before it settles in.

The cause is not minor. `CbtProc` catches `HCBT_CREATEWND` and **posts** — so by the time the
coordinator picks the window up, Word has created it where it chose and shown it there, and the stack
can only move it afterwards. "Afterwards" is a frame the user can see.

`lpcs` is the `CREATESTRUCT` Word passed to `CreateWindowEx`, it is writable, and writing it is the
documented purpose of `HCBT_CREATEWND`. So the stack's rectangle is written into it inside the hook,
before the window exists. There is no first position to flash from.

Checked on the `attach` line's rect, which is read after the window exists and before anything has
moved it. A one-frame flash cannot be photographed reliably; a window created at the right
coordinates cannot produce one.

---

## What running it found that reading it did not

**The debounce was counted in ticks, and two ticks happened in a quarter of a second.** The janitor is
not only the timer — `WM_SHOWWINDOW` calls it too — so hiding a window produces a pass from the hide
and another from the show, milliseconds apart, and a guard that wanted "two ticks" was satisfied by
both edges of one gesture. The first run of `check-row.ps1` evicted the window it was proving would
not be evicted. It is a **length of time** now, longer than the janitor's own period, so at least one
later tick has to agree.

**And `WM_SHOWWINDOW` is sent *before* the visibility changes**, so a janitor run from that handler
reads the state the window is leaving. On a show that means reading "hidden" about a window that is
being shown. It now chains first and asks afterwards.

### Which broke eight checks in a suite that had never gone red, and the A/B is what found it

`check-reorder` went from **118/118 to 110/118** — the tear-off drags, consistently, three runs
running. Not a flake, and nothing in the diff obviously touched dragging. The rule in this project is
to A/B against a rebuilt binary of the previous commit before believing anything, so: stash, rebuild,
install, re-run. **118/118.** It was mine.

**A membership pass was running inside another membership pass.** The janitor is not only the timer —
`WM_SHOWWINDOW` calls it, and `Join` and `Leave` both show and hide windows, so a pass can call
synchronously back into itself several frames deep. The array **is** the tab order and `MoveToEnd`
memmoves it, so a re-entrant `Join` leaves the outer pass holding a `Member*` into a slot that now
describes a different window. `StackTearOffTab` reads `member->tornOff` through such a pointer.

The hazard was always there. What reached it was moving the `WM_SHOWWINDOW` janitor call to *after*
Word had shown the window — because that is the first time the nested pass can answer "yes, join
this one", which is the only branch that moves the array. Both halves are right on their own and
wrong together.

So `StackJanitor` is now a guard around `JanitorPass`, and a nested call returns immediately. Nothing
is lost by skipping it: the timer comes round in half a second, and the outer pass is in the middle
of deciding the very thing the nested one would have decided.

**The generalisation: an event handler that calls a re-derivation, and a re-derivation that raises
events, is a loop — and the array it walks must not be mutable underneath it.** This is the same
family as reading `member->frame` after `MoveToEnd` two paragraphs up. Both are the array moving
under a pointer; one from below, one from the side.

**A log line named the wrong window.** `MoveToEnd` memmoves the array, so the `Member*` afterwards
points at whatever slid down into that slot — and the line reporting the move read `member->frame`
after the move. It had been naming an empty spare frame instead of the document that had just
arrived. Read before the move now. A log line that names the wrong window is worse than no line.

**And the suite counted its own guard's message as the failure.** `would have left the stack` contains
`left the stack`. Two spaces in the pattern separate them.

**The row line put document NAMES in the log, and a suite matched one.** `check-reorder` asserted
that nothing was written about the carried card with `Get-LogSince 'ghost'`, and the fixtures in that
very section are called `wordtab-noghost-Alpha.rtf`. The row line names every tab, so a green suite
went red on the name of its own scratch file. It is `ghost  hwnd=` now. **Anything that writes user
text into the log widens what every pattern in every suite can match** — that is the cost of the row
line, and it is worth it, but the next such line should come with the same sweep.

---

## Files

- `src\native\stack.cpp` — `Member::missedAt` and `STAY_GRACE_MS`, `g_closeAimed` /
  `CloseWasAskedFor`, `LogRow`, `rejoin` at `StackAttachFrame`, the `MoveToEnd` log fix,
  `StackCloseWindowCommand` and `TabCloseStack`, `StackProposeCreateRect`, and
  `StackJanitor`'s re-entrancy guard around `JanitorPass`
- `src\native\strip.cpp` — `PickWwf` and `WwfHuntProc` (replacing `FindChildOfClass`),
  `StripHasDocument`'s fallback, `Unbind`, the janitor's rebind, `StripDescribeDocumentFrames`,
  the parent-of-`_WwF` warning, and the dot poll's `TicksFor` / `AskModified` governor
- `src\native\frames.cpp` — `SC_CLOSE`, the `CREATESTRUCT` write in `CbtProc`, and
  `WM_SHOWWINDOW` chaining before it asks
- `src\native\wordtab.h` — four new declarations
- `install\settings.ps1` — `TabCloseStack`
- `tools\check-row.ps1` — 25 checks, including the window close over a document with unsaved changes
- `tools\probe-view.ps1` — the measurement that found nothing, kept as evidence
- `tools\WordLayout.cs` — `SysClose`
- `tools\check-reorder.ps1` — two patterns that were matching the wrong thing
- `tools\check-all.ps1` — the runner, now twelve suites
