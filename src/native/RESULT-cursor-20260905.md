# The IDC_SIZEALL regression: there was never a second ingredient — 2026-09-05, late

Input: `RESULT-battery-20260905.md` section 6 item 1 — *"The guard is necessary but not sufficient; the
second ingredient is unidentified and there is no committed snapshot of the passing tree. It blocks the
clean battery that blocks everything else."*

**It is fixed.** `check-reorder` is 118/118, and the fix is proved to be what makes it pass by building
a variant with the fix neutralised and watching the check go red.

`HEAD` is still `ef49c79`. Nothing is committed and nothing is published.

---

## 0. The snapshot that "did not exist"

The previous note's central difficulty was that nothing is committed, so the passing tree could not be
diffed against the failing one. **That was wrong, and the way it was wrong is worth keeping.**

Claude Code's own session transcripts record every `Edit` with its `old_string` and `new_string` and a
timestamp. `~/.claude/projects/D--WordTab/*.jsonl`. Session `6c024d7d` spans 09-04 16:40 → 09-05 17:13
local (transcripts are UTC, this machine is UTC−4) and therefore contains **both** the 09-04 build that
passed and every edit made after it.

Reading it out gave the answer in one step:

| | |
|---|---|
| the binary that PASSED at 09-04 18:40 | built at **09-04 17:44:21** |
| `strip.cpp` edits after that build | **exactly three** — 09-05 16:59:06, 16:59:10, 17:04:37 |
| what those three edits are | the `SampleChrome` BitBlt rewrite, and nothing else |
| `connect.cpp`, `stack.cpp`, `frames.cpp`, `wordtab.h` | last written 09-04 **≤17:14**, i.e. *before* that build |

So the four other files are byte-identical across the regression — which independently confirms the
previous session's bisect, and narrows `strip.cpp` to one change without building anything.

**Keep this.** An uncommitted tree is not an unrecoverable one while the transcripts are on disk.

---

## 1. The bisect, and what it actually proved

`tools\ab-binary.ps1 -Patch sampler` puts `0 &&` in front of the `BitBlt`, so `banded` stays FALSE and
all 24 points fall through to `GetPixel(screen, ...)` — the same points, the same `y`, the same order.
That *is* the old sampler, 750,000us and all.

```
variant AE17CB4AA2955D2B   deployed-was B780E41FF32B02F9
cursor during the TabGhost=0 tear-off 0x10015
PASS  118 checks, 0 failures
```

The prediction in the last note held. **But the conclusion it invited is wrong.** Reverting the sampler
would re-hide a real fault *and* restore a 744ms stall on Word's UI thread that the same session had
just correctly removed.

---

## 2. There was never a second ingredient

Every row of the evidence table — including two the last note could not place — is explained by one
question: **how busy is Word's UI thread during the 400ms of mouse silence the check measures in?**

| variant | the thread during that 400ms | result |
|---|---|---|
| `ef49c79` | 750ms sampler hogging it | PASS |
| the 09-04 17:44 build | 750ms sampler hogging it | PASS |
| `-Patch sampler` | 750ms sampler restored | PASS |
| `-Patch trybind` | frame binds, so each tick does MORE work | PASS (and 17 other checks red) |
| working tree | quiet | FAIL |
| `TabThemeSample=0` | quieter still | FAIL |
| `TabDot=0` | quieter still | FAIL |
| `-Patch dragstand2` (clean anchor) | quieter still | FAIL |

`-Patch trybind` is the row the last note read backwards. Neutralising the guard does not remove a
fault — it makes the frame bind, which makes every janitor tick **heavier**, which masks the fault the
same way the slow sampler did. That is also why it cost 17 other checks: it was never a fix.

**So the defect is single, and it is old.** It predates `ef49c79`. The 744ms chrome sampler was hiding
it in every build there has ever been, and fixing that sampler is what uncovered it.

`dragstand2` — the hypothesis the last note marked live — is now **refuted on a clean single-site
anchor**, which is what that note asked for. It fails.

---

## 3. Naming it instead of inferring it

Four hypotheses fitted the evidence equally well and none named a call. Three probes, each answering
the question the one before it raised:

1. **Per-stage probe in `JanitorProc`** — read the cursor around every stage while `g_dragging`.
   **Zero hits.** The cursor does not change inside any janitor stage.
2. **Baseline moved into `DragCursor`** so a change *between* ticks became visible:
   ```
   CURSOR TAKEN mid-drag: tick start (so it went BETWEEN ticks)
   put 0x10005 where 0x10015 was   [tick 9, 8 TryBind call(s) so far]
   ```
   Between ticks, and `g_dragging` was TRUE — which by itself refutes every "the drag ended early"
   story, because `DragForget` clears `g_dragging` and the probe is gated on it.
3. **`WH_CALLWNDPROCRET` hook** naming the message that changed it:

```
20:38:09.049  the tear - the strip sets IDC_SIZEALL                  cursor 0x10015
20:38:09.439  _WwG  WM_NCCREATE           (0x0081)                   cursor 0x10005
20:38:09.439  _WwG  WM_WINDOWPOSCHANGING  (0x0046)                   cursor 0x10007
20:38:09.441  _WwG  WM_STYLECHANGING      (0x007C)                   cursor 0x10005
20:38:09.871  the button is released
```

**Word builds a `_WwG` document view 390ms into a gesture it knows nothing about**, puts up a wait
cursor while it works, and then restores what *it* thinks the pointer should be over a document — the
I-beam.

Nothing of ours is asked. Nothing is lost. The capture never moves — which is why the
`WM_CAPTURECHANGED` handler is silent in the failing, the passing, *and* the `ef49c79` baseline logs.

---

## 4. The mechanism, and the comment that hid it for months

`strip.cpp`'s `DragCursor` used to carry this:

> *"Setting it while we hold the mouse capture is enough to keep it — WM_SETCURSOR is not sent to
> anybody while the mouse is captured, so nothing else is going to put the arrow back underneath us."*

The first half is true. **The conclusion does not follow.** `WM_SETCURSOR` is not sent while the mouse
is captured, so no other *window* is asked — but `SetCursor` is per-**thread** state and this add-in
runs on Word's own UI thread. Word does not have to be asked. It just runs, and the last `SetCursor` on
the thread wins.

The strip only ever asserted the cursor from `WM_MOUSEMOVE`. So the moment the hand stops, the gesture
stops defending its own pointer.

**A note on the classes:** the strip's own window class cursor is `IDC_ARROW`, so the I-beam could not
have arrived through `WM_SETCURSOR` reaching the strip either. `0x10005` was only ever explicable as a
direct `SetCursor` from Word's code on our thread.

## The fix

The cursor is **held** for the length of the gesture instead of set at its edges. `DragCursor` records
what it asked for; a 50ms timer (`ID_DRAG_CURSOR`) puts it back if anything else has taken it; it starts
where a gesture becomes a drag and stops everywhere a gesture can end, beside `DragScrollStop`.

It is not silent. Once per gesture it says what took the pointer, named against the cursors it could be:

```
strip  the pointer was taken mid-drag and put back: something on Word's UI thread
       set 0x10005 where this gesture had set 0x10015  (SIZEALL 0x10015, arrow 0x10003, I-beam 0x10005, NO 0x10017)
```

**That line fires in the ghost-ON sections too** — one run logged Word replacing `IDC_NO` with a Word
custom cursor `0x26F0227` during an ordinary in-row drag. So this was never about `TabGhost`; the
carried card was never hiding anything. The `TabGhost=0` step is simply **the only cursor read in the
whole suite taken during mouse silence**, and that accident is the only reason this was ever caught.

---

## 5. A check that looked right and pinned nothing — removed

Since the fault is not `TabGhost`-specific, the obvious move was a symmetric read in the ghost-ON
section: sleep 400ms after the tear, assert `IDC_SIZEALL`. It was written, and it passed.

Then it was tested the only way a new check may be trusted — `-Patch cursorhold`, which neutralises the
hold:

```
cursor after 400ms of stillness 0x10015
PASS  and it is still IDC_SIZEALL after the hand has been still for 400ms     <- WITH THE FIX REMOVED
cursor during the TabGhost=0 tear-off 0x10005
FAIL  and the pointer still says so - IDC_SIZEALL, unchanged
FAIL  1 of 119 checks failed
```

**It passes with the fix removed, so it asserts nothing.** The theft is opportunistic — it needs Word to
happen to build a window in that particular silence, and in the ghost-ON section it does not. The check
was removed rather than kept as decoration.

`check-reorder.ps1` is therefore back to the previous session's +11 lines, and the pin for this fix is
the `TabGhost=0` step, which **is** proved to go red without it. `-Patch cursorhold` is kept in
`ab-binary.ps1` as that proof.

**Do not re-add the symmetric check without re-running `-Patch cursorhold` against it.** See
[[green-suites-are-not-a-clean-ab]].

---

## 6. The battery

The clean battery `NOTE-next-20260905.md` asked for, and `RESULT-battery-20260905.md` could not get.
Console minimised; deployed DLL hash-checked against the built one before it started (`D2604A68…`, and
the wrapper throws if they differ, because a battery about the wrong binary proves nothing).

| | |
|---|---|
| governor 21 | palette 36 |
| stack 65 | **row 36, 8 failures** |
| rowsize 33 | onepage 18 |
| strip 67 | tabs 26 |
| menu 71 | **reorder 118, 0 failures** |
| startscreen 50 | look 48 |
| scroll 62 | title 105 |
| dot 71 | soak-stack 22 |

**`row` is a confound and the harness named it itself**, which is what the extra instrumentation in
these suites is for:

```
attempt 1: (1040,773) is over 0x806B2 (Windows.UI.Core.CoreWindow), not the button 0x320438
attempt 2: (1040,773) is over 0x806B2 (Windows.UI.Core.CoreWindow), not the button 0x320438
attempt 3: (1040,773) is over 0x806B2 (Windows.UI.Core.CoreWindow), not the button 0x320438
FAIL  the Cancel button was clicked
```

A Windows shell overlay sat over the dialog's Cancel button for all three attempts; the seven failures
after it are that one cascading. Zero `taking it back` lines, so it was not the console this time — it
is the same family as [[the-console-is-a-foreground-thief]] with a different thief. `row` re-ran
**42/42**, which is the count the `ef49c79` baseline and the previous battery both got, so nothing was
lost to the cascade.

**So: 16/16 suites, 855 checks**, counting `row` at its re-run.

The interesting number is `reorder`: **118/118**, where four consecutive runs on this tree had failed
it.

---

## 7. Still open

1. Nothing is committed and nothing is published — and publishing was refused once before. **Ask.**
2. The probes from `NOTE-next-20260905.md` (`ProbeUnaskedResize`, `StripDescribeDpiContext`) are still
   printed and still unread — they need a field report, not a suite.
3. `Set-Pointer` still confirms the wrong event (`RESULT-battery-20260905.md` section 6 item 2).
4. `scroll`'s out-of-step race and `check-stack.ps1` throwing on lost foreground are both unchanged.
