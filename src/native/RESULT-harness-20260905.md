# Two suite fixes: a confirmation that confirmed nothing, and a throw that cost a battery — 2026-09-05

Both were on the open list in `RESULT-battery-20260905.md` section 6 (items 2 and 4). Neither is the
add-in. Both are the reason batteries have been lying.

---

## 1. `Set-Pointer` returned true for moves nothing received

`WordTabHarness.ps1` confirmed a hover by asking `GetCursorPos` whether the pointer was where it had
been sent. That is a true statement about the **pointer** and says nothing about whether the window
under it was told — and the strip's hover state is driven by `WM_MOUSEMOVE`, not by where the cursor
is. The 2026-09-05 battery has it in its own words, two lines apart:

```
PASS  the pointer arrived on Beta
FAIL  hovering a modified tab brings the x back
```

with eight seconds of tooltip silence in the add-in log across the whole step, where the passing run
logged four show/hide pairs.

### Measured, not argued: `tools\move-delivery.cpp`

One window, one counter in its own window procedure, no Word and no add-in:

```
  1. arriving from wherever the pointer was      ->  1 WM_MOUSEMOVE   GetCursorPos says (300,300)
  2. a REAL move, B -> A                         ->  1 WM_MOUSEMOVE   GetCursorPos says (400,380)
  3. the SAME point again, A -> A                ->  0 WM_MOUSEMOVE   GetCursorPos says (400,380)
  4. and once more, A -> A                       ->  0 WM_MOUSEMOVE   GetCursorPos says (400,380)
  5. nudge away, A -> A+1                        ->  1 WM_MOUSEMOVE   GetCursorPos says (401,380)
  6. and back, A+1 -> A                          ->  1 WM_MOUSEMOVE   GetCursorPos says (400,380)
```

**Windows posts no move message when the position does not change, and `GetCursorPos` answers happily
in exactly that case.** So the old confirmation did not merely fail to prove delivery — it reported
SUCCESS for a move nothing received.

**And it made the retry loop decorative.** After attempt 1 the pointer is already at the target, so
each of the remaining four retries re-sent the same absolute move: zero messages, four times, then
`$true`. A loop that could not have worked is worse than no loop, because the caller reads it as five
tries.

### What it does now

- **Nudges off the target and back whenever the pointer is already on it**, so there is always a real
  transition for Windows to report. That covers the retries as well as the first call.
- **Checks what is UNDER the point**, when the caller says what it should be: `-Onto <class>`, checked
  with `Get-ClassAt` — the same authority `Invoke-ConfirmedClick` already uses for clicks. A move
  delivered to a window that is *covering* the strip and a move the strip ignored are otherwise the
  same evidence. The class it actually found is named in the note, because "not on the strip" narrows
  nothing down.
- `-Onto 'WordTabStrip'` is wired into the six call sites that aim AT the strip (`check-dot`,
  `check-look` ×2, `check-reorder`, `check-scroll`, `check-tabs`). The parks at (4,4) deliberately aim
  off it and pass nothing.

**What it still does not prove** is that the strip *acted* on the move; nothing outside the process
can see that, and the caller's own assertion is what tests it. What it does mean is that a failure
after a `$true` from here is about the add-in — which is the whole point.

### The pin: `tools\pin-pointer.ps1`

A harness fix cannot be A/B'd with `ab-binary.ps1`, which swaps the add-in. So the pin is a window that
counts what it was actually told, driven by the real `Set-Pointer`, with `-Old` redefining the pre-fix
body on top of it as the control arm:

| arm | WM_MOUSEMOVE delivered | what Set-Pointer returned |
|---|---|---|
| `-Old` (pre-fix) | **1** | True, True |
| the fix | **3** | True, True |

The control returns `$true` for a move nothing received. That is the defect, reproduced on demand.

**It refuses to report a number taken while somebody is using the mouse.** The first run of the pin had
the control arm reporting 2 and passing, with the pointer drifting through `(1371,557)` and
`(1358,543)` between calls — a hand on the mouse counts the same as an injected move. The watcher now
measures 600ms of nothing first and prints `QUIET n`; the pin skips with a yellow line if `n` is not
zero. An instrument that cannot say "I could not measure this" is worse than no instrument.

---

## 2. `check-stack` threw instead of failing, and took a battery with it

On 2026-09-04 another application took the foreground mid-suite:

```
==> Minimising the stack
check-stack.ps1: Index was outside the bounds of the array.
```

`Get-Frames` came back empty and `$before[0]` under `Set-StrictMode` is a hard error, not `$null`.
Three Word windows were left up, `check-row` started against them, and reported `the row has four tabs
(has 7)` eight ways. **One stolen foreground cost a whole battery, and most of a session's diagnosis,
because the suite that failed was not the suite that broke.**

Two fixes, and the second is the one that generalises:

- **The specific throw is now a named failure.** The step is skipped with
  `there are windows to minimise (Get-Frames found none - something closed them, or took the
  foreground away from Word)`. The two asserts that used to follow the restore are now inside the
  `else` as well: left outside they compare 0 against 0 and **PASS**, so a step that measured nothing
  would report that everything came back.
- **Closing Word is a `finally`, not a last line.** The cleanup was the final statement of the script,
  so any unhandled error skipped it. It is now `Invoke-StackCleanup`, called both on the normal path
  and from a script-wide `trap` that names the error, names its line, cleans up, and exits non-zero —
  so `check-all` records *this* suite as red instead of recording the next one as red.

A `trap` rather than a `try`/`finally` wrapped round eight hundred lines: it reaches every statement
without re-indenting any of them, which matters because the diff is what the next reader checks.

### The pin: `-PinThrow`

```
pwsh -File tools\check-stack.ps1 -PinThrow
```

throws where the crash threw. Observed:

```
check-stack stopped on an unhandled error: -PinThrow: pretending the foreground was stolen here
  at line 611: if ($PinThrow) { throw '-PinThrow: pretending the foreground was stolen here' }
==> Closing Word
FAIL  stopped after 43 checks, 2 failure(s), plus the error above
---- exit code: 2
WINWORD after : 0
```

Word closed, exit code non-zero, the error and its line named. `check-all` passes no such switch, so it
is inert in a battery.

---

## What these do NOT fix

**The `dot` hover failure is not proved to be either of these.** The move that failed there was a real
transition — `check-dot.ps1:304` parks the pointer at (4,4) before every button read — so the
same-point no-op does not explain it, and I did not find the cause. `-Onto` will now name a covering
window if that is what it was, and the nudge removes one whole class of silent non-delivery, but
neither is a demonstrated fix for that specific failure.

**A live lead, not chased:** the add-in's tip code already knows that *"a still pointer sends no
WM_MOUSEMOVE, so the wait that would have started the clock again never happens"* (`strip.cpp`, the
tip re-arm). If the strip's `hot` state can be cleared by the row moving under a still hand — the strip
does move, and the log carries `strip had drifted ... correcting` lines — then the dot failure is the
same shape as the drag-cursor defect fixed in `RESULT-cursor-20260905.md`: **state that depends on
mouse movement, on a UI where things move under a still pointer.** That is an add-in question and was
out of scope here.
