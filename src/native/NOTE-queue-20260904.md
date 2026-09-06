# The queue out of the 2026-09-04 report — do all of it

**They said: "do all next session." That means NO TRIAGE.** Do not come back with a ranked list and
ask which. The ordering below is mine and it is worst-first; work down it. If an item turns out to
depend on one they did not separately approve, do the minimum it needs and say what was left out.

**State when this was written:** `HEAD` = `ef49c79`. Working tree carries two *finished, verified*
fixes — `strip.cpp` (hover midpoint), `stack.cpp` (`ParkMinimised`), `tools/palette-test.cpp`,
`tools/check-stack.ps1`, plus `RESULT-unseen.md`. **All 16 suites green, 851 checks.** Not committed,
not published. Read `RESULT-unseen.md` first; do not redo that work.

**Before touching anything: run the full battery once to lay down the baseline logs.** Every suite
moves the add-in log to `%LOCALAPPDATA%\WordTab\history\<stamp>-<suite>.log` and keeps it, and that
is the only real A/B this project has. A green battery before and after is not one.

---

## 1. The row grows 21-27px on every drag, and the registry banks it one drag late

**HIGHEST HARM. It walks the row back toward the "3 pages wide" complaint they shouted about
three times.**

**Evidence.** Two drags, both tracing *0 resizes*: `1240x1397 → 1261x1410 → 1288x1437`. Their
registry holds `(-1,667 1261x1410)` while their live windows are `1288x1437`.

```
10:01:02.079  WM_EXITSIZEMOVE hwnd=0x111F6
10:01:02.079  stack  the row remembers its size: (-1,667 1261x1410)     <- correct, matches the drag
10:01:02.127  WM_SIZE 0x31376 restored 1266x1424 outer (-11,662 1288x1437)  WordTab moved it
10:01:02.166  WM_SIZE 0x1140C  ... same ...                                 WordTab moved it
10:01:02.190  WM_SIZE 0x111F6  ... same ...                             nobody here asked for it
```

**What is settled.** Word hands the *master* a bigger rectangle ~48ms after `WM_EXITSIZEMOVE`; the
`WM_SIZE` line added in `ef49c79` for exactly this says `nobody here asked for it`. We take the
proposal `const` (`frames.cpp:359`) and cannot be growing it. `StackOnFramePosChanging` then copies
`pos->cx/cy` verbatim onto the followers (`stack.cpp:1018`, batched at `stack.cpp:1050`), which is the
row working as designed — without it the followers would just be left behind.

**REFUTED, do not re-chase:** a coordinate-convention mix in `RememberRowRect`. `rcNormalPosition` and
`GetWindowRect` are provably the same space for this window — at 07:53:50 an apply of
`(21,717 1240x1397)` produced a `WM_SIZE` of exactly `(21,717 1240x1397)`. The registry round trip is
lossless and is the one part of this path that demonstrably does not grow.

**What we own: the ratchet.** `RememberRowRect` runs at `WM_EXITSIZEMOVE` (`frames.cpp:349`), i.e.
*before* the growth, so a drag banks the **previous** drag's growth. One drag in a session banks
nothing; two banks one growth. That is why 09-03 → 09-04 did not ratchet (one drag) and why their
registry now holds a grown size (two drags).

**PREDICTION TO CHECK FIRST IN THE NEXT REPORT:** their next Word start should log
`the row takes the size it was left at ... (-1,667 1261x1410)` — wider than the `1240x1397` it has
started at twice. If it does, the ratchet is confirmed from the field.

**Do the PROBE first, and only the probe, until it has been read.** The message that carries the
growth is recorded nowhere: `TraceRecord` is dead outside the modal loop (`frames.cpp:141-143`). In
`frames.cpp` `WM_WINDOWPOSCHANGING`, when `!g_inModalLoop && !StackIsSyncing() && !(pos->flags &
SWP_NOSIZE)` and `pos->cx/cy` differ from `GetWindowRect`, log once: hwnd, current rect, proposed
rect, `pos->flags`, `pos->hwndInsertAfter`, `IsZoomed`, `StripDpiOf`, **`GetWindowPlacement().rcNormalPosition`
beside `GetWindowRect`**, and **`AdjustWindowRectEx`'s own answer for that frame**.

**Two leads the probe should be able to settle:**
- Both sessions' *first* drag landed on exactly `(7,671 1261x1410)` from two **different** drag-end
  positions (`(21,717)` and `(16,677)`) — and that is also the rect every `attach (existing)` line
  reports as Word's own for a fresh frame. A landing rect independent of where the drag ended is the
  signature of something re-asserting a *stored* rectangle, not a delta.
- Drag 1's growth `+21x+13` is within a pixel of that frame's non-client margin (`22x13`), and its
  post-drag *client* size (`1239x1397`) is within one pixel of its pre-drag *outer* size
  (`1240x1397`) — the fingerprint of an outer rect read as a client rect. **Drag 2 (`+27x+27`) does
  not fit that story**, so it is a lead, not a finding.

**TRAP — do not write the guard as "pin the size outside the modal loop".** `frames.cpp:485` already
says it: *"Still not covered: Win+Left and Win+Right. They snap without the modal loop and without a
system command, and there is no message that says a person did it."* A guard keyed on
`!g_inModalLoop` silently springs those snaps back and the user gets no error. Gate any rewrite on a
short window (~250ms) since that frame's last `WM_EXITSIZEMOVE` — the growth arrived at +48ms and
+50ms in the three observed drags.

**Second trap:** any DPI exemption must be a **state** comparison (`GetDpiForWindow(frame) !=
state->dpi`), not an event. Item 6 shows geometry crossing a 120↔144 boundary with **no
`WM_DPICHANGED` at all**, so an event-keyed exemption would never open and would pin the row at the
wrong size across a real monitor change on their two-DPI rig.

**Do NOT "fix" it by re-banking the grown size after the dust settles.** That records a size the user
never chose and makes the ratchet immediate instead of one drag late.

---

## 2. The 500ms janitor — dot poll included — runs inside Word's modal drag loop

**Nobody was looking at this; it came out of the completeness sweep.**

`g_janitor = SetTimer(NULL, 0, 500, JanitorProc)` (`strip.cpp:6574`) is a **thread** timer, so its
`WM_TIMER` is dispatched by whatever pump is running — and during `SC_MOVE` that is Word's own drag
loop. `JanitorProc` (`strip.cpp:6191`) has no modal-loop guard and **cannot have one**:
`g_inModalLoop` is a file static at `frames.cpp:107` and `wordtab.h` exports no accessor.
`PollModified` stands down for `StackCloseInFlight()` (`strip.cpp:5942`) and for a disabled frame, but
not for a drag.

**It demonstrably happens in the field.** Between their `ENTERSIZEMOVE` and `EXITSIZEMOVE` sit a dot
poll and three `strip had drifted ... correcting` lines; drag 2 contains two complete janitor rounds
473ms apart. Their drag traces carry gaps of **260.8ms and 173.3ms** — 29% of a 908.9ms drag with the
row not following the mouse. A full row pass there costs **mean 179,188us, worst 520,183us**.

**Fix:** export a "is a drag in progress" query from `frames.cpp` and have `PollModified` stand down
for it exactly as it already does for `StackCloseInFlight`, and log the pass it skipped.
**Stand down `PollModified`, not the whole janitor** — otherwise the strip stops following the frame
mid-drag and item 6's drift becomes visible instead.

**Dev rig:** reproduces the *structure* (a tick lands inside a drag here too) but not the cost — a
local document answers in ~430us, three orders of magnitude below theirs. Complete as probe + guard +
log line; then read the next report's gap histograms. If the `>50ms` buckets empty out, that was it.

**Not measured and worth timing while in here:** `SampleChrome` does `GetDC(NULL)` (`strip.cpp:899`)
then 24 × `WindowFromPoint` + `GetPixel` on a DWM-composited screen DC (`strip.cpp:919, 926`) every
fourth tick, with **no QPC brackets at all** — the only janitor stage with no instrument. It is a live
candidate for those same gaps.

---

## 3. The active-window dot poll is uncapped and memoryless — 35s stale on the tab being typed in

The comment at `strip.cpp:6083-6094` says this line *"has never yet appeared after the first pass of a
process"*. **It has now, seven times across two sessions.** Worst: `52056us → rung 3 → 20000ms`, and
the observed gap between consecutive active passes was **35.35s**. The direction that misleads is a
*saved* document still showing the unsaved dot.

**Fix**, at `strip.cpp:6100`:
```c
int want = RungFor(us);
*rung = (want > 1) ? 1 : want;      // the typed-in tab is never asked less often than every 2s
```
Rungs 2-4 exist to protect the UI thread from the **row's** 90,000-518,000us passes; the active pass's
worst across two sessions is 52ms. Nothing in the report justifies a 20s or 60s cadence on the
foreground window.

**Be honest in the comment:** `strip.cpp:6046-6048` claims the typed-in tab "is still asked twice a
second however far the row has backed off" — 500ms. A 2s cap is a deliberate weakening of the code's
own stated invariant, not the "outer edge" of it. Say so rather than citing the comment as licence.

**Also add the hwnd to the active governor line** (`strip.cpp:6119-6121`). It does not currently say
*which* window it measured, so neither this report nor the next can tell whether two expensive samples
were even about the same document. `connect.cpp:549-573` walks all of `Application.Windows` regardless
of `count`, so an active pass is not self-evidently about the active document.

---

## 4. The governor ping-ponged 500↔2000ms 36 times in 3m50s

**Not** `f333349`'s SharePoint bistability — plain jitter across a bare 4000us threshold with no
hysteresis. Their 7-window row sits at 2400-3900us with spikes over 4000. Proof it is not the old
fault: **in the whole 230s episode not one line above "2000 ms" was written**, so every pass in it —
logged or not — cost ≤15000us, whereas f333349's produced 90,000-518,000us passes (which this same
log shows the ladder handling correctly a minute later: `510616us → 60000ms`).

**Costs 0.10 points of UI thread** against a 1% budget — no user harm. **Costs 49% of the log's
bytes/hour**, in the same words and shape as the one line that matters.

**Fix, one line at `strip.cpp:6079`:**
```c
else if (RungFor(basis * 2) < *rung)     // was: else if (want < *rung)
```
`basis * 2` can only map to a rung ≥ `RungFor(basis)`, so it is strictly harder to descend — the same
direction f333349's fix already went, so it cannot reintroduce it. Replaying the report's own numbers:
the 72-sample warm band goes from **34 cadence changes to 1**; the `12:16-12:19` tail is
**bit-identical**. Keep `want` used at 6077 or `-Werror` will bite.

**MIRROR IT BY HAND into `tools/governor-test.cpp:64`** — `check-governor.ps1:29-30` says in as many
words: *"Keep them in step by hand when PollModified changes."* Also check `tools/check-dot.ps1:556-561`,
which asserts the first cadence line reads 500ms (it survives — dev-rig passes never leave rung 0).

**Land 3 and 4 together.** They are one arithmetic in one `if/else` with one hand-maintained copy.
Do **not** apply hysteresis to the active branch — its memorylessness is deliberate and the cap
already bounds the damage.

---

## 5. OnePage ran six times and logged nothing, and `FALSE` has five meanings

`WordTabOnePageView` returns FALSE for five different reasons and the caller logs on none, so
"healthy, nothing to do" and **"the correction was attempted and Word refused the write"**
(`connect.cpp:441`) are byte-identical in the log. The second leaves the document three pages across
at 10% zoom — their original complaint, verbatim: *"every time i mean"* — and the next six reports
could not tell us.

**The right shape is NOT to re-derive the illness test in `stack.cpp`.** Two agents independently
drafted an `else` there and both got it wrong the same way: the real test is
`manyPages = (read && columns > 1 && columns < 99)` (`connect.cpp:415`) because **99 is Word's
"as many as fit" and IS the healthy value**. Dropping the `< 99` prints "WORD REFUSED THE WRITE" on
every healthy join. That is [[one-copy-of-what-decides-truth]] on its first writing.

**Do it properly:** have `WordTabOnePageView` return the *reason* it decided (a small enum set at each
of its five exits) and have `stack.cpp:758` print what it was handed. Note `columns==0 && percent==0`
is **not** unique to "never asked" — `connect.cpp:409-410` writes the out-params unconditionally after
the reads, so it also covers "Zoom was obtained and both `GetLongProperty` calls failed".

Then close the suite hole: `check-onepage.ps1:166` and `:183` only `Write-Note` two variables their
own comments (155-156) claim are asserted.

**Free, and worth doing before shipping anything:** their report script does not read zoom at all
(`tools/workrig-recon.ps1` has no `Zoom`/`PageColumns` reference). Adding `View.Zoom.PageColumns` and
`.Percentage` per window to it would settle this from the next report with no add-in change.

---

## 6. The strip briefly rendered at the other monitor's DPI

A clean ×1.2 then ÷1.2 across a minimize/restore, **no `WM_DPICHANGED`, no monitor crossing**
(every rect in the episode is inside DISPLAY2). Self-healed both times, nobody complained.
`1218×120/144 = 1015`, `48×120/144 = 40`; then `1218×144/120 = 1462`, `48×144/120 = 58`.

**REFUTED, do not re-chase:** that WordTab computed the wrong size. `ApplyMetrics` is the only writer
of `state->stripH` and `strip.cpp:1319` is the only sizer of the strip, always with `state->stripH`,
which was 48 throughout. **WordTab wrote the correct number both times — it fixed the strip, it did
not break it.** Also refuted: a thread-wide awareness flip. In the *same* `PlaceStrip` call,
`ChildRect(frame, _WwF)` came back clean while `ChildRect(frame, strip)` came back scaled, which no
thread-wide context can do.

**So: probe only.** Extend the drift line at `strip.cpp:1308` with `state->dpi`,
`GetDpiForWindow(frame)`, and the three awareness contexts —
`GetWindowDpiAwarenessContext(frame)`, `…(strip)`, `GetThreadDpiAwarenessContext()` — reached the way
`GetDpiForWindow` already is at `strip.cpp:538-546` (`GetProcAddress` off `user32`, because
w64devkit's headers predate them). If the strip's context differs from the frame's, this is
mixed-mode virtualisation of our own child and it is ours to fix.

**TRAP — do not ship the "suspect read, don't correct" guard as first drafted.** Keyed on the shape of
one read (`58/48 = 1.2083` vs `144/120 = 1.2`, inside 1%), it fires on the *restore*, where the strip
was **genuinely** 1462x58 — turning a self-healing 1s flicker into a permanent 244px overhang with
33px of strip over the document. Any such guard needs "second consecutive tick showing the same
suspect read wins, correct it anyway".

WordTab declares **no DPI awareness anywhere** — no manifest, no `.rc`, no `SetThreadDpiAwarenessContext`
— so every rect it reads is in whatever context Word's UI thread is in. Fold this probe and item 1's
into **one** probe with two callers.

---

## Refuted — do not schedule

**The log roll.** Measured 9,126 B/h against the 512KB cap is ~57 hours of Word, seven working days;
the session investigated used 8.4% of it. Its one live half (the strip drift line re-firing on
consecutive ticks) rides with item 6. **One free thing worth taking:** have the report collector print
the first line of `wordtab.log` verbatim, so the next reader can tell a rolled log from a fresh
install.

**The `tab name |Word|` line.** Real but paint-neutral: `0xA1736` is hidden, in no painter's tab list,
and never joined the row. `StripTabName`'s gate (`strip.cpp:5289`) rejects `""` and `"Word"`
identically, so the placeholder never reaches a pixel. Worth one guard in `ReadTitle` purely so the log
stops asserting a tab name for a window that has no tab — fold it into item 6's slice if convenient,
along with the same call chain's other half (`ApplyInitial` deriving a `natural` rect 43px short from
an unlaid-out frame). Both are `TryBind` running on a frame Word has created but not finished; **one
guard above both**, not two.
