# The clean battery that `NOTE-next-20260905.md` asked for, and what it caught — 2026-09-05 evening

> **SUPERSEDED ON ITS MAIN POINT, later the same evening — read `RESULT-cursor-20260905.md`.**
> The `IDC_SIZEALL` regression in section 3 is **FIXED**, and section 3's framing was wrong in a way
> worth knowing: there is **no second ingredient**. The defect is single, older than `ef49c79`, and was
> masked in every previous build by the 744ms chrome sampler. `-Patch trybind` "fixing" it is read
> backwards here — it makes each janitor tick HEAVIER, which masks the fault exactly as the slow
> sampler did, which is also why it cost 17 other checks.
> Everything else in this file stands, including all four harness lessons in section 4.

Input: `NOTE-next-20260905.md` section 5, item 1 - *"A clean full battery. Nothing else should ship
before it."* The battery was run, twice. **It is still not clean, and one of the reds is a real
regression this working tree has and `ef49c79` does not.**

`HEAD` is still `ef49c79`. Nothing is committed and nothing is published.

**Where the evidence lives**, because every quote below comes from one of these:

| | |
|---|---|
| suite console output, per suite per battery | `%LOCALAPPDATA%\WordTab\history\*.out.txt` |
| the add-in's own log, archived per suite | `%LOCALAPPDATA%\WordTab\history\*.log` |
| the `ef49c79` baseline battery, kept out of the prune window | `%LOCALAPPDATA%\WordTab\history\baseline-ef49c79-wt\` |
| A/B variant builds and their suite output | `%TEMP%\wordtab-ab\` |

**The history directory prunes to the newest 40 of each kind** (`WordTabHarness.ps1:131`,
`:328`). Two batteries write 32 `.out.txt` and 17 `.log` files, so **anything you want to keep must be
copied out**, the way the `ef49c79` baseline was.

---

## 0. The crash cost nothing

Checked first, because "resume, something crashed" could have meant a half-written file:

| | |
|---|---|
| `strip.cpp` last written | 17:04:37 |
| `WordTab.dll` built | 17:11:16 |
| `RESULT-queue.md` written | 17:10:45 |
| `NOTE-next-20260905.md` written | 17:13:17 |

The bank post-dates the build, and the **deployed** DLL was byte-identical (SHA256 `B780E41F…`) to
`src\native\build\WordTab.dll`. Nothing was lost.

**Check that pair every time.** Word loads `%LOCALAPPDATA%\Programs\WordTab\WordTab.dll`
(`install\install.ps1:71`, `:187`); `src\native\build.ps1` writes to `src\native\build\` and copies
nothing out. A battery can be green about the wrong binary.

```powershell
pwsh -File src\native\build.ps1            # -> src\native\build\WordTab.dll
pwsh -File install\install.ps1             # copies it to %LOCALAPPDATA%\Programs\WordTab\
pwsh -File tools\check-all.ps1             # the gate: all 16 suites. Word must be closed first.
pwsh -File tools\check-all.ps1 -Only reorder
```

---

## 1. Two batteries

| | battery 1 (17:22) | battery 2 (17:48, console minimised) |
|---|---|---|
| result | 14/16, 863 checks | 14/16 |
| red | `reorder`, `dot` | `reorder`, `scroll` |

`dot` was 71/71 in battery 2 and in two separate re-runs. `scroll` was 62/62 in battery 1. Only one
failure appears in **both**, and it is the one that matters.

---

## 2. Three findings that are NOT the add-in

### The console is a foreground thief, and it was mine

`check-reorder` reads a tab's title through the foreground window. Windows Terminal - the console
running the battery - took the foreground while it read:

```
tab 0: the foreground was "Terminal" - taking it back and re-reading
order: Terminal | wordtab-order-Echo.rtf | wordtab-order-Alpha.rtf | Document1
FAIL  the tabs either side of it kept their order
```

The harness noticed and took the foreground back, **and the value it had already read was still
wrong**. `CASCADIA_HOSTING_WINDOW_CLASS` appears in no retained 0904 output and in three suites of
battery 1; minimising the console gave `terminal=0` across all sixteen in battery 2.

### `dot`'s hover failure is an undelivered mouse move, and the add-in was right

`FAIL  hovering a modified tab brings the x back` - the close button stayed a dot with the pointer on
it. `strip.cpp:2553` is `if (modified && !hot && !downClose)`, so `hot` was FALSE.

**The tooltip proves it was.** The tip only appears when the pointer arms a tab, and the failing run
logs **eight seconds of silence** (`17:38:28.851` to `17:38:37.000`) across the whole hover step where
the passing run has four tip show/hide pairs. The strip never received the move.

The suite believed otherwise because `WordTabHarness.ps1:979` confirms with `[WordLayout]::Cursor()`
- GetCursorPos, which is not the event the strip needs. Same shape as the `rowsize` lesson in
`RESULT-queue.md`: *a test that waits for event A and then asserts on event B is not flaky, it is
wrong.* **Unfixed.**

### `scroll`'s widen is an old race that usually loses

The suite widens the window to 990 and measures. Battery 1 read `strip 964 px` and passed; battery 2
read `strip 874 px` and failed five checks off that one reading:

```
18:00:07.042  resize  hwnd=0x130304  NOBODY HERE ASKED: now (900x700) -> proposed (990x700)
18:00:07.079  stack   hwnd=0x130304  out of step: (990x700), master (900x700) - putting it back
```

**`out of step` appears six times in the `ef49c79` baseline's own `soak-stack` log**, so the mechanism
is not new. Not chased further.

---

## 3. THE REGRESSION

```
==> TabGhost=0 keeps the tear-off and drops only the carried card      (check-reorder.ps1:1269)
    PASS  the tab still comes out of the row with the card switched off
    cursor during the TabGhost=0 tear-off 0x10005  (SIZEALL 0x10015, arrow 0x10003, I-beam 0x10005, NO 0x10017)
    FAIL  and the pointer still says so - IDC_SIZEALL, unchanged       (the Assert is check-reorder.ps1:1351)
```

Cite the step name at `check-reorder.ps1:1269`, not the Assert - the Assert moved from 1340 to 1351
when this session added the diagnostic, and will move again.

Every neighbouring check passes: the tab tears off, nothing is carried, no window is made. Only the
cursor is wrong, and **it is Word's I-beam** - which `check-reorder.ps1` already names, in the
`TabTearOff=0` section above, as what the pointer shows over the document when the add-in never
touched it. The cursor is set at `strip.cpp:4049` on the tear transition and restated every mouse
move at `:4087`; the suite reads it 400ms after the last move, button still held.

### The evidence

| binary | cursor | check | run |
|---|---|---|---|
| `ef49c79` sources, built and deployed 09-05 | `0x10015` | **PASS, 118/118** | `%TEMP%\wordtab-ab` |
| `ef49c79` + tree `{connect.cpp, stack.cpp, wordtab.h}` | `0x10015` | **PASS, 118/118** | `%TEMP%\wordtab-ab` |
| that + tree `strip.cpp` (modal-loop stub) | `0x10005` | FAIL | noisy run, 7 red |
| full working tree | `0x10005` | FAIL, four runs | batteries 1-2 + two standalone |
| full tree, `TryBind` guard neutralised | `0x10015` | PASS, 17 of 118 red | `-Patch trybind` |
| **a working-tree build of 2026-09-04 18:40, guard present** | not printed | **PASS** | `history\20260904-184022-930-reorder.out.txt:190` |

**Word did not change under us**: `build=16.0.20326` in the baseline log and in today's.

### Ruled out by measurement

- **The chrome sampler.** `TabThemeSample=0` clears `g_sampleEnabled` (`strip.cpp:6880`), which gates
  the only call to `SampleChrome` (`:6672` guarding `:6725` - verified, there is exactly one caller).
  The check still failed. `CAPTUREBLT` compositing the cursor layer was the first hypothesis and it is
  **wrong**.
- **The dot poll.** `TabDot=0` returns out of `PollModified` at `strip.cpp:6173`; still failed.
- **The drag code itself.** The diff's hunks in `strip.cpp` run `…-1305,11…` and then `…-5668,6…`, so
  **old lines 1316 to 5667 are untouched** - and the tear-off path (`DragCursor`, the threshold, the
  transition block) sits at old ~3880-3925. Byte-identical to `ef49c79`. *(Stated loosely in the first
  draft as "no hunk between 1305 and 5668", which is false for 1305-1315 - the `PlaceStrip` logging
  change is in there.)*
- **`connect.cpp` / `stack.cpp` / `wordtab.h`.** Bisected out: 118/118 with all three from the tree.
- **`PlaceStrip`, `Restore`, `DpiOf`.** Logging, comments, and a pure refactor.
  `StripDescribeDpiContext` reads awareness contexts and never sets one.

### What is actually established, and what is NOT

**The guard is necessary. It is NOT sufficient, and the first draft of this file got that wrong.**

Neutralising `TryBind`'s unfinished-frame guard (`strip.cpp:5892`) makes the cursor right. But a
**working-tree build that already contained that guard PASSED this check on 2026-09-04 evening**:

- `history\20260904-184022-930-reorder.out.txt:190` - `PASS  and the pointer still says so`
- its log `20260904-184022-925-reorder.log` carries `Word has made this frame and not titled it`
  **three times**, so the guard was in that binary
- `git show HEAD:src/native/strip.cpp | grep -c "not titled it"` is **0**, so it was not `ef49c79`
- and that log's `chrome sample: 723570 us` / `750083 us` dates it **before** the `BitBlt` rewrite
  (the same line reads `30977 us` on 09-05)

So the guard leaves a frame unbound, and **something else that landed in `strip.cpp` between that
build and its `17:04:37` mtime on 09-05 turns that into a lost cursor.** Nothing is committed, so
**no snapshot of the passing tree exists to diff against** - that is the central difficulty.

**A hypothesis that fits every row, explicitly NOT tested:** the old per-pixel sampler stalled Word's
UI thread for ~0.75s every two seconds and may have been masking a timing-dependent race. That would
explain 09-04 passing (slow sampler), the tree failing (fast sampler), *and* `TabThemeSample=0`
failing (no sampler at all - also fast). **Do not write this up as the cause.** It is the cheapest
thing to test next: revert `SampleChrome` to per-pixel `GetPixel`, leave it enabled, and predict PASS.

**UNTESTED, not refuted - and the first draft of this file said "REFUTED", wrongly.** The idea that
`TryBind`'s per-tick `GetWindowTextW` re-enters Word's window procedure mid-drag was "disproved" by a
`dragstand2` build that still failed. That inference does not hold: the build was contaminated (below)
and **a superset failing does not prove the subset fails**, because the extra change can break it
independently. Here the extra change was not inert - it also suppressed `ReadTitle` for the whole of
every drag for any frame ever deferred, since `untitledLogged` is set once at `strip.cpp:5896` and
never cleared. **The hypothesis is live and needs a clean single-site run.**

### The contaminated experiments - do not cite their numbers

The patch anchor `wchar_t raw[256];` + `GetWindowTextW(state->frame, raw, 256);` occurs **twice** in
`strip.cpp` - in `ReadTitle` (~5491) and in `TryBind` (~5888) - and `.Replace()` rewrote **both**. So
the `dragstand` result ("fixes the cursor, 27 of 113 red") and the `dragstand2` result each tested a
second change nobody asked for. **`tools\ab-binary.ps1` now anchors on the comment line belonging only
to `TryBind` and throws unless the anchor matches exactly once.** Both need re-running.

The `trybind` row is **clean**: its anchor is the single line at `strip.cpp:5892` and occurs once.

### Two constraints on any fix

- **The guard is load-bearing.** Neutralising it cost **17 of 118** checks in the same run that fixed
  the cursor. It exists for the 116 empty-title binds in `RESULT-queue.md`. "Just remove it" is not
  available.
- **Do not conclude the default configuration is safe.** `TabGhost` defaults to on, and with it on the
  normal tear-off cursor check passes - but our own card window then sits under the pointer and would
  absorb a `WM_SETCURSOR` that leaked to Word, so a green check there may be the ghost hiding the same
  fault.

### Where to start, concretely

**The harness is `tools\ab-binary.ps1`.** It copies the tree (never edits it), hashes the deployed DLL
before anything moves, empties the log, minimises the console, runs a suite, force-closes Word if it
outstays the wait, and hash-checks the restore.

**It has been self-tested once, on 2026-09-05 at 20:0x**, and reproduced the second row of the
evidence table exactly - so the harness itself is known good, even though the throwaway scripts the
09-05 A/B runs actually used are gone:

```powershell
pwsh -File tools\ab-binary.ps1 -FromTree connect.cpp,stack.cpp,wordtab.h
#   variant A6136C39DFF17E5F   deployed-was B780E41FF32B02F9
#   RESTORED: the deployed DLL is what it was (B780E41FF32B02F9)
#   cursor during the TabGhost=0 tear-off 0x10015  (SIZEALL 0x10015, ...)
#   PASS  118 checks, 0 failures                    -> %TEMP%\wordtab-ab\fromtree-reorder.txt
```

Re-run that whenever the harness is edited: it is the one command whose answer is already known.
**Note the variant hash is not stable across build directories** (the same sources built under a
different scratch path gave `62B5AA3B…` earlier), so compare behaviour, not bytes, between runs.

Then, worst-first:

1. **Find the second ingredient.** What else changed in `strip.cpp` between the 09-04 18:40 build and
   `17:04:37` on 09-05? There is no committed snapshot, so this is reconstruction: `RESULT-queue.md`
   is the 09-04 account and the `BitBlt` rewrite is provably inside that window (723,570us vs
   30,977us). Test the sampler-revert prediction above first - it is one build.
2. **Instrument capture, which nothing currently reports.** Log `GetCapture()` and `GetCursor()` on
   every janitor tick while `g_dragging`, and log `WM_CAPTURECHANGED` on the strip with the window
   that took it. If capture is NULL when the suite reads the cursor, this is capture loss and the
   question is what cancels it; if capture is still ours and the cursor is already the I-beam, then
   something calls `SetCursor` on our own thread.
3. **Re-run the two contaminated patches** with the fixed single-site anchors.
4. **The `TabGhost=1` control** for the last bullet above. Note this is a **suite edit, not a registry
   edit**: `check-reorder.ps1` sets `TabGhost` itself for that section, so forcing the ghost on means
   changing that section or adding a switch.

---

## 4. Harness lessons paid for this session

- **Reset the add-in log before any direct suite run.** `check-all.ps1` empties it before each suite;
  running `check-reorder.ps1` on its own does not. Seven back-to-back runs took it to 524,176 bytes
  and the add-in deleted it mid-suite (`log.cpp:11` sets the 512KB cap, `log.cpp:52` does the delete).
  The harness caught it and threw - `750d49f` working as intended - but the run proved nothing.
  **This also falsifies `NOTE-next-20260905.md` section 4's "the log roll is still not a thing"**: it
  is not a thing *for a battery*, and it is very much a thing for repeated single-suite runs.
- **A binary-swap harness must force Word closed before restoring.** Word holds the DLL open; an
  aborted suite leaves it up; the polite wait expired, `Copy-Item` threw, and **the patched binary
  stayed deployed**. Verify the restore by hash and fail loudly if it does not match.
- **An anchor that matches twice is as bad as one that matches nothing.** See the contamination above.
  Require exactly one.
- **Three PowerShell traps, each caught only by a guard:** `Set-Content -NoNewline` joins a line ARRAY
  with no separator and collapsed every extracted source file onto one line; `pwsh -File` passes
  `-Param a,b` as one literal string (`check-all.ps1:39-42` says so in its own words); and `-like`
  reads `[` and `]` as a character class, so the anchor `raw[0] == ...` searched for `raw0`.

---

## 5. Changed in this session

- `tools\check-reorder.ps1`, **+11 lines**: the TabGhost=0 cursor check now prints the cursor it saw
  against SIZEALL, arrow, I-beam and NO, in the idiom the `TabTearOff=0` section already used. It had
  failed four times saying only "not SIZEALL". **That line is what turned the diagnosis** - `0x10005`
  named Word's I-beam immediately.
- `tools\ab-binary.ps1`, **new**: the variant-build harness described above, self-tested once against
  the known 118/118 case.
- This file, and a banner on `NOTE-next-20260905.md`.

**No add-in source was changed.** Nothing was committed and nothing was published.

**Deliberately NOT done:** `check-all.ps1` was not changed to minimise the console, though that is the
durable fix for section 2's thief. It hides the battery's output from whoever is watching, which is a
change to somebody else's tool. `ab-binary.ps1` does it for its own runs, with `-KeepConsole` to
disable.

---

## 6. Still open, worst first

1. **The `IDC_SIZEALL` regression.** The guard is necessary but not sufficient; the second ingredient
   is unidentified and there is no committed snapshot of the passing tree. It blocks the clean battery
   that blocks everything else.
2. **`Set-Pointer` confirms the wrong event.** Until it waits for something the strip actually did,
   `dot` and `reorder` will keep producing failures that are not about the add-in.
3. **`scroll`'s out-of-step race.** Pre-existing, now seen.
4. **`check-stack.ps1` throws instead of failing when it loses the foreground** - unchanged from
   `NOTE-next-20260905.md`.
5. **`RESULT-queue.md`'s closing "What is still open" list is stale**: its item 1 calls the chrome
   sampler *"Measured, understood, unfixed. Next slice."* while the same file's headline section
   documents the fix landing at 744ms -> 21ms with `check-look` 48/48. The headline is right.
6. Nothing is committed or published, and publishing was refused once before - **ask.**
