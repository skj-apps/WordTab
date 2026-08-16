# Hardening the test harness

**Status: the eleven suites now share one copy of every primitive that decides whether the harness is
telling the truth.** No product code changed. `src\native` is untouched; this slice is entirely about
`tools\`.

The brief was "retrofit `Set-Pointer` / `Invoke-StripClick` into the eight suites that still inject
clicks without confirming they landed, and narrow `Get-WordDialog` in `check-menu` and `check-title`".
The inventory that opened the slice found the problem was larger and more specific than that.

## What was measured before anything was changed

A full battery at `be0a2f2`, from a clean start: **409 checks, 0 failures, 19.1 minutes.** That run
exists because HEAD's status was genuinely unverified — the icons slice had rebuilt with `strip.cpp`
stashed and `be0a2f2` shipped only a note — and because this slice rewrites the proof machinery, so a
red suite afterwards would otherwise have been ambiguous between "the retrofit broke it" and "it was
already flaky". It cost 19 minutes and it paid for itself twice over.

Then an inventory of all eleven suites, one reader per file:

| | |
|---|---|
| injected inputs | **134** |
| …that confirmed they landed | **4** |
| `Start-Sleep` per battery | **8.7 minutes of 19.1** — 45% of the battery is literal sleeping |
| distinct `Get-WordDialog` implementations | **3** |
| …that were correct | **0** |

## The three dialog tests, and why none of them was right

"Is Word asking the user something?" is the most safety-critical measurement in the project:
`check-menu` deliberately provokes the save prompt, and a suite that cannot see a modal will drive
clicks and Escapes past it and can walk away leaving a real prompt up on the user's screen.

- `check-dot` excluded `OpusApp` and **`Net UI Tool Window`**, floor 150×60. It did **not** exclude
  `#32768`, so an open context menu read as a question.
- `check-menu` and `check-title` excluded `OpusApp` and **`#32768`**, floor 200×100. They did **not**
  exclude `Net UI Tool Window` — the exact window that cost a whole battery run.
- Eight suites had no dialog test at all in their close paths. They posted `WM_CLOSE` and threw on a
  timeout, so "a save prompt is up" and "Word is wedged" produced the same message.

**And all three would have called a context menu's *shadow* a question.** `tools\probe-dialogs.ps1`,
written for this slice, photographs nothing and enumerates everything: for our own 411×334 context
menu Word also puts up a `SysShadow` at **416×339**, which is not `OpusApp`, not `#32768`, and
comfortably over 200×100.

## The probe, and the design decision it settled

`tools\probe-dialogs.ps1` walks Word through six states and dumps every visible top-level window it
owns — class, size, title, and the text of every child control — with what the classifier calls each
one. Measured on this rig:

| state | what Word actually shows | kind |
|---|---|---|
| idle | `OpusApp` 900×700 | frame |
| our context menu | `#32768` 411×334 **and `SysShadow` 416×339** | menu, chrome |
| Backstage | seven `MSO_KEYTIP_WINDOW_CLASS` at 38×36 and 30×36 | chrome |
| save prompt | `NUIDialog` 920×713 | question |
| Save As (F12) | **`#32770` 1440×960, titled "Save As"** | question |

The classifier agreed with all six.

**The last row is the finding.** The classifier is a *denylist* — everything visible and big enough
that is not a frame, a menu or known chrome is a question — and that was written down as a deliberate
choice before the probe ran, on the argument that the two mistakes are not symmetrical: calling
chrome a question wastes a run, but calling a question chrome means driving input past a modal that
is really there. The probe turned that argument into a measurement. **An allowlist would have been
written as "a question is a `NUIDialog`", because that is the only prompt class this project had ever
seen — and it would have been blind to the `#32770` Save As, which is a real modal that really does
stop Word.**

The probe settled a second thing: **the save prompt has no window title and no child control with any
text in it at all.** So it cannot be identified by what it says even if matching on Word's English
were acceptable, which the tab-name slice already established it is not.

## Two live bugs in the suites, found by putting the copies side by side

Neither was found by reasoning about the code. Both fell out of the diff between three versions of
the same function.

**`check-menu` would minimise the desktop.** `Set-WordForeground`'s last resort is to minimise an
application that will not give the foreground up. `check-dot` and `check-title` guard that with
`class -ne 'Progman' -and class -ne 'WorkerW'`; `check-menu`'s copy had no guard at all — and
`check-menu` states the rule, with the reason, in a comment 260 lines further down and obeys it
there. Minimising `Progman` is Show Desktop, which takes Word down with everything else.

**Nothing guarded against minimising Word's own modal.** A save prompt is not an `OpusApp`, so
"not one of our frames, minimise it" minimises the question itself and leaves Word disabled behind a
modal nobody can see — the "Word beeping with nothing on screen" failure, reached from the opposite
direction to the one the tabs slice fixed. The shared `Set-WordForeground` now refuses to minimise
any window belonging to a WINWORD process.

**And two live `Kill()` calls.** `check-stack` and `check-strip` both ended with
`CloseMainWindow()` then `Kill()`. The rule against killing Word has been in the project notes since
the first slice, and the scrolling-row slice found and fixed a violation inside `soak-stack` — but
nobody went looking for the other two. `check-stack` runs **first** in `check-all`, so its cleanup is
the one with the most to poison. `check-tabs` was carrying the mirror image: its footer explains that
`CloseMainWindow` closes exactly one of N frames, and its own pre-run cleanup used `CloseMainWindow`.

## The rule about sleeps, and how it was broken within the hour

8.7 minutes of every 19.1-minute battery was `Start-Sleep`. The biggest single source was the same
line in eight files: `Start-Sleep -Seconds $(if ($i -eq 1) { 14 } else { 7 })` after launching Word on
a document. Those are now `Wait-WordReady`, which waits for the frame to exist **and to have a strip
on it** — faster when Word is quick, more correct when it is slow, and it says so on timeout instead
of sailing past.

The rule written down for which sleeps may be converted:

> Convert a sleep that is waiting for **Word** to do something — start, open a document, shut down.
> That is the harness getting into position, and how long it takes says nothing about the product.
>
> Do **not** convert a sleep that sits between an action under test and the assertion about it. That
> gap is part of the claim. A bounded wait there still goes red if the thing never happens, but it
> stays green if the add-in regresses from instant to nine seconds, and nothing reports that. It
> trades a timing check for an eventually check, silently, and the suite still prints PASS.

**Then the first battery after the retrofit failed on exactly that.** `check-stack`'s close-down loop
used to sleep 4 seconds after each `WM_CLOSE`; it became a wait for the window count to drop. The
count drops immediately. The assertion three lines later is "the last window is back in Alt+Tab" —
and putting `WS_EX_TOOLWINDOW` *back* is something the add-in's janitor does on its own half-second
cadence, after the close. The wait returned before the add-in had done it. The reasoning at the time
was "the assertions below are about the surviving window's strip and Alt+Tab entry, not about how
many windows there are", which is true and beside the point: what the sleep was really buying was
time for the add-in to reconcile.

Fixed by keeping the fast count-wait inside the loop and restoring an explicit 4-second settle before
the assertions — once, after the loop, rather than after every close — so the strength of the check
is exactly what it was and most of the saving is kept. `check-tabs` had the identical change and
passed that run, which is worse rather than better: the hazard was there and the battery happened not
to catch it.

**The general shape, worth more than the instance:** a sleep can be buying time for something that is
not visible in the code near it. Here it was the add-in's own janitor putting `WS_EX_TOOLWINDOW`
back. Before converting a sleep, ask what the *product* is doing during it, not only what the script
is waiting for.

## The regression the audit did not catch and the battery did

`Get-WordDialog` used to hand back a `WordLayout.Child`, which carries `Left/Top/Right/Bottom`. The
shared version returned a richer `pscustomobject` — class, size, title, pid, and the *kind* — and
dropped the rectangle. One caller in `check-menu` does `Format-Rect $asked`, in the branch that only
runs when Word actually asks where to save a never-saved document. It failed with "the property
'Left' cannot be found on this object".

Fixed by making the shared object a strict **superset** of what it replaced rather than by changing
the call site. A shared type that is not a superset of the thing it replaced is a breaking change
wearing the clothes of a refactor.

**What happened next is the best demonstration this slice produced.** That crash left Word up with
the Save As question on screen. `check-reorder` ran next and refused to start:

```
==> Closing 1 Word process(es) already running
    Word is asking something: class=NUIDialog 920x641 kind=question ||
Exception: Word would not close (question: class=NUIDialog 920x641 kind=question ||)
```

Under the old code that is "Word would not close (1 left)" with no idea why, and under the *loose*
old code the suite would have gone on and driven clicks past a modal. The size it named, 920×641, is
the second save-prompt size recorded in the notes from the battery run that started all this.

## What the module is, and what it is not

`tools\WordTabHarness.ps1`, dot-sourced by all eleven suites immediately after `Add-Type`. It holds
only the primitives that must not disagree between suites:

- **the classifier** — `Get-WordWindowKind`, `Get-WordWindows`, `Get-WordDialog`, `Get-WordMenu`,
  `Get-WordSavePrompt`, `Format-WordWindow`, `Format-WordChrome`
- **confirmed input** — `Set-Pointer` (cursor readback), `Get-ClassAt`, `Invoke-ConfirmedClick`,
  `Invoke-StripClick`, `Invoke-ConfirmedKey`, `Invoke-ConfirmedKeyOn`, `Start-ConfirmedDrag`
- **the foreground** — `Set-WordForeground` with its three guards, `Test-WordForeground`,
  `Test-WordHasFocus`
- **bounded waits** — `Wait-Until`, `Wait-WordFrames`, `Wait-WordReady`, `Wait-WordGone`,
  `Wait-WordDialog`
- **shutting Word down** — `Close-AllWord`, which never kills, never uses `CloseMainWindow`, stops on
  a question and hands it back rather than pressing anything

It deliberately does **not** call back into suite-local helpers. Every suite has its own
`Write-Note`, `Get-Frames` and `Set-Dirty`; a shared file whose behaviour depends on which suite
loaded it is not shared. It also does not assert — the suites decide what a failure means, and every
confirming function returns `$true`/`$false` so the arrival can be asserted as a check in its own
right. That is the point: **a hover that never happened and a hover that is not drawn produce
byte-identical evidence, and only one of them is a bug in the product.**

## What confirmation actually looks like at each kind of input

- **A pointer move** is confirmed by reading the cursor back. Never by what is under the point —
  five `MouseTo(4, 4)` calls across three suites park the pointer *deliberately off* the strip, and a
  "is the strip under it" check would be exactly wrong there.
- **A click** is confirmed by asking the system what window is under the point *before* clicking, and
  by re-running a scriptblock to recompute the point on every attempt, because the strip moves — 46px
  between two clicks a second apart, measured. Menu items expect `#32768`, not the strip: a
  right-click that lands on the document opens *Word's* context menu, which is also a visible
  `#32768`, so "a menu appeared" would otherwise pass while measuring the wrong menu.
- **A keystroke** is confirmed by checking a Word window has the foreground. Not specifically a
  *frame* — Word's dialogs, Backstage and menus are all legitimate targets for Escape, and demanding a
  frame would make it refuse to dismiss the very things it is most needed for.
- **A keystroke aimed at one window** (`Ctrl+W`, `Ctrl+S`) gets its own helper, because those close
  or save *whichever document is in front*. A `Focus()` that quietly did not take does not lose the
  keystroke — it applies it to the wrong document, and everything afterwards measures a window that
  should not exist.
- **A held-button gesture** is confirmed once, before the button goes down, and never again until it
  is up. Everything between `DragHold` and `DragRelease` is one gesture; re-aiming inside it produces
  a different gesture from the one named.
- **A click into the document page** is confirmed *passively* — ask what is under the point, change
  nothing. `Invoke-ConfirmedClick` takes the foreground back before it clicks, and the click into the
  page is what makes Word foreground *as user input*, which Windows honours unconditionally. Calling
  `Focus()` in that gap was measured leaving the window activated with the caret unplaced.

## Two negative tests that were passing for the wrong reason

`check-startscreen` clicks where the old "Word" tab used to be and asserts nothing happens.
`check-look` clicks the empty row's message and asserts it is a label, not a button. `check-scroll`
clicks a chevron that has gone dead and asserts the row does not move.

**"Nothing happened" is exactly what a click that never arrived also produces.** All three now assert
that the click landed on the strip as a check in its own right, so the negative means what it says.

## The `StrictMode` trap, from three directions

The project has recorded twice that a PowerShell function returning an **empty** array returns
nothing, so `.Count` on it is a hard error. This slice met the same trap in two more shapes.

**One item is not a one-element array.** The probe hit this on its first run: a function returning a
single value returns *that value*, so `$frames.Count` on a Word with exactly one window is the same
hard error — and one window is the only state the probe ever runs in.

**`@(...)[0]` on an empty match throws, and the guard after it is dead code.** Every suite's
`Get-Parts` built `Strip = @($kids | Where-Object { $_.Class -eq 'WordTabStrip' })[0]`, and every
suite's `Get-TopStrip` then wrote `if (-not $parts.Strip) { throw 'has no WordTab strip' }` — a guard
that can never run, because indexing the empty match throws first. Verified directly on this rig,
pwsh 7.6.5:

```
Set-StrictMode -Version Latest
@(1,2,3 | Where-Object { $_ -gt 99 })[0]    ->  Index was outside the bounds of the array.
@(1,2,3 | Where-Object { $_ -gt 99 }).Count ->  0
```

Sixteen occurrences across eight suites, now `| Select-Object -First 1`, which yields `$null` and
makes those guards reachable for the first time.

## The adversarial audit

Four reviewers, one lens each — PowerShell semantics, gesture invariants, the sleep conversions, and
behaviour changes hiding as refactors — then three independent skeptics per claim, each told to
refute it and to default to refuted when uncertain. **18 claims, 5 survived, and they are two
distinct defects.** Both were in code this slice added, and both were unanimous: 0 refutations of 3.

**A wait that throws instead of waiting.** The new `Open-Document` in `check-dot` and `check-title`
polled `Wait-Until { $null -ne (Get-Parts $new[0]).Strip }` — waiting for the add-in to carve its
band. `Get-Parts` is the function above: it throws on a strip-less frame rather than returning
`$null`. So a wait for "the strip has appeared" throws at exactly the moment the strip has not
appeared, which is the only moment it is ever called. It never fired in three battery runs, because
the janitor is usually quicker than the first 400ms poll. That is what a latent defect looks like:
green until the machine is busy. Now the guarded `@(...).Count -gt 0` shape, which is what
`Wait-WordReady` already used for the same reason.

**`Close-AllWord` read for the prompt too early.** `WM_CLOSE` is *posted*, so nothing is instant, and
reading for a dialog immediately after posting it reliably reads "no dialog" — after which the loop
posts a second `WM_CLOSE` at a live modal, which is precisely what the comment two lines above says
must not happen. The per-suite loops it replaced avoided this by sleeping 1200–2000ms first. Sleeping
is the wrong fix twice over: too short and the prompt is still missed, too long and it is paid on
every close in every suite. It now waits for **whichever of the two outcomes happens** — the window
goes, or Word puts a question up — and only then takes the double read.

Both were found by reading, not by running. The battery had already gone green over the first one.

## The defect the harness found, which is why the battery is not green

**`check-title` and `check-dot` fail, and they are right to.** The strip is landing **38 physical
pixels too high, overlapping Word's ribbon**, and the `+` is then drawn under `NetUIHWND` and cannot
be clicked.

Measured, from the failing run's own diagnostics: the strip reports itself at `(1536,432 874x64)`
and the `+`'s centre at `(2372,464)` — which is *inside* that rectangle, 836px into an 874px strip,
dead centre vertically. `WindowFromPoint` at that pixel answers `NetUIHWND`. Both reads are live. A
clean Word in the same 900×700 window puts the strip at `(1536,470)`: **38px lower.** In frame client
coordinates the good layout is ribbon `0..356`, `DropShadow 356..368`, strip `356..420`, `_WwF 420`;
the failing one has the strip at roughly `318..382`, under a ribbon that still ends at 356.

- **It is pre-existing, not a regression from this slice.** The icons slice recorded exactly this at
  `956ab7c` — "the `+`'s computed centre landed 38px high, inside the ribbon's `NetUIHWND`" — blamed
  it on the icon code, bisected it, and left "cause never diagnosed". No product code has changed
  since.
- **It has a reproducible trigger, found today:** `look` → `scroll` → `title` fails; `scroll` →
  `title` passes; `title` alone passes 36 of 36. So it needs `check-look` to have run first.
- **It does not reproduce from geometry alone.** A probe opening six documents at the same 900×700
  window, with the `+` at the identical position 38px from the strip's right edge, answers
  `WordTabStrip` at that pixel. Nor does a fresh Word straight after `check-look`: strip flush,
  registry restored, theme restored.
- **There is no user-facing repro yet.** The trigger is a test ordering, and it would be overclaiming
  to say a user would meet this by hand. That is the first thing the next slice should settle.

**What this slice contributed to it is the whole point of the slice.** Under the old harness this
was a silent mis-aimed click reported three assertions later as "the + opened a new document — FAIL",
with nothing in the add-in's log, and it consumed a bisect that reached the wrong conclusion. Under
the new one the suite says:

```
the +: (2372,464) is over "NetUIHWND", not WordTabStrip - taking the foreground back and re-measuring
  the foreground is "Meeting - Wordsmith notes.docx - Word" and its strip is at (1536,432 874x64);
  the point is 32px below its top
```

An undiagnosed mystery became a reproducible trigger and an exact symptom in one run.

## Cost

| | before | after |
|---|---|---|
| checks | 409 | **440** |
| battery | 19.1 min | **13.6 min** (−29%) |
| `Start-Sleep` in the suites | 254 calls, 8.7 min | **180 calls, 5.3 min** |
| suites green | 11 of 11 | **9 of 11** — the two reds are the defect above |

The 31 extra checks are confirmations that were previously assumptions: every one of them asserts
that an injected input landed, which is a claim the suites used to make silently and sometimes
wrongly.

Per-suite, final run: stack 35, strip 40, tabs 26, menu 37, reorder 37, startscreen 50, look 48,
scroll 49, title 35 (3 failed), dot 61 (1 failed), soak-stack 22.

`check-menu` gained the most (32 → 37) because its Save-As branch — the one that only runs when Word
really does ask where to save a never-saved document — now asserts that each menu-item click landed
on the menu. That branch is also where the `Format-Rect` regression above was hiding.

## Files

- `tools\WordTabHarness.ps1` — new, the shared primitives
- `tools\probe-dialogs.ps1` — new, the measurement behind the classifier. Re-run it before changing
  `Get-WordWindowKind`, and on any machine where the harness starts seeing dialogs that are not there
- all eleven suites — dot-source the module, local duplicates deleted, inputs confirmed

## Deployment, checked the same day

`uninstall.ps1` → `install.ps1` round trip run after all of the above and clean: the add-in
registration, the CLSID, the ProgId, `%LOCALAPPDATA%\Programs\WordTab`, `%LOCALAPPDATA%\WordTab` and
the settings key all removed, then reinstalled from clean HEAD with the activation smoke test
passing. That is the whole no-admin deployment story in two commands and it still holds.

## For the next slice

**The strip-placement defect above is the brief.** What is already in hand: a reproducible trigger
(`look` → `scroll` → `title`), an exact symptom (38 physical px too high, overlapping `NetUIHWND`,
the `+` unclickable), two states measured side by side, and the knowledge that it is not the geometry
alone and not a fresh Word after `check-look`. What is not in hand: the root cause, and whether a
user can reach it by hand — that second question decides how urgent it is and should be answered
first, because the answer may be "not reachable, so this is a test-only artefact".

Also still open from this slice, deliberately not done:

- `Get-Parts` returning `$null` for a missing strip is now possible, so the `if (-not $parts.Strip)`
  guards in eight suites are reachable for the first time. None of them has ever been exercised.
- 5.3 minutes of `Start-Sleep` remain. The rule above says which of them may go: the ones waiting for
  Word, not the ones between an action and its assertion. Most of what is left is the second kind.
- `check-scroll`'s `Use-Wheel` still injects `mouse_event` through a local P/Invoke rather than
  through `WordLayout`. The pointer move in front of it is confirmed now; the notches are not.
