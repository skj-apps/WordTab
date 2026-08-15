# Tab affordances: hover, close, middle-click, and a button for a new document

**PASSED 2026-08-15.** `pwsh -File tools\check-tabs.ps1` — **22 checks, 0 failures**, plus the three
existing suites re-run unchanged: `check-stack` 33, `check-strip` 40, `soak-stack` 22. **117 checks,
0 failures across the four.**

Before this slice a tab was a rectangle you could click. Now every tab carries a close button, the
tab under the pointer lights up, a middle click closes, and a `+` at the end of the row makes a new
document that joins the stack by itself.

This is the first slice of product rather than plumbing, and it is the first one where the
interesting decisions are about what happens to a *document* rather than to a window.

---

## What was built

**One layout function.** `ComputeLayout` in `strip.cpp` returns every clickable rectangle in the
strip — each tab, each tab's close button, and the `+` — in one structure. Painting draws what it
returns and hit-testing reads what it returns, so a click cannot land somewhere other than what the
user is looking at. It replaced `TabRect`, which computed one tab at a time and was already being
mirrored, slightly wrongly, by the check scripts.

**Hover as state, not as an effect.** `hotKind`/`hotFrame` on each strip, set from `WM_MOUSEMOVE`,
cleared by a `WM_MOUSELEAVE` armed through `TrackMouseEvent`. Repainting happens only when what is
under the pointer actually changes — `WM_MOUSEMOVE` arrives on every pixel of movement, and
invalidating on each one would repaint the strip continuously while the user crosses it on the way
to the ribbon.

**Press and release, on the same button.** A close button that fired on the press would mean one
mis-aimed click destroys a document. Both buttons capture the mouse on the way down and act only if
the release lands on the same target; `WM_CAPTURECHANGED` cancels anything half-pressed. The check
asserts this directly by pressing a close button and letting go on the tab beside it.

**Frames, not indices, all the way through.** The tab row can change between a press and its
release — that is exactly what closing a tab does — so an index that meant one document on the way
down can mean a different one on the way up. Every stored target is an `HWND`.

**Double buffering.** Hover means the strip now repaints whenever the pointer crosses a tab
boundary. Painting straight to the screen showed the background fill before the tabs landed on it,
which at that rate is a flicker directly under the pointer.

---

## The decisions worth keeping

### Closing a tab activates it first, and that is not cosmetic

`StackCloseTab` posts `WM_CLOSE` — the exact path Word takes when the user clicks its own close
button, so the save prompt, any document-close macros and Word's own bookkeeping all behave
identically and none of it is reimplemented here.

But it activates the tab first. **Every window in the stack sits at the same rectangle with the
active one on top, so a modal save prompt belonging to a window underneath can end up behind the
window in front of it** — and an invisible modal dialog is, from the user's side, Word beeping and
refusing to respond with nothing on screen to explain why. Activating first makes that class of
failure impossible.

**Confirmed by hand, since no script should sit waiting on a modal dialog.** A document was made
dirty, left as a background tab, and closed from its tab: Word's `NUIDialog` came up in front, named
the right file, and the "Don't Save" that followed left the user back on the document they had been
reading. Screenshot in the session; the mechanism is the `NUIDialog` being foreground at
(1057,251 690x532) inside the stack rect.

### ...and it hands the user back afterwards

Activating a tab in order to close it moves the user off whatever they were reading. `g_returnTo` in
`stack.cpp` is set before that switch and consumed in `Leave` when the active slot is vacated, so
closing a background tab leaves you exactly where you were rather than on whatever happens to follow
the tab that closed.

It is checked when the active slot is vacated rather than keyed to the frame we posted `WM_CLOSE`
to, and that is deliberate: **frame lifetime is not document lifetime**, and closing a document was
measured to hide one window and destroy a different one. What matters is that the active slot is
being vacated, not which window vacated it.

### A button with no room is a button that closes the wrong document

A tab narrower than three close buttons gets no close button at all, and the name wins. The `+`'s
width is taken out of the available space *before* the tabs are sized, so tabs shrink as documents
open and the button does not; past the point where tabs fill the strip it pins to the right edge.
Either way it stays inside the strip and clickable, which is the only property it has to have.

### The one thing taken from Word's object model

`WordTabNewDocument` in `connect.cpp` calls `Application.Documents.Add()`, late-bound. It is the
only use of the object model in the whole add-in, and it exists because making a document is the one
thing windows cannot do. Everything else WordTab does, it does to windows — and it has to, because
the object model knows nothing about the stack.

The new document then arrives as a new `OpusApp` frame, which the CBT hook sees, the strip binds and
the janitor joins. **Nothing places it and nothing makes it a tab**: it becomes one by the same route
every other document does. That is why the feature is three calls rather than a subsystem.

It is called from a posted message rather than from the click handler, because `Documents.Add`
creates a window and pumps messages while it does, and some of those messages come back to the same
window procedure.

---

## What the checking found, and it was all in the checks

Nothing in the add-in was wrong on the first run. Every failure was a defect in how it was being
measured, and two of them are worth remembering.

**The strip moves, so a rectangle goes stale in about a second.** The first version of
`check-tabs.ps1` measured the tab row once and clicked four tabs from it. Three of the four clicks
did nothing, and the add-in log showed they never arrived — because a relayout between clicks had
shifted the strip **46 pixels** and the later clicks were landing in the document. A click that
lands nowhere looks exactly like a feature that does not work.

The fix is structural rather than a longer sleep: `Get-Spot` recomputes from the strip that is on
screen at that instant, and every click and every hover goes through it. `check-stack.ps1` had the
same latent bug and now recomputes inside its loop too.

**Killing Word poisons the next run.** A killed Word offers to recover its documents on the next
launch, which adds a window nobody asked for and a recovery pane over the document. One run was
measured against "Document7 — AutoRecovered" without noticing. `check-tabs.ps1` now closes any
running Word gracefully *before* it starts and refuses to continue if it will not close, and it never
kills Word on the way out. This is the rule that was already written down, enforced in the script
rather than left to whoever is driving.

**`CloseMainWindow` closes one window, not one Word.** A process holding a stack has N top-level
windows; two rounds of `CloseMainWindow` left four documents open. Cleanup now closes frames one at a
time.

**The assertion that was wrong, not the code:** "the window behind that tab is gone" tested
`IsWindow`. It failed, correctly — Word had recycled the frame. The claim that matters is that the
document is out of the tab row, which is what it asserts now.

---

## Measured

- Tab layout at 144 DPI, 5 documents in a 1210px strip: tabs 160px wide, close buttons 24×24, `+`
  39×30 — all five tabs carrying a button, the `+` inside the strip.
- Hovering the close button of tab 1: **576 pixels changed inside the button** (every sampled pixel),
  **0 pixels changed on tab 0**. The hover is visible and it is local.
- `Documents.Add` to the new document being a tab in the stack: within one settle cycle, no
  intervention.
- Press-on-button, release-on-tab: 6 windows before, 6 after.

## Switches

`HKCU\Software\WordTab\TabButtons` (DWORD, default 1). Set to 0 and the tabs are exactly what they
were before this slice — no close buttons, no `+`, hover still live. Bisectable like every other
piece.

## What this slice does not do

No right-click menu, no drag to reorder, and the tabs still look like the placeholder rectangles
they have always been — the visual design is deliberately last, so it is done once against the
finished set of elements rather than four times. A close button on a tab whose document is dirty
gives you Word's own prompt, which is correct but is Word's dialog rather than ours.
