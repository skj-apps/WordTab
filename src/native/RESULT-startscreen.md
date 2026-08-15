# A Word window with no document is not a tab

**PASSED 2026-08-15.** `pwsh -File tools\check-startscreen.ps1` — **44 checks, 0 failures**, plus the
six existing suites re-run: `check-stack` 34, `check-strip` 40, `check-tabs` 22, `check-menu` 32,
`check-reorder` 30, `soak-stack` 22. **224 checks, 0 failures across the seven.**

This slice was queued as "Start-screen policy: decide what a Start-screen window should *do*, not
just that it is excluded". The first measurement showed it was not excluded at all.

---

## The premise was wrong, and it had been wrong for four slices

`stack.cpp` carried this comment, and `wordtab.h` repeated it:

> A Start-screen window — Word launched with no document — has no `_WwF` and is deliberately left
> out: it has no document to be a tab for.

Every word of the second half was true as a statement of intent. The first half was false.

**`_WwF` is Word's document *frame*, and Word keeps it for the life of the window.** Closing the last
document destroys the `_WwB` and `_WwG` *inside* it and leaves `_WwF` standing, empty. So
`StripHasDocumentFrame` — which asked exactly the question its name says, "is there a `_WwF`" —
answered `TRUE` for a window with nothing open in it. That window joined the stack and was given a
tab labelled **"Word"**, with a close button and everything: a tab for no document.

Measured on 16.0.20228, dev rig, with the shipped build before this slice:

| state | `_WwF` | inside `_WwF` | title | what WordTab did |
|---|---|---|---|---|
| launched with no file (Start screen) | yes | *nothing* | `Word` | joined the stack, one tab: "Word" |
| any document, print layout | yes | `_WwB` → `_WwG` | the document | one tab, correct |
| read mode / web layout / draft / outline | yes | `_WwB` → `_WwG` | the document | correct |
| Backstage open over a document | yes | `_WwB` → `_WwG` | the document | correct |
| minimised document | yes | `_WwB` → `_WwG` | the document | correct, and it must stay a member |
| **last document closed with Ctrl+W** | **yes** | ***nothing*** | `Word` | **joined, one tab: "Word"** |

The photograph of that last row is what the slice is really about: greyed-out ribbon, no document, a
black void — and our tab row across the top of it holding one tab called "Word" with a × and a +.

`_WwF` holds something in **every** document state measured and holds **nothing** in both no-document
states. That is a clean predicate, and it is the one now used.

---

## The policy

**The tab row lists documents. A Word window with no document open is not a tab.**

It is not a stack member: it keeps its own taskbar button and Alt+Tab entry, and the stack neither
moves it nor hides it nor lets it speak for the others' layout. Its strip **stays**, and draws an
empty row with only the **+** on it. The moment a document appears in that window it joins the row
normally.

Three things decided rather than fallen into:

**The strip stays rather than being torn down.** The alternative — no document, no strip — makes an
empty Word look exactly like stock Word, which is tidier on paper. It was rejected because the band
is already carved out of `_WwF`, and giving it back means giving it back *again* the moment a
document arrives, so the document would visibly jump down 32 logical pixels at the instant it
appeared. Keeping the strip costs an empty band and buys a transition with no movement in it at all.

**The + stays, and it is the one button still worth pressing.** Measured: `Documents.Add` on a Word
with no document open puts the new document **into that same window** rather than opening another.
So + on an empty row is a round trip back to one tab, in the window already in front of the user,
with no stray window left behind. An empty row with a + is also an honest statement: nothing is open,
here is how to open something.

**Nothing was done about the Start screen specifically.** It needs nothing. Word reuses that window
for whatever the user picks — measured for Escape (blank document), for the + , and for a file opened
from Explorer while the Start screen is up, all of which keep the same `HWND` and leave no second
window. And while the Start screen is up our row is behind Word's own `FullpageUIHost`, which covers
the whole client area exactly as Backstage does, so there is nothing of ours on screen to get wrong.
`check-startscreen` asserts that occlusion rather than assuming it.

---

## What changed

**`StripHasDocumentFrame` → `StripHasDocument`** (`strip.cpp`). Renamed, because the old name is the
reason the bug existed — it described what the code did, and what the code did was not what the
caller needed. It now answers `TRUE` only if the frame's `_WwF` has at least one child window.

`GetWindow(GW_CHILD)` rather than searching for `_WwB` by name, on two grounds. It is one call
instead of a recursive `EnumChildWindows` on a half-second timer — this predicate runs twice per
tick per window, so the janitor got *cheaper*. And "the document frame is empty" is a weaker
assumption about Word's internals than any particular class name inside it: a renamed child would
silently break a name test and cannot break this one.

**`StackTabs` returns 0 for a document-less frame** (`stack.cpp`). Its `!joined` branch used to
return 1 — the frame itself — which is what actually drew the "Word" tab. The new test is placed
*ahead* of the stacking test, because this is a fact about the document rather than about the stack,
so it holds with `Stack` switched off too.

**The `Leave` reason is now accurate.** `EligibleToStay` can fail two ways, and the janitor logged
both as "hidden or minimised". A visible frame that fails it has lost its document, and the log now
says `left the stack (no document open)`. `check-startscreen` asserts on that string, so it is a
tested claim rather than a comment.

No new switch. Every other slice added one; this one is a correction to a predicate, and a switch
that restored a tab labelled "Word" would exist only to reproduce the bug.

---

## Four hazards the audit found, and the one it found that is not mine

Before writing any of this, five agents read `strip.cpp`, `stack.cpp`, `frames.cpp`, `taskbar.cpp`
and the harness for everything that touches a document-less frame; then a second pass put 27 claimed
hazards through adversarial verification with instructions to refute. **Twenty-one were refuted.**
Six survived, and they were worth the pass.

**1. A stale menu command could destroy the empty window.** `WM_WORDTAB_CMD` is posted, never run
inside `TrackPopupMenu`'s modal loop — and the janitor keeps ticking inside that loop. So a frame can
lose its last document between the click and the command arriving, and `IsWindow` does not catch it
because Word leaves the window standing. `StackCloseTab`'s not-in-a-stack branch would then post
`WM_CLOSE` to it, on a comment that said *"Still a document, still closable"* — the exact premise
this slice falsifies. Since a surviving empty frame is Word's last window, that ends the Word
session. **Fixed:** `StackCloseTab` drops any close aimed at a window with no document. `Close All`
routes through the same branch, so it is covered by the same guard.

**2. A drag could freeze mid-gesture.** The drag state is global and the strip holding the mouse
capture is usually *not* the strip on screen, so the existing "the tab being carried has gone" test
does not catch the strip's *own* window losing its last document. `DragMove` then returned silently
on a zero-length row, every strip kept drawing the carried tab from the last position computed, and
the release committed it there. This was dead defensive code before this slice and is reachable
because of it. **Fixed:** it now cancels the drag, which is what "the row you were carrying this in
no longer exists" means.

**3. `MatchTo` would copy a minimised window's rectangle.** Windows parks a minimised window near
-32000. A window joining a minimised stack was matched to it, which puts the window the user is
looking at off-screen — and `Present()` then moves the taskbar button onto it. `check-stack`'s own
`StackOnFramePosChanging` already refuses that exact rectangle, with a comment saying why; there was
no counterpart in `MatchTo`. **Measured, and the obvious route does not reach it:** open a document
while the stack is minimised and Word restores the stack *first*, so the master is no longer iconic
by the time `Join` runs. But this slice creates a route Word has no reason to interrupt — the
document-less window is now the one window deliberately left on screen when the stack goes down, and
it carries a + that invites exactly this. **Guarded** rather than designed around: `MatchTo` leaves
the window where Word put it and says so in the log. That is not a complete answer to "a document
arrived while the stack was down"; it is the answer to "never make a window unreachable", which is
the rule that may not be broken while the better answer is worked out.

**4. The + could be drawn lit under nothing.** The hover flag is written on mouse movement, but the
+ *moves* when the tab count changes — and the count changes with the pointer sitting perfectly
still. Emptying the row moves it a whole tab width, so the strip lit a chip at the left edge a
tab-width from the pointer. Pre-existing and cosmetic; this slice made it maximal. **Fixed** by
checking the flag against where the button actually is, which is the same principle as the strip
measuring its own position rather than trusting where it last put itself.

**5. Not mine, confirmed, and worse than the known gap says it is.** When enough documents are open
that tabs hit the 70-logical-pixel minimum, the row runs past the +'s reserved space and `after` is
clamped — laying the + rectangle **on top of the last tab**. `HitTestStrip` tests every tab before it
tests the +, and `DrawStrip` paints the + last, over the tabs. So the user sees a + and clicks a tab.
At 200% on a 1280px-wide window, **nine tabs** is enough for the + to lie entirely inside the last
tab, with its glyph centre inside that tab's *close button* — clicking the drawn + closes a document.
Right-click and middle-click use the same hit test. Independent of everything in this slice: with
zero tabs there is nothing to clamp against. **Recorded, not fixed** — it belongs with the "no
scrolling tab row" gap, and it is the strongest argument yet for doing that one.

---

## The harness

**`check-startscreen.ps1`, 44 checks.** How it proves the row is empty is worth stating, because it
is not a pixel comparison. The +'s position is a function of the tab count: with no tabs it sits at
the left padding, with one tab it sits past that tab — 444 physical pixels apart on this rig. So the
script computes both, asserts they differ, and clicks the *no-tabs* one. If the row is really empty
that is the +, and a document appears in this very window; if the row still had the "Word" tab, the
same click lands on the tab and nothing happens. The reverse is asserted too: clicking where the
"Word" tab used to be must not close Word, open anything, or conjure a document.

Two measurements the sequence depends on, both of which cost a run to find:

- **Word replaces a pristine, unmodified blank document** rather than opening a second window beside
  it. So the second document has to come from the + button; opening a file leaves one window and
  quietly measures nothing.
- **On the cold Start screen the row is not clickable**, because Word's `FullpageUIHost` is over it.
  The suite asserts that and gets out of the Start screen the way a user would.

**Two changes to `WordLayout.cs`:** `CtrlPress` (Ctrl+W is Word's *document* close, and the only way
to reach this state — `WM_CLOSE` closes the *window*, and on the last window it closes Word), and
`FirstChild`, which mirrors `StripHasDocument`.

**`check-menu.ps1` started this slice failing 3 of 32, and none of it was the add-in.** Reproduced
identically against `1ffdd8b` with every change stashed. Two separate causes, and the second one is
the more useful lesson.

**The first was the `SetForegroundWindow` trap**, already recorded twice on this project. Run from a
terminal that holds the desktop, a click on a tab switches the document *inside* Word without Word
coming forward, and one reading of the tab row came back naming **"Terminal"** as a document.
`check-stack` and `check-reorder` were hardened against this long ago; `check-menu` never was. It has
the same `Set-WordForeground` handshake now, and the Save menu item is **clicked** rather than
activated by its access key — a popup menu is topmost, so a click at its rectangle lands on it
whatever owns the keyboard.

**The second was a hard-coded offset, and it had nothing to do with the desktop at all.**
`Set-Dirty` types into a document to modify it, and aimed at `strip.Bottom + 300` — which assumes a
large Word window. Word remembers its window size between sessions, and by the end of this slice's
runs it had come back small enough that the point was *below the window*. The click landed on the
**desktop**: `WindowAt` reported `SysListView32` under it and the foreground as `Program Manager`.

Follow that through, because the failure it produces is three steps from its cause. Nothing is typed,
so the document is never modified; Word's `Document.Save` on an unmodified document correctly writes
nothing; and the assertion that fails is **"Save wrote the document to disk"** — which reads exactly
like a broken Save command in the add-in. The add-in's own log said otherwise throughout:
`Document.Save returned (Word confirmed the window handle)`. **When a suite fails, read the add-in
log before believing the assertion.** `Set-Dirty` now aims at the `_WwF` rectangle and sizes
everything from it, so it works at any window size.

**And it proves the edit rather than assuming it.** It photographs the page, types, photographs
again, and retries. Two bounds, both load-bearing: the change must be at least ~100 sampled pixels
(a caret blink is a couple of pixels wide where five characters are not) *and* at most half the band,
with a control band below it that must be untouched. Without the upper bound the check reported a
**fully occluded window** — every sampled pixel different — as a large successful edit, which is a
false green rather than a missed failure. When it does detect occlusion it minimises the offending
window, except the shell's own `Progman`/`WorkerW`: minimising the desktop is "Show Desktop", which
would take Word down with it.

---

## What to preserve

- **`_WwF` is not a document.** It is the frame a document goes in, and Word keeps it empty. Anything
  that wants to know whether there is a document must look *inside* it.
- **A minimised document is still a document.** The stack keeps minimised windows as members on
  purpose — when the whole stack goes down together, dropping every window would leave nothing to
  bring back. Measured: `_WwF` keeps `_WwB` while minimised, so the predicate survives it, and
  `check-startscreen` asserts that so a future change cannot break it quietly.
- **Every mechanism that makes a window harder to reach has an inverse that cannot be skipped.**
  `Leave` restores the taskbar button and clears `WS_EX_TOOLWINDOW` first and unconditionally, which
  is what makes "it stops being a tab" safe to do to a window the user is looking at.
- **A window with no document is left exactly where it is.** `Leave` is called with
  `restorePosition = FALSE`: they closed a document, not a window, and moving it would be a change
  they did not ask for.
