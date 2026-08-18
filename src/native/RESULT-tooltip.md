# Pointing at a tab says what it could not fit

**Status: passed 2026-08-17.** Rest the pointer on a tab and, half a second later, a small panel
below it gives the document's name in full and, underneath, the folder the document is in.

`tools\check-title.ps1` 39 → 90 checks. `tools\WordTabHarness.ps1 -SelfTest` 23 → 30. **No new suite
and no new probe** — the fixtures this needed were already open in the suite that owns tab names.

New switch `TabTip` (DWORD, default 1) at `HKCU\Software\WordTab`.

---

## Why, and it is a measurement rather than a nicety

The icons slice ended with a finding that every slice since has depended on: **the name is the only
thing on a tab that identifies its document.** Every Word document gets byte-identical artwork from
the shell, so an icon carries no information; the name carries all of it.

And the name is the first thing the row spends when it runs short of room. Word's own default window
fits five tabs, measured in the scroll slice. A sixth document shortens every name in the row, and it
stays shortened. Past that point the strip is `Qua...`, `Me...`, `Draf...` — and there is nowhere
else in the window to look, because the thing the tabs replaced was Word's window list.

**The folder is the second line because it is the thing no tab width could ever show.** Two documents
called `Report.docx` from two different folders draw two identical tabs at any size. The name answers
"which document is this"; the folder answers "which of these two".

## A tooltip with nothing to add does not appear

If the name fitted *and* there is no folder to name, everything the panel would say is already on the
screen, and showing it anyway is half a second of the user's attention spent on a repetition. So it
stays away.

Both halves of that test are honest answers rather than guesses, and they come from different places:

- **"the name was cut" is recorded by the code that cut it.** Not recomputed from the tab rectangle
  somewhere else — the text box is derived from the close button's position, and a second derivation
  of one rectangle is the shape of defect this project keeps finding.
- **"there is no folder" is Word's own answer**, and it is distinguished from "nobody could ask" (see
  below).

Measured, one document, at the tab's full 220 logical px: `Document1` produces no panel and the log
says why — `nothing to add (name fits; no folder - never saved)`. The same tab in a seven-document
row produces one, because the name no longer fits. And a *saved* document on a wide tab produces one
too, because the folder is still something the row cannot say. Both reasons drive independently, and
the suite drives both.

## One place decides whether a name fitted

Both renderers used to draw the name themselves — same string, same `DT_END_ELLIPSIS`, two
`DrawTextW` calls eighty lines apart. That was tolerable while the answer was only ever pixels, and
it stopped being tolerable the moment something else needed to know whether the name had been cut.

`DrawTabName` is now the one place a tab's name reaches a device context, and it records the cut
where the cutting happens. `DT_CALCRECT` measures what the string wanted; the box is what it got.
Measuring is the only honest test — the alternative is reading pixels back looking for an ellipsis
glyph, which cannot tell an ellipsis Windows added from three dots the user typed.

## The two failure shapes, kept apart again

`WordTabReadDocumentPath` follows the rule `WordTabReadModified` established, and it matters here for
a reason that shows on screen:

- **FALSE — Word would not answer at all.** The panel shows the name alone.
- **TRUE with an empty string — Word answered and there is no folder.** Also the name alone, but it
  is a fact about the document rather than a failure to read one.

**Protected View turns out to be the second of those, not the first**, and that was measured rather
than assumed. Such a document is in none of this Application's collections, so the walk over
`Application.Windows` finds no window claiming that frame and falls off the end having answered.
Measured on the rig: `Downloaded report.docx` gets a panel **302x56 with one line**, against
**578x92 with two** for every ordinary document beside it in the same row. The height is the
assertion, because a folder line drawn empty would leave it unchanged.

## Read on demand, and that is a deliberate difference from the dot

The dot is polled twice a second because the thing it reports changes while nobody is looking. A
document's folder changes only when the user does Save As, and the only moment anything needs to know
it is the moment a panel is about to appear. Polling for it would be a `BSTR` allocated per window
per tick to answer a question nobody had asked. **The cost of asking is paid by the hover that
asked**, and `TabTip=0` stops the asking entirely.

`Document.Path`, not `FullName`: the name is already on the first line, and repeating it in the
second would spend the width that makes the folder readable.

## The re-entrancy this created, and it would not have shown up as a red suite

**A call into Word's object model pumps messages.** That is the rule this add-in already follows
everywhere it posts a command to itself rather than running one inside a window procedure — and
`TipShow` reads the folder from Word *in the middle of deciding whether to show a panel*.

A janitor tick dispatched inside that pump reaches `StripRefreshTabs`, which calls `TipStop`, which
clears exactly the two globals that record that a tooltip is up. Without re-asserting them after the
call, the panel is shown with nothing recording that it exists, and the next `TipStop` finds nothing
to hide: **a tooltip that stays on the screen until the pointer happens to arm another one.**

Two lines, and they are the interesting two lines in the file. Reasoned rather than caught — but
reasoned from a rule this project had already written down at cost, which is the argument for writing
them down.

## A window of our own, not a system tooltip

A system tooltip is painted in the *system's* colours, and Word's theme is not Windows' theme — this
rig runs a dark Word on a light Windows. The panel uses `menuBack`, `menuText`, `menuTextDim` and
`menuLine`: **the same four floating-surface colours the context menu already derives**, because a
tooltip and a menu are the same kind of object — a small panel over the document, a step above Word's
chrome rather than below it. No new palette entry was needed, which is the check that the abstraction
was the right one.

Other decisions in the drawing:

- **`DT_PATH_ELLIPSIS` on the folder line, `DT_END_ELLIPSIS` on the name.** A path cut at the end
  loses the folder the document is actually in, which is the only part anybody reads; cut in the
  middle it keeps both ends, which is what every file dialog in Windows does.
- **Under the tab, not under the pointer.** A panel that points at the thing it is about needs no
  arrow, and the tab is a fixed target where the pointer is not. Measured: panel left 319, tab left
  319.
- **Clamped to the monitor, not to the frame.** A Word window can hang off the right-hand edge of a
  screen, and a panel clamped to the window would then be drawn off it. If it will not fit below, it
  goes above.
- **`CS_DROPSHADOW`**, so the one floating surface here without a shadow is not this one.
- **`WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE | WS_EX_TRANSPARENT`, owned by the frame.** An owned popup
  is above its owner without being topmost over every other application, and an owned tool window
  gets neither a taskbar button nor an Alt+Tab entry. **This project spent a slice making Word one
  taskbar button; a panel that appears on a hover is not allowed to add a second**, and the suite
  asserts the style rather than trusting it. `WS_EX_TRANSPARENT` means the pointer reaches the
  document underneath, so moving onto the panel is moving *off* the strip — which is what takes it
  away.

## The delays are the system's, not mine

Windows defines its own tooltip timings in terms of the double-click time: a tooltip appears after
`GetDoubleClickTime()` and takes itself away after ten times that. WordTab reads the same value, so
its panel arrives when the user's other tooltips arrive — on a machine whose owner may have changed
that setting for a reason. The suite reads it through `[WordLayout]::DoubleClickTime()` rather than
hardcoding 500ms, or it would be asserting against the default of a setting instead of the product.

**The early return on "already this tab" is what makes the delay mean what it says.** `WM_MOUSEMOVE`
arrives on every pixel; re-arming on each one gives a tooltip that never appears unless the hand is
perfectly still for half a second, which is a different gesture and notably hard to perform on a
trackpad.

## One place takes it down

`StripRefreshTabs` calls `TipStop`. Every caller of that function is a real change — a title, a dot,
the stack, the scroll position — so the row's own repaint path is the one place that has to know
about the tooltip, rather than each of them. A tab that scrolls out from under its own panel, or
closes while it is up, would otherwise leave it naming a document that is no longer there for as long
as the auto-hide takes.

Buttons are handled the same way: one test at the top of `StripWndProc` covers all three, so a fourth
cannot be added without it. Measured leaving-latency: **5ms**, against an auto-hide of 5000ms — which
is the evidence that it goes on `WM_MOUSELEAVE` and not on the timer. A test that only asserted "it
is gone eventually" would pass on the timer alone.

---

# What it cost outside the feature

## The harness called our own window a question, and that would have poisoned every suite

**Found on the first drive, and it is the reason this section exists.** The tooltip is the add-in's
first *top-level* window — the strip is a child and so was never enumerated — and at 917x92 it is
comfortably past the harness's size floor. `Get-WordWindowKind` is a denylist of Word's own chrome
classes, so it classified the panel as a **question**.

Every dialog guard in the harness stops a run when Word is asking something, including the one
`Close-AllWord` uses to decide whether it may keep closing windows. **The first hover in any suite
would have reported a save prompt that does not exist**, and the diagnosis would have started at
Word.

The fix is a third category rather than a longer denylist: `ours`. "Word drew this and it is not
asking anything" and "WordTab drew this" are different facts, and folding ours into a list whose
comment says every entry was measured coming out of Word would have been a quiet lie about where the
list came from. `WordTabStrip` is in it too, even though it is a child today — the answer must not
depend on that staying true.

**It is now in the harness self-test**, which is where it belongs: `Get-WordWindowKind` is a pure
function of a class name and a size, so it runs with no Word and no desktop. Seven checks, including
the two prompts that must *stay* questions — a fix that made the classifier permissive would
otherwise pass.

## The suite passed alone and failed in the battery, and the reason is a rule worth keeping

**90 of 90 standalone; 81 of 84 inside the battery, three failures, and not one of them was the
product.** Every behaviour in that run was correct.

The difference was **how wide Word was**. Standalone, this suite inherited a 1911px window and seven
tabs fitted. Inside the battery it inherited whatever the previous suite left — **900px, because
`check-stack` resizes Word and Word persists its window placement on exit.** At that width seven tabs
go to their 70 logical px minimum, the row overflows and scrolls, and only the last five are on
screen. So:

- two documents were off the left-hand end, and the assertions naming them failed about tabs nobody
  could have hovered;
- the computed slots for tabs 5 and 6 landed on the button cluster past the end of a shorter row —
  **the same shape as `Get-Spot` assuming the row is as long as the window count**, which the
  tear-off slice had already recorded;
- and `Document1`'s name **was** cut at that width, so the panel was right to appear and the suite
  was wrong to say it should not.

The last of those is the interesting one. **"The name fits" is a fact about the width of the row, and
the width of the row is a fact about the window** — so an assertion that a particular document
produces no tooltip is an assertion about how wide Word happened to be when the suite ran. It was
written as though it were about the product.

Three changes came out of it, and the second is the general one:

1. **The width is now a fixture rather than an inheritance.** The section widens Word to 760 logical
   px if it is narrower, and puts it back in a `finally` — the same rule, for the same reason, as the
   maximized tear-off section in `check-reorder`. A section that needs a property of the environment
   should establish it, not hope for it.
2. **The suppression case moved to where the tab width is deterministic**: two documents in one
   window, which is the widest arrangement this suite can produce that still contains a never-saved
   document. `Document1` fits on a half-width tab in any window Word will open.
3. **Every saved fixture is now asserted, not one of them.** Checking a single tooltip would pass on
   a build that could only ever describe the tab it was pointed at.

And the diagnosis cost one `grep` and no machine time, because `check-all` keeps each suite's
transcript beside its log. That mechanism has now paid for itself on three consecutive batteries.

## Three more, all mine, all found by running it

1. **`@($lines | Select-Object -First 1)` returns the string `"System.Object[]"`.** The array wrapper
   survives the assignment, so the tooltip's `Name` compared unequal to every name there is and read
   in the transcript like a panel drawing garbage. The `@()` habit that correctly guards `.Count`
   elsewhere in the harness is wrong on a scalar — **the fourth shape of this trap in these notes.**
2. **The hide logged the wrong window.** `TipStop` named the frame whose strip was carrying the
   panel; `TipShow` named the frame the panel was *about*. Those are routinely different, because
   every window in the stack draws every tab and the panel belongs to whichever window is in front.
   Two meanings of `hwnd=` in one log family is a log that has to be decoded rather than read.
3. **`install.ps1` turns the load banner back on**, and a dialog over Word breaks any suite that
   clicks. Nothing to do with this slice; noted because the rig was left with it off and two installs
   in one session put it back both times.

## A shared type widened rather than a second one added

`[WordLayout]::TitleOf` read 256 characters. The tooltip carries its whole text on the window itself
— a document name and a full path with a newline between them — which is past 256 for any document a
few folders deep. Widened to 1024, which is a strict superset: no caller can tell the difference on a
string that already fitted. The alternative was a second reader, and the rule about shared types
being supersets was written down after exactly that mistake.

## Why the section went in the middle of `check-title`

Two slices running have found a real defect by putting a new section upstream of the rest of its
suite, and none has been found by a section that ran last. This one went in before Close Tabs to the
Right, so everything below it keeps using the same Word.

It also cost no fixtures at all. `check-title` already opens six saved documents from one folder, one
of them in Protected View, and then adds a never-saved one through the `+` — which is exactly the
three cases the tooltip has to tell apart, in one row, at a width narrow enough that names are cut.

**Every tab is hovered and the answers are matched by name afterwards**, rather than assuming which
tab is which: the row order is the open order today, and an assertion that depends on that is an
assertion about the stack's member array, which is `check-reorder`'s subject.

---

## Still open

- **Nothing tells you about the `+` or the chevrons.** They are unlabelled glyphs and a tooltip is
  the standard way to name them. Deliberately not in this slice: it is a different question (naming a
  control) from this one (saying what did not fit), and it would put three more English strings in
  the add-in.
- **The `FALSE` branch of `WordTabReadDocumentPath` is reasoned, not driven.** Protected View turned
  out to be the determinate-empty case, so "Word would not answer at all" needs a Word that has no
  Application object or refuses the collection — which nothing in the suite can provoke on purpose.
  The log line distinguishes the two, and the panel is the same either way.
- **One font for both lines.** A smaller one for the folder would read better and costs another font
  object per strip; not obviously worth it, and not measured.
- **The auto-hide is not asserted.** The suite asserts it goes on the pointer leaving, and asserts
  that it went in far less time than the auto-hide would take — which is the check that matters. That
  the panel *would* eventually go on its own is untested; it is one `SetTimer` and it would cost the
  suite five seconds of standing still.

---

## 2026-08-17: the intermittent, found and measured

`check-title`'s `title` suite went red once in a battery: **one tab of seven produced no tooltip in
4020ms and NO LOG LINE AT ALL**, and then never did it again in 208 further attempts. The reading
recorded at the time - "so `TipShow` was never reached" - was wrong, and the way it was wrong is the
useful part: **`TipShow` had four exits that returned without writing anything**, so a tooltip
abandoned before it was shown and a hover that never happened produced identical evidence. "No panel
and nothing logged" meant three different things and named none of them.

### The mechanism

`StripRefreshTabs` called `TipStop`, and `TipStop` does two jobs: it takes down a panel that is UP,
and it cancels a timer that is merely PENDING. Only the first is what a row change wants.

A pointer resting on a tab sends no further `WM_MOUSEMOVE`. So nothing re-arms it. **Any change to
the row inside that half second - a dot coming on in another document, a title Word rewrote, a window
activating - meant the tooltip never appeared at all, for as long as the hand stayed still.** Not
delayed: never. And silently, because a tooltip that was never shown has nothing to log about being
hidden.

Nothing about a pending arm can be made stale by a row change, which is what makes the cancel pure
loss. `TipShow` re-reads everything when it fires: it hit-tests where the pointer is *now*, reads the
tab's name *now*, and asks Word for the folder *now*.

### The measurement

`tools\probe-tip.ps1`, three documents, the pointer parked off the row before each hover. `-Churn`
flips `Document.Saved` on document 1 - **never the document being hovered** - immediately before the
pointer arrives, so the janitor's next 500ms poll refreshes the row somewhere inside the window where
the tooltip is being waited for.

| build | quiet | with churn |
|---|---|---|
| before | 12 of 12 tooltips | **0 of 12**, and **0 log lines** |
| after  | 12 of 12 tooltips | 12 of 12 tooltips |

The 0-of-12 with 0 log lines is the reported failure reproduced exactly, including its silence. The
quiet column is the control that says the churn is what did it.

### What changed

- **`TipRowChanged`**, which is what `StripRefreshTabs` now calls. It takes down a panel that is
  visible and leaves a pending arm alone.
- **Every exit in `TipShow` says so.** Four were silent.
- **A pointer that is on a different tab when the timer fires restarts the wait there** instead of
  dropping it. Same defect in a second dress: half a second is long enough for the row to have
  scrolled under a still hand, and a still hand sends no `WM_MOUSEMOVE` to start the clock again. It
  terminates - the next fire hit-tests the same point against the same layout and matches.

### What this does not explain

The `title` failure it reproduces is the same shape, but the archived log for that run shows **no
`dot`, `tab name`, `stack` or `relayout` line anywhere in the 6.5 seconds** the tab was being
hovered. Almost every caller of `StripRefreshTabs` logs before calling it; the two that can reach it
silently are a scroll step inside the 250ms log throttle and the drag re-layout in `DragMove`, and
neither was happening - nothing was being dragged, and the suite widens Word specifically so the
seven tabs do not scroll.

So the mechanism is measured but **the trigger on that particular day is still not named**, and it
may be a fourth thing rather than a refresh at all: a `WM_MOUSEMOVE` that never reached the strip
would look identical, and this project has already been bitten twice by injected input that did not
take. That is what the logging half is for. **If `title` goes red again the log now says which of
these it was** - a restart on another tab, an abandonment with its reason, or still nothing at all,
which would rule the whole `TipShow` family out by elimination and point at the arm never happening.
