# Ctrl+Tab: the research, before the feature

**2026-08-16. Queue item 5, the research half. NO PRODUCT CODE CHANGED** — `src\native` is
untouched, this is `tools\` only, and the installed DLL is still `8DA909D3C0A1570A84A0B14368C9B2FE`,
byte-identical to the one in `dist\WordTab-20260816-4abea68.zip`.

Instrument: `tools\probe-keyboard.ps1`. Eight measurements, three oracles each.

## Why this was a research slice and not a feature slice

The queue entry said the obstacle was that "the strip is `WS_EX_NOACTIVATE` and never has focus, so
keystrokes never arrive". That is true of the **strip** and it is not the question. The add-in lives
inside Word's own process and **already installs a thread-local `WH_CBT` hook on Word's UI thread**
(`frames.cpp:656`), with the `hMod = NULL` detail right and the unhook balanced in `FramesStop`; it
has logged `cbtHook=installed` on every run this project has ever made. So the mechanism was never
the hard part, and the three things that actually decide the shape of this feature had never been
measured.

All three came back differently from the obvious guess. Two of them change what gets built.

## 1. Word gives the keyboard to `_WwG`, which we do not subclass

`GetGUIThreadInfo` on Word's UI thread, with three documents up and a real caret in the page:

    hwndActive  0x8310BE  |OpusApp|      <- the frame, subclassed (frames.cpp)
    hwndFocus   0x9608EA  |_WwG|         <- the document pane, NOT subclassed
    _WwF        0xBE0864                 <- subclassed (strip.cpp)
    strip       0x7109E8                 <- WS_EX_NOACTIVATE

A `WM_KEYDOWN` is delivered to the focus window and to nothing else — **keyboard messages do not
travel up to parents the way `WM_CONTEXTMENU` does.** So neither of the two windows this add-in
subclasses can ever see a keystroke, however many more we subclass, and **a thread-level hook is the
only route**. `WH_GETMESSAGE` or `WH_KEYBOARD` on `g_uiThread`, alongside the `WH_CBT` that is
already there.

## 2. Ctrl+F6 ALREADY switches documents, and the tab row already follows

Measured, three presses, three documents:

    Charlie.rtf  ->  Bravo.rtf  ->  Alpha.rtf  ->  Charlie.rtf

and the add-in said `stack  active -> 0x…` matching the new foreground handle **every one of the
three times**. Ctrl+Shift+F6 cycles the other way. This works because `frames.cpp:344` turns
`WM_ACTIVATE` into `StackOnFrameActivate` — WordTab correctly tracks a switch **it did not
initiate**.

**So keyboard document switching is not absent today. It is unbound and undiscoverable.** That is a
user-visible fact worth telling people about right now, at zero cost: with WordTab installed, Ctrl+F6
moves between documents and the tabs keep up.

## 3. Word's window order is MOST-RECENTLY-USED, not the tab order — and this is the finding that changes the build

**Sections 2 and 3 cannot tell those two apart, and I nearly wrote down the wrong conclusion because
of it.** `Charlie -> Bravo -> Alpha` is exactly what "one tab to the left, wrapping" produces *and*
exactly what "the next most recently used window" produces, because the documents were opened in that
order and so were last touched in that order. Two hypotheses, one dataset, and the convenient one
would have been written up as fact.

The discriminator is to make the two orders disagree before pressing anything: activate Bravo, then
Alpha. Tab order is still Alpha, Bravo, Charlie — so "one left of Alpha, wrapping" is **Charlie** —
while "next most recent" is **Bravo**.

    before   Alpha.rtf   (Bravo most recently used before it)
    Ctrl+F6
    after    Bravo.rtf

**Bravo. Word's list is MRU.**

The consequence is not cosmetic. **Ctrl+Tab cannot be implemented by handing Word its own command**,
because "the tab to the right of this one" is not something Word's command can express — it will go
somewhere that depends on what the user did five minutes ago. The add-in has to compute the target
itself from `g_members`, which *is* the tab order (`RESULT-reorder.md`: the array is the order,
deliberately, with no index to keep consistent) and activate it with `StackActivate`.

That is also the answer to a question this project keeps asking: a tab strip's job is to make the
document order visible and stable, and Word's own keyboard command is tied to something else
entirely. **The two orders are different by design, not by accident.**

## 4. What binding Ctrl+Tab actually costs, and it is confined to tables

Word does own Ctrl+Tab: pressed in the document body it put a tab character in and `Document.Saved`
went `True -> False`. But the cost is not what that alone suggests.

**In ordinary body text, plain Tab inserts a tab as well** — measured, same transition. So there
Ctrl+Tab is a *second* route to something the user already has a first route to, and taking it costs
nothing anyone would notice.

**Inside a table they genuinely diverge**, and this is the whole of the real cost:

    Tab      in a cell:   caret 0 -> 1,  document saved=True    (moved to the next cell, no edit)
    Ctrl+Tab in a cell:                  document saved=False   (inserted a literal tab)

Inside a table, **Ctrl+Tab is the only way to put a tab character in a cell**. Nowhere else does it
do anything Tab does not.

Ctrl+Shift+Tab in the body does nothing Word cares about at all — it neither switched document nor
edited it. **It was NOT measured inside a table** and Shift+Tab is "previous cell" there, so treat
the reverse chord's cost as unmeasured rather than zero.

## The brief for the feature slice

1. A `WH_GETMESSAGE` or `WH_KEYBOARD` hook on `g_uiThread`, next to the existing `WH_CBT`. Same
   `hMod = NULL` rule, same balanced unhook in `FramesStop`, same `g_lockCount` discipline.
2. The chord must be **swallowed**, not observed and passed on, or Word puts a tab in the document.
3. The target is computed from `g_members` and activated with `StackActivate` — never delegated to
   Word's own command, because of §3.
4. Direction: Ctrl+Tab to the tab on the **right**, Ctrl+Shift+Tab to the **left**, wrapping. Note
   that this is the *opposite* of Word's own Ctrl+F6, which moves left; matching Word here would put
   WordTab at odds with every other tabbed application, and the tab row is the thing the user is
   looking at.
5. A switch, `TabKeys`, default 1 — and unlike most of this project's switches, `0` here is a real
   answer rather than a way of putting a bug back: it gives the user Word's literal-tab-in-a-cell
   behaviour untouched.

**The decision this slice deliberately does not make, because it is a product decision and it is
named rather than resolved silently:** whether Ctrl+Tab is swallowed *everywhere*, or everywhere
*except inside a table*. The second is possible — the add-in can ask `Selection.Information(12)`
(`wdWithInTable`), and the object model already costs ~0.3ms steady in the dot poll, which is
affordable on a keystroke — but it makes the chord's behaviour depend on where the caret is, which
is a thing users have to learn. My recommendation is to swallow it everywhere and let `TabKeys=0` be
the escape hatch, because a chord that works in three places out of four is worse than one that
works or doesn't.

## Near-misses worth more than the successes

- **The two-hypotheses trap in §3 is the important one.** A clean, repeatable, three-document cycle
  looked exactly like proof of a model it was not evidence for at all. What separated them was not
  more data of the same kind — it was *arranging for the two answers to differ before pressing the
  key*. A sequence that both hypotheses predict is not a measurement of either.
- **7b's guard caught my own setup bug, and that is the only reason this document is not wrong.**
  The first run built the table at `Selection.Range`, which was left mid-paragraph by section 7a, so
  Ctrl+Home landed in the text *above* the table. The probe asserts `Selection.Information(12)`
  *before* measuring anything and stopped the run with "the caret is not in a table cell" rather
  than measuring the body and reporting it as a cell. **A setup step needs its own assertion; "the
  keystroke did nothing" and "the fixture was never built" are the same evidence.**
- **`[WordLayout]::NativeOm` is why the object model was usable at all.** `New-Object -ComObject
  Word.Application` binds to a different instance and leaves a stray WINWORD behind — the QAT slice
  lost time to that. Asking one *named* window through `OBJID_NATIVEOM` has now been used to read
  `Document.Saved`, read `Selection.Range.Start`, ask `Selection.Information`, and author a table.
- **`Document.Saved` is a settable property and this probe sets it, which the add-in may never do.**
  The rule protects the *user's* documents from the *add-in*. A probe discarding a change it made
  itself, sixty seconds earlier, to a file it authored itself, in `%TEMP%`, in a directory it is
  about to delete, is a different act — and the alternative is leaving a modal question on the
  user's screen after the run. `Close-AllWord` correctly refuses to answer questions, so something
  had to give and this is the honest place for it. Nothing in `src\native` does this or may.
- **Two additions to shared files, both strict supersets**: `[WordLayout]::ThreadGui` and
  `CtrlShiftPress`, and `-Shift` on `Invoke-ConfirmedKey`. The chord went into the *harness* rather
  than into the probe so that keystroke confirmation keeps having exactly one implementation —
  `-Shift` without `-Ctrl` throws rather than quietly sending an unshifted key. Gated with
  `WordTabHarness.ps1 -SelfTest`: 23 checks, 0 failures.
