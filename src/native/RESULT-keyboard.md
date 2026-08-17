# Ctrl+Tab: the research, then the feature

**Queue item 5, done in two halves on 2026-08-16. Both halves are in this one file on purpose** —
what is known about Word and the keyboard should have one place to be read, not a research note and
a build note that agree until one of them is corrected.

- **The research half** (below, commit `b10a375`) changed **no product code**: `tools\` only, with
  the installed DLL still `8DA909D3C0A1570A84A0B14368C9B2FE`, byte-identical to the one in
  `dist\WordTab-20260816-4abea68.zip`. Instrument `tools\probe-keyboard.ps1`, eight measurements,
  three oracles each.
- **The feature half** is at the end of this file, under "The build". Ctrl+Tab and Ctrl+Shift+Tab
  now move along the tab row.

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

---

# The build

**2026-08-16, the same evening. Ctrl+Tab moves one tab to the right, Ctrl+Shift+Tab one to the
left, both wrapping.** The brief above was followed as written and nothing in it turned out to be
wrong, which is the return on having measured first.

**No new suite and no new probe.** The regression test is a section inside `tools\check-stack.ps1`,
which already has three documents up and already knows the tab order. `probe-keyboard.ps1` stays as
the measurement.

## What was built

Three edits, ~90 lines of product code.

- **`frames.cpp`: a `WH_GETMESSAGE` hook on `g_uiThread`,** next to the `WH_CBT` that has always been
  there. `GetMsgProc` rewrites a matching `WM_KEYDOWN` to `WM_NULL` and posts
  `WM_WORDTAB_SWITCH_TAB` to the coordinator window. Installed in `FramesStart` only when `TabKeys`
  is on, unhooked in `FramesStop` alongside the CBT hook.
- **`stack.cpp`: `StackNeighbourTab(frame, delta)`,** the tab `delta` positions along the row,
  wrapping; `NULL` when there is nowhere to go.
- **`CoordinatorProc` gains one case**, which recomputes the target and calls `StackActivate`.

## The five decisions that are not obvious from the code

**1. `WH_GETMESSAGE`, not `WH_KEYBOARD`, and the reason is the swallow.** In a `GetMessage` hook the
`MSG` is ours for the length of the call and rewriting it to `WM_NULL` is total. It is also the same
move the strip already makes on `WM_WINDOWPOSCHANGING` — change what is being *proposed* rather than
correct what has already happened — so it is an idiom this codebase already has rather than a new
one. The decisive part is *when* it runs: before the message loop's `TranslateMessage`. Swallow any
later and the synthesised `WM_CHAR` of `0x09` is still on its way, and Word types a tab into the
document with no keydown left to explain it.

**2. The hook computes nothing and activates nothing; it posts.** `CbtProc` has always worked this
way and the reason transfers exactly: this code runs *inside* `GetMessage`, and calling
`SetForegroundWindow` from in there re-enters message retrieval on the thread that is retrieving.
The post is pumped a moment later, at a point in the loop where switching a window is an ordinary
thing to do.

**And what is posted is the source and the direction, never the target.** Slice 4 lost a bug to a
posted command whose target had gone by the time it arrived, so the tab to move to is worked out at
the moment it is used. Same message, same mistake available, not made twice.

**3. The chord is swallowed on membership, not on "is there somewhere to go" — and that is the
product decision this slice makes.** The test is `StackTabIndex(root) >= 0`: did this key go to a
window that is a tab in our row. A Word with one document open therefore swallows Ctrl+Tab and does
nothing, rather than typing a tab character that it would not type one document later. The
alternative makes what the key *means* depend on how many documents happen to be open, which is the
same objection the research half raised against swallowing everywhere-except-in-a-table.

It also lands correctly in three places without any code for them, which is the sign the test is the
right one: **focus inside a dialog** roots at the dialog rather than at an `OpusApp`, so Ctrl+Tab
still moves between the pages of a property sheet; **a Word window with no document** has already
left the stack; and **`Stack=0`** means there is no row, so the key is Word's again.

**4. Auto-repeat is swallowed but not acted on, and this one is a judgement, not a measurement.**
Bit 30 of the keydown's `lParam` is the previous key state. A held chord repeats at up to ~30/sec and
each switch here is a real window activation plus a relayout of every window in the stack, so
honouring repeats would thrash Word; a row small enough to fit on screen is one nobody needs to hold
a key to cross. Tapping Tab with Ctrl held still works, because each tap is a fresh keydown.
**Swallowing the repeats is not optional** — passing them through would put a run of tab characters
into the document, which is worse than either answer. **Not driven by the suite**: injected input
does not carry the auto-repeat bit, so a test of it would be a test of the test. Recorded as
undriven rather than as covered.

**5. `GetKeyState`, not `GetAsyncKeyState`.** The message-synchronised state is the right question —
"which modifiers were down when this key was pressed" — and it is reliably populated here because
Ctrl goes down *first*, so its own `WM_KEYDOWN` has been retrieved and dispatched before Tab arrives.
This was written as the reasoned choice and then settled by the suite rather than by argument; if it
had been wrong, all four chords would have gone straight through into the document.

## How it is proved

**Section "Ctrl+Tab between documents" in `check-stack.ps1`, riding on the existing three-document
fixture.** `$activated` — which window each computed tab slot brought forward — *is* the tab order,
and it is the only way this project can know that order from outside Word.

**The assertion that matters is the direction, not that something switched.** Four presses,
`+1 +1 -1 -1` starting from the last tab, which is where the click loop leaves the run: that crosses
the right-hand end on the first press and the left-hand end on the last, so both wraps are driven
without a step that exists only to set them up. Then the second oracle: the add-in's own
`keys  ctrl+tab  … tab 2 -> … tab 0` lines, one per chord, because the foreground window says which
document came forward and the log says which tab WordTab decided on.

**And the cost that was avoided:** `Document.Saved` per window before and after, asserting nothing
went from saved to modified — with `at least one document was unmodified before the chords` as a
control, because a check that cannot fail is not a check.

**Two more, at the end of the run:** the one-document case (chord swallowed, log says
`nowhere (swallowed; no other tab)`, nothing typed), and **`TabKeys=0` driven positively** — Word
restarted with the switch off, the startup line asserted to read `msgHook=off`, then Ctrl+Tab
asserted to *reach Word and type*. "Nothing happened" is also what a keystroke that never landed
produces, and this project has had three negative tests pass for exactly that reason.

## Near-misses worth more than the successes

- **The test's own MRU comparison was wrong, and it was wrong in the shape this whole feature is
  about.** The first version computed Word's answer as "the tab to the left of the one we are on",
  which is right for the *first* press and wrong for every one after it, because Word's list
  reorders on every switch. Two of the four steps then printed `and not to Word's MRU tab 0` about a
  press whose correct answer was *also* tab 0 — a comparison against a number that is only sometimes
  the other model's answer, reading exactly like a discriminator. It now tracks the list properly and
  **counts the presses that genuinely disagree (3 of 4) and asserts that count**, so the section
  cannot quietly stop being able to tell the two implementations apart. The research half's central
  finding was that a sequence both hypotheses predict is not a measurement of either; the *test for
  it* made the same mistake one level up.
- **The suite now sets `Document.Saved`, which is the second place in the project to do it.** The
  standing rule — the add-in only ever reads it — protects the *user's* documents from the
  *add-in*, and is untouched: nothing in `src\native` does this or may. A suite marking its own
  scratch `.rtf` in `%TEMP%` saved is the act `probe-keyboard.ps1` already justified, and the
  alternative is a modal save prompt left on the user's screen, because `Close-AllWord` correctly
  refuses to answer a question it did not raise. It is also what makes the one-document check *able
  to fail*: the resize earlier in the same run dirties a document all by itself.
- **`TabKeys=0` uninstalls the hook rather than making it inert**, which costs the suite a Word
  restart, because the switch is read once at `FramesStart`. That is the right trade for a hook that
  sees every message Word retrieves: "off" has to mean out of the way. Same shape as `TabDot=0`,
  where the poll returns before it asks.
- **The startup line reports what `SetWindowsHookEx` returned, never the variable that asked for
  it** — `msgHook=installed` / `FAILED` / `off (…TabKeys=0)`. `TabThemeSample` stayed dead for two
  slices because its log line printed the variable, which kept its default.
- **Both outcomes of a chord are logged, including the one where nothing happens.** "The hook never
  fired" and "the hook fired and there was nowhere to go" are different facts and must not share a
  silence. That is what makes the one-document assertion possible at all.

## Still open

- **Ctrl+Shift+Tab inside a table is still unmeasured** — the research half says so and the feature
  did not change it. Shift+Tab is "previous cell" there, so treat the reverse chord's cost as
  unknown rather than zero.
- **Backstage.** Ctrl+Tab with Backstage open switches documents underneath it. Harmless, since the
  strip is hidden behind `FullpageUIHost` anyway, but it is not a considered behaviour.
- **Auto-repeat**, above: decided, not driven.
- **No keyboard reorder and no keyboard scroll.** Unchanged by this slice.
