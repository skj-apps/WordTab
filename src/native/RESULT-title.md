# The tab name is the document, and Close Tabs to the Right

**PASSED.** 35 checks in the new `tools\check-title.ps1`; the other nine suites re-run — **348 checks,
0 failures across the ten.** Writeup for the slice that made a tab read the document's name instead
of Word's chrome, and added the one context-menu command that was missing.

| suite | checks | | suite | checks |
|---|---|---|---|---|
| stack | 34 | | look | 43 |
| strip | 40 | | scroll | 46 |
| tabs | 22 | | **title** | **35** |
| menu | 32 | | soak | 22 |
| reorder | 30 | | startscreen | 44 |

Two things ship here:

- **A tab is named after the document.** `Quarterly report.doc` rather than
  `Quarterly report.doc  -  Compatibility Mode`, and the same for `Read-Only` and `Protected View`.
  **And a document whose own name contains `" - Word"` keeps all of it** — `Meeting - Wordsmith
  notes.docx` used to draw as `Meeting`, which was a much worse defect than any annotation.
- **Close Tabs to the Right**, on the tab menu between Close Others and Close All, greyed when there
  is nothing to the right.

---

## The measurement came first, and it corrected two things I believed

`tools\probe-titles.ps1` authors a fixture in each state Word annotates — with Word itself, so the
`.doc` and `.docx` are real files in real formats — and prints the resulting `OpusApp` title as text
**and as codepoints**. The codepoints matter: a separator that is an en dash or a non-breaking space
is invisible in the first form and fatal to a string match.

| file | title Word writes |
|---|---|
| `Quarterly report.docx` | `Quarterly report.docx - Word` |
| `Quarterly report.doc` | `Quarterly report.doc␣␣-␣␣Compatibility Mode - Word` |
| `Quarterly report.rtf` | `Quarterly report.rtf␣␣-␣␣Compatibility Mode - Word` |
| `Locked report.docx` (read-only attribute) | `Locked report.docx␣␣-␣␣Read-Only - Word` |
| `Downloaded report.docx` (mark-of-the-web) | `Downloaded report.docx␣␣-␣␣Protected View - Word` |
| `Meeting - Wordsmith notes.docx` | `Meeting - Wordsmith notes.docx - Word` |
| never saved | `Document1 - Word` |

Two beliefs recorded in the project notes were wrong, and the probe is the only reason that was
found before code was written against them:

- **The annotations are not bracketed.** The notes said `[Read-Only]`, `[Group]`, `[Shared]`. On this
  build they are dash-delimited. Code written against the bracket form would have done nothing at
  all and every assertion about it would have been written to match.
- **Word shows the extension for `.docx` too.** I had assumed it omitted it. It does not.

## The rule: two separators of different shapes

The one fact that makes this cheap:

> **The application suffix is single-spaced (`␣-␣Word`). The state annotations are double-spaced
> (`␣␣-␣␣X`).**

Two different separators, so an annotation can be recognised **by its shape rather than by an
English word list**. That matters more than it looks: `Compatibility Mode`, `Read-Only` and
`Protected View` are all localised strings, and whatever Word adds next is in no list that could be
written today. A table of English annotations would be a table that quietly stops working on a
French Word and on the next Word update alike.

So `WordTabFrameTitle` is: strip the application suffix from the end, then strip trailing
`␣␣-␣␣X` groups from the end.

**Bounded at three groups**, not `while (TRUE)`. `a  -  b  -  c.docx` is a legal filename, and an
unbounded rule would leave such a document with a one-word tab. The guard is what stops a filename
that happens to look like an annotation being consumed to nothing.

## From the end, not the first match — the defect that was worth more than the feature

The old rule was `wcsstr(out, L" - Word")`, which finds the **first** occurrence. A document named
`Meeting - Wordsmith notes.docx` therefore drew as `Meeting`. Two thirds of the name gone, silently,
on a perfectly ordinary filename — worse than anything the annotations cost, and it had been there
since the strip was written.

Everything now comes off the **end**. `" - Microsoft Word"` is tried before `" - Word"`, longest
first, because the short one matches inside the long one; that is for Word 2010 and earlier and is
unmeasured on this build, which is why it is one extra comparison rather than a mechanism.

## The extension decision dissolved into a measurement

This slice was offered with a decision attached: show the file extension on the tab, or drop it?
Neither, as it turns out — it was never ours:

> **Word follows Explorer's "hide extensions for known file types".** The same document reads
> `Quarterly report.docx - Word` with `HideFileExt=0` and `Quarterly report - Word` with it set to 1.
> Measured, by flipping it for real and restoring it in a `finally`.

So the user has already answered this question, to Windows, and Word already listens. Stripping an
extension here would override an answer they have given; and on a machine with extensions hidden
there would be nothing to strip anyway. **The tab takes exactly what Word gives it.**

This is the second time in this project a product decision has turned out to be a measurement
(`UI Theme` in the look slice was the first). It is worth reaching for the probe before the taste.

## What is deliberately not handled

**The bracketed form** (`[Read-Only]`) that older Word builds use. It could not be provoked on this
build, so any code for it would be unmeasured — and a trailing `[...]` rule would eat the tail of a
document genuinely named `Report [final].docx`. Recorded here rather than guessed at.

## `TabTitleTrim`, and the half of the rule that is not on it

New switch, `HKCU\Software\WordTab\TabTitleTrim`, default 1. Off, the tab reads what Word's title bar
reads minus the application suffix.

**The truncation fix is not on the switch.** Removing the suffix from the end rather than the first
match was a defect, and a setting that restored it would exist only to put the bug back — the same
argument the start-screen slice used for having no switch at all. What *is* on the switch is the
annotation trimming, because that half is a policy: somebody may want to see `Read-Only` on the tab,
and `a  -  b.docx` is the filename where the rule is wrong.

---

## Close Tabs to the Right

Genuinely half a slice, because the order model, the batch-close queue and the menu all existed.

**`g_members` order *is* the tab order** — that is the order model the reorder slice chose precisely
so that nothing could disagree with the array. So "to the right of this tab" is "later in the array",
and `CloseBatchStart` grew one parameter:

```c
static void CloseBatchStart(HWND keep, HWND after, const wchar_t* what)
```

`after == NULL` means the whole row, and with NULL the function does what it did before, statement
for statement — which is why `check-menu`'s 32 checks needed no thought. There is no index to
compute and none to keep consistent.

Two details that are not obvious:

- **The active tab goes last, but only if it is in the range.** Close Others and Close All always
  contain the active tab; Close Tabs to the Right must not drag the document the user is looking at
  into a batch that was never about it.
- **`from` is passed as both `keep` and `after`.** The range alone already excludes it. Passing it as
  `keep` too means a wrong answer from the range lookup still cannot close the tab the user pointed
  at.

**Greyed on the last tab, not hidden.** An item that comes and goes changes the menu's shape
depending on which tab was right-clicked, and moves everything below it — and `Close All` moving
under the pointer is the one thing on that menu worth being careful about.

`StackTabsRightOf` answers 0 both for the last tab and for a window that is not a tab in a stack,
because to the caller those mean the same thing. It returns a **count** rather than a boolean so the
log line can say how many, and a batch that claims "2 tab(s) queued" against a row the user read as
three is how a miscount gets noticed.

---

## How you measure a string that is only ever pixels

A tab name is drawn. It is not in a control, so there is nothing to read from outside the process.

The add-in logs the computed name (`strip hwnd=0x… tab name |X| from window title |Y|`) and the
suite asserts the exact string against that line. **But a log line saying what was *computed* is not
evidence of what was *drawn***, and this project has already been burned twice by measurements that
could not say what they measured. So it is paired with a photograph:

- the rightmost column of text pixels inside the tab is measured with the trim on, and again with
  `TabTitleTrim=0`
- trimmed: text runs to **x=410**, well short of its limit — so the trimmed name is not an ellipsis
- untrimmed: text runs to **x=553 of a 556 limit** — it fills the tab to the ellipsis

The log says *which string*; the pixels say *the string reached the screen*. Neither on its own would
do.

**The scan stops where the renderer stops drawing** — `close.left - 4` logical px, the same number
`DrawOneTab` uses. The first version scanned to the tab's right edge, ran onto the close button, and
answered with the x of the **button** for both strings: two different names, one identical
measurement, and an assertion that could never have failed for the right reason.

---

## What the suite cost, and the product bug it found

Three faults, all in the check script rather than the add-in — and the third was not.

**1. `.Count` on an empty return.** Already recorded in the project notes and it cost a run anyway:
`(Get-WordPids).Count` is a hard error under `Set-StrictMode` when nothing is running.

**2. A transient window read as a save prompt.** A window Word puts up while shutting down was taken
for "the user has unsaved work", and a run in which every check had passed aborted on its last step.
The guard now confirms twice, a second and a half apart: *a question being asked of the user is still
there when you look again.* This is [[transient-state-must-be-an-event]] from the other side — a
guard that can only sample must at least refuse to believe one frame of it.

**3. `New-Item -Path <regkey> -Force` recreates an existing key and destroys every value in it.**
The registry twin of `New-Item -ItemType File -Force` truncating a file.

That third one is the interesting one, because **the same line was in `install\install.ps1` (line
165), where it had been since the installer was written.** Every re-install silently reset every
switch the user had set — `TabScroll`, `TabStyle`, `TabThemeSample`, `TabDrag`, `ShowLoadBanner` —
back to its default. Fixed: the key is created only when it is not there. The add-in *registration*
key above it keeps its `-Force`, on purpose: every value in it is rewritten on the next two lines.

**How it presented is the part worth keeping.** The suite wiped the key, so `ShowLoadBanner` went
back to its default, so the add-in's own banner came up over Word — and the clicks aimed at the `+`
landed on a `Static` control inside that dialog. The failure printed was **"the + opened a new
document — FAIL"**, three sections and one subsystem away from the cause, with nothing in the
add-in's log at all because the message was never posted. It took a better error message to find it:
the guard now names the dialog rather than the control under the pointer, and it named it
immediately.

Same shape as `check-menu`'s hard-coded click offset, which spent a slice looking like a broken Save.
**When a suite fails, read the add-in log before believing the assertion, and check what is actually
under the click point.**

`Invoke-StripClick` is the durable answer: it re-measures the point on every attempt and refuses to
click unless `WordTabStrip` is what is under it. `SetForegroundWindow` is refused from a process that
is not already foreground, so a window that comes over Word does not merely steal the keystrokes —
**it receives the clicks**, at screen coordinates that were perfectly correct for Word. This run saw
`Program Manager` take the foreground mid-suite and recover from it on its own.

---

## Also new

**`tools\check-all.ps1`** runs every suite in turn and prints one table. They cannot run in parallel —
they all drive the same Word and the same desktop — and the runner reports each suite's own count
rather than merging them, because a single total hides which suite went quiet.

It cost two lessons of its own on its first outing, both about a runner's failure modes being more
dangerous than a suite's:

- **`-Only` matched nothing and it reported "All 0 suites green."** Running zero suites is the same
  *shape* as passing, which is exactly the answer a test runner must never be able to give. It now
  prints what it is about to run and exits 1 when nothing matches.
- **The summary regex was anchored and matched no suite at all**, so every row read
  "(no summary line)" while every suite passed. The suites do not agree on how they announce a
  total: some print `34 checks, all passed.` at column 0, others `PASS  40 checks, 0 failures`.

**Running the whole battery back to back is itself a hazard**, and the first full run showed it in two
different ways. Neither was a regression; both were the suites failing to get the input they asked
for onto the window they meant.

**`check-reorder`** printed `the foreground was "Terminal" - taking it back` immediately before the
one assertion it failed, with the row unchanged — the drag had gone somewhere else. It passes 30/30
on its own.

**`check-look`** was the more interesting one, because it failed *reproducibly*, which is normally
how a real regression behaves. Two hover assertions failed with "0 pixels changed", and the
"hovered" tab sampled `RGB(15,15,15)` — the well colour, 100% of the scan. That reads as a hover
highlight that is not being drawn.

It wasn't. **`SendInput`'s absolute mouse move had silently not taken.** A diagnostic line printing
where the pointer actually ended up settled it in one run:

```
aimed at (258,540); pointer is at (1894,459) over "SysListView32"; tab 0 is (172,511 259x58)
```

The pointer was never on the tab — it was on the desktop. Both photographs were therefore of the
same pointer position, so of course nothing changed between them. **A hover that never happened and a
hover that is not drawn produce byte-identical evidence**, and the assertion could only report the
second.

`Set-Pointer` is the fix: move, read `GetCursorPos` back, retry, and **assert that the pointer
arrived** as a check in its own right. The suite can now only fail the hover assertions for a reason
that is actually about hover. Same shape as `Invoke-StripClick` above, and the same rule underneath —
*an input you did not confirm landed is not an input.*

It earned its keep on the first run, and the log shows the whole story in four lines:

```
hovering tab 0: asked for (258,540), the pointer is at (1894,459) - trying again
PASS  the pointer could be put on tab 0
aimed at (258,540); pointer is at (258,540) over "WordTabStrip"; tab 0 is (172,511 259x58)
PASS  hovering tab 0 visibly changes it (3624 pixels)
```

The move needed a second attempt; with the pointer actually on the tab, hover changes 3624 pixels and
the hovered tab samples `RGB(28,28,28)` — between the well at 15 and the card at 41, exactly as the
look slice designed it. **check-look is 43 checks now** (41 plus the two that assert the pointer
arrived) **and 0 failures.**

The durable answer for the rest of the battery is the same, but retrofitting it into every suite is
a slice of its own. Until then: give the battery the desktop.

## Files

- `src\native\strip.cpp` — `WordTabFrameTitle`, the `TabTitleTrim` switch, `CMD_CLOSE_RIGHT`, the
  menu item, and the tab-name log line in `ReadTitle`
- `src\native\stack.cpp` — `CloseBatchStart`'s range parameter, `StackCloseToRight`,
  `StackTabsRightOf`
- `src\native\wordtab.h` — the two new declarations
- `install\install.ps1` — the settings key is no longer recreated on install
- `tools\probe-titles.ps1` — the measurement, kept as evidence
- `tools\check-title.ps1` — 35 checks
- `tools\check-all.ps1` — the runner
- `tools\WordLayout.cs` — `BroadcastSettingChange`, for the `HideFileExt` probe
- `tools\check-menu.ps1`, `tools\check-look.ps1` — both assert the menu's contents, so both learned
  the new item
