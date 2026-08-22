# Which tab is highlighted

**The complaint.** From the work rig, on `build-4c4c17a`: *"when opening a second doc it switch to
that but initial tab still highlighted"*.

Two windows, one row. Word opens the second document and puts it on screen - the user is looking at
it - and the tab row goes on drawing the FIRST document's tab as the selected one. Nothing is
broken enough to lose work, and everything about it is wrong: the row is claiming the user is
somewhere they are not.

## What the log said

Their report carries the whole thing three times over. A join that ends correctly looks like this:

```
13:17:27.847  stack  hwnd=0x...A1202  joined, snapped to 0x...B0E22  (2 in the stack)
13:17:27.849  stack  active -> 0x00000000000A1202
```

and one that does not looks like this - same shape, same second, one line missing:

```
13:22:47.576  stack  hwnd=0x...A1202  joined, snapped to 0x...B0E22  (2 in the stack)
13:22:47.592  stack  row, left to right: 0x...B0E22 |...|  0x...A1202 |BOR Procedures Dental|
```

No `active ->`, and none for the next three and a half seconds either - until the user clicked the
close button on the tab, at which point the row said `closing this tab (a background one - will
return to where the user was)`. The row thought the document the user was reading was in the
background. Twice more the same day: `DNOC & CRN Wording` at 09:00:38 waited 3.4s for the user to
click its tab, `F.U.N.` at 09:00:48 waited 13s for them to click into the document.

## Why the row did not know

`Join` never sets the active window - deliberately, because a window joining is not always a window
the user has moved to. What was supposed to catch this is the janitor, twice a second:

```c
// Word activates windows without always telling our frame procedure, so trust the system over
// our own bookkeeping.
HWND foreground = GetForegroundWindow();
```

That line is right about the principle and wrong about the question. **It asks who has the keyboard.
The tab that should be highlighted is the window the user can SEE.** The stack holds every window
at one rectangle, so the document on screen is the front-most of them - focus was only ever a proxy
for that, and Word breaks the proxy: it can raise a newly opened document's window in front without
handing it focus, and on their rig it does.

The palette sampler in the same log agrees, from the other side. At 13:22:49, two seconds into the
wrong state, it reported `palette sampled from Word's ribbon in 0x...A1202` - the new window. It
asks the front window first and falls through to the rest when the front one's ribbon cannot be
read, so its answer says the new window was readable, which means it was on top. The row and the
sampler had two different ideas of which window was in front, and the sampler's was right.

## It does not reproduce on the dev rig, and that is the point

Opening a second document here - `winword.exe path.rtf`, twice - produces this, every time:

```
21:26:30.133  stack  hwnd=0x...2F08BE  joined, snapped to 0x...C30438  (2 in the stack)
21:26:30.134  stack  active -> 0x00000000002F08BE
```

Word raises the new window and gives it the keyboard in the same breath, so both answers agree and
there is nothing to see. A check that opened a second document would have passed on this machine on
the day the bug was reported. So the check drives the MECHANISM instead: `SetWindowPos(HWND_TOP,
SWP_NOACTIVATE)` on a background tab - a window put in front with the keyboard left where it was,
which is exactly the state their Word leaves behind - and then asserts on the row from outside.

## What was built

**`FrontMostJoined` in `stack.cpp`**, and the janitor asks it instead of asking the keyboard.
`GetTopWindow` walks the desktop from the front; the first joined member met is the one the user is
looking at. When it answers, the row goes there.

The focus rule is kept for the case z-order cannot answer: with fewer than two tabs in the row there
is nothing to compare, and the walk is skipped entirely - which is also what keeps the cost off the
idle path, since a one-tab row is most of the time.

## What was checked

`tools\check-stack.ps1`, new section, and it is an A/B rather than an assertion written after the
fact:

| | before | after |
|---|---|---|
| the raise left the keyboard where it was | PASS | PASS |
| the window in front is the one presented to Alt+Tab | **FAIL** | PASS |
| the window behind it gave the presentation up | **FAIL** | PASS |
| the row said the active tab changed | **FAIL** | PASS |

The instrument is `WS_EX_TOOLWINDOW`, not the log: the row presents exactly one window to Alt+Tab
and it is the active one, so which window that is reads `g_active` back through the product. The
first assertion is the precondition - if the raise HAD moved the keyboard, the old rule would have
answered and a pass would mean nothing.

## What the battery found, which was not this

Running all fourteen suites against the change turned up a second thing, in `check-reorder`'s log
rather than in its assertions:

```
21:43:41.766  strip  hwnd=0x...A00946  tab clicked -> 0x...300916
21:43:41.770  stack  active -> 0x0000000000300916
21:43:41.788  ghost  hwnd=0x...A00946  carrying 0x...300916  (170x66, grabbed 51px in)
21:43:41.862  stack  active -> 0x0000000000A00946
```

A tab was clicked, the row went there, and 92ms later it came back. Between the two lines the
carried card appeared - and the card is a top-level window belonging to the window the press landed
on, so showing it put THAT window back in front. The new rule read the z-order honestly and undid
the user's switch.

So the rule does not run while a press is held on a tab: `StripTabPressHeld` in `strip.cpp`, asked
by `FrontMostJoined`. A gesture in progress moves windows for its own reasons, and z-order in the
middle of one is not a statement about which document anybody has chosen. When it declines, the
older focus rule answers - which is the answer that shipped, so a drag behaves exactly as it did.

**This is the second time a rule about "what is true right now" has been caught being asked during a
gesture that was still happening.** The first was the save prompt the janitor could not see; the
shape is the same and worth the name: *a fact read in the middle of a gesture is not a fact.*

**The A/B, which is what settled it.** Three runs of `check-reorder` on this machine, same session,
same screen, minutes apart:

| binary | result |
|---|---|
| the shipped build, rebuilt | 118 / 118 |
| the front-most rule, no guard | **113 / 118** |
| the front-most rule with the guard | 118 / 118 |

Without the middle row this would have been filed as the reorder suite's known flakiness with
injected input, which it looks exactly like: a carried card that does not follow the pointer, and
tabs that report the wrong document. It was not. The suite found a real defect in a change whose
own check was already green.

## What is not covered

The trigger. Nothing here opens a second document and asserts the tab follows, because on this rig
it always does - Word raises and focuses together and the check could not fail. What is asserted is
the rule that was wrong. If their Word ever puts a document in front AND gives it focus, that path
was already working and still is; if it puts one in front without focus, that is now handled.

## The battery, and the one check that stayed red

Fourteen suites on the final binary: **795 checks, 794 passed, 13 suites green**. `check-reorder`
finished 117/118,
failing *"the card follows the pointer, pixel for pixel"*.

That one is not this change, and it is not a hunch - the pre-change binary fails it too. Six runs,
alternating, all this evening:

| binary | the card moved, of 180px |
|---|---|
| before | 180 (pass), 0 (fail), 150 (fail) |
| after  | 180 (pass), 60 (fail), 30 (fail) |

The rig is in an RDP session and the injected `WM_MOUSEMOVE`s are being coalesced before Word sees
them; `GetCursorPos` confirms the pointer moved and proves nothing about what arrived. It also
cannot be this change by construction: while a press is held on a tab `FrontMostJoined` returns
before doing anything, so every code path under that check is the one that shipped.

**The shape of a failure is evidence, and these two shapes differ.** The real regression the
battery caught was 0px AND a spurious tear-off that corrupted the row order - five checks - and it
has not come back since the guard. This one is one check, wanders, and predates the change.
