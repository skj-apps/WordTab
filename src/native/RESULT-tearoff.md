# Taking a tab out of the stack

**Status: both halves passed 2026-08-16.** A document comes out of the stacked window and stands on
its own — by `Move to New Window` on the tab context menu, or by dragging its tab off the strip and
letting go. It lands offset from the stack, with its taskbar button and its Alt+Tab entry back.

Two slices, in this one file on purpose: the mechanism and the gesture are one feature and what is
known about taking a window out of a stack should have one place to be read. **The mechanism half is
first, below; the gesture half is at the end.**

**243 checks, 0 failures** across the five affected suites for the mechanism: menu 63 (was 37), look
48, stack 60, startscreen 50, soak 22. **No new suite and no new probe** — the section went into
`check-menu`, which already opens a stack of documents and already has the machinery to open and read
the menu.

---

## The queue entry's obstacle was the wrong one

The queue had this filed as *"every 'make this window harder to reach' mechanism needs its inverse
driven."* Reading `stack.cpp` before starting said otherwise: **the inverse was already built, and it
already runs.**

`Leave(member, why, restorePosition)` at `stack.cpp:312` puts back every single thing the stack does
to a window — the taskbar button through `PresentWindow`, `WS_EX_TOOLWINDOW`, the window's own
rectangle and zoom state, a `StripRefit` because Word will not lay out a window it is not focused on,
and the active slot handed to somebody else. It was written for shutdown, it fires today at uninstall
and at `Stack=0`, and it is exactly what tearing off needs. There is no second implementation of
"un-stack a window" in this slice.

**The real obstacle is one the queue did not mention, and it is smaller and sharper.** `StackJanitor`
re-derives membership twice a second from what is true about each window, and a torn-off window passes
every test it applies: visible, sized, holding a document. Without something to stop it, the window
snaps back into the stack within half a second and the command simply appears not to work.

So the slice is: a sticky flag, a placement rule, and a menu item.

## What stays out, and for how long — the bug this slice actually had

The decision named at the start was *"a torn-off window stays out until it is closed."* **That was
wrong, and the suite caught it three sections downstream.**

Word does not destroy a frame when its last document closes. It **hides the frame and reuses it for
the next document** — the same fact that once gave an empty Word window a tab called "Word", recorded
in `RESULT-startscreen.md`. So "until the window closes" is far longer than it sounds, and a flag
living on the `HWND` outlived the decision it recorded:

```
tear off tab 5 -> Ctrl+W closes it -> Word hides frame 0x8909B2
press + -> Documents.Add lands in 0x8909B2 -> it refuses to join the row
```

Five Word windows, four tabs, and nothing on screen to explain the difference. Every right-click after
that aimed past the end of the row and got the empty-strip menu instead, which is how it surfaced:
`check-menu`'s Save-never-saved section stopped finding a Save item, and the section after it died on
an empty match.

**The rule now: a torn-off window stops being torn off when it stops holding a document.** Tested on
the *document*, not on visibility — Word hides and re-shows frames of its own accord, and a rule keyed
to that would drop a torn-off window back into the stack while the user was looking at it. The
transient-state trap, avoided by picking the non-transient predicate.

Worth keeping the diagnosis route as well as the answer: **the add-in's own log said it in one line**
(`menu on 0x0000000000000000 -> command 0` — a menu with no target, i.e. a right-click on the empty
strip). No pixels, no bisect, no guessing.

## Where a torn-off window goes

**Not where it is.** Every window in the stack sits at the same rectangle, so a window that leaves and
keeps its position is a window the user cannot see has left — it is exactly on top of, or exactly
underneath, the thing it was pulled out of. *"Nothing happened"* and *"it worked perfectly"* would be
the same picture.

**Not `joinRect` either**, which is what `Leave` restores and is right there for free. That is where
the window was before it *ever* joined: usually wherever Word opened it on a cold start, which may be
the same place as the stack, off the side of the current monitor, or on a monitor that has since been
unplugged. `joinRect` answers *"undo the stacking"*, and this is not an undo.

So: same size, offset by `SM_CYCAPTION + SM_CYSIZEFRAME` — the step the shell itself cascades new
windows by, which puts the title bar of the window underneath in plain view. Measured at 50px on this
rig at 192dpi. Then clamped back onto the monitor work area, because a title bar below the bottom of
the screen cannot be dragged back and the stack is often near the bottom-right already.

**A maximized stack is the case that forced the rule to have a second branch.** Two maximized windows
are pixel-identical and an offset is not even possible, so a torn-off window comes down to three
quarters of the work area, centred, then cascaded. `SetZoomState(frame, FALSE)`, which changes the
state through `SetWindowPlacement` without activating.

## The other decisions

- **`StackCanTearOff()` rather than a registry read in the strip.** The menu asks the stack whether
  tearing off is available; the drag gesture will ask the same function. One copy of the answer.
- **The item is grouped with Save, not with the closes**, and greyed rather than hidden on a single
  tab — the rule already stated in `ShowTabMenu` about the menu keeping one shape so `Close All` never
  moves under the pointer. It is also the only other item on that menu that destroys nothing.
- **`&Move to New Window`** — `M`, which is free; `N` belongs to New Document.
- **Named `StackTearOffTab`, not `StackDetachTab`.** `StackDetachFrame` already exists and means
  something entirely different — a window being destroyed and forgotten. Two names that differ by one
  word, doing opposite things, in the same file, is how the wrong one gets called.
- **New switch `TabTearOff`** (default 1). Off, the item is not offered and the command is refused.
  Unlike most switches here it is not restoring a previous behaviour — nothing could leave the stack
  before — it is there because a torn-off window is the one thing WordTab does that the user cannot
  undo from inside WordTab yet.
- **Everything else falls out for free because the row keys off `joined`.** A torn-off window gets
  `StackTabIndex` -1, `StackNeighbourTab` NULL (so Ctrl+Tab in it correctly does nothing),
  `StackTabsRightOf` 0, `StackMoveTab` FALSE, and `StackTabs` answers 1 — itself, as its own single
  tab. No new code, and it is why the strip in a torn-off window draws one correct tab.

## Three things found on the way that were not this feature

- **The owner-drawn menu's tag array was `MenuItemTag[8]` and the menu was already eight items.** A
  ninth does not overflow — the bound is checked — it silently falls off the owner-draw path and
  Windows draws it in the default style: one light grey row at the bottom of a dark menu. **Every
  text and state assertion in `check-menu` still passes while that is on screen.** Raised to 12, and
  the overflow now logs, because what it produces is a rendering glitch rather than anything that
  names a limit.
- **`@($menu.Items | Where-Object {...})[0]` is a hard error under StrictMode when nothing matches** —
  the third shape of this trap, recorded in `RESULT-harness.md`, with four occurrences still live in
  `check-menu`. One of them killed the run instead of reporting a failure, which is the whole cost:
  the suite stopped three sections before the end and reported nothing about them. All four are now
  `| Select-Object -First 1`, which makes the `if ($item -and ...)` guards below them reachable for
  the first time.
- **`Get-Spot` assumed the tab row is as long as the Word window count.** True for the whole life of
  the project until this slice, and false the moment a window leaves the stack. It does not throw when
  it is wrong: the computed slots are merely narrower than the real ones and the click still lands, so
  it passes for the wrong reason. It now takes an explicit count.

And one thing hoisted: **the expected tab menu, which two suites hardcoded.** `check-menu` asserts the
contents; `check-look` asserts that owner-drawing did not cost the labels their `GetMenuString`
readback. Both listed all eight items, and adding one turned `check-look` red for a reason that had
nothing to do with what it measures. Now `Get-TabMenuItems` in `WordTabHarness.ps1`.

## What the suite asserts, and why in that shape

The section is in `check-menu` between the empty-strip menu and the Save sections, and it opens with
`$Documents = 5` rather than 4 — the fifth belongs to this section, which tears it off and then closes
it, leaving the four windows every section below was written against.

- **The tab is identified by clicking it first.** Which window a slot belongs to is knowable from
  outside Word only by activating it and reading the title, the same oracle `check-stack`'s click loop
  uses. *"Some window left the stack"* is a much weaker claim than *"the one whose tab I
  right-clicked left the stack"*, and only the second one is what the user asked for.
- **Reachability is asserted as a count of presented windows, not as a property of the torn-off one.**
  Exactly two windows are in Alt+Tab and the taskbar afterwards: the torn-off one and the stack's
  active one. Anything less is a window the user cannot reach.
- **The stickiness check waits three janitor ticks and asks again.** Every other assertion in the
  section runs inside the first tick, so all of them would pass against a build with no sticky flag at
  all.
- **The add-in's account is asserted as a count of one**, not as "at least one" — a command that ran
  twice would take two windows out and the geometry assertions would still pass.
- **The greyed case is free**: the torn-off window is now a stack of one, so its own menu is where
  "the last tab has nowhere to go" gets driven, with no extra fixture.
- **`Get-StripPlacement` over all five windows**, because the torn-off window is moved and refitted
  and that is exactly the operation the 38px bug lived in.

## Still open after the mechanism half

- ~~**The gesture.**~~ Built; see below.
- ~~**`TabTearOff=0` has not been driven.**~~ Driven positively by the gesture slice; see below.
- **Putting one back.** There is no way to drag a window into a stack, and a torn-off window rejoins
  only by losing its document. That is deliberate for now — a window that silently rejoined after the
  user pulled it out would be worse — but it is the obvious next question.
- **Tearing off a maximized stack is not driven either** — the three-quarters branch of `PlaceTornOff`
  is reasoned and logged, not measured. `check-menu` and `check-reorder` both run at 900x700.

---

# The gesture: dragging a tab out of the row

**Status: passed 2026-08-16.** Drag a tab a row-height clear of the strip and let go, and that
document comes out of the stack. The pointer changes shape while it is out there, so which of the two
things the drop will do is visible before letting go.

**`check-reorder` 37 → 72 checks.** **No new suite and no new probe** — the gesture is a drag, and the
suite that drives drags already had the fixture, the confirmed-input machinery and the only honest
oracle for the tab order. One new primitive in the shared harness (`CursorShape`/`SystemCursor` in
`WordLayout.cs`) because the feedback is a cursor and nothing could read one.

## One press, two meanings

The drop means *reorder* or *tear off* depending only on where the pointer is when the button comes
up, and it can change its mind as often as the user likes on the way. That is one state flag,
`g_dragTorn`, on top of the drag that already existed.

- **The threshold is a whole row-height clear of the strip, above or below** — `TEAROFF_LOGICAL_SLOP`,
  written as `STRIP_LOGICAL_H` because that is what it means rather than a number that happens to
  equal it. Sixteen times the reorder slop, and it is the one threshold in this file chosen for **the
  cost of being wrong** rather than for how it feels: a reorder aimed at the wrong slot is fixed by
  dragging again, and there is no gesture that puts a torn-off window back, so a tear the user did not
  mean is a window they have to go and find.
- **The way back in is half the way out.** Without hysteresis the boundary is a place a hand can come
  to rest, and a pointer sitting on the line flips the gesture several times a second with the tab
  jumping under the hand each time.
- **Starting the drag is deliberately asymmetric.** Horizontal travel past 4px starts a reorder, as it
  always has; vertical travel starts nothing at all until it is already past the tear-off threshold.
  A hand that slips downwards while clicking a tab must still be a click — but the one gesture with no
  horizontal component whatsoever, pulling a tab straight down, is precisely the one this slice exists
  for, and the old purely-horizontal test could never have seen it. The comment it replaced said in so
  many words *"there is no tear-off in this add-in, so dragging a tab downwards is not a different
  gesture"*.
- **The drop posts `CMD_TEAROFF`**, the same command the menu item raises, through the same
  `WM_WORDTAB_CMD` route. One implementation of "a tab was torn off" rather than one per way of asking
  for it, and that path already drops the command if the window has gone in the meantime.

## The feedback, and what it cost to not have a better one

**A tab dragged out of the row is somewhere the strip cannot draw.** The pointer is over Word's
document, or off the window entirely, and a child window may not paint there. There is no floating
window preview following the hand, the way a browser does it — that would mean moving the real window
live, which is a different and much larger slice.

So the two signals are:

- **The row lets go.** The order goes back to what it was when the tab was picked up, and the tab
  stops following the pointer: it sits still in the slot it came from while the pointer walks away.
  **A gesture that changes what it means has to show something at the moment it changes**, and this is
  the only change the strip can make that is visible whether or not the tab has been moved yet.
  Following the pointer would go on saying "this is being carried in the row", which is the one thing
  that has stopped being true.
- **`IDC_SIZEALL`** — Windows' own "you are relocating this", not a cursor invented here. Set directly
  rather than through `WM_SETCURSOR`, because that message is not sent to anybody while the mouse is
  captured; for the same reason nothing else is going to put the arrow back underneath us.

**The cursor turned out to be a proper test oracle, which was not obvious.** Standard cursors are
shared objects — `LoadCursorW` with a NULL instance hands every process the same handle for the same
one — so a check script can assert *identity* rather than compare bitmaps. Measured before it was
relied on: `IDC_ARROW` is `0x10003` and `IDC_SIZEALL` is `0x10015` in both processes. The assertion is
still written as a **change** first, because "some other cursor" would satisfy the change on its own.

**And it must be read while the button is still down.** The first version of the `TabTearOff=0`
section read it after the release and failed: by then the pointer is over a document and Word has set
its I-beam. The shape belongs to us only for as long as we hold the capture.

## Three things found on the way, and two of them were not this feature

### A new document was arriving in the middle of the row

**A real defect, pre-existing since the mechanism half, and three moves from a user.** Drag a tab out,
close that window, press `+`: the new document arrives where the torn-off one used to be. Measured as
tab 1 of 5.

The cause is one this project has now written down four times. **Word does not destroy a frame when
its last document closes — it hides it and puts the next document straight into it.** The row order
*is* the member array, and a recycled frame still holds its old place in that array, so a document the
user has never seen inherits the position of one that has gone. The add-in's own log said it in one
line: `joined, snapped to ...` for a handle that had been torn off two minutes earlier.

**The fix is narrow on purpose.** A member that is out of the row *and holding no document* is marked
`rejoin`, and a rejoining member is moved to the end of the array before it joins. The obvious wider
rule — "anything that rejoins goes to the end" — is wrong and would have been worse than the bug:
**a minimised window still holds its document**, and a user who minimises Word and restores it must
find the row exactly as they left it, not shuffled. The flag is therefore set from the same fact, on
the same tick, as the one that clears `tornOff`.

Two smaller things fell out of it:

- **`Join` now returns whether it compacted the array**, and the janitor steps its index back when it
  did. It has exactly one caller, which is why that is safe to make its business rather than hiding
  the move behind a deferred queue.
- **The move-to-end is logged.** The one writer in this project that moved something and said nothing
  cost a slice to find (`StripSetNatural`, `RESULT-chrome.md`), and this is the same shape: a silent
  reorder nobody asked for.

### The strip was a second behind the pointer

The suite failed three assertions with the pointer confirmed to be exactly where it was aimed, and the
add-in's log showed the transition arriving **a second late**. The cause was in the new code: while a
tab is out of the row it does not move, and I was calling `StripRefreshTabs()` on every mouse-move
anyway — a full composite of every strip in the stack, sixty times a second, to redraw a tab that was
already where it belonged. Now it repaints only when the tab is not already home, which out there is
at most once.

**The test lesson is the bigger half.** The assertion was `sleep 250ms, then read the log`, and that
is exactly the shape this project has already written a rule against: a fixed wait long enough to be
safe hides the difference between an add-in that reacts instantly and one that is a second behind,
and a wait that is too short fails while naming the wrong thing — *"the tab did not come back into
the row"* when what happened is that the strip had not processed the move yet. It is now a bounded
wait that **asserts the event and prints its latency**: 23ms out, 345ms back, 15ms out again.

The reason it was caught at all is that **the same script was run twice and disagreed with itself**.
One run did the whole round trip; the next never re-entered the row. A single red run would have been
read as a threshold that was wrong.

### `far` is still a macro

`BOOL far = ...` compiles as `BOOL = ...`: `windef.h` has `#define far` for the sake of 16-bit code,
and the compiler reports the error on the line *after* it. Renamed to `travelled`.

## What the suite asserts, and why in that shape

The section is in the **middle** of `check-reorder`, and the suite opens `$Documents = 5` rather than
4 — the fifth belongs to this section, which tears it off and closes it, leaving the four every
section below was written against. That placement is the reason the recycled-frame defect was found:
the `+` section runs downstream of it, and appending the new section at the end would have shipped
the bug.

- **The control comes first.** A tab carried to 0.4 of a row below the strip still reorders. Without
  it, "dragging a tab out of the row tears it off" is satisfied by a build that tears off whenever the
  pointer leaves the strip at all.
- **The threshold is measured off the strip, not copied from the source.** `TEAROFF_LOGICAL_SLOP` is
  `STRIP_LOGICAL_H`, so the strip's own height on screen *is* the threshold in physical pixels at
  whatever DPI the suite is running at — 64px here. A number mirrored from a `#define` goes stale
  quietly.
- **One continuous gesture, out and back and out again**, so both thresholds are driven by the route
  rather than by a second fixture.
- **The pointer is read back after every mid-gesture move.** An injected mouse move can silently not
  take, and "the add-in did not notice the pointer come back" and "the pointer never came back" are
  the same three failing assertions.
- **The row is read only after the torn-off window has been closed.** With five Word windows and four
  tabs, `Get-TabSpot` computes its slots from the window count and is wrong *silently* — the slots are
  merely narrower and the clicks still land. Same trap as `Get-Spot` in the mechanism half.
- **"It tore off INSTEAD of reordering" is a separate assertion.** The tab is carried two places along
  the row before it is taken out of it, so a build that committed both would leave the remaining tabs
  in a different order — and every geometry assertion would still pass.
- **`TabTearOff=0` is driven positively**: Word restarted with the switch off, the startup line
  asserted to read `detach=off`, then the same drag out past the same threshold asserted to **reorder**
  — the input landed, it just meant something else. The cursor is asserted to have stayed an arrow, so
  nothing was offered that would then be refused. Costs one Word restart, because the switch is read
  in `StackStart`.

## Still open after the gesture half

- **No live preview of the window being dragged**, and this is the honest gap. The feedback is the row
  letting go plus a cursor, not a window following the hand. Doing it properly means moving the real
  window during the drag, which is a slice of its own.
- ~~**Putting a window back into a stack**~~ — built; see below.
- ~~**The maximized branch of `PlaceTornOff`**~~ — measured 2026-08-16; see below.

## Tearing off a maximized stack — measured 2026-08-16

The one branch of `PlaceTornOff` that had never been driven, by either half. Everywhere else a
torn-off window keeps its size and is offset by one caption and border — but **two maximized windows
are pixel-identical and an offset is not even possible**, so a maximized stack hands the torn-off
window a window-sized window instead: three quarters of the work area, centred, then cascaded by the
same step.

Measured on this rig, and it is exactly what the code reasoned: work area **2848x1672**, three
quarters of it **2136x1254**, and the torn-off window landed at **(406,259 2136x1254)** — cascaded
**(50,50)** from centred, which is `SM_CYCAPTION + SM_CYSIZEFRAME` here. The windows left behind
stayed maximized, and the add-in's own log names the branch it took.

- **The section is the last one before the Word restart, on purpose.** It changes the shape of every
  window in the fixture, and **Word persists its window placement on exit** — so a suite that left it
  maximized would hand the next suite in the battery a maximized Word and a set of assertions written
  at 900x700. The restore is in a `finally` for the same reason.
- **The work area is read through the same API the add-in uses**, not taken as the screen size. A
  check that compared against the whole screen would be measuring a different rectangle from the one
  the product placed against. New `[WordLayout]::WorkArea`.

---

# Putting one back

**Status: passed 2026-08-16.** A window standing on its own goes back into the stack — drag its lone
tab onto another window's tab row and let go, or right-click it and choose **Move Back to the Tab
Row**. Until this existed, tearing off was a one-way door: the only way a torn-off window rejoined was
by losing its document, which is to say by being closed.

**`check-reorder` 72 → 90 checks, `check-menu` 63 → 71.** No new suite and no new probe.

## Almost all of it already existed

`Join()` has snapped windows onto the stack rectangle, put their tabs in the row and repainted every
strip since the first stacking slice. `StackJoinTab` clears the sticky flag, marks the member as
arriving, calls `Join`, and activates the window. That is the whole mechanism — the same shape as the
tear-off half, where `Leave()` turned out to already do everything.

What is genuinely new is in the input:

- **This is the first gesture in the add-in that acts on a window other than the one it started in.**
  The strip being dragged *from* owns the mouse capture for the whole gesture, so no other window is
  ever told the pointer is over it. `WindowFromPoint` answers regardless of capture, which is the only
  reason this can work at all.
- **The window under the point is resolved through our own `g_strips` array, not by class name.** A
  handle that came from `WindowFromPoint` could be anything; a class-name check believes whatever
  registered that class, and reading `GWLP_USERDATA` off a foreign window is reading somebody else's
  pointer as ours. `FindByStrip` can only answer with a strip this process made.
- **The target is re-asked on every mouse-move rather than cached from the press.** The row underneath
  can change while the gesture is in the air — a document closing takes a window out of it — and a
  drop aimed at a stack that is no longer there has to be refused rather than honoured against a stale
  answer.

## The tab lands at the END of the row, and that is a decision

Not where it left from. The window is **arriving**, not being undone: the user may have torn it off
ten minutes and three reorders ago, and a tab reappearing in the middle of a row they have rearranged
since would be a surprise. It reuses the flag written for recycled frames, which is the same rule for
the same reason — *a tab that was not in the row joins at the end of it*.

The suite asserts this specifically: the window was tab 0 when it was torn off, and it comes back as
the last tab, with the three that stayed behind still in their original order.

## IDC_NO is most of this gesture, and saying so is the point

There is still no window preview following the hand. So the pointer carries the whole of the feedback,
and unlike the tear-off gesture the answer is *usually no*: everywhere except one thin band of screen,
letting go does nothing.

- Over a row it can join → **`IDC_SIZEALL`**.
- Anywhere else → **`IDC_NO`**, Windows' own "you cannot drop that here".

**A gesture whose most common state is "this will do nothing" has to say so**, or the user is carrying
an invisible thing with no way of knowing whether releasing will achieve anything. Both states are
driven: the suite carries the tab down into the window's own document first, which is a real place a
hand passes through and one that can never be joined.

**The "over nothing" state had to be announced explicitly.** It is the state the gesture *starts* in,
so a bare `onto != g_dragOnto` treated the first evaluation as no change and logged nothing at all —
the most common state this gesture is ever in would have been the one state never announced. Hence
`g_dragOntoKnown`.

## Two routes, one command, and both driven

Tearing off has a menu item and a gesture; putting back now has both too, and for the same reason.
**The gesture alone is not discoverable**: a user who took a window out through the menu will look in
the menu to put it back, and a drag onto another window's tab row is not something anyone guesses.

- **`Move Back to the Tab Row` is hidden rather than greyed on ordinary tabs**, which is the opposite
  of every other item on this menu. That is deliberate: `Move to New Window` is something any tab
  *could* do, so it stays put and greys; this is something only a window standing on its own can do,
  and on every other tab it would be a permanently grey line describing a state that tab is not in.
- **The shared expected-menu list grew a second shape** rather than a footnote. `Get-TabMenuItems`
  documented in so many words that the list is the same whatever the row holds, because items grey
  rather than hide — and this item breaks that. `-OnItsOwn` returns the ten-item list, and the
  torn-off window's whole menu is asserted against it, because an item that hides changes the *shape*
  of the menu and the shape is what the rule about `Close All` never moving under the pointer is for.
- **The menu route is driven in `check-menu`, the drag route in `check-reorder`** — the same split as
  tearing off, and both raise the one `CMD_JOIN` so neither is a second implementation.

## What the suite asserts, and why in that shape

- **It tears the window off with the gesture proven a section earlier**, rather than building a second
  fixture, so the rejoin is tested against a window that got out the way a user gets it out.
- **The lone row's slot is computed against a count of ONE.** That window's row holds one tab while
  four Word windows exist. Asking for it wrong does not throw — the computed slots come out narrower
  and the press still lands on the strip — so `Get-TabSpot` now takes the count rather than assuming
  it. Third time this trap has been paid for in this project.
- **The drop point is confirmed against `WindowFromPoint` before the gesture starts.** That is the
  exact question the add-in asks on every move; if the answer here were not the stacked window's
  strip, every assertion after it would be about a drop that was never aimed at anything.
- **Reachability is asserted in reverse.** After a tear-off the claim is *two* windows presented to the
  shell; after a rejoin it is **one**, because a window that went back into the stack has to have
  given its taskbar button and Alt+Tab entry back up.
- **The section sits above everything else in the file**, so all the sections below run against a row
  containing a window that left and came back. That placement has now found the last two defects in
  this suite.

## Still open

- **A live preview** — still the honest gap, now on both gestures.
- **Two windows both standing on their own cannot be dropped onto each other** to form a stack: the
  target has to be a window that is currently *in* one. Stated rather than discovered; `StackCanRejoin`
  requires `JoinedCount() >= 1`.
- **Rejoining is gated on stacking being on, not on `TabTearOff`.** With tearing off switched off
  nothing can be outside the stack in the first place, so the case cannot arise — but it is a
  reasoned answer rather than a measured one.

---

# The tab you are carrying

**Passed 2026-08-17.** `check-reorder` 102 → 118 checks, harness self-test 30 → 32. **No new suite
and no new probe** — the tear-off section already drove the gesture out of the row, back into it, and
out again, which is exactly the three states this has.

## The gap this closes, in the words both halves of this file used

The mechanism half said it, and the gesture half said it again:

> The feedback is the row letting go plus `IDC_SIZEALL`, and the honest limitation is stated: a tab
> out of the row is over Word's document where the strip may not paint, so there is **no floating
> window preview** the way a browser does it.

It is worse than "no preview", and the code says so plainly. Out past the threshold the carried tab
is drawn **in the row, in the slot it came from, and deliberately not moving** — the torn branch of
`DragMove` stops updating `g_dragLeft` and stops repainting. So the pointer walks away from a picture
of the tab it is supposedly holding, and the only thing that moves with the hand is a cursor shape.

Now a card comes with it: one tab, painted by the same `DrawOneTab` the row paints a lifted tab with,
in a small popup owned by the frame.

## Decisions worth not re-deriving

- **A WINDOW, NOT SOMETHING DRAWN INTO THE STRIPS, AND THE REASON IS A MEASUREMENT THIS FILE ALREADY
  CONTAINS.** A carried tab *inside* the row repaints every strip in the stack on every mouse-move,
  which the gesture slice measured leaving the add-in a second behind the pointer — and which is
  precisely why the torn branch stopped doing that work. Moving a popup is one `SetWindowPos` that
  repaints nothing: not the strips, not Word. **The feedback that was missing went on the side of the
  gesture that had spare time, not the side that did not.** Had it been drawn into the row, this
  feature would have re-created the defect the last slice removed.
- **`WS_EX_TRANSPARENT` is load-bearing here, not tidiness.** The tooltip wants it so the pointer can
  pass through; this window is *under the pointer for the whole gesture*, and **the drop is resolved
  with `WindowFromPoint`**. A card that answered for itself would make every tear-off and every
  rejoin land on the thing being dragged rather than on what it was aimed at. Same three flags as the
  tooltip — `TOOLWINDOW | NOACTIVATE | TRANSPARENT` — for what turns out to be a sharper reason.
- **It is the width of the tab it came from, including a squeezed one.** The card is built from
  `layout.tab[home]`, so in an overflowing row it is as narrow as the row made it. A fixed size would
  look wrong in exactly the case where the user is most likely to be reorganising: too many documents
  open.
- **Held at the point it was grabbed by** — the same `g_dragGrabDx` the row carries a tab with, and
  `g_dragPressY` for the other axis. The moment the row lets go is already the instant when the most
  changes at once; a card that also jumped under the hand would read as the drag having gone wrong.
  Asserted rather than assumed: grabbed 51px in, holding the card 51px in.
- **Its own DIB per paint, not the strip's.** `EnsureSurface` keeps exactly one buffer per strip and
  rebuilds it whenever the size asked for changes, so borrowing it at a different size would throw
  the row's buffer away and rebuild it on every paint for the length of the drag. A card is painted
  when it appears and not again — moving a window does not repaint its contents — so a per-paint DIB
  is the cheaper of the two.
- **An opaque window with the strip's own background around the card, and that is a choice.** The
  margin exists because `SurfLift`'s shadow reaches 4 logical px past the card and would otherwise be
  drawn into pixels the window does not own, leaving it square-edged at the bottom — the one part of
  it that says it is off the row. Photographed over Word's workspace it reads as a piece of the row
  being carried. **The alternative is `UpdateLayeredWindow` with per-pixel alpha**, which would let
  the margin vanish and the shadow stay soft; it is not done because the compositing here writes
  colour and does not track coverage, so the alpha channel would have to be threaded through
  `SurfFill`/`Blend`/`SurfLift` — a change to the shared painter for a margin, on the day of an
  install.
- **No card on the REJOIN drag, deliberately.** Dragging a lone window's tab towards a stack already
  has something on screen representing the thing being dragged: the window itself. Tear-off is the
  case where the row has let go and *nothing* does. The card is needed in one and would be a second
  picture of a visible window in the other.
- **New switch `TabGhost`** (default 1), and it is the one switch here whose `0` removes feedback
  rather than behaviour — the gesture is untouched. It exists because this draws over Word's document
  with a window the add-in owns, which is the kind of thing that renders badly on a machine nobody
  has tried; tear-off should not have to be given up over it.
- **Driving `TabGhost=0` needed its own Word restart, and folding it into the `TabTearOff=0` section
  would have been a check that passes whatever the switch does.** With tearing off refused the tab
  never leaves the row at all, so "nothing is carried" is true there for a reason that has nothing to
  do with this switch. The assertion is a **pair**: the tab still comes out of the row *and* the
  pointer still shows `IDC_SIZEALL` (the gesture is untouched), **and** nothing is carried and the
  add-in never even created the window. Either half alone is satisfied by something being broken.
- **The control is the first assertion, not the last.** "A card is on screen" proves nothing without
  a moment in the same gesture where there is not one — so the suite asserts there is nothing carried
  while the drag is still in the row, and asserts it *after* the row has been rearranged, so the drag
  is provably real at that point.
- **Following is asserted as a DELTA against the pointer's own delta**, not as a position. A card
  parked at a fixed offset from the strip would satisfy a position check. Measured: pointer moved
  180px, card moved 180px.
- **Both harness readers are one function.** `Get-OurTopLevel` finds the add-in's own visible
  top-level window of a class, and the tooltip and the card share it; `Wait-OurWindow` is the bounded
  wait with the latency print, shared the same way. Two copies of "find our window" is the shape
  [[one-copy-of-what-decides-truth]] is about, and this project has paid for it three times.
- **THE SELF-TEST CHECKS THE CARD AT A SIZE THAT IS UNDER THE HARNESS'S SIZE FLOOR, AND THAT IS THE
  INTERESTING ONE.** A card in a squeezed row is 78x72 here — narrower than the 150x60 floor — so it
  would be classified `chrome` rather than `question` and would not stop a run. **But it would be
  right for the wrong reason**, and the day that floor moves it becomes a `question` in the middle of
  every tear-off drag, which is the exact failure the tooltip cost a slice to find. The class name is
  what keeps it out, at every size, and now that is asserted at both.

## What it measured

```
carried tab (627,783 170x66)  from a 154px tab at 192 dpi, shadow 8px
grabbed 51px into the tab; the pointer is holding the card 51px in
pointer moved 180px, the card moved 180px
the carried tab: up after 8ms / gone after 4ms / up after 4ms / gone after 6ms
```

`170 = 154 + 8 * 2`, and the four latencies are the row letting go, taking it back, letting go again,
and the release.

## Two mistakes of mine, both in the drive script rather than the product

Worth writing down because both produced a symptom that looked exactly like a broken threshold.

1. **Client and screen coordinates mixed in one expression.** `Children()` answers in frame-client
   space and `RectOf()` in screen space; taking the strip's bottom from the first and adding it to a
   point aimed at the second put the pointer 180px from where it was meant to be. The add-in
   correctly reported the tab leaving the row *upwards*, and the "come back into the row" step never
   came back, so the card correctly never hid. **Three assertions would have been reported against
   the product.**
2. **The strip was taken from whichever frame enumerated first, not from the foreground one.** Every
   stacked window has a strip and they share a rectangle — until one does not, and a leftover window
   from an earlier run made the measured tab rect belong to a window that was not in front. The Y
   offset it produced was 78px of nothing. [[word-lays-out-only-focused-window]], in the test rig
   rather than the product.

