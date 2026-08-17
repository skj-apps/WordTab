# Taking a tab out of the stack

**Status: passed 2026-08-16.** `Move to New Window` on the tab context menu takes one document out of
the stacked window and leaves it standing on its own, offset from the stack, with its taskbar button
and its Alt+Tab entry back.

This is the first half of tear-off. The second half is the gesture — dragging a tab off the strip —
and it is deliberately a separate slice, because it is a `strip.cpp` input change on top of a
mechanism that is by then already proven.

**243 checks, 0 failures** across the five affected suites: menu 63 (was 37), look 48, stack 60,
startscreen 50, soak 22. **No new suite and no new probe** — the section went into `check-menu`, which
already opens a stack of documents and already has the machinery to open and read the menu.

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

## Still open

- **The gesture.** Dragging a tab off the strip is the next slice; it calls `StackTearOffTab` and
  nothing in this file changes.
- **Putting one back.** There is no way to drag a window into a stack, and a torn-off window rejoins
  only by losing its document. That is deliberate for now — a window that silently rejoined after the
  user pulled it out would be worse — but it is the obvious next question after the gesture.
- **`TabTearOff=0` has not been driven.** The switch is read at `StackStart`, so proving it needs a
  Word restart, and this slice ran out of the machine time it had asked for. Every other switch in
  this project is driven positively and this one is not yet; it should be picked up by the gesture
  slice, which has to run the battery anyway.
- **Tearing off a maximized stack is not driven either** — the three-quarters branch of `PlaceTornOff`
  is reasoned and logged, not measured. `check-menu` runs at 900x700.
