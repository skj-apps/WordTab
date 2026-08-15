# Slice: one window to the rest of Windows - taskbar, Alt+Tab, minimise

**Date:** 2026-08-14
**Outcome: PASSED.** Three stacked Word windows now give one taskbar button and one Alt+Tab entry,
and the stack goes down to the taskbar and comes back as a single window.

Stacking made N windows look like one *inside* Word's frame. Everything outside it still counted N,
which gave the whole thing away at a glance - Windows 11 draws a multi-window app as a stacked card,
so you did not even have to hover.

## Proof

Neither the taskbar nor Alt+Tab can be enumerated through any API, so the mechanism is asserted and
the result is photographed.

    PASS  33 checks, 0 failures        (two consecutive runs, plus the strip slice's 40 unchanged)

    0x1E0ECE  toolwindow=False  <- active
    0x2105A6  toolwindow=True
    0x2005BC  toolwindow=True
    PASS  exactly one window is left in Alt+Tab (2 of 3 hidden)
    PASS  the one left is the active one
    PASS  every window went down with it (3 of 3)
    PASS  every window came back with it (3 of 3)
    PASS  the last window is back in Alt+Tab
    PASS  the last window still has its strip, correctly placed

The photographs, both taken by `check-stack.ps1 -Screenshot` with three documents open:

- **The taskbar:** a single flat Word icon with one dot under it. Windows 11 draws a second card
  edge behind the icon when an app has more than one window; there isn't one. Spike 2 used exactly
  this as its success signal.
- **Alt+Tab:** two entries on the whole desktop - `wordtab-stack-3.rtf - Word` and a terminal. Three
  Word windows, one entry. Visible behind the switcher, incidentally, is the point of the entire
  project: one Word window with three tabs in it.

## What it is

`src\native\taskbar.cpp` plus a presentation rule in `stack.cpp`. Two mechanisms, because these are
two different systems and only one of them is a window style:

- **The taskbar is told**, through `ITaskbarList::DeleteTab` / `AddTab`. No styles are touched, so
  there is nothing to undo badly. `CoCreateInstance` on Word's UI thread, which is an STA that has
  been initialised for COM since long before we loaded - initialising it ourselves would be wrong.
- **Alt+Tab reads `WS_EX_TOOLWINDOW`** off the window when it is invoked, so setting that style on
  the windows behind is enough, and deliberately *without* `SWP_FRAMECHANGED`: the shell's
  classification is what should change, not the window's frame. Word draws its own caption over the
  non-client area anyway, and the windows this applies to are underneath the active one where none
  of it is visible.

The rule both obey: **exactly one window is presented, and it is the active one.** Not the first
window - if the button stayed put, restoring the stack from the taskbar would bring back a document
the user was not looking at.

**Minimising is all-or-nothing in both directions.** Down, because otherwise the next document in
the stack is revealed underneath and the "one window" the user minimised is still on screen. Up,
because only the active window has a taskbar button, so any window that did not come back with it
would be stranded - minimised, no button, no Alt+Tab entry, no way to reach it. Restored windows are
shown with `SW_SHOWNOACTIVATE`, not `SW_RESTORE`, or the window the user actually clicked ends up
behind the ones it brought back with it.

Switchable, no rebuild: `HKCU\Software\WordTab\Taskbar` and `AltTab`, both default on.

## The safety property this slice is really about

Every mechanism here makes a window harder to reach, so each one needs an inverse that cannot be
skipped:

- a window that leaves the stack gets its button and its Alt+Tab entry back, before anything else
  happens in `Leave`
- `StackStop` restores every window it knows about, joined or not, then releases the shell object
- one window in a stack keeps everything: hiding the only window there is would leave the user with
  no way back to Word at all
- minimising restores all or none

`check-stack.ps1` now closes documents down to the last window and asserts it is reachable and still
correctly laid out. "The add-in ate my document" is not a recoverable first impression, and the
check exists so that failure mode has to get past an assertion before it gets to a user.

## Findings

**Membership needed splitting in two.** Joining requires a window we can measure and place: on
screen, not minimised, sensible size, has a document frame. *Staying* requires only that it still
exists and still has a document. Without that distinction, minimising the stack dropped every window
out of it, and the one window with a taskbar button came back alone.

**The layout freeze needed to cover minimised windows too.** Word squeezes a frame's interior to
nothing on the way down to the taskbar, exactly as it dismantles it on the way out. `IsWindowVisible`
is `TRUE` for a minimised window, so the freeze added last slice did not catch this: the test is
"on screen", which is `IsWindowVisible && !IsIconic`. Without it a minimise would have broadcast a
garbage layout to every other window in the stack - the same bug as the closing document, arriving
by a different road.

**Word health across the session:** `LoadBehavior` still 3, `Resiliency\DisabledItems` empty, no
`Application Error` events for WINWORD, no failed taskbar calls logged.

## Reproducing

    pwsh -File install\install.ps1 -NoBanner
    pwsh -File tools\check-stack.ps1 -Screenshot

By hand: open three documents, look at the taskbar - one Word button - and press Alt+Tab - one Word
entry. Minimise, then click the taskbar button: the whole stack comes back with the tab you left on.

## Not covered

- **`WS_EX_TOOLWINDOW` on a window Word owns is a style change to a host window**, which is a bigger
  liberty than anything else the add-in does. Nothing measured objects to it - Word runs normally,
  the windows activate, resize, maximise and tab-switch as before - but it is the first thing to
  suspect if the work rig behaves differently, and it can be switched off with `AltTab=0` without a
  rebuild.
- The taskbar button's **jump list, thumbnail and preview** are Word's own and are not touched, so
  they still describe whichever window holds the button.
- Restoring by clicking the taskbar button is exercised through `ShowWindow`, not through a real
  click on the taskbar.
- Still untested on the work rig.
