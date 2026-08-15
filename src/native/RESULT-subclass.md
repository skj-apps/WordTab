# Slice: subclass Word's frame window from inside the add-in

**Date:** 2026-08-14
**Outcome: PASSED.** WordTab now sits in the message path of every `OpusApp` frame in the Word
process, sees each step of a drag as Word performs it, and moves other windows in the same frame.

This is the slice the whole in-process move was for. Everything the earlier spikes proved was done
from another process on a 30ms poll, and dragging a window was the one case that could not survive
that. It no longer has to.

## Proof

The measurement that settles it, from `%LOCALAPPDATA%\WordTab\wordtab.log` during a 40-step drag
of one of two open documents:

    drag  hwnd=0x00000000006C0AD4  41 position updates (40 moved, 0 resized) over 616.1ms  [81 messages traced]
    drag  gap between updates: min 12.1ms  mean 15.4ms  max 16.8ms
    drag  gap histogram:  <=8ms:0  <=16ms:32  <=33ms:8  <=50ms:0  >50ms:0
    drag  VERDICT: a 30ms poll could take at most 21 samples of these 41 updates (48% unseen); 40 of 40 gaps were shorter than 30ms
    drag  lockstep: moved 1 other frame(s) on 40 of 41 updates, in the same WM_WINDOWPOSCHANGING

Read the two claims separately.

**Word moves a dragged window about every 15ms, and every single gap was under 30ms.** Not one
frame of that drag was slow enough for a 30ms poller to have been guaranteed to see it. At its
theoretical best a 30ms poll takes 21 samples where Word made 41 moves. That is the lag the user
saw as "it really glitches out", now a number rather than an impression. A longer 120-step drag
gives the same shape: 121 updates over 1853.7ms, mean gap 15.4ms, 120 of 120 gaps under 30ms.

**And we act on it.** The follower window moved on 40 of the 41 updates - all of them except the
first, which is the one where there is no delta yet - inside the same `WM_WINDOWPOSCHANGING`, before
Word had moved the dragged window at all. End to end:

    dragged  0x6C0AD4  (952,140) -> (1152,220)   delta (200,80)
    follower 0x521996  (918,106) -> (1118,186)   delta (200,80)

Timing is measured with `QueryPerformanceCounter`, not `GetTickCount`, whose ~16ms resolution would
have quantised away the very gaps being measured. Samples are buffered in memory and reported once
at `WM_EXITSIZEMOVE`; writing a log line per message would have made Word stutter and corrupted the
measurement with the cost of taking it.

Corroborating, over roughly a dozen Word launches with our window procedure in the path:

- `LoadBehavior` stays at 3 and `Resiliency\DisabledItems` stays empty.
- No `Application Error` events for `WINWORD`. Word starts, runs and exits normally.
- Attach, detach and shutdown are symmetric and logged:

      FramesStart  uiThread=27196  coordinator=0x00000000004814C4  cbtHook=installed  followDrag=1  qpc=10000000Hz
      module pinned in WINWORD: yes
      attach  hwnd=0x0000000000460AB4  (existing)  thread=27196 (ours)  visible=0  rect=(918,106 1108x1117)  frames=1
      attach  hwnd=0x00000000003F0656  (new frame)  thread=27196 (ours)  visible=0  rect=(952,140 1108x1117)  frames=2
      ...
      WM_NCDESTROY  hwnd=0x00000000003F0656  (frame closing)
      detach  hwnd=0x00000000003F0656  (destroyed)  removed=1  frames=2
      OnBeginShutdown
      FramesStop  thread=27196 (ui thread)  frames=2
      detach  hwnd=0x0000000000322892  (shutdown)  removed=1  frames=1
      detach  hwnd=0x0000000000460AB4  (shutdown)  removed=1  frames=0
      FramesStop  done  (lockCount=0)

## What it is

`src\native\frames.cpp`, plus two call sites in `connect.cpp`: `FramesStart` at
`OnStartupComplete`, `FramesStop` at `OnBeginShutdown` and again at `OnDisconnection`.

- **Discovery.** Frames that already exist are found with `FindWindowEx`. Frames created later -
  Word makes a new top-level `OpusApp` per document window - are caught by a thread-local `WH_CBT`
  hook. The hook does not subclass directly: at `HCBT_CREATEWND` the window has not had
  `WM_NCCREATE` yet and creation can still fail. It posts to a message-only coordinator window
  instead, which picks the window up once it is real.
- **Subclassing** uses comctl32's `SetWindowSubclass` / `DefSubclassProc` / `RemoveWindowSubclass`
  rather than swapping `GWLP_WNDPROC` by hand. The difference matters in a host that other add-ins
  also live in: `RemoveWindowSubclass` unlinks us wherever we sit in the chain, where restoring a
  saved procedure by hand corrupts the chain if anyone subclassed after us. These three are
  exported by name from the `comctl32.dll` in `System32`, checked before committing to them.
- **The move trace** is the measurement above.
- **The lockstep demonstration** moves every other frame by the same delta, batched through
  `BeginDeferWindowPos` so they move in one pass rather than one repaint each. It is a
  demonstration, not the stacking engine - switch it off with
  `HKCU\Software\WordTab\FollowDrag = 0`, which needs no rebuild.

`tools\drive-drag.ps1` drives a real drag headlessly so this can be re-measured without a hand
test.

## Findings worth carrying into the next slice

**Word leaves `SWP_NOSIZE` clear while a window is dragged by its caption**, even though the size
never changes: every frame arrives with flags `0x00080214` and a `cx`/`cy` identical to the
window's own. Telling a move from a resize therefore has to compare the numbers, not read the flag.
This cost a full test cycle - the followers simply never moved, silently, because the code believed
every drag was a resize.

**Every frame is on one UI thread.** Logged as `thread=NNN (ours)` on every attach, across every
run. Everything that assumes a single message loop is safe for now, and the log says loudly if that
ever stops being true.

**Frame lifetime is not document lifetime.** Closing one of two documents hid one frame - window
still alive, title reset to `Word` - and destroyed a *different* one. Word also creates frames it
never shows. So a tab strip keyed on window creation and destruction will show tabs for documents
that do not exist and miss ones that do; it has to follow `WM_SHOWWINDOW` and Word's own document
model instead. This is the largest open question for the stacking engine.

**Visibility at `OnStartupComplete` varies and cannot be relied on either way.** A cold start with
one document gives `total=1 visible=0`, as recorded in the previous slice. A start that restores two
documents gives `total=2 visible=2`. Code that waits for visibility hangs in the first case; code
that assumes invisibility is wrong in the second.

**Maximize and restore state is now in reach.** `WM_SYSCOMMAND` `SC_MAXIMIZE`/`SC_RESTORE` and
`WM_SIZE maximized 2560x1528` / `restored 1086x1104` all arrive in the subclass. Spike 2 could only
copy rectangles, which left a stacked window looking maximized without being maximized; the state
itself is available here.

**The DLL is pinned in the process** with `GetModuleHandleEx(GET_MODULE_HANDLE_EX_FLAG_PIN)` the
first time a window is subclassed, and `g_lockCount` is raised so `DllCanUnloadNow` agrees. COM
reference counting knows nothing about window procedures, and a DLL unloaded while its procedure is
in a subclass chain - or merely still on Word's stack - crashes the host.

**Word's title bar is not the title bar.** The whole caption row is covered by a `NetUIHWND` child:
Word draws its own. `WM_NCHITTEST` answers `HTCAPTION` for every point across that row, but a click
at most of them lands on the Quick Access Toolbar or the search box and the window never moves, and
a grab within a few pixels of the top edge comes out as `SC_SIZE` because the resize border wins.
Only a narrow strip between the toolbar and the search box actually drags. `drive-drag.ps1`
therefore calibrates by trying, rather than trusting the hit test. Anything that later needs to
reason about where Word's caption is should start from this rather than from `WM_NCHITTEST`.

## Reproducing

    pwsh -File install\install.ps1 -NoBanner     # build, install, register
    # open one or two documents in Word, then:
    pwsh -File tools\drive-drag.ps1 -Steps 40 -OneWay

Then read `%LOCALAPPDATA%\WordTab\wordtab.log`; the measurement is the lines starting `drag`.
`-OneWay` leaves the window displaced so the follower windows' final positions are evidence in
themselves - the default there-and-back drag ends where it started, where a perfect follow and no
follow at all look identical.

By hand, which is the test that matters: open two documents, drag one by its title bar, and watch
the other keep station with it exactly.

## Not covered

The lockstep demonstration moves windows that are wherever Word left them; it is not the stacking
engine and does nothing about z-order, sizing, or which window is on top. Resize drags are
deliberately excluded from it.

Nothing is drawn yet. Porting the spike geometry in-process - shrinking `_WwF` and painting the
strip into the gap - is the next slice, and the `WM_WINDOWPOSCHANGING` handler in `frames.cpp` is
where it goes.

Still untested on the work rig.
