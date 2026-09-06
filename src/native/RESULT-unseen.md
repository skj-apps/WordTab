# Two things nobody could see

From the work rig's report of **2026-09-04**, the first report taken on `ef49c79` — the installed
DLL's SHA256 (`31D88300…`) matched the local build byte for byte, so for once every line in it was
judgeable.

Neither of these was on the queue. One was invisible because it drew nothing; the other was
invisible because the user had stopped noticing it.

---

## 1. On Word's Black theme, hovering a tab did nothing at all

**The report says it in its own words**, on the one line where the palette is adopted:

```
07:53:51.643  strip  palette sampled from Word's ribbon in 0x1140C:
              chrome=RGB(10,10,10) dark  well=RGB(0,0,0) card=RGB(10,10,10) edge=RGB(36,36,36)
```

`chrome=RGB(10,10,10)` is Word's **Black** theme with the window inactive. `DerivePalette` was:

```c
out->back  = Step(chrome, -26);      // 10-26 -> clamps to 0
out->hover = Step(chrome, -13);      // 10-13 -> clamps to 0
```

Both `RGB(0,0,0)`. The hover fill painted the well colour onto the well. Every hover, on every
inactive tab, all day, since 2026-08-28 — and the same collapse took the *menu-is-open-on-this-tab*
highlight with it (`strip.cpp:2333` and `strip.cpp:2673` are the only two consumers of `hover`, and
both are direct fills).

**Why it was never seen here.** The scheme is measured off one number — 26, the step Word itself
takes from the ribbon to the workspace below it — and that measurement was taken against **Dark
Grey**, `RGB(41,41,41)`. Every theme Word ships sits above chrome 26 except Black. Nobody asked what
the arithmetic did one theme further down, and subtraction below its own step is where it went.

**Why a tab row where hover does nothing never gets reported:** it reads as *nothing happened*. That
is not a complaint people file.

### The fix, and the one it replaced

```c
out->hover = Mix(out->back, chrome, 50);      // was Step(chrome, -13)
```

Halfway between the well and the card, which is what a hover *is*: the tab, part-way brought up.
For any ribbon lighter than the step, `back` is `chrome-26` exactly and the midpoint of `chrome-26`
and `chrome` **is** `chrome-13` — so this is bit-identical at 41 (Dark Grey), 243 (Light Grey) and
255 (White). The two rules part company at chrome 0..25 and nowhere else: 23 values of 256.

**The first attempt was `Step(out->back, +13)` — up from the well instead of down from the chrome —
and it was wrong.** It also produces `chrome-13` for every ribbon above the step, it also cleared the
collapse, and it passed a test written for it. What it did on Black was give `well=0, hover=13,
card=10`: a hover **lighter than the tab it is hovering**.

`check-look.ps1:517-519` had already said so, before any of this was written:

```powershell
Assert (Get-Luma $hoverColour) -gt (Get-Luma $well)    # brighter than the well
Assert (Get-Luma $hoverColour) -lt (Get-Luma $card)    # and darker than the active card
```

The hover is not "visible"; it is **between**. A fixed step up from the well satisfies the first
assertion and breaks the second, and it would have overshot the card at 13 of the 256 ribbons. A
midpoint cannot overshoot by construction, at any ribbon, which is the property worth having.

The lesson is narrower than "write tests first". It is that **a test written to match the arithmetic
it is checking proves nothing.** The replacement asserts the *ordering* — `well <= hover <= card`,
strict wherever there is a unit of room — which is a property of the design rather than a
transcription of the code, and it is `check-look`'s property, not one invented here.

### What is asserted, and where

`tools/palette-test.cpp` grows two groups (`check-palette` 23 → 36 checks). It replays every ribbon
Word can give us, 0..255:

| | |
|---|---|
| `outOfOrder == 0` | the hover is never outside `well..card`, at any ribbon |
| strict wherever there is room | and strictly between the two whenever `card - well >= 2` |
| `overshootAlt == 13` | a fixed step up from the well would overshoot the card 13 times — the first attempt, kept so that shape cannot come back |
| `collapsedOld == 14` | the shipped rule collapsed hover onto the well below chrome 14 |
| `collapsedNew == 2` | the midpoint collapses only at chrome 0 and 1, where the card has collapsed onto the well too and there is nothing to be between |
| `differ == 23` | the blast radius, as a number, so it cannot quietly grow |

Plus the four themes by name: Black idle → hover 5 between 0 and 10; Black active → 4; Dark Grey,
Light Grey and White bit-identical to what shipped.

### What could NOT be verified here, and why that is the whole reason this suite exists

**This rig cannot render a black ribbon.** Tried, and recorded so nobody tries again:

- `HKCU\Software\Microsoft\Office\16.0\Common\UI Theme = 4` is **reverted to 6 by Word on startup**.
  The value that wins is the roaming one,
  `…\Common\Roaming\Identities\<id>_LiveId\Settings\1186\{0…0}\Placeholder`, and its encoding is not
  `UI Theme`'s: `Placeholder=2` gives `UI Theme=6`, and both `4` and `5` gave `UI Theme=0` (Colorful,
  light). Restored to `2` afterwards.
- A `check-look` run on this rig with the theme pushed to Black still photographed the ribbon at
  `RGB(41,41,41)` and passed 44 checks — i.e. it re-proved Dark Grey and could not see the defect.

That is the same wall `check-palette.ps1` was created against and says so in its own header: *the dev
rig cannot produce the input.* The verification is therefore the arithmetic, plus the blast-radius
number, plus 16 green suites proving nothing moved on the themes this rig **can** produce. A
`TabChrome` override alongside `TabDpi` would make the Black path drivable here forever; not built,
because this slice was not asked to grow a production setting.

---

## 2. Minimising Word left the other windows drawn on the desktop

`PresentWindow` puts `WS_EX_TOOLWINDOW` on every window in the row except the active one, to keep
them out of the taskbar and Alt+Tab. **A tool window is not shell-managed, so when it is minimised
Windows does not park it near -32000** — it tiles it in the old minimised-window area, at a real
desktop coordinate, and draws it there.

The report contains its own control experiment, three lines apart:

```
07:57:04.491  WM_SIZE hwnd=0x111F6  minimized 0x0 outer (-32000,-32000 237x39)  nobody here asked for it
07:57:04.518  WM_SIZE hwnd=0x31376  minimized 0x0 outer (0,1101 237x39)         WordTab moved it
07:57:04.572  WM_SIZE hwnd=0x1140C  minimized 0x0 outer (237,1101 237x39)       WordTab moved it
```

The one window still shell-managed parked properly. The two followers landed on screen, one
`SM_CXMINSPACING` apart.

**The user confirmed it: "i can confirm ive seen it but forgot".** That is what a small ugly thing
gets — seen, shrugged at, never reported. It was ranked as *medium if Windows paints them, cosmetic
if not*, and the only open link in the chain was whether they were composited. They are.

### The fix

`ParkMinimised` in `stack.cpp`, called from the minimise branch of `StackOnFrameSize` immediately
after each follower goes down. `ptMinPosition` + `WPF_SETMINPOSITION`, because that is the documented
way to say where a minimised window sits — with `showCmd` forced back to `SW_SHOWMINNOACTIVE` first,
since `SetWindowPlacement` would otherwise re-apply `SW_SHOWMINIMIZED`, which **activates**, and
handing focus to a window on its way down is exactly what the caller's `SW_SHOWMINNOACTIVE` was
chosen to avoid.

Nothing touches `rcNormalPosition`, so the size the row comes back at is untouched — and the restore
path does not depend on it either way: it already calls `MatchTo` on every follower after bringing
them up, which overwrites the rectangle regardless of where they were parked.

It logs the rectangle the window **actually ended at**, not the one asked for. This rig is not the one
that matters, and "the park took" is not a thing to assume about a window the shell has opinions
about.

### The A/B, which is the point

`check-stack.ps1` already minimised the stack and asserted every window went down and came back. It
never asked **where**. That assertion now exists, written as *does not intersect the screen* rather
than *is at -32000*, because the harm is that the user can see it; -32000 is only how Windows happens
to say the same thing.

Same rig, same suite, one binary apart:

```
ef49c79 (shipped)     FAIL  every minimised window is parked off-screen (2 still visible)
                        still on screen: 0x3049A tool=True at (0,1537 83x28)
                        still on screen: 0x10544 tool=True at (83,1537 83x28)
                      FAIL  1 of 65 checks failed

this build            PASS  every minimised window is parked off-screen (0 still visible)
                      PASS  65 checks, 0 failures
```

83px apart here against 237px on the work rig — `SM_CXMINSPACING` scaled by DPI, the same mechanism
seen from two machines. **The field defect reproduced on the dev rig, unprompted, on the first run of
the new assertion.** And the add-in log carries the same A/B by itself:

```
13:41:04.318  stack  minimised with the active window: 2 other window(s) went down too      <- shipped, no park
13:42:14.131  stack  hwnd=0x…002600A4  parked while minimised: (-32000,-32000 83x28)        <- this build
13:42:14.134  stack  hwnd=0x…0004056A  parked while minimised: (-32000,-32000 83x28)
```

---

## Battery

**All 16 suites green, 848 checks** (was 837): `check-palette` 23 → 36, `check-stack` 64 → 65.
`check-reorder`, the flaky one, ran 118/118.

---

## What this report also produced, and did not get fixed here

Ranked, worst first, from a nine-way investigation of the same log. None of it is in this build.

1. **The row grows 21-27px on every drag** (`1240x1397 → 1261x1410 → 1288x1437` across two drags that
   both traced *0 resizes*). Word hands the master a bigger rectangle ~48ms after `WM_EXITSIZEMOVE`
   — the `WM_SIZE` line added in `ef49c79` for exactly this says `nobody here asked for it` — and the
   row copies it to every follower. What we own is the **ratchet**: `RememberRowRect` runs at
   `EXITSIZEMOVE`, i.e. before the growth, so a drag banks the *previous* drag's growth. Their
   registry holds `1261x1410` while their windows are `1288x1437`, so their next Word start comes
   back wider. **Falsifiable in the next report.** Why Word does it is unknown; that message is
   recorded nowhere, because `TraceRecord` is dead outside the modal loop. Probe first. Any guard
   must not be conditioned on `!g_inModalLoop` — that fires on Win+Left/Win+Right, which
   `frames.cpp:485` already names as unguardable.
2. **The 500ms janitor, dot poll included, runs inside Word's modal drag loop** and cannot ask whether
   a drag is in progress — `g_inModalLoop` is a file static with no accessor. Their drag traces carry
   260.8ms and 173.3ms gaps; a full row pass there costs a mean of 179,188us.
3. **The active-window dot poll is uncapped and memoryless.** One 52ms sample bought 20s nominal and
   35s observed of silence on the tab with the keyboard — a saved document still showing the unsaved
   dot. The comment at `strip.cpp:6083-6094` says this line "has never yet appeared after the first
   pass of a process". It has now, seven times across two sessions.
4. **The row governor ping-ponged 500↔2000ms 36 times in 3m50s** — not the SharePoint bistability
   `f333349` fixed, but plain jitter across a bare 4000us threshold with no hysteresis. Costs 0.10
   points of UI thread; costs 49% of the log's bytes/hour. One line: `else if (RungFor(basis * 2) < *rung)`,
   mirrored by hand into `tools/governor-test.cpp` (`check-governor.ps1` pins the copy).
5. **`OnePage` ran six times and logged nothing**, and `FALSE` has five meanings in
   `WordTabOnePageView` — "healthy, nothing to do" and "the correction was attempted and Word refused
   the write" are byte-identical in the log. One `else` at `stack.cpp:758`, six lines a day.
6. **The strip briefly rendered at the other monitor's DPI** across a minimize/restore (a clean ×1.2
   then ÷1.2, no `WM_DPICHANGED`, no monitor crossing), self-healed both times. Probe, not a guard —
   the guard as first drafted would have turned a 1s flicker into a permanent 244px overhang.

**Not a thing:** the log roll. Measured 9,126 B/h against a 512KB cap is ~57 hours of Word, seven
working days. The session under investigation used 8.4% of it.
