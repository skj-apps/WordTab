# The row's own size

**The complaint.** From the work rig, on `build-00f68f2`: *"also still opening up 4 pages wide and
word tab dialog appears every open change that. very annoying along w/ multiple pages opening."*

Two complaints, not three - "4 pages wide" and "multiple pages opening" are the same thing said
twice, and the dialog is separate (see below).

## What the log said

Their `settings.ps1 -Report` dump settled it without needing a guess. Every fresh Word start in it
begins the same way:

```
attach  hwnd=0x...  (existing)  thread=...  visible=0  rect=(418,418 3840x1536)  frames=1
```

That is **Word** restoring a frame about 3840px wide, on a 5120x2160 primary at 150%, **before the
add-in has touched anything**. At that width Word's own layout puts about four pages side by side.
The multi-page view is Word's, not our drawing, and it is not a WordTab defect.

What *is* ours is that it spreads. `Join` in `stack.cpp` snaps every window that joins to the first
one:

```
stack  hwnd=0x...  joined, snapped to 0x...  (2 in the stack)
```

and the live section of their report showed all three frames at exactly the same rectangle,
`(-8,668 1292x1418)`. So **whatever size window #1 happens to open at becomes the size of the whole
row.** Sizing the row fixed it for that session; the next start took Word's rectangle again.

## What was built

`RowSize` (on by default). The row records the size the user settles on and applies it to the window
that defines the row, over the top of whatever Word restored.

- **Recorded at `WM_EXITSIZEMOVE`** - the moment a person stops dragging - and only for the window
  that currently defines the row. Not from the intermediate rectangles a drag passes through, of
  which the trace shows dozens per drag.
- **`GetWindowPlacement().rcNormalPosition`, not `GetWindowRect`.** The restored rectangle survives
  underneath a maximized window, so maximizing the row does not overwrite the size chosen for it;
  the maximized state is stored beside it as a state. Minimised is refused outright - Windows parks
  a minimised window near -32000, which is the same reason `MatchTo` will not copy a minimised
  master's rectangle.
- **Applied in exactly one place**: the `if (!master)` branch of `Join`, the branch whose comment
  already said *"First window in: it defines where the stack is."* That is the only window whose
  size matters, because every other one is snapped to it.
- **Refused when unreachable.** `MonitorFromRect(..., MONITOR_DEFAULTTONULL)`. Monitors get
  unplugged and laptops get undocked, and a row restored onto a screen that no longer exists is a
  row the user cannot reach. Same rule `MatchTo` keeps for a minimised master; the log says so when
  it happens.
- **Four DWORDs, not a `REG_BINARY` blob** - `RowLeft` / `RowTop` / `RowRight` / `RowBottom` /
  `RowMaximized`. Every other value under the key is a readable DWORD that `settings.ps1` and regedit
  can both show. Negative coordinates survive as two's complement, which is not a corner case: their
  own row sat at `left = -8`.

**This does not explain why Word restores the rectangle it does, and does not need to.** That could
not be reproduced here - this machine has neither a 5120px screen nor a Word that restores wide - and
the fix writes over the answer rather than correcting it. Same position `tools\probe-view.ps1` is in.

## What was checked

`tools\check-rowsize.ps1`, new, **24 checks, 0 failures**, wired into `check-all.ps1` after `row`.

1. Nothing remembered: the add-in says nothing, moves nothing, and invents no rectangle.
2. A settled size is stored, and what is stored matches what `GetWindowPlacement` reports.
3. It is applied on the next start, and the window comes back at it.
4. A rectangle at `(-30000,-30000)` is refused, logged, and the window stays reachable.
5. `RowSize=0` turns the whole thing off, back to Word's rectangle.

**Two failures in the suite's first run were the suite's fault, and both are worth keeping.**

- `Wait-WordReady` returns when the **strip binds**. The stack **join** is a separate, later event -
  measured about half a second later - so asserting straight after `Wait-WordReady` reported "no
  window joined" while the log plainly showed one 500ms afterwards. The suite now waits for the join
  itself.
- `Close-AllWord` *asks*; it does not wait for the process to be gone. Opening the next document
  while the last WINWORD was still dying got the new one handed off and exited ~250ms after startup,
  so three steps ran against a window that never joined anything. The suite now waits for the
  process.

And the shape worth remembering: in that first run the **geometry** assertions passed while the
**log** assertions failed, because Word's own window memory happened to produce the same rectangle.
A geometry check alone would have reported success for a feature that had not run at all.

## Also in this slice: the load banner

Separate complaint, same report. `install.ps1` wrote `ShowLoadBanner` with an unconditional
`-Force` on every install:

```powershell
New-ItemProperty ... -Name 'ShowLoadBanner' -Value ([int](-not $NoBanner)) -PropertyType DWord -Force
```

The comment **directly above that line** sets out the rule it breaks - that `-Force` must never
rewrite a switch the user set, because this key is the user's. `ShowLoadBanner` was the one value in
it the installer overwrote, so a user who turned the dialog off got it back on the next build, and
the shipped one-click `.cmd` never passed `-NoBanner`.

Now: seeded when absent, left alone when present, `-NoBanner` / `-Banner` to say it explicitly, and
**the seed is off**. The banner earned its place when nothing else proved the add-in had loaded; the
installer's smoke test and the log both say so now, and a dialog on every Word start is a cost paid
forever for a fact learned once. A/B against the real installer, all four cases as expected: absent
seeds off, an existing 1 is preserved, `-Banner` overrides a 0, `-NoBanner` writes 0.

### And that fix did nothing for the machine that reported it

Their report, on a build carrying it: `ShowLoadBanner  on (1)` - the only value set in the entire
key - and the log has the dialog opening at 07:51, 08:58 and 13:11, each one clicked away by hand.

Preserving an existing 1 is right when the 1 is a choice. Theirs was not a choice: the old
installer stamped it, and "never rewrite a switch the user set" cannot tell the two apart, so the
fix stopped the cause and left every machine that already had it exactly as annoyed as before.

So the stamp is cleared **once**, guarded by a `BannerSeedFixed` marker written on every install
whichever way the decision went. Driven against the real installer, three cases:

| starting state | what the installer did |
|---|---|
| `ShowLoadBanner=1`, no marker | cleared it to 0, said so, wrote the marker |
| the user sets it back to 1 | left alone - *"left as you set it"* |
| neither value present | seeded 0 as a first install, not reported as a clearing |

## 2026-08-21: the report came back "still opened 3 pages wide", and the fix above is why

Third time of asking. The remembering works - the suite proved it and still does - but remembering
is not a fix, because **until the user settles on a size there is nothing to remember, and Word's
rectangle stands.** The shipped answer required them to drag the window once and let go. The report
that came back says they did not, and a fix that depends on the user doing something is not a fix
for a complaint that has now been made three times.

**The arithmetic, which is what was missing.** Word draws a page at its true size at 100% zoom, so a
document frame N page-widths across shows N pages side by side. Nothing here is our drawing:

| | their rig | why |
|---|---|---|
| screen | 5120x2160 at 150% | so 144 dpi |
| a page | 8.5in x 144 = **1224px** | Letter at 100% zoom |
| Word restores | **3840px** | 3840 / 1224 = 3.1 -> **three pages** |
| Word maximized | **5120px** | 5120 / 1224 = 4.2 -> **four pages** |

"4 pages wide" and "still 3 pages wide" are **one complaint at two window sizes**, and the numbers
fall out of the sum exactly. They were logged as two separate readings for three days.

**What is added: a default width, applied when nothing is remembered.** `ApplyDefaultRowRect` in
`stack.cpp`, reached from the one branch of `ApplyRememberedRowRect` that used to give up. If the
window Word opened is **three or more pages across**, it is narrowed to about **1.4 pages** - one
page and room around it - keeping its height and its corner, un-maximizing if it has to, clamped
back onto the monitor it was already on.

Three decisions worth arguing with:

- **Three pages, not two.** A maximized Word on an ordinary 1920x1080 screen at 100% is about 2.3
  pages across, and two pages side by side is what Word has always done there and what nobody has
  complained about. Three means every ordinary screen is untouched and only the wide ones - which
  is where all three reports came from - are reached. That is the whole blast radius.
- **Letter (8.5in), not A4.** A4 is *narrower*, so a threshold computed from Letter is reached later
  and a width computed from Letter is wider. Both errors fall towards leaving the window alone.
- **It is not written to the registry.** "Remembered" has to keep meaning *a size the user chose*.
  A default that wrote itself into that slot would be indistinguishable from one, and would then be
  applied forever on a machine where it was wrong, with no way for the user to tell which of the two
  they were looking at. A remembered size beats the default, always, from the moment it exists.

**And the second hole, which would have kept the first fix from ever working.** `RememberRowRect`
was reached from `WM_EXITSIZEMOVE` and nowhere else. That message ends the modal move/size loop, so
it catches a **drag** and nothing else: double-clicking the title bar and Win+Up never enter that
loop. A user who did exactly what they were asked - settle on a size - by maximizing was never
recorded, and got Word's rectangle back on the next start. Now recorded from `WM_SIZE` as well, on
`SIZE_MAXIMIZED` and `SIZE_RESTORED`, guarded on `g_inModalLoop` so a drag is still recorded once at
its end, and on `StackIsSyncing` so WordTab's own moves - `MatchTo`, the remembered rect, this very
default - are never read back as the user's choice.

**And the second hole was first plugged with the wrong message, which the battery caught.** The
first attempt keyed the recording on **`WM_SIZE`**, and it passed its own suite 39/39 - then took
`title` and `dot` down, 14 failures between them, in the battery straight afterwards. `WM_SIZE` is
*every* resize there is: Word's own, another program's, a check suite's. Suites that size Word to a
known rectangle were now "settling on" it, the row was remembered across suites, and later ones ran
against windows dragged to 618x700 where Word's ribbon wraps differently and the document frame sits
at 637 instead of 356. Every one of those 14 failures read as a Protected View bug.

It is worse than untidy. On a wide screen, **Word's own startup rectangle would install itself as
the size the user chose** - and the remembered size beats the default, so the fix would quietly
disable itself on exactly the machine it was written for.

The right message is **`WM_SYSCOMMAND`** with `SC_MAXIMIZE` / `SC_RESTORE`, which is the same
distinction `SC_CLOSE` already turns on eleven lines above it in that file: *this message is a person
operating the window*. `ShowWindow(SW_MAXIMIZE)` maximizes the same window without it. The suite now
checks both halves in one step - a bare `SetWindowPos` resize is **not** recorded, the same window
maximized by system command **is** - so the regression cannot come back silently. Still not covered:
Win+Left and Win+Right, which snap without the modal loop and without a system command, and there is
no message that says a person did it.

**One reader for the DPI.** `StripDpiOf` exports what `strip.cpp` already computes, so the two files
cannot disagree about what screen this is - and because it honours `TabDpi`, the page arithmetic
became **testable on a one-monitor rig**. At a forced 72dpi a page is 612px, three pages is 1836px
and the default is 856px: the same sum as their 1224/3672/1714, in numbers this screen can hold.

`check-rowsize.ps1` is **41/41** (was 24). New: a window Word opens too wide is narrowed and says so;
the narrowed width is about one page; nothing is written to the registry; a remembered size still
beats it; a window that was *not* too wide is left alone; and maximizing without a drag is recorded.

**Two bugs in the new checks, both the same shape and both this suite's speciality.** Step 8 reused
the *previous* step's log mark, so `Wait-Joined` was satisfied by an old join line and returned at
once - the window was maximized about half a second before it was a member of anything, and the
suite duly reported that maximizing is not recorded. Step 1 was made deterministic the other way:
Word is primed narrow first, so "Word's rectangle stands" is tested as *left alone because there was
nothing wrong with it*, not as *nothing looked*. Priming Word's own window memory - size it, close
Word, reopen - is what lets a machine with one ordinary screen reproduce "Word opened it far too
wide" on demand.

## Still open

- **The first ~500ms.** The default is applied at `Join`, which is about half a second after Word
  has already put the window on screen, so a wide window is briefly visible before it narrows. The
  persistent width was the complaint and this fixes that; the flash is not fixed.
- **Whether this actually fixes the complaint on their rig.** Still cannot be reproduced here - but
  unlike last time the mechanism is arithmetic rather than a guess, and the log now prints the sum
  it did: the width Word opened at, how many pages across that is, and what a page measures. If it
  is still wrong there, that one line says which of the three numbers is not what we think.
- ~~**Office Tab is half-live on that machine**~~ - **closed 2026-08-21.** Their report has all
  three ProgIds with an empty `LoadBehavior` and nothing in Word's Disabled Items: it is gone, and
  what that rig reports from now on is WordTab alone in the frame.

## 2026-08-21 evening: the report came back, and the mechanism had worked

At **08:59:45** they dragged the window to a size they wanted, and `4c4c17a` recorded it:

```
stack  the row remembers its size: (-10,673 1305x1419)
```

At **13:11:28**, a fresh Word process, before any of the default-width work below was on that
machine:

```
stack  hwnd=0x...A1202  the row takes the size it was left at, not the one Word restored: (-10,673 1305x1419)
```

1305px at 144dpi is **1.07 pages**. The row is one page wide and stays one page wide across
restarts. The complaint that took three reports is over on that machine - by the one manual step,
five days late.

**Which does not make the default width optional.** It is what stands in for that step on a machine
where it has not happened, and their `RowLeft` is one profile reset from being gone. Both halves
ship together.

### The report was printing the answer as gibberish

The same report said:

```
Row remembered  (4294967286,673 -4294965991x1419)
RowLeft         4294967286
```

`4294967286` is `-10` read as unsigned. The four values are `REG_DWORD`s and a window's left edge is
signed - theirs is genuinely at x=-10 - so whatever reads them back in PowerShell hands over the
raw 32 bits and the width came out as `-4294965991`. The add-in itself was never confused: it reads
them as `LONG` and its own log line says `(-10,673 1305x1419)`, which is how the two could be
compared at all.

`Get-Signed32` in `settings.ps1`, applied to the four row values, and checked against the exact
numbers from that report: `(-10,673 1305x1419)`. The raw list underneath still prints what is
literally in the registry, because that section's job is to show anything unexpected as it is.

### The suite could not run two of its steps, and now says so

`check-rowsize` steps 6 and 7 force `TabDpi=72` so that three pages is 1836px, small enough for a
window this rig can hold. **72 is the floor**: `PageWidthPx` clamps a forced dpi outside 72..480
back to 96, so a smaller page cannot be asked for - which was found by trying, when a screen-derived
63dpi silently produced no narrowing at all.

And the screen is not a constant. This rig measured **2856px** across in the morning and **1740px**
the same evening - it is in an RDP session now - so `Prime-WordWidth` could no longer make Word
wide enough to trigger anything, and three checks went red saying exactly that: *"Word was primed
1660px wide, which is over the 1836px that is three pages here"*. A precondition failing, not the
product. Those steps are now skipped by name, with the number, when the work area is under 1916px.
A check that cannot be run is a check that cannot fail, and it should say which one it is.
