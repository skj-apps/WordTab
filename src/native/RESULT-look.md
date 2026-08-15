# The look: tab cards, and a palette read off Word's own ribbon

**PASSED 2026-08-15.** `pwsh -File tools\check-look.ps1` — **41 checks, 0 failures**, plus the seven
existing suites re-run unchanged: `check-stack` 34, `check-strip` 40, `check-tabs` 22, `check-menu`
32, `check-reorder` 30, `check-startscreen` 44, `soak-stack` 22. **265 checks, 0 failures across the
eight.**

Before this slice the tabs were the placeholder rectangles they had been since the strip first drew
anything: a flat fill, a one-pixel border all the way round, a name, and an `x` drawn as two
one-pixel diagonals that stayed one pixel wide however far the rest of the window scaled. The band
they sat in was one of two greys chosen by eye, selected between by reading Office's `UI Theme`
registry value once at startup. The context menu was the system's, which meant it was white on a
black Word. Every one of those was deliberate: the look was left until last so it could be done once
against the finished set of elements rather than five times against a moving one.

This is that slice, and it turned out to be about colour far more than about shape. The shape work is
what it looks like — rounded cards, anti-aliased, a card that runs down to meet the document, a lift
under a tab being carried. The colour work is the part with the measurements in it, because the two
sources this add-in had for "what colour is Word" both turned out to be wrong, in different ways, and
finding that out took three attempts and two of them looked like they were working.

**The one-line summary of the whole slice: the strip no longer has a palette. It has Word's.**

---

## What was built

**A palette derived from one colour, and that colour is measured rather than looked up.** Word paints
a flat band along the bottom of its ribbon, immediately above our strip. The strip samples it and
derives everything from it: the well the tabs sit in is that colour minus 26 levels, the active tab's
card *is* that colour, the hover fill is halfway between, the border is a step the other way, and the
text is black or white by luminance. The 26 is not a taste: it is the step Word itself takes from the
ribbon to the workspace below it — 255 to 228 in a light theme, 41 to 9 in a dark one — so the well
lands on Word's own workspace grey without ever being told what it is. Colorful, White, Dark Grey,
Black and whatever Office ships next all come out right with no table listing them.

**Tab cards, composited by hand.** GDI has no anti-aliasing, so the strip now draws into a 32-bit DIB
it owns and works out coverage per pixel. A card is a rounded rectangle — round on top, square at the
bottom, because it meets the document there and a tab that is round at the bottom is a lozenge. At
rest an inactive tab is not a shape at all: it is a name on the well with a thin rule between it and
its neighbour, which is how a browser draws them. Hover fills it. The active one is a card.

**The hairline along the bottom of the strip stops under the active card.** It runs the full width
otherwise. That gap is the entire difference between a row of buttons and a row of tabs: one of them
belongs to what is underneath it.

**A lift under a tab being carried**, drawn by the same function that draws every other tab — it is
an argument to `DrawOneTab`, not a second drawing path, because a second drawing path is how a
dragged tab starts looking like a different kind of object.

**The empty row says something.** A Word window with no document keeps its strip and draws only the
`+`; since the last slice that is a designed state rather than an accident, and it now reads
`No document open` in muted text beside the button. The message is painted and nothing else — see
below for why that is load-bearing.

**The context menu is owner-drawn**, so it is dark on a dark Word instead of the system's white
rectangle. It is still a real `HMENU` with real strings and real `MF_GRAYED`, which is what keeps
`check-menu`'s 32 assertions reading it from outside the process.

**One place that knows what depends on DPI.** There were three — `StripAttachFrame`, `TryBind` and
`StripOnFrameDpiChanged` each set `dpi` and `stripH`, and only two of them rebuilt the font; `TryBind`
guarded it with `if (!state->font)`, so a DPI change between attach and bind left the strip at the new
height with the old font. Latent while the font was the only thing derived from DPI. This slice
derives the corner radius, the glyph stroke, the chip radius, the shadow spread and the size of the
back buffer from it too, so three copies of "what depends on DPI" became `ApplyMetrics`.

---

## The decisions worth keeping

### Not one rectangle moved, and that is why 224 checks are still measuring what they measured

`ComputeLayout` is the single source of truth for the tab row: painting draws what it returns and
hit-testing reads what it returns, and `tools\WordLayout.cs` is a second, hand-maintained copy of it
in C# that the check scripts click through. Every geometry assertion in six suites is anchored to
those numbers.

**So the restyle was confined to what is drawn inside the rectangles, and changed none of them.** The
card is drawn one logical pixel inside its tab rect rather than the rect being narrowed; the corner
radius takes pixels out of a corner that is already there; the close button and the `+` are where
they were. `WordLayout.cs` needed no edit at all. That is not tidiness — it is the reason a slice
that rewrote the whole drawing path could be checked against a suite of 224 assertions that were
written before it and still mean the same thing.

The one place this bit was the empty row. A `New document` label to the right of the `+`, with the
`+`'s hit rectangle widened to cover it, is the obvious design — and it would have broken
`check-startscreen.ps1`. That suite proves the row is empty **without looking at a single pixel**: the
`+`'s position is a function of the tab count, so it computes where the `+` would be with no tabs and
where a tab would be with one, clicks the *tab* position, and asserts nothing happens. A widened hit
rectangle reaches past that point, the click lands on the `+`, a document appears, and the assertion
that fails is three steps downstream of the cause. So the message is a label and not a button, the
hit rectangle is untouched, and `check-look` asserts the label does nothing when clicked.

### The registry is not the oracle, and Word rewrites it behind you

The palette used to come from `HKCU\Software\Microsoft\Office\16.0\Common` → `UI Theme`. Two things
measured for this slice retired that.

It is legitimately **6** — "use system setting" — on a default install, so the answer is somewhere
else anyway. And **Word rewrites it at startup from the roaming account setting**: a probe wrote 5
(White), launched Word, and read the value back as 6. A value read once while the add-in is loading
may be the previous session's answer to a question that has already been asked again.

The registry read survives as the fallback, because at the moment `StripStart` runs there is no Word
window to take a colour off. It is now stocked with the two colours that were actually measured
(41,41,41 and 255,255,255) rather than the two that were chosen by eye.

### Word's ribbon is not a child of Word's frame

The sampler looks for the ribbon by position: the widest visible `NetUIHWND` whose bottom edge is the
strip's top edge. The first version walked `GetWindow(GW_CHILD)` and `GW_HWNDNEXT` — the frame's
direct children — and found **zero candidates**, every time.

The ribbon is five windows deep: `MsoCommandBarDock` → `MsoCommandBar` → `MsoWorkPane` → `NUIPane` →
`NetUIHWND`, **all five reporting the same rectangle**. What made the mistake so easy is that the
enumeration the check scripts use (`WordLayout.Children`) is recursive, so in every listing taken
while this was being designed the ribbon looked exactly like a direct child. It is `EnumChildWindows`
now.

Only the innermost of that chain is worth sampling anyway. The outer four are `WS_CLIPCHILDREN`, so
their device contexts exclude every pixel belonging to a child, which up there is all of them — the
same reason the *frame's* own client DC answers `CLR_INVALID` at every pixel, which is where this
investigation started.

### Three ways to read a colour off another application's window, and only one of them is honest

This is the part worth not re-deriving, because two of the three look like they work.

| how | in front | covered | after a live theme change |
| --- | --- | --- | --- |
| the frame's client DC | `CLR_INVALID` everywhere | `CLR_INVALID` | — |
| the ribbon's own window DC | correct | **correct** | **stale forever** |
| the screen DC | correct | **reads the covering window** | correct |

The frame's DC is useless for the reason above. The ribbon's own DC was the obvious choice and was
what the first working version used: it reads correctly even with Notepad maximised on top, which is
the failure everybody expects. It is also a lie, and this is the measurement the slice turns on:

```
  before the flip        window DC: RGB( 41, 41, 41)    screen: RGB(41,41,41)
  3s after the flip      window DC: RGB( 41, 41, 41)    screen: RGB(255,255,255)
  12s after the flip     window DC: RGB( 41, 41, 41)    screen: RGB(255,255,255)
  after putting it back  window DC: RGB( 41, 41, 41)    screen: RGB(41,41,41)
```

**Word re-renders its ribbon somewhere GDI cannot follow.** The redirection surface behind that HWND
keeps whatever GDI last drew into it and never changes again, so the window DC answers correctly
exactly once — at startup — and is stale from then on. Which is precisely the case that mattered: a
palette that is only right at startup is the registry read this was built to replace.

So it is the screen DC, and its hazard is handled rather than hoped about. **Every sample point is
asked who owns it — `WindowFromPoint` — before it is read, and points that belong to anything but the
ribbon are not read at all.** That is not a guess about how likely occlusion is; it is the question
the hazard actually poses, answered per pixel. Measured against the earlier failure: with Notepad
maximised over Word the screen DC returned RGB(39,39,39) for Word's real RGB(41,41,41) — a wrong
answer two levels from the right one, which no plausibility check would ever have caught.

**The general rule, and it is the same one this project keeps relearning in new clothes: a
measurement that cannot say what it measured is not a measurement.** The ribbon-DC version returned a
bare `FALSE` when it could not answer, the palette silently kept its fallback, and because on this rig
the fallback and the sample are the *same number* the only symptom was a theme change that did not
take. The sampler now reports why it failed, and that line is what found the `NetUIHWND` depth in one
build instead of five.

### A theme is sticky, so polling it is honest

The sampler runs on the janitor, every fourth tick, once globally rather than once per window — a
`GetDC` and twenty-four `GetPixel`s every two seconds, off the first strip that can answer.

This project's standing rule is that transient state must be heard rather than sampled, and it was got
wrong twice over Word's save prompt at real cost. The rule does not apply here and it is worth being
precise about why: **it is about state that can appear and disappear inside one tick**, which a
sampler cannot see at any cadence. A theme is not that. Somebody changes it in File > Account, it
stays changed, and a sampler that is one tick late is one tick late rather than wrong. There is also
no event to hear: the strip is a `WS_CHILD`, and `WM_SETTINGCHANGE` is broadcast to top-level windows
only.

`ApplyPalette` is compare-before-act, so a redundant call costs one comparison, and it logs only when
the colour actually changes.

### The owner-drawn menu keeps its strings, and that was measured before a line of it was written

Owner-drawing was the mandate this slice inherited, with the alternative named and rejected: the
undocumented uxtheme ordinals would restyle menus for the whole of Word, from an add-in, which is not
ours to do.

The risk was `check-menu`, which reads the live `HMENU` from another process and asserts seven exact
labels, two separators and one greyed item — and which is the only suite that deliberately provokes
Word's save prompt, so not one to break casually. Four things were measured first, with a
throwaway Win32 program, before any of it went into `strip.cpp`:

- **An owner-drawn item keeps its string.** The documentation says `GetMenuString` has nothing to
  return for an `MFT_OWNERDRAW` item. With the string supplied through `MIIM_STRING` in the same call
  it returns it: `GetMenuString=5 '&Save'`. Every label assertion survived untouched.
- **Mnemonics still work.** Typing `n` over the owner-drawn menu returned `CMD_NEW` exactly as over a
  system one, and `WM_MENUCHAR` was never sent — because the string is still there to match against.
  No third case in the frame subclass.
- **`MFT_SEPARATOR | MFT_OWNERDRAW` is a real combination**, undocumented though it is. The item still
  reports `MF_SEPARATOR` to `GetMenuState` — which is what `check-menu` reads to expect a `-` — *and*
  it is sent to us to draw. A plain `MF_SEPARATOR` is drawn by Windows as a bright rule running the
  full width of the menu, which on a dark background is exactly wrong.
- **`WM_MEASUREITEM` and `WM_DRAWITEM` arrive with `CtlType == ODT_MENU` and `dwItemData` intact**,
  `TPM_NONOTIFY` notwithstanding — those are requests, not notifications.

They arrive at the menu's **owner**, which must be Word's frame and not the strip (the strip is
`WS_EX_NOACTIVATE` and a popup owned by a non-foreground window does not dismiss on an outside
click), so they land in `frames.cpp`'s subclass and are forwarded back. Those are the first two cases
in that file that return without chaining, and the condition is the whole of their safety.

**Ours are identified by a pointer into our own array, range-checked before it is dereferenced.** Not
by item id — the ids are 1 to 5 and would collide with anything. Not by reading a magic number
straight out of `itemData` either: that value is chosen by whoever built the menu, menus owned by this
frame are not all ours (Word's own Alt+Space menu is a real `HMENU`, and Office Tab is installed on
these machines and could be doing exactly what we are doing), and dereferencing a foreign pointer to
check a magic is a wild read inside Word's process. So: prove it points into our array, on an element
boundary, *then* read the magic, and on the draw side compare `hwndItem` against the `HMENU` we
currently have up.

Owner-drawing draws the *items*. The window behind them, its border and its Windows 11 rounded
corners stay the system's; `SetMenuInfo` with `MIM_BACKGROUND` is the only handle on the gutter and
margins, and photographed, it is enough — the menu is dark edge to edge.

### A shadow is worth nothing on a dark theme

The lift went in as a drop shadow, which is what a lift is. Photographed on a dark Word it was
invisible: the well behind it is already RGB(15,15,15) and black on near-black is nothing at all.

So the carried card is **raised as well as shadowed** — a step lighter than the active card in a dark
theme, and not raised at all in a light one, because there the card is already white and there is
nowhere to raise it to. The shadow is drawn in both. It only earns its keep in one.

---

## What the checking found

**The sampler had never once worked, and nothing said so.** Every palette line in the log read
`from the Office theme setting` — the fallback. Because the fallback and the sample agree on this rig
(both 41,41,41), the tabs looked exactly right, and the only symptom was in the one assertion that
flipped the theme. Adding a reason string to the failure path found the cause — `0 candidates` — in a
single build. The lesson is written up above; it is the same one the strip and the reconciler already
learned, in a new costume.

**Two of the six first-run failures were the test, not the code, and both were the same mistake:
measuring while something else was moving.**

- *The lift test measured a selection change.* It photographed the row at rest, pressed a tab, and
  compared. Pressing a tab **activates** it — so the previously active tab stopped being a card, and
  6242 pixels changed on a tab the lift never touched. `check-reorder` already knew this and holds
  selection and hover constant across its two photographs; `check-look` now does the same. The fixed
  number is 0.
- *The corner probe walked a diagonal that was not on the arc.* It sampled 14 pixels along a line
  computed from a radius it had guessed, found one blended pixel, and reported a staircase. It now
  reads the corner radius from the DPI the way the add-in scales it, samples the whole corner box, and
  counts both the cut-away quadrant and the blended edge between it and the card: 144 pixels, 27 well,
  21 blended.

**Word's own drop shadow sits on top of our strip, and a colour read up there is Word's shadow, not
ours.** `RESULT-strip.md` recorded the `DropShadow` child overlapping our top nine pixels as harmless
and "worth remembering when the strip is drawn properly". This is that moment. A scan taken six pixels
below the strip's top edge — which looks like the obvious place to read the well — came back
RGB(61,61,61) for a well that is RGB(15,15,15). The suite now reads the well from the empty run to the
*right of the `+`*, at mid-height, where nothing of Word's overlaps it.

**One `soak-stack` failure was a run that began with the previous suite's Word still up.** Backstage
failed to open, the suite reported 19 checks instead of 22, and Word then held its scratch files open
so the next run could not even start. This is exactly the trap the project notes already warn about —
a suite that begins with leftovers measures something else. From a clean start: 22 of 22.

**`check-menu` reported its `Set-Dirty` warnings again** (`the keystrokes did not reach the document`)
and still passed 32 of 32, because the assertions that needed a dirty document got one by another
route. That is the known fragility of that helper rather than anything this slice touched, and it is
still the first thing to suspect if that suite ever fails oddly.

---

## Measured

- **Word's chrome, sampled from the ribbon at 144 dpi:** RGB(41,41,41) at 98–99% of a 150-pixel scan
  with Windows dark; RGB(255,255,255) at 100% with Windows light. Both flat enough that the 75%
  agreement threshold is not close to being the binding constraint.
- **The derivation, checked against the screen rather than against itself:** ribbon RGB(41,41,41) →
  well RGB(15,15,15), a step of **26 levels**, against the code's 26. After a live flip to light:
  ribbon RGB(255,255,255) → well RGB(229,229,229), **26 levels** again.
- **A live theme change, Word running throughout:** well RGB(15,15,15) → RGB(229,229,229), and the log
  says `palette sampled from Word's ribbon` rather than the fallback. Two theme changes in one probe,
  dark → light → dark, both followed within one janitor sample.
- **The corner at 192 dpi (this rig, 200% scaling):** radius 12 physical pixels. Of the 144 pixels in
  the corner box, **27 are the well** (the quadrant cut away) and **21 are blended** — neither the
  well nor the card. With `TabStyle=0`, the same box: **0 well, 0 cut away.**
- **The hairline:** RGB(67,67,67) beside the active card, RGB(41,41,41) — the card itself — underneath
  it.
- **Hover is local:** hovering tab 0 changes 6108 pixels on tab 0 and **0 pixels on tab 1**. Hovering
  the close button of tab 1 (`check-tabs`): 256 pixels on the button, **0 on tab 0**.
- **The lift is local:** a tab carried a third of a slot changes 6142 pixels where it was and **0 on
  the last tab in the row** — the shadow reaches 4 logical pixels and the nearest thing anything
  asserts about is a whole tab away.
- **The menu:** 330×284 at 144 dpi, background RGB(43,43,43) at 100% of the scan, against the system
  menu colour RGB(240,240,240) it used to be. Seven entries, two separators, `Close Others` greyed,
  every label still readable through `GetMenuString`.
- **The empty row:** 275 inked pixels in the 180 physical pixels to the right of the `+`, and **0** in
  the 180 pixels beyond that. Clicking the message opens nothing.
- **Cost:** `check-strip`'s "the strip stayed quiet during the drag" assertion still passes — the
  compositing added no logging to the paint path, and the back buffer is built once per strip rather
  than once per paint, which is what keeps `DragMove`'s repaint-every-strip-per-mouse-move affordable.

---

## Switches

`HKCU\Software\WordTab\TabStyle` (DWORD, default 1). Set to 0 and the strip is what it was before
this slice: flat rectangles with a full border, an aliased one-pixel `x` and `+`, a bare empty row,
and a context menu drawn by the system in the system's colours. Bisectable like every other piece.
The palette is still derived with the switch off, because that part is a correction rather than a
style — the tabs are the right colours, just not shaped. `check-look` restarts Word with it set and
asserts the corner that was carved is square again.

`HKCU\Software\WordTab\TabThemeSample` (DWORD, default 1). The sampler on its own switch, because it
is the one part of this that reads pixels out of a window belonging to Word. Set to 0 and the palette
comes from the Office theme registry value the way it always did. **This is the setting to try first
if the strip is ever the wrong colour on a machine this has not been run on** — the work rig, for
instance.

---

## What this slice does not do

- **No icons on the tabs, and no modified-document indicator.** A dot for unsaved changes is a real
  feature and wants the document's dirty state, which is a different kind of work from painting.
- **The tab name is still Word's window title.** With a `.rtf` open that reads
  `Quarterly report.rtf - Compatibility Mode`, which eats most of a tab. Only the ` - Word` suffix is
  stripped today. Worth revisiting, but it is a title-parsing change rather than a look change and
  every check script reads those titles.
- **No animation.** Hover and selection change between frames; there is no per-strip clock and
  `SetHot` repaints only when what is under the pointer actually changes, which is what keeps the
  strip quiet.
- **Word's own `DropShadow` still overlaps the top twelve physical pixels of the strip.** Harmless —
  it draws over the well, above the cards — but it is why nothing should ever read a colour from up
  there again.
- **The `+` still sits on top of the last tab once the row overflows** (`RESULT-startscreen.md`). Not
  touched here, not made worse: the `+` is drawn as a glyph and a hover chip, not as a filled card, so
  it is no more inviting than it was. Still the strongest argument for a scrolling tab row.
- **A DPI change is still untested.** This slice put the corner radius, the glyph stroke, the chip
  radius, the shadow spread and the back buffer's size on the DPI path, and collapsed the three places
  that rebuild DPI-derived state into one — so it is better organised than it was and no better
  tested. This rig has one monitor.
- **The light palette was exercised by flipping Windows' app theme, not Office's.** Office's own
  `UI Theme` value cannot be driven from outside — Word overwrites it — so Colorful, Dark Grey and
  Black as distinct Office settings have never been seen. The derivation does not care, since it reads
  the colour rather than the setting, but it has not been watched doing so.

---

## Running it

```
pwsh -File tools\check-look.ps1
pwsh -File tools\check-look.ps1 -KeepOpen -Screenshot
pwsh -File tools\check-look.ps1 -NoThemeSwitch
```

`check-look` is the first suite that reads **absolute colours** rather than counting changed pixels,
and the first that changes a machine-wide setting: it flips Windows' app theme to light and back to
prove the palette follows Word without a restart. The restore is in the `trap` as well as at the end,
so a crash mid-run does not leave the desktop somewhere the user did not put it. `-NoThemeSwitch`
skips that section.

Because it reads absolute colours it can fail for a reason the other suites cannot — a window over
Word photographs as a perfectly plausible restyle — so every colour assertion has a control taken from
the same photograph: the ribbon for the well, an inactive tab for the active card, the system menu
colour for the menu, the empty row's far end for its message.
