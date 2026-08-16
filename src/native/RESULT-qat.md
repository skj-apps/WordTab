# The Quick Access Toolbar below the ribbon: measured, and there is nothing to fix

Queue item 9, raised by the palette slice's own audit and deliberately left unmeasured at the time.
**Measured 2026-08-16. The guess that put it on the queue was wrong in both halves, the current code
is correct in this configuration, and no product code changed.**

## What the queue said

`RESULT-palette.md` recorded this as an accepted, unmeasured cost of the new sampling rule:

> **Quick Access Toolbar below the ribbon.** If a user turns that on, Word docks a second full-width
> row between the ribbon and the document. Structurally that is the same shape as the message bar, so
> the new rule declines — but unlike the message bar, that row is painted in the ribbon's own colour,
> so the *old* rule sampled it and got the right answer by luck. On such a machine the strip now
> falls back to the Office-theme palette instead of the sampled one.

Three claims, all plausible, all written ahead of a measurement:

1. Word docks it as **a second full-width row** (i.e. a second `MsoCommandBar` chain).
2. Therefore the sampler **declines** and the strip falls back to the Office-theme palette.
3. The row is **painted in the ribbon's own colour**.

**One, two and three are all false.**

## What Word actually does

`HKCU\Software\Microsoft\Office\16.0\Common\Toolbars\Word\QuickAccessToolbarStyle`, one DWORD, read
by Word at startup. Four values driven, one Word restart each, this rig at 192 dpi:

| value | what the window looks like | chrome height |
|-------|----------------------------|---------------|
| 0     | QAT above the ribbon        | 356 |
| 16    | **this rig's own setting** — QAT collapsed to a `»` chevron in the title bar | 356 |
| 32    | QAT above the ribbon        | 356 |
| 1     | **QAT docked below the ribbon** | 426 |
| 17    | QAT docked below the ribbon | 426 |

So **bit `0x1` is the below-the-ribbon flag** and the chrome grows by **70 physical px** (35 logical)
to carry the row.

**And the row is not a second chain. It is inside the ribbon's own `NetUIHWND`.**

```
QAT ABOVE (or collapsed)                   QAT BELOW THE RIBBON
client 874x685  dpi 192                    client 874x685  dpi 192
strip  y 356..420                          strip  y 426..490
_WwF   y 420..641                          _WwF   y 490..641

MsoCommandBarDock   y 0..356               MsoCommandBarDock   y 0..426
  MsoCommandBar     y 0..356                 MsoCommandBar     y 0..426
    MsoWorkPane     y 0..356                   MsoWorkPane     y 0..426
      NUIPane       y 0..356                     NUIPane       y 0..426
        NetUIHWND   y 0..356                       NetUIHWND   y 0..426
```

**One chain in both.** Compare the Protected View window that caused the original bug, which really
does hold *two* — ribbon `y 0..156` and the message bar `y 156..318`. That is the shape
`RibbonHuntProc` declines on, and the QAT is not it.

Consequence: the hunt's three tests all pass. The `NetUIHWND` above the strip is full width, its
bottom edge is the strip's top edge, and **its top edge is 0**, so there is no intruder and nothing
declines. The add-in says so itself:

```
strip  chrome sample: ok
```

## The one thing that is true, and it is not a defect

The sample point is `ribbonRect.bottom - 3`, so with the QAT below the ribbon the sampler reads the
**QAT row** rather than the ribbon body. Those are different colours — measured in both themes, every
row of chrome scanned from the top of the window down:

```
                      DARK                          LIGHT
ribbon body           y 156..355  RGB( 41, 41, 41)   y 156..352  RGB(255,255,255)
QAT row               y 356..425  RGB( 31, 31, 31)   y 356..425  RGB(250,250,250)
```

Ten levels apart in dark, five in light — so **"painted in the ribbon's own colour" is wrong too**,
and the old rule was not right by luck; both rules pick the same window here and always did.

What the strip then derives:

```
DARK    sampled RGB(31,31,31)     well RGB(5,5,5)       card RGB(31,31,31)   edge RGB(57,57,57)
LIGHT   sampled RGB(250,250,250)  well RGB(224,224,224) card RGB(250,250,250) edge RGB(205,205,205)
```

**This is correct, and it is correct for the reason the subsystem was built on rather than by
accident.** The palette's rule is *the colour Word is painting immediately above us* (strip.cpp:317),
and with the QAT below the ribbon that colour **is** the QAT row. The active card is drawn directly
under that band and matches it exactly — the same relationship it has to the ribbon body in every
other configuration. Photographed at both themes; the light scan reads the strip's own pixels back as
`RGB(250,250,250)` for the card and `RGB(224,224,224)` for the well, which is the log's numbers
arriving on the screen.

The well is the sampled colour minus 26, so it lands at 5 where Word's dark workspace is 9, and at
224 where Word's light workspace is 228. **Four levels, in both directions, at both ends of the
range.** Invisible, and it is the same four levels the design already accepts.

## The decision

**No code change.** The queue item asked for a measurement and then a decision; the measurement says
the general rule already covers this configuration, so writing anything specific to it would be a
rule invented for a case that does not need one — which is exactly how this subsystem's last two
defects started.

**And no new suite or probe.** The durable instrument is the add-in's own `strip  chrome sample:`
line, which already distinguishes every outcome this measurement cared about; a suite for this would
cost a Word restart and a registry flip per run to guard a configuration that the general rule
handles. The reproduction is two commands and is written down below instead.

## What is worth knowing about it anyway

- **`strip  chrome sample:` is the line to read, not the palette line.** The palette line is logged
  only when the colour *changes*, and on a dark rig the fallback and the sample are the same number —
  so at `QuickAccessToolbarStyle=0` there is no palette line at all and its absence means nothing.
  The first version of this measurement grepped for `palette` and `ribbon`, found neither, and would
  have concluded the sampler had declined. It had not.
- **The sample row now lands inside a row of controls.** At 192 dpi the QAT row is 70px tall and
  `bottom - 3` is comfortably below the buttons, but a taller QAT, a different DPI or a user with many
  QAT commands could put glyphs on that row. The flatness test (three quarters of 24 points must
  agree) is what catches that, and the outcome is a decline and the fallback palette — the documented
  honest behaviour, not a wrong colour. Nothing to do; worth recognising if a machine ever reports the
  strip on the fallback with the QAT below the ribbon.
- **Geometry is fine and it was checked, not assumed.** The strip sits at `y 426..490` with `_WwF` at
  `490..641`: between the chrome and the document, which is `Get-StripPlacement`'s invariant. The
  taller chrome is a property of *every* window in this configuration rather than of one window, so
  it does not interact with the 38px Protected View fix — that fix stopped one window's chrome height
  being broadcast onto others, and here they all agree.
- **`New-Object -ComObject Word.Application` does not attach to the running Word.** It starts a
  second, invisible instance, which loads the add-in and logs its own `StripStart`. It cost one
  confusing pair of log lines here, and it leaves a document-less WINWORD behind that `Close-AllWord`
  cannot reach with `WM_CLOSE` because it has no main window. Measure Word from outside.

## Reproducing it in thirty seconds

Word closed, then:

```powershell
$k = 'HKCU:\Software\Microsoft\Office\16.0\Common\Toolbars\Word'
Set-ItemProperty $k QuickAccessToolbarStyle 1 -Type DWord   # 1 = below the ribbon
# start Word, look at %LOCALAPPDATA%\WordTab\wordtab.log for `strip  chrome sample:`
Set-ItemProperty $k QuickAccessToolbarStyle 16 -Type DWord  # this rig's own value
```

**Set it with Word closed.** Word writes its settings on exit and will undo a live edit.
