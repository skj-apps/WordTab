# The strip at a scale this machine does not run at

**Passed 2026-08-17.** `StripOnFrameDpiChanged` had never executed — not on this rig, not on any rig,
not once in the life of the project. It does now, four times a run, at 96, 120, 144 and 192.

This is the project's named blind spot. Every scaled size in `strip.cpp` goes through one
`Scaled(logical, dpi)` helper, and until this slice that helper had only ever been called with one
number. The work rig — a laptop plus a 40" second monitor — is where a second DPI first happens, and
it is the machine nobody here can reach.

## First: can the desktop just be rescaled?

The project's notes had this filed as a blocker — *"changing the desktop's scaling needs an
undocumented API"*. That was true and it was not the end of the sentence. The API is undocumented and
it is reachable: `DISPLAYCONFIG_DEVICE_INFO_GET_SOURCE_DPI_SCALE` (`-3`) and `_SET_` (`-4`), the two
private DisplayConfig packets the Settings app itself uses. `tools\probe-dpi.ps1` speaks them.

**The answer is still no, and now it is a measurement rather than an assumption.** One display
source, relative range **`0..0`**. There is no second scale factor on offer, so there is nothing to
switch to — by this route or any other.

### The probe measured a caution about itself, and it is the more useful half

The packet returns a **success code while answering `0/0/0`**. Mapped through the percentage table
that comes out as *"current 100%"* — on a monitor Windows independently reports at **192 dpi, which
is 200%**. Both cannot be true.

The **relative indices are sound**: they are read straight out of the packet, and `0..0` is what
carries the finding. The **percentages are not**, here: they go through a hardcoded table that no API
exposes. So `-Report` cross-checks against `GetDpiForMonitor` and **prints the disagreement instead
of the number**.

An undocumented packet that succeeds while saying nothing is indistinguishable from one that
genuinely means "this monitor is at the first entry in the table" *unless something else is asked*.
So ask something else.

### And then the control needed a control

I wrote in this file's own source that `GetDpiForMonitor` "does not depend on the DPI awareness the
calling process declared". **That was an assertion, not a measurement, and it is false.**
`MDT_EFFECTIVE_DPI` is *effective* in the literal sense — it is reported through the calling
process's DPI virtualisation. The same monitor, one minute apart:

| shell | `GetDpiForMonitor` says | verdict printed |
|-------|-------------------------|-----------------|
| pwsh 7 (per-monitor aware in its manifest) | **192 dpi (200%)** | disagreement reported |
| Windows PowerShell 5.1 (not aware)         | **96 dpi (100%)**  | **"they agree"** |

**5.1 is the shell the work rig is guaranteed to have.** Unfixed, this script would have gone to the
one machine it exists for and reported *"packet 100%, monitor 100%, they agree"* — while both numbers
were wrong. **A broken cross-check that agrees is worse than no cross-check**, because it converts
"unverified" into "verified".

Fixed by declaring `PER_MONITOR_AWARE_V2` before asking; it fails harmlessly when the host already
set one, which is the pwsh 7 case. **Both shells now answer 192 and both print the disagreement.**
The lesson is the ordinary one and it caught me inside the very slice about it: *the thing you are
checking with is also a thing that can be wrong, and "it runs" is not "it agrees".*

## So the way in is an override, not a faked message

**`HKCU\Software\WordTab\TabDpi`.** Absent or 0 means "ask Windows", which is every real install.

The obvious alternative was to synthesise the real event, and it is closed: **`WM_DPICHANGED` cannot
be posted from another process.** It carries a suggested `RECT` **by pointer**, and the frame proc
chains it on to Word — which would dereference an address from the wrong address space. This is
written down in three places now because it is the trap a future reader walks back into.

`DpiOverride()` is read **on every call rather than cached at `StripStart`** like every other switch,
and the re-reading is the point rather than an accepted cost: it is what lets a value written into
the registry mid-run change the answer. The janitor notices the change and calls
`StripOnFrameDpiChanged` — the same function the real message reaches. **The trigger is stood in for;
everything downstream of it is real.**

Three call sites: a strip being built, a DPI change, and the janitor's watch. `g_dpiWatch` is set at
`StripStart` **only if `TabDpi` already held a usable value**, so a machine that has never set it
never reaches that line and never pays the registry read. **Production behaviour and production cost
are both exactly unchanged.**

## What it drives that nothing else did

`StripOnFrameDpiChanged` re-derives the height, the font, the corner radius, the glyph stroke and the
back buffer; takes down a tooltip measured at the old scale (taken away rather than rebuilt — the
pointer has just crossed monitors, which is not a hover); and un-applies the `_WwF` shift it made at
the old scale so the next layout can re-apply at the new one. All of it was reasoned. None of it was
run.

`check-strip` **40 → 67 checks** in the battery (70 run standalone, where an extra frame is left over
from a previous suite and the "every frame has its own strip" section fires). The section owns its
own Word, because the override has to be in the registry *before* `StripStart` reads it.

| dpi | strip height | strip bottom = document top |
|-----|--------------|------------------------------|
| 192 | 64px | 420 |
| 96  | 32px | 388 |
| 120 | 40px | 396 |
| 144 | 48px | 404 |
| 192 | 64px | 420 |

The height is the obvious assertion and the weakest one. **The seam is the real one** — the strip has
to still meet the document frame exactly and still sit directly under Word's chrome, with no gap and
no overlap, at every size. A wrong scale shows up there.

Three things the section is careful about:

- **It starts at the DPI the machine really is.** That first step is a control, not a restatement:
  nothing about the strip differs from an ordinary run at that point.
- **The add-in must *say* it is overridden.** Without that assertion the whole section could pass
  having measured a strip that was never overridden at all — every height below is compared against a
  number computed from the same DPI, so an override that silently did nothing at 192 would look
  identical to one that worked.
- **The transition must be reported by the function under test.** A height that happens to be right
  because Word rebuilt the window is a different thing from `StripOnFrameDpiChanged` having run.

And the restore is asserted rather than assumed: the value is removed in a `finally`, Word comes back
with no override, and the strip must be at the scale Windows reports with the add-in no longer
claiming otherwise. **A `TabDpi` left behind would make every suite after this one in the battery run
against a strip built at a scale the machine is not at — silently**, because the add-in only mentions
the override on a startup line nothing downstream reads.

## The second copy of "what DPI is this"

`WordLayout::Dpi` had to learn the same override, and this is not tidiness. It is the harness's copy
of `ComputeLayout` — same pad, gap, minimum, desired, plus, close and chevron constants — and its
whole job is to predict where the add-in put a tab so a suite can click it. **The moment the add-in
builds itself at a DPI the window is not running at, a harness that asks Windows is computing slots
for a strip that does not exist**, and every click lands somewhere else. Silently: the click still
hits the row, just the wrong tab.

**It was written and then never executed.** `check-strip` asked for the DPI once, deliberately
*before* the override was written, and no other suite runs with `TabDpi` set — so the second copy was
dead as far as the battery was concerned. It now gets an assertion inside the loop, and it needs one,
because **its failure is silent in the worst way**: a bad `RegGetValueW` p/invoke returns non-zero,
`ForcedDpi` answers 0, and `Dpi` falls through to asking Windows — which on this rig returns 192, the
same number the override is set to for two of the five steps. **Only the 96, 120 and 144 steps can
tell a working read from a dead one.**

`RegGetValueW` rather than `Microsoft.Win32.Registry`, for a harness-wide reason: naming any assembly
in `Add-Type` replaces PowerShell's default reference set, so every suite passes its own list, and
reaching for the `Registry` class would mean editing all of them to add one more.

## What this does not cover, stated rather than discovered later

- **The trigger.** In production the trigger is `WM_DPICHANGED` arriving from Windows. What is
  covered is everything downstream of the two lines in `frames.cpp` that receive it.
- **Two monitors at once.** A DPI *change* is now driven; a stack straddling two scales is not, and
  cannot be here. The code review remains reassuring rather than conclusive: every size goes through
  the one `Scaled` helper, `WM_DPICHANGED` → `StripOnFrameDpiChanged` rebuilds them in one place, and
  tear-off placement uses `MonitorFromWindow`/`MonitorFromRect` with `MONITOR_DEFAULTTONEAREST` — no
  `SPI_GETWORKAREA`, no `SM_CXSCREEN`, none of the primary-only APIs. **The obvious multi-monitor bug
  is not there. That is not the same as there being none.**
- **`TabDpi` is not a user setting** and is not offered as one. It makes the add-in disagree with the
  machine on purpose, which is why the startup line announcing it is written **only when it is set** —
  a log that mentioned it on every ordinary startup would train the reader to skip the line that
  matters.

## The one thing worth carrying to the work rig

`probe-dpi.ps1 -Report` there answers the question this rig cannot: **two monitors, two scale
factors, and a real relative range with something in it.** It changes nothing and it is one command.
If the strip misbehaves at the wrong size on that machine, `TabDpi` is also the escape hatch that
pins it to a known-good scale while the real cause is found.
