# Spike 1 result: PASS

Run on the dev rig, 2026-08-14. M365 x64, Word 16.0.20228.20190, 150% DPI.
Office Tab trial disabled for the run.

## Question

Can we carve a 32px strip out of Word's own layout — shrink `_WwF` from the top and
put our window in the gap — and have it stay correct through resize, maximize/restore
and Backstage (the full-screen File menu)?

Backstage was the specific risk: it is reported to have broken the old "Doc Tabs"
add-in, because leaving it makes Word rebuild its layout from scratch.

## Answer

Yes.

Measured live, mid-session, with the spike running:

```
ribbon (NetUIHWND)   y  72 .. 250   h=178
WordTabStripSpike    y 250 .. 282   h= 32   <- ours
_WwF document frame  y 282 .. 786   h=504
status bar           y 786 .. 808   h= 22
```

Ribbon bottom meets strip top at 250. Strip bottom meets document top at 282.
Contiguous, no gap, no overlap — structurally identical to the stack Office Tab
produces (ribbon 178, strip 32, then the document).

Backstage behaviour, observed by hand: opening File hides the strip (correct — it is
a child of `OpusApp` and Backstage covers the client area); leaving Backstage brings
it back in the right place.

## The finding that actually matters

Word does not leave `_WwF` alone. From the log:

```
relayout: _WwF natural (0, 78 724x636)  -> (0,110 724x604)
relayout: _WwF natural (0, 78 1081x636) -> (0,110 1081x604)
relayout: _WwF natural (0,178 1081x536) -> (0,210 1081x504)   x6
```

Word resets the frame to its own natural rect repeatedly — six times at the tail of
the session. So the Doc Tabs failure mode is real and reproducible.

What survives it is the sync being **idempotent against Word's layout** rather than a
repeated delta: on each pass we read `_WwF`, and if it is not exactly where we last
put it we treat what we see as Word's natural rect and shift from there. Nudging by
-32 each time would march the frame off the bottom of the window within seconds.

Six resets, six clean recoveries, no thrash in between.

## What this does NOT prove

- **In-process.** This runs outside WINWORD, watching via `SetWinEventHook` plus a
  30ms poll. There is therefore a sub-frame window where `_WwF` sits unshifted before
  we correct it. An in-process subclass handling `WM_SIZE` synchronously closes that
  gap and cannot lose the race. The out-of-process design was a shortcut to isolate
  the geometry question from COM registration and the missing C++ toolchain — not a
  design decision.
- **Multiple windows.** The spike drives one `OpusApp`. Real tabs mean N stacked
  windows kept in geometric lockstep.
- **Z-order and taskbar.** Untouched. N stacked windows produce N taskbar buttons
  until `ITaskbarList::DeleteTab` suppresses the inactive ones.

## Reproducing

```
dotnet build -c Release
bin\Release\net9.0-windows\win-x64\StripSpike.exe
```

Disable the Office Tab trial first (Word > File > Options > Add-ins >
Manage: COM Add-ins > Go > untick, then restart Word). It manipulates the same `_WwF`
and the two are otherwise indistinguishable.

Ctrl+C restores Word's layout on exit.
