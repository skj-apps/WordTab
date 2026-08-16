# Why there are no icons on the tabs

**Not a RESULT.** This records a slice that was built, measured, and **thrown away on 2026-08-16** —
so that the next triage does not spend another ninety minutes re-deriving it. "Icons on tabs" had been
on the candidate list for three triages as "the last purely visual gap". It is now closed.

The code is gone. The measurement is kept: `tools\probe-icons.ps1`.

---

## The measurement that should have ended it before any code was written

`probe-icons.ps1` asks the shell for the icon of nine extensions, renders each into a DIB and hashes
the pixels. On this rig:

| icon | extensions |
|---|---|
| `D056C39A` | **`.docx`, `.doc`, `.rtf`** |
| `33F33519` | `.odt`, `.dotx`, `.docm`, *(no extension)* |
| `1F88ED40` | `.txt` |
| `3B4270ED` | `.pdf` |

**Every real Word document gets byte-identical artwork**, and `.docm` / `.dotx` come back as the
*generic blank page* rather than a Word document. So:

- **A per-file-type icon buys nothing.** It would differ only for `.txt` and `.pdf`, and would be
  actively worse for macro-enabled documents.
- **A single icon carries no information at all.** If every tab shows the same picture, the picture
  distinguishes nothing. The *name* is the only thing on a tab that identifies its document.

That second point is the whole argument, and it was under-weighted the first time round because
"icons make a row read as tabs" is true and appealing. It is also purely decorative — and this
decoration is not free.

---

## An icon is the only thing this project ever drew that takes width from something useful

The unsaved-changes dot lives inside the close button's existing square and costs nothing. An icon has
to come from somewhere, and the only place is the name.

Built and photographed at 200% on the dev rig, five documents in a default-sized Word window:

```
[icon] Qua...  x   [icon] Me...  x   [icon] Draf...  x   [icon] Bud...  x   [icon] App...  x
```

Identical icon on every tab; every name destroyed. A rationing rule was then invented to contain the
damage — *the icon may cost the name at most one part in three of what it would otherwise have had* —
which worked, and produced this:

```
[icon] Quarterl... x   [icon] Meeting... x   [icon] Draft pr... x   [icon] Budget ... x
```

**But needing a rationing rule at all is the finding.** The feature's own cost had to be capped
against the feature it was damaging, to deliver a picture that is the same on every tab.

---

## And it caused a real regression

With the icon code installed, `tools\check-title.ps1` failed 3 of its checks: the **`+` button's**
computed centre landed at y=464, inside Word's ribbon (`NetUIHWND`, which ends at y=470) instead of
inside the strip, so the click never landed and no new document opened. The other eleven suites were
green.

**Confirmed by bisect, not by argument** — which matters, because the reasoning pointed the other way
and was wrong. `ComputeLayout` and `tools\WordLayout.cs` were both untouched by the icon work, and a
standalone six-document diagnostic put the `+` centre correctly at y=502 with `WindowFromPoint`
returning `WordTabStrip`. It looked like somebody else's flake. It was not:

| binary | `check-title` |
|---|---|
| with the icon code | **34 checks, 3 FAILED** |
| same tree, `strip.cpp` stashed, rebuilt, reinstalled | **35 checks, all passed** |

**The cause was never diagnosed.** The strip's position differed by 38 physical px in that scenario,
and `check-title` is the only suite whose fixtures include a **Protected View** document — which lives
in a second `WINWORD` process. Whether drawing an icon perturbs that, or whether the once-per-process
`SHGetFileInfoW` call inside the first `WM_PAINT` does, is unknown. **If icons are ever revived, start
here**, and suspect the shell call on Word's UI thread before the drawing.

---

## What was kept

- **`tools\probe-icons.ps1`** — the evidence. Also measured two things worth having if any future
  slice wants shell artwork: `SHGetFileInfoW` with `SHGFI_USEFILEATTRIBUTES` answers from the *name*
  and never touches the disk (safe for documents on corporate shares), and `DrawIconEx` blends a
  32-bit icon onto the strip's own top-down 32-bpp DIB correctly on both light and dark — the corner
  pixel stays the background colour, so no hand-rolled compositing is needed. 0.147 ms per lookup.
- Two test-harness lessons that outlived the feature and belong to whatever suite meets them next:
  - **ClearType defeats a naive colour probe.** Word's UI font is subpixel-antialiased, so every glyph
    edge carries a fringe reaching a blue-minus-red of about 100. A probe thresholded at 40 reported a
    square of plain tab *name* as 11% "blue" and a tab with no icon at all as 5%.
  - **Word treats a document's view as part of the document, so resizing a window dirties it.** A
    suite that sweeps window widths will meet a save prompt on documents nothing ever typed into.

---

## What to do if this comes back

Re-run `probe-icons.ps1` first. If the shell still returns one image for `.docx`, `.doc` and `.rtf`,
the answer is still no, and it is still no for the same reason: **a mark that is identical on every
tab is not information, and on this row it is paid for in the only thing that is.**
