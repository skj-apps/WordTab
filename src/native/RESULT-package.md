# WordTab can be installed on a machine with no compiler

**Passed 2026-08-16, commit `1717b56`.** No product code changed — `src\native\*.cpp` untouched, this
is `install\` and `.gitignore` — so the DLL is byte-identical to the one the 457-check battery
tested.

## The problem, found by reading the installer rather than remembering it

The work-rig install was scheduled by the user for the morning of 2026-08-17. **As the repo stood, it
could not have happened**, and nothing in the project's notes said so:

- `install.ps1` **builds from source by default**, and the work rig has **no C++ toolchain**.
  w64devkit is portable and admin-free, but it is 577MB unpacked / ~90MB downloaded from GitHub, on a
  locked-down corporate machine that may not permit the download at all.
- **`src/native/build/` is gitignored**, so however the repo gets over there — clone, zip, OneDrive —
  the DLL does not come with it.
- `-SkipBuild` exists, and **on a cloned repo it was guaranteed to throw.** Its staleness guard
  compares the DLL's `LastWriteTime` against the newest `.cpp/.h/.def`, and a `git clone` stamps
  every source file at the moment of the clone — which is always newer than a DLL copied in
  beforehand. It passes on the dev rig only because these files' real timestamps happen to line up.
- **`install.ps1` had never been run under Windows PowerShell 5.1**, which is the only shell the work
  rig is guaranteed to have. `tools\workrig-recon.ps1` carries the comment *"Windows PowerShell 5.1
  compatible"*, so the project already knew not to assume `pwsh` 7 was there — and then assumed it
  everywhere else.

**None of this is visible from the outside.** Every one of these fails at a different point on the
morning it matters, on a machine with no toolchain and nobody to diagnose it.

## What was built

`install\package.ps1` assembles **this repo with the compiler-shaped hole filled in**:

```
WordTab-<date>-<commit>\
  PAYLOAD.txt                   which build this is, and the SHA256 to prove it
  README.txt                    what to do, in order, and every failure path with its check
  install\install.ps1           VERBATIM COPY
  install\uninstall.ps1         VERBATIM COPY
  src\native\wordtab.h          the CLSID install.ps1 checks itself against
  src\native\build\WordTab.dll  the build
```

113KB zipped. Small enough to email, and transport-agnostic — USB, OneDrive, attachment.

**The scripts are copied, never rewritten, and that is the whole design.** A package carrying its own
installer would be a second implementation of *how WordTab is registered*, and the two would agree
right up until the day one of them was fixed. This project already has that scar:
[[one-copy-of-what-decides-truth]] — three implementations of "is Word asking something", none
correct. `install.ps1` recognises a package by its `PAYLOAD.txt` and skips the build; the payload's
directory layout is exactly the repo layout it already expects, so nothing about it is special-cased
beyond that one test.

## The two changes to `install.ps1`, and why a repo does not mind either

**Provenance by hash, not by timestamp.** In a source tree the mtime comparison is right and stays.
In a package it is answering a question it cannot answer: there are no sources to compare against,
and a zip, an email or a clone rewrites file times anyway. So a package verifies the DLL's **SHA256
against the manifest** instead. That is not an extra check bolted on — it *replaces* the proxy with
the thing the proxy was standing in for. The proxy's whole purpose is "are these the bytes built from
these sources"; the hash answers it directly, and cannot fail on a correct payload the way the
timestamp can.

**`Unblock-File` on the installed DLL.** A package that arrives by OneDrive, browser or email carries
a mark-of-the-web, and `Copy-Item` carries the stream across into `%LOCALAPPDATA%`. Policy refusing to
load a marked DLL looks *exactly* like a bad registration, three steps from the cause. **Driven, not
reasoned**: a `Zone.Identifier` was written onto the DLL inside an extracted zip, and the installed
copy came out with `:$DATA` and nothing else. It costs one call and nothing at all when the file never
left the machine.

Two smaller ones: the other-tabbed-add-in warning is now **one block** rather than one warning per
registration — the dev rig has three of those ProgIds across two hives and produced five consecutive
`WARNING:` lines, which after a successful install reads as a failed one — and the closing message no
longer promises a load banner when `-NoBanner` was passed.

## How it was proved

The round trip was driven **from an extracted zip, under Windows PowerShell 5.1, with a
mark-of-the-web on it** — the exact shape of the morning it is for, not an approximation of it:

```
install (5.1, from the zip)   -> manifest hash matched, MOTW cleared, CoCreate succeeded
Word, two documents           -> load banner appeared (#32770 865x497 'WordTab')
                                 two frames, both with strips, sharing one rectangle
                                 Get-StripPlacement: OK, 2 of 2 measured
                                 6 checks, 0 failures, photographed
uninstall (5.1, from the zip) -> no residue: neither directory, neither hive, no settings key
rebuild + reinstall from repo -> WordTab.dll MD5 8DA909D3C0A1570A84A0B14368C9B2FE, UNCHANGED
```

**The build is reproducible**, and that last line is worth more than it looks: it means the DLL going
to the work rig is provably the same bytes the green battery tested, rather than a rebuild that is
merely believed to be equivalent.

## Decisions worth not re-deriving

- **The package is gitignored, deliberately.** `dist/` holds a build output like any other, and a
  committed binary would drift from the tree it claims to be. The zip is the transport; the manifest
  is how it stays identifiable once it has left.
- **`package.ps1` refuses to be vague about what it packaged.** A dirty working tree labels the
  package `<commit>+dirty` **and warns that nobody will be able to reconstruct it from that label**.
  A package is a thing that leaves the machine and comes back as a bug report, and "the tip of main,
  probably" is not an answer to "which build is that".
- **The staleness check moved rather than being deleted.** `package.ps1` runs the same mtime
  comparison at packaging time, where the sources are still present, and refuses to package a stale
  DLL. Past that point the DLL travels without its sources and the check has nothing to stand on —
  so it happens at the last moment it can mean anything.
- **The installer warns about other tabbed-Word add-ins; it does not disable them.** Turning
  somebody else's software off on their machine is their call. ~~And the honest framing is a warning
  rather than a blocker: WordTab and Office Tab have coexisted through every check suite on the dev
  rig without the strip's tripwire firing once.~~ **That second sentence was false and was retracted
  on 2026-08-17 — see "What the installer says about other add-ins" below.**
- **`README.txt` is written for someone with no context and no me.** Ordered steps, the expected
  outcome stated so a wrong one is recognisable, and every failure path paired with the thing to
  check — `LoadBehavior` rewritten to 2, `Resiliency\DisabledItems`, and where the log is.
- **The load banner is on by default in a package**, and the README says it will appear. It is the
  one thing that proves the add-in loaded before any document is open, and the first run is exactly
  when that is worth a click.

---

# What the installer says about other add-ins, and why it used to be wrong

Measured 2026-08-17, the morning of the first install on a real user's machine. **No product code
changed** — `install\install.ps1` and `install\package.ps1` only, so the DLL stays byte-identical to
the build the last green battery tested (MD5 `9C5015588D4838D7EA27733B96DE5D53`).

## The claim that started it

The installer told the user, and this file repeated it:

> That is not known to be a problem - WordTab and Office Tab have coexisted through every check
> suite on the dev rig

**That sentence was read off a registry key, not off a running Word, and it is false.** The dev rig
has `OfficeTab.TabsforWord2013` and `TabsforOfficeHelper.Helper` at `LoadBehavior=3`, which is what
"is installed and enabled" looks like from the registry. It says nothing about whether Word loaded
them.

It matters because the work rig is a machine where **Office Tab is a tool the user actually uses**,
and the installer was about to tell them the combination was proven.

## What is actually true on the dev rig

Word's window tree, two documents open, WordTab installed:

```
--- frame 0x1D0630 ---
vis  NetUIHWND                (0,0 874x356)      <- the ribbon
vis  _WwF                     (0,420 874x221)
vis  WordTabStrip             (0,356 874x64)     <- ours, and the only one
==> Office Tab processes
    none
==> Modules loaded into WINWORD matching Tab
    C:\Users\skj81\AppData\Local\Programs\WordTab\WordTab.dll
```

No `TChromeTabs`, no Office Tab process, no Office Tab DLL in `WINWORD` at all. **Office Tab has
never once been loaded during a check suite.** The reason is two entries in

```
HKCU\Software\Microsoft\Office\16.0\Word\Resiliency\DisabledItems
    A24457 -> c:\program files (x86)\extendoffice\office tab\tabsforoffice64.dll
              office tab (classic) 19.00
    A39205 -> c:\program files (x86)\extendoffice\office tab\tabsforoffice32.dll
```

written **2026-08-15 at 20:32:37**, mid-development, at the same second Word rewrote
`OfficeTabs.Connect` to `LoadBehavior=2`.

**And the cause of that is NOT established.** Two stories fit the timestamp — WordTab and Office Tab
colliding, or one of the suites' `Kill()` calls (live in the tree until the harness slice removed
them) taking Word down while Office Tab was loading, which is exactly how Word decides an add-in
hung. **It is not written down as a WordTab defect, because nothing measured says it is one.**

## The attempt to measure the collision, and why it failed

With the user's agreement, Office Tab was fully re-enabled and Word restarted, WordTab switched off
as a control: both `DisabledItems` entries cleared, `LoadBehavior=3` on all three of its ProgIds, its
own `Expired=0 ExpiredDays=30`, its own `Office2013\TabsforWord\Enable=1`, and finally
`OfficeTabLauncher.exe` started by hand because nothing autostarts it here.

`TabsforOffice64.dll` **did** get into `WINWORD` — so this was not a registration problem — and it
**still drew nothing**: two separate `OpusApp` frames, `_WwF` at full height with no carve, no
`TChromeTabs`, screenshot confirming plain Word.

**So the collision cannot be measured on this rig at all**, and the honest report is "untested".
Everything was restored: both `DisabledItems` values back, `LoadBehavior` 3/2/3/3 as found, no Office
Tab process left running, no `Run` key added.

## What the installer does now

Three defects, all of them in the *reporting* rather than the install, all found by looking at a
machine instead of at the code:

- **Disabled Items are NAMED, not counted.** Each value is a fixed-width binary blob holding the DLL
  path and the vendor's friendly name as NUL-separated UTF-16, so the question "is it *us* Word
  disabled, or somebody else" can be answered. Those are different situations with different fixes
  and the old code printed the same warning for both. Only the `.dll`/`.vsto`/`.xll`/`.wll` fields
  are kept — keeping the friendly names would double the count and give the path matching a second
  thing to trip on.
- **`HKCU` beats `HKLM` for the same ProgId.** Word reads the per-user value and stops. The old code
  scanned both hives and reported a hit in either, so the dev rig — `HKLM` 3 under `HKCU` 2 — was
  told an add-in would load that Word had already given up on.
- **Disabled Items beats `LoadBehavior`.** A rival Word has killed stays dead at `LoadBehavior=3`,
  and matching it needs the ProgId resolved through its CLSID to a **path**, because the name in
  Disabled Items is the vendor's and not the ProgId.

The result on this rig, which is now three true statements where there was one false one:

```
    Word has not disabled WordTab (2 other entry/entries in Disabled Items)
      disabled: c:\program files (x86)\extendoffice\office tab\tabsforoffice64.dll
      disabled: c:\program files (x86)\extendoffice\office tab\tabsforoffice32.dll
WARNING: Another tabbed-Word add-in is registered to load with Word: TabsforOfficeHelper.Helper
    WordTab has never been tested beside a working one ...
```

## Decisions worth not re-deriving

- **"Registered to load", not "will load".** `TabsforOfficeHelper.Helper` is at `LoadBehavior=3`,
  absent from Disabled Items, and its DLL does not appear in `WINWORD`'s module list. The registry
  says what Word has been *asked* to do; claiming the stronger thing would be an assertion about the
  product made from a fact about a registry key — the same mistake the retracted sentence made.
- **An entry naming OUR OWN DLL is the one case in that block that means the install will not
  work**, so it is the one that gets a warning and the exact dialog path. Everything else there is
  information.
- **The README now says turn the other one off, and says why it is being cautious.** It also gives
  the user a way to try both anyway and a description of what failure would look like — one row of
  tabs is fine, two rows or a document that jumps is not. A warning nobody can act on is a warning
  nobody reads.
- **The rig was left exactly as found**, checked value by value rather than assumed: the `.reg`
  exports were taken before the first change and re-imported after the last.
- **This is the second time reading the installer rather than recalling it has produced real
  findings**, and it is the same lesson as the packaging slice: a verification is only as good as the
  environment it ran in. "They have coexisted" was true of a rig where the rival never ran.
