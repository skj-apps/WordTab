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
  somebody else's software off on their machine is their call. And the honest framing is a warning
  rather than a blocker: WordTab and Office Tab have coexisted through every check suite on the dev
  rig without the strip's tripwire firing once.
- **`README.txt` is written for someone with no context and no me.** Ordered steps, the expected
  outcome stated so a wrong one is recognisable, and every failure path paired with the thing to
  check — `LoadBehavior` rewritten to 2, `Resiliency\DisabledItems`, and where the log is.
- **The load banner is on by default in a package**, and the README says it will appear. It is the
  one thing that proves the add-in loaded before any document is open, and the first run is exactly
  when that is worth a click.
