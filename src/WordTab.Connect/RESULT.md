# Slice: get a COM add-in to load into Word

**Date:** 2026-08-14
**Outcome:** the add-in does not load, and the reason is architectural rather than a bug.
**Headline: a managed (.NET) COM server cannot be activated from a per-user HKCU registration.
Since WordTab must install without admin rights, its COM entry point has to be a native DLL.**

This retires the plan of record — "build the add-in against .NET Framework 4.8.1 and register it
under HKCU" — on measurement rather than opinion. The registration half of that plan is
impossible; the .NET Framework half is a separate question that is now open again.

## What was built

A complete, working per-user COM add-in, everything except the one thing that turned out to be
blocked:

- `WordTab.Connect.csproj` — .NET Framework 4.8.1 class library. **Builds clean with no Visual
  Studio** via `Microsoft.NETFramework.ReferenceAssemblies`, which was the flagged unknown behind
  the net481 decision. That unknown is resolved: it builds.
- `Interop.cs` — `IDTExtensibility2` hand-declared. No dependency on `Extensibility.dll`, which is
  not in the GAC on this rig and cannot be assumed on the work rig either.
- `Connect.cs` — the COM class. Logs every callback, never lets an exception cross the COM
  boundary, and shows a one-shot "WordTab is loaded" banner on a background thread.
- `Log.cs` — best-effort file log at `%LOCALAPPDATA%\WordTab\wordtab.log`.
- `install\install.ps1` / `install\uninstall.ps1` — per-user install to
  `%LOCALAPPDATA%\Programs\WordTab\` plus HKCU registration, copying the shape measured on the
  work rig's own Office Tab install. No admin, no HKLM, no `Program Files`.

The code is sound and mostly survives the pivot. What changes is *how the class is reached*, not
what it does once reached.

## How the failure was found

`install.ps1` ends by activating the class outside Word (`CoCreate` from Windows PowerShell). That
check cost a second and caught the problem before Word was ever launched, which is worth keeping:
it separates "the registration is wrong" from "Word declined to load us", two different failures
with different fixes.

    CoCreate('WordTab.Connect') failed: 0x80070002 (ERROR_FILE_NOT_FOUND)

Everything else checked out, which is what made it interesting:

- The DLL and type load and construct fine — proven by loading the assembly directly and calling
  the constructor, which wrote its log line.
- The registration matched, key for key, what `RegAsm.exe /codebase /regfile:` generates. It was
  compared against that reference output rather than eyeballed.
- The keys were visible in the merged `HKEY_CLASSES_ROOT` view.

Falsifying the obvious explanations:

| Change | Result |
| --- | --- |
| `InprocServer32` = full path to `mscoree.dll` vs bare `mscoree.dll` | same error |
| `CodeBase` pointing at a real file vs a nonexistent one | **same error** |
| Activating from the install directory | same error |
| A deliberately bogus `Class` value | same error |
| A deliberately bogus `RuntimeVersion` | same error |

A bogus `RuntimeVersion` producing an identical failure is the tell: the CLR shim was never
reading these values at all.

## The controlled experiment

Three activations, in `tools\probe-managed-com-hkcu.ps1`, which reproduces all of this in about
five seconds:

| Test | Result |
| --- | --- |
| **Native** DLL (`scrrun.dll`) under an invented **HKCU** CLSID | `0x80040111` `CLASS_E_CLASSNOTAVAILABLE` |
| **Managed** class (`System.Collections.SortedList`) under an invented **HKCU** CLSID | `0x80070002` |
| **Managed** class, same type, via .NET's own **HKLM** registration | **activates** |

The native result is a pass, not a failure: `CLASS_E_CLASSNOTAVAILABLE` can only come from a DLL
that COM found, loaded and called. So per-user COM registration works fine on this machine — for
native servers.

The managed pair is the finding. Identical class, identical values, complete registration, a GAC
assembly so no `CodeBase` is involved and assembly resolution cannot be blamed. The only
difference is the hive, and only HKLM works.

**Conclusion: the CLR's COM activation path does not see per-user registrations.** The precise
mechanism was not pinned down — the suspicion is that it opens `HKEY_CLASSES_ROOT` in a way that
suppresses the HKCU merge — and it was not chased further, because no amount of understanding it
gives us write access to HKLM on the target machine.

## Confirmed in the real host

Word behaves exactly as the probe predicts. Launching it with the add-in registered:

- nothing in the log — our class was never constructed;
- Word rewrote `LoadBehavior` from `3` to `2`, its record that loading was attempted and failed.

## Why the obvious escape routes are closed

- **Register under HKLM.** Needs admin. This is the one constraint the project cannot trade away.
- **Modern .NET `EnableComHosting`.** Genuinely promising — it generates a *native* `comhost.dll`
  that reads an embedded CLSID map instead of the registry, so HKCU registration would work. It
  fails on a different axis: `NETSDK1128: COM hosting does not support self-contained deployments`.
  Framework-dependent is the only supported form, and the work rig has no .NET runtime installed
  at all. Tested anyway, and the failure was informative — activation reached `comhost` and died
  inside it with a **hostfxr** error (`0x80008093`), not a COM error. COM had already found and
  called our native DLL from HKCU. That is the positive proof that a native shim works.
- **Bootstrapping the CLR from VBA** via the framework types .NET registers under HKLM. Rejected
  deliberately. It is the load path application-control bypasses use, it would sit badly on a
  managed corporate machine with AppLocker logging, and it is fragile. Not worth doing.

## Where this leaves the architecture

The COM entry point must be native. Word's `Addins` key and the HKCU install location are
unaffected — the blueprint read off the work rig still holds — and Office Tab, whose DLLs are
native and registered under HKCU, is consistent with everything measured here.

Two ways to produce that native DLL, and **there is no native compiler on this rig at all**
(no `cl`, `clang`, `gcc`, `zig`, or any linker), though the Windows SDK headers and libs
(10.0.28000.0) are present:

1. **Write the add-in in native C++.** What every comparable product does, Office Tab included.
   No CLR in the process, no runtime dependency on the work rig, smallest and fastest. The spikes'
   work is already Win32 calls wearing a C# coat, so little is genuinely lost in translation.
2. **Keep the C# and give the assembly native exports** (`DllExport`, an IL rewrite via
   ildasm/ilasm — both present on this rig). Preserves the existing code and needs no new
   compiler, but adds non-standard build machinery: it wants a `.sln`, a configuration wizard, and
   Mono.Cecil rewriting on every build.

Either way a compiler question has to be answered first, and a per-user LLVM/clang install
(no admin, using the already-installed Windows SDK) would satisfy option 1 without breaking the
project's own no-admin ethos.

## Reproducing

    pwsh -File tools\probe-managed-com-hkcu.ps1     # the finding, in about five seconds
    pwsh -File install\install.ps1                  # build, install, register, smoke-test
    pwsh -File install\uninstall.ps1                # remove everything

`install.ps1` will report the smoke-test failure described above. That is the expected result
today, not a regression.

**Run the probe on the work rig.** It needs no admin and takes seconds. The dev rig and work rig
run byte-identical Word builds, but this behaviour is about the CLR and the registry rather than
about Word, and a different answer there would reopen the whole question.
