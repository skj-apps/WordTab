# Slice: get an add-in to load into Word — natively

**Date:** 2026-08-14
**Outcome: PASSED.** WordTab loads inside WINWORD from a per-user install with no admin rights,
and Word drives it through the full add-in lifecycle.

This is the first product code that actually runs inside Word. It closes the gap opened by the
previous slice, whose finding — a managed COM server cannot be activated from an HKCU
registration — is written up in `src\WordTab.Connect\RESULT.md`.

## Proof

From `%LOCALAPPDATA%\WordTab\wordtab.log`, one Word session start to finish:

    ---- DllGetClassObject ----  module=C:\Users\skj81\AppData\Local\Programs\WordTab\WordTab.dll
                                 host=C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE  bits=64
    OnConnection  mode=1  app=Microsoft Word version=16.0 build=16.0.20228 documents=0
    banner: enabled=1 alreadyShown=0
    banner: thread created
    OnAddInsUpdate
    banner: thread running, calling MessageBoxW
    OnStartupComplete  OpusApp windows: total=1 visible=0 first=0x0000000000301E44
    OnBeginShutdown
    OnDisconnection  mode=0

Everything that matters is in there: our DLL loaded from `%LOCALAPPDATA%` into `WINWORD.EXE` at
64-bit, all five `IDTExtensibility2` callbacks fired in order, and the shutdown pair arrived
cleanly. `app=Microsoft Word build=16.0.20228` comes from interrogating the `Application` object
Word handed us, so we are demonstrably talking to the real host and not merely being instantiated
— and the build matches the one recorded for this rig.

Corroborating:

- `WordTab.dll` appears in `(Get-Process WINWORD).Modules`.
- `LoadBehavior` **stays at 3** after a Word launch. The previous, managed attempt was demoted to
  2 — Word's record that loading was tried and failed.
- The banner dialog appears about 0.5s into startup and stays until dismissed.
- `install.ps1`'s smoke test (`CoCreate` outside Word) now succeeds, where the managed build
  failed with `0x80070002`.
- Install → uninstall → reinstall round-trips clean, leaving no keys and no files.

## What it is

A 44KB native COM in-process server, built with GCC 16.2.0 from **w64devkit** — a portable
MinGW-w64 toolchain that unzips into `%LOCALAPPDATA%` and needs no admin rights, no installer and
no registry entries. Deleting its folder uninstalls it. That matters beyond convenience: the
build toolchain now honours the same no-admin constraint the product does.

- `wordtab.h` — the CLSID, and `IDTExtensibility2` declared by hand. No type library and no Office
  PIA, so the build does not require Office to be installed.
- `dllmain.cpp` — the GUIDs, the class factory, and the only two exports COM needs
  (`DllGetClassObject`, `DllCanUnloadNow`). `DllMain` deliberately does nothing but record the
  module handle: it runs under the loader lock, and the logger reaches for shell32 to find
  `%LOCALAPPDATA%`, which is exactly how a process deadlocks at load time.
- `connect.cpp` — the object Word instantiates. Implements `IUnknown`, `IDispatch` and
  `IDTExtensibility2` off one vtable, and routes `Invoke` DISPIDs 1–5 to the same methods, so it
  cannot matter which of the two ways Word chooses to call us.
- `log.cpp` — the file log. This is how we see inside WINWORD, where there is no console.
- `build.ps1` — the build. Statically linked; the DLL imports only `KERNEL32`, `USER32`,
  `ADVAPI32`, `OLEAUT32`, `SHELL32` and `msvcrt`, all of which ship with Windows. Run with
  `-Verify` to have that checked rather than assumed.

Registration is unchanged in shape from the managed attempt and still matches the blueprint read
off the work rig's Office Tab: files in `%LOCALAPPDATA%\Programs\WordTab\`, everything else under
HKCU. What changed is only what `InprocServer32` points at — our own DLL, rather than `mscoree.dll`
plus the managed-class values the CLR turned out never to read from the per-user hive.

## Findings worth carrying into the next slice

**At `OnStartupComplete` the `OpusApp` window exists but is not yet visible** — logged as
`total=1 visible=0`. Reproduced on every run. Any window work that waits for `IsWindowVisible`
before acting will wait forever at this point, so the next slice must either act on the invisible
window or hook the moment it is shown. This is the first thing that would have silently wasted an
afternoon.

**`documents=0` at `OnConnection`.** Word launched with no document, which is the Start-screen
case already known to need a policy of its own — the tabbed stack has nothing to put in a tab yet.

**Word calls us through the vtable, not `Invoke`.** No `Invoke dispid=` lines appear in any log,
so the dispatch path is unused so far. It stays in as insurance: it is cheap, and an Office host
that calls by DISPID would otherwise fail silently.

**Nothing may throw or escape.** The rule that shaped the managed version carries over unchanged,
which is why the build uses `-fno-exceptions` and every callback returns an HRESULT. An error
crossing the COM boundary during load gets the add-in put in Word's `Resiliency\DisabledItems`
list, which is silent, sticky, and easy to mistake for a registration fault later.

## Reproducing

    pwsh -File src\native\build.ps1 -Verify   # build, and check the exports and imports
    pwsh -File install\install.ps1            # build, install, register, smoke-test
    pwsh -File install\uninstall.ps1          # remove everything

Then start Word: a dialog says WordTab is loaded, and
`%LOCALAPPDATA%\WordTab\wordtab.log` records the session. `install.ps1 -NoBanner` turns the dialog
off without a rebuild once it has served its purpose.

The toolchain is not in the repo. To set it up on a fresh machine, download
`w64devkit-x64-<version>.7z.exe` from https://github.com/skeeto/w64devkit/releases and run it as
`<file> -o"$env:LOCALAPPDATA\Programs\w64devkit" -y`. `build.ps1` prints these instructions if it
cannot find it. Version used here: 2.9.1, SHA-256
`9208C19755CD4964B7915B9AFCF02C66D493A4C870C4B3E83F6C538D9C1237A5`.

## Not covered

Untested on the work rig — nothing has been installed there yet. Everything measured about that
machine says this should work (native COM, HKCU registration, AppLocker in audit mode, no add-in
signing policy), but "should" is not "did".

The add-in does nothing yet beyond announcing itself. Subclassing `OpusApp` and moving the window
work in-process — the whole point of getting here — is the next slice.
