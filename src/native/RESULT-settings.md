# A README, and a way to turn things off without regedit

**Status: done 2026-08-16.** No product code changed. New `README.md` at the repository root — there
was none at all, the only file at the top level was `.gitignore` — and new `install\settings.ps1`,
which ships inside the package.

## The gap

WordTab has **fifteen switches**. Every one is a DWORD under `HKCU\Software\WordTab`, every one
defaults to on, and until now not one of them was written down anywhere a user could reach: they
existed in `.cpp` comments, in the `RESULT-*.md` writeups, and in the suites that drive them.

That matters because of where this thing is going. The whole project exists to be installable on a
corporate machine with no admin rights and nobody to diagnose it — and the answer to *"the tab row
has come out the wrong colour"* was **regedit**, on a locked-down work PC, from notes that only exist
in this repository.

## The decision

**A script and a README, not an options dialog.** A dialog is a lot of Win32 for a screen almost
nobody opens, and it would be the first piece of UI in this project that is not part of Word's own
window. The escape hatch has to be reachable and legible; it does not have to be pretty.

`install\settings.ps1`:

- **lists every switch, its current value, and one line on what it does** — in the terms somebody
  turning it off would think in, not in the terms the code uses;
- takes `-Set TabDot=0` (several at a time) and `-Reset`;
- **refuses a name it does not know**, because a typo would otherwise become a registry value that
  looks like a setting, is read by nothing, and is indistinguishable from a switch that does not work;
- says, every time, that Word reads these once at startup and has to be closed completely.

**It decides nothing, and the script says so.** The table of switch names is documentation and can go
stale; the authority is the add-in. So the script finishes by printing the add-in's own startup lines
from the log — `StripStart … tab drag=on tab scroll=on …`, `StackStart … detach=on` — and states that
if the two disagree, the log is right. That is [[one-copy-of-what-decides-truth]] applied to a thing
that could not avoid being a second copy: make the second copy visibly subordinate to the first.

## It travels with the package

`package.ps1` copies it verbatim alongside `install.ps1` and `uninstall.ps1`. **The escape hatch has
to be in the package it is an escape from** — a machine whose strip comes out wrong is exactly the
case it exists for, and a copy that only lives in the repository is a copy the person with the problem
does not have.

The package's `README.txt` grew two sections with it: what the tabs actually *do* (which it never
said — it only ever said how to install), and the two switches most likely to be wanted, named
explicitly:

- **`TabThemeSample=0`** if the row is the wrong colour on a machine this has not run on;
- **`TabKeys=0`** if Ctrl+Tab is needed to type a tab inside a table.

The load-banner instruction in `README.txt` was a raw `New-ItemProperty` command; it is now
`settings.ps1 -Set ShowLoadBanner=0`, so the first thing a new user is told to change is also the
thing that teaches them where all the other switches are.

## The repository README

Written for somebody who has never seen this project: what it does in a paragraph, **why it is
native C++ and installs under `HKCU`** (the constraint is the reason for nearly every technical
decision here), how to install with and without a compiler, how to turn parts off, how it works in
two paragraphs, how to build and how to run the checks — including that **they take over the desktop
and cannot run in parallel**, which is the single most likely way for someone to waste an afternoon.

It ends with the known limitations, stated rather than left to be discovered: no live preview while
dragging a tab out, no dropping two lone windows onto each other, the Protected View palette
fallback, and **multi-monitor behaviour being unproven because this has only ever run on a
single-monitor machine.**
