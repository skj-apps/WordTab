# The log rolls, and six of the seven readers passed for having read nothing

**Queue item 8. No product code changed** — `src\native` is untouched; this is entirely `tools\`.
The add-in's log is the instrument every suite uses to see inside Word, and seven copies of "read
the log from a mark" had grown across the eleven suites. All seven were wrong in the same place.

## What was wrong

`src\native\log.cpp:42` **deletes the log** once it passes 512KB, before every write. A full battery
writes about 250KB, so the second battery in a session rolls it — and it rolled in the middle of
`check-title` during the battery that found this.

Every copy marked its place with a **byte offset**, which after that deletion means nothing. The
copies then split into two ways of being wrong, and both are silent:

| copy | what it did after a roll | what that does to an assertion |
|---|---|---|
| `check-reorder`, `check-startscreen`, `check-scroll`, `check-look` | `if ($offset -gt $len) { $offset = 0 }` — read the whole of the fresh file | everything between the mark and the roll is gone. **"The log says nothing since the mark" goes GREEN having lost the evidence that would have failed it** |
| `check-title` | `[Math]::Min($mark, $len)` — read from the END of the fresh file | returns nothing, always. Every "says nothing" check passes; every "says X" check fails for a reason that is not the product |
| `check-dot` | reported the shrunk file and stopped | correct, and its note said where the rest of the fix belonged |
| `check-menu` | **no mark at all** — last match in the last 600 lines | could answer with a line from a **previous run**, and whether it did depended on how much the add-in happened to have logged since |

The seventh is the one that matters most: `check-menu`'s reading sits under **the most safety-critical
assertion in this project** — "the batch stopped *because the user declined*, not because it ran out
of patience". That is the assertion that exists because two earlier versions of the add-in's
save-prompt detection were wrong and *every outside-visible check passed anyway*. Its evidence was a
coin toss dressed as a measurement.

The dangerous shape in all of them is the same, and it is this project's recurring one: **a read that
cannot say it read nothing.**

## The fix

A mark is no longer a number. It is an offset **plus the 256 bytes that were at it** — and a read
verifies those bytes are still there before it believes the offset. A rolled log is a hard error
carrying a full explanation, not an empty result.

Three detectors, because a roll is reachable three ways:

- **the file is shorter than the mark** — the roll happened and the file has not caught up
- **the anchor bytes have changed** — it rolled *and grew back past the mark*. **No offset can ever
  see this**: the file is longer than the mark, so every length test passes. This is the hole
  `check-dot`'s note called out and could not close on its own
- **the file is shorter than the high-water mark this mark has already seen** — covers a mark taken
  at 0 bytes, which has no anchor. Every read updates it, and the suites read inside `Wait-Until`
  loops, so there is almost always an earlier read to have seen the file bigger

**Marking the log by appending a line to it was the other candidate and was rejected on a
measurement.** The add-in appends with `FILE_APPEND_DATA` (`log.cpp:102`), which is atomic against
other appenders; .NET's `FileMode.Append` is a *seek-to-end and then write*, which is not, and can
land on top of a line Word's thread wrote in between. **A test instrument may not corrupt the
evidence it is reading.** Reading the bytes that are already there costs nothing and cannot.

`Reset-LogFile` is the other half, and the cheaper one: `check-all.ps1` archives the log after each
suite, so the next one starts at 0 bytes and **cannot reach 512KB on its own** — measured across a
whole battery, the suites write **9.0KB (`startscreen`) to 47.8KB (`scroll`)** each. The roll stops being reachable; the detectors stay
as the backstop. It also means the battery now leaves **one log file per suite** in
`%LOCALAPPDATA%\WordTab\history\` instead of whatever the last 512KB happened to be, which is exactly
the evidence *"the suite that reports a failure is not necessarily the suite that caused it"* needs.
`uninstall.ps1` still takes it all with one `Remove-Item -Recurse`, so the no-admin deployment story
is unchanged.

## The API, and why it has this shape

    Set-LogMark [[-Mark] <mark>]   take a mark, or put a stashed one back. Returns NOTHING
    Get-CurrentLogMark             the mark in force, to stash
    Get-LogSince <pattern>         every line since the mark that matches
    Get-LogCount <pattern>         how many
    Get-LogLast  <pattern>         the last one, or $null
    Reset-LogFile [-Label <s>]     move the log to history\; best-effort, never throws

- **`Set-LogMark` returns nothing on purpose.** The suites call it as a bare statement, and a
  function that returned the mark would dump a `pscustomobject` into the middle of their output —
  which `check-all.ps1` then scrapes for the summary line.
- **There is deliberately no `Get-LogMark`.** Four suites called `$mark = Get-LogMark` and got a
  number. A function of that name returning anything else would hand every un-migrated call site a
  plausible wrong answer; removing the name makes a missed call site an error.
- **`Get-LogSince` type-checks its argument.** Four suites called it as `Get-LogSince $mark 'pattern'`
  — mark first — and with a `[string]$Pattern` PowerShell would cheerfully stringify the mark and
  match nothing. That is the exact silent-empty-result this whole primitive exists to remove, so it
  throws instead.

## Proved without Word: `pwsh -File tools\WordTabHarness.ps1 -SelfTest`

**23 checks, about a second, no Word and no desktop** — and `check-all.ps1` runs it as a gate before
the first suite.

It exists because **the fault cannot be reached inside a real run on purpose**: it needs a 512KB log,
which is two whole batteries, and what it produces is silence. A scratch file rolls in a millisecond.
Every branch is reachable there, including the one an offset can never see — and that case is
asserted in two parts, first that the regrown file *is longer than the mark* (so no length test could
have caught it), then that the anchor catches it anyway.

## Three mistakes worth more than the successes

**1. A `param()` block on a dot-sourced script leaves TYPE-CONSTRAINED variables in the caller's
scope — and that broke the very gate this slice added.** The harness declares
`param([switch]$SelfTest)`, and dot-sourcing it puts `$SelfTest` into all eleven suites' scopes *with
its `[switch]` type still attached*. `check-all.ps1` then did `$selfTest = & pwsh ... -SelfTest` and
died with **"Cannot convert System.Object[] to SwitchParameter"** — PowerShell variable names are
case-insensitive, so it was assigning to the parameter. **Found by a run, not by reading it**, having
noticed the pollution while writing it and waved it away as harmless. The harness now copies the flag
to `$script:HarnessSelfTest` and `Remove-Variable`s the name: a shared file may not leave names in
the scope of everything that loads it, least of all names with types attached.

**2. The gate carried on after failing to read its own result.** It printed `harness self-test: (no
summary line)` and ran the battery anyway. That is the same shape as `check-all -Only` matching no
suites and reporting *"All 0 suites green"*, which this project already fixed once. It now exits 1
when the self-test says nothing it can read, not only when the self-test fails.

**3. The archive was labelled with the wrong suite, and only looking at the output showed it.**
`Reset-LogFile -Label $name` ran *before* suite `$name`, so every file was named after the suite that
was **about to start** rather than the one that wrote it — a file called `title.log` holding
`scroll`'s output. In a slice whose entire subject is evidence that quietly describes the wrong
thing, that is worse than keeping no file at all. It now runs after each suite, with one
`before-battery` archive ahead of the first. **The proof is the file sizes, not the names**: in the
battery that had the bug, every archive's size matches the suite one position *earlier* in the run
order (`title.log` 48.1KB against `scroll`'s 47.8KB); in the battery after it they line up with their
own names.

Also corrected: a comment claiming the add-in's log lines "run 90-200 bytes". **Measured over a whole
battery's log — 2030 lines, 62 to 315 bytes, mean 121** — so the longest line is longer than the
anchor. That is fine, and the comment now says why: the anchor identifies a byte range, it does not
have to be a whole line.

## Measured while doing it

- The add-in opens the log per line and closes it again, so nothing holds it between writes and
  `Reset-LogFile` can move it while Word runs. Word being closed between suites makes it certain.
- A harness read holds the file with `FileShare.ReadWrite` — **not `Delete`** — so while a read is in
  flight the add-in's own `DeleteFileW` fails and `log.cpp` ignores the return. The harness can
  therefore *suppress* a roll for a moment. Benign, true of every previous copy, and unreachable now
  that the log is emptied per suite; recorded so nobody rediscovers it as a mystery.

## What this does not do

- `Get-LogSince` matches with `-like "*$pattern*"`, which is a **wildcard** match, not a substring
  one. Six of the seven copies did the same and every pattern in the suites is plain text, so nothing
  changes — but `check-menu`'s copy used `Select-String -SimpleMatch`, a genuine substring, and the
  two only agree while the patterns stay free of `[`, `]`, `*` and `?`.
- A mark taken at 0 bytes on a file that then rolls **before any read** is still undetectable. It
  needs 512KB inside one suite, and the largest writes 40KB.
