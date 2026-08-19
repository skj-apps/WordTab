// The dot poll's cadence governor, old against new, driven by the work rig's own measurements.
//
// Compiled and run by tools\check-governor.ps1; see that file for why this is a suite of its own.
// The short version: **the dev rig cannot produce a slow pass.** Local files answer Document.Saved in
// a few hundred microseconds, so the whole path this governor exists for never executes here, and
// check-dot.ps1 - which drives real Word - can only ever see the cheap case. That is exactly how the
// broken first version shipped: it passed every suite on this machine.
//
// So the arithmetic is lifted out and replayed against numbers taken from the work rig's log. Nothing
// is #included from the project, deliberately: if this copy ever drifts from strip.cpp then the copy
// is wrong, and the copy is what is being asserted. Keep the three step functions below in step with
// PollModified by hand.

#include <stdio.h>
typedef long long LONGLONG;

static const int kRungTicks[] = { 1, 4, 12, 40, 120 };
static const int kRungCount   = 5;

static int RungFor(LONGLONG us)
{
    if (us > 200000) return 4;
    if (us > 50000)  return 3;
    if (us > 15000)  return 2;
    if (us > 4000)   return 1;
    return 0;
}

static int TicksFor(LONGLONG us)          // f333349, kept only so the two can be compared
{
    if (us > 200000) return 120;
    if (us > 50000)  return 40;
    if (us > 15000)  return 12;
    if (us > 4000)   return 4;
    return 1;
}

// f333349: the cadence from the CHEAPER of the last two passes, and one cheap pass restores it.
struct Old { LONGLONG prev; Old() : prev(0) {} };
static int OldStep(Old* g, LONGLONG us)
{
    LONGLONG lastTime = g->prev;
    g->prev = us;
    LONGLONG basis = (us < lastTime) ? us : lastTime;
    return TicksFor(basis) * 500;                         // ms
}

// Now, for the whole-row pass: the DEARER of the last two, climb in one step, descend a rung at a
// time, first pass of a process discarded and not stored.
struct New { int rung; LONGLONG prev; bool seen; New() : rung(0), prev(0), seen(false) {} };
static int NewStep(New* g, LONGLONG us)
{
    LONGLONG lastTime = g->prev;
    if (!g->seen)
    {
        g->seen = true;
    }
    else
    {
        g->prev = us;
        LONGLONG basis = (us > lastTime) ? us : lastTime;
        int want = RungFor(basis);
        if (want > g->rung)      g->rung = want;
        else if (want < g->rung) g->rung--;
    }
    if (g->rung < 0)           g->rung = 0;
    if (g->rung >= kRungCount) g->rung = kRungCount - 1;
    return kRungTicks[g->rung] * 500;
}

// ...and for the active-window pass, which carries nothing between passes because the subject
// changes underneath it: switching tabs changes which document is being asked about.
struct Act { int rung; bool seen; Act() : rung(0), seen(false) {} };
static int ActStep(Act* g, LONGLONG us)
{
    if (!g->seen) g->seen = true;
    else          g->rung = RungFor(us);
    return kRungTicks[g->rung] * 500;
}

static int failures = 0;
static void Check(bool ok, const char* what)
{
    printf("    %s  %s\n", ok ? "PASS" : "FAIL", what);
    if (!ok) failures++;
}

int main(void)
{
    // ---- 1. a healthy rig is not disturbed -------------------------------------------------------
    // This rig's own numbers: cold starts of 26,621us, 27,898us and 29,977us on three consecutive
    // runs, against a steady state three orders of magnitude below them.
    printf("==> A cold start must not put a healthy machine on a slow cadence\n");
    {
        New n; Old o;
        int nFirst = NewStep(&n, 27547), oFirst = OldStep(&o, 27547);
        Check(nFirst == 500, "new: the first pass leaves the cadence at 500ms");
        Check(oFirst == 500, "old: same - min(anything, 0) was 0, so it was free there too");

        int nLast = 0, oLast = 0;
        for (int i = 0; i < 20; i++) { nLast = NewStep(&n, 400); oLast = OldStep(&o, 400); }
        Check(nLast == 500, "new: stays at 500ms across 20 cheap passes");
        Check(oLast == 500, "old: likewise - this rig sees no change at all");
    }

    // ---- 2. the thrash ---------------------------------------------------------------------------
    // The work rig alternates: cheap while SharePoint's answer is warm, about half a second once it
    // has gone cold. These are the values its log carries around 10:20 on 2026-08-19.
    printf("==> Alternating slow and fast - the pattern that produced the thrash\n");
    {
        New n; Old o;
        const LONGLONG seq[] = { 494927, 268, 486406, 95044, 344, 501877, 92637, 760,
                                 499611, 95902, 502, 502810, 109829, 503, 495682, 352 };
        const int      count = (int)(sizeof(seq) / sizeof(seq[0]));

        NewStep(&n, seq[0]); OldStep(&o, seq[0]);         // pass 1: discarded by the new governor

        // Counted from pass 3 on. Pass 2 is expected to read 500ms and is not interesting: pass 1
        // was discarded WITHOUT being stored, so pass 2 has nothing to be the dearer of and is
        // judged on itself. That costs one extra half-second tick per process, at startup, once.
        int oBackToFloor = 0, nBackToFloor = 0;
        for (int i = 1; i < count; i++)
        {
            int oms = OldStep(&o, seq[i]);
            int nms = NewStep(&n, seq[i]);
            if (oms == 500)           oBackToFloor++;
            if (nms == 500 && i >= 2) nBackToFloor++;
        }
        printf("    old returned to 500ms %d times in %d passes; new %d after pass 2\n",
               oBackToFloor, count - 1, nBackToFloor);
        Check(oBackToFloor >= 5, "old: snaps back to 500ms repeatedly - this is the bug");
        Check(nBackToFloor == 0, "new: never returns to 500ms while slow passes keep arriving");
    }

    // ---- 3. and it does come back, on sustained evidence -----------------------------------------
    printf("==> Recovery takes sustained cheap passes, not one lucky sample\n");
    {
        New n;
        NewStep(&n, 400);                                 // discarded
        int after = NewStep(&n, 518726);                  // the worst pass the work rig ever logged
        Check(after == 60000, "one slow pass backs off to 60s at once");

        int steps = 0, ms = after;
        while (ms != 500 && steps < 20) { ms = NewStep(&n, 400); steps++; }
        printf("    walked 60000ms -> 500ms in %d cheap passes\n", steps);
        Check(ms == 500, "it does return to 500ms once Word is genuinely quick again");
        Check(steps == 5, "and takes five - one to flush the slow sample, then a rung each");

        // A lucky warm answer in the middle of trouble must not undo the backoff.
        New m;
        NewStep(&m, 400);
        NewStep(&m, 518726);
        int mid = NewStep(&m, 300);
        Check(mid == 60000, "one cheap pass does not move it at all - max() still sees the slow one");
        int mid2 = NewStep(&m, 300);
        Check(mid2 == 20000, "two consecutive cheap passes step down exactly one rung");
    }

    // ---- 4. the active window is not the row -----------------------------------------------------
    printf("==> The active pass is judged on itself, so one slow tab cannot poison the next\n");
    {
        Act a;
        ActStep(&a, 300);                                 // discarded
        Check(ActStep(&a, 300)    == 500,   "a warm foreground document is asked twice a second");
        Check(ActStep(&a, 518726) == 60000, "a slow one backs the ACTIVE pass off too - cost still wins");
        Check(ActStep(&a, 300)    == 500,   "and the next cheap pass restores it, with no walk down");

        // The failure this rule exists to prevent: moving to a tab whose document is warm must not
        // inherit a cadence that a different, slow document set.
        Act b; New n;
        ActStep(&b, 300);    NewStep(&n, 300);
        ActStep(&b, 518726); NewStep(&n, 518726);         // document A, cold
        int aMs = ActStep(&b, 300);                        // document B, warm
        int nMs = NewStep(&n, 300);
        printf("    slow tab then a warm one: active %dms; the row's sticky rule would give %dms\n",
               aMs, nMs);
        Check(aMs == 500,   "the warm tab is asked twice a second");
        Check(nMs == 60000, "the sticky rule would have pinned it at 60s - why the row's rule is not used here");
    }

    printf("\n%s\n", failures ? "FAILED" : "All good.");
    return failures ? 1 : 0;
}
