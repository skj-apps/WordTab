// The strip palette's adoption gate, driven by the work rig's own sampled colours.
//
// Compiled and run by tools\check-palette.ps1; see that file for why this is a suite of its own.
// The short version is the same as check-governor's: **the dev rig cannot produce the input.** The
// fault this gate exists for is Word's ribbon reading RGB(9,9,9) while a window is active and
// RGB(10,10,10) while it is not, and a sampler that once in a working day reads RGB(0,0,0) instead
// of either. Nothing here can make Word's ribbon do that on demand, so check-look.ps1 - which drives
// real Word - sees a ribbon that holds one colour all run and cannot tell the old rule from the new
// one. That is exactly how the old rule shipped: it passed every suite on this machine and then
// rebuilt the palette over a hundred times in one day on theirs.
//
// So the decision is lifted out and replayed against colours copied from the work rig's log. Nothing
// is #included from the project, deliberately: if this copy ever drifts from strip.cpp then the copy
// is wrong, and the copy is what is being asserted. Keep ChromeNear and AdoptStep below in step with
// ChromeNear and AdoptSampledChrome by hand.

#include <stdio.h>

typedef unsigned long COLORREF;
#define RGBV(r, g, b) ((COLORREF)((r) | ((g) << 8) | ((b) << 16)))
#define RV(c) ((int)((c) & 0xFF))
#define GV(c) ((int)(((c) >> 8) & 0xFF))
#define BV(c) ((int)(((c) >> 16) & 0xFF))

static const int kSameWithin = 4;                 // CHROME_SAME_WITHIN

static bool ChromeNear(COLORREF a, COLORREF b)
{
    int dr = RV(a) - RV(b); if (dr < 0) dr = -dr;
    int dg = GV(a) - GV(b); if (dg < 0) dg = -dg;
    int db = BV(a) - BV(b); if (db < 0) db = -db;
    return dr <= kSameWithin && dg <= kSameWithin && db <= kSameWithin;
}

// The gate. `applied` stands in for g_palette.chrome, and the return value is whether this sample
// caused a rebuild - which on the real thing means two brushes destroyed and remade and every strip
// in the row invalidated.
struct Gate
{
    COLORREF applied;
    bool     ready;
    bool     everSampled;
    bool     pendingSet;
    COLORREF pending;

    // Constructed the way StripStart leaves it: the theme registry's guess is already applied, and
    // nothing has been sampled yet.
    Gate(COLORREF fallback)
        : applied(fallback), ready(true), everSampled(false), pendingSet(false), pending(0) {}
};

static bool AdoptStep(Gate* g, COLORREF chrome)
{
    if (g->ready && g->everSampled && ChromeNear(g->applied, chrome))
    {
        g->pendingSet = false;
        return false;
    }

    if (g->everSampled && (!g->pendingSet || !ChromeNear(g->pending, chrome)))
    {
        g->pending    = chrome;
        g->pendingSet = true;
        return false;
    }

    g->everSampled = true;
    g->pendingSet  = false;

    if (g->ready && g->applied == chrome)
        return false;                             // ApplyPalette's own compare-before-act

    g->applied = chrome;
    g->ready   = true;
    return true;
}

// The old rule, kept only so the two can be compared: ApplyPalette called directly, believing any
// value that is not bit-identical to the one already applied.
struct Old { COLORREF applied; Old(COLORREF fallback) : applied(fallback) {} };
static bool OldStep(Old* g, COLORREF chrome)
{
    if (g->applied == chrome)
        return false;
    g->applied = chrome;
    return true;
}

static int failures = 0;
static void Check(bool ok, const char* what)
{
    printf("    %s  %s\n", ok ? "PASS" : "FAIL", what);
    if (!ok) failures++;
}

// Word's four measured ribbon bodies. Black is the theme the work rig runs; dark grey is what this
// project measured first and what the registry fallback still names; the two light ones are the far
// end of the range the dead band must not swallow.
static const COLORREF kBlack     = RGBV(9, 9, 9);
static const COLORREF kBlackIdle = RGBV(10, 10, 10);
static const COLORREF kDarkGrey  = RGBV(41, 41, 41);
static const COLORREF kLightGrey = RGBV(243, 242, 241);
static const COLORREF kWhite     = RGBV(255, 255, 255);

int main(void)
{
    // ---- 1. the first measurement still wins immediately -----------------------------------------
    printf("==> The registry's guess is replaced by the first real sample, with no wait\n");
    {
        Gate g(kDarkGrey);                        // DarkThemeInUse() -> RGB(41,41,41)
        Check(AdoptStep(&g, kBlack), "the first sample is adopted at once, however far it has moved");
        Check(g.applied == kBlack,   "and it is the sampled colour that is now applied");
    }

    // ---- 2. the flap -----------------------------------------------------------------------------
    // The work rig's log for 2026-08-28, in order: the ribbon reads 9 while its window is active and
    // 10 while it is not, so every switch between two documents moved it by one unit in each channel.
    printf("==> Word's active/inactive ribbon must not read as a theme change\n");
    {
        Gate g(kDarkGrey); Old o(kDarkGrey);
        AdoptStep(&g, kBlack); OldStep(&o, kBlack);

        int rebuilds = 0, oldRebuilds = 0;
        for (int i = 0; i < 60; i++)
        {
            COLORREF c = (i & 1) ? kBlackIdle : kBlack;
            if (AdoptStep(&g, c)) rebuilds++;
            if (OldStep(&o, c))   oldRebuilds++;
        }
        printf("    60 alternating samples: new rule %d rebuilds, old rule %d\n", rebuilds, oldRebuilds);
        Check(rebuilds == 0,     "one unit in each channel is the same colour - nothing is rebuilt");
        // 59 and not 60: the first of the sixty repeats the colour already applied, and equality
        // caught that one. Every other sample - every switch of the active window, all working day -
        // went through.
        Check(oldRebuilds == 59, "the old rule rebuilt on all but the first, which is the field report");
        Check(g.applied == kBlack, "and the palette still holds the colour it first measured");
    }

    // ---- 3. the outlier --------------------------------------------------------------------------
    // 14:05:24.848 sampled RGB(0,0,0) with RGB(10,10,10) either side of it, and again at 14:09:19.310.
    // Ten units is well outside the dead band, so only the second half of the gate can catch it.
    printf("==> A single sample that is not the ribbon must not reach the screen\n");
    {
        Gate g(kDarkGrey);
        AdoptStep(&g, kBlackIdle);

        Check(!AdoptStep(&g, RGBV(0, 0, 0)), "the outlier is not adopted on first sight");
        Check(g.applied == kBlackIdle,       "so the row keeps the colour it was drawing");
        Check(!AdoptStep(&g, kBlackIdle),    "the next sample agrees with what is applied, and nothing moves");

        // ...and the pending candidate must not survive to be confirmed by a later, unrelated one.
        Check(!AdoptStep(&g, RGBV(0, 0, 0)), "a second, separate outlier is also refused first time");
        Check(g.applied == kBlackIdle,       "still unchanged");
    }

    // ---- 4. a real theme change still gets through ------------------------------------------------
    printf("==> A person in File > Account is believed, one sample later\n");
    {
        Gate g(kDarkGrey);
        AdoptStep(&g, kBlack);

        Check(!AdoptStep(&g, kWhite), "the first sight of white is held back - it could be an outlier");
        Check(AdoptStep(&g, kWhite),  "the second sight two seconds later is believed");
        Check(g.applied == kWhite,    "and the palette is white");

        Check(!AdoptStep(&g, kDarkGrey), "changing again: held once more");
        Check(AdoptStep(&g, kDarkGrey),  "and adopted on the second reading");
    }

    // ---- 5. the dead band is nowhere near the themes it must tell apart ---------------------------
    printf("==> The four themes Word actually has are all further apart than the dead band\n");
    {
        Check(!ChromeNear(kBlack, kDarkGrey),     "black and dark grey are different colours (9 vs 41)");
        Check(!ChromeNear(kLightGrey, kWhite),    "light grey and white are different colours (243 vs 255)");
        Check(!ChromeNear(kDarkGrey, kLightGrey), "dark grey and light grey are different colours");
        Check(ChromeNear(kBlack, kBlackIdle),     "active and inactive black are the same colour");

        // The smallest real gap is light grey to white at 12; the dead band is 4. If a future Word
        // ships two themes closer together than this, THIS is the assertion that says so.
        Check(kSameWithin < 12, "the dead band is below the smallest gap between two real themes");
    }

    // ---- 6. the flap around a theme change -------------------------------------------------------
    // The nastiest input the real thing can see: a genuine change, arriving in the middle of the
    // one-unit jitter. The pending candidate has to survive the jitter without being cleared by it.
    printf("==> A theme change during the usual jitter is still adopted\n");
    {
        Gate g(kDarkGrey);
        AdoptStep(&g, kBlack);
        AdoptStep(&g, kBlackIdle);                // jitter, absorbed

        Check(!AdoptStep(&g, kWhite), "white, first sight: held");
        // Word's ribbon jitters in the new theme too, and the confirming sample is allowed to be one
        // unit off the candidate - ChromeNear, not equality, is what the second reading is judged by.
        Check(AdoptStep(&g, RGBV(253, 253, 253)), "a confirming sample one unit off still confirms it");
        Check(g.applied == RGBV(253, 253, 253),   "and what is applied is the sample that confirmed, not the candidate");
    }

    printf("\n%s\n", failures ? "FAILED" : "All good.");
    return failures ? 1 : 0;
}
