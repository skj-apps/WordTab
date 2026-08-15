// WordTab - the strip: carving 32 logical pixels out of Word's own layout, from inside Word.
//
// Spike 1 proved the geometry works (spikes\StripSpike\RESULT.md): shrink Word's `_WwF` document
// frame from the top, park our own window in the gap, and the result survives resize, maximize and
// Backstage. It did that from another process on a 30ms poll, which meant every relayout Word made
// was visible for a frame or two before we corrected it.
//
// This is the same geometry with the poll taken out. The move: rather than watch `_WwF` and correct
// it afterwards, we subclass `_WwF` itself and rewrite the *proposed* rectangle in
// WM_WINDOWPOSCHANGING. Word never gets to move it to the wrong place at all - there is no
// correction, because there is nothing to correct. Any SetWindowPos or MoveWindow on that window,
// from anywhere, comes through here first, so there is no path by which Word can lay out behind our
// back.
//
// The two rules that carried over from the spikes, both learned by watching them fail:
//
//   1. The sync must be idempotent against Word's layout. Word rewrites `_WwF` to its own natural
//      rect repeatedly - six times in one logged spike session - and a design that nudges by -32
//      each time walks the document frame off the bottom of the window in seconds. That is the old
//      "Doc Tabs" failure mode. So: what Word proposes is treated as natural, and the shift is
//      computed from it, never accumulated onto it.
//
//   2. Idempotence against Word is not idempotence against another add-in. Office Tab shifts this
//      same `_WwF`. Two processes each reading the other's +32 as a fresh natural rect squeezed the
//      document frame to 25px tall in seconds - measured, spike 2. Hence the top-edge test below
//      (if the top is exactly where we left it, nobody re-laid it out, so keep it) and a height
//      tripwire that refuses to shrink the document below twice the strip.
//
// The same house rules as frames.cpp apply: nothing throws, nothing escapes, and nothing here is
// slow. This code runs between Word and its own window procedure during a live resize.

#include "wordtab.h"
#include <commctrl.h>
#include <string.h>
#include <wchar.h>
#include <stdio.h>

// Height of the strip in logical pixels, scaled per-window by DPI. 32 is what spike 1 used and what
// Office Tab's own strip measures at 100%.
#define STRIP_LOGICAL_H   32
#define TAB_LOGICAL_W    220
#define MAX_STRIPS       256

// The affordances on a tab, all in logical pixels and all scaled per window.
#define TAB_LOGICAL_MIN_W  70    // below this a tab is a colour, not a label
#define TAB_LOGICAL_PAD     6    // the strip's left and right inset
#define TAB_LOGICAL_GAP     4    // between the last tab and the new-document button
#define CLOSE_LOGICAL      16    // the close button's hit target, a square
#define PLUS_LOGICAL       26    // the new-document button's width

// How far a press has to travel before it is a drag rather than a click. Four logical pixels is what
// Windows itself uses (SM_CXDRAG's default), but taken as our own scaled constant rather than read
// from the system: SM_CXDRAG is not per-monitor DPI scaled, so on this 150% rig it would be a third
// smaller than everything else in this file, and the check script mirrors these numbers.
#define DRAG_LOGICAL_SLOP   4

// The tab row is bounded independently of the stack. 128 tabs at the 70px minimum is wider than any
// monitor sold, so a layout array larger than this could only describe tabs nobody can see - and an
// unbounded one on the stack of a WM_MOUSEMOVE handler is a different kind of problem.
#define MAX_TABS         128

// What the pointer is over, or what a button press is claiming. Used for both, which is why HIT_TAB
// appears as a press kind: it means a *middle* press, since a left press on a tab acts immediately
// and never waits for a release.
enum { HIT_NONE = 0, HIT_TAB, HIT_CLOSE, HIT_PLUS };

// Posted to a strip so that a command runs after the handler that raised it has returned. Calling
// into Word's object model from inside our own window procedure would re-enter it: Documents.Add
// creates a window and pumps messages, Document.Save can put up a dialog that runs its own message
// loop, and some of those messages are ours.
//
//   WM_WORDTAB_CMD   wParam = CMD_*, lParam = the tab it is about (NULL for the ones that are not)
//   WM_WORDTAB_SAVE  lParam = the tab, once it has been activated and Word has processed that
#define WM_WORDTAB_CMD   (WM_APP + 10)
#define WM_WORDTAB_SAVE  (WM_APP + 11)

// The menu's command ids, which are also the ids TrackPopupMenu hands back. They start at 1 because
// TPM_RETURNCMD answers 0 for "the user dismissed it without choosing".
enum { CMD_NEW = 1, CMD_SAVE, CMD_CLOSE, CMD_CLOSE_OTHERS, CMD_CLOSE_ALL };

static const wchar_t* const kWwfClass    = L"_WwF";
static const wchar_t* const kStripClass  = L"WordTabStrip";

// Our link in `_WwF`'s subclass chain. Distinct from the frame's - see frames.cpp.
static const UINT_PTR kWwfSubclassId = 0x57544143;   // 'WTAC'

// ---------------------------------------------------------------------------------------------
// Per-frame state.
//
// One entry per OpusApp frame we have subclassed. Kept here rather than in frames.cpp's table so
// that file stays about the frame subclass and this one about geometry; they are joined only by
// the four Strip* calls declared in wordtab.h.
//
// All of it is touched on Word's UI thread only.
// ---------------------------------------------------------------------------------------------

struct StripState
{
    HWND  frame;         // Word's OpusApp
    HWND  wwf;           // Word's document frame, subclassed by us
    HWND  strip;         // ours, a child of the frame

    int   stripH;        // STRIP_LOGICAL_H scaled to this window's DPI
    int   dpi;
    HFONT font;

    BOOL  enabled;       // cleared while restoring, so the handler stops rewriting
    BOOL  hasApplied;
    RECT  natural;       // where Word wants `_WwF`, in frame client coordinates
    RECT  applied;       // where we put it: natural, top pushed down by stripH

    BOOL  stripPlaced;
    RECT  stripAt;
    SIZE  clientAtNatural;   // the frame's client size when `natural` was recorded - see StripRefit

    BOOL  trippedLogged; // the height tripwire says its piece once per frame, not once per message
    BOOL  wasVisible;    // to catch hidden -> shown, where the layout has to be re-derived

    // Logging a relayout costs a file write, and Word relayouts on every mouse movement of a resize
    // drag. So they are throttled and counted, and the count is reported with the next line that
    // does get written - the alternative is either a stutter the user can feel or a silence that
    // hides how often this happens.
    DWORD lastLogTick;
    int   suppressed;

    // What the pointer is over and what is being pressed, held per strip rather than in one global.
    // Only the active window's strip is on top, so only it receives mouse messages - but a strip
    // that was hovered and then covered would otherwise keep drawing a highlight under a pointer
    // that is somewhere else entirely, and show it again the moment its tab came forward.
    //
    // Frames, not indices. The tab row can change between a press and its release - that is exactly
    // what closing a tab does - and an index that meant one document on the way down can mean a
    // different one on the way up.
    int   hotKind;
    HWND  hotFrame;
    int   pressKind;      // a left press on a close or new button, waiting for its release
    HWND  pressFrame;
    HWND  middleFrame;    // a middle press on a tab, likewise
    BOOL  rightDown;      // a right press, waiting to become a context menu on release
    HWND  rightFrame;     // ...and the tab it landed on, NULL for the empty part of the strip
    HWND  menuFrame;      // the tab a context menu is open for, kept lit while it is up
    BOOL  tracking;       // TrackMouseEvent armed, so WM_MOUSELEAVE will arrive

    wchar_t title[256];
};

// Where every clickable thing in a strip is, in the strip's client coordinates. One structure
// computed by one function and used by both painting and hit-testing, so a click can never land
// somewhere other than what was drawn.
struct StripLayout
{
    int  count;
    RECT tab[MAX_TABS];
    RECT close[MAX_TABS];   // empty when the tab is too narrow to carry a button honestly
    RECT plus;
    BOOL hasPlus;
};

// What a point in a strip is over.
struct StripHit
{
    int  kind;
    HWND frame;             // for HIT_TAB and HIT_CLOSE
    int  index;             // ...and where in the row it is, -1 otherwise
    RECT tab;               // ...and the tab's rectangle, so picking one up needs no second layout
};

static StripState g_strips[MAX_STRIPS];
static int  g_stripCount = 0;
static BOOL g_stripEnabled = TRUE;
static BOOL g_buttonsEnabled = TRUE;     // HKCU\Software\WordTab\TabButtons
static BOOL g_menuEnabled = TRUE;        // HKCU\Software\WordTab\TabMenu
static BOOL g_dragEnabled = TRUE;        // HKCU\Software\WordTab\TabDrag

// ---------------------------------------------------------------------------------------------
// A tab being dragged.
//
// Global, unlike hover and the pressed buttons, which are per strip. The reason is that a press on a
// tab activates that document *before* the drag begins - and activating raises a different window,
// whose strip is now the one in front. So the strip holding the mouse capture is very often not the
// strip the user can see. Per-strip state would draw the tab travelling on a window that is
// underneath another one, and nothing at all on the window they are looking at.
//
// Held as the add-in's state instead, every strip can draw the same carried tab, and the one on top
// is by construction the one that shows it. They are all the same width at the same position - that
// is what the stack guarantees - so one x coordinate is meaningful in all of them.
//
// g_dragStrip doubles as "this strip is holding the mouse capture for a tab gesture". g_dragFrame is
// cleared on cancel while the capture is kept, because the left button is still down and letting go
// of the mouse mid-gesture would deliver the release to whatever is underneath.
// ---------------------------------------------------------------------------------------------

static HWND g_dragStrip  = NULL;    // the strip that owns the capture, NULL when no press is held
static HWND g_dragFrame  = NULL;    // the tab under that press, NULL once cancelled
static int  g_dragPressX = 0;       // where the press landed, in strip client coordinates
static int  g_dragGrabDx = 0;       // how far into the tab, so it does not jump when picked up
static int  g_dragLeft   = 0;       // the carried tab's left edge right now
static int  g_dragFrom   = 0;       // the position it was picked up from - the log, and the undo
static BOOL g_dragging   = FALSE;   // past the slop: this is a drag, not a click that has not ended
static ATOM g_stripClass = 0;
static UINT_PTR g_janitor = 0;

static HBRUSH   g_backBrush    = NULL;   // strip background
static HBRUSH   g_tabBrush     = NULL;   // the selected tab
static HBRUSH   g_tabIdleBrush = NULL;   // the others
static HBRUSH   g_tabHotBrush  = NULL;   // an unselected tab under the pointer
static HBRUSH   g_chipHotBrush = NULL;   // a close or new button under the pointer
static HBRUSH   g_chipDownBrush = NULL;  // ...and while it is held down
static HPEN     g_edgePen      = NULL;
static COLORREF g_edgeColor    = RGB(200, 198, 196);
static COLORREF g_textColor    = RGB(50, 49, 48);
static COLORREF g_idleTextColor = RGB(96, 94, 92);
static COLORREF g_glyphColor    = RGB(96, 94, 92);
static COLORREF g_glyphHotColor = RGB(32, 31, 30);
static BOOL     g_darkTheme    = FALSE;

static StripState* FindByFrame(HWND frame)
{
    for (int i = 0; i < g_stripCount; i++)
        if (g_strips[i].frame == frame)
            return &g_strips[i];
    return NULL;
}

// ---------------------------------------------------------------------------------------------
// DPI. This rig runs at 150%, so a hard-coded 32 is 32 physical pixels here - two thirds of the
// height it should be. GetDpiForWindow is Windows 10 1607 and is reached through GetProcAddress
// rather than a header, because w64devkit's headers are older than the API and the failure would
// be a link error at the end of a long build.
// ---------------------------------------------------------------------------------------------

typedef UINT (WINAPI *GetDpiForWindowFn)(HWND);

static int DpiOf(HWND hwnd)
{
    static GetDpiForWindowFn fn = NULL;
    static BOOL looked = FALSE;
    if (!looked)
    {
        looked = TRUE;
        HMODULE user32 = GetModuleHandleW(L"user32.dll");
        if (user32)
            fn = (GetDpiForWindowFn)(void*)GetProcAddress(user32, "GetDpiForWindow");
    }

    if (fn)
    {
        UINT dpi = fn(hwnd);
        if (dpi >= 72 && dpi <= 480)
            return (int)dpi;
    }

    // Pre-1607, or a window the system will not answer for: the system DPI is close enough, and on
    // a single-monitor rig it is identical.
    HDC screen = GetDC(NULL);
    int dpi = screen ? GetDeviceCaps(screen, LOGPIXELSY) : 96;
    if (screen)
        ReleaseDC(NULL, screen);
    return dpi > 0 ? dpi : 96;
}

static int Scaled(int logical, int dpi)
{
    return MulDiv(logical, dpi, 96);
}

// ---------------------------------------------------------------------------------------------
// Theme.
//
// Not a design decision - the real look of the tabs is a product question for later. This is the
// minimum needed for the strip not to read as broken: a band of light grey across a black Word
// window looks like a failure even when the geometry underneath it is perfect.
//
// Office keeps its own theme setting; "use system setting" defers to Windows.
// ---------------------------------------------------------------------------------------------

static DWORD ReadDword(HKEY root, const wchar_t* key, const wchar_t* name, DWORD fallback)
{
    DWORD value = 0;
    DWORD size = sizeof(value);
    if (RegGetValueW(root, key, name, RRF_RT_REG_DWORD, NULL, &value, &size) != ERROR_SUCCESS)
        return fallback;
    return value;
}

static BOOL DarkThemeInUse(void)
{
    // 0 colorful, 3 dark grey, 4 black, 5 white, 6 follow Windows. Colorful has a coloured title
    // bar but a light ribbon, so it belongs with the light palette.
    DWORD theme = ReadDword(HKEY_CURRENT_USER,
                            L"Software\\Microsoft\\Office\\16.0\\Common", L"UI Theme", 6);

    if (theme == 3 || theme == 4)
        return TRUE;
    if (theme == 0 || theme == 5)
        return FALSE;

    return ReadDword(HKEY_CURRENT_USER,
                     L"Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize",
                     L"AppsUseLightTheme", 1) == 0;
}

static void MakeFont(StripState* state)
{
    if (state->font)
    {
        DeleteObject(state->font);
        state->font = NULL;
    }

    LOGFONTW lf;
    memset(&lf, 0, sizeof(lf));
    lf.lfHeight  = -Scaled(12, state->dpi);
    lf.lfWeight  = FW_NORMAL;
    lf.lfCharSet = DEFAULT_CHARSET;
    lf.lfQuality = CLEARTYPE_QUALITY;
    wcscpy(lf.lfFaceName, L"Segoe UI");

    state->font = CreateFontIndirectW(&lf);
}

// ---------------------------------------------------------------------------------------------
// Geometry helpers.
// ---------------------------------------------------------------------------------------------

// A child's rectangle in its parent's client coordinates - the same space WINDOWPOS.x/y use, so
// everything below can be compared without converting twice.
static BOOL ChildRect(HWND parent, HWND child, RECT* out)
{
    RECT screen;
    if (!GetWindowRect(child, &screen))
        return FALSE;

    POINT topLeft;
    topLeft.x = screen.left;
    topLeft.y = screen.top;
    if (!ScreenToClient(parent, &topLeft))
        return FALSE;

    out->left   = topLeft.x;
    out->top    = topLeft.y;
    out->right  = topLeft.x + (screen.right - screen.left);
    out->bottom = topLeft.y + (screen.bottom - screen.top);
    return TRUE;
}

static BOOL SameRect(const RECT* a, const RECT* b)
{
    return a->left == b->left && a->top == b->top &&
           a->right == b->right && a->bottom == b->bottom;
}

// Remember how big the frame was when this natural rect was true. The stack needs it to refit a
// window's document frame to a new size that Word will not lay out for it - see StripRefit.
static void RememberClient(StripState* state)
{
    RECT client;
    if (!GetClientRect(state->frame, &client))
        return;
    state->clientAtNatural.cx = client.right - client.left;
    state->clientAtNatural.cy = client.bottom - client.top;
}

static void LogRelayout(StripState* state, const wchar_t* why)
{
    DWORD now = GetTickCount();
    if (state->lastLogTick != 0 && (now - state->lastLogTick) < 250)
    {
        state->suppressed++;
        return;
    }
    state->lastLogTick = now;

    wchar_t extra[64] = L"";
    if (state->suppressed > 0)
    {
        _snwprintf(extra, 64, L"  [+%d more not logged]", state->suppressed);
        extra[63] = L'\0';
        state->suppressed = 0;
    }

    LogWrite(L"strip  hwnd=0x%p  %s: _WwF natural (%ld,%ld %ldx%ld) -> (%ld,%ld %ldx%ld)  strip h=%d%s",
             (void*)state->frame, why,
             state->natural.left, state->natural.top,
             state->natural.right - state->natural.left,
             state->natural.bottom - state->natural.top,
             state->applied.left, state->applied.top,
             state->applied.right - state->applied.left,
             state->applied.bottom - state->applied.top,
             state->stripH, extra);
}

// Put our window in the gap. Cheap to call repeatedly: it does nothing unless the rectangle moved.
//
// The strip's position is derived from where the document frame **actually is**, not from the
// natural rect we computed for it. The two ought to be the same and are almost always, but "almost"
// was worth a visible seam: Word occasionally reports a child a pixel away from where we put it -
// the frame's client origin shifts by one when Windows redraws the border - and a strip positioned
// from stored numbers then sits a pixel off the document frame it is supposed to be flush against.
// Derived from the live rect, the two cannot disagree, whatever either of them is doing.
//
// `wwfRect` overrides that when the caller already knows where the document frame is about to be:
// during WM_WINDOWPOSCHANGING it has not moved yet, so reading it would place the strip a message
// behind.
static void PlaceStrip(StripState* state, const RECT* wwfRect)
{
    if (!state->strip || !state->hasApplied)
        return;

    // Not while the window is off screen: its layout is frozen and deliberately not trusted (see
    // AdjustProposed), so the document frame down there is not something to line up against.
    if (!wwfRect && (!IsWindowVisible(state->frame) || IsIconic(state->frame)))
        return;

    RECT document = state->applied;
    if (wwfRect)
    {
        document = *wwfRect;
    }
    else if (state->wwf && IsWindow(state->wwf))
    {
        RECT live;
        if (ChildRect(state->frame, state->wwf, &live))
        {
            // The live rect is trusted for pixel-level disagreement, not for a different layout.
            // A document frame a pixel from where we put it is the border artifact this whole
            // derivation exists to absorb; one 48 pixels away is Word mid-transition - it has reset
            // the frame and we have not shifted it back yet - and following it there would put the
            // strip over the ribbon or off the top of the window. Measured, both.
            LONG drift = live.top - state->applied.top;
            if (drift < 0) drift = -drift;
            if (drift < state->stripH)
                document = live;
        }
    }

    RECT want;
    want.left   = document.left;
    want.top    = document.top - state->stripH;
    want.right  = document.right;
    want.bottom = document.top;

    // Compared against where the strip *actually is*, not against where we last put it. Those are
    // not the same thing - Word, another add-in, or a message we did not see can move it - and
    // trusting our own bookkeeping leaves the strip drawn a pixel or two away from the document
    // frame it is supposed to be flush against, which is visible and was.
    RECT actual;
    if (ChildRect(state->frame, state->strip, &actual) && SameRect(&want, &actual))
    {
        state->stripAt = want;
        state->stripPlaced = TRUE;
        return;
    }

    if (state->stripPlaced && !SameRect(&actual, &state->stripAt))
    {
        LogWrite(L"strip  hwnd=0x%p  strip had drifted: at (%ld,%ld %ldx%ld), expected "
                 L"(%ld,%ld %ldx%ld) - correcting",
                 (void*)state->frame,
                 actual.left, actual.top, actual.right - actual.left, actual.bottom - actual.top,
                 want.left, want.top, want.right - want.left, want.bottom - want.top);
    }

    // SWP_NOZORDER after the first placement, deliberately. The strip is created at the top of the
    // z-order, which is where it needs to be to sit above the frame's background; forcing it back
    // there on every move would also put it above Backstage, which covers the whole client area and
    // is meant to cover us with it.
    SetWindowPos(state->strip, NULL,
                 want.left, want.top,
                 want.right - want.left, state->stripH,
                 SWP_NOZORDER | SWP_NOACTIVATE);

    state->stripAt = want;
    state->stripPlaced = TRUE;
    InvalidateRect(state->strip, NULL, FALSE);
}

// The heart of it: rewrite the position Word is *about* to move `_WwF` to.
//
// Called from inside `_WwF`'s WM_WINDOWPOSCHANGING, so the numbers here are a proposal, not a fact,
// and changing them changes where the window actually lands. Nothing ever flickers because nothing
// is ever wrong for a frame.
static void AdjustProposed(StripState* state, WINDOWPOS* pos)
{
    if (!state->enabled || !pos || !state->wwf)
        return;

    // Freeze while the window is not on screen - hidden, or minimised. Word dismantles a frame's
    // layout on the way out (closing a document was measured to leave `_WwF` filling the whole
    // client area, top edge at 0) and squeezes it to nothing on the way down to the taskbar.
    // Accepting either as a natural rect puts the strip over the ribbon, and if that window is the
    // active one it is then broadcast to every other window in the stack. Nobody can see a window
    // that is not on screen, so there is nothing to gain by shifting it; the janitor re-derives
    // when it comes back.
    if (!IsWindowVisible(state->frame) || IsIconic(state->frame))
        return;

    RECT current;
    if (!ChildRect(state->frame, state->wwf, &current))
        return;

    // Fill in whatever the flags say to leave alone, so the rest of this works on a whole rectangle
    // rather than on four values of which some may be meaningless.
    RECT proposed;
    if (pos->flags & SWP_NOMOVE)
    {
        proposed.left = current.left;
        proposed.top  = current.top;
    }
    else
    {
        proposed.left = pos->x;
        proposed.top  = pos->y;
    }

    LONG width  = (pos->flags & SWP_NOSIZE) ? (current.right - current.left)  : pos->cx;
    LONG height = (pos->flags & SWP_NOSIZE) ? (current.bottom - current.top)  : pos->cy;
    proposed.right  = proposed.left + width;
    proposed.bottom = proposed.top + height;

    if (width <= 0 || height <= 0)
        return;

    // Exactly where we last put it: this is our own placement coming back round, or Word confirming
    // it. Nothing to do, and doing something would be the accumulating shift that walks the frame
    // off the window.
    if (state->hasApplied && SameRect(&proposed, &state->applied))
        return;

    RECT natural, applied;

    if (state->hasApplied && proposed.top == state->applied.top)
    {
        // The top edge is still exactly where we put it, so nobody has re-laid out the top - this
        // proposal was computed from our shifted rect (a width change, a height change, or another
        // add-in shifting the same window). Keep the top, take the other edges. Treating this as a
        // fresh natural rect is what let two processes walk `_WwF` down to 25px tall in spike 2.
        natural = proposed;
        natural.top = state->natural.top;
        applied = proposed;
    }
    else
    {
        // Anything else is Word laying the window out from scratch. What it proposes is natural by
        // definition, and the shift is computed from it - never added to what is already there.
        natural = proposed;
        applied = proposed;
        applied.top = proposed.top + state->stripH;
    }

    // Tripwire. If honouring the shift would leave the document less than two strips tall,
    // something is wrong - most likely someone else shifting the same window - and the right answer
    // is to stop rather than to keep squeezing.
    if ((applied.bottom - applied.top) < (2 * state->stripH))
    {
        if (!state->trippedLogged)
        {
            state->trippedLogged = TRUE;
            LogWrite(L"strip  hwnd=0x%p  TRIPWIRE: shift would leave _WwF %ldpx tall (< 2 x %d) - "
                     L"leaving Word's layout alone. Is another tab add-in shifting the same window?",
                     (void*)state->frame, applied.bottom - applied.top, state->stripH);
        }
        return;
    }

    BOOL naturalMoved = !SameRect(&natural, &state->natural);

    state->natural    = natural;
    state->applied    = applied;
    state->hasApplied = TRUE;
    RememberClient(state);

    pos->x  = applied.left;
    pos->y  = applied.top;
    pos->cx = applied.right - applied.left;
    pos->cy = applied.bottom - applied.top;
    pos->flags &= ~(SWP_NOMOVE | SWP_NOSIZE);

    // In the same message as the document frame moves, rather than waiting for the CHANGED that
    // follows: the two are meant to be flush against each other, so they should move together. The
    // rect is passed in because the document frame has not moved yet - reading it here would place
    // the strip against where it used to be.
    PlaceStrip(state, &applied);

    if (naturalMoved)
    {
        LogRelayout(state, L"relayout");

        // Word has just laid this window out. If it is the focused one it is the only window in
        // the stack Word will lay out at all, so the rest are given the same interior from here.
        StackOnActiveLayout(state->frame, &natural);
    }
}

// ---------------------------------------------------------------------------------------------
// `_WwF`'s subclass procedure.
// ---------------------------------------------------------------------------------------------

static LRESULT CALLBACK WwfSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                                        UINT_PTR idSubclass, DWORD_PTR refData)
{
    (void)idSubclass;

    StripState* state = (StripState*)refData;

    switch (msg)
    {
    case WM_WINDOWPOSCHANGING:
        if (state && state->wwf == hwnd)
            AdjustProposed(state, (WINDOWPOS*)lParam);
        break;

    case WM_WINDOWPOSCHANGED:
        if (state && state->wwf == hwnd)
        {
            const WINDOWPOS* pos = (const WINDOWPOS*)lParam;
            PlaceStrip(state, NULL);

            // Backstage and window minimisation both take `_WwF` away; the strip belongs with it,
            // not with the frame.
            if (pos && (pos->flags & SWP_HIDEWINDOW) && state->strip)
                ShowWindow(state->strip, SW_HIDE);
            else if (pos && (pos->flags & SWP_SHOWWINDOW) && state->strip)
                ShowWindow(state->strip, SW_SHOWNA);
        }
        break;

    case WM_NCDESTROY:
        // Word replaces `_WwF` rather than only moving it - closing a document was measured to
        // destroy one window and reuse another. Unbinding here lets the janitor rebind to whatever
        // replaces it.
        if (state && state->wwf == hwnd)
        {
            LogWrite(L"strip  hwnd=0x%p  _WwF 0x%p destroyed - unbinding", (void*)state->frame, (void*)hwnd);
            state->wwf = NULL;
            state->hasApplied = FALSE;
            state->stripPlaced = FALSE;
        }
        RemoveWindowSubclass(hwnd, WwfSubclassProc, kWwfSubclassId);
        break;

    default:
        break;
    }

    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

// ---------------------------------------------------------------------------------------------
// The strip window.
//
// Painted rather than composed of controls: at this stage it exists to show, without reading a log,
// that we own a horizontal band of Word's layout and that the band tracks the document frame
// through every relayout. The tab it draws is a placeholder for the real tab strip.
// ---------------------------------------------------------------------------------------------

// Where everything in the strip sits, in the strip's client coordinates.
//
// This is the single source of truth for the tab row: painting draws what it returns and
// hit-testing reads what it returns, so there is no way for a click to land somewhere other than
// what the user is looking at. tools\WordLayout.cs mirrors it for the check scripts, and if the two
// ever drift the injected clicks miss and the assertions fail loudly - which is the intended
// failure, rather than a test that quietly clicks the wrong tab and passes.
static void ComputeLayout(StripState* state, const RECT* client, int count, StripLayout* out)
{
    int pad     = Scaled(TAB_LOGICAL_PAD, state->dpi);
    int gap     = Scaled(TAB_LOGICAL_GAP, state->dpi);
    int minimum = Scaled(TAB_LOGICAL_MIN_W, state->dpi);
    int desired = Scaled(TAB_LOGICAL_W, state->dpi);
    int plusW   = Scaled(PLUS_LOGICAL, state->dpi);
    int closeW  = Scaled(CLOSE_LOGICAL, state->dpi);

    if (count > MAX_TABS)
        count = MAX_TABS;
    out->count = count;
    out->hasPlus = FALSE;
    SetRectEmpty(&out->plus);

    // The new-document button's width comes out of the space before the tabs are sized, not after.
    // Tabs shrink as documents are opened; a button does not, and a button that has been squeezed
    // off the end of the strip is a feature the user cannot reach.
    int available = (client->right - client->left) - pad * 2 - plusW - gap;
    if (available < minimum)
        available = minimum;

    int width = desired;
    if (count > 0 && width * count > available)
        width = available / count;
    if (width < minimum)
        width = minimum;

    for (int i = 0; i < count; i++)
    {
        RECT* tab = &out->tab[i];
        tab->left   = client->left + pad + i * width;
        tab->right  = tab->left + width - Scaled(2, state->dpi);   // a hairline between tabs
        tab->top    = client->top + Scaled(3, state->dpi);
        tab->bottom = client->bottom;

        // A close button, but only where there is honestly room for one. A tab narrow enough that
        // the button covers the name is a tab whose button closes a document the user cannot
        // identify, so below that width the name wins and there is no button at all.
        SetRectEmpty(&out->close[i]);
        if (g_buttonsEnabled && (tab->right - tab->left) >= closeW * 3)
        {
            int middle = (tab->top + tab->bottom) / 2;
            out->close[i].right  = tab->right - Scaled(6, state->dpi);
            out->close[i].left   = out->close[i].right - closeW;
            out->close[i].top    = middle - closeW / 2;
            out->close[i].bottom = out->close[i].top + closeW;
        }
    }

    if (!g_buttonsEnabled)
        return;

    // After the last tab while there is room for it, pinned to the right edge once the tabs have
    // filled the strip. Either way it ends up inside the strip and clickable, which is the only
    // property it has to have.
    int after = (count > 0) ? (out->tab[count - 1].right + gap) : (client->left + pad);
    int limit = client->right - pad - plusW;
    if (after > limit)
        after = limit;
    if (after < client->left + pad)
        after = client->left + pad;

    out->plus.left   = after;
    out->plus.right  = after + plusW;
    out->plus.top    = client->top + Scaled(6, state->dpi);
    out->plus.bottom = client->bottom - Scaled(6, state->dpi);
    out->hasPlus = (out->plus.right <= client->right && out->plus.bottom > out->plus.top);
}

// What a point is over. The close button is tested before the tab it sits on: they overlap by
// definition, and the smaller target is the more specific intent.
static StripHit HitTestStrip(StripState* state, HWND hwnd, POINT point)
{
    StripHit hit;
    hit.kind  = HIT_NONE;
    hit.frame = NULL;
    hit.index = -1;
    SetRectEmpty(&hit.tab);

    RECT client;
    if (!GetClientRect(hwnd, &client))
        return hit;

    HWND frames[MAX_STRIPS];
    int count = StackTabs(state->frame, frames, MAX_STRIPS, NULL);

    StripLayout layout;
    ComputeLayout(state, &client, count, &layout);

    for (int i = 0; i < layout.count; i++)
    {
        if (!IsRectEmpty(&layout.close[i]) && PtInRect(&layout.close[i], point))
        {
            hit.kind  = HIT_CLOSE;
            hit.frame = frames[i];
            hit.index = i;
            hit.tab   = layout.tab[i];
            return hit;
        }
        if (PtInRect(&layout.tab[i], point))
        {
            hit.kind  = HIT_TAB;
            hit.frame = frames[i];
            hit.index = i;
            hit.tab   = layout.tab[i];
            return hit;
        }
    }

    if (layout.hasPlus && PtInRect(&layout.plus, point))
        hit.kind = HIT_PLUS;

    return hit;
}

static POINT PointOf(LPARAM lParam)
{
    POINT point;
    point.x = (short)LOWORD(lParam);
    point.y = (short)HIWORD(lParam);
    return point;
}

// Repaint only when what is under the pointer has actually changed. WM_MOUSEMOVE arrives on every
// pixel of movement; invalidating on each one would repaint the strip continuously while the user
// is doing nothing but crossing it on the way to the ribbon.
static void SetHot(StripState* state, HWND hwnd, int kind, HWND frame)
{
    if (state->hotKind == kind && state->hotFrame == frame)
        return;
    state->hotKind  = kind;
    state->hotFrame = frame;
    InvalidateRect(hwnd, NULL, FALSE);
}

// Ask for the one WM_MOUSELEAVE that tells us the pointer has gone. Without it a highlight stays
// lit under a pointer that left the strip a minute ago - there is no "mouse exited" message
// otherwise, and polling for it would be a timer for something an API already answers.
static void ArmLeaveTracking(StripState* state, HWND hwnd)
{
    if (state->tracking)
        return;

    TRACKMOUSEEVENT track;
    memset(&track, 0, sizeof(track));
    track.cbSize    = sizeof(track);
    track.dwFlags   = TME_LEAVE;
    track.hwndTrack = hwnd;
    if (TrackMouseEvent(&track))
        state->tracking = TRUE;
}

// The two glyphs, drawn as lines rather than as characters. A font is not guaranteed to have a
// multiplication sign or a heavy plus at any particular weight, and one that substitutes silently
// gives a close button that looks like a lowercase x. Two lines cannot be substituted.
static void DrawGlyphLines(HDC dc, const RECT* box, COLORREF color, int dpi, BOOL cross)
{
    HPEN pen = CreatePen(PS_SOLID, Scaled(1, dpi), color);
    if (!pen)
        return;

    HGDIOBJ oldPen = SelectObject(dc, pen);

    if (cross)
    {
        int inset = Scaled(5, dpi);
        MoveToEx(dc, box->left + inset, box->top + inset, NULL);
        LineTo(dc, box->right - inset, box->bottom - inset);
        MoveToEx(dc, box->right - inset - 1, box->top + inset, NULL);
        LineTo(dc, box->left + inset - 1, box->bottom - inset);
    }
    else
    {
        int cx  = (box->left + box->right) / 2;
        int cy  = (box->top + box->bottom) / 2;
        int arm = Scaled(5, dpi);
        MoveToEx(dc, cx - arm, cy, NULL);
        LineTo(dc, cx + arm + 1, cy);
        MoveToEx(dc, cx, cy - arm, NULL);
        LineTo(dc, cx, cy + arm + 1);
    }

    SelectObject(dc, oldPen);
    DeleteObject(pen);
}

// A close or new button's background: nothing at rest, a chip under the pointer, a darker one while
// it is held. Slightly larger than the hit rectangle so the glyph is not touching its own edge.
static void DrawChip(HDC dc, const RECT* box, BOOL hot, BOOL down, int dpi)
{
    if (!hot && !down)
        return;
    RECT chip = *box;
    InflateRect(&chip, Scaled(2, dpi), Scaled(2, dpi));
    FillRect(dc, &chip, down ? g_chipDownBrush : g_chipHotBrush);
}

// One tab: its background, its border, its name and its close button.
//
// Its rectangle is a parameter rather than an index into the layout, and that is the whole point: a
// tab being carried is drawn by this same function at wherever the pointer has taken it, so a
// dragged tab cannot end up looking like a different kind of object from a tab sitting still.
static void DrawOneTab(StripState* state, HDC dc, HWND frame,
                       RECT tab, RECT close, BOOL selected, BOOL hot)
{
    HBRUSH fill = selected ? g_tabBrush : (hot ? g_tabHotBrush : g_tabIdleBrush);
    FillRect(dc, &tab, fill);

    HGDIOBJ oldPen   = SelectObject(dc, g_edgePen);
    HGDIOBJ oldBrush = SelectObject(dc, GetStockObject(NULL_BRUSH));
    Rectangle(dc, tab.left, tab.top, tab.right, tab.bottom + 1);
    SelectObject(dc, oldBrush);
    SelectObject(dc, oldPen);

    BOOL hasClose = !IsRectEmpty(&close) && close.right <= tab.right;

    wchar_t title[256];
    WordTabFrameTitle(frame, title, 256);

    RECT text = tab;
    text.left += Scaled(10, state->dpi);
    // The name stops before the button rather than running under it. A title clipped by an
    // ellipsis reads as a long name; one running under a close button reads as a bug.
    text.right = hasClose ? (close.left - Scaled(4, state->dpi))
                          : (tab.right - Scaled(8, state->dpi));
    if (text.right > text.left)
    {
        SetTextColor(dc, selected ? g_textColor : g_idleTextColor);
        DrawTextW(dc, title, -1, &text,
                  DT_SINGLELINE | DT_VCENTER | DT_LEFT | DT_END_ELLIPSIS | DT_NOPREFIX);
    }

    if (hasClose)
    {
        BOOL hotClose  = (state->hotKind == HIT_CLOSE && state->hotFrame == frame);
        BOOL downClose = (state->pressKind == HIT_CLOSE && state->pressFrame == frame);
        DrawChip(dc, &close, hotClose, downClose, state->dpi);
        DrawGlyphLines(dc, &close, hotClose || downClose ? g_glyphHotColor : g_glyphColor,
                       state->dpi, TRUE);
    }
}

// Everything in the strip, onto whatever device context is handed in.
//
// Separate from PaintStrip so the same drawing serves WM_PAINT, the off-screen bitmap it paints
// through, and WM_PRINTCLIENT - which is how the check scripts photograph a strip without a camera.
static void DrawStrip(StripState* state, HDC dc, const RECT* client)
{
    FillRect(dc, client, g_backBrush);

    // The tabs are the stack's, not this window's. Every window in the stack draws the same row
    // with the same one selected, which is what makes switching look like a strip standing still
    // while the page behind it changes. Off a stack this is simply one tab: our own document.
    HWND frames[MAX_STRIPS];
    int  activeIndex = 0;
    int  count = StackTabs(state->frame, frames, MAX_STRIPS, &activeIndex);

    StripLayout layout;
    ComputeLayout(state, client, count, &layout);

    HGDIOBJ oldFont = SelectObject(dc, state->font ? (HGDIOBJ)state->font
                                                   : GetStockObject(DEFAULT_GUI_FONT));
    SetBkMode(dc, TRANSPARENT);

    // The tab being carried, if it is one of ours. Held out of the loop and drawn afterwards, so it
    // is on top of the tabs it is passing over rather than half under them.
    int carried = -1;
    if (g_dragging && g_dragFrame)
    {
        for (int i = 0; i < layout.count; i++)
            if (frames[i] == g_dragFrame)
                carried = i;
    }

    for (int i = 0; i < layout.count; i++)
    {
        if (i == carried)
            continue;

        RECT tab = layout.tab[i];
        if (tab.right <= tab.left || tab.left >= client->right)
            break;
        if (tab.right > client->right)
            tab.right = client->right;

        // A tab with its context menu open is drawn hot for as long as the menu is up, which is
        // the only thing on screen saying which document those commands are about.
        BOOL hotTab = ((state->hotFrame == frames[i]) &&
                       (state->hotKind == HIT_TAB || state->hotKind == HIT_CLOSE)) ||
                      (state->menuFrame && state->menuFrame == frames[i]);

        DrawOneTab(state, dc, frames[i], tab, layout.close[i], (i == activeIndex), hotTab);
    }

    if (carried >= 0)
    {
        // Offset by however far the tab has been taken from the slot it currently occupies. The row
        // has already rearranged underneath it - StackMoveTab runs live during the drag - so this
        // offset shrinks back to nothing as the tab arrives over its new position, and the tab is
        // never drawn in two places or missing from one.
        LONG shift = g_dragLeft - layout.tab[carried].left;

        RECT tab = layout.tab[carried];
        OffsetRect(&tab, shift, 0);

        RECT close = layout.close[carried];
        if (!IsRectEmpty(&close))
            OffsetRect(&close, shift, 0);

        if (tab.left < client->right && tab.right > tab.left)
        {
            if (tab.right > client->right)
                tab.right = client->right;
            DrawOneTab(state, dc, g_dragFrame, tab, close, (carried == activeIndex), TRUE);
        }
    }

    if (layout.hasPlus)
    {
        // Checked against where the button actually is, not just against the flag. The hot flag is
        // written on mouse movement, but the plus *moves* when the number of tabs changes - and the
        // number of tabs changes with the pointer sitting perfectly still, every time a document
        // opens or closes. Emptying the row moves it a whole tab width to the left edge, and without
        // this the strip would light a chip there under nothing at all. Same principle as the strip
        // measuring where it really is rather than trusting where it last put itself.
        BOOL hotPlus  = (state->hotKind == HIT_PLUS);
        if (hotPlus && state->strip && IsWindow(state->strip))
        {
            POINT cursor;
            if (GetCursorPos(&cursor) && ScreenToClient(state->strip, &cursor))
                hotPlus = PtInRect(&layout.plus, cursor) ? TRUE : FALSE;
        }
        BOOL downPlus = (state->pressKind == HIT_PLUS);
        DrawChip(dc, &layout.plus, hotPlus, downPlus, state->dpi);
        DrawGlyphLines(dc, &layout.plus, hotPlus || downPlus ? g_glyphHotColor : g_glyphColor,
                       state->dpi, FALSE);
    }

    SelectObject(dc, oldFont);

    // A hairline along the bottom, so the strip reads as part of Word's chrome rather than as a
    // rectangle dropped on top of it.
    RECT line = *client;
    line.top = line.bottom - 1;
    HBRUSH edge = CreateSolidBrush(g_edgeColor);
    if (edge)
    {
        FillRect(dc, &line, edge);
        DeleteObject(edge);
    }
}

static void PaintStrip(StripState* state, HWND hwnd)
{
    PAINTSTRUCT ps;
    HDC dc = BeginPaint(hwnd, &ps);
    if (!dc)
        return;

    RECT client;
    GetClientRect(hwnd, &client);

    // Drawn into an off-screen bitmap and blitted once. Hover means the strip now repaints whenever
    // the pointer crosses a tab boundary, and painting straight to the screen shows the background
    // fill before the tabs land on it - which at 60 crossings a second is a flicker under the
    // pointer, exactly where the user is looking.
    HDC     mem = CreateCompatibleDC(dc);
    HBITMAP bmp = mem ? CreateCompatibleBitmap(dc, client.right, client.bottom) : NULL;

    if (mem && bmp)
    {
        HGDIOBJ oldBitmap = SelectObject(mem, bmp);
        DrawStrip(state, mem, &client);
        BitBlt(dc, 0, 0, client.right, client.bottom, mem, 0, 0, SRCCOPY);
        SelectObject(mem, oldBitmap);
    }
    else
    {
        DrawStrip(state, dc, &client);
    }

    if (bmp) DeleteObject(bmp);
    if (mem) DeleteDC(mem);

    EndPaint(hwnd, &ps);
}

// The context menu.
//
// A plain Win32 popup, built and thrown away each time it is shown, because everything on it depends
// on what is true at that instant: how many tabs there are, and whether the pointer was over one.
// There is no menu to keep in sync with the tab row if there is no menu between right-clicks.
//
// TPM_RETURNCMD is what makes this small. The chosen id comes back as the return value, so there is
// no WM_COMMAND to route, no id space to keep clear of Word's own thousands of command ids, and no
// window that has to still exist by the time a command arrives.
//
// **Nothing is done from inside here.** TrackPopupMenu runs a modal loop - Word's timers fire, the
// janitor runs, documents can open and close - so by the time it returns, the state this function
// started with may be gone, including `state` itself and the strip window. The command is posted and
// acted on in a fresh message, where everything is looked up again.
static void ShowTabMenu(HWND hwnd, POINT client, HWND target)
{
    StripState* state = (StripState*)GetWindowLongPtrW(hwnd, GWLP_USERDATA);
    if (!state)
        return;

    HMENU menu = CreatePopupMenu();
    if (!menu)
        return;

    int count = StackTabs(state->frame, NULL, MAX_TABS, NULL);

    if (target)
    {
        AppendMenuW(menu, MF_STRING, CMD_SAVE, L"&Save");
        AppendMenuW(menu, MF_SEPARATOR, 0, NULL);
        AppendMenuW(menu, MF_STRING, CMD_CLOSE, L"&Close");
        AppendMenuW(menu, MF_STRING | (count > 1 ? MF_ENABLED : MF_GRAYED),
                    CMD_CLOSE_OTHERS, L"Close &Others");
        AppendMenuW(menu, MF_STRING, CMD_CLOSE_ALL, L"Close &All");
        AppendMenuW(menu, MF_SEPARATOR, 0, NULL);
    }

    // On the empty part of the strip this is the whole menu. A right-click that produces nothing at
    // all reads as a dead area rather than as a deliberate one, and this is the command that has
    // nothing to do with any particular tab.
    AppendMenuW(menu, MF_STRING, CMD_NEW, L"&New Document");

    // The tab stays lit for as long as the menu is up. A tab can be narrow enough that its name is
    // an ellipsis, and "Close All" arriving from a menu the user is no longer sure they aimed
    // correctly is not a comfortable thing to click.
    state->menuFrame = target;
    InvalidateRect(hwnd, NULL, FALSE);
    UpdateWindow(hwnd);                 // painted before the modal loop, not after it

    POINT screen = client;
    ClientToScreen(hwnd, &screen);

    // The owner is Word's frame, not the strip: the strip is WS_EX_NOACTIVATE and can never be the
    // foreground window, and a popup menu whose owner is not foreground does not dismiss when the
    // user clicks away from it. The WM_NULL afterwards is the other half of that rule.
    SetForegroundWindow(state->frame);

    int chosen = (int)TrackPopupMenu(menu,
                                     TPM_RETURNCMD | TPM_NONOTIFY | TPM_LEFTALIGN | TPM_TOPALIGN |
                                     TPM_RIGHTBUTTON,
                                     screen.x, screen.y, 0, state->frame, NULL);
    DestroyMenu(menu);

    // Everything from before the modal loop is re-derived: the strip may have been detached and
    // destroyed while the menu was open, and the state array may have been compacted under us.
    if (!IsWindow(hwnd))
        return;
    state = (StripState*)GetWindowLongPtrW(hwnd, GWLP_USERDATA);
    if (state)
    {
        state->menuFrame = NULL;
        state->hotKind   = HIT_NONE;    // the pointer spent the last few seconds over a menu
        state->hotFrame  = NULL;
        InvalidateRect(hwnd, NULL, FALSE);
        PostMessageW(state->frame, WM_NULL, 0, 0);
    }

    LogWrite(L"strip  hwnd=0x%p  menu on 0x%p -> command %d", (void*)hwnd, (void*)target, chosen);

    if (chosen != 0)
        PostMessageW(hwnd, WM_WORDTAB_CMD, (WPARAM)chosen, (LPARAM)target);
}

// What a right-click is about: the tab under it, or nothing. A close button counts as its own tab -
// the two overlap, and a right-click is not aimed at a button.
static HWND MenuTargetAt(StripState* state, HWND hwnd, POINT point)
{
    StripHit hit = HitTestStrip(state, hwnd, point);
    if (hit.kind == HIT_TAB || hit.kind == HIT_CLOSE)
        return hit.frame;
    return NULL;
}

// ---------------------------------------------------------------------------------------------
// Dragging a tab to reorder it.
//
// The row rearranges *live*, as the tab is carried, rather than showing an insertion marker and
// rearranging on the drop. Both are defensible; live wins here because every window in the stack
// draws the same row, so a live reorder is a thing all of them already know how to show, whereas an
// insertion marker would be state the drawing code would have to be taught. It also means the drop
// itself has nothing to do: by the time the button is released the order is already what the user
// can see, and releasing is only letting go.
//
// The consequence is that cancelling has to undo, which is why the position the tab was picked up
// from is kept for the whole gesture.
// ---------------------------------------------------------------------------------------------

// Forget the gesture without touching the row: a drop that was agreed to, or a strip destroyed
// underneath one.
static void DragForget(void)
{
    g_dragStrip = NULL;
    g_dragFrame = NULL;
    g_dragging  = FALSE;
}

// Put the row back exactly as it was and stop carrying the tab - but keep the capture, because the
// left button is still down and letting the mouse go now would deliver its release to whatever
// happens to be underneath the pointer.
//
// There is deliberately no Escape. The strip is WS_EX_NOACTIVATE and never holds the keyboard focus,
// so a keypress never reaches it, and reading GetAsyncKeyState between two mouse movements is
// exactly the mistake the batch close was written twice to avoid: a state that can come and go
// inside a polling interval has to arrive as an event or it is not being observed at all. The two
// cancels that *are* events - the right button, and losing the capture - both come through here, and
// dragging the tab back where it came from is the third.
static void DragUndo(HWND hwnd, const wchar_t* why)
{
    HWND frame = g_dragFrame;
    BOOL was   = g_dragging;
    int  from  = g_dragFrom;

    g_dragFrame = NULL;
    g_dragging  = FALSE;

    if (!was || !frame || !IsWindow(frame))
        return;

    StackMoveTab(frame, from);
    LogWrite(L"strip  hwnd=0x%p  drag cancelled (%s) - 0x%p back at tab %d",
             (void*)hwnd, why, (void*)frame, from);
    StripRefreshTabs();
}

// The pointer has moved with a tab held down. Below the slop this is still a click that has not
// finished; past it the tab is being carried, and the row rearranges under it.
static void DragMove(StripState* state, HWND hwnd, POINT point)
{
    // The document behind a carried tab can go away mid-gesture: Word hiding the window, a close
    // that was already in flight, the janitor dropping it from the stack. There is then nothing to
    // carry and nowhere to put it back, so the gesture is simply over.
    if (!IsWindow(g_dragFrame) || StackTabIndex(g_dragFrame) < 0)
    {
        BOOL was = g_dragging;
        g_dragFrame = NULL;
        g_dragging  = FALSE;
        if (was)
        {
            LogWrite(L"strip  hwnd=0x%p  drag abandoned - that tab is no longer in the row",
                     (void*)hwnd);
            StripRefreshTabs();
        }
        return;
    }

    int slop = Scaled(DRAG_LOGICAL_SLOP, state->dpi);
    int dx   = point.x - g_dragPressX;

    if (!g_dragging)
    {
        // Horizontal distance only. The row has no vertical meaning - there is no tear-off in this
        // add-in, so dragging a tab downwards is not a different gesture, it is the same one done
        // untidily - and a threshold that counted vertical movement would start a reorder from a
        // hand that slipped while clicking.
        if (dx > -slop && dx < slop)
            return;

        g_dragging = TRUE;
        LogWrite(L"strip  hwnd=0x%p  drag started on 0x%p (tab %d)",
                 (void*)hwnd, (void*)g_dragFrame, g_dragFrom);
    }

    RECT client;
    if (!GetClientRect(hwnd, &client))
        return;

    HWND frames[MAX_STRIPS];
    int count = StackTabs(state->frame, frames, MAX_STRIPS, NULL);

    StripLayout layout;
    ComputeLayout(state, &client, count, &layout);

    // The row this gesture is happening *in* can empty underneath it. The drag state is global and
    // the strip holding the capture is usually not the strip on screen, so `g_dragFrame` above can
    // still be a perfectly good tab in some other window's row while *this* window loses its last
    // document. There is then nowhere to draw the carried tab and no row to drop it into, and simply
    // returning would freeze the gesture: every strip would keep drawing the tab at the last position
    // this function computed, and the release would commit it there.
    if (layout.count <= 0)
    {
        DragUndo(hwnd, L"the row it was being carried in has no documents left");
        return;
    }

    // Where the tab is now: carried from the point inside it that was grabbed, so it does not jump
    // under the pointer when it is picked up, and never past either end of the row - there is
    // nowhere further to go, and a tab drawn off the strip is one being aimed blind.
    int left = point.x - g_dragGrabDx;
    if (left < layout.tab[0].left)
        left = layout.tab[0].left;
    if (left > layout.tab[layout.count - 1].left)
        left = layout.tab[layout.count - 1].left;
    g_dragLeft = left;

    // Which slot it belongs in: the one containing the carried tab's own centre. Measured from the
    // tab rather than from the pointer, so the row swaps when the tab is visibly half way past its
    // neighbour - the pointer can be anywhere along it, and swapping on the pointer makes a tab
    // grabbed by its right-hand edge jump a place the instant it is picked up.
    int centre = left + (layout.tab[0].right - layout.tab[0].left) / 2;
    int target = 0;
    for (int i = 0; i < layout.count; i++)
        if (layout.tab[i].left <= centre)
            target = i;

    StackMoveTab(g_dragFrame, target);   // free, and silent, while the tab is already there
    StripRefreshTabs();
}

static LRESULT CALLBACK StripWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    StripState* state = (StripState*)GetWindowLongPtrW(hwnd, GWLP_USERDATA);

    switch (msg)
    {
    case WM_ERASEBKGND:
        return 1;              // WM_PAINT covers every pixel; erasing first only flickers

    case WM_PAINT:
        if (state)
        {
            PaintStrip(state, hwnd);
            return 0;
        }
        break;

    // How a strip is photographed. DefWindowProc cannot draw a window's client area for it, so
    // without this a PrintWindow of a Word frame comes back with a blank band where the tabs are -
    // and a screenshot that silently omits the thing under test is worse than no screenshot.
    case WM_PRINTCLIENT:
        if (state && wParam)
        {
            RECT client;
            GetClientRect(hwnd, &client);
            DrawStrip(state, (HDC)wParam, &client);
            return 0;
        }
        break;

    case WM_MOUSEMOVE:
        if (state)
        {
            // A held tab takes the whole message. Hover means nothing while the button is down - the
            // pointer is on the tab it is carrying - and the carried tab is its own feedback.
            if (g_dragStrip == hwnd)
            {
                if (g_dragFrame)
                    DragMove(state, hwnd, PointOf(lParam));
                return 0;
            }

            ArmLeaveTracking(state, hwnd);
            StripHit hit = HitTestStrip(state, hwnd, PointOf(lParam));
            SetHot(state, hwnd, hit.kind, hit.frame);
            return 0;
        }
        break;

    case WM_MOUSELEAVE:
        if (state)
        {
            state->tracking = FALSE;
            SetHot(state, hwnd, HIT_NONE, NULL);
            return 0;
        }
        break;

    case WM_LBUTTONDOWN:
        if (state)
        {
            StripHit hit = HitTestStrip(state, hwnd, PointOf(lParam));

            if (hit.kind == HIT_CLOSE || hit.kind == HIT_PLUS)
            {
                // Press and release on the same button, the way every other button in Windows
                // behaves: pressing one and sliding off cancels it. Acting on the press would mean
                // a mis-aimed click closes a document with no way to change your mind.
                state->pressKind  = hit.kind;
                state->pressFrame = hit.frame;
                SetCapture(hwnd);
                InvalidateRect(hwnd, NULL, FALSE);
            }
            else if (hit.kind == HIT_TAB)
            {
                // Switching, though, happens on the press. It is instant, it is reversible by
                // clicking the tab you came from, and waiting for the release makes it feel slow.
                LogWrite(L"strip  hwnd=0x%p  tab clicked -> 0x%p",
                         (void*)state->frame, (void*)hit.frame);
                StackActivate(hit.frame);

                // ...and the same press may turn out to be a drag. Nothing is committed here: the
                // tab is only claimed, and it stays a plain click until the pointer travels far
                // enough to mean something else.
                //
                // The capture is taken *after* the activation, not before. Activating raises a
                // different window, and the one thing that has to be true when this returns is that
                // this strip owns the mouse. A tab that is not in a stack has no row to be reordered
                // within, so it is not picked up at all.
                if (g_dragEnabled && StackTabIndex(hit.frame) >= 0)
                {
                    POINT point  = PointOf(lParam);
                    g_dragStrip  = hwnd;
                    g_dragFrame  = hit.frame;
                    g_dragPressX = point.x;
                    g_dragGrabDx = point.x - hit.tab.left;
                    g_dragLeft   = hit.tab.left;
                    g_dragFrom   = hit.index;
                    g_dragging   = FALSE;
                    SetCapture(hwnd);
                }
            }
            return 0;
        }
        break;

    // Letting go of a tab. The drop has nothing to commit - the row rearranged as the tab was
    // carried - so this is bookkeeping and a log line. It is also where a press that never became a
    // drag ends, which is a plain click and was already handled on the way down.
    case WM_LBUTTONUP:
        if (g_dragStrip == hwnd)
        {
            HWND frame = g_dragFrame;
            BOOL was   = g_dragging;
            int  from  = g_dragFrom;

            // Cleared *before* the capture goes back. ReleaseCapture sends this window a
            // WM_CAPTURECHANGED, and that handler's job is to put an interrupted drag back where it
            // started - which is the exact opposite of what a completed drop means.
            DragForget();
            if (GetCapture() == hwnd)
                ReleaseCapture();

            if (was && frame)
            {
                LogWrite(L"strip  hwnd=0x%p  drag ended: 0x%p is tab %d (was %d)",
                         (void*)hwnd, (void*)frame, StackTabIndex(frame), from);
                StripRefreshTabs();
            }
            return 0;
        }
        if (state && state->pressKind != HIT_NONE)
        {
            int  kind  = state->pressKind;
            HWND frame = state->pressFrame;

            state->pressKind  = HIT_NONE;
            state->pressFrame = NULL;
            if (GetCapture() == hwnd)
                ReleaseCapture();

            StripHit hit = HitTestStrip(state, hwnd, PointOf(lParam));
            if (hit.kind == kind && hit.frame == frame)
            {
                if (kind == HIT_CLOSE)
                {
                    LogWrite(L"strip  hwnd=0x%p  close clicked on 0x%p",
                             (void*)state->frame, (void*)frame);
                    StackCloseTab(frame);
                }
                else
                {
                    LogWrite(L"strip  hwnd=0x%p  new-document button clicked",
                             (void*)state->frame);
                    PostMessageW(hwnd, WM_WORDTAB_CMD, CMD_NEW, 0);
                }
            }
            else
            {
                LogWrite(L"strip  hwnd=0x%p  button released off target - cancelled",
                         (void*)state->frame);
            }

            InvalidateRect(hwnd, NULL, FALSE);
            return 0;
        }
        break;

    // Middle-click closes, which is what a middle click does to a tab everywhere else. Paired the
    // same way as the close button: the release has to land on the tab the press did.
    case WM_MBUTTONDOWN:
        if (g_dragStrip == hwnd)
            return 0;              // a tab is being held; a second button is not a second gesture
        if (state)
        {
            StripHit hit = HitTestStrip(state, hwnd, PointOf(lParam));
            if (hit.kind == HIT_TAB || hit.kind == HIT_CLOSE)
            {
                state->middleFrame = hit.frame;
                SetCapture(hwnd);
            }
            return 0;
        }
        break;

    case WM_MBUTTONUP:
        if (state && state->middleFrame)
        {
            HWND frame = state->middleFrame;
            state->middleFrame = NULL;
            if (GetCapture() == hwnd)
                ReleaseCapture();

            StripHit hit = HitTestStrip(state, hwnd, PointOf(lParam));
            if ((hit.kind == HIT_TAB || hit.kind == HIT_CLOSE) && hit.frame == frame)
            {
                LogWrite(L"strip  hwnd=0x%p  middle-clicked 0x%p", (void*)state->frame, (void*)frame);
                StackCloseTab(frame);
            }
            return 0;
        }
        break;

    // A right press claims a target the same way the other two buttons do, and the menu appears on
    // the release. Consistent with them on purpose: a press that lands on the wrong tab is still
    // taken back by sliding off it before letting go.
    case WM_RBUTTONDOWN:
        // The right button while a tab is held is the cancel, not a menu. The capture is kept: the
        // left button is still down, and the gesture is not over until it is let go of.
        if (g_dragStrip == hwnd)
        {
            DragUndo(hwnd, L"right button");
            return 0;
        }
        if (state && g_menuEnabled)
        {
            state->rightDown  = TRUE;
            state->rightFrame = MenuTargetAt(state, hwnd, PointOf(lParam));
            SetCapture(hwnd);
            return 0;
        }
        break;

    case WM_RBUTTONUP:
        // Swallowed while a tab is held, and not merely ignored. DefWindowProc turns a right release
        // into WM_CONTEXTMENU, and a child window's WM_CONTEXTMENU goes to its parent - so letting
        // this one through would end a cancelled drag by opening Word's own context menu.
        if (g_dragStrip == hwnd)
            return 0;
        if (state && state->rightDown)
        {
            POINT point  = PointOf(lParam);
            HWND  target = state->rightFrame;

            // Capture goes back *before* the menu opens: TrackPopupMenu takes capture itself, and
            // two owners of the mouse is one too many.
            state->rightDown  = FALSE;
            state->rightFrame = NULL;
            if (GetCapture() == hwnd)
                ReleaseCapture();

            if (MenuTargetAt(state, hwnd, point) == target)
                ShowTabMenu(hwnd, point, target);
            else
                LogWrite(L"strip  hwnd=0x%p  right button released off target - no menu",
                         (void*)state->frame);
            return 0;
        }
        break;

    // Capture can be taken away without a button release - a dialog appearing, Alt+Tab, Word
    // starting a modal loop. Anything half-pressed at that point is cancelled, not completed.
    case WM_CAPTURECHANGED:
        // Something took the mouse away mid-gesture: a dialog, Alt+Tab, Word starting a modal loop.
        // A drag that was interrupted was never agreed to, so the row goes back exactly as it was.
        // A drop clears the drag before releasing the capture, so it does not arrive here.
        if (g_dragStrip == hwnd)
        {
            DragUndo(hwnd, L"the mouse capture was taken away");
            DragForget();
        }
        if (state && (state->pressKind != HIT_NONE || state->middleFrame || state->rightDown))
        {
            state->pressKind   = HIT_NONE;
            state->pressFrame  = NULL;
            state->middleFrame = NULL;
            state->rightDown   = FALSE;
            state->rightFrame  = NULL;
            InvalidateRect(hwnd, NULL, FALSE);
        }
        break;

    // Every command the strip can raise, run here rather than where it was chosen: Documents.Add
    // makes a window and pumps messages while it does, Document.Save can put a dialog up, and a
    // close destroys the window whose procedure we would still be inside. All of them come back
    // round to this procedure, so all of them wait for it to have returned.
    case WM_WORDTAB_CMD:
    {
        HWND target = (HWND)lParam;
        if (target && !IsWindow(target))
        {
            LogWrite(L"strip  hwnd=0x%p  command %d dropped - that tab has gone",
                     (void*)hwnd, (int)wParam);
            return 0;
        }

        switch ((int)wParam)
        {
        case CMD_NEW:
            if (!WordTabNewDocument())
                LogWrite(L"strip  hwnd=0x%p  new document declined by Word", (void*)hwnd);
            break;

        case CMD_SAVE:
            // Activate here, save in the next message. Word updates which window is active while it
            // processes the activation, so asking it in this one would be asking before it knows -
            // and the check in WordTabSaveDocument would then refuse a save that was perfectly
            // legitimate. The activation is not optional either way: a Save As dialog owned by a
            // window underneath another at the same rectangle cannot be seen.
            StackActivate(target);
            PostMessageW(hwnd, WM_WORDTAB_SAVE, 0, (LPARAM)target);
            break;

        case CMD_CLOSE:        StackCloseTab(target);    break;
        case CMD_CLOSE_OTHERS: StackCloseOthers(target); break;
        case CMD_CLOSE_ALL:    StackCloseAll(target);    break;
        default: break;
        }
        return 0;
    }

    case WM_WORDTAB_SAVE:
    {
        HWND target = (HWND)lParam;
        if (target && IsWindow(target))
            WordTabSaveDocument(target);
        return 0;
    }

    default:
        break;
    }

    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

static BOOL CreateStripWindow(StripState* state)
{
    if (state->strip && IsWindow(state->strip))
        return TRUE;

    state->strip = CreateWindowExW(WS_EX_NOACTIVATE, kStripClass, L"WordTab",
                                   WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS,
                                   0, 0, 10, state->stripH,
                                   state->frame, NULL, g_module, NULL);
    if (!state->strip)
    {
        LogWrite(L"strip  hwnd=0x%p  CreateWindowEx FAILED (lastError=%lu)",
                 (void*)state->frame, GetLastError());
        return FALSE;
    }

    SetWindowLongPtrW(state->strip, GWLP_USERDATA, (LONG_PTR)state);
    state->stripPlaced = FALSE;
    return TRUE;
}

// ---------------------------------------------------------------------------------------------
// Binding a frame to its `_WwF`.
// ---------------------------------------------------------------------------------------------

struct FindChildArgs
{
    const wchar_t* cls;
    HWND found;
};

static BOOL CALLBACK FindChildProc(HWND hwnd, LPARAM param)
{
    FindChildArgs* args = (FindChildArgs*)param;
    wchar_t cls[64] = L"";
    GetClassNameW(hwnd, cls, 64);
    if (_wcsicmp(cls, args->cls) == 0)
    {
        args->found = hwnd;
        return FALSE;
    }
    return TRUE;
}

static HWND FindChildOfClass(HWND parent, const wchar_t* cls)
{
    FindChildArgs args;
    args.cls = cls;
    args.found = NULL;
    EnumChildWindows(parent, FindChildProc, (LPARAM)&args);
    return args.found;
}

// The name to put on a tab. Word's frame title is "<document> - Word", and the suffix is identical
// on every tab, so it would only eat the width the document name needs.
void WordTabFrameTitle(HWND frame, wchar_t* out, int chars)
{
    if (!out || chars <= 0)
        return;
    out[0] = L'\0';
    if (!frame || !IsWindow(frame))
        return;

    GetWindowTextW(frame, out, chars);
    out[chars - 1] = L'\0';

    wchar_t* suffix = wcsstr(out, L" - Word");
    if (suffix)
        *suffix = L'\0';
    if (out[0] == L'\0')
        wcscpy(out, L"Word");
}

static void ReadTitle(StripState* state)
{
    wchar_t title[256];
    WordTabFrameTitle(state->frame, title, 256);

    if (wcscmp(title, state->title) != 0)
    {
        wcsncpy(state->title, title, 255);
        state->title[255] = L'\0';

        // Every strip shows every tab, so one window's title changing has to repaint all of them.
        StripRefreshTabs();
    }
}

// Apply the shift for the first time. Everything after this is driven by WM_WINDOWPOSCHANGING;
// this is the one place that has to push, because Word will not spontaneously lay out a window just
// because we started caring about it.
static void ApplyInitial(StripState* state)
{
    RECT current;
    if (!ChildRect(state->frame, state->wwf, &current))
        return;

    // Idempotent, like everything else here. If the document frame is already exactly where we put
    // it then nothing has happened that needs re-deriving, and treating what we see as a fresh
    // natural rect would shift it a second time - the strip lands 32px lower and the document with
    // it. That is the accumulating-shift failure the whole design exists to avoid, and it does not
    // stop being possible just because this path is called "initial".
    if (state->hasApplied && SameRect(&current, &state->applied))
    {
        PlaceStrip(state, NULL);
        return;
    }

    RECT applied = current;
    applied.top = current.top + state->stripH;

    if ((applied.bottom - applied.top) < (2 * state->stripH))
    {
        LogWrite(L"strip  hwnd=0x%p  initial shift skipped: _WwF is only %ldpx tall",
                 (void*)state->frame, current.bottom - current.top);
        return;
    }

    // Set the state *before* moving, so the WM_WINDOWPOSCHANGING this call is about to trigger sees
    // its own proposal as already-ours and leaves it alone rather than shifting it a second time.
    state->natural    = current;
    state->applied    = applied;
    state->hasApplied = TRUE;

    SetWindowPos(state->wwf, NULL,
                 applied.left, applied.top,
                 applied.right - applied.left, applied.bottom - applied.top,
                 SWP_NOZORDER | SWP_NOACTIVATE);

    RememberClient(state);
    LogRelayout(state, L"initial");
    PlaceStrip(state, NULL);
    StackOnActiveLayout(state->frame, &state->natural);
}

// ---------------------------------------------------------------------------------------------
// What the stack needs from us.
//
// These exist because of the layout-oracle rule: Word lays out only the focused window, so the
// stack has to take that window's interior and hand it to the others itself. See stack.cpp.
// ---------------------------------------------------------------------------------------------

// Is there a document open in this window?
//
// The obvious test - "does the frame have a `_WwF`" - is wrong, and was wrong here for four slices.
// `_WwF` is the document *frame*, and Word keeps it for the life of the window: closing the last
// document destroys the `_WwB` and `_WwG` inside it and leaves `_WwF` behind, empty. So a Word window
// with nothing open passed that test, joined the stack, and was given a tab labelled "Word".
//
// Measured on 16.0.20228: `_WwF` holds `_WwB` -> `_WwG` in print layout, read mode, web layout,
// draft and outline, with Backstage open, and while minimised - and holds nothing whatsoever when
// there is no document. So the honest question is whether anything is *in* the document frame.
//
// `GetWindow(GW_CHILD)` rather than looking for `_WwB` by name, on two grounds: it is one call
// instead of a recursive enumeration on a half-second timer, and "the document frame is empty" is a
// weaker assumption about Word's internals than any particular class name inside it - a renamed
// child would silently break the name test and cannot break this one.
BOOL StripHasDocument(HWND frame)
{
    StripState* state = FindByFrame(frame);
    HWND wwf = (state && state->wwf && IsWindow(state->wwf))
             ? state->wwf
             // Not bound yet - the janitor may not have come round. Ask the window itself rather
             // than reporting "no document" for what is really "not looked at yet".
             : FindChildOfClass(frame, kWwfClass);

    return (wwf != NULL) && (GetWindow(wwf, GW_CHILD) != NULL);
}

BOOL StripGetNatural(HWND frame, RECT* natural)
{
    StripState* state = FindByFrame(frame);
    if (!state || !state->hasApplied || !natural)
        return FALSE;
    *natural = state->natural;
    return TRUE;
}

// Give a window the document-frame rect that Word gave the focused one. Sound only because the
// stack has already made them the same size.
void StripSetNatural(HWND frame, const RECT* natural)
{
    StripState* state = FindByFrame(frame);
    if (!state || !state->enabled || !natural || !state->wwf || !IsWindow(state->wwf))
        return;

    RECT applied = *natural;
    applied.top = natural->top + state->stripH;
    if ((applied.bottom - applied.top) < (2 * state->stripH))
        return;

    if (state->hasApplied && SameRect(natural, &state->natural) && SameRect(&applied, &state->applied))
        return;

    // State first, then move: the WM_WINDOWPOSCHANGING this triggers then sees its own proposal as
    // already ours and leaves it alone, instead of shifting it a second time.
    state->natural    = *natural;
    state->applied    = applied;
    state->hasApplied = TRUE;
    RememberClient(state);

    SetWindowPos(state->wwf, NULL,
                 applied.left, applied.top,
                 applied.right - applied.left, applied.bottom - applied.top,
                 SWP_NOZORDER | SWP_NOACTIVATE);

    PlaceStrip(state, NULL);
}

// Refit a window's document frame to the window's *own* current size. Needed when a window leaves
// the stack and goes back to the size it had before: Word will not lay out a window it is not
// focused on, so the insets it had are re-applied to the new client area by hand.
void StripRefit(HWND frame)
{
    StripState* state = FindByFrame(frame);
    if (!state || !state->enabled || !state->hasApplied)
        return;

    RECT client;
    if (!GetClientRect(state->frame, &client))
        return;
    if (state->clientAtNatural.cx <= 0 || state->clientAtNatural.cy <= 0)
        return;

    RECT natural;
    natural.left   = state->natural.left;
    natural.top    = state->natural.top;
    natural.right  = (client.right - client.left) - (state->clientAtNatural.cx - state->natural.right);
    natural.bottom = (client.bottom - client.top) - (state->clientAtNatural.cy - state->natural.bottom);

    if (natural.right <= natural.left || natural.bottom <= natural.top)
        return;

    StripSetNatural(frame, &natural);
}

// One window's tab row is every window's tab row, so anything that changes it repaints them all.
void StripRefreshTabs(void)
{
    for (int i = 0; i < g_stripCount; i++)
    {
        StripState* state = &g_strips[i];
        if (state->strip && IsWindow(state->strip))
            InvalidateRect(state->strip, NULL, FALSE);
    }
}

static void TryBind(StripState* state)
{
    if (!g_stripEnabled || !state->enabled)
        return;
    if (state->wwf && IsWindow(state->wwf))
        return;
    if (!IsWindow(state->frame))
        return;

    HWND wwf = FindChildOfClass(state->frame, kWwfClass);
    if (!wwf)
        return;                 // not built yet - the janitor will come back

    state->wwf = wwf;
    state->hasApplied = FALSE;

    // Seeded from what is true now, not left at zero: otherwise the first janitor tick reads a
    // window that has been visible all along as having just appeared, and re-derives a layout that
    // did not need re-deriving.
    state->wasVisible = (IsWindowVisible(state->frame) && !IsIconic(state->frame)) ? TRUE : FALSE;
    state->dpi    = DpiOf(state->frame);
    state->stripH = Scaled(STRIP_LOGICAL_H, state->dpi);
    if (!state->font)
        MakeFont(state);

    // refData carries the state pointer, so the procedure never has to search for it. The entries
    // live in a static array and are never moved, which is what makes that safe.
    if (!SetWindowSubclass(wwf, WwfSubclassProc, kWwfSubclassId, (DWORD_PTR)state))
    {
        LogWrite(L"strip  hwnd=0x%p  SetWindowSubclass on _WwF FAILED (lastError=%lu)",
                 (void*)state->frame, GetLastError());
        state->wwf = NULL;
        return;
    }

    if (!CreateStripWindow(state))
    {
        RemoveWindowSubclass(wwf, WwfSubclassProc, kWwfSubclassId);
        state->wwf = NULL;
        return;
    }

    ReadTitle(state);

    RECT rect;
    ChildRect(state->frame, wwf, &rect);
    LogWrite(L"strip  hwnd=0x%p  bound _WwF=0x%p strip=0x%p  dpi=%d stripH=%d  "
             L"_WwF at (%ld,%ld %ldx%ld)",
             (void*)state->frame, (void*)wwf, (void*)state->strip,
             state->dpi, state->stripH,
             rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top);

    ApplyInitial(state);
}

// Put Word's layout back exactly as we found it. Called on detach and on shutdown - an add-in that
// leaves the host's window layout altered after it is switched off is a bug report waiting to
// happen, and `enabled` has to go down first or the handler would simply shift it again.
static void Restore(StripState* state)
{
    state->enabled = FALSE;

    // A drag whose strip is being destroyed is over. Forgotten rather than undone: a window going
    // away says nothing about whether the user meant the moves they had already made, and
    // DestroyWindow releases the capture, which would otherwise arrive as a cancel.
    if (state->strip && g_dragStrip == state->strip)
        DragForget();

    if (state->strip && IsWindow(state->strip))
    {
        DestroyWindow(state->strip);
        state->strip = NULL;
    }

    if (state->wwf && IsWindow(state->wwf))
    {
        if (state->hasApplied)
        {
            SetWindowPos(state->wwf, NULL,
                         state->natural.left, state->natural.top,
                         state->natural.right - state->natural.left,
                         state->natural.bottom - state->natural.top,
                         SWP_NOZORDER | SWP_NOACTIVATE);
        }
        RemoveWindowSubclass(state->wwf, WwfSubclassProc, kWwfSubclassId);
    }
    state->wwf = NULL;
    state->hasApplied = FALSE;

    if (state->font)
    {
        DeleteObject(state->font);
        state->font = NULL;
    }
}

// ---------------------------------------------------------------------------------------------
// The janitor.
//
// The geometry itself is event-driven and needs no timer. This exists for the things that have no
// event: `_WwF` not existing yet when a frame is first subclassed, `_WwF` being replaced under us,
// and the document title changing (Word does not always send the frame a WM_SETTEXT we can see).
// Half a second is slow enough to be free and fast enough that nothing is visibly late.
// ---------------------------------------------------------------------------------------------

static void CALLBACK JanitorProc(HWND hwnd, UINT msg, UINT_PTR id, DWORD tick)
{
    (void)hwnd; (void)msg; (void)id; (void)tick;

    for (int i = 0; i < g_stripCount; i++)
    {
        StripState* state = &g_strips[i];
        if (!state->enabled || !IsWindow(state->frame))
            continue;

        if (!state->wwf || !IsWindow(state->wwf))
        {
            TryBind(state);
            continue;
        }

        // Hidden -> shown. Everything Word did to this window's layout while it was hidden was
        // ignored on purpose (see AdjustProposed), so what is there now is Word's own idea of the
        // layout and is exactly what "natural" means. Re-derive from it rather than carrying
        // forward a rect from before the window went away.
        // "On screen", not merely WS_VISIBLE: a minimised window keeps that style, and its layout
        // is frozen for the same reason a hidden one's is.
        BOOL visible = (IsWindowVisible(state->frame) && !IsIconic(state->frame)) ? TRUE : FALSE;
        if (visible && !state->wasVisible)
        {
            // ApplyInitial does nothing if the document frame is still where we put it, which is
            // the common case - Word usually hides and shows a window without touching its layout.
            // It only re-derives when Word actually did rearrange things while we were not looking.
            state->stripPlaced = FALSE;
            ApplyInitial(state);
        }
        state->wasVisible = visible;

        ReadTitle(state);

        // Cheap when nothing has moved - it compares against the strip's real rect and returns.
        // This is what makes any drift heal within half a second rather than staying wrong.
        if (visible)
            PlaceStrip(state, NULL);

        // The strip belongs with the document frame: shown when it is shown, hidden when Backstage
        // or a minimise takes it away.
        if (state->strip && IsWindow(state->strip))
        {
            BOOL wantVisible = IsWindowVisible(state->wwf) ? TRUE : FALSE;
            BOOL isVisible   = IsWindowVisible(state->strip) ? TRUE : FALSE;
            if (wantVisible != isVisible)
                ShowWindow(state->strip, wantVisible ? SW_SHOWNA : SW_HIDE);
        }
    }

    // Membership is decided from what is true right now - visible, has a document frame - so it is
    // re-decided on the same cadence rather than tracked through events that Word does not always
    // send.
    StackJanitor();
}

// ---------------------------------------------------------------------------------------------
// Entry points, called from frames.cpp.
// ---------------------------------------------------------------------------------------------

void StripStart(void)
{
    g_stripEnabled = WordTabReadFlag(L"Strip", TRUE);

    if (!g_stripEnabled)
    {
        LogWrite(L"StripStart  disabled by HKCU\\Software\\WordTab\\Strip=0");
        return;
    }

    // The close and new-document buttons, switchable like every other piece. Off, the tabs are
    // exactly what they were before this slice - which is what makes them bisectable if something
    // about them ever misbehaves.
    g_buttonsEnabled = WordTabReadFlag(L"TabButtons", TRUE);

    // The context menu, on its own switch. Off, a right-click on the strip does nothing at all -
    // which is what it did before this slice.
    g_menuEnabled = WordTabReadFlag(L"TabMenu", TRUE);

    // Dragging a tab to reorder it. Off, a press on a tab switches to it and nothing else, which is
    // what it did before this slice - and the mouse capture is never taken, so the whole mechanism
    // is out of the way rather than merely inert.
    g_dragEnabled = WordTabReadFlag(L"TabDrag", TRUE);

    if (!g_stripClass)
    {
        WNDCLASSEXW wc;
        memset(&wc, 0, sizeof(wc));
        wc.cbSize        = sizeof(wc);
        wc.style         = CS_HREDRAW | CS_VREDRAW;
        wc.lpfnWndProc   = StripWndProc;
        wc.hInstance     = g_module;
        wc.hCursor       = LoadCursorW(NULL, IDC_ARROW);
        wc.lpszClassName = kStripClass;
        g_stripClass = RegisterClassExW(&wc);
    }

    g_darkTheme = DarkThemeInUse();
    if (g_darkTheme)
    {
        g_edgeColor     = RGB(77, 77, 77);
        g_textColor     = RGB(255, 255, 255);
        g_idleTextColor = RGB(186, 186, 186);
        g_glyphColor    = RGB(186, 186, 186);
        g_glyphHotColor = RGB(255, 255, 255);
        if (!g_backBrush)     g_backBrush     = CreateSolidBrush(RGB(38, 38, 38));
        if (!g_tabBrush)      g_tabBrush      = CreateSolidBrush(RGB(66, 66, 66));
        if (!g_tabIdleBrush)  g_tabIdleBrush  = CreateSolidBrush(RGB(45, 45, 45));
        if (!g_tabHotBrush)   g_tabHotBrush   = CreateSolidBrush(RGB(55, 55, 55));
        if (!g_chipHotBrush)  g_chipHotBrush  = CreateSolidBrush(RGB(90, 90, 90));
        if (!g_chipDownBrush) g_chipDownBrush = CreateSolidBrush(RGB(112, 112, 112));
    }
    else
    {
        g_edgeColor     = RGB(200, 198, 196);
        g_textColor     = RGB(50, 49, 48);
        g_idleTextColor = RGB(96, 94, 92);
        g_glyphColor    = RGB(96, 94, 92);
        g_glyphHotColor = RGB(32, 31, 30);
        if (!g_backBrush)     g_backBrush     = CreateSolidBrush(RGB(237, 235, 233));
        if (!g_tabBrush)      g_tabBrush      = CreateSolidBrush(RGB(255, 255, 255));
        if (!g_tabIdleBrush)  g_tabIdleBrush  = CreateSolidBrush(RGB(225, 223, 221));
        if (!g_tabHotBrush)   g_tabHotBrush   = CreateSolidBrush(RGB(240, 238, 236));
        if (!g_chipHotBrush)  g_chipHotBrush  = CreateSolidBrush(RGB(205, 203, 201));
        if (!g_chipDownBrush) g_chipDownBrush = CreateSolidBrush(RGB(188, 186, 184));
    }
    if (!g_edgePen) g_edgePen = CreatePen(PS_SOLID, 1, g_edgeColor);

    // A thread timer rather than a window timer: it needs no window of its own, and Word's message
    // loop dispatches it to the callback like any other.
    if (!g_janitor)
        g_janitor = SetTimer(NULL, 0, 500, JanitorProc);

    LogWrite(L"StripStart  class=%s janitor=%s  stripH=%d logical px  theme=%s  tab buttons=%s  "
             L"tab menu=%s  tab drag=%s",
             g_stripClass ? L"registered" : L"FAILED",
             g_janitor ? L"running" : L"FAILED", STRIP_LOGICAL_H,
             g_darkTheme ? L"dark" : L"light",
             g_buttonsEnabled ? L"on" : L"off (HKCU\\Software\\WordTab\\TabButtons=0)",
             g_menuEnabled ? L"on" : L"off (HKCU\\Software\\WordTab\\TabMenu=0)",
             g_dragEnabled ? L"on" : L"off (HKCU\\Software\\WordTab\\TabDrag=0)");
}

void StripAttachFrame(HWND frame)
{
    if (!g_stripEnabled || !frame)
        return;
    if (FindByFrame(frame))
        return;
    if (g_stripCount >= MAX_STRIPS)
        return;

    StripState* state = &g_strips[g_stripCount++];
    memset(state, 0, sizeof(*state));
    state->frame   = frame;
    state->enabled = TRUE;
    state->dpi     = DpiOf(frame);
    state->stripH  = Scaled(STRIP_LOGICAL_H, state->dpi);
    MakeFont(state);

    TryBind(state);
}

void StripDetachFrame(HWND frame)
{
    StripState* state = FindByFrame(frame);
    if (!state)
        return;

    // If the frame itself is on its way out there is nothing to restore and its children are gone
    // or going; touching them is at best pointless and at worst a crash.
    if (IsWindow(frame))
        Restore(state);
    else
    {
        state->enabled = FALSE;
        if (state->strip && g_dragStrip == state->strip)
            DragForget();
        state->wwf = NULL;
        state->strip = NULL;
        if (state->font) { DeleteObject(state->font); state->font = NULL; }
    }

    int index = (int)(state - g_strips);
    for (int i = index; i < g_stripCount - 1; i++)
        g_strips[i] = g_strips[i + 1];
    g_stripCount--;

    // The entries after the removed one have shifted, so every `_WwF` subclass still carrying a
    // refData pointer into this array now points at the wrong entry. Re-point them.
    for (int i = index; i < g_stripCount; i++)
    {
        if (g_strips[i].wwf && IsWindow(g_strips[i].wwf))
        {
            SetWindowSubclass(g_strips[i].wwf, WwfSubclassProc, kWwfSubclassId,
                              (DWORD_PTR)&g_strips[i]);
        }
        if (g_strips[i].strip && IsWindow(g_strips[i].strip))
            SetWindowLongPtrW(g_strips[i].strip, GWLP_USERDATA, (LONG_PTR)&g_strips[i]);
    }
}

void StripOnFrameDpiChanged(HWND frame)
{
    StripState* state = FindByFrame(frame);
    if (!state || !state->enabled)
        return;

    int dpi = DpiOf(frame);
    if (dpi == state->dpi)
        return;

    LogWrite(L"strip  hwnd=0x%p  DPI %d -> %d, re-scaling the strip", (void*)frame, state->dpi, dpi);

    state->dpi    = dpi;
    state->stripH = Scaled(STRIP_LOGICAL_H, dpi);
    MakeFont(state);

    // The old shift was computed at the old scale, so the current rect is not a natural one. Undo
    // it, then let the next layout - or the janitor - re-apply at the new height.
    if (state->wwf && IsWindow(state->wwf) && state->hasApplied)
    {
        state->hasApplied = FALSE;
        state->stripPlaced = FALSE;
        SetWindowPos(state->wwf, NULL,
                     state->natural.left, state->natural.top,
                     state->natural.right - state->natural.left,
                     state->natural.bottom - state->natural.top,
                     SWP_NOZORDER | SWP_NOACTIVATE);
        ApplyInitial(state);
    }
}

void StripStop(void)
{
    if (g_janitor)
    {
        KillTimer(NULL, g_janitor);
        g_janitor = 0;
    }

    for (int i = 0; i < g_stripCount; i++)
        Restore(&g_strips[i]);
    g_stripCount = 0;

    // The window class is not unregistered here: frames.cpp's FramesStop does the same for its
    // coordinator class, and the module is pinned in the process anyway.
    LogWrite(L"StripStop  done");
}
