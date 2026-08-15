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

    wchar_t title[256];
};

static StripState g_strips[MAX_STRIPS];
static int  g_stripCount = 0;
static BOOL g_stripEnabled = TRUE;
static ATOM g_stripClass = 0;
static UINT_PTR g_janitor = 0;

static HBRUSH   g_backBrush    = NULL;   // strip background
static HBRUSH   g_tabBrush     = NULL;   // the selected tab
static HBRUSH   g_tabIdleBrush = NULL;   // the others
static HPEN     g_edgePen      = NULL;
static COLORREF g_edgeColor    = RGB(200, 198, 196);
static COLORREF g_textColor    = RGB(50, 49, 48);
static COLORREF g_idleTextColor = RGB(96, 94, 92);
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

// Where tab `index` of `count` sits, in the strip's client coordinates. One function for painting
// and for hit-testing, so a click can never land somewhere other than what was drawn.
static void TabRect(StripState* state, int index, int count, const RECT* client, RECT* out)
{
    int pad     = Scaled(6, state->dpi);
    int minimum = Scaled(70, state->dpi);
    int desired = Scaled(TAB_LOGICAL_W, state->dpi);

    int available = (client->right - client->left) - pad * 2;
    int width = desired;
    if (count > 0 && width * count > available)
        width = available / count;
    if (width < minimum)
        width = minimum;

    out->left   = client->left + pad + index * width;
    out->right  = out->left + width - Scaled(2, state->dpi);   // a hairline between tabs
    out->top    = client->top + Scaled(3, state->dpi);
    out->bottom = client->bottom;
}

static void PaintStrip(StripState* state, HWND hwnd)
{
    PAINTSTRUCT ps;
    HDC dc = BeginPaint(hwnd, &ps);
    if (!dc)
        return;

    RECT client;
    GetClientRect(hwnd, &client);

    FillRect(dc, &client, g_backBrush);

    // The tabs are the stack's, not this window's. Every window in the stack draws the same row
    // with the same one selected, which is what makes switching look like a strip standing still
    // while the page behind it changes. Off a stack this is simply one tab: our own document.
    HWND frames[MAX_STRIPS];
    int  activeIndex = 0;
    int  count = StackTabs(state->frame, frames, MAX_STRIPS, &activeIndex);

    HGDIOBJ oldFont = SelectObject(dc, state->font ? (HGDIOBJ)state->font
                                                   : GetStockObject(DEFAULT_GUI_FONT));
    SetBkMode(dc, TRANSPARENT);

    for (int i = 0; i < count; i++)
    {
        RECT tab;
        TabRect(state, i, count, &client, &tab);
        if (tab.right <= tab.left || tab.left >= client.right)
            break;
        if (tab.right > client.right)
            tab.right = client.right;

        BOOL selected = (i == activeIndex);
        FillRect(dc, &tab, selected ? g_tabBrush : g_tabIdleBrush);

        HGDIOBJ oldPen   = SelectObject(dc, g_edgePen);
        HGDIOBJ oldBrush = SelectObject(dc, GetStockObject(NULL_BRUSH));
        Rectangle(dc, tab.left, tab.top, tab.right, tab.bottom + 1);
        SelectObject(dc, oldBrush);
        SelectObject(dc, oldPen);

        wchar_t title[256];
        WordTabFrameTitle(frames[i], title, 256);

        RECT text = tab;
        text.left  += Scaled(10, state->dpi);
        text.right -= Scaled(8, state->dpi);
        SetTextColor(dc, selected ? g_textColor : g_idleTextColor);
        DrawTextW(dc, title, -1, &text,
                  DT_SINGLELINE | DT_VCENTER | DT_LEFT | DT_END_ELLIPSIS | DT_NOPREFIX);
    }

    SelectObject(dc, oldFont);

    // A hairline along the bottom, so the strip reads as part of Word's chrome rather than as a
    // rectangle dropped on top of it.
    RECT line = client;
    line.top = line.bottom - 1;
    HBRUSH edge = CreateSolidBrush(g_edgeColor);
    if (edge)
    {
        FillRect(dc, &line, edge);
        DeleteObject(edge);
    }

    EndPaint(hwnd, &ps);
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

    case WM_LBUTTONDOWN:
        if (state)
        {
            RECT client;
            GetClientRect(hwnd, &client);

            POINT point;
            point.x = (short)LOWORD(lParam);
            point.y = (short)HIWORD(lParam);

            HWND frames[MAX_STRIPS];
            int count = StackTabs(state->frame, frames, MAX_STRIPS, NULL);
            for (int i = 0; i < count; i++)
            {
                RECT tab;
                TabRect(state, i, count, &client, &tab);
                if (PtInRect(&tab, point))
                {
                    LogWrite(L"strip  hwnd=0x%p  tab %d clicked -> 0x%p",
                             (void*)state->frame, i, (void*)frames[i]);
                    StackActivate(frames[i]);
                    break;
                }
            }
            return 0;
        }
        break;

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

BOOL StripHasDocumentFrame(HWND frame)
{
    StripState* state = FindByFrame(frame);
    if (state && state->wwf && IsWindow(state->wwf))
        return TRUE;

    // Not bound yet - the janitor may not have come round. Ask the window itself rather than
    // reporting "no document" for what is really "not looked at yet".
    return FindChildOfClass(frame, kWwfClass) != NULL;
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
        if (!g_backBrush)    g_backBrush    = CreateSolidBrush(RGB(38, 38, 38));
        if (!g_tabBrush)     g_tabBrush     = CreateSolidBrush(RGB(66, 66, 66));
        if (!g_tabIdleBrush) g_tabIdleBrush = CreateSolidBrush(RGB(45, 45, 45));
    }
    else
    {
        g_edgeColor     = RGB(200, 198, 196);
        g_textColor     = RGB(50, 49, 48);
        g_idleTextColor = RGB(96, 94, 92);
        if (!g_backBrush)    g_backBrush    = CreateSolidBrush(RGB(237, 235, 233));
        if (!g_tabBrush)     g_tabBrush     = CreateSolidBrush(RGB(255, 255, 255));
        if (!g_tabIdleBrush) g_tabIdleBrush = CreateSolidBrush(RGB(225, 223, 221));
    }
    if (!g_edgePen) g_edgePen = CreatePen(PS_SOLID, 1, g_edgeColor);

    // A thread timer rather than a window timer: it needs no window of its own, and Word's message
    // loop dispatches it to the callback like any other.
    if (!g_janitor)
        g_janitor = SetTimer(NULL, 0, 500, JanitorProc);

    LogWrite(L"StripStart  class=%s janitor=%s  stripH=%d logical px  theme=%s",
             g_stripClass ? L"registered" : L"FAILED",
             g_janitor ? L"running" : L"FAILED", STRIP_LOGICAL_H,
             g_darkTheme ? L"dark" : L"light");
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
