// WordTab - the stack: several Word windows made to look like one window with tabs.
//
// This is the product claim. Spike 1 carved a strip out of one window; the previous slice moved
// that in-process. A strip on one window is a stripe, not a tabbed interface. What makes it a
// tabbed interface is that N real OpusApp windows sit at exactly one rectangle, in z-order, with
// every strip drawing the same row of tabs - so switching documents looks like a tab strip staying
// still while the page behind it changes, rather than three windows swapping places.
//
// Spike 2 proved all of that out of process (spikes\StackSpike\RESULT.md). Two things it found are
// load-bearing here and are not obvious:
//
//   1. **Word only lays out a window's interior while that window has focus.** Force-resize a
//      stacked window that is not focused and Word moves its ribbon and status bar but leaves
//      `_WwF` completely alone - measured, and activating the window later does not repair it. In a
//      stack only one window is ever focused, so waiting for Word to lay out the other N-1 is
//      waiting forever. The fix: the focused window is the **layout oracle**, and every other
//      window's `_WwF` is set to the rect Word gave the focused one. Sound only because stacking
//      already guarantees they are the same size.
//
//   2. **Dragging is where out-of-process lost.** Spike 2 could only reposition the followers once
//      per 30ms poll, so they lagged and peeked out from under the dragged window - the user's
//      "it really glitches out". Its workaround was to hide them for the duration. In-process we
//      are called inside Word's modal loop, before each move, so the followers move in the same
//      frame and there is nothing to hide.
//
// Membership rules, which are subtler than they look. Frame lifetime is not document lifetime:
// closing one of two documents was measured to hide one frame and destroy a *different* one, and
// Word creates frames it never shows. So a window is in the stack while it is visible and has a
// document open in it, and it leaves when either stops being true - never keyed on creation and
// destruction.
//
// And the *document frame*'s lifetime is not the document's either, which cost this file four
// slices of being quietly wrong. `_WwF` is where a document goes, and Word keeps it standing and
// empty after the last document in that window closes - so "has a `_WwF`" was answering TRUE for a
// window with nothing in it, which then joined the stack and was given a tab labelled "Word". The
// test is StripHasDocument, and it looks *inside* the document frame. See RESULT-startscreen.md.

#include "wordtab.h"
#include <commctrl.h>
#include <string.h>
#include <wchar.h>
#include <stdio.h>

#define MAX_MEMBERS 256

struct Member
{
    HWND frame;
    BOOL joined;
    RECT joinRect;      // where the window was before we stacked it, so it can be put back
    BOOL joinZoomed;
    int  repairs;       // consecutive attempts to put this one back in step - see Reconcile

    // The user pulled this window out of the stack and it is to stay out. Every other reason a
    // member is not joined is a fact about the window that the janitor re-tests twice a second -
    // hidden, minimised, no document - and it rejoins the moment that fact changes. This one is not
    // a fact about the window; it is a decision about it, and nothing the window does can revoke it.
    // See StackTearOffTab.
    BOOL tornOff;

    // Whatever document turns up in this frame next is a NEW one, so its tab belongs at the end of
    // the row rather than at the place this frame used to hold.
    //
    // Set while a frame is out of the row and holding no document - Word does not destroy a frame
    // when its last document closes, it hides it and puts the next document straight into it, so a
    // member can come back carrying something the user has never seen. Without this the new document
    // silently inherits the old one's position: drag a tab out, close it, press + and the new
    // document arrives in the middle of the row.
    //
    // Deliberately NOT set for every leave. A minimised window still holds its document, and a user
    // who minimises Word and restores it must find the row exactly as they left it - a rule that sent
    // every rejoining window to the end would shuffle the tabs every time one blinked.
    BOOL rejoin;

    // Consecutive janitor ticks on which this window - already in the row - has failed the test to
    // stay in it. A tab is not given up on the first one.
    //
    // The user's report: "changing setting like view>page width closes tabs", and then, when asked
    // whether the document had gone with it, "oh no it didnt it just opened a new word w/out it".
    // So the window is EVICTED and comes back a plain Word window with no strip. A routine Word UI
    // action costing somebody the whole row is the worst thing in the queue, and the shape of it is
    // one this project has now met four times: a state that is true for a moment is read as a state
    // that is true. See [[transient-state-must-be-an-event]] - the honest fix is to hear the event,
    // and there is no event here, so the next best thing is to refuse to act on one reading.
    //
    // Held as the MOMENT it started rather than as a count of ticks, and that is not a detail. The
    // janitor is not only the timer: WM_SHOWWINDOW calls it too, so a window being hidden produces
    // two passes a few milliseconds apart, and a count of two was reached inside a quarter of a
    // second by a window that was fine - measured, tools\check-row.ps1, first run. A window has to be
    // wrong for a length of TIME, not for a number of looks, or the guard is only as good as how
    // often something happened to ask.
    DWORD missedAt;
};

// Longer than the janitor's own half second, so that at least one later tick has to agree before a
// tab goes. Anything shorter can be satisfied by two passes of the same moment.
#define STAY_GRACE_MS 700

static void Reconcile(void);
static void CloseBatchStep(void);
static void CloseBatchEnd(const wchar_t* why);


// **The order of this array is the order of the tabs.** There is no separate order field, and that
// is a decision rather than an omission: an `order` int has to be kept consistent with the array by
// every path that adds a member, drops one, or compacts the array after a close, and any one of them
// getting it wrong produces two tabs claiming the same position. An array cannot disagree with
// itself about what order it is in.
//
// It also makes the rest of this file read in tab order for free: FirstJoined becomes the leftmost
// tab rather than the earliest-joined one, and a batch close runs left to right. Joining appends, so
// a new document arrives at the end of the row, which is where every other tabbed application puts
// one. See StackMoveTab.
static Member g_members[MAX_MEMBERS];
static int    g_memberCount = 0;      // frames we know about, joined or not
static HWND   g_active      = NULL;
static BOOL   g_onePage     = TRUE;    // HKCU\Software\WordTab\OnePage
static BOOL   g_enabled     = TRUE;
static BOOL   g_started     = FALSE;
static BOOL   g_inSync      = FALSE;  // our own SetWindowPos calls come back through the subclass
static BOOL   g_altTab      = TRUE;
static BOOL   g_tearOff      = TRUE;
static BOOL   g_rowSize      = TRUE;  // the row comes back the size it was left, not Word's
static int    g_closeStack   = 1;     // 0 = one document, 1 = ask, 2 = the whole stack
static BOOL   g_minimized   = FALSE;  // the whole stack is down on the taskbar

// Where to put the user back when the active tab goes away, set only by StackCloseTab. Closing a
// background tab has to activate it first (see there), which moves the user off the document they
// were reading; this is how they get back to it rather than to whatever happens to be next in the
// row. Always re-validated before use - the window may have closed in the meantime.
static HWND   g_returnTo    = NULL;

// The last window WE posted a WM_CLOSE to, single tab or batch.
//
// It exists so the janitor can tell a document that has gone because the user closed it from one
// that merely looks gone for a moment. The row waits a tick before giving up on a window (see
// STAY_MISSES); this is the one case where waiting would be wrong, because the disappearance is the
// answer to something we asked and the user is watching for it.
//
// Deliberately not cleared on success - it is only ever compared against a window that is still a
// member, and a handle Word has reused belongs to a window that has just been attached and so cannot
// be in the row yet. It is overwritten when the next close is aimed somewhere else.
static HWND   g_closeAimed  = NULL;

static BOOL CloseWasAskedFor(HWND frame)
{
    return (frame && frame == g_closeAimed) ? TRUE : FALSE;
}

// The row, as the user sees it, written out when it changes and never otherwise.
//
// Two of the things in the queue are about ORDER - a window that leaves the row and does not come
// back, and a new document whose tab arrives at the front instead of the end - and neither of them
// was answerable from the log that reported them. There were lines for a window joining and a window
// leaving, and none at all for what the row then WAS. A membership line that does not say what the
// membership is cannot settle an argument about ordering.
//
// Only on change, by comparing the line it would write against the last one it wrote, so a row that
// is sitting still costs a snprintf twice a second and no file write at all.
static void LogRow(void)
{
    wchar_t line[1024];
    int at = 0;
    line[0] = L'\0';

    for (int i = 0; i < g_memberCount; i++)
    {
        if (!g_members[i].joined)
            continue;

        wchar_t name[128];
        StripTabName(g_members[i].frame, name, 128);

        int wrote = _snwprintf(line + at, (size_t)(1024 - at), L"%s0x%p |%s|",
                               at ? L"  " : L"", (void*)g_members[i].frame, name);
        if (wrote < 0 || at + wrote >= 1023)
            break;
        at += wrote;
    }

    static wchar_t last[1024] = L"";
    if (wcscmp(line, last) == 0)
        return;

    wcsncpy(last, line, 1023);
    last[1023] = L'\0';

    if (at == 0)
        LogWrite(L"stack  row: empty");
    else
        LogWrite(L"stack  row, left to right: %s", line);
}

static Member* Find(HWND frame)
{
    for (int i = 0; i < g_memberCount; i++)
        if (g_members[i].frame == frame)
            return &g_members[i];
    return NULL;
}

static int JoinedCount(void)
{
    int count = 0;
    for (int i = 0; i < g_memberCount; i++)
        if (g_members[i].joined)
            count++;
    return count;
}

// Joining and staying are different tests, and the difference is minimising.
//
// To *join*, a window has to be one we can measure and place: on screen, not minimised, a sensible
// size, and with a document open in it. A window with no document is deliberately left out - Word's
// Start screen, and the empty frame Word leaves behind when you close its last document - because
// it has no document to be a tab for. It keeps its own taskbar button and Alt+Tab entry, the stack
// never moves it, and it joins the row by itself the moment a document appears in it.
//
// That was always the intent; until this slice it was not the behaviour, because the test asked
// whether the window had a `_WwF` and Word keeps the document frame after the document has gone.
// See StripHasDocument in strip.cpp for what is actually measured.
//
// To *stay*, all that is required is that it still exists and still has a document. A minimised
// window is still a document and still deserves its tab; more to the point, when the whole stack
// goes down to the taskbar together, dropping every window out of the stack would leave nothing to
// bring back and the only window with a taskbar button would come back alone. (Measured: a
// minimised document keeps `_WwB` inside its `_WwF`, so it still passes the document test.)
static BOOL EligibleToStay(HWND frame)
{
    if (!frame || !IsWindow(frame) || !IsWindowVisible(frame))
        return FALSE;
    return StripHasDocument(frame);
}

static BOOL EligibleToJoin(HWND frame)
{
    if (!EligibleToStay(frame) || IsIconic(frame))
        return FALSE;

    RECT rect;
    if (!GetWindowRect(frame, &rect))
        return FALSE;
    if ((rect.right - rect.left) < 200 || (rect.bottom - rect.top) < 200)
        return FALSE;

    return TRUE;
}

static HWND FirstJoined(HWND except)
{
    for (int i = 0; i < g_memberCount; i++)
        if (g_members[i].joined && g_members[i].frame != except && IsWindow(g_members[i].frame))
            return g_members[i].frame;
    return NULL;
}

// ---------------------------------------------------------------------------------------------
// Geometry.
// ---------------------------------------------------------------------------------------------

// Maximize or restore a window *without activating it*.
//
// `ShowWindow(SW_MAXIMIZE)` and `SW_RESTORE` both activate, and activating a window behind the one
// the user is looking at pulls it in front and hands it the keyboard. `SetWindowPlacement` changes
// the same state and does not activate, which is the only reason it is used here.
static void SetZoomState(HWND frame, BOOL zoom)
{
    WINDOWPLACEMENT placement;
    memset(&placement, 0, sizeof(placement));
    placement.length = sizeof(placement);
    if (!GetWindowPlacement(frame, &placement))
        return;

    UINT wanted = zoom ? SW_SHOWMAXIMIZED : SW_SHOWNORMAL;
    if (placement.showCmd == wanted)
        return;

    placement.showCmd = wanted;
    placement.flags = 0;
    SetWindowPlacement(frame, &placement);
}

// Make `frame` match `master` exactly: same rect, and the same maximized-or-not *state* rather
// than only the same rectangle. Spike 2 could copy rectangles only, which left a stacked window
// looking maximized without being maximized - and it is not even the same rectangle: a maximized
// window covers the taskbar's strip of screen and a window merely resized to "maximized size" does
// not, so the two differ by the height of the taskbar. In-process the state itself is reachable.
static void MatchTo(HWND master, HWND frame)
{
    if (!IsWindow(master) || !IsWindow(frame) || master == frame)
        return;

    g_inSync = TRUE;

    // State first, then rectangle, and the rectangle **always**. Two windows can both be maximized
    // and still not be the same size - measured: a maximized window that the shell classifies
    // differently gets the full screen where another gets the work area, 72px shorter. Setting the
    // state alone leaves them looking like two windows; copying the rect alone leaves a window that
    // looks maximized without being maximized, which is what spike 2 could not fix.
    SetZoomState(frame, IsZoomed(master) ? TRUE : FALSE);

    // A minimised master has no rectangle worth copying. Windows parks a minimised window off-screen
    // near -32000, so copying it would put the joining window - which is the one the user is looking
    // at, since a document just appeared in it - somewhere they cannot see or reach, and Present()
    // would then move the taskbar button onto it. StackOnFramePosChanging already refuses this exact
    // rectangle for the same reason; this is the other half of that guard.
    //
    // Left where Word put it instead. That is not a complete answer to "a document arrived while the
    // stack was down" - it is the answer to "never make a window unreachable", which is the rule that
    // may not be broken while the better answer is worked out.
    if (IsIconic(master))
    {
        g_inSync = FALSE;
        LogWrite(L"stack  hwnd=0x%p  joined while the stack is minimised - left where it is, "
                 L"not snapped to an off-screen rectangle", (void*)frame);
        return;
    }

    RECT rect;
    if (GetWindowRect(master, &rect))
    {
        SetWindowPos(frame, NULL, rect.left, rect.top,
                     rect.right - rect.left, rect.bottom - rect.top,
                     SWP_NOZORDER | SWP_NOACTIVATE);
    }

    g_inSync = FALSE;

    // Word will not lay this window's interior out - it is not focused - so hand it the rect the
    // focused window got. This is the layout-oracle rule; without it the document frame inside a
    // background window keeps the size it had before it joined the stack.
    RECT natural;
    if (StripGetNatural(master, &natural))
        StripSetNatural(frame, &natural, L"joined the stack, taking the master's interior");
}

// ---------------------------------------------------------------------------------------------
// Presentation: how many windows the rest of Windows is allowed to see.
//
// Stacking makes N windows look like one *inside* Word's frame. Everything outside it - the
// taskbar, Alt+Tab - still counts N, which gives the whole thing away at a glance. So the stack
// presents exactly one window: the active one.
//
// Two different mechanisms, because they are two different systems:
//   - the taskbar is told, through ITaskbarList (see taskbar.cpp). Styles are not involved.
//   - Alt+Tab reads WS_EX_TOOLWINDOW off the window when it is invoked, so setting that style on
//     the windows behind is enough. Deliberately *without* SWP_FRAMECHANGED: we want the shell's
//     classification to change, not the window's frame to be recalculated. Word draws its own
//     caption over the non-client area anyway, and the windows this is applied to are underneath
//     the active one where nothing about them is visible.
//
// The safety rule for both: a window that leaves the stack gets everything back. A window with no
// taskbar button and no Alt+Tab entry that is also underneath another window is unreachable by any
// means the user has.
// ---------------------------------------------------------------------------------------------

static void PresentWindow(HWND frame, BOOL show)
{
    TaskbarShow(frame, show);

    if (!g_altTab || !IsWindow(frame))
        return;

    LONG_PTR style = GetWindowLongPtrW(frame, GWL_EXSTYLE);
    LONG_PTR wanted = show ? (style & ~(LONG_PTR)WS_EX_TOOLWINDOW)
                           : (style |  (LONG_PTR)WS_EX_TOOLWINDOW);
    if (wanted != style)
        SetWindowLongPtrW(frame, GWL_EXSTYLE, wanted);
}

static void Present(void)
{
    if (!g_enabled)
        return;

    int joined = JoinedCount();

    for (int i = 0; i < g_memberCount; i++)
    {
        if (!g_members[i].joined)
            continue;

        // One window in the stack is just Word: it keeps its button and its Alt+Tab entry. Hiding
        // the only window there is would leave the user with no way back to Word at all.
        BOOL show = (joined <= 1) || (g_members[i].frame == g_active);
        PresentWindow(g_members[i].frame, show);
    }
}

// ---------------------------------------------------------------------------------------------
// Joining and leaving.
// ---------------------------------------------------------------------------------------------

// Take a member out of the middle of the array and put it at the end, keeping everything else in
// order. The array *is* the tab order, so this is how a tab is made the last one - and it is the same
// operation StackMoveTab performs, done to a member that is not in the row yet.
//
// Returns where the member now lives: the caller's pointer is into the array and the memmove moves it.
static Member* MoveToEnd(Member* member)
{
    int index = (int)(member - g_members);
    if (index < 0 || index >= g_memberCount || index == g_memberCount - 1)
        return member;

    Member moving = *member;
    memmove(&g_members[index], &g_members[index + 1],
            (size_t)(g_memberCount - index - 1) * sizeof(Member));
    g_members[g_memberCount - 1] = moving;
    return &g_members[g_memberCount - 1];
}

// Returns TRUE if the member array was compacted, which the caller has to know about because it is
// iterating over it. There is one caller - the janitor - and that is checked by the compiler rather
// than by hoping: this is a static with a single call site.
// ---------------------------------------------------------------------------------------------
// The row's own size, kept across restarts.
//
// Without this the first window in defines the stack rect - see Join - and the first window in
// is whatever size WORD restored it to. So the row is the size of Word's memory rather than the
// size the user chose, and on a wide screen that is a visible complaint rather than a detail:
// Word restores a near-full-width frame, Word's own layout puts four pages side by side inside
// it, and every tab that joins inherits it through MatchTo. Resizing the row fixes it until the
// next restart, when Word's rectangle comes back and takes the whole row with it.
//
// This does not try to work out WHY Word restores the rectangle it does - that could not be
// reproduced here, and it does not need to be. The remembered size is applied over the top of
// whatever Word chose, so the answer does not depend on knowing.
//
// Four DWORDs rather than one REG_BINARY blob because every other value under this key is a
// readable DWORD that settings.ps1 and regedit can both show. Negative coordinates survive the
// round trip as two's complement, which is not a corner case: a window on a monitor left of the
// primary one has a negative left, and the rig this was reported from has exactly that.
// ---------------------------------------------------------------------------------------------

static const wchar_t* const ROW_LEFT   = L"RowLeft";
static const wchar_t* const ROW_TOP    = L"RowTop";
static const wchar_t* const ROW_RIGHT  = L"RowRight";
static const wchar_t* const ROW_BOTTOM = L"RowBottom";
static const wchar_t* const ROW_MAX    = L"RowMaximized";

// Absent has to be told apart from any rectangle a user could actually have, and every DWORD
// here is a legitimate coordinate, so the reader's default doubles as the sentinel.
static const DWORD ROW_ABSENT = 0x7FFFFFFFu;

// Small enough that it cannot be a window someone sized on purpose. A row this size would be a
// row the user cannot use, so it is not worth storing and not worth restoring.
#define ROW_MIN_EDGE 200

static void RememberRowRect(void)
{
    if (!g_rowSize || !g_active || !IsWindow(g_active))
        return;

    WINDOWPLACEMENT placement;
    placement.length = sizeof(placement);
    if (!GetWindowPlacement(g_active, &placement))
        return;

    // rcNormalPosition rather than GetWindowRect, so that maximizing the row does not overwrite
    // the size chosen for it: the restored rectangle is kept underneath, and the maximized state
    // is stored beside it as a state. Minimised is refused outright - Windows parks a minimised
    // window near -32000 and that is not a size anyone asked for, the same reason MatchTo
    // refuses to copy a minimised master's rectangle.
    if (placement.showCmd == SW_SHOWMINIMIZED)
        return;

    const RECT r = placement.rcNormalPosition;
    if (r.right - r.left < ROW_MIN_EDGE || r.bottom - r.top < ROW_MIN_EDGE)
        return;

    WordTabWriteNumber(ROW_LEFT,   (DWORD)r.left);
    WordTabWriteNumber(ROW_TOP,    (DWORD)r.top);
    WordTabWriteNumber(ROW_RIGHT,  (DWORD)r.right);
    WordTabWriteNumber(ROW_BOTTOM, (DWORD)r.bottom);
    WordTabWriteNumber(ROW_MAX,    placement.showCmd == SW_SHOWMAXIMIZED ? 1u : 0u);

    LogWrite(L"stack  the row remembers its size: (%ld,%ld %ldx%ld)%s",
             r.left, r.top, r.right - r.left, r.bottom - r.top,
             placement.showCmd == SW_SHOWMAXIMIZED ? L" maximized" : L"");
}

// ---------------------------------------------------------------------------------------------
// The row's width when nothing has been remembered yet.
//
// The field report this exists for, three times over: "opening up 4 pages wide", then "still opened
// 3 pages wide". Neither number is anything WordTab draws. Word draws a page at its true size at
// 100% zoom, so a document frame N page-widths across shows N pages side by side - and on that rig
// Word restores its own frame 3840px wide on a 5120x2160 screen at 150%. A page there is
// 8.5in x 144dpi = 1224px, so 3840px is three pages across and the maximized 5120px is four. The
// two reports are one complaint seen at two window sizes, and the arithmetic says so.
//
// Remembering the size the user settles on - which is what shipped first - does not answer it. Until
// they settle on one there is nothing to remember, and Word's rectangle stands: the fix was made to
// depend on the user doing something, and the report that came back says they did not do it. A
// default that needs no manual step is the fix. The remembered size still wins the moment there is
// one, which is the point - this is what happens in its absence, not instead of it.
//
// Deliberately NOT written to the registry. "Remembered" means a width the user chose, and a default
// that wrote itself into that slot would be indistinguishable from one - it would then be applied
// forever on a machine where it was wrong, and the user would have no way to tell which of the two
// they were looking at.
// ---------------------------------------------------------------------------------------------

// A page in tenths of an inch. Letter (8.5in) rather than A4 (8.27in) on purpose: A4 is NARROWER,
// so a threshold computed from Letter is reached later and a width computed from Letter is wider.
// Both errors fall the same way - towards leaving the window alone.
#define ROW_PAGE_TENTHS 85

// Three pages across before this touches anything, and one-and-a-bit after it does.
//
// Three rather than two is the whole blast radius of this change, so it is the number to argue with
// first. A maximized Word on an ordinary 1920x1080 screen at 100% is about 2.3 pages across, and two
// pages side by side is what Word has always done there and what nobody has ever complained about.
// Starting at three leaves every ordinary screen untouched and catches the wide ones - which is
// where every one of these reports has come from.
#define ROW_CAP_TRIGGER_TENTHS 30
#define ROW_CAP_TARGET_TENTHS  14

// How wide one page is on this window's screen, in the same physical pixels GetWindowRect speaks.
static LONG PageWidthPx(HWND frame)
{
    int dpi = StripDpiOf(frame);
    if (dpi < 72 || dpi > 480)
        dpi = 96;
    return ((LONG)ROW_PAGE_TENTHS * (LONG)dpi) / 10;
}

static void ApplyDefaultRowRect(HWND frame)
{
    const LONG page = PageWidthPx(frame);
    if (page <= 0)
        return;

    // Measured as the window is SHOWN, not as it is stored. A maximized window covers the screen
    // whatever its rcNormalPosition says, and it is what is on the screen that decides how many
    // pages Word lays out - the "4 pages wide" report is the maximized case of the "3 pages" one.
    RECT shown;
    if (!GetWindowRect(frame, &shown))
        return;

    const LONG width = shown.right - shown.left;
    if (width < (page * ROW_CAP_TRIGGER_TENTHS) / 10)
        return;                 // under three pages across: Word's rectangle is not the complaint

    WINDOWPLACEMENT placement;
    placement.length = sizeof(placement);
    if (!GetWindowPlacement(frame, &placement))
        return;
    if (placement.showCmd == SW_SHOWMINIMIZED)
        return;                 // no rectangle worth correcting, same rule as MatchTo and Remember

    MONITORINFO info;
    info.cbSize = sizeof(info);
    HMONITOR monitor = MonitorFromWindow(frame, MONITOR_DEFAULTTONEAREST);
    if (!monitor || !GetMonitorInfoW(monitor, &info))
        return;

    // Where it goes when it stops being maximized is the user's own restored rectangle if they have
    // one and the work area if they do not - and either way only its WIDTH is overruled. Height,
    // and the corner it sits in, are left as they were found.
    RECT r = placement.rcNormalPosition;
    if (r.right - r.left < ROW_MIN_EDGE || r.bottom - r.top < ROW_MIN_EDGE)
        r = info.rcWork;

    const LONG target = (page * ROW_CAP_TARGET_TENTHS) / 10;
    if (target < ROW_MIN_EDGE)
        return;
    r.right = r.left + target;

    // Back onto the monitor it was already on. rcNormalPosition is in workspace coordinates and
    // rcWork is in screen coordinates; the two differ by the taskbar's edge, which is close enough
    // for a clamp whose whole job is to keep the window reachable.
    if (r.right > info.rcWork.right)
    {
        const LONG shift = r.right - info.rcWork.right;
        r.left  -= shift;
        r.right -= shift;
    }
    if (r.left < info.rcWork.left)
    {
        const LONG shift = info.rcWork.left - r.left;
        r.left  += shift;
        r.right += shift;
    }

    placement.rcNormalPosition = r;
    placement.showCmd = SW_SHOWNORMAL;
    placement.flags = 0;

    // Under g_inSync for the same reason every other move here is: this is WordTab moving the
    // window, so neither the position hook nor the WM_SIZE that records the row's size may read it
    // back as the user having chosen this width.
    g_inSync = TRUE;
    SetWindowPlacement(frame, &placement);
    g_inSync = FALSE;

    LogWrite(L"stack  hwnd=0x%p  no row size remembered and Word opened this window %ldpx wide, "
             L"which is %ld.%ld pages across at %ldpx to a page - narrowed to (%ld,%ld %ldx%ld). "
             L"Size the window yourself and that is what comes back instead (RowSize=0 turns this "
             L"off entirely).",
             (void*)frame, width, width / page, ((width * 10) / page) % 10, page,
             r.left, r.top, r.right - r.left, r.bottom - r.top);
}

static void ApplyRememberedRowRect(HWND frame)
{
    if (!g_rowSize || !IsWindow(frame))
        return;

    const DWORD stored = WordTabReadNumber(ROW_RIGHT, ROW_ABSENT);
    if (stored == ROW_ABSENT)
    {
        // Nothing remembered. Word's rectangle used to stand here unconditionally - and on a wide
        // screen Word's rectangle is the "3 pages wide" report. It stands only while it is a size
        // somebody would recognise as a Word window.
        ApplyDefaultRowRect(frame);
        return;
    }

    RECT r;
    r.left   = (LONG)WordTabReadNumber(ROW_LEFT,   0);
    r.top    = (LONG)WordTabReadNumber(ROW_TOP,    0);
    r.right  = (LONG)stored;
    r.bottom = (LONG)WordTabReadNumber(ROW_BOTTOM, 0);

    if (r.right - r.left < ROW_MIN_EDGE || r.bottom - r.top < ROW_MIN_EDGE)
        return;

    // A remembered rectangle is only good while the screen it was on still exists. Monitors get
    // unplugged and laptops get undocked, and a row restored onto a monitor that is gone is a row
    // the user cannot reach - which is the one rule MatchTo will not break either.
    if (!MonitorFromRect(&r, MONITOR_DEFAULTTONULL))
    {
        LogWrite(L"stack  hwnd=0x%p  the remembered row (%ld,%ld %ldx%ld) is on no monitor that "
                 L"exists now - leaving it where Word put it", (void*)frame,
                 r.left, r.top, r.right - r.left, r.bottom - r.top);
        return;
    }

    WINDOWPLACEMENT placement;
    placement.length = sizeof(placement);
    if (!GetWindowPlacement(frame, &placement))
        return;

    placement.rcNormalPosition = r;
    placement.showCmd = WordTabReadNumber(ROW_MAX, 0) ? SW_SHOWMAXIMIZED : SW_SHOWNORMAL;
    placement.flags = 0;

    // Under g_inSync for the same reason MatchTo is: this is WordTab moving the window, not the
    // user, and the position hook must not read it back as a drag to follow.
    g_inSync = TRUE;
    SetWindowPlacement(frame, &placement);
    g_inSync = FALSE;

    LogWrite(L"stack  hwnd=0x%p  the row takes the size it was left at, not the one Word "
             L"restored: (%ld,%ld %ldx%ld)%s", (void*)frame,
             r.left, r.top, r.right - r.left, r.bottom - r.top,
             placement.showCmd == SW_SHOWMAXIMIZED ? L" maximized" : L"");
}

void StackRememberRowSize(HWND frame)
{
    Member* member = Find(frame);
    if (!member || !member->joined || frame != g_active)
        return;
    RememberRowRect();
}

// Is this frame inside the grace window - the stack has noticed it has no document and is waiting
// to see whether that is a moment or a fact?
//
// Exported for the strip, which has the same question to answer about the same half second and no
// way of its own to answer it. A frame that loses its document reverts to a bare "Word" title
// immediately, and the strip's title poll runs ABOVE StackJanitor in the same tick - so without
// this the row renames a tab to "Word" while the stack is still deciding whether the tab should be
// there at all. Two things saying different things about the same window in the same tick is the
// shape that put a tab called "Word" in the row for a second and a half in the field.
BOOL StackIsWaitingFor(HWND frame)
{
    Member* member = Find(frame);
    return (member && member->joined && member->missedAt != 0) ? TRUE : FALSE;
}

static BOOL Join(Member* member)
{
    if (member->joined)
        return FALSE;

    // The one place a torn-off window is kept out, deliberately here rather than in the janitor's
    // eligibility test. Eligibility answers "could this window be a tab", which is still yes - it is
    // visible, sized and holds a document, which is exactly why the janitor would put it straight
    // back half a second after the user pulled it out. This answers the different question of whether
    // it may be, and putting it on the only path into the stack means a future caller cannot miss it.
    if (member->tornOff)
        return FALSE;

    // A frame that lost its document and has been handed a new one is carrying a document the user
    // has never seen, so its tab goes where every new document's tab goes: the end. Without this the
    // new document inherits the position the old one held, and Word chooses which frame to recycle,
    // so from the user's side a new document arrives in the middle of the row for no visible reason.
    BOOL compacted = FALSE;
    if (member->rejoin)
    {
        member->rejoin = FALSE;

        // Read before the move, not after. MoveToEnd memmoves the array, so `member` afterwards
        // points at whatever slid down into that slot - a different window entirely - and the line
        // this used to write named it. A log line that names the wrong window is worse than no line
        // at all; the first run of check-row.ps1 caught this one naming an empty frame.
        HWND moving = member->frame;

        Member* end = MoveToEnd(member);
        if (end != member)
        {
            compacted = TRUE;
            LogWrite(L"stack  hwnd=0x%p  a document that has never had a tab - it goes to the end "
                     L"of the row", (void*)moving);
            member = end;
        }
    }

    member->joined = TRUE;
    member->missedAt = 0;
    GetWindowRect(member->frame, &member->joinRect);
    member->joinZoomed = IsZoomed(member->frame) ? TRUE : FALSE;

    HWND master = (g_active && g_active != member->frame && Find(g_active) && Find(g_active)->joined)
                ? g_active : FirstJoined(member->frame);

    if (!master)
    {
        // First window in: it defines where the stack is - which is why the size it happens to
        // have is the size of everything that joins it, and why the remembered one is applied
        // here and nowhere else.
        g_active = member->frame;
        LogWrite(L"stack  hwnd=0x%p  joined as the first window (it defines the stack rect)",
                 (void*)member->frame);
        ApplyRememberedRowRect(member->frame);
    }
    else
    {
        MatchTo(master, member->frame);
        LogWrite(L"stack  hwnd=0x%p  joined, snapped to 0x%p  (%d in the stack)",
                 (void*)member->frame, (void*)master, JoinedCount());
    }

    // The view, once, now that this window is a tab.
    //
    // Here rather than on a timer, and this is the distinction that matters: a document arriving
    // in the row is an EVENT, and asking Word one question per event is not the same animal as the
    // dot poll, which asks per window twice a second and had to be governed for it. Reading and
    // writing the two zoom properties measured 3ms here. If it is slower on a SharePoint machine it
    // is slower once, as the document opens, which is already the slowest moment there is.
    if (g_onePage)
    {
        LONG columns = 0;
        LONG percent = 0;
        if (WordTabOnePageView(member->frame, &columns, &percent))
        {
            // Two illnesses, one cure, and the log says which one it found. They were one line for
            // an hour and that line would have reported "1 pages across", which is not a thing that
            // happens and would have sent the next person looking in the wrong place.
            if (columns > 1)
                LogWrite(L"view  hwnd=0x%p  this document opened %ld pages across at %ld%% - put "
                         L"back to one page at 100%% (OnePage=0 leaves it alone)",
                         (void*)member->frame, columns, percent);
            else
                LogWrite(L"view  hwnd=0x%p  this document opened at %ld%%, which is what fitting "
                         L"several pages across leaves behind - put back to 100%% (OnePage=0 leaves "
                         L"it alone)", (void*)member->frame, percent);
        }
    }

    Present();
    StripRefreshTabs();
    return compacted;
}

static void Leave(Member* member, const wchar_t* why, BOOL restorePosition)
{
    if (!member->joined)
        return;
    member->joined = FALSE;

    // Everything back, first and unconditionally. A window with no taskbar button, no Alt+Tab entry
    // and another window on top of it cannot be reached by any means the user has.
    PresentWindow(member->frame, TRUE);

    // Put it back where it was before we stacked it. Without this a window that leaves the stack -
    // switched off, or hidden and shown again - is left sitting exactly under the others, which
    // from the user's side looks like a document that vanished.
    if (restorePosition && IsWindow(member->frame))
    {
        g_inSync = TRUE;
        if (member->joinZoomed)
        {
            if (!IsZoomed(member->frame))
                ShowWindow(member->frame, SW_MAXIMIZE);
        }
        else
        {
            if (IsZoomed(member->frame))
                ShowWindow(member->frame, SW_RESTORE);
            SetWindowPos(member->frame, NULL,
                         member->joinRect.left, member->joinRect.top,
                         member->joinRect.right - member->joinRect.left,
                         member->joinRect.bottom - member->joinRect.top,
                         SWP_NOZORDER | SWP_NOACTIVATE);
        }
        g_inSync = FALSE;

        // Word will not re-lay out a window it is not focused on, so its own document frame has to
        // be refitted to the size we just gave it. Restore has the oracle problem in reverse.
        StripRefit(member->frame);
    }

    if (g_active == member->frame)
    {
        // Where the user was before a close took them somewhere else, if that is still a live tab.
        // Deliberately checked here rather than keyed to the frame we posted WM_CLOSE to: frame
        // lifetime is not document lifetime, and closing a document was measured to hide one window
        // and destroy a different one. What matters is that the active slot is being vacated, not
        // which window vacated it.
        HWND next = NULL;
        if (g_returnTo && g_returnTo != member->frame && IsWindow(g_returnTo))
        {
            Member* back = Find(g_returnTo);
            if (back && back->joined)
                next = g_returnTo;
        }

        BOOL returning = (next != NULL);
        if (!next)
            next = FirstJoined(member->frame);

        g_active = next;
        if (g_active)
        {
            SetWindowPos(g_active, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
            if (returning)
            {
                // Raising is not enough here: the window the user was on has to take the keyboard
                // back too, or they are looking at one document and typing into the z-order.
                SetForegroundWindow(g_active);
                LogWrite(L"stack  hwnd=0x%p  active again - back to the tab the user was on",
                         (void*)g_active);
            }
        }
        g_returnTo = NULL;
    }

    LogWrite(L"stack  hwnd=0x%p  left the stack (%s)  (%d remain)",
             (void*)member->frame, why, JoinedCount());

    Present();
    StripRefreshTabs();
}

// ---------------------------------------------------------------------------------------------
// Entry points from frames.cpp.
// ---------------------------------------------------------------------------------------------

void StackStart(void)
{
    if (g_started)
        return;
    g_started = TRUE;

    g_enabled = WordTabReadFlag(L"Stack", TRUE);
    g_altTab  = WordTabReadFlag(L"AltTab", TRUE);

    // Off, a tab cannot be taken out of the stack at all: the menu item is not offered and the
    // command is refused. Unlike most of these switches this one is not here to put a defect back -
    // there is no previous behaviour to restore, because nothing could leave the stack before - it is
    // here because a torn-off window is the one thing WordTab does that the user cannot undo from
    // inside WordTab, and a machine where it misbehaves needs a way to stop offering it.
    g_tearOff  = WordTabReadFlag(L"TabTearOff", TRUE);

    // Off, the row is whatever size Word restored the first window to - the behaviour before
    // the row remembered anything. Here because overriding the host's own window placement is
    // the kind of thing that wants an off switch on a machine where it goes wrong.
    g_rowSize  = WordTabReadFlag(L"RowSize", TRUE);

    // Off, a document that opens showing several pages side by side is left showing them. On is
    // the default because it is the only setting the person this was built for was correcting by
    // hand on every single Word start - see WordTabOnePageView for what Word is doing and why the
    // ribbon buttons never made it stop.
    g_onePage  = WordTabReadFlag(L"OnePage", TRUE);

    // What the title bar's x does, and the one switch here that is not a boolean:
    //
    //   1 (default)  ask - "close all N tabs", "close only this document", or cancel
    //   2            close the whole stack without asking, which is what this did before it asked
    //   0            close only the document in front, which is what Word does without WordTab
    //
    // Three states because the two ends are both real answers somebody might want standing, and the
    // question is only worth asking of a person who has not already answered it. Read as a number
    // rather than a flag for that reason; `settings.ps1` still shows it as on or off, because that
    // script speaks in on and off and 2 is documented in the README beside TabFontSize.
    g_closeStack = (int)WordTabReadNumber(L"TabCloseStack", 1);

    TaskbarStart();

    LogWrite(L"StackStart  stacking=%s  altTab suppression=%s  detach=%s  the window's x=%s"
             L"  row size=%s  one page=%s",
             g_enabled ? L"on" : L"off (HKCU\\Software\\WordTab\\Stack=0)",
             g_altTab ? L"on" : L"off",
             g_tearOff ? L"on" : L"off (HKCU\\Software\\WordTab\\TabTearOff=0)",
             g_closeStack == 2 ? L"the whole stack, no question (TabCloseStack=2)"
           : g_closeStack == 1 ? L"asks: all, this one, or cancel"
                               : L"this document only (TabCloseStack=0)",
             g_rowSize ? L"remembered" : L"off (RowSize=0) - Word's rectangle stands",
             g_onePage ? L"on" : L"off (OnePage=0) - Word's remembered column count stands");
}

void StackAttachFrame(HWND frame)
{
    if (!g_enabled || !frame || Find(frame) || g_memberCount >= MAX_MEMBERS)
        return;

    Member* member = &g_members[g_memberCount++];
    memset(member, 0, sizeof(*member));
    member->frame = frame;

    // **A window we have only just met has never had a tab, so its first one goes at the end of the
    // row - wherever in the array this member ends up sitting.**
    //
    // Appending to the array already puts it there, and for a frame Word creates while the user
    // watches that is the whole story. It is not the whole story for a frame Word made EARLIER and
    // is only now filling: the user's report is "now newly opened docs popping infront lets make
    // them pop to end", and their log shows empty frames attaching with no title at all - a frame
    // that attached before the documents around it holds an early slot in the array and hands it to
    // whatever document lands in it. This flag is what MoveToEnd keys off, and setting it here says
    // the same thing for a frame we have just met that it already says for one Word has recycled:
    // this document has never been a tab, so it becomes the last one.
    //
    // Order-preserving when several frames arrive together, which is every cold start with more than
    // one document: they join in array order and each moves to the end in turn, which leaves them in
    // the order they were in. Driven, not reasoned - see tools\check-reorder.ps1.
    member->rejoin = TRUE;

    // Not joined here. At the moment a frame is subclassed it is usually not visible yet and has no
    // `_WwF` - measured, every cold start - so eligibility is decided by the janitor instead.
    StackJanitor();
}

void StackDetachFrame(HWND frame)
{
    Member* member = Find(frame);
    if (!member)
        return;

    // The window is going away, so there is nothing to put back and its children may already be
    // gone. Restoring here would be at best pointless.
    Leave(member, L"frame closing", FALSE);

    int index = (int)(member - g_members);
    for (int i = index; i < g_memberCount - 1; i++)
        g_members[i] = g_members[i + 1];
    g_memberCount--;
}

void StackOnFrameActivate(HWND frame)
{
    if (!g_enabled)
        return;

    Member* member = Find(frame);
    if (!member || !member->joined || g_active == frame)
        return;

    g_active = frame;
    LogWrite(L"stack  active -> 0x%p", (void*)frame);

    // The taskbar button and the Alt+Tab entry move with the active tab. If they stayed on one
    // window, restoring the stack from the taskbar would bring back a document the user was not
    // looking at.
    Present();

    // The newly focused window is now the layout oracle, and it is the only one Word will lay out.
    // Push what it has to the others so a stale background window is corrected on the way in.
    RECT natural;
    if (StripGetNatural(frame, &natural))
    {
        for (int i = 0; i < g_memberCount; i++)
            if (g_members[i].joined && g_members[i].frame != frame)
                StripSetNatural(g_members[i].frame, &natural,
                                L"the active window changed and pushed its interior out");
    }

    StripRefreshTabs();
}

// The lockstep. Called from the active frame's WM_WINDOWPOSCHANGING - a proposal, not a fact - so
// the other windows are moved to where this one is *about* to be, before it gets there. That is
// the whole reason for being inside Word: from outside, the followers can only ever be one poll
// behind, which is what the user saw as glitching.
int StackOnFramePosChanging(HWND frame, const WINDOWPOS* pos)
{
    if (!g_enabled || g_inSync || !pos || frame != g_active)
        return 0;
    if (pos->flags & (SWP_HIDEWINDOW | SWP_SHOWWINDOW))
        return 0;
    if (IsIconic(frame))
        return 0;

    // A maximized window is not a rectangle to be copied. Its rect covers the strip of screen the
    // taskbar sits on, which a window merely *resized* to those numbers does not - measured, the
    // two differ by 72px on this rig. Maximizing is a state, and it is propagated as one from
    // WM_SIZE below.
    if (IsZoomed(frame))
        return 0;

    RECT current;
    if (!GetWindowRect(frame, &current))
        return 0;

    LONG x  = (pos->flags & SWP_NOMOVE) ? current.left : pos->x;
    LONG y  = (pos->flags & SWP_NOMOVE) ? current.top  : pos->y;
    LONG cx = (pos->flags & SWP_NOSIZE) ? (current.right - current.left)  : pos->cx;
    LONG cy = (pos->flags & SWP_NOSIZE) ? (current.bottom - current.top)  : pos->cy;

    // A minimise parks the window off-screen near -32000 (scaled by DPI). Following it there would
    // take every other document with it.
    if (x < -30000 || y < -30000 || cx <= 0 || cy <= 0)
        return 0;

    HWND followers[MAX_MEMBERS];
    int count = 0;
    for (int i = 0; i < g_memberCount; i++)
    {
        HWND candidate = g_members[i].frame;
        if (!g_members[i].joined || candidate == frame)
            continue;
        if (!IsWindow(candidate) || !IsWindowVisible(candidate) || IsIconic(candidate))
            continue;
        followers[count++] = candidate;
    }
    if (count == 0)
        return 0;

    g_inSync = TRUE;

    // One atomic pass rather than one repaint each.
    HDWP batch = BeginDeferWindowPos(count);
    for (int i = 0; i < count; i++)
    {
        UINT flags = SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOREDRAW;
        if (batch)
            batch = DeferWindowPos(batch, followers[i], NULL, x, y, cx, cy, flags);
        else
            SetWindowPos(followers[i], NULL, x, y, cx, cy, flags);
    }
    if (batch)
        EndDeferWindowPos(batch);

    g_inSync = FALSE;
    return count;
}

BOOL StackIsSyncing(void)
{
    return g_inSync;
}

// Called when the *active* window's document frame has been laid out by Word. Everything the other
// windows know about their own interior comes from here: Word never lays them out at all.
void StackOnActiveLayout(HWND frame, const RECT* natural)
{
    if (!g_enabled || g_inSync || frame != g_active || !natural)
        return;

    // Only a window Word is actually showing can speak for the stack. A frame on its way out gets
    // its layout dismantled first, and one on its way to the taskbar gets squeezed to nothing;
    // broadcasting either would take every other document's layout down with it - measured, once,
    // and it put every strip on top of its ribbon.
    if (!IsWindowVisible(frame) || IsIconic(frame))
        return;

    for (int i = 0; i < g_memberCount; i++)
    {
        if (!g_members[i].joined || g_members[i].frame == frame)
            continue;
        StripSetNatural(g_members[i].frame, natural, L"Word laid the active window out");
    }
}

// Minimising. The stack is one window to the user, so it goes down and comes back as one.
//
// This has to be all-or-nothing in both directions, and the reason is the taskbar. Only the active
// window has a button, so only the active window can be restored by clicking it. If minimising
// took just that window down, the next document in the stack would be revealed underneath - the
// illusion collapses, and the "one window" the user minimised is still on screen. And if restoring
// brought back only the active window, every other document would be left minimised with no
// taskbar button and no Alt+Tab entry: unreachable.
void StackOnFrameSize(HWND frame, WPARAM sizeType)
{
    if (!g_enabled || g_inSync)
        return;

    Member* member = Find(frame);
    if (!member || !member->joined || frame != g_active)
        return;

    if (sizeType == SIZE_MINIMIZED && !g_minimized)
    {
        g_minimized = TRUE;
        g_inSync = TRUE;
        int count = 0;
        for (int i = 0; i < g_memberCount; i++)
        {
            if (!g_members[i].joined || g_members[i].frame == frame)
                continue;
            if (IsWindow(g_members[i].frame) && !IsIconic(g_members[i].frame))
            {
                // SW_SHOWMINNOACTIVE, not SW_MINIMIZE: the latter activates the next window in the
                // z-order on its way down, which here means handing focus to another document in
                // the same stack while the user is trying to put the whole thing away.
                ShowWindow(g_members[i].frame, SW_SHOWMINNOACTIVE);
                count++;
            }
        }
        g_inSync = FALSE;
        LogWrite(L"stack  minimised with the active window: %d other window(s) went down too", count);
        return;
    }

    // Maximize and restore are *states*, and they have to be propagated as states rather than as
    // rectangles - see MatchTo. This is the gap spike 2 could not close, and it is visible: a
    // window given a maximized window's rectangle is 72px shorter than a maximized one on this rig,
    // because a real maximized window covers the taskbar's strip of screen.
    if (!g_minimized && (sizeType == SIZE_MAXIMIZED || sizeType == SIZE_RESTORED))
    {
        BOOL zoom = (sizeType == SIZE_MAXIMIZED);
        g_inSync = TRUE;
        int count = 0;
        for (int i = 0; i < g_memberCount; i++)
        {
            HWND other = g_members[i].frame;
            if (!g_members[i].joined || other == frame || !IsWindow(other) || IsIconic(other))
                continue;
            if ((IsZoomed(other) ? TRUE : FALSE) != zoom)
            {
                SetZoomState(other, zoom);
                count++;
            }
        }
        g_inSync = FALSE;

        if (count > 0)
        {
            for (int i = 0; i < g_memberCount; i++)
                if (g_members[i].joined && g_members[i].frame != frame)
                    MatchTo(frame, g_members[i].frame);

            LogWrite(L"stack  %s with the active window: %d other window(s) followed",
                     zoom ? L"maximized" : L"restored", count);
        }
        return;
    }

    if ((sizeType == SIZE_RESTORED || sizeType == SIZE_MAXIMIZED) && g_minimized)
    {
        g_minimized = FALSE;
        g_inSync = TRUE;
        int count = 0;
        for (int i = 0; i < g_memberCount; i++)
        {
            if (!g_members[i].joined || g_members[i].frame == frame)
                continue;
            if (IsWindow(g_members[i].frame) && IsIconic(g_members[i].frame))
            {
                // Not SW_RESTORE: that would activate them, and the window the user clicked on
                // would end up behind the ones it brought back with it.
                ShowWindow(g_members[i].frame, SW_SHOWNOACTIVATE);
                count++;
            }
        }
        g_inSync = FALSE;

        // They come back wherever Windows left them, so put the stack back together.
        for (int i = 0; i < g_memberCount; i++)
        {
            if (g_members[i].joined && g_members[i].frame != frame)
                MatchTo(frame, g_members[i].frame);
        }

        SetWindowPos(frame, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
        LogWrite(L"stack  restored with the active window: %d other window(s) came back", count);
        Present();
    }
}

// Membership, re-decided from what is true right now rather than from what happened. Cheap enough
// to run on the strip's half-second janitor, and that cadence is what makes it robust against
// Word hiding and showing frames without telling anyone.
// Which of the row's windows the user is actually looking at.
//
// The stack holds every window at one rectangle, so the document on screen is the front-most of
// them and nothing else. That is the whole definition, and it is worth writing down because the
// row spent a long time reading a different answer: the KEYBOARD. GetForegroundWindow and
// WM_ACTIVATE are both about focus, and focus is only a proxy for what is in front.
//
// The proxy failed in the field. From the work rig, their words: "when opening a second doc it
// switch to that but initial tab still highlighted" - and their log has it three times in one day,
// `joined, snapped to 0x...` with no `active ->` after it. Word had put the new document's window
// in front without handing it the keyboard, so every question the row knew how to ask returned the
// window the user could no longer see, and the highlight stayed there until they clicked something.
// It does not reproduce on the dev rig, where Word raises and focuses in the same breath - which is
// exactly why the check for it drives the mechanism (SWP_NOACTIVATE) rather than the trigger.
//
// GetTopWindow walks the desktop from the front, so the first joined member met is the one on top.
// The walk stops there, and nothing is walked at all while the row holds a single tab - with nothing
// to switch between there is nothing to decide, and that is the idle case this runs in most of the
// time. What is left is one GetWindow per window in front of ours, twice a second: a local call with
// no cross-process work in it, which is what makes this safe to put on a timer at all. The dot poll
// is the standing warning about the other kind - see RESULT-dot.md, where one object-model call
// measured 430us here and half a second on the work rig.
static HWND FrontMostJoined(void)
{
    if (JoinedCount() < 2)
        return NULL;

    // Not while a tab is being carried. The gesture moves windows for its own reasons - showing
    // the card under the pointer was measured raising the window the tab was picked up from, over
    // the one the same press had just activated - so z-order during a drag is not a statement
    // about which document the user has chosen. Seen once in the reorder suite, 92ms after a
    // "tab clicked ->" line, undoing the switch. Answering NULL here leaves the older focus rule
    // to reply, which is exactly what it replied before any of this existed.
    if (StripTabPressHeld())
        return NULL;

    for (HWND h = GetTopWindow(NULL); h; h = GetWindow(h, GW_HWNDNEXT))
    {
        Member* member = Find(h);
        if (member && member->joined && IsWindowVisible(h) && !IsIconic(h))
            return h;
    }
    return NULL;
}

static void JanitorPass(void)
{

    for (int i = 0; i < g_memberCount; i++)
    {
        Member* member = &g_members[i];

        if (!member->joined)
        {
            // A torn-off window stops being torn off when it stops holding a document.
            //
            // "It stays out until the window closes" was the intended rule and it was wrong, because
            // the window does not close. Word hides the frame when its last document goes and reuses
            // it for the next one - the same thing that gave an empty Word window a tab called "Word"
            // for four slices - so the flag outlived the document it was a decision about. Measured:
            // tear a tab off, close it, press +, and Word puts the new document straight back into
            // that same HWND, which then refused to join the row for the rest of the process. Five
            // windows and four tabs, with no way for the user to work out why.
            //
            // Tested on the document rather than on visibility on purpose. Word hides and re-shows
            // frames of its own accord - closing one document was measured to hide a *different*
            // window for a moment - and a rule keyed to that would drop a torn-off window back into
            // the stack while the user was looking at it.
            BOOL empty = !IsWindow(member->frame) || !StripHasDocument(member->frame);

            if (member->tornOff && empty)
            {
                member->tornOff = FALSE;
                LogWrite(L"stack  hwnd=0x%p  no longer torn off - its document has gone and the "
                         L"frame is Word's to reuse", (void*)member->frame);
            }

            // Recorded from the same fact, and while it is still true. A frame sitting out of the row
            // with no document in it is one Word is free to hand the next document to, and that
            // document has never had a tab - so it gets a new one, at the end, rather than the place
            // this frame used to hold. Read again on every tick because it is a state, not an event:
            // the frame can be emptied and refilled between two of them.
            //
            // Not set for a member that still holds a document. That is a window Word has merely
            // hidden or the user has minimised, and it must come back exactly where it was.
            if (empty)
                member->rejoin = TRUE;

            if (EligibleToJoin(member->frame))
            {
                // Join can move this member to the end of the array, which shifts everything after it
                // down one - so the entry now at i is the one that was at i+1 and has not been looked
                // at. Stepping back is what makes the pass complete rather than leaving a window's
                // membership half a second stale.
                if (Join(member))
                    i--;
            }
        }
        else if (EligibleToStay(member->frame))
        {
            member->missedAt = 0;
        }
        else
        {
            // Left where it is, deliberately, and for two different reasons now.
            //
            // Hidden: Word hides and re-shows frames of its own accord - closing one document was
            // measured to hide a *different* window for a moment - and moving a window back to its
            // old position every time it blinks makes documents jump around the screen for no
            // reason the user can see. Nobody is looking at it, and when it comes back the join
            // snaps it to the stack again.
            //
            // No document: the user *is* looking at this one, which makes moving it worse rather
            // than better - they closed a document, not a window. It keeps its place, its taskbar
            // button and its Alt+Tab entry, and its strip is left showing an empty row and a +.
            //
            // Only StackStop puts windows back where they came from.
            BOOL alive = IsWindow(member->frame) ? TRUE : FALSE;
            const wchar_t* why = L"gone";
            if (alive)
                why = IsWindowVisible(member->frame) ? L"no document open" : L"hidden or minimised";

            DWORD now = GetTickCount();
            BOOL  first = (member->missedAt == 0) ? TRUE : FALSE;
            if (first)
            {
                member->missedAt = now ? now : 1;   // 0 is "eligible", so never store it as a time
            }

            // A window that has actually gone is not waited for: there is nothing to come back, and
            // holding a tab for a destroyed HWND would draw a name read off a dead window. Nor is a
            // close we asked for ourselves - the tab the user clicked the x on must not sit there
            // for another second looking like it did not work, and we know that one is real because
            // we posted the WM_CLOSE.
            BOOL immediate = (!alive || CloseWasAskedFor(member->frame)) ? TRUE : FALSE;

            DWORD waited = now - member->missedAt;

            if (!immediate && waited < STAY_GRACE_MS)
            {
                // Once, on the first miss, with everything needed to tell the two mechanisms apart
                // next time this is reported from a machine I cannot reach: which document frames
                // the window has, which one the strip is bound to, and what is inside them. If the
                // window comes back on the next tick this line is the only trace it left, and that
                // is exactly the trace that was missing from the report that queued this.
                if (first)
                {
                    wchar_t frames[512];
                    StripDescribeDocumentFrames(member->frame, frames, 512);
                    LogWrite(L"stack  hwnd=0x%p  would have left the stack (%s) - waiting %d ms "
                             L"in case it is a moment rather than a fact.  %s",
                             (void*)member->frame, why, STAY_GRACE_MS, frames);
                }
                continue;
            }

            if (!immediate)
            {
                LogWrite(L"stack  hwnd=0x%p  still %s %lu ms later - it is a fact, not a moment",
                         (void*)member->frame, why, (unsigned long)waited);
            }

            member->missedAt = 0;
            Leave(member, why, FALSE);
        }
    }

    // Which tab is drawn selected, read back from the system rather than tracked through events
    // Word does not always send. The window in front is the one the user is looking at, so it is
    // the one the row is on - see FrontMostJoined for why that is the question and focus was not.
    HWND showing = FrontMostJoined();

    // When z-order declines to answer - one tab or none, or a gesture still in progress - the
    // older rule stands: Word activates windows without always telling our frame procedure, so
    // trust the system over our own bookkeeping. Both rules are the same sentence - believe what
    // is there over what we last wrote down - asked of the two things that can answer it.
    if (!showing)
    {
        HWND foreground = GetForegroundWindow();
        Member* member  = foreground ? Find(foreground) : NULL;
        if (member && member->joined)
            showing = foreground;
    }

    if (showing && showing != g_active)
        StackOnFrameActivate(showing);

    // After the membership pass above, deliberately: whether the tab we asked Word to close has
    // actually gone is a membership question, and this reads the answer that loop just wrote.
    CloseBatchStep();

    Reconcile();

    // Last, so it describes the row as this tick left it. Silent unless it changed.
    LogRow();
}

// **A membership pass never runs inside another membership pass.**
//
// The janitor is not only the timer. `WM_SHOWWINDOW` calls it, and Join and Leave both show and hide
// windows - so a pass can call, synchronously and several frames deep, straight back into itself.
// The array is the tab order and `MoveToEnd` memmoves it, so a re-entrant Join leaves the outer pass
// holding a `Member*` into a slot that now describes a different window.
//
// Measured rather than reasoned, and it cost a green suite: `check-reorder` went from 118/118 to
// 110/118 the moment the WM_SHOWWINDOW handler started re-deciding membership *after* Word had shown
// the window instead of before. That is the correct place to ask - the state before a show is the
// state the window is leaving - but it is also the first time the answer could be "yes, join it",
// which is the only branch that moves the array. The hazard was always there; the reordering is what
// reached it. An A/B against a rebuilt binary of the previous commit is what found it, which is the
// standing rule in this project for a reason.
//
// Skipping the nested pass loses nothing: the timer comes round in half a second, and the outer pass
// is in the middle of deciding the very thing the nested one would have decided.
static BOOL g_inJanitor = FALSE;

void StackJanitor(void)
{
    if (!g_enabled || g_inJanitor)
        return;

    g_inJanitor = TRUE;
    JanitorPass();
    g_inJanitor = FALSE;
}

// Put back together anything that has come apart.
//
// Every divergence found so far had its own cause and its own event - a maximize that propagated as
// a rectangle rather than a state, a window whose restore arrived while it was not the active one -
// and fixing each one individually is a losing game: the next cause is one Word update away. So
// rather than trust that every event was caught, the stack compares itself against the active
// window twice a second and repairs what does not match. The same principle as the strip comparing
// itself against where it actually is rather than where it last put itself.
//
// Cheap when nothing is wrong: two rectangle comparisons per window and no calls at all.
static void Reconcile(void)
{
    if (!g_enabled || g_minimized || g_inSync || !g_active || !IsWindow(g_active))
        return;
    if (IsIconic(g_active) || JoinedCount() < 2)
        return;

    RECT master;
    if (!GetWindowRect(g_active, &master))
        return;

    BOOL masterZoomed = IsZoomed(g_active) ? TRUE : FALSE;

    for (int i = 0; i < g_memberCount; i++)
    {
        HWND frame = g_members[i].frame;
        if (!g_members[i].joined || frame == g_active || !IsWindow(frame) || IsIconic(frame))
            continue;

        RECT rect;
        if (!GetWindowRect(frame, &rect))
            continue;

        BOOL zoomed = IsZoomed(frame) ? TRUE : FALSE;
        if (zoomed == masterZoomed &&
            rect.left == master.left && rect.top == master.top &&
            rect.right == master.right && rect.bottom == master.bottom)
        {
            g_members[i].repairs = 0;
            continue;
        }

        // A window that will not stay put is worse than one that is wrong: forcing it every half
        // second forever would be a permanent cost for no gain. Log it once and leave it.
        if (g_members[i].repairs >= 10)
            continue;

        if (g_members[i].repairs == 0)
        {
            LogWrite(L"stack  hwnd=0x%p  out of step: (%ld,%ld %ldx%ld) zoomed=%d, master "
                     L"(%ld,%ld %ldx%ld) zoomed=%d - putting it back",
                     (void*)frame,
                     rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top, (int)zoomed,
                     master.left, master.top, master.right - master.left, master.bottom - master.top,
                     (int)masterZoomed);
        }

        g_members[i].repairs++;
        if (g_members[i].repairs == 10)
        {
            LogWrite(L"stack  hwnd=0x%p  will not stay in step after 10 attempts - leaving it "
                     L"alone rather than fighting it every half second", (void*)frame);
        }

        MatchTo(g_active, frame);
    }
}

void StackStop(void)
{
    if (!g_started)
        return;
    g_started = FALSE;

    // Before anything is put back: a queue of tabs still to close is a queue of WM_CLOSEs about to
    // be posted to windows we are in the middle of letting go of.
    CloseBatchEnd(L"abandoned - the add-in is shutting down");

    for (int i = 0; i < g_memberCount; i++)
        Leave(&g_members[i], L"shutdown", TRUE);

    // Belt and braces over Leave, which has already done this for every joined member: nothing may
    // be left without a taskbar button or an Alt+Tab entry, including a window that had somehow
    // stopped being a member without going through Leave.
    for (int i = 0; i < g_memberCount; i++)
        PresentWindow(g_members[i].frame, TRUE);

    TaskbarStop();

    g_memberCount = 0;
    g_active = NULL;
    g_returnTo = NULL;
    g_minimized = FALSE;
    LogWrite(L"StackStop  done");
}

// ---------------------------------------------------------------------------------------------
// What the strip needs: the tab row, and what to do when one is clicked.
// ---------------------------------------------------------------------------------------------

int StackTabs(HWND frame, HWND* out, int max, int* activeIndex)
{
    if (activeIndex)
        *activeIndex = 0;

    // No document, no tab. Word leaves the frame alive and on screen after its last document is
    // closed, and that window is a real thing the user is looking at - but there is nothing for a
    // tab to name, select or close. The row is drawn empty and only the + is left, which is both an
    // honest statement that nothing is open and the one button that is still worth pressing:
    // measured, Word puts the new document into this very window rather than opening another.
    //
    // Ahead of the stacking test on purpose. This is a fact about the document, not about the
    // stack, so it holds with `Stack` switched off too.
    if (!StripHasDocument(frame))
        return 0;

    // Stacking off, or this window is not in a stack: it is its own single tab. The strip then
    // still shows the document's name, which is the previous slice's behaviour.
    Member* member = Find(frame);
    if (!g_enabled || !member || !member->joined)
    {
        if (max > 0 && out)
            out[0] = frame;
        return max > 0 ? 1 : 0;
    }

    int count = 0;
    for (int i = 0; i < g_memberCount && count < max; i++)
    {
        if (!g_members[i].joined)
            continue;
        if (out)
            out[count] = g_members[i].frame;
        if (g_members[i].frame == g_active && activeIndex)
            *activeIndex = count;
        count++;
    }
    return count;
}

// Where a tab sits in the row: joined members, counted left to right. -1 for a window that is not a
// tab in a stack - stacking switched off, a lone window, one Word has hidden - which is the answer
// the strip needs, because none of those can be reordered.
int StackTabIndex(HWND frame)
{
    if (!g_enabled || !frame)
        return -1;

    int index = 0;
    for (int i = 0; i < g_memberCount; i++)
    {
        if (!g_members[i].joined)
            continue;
        if (g_members[i].frame == frame)
            return index;
        index++;
    }
    return -1;
}

// How many tabs sit to the right of this one. 0 for the last tab, and 0 for a window that is not a
// tab in a stack - both of which mean the same thing to the caller: there is nothing for "Close Tabs
// to the Right" to close, so the menu item is greyed and the command is refused. Counted rather than
// answered as a boolean because the log line says how many, and a batch that says "3 tab(s) queued"
// against a menu item the user believed applied to two is how a miscount gets noticed.
int StackTabsRightOf(HWND frame)
{
    int index = StackTabIndex(frame);
    if (index < 0)
        return 0;
    return JoinedCount() - index - 1;
}

// The tab `delta` positions along the row, wrapping. Ctrl+Tab is +1 and Ctrl+Shift+Tab is -1.
//
// The reason this is a function on the stack rather than three lines inside the keyboard hook is the
// same reason StackMoveTab exists: the row is g_members in array order, and the translation from a
// tab position to the array slot that holds it is not the identity - g_members also carries windows
// that are not joined. Every caller reaches the row through an HWND and nothing outside this file
// indexes the array.
//
// NULL for "there is nowhere to go", which covers a window that is not a tab in a stack (stacking
// switched off, a lone window, one Word has hidden) and a stack holding a single tab. The caller
// swallows the chord either way - see the note on GetMsgProc about why a key that sometimes types a
// tab character is worse than one that sometimes does nothing.
HWND StackNeighbourTab(HWND frame, int delta)
{
    int from = StackTabIndex(frame);
    if (from < 0)
        return NULL;

    int joined = JoinedCount();
    if (joined < 2)
        return NULL;

    // Two modulos: C's % keeps the sign of the dividend, so a step off the left end of a three-tab
    // row is -1 and would index nothing. Written for any delta rather than for the two it is called
    // with, because a keyboard-scroll slice would otherwise find a function that only handles +-1.
    int to = ((from + delta) % joined + joined) % joined;
    if (to == from)
        return NULL;

    int seen = 0;
    for (int i = 0; i < g_memberCount; i++)
    {
        if (!g_members[i].joined)
            continue;
        if (seen == to)
            return g_members[i].frame;
        seen++;
    }
    return NULL;
}

// Move a tab to a position in the row.
//
// The move is expressed in tab positions and performed in array positions, and the two are not the
// same: g_members also holds windows that are not joined - one Word has hidden, one on its way in or
// out - and those have no place in the row. So the tab index is translated to the array slot the tab
// at that position actually occupies, and the entry is moved there.
//
// Returns TRUE only when the row really changed. A drag calls this on every mouse movement, so "it
// is already there" has to be free and has to be silent - otherwise the log fills with a line per
// pixel of a gesture that did nothing.
BOOL StackMoveTab(HWND frame, int toIndex)
{
    int from = StackTabIndex(frame);
    if (from < 0)
        return FALSE;

    int joined = JoinedCount();
    if (joined < 2)
        return FALSE;

    if (toIndex < 0)
        toIndex = 0;
    if (toIndex > joined - 1)
        toIndex = joined - 1;
    if (toIndex == from)
        return FALSE;

    int fromSlot = -1, toSlot = -1, seen = 0;
    for (int i = 0; i < g_memberCount; i++)
    {
        if (!g_members[i].joined)
            continue;
        if (seen == from)
            fromSlot = i;
        if (seen == toIndex)
            toSlot = i;
        seen++;
    }
    if (fromSlot < 0 || toSlot < 0)
        return FALSE;

    // Whole-struct moves, so joinRect, joinZoomed and the repair counter travel with the window they
    // describe. Nothing outside this function may hold a Member* across the call - the entries move -
    // and nothing does: every caller reaches the stack through an HWND.
    Member moving = g_members[fromSlot];
    if (fromSlot < toSlot)
    {
        for (int i = fromSlot; i < toSlot; i++)
            g_members[i] = g_members[i + 1];
    }
    else
    {
        for (int i = fromSlot; i > toSlot; i--)
            g_members[i] = g_members[i - 1];
    }
    g_members[toSlot] = moving;

    LogWrite(L"stack  hwnd=0x%p  tab moved %d -> %d (of %d)", (void*)frame, from, toIndex, joined);
    LogRow();

    // Every window in the stack draws the same row, so a reorder is a repaint of all of them.
    StripRefreshTabs();
    return TRUE;
}

// Every strip draws the *stack's* active tab as selected, not its own - which is what makes a
// switch look like one strip standing still while the page changes, rather than two strips
// swapping. It is the whole visual trick, and it costs one line.
void StackActivate(HWND frame)
{
    if (!IsWindow(frame))
        return;

    Member* member = Find(frame);
    if (!member || !member->joined)
        return;

    // Same process and we are being called from a click in the foreground window, so this is
    // allowed - no AttachThreadInput handshake needed, which is what spike 2 had to do from
    // outside.
    SetWindowPos(frame, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
    SetForegroundWindow(frame);

    StackOnFrameActivate(frame);
}

// ---------------------------------------------------------------------------------------------
// Taking a tab out of the stack.
// ---------------------------------------------------------------------------------------------

// Whether detaching is available at all. Asked by the strip when it builds the menu, so the switch
// is read in one place and the drag gesture that will call StackTearOffTab next gets the same answer
// without reading the registry a second time - [[one-copy-of-what-decides-truth]] applied to a flag.
BOOL StackCanTearOff(void)
{
    return g_enabled && g_tearOff;
}

// Where a torn-off window goes.
//
// **Not where it is now, and this is the whole reason there is a placement rule at all.** Every
// window in the stack sits at the same rectangle, so a window that leaves the stack and keeps its
// position is a window the user cannot see has left: it is exactly on top of, or exactly underneath,
// the thing it was pulled out of. "Nothing happened" and "it worked perfectly" would be the same
// picture. So it is offset, by the step Windows itself cascades new windows by.
//
// Not `joinRect` either, which is what Leave restores and is right there. That is where the window
// was before it *ever* joined - usually wherever Word opened it on a cold start, which may be the
// same place as the stack, off the side of the current monitor, or on a monitor that has since been
// unplugged. It answers "undo the stacking", and this is not an undo.
static void PlaceTornOff(HWND frame, const RECT* stackRect, BOOL stackZoomed)
{
    // The caption plus one border: the same step the shell uses to cascade, so a torn-off window
    // lands where a new window would and the title bar of the one underneath stays visible.
    int step = GetSystemMetrics(SM_CYCAPTION) + GetSystemMetrics(SM_CYSIZEFRAME);
    if (step < 24)
        step = 24;

    RECT target = *stackRect;

    // A maximized stack is the case where "it did nothing" is most convincing: two maximized windows
    // are pixel-identical and an offset is not even possible. So the torn-off window comes down to a
    // window-sized window. Three quarters of the work area, which is roughly what Word opens at.
    if (stackZoomed)
    {
        MONITORINFO info;
        memset(&info, 0, sizeof(info));
        info.cbSize = sizeof(info);
        HMONITOR monitor = MonitorFromWindow(frame, MONITOR_DEFAULTTONEAREST);
        if (monitor && GetMonitorInfoW(monitor, &info))
        {
            LONG width  = (info.rcWork.right - info.rcWork.left) * 3 / 4;
            LONG height = (info.rcWork.bottom - info.rcWork.top) * 3 / 4;
            target.left   = info.rcWork.left + (info.rcWork.right - info.rcWork.left - width) / 2;
            target.top    = info.rcWork.top + (info.rcWork.bottom - info.rcWork.top - height) / 2;
            target.right  = target.left + width;
            target.bottom = target.top + height;
        }
    }

    OffsetRect(&target, step, step);

    // Back onto the work area if the offset pushed it off. A window whose title bar is below the
    // bottom of the screen cannot be dragged back, and the stack is often near the bottom-right
    // already because that is where the user left it.
    MONITORINFO info;
    memset(&info, 0, sizeof(info));
    info.cbSize = sizeof(info);
    HMONITOR monitor = MonitorFromRect(&target, MONITOR_DEFAULTTONEAREST);
    if (monitor && GetMonitorInfoW(monitor, &info))
    {
        LONG overRight  = target.right - info.rcWork.right;
        LONG overBottom = target.bottom - info.rcWork.bottom;
        if (overRight > 0)
            OffsetRect(&target, -overRight, 0);
        if (overBottom > 0)
            OffsetRect(&target, 0, -overBottom);
        if (target.left < info.rcWork.left)
            OffsetRect(&target, info.rcWork.left - target.left, 0);
        if (target.top < info.rcWork.top)
            OffsetRect(&target, info.rcWork.top - target.top, 0);
    }

    g_inSync = TRUE;
    SetZoomState(frame, FALSE);
    SetWindowPos(frame, NULL, target.left, target.top,
                 target.right - target.left, target.bottom - target.top,
                 SWP_NOZORDER | SWP_NOACTIVATE);
    g_inSync = FALSE;

    LogWrite(L"stack  hwnd=0x%p  torn off to (%ld,%ld %ldx%ld)%s",
             (void*)frame, target.left, target.top,
             target.right - target.left, target.bottom - target.top,
             stackZoomed ? L" (the stack was maximized, so it comes down to a window)" : L"");
}

// Pull one tab out of the stack and make it a window of its own.
//
// The mechanism is the one that already existed: Leave puts back every single thing the stack did to
// a window - its taskbar button, its Alt+Tab entry, its own rectangle - because a window with none of
// those, sitting underneath another one, is unreachable by any means the user has. That rule was
// written for shutdown and it is exactly what tearing off needs, so there is no second implementation
// of "un-stack a window" here. What this adds is the two things shutdown does not need: the window
// has to go somewhere the user can see it, and it has to *stay* out.
//
// Staying out is the part with teeth. The janitor re-tests membership twice a second and a torn-off
// window still passes every test - visible, sized, holding a document - so without the sticky flag it
// would snap back into the stack within half a second, and from the user's side the command would
// simply not work. See Member::tornOff and the guard in Join.
void StackTearOffTab(HWND frame)
{
    if (!frame || !IsWindow(frame))
        return;

    if (!StackCanTearOff())
    {
        LogWrite(L"stack  hwnd=0x%p  detach refused - switched off", (void*)frame);
        return;
    }

    Member* member = Find(frame);
    if (!member || !member->joined)
    {
        // Already its own window: stacking off, a lone window, or a tab torn off twice because the
        // command was posted from a menu the janitor ticked underneath. Nothing to do, and doing
        // nothing is the right answer rather than an error - the user asked for a state it is in.
        LogWrite(L"stack  hwnd=0x%p  detach ignored - not a tab in a stack", (void*)frame);
        return;
    }

    if (JoinedCount() < 2)
    {
        // The last tab has nowhere to go. Detaching it would produce the window that is already
        // there, with the difference that nothing could ever put it back.
        LogWrite(L"stack  hwnd=0x%p  detach ignored - it is the only tab", (void*)frame);
        return;
    }

    RECT stackRect;
    if (!GetWindowRect(frame, &stackRect))
        return;
    BOOL stackZoomed = IsZoomed(frame) ? TRUE : FALSE;

    // Set before Leave, not after: Leave calls Present() and StripRefreshTabs(), and both of those
    // ask what the membership is. Setting it afterwards would give them one pass over a window that
    // is out of the stack but not yet known to be staying out.
    member->tornOff = TRUE;

    // FALSE, so Leave does not restore joinRect - the placement below is this command's answer to
    // where the window goes, and letting Leave move it first would be two moves the user can see.
    Leave(member, L"torn off by the user", FALSE);

    PlaceTornOff(frame, &stackRect, stackZoomed);

    // Word lays out only the window it is focused on, so the interior has to be refitted to the size
    // it was just given or the document frame keeps the stack's dimensions. Same call and same reason
    // as Leave's own restore path.
    StripRefit(frame);

    // The user asked for this window, so it is the one they get. Raising without focusing would leave
    // them typing into a document that is no longer in front.
    SetWindowPos(frame, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
    SetForegroundWindow(frame);

    LogWrite(L"stack  hwnd=0x%p  torn off, now its own window  (%d left in the stack)",
             (void*)frame, JoinedCount());
}

// ---------------------------------------------------------------------------------------------
// Putting one back.
//
// The inverse of StackTearOffTab, and the reason tearing off is no longer a one-way door. Until this
// existed the only way a torn-off window returned to the row was by losing its document, which is to
// say by being closed - so "I pulled out the wrong one" had no answer that kept the document open.
//
// Almost all of it is machinery that was already here. Join snaps a window onto the stack rectangle,
// puts its tab in the row and repaints every strip; the sticky flag is one line; and where the tab
// lands is the rule written for recycled frames. What this adds is the decision that the user is
// allowed to revoke a decision.
// ---------------------------------------------------------------------------------------------

BOOL StackCanRejoin(HWND frame)
{
    if (!g_enabled || !frame || !IsWindow(frame))
        return FALSE;

    Member* member = Find(frame);
    if (!member || member->joined)
        return FALSE;                  // not one of ours, or already a tab in the row

    // The same test the janitor applies, so a window cannot be dropped into a stack that would then
    // refuse to hold it: on screen, not minimised, big enough to place, and holding a document.
    if (!EligibleToJoin(frame))
        return FALSE;

    // And there has to be something to join. A lone window "rejoining" would arrive at a stack of
    // one, which is the state it is already standing in.
    return JoinedCount() >= 1;
}

void StackJoinTab(HWND frame)
{
    if (!StackCanRejoin(frame))
    {
        // Silent in the sense that nothing is greyed - the strip does not start the gesture at all
        // when this is false - but logged, because the drop is posted and the row can change between
        // the release and the command arriving.
        LogWrite(L"stack  hwnd=0x%p  rejoin refused - already a tab, not placeable, or no stack to "
                 L"join", (void*)frame);
        return;
    }

    Member* member = Find(frame);
    if (!member)
        return;

    // Staying out was a decision about this window and this is the user revoking it. Cleared here
    // rather than inside Join, because Join is also the janitor's path and a torn-off window has to
    // go on failing there twice a second.
    member->tornOff = FALSE;

    // Its tab goes to the END of the row, not back to the place it used to hold - see the header.
    // Same flag, and the same reasoning, as a frame Word has recycled for a new document.
    member->rejoin = TRUE;

    Join(member);

    // After Join, not before: Join is what puts the window on the stack rectangle, and activating
    // first would raise it while it was still standing where it was torn off to.
    StackActivate(frame);

    LogWrite(L"stack  hwnd=0x%p  put back into the stack  (%d in the stack)",
             (void*)frame, JoinedCount());
}

// Close the document behind a tab.
//
// WM_CLOSE, posted, rather than the object model. That is the exact path Word takes when the user
// clicks its own close button, so the save prompt, any document-close macros and Word's own
// bookkeeping all behave identically and none of it has to be reimplemented here. Posted rather than
// sent because we are called from inside a click handler on a window Word is about to destroy.
//
// **The tab is activated first, and that is not cosmetic.** Every window in the stack sits at the
// same rectangle with the active one on top. A modal save prompt belonging to a window underneath
// can end up behind the window in front of it - and an invisible modal dialog is, from the user's
// side, Word beeping and refusing to respond with nothing on screen to explain why. Activating first
// makes that impossible: the prompt is always about the document that is visible. The cost is one
// window switch, and g_returnTo pays it back.
void StackCloseTab(HWND frame)
{
    if (!frame || !IsWindow(frame))
        return;

    // A close aimed at a tab that no longer exists. Commands are posted, not run where they are
    // raised - the tab menu's TrackPopupMenu is a modal loop and the janitor keeps ticking inside it
    // - so a window can lose its last document between the click and the command arriving. IsWindow
    // is not enough to catch that, because Word leaves the window standing when its document closes.
    //
    // Closing it anyway would shut Word down over a tab that had already gone. They closed a
    // document, not a window.
    if (!StripHasDocument(frame))
    {
        LogWrite(L"stack  hwnd=0x%p  close dropped - that window has no document open any more",
                 (void*)frame);
        return;
    }

    Member* member = Find(frame);
    if (!member || !member->joined)
    {
        // Not stacked - a lone window with a strip of its own, or stacking switched off. Still a
        // document, still closable, and there is no z-order to worry about.
        LogWrite(L"stack  hwnd=0x%p  close requested (not in a stack)", (void*)frame);
        g_closeAimed = frame;
        PostMessageW(frame, WM_CLOSE, 0, 0);
        return;
    }

    g_returnTo = NULL;
    if (g_active && g_active != frame && IsWindow(g_active))
    {
        g_returnTo = g_active;
        StackActivate(frame);
    }

    LogWrite(L"stack  hwnd=0x%p  closing this tab%s",
             (void*)frame,
             g_returnTo ? L" (a background one - will return to where the user was)" : L"");

    // Before the post, so the janitor cannot see the document go and read it as the transient it
    // waits a tick for. This close is one we asked for.
    g_closeAimed = frame;
    PostMessageW(frame, WM_CLOSE, 0, 0);
}

// ---------------------------------------------------------------------------------------------
// Closing several tabs at once.
//
// "Close Others" and "Close All" are not a loop over StackCloseTab, and the reason is the save
// prompt. Every close may raise one; it is modal, it runs its own message loop, and it is the user's
// question to answer. Post WM_CLOSE to six windows at once and the second prompt arrives on top of
// the first, for a document the user cannot see behind it - and "Cancel" on the first one closes the
// other five anyway, which is the opposite of what cancel means.
//
// So a batch is a queue with exactly one close in flight, stepped by the janitor. Each tick sorts
// the in-flight tab into one of these states:
//
//   - it is out of the tab row      -> it closed; start the next one
//   - Word is asking about it       -> wait, for as long as it takes. Someone typing a filename into
//                                      Save As is not a hang, so this state has no timeout at all
//   - it was asked about, the
//     question has gone, and the
//     document is still here        -> the user answered and said no. Abandon the rest of the batch
//   - nothing has happened yet      -> wait, but not forever
//
// **The third state requires having seen the question**, and seeing it is the subtle part. The
// first version concluded "declined" from "still open two ticks after WM_CLOSE", and the log caught
// it doing so one second after posting the close, before the prompt had even appeared: Word is
// slower than that, so "still open" on its own is evidence of nothing. The second version required
// the question but looked for it on the janitor's tick - and a prompt that went up and came down
// inside half a second was never seen at all, so the batch decided nothing had happened. Both
// passed every assertion that watched from outside.
//
// So the question is heard as an **event**: a modal dialog disables the window that owns it, and
// EnableWindow sends WM_ENABLE, which the frame subclass hands to StackOnFrameEnable below. A
// message cannot be missed by being quick. The poll is kept as well, for a dialog that is modal
// without disabling anything, but nothing depends on it alone.
//
// The grace after the question disappears is for the other direction: "Save" dismisses the dialog
// and *then* writes the file and closes the document, and concluding "declined" in that gap would
// stop a batch the user had just agreed to. Waiting costs nothing - a document that does close is
// caught by the first state on an earlier tick.
// ---------------------------------------------------------------------------------------------

#define CLOSE_TICKS_AFTER_ANSWER  6     // ~3s for Word to write the file and close the document
#define CLOSE_TICKS_NO_ANSWER    24     // ~12s of nothing at all before giving up on a tab

static HWND g_closeQueue[MAX_MEMBERS];
static int  g_closeCount  = 0;      // how many were queued
static int  g_closeNext   = 0;      // where the queue has got to
static HWND g_closeFlight = NULL;   // the one WM_CLOSE is out for
static BOOL g_closeAsked  = FALSE;  // Word has put a question up about it at some point
static int  g_closeIdle   = 0;      // consecutive ticks with no question and no progress


// The event half of "Word is asking about this one", and the half that can be relied on. A modal
// dialog disables the window that owns it, and EnableWindow sends WM_ENABLE - so the frame's
// subclass hears about the prompt going up whether or not anything happens to be looking at that
// moment. See the poll below for why that matters.
void StackOnFrameEnable(HWND frame, BOOL enabled)
{
    if (!g_closeFlight || frame != g_closeFlight || enabled)
        return;

    if (!g_closeAsked)
    {
        g_closeAsked = TRUE;
        g_closeIdle  = 0;
        LogWrite(L"stack  hwnd=0x%p  Word is asking the user about this document (the window was "
                 L"disabled) - the batch waits", (void*)frame);
    }
}

// The poll half, kept as well as the event and not instead of it. A dialog that is modal by some
// other means than disabling its owner would produce no WM_ENABLE, and this catches it: Word's
// prompts are top-level windows of this process that are not `OpusApp` frames. Neither signal alone
// covers the ground, and the cost of both is two API calls twice a second.
static BOOL WordIsAsking(HWND target)
{
    if (target && IsWindow(target) && !IsWindowEnabled(target))
        return TRUE;

    HWND foreground = GetForegroundWindow();
    if (!foreground)
        return FALSE;

    DWORD pid = 0;
    GetWindowThreadProcessId(foreground, &pid);
    if (pid != GetCurrentProcessId())
        return FALSE;               // another application entirely - not ours to interpret

    wchar_t cls[64] = L"";
    GetClassNameW(foreground, cls, 64);
    return _wcsicmp(cls, L"OpusApp") != 0;
}

// The same test CloseBatchEnd opens with, given a name and a declaration so that work outside this
// file can stand off while a batch runs.
BOOL StackCloseInFlight(void)
{
    return (g_closeCount > 0 || g_closeFlight != NULL) ? TRUE : FALSE;
}

static void CloseBatchEnd(const wchar_t* why)
{
    if (g_closeCount == 0 && !g_closeFlight)
        return;

    int left = g_closeCount - g_closeNext;
    LogWrite(L"stack  close batch %s (%d tab(s) left unclosed)", why, left > 0 ? left : 0);

    g_closeCount  = 0;
    g_closeNext   = 0;
    g_closeFlight = NULL;
    g_closeAsked  = FALSE;
    g_closeIdle   = 0;
}

static void CloseBatchStep(void)
{
    if (g_closeFlight)
    {
        Member* member = Find(g_closeFlight);
        BOOL stillATab = (member && member->joined && EligibleToStay(g_closeFlight)) ? TRUE : FALSE;

        if (!stillATab)
        {
            g_closeFlight = NULL;
            g_closeAsked  = FALSE;
            g_closeIdle   = 0;
        }
        else if (WordIsAsking(g_closeFlight))
        {
            if (!g_closeAsked)
            {
                g_closeAsked = TRUE;
                LogWrite(L"stack  hwnd=0x%p  Word is asking the user about this document "
                         L"(seen by the janitor) - the batch waits", (void*)g_closeFlight);
            }
            g_closeIdle = 0;
            return;
        }
        else if (++g_closeIdle < (g_closeAsked ? CLOSE_TICKS_AFTER_ANSWER : CLOSE_TICKS_NO_ANSWER))
        {
            return;
        }
        else if (g_closeAsked)
        {
            CloseBatchEnd(L"stopped - the question was answered and the document is still open, "
                          L"so the user declined");
            return;
        }
        else
        {
            CloseBatchEnd(L"stopped - WM_CLOSE produced neither a closed document nor a question");
            return;
        }
    }

    while (g_closeNext < g_closeCount)
    {
        HWND next = g_closeQueue[g_closeNext++];
        Member* member = Find(next);

        // It may have closed on its own while it waited its turn - the user's own close button, or
        // Word recycling the frame. Membership now is the only thing that decides.
        if (!member || !member->joined || !IsWindow(next))
            continue;

        g_closeFlight = next;
        g_closeAsked  = FALSE;
        g_closeIdle   = 0;
        StackCloseTab(next);
        return;
    }

    if (g_closeCount > 0)
        CloseBatchEnd(L"finished");
}

// Queue every joined tab except `keep`, **with the active one last**. The user keeps looking at the
// document they were on for as long as the batch allows, and the last prompt they answer is about
// the document they were actually reading rather than one they have never seen.
//
// `after` narrows the range to the tabs that sit after it in `g_members`, which is what "Close Tabs
// to the Right" needs. **`g_members` order is the tab order** - that is the order model, chosen so
// nothing can disagree with the array - so "later in the array" *is* "further right on the row",
// with no index to compute and none to keep consistent. NULL means the whole row, and with NULL
// this function does exactly what it did before, statement for statement.
static void CloseBatchStart(HWND keep, HWND after, const wchar_t* what)
{
    CloseBatchEnd(L"replaced");

    int from = 0;
    if (after)
    {
        from = -1;
        for (int i = 0; i < g_memberCount; i++)
        {
            if (g_members[i].frame == after)
            {
                from = i + 1;
                break;
            }
        }
        if (from < 0)
        {
            // The tab the command was aimed at has left the stack between the menu closing and the
            // command arriving. Nothing here is a safe guess at what the user meant, so nothing is
            // closed - the same rule the stale-context-menu fix follows.
            LogWrite(L"stack  %s: the tab it was invoked on is no longer in the row - nothing closed",
                     what);
            return;
        }
    }

    for (int i = from; i < g_memberCount; i++)
    {
        HWND frame = g_members[i].frame;
        if (!g_members[i].joined || frame == keep || frame == g_active)
            continue;
        if (g_closeCount < MAX_MEMBERS)
            g_closeQueue[g_closeCount++] = frame;
    }

    if (g_active && g_active != keep && g_closeCount < MAX_MEMBERS)
    {
        Member* member = Find(g_active);

        // ...but only if the active tab is inside the range at all. Close Others and Close All pass
        // `after` as NULL and so always are; Close Tabs to the Right must not drag the document the
        // user is looking at into a batch that was never about it.
        BOOL inRange = FALSE;
        for (int i = from; i < g_memberCount && !inRange; i++)
            if (g_members[i].frame == g_active)
                inRange = TRUE;

        if (member && member->joined && inRange)
            g_closeQueue[g_closeCount++] = g_active;
    }

    LogWrite(L"stack  %s: %d tab(s) queued, one at a time", what, g_closeCount);
    CloseBatchStep();
}

void StackCloseOthers(HWND keep)
{
    Member* member = Find(keep);
    if (!g_enabled || !member || !member->joined)
    {
        LogWrite(L"stack  hwnd=0x%p  close others: not a tab in a stack - nothing to close",
                 (void*)keep);
        return;
    }

    if (JoinedCount() < 2)
    {
        LogWrite(L"stack  hwnd=0x%p  close others: it is the only tab", (void*)keep);
        return;
    }

    // The kept tab first, so it is where the user is left standing between one close and the next -
    // g_returnTo is set from whatever is active when each close begins, and that should be the tab
    // they chose to keep rather than whichever document happened to be closing before it.
    StackActivate(keep);
    CloseBatchStart(keep, NULL, L"close others");
}

// Close every tab to the right of this one. The cheap sibling of Close Others: same queue, same
// prompt handling, same left-to-right order - the only new thing is where the range starts.
void StackCloseToRight(HWND from)
{
    Member* member = Find(from);
    if (!g_enabled || !member || !member->joined)
    {
        LogWrite(L"stack  hwnd=0x%p  close to the right: not a tab in a stack - nothing to close",
                 (void*)from);
        return;
    }

    if (StackTabsRightOf(from) == 0)
    {
        LogWrite(L"stack  hwnd=0x%p  close to the right: it is the last tab", (void*)from);
        return;
    }

    // Stand the user on the tab they kept, for the same reason Close Others does: g_returnTo is
    // taken from whatever is active when each close begins.
    StackActivate(from);

    // `from` is passed as both the kept tab and the start of the range. The range alone already
    // excludes it; passing it as `keep` too means a wrong answer from the range lookup still cannot
    // close the tab the user pointed at.
    CloseBatchStart(from, from, L"close to the right");
}

// ---------------------------------------------------------------------------------------------
// What the x means, asked rather than assumed.
//
// The first version of this closed the whole stack outright, which is what the user asked for -
// "you have to close all tabs individually - that needs fixing" - and it is also the one thing
// WordTab does that ends with several of somebody's documents shut. Every unsaved one still gets
// Word's own prompt and a cancel still stops the rest, so nothing can be lost silently; but a person
// who meant "close this one" and got five closes has still been surprised by their own window.
//
// So the x asks. Three answers, and the middle one is the whole point of the change:
//
//   Close all N tabs          - what the button used to do on its own
//   Close only this document  - what Word does without WordTab
//   Cancel                    - nothing happens
//
// A task dialog rather than a MessageBox, because "Yes / No / Cancel" over the question "close all
// tabs?" is exactly the shape where people click the wrong one: the answers here are two different
// actions, not a yes and a no, and command links let each one say what it does. Reached through
// GetProcAddress rather than by linking it: TaskDialogIndirect exists only in version 6 of
// comctl32, which is present through Word's own activation context on every machine this will meet
// - and "almost certainly present" is not a thing to stake a close button on. If it is not there,
// the MessageBox below asks the same question in the words that fit three fixed buttons.
// ---------------------------------------------------------------------------------------------

#define CLOSE_ASK_ALL     101
#define CLOSE_ASK_ONE     102

typedef HRESULT (WINAPI *TaskDialogIndirectFn)(const TASKDIALOGCONFIG*, int*, int*, BOOL*);

// Which answer, as one of the three IDs above. IDCANCEL for anything that is not a clear yes to one
// of the two actions - a dialog that failed to appear included, because a close button that acts on
// an answer nobody gave is worse than one that does nothing.
static int AskWhatToClose(HWND frame, int tabs)
{
    wchar_t name[128];
    StripTabName(frame, name, 128);

    wchar_t heading[160];
    _snwprintf(heading, 160, L"Close all %d tabs in this window?", tabs);
    heading[159] = L'\0';

    wchar_t allText[160];
    _snwprintf(allText, 160, L"Close all %d tabs\nWord will ask about any with unsaved changes.", tabs);
    allText[159] = L'\0';

    wchar_t oneText[320];
    _snwprintf(oneText, 320, L"Close only this document\n\"%s\" closes; the other %d stay open.",
               name, tabs - 1);
    oneText[319] = L'\0';

    HMODULE comctl = GetModuleHandleW(L"comctl32.dll");
    TaskDialogIndirectFn ask = comctl
        ? (TaskDialogIndirectFn)(void*)GetProcAddress(comctl, "TaskDialogIndirect")
        : NULL;

    if (ask)
    {
        TASKDIALOG_BUTTON buttons[2];
        buttons[0].nButtonID     = CLOSE_ASK_ALL;
        buttons[0].pszButtonText = allText;
        buttons[1].nButtonID     = CLOSE_ASK_ONE;
        buttons[1].pszButtonText = oneText;

        TASKDIALOGCONFIG config;
        memset(&config, 0, sizeof(config));
        config.cbSize             = sizeof(config);
        config.hwndParent         = frame;
        config.hInstance          = g_module;
        config.dwFlags            = TDF_USE_COMMAND_LINKS | TDF_ALLOW_DIALOG_CANCELLATION |
                                    TDF_POSITION_RELATIVE_TO_WINDOW;
        config.dwCommonButtons    = TDCBF_CANCEL_BUTTON;
        config.pszWindowTitle     = L"WordTab";
        config.pszMainInstruction = heading;
        config.pButtons           = buttons;
        config.cButtons           = 2;
        config.nDefaultButton     = CLOSE_ASK_ALL;

        int pressed = 0;
        HRESULT hr = ask(&config, &pressed, NULL, NULL);
        if (SUCCEEDED(hr))
        {
            LogWrite(L"stack  hwnd=0x%p  asked what the x meant: %s", (void*)frame,
                     pressed == CLOSE_ASK_ALL ? L"close all"
                   : pressed == CLOSE_ASK_ONE ? L"close only this one"
                                              : L"cancelled");
            return pressed;
        }

        LogWrite(L"stack  hwnd=0x%p  TaskDialogIndirect failed (hr=0x%08lX) - falling back to a "
                 L"message box", (void*)frame, (unsigned long)hr);
    }

    // The fallback, and it is worded for the buttons it has rather than being the same sentence with
    // worse controls. Yes/No/Cancel can carry two actions and an escape only if the question names
    // which is which.
    wchar_t text[420];
    _snwprintf(text, 420,
               L"This window has %d tabs in it.\n\n"
               L"Yes\tclose all %d\n"
               L"No\tclose only \"%s\"\n"
               L"Cancel\tleave everything open",
               tabs, tabs, name);
    text[419] = L'\0';

    int answer = MessageBoxW(frame, text, L"WordTab", MB_YESNOCANCEL | MB_ICONQUESTION);
    LogWrite(L"stack  hwnd=0x%p  asked what the x meant (message box): %s", (void*)frame,
             answer == IDYES ? L"close all" : answer == IDNO ? L"close only this one" : L"cancelled");

    if (answer == IDYES) return CLOSE_ASK_ALL;
    if (answer == IDNO)  return CLOSE_ASK_ONE;
    return IDCANCEL;
}

// The title bar's x, Alt+F4, and the window menu's Close - all of which arrive as SC_CLOSE.
//
// The user's words: "you have to close all tabs individually - that needs fixing next time we edit."
// Confirmed in their log at 14:27-14:28, four frames going down one at a time. The stack is one
// window to the user in every other respect - one taskbar button, one Alt+Tab entry, it moves and
// resizes as one - so the button that closes a window has to close the window they can see, which is
// all of it. Every tabbed application works this way.
//
// TRUE means the command has been taken over and the caller must swallow it. Everything that is not
// unambiguously "the user pressed close on a stack of several documents" answers FALSE and lets
// Word do exactly what it did before:
//
//   - not stacked, or a row of one: Word's own close is already right, and routing a single document
//     through the batch machinery would only add a tick of latency
//   - a batch already running: this IS the batch closing the tabs, and the WM_CLOSEs it posts must
//     reach Word. They arrive as WM_CLOSE rather than SC_CLOSE, so they do not come through here at
//     all - but a user pressing x again while a batch is part-way through must not start a second
//     one on top of it
//
// The batch itself is what makes this safe: one close at a time, each with its own save prompt, and
// a prompt the user cancels abandons the rest. Closing a stack of five documents with unsaved work
// asks five questions in turn, and Cancel on the second one leaves the remaining three open.
BOOL StackCloseWindowCommand(HWND frame)
{
    if (!g_enabled || !g_closeStack)
        return FALSE;

    Member* member = Find(frame);
    if (!member || !member->joined)
        return FALSE;

    if (JoinedCount() < 2)
        return FALSE;

    if (StackCloseInFlight())
    {
        LogWrite(L"stack  hwnd=0x%p  close pressed while a batch is already running - left alone",
                 (void*)frame);
        return FALSE;
    }

    int tabs = JoinedCount();

    // Straight through when the user has said they never want the question. See StackStart: 1 asks,
    // 2 does what the first version of this did.
    if (g_closeStack == 2)
    {
        LogWrite(L"stack  hwnd=0x%p  the window's own close: %d tab(s) go with it "
                 L"(TabCloseStack=2, not asking)", (void*)frame, tabs);
        StackCloseAll(frame);
        return TRUE;
    }

    int answer = AskWhatToClose(frame, tabs);

    // **Everything is re-read after the dialog, nothing is carried across it.** A task dialog runs a
    // modal loop, the janitor ticks inside it, and the row can be a different row by the time an
    // answer comes back - Word can close a document, or open one, while the question is on screen.
    // The same rule the tab context menu follows, and for the same reason.
    member = Find(frame);
    if (!member || !member->joined || !IsWindow(frame))
    {
        LogWrite(L"stack  hwnd=0x%p  the tab the x was pressed on has left the row while the "
                 L"question was up - nothing closed", (void*)frame);
        return TRUE;
    }

    if (answer == CLOSE_ASK_ALL)
    {
        LogWrite(L"stack  hwnd=0x%p  the window's own close: %d tab(s) go with it",
                 (void*)frame, JoinedCount());
        StackCloseAll(frame);
        return TRUE;
    }

    if (answer == CLOSE_ASK_ONE)
    {
        // Word's own behaviour, routed through StackCloseTab rather than let through to Word: that
        // is the one path that knows about the row - which tab to leave the user standing on, and
        // that a close aimed at a window whose document has already gone must be dropped rather
        // than shutting Word down.
        LogWrite(L"stack  hwnd=0x%p  the window's own close: this document only", (void*)frame);
        StackCloseTab(frame);
        return TRUE;
    }

    LogWrite(L"stack  hwnd=0x%p  the window's own close: cancelled, nothing closed", (void*)frame);
    return TRUE;
}

// Where a window about to be created should be put, so that it is never seen anywhere else.
//
// The queue's last item: opening an EXISTING document shows the new window a few inches above the
// stack for a moment before it joins. The user called it minor and it is, but its cause is not - the
// window is created wherever Word asked for it, shown, and only then moved, because the CBT hook
// POSTS and the placement happens after Word has already put it on screen. The fix is to answer the
// question at the moment it is asked: a window created at the stack's rectangle has nowhere to flash
// from. See CbtProc.
//
// FALSE when there is nothing to match - no stack, or one that is minimised or being moved by us -
// and then Word's own choice stands, which is what happens today.
BOOL StackProposeCreateRect(RECT* out)
{
    if (!g_enabled || !out || g_minimized || g_inSync)
        return FALSE;
    if (!g_active || !IsWindow(g_active) || IsIconic(g_active))
        return FALSE;
    if (JoinedCount() < 1)
        return FALSE;

    return GetWindowRect(g_active, out) ? TRUE : FALSE;
}

void StackCloseAll(HWND anyTab)
{
    Member* member = Find(anyTab);
    if (!g_enabled || !member || !member->joined)
    {
        // No stack to enumerate - stacking switched off, or a window that never joined. Its strip
        // shows exactly one tab, so "close all" means that one document and there is no batch.
        // Or it shows none, because the window has no document left: StackCloseTab drops the
        // command in that case rather than closing the window, and this is the route a stale menu
        // command takes to get there.
        LogWrite(L"stack  hwnd=0x%p  close all: not in a stack, so this is its one document",
                 (void*)anyTab);
        StackCloseTab(anyTab);
        return;
    }

    CloseBatchStart(NULL, NULL, L"close all");
}
