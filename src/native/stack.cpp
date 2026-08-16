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
#include <string.h>
#include <wchar.h>

#define MAX_MEMBERS 256

struct Member
{
    HWND frame;
    BOOL joined;
    RECT joinRect;      // where the window was before we stacked it, so it can be put back
    BOOL joinZoomed;
    int  repairs;       // consecutive attempts to put this one back in step - see Reconcile
};

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
static BOOL   g_enabled     = TRUE;
static BOOL   g_started     = FALSE;
static BOOL   g_inSync      = FALSE;  // our own SetWindowPos calls come back through the subclass
static BOOL   g_altTab      = TRUE;
static BOOL   g_minimized   = FALSE;  // the whole stack is down on the taskbar

// Where to put the user back when the active tab goes away, set only by StackCloseTab. Closing a
// background tab has to activate it first (see there), which moves the user off the document they
// were reading; this is how they get back to it rather than to whatever happens to be next in the
// row. Always re-validated before use - the window may have closed in the meantime.
static HWND   g_returnTo    = NULL;

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
        StripSetNatural(frame, &natural);
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

static void Join(Member* member)
{
    if (member->joined)
        return;

    member->joined = TRUE;
    GetWindowRect(member->frame, &member->joinRect);
    member->joinZoomed = IsZoomed(member->frame) ? TRUE : FALSE;

    HWND master = (g_active && g_active != member->frame && Find(g_active) && Find(g_active)->joined)
                ? g_active : FirstJoined(member->frame);

    if (!master)
    {
        // First window in: it defines where the stack is.
        g_active = member->frame;
        LogWrite(L"stack  hwnd=0x%p  joined as the first window (it defines the stack rect)",
                 (void*)member->frame);
    }
    else
    {
        MatchTo(master, member->frame);
        LogWrite(L"stack  hwnd=0x%p  joined, snapped to 0x%p  (%d in the stack)",
                 (void*)member->frame, (void*)master, JoinedCount());
    }

    Present();
    StripRefreshTabs();
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
    TaskbarStart();

    LogWrite(L"StackStart  stacking=%s  altTab suppression=%s",
             g_enabled ? L"on" : L"off (HKCU\\Software\\WordTab\\Stack=0)",
             g_altTab ? L"on" : L"off");
}

void StackAttachFrame(HWND frame)
{
    if (!g_enabled || !frame || Find(frame) || g_memberCount >= MAX_MEMBERS)
        return;

    Member* member = &g_members[g_memberCount++];
    memset(member, 0, sizeof(*member));
    member->frame = frame;

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
                StripSetNatural(g_members[i].frame, &natural);
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
        StripSetNatural(g_members[i].frame, natural);
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
void StackJanitor(void)
{
    if (!g_enabled)
        return;

    for (int i = 0; i < g_memberCount; i++)
    {
        Member* member = &g_members[i];

        if (!member->joined)
        {
            if (EligibleToJoin(member->frame))
                Join(member);
        }
        else if (!EligibleToStay(member->frame))
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
            const wchar_t* why = L"gone";
            if (IsWindow(member->frame))
                why = IsWindowVisible(member->frame) ? L"no document open" : L"hidden or minimised";
            Leave(member, why, FALSE);
        }
    }

    // Word activates windows without always telling our frame procedure, so trust the system over
    // our own bookkeeping.
    HWND foreground = GetForegroundWindow();
    if (foreground && foreground != g_active)
    {
        Member* member = Find(foreground);
        if (member && member->joined)
            StackOnFrameActivate(foreground);
    }

    // After the membership pass above, deliberately: whether the tab we asked Word to close has
    // actually gone is a membership question, and this reads the answer that loop just wrote.
    CloseBatchStep();

    Reconcile();
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
