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
// `_WwF`, and it leaves when either stops being true - never keyed on creation and destruction.

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

static Member g_members[MAX_MEMBERS];
static int    g_memberCount = 0;      // frames we know about, joined or not
static HWND   g_active      = NULL;
static BOOL   g_enabled     = TRUE;
static BOOL   g_started     = FALSE;
static BOOL   g_inSync      = FALSE;  // our own SetWindowPos calls come back through the subclass
static BOOL   g_altTab      = TRUE;
static BOOL   g_minimized   = FALSE;  // the whole stack is down on the taskbar

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
// size, and carrying a document frame. A Start-screen window - Word launched with no document -
// has no `_WwF` and is deliberately left out: it has no document to be a tab for.
//
// To *stay*, all that is required is that it still exists and still has a document. A minimised
// window is still a document and still deserves its tab; more to the point, when the whole stack
// goes down to the taskbar together, dropping every window out of the stack would leave nothing to
// bring back and the only window with a taskbar button would come back alone.
static BOOL EligibleToStay(HWND frame)
{
    if (!frame || !IsWindow(frame) || !IsWindowVisible(frame))
        return FALSE;
    return StripHasDocumentFrame(frame);
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
        g_active = FirstJoined(member->frame);
        if (g_active)
            SetWindowPos(g_active, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
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
            // Left where it is, deliberately. Word hides and re-shows frames of its own accord -
            // closing one document was measured to hide a *different* window for a moment - and
            // moving a window back to its old position every time it blinks makes documents jump
            // around the screen for no reason the user can see. It is hidden; nobody is looking at
            // it; and when it comes back the join snaps it to the stack again. Only StackStop puts
            // windows back where they came from.
            Leave(member, IsWindow(member->frame) ? L"hidden or minimised" : L"gone", FALSE);
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
