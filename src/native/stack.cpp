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
};

static Member g_members[MAX_MEMBERS];
static int    g_memberCount = 0;      // frames we know about, joined or not
static HWND   g_active      = NULL;
static BOOL   g_enabled     = TRUE;
static BOOL   g_started     = FALSE;
static BOOL   g_inSync      = FALSE;  // our own SetWindowPos calls come back through the subclass

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

// Is this window one we should be stacking? Visible, not minimised, and carrying a document frame.
// A Start-screen window - Word launched with no document - has no `_WwF` and is deliberately left
// out: it has no document to be a tab for.
static BOOL Eligible(HWND frame)
{
    if (!frame || !IsWindow(frame) || !IsWindowVisible(frame) || IsIconic(frame))
        return FALSE;

    RECT rect;
    if (!GetWindowRect(frame, &rect))
        return FALSE;
    if ((rect.right - rect.left) < 200 || (rect.bottom - rect.top) < 200)
        return FALSE;

    return StripHasDocumentFrame(frame);
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

// Make `frame` match `master` exactly: same rect, and the same maximized-or-not *state* rather
// than only the same rectangle. Spike 2 could copy rectangles only, which left a stacked window
// looking maximized without being maximized - its title bar then behaved oddly on a double click.
// In-process the state itself is reachable, so it is used.
static void MatchTo(HWND master, HWND frame)
{
    if (!IsWindow(master) || !IsWindow(frame) || master == frame)
        return;

    g_inSync = TRUE;

    if (IsZoomed(master))
    {
        if (!IsZoomed(frame))
            ShowWindow(frame, SW_MAXIMIZE);
    }
    else
    {
        if (IsZoomed(frame))
            ShowWindow(frame, SW_RESTORE);

        RECT rect;
        if (GetWindowRect(master, &rect))
        {
            SetWindowPos(frame, NULL, rect.left, rect.top,
                         rect.right - rect.left, rect.bottom - rect.top,
                         SWP_NOZORDER | SWP_NOACTIVATE);
        }
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

    StripRefreshTabs();
}

static void Leave(Member* member, const wchar_t* why, BOOL restorePosition)
{
    if (!member->joined)
        return;
    member->joined = FALSE;

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
    LogWrite(L"StackStart  stacking=%s", g_enabled ? L"on" : L"off (HKCU\\Software\\WordTab\\Stack=0)");
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
    // its layout dismantled first, and broadcasting that would take every other document's layout
    // down with it - measured, once, and it put every strip on top of its ribbon.
    if (!IsWindowVisible(frame))
        return;

    for (int i = 0; i < g_memberCount; i++)
    {
        if (!g_members[i].joined || g_members[i].frame == frame)
            continue;
        StripSetNatural(g_members[i].frame, natural);
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
        BOOL eligible = Eligible(member->frame);

        if (eligible && !member->joined)
            Join(member);
        else if (!eligible && member->joined)
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
}

void StackStop(void)
{
    if (!g_started)
        return;
    g_started = FALSE;

    for (int i = 0; i < g_memberCount; i++)
        Leave(&g_members[i], L"shutdown", TRUE);

    g_memberCount = 0;
    g_active = NULL;
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
