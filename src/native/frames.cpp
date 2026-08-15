// WordTab - the in-process subclass of Word's OpusApp frame window.
//
// This is the reason the add-in had to get inside WINWORD at all. Both spikes watched Word from
// another process on a 30ms poll, and that is good enough for resize, maximize and Backstage - but
// not for a drag. Word moves a dragged window inside a modal move/size loop at its own frame rate,
// and a poller sampling every 30ms cannot keep other windows with it: they lag and peek out from
// under the one being dragged. In-process we are called *inside* that loop, on Word's own thread,
// before each move happens.
//
// What is here:
//   - discovery: every OpusApp frame this process owns, including ones created later
//   - a subclass on each, chained properly so other add-ins are not broken
//   - a move trace: every position change during a drag, timed with QueryPerformanceCounter, so
//     the "a 30ms poll cannot keep up" claim is a measurement rather than an assertion
//   - a small, switchable demonstration: during a drag, move the other frames in the same frame
//
// Rules this file lives by, all of them learned the expensive way:
//   - Nothing throws and nothing escapes. An exception crossing back into Word gets the add-in put
//     in Resiliency\DisabledItems, which is silent and sticky.
//   - Nothing here is slow. This code runs between Word and its own window procedure; a file write
//     per message would make Word visibly stutter, so the trace goes to memory and is flushed once
//     the drag ends.
//   - Detach must be real and must survive being done out of order. Another add-in may have
//     subclassed the same window after us.

#include "wordtab.h"
#include <commctrl.h>
#include <string.h>
#include <stdio.h>

// The class name of Word's top-level frame. One frame per document window; several exist at once.
static const wchar_t* const kFrameClass = L"OpusApp";

// Identifies our link in a window's subclass chain. Any value, as long as it is ours alone.
static const UINT_PTR kSubclassId = 0x57544142;   // 'WTAB'

// Posted to the coordinator window when the CBT hook sees a frame being created.
#define WM_WORDTAB_FRAME_CREATED  (WM_APP + 1)

static const wchar_t* const kCoordinatorClass = L"WordTabCoordinator";

// ---------------------------------------------------------------------------------------------
// Module state.
//
// Everything below runs on Word's UI thread: the callbacks Word makes into us, the CBT hook, the
// coordinator's messages and the subclass procedure are all on that one thread. The lock exists
// only because the COM object is registered ThreadingModel=Both, so OnDisconnection is not
// *guaranteed* to arrive on the UI thread even though it always has.
// ---------------------------------------------------------------------------------------------

#define MAX_FRAMES 256

static CRITICAL_SECTION g_lock;
static BOOL   g_lockReady   = FALSE;
static BOOL   g_started     = FALSE;
static DWORD  g_uiThread    = 0;
static HHOOK  g_cbtHook     = NULL;
static HWND   g_coordinator = NULL;
static ATOM   g_coordClass  = 0;

static HWND   g_frames[MAX_FRAMES];
static int    g_frameCount  = 0;

// The lockstep demonstration's state - see FollowDrag below. Declared up here because the move
// trace reports on it, and the trace comes first.
static BOOL   g_followDrag  = TRUE;
static BOOL   g_inFollow    = FALSE; // re-entry guard: our own SetWindowPos calls come back to us
static int    g_followCount = 0;     // frames of the drag on which we moved the other windows
static int    g_followMax   = 0;     // most followers moved in a single frame

static void EnsureLock(void)
{
    // First call is on Word's single loading thread, long before anything else can race it.
    if (!g_lockReady)
    {
        InitializeCriticalSection(&g_lock);
        g_lockReady = TRUE;
    }
}

// ---------------------------------------------------------------------------------------------
// The move trace.
//
// A drag is a modal loop: Word does not return to its message pump, it runs its own, and it sends
// us WM_WINDOWPOSCHANGING before every single step. Recording those in memory and reporting once
// at the end costs nothing during the drag, which is the point - the measurement must not perturb
// the thing being measured.
// ---------------------------------------------------------------------------------------------

struct MoveSample
{
    LONGLONG stamp;      // QueryPerformanceCounter, not GetTickCount: that has ~16ms resolution
    UINT     msg;        // and would quantise away the very gaps we are trying to measure
    UINT     flags;      // WINDOWPOS flags
    LONG     x, y, cx, cy;
};

#define TRACE_CAPACITY 1024

static MoveSample g_trace[TRACE_CAPACITY];
static int        g_traceCount  = 0;
static int        g_traceLost   = 0;   // samples past capacity, so a long drag reports honestly
static BOOL       g_inModalLoop = FALSE;
static HWND       g_modalWindow = NULL;
static LONGLONG   g_modalStart  = 0;
static LONGLONG   g_qpcFreq     = 0;

static LONGLONG Now(void)
{
    LARGE_INTEGER value;
    if (!QueryPerformanceCounter(&value))
        return 0;
    return value.QuadPart;
}

// Ticks to microseconds. Integer maths throughout: MinGW linked against msvcrt cannot be trusted
// to format %f, and this is evidence, so it has to be right.
static LONGLONG ToMicros(LONGLONG ticks)
{
    if (g_qpcFreq <= 0)
        return 0;
    return (ticks * 1000000LL) / g_qpcFreq;
}

static void TraceReset(HWND hwnd)
{
    g_traceCount  = 0;
    g_traceLost   = 0;
    g_followCount = 0;
    g_followMax   = 0;
    g_modalWindow = hwnd;
    g_modalStart  = Now();
    g_inModalLoop = TRUE;
}

static void TraceRecord(UINT msg, const WINDOWPOS* pos)
{
    if (!g_inModalLoop)
        return;
    if (g_traceCount >= TRACE_CAPACITY)
    {
        g_traceLost++;
        return;
    }

    MoveSample* sample = &g_trace[g_traceCount++];
    sample->stamp = Now();
    sample->msg   = msg;
    sample->flags = pos ? pos->flags : 0;
    sample->x     = pos ? pos->x  : 0;
    sample->y     = pos ? pos->y  : 0;
    sample->cx    = pos ? pos->cx : 0;
    sample->cy    = pos ? pos->cy : 0;
}

// Report the drag. This is the slice's central measurement, so it says plainly what a 30ms poll
// would have seen - the same number the out-of-process spikes were living with.
static void TraceFlush(void)
{
    if (!g_inModalLoop)
        return;
    g_inModalLoop = FALSE;

    LONGLONG endStamp   = Now();
    LONGLONG totalMicros = ToMicros(endStamp - g_modalStart);

    if (g_traceCount < 2)
    {
        LogWrite(L"drag  hwnd=0x%p  %d position change(s) in %lld.%lldms - too few to time",
                 (void*)g_modalWindow, g_traceCount,
                 totalMicros / 1000, (totalMicros % 1000) / 100);
        return;
    }

    LONGLONG spanMicros = ToMicros(g_trace[g_traceCount - 1].stamp - g_trace[0].stamp);

    // Every step of the drag shows up twice, as CHANGING then CHANGED. The honest count of "how
    // often did Word move the window" is the number of CHANGINGs, and that is what the poll
    // comparison has to be measured against - counting both would double the claim.
    LONGLONG minGap = 0x7FFFFFFFFFFFFFFFLL;
    LONGLONG maxGap = 0;
    int under8 = 0, under16 = 0, under33 = 0, under50 = 0, over50 = 0;
    int gapsUnder30 = 0, gaps = 0;
    int updates = 0, moved = 0, resized = 0;
    LONGLONG previous = 0;

    for (int i = 0; i < g_traceCount; i++)
    {
        if (g_trace[i].msg != WM_WINDOWPOSCHANGING)
            continue;

        // Compare against the first frame rather than reading SWP_NOSIZE: Word leaves that flag
        // clear throughout a caption drag, so trusting it reports every move as a resize too.
        if (!(g_trace[i].flags & SWP_NOMOVE) &&
            (g_trace[i].x != g_trace[0].x || g_trace[i].y != g_trace[0].y)) moved++;
        if (g_trace[i].cx != g_trace[0].cx || g_trace[i].cy != g_trace[0].cy) resized++;
        updates++;

        if (updates > 1)
        {
            LONGLONG gap = ToMicros(g_trace[i].stamp - previous);
            gaps++;
            if (gap < minGap) minGap = gap;
            if (gap > maxGap) maxGap = gap;

            if      (gap <=  8000) under8++;
            else if (gap <= 16000) under16++;
            else if (gap <= 33000) under33++;
            else if (gap <= 50000) under50++;
            else                   over50++;

            // The whole argument in one counter: a gap shorter than the poll interval is a frame
            // the poller could not have seen.
            if (gap < 30000) gapsUnder30++;
        }
        previous = g_trace[i].stamp;
    }

    if (gaps == 0)
        minGap = 0;
    LONGLONG meanGap = gaps > 0 ? (spanMicros / gaps) : 0;

    // What a 30ms poller could have managed over the same span, at its theoretical best.
    LONGLONG pollSamples = (spanMicros / 30000) + 1;
    int missedPercent = 0;
    if (updates > 0 && pollSamples < updates)
        missedPercent = (int)(((LONGLONG)updates - pollSamples) * 100 / updates);

    LogWrite(L"drag  hwnd=0x%p  %d position updates (%d moved, %d resized) over %lld.%lldms"
             L"  [%d messages traced]%s",
             (void*)g_modalWindow, updates, moved, resized,
             spanMicros / 1000, (spanMicros % 1000) / 100, g_traceCount,
             g_traceLost > 0 ? L"  [trace full - more were dropped]" : L"");

    LogWrite(L"drag  gap between updates: min %lld.%lldms  mean %lld.%lldms  max %lld.%lldms",
             minGap / 1000, (minGap % 1000) / 100,
             meanGap / 1000, (meanGap % 1000) / 100,
             maxGap / 1000, (maxGap % 1000) / 100);

    LogWrite(L"drag  gap histogram:  <=8ms:%d  <=16ms:%d  <=33ms:%d  <=50ms:%d  >50ms:%d",
             under8, under16, under33, under50, over50);

    LogWrite(L"drag  VERDICT: a 30ms poll could take at most %lld samples of these %d updates "
             L"(%d%% unseen); %d of %d gaps were shorter than 30ms",
             pollSamples, updates, missedPercent, gapsUnder30, gaps);

    // The other half of the point: not only did we see every frame, we acted on every frame.
    LogWrite(L"drag  lockstep: moved %d other frame(s) on %d of %d updates, in the same "
             L"WM_WINDOWPOSCHANGING%s",
             g_followMax, g_followCount, updates,
             g_followDrag ? L"" : L"  (FollowDrag is off)");

    // A short sample of the raw frames, as evidence that the numbers above came from somewhere.
    int show = g_traceCount < 8 ? g_traceCount : 8;
    for (int i = 0; i < show; i++)
    {
        LONGLONG at = ToMicros(g_trace[i].stamp - g_trace[0].stamp);
        LogWrite(L"drag    +%lld.%lldms  %s  pos=(%ld,%ld) size=(%ldx%ld) flags=0x%08X",
                 at / 1000, (at % 1000) / 100,
                 g_trace[i].msg == WM_WINDOWPOSCHANGING ? L"CHANGING" : L"CHANGED ",
                 g_trace[i].x, g_trace[i].y, g_trace[i].cx, g_trace[i].cy,
                 g_trace[i].flags);
    }
    if (g_traceCount > show)
        LogWrite(L"drag    ... %d more", g_traceCount - show);
}

// ---------------------------------------------------------------------------------------------
// The lockstep demonstration.
//
// Not the stacking engine - that is a later slice. This is the smallest thing that makes the
// in-process win visible without reading a log: while one frame is dragged, every other frame
// moves by the same delta, in the same frame, before either is painted. Two Word windows side by
// side move as one. The out-of-process spike could not do this at all; its workaround was to hide
// the other windows for the duration of the drag.
//
// Switchable at HKCU\Software\WordTab\FollowDrag (default on), so it can be turned off without a
// rebuild - same shape as ShowLoadBanner.
// ---------------------------------------------------------------------------------------------

// Declared in wordtab.h and shared with strip.cpp: every switch WordTab has is a DWORD under the
// same key, and one reader for all of them is one place for the "absent means default" rule.
BOOL WordTabReadFlag(const wchar_t* name, BOOL defaultValue)
{
    DWORD value = 0;
    DWORD size = sizeof(value);
    if (RegGetValueW(HKEY_CURRENT_USER, L"Software\\WordTab", name,
                     RRF_RT_REG_DWORD, NULL, &value, &size) != ERROR_SUCCESS)
    {
        return defaultValue;
    }
    return value != 0;
}

static void FollowDrag(HWND dragged, const WINDOWPOS* pos)
{
    if (!g_followDrag || g_inFollow || !pos)
        return;
    if (pos->flags & SWP_NOMOVE)
        return;

    RECT current;
    if (!GetWindowRect(dragged, &current))
        return;

    // Word does not set SWP_NOSIZE while a window is dragged by its caption, even though the size
    // never changes - measured: every frame of a move drag arrives with flags 0x00080214 and a cx
    // and cy identical to the window's own. So the flag cannot be trusted to separate a move from
    // a resize, and the numbers have to be compared instead. Getting this wrong is silent: the
    // followers simply never move.
    if (!(pos->flags & SWP_NOSIZE))
    {
        LONG width  = current.right - current.left;
        LONG height = current.bottom - current.top;
        if (pos->cx != width || pos->cy != height)
            return;                     // a genuine resize - out of scope for the demonstration
    }

    LONG dx = pos->x - current.left;
    LONG dy = pos->y - current.top;
    if (dx == 0 && dy == 0)
        return;

    // Snapshot the table: SetWindowPos below re-enters this file, and holding the lock across it
    // would be a deadlock waiting to happen.
    HWND followers[MAX_FRAMES];
    int  count = 0;

    EnsureLock();
    EnterCriticalSection(&g_lock);
    for (int i = 0; i < g_frameCount; i++)
    {
        HWND candidate = g_frames[i];
        if (candidate == dragged || !IsWindow(candidate))
            continue;
        if (!IsWindowVisible(candidate) || IsIconic(candidate))
            continue;
        followers[count++] = candidate;
    }
    LeaveCriticalSection(&g_lock);

    if (count == 0)
        return;

    g_inFollow = TRUE;

    // DeferWindowPos so every window moves in one atomic pass rather than one repaint each. This
    // is the primitive the stacking engine will want, so it is worth using here.
    HDWP batch = BeginDeferWindowPos(count);
    for (int i = 0; i < count; i++)
    {
        RECT rect;
        if (!GetWindowRect(followers[i], &rect))
            continue;

        if (batch)
        {
            batch = DeferWindowPos(batch, followers[i], NULL,
                                   rect.left + dx, rect.top + dy, 0, 0,
                                   SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOREDRAW);
        }
        else
        {
            SetWindowPos(followers[i], NULL, rect.left + dx, rect.top + dy, 0, 0,
                         SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOREDRAW);
        }
    }
    if (batch)
        EndDeferWindowPos(batch);

    g_followCount++;
    if (count > g_followMax)
        g_followMax = count;

    g_inFollow = FALSE;
}

// ---------------------------------------------------------------------------------------------
// The subclass procedure. Everything Word's frame receives passes through here first.
// ---------------------------------------------------------------------------------------------

static void DetachFrame(HWND hwnd, const wchar_t* why);

static const wchar_t* SizeTypeName(WPARAM type)
{
    switch (type)
    {
    case SIZE_RESTORED:  return L"restored";
    case SIZE_MINIMIZED: return L"minimized";
    case SIZE_MAXIMIZED: return L"maximized";
    case SIZE_MAXSHOW:   return L"maxshow";
    case SIZE_MAXHIDE:   return L"maxhide";
    default:             return L"?";
    }
}

static LRESULT CALLBACK FrameSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                                          UINT_PTR idSubclass, DWORD_PTR refData)
{
    (void)idSubclass;
    (void)refData;

    switch (msg)
    {
    case WM_ENTERSIZEMOVE:
        // The modal loop starts here. From now until WM_EXITSIZEMOVE, Word runs its own pump and
        // never returns to its message loop - which is exactly why an outside watcher loses.
        LogWrite(L"WM_ENTERSIZEMOVE  hwnd=0x%p  (modal move/size loop begins)", (void*)hwnd);
        TraceReset(hwnd);
        break;

    case WM_EXITSIZEMOVE:
        TraceFlush();
        LogWrite(L"WM_EXITSIZEMOVE   hwnd=0x%p", (void*)hwnd);
        break;

    case WM_WINDOWPOSCHANGING:
        // The hook point that matters. We are handed the *proposed* position, before the move
        // happens, and may read or change it. This is where the stacking engine will keep every
        // window together.
        if (!g_inFollow)
        {
            const WINDOWPOS* pos = (const WINDOWPOS*)lParam;
            TraceRecord(msg, pos);
            if (g_inModalLoop && hwnd == g_modalWindow)
                FollowDrag(hwnd, pos);
        }
        break;

    case WM_WINDOWPOSCHANGED:
        if (!g_inFollow)
            TraceRecord(msg, (const WINDOWPOS*)lParam);
        break;

    case WM_SIZE:
        // Logged unconditionally *outside* a drag only: inside one it floods, and the trace has it
        // covered anyway. The maximize/restore state carried here is the thing spike 2 could not
        // propagate - it copied rectangles, so a stacked window looked maximized without being it.
        if (!g_inModalLoop)
        {
            LogWrite(L"WM_SIZE  hwnd=0x%p  %s  %dx%d", (void*)hwnd,
                     SizeTypeName(wParam), (int)LOWORD(lParam), (int)HIWORD(lParam));
        }
        break;

    case WM_SYSCOMMAND:
    {
        WPARAM command = wParam & 0xFFF0;
        if (command == SC_MAXIMIZE || command == SC_RESTORE ||
            command == SC_MINIMIZE || command == SC_MOVE || command == SC_SIZE)
        {
            const wchar_t* name = command == SC_MAXIMIZE ? L"SC_MAXIMIZE"
                                : command == SC_RESTORE  ? L"SC_RESTORE"
                                : command == SC_MINIMIZE ? L"SC_MINIMIZE"
                                : command == SC_MOVE     ? L"SC_MOVE"
                                                         : L"SC_SIZE";
            LogWrite(L"WM_SYSCOMMAND  hwnd=0x%p  %s", (void*)hwnd, name);
        }
        break;
    }

    case WM_SHOWWINDOW:
        // Frame lifetime is not document lifetime, and this is where that shows up. Closing one of
        // two documents was measured to hide one frame (title reset to "Word", window still alive)
        // and destroy a different one - Word moves documents between frames rather than pairing
        // them. It also creates frames it never shows. So a tab strip must follow shown/hidden,
        // not creation and destruction, or it will show tabs for documents that do not exist and
        // miss ones that do.
        LogWrite(L"WM_SHOWWINDOW  hwnd=0x%p  %s", (void*)hwnd, wParam ? L"shown" : L"hidden");
        break;

    case WM_DPICHANGED:
        // This rig runs at 150%. Every rectangle we compute has to survive a monitor change.
        LogWrite(L"WM_DPICHANGED  hwnd=0x%p  dpi=%d", (void*)hwnd, (int)LOWORD(wParam));
        StripOnFrameDpiChanged(hwnd);
        break;

    case WM_NCDESTROY:
        // Last message a window ever gets. Detaching here is not optional: leaving our procedure
        // on a dead window, or in the chain after the DLL unloads, is a crash in Word.
        LogWrite(L"WM_NCDESTROY  hwnd=0x%p  (frame closing)", (void*)hwnd);
        DetachFrame(hwnd, L"destroyed");
        break;

    default:
        break;
    }

    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

// ---------------------------------------------------------------------------------------------
// Attach / detach.
// ---------------------------------------------------------------------------------------------

static BOOL IsTracked(HWND hwnd)
{
    for (int i = 0; i < g_frameCount; i++)
        if (g_frames[i] == hwnd)
            return TRUE;
    return FALSE;
}

static void PinModule(void)
{
    // Once our procedure is in a window's subclass chain, this DLL must never leave the process.
    // COM's reference counting knows nothing about window procedures, so pinning is the only
    // honest way to make that true - and it is deliberately permanent, because Word may hold a
    // return address into us on a stack long after the last window is detached.
    static LONG pinned = 0;
    if (InterlockedExchange(&pinned, 1) != 0)
        return;

    HMODULE ignored = NULL;
    BOOL ok = GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_PIN |
                                 GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS,
                                 (LPCWSTR)(void*)&FrameSubclassProc, &ignored);
    LogWrite(L"module pinned in WINWORD: %s", ok ? L"yes" : L"FAILED");
}

static void AttachFrame(HWND hwnd, const wchar_t* why)
{
    if (!hwnd || !IsWindow(hwnd))
        return;

    DWORD pid = 0;
    DWORD thread = GetWindowThreadProcessId(hwnd, &pid);
    if (pid != GetCurrentProcessId())
        return;

    wchar_t cls[64] = L"";
    GetClassNameW(hwnd, cls, 64);
    if (_wcsicmp(cls, kFrameClass) != 0)
        return;

    EnsureLock();
    EnterCriticalSection(&g_lock);

    if (IsTracked(hwnd) || g_frameCount >= MAX_FRAMES)
    {
        LeaveCriticalSection(&g_lock);
        return;
    }

    BOOL attached = SetWindowSubclass(hwnd, FrameSubclassProc, kSubclassId, 0);
    if (attached)
        g_frames[g_frameCount++] = hwnd;
    int total = g_frameCount;

    LeaveCriticalSection(&g_lock);

    if (!attached)
    {
        LogWrite(L"attach FAILED  hwnd=0x%p  (%s)  lastError=%lu", (void*)hwnd, why, GetLastError());
        return;
    }

    PinModule();

    RECT rect = { 0, 0, 0, 0 };
    GetWindowRect(hwnd, &rect);

    // The thread comparison is the interesting part: if Word ever puts a frame on a second UI
    // thread, everything that assumes one message loop has to be revisited.
    LogWrite(L"attach  hwnd=0x%p  (%s)  thread=%lu%s  visible=%d  rect=(%ld,%ld %ldx%ld)  frames=%d",
             (void*)hwnd, why, thread,
             thread == g_uiThread ? L" (ours)" : L" (DIFFERENT THREAD)",
             (int)IsWindowVisible(hwnd),
             rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top,
             total);

    // Give the frame its strip. Done after the log line above so the two read in the order they
    // happened, and after the subclass so a strip can never exist on a frame we are not watching.
    StripAttachFrame(hwnd);
}

static void DetachFrame(HWND hwnd, const wchar_t* why)
{
    // First, because it puts Word's layout back and destroys our child window, and both need the
    // frame to still be in a state where its children can be touched.
    StripDetachFrame(hwnd);

    EnsureLock();
    EnterCriticalSection(&g_lock);

    int found = -1;
    for (int i = 0; i < g_frameCount; i++)
    {
        if (g_frames[i] == hwnd) { found = i; break; }
    }
    if (found >= 0)
    {
        for (int i = found; i < g_frameCount - 1; i++)
            g_frames[i] = g_frames[i + 1];
        g_frameCount--;
    }
    int remaining = g_frameCount;

    LeaveCriticalSection(&g_lock);

    if (found < 0)
        return;

    // RemoveWindowSubclass unlinks us wherever we sit in the chain, so an add-in that subclassed
    // after us keeps working. This is the whole reason for using comctl32's subclassing rather
    // than swapping GWLP_WNDPROC by hand, where removing out of order corrupts the chain.
    BOOL removed = IsWindow(hwnd) ? RemoveWindowSubclass(hwnd, FrameSubclassProc, kSubclassId) : TRUE;

    LogWrite(L"detach  hwnd=0x%p  (%s)  removed=%d  frames=%d",
             (void*)hwnd, why, (int)removed, remaining);
}

// ---------------------------------------------------------------------------------------------
// Discovery of frames created later.
//
// Word makes a new top-level OpusApp window for every document window, so subclassing whatever
// exists at startup covers only the first one. A thread-local WH_CBT hook sees each new window on
// Word's UI thread; the class check is cheap and is reached only for top-level windows.
//
// The hook does not subclass directly. At HCBT_CREATEWND the window exists but has not yet had
// WM_NCCREATE, and CreateWindowEx may still fail and destroy it. So the hook posts, and the
// coordinator picks the window up once it is real - by which time the post has been pumped, which
// can only happen after creation returned.
// ---------------------------------------------------------------------------------------------

static LRESULT CALLBACK CbtProc(int code, WPARAM wParam, LPARAM lParam)
{
    if (code == HCBT_CREATEWND && lParam && g_coordinator)
    {
        CBT_CREATEWNDW* create = (CBT_CREATEWNDW*)lParam;
        if (create->lpcs && create->lpcs->hwndParent == NULL &&
            (create->lpcs->style & WS_CHILD) == 0)
        {
            HWND hwnd = (HWND)wParam;
            wchar_t cls[64] = L"";
            if (GetClassNameW(hwnd, cls, 64) > 0 && _wcsicmp(cls, kFrameClass) == 0)
                PostMessageW(g_coordinator, WM_WORDTAB_FRAME_CREATED, (WPARAM)hwnd, 0);
        }
    }

    // A negative code must be passed on untouched, per the hook contract.
    return CallNextHookEx(NULL, code, wParam, lParam);
}

static LRESULT CALLBACK CoordinatorProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    if (msg == WM_WORDTAB_FRAME_CREATED)
    {
        AttachFrame((HWND)wParam, L"new frame");
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

// ---------------------------------------------------------------------------------------------
// Start / stop.
// ---------------------------------------------------------------------------------------------

void FramesStart(void)
{
    if (g_started)
        return;
    g_started = TRUE;

    LARGE_INTEGER freq;
    g_qpcFreq = QueryPerformanceFrequency(&freq) ? freq.QuadPart : 0;

    g_uiThread   = GetCurrentThreadId();
    g_followDrag = WordTabReadFlag(L"FollowDrag", TRUE);

    // Before any frame is attached: AttachFrame hands each one to the strip code, which has to be
    // ready to receive it.
    StripStart();

    // A message-only window: no pixels, no taskbar, no z-order. It exists to give the CBT hook
    // somewhere to post to, and it is where the coordinator's timers and state will live later.
    WNDCLASSEXW wc;
    memset(&wc, 0, sizeof(wc));
    wc.cbSize        = sizeof(wc);
    wc.lpfnWndProc   = CoordinatorProc;
    wc.hInstance     = g_module;
    wc.lpszClassName = kCoordinatorClass;
    g_coordClass = RegisterClassExW(&wc);

    if (g_coordClass)
    {
        g_coordinator = CreateWindowExW(0, kCoordinatorClass, L"WordTab", 0, 0, 0, 0, 0,
                                        HWND_MESSAGE, NULL, g_module, NULL);
    }

    if (!g_coordinator)
    {
        LogWrite(L"FramesStart: coordinator window FAILED (lastError=%lu) - new frames will not "
                 L"be picked up", GetLastError());
    }

    // hMod must be NULL for a thread-local hook whose procedure lives in this process - passing a
    // module handle here is the documented way to have SetWindowsHookEx quietly refuse.
    g_cbtHook = SetWindowsHookExW(WH_CBT, CbtProc, NULL, g_uiThread);

    // While we hold window subclasses and a hook, DllCanUnloadNow must say no. Balanced in
    // FramesStop; kept separate from the pin above, which never comes back.
    InterlockedIncrement(&g_lockCount);

    LogWrite(L"FramesStart  uiThread=%lu  coordinator=0x%p  cbtHook=%s  followDrag=%d  qpc=%lldHz",
             g_uiThread, (void*)g_coordinator,
             g_cbtHook ? L"installed" : L"FAILED", (int)g_followDrag, g_qpcFreq);

    // Whatever already exists. At OnStartupComplete the first frame is created but not yet
    // visible - measured, every run - so this must not filter on IsWindowVisible.
    HWND window = NULL;
    int seen = 0;
    while ((window = FindWindowExW(NULL, window, kFrameClass, NULL)) != NULL)
    {
        DWORD pid = 0;
        GetWindowThreadProcessId(window, &pid);
        if (pid != GetCurrentProcessId())
            continue;
        seen++;
        AttachFrame(window, L"existing");
    }

    LogWrite(L"FramesStart  done: %d existing frame(s) seen, %d attached", seen, g_frameCount);
}

void FramesStop(void)
{
    if (!g_started)
        return;
    g_started = FALSE;

    DWORD thread = GetCurrentThreadId();
    LogWrite(L"FramesStop  thread=%lu%s  frames=%d", thread,
             thread == g_uiThread ? L" (ui thread)" : L" (NOT the ui thread)", g_frameCount);

    if (g_cbtHook)
    {
        UnhookWindowsHookEx(g_cbtHook);
        g_cbtHook = NULL;
    }

    // Puts Word's layout back on every frame while the frames are all still alive. After this the
    // per-frame DetachFrame calls below find nothing left to restore, which is what we want: by
    // then Word may already be tearing windows down.
    StripStop();

    // Detach back-to-front: DetachFrame compacts the table as it goes.
    for (;;)
    {
        EnsureLock();
        EnterCriticalSection(&g_lock);
        HWND hwnd = g_frameCount > 0 ? g_frames[g_frameCount - 1] : NULL;
        LeaveCriticalSection(&g_lock);

        if (!hwnd)
            break;
        DetachFrame(hwnd, L"shutdown");
    }

    if (g_coordinator)
    {
        DestroyWindow(g_coordinator);
        g_coordinator = NULL;
    }
    if (g_coordClass)
    {
        UnregisterClassW(kCoordinatorClass, g_module);
        g_coordClass = 0;
    }

    // Balances the increment in FramesStart. The module pin is not undone - see PinModule.
    InterlockedDecrement(&g_lockCount);

    LogWrite(L"FramesStop  done  (lockCount=%ld)", g_lockCount);
}
