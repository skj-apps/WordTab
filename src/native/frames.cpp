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

// Posted to the coordinator when the keyboard hook swallows Ctrl+Tab.
//   wParam = the frame the keystroke went to, lParam = +1 for the tab to the right, -1 for the left
#define WM_WORDTAB_SWITCH_TAB     (WM_APP + 2)

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
static HHOOK  g_msgHook     = NULL;
static BOOL   g_keysEnabled = FALSE;
static HWND   g_coordinator = NULL;
static ATOM   g_coordClass  = 0;

static HWND   g_frames[MAX_FRAMES];
static int    g_frameCount  = 0;

// How much of the drag the stack kept up with. Declared here because the move trace reports on it
// and the trace comes first; fed by the return value of StackOnFramePosChanging.
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
             L"WM_WINDOWPOSCHANGING",
             g_followMax, g_followCount, updates);

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
// Declared in wordtab.h and shared with strip.cpp: every switch WordTab has is a DWORD under the
// same key, and one reader for all of them is one place for the "absent means default" rule.
// The numeric form, and the only one that touches the registry. Every switch WordTab has lives under
// this one key, so the key and the absent-means-default rule are written once here rather than once
// per return type - two entry points is the point (see the header), two reads would not be.
DWORD WordTabReadNumber(const wchar_t* name, DWORD defaultValue)
{
    DWORD value = 0;
    DWORD size = sizeof(value);
    if (RegGetValueW(HKEY_CURRENT_USER, L"Software\\WordTab", name,
                     RRF_RT_REG_DWORD, NULL, &value, &size) != ERROR_SUCCESS)
    {
        return defaultValue;
    }
    return value;
}

BOOL WordTabReadFlag(const wchar_t* name, BOOL defaultValue)
{
    // A value that is present and 0 is OFF even when the default is on - which is the whole point of
    // every TabSomething=0 escape hatch - so the default is passed down as the number to return when
    // the value is ABSENT, and the != 0 is applied to whatever comes back either way.
    return WordTabReadNumber(name, defaultValue ? 1u : 0u) != 0;
}

// The write half. Same key, same DWORD shape, so anything WordTab records for itself can be
// read back by WordTabReadNumber and shown by settings.ps1 and regedit like every other value
// here. RegSetKeyValue creates the key if it is missing, which is the case on a machine where
// the installer has not run since the key was last cleaned out.
void WordTabWriteNumber(const wchar_t* name, DWORD value)
{
    LSTATUS status = RegSetKeyValueW(HKEY_CURRENT_USER, L"Software\\WordTab", name,
                                     REG_DWORD, &value, sizeof(value));
    if (status != ERROR_SUCCESS)
        LogWrite(L"settings: could not write %s (error %ld)", name, (long)status);
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

        // The user has settled on a size. If this is the window the row is measured from,
        // that is the size the row should come back as next time Word starts.
        StackRememberRowSize(hwnd);
        break;

    case WM_WINDOWPOSCHANGING:
        // The hook point that matters. We are handed the *proposed* position, before the move
        // happens, and may read or change it. This is where the stack keeps every window together:
        // the followers are moved to where this window is *about* to be, in this same message, so
        // they never lag behind it by even one frame.
        if (!StackIsSyncing())
        {
            const WINDOWPOS* pos = (const WINDOWPOS*)lParam;
            TraceRecord(msg, pos);

            int moved = StackOnFramePosChanging(hwnd, pos);
            if (moved > 0)
            {
                g_followCount++;
                if (moved > g_followMax)
                    g_followMax = moved;
            }
        }
        break;

    case WM_WINDOWPOSCHANGED:
        if (!StackIsSyncing())
            TraceRecord(msg, (const WINDOWPOS*)lParam);
        break;

    case WM_ACTIVATE:
        // Which window is on top is which tab is selected, and the active window is also the only
        // one Word will lay out - so this is where the stack's layout oracle changes hands.
        if (LOWORD(wParam) != WA_INACTIVE)
            StackOnFrameActivate(hwnd);
        break;

    case WM_ENABLE:
        // A modal dialog disables the window that owns it, and EnableWindow sends this. So this is
        // Word putting a question to the user about this document, and taking it away again -
        // delivered as an event rather than something to be noticed by looking.
        //
        // That distinction is the whole reason this case exists. The batch close first tried to
        // spot the save prompt by testing IsWindowEnabled on the janitor's half-second tick, and a
        // prompt that came and went inside one tick was invisible to it - which a script does
        // routinely and an impatient user will do eventually. A message cannot be missed.
        StackOnFrameEnable(hwnd, wParam ? TRUE : FALSE);
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
        // The stack is one window to the user, so it goes down to the taskbar and comes back as
        // one. This is where that is noticed.
        StackOnFrameSize(hwnd, wParam);
        break;

    case WM_SYSCOMMAND:
    {
        WPARAM command = wParam & 0xFFF0;

        // The window's own close - the x in the title bar, Alt+F4, the window menu. The stack is
        // one window to the user, so closing it closes every document in it.
        //
        // SC_CLOSE and not WM_CLOSE, and the difference is the whole of the safety. WM_CLOSE is what
        // the add-in itself posts to close one tab, and what the batch posts to close them in turn;
        // intercepting that would be the stack answering its own question and closing everything
        // over and over. SC_CLOSE is only ever the user pressing close on this window.
        //
        // Returning without chaining, which this file does for exactly two other messages and under
        // the same rule: only when the answer is provably ours. StackCloseWindowCommand says FALSE
        // for anything that is not a stack of several documents, and then this falls through to
        // Word untouched.
        if (command == SC_CLOSE)
        {
            LogWrite(L"WM_SYSCOMMAND  hwnd=0x%p  SC_CLOSE", (void*)hwnd);
            if (StackCloseWindowCommand(hwnd))
                return 0;
            break;
        }

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
        //
        // **Chained first, then asked.** WM_SHOWWINDOW is sent *before* the window's visibility
        // actually changes, so a janitor run from here reads the state the window is leaving rather
        // than the one it is arriving at - and on a show that means reading "hidden" about a window
        // that is being shown. Measured, tools\check-row.ps1: a window hidden and re-shown a quarter
        // of a second later produced two "this window is not visible" readings a few milliseconds
        // apart, one from each edge of the gesture, and the second one evicted it from the row.
        LogWrite(L"WM_SHOWWINDOW  hwnd=0x%p  %s", (void*)hwnd, wParam ? L"shown" : L"hidden");
        {
            LRESULT chained = DefSubclassProc(hwnd, msg, wParam, lParam);
            StackJanitor();      // membership is decided by what is visible, so re-decide it now
            return chained;
        }

    case WM_DPICHANGED:
        // This rig runs at 150%. Every rectangle we compute has to survive a monitor change.
        LogWrite(L"WM_DPICHANGED  hwnd=0x%p  dpi=%d", (void*)hwnd, (int)LOWORD(wParam));
        StripOnFrameDpiChanged(hwnd);
        break;

    // The tab context menu is owner-drawn, and an owner-drawn menu asks the window that owns it to
    // measure and paint each item. Its owner is this frame - it has to be; see StripOnMenuMeasure in
    // wordtab.h - so the requests arrive here.
    //
    // These are the first two cases in this file that return without chaining, and the condition is
    // the whole of their safety. Word owner-draws its own menus and controls on this same window,
    // and any of them would arrive here too; answering for one of those would stop it drawing. So
    // the strip answers TRUE only for items it can prove are its own, and everything else falls
    // through the `break` to DefSubclassProc exactly as before.
    //
    // Nothing is logged from either: WM_DRAWITEM fires per item, and again per item every time the
    // highlight moves, and a log line here is a file write inside a modal loop.
    case WM_MEASUREITEM:
        if (StripOnMenuMeasure(hwnd, (MEASUREITEMSTRUCT*)lParam))
            return TRUE;
        break;

    case WM_DRAWITEM:
        if (StripOnMenuDraw(hwnd, (DRAWITEMSTRUCT*)lParam))
            return TRUE;
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

    // Give the frame its strip, then offer it to the stack. Strip first: the stack decides
    // membership by asking whether the window has a document frame, which is the strip's business.
    StripAttachFrame(hwnd);
    StackAttachFrame(hwnd);
}

static void DetachFrame(HWND hwnd, const wchar_t* why)
{
    // Stack first, so the window is out of the tab row before its strip is destroyed. Both need the
    // frame to still be in a state where its children can be touched, which is why this happens
    // here rather than after the subclass is removed.
    StackDetachFrame(hwnd);
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
            {
                // **Where it is born, decided here, because everything else happens too late.**
                //
                // Opening an existing document showed its window a few inches above the stack before
                // it settled into it - the user called it minor, and its cause is that this hook
                // posts. By the time the coordinator picks the window up, Word has created it at the
                // position it chose and shown it there; the stack can only move it afterwards, and
                // "afterwards" is a frame the user can see.
                //
                // `lpcs` is the CREATESTRUCT Word passed to CreateWindowEx and it is writable: this
                // is the documented purpose of HCBT_CREATEWND. Writing the stack's own rectangle
                // into it means the window is created where it belongs, and there is no first
                // position to flash from. Word may still lay it out differently afterwards - a
                // maximized stack maximizes it - and that is fine, because none of those is a move
                // from somewhere else on screen.
                RECT where;
                if (StackProposeCreateRect(&where))
                {
                    create->lpcs->x  = where.left;
                    create->lpcs->y  = where.top;
                    create->lpcs->cx = where.right - where.left;
                    create->lpcs->cy = where.bottom - where.top;
                }

                PostMessageW(g_coordinator, WM_WORDTAB_FRAME_CREATED, (WPARAM)hwnd, 0);
            }
        }
    }

    // A negative code must be passed on untouched, per the hook contract.
    return CallNextHookEx(NULL, code, wParam, lParam);
}

// ---------------------------------------------------------------------------------------------
// Ctrl+Tab between documents.
//
// Word gives the keyboard focus to _WwG, the document pane, which this add-in does not subclass -
// and a WM_KEYDOWN goes to the focus window and nowhere else. Unlike WM_CONTEXTMENU it does not
// travel up to parents, so no number of extra window subclasses could ever see a keystroke and a
// thread hook is the only route. Measured with GetGUIThreadInfo; RESULT-keyboard.md §1.
//
// WH_GETMESSAGE rather than WH_KEYBOARD, for one reason: here the MSG belongs to us for the length
// of the call and rewriting it to WM_NULL is a documented, total swallow. That is the same move the
// strip makes on WM_WINDOWPOSCHANGING - change what is being proposed rather than correct what
// already happened - and it is the only reliable one, because it runs *before* the message loop's
// TranslateMessage. Swallowing later would leave the synthesised WM_CHAR of 0x09 behind and Word
// would type a tab into the document with no keydown to explain it.
//
// The hook computes nothing and activates nothing. It posts, exactly as CbtProc does and for the
// same reason: this runs inside GetMessage, and SetForegroundWindow from in there re-enters message
// retrieval on the thread that is retrieving. By the time the post is pumped the loop is back at a
// place where a window switch is an ordinary thing to do. What is posted is the *source and the
// direction*, never the target - slice 4 lost a bug to a posted command whose target had gone by the
// time it arrived, so the tab to move to is worked out at the moment it is used.
// ---------------------------------------------------------------------------------------------

static LRESULT CALLBACK GetMsgProc(int code, WPARAM wParam, LPARAM lParam)
{
    // PM_NOREMOVE is somebody peeking: the message stays in the queue and will be back. Rewriting it
    // now would be discarded and acting on it would switch tabs twice, once on the look and once on
    // the read.
    if (code == HC_ACTION && wParam == PM_REMOVE && lParam)
    {
        MSG* msg = (MSG*)lParam;

        // GetKeyState and not GetAsyncKeyState: this is the modifier state as of the last message
        // this thread took off its queue, which is the right question - Ctrl goes down first, so its
        // own WM_KEYDOWN has already been retrieved and dispatched by the time Tab arrives. The
        // physical-state call would answer about the instant the hook happened to run instead.
        if (msg->message == WM_KEYDOWN && msg->wParam == VK_TAB &&
            (GetKeyState(VK_CONTROL) & 0x8000) != 0)
        {
            HWND frame = GetAncestor(msg->hwnd, GA_ROOT);

            // The test is "did this key go to a window that is a tab in our row", and it decides the
            // swallow on its own - not "is there somewhere to go". A single-document Word swallows
            // Ctrl+Tab and does nothing, rather than typing a tab character that one document later
            // it would not have typed. A chord whose meaning depends on how many documents are open
            // is worse than one that is sometimes a no-op.
            //
            // It also falls out correctly everywhere it needs to: focus inside a dialog roots at the
            // dialog and not at an OpusApp, so Ctrl+Tab still moves between the pages of a property
            // sheet; a Word window with no document has left the stack; and Stack=0 means there is no
            // row, so the key is Word's again.
            if (StackTabIndex(frame) >= 0)
            {
                // Bit 30 is the previous key state: set means this is an auto-repeat. Repeats are
                // swallowed but not acted on, and that is deliberate rather than an oversight. A
                // switch here is a real window activation and a relayout of every window in the
                // stack; at the 30-a-second repeat rate a held chord would thrash Word, and a row
                // small enough to fit on screen is one nobody needs to hold a key to cross. Tapping
                // Tab with Ctrl held still works - each tap is a fresh keydown.
                if ((msg->lParam & (1 << 30)) == 0 && g_coordinator)
                {
                    LPARAM delta = (GetKeyState(VK_SHIFT) & 0x8000) ? -1 : +1;
                    PostMessageW(g_coordinator, WM_WORDTAB_SWITCH_TAB, (WPARAM)frame, delta);
                }

                msg->message = WM_NULL;
                msg->wParam  = 0;
                msg->lParam  = 0;
            }
        }
    }

    return CallNextHookEx(NULL, code, wParam, lParam);
}

static LRESULT CALLBACK CoordinatorProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    if (msg == WM_WORDTAB_FRAME_CREATED)
    {
        AttachFrame((HWND)wParam, L"new frame");
        return 0;
    }

    if (msg == WM_WORDTAB_SWITCH_TAB)
    {
        HWND source = (HWND)wParam;
        int  delta  = (int)(LONG_PTR)lParam;
        HWND target = StackNeighbourTab(source, delta);

        // Both outcomes are logged, including the one where nothing happens. "The hook never fired"
        // and "the hook fired and there was nowhere to go" are different facts and they must not
        // share a silence - that is the same rule the modified-flag reader is built on.
        if (target)
        {
            LogWrite(L"keys  ctrl%s+tab  0x%p tab %d -> 0x%p tab %d",
                     delta < 0 ? L"+shift" : L"", (void*)source, StackTabIndex(source),
                     (void*)target, StackTabIndex(target));
            StackActivate(target);
        }
        else
        {
            LogWrite(L"keys  ctrl%s+tab  0x%p tab %d -> nowhere (swallowed; no other tab)",
                     delta < 0 ? L"+shift" : L"", (void*)source, StackTabIndex(source));
        }
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

    g_uiThread = GetCurrentThreadId();

    // Before any frame is attached: AttachFrame hands each one to both of these, which have to be
    // ready to receive it.
    StripStart();
    StackStart();

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

    // Ctrl+Tab / Ctrl+Shift+Tab between documents. Off, the hook is not installed at all rather than
    // installed and inert - the same shape as TabDot=0, where the poll returns before it asks. This
    // is a hook on every message Word retrieves, so "off" has to mean it is out of the way.
    //
    // And unlike most switches in this project, 0 here is a real answer rather than a way of putting
    // a bug back. Word owns Ctrl+Tab, and inside a table it is the only way to type a literal tab
    // into a cell - measured, RESULT-keyboard.md §4. Everywhere else plain Tab already does the same
    // thing, so the cost of taking the chord is confined to that one place; a user who works in
    // tables all day is the person this switch is for.
    g_keysEnabled = WordTabReadFlag(L"TabKeys", TRUE);
    if (g_keysEnabled)
        g_msgHook = SetWindowsHookExW(WH_GETMESSAGE, GetMsgProc, NULL, g_uiThread);

    // While we hold window subclasses and a hook, DllCanUnloadNow must say no. Balanced in
    // FramesStop; kept separate from the pin above, which never comes back.
    InterlockedIncrement(&g_lockCount);

    // Both hooks are reported by what SetWindowsHookEx actually returned, never by the variable that
    // asked for them. A switch that reads back its own default is how TabThemeSample stayed dead for
    // two slices while the log said it was on.
    LogWrite(L"FramesStart  uiThread=%lu  coordinator=0x%p  cbtHook=%s  msgHook=%s  qpc=%lldHz",
             g_uiThread, (void*)g_coordinator,
             g_cbtHook ? L"installed" : L"FAILED",
             !g_keysEnabled ? L"off (HKCU\\Software\\WordTab\\TabKeys=0)"
                            : (g_msgHook ? L"installed" : L"FAILED"),
             g_qpcFreq);

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

    // Every hook installed in FramesStart comes out here. The keyboard one especially: left behind,
    // it would rewrite messages through a callback whose DLL has gone.
    if (g_cbtHook)
    {
        UnhookWindowsHookEx(g_cbtHook);
        g_cbtHook = NULL;
    }
    if (g_msgHook)
    {
        UnhookWindowsHookEx(g_msgHook);
        g_msgHook = NULL;
    }

    // Both put Word back the way it was while every frame is still alive: the stack returns each
    // window to the position it had before it was stacked, then the strip gives back the band it
    // took. Order matters - the strip refits the document frame to whatever size the window ends
    // up at, so the windows have to be moved first.
    StackStop();
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
