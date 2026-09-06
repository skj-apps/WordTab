// Does an injected mouse move to the SAME point deliver a WM_MOUSEMOVE?
//
// Written 2026-09-05 to settle one question with a number rather than with folklore, because the
// answer decides how Set-Pointer in tools\WordTabHarness.ps1 has to be written.
//
// The harness confirms a hover by asking GetCursorPos whether the pointer is where it was sent. That
// is a true statement about the POINTER and says nothing about whether the target window was told -
// and the strip's hover state is driven by WM_MOUSEMOVE, not by where the cursor happens to be. The
// battery of 2026-09-05 has the failure in its own words: `hovering a modified tab brings the x back`
// failed with the pointer provably on the tab, and the add-in logged eight seconds of tooltip silence
// across the whole step where the passing run logged four show/hide pairs. The strip was never told.
//
// One candidate mechanism is that Windows posts no move message when the position does not change, so
// a re-send to a point the cursor already occupies is silently nothing. That would make the harness's
// confirmation not merely weak but actively misleading: it would report success in exactly the case
// where nothing was delivered.
//
// This measures it directly. No Word, no add-in: one window, one counter in its window procedure, and
// three injections.
//
//   build:  g++ -O2 -municode -mwindows -o move-delivery.exe move-delivery.cpp -luser32
//           (or -mconsole; it prints, so console is easier to read)
//   run:    move-delivery.exe          - takes over the pointer for about a second, then puts it back

#include <windows.h>
#include <stdio.h>
#include <string.h>
#include <stdlib.h>

static int g_moves = 0;

static LRESULT CALLBACK WndProc(HWND h, UINT m, WPARAM w, LPARAM l)
{
    if (m == WM_MOUSEMOVE)
        g_moves++;
    return DefWindowProcW(h, m, w, l);
}

// The harness's own conversion, copied deliberately: measuring a different injection from the one the
// suites use would answer a question nobody asked. See WordLayout.cs Move().
static void Inject(int x, int y)
{
    int vx = GetSystemMetrics(SM_XVIRTUALSCREEN), vy = GetSystemMetrics(SM_YVIRTUALSCREEN);
    int vw = GetSystemMetrics(SM_CXVIRTUALSCREEN), vh = GetSystemMetrics(SM_CYVIRTUALSCREEN);

    INPUT in;
    ZeroMemory(&in, sizeof(in));
    in.type = INPUT_MOUSE;
    in.mi.dx = (int)(((LONGLONG)(x - vx) * 65535) / (vw - 1));
    in.mi.dy = (int)(((LONGLONG)(y - vy) * 65535) / (vh - 1));
    in.mi.dwFlags = MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_VIRTUALDESK;
    SendInput(1, &in, sizeof(in));
}

// Injected input is delivered asynchronously; a count read straight after SendInput would be reading
// before the message arrived and would report zero for every case alike.
static void Settle(int ms)
{
    DWORD until = GetTickCount() + ms;
    MSG msg;
    for (;;)
    {
        while (PeekMessageW(&msg, NULL, 0, 0, PM_REMOVE))
        {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
        if (GetTickCount() >= until)
            break;
        Sleep(5);
    }
}

static int Round(const wchar_t* what, int x, int y)
{
    g_moves = 0;
    Inject(x, y);
    Settle(300);

    POINT at;
    GetCursorPos(&at);
    // %ls, not %s: mingw's wprintf reads %s as a NARROW string and the first run of this printed the
    // numbers correctly beside six blank labels - a small instance of the thing the whole file is
    // about, an instrument that looks like it is reporting and is not.
    wprintf(L"  %-46ls -> %2d WM_MOUSEMOVE   GetCursorPos says (%d,%d)\n", what, g_moves, at.x, at.y);
    return g_moves;
}

// `watch <ms>` - the same window and the same counter, but driven from OUTSIDE instead of injecting
// anything itself. It prints the rectangle it is sitting at and then, after <ms>, the number of
// WM_MOUSEMOVEs it was sent.
//
// This is the pin for Set-Pointer in tools\WordTabHarness.ps1. A harness fix cannot be A/B'd with
// ab-binary.ps1, which swaps the ADD-IN - so the proof has to be a window that counts what it was
// actually told, and the test is: call Set-Pointer twice at the same point and see whether the second
// call delivers anything. Before the fix it delivers nothing and returns $true regardless.
static int Watch(int ms)
{
    WNDCLASSEXW wc;
    ZeroMemory(&wc, sizeof(wc));
    wc.cbSize        = sizeof(wc);
    wc.lpfnWndProc   = WndProc;
    wc.hInstance     = GetModuleHandleW(NULL);
    wc.hCursor       = LoadCursorW(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.lpszClassName = L"WordTabMoveDelivery";
    RegisterClassExW(&wc);

    HWND h = CreateWindowExW(WS_EX_TOPMOST, L"WordTabMoveDelivery", L"move delivery",
                             WS_OVERLAPPEDWINDOW, 200, 200, 500, 400,
                             NULL, NULL, wc.hInstance, NULL);
    if (!h)
        return 1;
    ShowWindow(h, SW_SHOW);
    Settle(400);

    RECT r;
    GetWindowRect(h, &r);
    // Printed BEFORE the wait and flushed, so the caller can read the rectangle, aim at it, and only
    // then wait for the count.
    wprintf(L"RECT %d %d %d %d\n", (int)r.left, (int)r.top, (int)r.right, (int)r.bottom);
    fflush(stdout);

    // Is anything ELSE moving the pointer?
    //
    // This counts every WM_MOUSEMOVE the window is sent, and it cannot tell an injected one from a
    // hand on the mouse. A run taken while somebody is using the machine reports extra moves and the
    // pin passes for the wrong reason - which happened on the first run of this: the control arm,
    // which must report 1, reported 2, and the pointer was drifting through (1371,557) and (1358,543)
    // between calls. An instrument that cannot say "I could not measure this" is worse than no
    // instrument, so it says it.
    g_moves = 0;
    Settle(600);
    wprintf(L"QUIET %d\n", g_moves);
    fflush(stdout);

    g_moves = 0;
    Settle(ms);
    wprintf(L"MOVES %d\n", g_moves);
    fflush(stdout);

    DestroyWindow(h);
    return 0;
}

int main(int argc, char** argv)
{
    SetProcessDPIAware();

    if (argc >= 2 && strcmp(argv[1], "watch") == 0)
        return Watch(argc >= 3 ? atoi(argv[2]) : 3000);

    POINT restore;
    GetCursorPos(&restore);

    WNDCLASSEXW wc;
    ZeroMemory(&wc, sizeof(wc));
    wc.cbSize        = sizeof(wc);
    wc.lpfnWndProc   = WndProc;
    wc.hInstance     = GetModuleHandleW(NULL);
    wc.hCursor       = LoadCursorW(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.lpszClassName = L"WordTabMoveDelivery";
    RegisterClassExW(&wc);

    HWND h = CreateWindowExW(WS_EX_TOPMOST, L"WordTabMoveDelivery", L"move delivery",
                             WS_OVERLAPPEDWINDOW, 200, 200, 500, 400,
                             NULL, NULL, wc.hInstance, NULL);
    if (!h)
    {
        wprintf(L"could not make the window\n");
        return 1;
    }
    ShowWindow(h, SW_SHOW);
    Settle(400);

    const int ax = 400, ay = 380;      // inside the window
    const int bx = 300, by = 300;      // also inside it, somewhere else

    wprintf(L"\nA window of our own, counting WM_MOUSEMOVE in its own window procedure.\n");
    wprintf(L"Each line injects one absolute move and then pumps for 300ms.\n\n");

    Round(L"1. arriving from wherever the pointer was", bx, by);
    Round(L"2. a REAL move, B -> A", ax, ay);
    int same = Round(L"3. the SAME point again, A -> A", ax, ay);
    int again = Round(L"4. and once more, A -> A", ax, ay);
    Round(L"5. nudge away, A -> A+1", ax + 1, ay);
    int back = Round(L"6. and back, A+1 -> A", ax, ay);

    wprintf(L"\n");
    if (same == 0 && again == 0)
        wprintf(L"ANSWER: a re-send to the point the cursor already occupies delivers NOTHING.\n"
                L"        GetCursorPos still says the pointer is there, so a harness that confirms\n"
                L"        with GetCursorPos reports success for a move the window never received.\n"
                L"        A nudge first (%d message(s) on the way back) is what makes it real.\n", back);
    else
        wprintf(L"ANSWER: a re-send to the same point DOES deliver (%d and %d). The no-op theory is\n"
                L"        wrong and the harness's problem is elsewhere - look at what is UNDER the\n"
                L"        pointer, not at whether it moved.\n", same, again);

    DestroyWindow(h);
    SetCursorPos(restore.x, restore.y);
    return 0;
}
