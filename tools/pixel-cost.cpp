// What the chrome sampler's three calls actually cost on this machine, off the add-in entirely.
//
// Written because tools\check-governor.ps1's argument applies here too: the expensive thing is not
// reachable from a suite that drives Word, and the number that matters is a property of THIS rig's
// compositor rather than of WordTab. SampleChrome (strip.cpp) does one GetDC(NULL), then per strip
// twenty-four WindowFromPoint + GetPixel pairs against that screen DC, every fourth janitor tick.
//
// Measured 2026-09-04, immediately after QPC brackets went round SampleChrome and reported a mean of
// 512,856-743,963us per sample:
//
//     round 0:  GetDC(NULL) 34 us   24x GetPixel 929092 us   24x WindowFromPoint 87 us
//     round 1:  GetDC(NULL)  9 us   24x GetPixel 968250 us   24x WindowFromPoint 79 us
//     round 2:  GetDC(NULL) 10 us   24x GetPixel 1284716 us  24x WindowFromPoint 1336 us
//
// GetDC is free and WindowFromPoint is free. **GetPixel against a screen DC is about 40ms a call**,
// because each one forces a readback out of a DWM-composited surface. That is the whole finding, and
// it says where the fix goes: one BitBlt of the sample band into a memory DC is one readback instead
// of twenty-four.
//
// Build and run:
//   %LOCALAPPDATA%\Programs\w64devkit\w64devkitin\g++.exe -O2 -o pixel-cost.exe ^
//       tools\pixel-cost.cpp -lgdi32 -luser32

#include <windows.h>
#include <stdio.h>
static long long us(LARGE_INTEGER a, LARGE_INTEGER b, LARGE_INTEGER f)
{ return ((b.QuadPart - a.QuadPart) * 1000000) / f.QuadPart; }
int main(void)
{
    LARGE_INTEGER f, t0, t1; QueryPerformanceFrequency(&f);
    POINT pt; GetCursorPos(&pt);
    for (int round = 0; round < 3; round++)
    {
        QueryPerformanceCounter(&t0);
        HDC screen = GetDC(NULL);
        QueryPerformanceCounter(&t1);
        long long dcUs = us(t0, t1, f);

        QueryPerformanceCounter(&t0);
        for (int i = 0; i < 24; i++) { volatile COLORREF c = GetPixel(screen, 100 + i * 7, 200); (void)c; }
        QueryPerformanceCounter(&t1);
        long long pixUs = us(t0, t1, f);

        QueryPerformanceCounter(&t0);
        for (int i = 0; i < 24; i++) { volatile HWND w = WindowFromPoint(pt); (void)w; }
        QueryPerformanceCounter(&t1);
        long long wfpUs = us(t0, t1, f);

        ReleaseDC(NULL, screen);
        printf("round %d:  GetDC(NULL) %lld us   24x GetPixel %lld us   24x WindowFromPoint %lld us\n",
               round, dcUs, pixUs, wfpUs);
    }
    return 0;
}
