// WordTab - one taskbar button for the whole stack.
//
// N stacked windows give N taskbar buttons, and that is what gives the trick away: the windows
// look like one window until you glance at the taskbar and see three of them. Windows 11 draws the
// difference as a stacked card behind the icon, so it is visible without even hovering.
//
// The fix is `ITaskbarList::DeleteTab`, which tells the shell to stop showing a button for a
// window without touching the window itself. No styles are changed, so there is nothing to undo
// badly, and a window that leaves the stack simply gets its button back. Spike 2 established that
// this works and that it works cross-process; in-process it is the same two calls with less
// ceremony.
//
// The rule this file exists to enforce: **exactly one button, and it follows the active tab.** Not
// "the first window's button" - if the button stayed on one window, minimising and restoring the
// stack from the taskbar would take you to a document you were not looking at.
//
// Everything here is best-effort. A shell that declines to co-operate is a cosmetic problem; it
// must never be able to take Word down with it, so every call checks and nothing is assumed.

#include "wordtab.h"
#include <shobjidl.h>

static ITaskbarList* g_taskbar = NULL;
static BOOL          g_tried   = FALSE;
static BOOL          g_enabled = TRUE;

static ITaskbarList* Taskbar(void)
{
    if (g_tried)
        return g_taskbar;
    g_tried = TRUE;

    if (!g_enabled)
        return NULL;

    // Word's UI thread is an STA and is already initialised for COM - it is an OLE application, so
    // this runs long after CoInitialize. Initialising it ourselves would be wrong and could change
    // the apartment out from under the host.
    HRESULT hr = CoCreateInstance(CLSID_TaskbarList, NULL, CLSCTX_INPROC_SERVER,
                                  IID_ITaskbarList, (void**)&g_taskbar);
    if (FAILED(hr) || !g_taskbar)
    {
        LogWrite(L"taskbar  CoCreateInstance failed (hr=0x%08lX) - leaving the taskbar alone", hr);
        g_taskbar = NULL;
        return NULL;
    }

    hr = g_taskbar->HrInit();
    if (FAILED(hr))
    {
        LogWrite(L"taskbar  HrInit failed (hr=0x%08lX)", hr);
        g_taskbar->Release();
        g_taskbar = NULL;
        return NULL;
    }

    LogWrite(L"taskbar  ITaskbarList ready");
    return g_taskbar;
}

void TaskbarStart(void)
{
    g_enabled = WordTabReadFlag(L"Taskbar", TRUE);
    if (!g_enabled)
        LogWrite(L"TaskbarStart  disabled by HKCU\\Software\\WordTab\\Taskbar=0");
}

void TaskbarShow(HWND frame, BOOL show)
{
    if (!g_enabled || !frame || !IsWindow(frame))
        return;

    ITaskbarList* taskbar = Taskbar();
    if (!taskbar)
        return;

    HRESULT hr = show ? taskbar->AddTab(frame) : taskbar->DeleteTab(frame);
    if (FAILED(hr))
    {
        LogWrite(L"taskbar  %s(0x%p) failed (hr=0x%08lX)",
                 show ? L"AddTab" : L"DeleteTab", (void*)frame, hr);
    }
}

void TaskbarStop(void)
{
    // Callers put every window's button back before this; releasing is all that is left. Note that
    // a window whose button was deleted and never restored is unreachable from the taskbar, which
    // is why restoring is the caller's responsibility and not an afterthought here.
    if (g_taskbar)
    {
        g_taskbar->Release();
        g_taskbar = NULL;
    }
    g_tried = FALSE;
}
