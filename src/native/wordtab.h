// WordTab - shared declarations for the native COM add-in.
//
// Native rather than .NET, and not by preference: a managed COM server cannot be activated from a
// per-user (HKCU) registration, and per-user is all we get on a machine without admin rights.
// See src\WordTab.Connect\RESULT.md for the measurement that forced this.

#pragma once

// Target Windows 10 and later. Stated rather than left to the toolchain's default, which is old
// enough to hide WM_DPICHANGED and comctl32's window subclassing behind version guards - both of
// which the frame code needs, and neither of which fails loudly if the guard silently excludes it.
#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0A00
#endif
#ifndef WINVER
#define WINVER 0x0A00
#endif

#define WIN32_LEAN_AND_MEAN
#define COBJMACROS
#include <windows.h>
#include <objbase.h>
#include <oleauto.h>

// ---------------------------------------------------------------------------------------------
// Identity. install\install.ps1 parses the CLSID straight out of this file and refuses to run if
// it disagrees with its own copy, so the code and the registration cannot drift apart.
// ---------------------------------------------------------------------------------------------

// {4BF75ED9-10EE-4866-BF4A-3D663A4149A1}
#define WORDTAB_CLSID_STRING L"{4BF75ED9-10EE-4866-BF4A-3D663A4149A1}"
extern const CLSID CLSID_WordTabConnect;

#define WORDTAB_PROGID L"WordTab.Connect"

// ---------------------------------------------------------------------------------------------
// IDTExtensibility2 - the interface Office calls an add-in through.
//
// Declared here rather than pulled from a type library. It is a fixed, published contract that
// has not changed since Office 2000, and declaring it keeps the build free of any dependency on
// Office being installed on the build machine.
//
// It is a *dual* interface, so Word may call it either way: through the vtable after querying for
// IID_IDTExtensibility2, or through IDispatch::Invoke by DISPID. We support both - see
// Connect::Invoke, which routes DISPIDs 1..5 to the very same methods.
// ---------------------------------------------------------------------------------------------

// {B65AD801-ABAF-11D0-BB8B-00A0C90F2744}
extern const IID IID_IDTExtensibility2;

enum ext_ConnectMode
{
    ext_cm_AfterStartup = 0,   // switched on mid-session: no OnStartupComplete is coming
    ext_cm_Startup      = 1,
    ext_cm_External     = 2,
    ext_cm_CommandLine  = 3,
    ext_cm_Solution     = 4,
    ext_cm_UISetup      = 5
};

enum ext_DisconnectMode
{
    ext_dm_HostShutdown     = 0,   // Word is closing
    ext_dm_UserClosed       = 1,   // switched off while Word keeps running - teardown must be real
    ext_dm_UISetupComplete  = 2,
    ext_dm_SolutionClosed   = 3
};

struct IDTExtensibility2 : public IDispatch
{
    virtual HRESULT STDMETHODCALLTYPE OnConnection(
        IDispatch* application, enum ext_ConnectMode connectMode,
        IDispatch* addInInst, SAFEARRAY** custom) = 0;

    virtual HRESULT STDMETHODCALLTYPE OnDisconnection(
        enum ext_DisconnectMode removeMode, SAFEARRAY** custom) = 0;

    virtual HRESULT STDMETHODCALLTYPE OnAddInsUpdate(SAFEARRAY** custom) = 0;
    virtual HRESULT STDMETHODCALLTYPE OnStartupComplete(SAFEARRAY** custom) = 0;
    virtual HRESULT STDMETHODCALLTYPE OnBeginShutdown(SAFEARRAY** custom) = 0;
};

// DISPIDs, fixed by the same contract. Do not renumber.
#define DISPID_OnConnection       1
#define DISPID_OnDisconnection    2
#define DISPID_OnAddInsUpdate     3
#define DISPID_OnStartupComplete  4
#define DISPID_OnBeginShutdown    5

// ---------------------------------------------------------------------------------------------
// Module state. DllCanUnloadNow answers from these, so every object and every LockServer call
// has to be accounted for or Word may unload us out from under a live pointer.
// ---------------------------------------------------------------------------------------------

extern LONG g_objectCount;
extern LONG g_lockCount;
extern HMODULE g_module;

// ---------------------------------------------------------------------------------------------
// Logging. This is how we see inside WINWORD, where there is no console. Every call is
// best-effort and swallows its own errors: a logging failure must never take the add-in down.
// ---------------------------------------------------------------------------------------------

void LogWrite(const wchar_t* format, ...);
const wchar_t* LogFilePath(void);

// The object Word instantiates.
HRESULT WordTabCreateConnect(REFIID riid, void** ppv);

// Ask Word for a new blank document, through the Application object it handed us at OnConnection.
// Returns FALSE if Word declined or there is no Application - it never throws and never blocks.
BOOL WordTabNewDocument(void);

// Which of these frames have a document with unsaved changes, by one pass over Application.Windows.
// `frames` and `modified` are parallel arrays of `count` entries.
//
// TRUE means Word answered and every entry of `modified` was written - including FALSE for a frame
// no Word window claims, which is a real answer rather than a missing one (a Protected View document
// is in none of this Application's collections, and cannot be edited anyway).
//
// FALSE means Word would not answer at all, and `modified` was NOT touched: the caller keeps
// whatever it had. A tick that failed to measure must not be able to pass for a measurement.
BOOL WordTabReadModified(const HWND* frames, int count, BOOL* modified);

// Save the document behind a tab, through Word's own Document.Save - so an unchanged document is
// untouched and one that has never been saved gets Word's Save As dialog, exactly as Ctrl+S would.
//
// `frame` must already be the window Word considers active. This does not activate it: it *checks*
// that Word agrees, and saves nothing if it does not. Saving the wrong document is the one failure
// in this add-in that reaches the user's data, so it is made impossible rather than unlikely.
BOOL WordTabSaveDocument(HWND frame);

// ---------------------------------------------------------------------------------------------
// Frame windows - the in-process subclass of Word's OpusApp frames. See frames.cpp.
//
// Both calls are idempotent and must both happen on Word's UI thread. FramesStop must run before
// Word tears its windows down, or our window procedure is left in a chain we no longer control.
// ---------------------------------------------------------------------------------------------

void FramesStart(void);
void FramesStop(void);

// ---------------------------------------------------------------------------------------------
// The strip - our band of Word's layout, carved out of the top of the `_WwF` document frame.
// See strip.cpp. Driven entirely by frames.cpp: there is one strip per subclassed frame, and it
// lives and dies with it.
// ---------------------------------------------------------------------------------------------

void StripStart(void);
void StripAttachFrame(HWND frame);
void StripDetachFrame(HWND frame);
void StripOnFrameDpiChanged(HWND frame);

// The tab context menu is owner-drawn, so that it is dark on a dark Word instead of being a white
// rectangle the system chose. An owner-drawn menu sends its measure and draw requests to the window
// that *owns* it, and that has to be Word's frame rather than the strip: the strip is
// WS_EX_NOACTIVATE and can never be foreground, and a popup menu whose owner is not foreground does
// not dismiss when the user clicks away from it. So they arrive at frames.cpp and come back here.
//
// Both answer TRUE only for our own items, identified by a pointer into our own array rather than by
// item id - the ids are 1 to 5 and would collide with anything. The frame must return TRUE without
// chaining for those, and pass everything else through untouched: Word owner-draws its own menus on
// this window, and so, for all we know, does the other tab add-in that is installed on these
// machines.
BOOL StripOnMenuMeasure(HWND frame, MEASUREITEMSTRUCT* item);
BOOL StripOnMenuDraw(HWND frame, DRAWITEMSTRUCT* item);

void StripStop(void);

// What the stack needs from the strip. All of these exist because Word lays out only the focused
// window: the stack has to take that window's document-frame rect and hand it to the others.

// Is there a document open in this window? This is the membership test - a window with no document
// is not a tab - and it is deliberately *not* the same question as "does it have a `_WwF`". Word
// keeps the document frame for the life of the window and empties it when the last document closes.
BOOL StripHasDocument(HWND frame);
BOOL StripGetNatural(HWND frame, RECT* natural);
// `why` names the path that asked, because this is the one writer of a window's natural rect that
// did not come from Word laying that window out - see the note on the definition.
void StripSetNatural(HWND frame, const RECT* natural, const wchar_t* why);
void StripRefit(HWND frame);
void StripRefreshTabs(void);

// The name to put on a tab: the frame's title with Word's " - Word" suffix removed.
void WordTabFrameTitle(HWND frame, wchar_t* out, int chars);

// ---------------------------------------------------------------------------------------------
// The stack - several Word windows held at one rectangle so they read as one window with tabs.
// See stack.cpp. Membership is decided by what is true now (visible, has a document open in it),
// never by window creation and destruction: frame lifetime is not document lifetime in Word, and
// neither is document-frame lifetime - `_WwF` outlives the document inside it.
// ---------------------------------------------------------------------------------------------

void StackStart(void);
void StackAttachFrame(HWND frame);
void StackDetachFrame(HWND frame);
void StackOnFrameActivate(HWND frame);
void StackOnFrameSize(HWND frame, WPARAM sizeType);

// A frame has been disabled or re-enabled, which is what a modal dialog does to the window that
// owns it. The batch close needs this as an *event*: a save prompt can come and go inside one
// janitor tick, and one that is never seen is one that is never waited for.
void StackOnFrameEnable(HWND frame, BOOL enabled);
void StackJanitor(void);
void StackStop(void);

// Is a batch close - Close Others, Close All, Close Tabs to the Right - part-way through?
//
// Exported so that the janitor's other work can stand off while it runs. A batch is a sequence of
// WM_CLOSEs with Word's save prompt appearing between them, which makes it the one stretch where
// Word is repeatedly in the middle of something it was asked to do by us.
BOOL StackCloseInFlight(void);

// ---------------------------------------------------------------------------------------------
// The taskbar - one button for the whole stack, following the active tab. See taskbar.cpp.
// TaskbarShow(frame, TRUE) must be reachable for every window the add-in ever hid: a window with
// no taskbar button, underneath another window, cannot be reached by the user at all.
// ---------------------------------------------------------------------------------------------

void TaskbarStart(void);
void TaskbarShow(HWND frame, BOOL show);
void TaskbarStop(void);

// Returns how many other windows were moved with this one, so the drag trace can report it.
int  StackOnFramePosChanging(HWND frame, const WINDOWPOS* pos);

// Word laid out the focused window; every other window in the stack is given the same interior.
void StackOnActiveLayout(HWND frame, const RECT* natural);

// TRUE while the stack is moving windows itself, so our own SetWindowPos calls coming back through
// the subclass are not mistaken for Word's.
BOOL StackIsSyncing(void);

// The tab row, and what a click on one does. `out` receives the frames, left to right.
int  StackTabs(HWND frame, HWND* out, int max, int* activeIndex);
void StackActivate(HWND frame);

// Where a tab sits in the row, and how to move it. The index space is the one StackTabs hands out:
// joined windows, left to right, which is exactly what the strip draws. StackTabIndex answers -1 for
// a window that is not a tab in a stack, and StackMoveTab returns TRUE only if the row changed - so
// a drag can call it on every mouse movement and it costs nothing while the tab is already there.
int  StackTabIndex(HWND frame);
BOOL StackMoveTab(HWND frame, int toIndex);

// How many tabs are to the right of this one. 0 both for the last tab and for a window that is not
// a tab in a stack, because to the caller those mean the same thing: nothing to close to the right.
int  StackTabsRightOf(HWND frame);

// Close the document behind a tab. Activates it first - see the comment on the definition, which is
// about where a modal save prompt ends up in the z-order - and returns the user to the tab they were
// on if it was not the one they closed.
void StackCloseTab(HWND frame);

// Close every tab but one, or every tab. Deliberately not a loop over StackCloseTab: any of those
// closes can raise Word's own save prompt, and six prompts stacked on top of each other for
// documents the user cannot see is not a feature. The batch runs one close at a time, stepped by
// the janitor, and a prompt the user cancels abandons the rest of it. See the definitions.
//
// `anyTab` is the frame the command came from, and is only used when there is no stack to enumerate
// - stacking switched off, or a lone window - where "close all" means that one document.
void StackCloseOthers(HWND keep);
void StackCloseAll(HWND anyTab);

// The same queue, narrowed to the tabs after this one in the row. Refused when there are none, so
// the caller does not have to check before asking.
void StackCloseToRight(HWND from);

// Read a DWORD switch from HKCU\Software\WordTab as a boolean. Absent means the default, so a
// fresh install behaves like a configured one. Used for the switches that must be flippable
// without a rebuild.
BOOL WordTabReadFlag(const wchar_t* name, BOOL defaultValue);
