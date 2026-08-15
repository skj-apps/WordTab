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
void StripStop(void);

// What the stack needs from the strip. All of these exist because Word lays out only the focused
// window: the stack has to take that window's document-frame rect and hand it to the others.
BOOL StripHasDocumentFrame(HWND frame);
BOOL StripGetNatural(HWND frame, RECT* natural);
void StripSetNatural(HWND frame, const RECT* natural);
void StripRefit(HWND frame);
void StripRefreshTabs(void);

// The name to put on a tab: the frame's title with Word's " - Word" suffix removed.
void WordTabFrameTitle(HWND frame, wchar_t* out, int chars);

// ---------------------------------------------------------------------------------------------
// The stack - several Word windows held at one rectangle so they read as one window with tabs.
// See stack.cpp. Membership is decided by what is true now (visible, has a document frame), never
// by window creation and destruction: frame lifetime is not document lifetime in Word.
// ---------------------------------------------------------------------------------------------

void StackStart(void);
void StackAttachFrame(HWND frame);
void StackDetachFrame(HWND frame);
void StackOnFrameActivate(HWND frame);
void StackJanitor(void);
void StackStop(void);

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

// Read a DWORD switch from HKCU\Software\WordTab as a boolean. Absent means the default, so a
// fresh install behaves like a configured one. Used for the switches that must be flippable
// without a rebuild.
BOOL WordTabReadFlag(const wchar_t* name, BOOL defaultValue);
