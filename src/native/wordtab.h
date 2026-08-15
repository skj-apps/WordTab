// WordTab - shared declarations for the native COM add-in.
//
// Native rather than .NET, and not by preference: a managed COM server cannot be activated from a
// per-user (HKCU) registration, and per-user is all we get on a machine without admin rights.
// See src\WordTab.Connect\RESULT.md for the measurement that forced this.

#pragma once

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
