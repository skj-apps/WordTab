// WordTab - the COM object Word instantiates. This is the entry point for everything WordTab
// will ever do inside WINWORD.
//
// It logs every callback, shows a one-shot banner, and hands off to the frame code in frames.cpp,
// which subclasses Word's OpusApp windows. Shrinking _WwF and painting the strip come next and
// hang off the same two points: FramesStart at OnStartupComplete, FramesStop at shutdown.
//
// Rule for every callback below: nothing escapes and nothing throws. An error returned or an
// exception raised across the COM boundary during load makes Word add us to its
// Resiliency\DisabledItems list, which is silent, sticky, and miserable to diagnose later.

#include "wordtab.h"
#include <new>
#include <stdio.h>
#include <string.h>

static LONG g_bannerShown = 0;

// ---------------------------------------------------------------------------------------------
// Late-bound helpers for talking to Word's Application object.
// ---------------------------------------------------------------------------------------------

// Read a property by name off an IDispatch. Returns NULL on any failure - Word declining to
// answer is information, not an error, and callers print "(?)".
// Caller frees with SysFreeString.
static BSTR GetStringProperty(IDispatch* disp, const wchar_t* name)
{
    if (!disp)
        return NULL;

    DISPID dispid = 0;
    LPOLESTR nameCopy = (LPOLESTR)name;
    if (FAILED(disp->GetIDsOfNames(IID_NULL, &nameCopy, 1, LOCALE_USER_DEFAULT, &dispid)))
        return NULL;

    DISPPARAMS noArgs = { NULL, NULL, 0, 0 };
    VARIANT result;
    VariantInit(&result);

    HRESULT hr = disp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                              DISPATCH_PROPERTYGET, &noArgs, &result, NULL, NULL);
    if (FAILED(hr))
    {
        VariantClear(&result);
        return NULL;
    }

    BSTR text = NULL;
    VARIANT asString;
    VariantInit(&asString);
    if (SUCCEEDED(VariantChangeType(&asString, &result, 0, VT_BSTR)))
        text = SysAllocString(asString.bstrVal ? asString.bstrVal : L"");

    VariantClear(&asString);
    VariantClear(&result);
    return text;
}

// Read Documents.Count. Separate from the above because it is a property on a property.
static BSTR GetDocumentCount(IDispatch* app)
{
    if (!app)
        return NULL;

    DISPID dispid = 0;
    LPOLESTR name = (LPOLESTR)L"Documents";
    if (FAILED(app->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid)))
        return NULL;

    DISPPARAMS noArgs = { NULL, NULL, 0, 0 };
    VARIANT documents;
    VariantInit(&documents);
    if (FAILED(app->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                           DISPATCH_PROPERTYGET, &noArgs, &documents, NULL, NULL)))
    {
        VariantClear(&documents);
        return NULL;
    }

    BSTR count = NULL;
    if (documents.vt == VT_DISPATCH && documents.pdispVal)
        count = GetStringProperty(documents.pdispVal, L"Count");

    VariantClear(&documents);
    return count;
}

// ---------------------------------------------------------------------------------------------
// Word's Application object, held for the life of the connection.
//
// WordTab is a window program: it subclasses Word's frames, moves them and paints on them, and the
// object model knows nothing about any of it. Two things are not reachable that way - making a
// document and saving one - so the Application object is kept for those two and nothing else.
// Keeping it in a file static rather than passing it around is deliberate: it is set once, on
// Word's UI thread, and read from one place.
// ---------------------------------------------------------------------------------------------

static IDispatch* g_application = NULL;

static void SetApplication(IDispatch* application)
{
    if (g_application)
    {
        g_application->Release();
        g_application = NULL;
    }
    if (application)
    {
        application->AddRef();
        g_application = application;
    }
}

// Free the strings Word may have put in an EXCEPINFO. Skipping this leaks a BSTR every time a call
// into Word fails, which is exactly when nobody is looking.
static void ClearExceptionInfo(EXCEPINFO* error)
{
    if (!error)
        return;
    if (error->bstrSource)      { SysFreeString(error->bstrSource);      error->bstrSource = NULL; }
    if (error->bstrDescription) { SysFreeString(error->bstrDescription); error->bstrDescription = NULL; }
    if (error->bstrHelpFile)    { SysFreeString(error->bstrHelpFile);    error->bstrHelpFile = NULL; }
}

// Read a property that answers with an object - Application.ActiveWindow, Window.Document. Returns
// NULL on any failure; the caller logs, because only the caller knows what it was asking for.
// Caller Releases.
static IDispatch* GetObjectProperty(IDispatch* disp, const wchar_t* name)
{
    if (!disp)
        return NULL;

    DISPID dispid = 0;
    LPOLESTR nameCopy = (LPOLESTR)name;
    if (FAILED(disp->GetIDsOfNames(IID_NULL, &nameCopy, 1, LOCALE_USER_DEFAULT, &dispid)))
        return NULL;

    DISPPARAMS noArgs = { NULL, NULL, 0, 0 };
    VARIANT result;
    VariantInit(&result);
    EXCEPINFO error;
    memset(&error, 0, sizeof(error));

    HRESULT hr = disp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                              DISPATCH_PROPERTYGET, &noArgs, &result, &error, NULL);
    ClearExceptionInfo(&error);

    IDispatch* object = NULL;
    if (SUCCEEDED(hr) && result.vt == VT_DISPATCH && result.pdispVal)
    {
        object = result.pdispVal;
        object->AddRef();
    }

    VariantClear(&result);
    return object;
}

// Read a numeric property. FALSE means absent or not a number, which for Window.Hwnd below is
// information rather than an error: it is a property an object model may simply not have, and the
// caller has a weaker but still sound answer to fall back on.
static BOOL GetLongProperty(IDispatch* disp, const wchar_t* name, LONG* value)
{
    if (!disp || !value)
        return FALSE;

    DISPID dispid = 0;
    LPOLESTR nameCopy = (LPOLESTR)name;
    if (FAILED(disp->GetIDsOfNames(IID_NULL, &nameCopy, 1, LOCALE_USER_DEFAULT, &dispid)))
        return FALSE;

    DISPPARAMS noArgs = { NULL, NULL, 0, 0 };
    VARIANT result;
    VariantInit(&result);
    EXCEPINFO error;
    memset(&error, 0, sizeof(error));

    HRESULT hr = disp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                              DISPATCH_PROPERTYGET, &noArgs, &result, &error, NULL);
    ClearExceptionInfo(&error);
    if (FAILED(hr))
    {
        VariantClear(&result);
        return FALSE;
    }

    VARIANT asLong;
    VariantInit(&asLong);
    BOOL ok = SUCCEEDED(VariantChangeType(&asLong, &result, 0, VT_I4)) ? TRUE : FALSE;
    if (ok)
        *value = asLong.lVal;

    VariantClear(&asLong);
    VariantClear(&result);
    return ok;
}

// Call a method that takes no arguments, and say what Word said if it objects.
static BOOL CallMethodNoArgs(IDispatch* disp, const wchar_t* name)
{
    if (!disp)
        return FALSE;

    DISPID dispid = 0;
    LPOLESTR nameCopy = (LPOLESTR)name;
    HRESULT hr = disp->GetIDsOfNames(IID_NULL, &nameCopy, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr))
    {
        LogWrite(L"save: no %s method on that object (hr=0x%08lX)", name, (unsigned long)hr);
        return FALSE;
    }

    DISPPARAMS noArgs = { NULL, NULL, 0, 0 };
    VARIANT result;
    VariantInit(&result);
    EXCEPINFO error;
    memset(&error, 0, sizeof(error));

    hr = disp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                      &noArgs, &result, &error, NULL);

    if (hr == DISP_E_EXCEPTION)
    {
        // Expected, routinely: Word raises "Command failed" when the user presses Cancel in the
        // Save As dialog. That is an answer, not a fault, and the only right response is to log it
        // and leave the document exactly as it is.
        LogWrite(L"save: Word raised an error on %s - %s", name,
                 error.bstrDescription ? error.bstrDescription : L"(no description)");
    }
    ClearExceptionInfo(&error);
    VariantClear(&result);

    if (FAILED(hr))
        return FALSE;
    return TRUE;
}

// Save the document behind a tab.
//
// Word's own Document.Save, so everything about saving is Word's: an unchanged document is not
// written, one that has never been saved gets the Save As dialog, and AutoRecover, macros and the
// read-only cases all behave exactly as they do from Ctrl+S. None of it is reimplemented here.
//
// **Through the active window, and only after checking that Word agrees which one that is.** The
// caller activates the tab first - it has to, for the same reason closing one does: a Save As dialog
// belonging to a window underneath another at the same rectangle is invisible, and an invisible
// modal dialog is Word beeping with nothing on screen. Having activated it, Application.ActiveWindow
// *is* the tab, and Window.Hwnd is one property read that turns "should be" into "is". If they
// disagree, nothing is saved: this is the only path in the add-in that touches the user's data, and
// saving the wrong document silently is not a failure mode worth leaving open.
BOOL WordTabSaveDocument(HWND frame)
{
    if (!g_application)
    {
        LogWrite(L"save: no Application object - was OnConnection ever called?");
        return FALSE;
    }
    if (!frame || !IsWindow(frame))
        return FALSE;

    IDispatch* window = GetObjectProperty(g_application, L"ActiveWindow");
    if (!window)
    {
        LogWrite(L"save: Word has no ActiveWindow - nothing saved");
        return FALSE;
    }

    LONG reported = 0;
    BOOL known = GetLongProperty(window, L"Hwnd", &reported);
    if (known && reported != (LONG)(LONG_PTR)frame)
    {
        LogWrite(L"save: Word's active window is 0x%08lX but the tab is 0x%p - not saving",
                 (unsigned long)reported, (void*)frame);
        window->Release();
        return FALSE;
    }

    IDispatch* document = GetObjectProperty(window, L"Document");
    window->Release();

    if (!document)
    {
        LogWrite(L"save: the active window has no Document - nothing saved");
        return FALSE;
    }

    BOOL saved = CallMethodNoArgs(document, L"Save");
    document->Release();

    LogWrite(L"save: hwnd=0x%p  Document.Save %s%s", (void*)frame,
             saved ? L"returned" : L"did not complete",
             known ? L"  (Word confirmed the window handle)"
                   : L"  (Word does not report Window.Hwnd - went on the activation alone)");
    return saved;
}

// Application.Documents.Add(), late-bound like everything else here so the build needs nothing from
// Office.
//
// Word 2013 and later gives every document its own OpusApp frame, so the new document arrives as a
// new top-level window - which the CBT hook sees, the strip binds and the janitor joins to the
// stack. Nothing here has to place it or make it a tab: it becomes one by the same route every
// other document does, which is the whole reason this is three calls rather than a feature.
BOOL WordTabNewDocument(void)
{
    if (!g_application)
    {
        LogWrite(L"new document: no Application object - was OnConnection ever called?");
        return FALSE;
    }

    DISPID documentsId = 0;
    LPOLESTR name = (LPOLESTR)L"Documents";
    HRESULT hr = g_application->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &documentsId);
    if (FAILED(hr))
    {
        LogWrite(L"new document: Application has no Documents (hr=0x%08lX)", (unsigned long)hr);
        return FALSE;
    }

    DISPPARAMS noArgs = { NULL, NULL, 0, 0 };
    VARIANT documents;
    VariantInit(&documents);

    hr = g_application->Invoke(documentsId, IID_NULL, LOCALE_USER_DEFAULT,
                               DISPATCH_PROPERTYGET, &noArgs, &documents, NULL, NULL);
    if (FAILED(hr) || documents.vt != VT_DISPATCH || !documents.pdispVal)
    {
        LogWrite(L"new document: Documents unavailable (hr=0x%08lX vt=%d)",
                 (unsigned long)hr, (int)documents.vt);
        VariantClear(&documents);
        return FALSE;
    }

    DISPID addId = 0;
    name = (LPOLESTR)L"Add";
    hr = documents.pdispVal->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &addId);
    if (SUCCEEDED(hr))
    {
        VARIANT created;
        VariantInit(&created);
        EXCEPINFO error;
        memset(&error, 0, sizeof(error));

        // No arguments at all: a blank document on the Normal template, which is what Ctrl+N does.
        // Every parameter of Documents.Add is optional, and leaving them out is how you say so
        // through IDispatch.
        hr = documents.pdispVal->Invoke(addId, IID_NULL, LOCALE_USER_DEFAULT,
                                        DISPATCH_METHOD, &noArgs, &created, &error, NULL);

        if (hr == DISP_E_EXCEPTION)
        {
            LogWrite(L"new document: Word raised an error - %s",
                     error.bstrDescription ? error.bstrDescription : L"(no description)");
        }
        ClearExceptionInfo(&error);
        VariantClear(&created);
    }

    VariantClear(&documents);

    if (FAILED(hr))
    {
        LogWrite(L"new document: Documents.Add failed (hr=0x%08lX)", (unsigned long)hr);
        return FALSE;
    }

    LogWrite(L"new document: Documents.Add succeeded - the new frame joins the stack by itself");
    return TRUE;
}

// ---------------------------------------------------------------------------------------------
// The load banner - the visible proof that we are inside Word.
//
// Shown on a background thread on purpose. A modal dialog on Word's UI thread during startup
// would block Word until it is dismissed, and an add-in that can hang its host while proving
// itself is not proving much. Switch it off with install.ps1 -NoBanner once it has served its
// purpose; that flips a registry value, so it needs no rebuild.
// ---------------------------------------------------------------------------------------------

static DWORD WINAPI BannerThread(LPVOID parameter)
{
    wchar_t* text = (wchar_t*)parameter;
    LogWrite(L"banner: thread running, calling MessageBoxW");

    int result = MessageBoxW(NULL, text, L"WordTab",
                             MB_OK | MB_ICONINFORMATION | MB_SETFOREGROUND | MB_TOPMOST);

    LogWrite(L"banner: MessageBoxW returned %d (lastError=%lu)", result, GetLastError());
    HeapFree(GetProcessHeap(), 0, text);
    return 0;
}

static BOOL BannerEnabled(void)
{
    DWORD value = 1;
    DWORD size = sizeof(value);
    if (RegGetValueW(HKEY_CURRENT_USER, L"Software\\WordTab", L"ShowLoadBanner",
                     RRF_RT_REG_DWORD, NULL, &value, &size) != ERROR_SUCCESS)
    {
        return TRUE;   // absent means on: a fresh install should announce itself
    }
    return value != 0;
}

static void ShowBannerOnce(enum ext_ConnectMode connectMode)
{
    BOOL enabled = BannerEnabled();
    LogWrite(L"banner: enabled=%d alreadyShown=%d", (int)enabled, (int)g_bannerShown);

    if (!enabled)
        return;
    if (InterlockedExchange(&g_bannerShown, 1) != 0)
        return;

    const SIZE_T kChars = 1024;
    wchar_t* text = (wchar_t*)HeapAlloc(GetProcessHeap(), 0, kChars * sizeof(wchar_t));
    if (!text)
    {
        LogWrite(L"banner: HeapAlloc failed");
        return;
    }

    wchar_t modulePath[MAX_PATH] = L"(unknown)";
    GetModuleFileNameW(g_module, modulePath, MAX_PATH);

    _snwprintf(text, kChars,
               L"WordTab is loaded inside Word.\r\n\r\n"
               L"native build, no .NET runtime in this process\r\n"
               L"connect mode: %d\r\n"
               L"module: %s\r\n"
               L"log: %s",
               (int)connectMode, modulePath, LogFilePath());
    text[kChars - 1] = L'\0';

    HANDLE thread = CreateThread(NULL, 0, BannerThread, text, 0, NULL);
    if (thread)
    {
        LogWrite(L"banner: thread created");
        CloseHandle(thread);
    }
    else
    {
        LogWrite(L"banner: CreateThread failed (lastError=%lu)", GetLastError());
        HeapFree(GetProcessHeap(), 0, text);
    }
}

// ---------------------------------------------------------------------------------------------
// The object itself.
// ---------------------------------------------------------------------------------------------

class Connect : public IDTExtensibility2
{
public:
    Connect() : m_refCount(1), m_application(NULL), m_addInInst(NULL)
    {
        InterlockedIncrement(&g_objectCount);
    }

    virtual ~Connect()
    {
        ReleaseHostObjects();
        InterlockedDecrement(&g_objectCount);
    }

    // -- IUnknown ------------------------------------------------------------------------------

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppv)
    {
        if (!ppv)
            return E_POINTER;
        *ppv = NULL;

        // One vtable serves all three. IDTExtensibility2 derives from IDispatch, so a caller that
        // asks for either gets something it can use, whichever way it decides to call us.
        if (IsEqualIID(riid, IID_IUnknown) ||
            IsEqualIID(riid, IID_IDispatch) ||
            IsEqualIID(riid, IID_IDTExtensibility2))
        {
            *ppv = static_cast<IDTExtensibility2*>(this);
            AddRef();
            return S_OK;
        }

        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef()
    {
        return (ULONG)InterlockedIncrement(&m_refCount);
    }

    ULONG STDMETHODCALLTYPE Release()
    {
        LONG remaining = InterlockedDecrement(&m_refCount);
        if (remaining == 0)
        {
            this->~Connect();
            HeapFree(GetProcessHeap(), 0, this);
        }
        return (ULONG)remaining;
    }

    // -- IDispatch -----------------------------------------------------------------------------
    //
    // No type library, so no type info. Word does not need it: it knows IDTExtensibility2's
    // DISPIDs, and GetIDsOfNames below covers a caller that works by name instead.

    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT* count)
    {
        if (!count)
            return E_POINTER;
        *count = 0;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT, LCID, ITypeInfo** typeInfo)
    {
        if (typeInfo)
            *typeInfo = NULL;
        return E_NOTIMPL;
    }

    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID, LPOLESTR* names, UINT nameCount,
                                            LCID, DISPID* dispIds)
    {
        if (!names || !dispIds)
            return E_POINTER;

        HRESULT result = S_OK;
        for (UINT i = 0; i < nameCount; i++)
        {
            dispIds[i] = DISPID_UNKNOWN;
            if      (_wcsicmp(names[i], L"OnConnection")      == 0) dispIds[i] = DISPID_OnConnection;
            else if (_wcsicmp(names[i], L"OnDisconnection")   == 0) dispIds[i] = DISPID_OnDisconnection;
            else if (_wcsicmp(names[i], L"OnAddInsUpdate")    == 0) dispIds[i] = DISPID_OnAddInsUpdate;
            else if (_wcsicmp(names[i], L"OnStartupComplete") == 0) dispIds[i] = DISPID_OnStartupComplete;
            else if (_wcsicmp(names[i], L"OnBeginShutdown")   == 0) dispIds[i] = DISPID_OnBeginShutdown;
            else result = DISP_E_UNKNOWNNAME;
        }
        return result;
    }

    // Route a by-DISPID call to the same methods the vtable path uses, so it cannot matter which
    // way Word chooses to call us. Arguments arrive in DISPPARAMS in reverse order.
    HRESULT STDMETHODCALLTYPE Invoke(DISPID dispId, REFIID, LCID, WORD flags,
                                     DISPPARAMS* params, VARIANT*, EXCEPINFO*, UINT*)
    {
        if (!(flags & DISPATCH_METHOD))
            return DISP_E_MEMBERNOTFOUND;

        UINT argCount = params ? params->cArgs : 0;
        VARIANT* args = params ? params->rgvarg : NULL;

        LogWrite(L"Invoke  dispid=%d argc=%u  (Word is calling us by DISPID, not vtable)",
                 (int)dispId, argCount);

        switch (dispId)
        {
        case DISPID_OnConnection:
        {
            // (application, connectMode, addInInst, custom) reversed => [3],[2],[1],[0]
            IDispatch* application = (argCount >= 4 && args[3].vt == VT_DISPATCH) ? args[3].pdispVal : NULL;
            long mode = (argCount >= 3 && args[2].vt == VT_I4) ? args[2].lVal : 0;
            IDispatch* addInInst = (argCount >= 2 && args[1].vt == VT_DISPATCH) ? args[1].pdispVal : NULL;
            return OnConnection(application, (enum ext_ConnectMode)mode, addInInst, NULL);
        }
        case DISPID_OnDisconnection:
        {
            long mode = (argCount >= 2 && args[1].vt == VT_I4) ? args[1].lVal : 0;
            return OnDisconnection((enum ext_DisconnectMode)mode, NULL);
        }
        case DISPID_OnAddInsUpdate:    return OnAddInsUpdate(NULL);
        case DISPID_OnStartupComplete: return OnStartupComplete(NULL);
        case DISPID_OnBeginShutdown:   return OnBeginShutdown(NULL);
        default:                       return DISP_E_MEMBERNOTFOUND;
        }
    }

    // -- IDTExtensibility2 ---------------------------------------------------------------------

    HRESULT STDMETHODCALLTYPE OnConnection(IDispatch* application, enum ext_ConnectMode connectMode,
                                           IDispatch* addInInst, SAFEARRAY**)
    {
        ReleaseHostObjects();

        m_application = application;
        if (m_application) m_application->AddRef();
        m_addInInst = addInInst;
        if (m_addInInst) m_addInInst->AddRef();

        // The strip's new-tab button reaches Word through this, and the strip is not a COM object.
        SetApplication(application);

        // Interrogating the Application object is the proof we are talking to the real Word
        // rather than merely having been instantiated. The build number should match the one
        // recorded for this rig.
        BSTR name    = GetStringProperty(application, L"Name");
        BSTR version = GetStringProperty(application, L"Version");
        BSTR build   = GetStringProperty(application, L"Build");
        BSTR docs    = GetDocumentCount(application);

        LogWrite(L"OnConnection  mode=%d  app=%s version=%s build=%s documents=%s",
                 (int)connectMode,
                 name    ? name    : L"(?)",
                 version ? version : L"(?)",
                 build   ? build   : L"(?)",
                 docs    ? docs    : L"(?)");

        SysFreeString(name);
        SysFreeString(version);
        SysFreeString(build);
        SysFreeString(docs);

        ShowBannerOnce(connectMode);

        // ext_cm_AfterStartup means we were switched on mid-session: Word is already up and no
        // OnStartupComplete is coming, so the frame work has to start here instead. In the normal
        // startup case it waits, because Word has not finished making its first window yet.
        if (connectMode != ext_cm_Startup)
        {
            LogWrite(L"OnConnection  mode is not ext_cm_Startup - starting frame work here, "
                     L"since no OnStartupComplete will arrive");
            FramesStart();
        }

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE OnStartupComplete(SAFEARRAY**)
    {
        // Word's UI exists by now - or nearly. This is where the window work will start, so log
        // what we can see: the OpusApp windows this process owns, and whether they are visible
        // yet. Visibility matters because "the window exists but is still hidden at
        // OnStartupComplete" and "the window does not exist" need different handling, and the
        // next slice has to know which one it is dealing with.
        int total = 0;
        int visible = 0;
        HWND first = NULL;
        HWND window = NULL;

        while ((window = FindWindowExW(NULL, window, L"OpusApp", NULL)) != NULL)
        {
            DWORD pid = 0;
            GetWindowThreadProcessId(window, &pid);
            if (pid != GetCurrentProcessId())
                continue;

            total++;
            if (!first)
                first = window;
            if (IsWindowVisible(window))
                visible++;
        }

        LogWrite(L"OnStartupComplete  OpusApp windows: total=%d visible=%d first=0x%p",
                 total, visible, (void*)first);

        // Subclass the frame(s) and start watching for new ones. Note this does not wait for the
        // window to become visible: at this point it reliably is not (total=1 visible=0, every
        // run), and waiting for visibility here waits forever.
        FramesStart();
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE OnAddInsUpdate(SAFEARRAY**)
    {
        LogWrite(L"OnAddInsUpdate");
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE OnBeginShutdown(SAFEARRAY**)
    {
        LogWrite(L"OnBeginShutdown");

        // The last safe moment to take our window procedure back out of Word's frames: they still
        // exist here, and by OnDisconnection some of them may not.
        FramesStop();
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE OnDisconnection(enum ext_DisconnectMode removeMode, SAFEARRAY**)
    {
        LogWrite(L"OnDisconnection  mode=%d", (int)removeMode);

        // ext_dm_UserClosed means Word keeps running without us, so teardown has to be real and
        // cannot rely on the process exiting. FramesStop is idempotent - on a normal shutdown
        // OnBeginShutdown has already run it.
        FramesStop();
        ReleaseHostObjects();
        return S_OK;
    }

private:
    void ReleaseHostObjects()
    {
        SetApplication(NULL);
        if (m_addInInst)   { m_addInInst->Release();   m_addInInst = NULL; }
        if (m_application) { m_application->Release(); m_application = NULL; }
    }

    LONG m_refCount;
    IDispatch* m_application;
    IDispatch* m_addInInst;
};

// Allocated from the process heap with placement new rather than the CRT's operator new: no
// exceptions to leak across the COM boundary and no C++ runtime dependency in the DLL.
HRESULT WordTabCreateConnect(REFIID riid, void** ppv)
{
    if (!ppv)
        return E_POINTER;
    *ppv = NULL;

    void* storage = HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, sizeof(Connect));
    if (!storage)
        return E_OUTOFMEMORY;

    Connect* object = new (storage) Connect();

    HRESULT hr = object->QueryInterface(riid, ppv);
    object->Release();          // drop the construction reference; QI took its own on success
    return hr;
}
