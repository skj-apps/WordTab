// WordTab - COM server plumbing: the GUIDs, the class factory, and the two exports COM needs.

#include "wordtab.h"
#include <new>

// ---------------------------------------------------------------------------------------------
// GUIDs, defined here rather than via initguid.h so there is exactly one definition to find.
// ---------------------------------------------------------------------------------------------

// {4BF75ED9-10EE-4866-BF4A-3D663A4149A1}  - keep in step with WORDTAB_CLSID_STRING in wordtab.h
const CLSID CLSID_WordTabConnect =
    { 0x4BF75ED9, 0x10EE, 0x4866, { 0xBF, 0x4A, 0x3D, 0x66, 0x3A, 0x41, 0x49, 0xA1 } };

// {B65AD801-ABAF-11D0-BB8B-00A0C90F2744}  - IDTExtensibility2, fixed by Office
const IID IID_IDTExtensibility2 =
    { 0xB65AD801, 0xABAF, 0x11D0, { 0xBB, 0x8B, 0x00, 0xA0, 0xC9, 0x0F, 0x27, 0x44 } };

LONG g_objectCount = 0;
LONG g_lockCount = 0;
HMODULE g_module = NULL;

// ---------------------------------------------------------------------------------------------
// Class factory.
// ---------------------------------------------------------------------------------------------

class ClassFactory : public IClassFactory
{
public:
    ClassFactory() : m_refCount(1) { }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppv)
    {
        if (!ppv)
            return E_POINTER;
        *ppv = NULL;

        if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IClassFactory))
        {
            *ppv = static_cast<IClassFactory*>(this);
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
            this->~ClassFactory();
            HeapFree(GetProcessHeap(), 0, this);
        }
        return (ULONG)remaining;
    }

    HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown* outer, REFIID riid, void** ppv)
    {
        if (!ppv)
            return E_POINTER;
        *ppv = NULL;

        // We do not support aggregation, and saying so plainly is required by the contract.
        if (outer)
            return CLASS_E_NOAGGREGATION;

        return WordTabCreateConnect(riid, ppv);
    }

    HRESULT STDMETHODCALLTYPE LockServer(BOOL lock)
    {
        if (lock)
            InterlockedIncrement(&g_lockCount);
        else
            InterlockedDecrement(&g_lockCount);
        return S_OK;
    }

private:
    LONG m_refCount;
};

// ---------------------------------------------------------------------------------------------
// Exports.
// ---------------------------------------------------------------------------------------------

extern "C" HRESULT STDAPICALLTYPE DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppv)
{
    if (!ppv)
        return E_POINTER;
    *ppv = NULL;

    // The first point at which it is safe to touch the filesystem or load anything: DllMain runs
    // under the loader lock, so the log deliberately starts here rather than at attach.
    if (!IsEqualCLSID(rclsid, CLSID_WordTabConnect))
        return CLASS_E_CLASSNOTAVAILABLE;

    wchar_t modulePath[MAX_PATH] = L"(unknown)";
    GetModuleFileNameW(g_module, modulePath, MAX_PATH);

    wchar_t hostPath[MAX_PATH] = L"(unknown)";
    GetModuleFileNameW(NULL, hostPath, MAX_PATH);

    LogWrite(L"---- DllGetClassObject ----  module=%s  host=%s  bits=%d",
             modulePath, hostPath, (int)(sizeof(void*) * 8));

    void* storage = HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, sizeof(ClassFactory));
    if (!storage)
        return E_OUTOFMEMORY;

    ClassFactory* factory = new (storage) ClassFactory();

    HRESULT hr = factory->QueryInterface(riid, ppv);
    factory->Release();
    return hr;
}

extern "C" HRESULT STDAPICALLTYPE DllCanUnloadNow(void)
{
    return (g_objectCount == 0 && g_lockCount == 0) ? S_OK : S_FALSE;
}

BOOL WINAPI DllMain(HINSTANCE instance, DWORD reason, LPVOID)
{
    // Nothing here may allocate, log, or load a library. DllMain runs under the loader lock, and
    // the logger reaches for shell32 to find %LOCALAPPDATA%, which is exactly the kind of call
    // that deadlocks a process at load time. All real work waits for DllGetClassObject.
    switch (reason)
    {
    case DLL_PROCESS_ATTACH:
        g_module = (HMODULE)instance;
        DisableThreadLibraryCalls(instance);
        break;
    default:
        break;
    }
    return TRUE;
}
