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

// One element of a collection: Application.Windows.Item(i), one-based like every Office collection.
//
// The first call in this add-in that passes an argument through IDispatch::Invoke, so the shape is
// worth stating once. DISPPARAMS carries the arguments in an array that is *reversed* - last
// parameter first - which does not show with one argument and is exactly the kind of thing that
// works by accident until a second one is added. `Item` is asked for by name rather than through
// DISPID_VALUE: a collection's default member is a convention, and a named lookup that fails says so
// instead of invoking something else.
//
// NULL on any failure, and the caller Releases what it gets, like GetObjectProperty.
static IDispatch* GetItemAt(IDispatch* collection, LONG index)
{
    if (!collection)
        return NULL;

    DISPID dispid = 0;
    LPOLESTR nameCopy = (LPOLESTR)L"Item";
    if (FAILED(collection->GetIDsOfNames(IID_NULL, &nameCopy, 1, LOCALE_USER_DEFAULT, &dispid)))
        return NULL;

    VARIANT arg;
    VariantInit(&arg);
    arg.vt   = VT_I4;
    arg.lVal = index;

    DISPPARAMS args = { &arg, NULL, 1, 0 };
    VARIANT result;
    VariantInit(&result);
    EXCEPINFO error;
    memset(&error, 0, sizeof(error));

    HRESULT hr = collection->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                                    DISPATCH_METHOD | DISPATCH_PROPERTYGET,
                                    &args, &result, &error, NULL);
    ClearExceptionInfo(&error);

    IDispatch* item = NULL;
    if (SUCCEEDED(hr) && result.vt == VT_DISPATCH && result.pdispVal)
    {
        item = result.pdispVal;
        item->AddRef();
    }

    VariantClear(&result);
    return item;
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

// Write a numeric property. The first write into Word's object model from this add-in, so the shape
// is worth stating: a property PUT carries its value as an argument NAMED DISPID_PROPERTYPUT, which
// is not something DISPPARAMS makes obvious - cNamedArgs is 1 and rgdispidNamedArgs points at that
// one id. Getting it wrong does not fail loudly; Word returns DISP_E_PARAMNOTOPTIONAL, or quietly
// does nothing at all.
static BOOL SetLongProperty(IDispatch* disp, const wchar_t* name, LONG value)
{
    if (!disp)
        return FALSE;

    DISPID dispid = 0;
    LPOLESTR nameCopy = (LPOLESTR)name;
    if (FAILED(disp->GetIDsOfNames(IID_NULL, &nameCopy, 1, LOCALE_USER_DEFAULT, &dispid)))
        return FALSE;

    VARIANT arg;
    VariantInit(&arg);
    arg.vt   = VT_I4;
    arg.lVal = value;

    DISPID    putId = DISPID_PROPERTYPUT;
    DISPPARAMS args = { &arg, &putId, 1, 1 };
    EXCEPINFO error;
    memset(&error, 0, sizeof(error));

    HRESULT hr = disp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                              DISPATCH_PROPERTYPUT, &args, NULL, &error, NULL);
    ClearExceptionInfo(&error);
    VariantClear(&arg);

    return SUCCEEDED(hr) ? TRUE : FALSE;
}

// The Word window that owns a frame, found by handle. Caller Releases.
//
// Not Application.ActiveWindow: this is asked as a window JOINS the row, which is not always the one
// Word considers active, and there is no reason to activate anything merely to look at its view. A
// frame that no window claims is the Protected View case - such a document lives in
// ProtectedViewWindows and in no collection reachable from here - and NULL is the honest answer.
static IDispatch* WindowForFrame(HWND frame)
{
    if (!g_application || !frame)
        return NULL;

    IDispatch* windows = GetObjectProperty(g_application, L"Windows");
    if (!windows)
        return NULL;

    LONG total = 0;
    if (!GetLongProperty(windows, L"Count", &total))
    {
        windows->Release();
        return NULL;
    }

    IDispatch* found = NULL;
    for (LONG index = 1; index <= total && !found; index++)
    {
        IDispatch* window = GetItemAt(windows, index);
        if (!window)
            continue;

        // The same (LONG)(LONG_PTR) narrowing as everywhere else here: window handles are 32-bit
        // values sign-extended into a pointer, and Word reports Hwnd as a long.
        LONG reported = 0;
        if (GetLongProperty(window, L"Hwnd", &reported) && reported == (LONG)(LONG_PTR)frame)
            found = window;                  // the reference is the caller's now
        else
            window->Release();
    }

    windows->Release();
    return found;
}

// Below this, a zoom is not a choice: it is what fitting several pages across left behind.
#define ONEPAGE_MIN_ZOOM 50

// One page across, which is what the user does by hand on every single Word start.
//
// Their words, and it took three of my answers before I heard them: "i have to chage view to single
// page on word open", "every time i mean". Everything built for "it opens 3 pages wide" was about
// the WIDTH of the window, and this is not that. Measured on 16.0.20228:
//
//   - A healthy Word reports Zoom.PageColumns = 99, its "as many as fit", and pages side by side are
//     then purely a function of how wide the window is. That case the row's own width answers.
//   - Once anything sets a COLUMN COUNT, Word keeps it as a default and crushes the zoom to fit that
//     many pages: PageColumns=2 in a 1305px window took the zoom to 10%. It survives closing the
//     document, closing Word, and opening a document Word has never seen - including a brand new
//     blank one from the template. It lives in HKCU\...\Word\Data, which is an opaque blob.
//   - The ribbon's One Page and 100% buttons fix the window in front of you and do NOT change that
//     default. That is the whole of "every time": the button being pressed could never have stuck.
//
// So the add-in does what they do, at the moment a document arrives in the row. PageColumns first
// and the zoom second, because setting the columns alone leaves the crushed zoom exactly where it
// was - measured: three columns at 10% became one column at 10%, which is one page and still
// unreadable.
//
// It acts ONLY when the view is showing more than one page across. A window already on one page is
// left alone entirely, zoom included, because a zoom somebody chose for a document is theirs.
const wchar_t* WordTabOnePageWhyName(enum WordTabOnePageWhy why)
{
    switch (why)
    {
    case OnePage_Corrected:    return L"it was put back to one page at 100%";
    case OnePage_AlreadyFine:  return L"already on one page at a readable zoom - nothing to do";
    case OnePage_WordRefused:  return L"IT WAS WRONG AND WORD REFUSED THE WRITE";
    case OnePage_CouldNotRead: return L"Word gave up its Zoom object but not its PageColumns/Percentage";
    case OnePage_NoZoom:       return L"View.Zoom could not be obtained";
    case OnePage_NoView:       return L"Window.View could not be obtained";
    case OnePage_NoWindow:     return L"no Word window claims this frame (Protected View reads like this)";
    }
    return L"unknown";
}

BOOL WordTabOnePageView(HWND frame, LONG* wasColumns, LONG* wasZoom, enum WordTabOnePageWhy* why)
{
    if (wasColumns) *wasColumns = 0;
    if (wasZoom)    *wasZoom    = 0;
    if (why)        *why        = OnePage_NoWindow;

    IDispatch* window = WindowForFrame(frame);
    if (!window)
        return FALSE;

    IDispatch* view = GetObjectProperty(window, L"View");
    window->Release();
    if (!view)
    {
        if (why) *why = OnePage_NoView;
        return FALSE;
    }

    IDispatch* zoom = GetObjectProperty(view, L"Zoom");
    view->Release();
    if (!zoom)
    {
        if (why) *why = OnePage_NoZoom;
        return FALSE;
    }

    // **Both reads are kept, and that is the point of the fourth reason.** The out-params below are
    // written unconditionally, so `columns == 0 && percent == 0` is NOT "nobody asked": it is also
    // what a Zoom object that answered and then refused both property reads leaves behind. Those two
    // are a healthy-looking log line and a total failure to see the document, and until now they
    // were the same three numbers.
    LONG columns = 0;
    LONG percent = 0;
    BOOL read     = GetLongProperty(zoom, L"PageColumns", &columns);
    BOOL readZoom = GetLongProperty(zoom, L"Percentage", &percent);

    if (wasColumns) *wasColumns = columns;
    if (wasZoom)    *wasZoom    = percent;

    // 99 is Word's "as many as fit", not a request for ninety-nine pages, and it is the healthy
    // state. Left alone: what it draws is decided by the window's width, which is the row's business
    // and is answered elsewhere.
    BOOL manyPages = (read && columns > 1 && columns < 99) ? TRUE : FALSE;

    // The other half of the same illness, and the suite for this found it by printing a number an
    // assertion was not looking at: correcting the columns leaves Word's remembered ZOOM crushed.
    // The second document opened afterwards came up "1 page across at 10%" - one page, and
    // unreadable, which is not what anybody meant by fixing it. Word recomputed the zoom when it was
    // asked for three columns and does not recompute it when it is asked for one.
    //
    // A zoom under half size is not a reading choice, it is the leftover of a fit-many-pages
    // calculation - Word's own floor is 10%, which is what a three-column fit produced here. Above
    // that, whatever zoom a document opens at is somebody's business and is left alone.
    BOOL tooSmall = (!manyPages && percent > 0 && percent < ONEPAGE_MIN_ZOOM) ? TRUE : FALSE;

    if (!manyPages && !tooSmall)
    {
        // Two different silences, told apart. Nothing is done in either case and the control flow is
        // exactly what it was - but "I looked and it was fine" and "I could not read the view at all"
        // are not the same report, and one of them means the next line in the log is about a document
        // nobody actually measured.
        if (why) *why = (read && readZoom) ? OnePage_AlreadyFine : OnePage_CouldNotRead;
        zoom->Release();
        return FALSE;
    }

    BOOL ok = TRUE;
    if (manyPages)
        ok = SetLongProperty(zoom, L"PageColumns", 1);
    if (ok)
        ok = SetLongProperty(zoom, L"Percentage", 100);

    if (why) *why = ok ? OnePage_Corrected : OnePage_WordRefused;

    zoom->Release();
    return ok;
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

// Which of these tabs have a document with unsaved changes.
//
// One pass over Application.Windows for the whole row rather than a lookup per tab: the collection
// has to be walked either way, and walking it once is the difference between four Invokes and four
// Invokes *per document*. This runs twice a second on Word's own UI thread, so its cost is Word's
// own responsiveness.
//
// **Window.Hwnd is the join, and it is the OpusApp frame handle.** Measured for every window at once
// rather than only for the active one - tools\probe-saved.ps1 - which is what makes it possible to
// find a document without activating anything, and therefore safe to do on a timer at all.
//
// Two failures, deliberately not merged into one:
//   - **Word would not answer at all** (no Application, no Windows, no Count). Returns FALSE having
//     touched nothing, so every flag keeps the value it already had. A tick that could not take a
//     measurement must not be allowed to look like a measurement that came back "clean".
//   - **Word answered and no window claims this frame.** That is determinate, and the answer is
//     FALSE. It is the Protected View case: such a document lives in a sandboxed WINWORD of its own
//     and is in neither Application.Windows nor Documents, only in ProtectedViewWindows - measured.
//     It is also the right answer on its own merits, because a document in Protected View cannot be
//     edited and so can never have unsaved changes.
//
// Document.Saved is TRUE when the document has NOT changed since it was last saved, so the dot is
// its negation. Saved is also *settable*, and nothing here ever writes it.
BOOL WordTabReadModified(const HWND* frames, int count, BOOL* modified)
{
    if (!frames || !modified || count <= 0 || !g_application)
        return FALSE;

    IDispatch* windows = GetObjectProperty(g_application, L"Windows");
    if (!windows)
        return FALSE;

    LONG total = 0;
    if (!GetLongProperty(windows, L"Count", &total))
    {
        windows->Release();
        return FALSE;
    }

    // Past here every frame gets a determinate answer, and the answer for one no window claims is
    // "clean" rather than "unknown".
    for (int i = 0; i < count; i++)
        modified[i] = FALSE;

    int unreadable = 0;

    for (LONG index = 1; index <= total; index++)
    {
        IDispatch* window = GetItemAt(windows, index);
        if (!window)
            continue;

        LONG reported = 0;
        if (!GetLongProperty(window, L"Hwnd", &reported))
        {
            window->Release();
            continue;
        }

        // Which tab this is, if it is one at all. A search rather than an index: one document can
        // own several windows - Word's own New Window - and Windows.Count is not the tab count.
        // The same (LONG)(LONG_PTR) narrowing as WordTabSaveDocument, for the same reason: window
        // handles are 32-bit values sign-extended into a pointer, and Word reports Hwnd as a long.
        int slot = -1;
        for (int i = 0; i < count; i++)
            if ((LONG)(LONG_PTR)frames[i] == reported)
                slot = i;

        if (slot < 0)
        {
            window->Release();
            continue;
        }

        IDispatch* document = GetObjectProperty(window, L"Document");
        window->Release();
        if (!document)
        {
            unreadable++;
            continue;
        }

        LONG saved = 0;
        if (GetLongProperty(document, L"Saved", &saved))
            modified[slot] = saved ? FALSE : TRUE;   // VARIANT_TRUE is -1, so never compare against TRUE
        else
            unreadable++;

        document->Release();
    }

    windows->Release();

    // Said when it starts and when it stops, never per tick. A window that is in the collection but
    // will not answer Document.Saved reads as clean above, and a tab silently stuck without its dot
    // is precisely the failure this project keeps meeting: a plausible answer from a question that
    // was never asked. Twice a second, so it is on change or it is noise.
    static BOOL complaining = FALSE;
    if (unreadable > 0 && !complaining)
    {
        complaining = TRUE;
        LogWrite(L"dot: %d of %ld window(s) would not answer Document.Saved - those tabs read as "
                 L"clean until they do", unreadable, total);
    }
    else if (unreadable == 0 && complaining)
    {
        complaining = FALSE;
        LogWrite(L"dot: every window answers Document.Saved again");
    }

    return TRUE;
}

// Which folder the document behind one tab lives in.
//
// **Read on demand rather than folded into the janitor's pass, and that is a deliberate difference
// from the dot.** The dot is polled because the thing it reports changes while nobody is looking; a
// document's folder changes only when the user does Save As, and the only moment anything needs to
// know it is the moment a tooltip is about to appear. Polling for it would mean a BSTR allocated per
// window twice a second to answer a question nobody had asked. The cost of asking here is paid by
// the hover that asked.
//
// One window, not the whole row, for the same reason: a hover is about one tab. The collection still
// has to be walked to find it, because Window.Hwnd is the only join between Word's windows and our
// frames - the same join, and the same (LONG)(LONG_PTR) narrowing, as WordTabReadModified.
//
// **Document.Path, not FullName.** The name is already on the tab and on the first line of the
// tooltip; repeating it in the second line would spend the width that makes the folder readable.
// Word returns it with no trailing separator, and an empty string for a document that has never been
// saved - which is a fact about that document, not a failure to read it.
BOOL WordTabReadDocumentPath(HWND frame, wchar_t* out, int chars)
{
    if (!out || chars <= 0)
        return FALSE;
    out[0] = L'\0';

    if (!frame || !IsWindow(frame) || !g_application)
        return FALSE;

    IDispatch* windows = GetObjectProperty(g_application, L"Windows");
    if (!windows)
        return FALSE;

    LONG total = 0;
    if (!GetLongProperty(windows, L"Count", &total))
    {
        windows->Release();
        return FALSE;
    }

    // Past here the answer is determinate. Falling off the end of the loop without finding the frame
    // means Word answered and no window of its claims it, and an empty folder is the truth about a
    // document this Application cannot see.
    BOOL answered = TRUE;

    for (LONG index = 1; index <= total; index++)
    {
        IDispatch* window = GetItemAt(windows, index);
        if (!window)
            continue;

        LONG reported = 0;
        if (!GetLongProperty(window, L"Hwnd", &reported) ||
            reported != (LONG)(LONG_PTR)frame)
        {
            window->Release();
            continue;
        }

        IDispatch* document = GetObjectProperty(window, L"Document");
        window->Release();
        if (!document)
        {
            // The window is Word's and it will not hand over its document. That is a failure to
            // read, not a document without a folder, so it is reported as one.
            answered = FALSE;
            break;
        }

        BSTR path = GetStringProperty(document, L"Path");
        document->Release();

        if (!path)
        {
            answered = FALSE;
        }
        else
        {
            wcsncpy(out, path, (size_t)(chars - 1));
            out[chars - 1] = L'\0';
            SysFreeString(path);
        }
        break;
    }

    windows->Release();
    return answered;
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
