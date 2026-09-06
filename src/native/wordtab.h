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

// The folder holding the document behind one frame, written into `out` with no trailing separator.
//
// The same two failure shapes as WordTabReadModified, kept apart for the same reason:
//
// TRUE means Word answered. `out` is the folder, or an EMPTY STRING when there genuinely is none -
// a document that has never been saved has no path, and a frame no window in the collection claims
// is the Protected View case. Both are determinate answers about a document with no folder.
//
// FALSE means Word would not answer at all and `out` is empty. The caller must not draw that as
// "this document has no folder": it is "nobody asked successfully".
BOOL WordTabReadDocumentPath(HWND frame, wchar_t* out, int chars);

// Save the document behind a tab, through Word's own Document.Save - so an unchanged document is
// untouched and one that has never been saved gets Word's Save As dialog, exactly as Ctrl+S would.
//
// `frame` must already be the window Word considers active. This does not activate it: it *checks*
// that Word agrees, and saves nothing if it does not. Saving the wrong document is the one failure
// in this add-in that reaches the user's data, so it is made impossible rather than unlikely.
// Put a document back to one page across, if it opened showing several side by side.
//
// Word keeps a column count as a DEFAULT and crushes the zoom to fit it, and the ribbon buttons
// that appear to fix it only fix the window in front of you - which is why the person this was
// built for was setting the view by hand on every Word start. Reports what it found, so the log
// can say what was wrong rather than only that something was.
// Why it decided what it decided. **FALSE had five meanings and the caller logged on none of them**,
// so "healthy, nothing to do" and "the correction was attempted and Word refused the write" were
// byte-identical in the log - and the second of those leaves the document three pages across at 10%
// zoom, which is the complaint this whole path exists to answer. OnePage ran six times in the
// 2026-09-04 report and said nothing about any of them.
//
// The reason is decided HERE, at the five exits, and never re-derived by the caller. That is not a
// style preference: the illness test is `columns > 1 && columns < 99`, because 99 is Word's "as many
// as fit" and IS the healthy value, and two separate attempts at rebuilding that test in stack.cpp
// dropped the `< 99` and would have printed "Word refused the write" on every healthy join.
enum WordTabOnePageWhy
{
    OnePage_Corrected = 0,   // it was wrong, it was set, Word accepted it
    OnePage_AlreadyFine,     // nothing to do - one page across already, at a readable zoom
    OnePage_WordRefused,     // it WAS wrong, the write was attempted, and Word said no
    OnePage_CouldNotRead,    // the Zoom object answered, but its properties did not
    OnePage_NoZoom,          // View.Zoom could not be obtained
    OnePage_NoView,          // Window.View could not be obtained
    OnePage_NoWindow         // no Word window claims this frame - Protected View looks like this
};

// The word for one of the above, for a log line. Never NULL.
const wchar_t* WordTabOnePageWhyName(enum WordTabOnePageWhy why);

// `why` may be NULL, and receives the reason whatever the return value is.
BOOL WordTabOnePageView(HWND frame, LONG* wasColumns, LONG* wasZoom, enum WordTabOnePageWhy* why);

BOOL WordTabSaveDocument(HWND frame);

// ---------------------------------------------------------------------------------------------
// Frame windows - the in-process subclass of Word's OpusApp frames. See frames.cpp.
//
// Both calls are idempotent and must both happen on Word's UI thread. FramesStop must run before
// Word tears its windows down, or our window procedure is left in a chain we no longer control.
// ---------------------------------------------------------------------------------------------

void FramesStart(void);
void FramesStop(void);

// Is Word inside its own modal move/size loop right now - a frame being dragged by its caption, or
// moved or sized from the window menu?
//
// Exported for the janitor, which runs on a THREAD timer and is therefore dispatched by whatever
// pump happens to be running. Inside this loop that pump is Word's own, so a half-second tick lands
// between two frames of a window the user is watching move, and whatever it does is paid for out of
// them. Anything on that timer that is expensive and can wait a gesture should ask this first.
//
// It says nothing about Win+Left and Win+Right. Those snap a window without entering the loop and
// without a system command, which frames.cpp already names as the case it cannot see - so FALSE here
// is not "nobody is moving this window".
BOOL FramesInModalMoveLoop(void);

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

// Is a press being held on a tab? A gesture in progress, so nothing may be concluded from where
// the windows have ended up - the carried card raises the window it was picked up from.
BOOL StripTabPressHeld(void);

// Is there a document open in this window? This is the membership test - a window with no document
// is not a tab - and it is deliberately *not* the same question as "does it have a `_WwF`". Word
// keeps the document frame for the life of the window and empties it when the last document closes.
BOOL StripHasDocument(HWND frame);

// Every `_WwF` this window has, which one the strip is bound to, and what is inside each - as one
// line of text for the log.
//
// This is a diagnostic and it is here because the defect it is for cannot be reproduced on the
// development machine. "A View change closes tabs" was reported from a rig I cannot reach, and the
// log it produced said a window left the stack without saying anything about the state that decided
// it. A window with two document frames and a window with one empty one are the same line otherwise,
// and they are different bugs.
void StripDescribeDocumentFrames(HWND frame, wchar_t* out, int chars);

// The DPI this window is being drawn at, as the strip computes it - honouring TabDpi. The stack
// uses it to work out how wide a printed page is on this screen; see ApplyDefaultRowRect.
int StripDpiOf(HWND frame);

// One line describing what DPI context a window is being read in: its own awareness context, the
// calling thread's, and the DPI Windows reports for it with no TabDpi override in front of it.
//
// A PROBE. Nothing branches on it and nothing should - it exists so the next report can answer two
// questions that the 2026-09-04 one could not. `label` is prefixed verbatim so one log line can
// carry two windows and say which is which. It always writes something, including on Windows older
// than the APIs it needs.
void StripDescribeDpiContext(HWND hwnd, const wchar_t* label, wchar_t* out, int chars);

BOOL StripGetNatural(HWND frame, RECT* natural);
// `why` names the path that asked, because this is the one writer of a window's natural rect that
// did not come from Word laying that window out - see the note on the definition.
void StripSetNatural(HWND frame, const RECT* natural, const wchar_t* why);
void StripRefit(HWND frame);
void StripRefreshTabs(void);

// The name to put on a tab: the frame's title with Word's " - Word" suffix removed.
void WordTabFrameTitle(HWND frame, wchar_t* out, int chars);

// What the tab is CALLED - the name the row is drawing for this frame right now.
//
// The same as the above except in one case: a frame whose document has gone, whose title Word has
// already reverted to a bare "Word", and about which the stack has not yet decided whether the tab
// goes too. A tab that took that name there would be announcing a decision nobody has made. Live in
// every other case, deliberately - see the definition for the regression that rule was written from.
//
// Everything that shows a tab's name to anybody - the two painters, the tooltip, the close dialog,
// the row's own log line - asks this, so they cannot disagree.
void StripTabName(HWND frame, wchar_t* out, int chars);

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

// Called when the user finishes moving or sizing a frame. The row keeps the size it is given,
// so the size has to be captured at the moment a person settles on one - not from every
// intermediate rectangle a drag passes through. Ignores any frame that is not the one defining
// the row.
void StackRememberRowSize(HWND frame);

// Is the stack waiting to see whether this frame's missing document is a moment or a fact?
//
// TRUE only inside that grace window. FALSE for a frame that is not in the row, for one that is
// there and holding a document, and for one whose loss has already been settled either way.
BOOL StackIsWaitingFor(HWND frame);

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

// The tab `delta` positions along the row from this one, wrapping at both ends. This is the whole of
// what Ctrl+Tab needs and it deliberately lives here rather than in the keyboard code, because the
// row order is g_members' order and nothing outside this file may walk that array.
//
// NULL when there is nowhere to go: not a tab in a stack, or the only tab in one. Word's own Ctrl+F6
// cannot be used for this - Word orders its windows most-recently-used, the row is ordered by
// position, and the two agree only until the user switches documents once. RESULT-keyboard.md §3.
HWND StackNeighbourTab(HWND frame, int delta);

// Take a tab out of the stack and leave it standing as a window of its own.
//
// Note the neighbour above: StackDetachFrame is a different operation with a similar name - that one
// is a window being destroyed and forgotten, this one is a window the user still has. Nothing about
// the subclass changes here; the window keeps its strip and draws itself as a single tab.
//
// Refused for a window that is not a tab in a stack and for the only tab in one, both silently -
// the caller greys the menu item, and a command posted from a menu can arrive after the row has
// changed underneath it.
//
// **It stays out.** Membership is otherwise re-tested twice a second from what is true about the
// window, and a torn-off window passes every one of those tests. See the definition.
void StackTearOffTab(HWND frame);

// Whether tearing off is available: stacking on, and HKCU\Software\WordTab\TabTearOff not 0. Asked
// by the strip rather than read there, so the menu item and the drag gesture cannot disagree.
BOOL StackCanTearOff(void);

// Put a window that is standing on its own back into the stack - the inverse of StackTearOffTab, and
// the reason tearing off is no longer a one-way door.
//
// Its tab arrives at the END of the row rather than at the place it used to hold. The window is
// arriving, not being undone: the user may have torn it off ten minutes and three reorders ago, and a
// tab that reappeared in the middle of a row they have since rearranged would be a surprise. Same
// rule, and the same flag, as a recycled frame coming back with a new document in it.
//
// Refused for a window that is already a tab in a stack, for one Word will not let us place, and when
// there is no stack to join - all silently, because the strip greys nothing here: it simply does not
// start the gesture. See StackCanRejoin.
void StackJoinTab(HWND frame);

// Whether this window could be dropped back into a stack right now: stacking on, this window is a
// member that is not joined, it is placeable, and there is at least one joined window for it to join.
// Asked by the strip on every mouse-move of a rejoin drag, so it decides both what the pointer looks
// like and what the drop does - one answer, not two that can disagree.
BOOL StackCanRejoin(HWND frame);

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

// The window's own close button, Alt+F4, or the window menu - anything that arrives as SC_CLOSE.
// TRUE means the stack has taken the command over and the frame must swallow it; FALSE means this
// was not a stack of several documents and Word's own close is exactly right. See the definition.
BOOL StackCloseWindowCommand(HWND frame);

// Where a frame that is about to be created should be put, so that a new document's window is never
// drawn anywhere but on the stack. FALSE when there is no stack to match, and then Word's own
// choice stands. Asked from inside the CBT hook, before the window exists.
BOOL StackProposeCreateRect(RECT* out);

// The same queue, narrowed to the tabs after this one in the row. Refused when there are none, so
// the caller does not have to check before asking.
void StackCloseToRight(HWND from);

// Read a DWORD switch from HKCU\Software\WordTab as a boolean. Absent means the default, so a
// fresh install behaves like a configured one. Used for the switches that must be flippable
// without a rebuild.
BOOL WordTabReadFlag(const wchar_t* name, BOOL defaultValue);

// The same key and the same "absent means the default" rule, but the value kept as a number rather
// than flattened to on/off. Separate from WordTabReadFlag rather than a parameter on it, because
// every existing caller wants the boolean and a shared reader returning DWORD would put a
// `!= 0` at seventeen call sites.
DWORD WordTabReadNumber(const wchar_t* name, DWORD defaultValue);

// The other direction, and the only thing here that writes to the registry. WordTab's switches
// are the user's to set and this does not touch them; it exists for the handful of values
// WordTab records for itself - where the row was left - which have to survive Word closing.
void WordTabWriteNumber(const wchar_t* name, DWORD value);
