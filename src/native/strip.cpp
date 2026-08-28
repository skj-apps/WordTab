// WordTab - the strip: carving 32 logical pixels out of Word's own layout, from inside Word.
//
// Spike 1 proved the geometry works (spikes\StripSpike\RESULT.md): shrink Word's `_WwF` document
// frame from the top, park our own window in the gap, and the result survives resize, maximize and
// Backstage. It did that from another process on a 30ms poll, which meant every relayout Word made
// was visible for a frame or two before we corrected it.
//
// This is the same geometry with the poll taken out. The move: rather than watch `_WwF` and correct
// it afterwards, we subclass `_WwF` itself and rewrite the *proposed* rectangle in
// WM_WINDOWPOSCHANGING. Word never gets to move it to the wrong place at all - there is no
// correction, because there is nothing to correct. Any SetWindowPos or MoveWindow on that window,
// from anywhere, comes through here first, so there is no path by which Word can lay out behind our
// back.
//
// The two rules that carried over from the spikes, both learned by watching them fail:
//
//   1. The sync must be idempotent against Word's layout. Word rewrites `_WwF` to its own natural
//      rect repeatedly - six times in one logged spike session - and a design that nudges by -32
//      each time walks the document frame off the bottom of the window in seconds. That is the old
//      "Doc Tabs" failure mode. So: what Word proposes is treated as natural, and the shift is
//      computed from it, never accumulated onto it.
//
//   2. Idempotence against Word is not idempotence against another add-in. Office Tab shifts this
//      same `_WwF`. Two processes each reading the other's +32 as a fresh natural rect squeezed the
//      document frame to 25px tall in seconds - measured, spike 2. Hence the top-edge test below
//      (if the top is exactly where we left it, nobody re-laid it out, so keep it) and a height
//      tripwire that refuses to shrink the document below twice the strip.
//
// The same house rules as frames.cpp apply: nothing throws, nothing escapes, and nothing here is
// slow. This code runs between Word and its own window procedure during a live resize.

#include "wordtab.h"
#include <commctrl.h>
#include <string.h>
#include <wchar.h>
#include <stdio.h>
#include <math.h>       // sqrtf, for the one anti-aliased stroke in DrawGlyph

// Height of the strip in logical pixels, scaled per-window by DPI. 32 is what spike 1 used and what
// Office Tab's own strip measures at 100%.
#define STRIP_LOGICAL_H   32
#define TAB_LOGICAL_W    220
#define MAX_STRIPS       256

// The affordances on a tab, all in logical pixels and all scaled per window.
#define TAB_LOGICAL_MIN_W  70    // below this a tab is a colour, not a label
#define TAB_LOGICAL_PAD     6    // the strip's left and right inset
#define TAB_LOGICAL_GAP     4    // between the last tab and the new-document button
#define CLOSE_LOGICAL      16    // the close button's hit target, a square
#define PLUS_LOGICAL       26    // the new-document button's width
#define CHEVRON_LOGICAL    20    // one scroll button, and there are two of them

// The label's type size, in logical pixels. A tab is TAB_LOGICAL_W wide whatever this is - the row's
// geometry is not derived from the text - so this is purely how many characters of a document's name
// survive before the ellipsis.
//
// 10 rather than 11, and that is the measurement rather than a preference. Segoe UI in the label rect
// this layout leaves (TAB_LOGICAL_W less the inset, the close button and the pad), against three
// ordinary Word filenames:
//
//     12px   31/34   32/47   29/31     every one of them cut
//     11px   33/34   33/47   30/31     every one of them still cut
//     10px   34/34   40/47   31/31     two of the three now whole
//
// 11 is the size that looks like it helps and does not: it buys one or two characters and leaves the
// ellipsis exactly where it was. The step that crosses the threshold for a normal filename is 10.
//
// Deliberately not Word's 12. Matching the ribbon is the right instinct for a control sitting under
// it, and the wrong trade here: this strip's job is to tell one document from another at a glance,
// which wants more of the name rather than type that matches. Anybody who disagrees has TabFontSize.
#define TAB_LOGICAL_FONT   10
#define TAB_FONT_MIN        8    // below this the name is a texture, not a word
#define TAB_FONT_MAX       16    // above this the descenders meet STRIP_LOGICAL_H

// How fast the row scrolls while a tab is being carried against one end of it. Per tick of
// DRAG_SCROLL_MS, so this is 16 logical px every 60ms - about a tab a second at the minimum width,
// which is fast enough to cross a full row while the hand stays still and slow enough to stop on the
// slot you meant.
#define DRAG_SCROLL_LOGICAL 16
#define DRAG_SCROLL_MS      60
#define DRAG_SCROLL_EDGE    24   // how close to the end of the track counts as "against it"

// The timer that does it. A drag is an event with a start and an end, so this is armed and killed by
// the gesture rather than left running - and it is the pointer's *position* it samples, which is a
// continuous quantity, not the kind of come-and-go state that has to be heard rather than polled.
#define ID_DRAG_SCROLL   1

// Holding a scroll chevron down keeps scrolling, the way a scrollbar arrow does.
//
// Two intervals, because one would be wrong in both directions: long enough not to fire on an
// ordinary click, and the row would crawl; short enough to cross a long row, and every click would
// scroll twice. So nothing happens for CHEVRON_HOLD_MS - which is comfortably longer than a click -
// and after that a step every CHEVRON_REPEAT_MS.
//
// The click itself still happens on RELEASE and is unchanged. Acting on the press would have been the
// smaller change, but it would have taken the press-and-slide-off cancel away from these two buttons
// and left the close button as the only one in the strip that still had it.
#define ID_CHEVRON_REPEAT  2
#define CHEVRON_HOLD_MS    400
#define CHEVRON_REPEAT_MS   90

// How far a press has to travel before it is a drag rather than a click. Four logical pixels is what
// Windows itself uses (SM_CXDRAG's default), but taken as our own scaled constant rather than read
// from the system: SM_CXDRAG is not per-monitor DPI scaled, so on this 150% rig it would be a third
// smaller than everything else in this file, and the check script mirrors these numbers.
#define DRAG_LOGICAL_SLOP   4

// How far *out* of the row a carried tab has to go before letting go of it means "take this document
// out of the stack" instead of "put it here". A whole row-height clear of the strip, above or below.
//
// Deliberately many times the reorder slop, and it is the one threshold in this file chosen for the
// cost of being wrong rather than for how it feels: a reorder aimed at the wrong slot is fixed by
// dragging again, but there is no gesture that puts a torn-off window back, so a tear the user did
// not mean is a window they have to go and find. Written as the strip's own height because that is
// what it means - one row clear of the row - and not as a number that happens to equal it today.
#define TEAROFF_LOGICAL_SLOP  STRIP_LOGICAL_H

// The tab row is bounded independently of the stack. 128 tabs at the 70px minimum is wider than any
// monitor sold, so a layout array larger than this could only describe tabs nobody can see - and an
// unbounded one on the stack of a WM_MOUSEMOVE handler is a different kind of problem.
#define MAX_TABS         128

// The look. Also logical pixels, also scaled per window.
//
// Not one of these appears in ComputeLayout. That is deliberate and it is the thing that made this
// slice safe to do: the restyle changes what is drawn *inside* the rectangles and never the
// rectangles themselves, so tools\WordLayout.cs - the second, hand-maintained copy of ComputeLayout
// that the check scripts click through - needed no edit at all, and the 224 checks that were passing
// before this slice are still measuring the same things afterwards.
#define TAB_LOGICAL_RADIUS    6   // the rounded top corners of a tab card
#define TAB_LOGICAL_INSET     1   // a card is drawn narrower than its rect, so cards do not touch
#define CHIP_LOGICAL_RADIUS   4   // the close and new-document buttons' hover chip
#define LIFT_LOGICAL_SPREAD   4   // how far the carried tab's shadow reaches past it
#define LIFT_ALPHA           64   // and how dark it is where it is darkest

// The x and the + are drawn as strokes rather than typed as characters, for the reason at
// DrawGlyphLines. A stroke of exactly one physical pixel is what made them read as placeholder
// scratches at 150%: everything around them scales and they did not. Tenths of a logical pixel,
// because 1 is too thin and 2 is a felt tip, and this number is ours alone - nothing mirrors it.
#define GLYPH_LOGICAL_STROKE_TENTHS 12

// The unsaved-changes dot, drawn in the close button's own square when the pointer is somewhere
// else. Radius, not diameter. The x's arms reach 4 logical px from the centre in each direction, so
// a radius of 3 puts the dot inside the same optical square and at about the same weight of ink -
// tuned against a photograph rather than reasoned about, because a filled disc reads heavier than
// two strokes at equal size. Nothing mirrors this number: the *rectangle* is the close button's, and
// that is the one the check scripts know.
#define DOT_LOGICAL_RADIUS 3

// The tooltip that appears when the pointer rests on a tab.
//
// It exists because the name is the only thing on a tab that identifies its document - that was the
// finding that closed the icons slice - and the name is exactly what runs out of room first. A
// default Word window fits five tabs; the sixth document makes every name shorter, and past that
// point the row is full of `Qua...` and `Draf...` with nothing anywhere that says which is which.
//
// Two timers rather than one, and both intervals come from the system rather than from taste.
// Windows' own tooltip delays are defined in terms of the double-click time: a tooltip appears after
// GetDoubleClickTime() and takes itself away after ten times that. Using those means the strip's
// tooltip arrives when the user's other tooltips arrive, on a machine whose owner may have changed
// that setting for a reason.
#define ID_TIP_SHOW  3
#define ID_TIP_HIDE  4

#define TIP_LOGICAL_PADX  10
#define TIP_LOGICAL_PADY   6
#define TIP_LOGICAL_GAP    2     // between the name and the folder under it
#define TIP_LOGICAL_MAXW 520     // a deep path is cut in the middle, not allowed to cross the screen
#define TIP_LOGICAL_DROP   2     // how far below the strip it hangs

// What the pointer is over, or what a button press is claiming. Used for both, which is why HIT_TAB
// appears as a press kind: it means a *middle* press, since a left press on a tab acts immediately
// and never waits for a release.
enum { HIT_NONE = 0, HIT_TAB, HIT_CLOSE, HIT_PLUS, HIT_PREV, HIT_NEXT };

// Posted to a strip so that a command runs after the handler that raised it has returned. Calling
// into Word's object model from inside our own window procedure would re-enter it: Documents.Add
// creates a window and pumps messages, Document.Save can put up a dialog that runs its own message
// loop, and some of those messages are ours.
//
//   WM_WORDTAB_CMD   wParam = CMD_*, lParam = the tab it is about (NULL for the ones that are not)
//   WM_WORDTAB_SAVE  lParam = the tab, once it has been activated and Word has processed that
#define WM_WORDTAB_CMD   (WM_APP + 10)
#define WM_WORDTAB_SAVE  (WM_APP + 11)

// The menu's command ids, which are also the ids TrackPopupMenu hands back. They start at 1 because
// TPM_RETURNCMD answers 0 for "the user dismissed it without choosing".
enum { CMD_NEW = 1, CMD_SAVE, CMD_CLOSE, CMD_CLOSE_OTHERS, CMD_CLOSE_ALL, CMD_CLOSE_RIGHT,
       CMD_TEAROFF, CMD_JOIN };

static const wchar_t* const kWwfClass    = L"_WwF";
static const wchar_t* const kStripClass  = L"WordTabStrip";
static const wchar_t* const kTipClass    = L"WordTabTip";
static const wchar_t* const kGhostClass  = L"WordTabGhost";

// Our link in `_WwF`'s subclass chain. Distinct from the frame's - see frames.cpp.
static const UINT_PTR kWwfSubclassId = 0x57544143;   // 'WTAC'

// ---------------------------------------------------------------------------------------------
// Per-frame state.
//
// One entry per OpusApp frame we have subclassed. Kept here rather than in frames.cpp's table so
// that file stays about the frame subclass and this one about geometry; they are joined only by
// the four Strip* calls declared in wordtab.h.
//
// All of it is touched on Word's UI thread only.
// ---------------------------------------------------------------------------------------------

struct StripState
{
    HWND  frame;         // Word's OpusApp
    HWND  wwf;           // Word's document frame, subclassed by us
    HWND  strip;         // ours, a child of the frame
    HWND  tip;           // ours, a popup owned by the frame; created the first time one is needed
    HWND  ghost;         // ours, the tab card carried under the pointer once the row has let go

    int   stripH;        // STRIP_LOGICAL_H scaled to this window's DPI
    int   dpi;
    HFONT font;

    // The back buffer, and it is a 32-bit DIB rather than a compatible bitmap because the strip now
    // composites: anti-aliased corners and the carried tab's shadow are coverage arithmetic against
    // whatever is already underneath them, and a screen-compatible bitmap has no channel to do that
    // in. Kept for the life of the strip rather than made per paint - DragMove repaints every strip
    // in the stack on every mouse movement, and a CreateDIBSection per movement is a cost paid
    // inside Word's own input loop. Rebuilt when the strip changes size or DPI.
    HDC     memDc;
    HBITMAP dib;
    HBITMAP dibOld;      // what the memory DC came with, put back before the DC is destroyed
    BYTE*   bits;        // BGRA, top-down, owned by the DIB section
    int     dibW, dibH;

    BOOL  enabled;       // cleared while restoring, so the handler stops rewriting
    BOOL  hasApplied;
    RECT  natural;       // where Word wants `_WwF`, in frame client coordinates
    RECT  applied;       // where we put it: natural, top pushed down by stripH

    BOOL  stripPlaced;
    RECT  stripAt;
    SIZE  clientAtNatural;   // the frame's client size when `natural` was recorded - see StripRefit

    BOOL  trippedLogged; // the height tripwire says its piece once per frame, not once per message
    BOOL  wasVisible;    // to catch hidden -> shown, where the layout has to be re-derived

    // Logging a relayout costs a file write, and Word relayouts on every mouse movement of a resize
    // drag. So they are throttled and counted, and the count is reported with the next line that
    // does get written - the alternative is either a stutter the user can feel or a silence that
    // hides how often this happens.
    DWORD lastLogTick;
    int   suppressed;
    LONG  lastLoggedTop;  // the throttle may drop a width change; it may not drop a top-edge move

    // What the pointer is over and what is being pressed, held per strip rather than in one global.
    // Only the active window's strip is on top, so only it receives mouse messages - but a strip
    // that was hovered and then covered would otherwise keep drawing a highlight under a pointer
    // that is somewhere else entirely, and show it again the moment its tab came forward.
    //
    // Frames, not indices. The tab row can change between a press and its release - that is exactly
    // what closing a tab does - and an index that meant one document on the way down can mean a
    // different one on the way up.
    int   hotKind;
    HWND  hotFrame;
    int   pressKind;      // a left press on a close or new button, waiting for its release
    HWND  pressFrame;
    BOOL  pressRepeated;  // ...and, on a scroll chevron, whether holding it has already scrolled
    HWND  middleFrame;    // a middle press on a tab, likewise
    BOOL  rightDown;      // a right press, waiting to become a context menu on release
    HWND  rightFrame;     // ...and the tab it landed on, NULL for the empty part of the strip
    HWND  menuFrame;      // the tab a context menu is open for, kept lit while it is up
    BOOL  tracking;       // TrackMouseEvent armed, so WM_MOUSELEAVE will arrive

    wchar_t title[256];

    // Has this frame's title been seen collapsed to Word's placeholder without a decision about it
    // yet? Cleared the moment a real title is taken. See ReadTitle for what the two states are for.
    BOOL namelessSeen;

    // Does this window's document have unsaved changes? Read from Word's object model by the
    // janitor, cached here and compared against last tick exactly as the title is.
    //
    // On the strip rather than on the stack's Member, and that is not arbitrary: StackTabs answers
    // with the frame as its own single tab when stacking is off or a member has not joined, so a
    // flag living on Member would be correct everywhere except with Stack=0, where it would simply
    // never appear. StripState is the thing that exists whenever there is a tab to draw.
    //
    // FALSE until Word has been asked, which is what StripAttachFrame's memset gives for free, and
    // the honest starting value: no dot is what a document that has not been measured looks like.
    BOOL modified;

    // Whether each tab's name had to be cut to fit, written by the one function that draws a name
    // and read by the tooltip. Recorded where the cutting happens rather than worked out a second
    // time from a rectangle that might not be the same one - the tab's text box is derived from the
    // close button's position, and a second copy of that derivation would be a second answer to
    // "does this name fit". Indexed by position in the row, like everything else the layout knows.
    BOOL nameCut[MAX_TABS];
};

// Where every clickable thing in a strip is, in the strip's client coordinates. One structure
// computed by one function and used by both painting and hit-testing, so a click can never land
// somewhere other than what was drawn.
struct StripLayout
{
    int  count;
    RECT tab[MAX_TABS];
    RECT close[MAX_TABS];   // empty when the tab is too narrow to carry a button honestly, or when
                            // it would fall outside the track and so be drawn on nothing
    RECT plus;
    BOOL hasPlus;

    // The band the tabs live in. Everything to the right of it - the two scroll buttons and the
    // new-document button - is a fixed cluster that the row is never allowed to reach, which is the
    // whole of this slice: `plus` and `tab[i]` cannot intersect, by construction rather than by a
    // clamp that gives up when it runs out of room.
    RECT track;
    RECT prev, next;        // the scroll buttons; empty unless the row overflows
    BOOL hasNav;
    BOOL canPrev, canNext;  // ...and whether there is anywhere left to go in that direction

    int  scroll;            // how far the row has been carried left, in pixels. 0 unless overflowing
    int  maxScroll;
    int  width;             // one tab's pitch, which is also one step of the scroll
};

// What a point in a strip is over.
struct StripHit
{
    int  kind;
    HWND frame;             // for HIT_TAB and HIT_CLOSE
    int  index;             // ...and where in the row it is, -1 otherwise
    RECT tab;               // ...and the tab's rectangle, so picking one up needs no second layout
};

static StripState g_strips[MAX_STRIPS];
static int  g_stripCount = 0;
static BOOL g_stripEnabled = TRUE;
static BOOL g_buttonsEnabled = TRUE;     // HKCU\Software\WordTab\TabButtons
static BOOL g_menuEnabled = TRUE;        // HKCU\Software\WordTab\TabMenu
static BOOL g_dragEnabled = TRUE;        // HKCU\Software\WordTab\TabDrag
static BOOL g_lookEnabled = TRUE;        // HKCU\Software\WordTab\TabStyle
static BOOL g_sampleEnabled = TRUE;      // HKCU\Software\WordTab\TabThemeSample
static BOOL g_scrollEnabled = TRUE;      // HKCU\Software\WordTab\TabScroll
static BOOL g_titleTrimEnabled = TRUE;   // HKCU\Software\WordTab\TabTitleTrim
static BOOL g_dotEnabled = TRUE;         // HKCU\Software\WordTab\TabDot
static BOOL g_tipEnabled = TRUE;         // HKCU\Software\WordTab\TabTip
static BOOL g_ghostEnabled = TRUE;       // HKCU\Software\WordTab\TabGhost

// HKCU\Software\WordTab\TabFontSize - the label's size in logical pixels, absent meaning the default.
// A number rather than a switch, and offered because "small enough to read the name, large enough to
// read at all" is a judgement about one person's monitor and eyes, not a fact this file can hold. The
// work rig is a laptop beside a 40" panel and nobody here can see either of them.
static int g_fontLogical = TAB_LOGICAL_FONT;

// Set once at StripStart, and only when HKCU\Software\WordTab\TabDpi already held a usable value.
// It makes the janitor watch for that value changing; see the janitor for why the real event cannot
// be used on a one-monitor rig. FALSE on every machine that has never set TabDpi, which is all of
// them, so nothing below it runs in an ordinary install.
static BOOL g_dpiWatch = FALSE;

// ---------------------------------------------------------------------------------------------
// How far the row is scrolled.
//
// Global, like the drag and for the same reason: every window in the stack draws the same row, and a
// row that stood at a different scroll position in each of them would stop being one row the moment
// there were enough documents for it to matter. One number, every strip.
//
// It is clamped by ComputeLayout rather than by whoever moved it, which is what makes it
// self-healing: widen the window and the next paint discovers there is less to scroll and shortens
// it, with nothing having to notice the resize. ComputeLayout is the only writer, and it only writes
// when the strip it is laying out actually has tabs - a window with no document paints an empty row
// too, and letting that one reset the number would scroll the real row back to the start every time
// anything repainted.
//
// g_scrollShown is the tab that was active when the row was last scrolled to reveal one. Revealing is
// an event - "the active document changed" - not a rule applied on every layout: applied on every
// layout it would drag the row back to the active tab a frame after the user scrolled away from it,
// and the wheel would appear not to work.
// ---------------------------------------------------------------------------------------------

static int  g_scroll      = 0;
static HWND g_scrollShown = NULL;
static int  g_scrollMax   = -1;    // the row's shape last time it was revealed against

// ---------------------------------------------------------------------------------------------
// A tab being dragged.
//
// Global, unlike hover and the pressed buttons, which are per strip. The reason is that a press on a
// tab activates that document *before* the drag begins - and activating raises a different window,
// whose strip is now the one in front. So the strip holding the mouse capture is very often not the
// strip the user can see. Per-strip state would draw the tab travelling on a window that is
// underneath another one, and nothing at all on the window they are looking at.
//
// Held as the add-in's state instead, every strip can draw the same carried tab, and the one on top
// is by construction the one that shows it. They are all the same width at the same position - that
// is what the stack guarantees - so one x coordinate is meaningful in all of them.
//
// g_dragStrip doubles as "this strip is holding the mouse capture for a tab gesture". g_dragFrame is
// cleared on cancel while the capture is kept, because the left button is still down and letting go
// of the mouse mid-gesture would deliver the release to whatever is underneath.
// ---------------------------------------------------------------------------------------------

static HWND g_dragStrip  = NULL;    // the strip that owns the capture, NULL when no press is held
static HWND g_dragFrame  = NULL;    // the tab under that press, NULL once cancelled
static int  g_dragPressX = 0;       // where the press landed, in strip client coordinates
static int  g_dragPressY = 0;       // ...and its y, which only the rejoin gesture has a use for
static int  g_dragGrabDx = 0;       // how far into the tab, so it does not jump when picked up
static int  g_dragLeft   = 0;       // the carried tab's left edge right now
static int  g_dragFrom   = 0;       // the position it was picked up from - the log, and the undo
static BOOL g_dragging   = FALSE;   // past the slop: this is a drag, not a click that has not ended
static BOOL g_dragTorn   = FALSE;   // carried clear of the row: letting go now takes it out of the stack
static BOOL g_dragJoin   = FALSE;   // the other gesture: a window on its own being carried INTO a stack
static HWND g_dragOnto   = NULL;    // the frame whose row it is over, NULL when it is over nothing
static BOOL g_dragOntoKnown = FALSE; // whether g_dragOnto has been worked out yet this gesture
static HWND g_dragScroll = NULL;    // the strip running the auto-scroll timer, NULL when it is off
static ATOM g_stripClass = 0;
static UINT_PTR g_janitor = 0;

// ---------------------------------------------------------------------------------------------
// The palette.
//
// One structure derived from one colour: the one Word is painting immediately above us. Every value
// below is a fixed step from it, and the steps are measurements rather than taste - see
// DerivePalette. That is what replaced the two hand-picked triples this file used to carry, and it
// is why Colorful, White, Dark Grey, Black and whatever Office ships next all come out right
// without a table listing them.
// ---------------------------------------------------------------------------------------------

struct Palette
{
    COLORREF chrome;      // what it was all derived from: Word's ribbon
    BOOL     dark;

    COLORREF back;        // the well the tabs sit in
    COLORREF selected;    // the active tab's card
    COLORREF lifted;      // ...and the same card while it is being carried
    COLORREF hover;       // an inactive tab under the pointer
    COLORREF edge;        // the card's border, and the hairline along the bottom
    COLORREF separator;   // between two inactive tabs
    COLORREF text;        // the active tab's name
    COLORREF textIdle;    // the others, and the empty row's message
    COLORREF glyph;       // the x and the +
    COLORREF glyphHot;
    COLORREF chip;        // a button's hover chip
    COLORREF chipDown;

    COLORREF menuBack;    // the context menu, which is ours to draw now
    COLORREF menuText;
    COLORREF menuTextDim;
    COLORREF menuHot;
    COLORREF menuLine;
};

static Palette g_palette;
static BOOL    g_paletteReady = FALSE;

// The two GDI objects the palette still needs as handles: everything else is composited by hand.
static HBRUSH g_backBrush     = NULL;   // the flat fallback renderer's ground
static HBRUSH g_menuBackBrush = NULL;   // SetMenuInfo's MIM_BACKGROUND, which is not ours to paint

// Which of our strips a window handle is, if it is one of ours at all.
//
// Answered by looking in our own array rather than by asking the window - a class-name check on a
// handle that came from WindowFromPoint would believe anything that registered the same class, and
// GWLP_USERDATA on a foreign window is somebody else's pointer being read as ours. These are exactly
// the strips this process made, so the question cannot be answered wrongly.
static StripState* FindByStrip(HWND strip)
{
    if (!strip)
        return NULL;
    for (int i = 0; i < g_stripCount; i++)
        if (g_strips[i].strip == strip)
            return &g_strips[i];
    return NULL;
}

static StripState* FindByFrame(HWND frame)
{
    for (int i = 0; i < g_stripCount; i++)
        if (g_strips[i].frame == frame)
            return &g_strips[i];
    return NULL;
}

// Does the document behind this tab have unsaved changes? Every strip draws every tab, so the answer
// belongs to the frame being drawn and not to the strip doing the drawing - which is why this is a
// lookup rather than a field read. FALSE for a frame with no strip of its own, and for every frame
// while the dot is switched off, so one test here turns the whole feature off in both renderers.
static BOOL FrameModified(HWND frame)
{
    if (!g_dotEnabled)
        return FALSE;
    StripState* owner = FindByFrame(frame);
    return (owner && owner->modified) ? TRUE : FALSE;
}

// ---------------------------------------------------------------------------------------------
// DPI. This rig runs at 150%, so a hard-coded 32 is 32 physical pixels here - two thirds of the
// height it should be. GetDpiForWindow is Windows 10 1607 and is reached through GetProcAddress
// rather than a header, because w64devkit's headers are older than the API and the failure would
// be a link error at the end of a long build.
// ---------------------------------------------------------------------------------------------

typedef UINT (WINAPI *GetDpiForWindowFn)(HWND);

// HKCU\Software\WordTab\TabDpi - absent or 0 means "ask Windows", which is every real install.
//
// This exists because until it did, every scaled size in this file had only ever been computed at
// one number. The dev rig runs at 192 and Windows offers exactly one scale factor for its display,
// so 96 and 144 were unreachable here - measured, not assumed: the private DisplayConfig scaling
// packet answers with min == cur == max. The work rig is a laptop plus a 40" second monitor, so it
// is the machine where a second DPI first happens, and it is the machine nobody here can reach.
//
// Deliberately read on every call rather than cached at StripStart like the other switches, and the
// re-reading is the whole point rather than an accepted cost: it is what lets a value written into
// the registry mid-run change the answer, which is the only way StripOnFrameDpiChanged can be
// reached on a rig with one monitor. That path had never executed on any rig before this.
//
// Note what this does NOT do: it does not let the real message be faked. WM_DPICHANGED cannot be
// posted from another process at all - see the janitor for why - so the override is the way in, and
// the janitor noticing the value change is the trigger that stands in for the message.
//
// Three call sites: a strip being built, a DPI change, and the janitor's watch. The first two are
// rare. The third runs twice a second per strip, but only when TabDpi was already set when Word
// started, so no ordinary install ever reaches it and none of them pays for this read.
static int DpiOverride(void)
{
    DWORD forced = WordTabReadNumber(L"TabDpi", 0);
    if (forced >= 72 && forced <= 480)
        return (int)forced;
    return 0;
}

static int DpiOf(HWND hwnd)
{
    int forced = DpiOverride();
    if (forced)
        return forced;

    static GetDpiForWindowFn fn = NULL;
    static BOOL looked = FALSE;
    if (!looked)
    {
        looked = TRUE;
        HMODULE user32 = GetModuleHandleW(L"user32.dll");
        if (user32)
            fn = (GetDpiForWindowFn)(void*)GetProcAddress(user32, "GetDpiForWindow");
    }

    if (fn)
    {
        UINT dpi = fn(hwnd);
        if (dpi >= 72 && dpi <= 480)
            return (int)dpi;
    }

    // Pre-1607, or a window the system will not answer for: the system DPI is close enough, and on
    // a single-monitor rig it is identical.
    HDC screen = GetDC(NULL);
    int dpi = screen ? GetDeviceCaps(screen, LOGPIXELSY) : 96;
    if (screen)
        ReleaseDC(NULL, screen);
    return dpi > 0 ? dpi : 96;
}

static int Scaled(int logical, int dpi)
{
    return MulDiv(logical, dpi, 96);
}

// The stack asks this file what the DPI is rather than asking Windows itself. Two reasons, and the
// second is the one that matters: the two files cannot then disagree, and TabDpi - which is the only
// way to reach a second DPI on a one-monitor rig - moves both of them at once, so the page-width
// arithmetic in stack.cpp can be driven from a check suite on a machine that has one screen.
int StripDpiOf(HWND frame)
{
    return DpiOf(frame);
}

// ---------------------------------------------------------------------------------------------
// Theme.
//
// The strip used to carry two hand-picked palettes, light and dark, chosen by eye and selected
// between by reading Office's `UI Theme` registry value. Two things measured for this slice retired
// that arrangement.
//
// **The registry value is not a reliable oracle.** It is legitimately 6 - "use system setting" - on
// a default install, so the answer is somewhere else anyway; and Word *rewrites it at startup* from
// the roaming account setting. Writing 5 (White) and launching Word produced a Word that read back
// 6. A value we read once while the add-in is loading may be the previous session's answer.
//
// **The colour itself is readable, and exactly.** Word paints a flat band along the bottom of the
// ribbon, immediately above our strip. Sampled from the ribbon window's own device context it comes
// back at 98-99% of a 150-pixel scan: RGB(41,41,41) with Windows dark, RGB(255,255,255) with
// Windows light. So the palette is derived from what Word is actually painting, and every Office
// theme - including ones that do not exist yet - comes out right without a table listing them.
//
// The registry read stays as the fallback for when the sample cannot be taken.
// ---------------------------------------------------------------------------------------------

static DWORD ReadDword(HKEY root, const wchar_t* key, const wchar_t* name, DWORD fallback)
{
    DWORD value = 0;
    DWORD size = sizeof(value);
    if (RegGetValueW(root, key, name, RRF_RT_REG_DWORD, NULL, &value, &size) != ERROR_SUCCESS)
        return fallback;
    return value;
}

static BOOL DarkThemeInUse(void)
{
    // 0 colorful, 3 dark grey, 4 black, 5 white, 6 follow Windows. Colorful has a coloured title
    // bar but a light ribbon, so it belongs with the light palette.
    DWORD theme = ReadDword(HKEY_CURRENT_USER,
                            L"Software\\Microsoft\\Office\\16.0\\Common", L"UI Theme", 6);

    if (theme == 3 || theme == 4)
        return TRUE;
    if (theme == 0 || theme == 5)
        return FALSE;

    return ReadDword(HKEY_CURRENT_USER,
                     L"Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize",
                     L"AppsUseLightTheme", 1) == 0;
}

// Rec. 709 luma, which is the one that matches how bright a colour looks rather than how much ink
// it is. It decides one thing - whether this is a light Word or a dark one - and that decision
// flips the direction of every step below.
static int Luma(COLORREF c)
{
    return (GetRValue(c) * 54 + GetGValue(c) * 183 + GetBValue(c) * 19) >> 8;
}

// A fixed number of levels toward white (positive) or black (negative), per channel, clamped. Steps
// rather than percentages: Word's own steps are absolute, and a percentage of 41 is not a step at
// all.
static int Clamp255(int v)
{
    if (v < 0)   return 0;
    if (v > 255) return 255;
    return v;
}

static COLORREF Step(COLORREF c, int delta)
{
    return RGB(Clamp255(GetRValue(c) + delta),
               Clamp255(GetGValue(c) + delta),
               Clamp255(GetBValue(c) + delta));
}

static COLORREF Mix(COLORREF a, COLORREF b, int percentB)
{
    int r = (GetRValue(a) * (100 - percentB) + GetRValue(b) * percentB) / 100;
    int g = (GetGValue(a) * (100 - percentB) + GetGValue(b) * percentB) / 100;
    int bl = (GetBValue(a) * (100 - percentB) + GetBValue(b) * percentB) / 100;
    return RGB(r, g, bl);
}

// Everything from one colour.
//
// The one number that is not arbitrary is 26, and it is the measurement this whole scheme rests on:
// the step Word itself takes from the ribbon to the workspace below it. Light, that is 255 -> 228.
// Dark, 41 -> 9. Twenty-six and thirty-two - call it 26 in both directions, and the well our tabs
// sit in lands on Word's own workspace grey without being told what it is.
//
// The active tab is then the chrome colour exactly: it is a piece of the ribbon, brought down.
static void DerivePalette(COLORREF chrome, Palette* out)
{
    BOOL dark = (Luma(chrome) < 128);

    out->chrome   = chrome;
    out->dark     = dark;

    out->selected = chrome;
    out->back     = Step(chrome, -26);
    out->hover    = Step(chrome, -13);

    // A picked-up tab is lifted by a shadow, and a shadow is black - which is worth nothing at all
    // on a dark Word, where the well behind it is already RGB(15,15,15). Photographed: the lift was
    // invisible. So in a dark theme the card is *raised* as well, and in a light one it is not,
    // because the card there is already white and there is nowhere to raise it to. The shadow is
    // drawn in both; it only earns its keep in one.
    out->lifted = dark ? Step(chrome, +14) : chrome;

    // A border has to go the other way in a dark theme or it is not a border. Light Word gets a
    // definite edge because a white card on a near-white ground needs one; dark Word needs less,
    // because the card is already lighter than everything around it.
    out->edge      = dark ? Step(chrome, +26) : Step(chrome, -45);
    out->separator = dark ? Step(out->back, +22) : Step(out->back, -22);

    out->text      = dark ? RGB(255, 255, 255) : RGB(32, 31, 30);
    out->textIdle  = Mix(out->text, out->back, 38);
    out->glyph     = out->textIdle;
    out->glyphHot  = out->text;
    out->chip      = dark ? Step(chrome, +40) : Step(chrome, -50);
    out->chipDown  = dark ? Step(chrome, +62) : Step(chrome, -67);

    // The menu is a floating surface, not part of the band, so it sits a step *above* the chrome
    // rather than below it - which is what every menu in Windows does against its own window.
    out->menuBack    = dark ? Step(chrome, +2) : RGB(255, 255, 255);
    out->menuText    = out->text;
    out->menuTextDim = Mix(out->text, out->menuBack, 45);
    out->menuHot     = dark ? Step(out->menuBack, +22) : Step(out->menuBack, -18);
    out->menuLine    = dark ? Step(out->menuBack, +34) : Step(out->menuBack, -30);
}

// The colour Word is painting immediately above the strip, read from the ribbon's own device
// context.
//
// Three ways of taking this sample were measured and only one of them tells the truth:
//
//   - the *frame's* client DC answers CLR_INVALID at every pixel. The frame is WS_CLIPCHILDREN and
//     everything up there belongs to a child, so there is nothing of the frame's own to read.
//   - the *screen* DC works while Word is in front and lies when it is not. With Notepad maximised
//     over Word it returned RGB(39,39,39) - Notepad's background - against Word's real
//     RGB(41,41,41). A wrong answer two levels away from the right one is worse than no answer.
//   - the *ribbon's own* window DC returned RGB(41,41,41) at 98% of the scan, unchanged, with
//     Notepad maximised on top of it. That is the one.
//
// A scan and a mode rather than a single pixel: one GetPixel lands on a separator or the edge of a
// button often enough to matter, and a band that is genuinely flat says so by agreeing with itself.
static BOOL ChildRect(HWND parent, HWND child, RECT* out);   // defined with the geometry helpers

// Finding the ribbon among the frame's descendants: the widest visible NetUIHWND that runs from the
// top of the window down to the strip's top edge. By position rather than by order, because a Word
// frame has three of them - the ribbon, the vertical scrollbar and the status bar - and their order
// is not ours to rely on.
//
// **Both edges are checked, and the top one is not decoration.** This rule used to be "whose bottom
// edge is the strip's top edge", which reads as "the thing directly above us" and is a different
// claim from "the ribbon". Measured on a Protected View window: its MsoCommandBarDock holds TWO
// MsoCommandBar chains, the ribbon at y 0..156 and the yellow message bar at y 156..318, and the
// strip goes under the *second* one. So the old rule sampled the message bar - RGB(109,87,0) on a
// rig whose Word is RGB(41,41,41) - and because the palette is one stack-wide object, one downloaded
// document turned every tab in every window yellow.
//
// The window is then unable to answer, which is the honest outcome and not a fallback: the same
// measurement scanned that window's whole chrome band and found no ribbon colour anywhere in it
// (RGB(9,9,9) for all 64 rows above the message bar). There is nothing there to sample. Saying so
// lets the poll move on to a window that can answer, and lets the palette keep what it has if none
// can. Nothing here is specific to Protected View: any bar Word puts between its ribbon and the
// document has the same shape and gets the same answer.

// The one message that names this structural case. A literal with a name rather than an inline
// string, because the poll compares reasons by POINTER to decide which one to report - see the
// precedence in the janitor - and because check-dot matches on it.
static const wchar_t kBandDoesNotReachTop[] =
    L"something is docked between Word's ribbon and this window's strip - the ribbon cannot be sampled here";

struct RibbonHunt
{
    HWND frame;
    int  clientW;
    LONG stripTop;
    HWND found;         // the ribbon: directly above the strip AND starting at the top of the window
    HWND intruder;      // directly above the strip but starting further down - a bar of some kind
};

static BOOL CALLBACK RibbonHuntProc(HWND child, LPARAM param)
{
    RibbonHunt* hunt = (RibbonHunt*)param;

    wchar_t cls[32];
    if (!GetClassNameW(child, cls, 32) || wcscmp(cls, L"NetUIHWND") != 0)
        return TRUE;
    if (!IsWindowVisible(child))
        return TRUE;

    RECT at;
    if (!ChildRect(hunt->frame, child, &at))
        return TRUE;

    if ((at.right - at.left) * 10 < hunt->clientW * 6)
        return TRUE;                                    // too narrow to be the ribbon
    LONG gap = at.bottom - hunt->stripTop;
    if (gap < -4 || gap > 4)
        return TRUE;                                    // not the thing directly above us

    // Word draws its own title bar inside the client area, so the ribbon's chain starts at y=0.
    // Anything directly above the strip that does *not* is something Word has put in between.
    if (at.top > 4)
    {
        hunt->intruder = child;
        return TRUE;
    }

    hunt->found = child;
    return FALSE;
}

//
// It reports why it could not answer as well as whether it could. That is not decoration: the first
// version of this returned a bare FALSE, the palette silently stayed on its fallback, and because the
// fallback and the sample agree on this rig the only symptom was a theme change that did not take -
// three steps away from the cause.
static BOOL SampleChrome(StripState* state, COLORREF* out, const wchar_t** why)
{
    *why = L"ok";

    if (!state->frame || !IsWindow(state->frame) || !state->strip)
    {
        *why = L"no frame or no strip";
        return FALSE;
    }

    RECT client;
    if (!GetClientRect(state->frame, &client))
    {
        *why = L"the frame has no client rect";
        return FALSE;
    }
    int clientW = client.right - client.left;
    if (clientW < 200)
    {
        *why = L"the frame is too narrow";
        return FALSE;
    }

    RECT stripAt;
    if (!ChildRect(state->frame, state->strip, &stripAt))
    {
        *why = L"the strip has no rect";
        return FALSE;
    }

    // The ribbon: a direct NetUIHWND child of the frame, as wide as the window, whose bottom edge is
    // where our top edge is. The frame has three NetUIHWNDs - the ribbon, the vertical scrollbar and
    // the status bar - and this picks the ribbon out of them by position rather than by order.
    RibbonHunt hunt;
    hunt.frame       = state->frame;
    hunt.clientW     = clientW;
    hunt.stripTop    = stripAt.top;
    hunt.found       = NULL;
    hunt.intruder    = NULL;

    // EnumChildWindows, not a walk of GetWindow(GW_CHILD)/GW_HWNDNEXT, and that distinction cost a
    // build to find. **The ribbon is not a child of the frame.** It is the innermost of a chain -
    // MsoCommandBarDock, MsoCommandBar, MsoWorkPane, NUIPane, NetUIHWND - five windows deep, all
    // reporting the *same* rectangle, which is what made the mistake so plausible: the enumeration
    // the check scripts use is recursive, so the ribbon looked like a direct child in every listing
    // taken while this was being designed. A sibling walk found nothing at all and said so, in a
    // failure that looked exactly like a sampler that simply agreed with the fallback.
    //
    // Only the innermost of that chain is worth sampling anyway: the outer four are WS_CLIPCHILDREN
    // and their device contexts exclude every pixel that belongs to a child, which up there is all
    // of them. That is the same reason the frame's own DC answered CLR_INVALID.
    EnumChildWindows(state->frame, RibbonHuntProc, (LPARAM)&hunt);

    HWND ribbon = hunt.found;
    if (!ribbon)
    {
        // The two reasons are different and must not share a message. Something docked between the
        // ribbon and the strip is a normal, correct state that this window simply cannot answer
        // from; "nothing of the right width is above the strip" means the geometry is not what this
        // code believes and somebody should look.
        //
        // The first message says what was MEASURED - a full-width band above the strip that does not
        // reach the top of the window - rather than naming it a message bar. Protected View is the
        // case it was measured on, but the same shape is anything Word docks in there, and a message
        // that asserts a cause the evidence does not carry is how a wrong diagnosis gets started.
        *why = hunt.intruder ? kBandDoesNotReachTop
                             : L"no NetUIHWND of the right width sits directly above the strip";
        return FALSE;
    }

    RECT ribbonRect;
    if (!GetWindowRect(ribbon, &ribbonRect))
    {
        *why = L"the ribbon has no rect";
        return FALSE;
    }
    int rw = ribbonRect.right - ribbonRect.left;
    int rh = ribbonRect.bottom - ribbonRect.top;
    if (rw < 200 || rh < 12)
    {
        *why = L"the ribbon is too small to sample";
        return FALSE;
    }

    // The screen, and it has to be the screen. This was written the other way first - GetDC on the
    // ribbon itself, so that an occluded Word could still be read - and measured, that DC is a lie:
    //
    //     before the flip        window DC: RGB( 41, 41, 41)    screen: RGB(41,41,41)
    //     3s after the flip      window DC: RGB( 41, 41, 41)    screen: RGB(255,255,255)
    //     12s after the flip     window DC: RGB( 41, 41, 41)    screen: RGB(255,255,255)
    //     after putting it back  window DC: RGB( 41, 41, 41)    screen: RGB(41,41,41)
    //
    // **Word re-renders its ribbon somewhere GDI cannot follow.** The redirection surface behind that
    // HWND keeps whatever was last drawn into it by GDI and never changes again, so the window DC
    // answers correctly exactly once - at startup - and then goes stale for the life of the process.
    // Which is the one case that mattered, because a palette that is only right at startup is the
    // registry read this was meant to replace.
    //
    // The screen DC tells the truth, and its hazard is the opposite one: it reads whatever is on
    // top, and with Notepad maximised over Word it returned RGB(39,39,39) against Word's real
    // RGB(41,41,41) - a wrong answer two levels from the right one. So every sample point is asked
    // *who owns this pixel* before it is read, and points that belong to anything but the ribbon are
    // not read at all. That is not a heuristic about how likely occlusion is; it is the question the
    // hazard actually poses, answered per pixel.
    HDC screen = GetDC(NULL);
    if (!screen)
    {
        *why = L"no screen DC";
        return FALSE;
    }

    COLORREF seen[32];
    int      count[32];
    int      kinds = 0, total = 0, foreign = 0;
    int      y = ribbonRect.bottom - 3;
    int      step = (rw / 2) / 24;
    if (step < 1) step = 1;

    for (int i = 0; i < 24; i++)
    {
        POINT pt;
        pt.x = ribbonRect.left + 20 + i * step;
        pt.y = y;

        HWND owner = WindowFromPoint(pt);
        if (owner != ribbon && !IsChild(ribbon, owner))
        {
            foreign++;
            continue;
        }

        COLORREF c = GetPixel(screen, pt.x, pt.y);
        if (c == CLR_INVALID)
            continue;

        total++;
        int found = -1;
        for (int k = 0; k < kinds; k++)
            if (seen[k] == c) { found = k; break; }
        if (found >= 0)
            count[found]++;
        else if (kinds < 32)
        {
            seen[kinds]  = c;
            count[kinds] = 1;
            kinds++;
        }
    }
    ReleaseDC(NULL, screen);

    if (total < 12)
    {
        *why = (foreign > 0) ? L"something is covering Word's ribbon"
                             : L"too few readable pixels";
        return FALSE;
    }

    int best = 0;
    for (int k = 1; k < kinds; k++)
        if (count[k] > count[best])
            best = k;

    // Three quarters of a flat band is flat. Anything less and we are looking at something that is
    // not the ribbon's background - a contextual tab, a mid-repaint, a theme we do not understand -
    // and the honest answer is to keep the palette we already have.
    if (count[best] * 4 < total * 3)
    {
        *why = L"the band under the ribbon is not flat";
        return FALSE;
    }

    *out = seen[best];
    return TRUE;
}

// Adopt a chrome colour, rebuild everything derived from it, and repaint.
//
// Written to be called again, which the old code could not be: every brush there was created behind
// an `if (!brush)` guard, so a second call changed the COLORREFs and kept the first theme's handles.
// That produced a palette half in one theme and half in the other - and specifically a bottom
// hairline that followed the change (it built its brush per paint) above a tab border that did not.
// Nothing here is guarded, everything is deleted before it is remade, and the whole thing is
// compare-before-act so a redundant call costs one comparison.
static BOOL ApplyPalette(COLORREF chrome, const wchar_t* why)
{
    if (g_paletteReady && g_palette.chrome == chrome)
        return FALSE;

    DerivePalette(chrome, &g_palette);
    g_paletteReady = TRUE;

    if (g_backBrush)     { DeleteObject(g_backBrush);     g_backBrush = NULL; }
    if (g_menuBackBrush) { DeleteObject(g_menuBackBrush); g_menuBackBrush = NULL; }
    g_backBrush     = CreateSolidBrush(g_palette.back);
    g_menuBackBrush = CreateSolidBrush(g_palette.menuBack);

    LogWrite(L"strip  palette %s: chrome=RGB(%d,%d,%d) %s  well=RGB(%d,%d,%d) "
             L"card=RGB(%d,%d,%d) edge=RGB(%d,%d,%d)",
             why,
             GetRValue(chrome), GetGValue(chrome), GetBValue(chrome),
             g_palette.dark ? L"dark" : L"light",
             GetRValue(g_palette.back), GetGValue(g_palette.back), GetBValue(g_palette.back),
             GetRValue(g_palette.selected), GetGValue(g_palette.selected), GetBValue(g_palette.selected),
             GetRValue(g_palette.edge), GetGValue(g_palette.edge), GetBValue(g_palette.edge));

    StripRefreshTabs();
    return TRUE;
}

// The palette when no sample can be taken: Word's two measured ribbon colours, chosen between by the
// registry read this file has always done.
static void ApplyFallbackPalette(const wchar_t* why)
{
    ApplyPalette(DarkThemeInUse() ? RGB(41, 41, 41) : RGB(255, 255, 255), why);
}

// ---------------------------------------------------------------------------------------------
// Which sampled colours are believed.
//
// `ApplyPalette` above compares for equality and rebuilds on any difference, which is right for the
// callers that hand it a colour they already know - the theme registry, and a caller that has
// decided. It is wrong for the sampler, and the work rig's log says so twice over in one day. What
// arrives from the sampler is a *measurement*, and this is where a measurement becomes a decision.
// ---------------------------------------------------------------------------------------------

// How far a sampled colour has to move before it is a different colour.
//
// Measured, and it is the whole reason the first half of this gate exists: Word's ribbon reads
// RGB(9,9,9) while its window is active and RGB(10,10,10) while it is not. Every switch between two
// documents therefore moved the palette by one unit in each channel, and equality called that a new
// theme - so one working day's log holds over a hundred palette rebuilds, each destroying and
// remaking two brushes and invalidating every strip in the row, for a difference no display can
// show.
//
// Four, because the Office themes this has to tell apart are nowhere near that close. Black samples
// at 9, dark grey at 41, and the two light ones in the 240s and at 255: the smallest real gap is an
// order of magnitude above this number and the observed jitter an order of magnitude below it.
// There is no value in between that would be a hard call.
#define CHROME_SAME_WITHIN 4

static BOOL ChromeNear(COLORREF a, COLORREF b)
{
    int dr = (int)GetRValue(a) - (int)GetRValue(b);
    int dg = (int)GetGValue(a) - (int)GetGValue(b);
    int db = (int)GetBValue(a) - (int)GetBValue(b);
    if (dr < 0) dr = -dr;
    if (dg < 0) dg = -dg;
    if (db < 0) db = -db;
    return (dr <= CHROME_SAME_WITHIN && dg <= CHROME_SAME_WITHIN && db <= CHROME_SAME_WITHIN)
           ? TRUE : FALSE;
}

// A new colour is adopted on the second reading, never the first.
//
// The dead band handles a palette jittering between two neighbouring values. This handles the other
// thing the log shows: a single sample that is not the ribbon at all. Twice in one day the sampler
// returned RGB(0,0,0) with every reading either side of it RGB(10,10,10), and because ApplyPalette
// believes whatever it is handed, the whole row went a different colour for the two seconds until
// the next sample put it back.
//
// Nothing about that sample is detectable from inside itself. It is flat, it is owned by the ribbon,
// it passes every test SampleChrome makes - the one thing separating it from a real theme change is
// that a real theme change is still true two seconds later. So that is the test, because it is the
// only one there is.
//
// What it costs is that a genuine theme change reaches the strip one sample - two seconds - later
// than it used to. A theme change is a person in File > Account, so that is not a delay anybody is
// waiting on. The FIRST sample of a process is exempt: until then the palette is the registry's
// guess rather than a measurement, and making the first measurement wait would leave a colour that
// may be wrong up for twice as long as it is now, which is the opposite of the point.
static BOOL AdoptSampledChrome(COLORREF chrome, HWND from)
{
    static BOOL     everSampled = FALSE;
    static BOOL     pendingSet  = FALSE;
    static COLORREF pending     = 0;

    // Already showing this colour, near enough. The common case by a very long way, and the one the
    // dead band exists for: it has to cost nothing and say nothing.
    if (g_paletteReady && everSampled && ChromeNear(g_palette.chrome, chrome))
    {
        pendingSet = FALSE;
        return FALSE;
    }

    if (everSampled && (!pendingSet || !ChromeNear(pending, chrome)))
    {
        // Once per candidate rather than once per sample, which matters: a colour that keeps
        // arriving and keeps failing to be confirmed would otherwise write a line every two seconds
        // for as long as Word is open. A second reading either adopts it - and ApplyPalette says so
        // in the line this project already greps for - or is replaced by a different candidate,
        // which writes its own line naming that one.
        LogWrite(L"strip  chrome read as RGB(%d,%d,%d) in 0x%p against a palette of RGB(%d,%d,%d) - "
                 L"waiting for a second reading before believing it",
                 GetRValue(chrome), GetGValue(chrome), GetBValue(chrome), (void*)from,
                 GetRValue(g_palette.chrome), GetGValue(g_palette.chrome),
                 GetBValue(g_palette.chrome));
        pending    = chrome;
        pendingSet = TRUE;
        return FALSE;
    }

    everSampled = TRUE;
    pendingSet  = FALSE;

    wchar_t why[96];
    _snwprintf(why, 96, L"sampled from Word's ribbon in 0x%p", (void*)from);
    why[95] = 0;
    return ApplyPalette(chrome, why);
}

static void MakeFont(StripState* state)
{
    if (state->font)
    {
        DeleteObject(state->font);
        state->font = NULL;
    }

    LOGFONTW lf;
    memset(&lf, 0, sizeof(lf));
    lf.lfHeight  = -Scaled(g_fontLogical, state->dpi);
    lf.lfWeight  = FW_NORMAL;
    lf.lfCharSet = DEFAULT_CHARSET;
    lf.lfQuality = CLEARTYPE_QUALITY;
    wcscpy(lf.lfFaceName, L"Segoe UI");

    state->font = CreateFontIndirectW(&lf);
}

// The back buffer. Freed here and rebuilt on demand rather than resized, because a strip changes
// size rarely and a DIB section cannot be resized anyway.
static void ReleaseSurface(StripState* state)
{
    // The bitmap goes back before the DC does. A bitmap still selected into a device context is not
    // deleted by DeleteObject, it is only marked - and the leak is per strip, inside Word, for the
    // rest of the session.
    if (state->memDc && state->dibOld)
        SelectObject(state->memDc, state->dibOld);
    if (state->dib)   { DeleteObject(state->dib); state->dib = NULL; }
    if (state->memDc) { DeleteDC(state->memDc);   state->memDc = NULL; }

    state->dibOld = NULL;
    state->bits = NULL;
    state->dibW = 0;
    state->dibH = 0;
}

// Everything about this strip that is a function of its DPI, in one place.
//
// It used to be three: StripAttachFrame, TryBind and StripOnFrameDpiChanged each set `dpi` and
// `stripH`, and only two of them remade the font - TryBind guarded it with `if (!state->font)`, so a
// DPI change between attach and bind left the strip at the new height with the old font. That was
// latent while the font was the only thing derived from DPI. This slice derives the corner radius,
// the glyph stroke, the chip radius, the shadow spread and the size of the back buffer from it too,
// so three copies of "what depends on DPI" was three chances to forget one.
static void ApplyMetrics(StripState* state)
{
    state->dpi    = DpiOf(state->frame);
    state->stripH = Scaled(STRIP_LOGICAL_H, state->dpi);
    MakeFont(state);
    ReleaseSurface(state);        // its size is in physical pixels, so it is DPI-dependent too

    // The numbers every other size in this file is derived from. Logged here rather than at the call
    // sites because this is the one place they are set, and a suite that wants to know what scale a
    // strip was built at should not have to infer it from a rectangle. The type size is here too for
    // the same reason: a label that comes out the wrong size on a machine nobody here can see is
    // answered by reading what it was built at, not by measuring a screenshot.
    LogWrite(L"strip  hwnd=0x%p  metrics: dpi=%d  stripH=%dpx (%d logical)  font=%d logical",
             (void*)state->frame, state->dpi, state->stripH, STRIP_LOGICAL_H, g_fontLogical);
}

// ---------------------------------------------------------------------------------------------
// Geometry helpers.
// ---------------------------------------------------------------------------------------------

// A child's rectangle in its parent's client coordinates - the same space WINDOWPOS.x/y use, so
// everything below can be compared without converting twice.
static BOOL ChildRect(HWND parent, HWND child, RECT* out)
{
    RECT screen;
    if (!GetWindowRect(child, &screen))
        return FALSE;

    POINT topLeft;
    topLeft.x = screen.left;
    topLeft.y = screen.top;
    if (!ScreenToClient(parent, &topLeft))
        return FALSE;

    out->left   = topLeft.x;
    out->top    = topLeft.y;
    out->right  = topLeft.x + (screen.right - screen.left);
    out->bottom = topLeft.y + (screen.bottom - screen.top);
    return TRUE;
}

static BOOL SameRect(const RECT* a, const RECT* b)
{
    return a->left == b->left && a->top == b->top &&
           a->right == b->right && a->bottom == b->bottom;
}

// Remember how big the frame was when this natural rect was true. The stack needs it to refit a
// window's document frame to a new size that Word will not lay out for it - see StripRefit.
static void RememberClient(StripState* state)
{
    RECT client;
    if (!GetClientRect(state->frame, &client))
        return;
    state->clientAtNatural.cx = client.right - client.left;
    state->clientAtNatural.cy = client.bottom - client.top;
}

static void LogRelayout(StripState* state, const wchar_t* why)
{
    // The throttle may never swallow a move of the **top** edge. That edge is the entire subject of
    // this file - the strip is drawn exactly at `natural.top` - and the one recorded occasion when it
    // went to the wrong place, the evidence was dropped by this limit and the cause went undiagnosed
    // for two slices. The flood the throttle exists for is a resize drag, which moves the other three
    // edges every ~15ms and leaves the top alone, so nothing is given back by exempting it.
    DWORD now = GetTickCount();
    BOOL topMoved = (state->natural.top != state->lastLoggedTop);
    if (!topMoved && state->lastLogTick != 0 && (now - state->lastLogTick) < 250)
    {
        state->suppressed++;
        return;
    }
    state->lastLogTick   = now;
    state->lastLoggedTop = state->natural.top;

    wchar_t extra[64] = L"";
    if (state->suppressed > 0)
    {
        _snwprintf(extra, 64, L"  [+%d more not logged]", state->suppressed);
        extra[63] = L'\0';
        state->suppressed = 0;
    }

    LogWrite(L"strip  hwnd=0x%p  %s: _WwF natural (%ld,%ld %ldx%ld) -> (%ld,%ld %ldx%ld)  strip h=%d%s",
             (void*)state->frame, why,
             state->natural.left, state->natural.top,
             state->natural.right - state->natural.left,
             state->natural.bottom - state->natural.top,
             state->applied.left, state->applied.top,
             state->applied.right - state->applied.left,
             state->applied.bottom - state->applied.top,
             state->stripH, extra);
}

// Put our window in the gap. Cheap to call repeatedly: it does nothing unless the rectangle moved.
//
// The strip's position is derived from where the document frame **actually is**, not from the
// natural rect we computed for it. The two ought to be the same and are almost always, but "almost"
// was worth a visible seam: Word occasionally reports a child a pixel away from where we put it -
// the frame's client origin shifts by one when Windows redraws the border - and a strip positioned
// from stored numbers then sits a pixel off the document frame it is supposed to be flush against.
// Derived from the live rect, the two cannot disagree, whatever either of them is doing.
//
// `wwfRect` overrides that when the caller already knows where the document frame is about to be:
// during WM_WINDOWPOSCHANGING it has not moved yet, so reading it would place the strip a message
// behind.
static void PlaceStrip(StripState* state, const RECT* wwfRect)
{
    if (!state->strip || !state->hasApplied)
        return;

    // Not while the window is off screen: its layout is frozen and deliberately not trusted (see
    // AdjustProposed), so the document frame down there is not something to line up against.
    if (!wwfRect && (!IsWindowVisible(state->frame) || IsIconic(state->frame)))
        return;

    RECT document = state->applied;
    if (wwfRect)
    {
        document = *wwfRect;
    }
    else if (state->wwf && IsWindow(state->wwf))
    {
        RECT live;
        if (ChildRect(state->frame, state->wwf, &live))
        {
            // The live rect is trusted for pixel-level disagreement, not for a different layout.
            // A document frame a pixel from where we put it is the border artifact this whole
            // derivation exists to absorb; one 48 pixels away is Word mid-transition - it has reset
            // the frame and we have not shifted it back yet - and following it there would put the
            // strip over the ribbon or off the top of the window. Measured, both.
            LONG drift = live.top - state->applied.top;
            if (drift < 0) drift = -drift;
            if (drift < state->stripH)
                document = live;
        }
    }

    RECT want;
    want.left   = document.left;
    want.top    = document.top - state->stripH;
    want.right  = document.right;
    want.bottom = document.top;

    // Compared against where the strip *actually is*, not against where we last put it. Those are
    // not the same thing - Word, another add-in, or a message we did not see can move it - and
    // trusting our own bookkeeping leaves the strip drawn a pixel or two away from the document
    // frame it is supposed to be flush against, which is visible and was.
    RECT actual;
    if (ChildRect(state->frame, state->strip, &actual) && SameRect(&want, &actual))
    {
        state->stripAt = want;
        state->stripPlaced = TRUE;
        return;
    }

    if (state->stripPlaced && !SameRect(&actual, &state->stripAt))
    {
        LogWrite(L"strip  hwnd=0x%p  strip had drifted: at (%ld,%ld %ldx%ld), expected "
                 L"(%ld,%ld %ldx%ld) - correcting",
                 (void*)state->frame,
                 actual.left, actual.top, actual.right - actual.left, actual.bottom - actual.top,
                 want.left, want.top, want.right - want.left, want.bottom - want.top);
    }

    // SWP_NOZORDER after the first placement, deliberately. The strip is created at the top of the
    // z-order, which is where it needs to be to sit above the frame's background; forcing it back
    // there on every move would also put it above Backstage, which covers the whole client area and
    // is meant to cover us with it.
    SetWindowPos(state->strip, NULL,
                 want.left, want.top,
                 want.right - want.left, state->stripH,
                 SWP_NOZORDER | SWP_NOACTIVATE);

    state->stripAt = want;
    state->stripPlaced = TRUE;
    InvalidateRect(state->strip, NULL, FALSE);
}

// The heart of it: rewrite the position Word is *about* to move `_WwF` to.
//
// Called from inside `_WwF`'s WM_WINDOWPOSCHANGING, so the numbers here are a proposal, not a fact,
// and changing them changes where the window actually lands. Nothing ever flickers because nothing
// is ever wrong for a frame.
static void AdjustProposed(StripState* state, WINDOWPOS* pos)
{
    if (!state->enabled || !pos || !state->wwf)
        return;

    // Freeze while the window is not on screen - hidden, or minimised. Word dismantles a frame's
    // layout on the way out (closing a document was measured to leave `_WwF` filling the whole
    // client area, top edge at 0) and squeezes it to nothing on the way down to the taskbar.
    // Accepting either as a natural rect puts the strip over the ribbon, and if that window is the
    // active one it is then broadcast to every other window in the stack. Nobody can see a window
    // that is not on screen, so there is nothing to gain by shifting it; the janitor re-derives
    // when it comes back.
    if (!IsWindowVisible(state->frame) || IsIconic(state->frame))
        return;

    RECT current;
    if (!ChildRect(state->frame, state->wwf, &current))
        return;

    // Fill in whatever the flags say to leave alone, so the rest of this works on a whole rectangle
    // rather than on four values of which some may be meaningless.
    RECT proposed;
    if (pos->flags & SWP_NOMOVE)
    {
        proposed.left = current.left;
        proposed.top  = current.top;
    }
    else
    {
        proposed.left = pos->x;
        proposed.top  = pos->y;
    }

    LONG width  = (pos->flags & SWP_NOSIZE) ? (current.right - current.left)  : pos->cx;
    LONG height = (pos->flags & SWP_NOSIZE) ? (current.bottom - current.top)  : pos->cy;
    proposed.right  = proposed.left + width;
    proposed.bottom = proposed.top + height;

    if (width <= 0 || height <= 0)
        return;

    // Exactly where we last put it: this is our own placement coming back round, or Word confirming
    // it. Nothing to do, and doing something would be the accumulating shift that walks the frame
    // off the window.
    if (state->hasApplied && SameRect(&proposed, &state->applied))
        return;

    RECT natural, applied;

    if (state->hasApplied && proposed.top == state->applied.top)
    {
        // The top edge is still exactly where we put it, so nobody has re-laid out the top - this
        // proposal was computed from our shifted rect (a width change, a height change, or another
        // add-in shifting the same window). Keep the top, take the other edges. Treating this as a
        // fresh natural rect is what let two processes walk `_WwF` down to 25px tall in spike 2.
        natural = proposed;
        natural.top = state->natural.top;
        applied = proposed;
    }
    else
    {
        // Anything else is Word laying the window out from scratch. What it proposes is natural by
        // definition, and the shift is computed from it - never added to what is already there.
        natural = proposed;
        applied = proposed;
        applied.top = proposed.top + state->stripH;
    }

    // Tripwire. If honouring the shift would leave the document less than two strips tall,
    // something is wrong - most likely someone else shifting the same window - and the right answer
    // is to stop rather than to keep squeezing.
    if ((applied.bottom - applied.top) < (2 * state->stripH))
    {
        if (!state->trippedLogged)
        {
            state->trippedLogged = TRUE;
            LogWrite(L"strip  hwnd=0x%p  TRIPWIRE: shift would leave _WwF %ldpx tall (< 2 x %d) - "
                     L"leaving Word's layout alone. Is another tab add-in shifting the same window?",
                     (void*)state->frame, applied.bottom - applied.top, state->stripH);
        }
        return;
    }

    BOOL naturalMoved = !SameRect(&natural, &state->natural);

    state->natural    = natural;
    state->applied    = applied;
    state->hasApplied = TRUE;
    RememberClient(state);

    pos->x  = applied.left;
    pos->y  = applied.top;
    pos->cx = applied.right - applied.left;
    pos->cy = applied.bottom - applied.top;
    pos->flags &= ~(SWP_NOMOVE | SWP_NOSIZE);

    // In the same message as the document frame moves, rather than waiting for the CHANGED that
    // follows: the two are meant to be flush against each other, so they should move together. The
    // rect is passed in because the document frame has not moved yet - reading it here would place
    // the strip against where it used to be.
    PlaceStrip(state, &applied);

    if (naturalMoved)
    {
        LogRelayout(state, L"relayout");

        // Word has just laid this window out. If it is the focused one it is the only window in
        // the stack Word will lay out at all, so the rest are given the same interior from here.
        StackOnActiveLayout(state->frame, &natural);
    }
}

// ---------------------------------------------------------------------------------------------
// `_WwF`'s subclass procedure.
// ---------------------------------------------------------------------------------------------

static LRESULT CALLBACK WwfSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                                        UINT_PTR idSubclass, DWORD_PTR refData)
{
    (void)idSubclass;

    StripState* state = (StripState*)refData;

    switch (msg)
    {
    case WM_WINDOWPOSCHANGING:
        if (state && state->wwf == hwnd)
            AdjustProposed(state, (WINDOWPOS*)lParam);
        break;

    case WM_WINDOWPOSCHANGED:
        if (state && state->wwf == hwnd)
        {
            const WINDOWPOS* pos = (const WINDOWPOS*)lParam;
            PlaceStrip(state, NULL);

            // Backstage and window minimisation both take `_WwF` away; the strip belongs with it,
            // not with the frame.
            if (pos && (pos->flags & SWP_HIDEWINDOW) && state->strip)
                ShowWindow(state->strip, SW_HIDE);
            else if (pos && (pos->flags & SWP_SHOWWINDOW) && state->strip)
                ShowWindow(state->strip, SW_SHOWNA);
        }
        break;

    case WM_NCDESTROY:
        // Word replaces `_WwF` rather than only moving it - closing a document was measured to
        // destroy one window and reuse another. Unbinding here lets the janitor rebind to whatever
        // replaces it.
        if (state && state->wwf == hwnd)
        {
            LogWrite(L"strip  hwnd=0x%p  _WwF 0x%p destroyed - unbinding", (void*)state->frame, (void*)hwnd);
            state->wwf = NULL;
            state->hasApplied = FALSE;
            state->stripPlaced = FALSE;
        }
        RemoveWindowSubclass(hwnd, WwfSubclassProc, kWwfSubclassId);
        break;

    default:
        break;
    }

    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

// ---------------------------------------------------------------------------------------------
// The strip window.
//
// Painted rather than composed of controls: at this stage it exists to show, without reading a log,
// that we own a horizontal band of Word's layout and that the band tracks the document frame
// through every relayout. The tab it draws is a placeholder for the real tab strip.
// ---------------------------------------------------------------------------------------------

// Where everything in the strip sits, in the strip's client coordinates.
//
// This is the single source of truth for the tab row: painting draws what it returns and
// hit-testing reads what it returns, so there is no way for a click to land somewhere other than
// what the user is looking at. tools\WordLayout.cs mirrors it for the check scripts, and if the two
// ever drift the injected clicks miss and the assertions fail loudly - which is the intended
// failure, rather than a test that quietly clicks the wrong tab and passes.
//
// There are three ways the row can be laid out, and which one is in force is decided here and
// nowhere else:
//
//   FIT      Every tab is at least the minimum width and they all fit. The + sits after the last
//            one. This is what the strip has always done, and it is unchanged to the pixel - the
//            arithmetic below is the arithmetic that was here before, in the same order.
//
//   SCROLL   There are more documents than fit at the minimum width. The + is pinned to the right
//            edge with two scroll buttons beside it, and the tabs live in a `track` that stops
//            before them and scrolls inside it.
//
//   SQUEEZE  HKCU\Software\WordTab\TabScroll=0. No minimum: the tabs divide the track between them
//            however many there are, so every document is on screen at once however narrow that
//            makes it. See StripStart for why this exists.
//
// The condition separating FIT from the other two is `width * count > available` *after* the minimum
// has been applied - which is exactly, and only, the case in which the old code clamped the + back
// on top of the last tab. Everything this slice changes is inside that condition. That is what makes
// it safe: the two hundred and sixty-five assertions written before it all run in FIT.
static void ComputeLayout(StripState* state, const RECT* client,
                          HWND* frames, int count, int activeIndex, StripLayout* out)
{
    int pad     = Scaled(TAB_LOGICAL_PAD, state->dpi);
    int gap     = Scaled(TAB_LOGICAL_GAP, state->dpi);
    int minimum = Scaled(TAB_LOGICAL_MIN_W, state->dpi);
    int desired = Scaled(TAB_LOGICAL_W, state->dpi);
    int plusW   = Scaled(PLUS_LOGICAL, state->dpi);
    int closeW  = Scaled(CLOSE_LOGICAL, state->dpi);
    int chevW   = Scaled(CHEVRON_LOGICAL, state->dpi);
    int hair    = Scaled(2, state->dpi);          // the sliver of well left between two tabs

    if (count > MAX_TABS)
        count = MAX_TABS;
    out->count     = count;
    out->hasPlus   = FALSE;
    out->hasNav    = FALSE;
    out->canPrev   = FALSE;
    out->canNext   = FALSE;
    out->scroll    = 0;
    out->maxScroll = 0;
    SetRectEmpty(&out->plus);
    SetRectEmpty(&out->prev);
    SetRectEmpty(&out->next);

    // The new-document button's width comes out of the space before the tabs are sized, not after.
    // Tabs shrink as documents are opened; a button does not, and a button that has been squeezed
    // off the end of the strip is a feature the user cannot reach.
    int reserved = g_buttonsEnabled ? (plusW + gap) : 0;
    int available = (client->right - client->left) - pad * 2 - reserved;
    if (available < minimum)
        available = minimum;

    int width = desired;
    if (count > 0 && width * count > available)
        width = available / count;
    if (width < minimum)
        width = minimum;

    BOOL overflow = (count > 0 && width * count > available);

    // The track. In FIT it is simply the padded strip and nothing is clipped by it; in the two
    // overflow modes it stops short of the button cluster, and that is the line the row is not
    // allowed to cross.
    out->track.left   = client->left + pad;
    out->track.right  = client->right - pad;
    out->track.top    = client->top;
    out->track.bottom = client->bottom;

    if (overflow && g_buttonsEnabled)
    {
        // [ tabs ... ] gap [ ‹ ][ › ] gap [ + ]
        //
        // The two chevrons touch each other on purpose: they are one control with two directions,
        // and a gap between them reads as two unrelated buttons that happen to be adjacent.
        out->plus.right  = client->right - pad;
        out->plus.left   = out->plus.right - plusW;
        out->plus.top    = client->top + Scaled(6, state->dpi);
        out->plus.bottom = client->bottom - Scaled(6, state->dpi);

        out->next.right = out->plus.left - gap;
        out->next.left  = out->next.right - chevW;
        out->prev.right = out->next.left;
        out->prev.left  = out->prev.right - chevW;

        out->next.top    = out->prev.top    = client->top + Scaled(6, state->dpi);
        out->next.bottom = out->prev.bottom = client->bottom - Scaled(6, state->dpi);

        out->track.right = out->prev.left - gap;
        out->hasNav = g_scrollEnabled;

        if (!g_scrollEnabled)
        {
            // SQUEEZE. The buttons still need their space reserved - they are why the track is
            // short - but there is nothing to scroll, so the chevrons are not drawn and the room
            // they would have taken goes back to the tabs.
            out->track.right = out->plus.left - gap;
            SetRectEmpty(&out->prev);
            SetRectEmpty(&out->next);
        }
    }
    else if (overflow)
    {
        // Buttons switched off entirely: no cluster, so the track is the whole padded strip. The row
        // still scrolls - the wheel is the only way to reach the far end, and that is what TabButtons
        // being off means.
        out->hasNav = FALSE;
    }

    int trackW = out->track.right - out->track.left;
    if (trackW < 0)
        trackW = 0;

    if (overflow && !g_scrollEnabled)
    {
        // SQUEEZE has no minimum and therefore cannot overflow at any count: the row is exactly as
        // wide as the track by construction. A tab can end up narrower than its own close button,
        // which is why that button drops out below three times its width - the tab becomes a colour
        // with a name in it, and it is still there, still clickable, still yours to switch to.
        width = (count > 0) ? (trackW / count) : desired;
        if (width < 1)
            width = 1;
        overflow = FALSE;
    }

    if (overflow)
    {
        out->maxScroll = count * width - trackW;
        if (out->maxScroll < 0)
            out->maxScroll = 0;
    }

    // The scroll position. Written back to the global rather than merely read, so that a window
    // widened until the row fits does not keep a scroll offset that no longer means anything - and
    // only when this strip has tabs, because an empty row belongs to a window that is not in the
    // stack and its layout must not speak for the row every other window is showing.
    if (count > 0)
    {
        if (g_scroll > out->maxScroll) g_scroll = out->maxScroll;
        if (g_scroll < 0)              g_scroll = 0;

        // Reveal the active tab. A tab row whose selected tab is off screen has stopped answering
        // the one question it exists to answer, so this is not optional - but it cannot be done on
        // every layout either, or the row would snap back to the active tab a frame after the wheel
        // moved it and the wheel would appear not to work at all.
        //
        // So it fires on the two things that can put the active tab off screen without the user
        // having asked for it:
        //
        //   the active document changed - a new one appends a tab at the far end and switches to it,
        //   Ctrl+F6 walks the windows, closing a tab moves the selection;
        //
        //   the shape of the row changed - `maxScroll` is a function of the tab count and the width
        //   of the track, so it moves when a document opens or closes and on every step of a resize
        //   drag. Narrowing the window until three tabs fit used to leave the user looking at tabs
        //   one to three with tab six selected and nothing on screen saying so. Photographed.
        //
        // Neither is a poll: both are quantities this function has already computed, compared with
        // what they were the last time it did. A wheel scroll changes neither of them, which is
        // exactly the property that makes the wheel work.
        HWND activeFrame = (frames && activeIndex >= 0 && activeIndex < count)
                           ? frames[activeIndex] : NULL;
        if (activeFrame && (activeFrame != g_scrollShown || out->maxScroll != g_scrollMax))
        {
            g_scrollShown = activeFrame;
            g_scrollMax   = out->maxScroll;

            int left  = activeIndex * width;
            int right = left + width;
            if (left < g_scroll)
                g_scroll = left;
            if (right > g_scroll + trackW)
                g_scroll = right - trackW;

            if (g_scroll > out->maxScroll) g_scroll = out->maxScroll;
            if (g_scroll < 0)              g_scroll = 0;
        }

        out->scroll = g_scroll;
    }

    out->width   = width;
    out->canPrev = (out->scroll > 0);
    out->canNext = (out->scroll < out->maxScroll);

    for (int i = 0; i < count; i++)
    {
        RECT* tab = &out->tab[i];
        tab->left   = out->track.left + i * width - out->scroll;
        tab->right  = tab->left + width - hair;                    // a hairline between tabs
        if (tab->right <= tab->left)
            tab->right = tab->left + 1;
        tab->top    = client->top + Scaled(3, state->dpi);
        tab->bottom = client->bottom;

        // A close button, but only where there is honestly room for one. A tab narrow enough that
        // the button covers the name is a tab whose button closes a document the user cannot
        // identify, so below that width the name wins and there is no button at all.
        //
        // ...and only where the whole of it is inside the track. A close button on a tab that is
        // half scrolled under the buttons beside it is a target the user cannot see the edges of,
        // and this is the second half of the defect this slice is about: the first half was a +
        // drawn over a tab, and both come from the same habit of clamping a rectangle instead of
        // deciding it does not belong.
        SetRectEmpty(&out->close[i]);
        if (g_buttonsEnabled && (tab->right - tab->left) >= closeW * 3)
        {
            RECT close;
            int middle = (tab->top + tab->bottom) / 2;
            close.right  = tab->right - Scaled(6, state->dpi);
            close.left   = close.right - closeW;
            close.top    = middle - closeW / 2;
            close.bottom = close.top + closeW;

            if (close.left >= out->track.left && close.right <= out->track.right)
                out->close[i] = close;
        }
    }

    if (!g_buttonsEnabled)
        return;

    if (overflow)
    {
        // Already placed with the cluster above: pinned, and the track was cut short to clear it.
        out->hasPlus = (out->plus.right <= client->right && out->plus.bottom > out->plus.top);
        return;
    }

    // FIT and SQUEEZE: after the last tab. There is no clamp here any more, and its absence is the
    // fix. It used to exist for the overflowing case and it answered by putting the + on top of a
    // tab - a button drawn over a target that was hit-tested first, so the user pressed one thing
    // and got another. Overflow is now a layout of its own and never arrives here.
    int after = (count > 0) ? (out->tab[count - 1].right + gap) : (client->left + pad);
    if (after > client->right - pad - plusW)
        after = client->right - pad - plusW;
    if (after < client->left + pad)
        after = client->left + pad;

    out->plus.left   = after;
    out->plus.right  = after + plusW;
    out->plus.top    = client->top + Scaled(6, state->dpi);
    out->plus.bottom = client->bottom - Scaled(6, state->dpi);
    out->hasPlus = (out->plus.right <= client->right && out->plus.bottom > out->plus.top);
}

// The row's layout for this strip, tabs and all. Every caller wanted the same four lines before it
// could call ComputeLayout, and one of them - the hit test - used to pass NULL for the active tab and
// so could not have revealed it. Returns the number of tabs; `frames` must have room for MAX_STRIPS.
static int LayoutOf(StripState* state, const RECT* client, HWND* frames, StripLayout* out)
{
    int activeIndex = 0;
    int count = StackTabs(state->frame, frames, MAX_STRIPS, &activeIndex);
    ComputeLayout(state, client, frames, count, activeIndex, out);
    return count;
}

// What a point is over. The close button is tested before the tab it sits on: they overlap by
// definition, and the smaller target is the more specific intent.
//
// Tabs are tested against the part of them that is inside the track, never against the whole
// rectangle. In an overflowing row a tab runs on underneath the scroll buttons and the +, and testing
// the whole thing would put a tab's hit area under a button that is drawn on top of it - which is the
// defect this slice exists to remove, arriving from the other direction.
static StripHit HitTestStrip(StripState* state, HWND hwnd, POINT point)
{
    StripHit hit;
    hit.kind  = HIT_NONE;
    hit.frame = NULL;
    hit.index = -1;
    SetRectEmpty(&hit.tab);

    RECT client;
    if (!GetClientRect(hwnd, &client))
        return hit;

    HWND frames[MAX_STRIPS];
    StripLayout layout;
    LayoutOf(state, &client, frames, &layout);

    for (int i = 0; i < layout.count; i++)
    {
        RECT visible;
        if (!IntersectRect(&visible, &layout.tab[i], &layout.track))
            continue;

        if (!IsRectEmpty(&layout.close[i]) && PtInRect(&layout.close[i], point))
        {
            hit.kind  = HIT_CLOSE;
            hit.frame = frames[i];
            hit.index = i;
            hit.tab   = layout.tab[i];
            return hit;
        }
        if (PtInRect(&visible, point))
        {
            hit.kind  = HIT_TAB;
            hit.frame = frames[i];
            hit.index = i;
            hit.tab   = layout.tab[i];
            return hit;
        }
    }

    // A scroll button with nowhere to go is not a target. It is drawn dimmed, and a dimmed button
    // that lights up under the pointer and then does nothing when pressed is worse than one that
    // ignores the pointer entirely.
    if (layout.hasNav && layout.canPrev && PtInRect(&layout.prev, point))
        hit.kind = HIT_PREV;
    else if (layout.hasNav && layout.canNext && PtInRect(&layout.next, point))
        hit.kind = HIT_NEXT;
    else if (layout.hasPlus && PtInRect(&layout.plus, point))
        hit.kind = HIT_PLUS;

    return hit;
}

// Move the row, and say whether it actually moved. The clamp is ComputeLayout's, not this function's:
// everything here does is offer a number, and the next layout decides how much of it was possible.
static BOOL ScrollBy(StripState* state, HWND hwnd, int delta)
{
    RECT client;
    if (!GetClientRect(hwnd, &client))
        return FALSE;

    HWND frames[MAX_STRIPS];
    StripLayout layout;
    LayoutOf(state, &client, frames, &layout);

    if (layout.maxScroll <= 0)
        return FALSE;

    int want = layout.scroll + delta;
    if (want > layout.maxScroll) want = layout.maxScroll;
    if (want < 0)                want = 0;
    if (want == layout.scroll)
        return FALSE;

    g_scroll = want;

    // Throttled the same way a relayout is, and for the same reason: the auto-scroll timer comes
    // through here sixteen times a second. The two positions that always get a line are the ends,
    // because "it will not scroll any further" is the complaint this log would be read to answer.
    {
        static DWORD lastTick = 0;
        static int   suppressed = 0;
        DWORD now = GetTickCount();
        BOOL atEnd = (want == 0 || want == layout.maxScroll);

        if (!atEnd && lastTick != 0 && (now - lastTick) < 250)
        {
            suppressed++;
        }
        else
        {
            lastTick = now;
            LogWrite(L"strip  hwnd=0x%p  row scrolled to %d of %d%s", (void*)hwnd, want,
                     layout.maxScroll,
                     suppressed ? L" (+ steps not logged)" : L"");
            suppressed = 0;
        }
    }

    StripRefreshTabs();
    return TRUE;
}

static POINT PointOf(LPARAM lParam)
{
    POINT point;
    point.x = (short)LOWORD(lParam);
    point.y = (short)HIWORD(lParam);
    return point;
}

// Repaint only when what is under the pointer has actually changed. WM_MOUSEMOVE arrives on every
// pixel of movement; invalidating on each one would repaint the strip continuously while the user
// is doing nothing but crossing it on the way to the ribbon.
static void SetHot(StripState* state, HWND hwnd, int kind, HWND frame)
{
    if (state->hotKind == kind && state->hotFrame == frame)
        return;
    state->hotKind  = kind;
    state->hotFrame = frame;
    InvalidateRect(hwnd, NULL, FALSE);
}

// Ask for the one WM_MOUSELEAVE that tells us the pointer has gone. Without it a highlight stays
// lit under a pointer that left the strip a minute ago - there is no "mouse exited" message
// otherwise, and polling for it would be a timer for something an API already answers.
static void ArmLeaveTracking(StripState* state, HWND hwnd)
{
    if (state->tracking)
        return;

    TRACKMOUSEEVENT track;
    memset(&track, 0, sizeof(track));
    track.cbSize    = sizeof(track);
    track.dwFlags   = TME_LEAVE;
    track.hwndTrack = hwnd;
    if (TrackMouseEvent(&track))
        state->tracking = TRUE;
}

// ---------------------------------------------------------------------------------------------
// Compositing.
//
// GDI has no anti-aliasing. A rounded tab drawn with RoundRect is a staircase, and at 150% the
// staircase is three physical pixels tall - which is exactly the sort of thing that makes a piece of
// UI read as somebody's first draft. So the strip composites its own pixels: a 32-bit DIB, coverage
// worked out per pixel, blended by hand.
//
// Not GDI+ and not AlphaBlend. GDI+ has a process-wide startup and shutdown that we would be sharing
// with Word's own use of it, and AlphaBlend lives in msimg32, which is a new import into someone
// else's process. Neither is worth it for a few hundred pixels of arithmetic that fits on one
// screen, and hand-rolled coverage is in the same spirit as DrawGlyph below refusing to trust a font
// to have a multiplication sign.
//
// The surface carries its own DC as well as its own bytes, because the tab names are still drawn by
// GDI - laying out and hinting text is not something to reimplement. That mixing is the one hazard
// here: GDI batches, so anything that reads the pixels back has to GdiFlush first. DrawOneTab does
// that at its start, which is the only ordering rule in this file.
// ---------------------------------------------------------------------------------------------

struct Surface
{
    BYTE* bits;      // BGRA, top-down
    int   w, h;
    HDC   dc;        // the same pixels, for the text

    // Nothing outside this is written. It is the whole surface except while the tabs are being
    // drawn, when it is the track - which is how a tab that is half scrolled off the end comes out
    // cut square at the edge rather than with a rounded corner in the middle of the row. Clipping
    // the *rectangle* instead would round the corners of the cut, and a tab that appears to end
    // neatly where it has in fact been truncated is a tab that hides the fact there is more of it.
    RECT clip;
};

static void SurfClip(Surface* s, const RECT* box)
{
    if (box)
    {
        s->clip = *box;
        if (s->clip.left   < 0)    s->clip.left   = 0;
        if (s->clip.top    < 0)    s->clip.top    = 0;
        if (s->clip.right  > s->w) s->clip.right  = s->w;
        if (s->clip.bottom > s->h) s->clip.bottom = s->h;
    }
    else
    {
        s->clip.left = s->clip.top = 0;
        s->clip.right = s->w;
        s->clip.bottom = s->h;
    }

    // The text is GDI's, so GDI needs telling separately. The DC outlives the paint - it belongs to
    // the strip, not to this drawing pass - so a clip left behind here would silently truncate the
    // next one.
    if (s->dc)
    {
        if (box)
        {
            HRGN region = CreateRectRgn(s->clip.left, s->clip.top, s->clip.right, s->clip.bottom);
            SelectClipRgn(s->dc, region);
            if (region)
                DeleteObject(region);
        }
        else
        {
            SelectClipRgn(s->dc, NULL);
        }
    }
}

static inline void Blend(Surface* s, int x, int y, COLORREF color, int alpha)
{
    if (alpha <= 0 || x < 0 || y < 0 || x >= s->w || y >= s->h)
        return;
    if (x < s->clip.left || x >= s->clip.right || y < s->clip.top || y >= s->clip.bottom)
        return;

    BYTE* p = s->bits + ((size_t)y * (size_t)s->w + (size_t)x) * 4;
    int r = GetRValue(color), g = GetGValue(color), b = GetBValue(color);

    if (alpha >= 255)
    {
        p[0] = (BYTE)b; p[1] = (BYTE)g; p[2] = (BYTE)r; p[3] = 255;
        return;
    }
    p[0] = (BYTE)((p[0] * (255 - alpha) + b * alpha) / 255);
    p[1] = (BYTE)((p[1] * (255 - alpha) + g * alpha) / 255);
    p[2] = (BYTE)((p[2] * (255 - alpha) + r * alpha) / 255);
    p[3] = 255;
}

static void SurfFill(Surface* s, const RECT* box, COLORREF color)
{
    int x0 = box->left   < 0 ? 0 : box->left;
    int y0 = box->top    < 0 ? 0 : box->top;
    int x1 = box->right  > s->w ? s->w : box->right;
    int y1 = box->bottom > s->h ? s->h : box->bottom;

    for (int y = y0; y < y1; y++)
        for (int x = x0; x < x1; x++)
            Blend(s, x, y, color, 255);
}

// How much of one pixel falls inside a circle, by sixteen samples on an even grid. Analytic coverage
// of a circle against a square is a page of algebra for an arc nine pixels long; sixteen samples are
// indistinguishable from it at this size and the whole thing stays in integers.
static int CircleCoverage(int x, int y, int cx, int cy, int radius)
{
    long long rr = (long long)(radius * 8) * (long long)(radius * 8);
    int hits = 0;

    for (int j = 0; j < 4; j++)
    {
        long long dy = (long long)(y * 8 + 2 * j + 1) - (long long)cy * 8;
        for (int i = 0; i < 4; i++)
        {
            long long dx = (long long)(x * 8 + 2 * i + 1) - (long long)cx * 8;
            if (dx * dx + dy * dy <= rr)
                hits++;
        }
    }
    return hits * 255 / 16;
}

// A filled rectangle with independently rounded top and bottom corners. A tab card is round on top
// and square on the bottom, because it meets the document there and a tab that is round at the
// bottom is a lozenge.
static void SurfRoundRect(Surface* s, const RECT* box, int rTop, int rBottom,
                          COLORREF color, int alpha)
{
    int x0 = box->left, x1 = box->right, y0 = box->top, y1 = box->bottom;
    if (x1 <= x0 || y1 <= y0 || alpha <= 0)
        return;

    int w = x1 - x0, h = y1 - y0;
    if (rTop    > w / 2) rTop    = w / 2;
    if (rBottom > w / 2) rBottom = w / 2;
    if (rTop    > h)     rTop    = h;
    if (rBottom > h)     rBottom = h;
    if (rTop < 0) rTop = 0;
    if (rBottom < 0) rBottom = 0;

    int clipY0 = y0 < 0 ? 0 : y0, clipY1 = y1 > s->h ? s->h : y1;
    int clipX0 = x0 < 0 ? 0 : x0, clipX1 = x1 > s->w ? s->w : x1;

    for (int y = clipY0; y < clipY1; y++)
    {
        BOOL inTop    = (y < y0 + rTop);
        BOOL inBottom = (y >= y1 - rBottom);

        for (int x = clipX0; x < clipX1; x++)
        {
            int cov = 255;

            if (inTop && x < x0 + rTop)
                cov = CircleCoverage(x, y, x0 + rTop, y0 + rTop, rTop);
            else if (inTop && x >= x1 - rTop)
                cov = CircleCoverage(x, y, x1 - rTop, y0 + rTop, rTop);
            else if (inBottom && x < x0 + rBottom)
                cov = CircleCoverage(x, y, x0 + rBottom, y1 - rBottom, rBottom);
            else if (inBottom && x >= x1 - rBottom)
                cov = CircleCoverage(x, y, x1 - rBottom, y1 - rBottom, rBottom);

            if (cov > 0)
                Blend(s, x, y, color, alpha * cov / 255);
        }
    }
}

// A filled circle, centred in a box. The same sixteen-sample coverage as the rounded corners above,
// so the dot is anti-aliased by the same arithmetic as everything else in the strip and cannot end
// up being the one shape with a staircase on it.
//
// Its own function rather than SurfRoundRect with the radius set to half the side: that would work,
// the clamps at the top of SurfRoundRect allow it, but it would express "a circle" as "a rectangle
// so rounded it stopped being one" and would round to an even diameter whatever radius it was given.
static void SurfDisc(Surface* s, const RECT* box, int radius, COLORREF color)
{
    if (radius <= 0)
        return;

    int cx = (box->left + box->right) / 2;
    int cy = (box->top + box->bottom) / 2;

    int x0 = cx - radius - 1, x1 = cx + radius + 1;
    int y0 = cy - radius - 1, y1 = cy + radius + 1;

    if (x0 < 0) x0 = 0;
    if (y0 < 0) y0 = 0;
    if (x1 > s->w) x1 = s->w;
    if (y1 > s->h) y1 = s->h;

    for (int y = y0; y < y1; y++)
        for (int x = x0; x < x1; x++)
        {
            int cov = CircleCoverage(x, y, cx, cy, radius);
            if (cov > 0)
                Blend(s, x, y, color, cov);
        }
}

// The lift under a carried tab: the same shape, drawn a few times, each one larger and offset a
// little further down, each one faint. The overlap is what makes the falloff - there is no blur
// kernel here and none is needed at four pixels.
//
// It reaches past the tab it belongs to, which is the only thing in the strip that does. That is
// bounded on purpose: LIFT_LOGICAL_SPREAD is 4 logical pixels, and the nearest thing any check
// asserts must not change is a whole tab away.
static void SurfLift(Surface* s, const RECT* box, int radius, int spread)
{
    if (spread <= 0)
        return;

    for (int i = spread; i >= 1; i--)
    {
        RECT ring = *box;
        InflateRect(&ring, i, i);
        OffsetRect(&ring, 0, (i + 1) / 2);
        SurfRoundRect(s, &ring, radius + i, radius + i, RGB(0, 0, 0), LIFT_ALPHA / spread);
    }
}

// An anti-aliased stroke, by distance to the segment. Used for the two glyphs and nothing else, so
// it is written for short lines and does not try to be a rasteriser.
static void SurfStroke(Surface* s, float ax, float ay, float bx, float by,
                       float thickness, COLORREF color)
{
    float half = thickness * 0.5f;
    float dx = bx - ax, dy = by - ay;
    float len2 = dx * dx + dy * dy;
    if (len2 <= 0.0f)
        return;

    int x0 = (int)((ax < bx ? ax : bx) - half - 1.0f);
    int x1 = (int)((ax > bx ? ax : bx) + half + 2.0f);
    int y0 = (int)((ay < by ? ay : by) - half - 1.0f);
    int y1 = (int)((ay > by ? ay : by) + half + 2.0f);

    for (int y = y0; y < y1; y++)
    {
        for (int x = x0; x < x1; x++)
        {
            float px = (float)x + 0.5f, py = (float)y + 0.5f;
            float t = ((px - ax) * dx + (py - ay) * dy) / len2;
            if (t < 0.0f) t = 0.0f;
            if (t > 1.0f) t = 1.0f;
            float qx = ax + t * dx - px, qy = ay + t * dy - py;
            float dist = sqrtf(qx * qx + qy * qy);

            float cover = half + 0.5f - dist;
            if (cover <= 0.0f) continue;
            if (cover > 1.0f) cover = 1.0f;
            Blend(s, x, y, color, (int)(cover * 255.0f));
        }
    }
}

// The glyphs, drawn as strokes rather than as characters. A font is not guaranteed to have a
// multiplication sign or a heavy plus at any particular weight, and one that substitutes silently
// gives a close button that looks like a lowercase x. Two strokes cannot be substituted.
//
// The stroke used to be exactly one physical pixel at every DPI, which is how a 150% rig ended up
// with a hairline x on a full-size button. It scales now, and it is anti-aliased, which for a
// diagonal is most of the difference.
enum { GLYPH_CROSS = 0, GLYPH_PLUS, GLYPH_PREV, GLYPH_NEXT };

static void DrawGlyph(Surface* s, const RECT* box, COLORREF color, int dpi, int kind)
{
    float thickness = (float)(GLYPH_LOGICAL_STROKE_TENTHS * dpi) / 960.0f;
    if (thickness < 1.0f) thickness = 1.0f;

    float cx = (float)(box->left + box->right) * 0.5f;
    float cy = (float)(box->top + box->bottom) * 0.5f;
    float arm = (float)Scaled(4, dpi);

    // A chevron is drawn narrower than it is tall - the arms are the same length as the + 's, but
    // half as far apart horizontally. Square, it reads as a "greater than" sign rather than as a
    // direction.
    float half = arm * 0.55f;

    switch (kind)
    {
    case GLYPH_PLUS:
        SurfStroke(s, cx - arm, cy, cx + arm, cy, thickness, color);
        SurfStroke(s, cx, cy - arm, cx, cy + arm, thickness, color);
        break;

    case GLYPH_PREV:
        SurfStroke(s, cx + half, cy - arm, cx - half, cy, thickness, color);
        SurfStroke(s, cx - half, cy, cx + half, cy + arm, thickness, color);
        break;

    case GLYPH_NEXT:
        SurfStroke(s, cx - half, cy - arm, cx + half, cy, thickness, color);
        SurfStroke(s, cx + half, cy, cx - half, cy + arm, thickness, color);
        break;

    default:
        SurfStroke(s, cx - arm, cy - arm, cx + arm, cy + arm, thickness, color);
        SurfStroke(s, cx + arm, cy - arm, cx - arm, cy + arm, thickness, color);
        break;
    }
}

// A close or new button's background: nothing at rest, a rounded chip under the pointer, a darker
// one while it is held. Slightly larger than the hit rectangle so the glyph is not touching its own
// edge.
static void DrawChip(Surface* s, const RECT* box, BOOL hot, BOOL down, int dpi)
{
    if (!hot && !down)
        return;

    RECT chip = *box;
    InflateRect(&chip, Scaled(2, dpi), Scaled(2, dpi));
    int radius = Scaled(CHIP_LOGICAL_RADIUS, dpi);
    SurfRoundRect(s, &chip, radius, radius, down ? g_palette.chipDown : g_palette.chip, 255);
}

// One tab: its card, its name and its close button.
//
// Its rectangle is a parameter rather than an index into the layout, and that is still the whole
// point: a tab being carried is drawn by this same function at wherever the pointer has taken it, so
// a dragged tab cannot end up looking like a different kind of object from a tab sitting still. The
// lift is a parameter for the same reason - it is an argument to this function, not a second
// drawing path for dragged tabs.
// The one place a tab's name is put on a device context, and therefore the one place that knows
// whether it fitted.
//
// Both renderers used to do this themselves - the same string, the same flags, two DrawTextW calls
// eighty lines apart. That was tolerable while the answer was only ever pixels, and it stopped being
// tolerable when the tooltip needed to know whether the name had been cut: worked out anywhere else
// it would be a second derivation of the text box, and two derivations of one rectangle is the shape
// of defect this project keeps finding. So the cut is recorded by the code that does the cutting.
//
// DT_CALCRECT measures what the string wanted; the box is what it got. Measuring is the only honest
// test - the alternative is reading pixels back looking for an ellipsis glyph, which cannot tell an
// ellipsis Windows added from three dots the user typed.
static void DrawTabName(StripState* state, HDC dc, const RECT* box, int index,
                        const wchar_t* title, COLORREF color)
{
    if (!dc || !box || box->right <= box->left)
        return;

    SetTextColor(dc, color);

    RECT text = *box;
    DrawTextW(dc, title, -1, &text,
              DT_SINGLELINE | DT_VCENTER | DT_LEFT | DT_END_ELLIPSIS | DT_NOPREFIX);

    if (state && index >= 0 && index < MAX_TABS)
    {
        RECT wanted = *box;
        DrawTextW(dc, title, -1, &wanted, DT_SINGLELINE | DT_LEFT | DT_NOPREFIX | DT_CALCRECT);
        state->nameCut[index] = ((wanted.right - wanted.left) > (box->right - box->left));
    }
}

static void DrawOneTab(StripState* state, Surface* s, HWND frame, int index,
                       RECT tab, RECT close, BOOL selected, BOOL hot, BOOL lifted, BOOL first,
                       BOOL modified)
{
    // The one ordering rule in this file. Tab names are drawn by GDI and GDI batches; everything
    // below reads the pixels back to blend against them, and a batch still in flight would be
    // composited over after it lands rather than before.
    GdiFlush();

    int dpi    = state->dpi;
    int radius = Scaled(TAB_LOGICAL_RADIUS, dpi);
    int inset  = Scaled(TAB_LOGICAL_INSET, dpi);

    RECT card = tab;
    card.left  += inset;
    card.right -= inset;
    if (card.right <= card.left)
        card = tab;

    if (lifted)
        SurfLift(s, &card, radius, Scaled(LIFT_LOGICAL_SPREAD, dpi));

    if (selected || lifted)
    {
        // The active card is Word's own chrome colour, brought down out of the ribbon. Its border is
        // a single step away from that, and it exists because in a light theme the card is white on
        // near-white and without an edge it is a smudge rather than a shape.
        COLORREF fill = lifted ? g_palette.lifted : g_palette.selected;

        SurfRoundRect(s, &card, radius, 0, g_palette.edge, 255);

        RECT inner = card;
        InflateRect(&inner, -1, 0);
        inner.top += 1;
        SurfRoundRect(s, &inner, radius - 1, 0, fill, 255);
    }
    else if (hot)
    {
        SurfRoundRect(s, &card, radius, 0, g_palette.hover, 255);
    }
    else if (!first)
    {
        // At rest an inactive tab is not a shape at all - it is a name on the well, the way a
        // browser draws them. What separates two of them is a rule, and it lives inside the right
        // tab's own rectangle so that hovering one tab can never change a pixel of another.
        RECT rule;
        rule.left   = card.left;
        rule.right  = card.left + 1;
        rule.top    = card.top + Scaled(7, dpi);
        rule.bottom = card.bottom - Scaled(6, dpi);
        SurfFill(s, &rule, g_palette.separator);
    }

    BOOL hasClose = !IsRectEmpty(&close) && close.right <= tab.right;

    wchar_t title[256];
    StripTabName(frame, title, 256);

    RECT text = tab;
    text.left += Scaled(12, dpi);
    // The name stops before the button rather than running under it. A title clipped by an
    // ellipsis reads as a long name; one running under a close button reads as a bug.
    text.right = hasClose ? (close.left - Scaled(4, dpi))
                          : (tab.right - Scaled(10, dpi));
    DrawTabName(state, s->dc, &text, index, title,
                (selected || lifted) ? g_palette.text : g_palette.textIdle);

    if (hasClose)
    {
        GdiFlush();

        // The dot is the close button, until the pointer arrives. VS Code's trick, and it is taken
        // for a measured reason rather than a stylistic one: the row already scrolls at six
        // documents in a Word window of Word's own default size, so a mark that took width of its
        // own would make every row scroll sooner, permanently, to show something that is usually
        // "no". This one takes none - the rectangle is the close button's, unchanged, which is also
        // why ComputeLayout and tools\WordLayout.cs needed no edit for this slice.
        //
        // `hot` and not `hotClose`: hovering anywhere on the tab brings the x back, so the button is
        // reached by moving towards the tab rather than by finding the dot first. A carried tab and
        // a tab with its menu open are both hot, and so both show the x, which is the same answer as
        // "the pointer is on it".
        //
        // `!downClose` as well, and it is not redundant. An armed press is NOT a subset of hot: a
        // press on the close button takes capture and sets pressKind, then sliding the pointer off
        // the tab while still holding clears hotFrame - only WM_CAPTURECHANGED clears pressKind - so
        // the button is still armed with hot FALSE. Without this term a modified tab would drop its
        // held chip and go back to looking at rest, while an unmodified tab under the identical
        // gesture keeps it. The press is live either way: moving back onto the button and releasing
        // still closes the document. One gesture must not be drawn two ways depending on whether the
        // document happens to be saved.
        BOOL hotClose  = (state->hotKind == HIT_CLOSE && state->hotFrame == frame);
        BOOL downClose = (state->pressKind == HIT_CLOSE && state->pressFrame == frame);

        if (modified && !hot && !downClose)
        {
            SurfDisc(s, &close, Scaled(DOT_LOGICAL_RADIUS, dpi), g_palette.glyph);
        }
        else
        {
            DrawChip(s, &close, hotClose, downClose, dpi);
            DrawGlyph(s, &close, hotClose || downClose ? g_palette.glyphHot : g_palette.glyph,
                      dpi, GLYPH_CROSS);
        }
    }
}

// Everything in the strip, composited into the surface.
//
// Separate from PaintStrip so the same drawing serves WM_PAINT and WM_PRINTCLIENT - which is how the
// check scripts photograph a strip without a camera. Anything that only happened on the WM_PAINT
// path would be a thing the photographs could not see.
static void DrawStripSoft(StripState* state, Surface* s, const RECT* client)
{
    SurfFill(s, client, g_palette.back);

    // The tabs are the stack's, not this window's. Every window in the stack draws the same row
    // with the same one selected, which is what makes switching look like a strip standing still
    // while the page behind it changes. Off a stack this is simply one tab: our own document.
    HWND frames[MAX_STRIPS];
    int  activeIndex = 0;
    int  count = StackTabs(state->frame, frames, MAX_STRIPS, &activeIndex);

    StripLayout layout;
    ComputeLayout(state, client, frames, count, activeIndex, &layout);

    HGDIOBJ oldFont = SelectObject(s->dc, state->font ? (HGDIOBJ)state->font
                                                      : GetStockObject(DEFAULT_GUI_FONT));
    SetBkMode(s->dc, TRANSPARENT);

    // Where the active card ends up, so the hairline below can be drawn around it rather than
    // through it. Left at nothing when no card is active, which is the empty row.
    int skipLeft = 0, skipRight = 0;

    // The tab being carried, if it is one of ours. Held out of the loop and drawn afterwards, so it
    // is on top of the tabs it is passing over rather than half under them.
    int carried = -1;
    if (g_dragging && g_dragFrame)
    {
        for (int i = 0; i < layout.count; i++)
            if (frames[i] == g_dragFrame)
                carried = i;
    }

    // Everything from here to the matching SurfClip(s, NULL) is confined to the track. A tab is
    // drawn at its full width and the pixels past the end are discarded, which is what a scrolled
    // row looks like: the tab at the edge is *cut*, not shortened.
    SurfClip(s, &layout.track);

    for (int i = 0; i < layout.count; i++)
    {
        if (i == carried)
            continue;

        RECT tab = layout.tab[i];
        if (tab.right <= tab.left || tab.left >= layout.track.right)
            break;
        if (tab.right <= layout.track.left)
            continue;                       // scrolled off the left-hand end

        // A tab with its context menu open is drawn hot for as long as the menu is up, which is
        // the only thing on screen saying which document those commands are about.
        BOOL hotTab = ((state->hotFrame == frames[i]) &&
                       (state->hotKind == HIT_TAB || state->hotKind == HIT_CLOSE)) ||
                      (state->menuFrame && state->menuFrame == frames[i]);

        // `first` suppresses the separator rule on the leftmost tab. In a scrolled row the leftmost
        // tab on screen is not tab zero, and a rule drawn against the edge of the track reads as a
        // border on the strip rather than as a divider between two tabs.
        BOOL first = (i == 0) || (tab.left <= layout.track.left);

        DrawOneTab(state, s, frames[i], i, tab, layout.close[i],
                   (i == activeIndex), hotTab, FALSE, first, FrameModified(frames[i]));

        if (i == activeIndex)
        {
            int inset = Scaled(TAB_LOGICAL_INSET, state->dpi);
            skipLeft  = tab.left + inset;
            skipRight = tab.right - inset;
            if (skipLeft  < layout.track.left)  skipLeft  = layout.track.left;
            if (skipRight > layout.track.right) skipRight = layout.track.right;
        }
    }

    if (carried >= 0)
    {
        // Offset by however far the tab has been taken from the slot it currently occupies. The row
        // has already rearranged underneath it - StackMoveTab runs live during the drag - so this
        // offset shrinks back to nothing as the tab arrives over its new position, and the tab is
        // never drawn in two places or missing from one.
        LONG shift = g_dragLeft - layout.tab[carried].left;

        RECT tab = layout.tab[carried];
        OffsetRect(&tab, shift, 0);

        RECT close = layout.close[carried];
        if (!IsRectEmpty(&close))
            OffsetRect(&close, shift, 0);

        if (tab.left < layout.track.right && tab.right > tab.left)
        {
            DrawOneTab(state, s, g_dragFrame, carried, tab, close,
                       (carried == activeIndex), TRUE, TRUE, FALSE, FrameModified(g_dragFrame));

            if (carried == activeIndex)
            {
                int inset = Scaled(TAB_LOGICAL_INSET, state->dpi);
                skipLeft  = tab.left + inset;
                skipRight = tab.right - inset;
                if (skipLeft  < layout.track.left)  skipLeft  = layout.track.left;
                if (skipRight > layout.track.right) skipRight = layout.track.right;
            }
        }
    }

    // Out of the track: the buttons beside it are drawn on the strip itself, and they are the one
    // thing in this row that a scroll can never move.
    SurfClip(s, NULL);

    if (layout.hasNav)
    {
        // A direction with nowhere left to go is drawn faint and is not a target - see HitTestStrip.
        // Half way to the well is enough to read as unavailable without the button disappearing,
        // which would make the row jump sideways every time it reached an end.
        COLORREF dim = Mix(g_palette.glyph, g_palette.back, 55);

        BOOL hotPrev  = (state->hotKind == HIT_PREV);
        BOOL downPrev = (state->pressKind == HIT_PREV);
        BOOL hotNext  = (state->hotKind == HIT_NEXT);
        BOOL downNext = (state->pressKind == HIT_NEXT);

        GdiFlush();
        if (layout.canPrev)
        {
            DrawChip(s, &layout.prev, hotPrev, downPrev, state->dpi);
            DrawGlyph(s, &layout.prev, (hotPrev || downPrev) ? g_palette.glyphHot : g_palette.glyph,
                      state->dpi, GLYPH_PREV);
        }
        else
        {
            DrawGlyph(s, &layout.prev, dim, state->dpi, GLYPH_PREV);
        }

        if (layout.canNext)
        {
            DrawChip(s, &layout.next, hotNext, downNext, state->dpi);
            DrawGlyph(s, &layout.next, (hotNext || downNext) ? g_palette.glyphHot : g_palette.glyph,
                      state->dpi, GLYPH_NEXT);
        }
        else
        {
            DrawGlyph(s, &layout.next, dim, state->dpi, GLYPH_NEXT);
        }
    }

    if (layout.hasPlus)
    {
        // Checked against where the button actually is, not just against the flag. The hot flag is
        // written on mouse movement, but the plus *moves* when the number of tabs changes - and the
        // number of tabs changes with the pointer sitting perfectly still, every time a document
        // opens or closes. Emptying the row moves it a whole tab width to the left edge, and without
        // this the strip would light a chip there under nothing at all. Same principle as the strip
        // measuring where it really is rather than trusting where it last put itself.
        BOOL hotPlus  = (state->hotKind == HIT_PLUS);
        if (hotPlus && state->strip && IsWindow(state->strip))
        {
            POINT cursor;
            if (GetCursorPos(&cursor) && ScreenToClient(state->strip, &cursor))
                hotPlus = PtInRect(&layout.plus, cursor) ? TRUE : FALSE;
        }
        BOOL downPlus = (state->pressKind == HIT_PLUS);
        GdiFlush();
        DrawChip(s, &layout.plus, hotPlus, downPlus, state->dpi);
        DrawGlyph(s, &layout.plus, hotPlus || downPlus ? g_palette.glyphHot : g_palette.glyph,
                  state->dpi, GLYPH_PLUS);
    }

    // The empty row, which since the last slice is a state rather than an accident: this window has
    // no document open, it is not in the stack, and the row has nothing to list. It used to be a
    // bare band with a + in the corner, which reads as a row that has failed to draw.
    //
    // The message is painted and nothing more. It is deliberately not part of the +'s hit rectangle,
    // because that rectangle is the one thing tools\check-startscreen.ps1 uses to prove the row is
    // empty without looking at a single pixel, and widening it would put the + under the point that
    // suite clicks to prove nothing is there.
    if (count == 0 && s->dc)
    {
        RECT text = *client;
        text.left = (layout.hasPlus ? layout.plus.right : client->left)
                    + Scaled(TAB_LOGICAL_PAD * 2, state->dpi);
        text.right -= Scaled(TAB_LOGICAL_PAD, state->dpi);
        if (text.right > text.left)
        {
            SetTextColor(s->dc, g_palette.textIdle);
            DrawTextW(s->dc, L"No document open", -1, &text,
                      DT_SINGLELINE | DT_VCENTER | DT_LEFT | DT_END_ELLIPSIS | DT_NOPREFIX);
        }
    }

    SelectObject(s->dc, oldFont);
    GdiFlush();

    // A hairline along the bottom, so the strip reads as part of Word's chrome rather than as a
    // rectangle dropped on top of it - except under the active card, which runs down to meet the
    // document instead. That gap is the whole difference between a row of buttons and a row of tabs:
    // one of them belongs to what is underneath it.
    RECT line = *client;
    line.top = line.bottom - 1;

    if (skipRight > skipLeft)
    {
        RECT left = line;
        left.right = skipLeft;
        if (left.right > left.left)
            SurfFill(s, &left, g_palette.edge);

        RECT right = line;
        right.left = skipRight;
        if (right.right > right.left)
            SurfFill(s, &right, g_palette.edge);
    }
    else
    {
        SurfFill(s, &line, g_palette.edge);
    }
}

// The flat renderer, which is what the strip looked like before this slice.
//
// Two jobs: it is what `HKCU\Software\WordTab\TabStyle=0` gives back, so a visual regression can be
// bisected the way every other piece of this add-in can; and it is what happens if the surface
// cannot be made at all, because a band that fails to draw inside Word is worse than a plain one.
static void DrawStripFlat(StripState* state, HDC dc, const RECT* client)
{
    HBRUSH back = CreateSolidBrush(g_palette.back);
    HBRUSH card = CreateSolidBrush(g_palette.selected);
    HBRUSH hover = CreateSolidBrush(g_palette.hover);
    HBRUSH edge = CreateSolidBrush(g_palette.edge);
    HPEN   pen  = CreatePen(PS_SOLID, 1, g_palette.edge);

    if (back)
        FillRect(dc, client, back);

    HWND frames[MAX_STRIPS];
    int  activeIndex = 0;
    int  count = StackTabs(state->frame, frames, MAX_STRIPS, &activeIndex);

    StripLayout layout;
    ComputeLayout(state, client, frames, count, activeIndex, &layout);

    HGDIOBJ oldFont = SelectObject(dc, state->font ? (HGDIOBJ)state->font
                                                   : GetStockObject(DEFAULT_GUI_FONT));
    SetBkMode(dc, TRANSPARENT);

    // The same rule as the composited renderer: a scrolled row is cut off at the track, and the
    // buttons beside it are drawn afterwards on the unclipped device context. GDI does this for us -
    // this is the one place where having a real DC is simpler than owning the pixels.
    // Applied only if the state could be saved first. A clip left on a device context that is not
    // ours - WM_PRINTCLIENT arrives with somebody else's - would truncate whatever they drew next.
    int savedDc = SaveDC(dc);
    if (savedDc)
        IntersectClipRect(dc, layout.track.left, layout.track.top,
                          layout.track.right, layout.track.bottom);

    for (int i = 0; i < layout.count; i++)
    {
        RECT tab = layout.tab[i];
        if (tab.right <= tab.left || tab.left >= layout.track.right)
            break;
        if (tab.right <= layout.track.left)
            continue;

        BOOL selected = (i == activeIndex);
        BOOL hotTab = ((state->hotFrame == frames[i]) &&
                       (state->hotKind == HIT_TAB || state->hotKind == HIT_CLOSE)) ||
                      (state->menuFrame && state->menuFrame == frames[i]);

        if (selected && card)     FillRect(dc, &tab, card);
        else if (hotTab && hover) FillRect(dc, &tab, hover);

        if (pen)
        {
            HGDIOBJ oldPen   = SelectObject(dc, pen);
            HGDIOBJ oldBrush = SelectObject(dc, GetStockObject(NULL_BRUSH));
            Rectangle(dc, tab.left, tab.top, tab.right, tab.bottom + 1);
            SelectObject(dc, oldBrush);
            SelectObject(dc, oldPen);
        }

        wchar_t title[256];
        StripTabName(frames[i], title, 256);

        BOOL hasClose = !IsRectEmpty(&layout.close[i]);
        RECT text = tab;
        text.left += Scaled(10, state->dpi);
        text.right = hasClose ? (layout.close[i].left - Scaled(4, state->dpi))
                              : (tab.right - Scaled(8, state->dpi));
        DrawTabName(state, dc, &text, i, title,
                    selected ? g_palette.text : g_palette.textIdle);

        if (hasClose && FrameModified(frames[i]) && !hotTab)
        {
            // The same swap as the composited renderer, in this one's idiom: a GDI ellipse from a
            // brush, aliased like the x it replaces. The dot ships here rather than only on the
            // styled path because TabStyle=0 exists so that a visual regression can be bisected -
            // and a fallback that quietly drops a feature is a fallback that cannot be compared
            // against.
            HBRUSH ink = CreateSolidBrush(g_palette.glyph);
            if (ink)
            {
                RECT box = layout.close[i];
                int cx = (box.left + box.right) / 2;
                int cy = (box.top + box.bottom) / 2;
                int r  = Scaled(DOT_LOGICAL_RADIUS, state->dpi);

                HGDIOBJ oldBrush = SelectObject(dc, ink);
                HGDIOBJ oldPen   = SelectObject(dc, GetStockObject(NULL_PEN));
                Ellipse(dc, cx - r, cy - r, cx + r + 1, cy + r + 1);
                SelectObject(dc, oldPen);
                SelectObject(dc, oldBrush);
                DeleteObject(ink);
            }
        }
        else if (hasClose)
        {
            HPEN glyph = CreatePen(PS_SOLID, Scaled(1, state->dpi), g_palette.glyph);
            if (glyph)
            {
                RECT box = layout.close[i];
                int inset = Scaled(5, state->dpi);
                HGDIOBJ oldPen = SelectObject(dc, glyph);
                MoveToEx(dc, box.left + inset, box.top + inset, NULL);
                LineTo(dc, box.right - inset, box.bottom - inset);
                MoveToEx(dc, box.right - inset - 1, box.top + inset, NULL);
                LineTo(dc, box.left + inset - 1, box.bottom - inset);
                SelectObject(dc, oldPen);
                DeleteObject(glyph);
            }
        }
    }

    if (savedDc)
        RestoreDC(dc, savedDc);

    if (layout.hasNav)
    {
        // Two chevrons, in the flat renderer's idiom: aliased lines from a pen, no chip, and dim
        // rather than absent at the ends of the row.
        for (int which = 0; which < 2; which++)
        {
            RECT box = which ? layout.next : layout.prev;
            BOOL live = which ? layout.canNext : layout.canPrev;

            HPEN glyph = CreatePen(PS_SOLID, Scaled(1, state->dpi),
                                   live ? g_palette.glyph : Mix(g_palette.glyph, g_palette.back, 55));
            if (!glyph)
                continue;

            int cx  = (box.left + box.right) / 2;
            int cy  = (box.top + box.bottom) / 2;
            int arm = Scaled(4, state->dpi);
            int half = Scaled(2, state->dpi);
            int tip = which ? (cx + half) : (cx - half);
            int back = which ? (cx - half) : (cx + half);

            HGDIOBJ oldPen = SelectObject(dc, glyph);
            MoveToEx(dc, back, cy - arm, NULL);
            LineTo(dc, tip, cy);
            LineTo(dc, back, cy + arm + 1);
            SelectObject(dc, oldPen);
            DeleteObject(glyph);
        }
    }

    if (layout.hasPlus)
    {
        HPEN glyph = CreatePen(PS_SOLID, Scaled(1, state->dpi), g_palette.glyph);
        if (glyph)
        {
            int cx  = (layout.plus.left + layout.plus.right) / 2;
            int cy  = (layout.plus.top + layout.plus.bottom) / 2;
            int arm = Scaled(5, state->dpi);
            HGDIOBJ oldPen = SelectObject(dc, glyph);
            MoveToEx(dc, cx - arm, cy, NULL);
            LineTo(dc, cx + arm + 1, cy);
            MoveToEx(dc, cx, cy - arm, NULL);
            LineTo(dc, cx, cy + arm + 1);
            SelectObject(dc, oldPen);
            DeleteObject(glyph);
        }
    }

    SelectObject(dc, oldFont);

    RECT line = *client;
    line.top = line.bottom - 1;
    if (edge)
        FillRect(dc, &line, edge);

    if (back)  DeleteObject(back);
    if (card)  DeleteObject(card);
    if (hover) DeleteObject(hover);
    if (edge)  DeleteObject(edge);
    if (pen)   DeleteObject(pen);
}

// The back buffer, made on demand and kept. `reference` is only used for its colour format, so it is
// as valid to build one against the foreign DC WM_PRINTCLIENT arrives with as against our own.
static BOOL EnsureSurface(StripState* state, HDC reference, int w, int h)
{
    if (state->dib && state->bits && state->memDc && state->dibW == w && state->dibH == h)
        return TRUE;

    ReleaseSurface(state);
    if (w <= 0 || h <= 0 || w > 32768 || h > 4096)
        return FALSE;

    BITMAPINFO info;
    memset(&info, 0, sizeof(info));
    info.bmiHeader.biSize        = sizeof(info.bmiHeader);
    info.bmiHeader.biWidth       = w;
    info.bmiHeader.biHeight      = -h;              // top-down, so row 0 is the top one
    info.bmiHeader.biPlanes      = 1;
    info.bmiHeader.biBitCount    = 32;
    info.bmiHeader.biCompression = BI_RGB;

    HDC dc = CreateCompatibleDC(reference);
    if (!dc)
        return FALSE;

    void* bits = NULL;
    HBITMAP dib = CreateDIBSection(dc, &info, DIB_RGB_COLORS, &bits, NULL, 0);
    if (!dib || !bits)
    {
        if (dib) DeleteObject(dib);
        DeleteDC(dc);
        return FALSE;
    }

    state->dibOld = (HBITMAP)SelectObject(dc, dib);
    state->memDc  = dc;
    state->dib    = dib;
    state->bits   = (BYTE*)bits;
    state->dibW   = w;
    state->dibH   = h;
    return TRUE;
}

static void DrawStrip(StripState* state, HDC dc, const RECT* client)
{
    // A paint can be the first thing that happens to a strip, before any janitor tick has had a
    // chance to look at the ribbon. Something has to be on the palette by the time anything is
    // drawn with it.
    if (!g_paletteReady)
        ApplyFallbackPalette(L"first paint");

    int w = client->right - client->left;
    int h = client->bottom - client->top;

    if (g_lookEnabled && EnsureSurface(state, dc, w, h))
    {
        Surface surface;
        surface.bits = state->bits;
        surface.w    = w;
        surface.h    = h;
        surface.dc   = state->memDc;
        SurfClip(&surface, NULL);          // and the drawing narrows it to the track and back again

        DrawStripSoft(state, &surface, client);
        SurfClip(&surface, NULL);          // the DC belongs to the strip, not to this paint
        BitBlt(dc, client->left, client->top, w, h, state->memDc, 0, 0, SRCCOPY);
        return;
    }

    DrawStripFlat(state, dc, client);
}

static void PaintStrip(StripState* state, HWND hwnd)
{
    PAINTSTRUCT ps;
    HDC dc = BeginPaint(hwnd, &ps);
    if (!dc)
        return;

    RECT client;
    GetClientRect(hwnd, &client);

    // The double buffer used to live here: a compatible bitmap made and thrown away every paint,
    // because hover repaints the strip whenever the pointer crosses a tab boundary and painting
    // straight to the screen shows the background fill before the tabs land on it - a flicker under
    // the pointer, exactly where the user is looking.
    //
    // It has moved into DrawStrip, and is now a 32-bit DIB kept for the life of the strip. Two
    // reasons. Compositing needs a buffer it can read back, which a screen-compatible bitmap is not;
    // and WM_PRINTCLIENT went through DrawStrip *without* the buffer, so the photographs the check
    // scripts take were of a different code path from the one on screen. Now there is one path.
    DrawStrip(state, dc, &client);

    EndPaint(hwnd, &ps);
}

// ---------------------------------------------------------------------------------------------
// The carried tab, once the row has let go of it.
//
// Dragging a tab clear of the strip takes its document out into a window of its own, and until now
// the only two things that said so were the row snapping back to its pick-up order and the cursor
// becoming IDC_SIZEALL. Both halves of the tear-off writeup name the same gap in the same words: a
// browser shows the thing being carried, and this showed a cursor. It is worse than it sounds,
// because out there the tab is drawn *in the row, in the slot it came from, not moving* - so the
// pointer walks away from a picture of the tab it is supposedly holding.
//
// This is that picture: one tab card, painted by the same DrawOneTab the row paints a lifted tab
// with, in a small popup that follows the pointer.
//
// **It costs nothing per mouse-move in the strip, and that is why it is a window rather than
// something drawn into the strips.** A carried tab INSIDE the row repaints every strip in the stack
// on every movement - measured, in the gesture slice, leaving the add-in a second behind the hand -
// which is exactly why the torn branch of DragMove stopped doing that work. Moving a popup is one
// SetWindowPos that repaints nothing: not the strips, not Word. The feedback that was missing is
// added on the side of the gesture that had spare time, not the side that did not.
//
// Owned by the frame, and WS_EX_TRANSPARENT | WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW for the tooltip's
// three reasons - no taskbar button, no Alt+Tab entry, the pointer passes through. Here the third is
// not cosmetic: **the drop is resolved with WindowFromPoint**, and a window sitting under the cursor
// that answered for itself would make every tear-off and every rejoin land on the thing being
// dragged rather than on what it was aimed at.
// ---------------------------------------------------------------------------------------------

static ATOM g_ghostClass = 0;

// The one ghost that can be up, and the strip carrying it. Global for the same reason the drag is:
// there is one pointer, so there is one tab being carried. The *window* is per strip, because a
// popup's owner is fixed at creation and each strip belongs to a different frame.
static HWND g_ghostStrip = NULL;

// Where the close button sits inside the card, so the ghost draws the same tab the row drew rather
// than a second guess at where a close button goes. Empty when that tab had none - a tab squeezed
// below the width that fits one, or a document with unsaved changes showing its dot instead.
static RECT g_ghostClose = { 0, 0, 0, 0 };

static void GhostStop(void)
{
    if (g_ghostStrip && IsWindow(g_ghostStrip))
    {
        StripState* state = FindByStrip(g_ghostStrip);
        if (state && state->ghost && IsWindow(state->ghost) && IsWindowVisible(state->ghost))
        {
            ShowWindow(state->ghost, SW_HIDE);
            LogWrite(L"ghost  hwnd=0x%p  put away", (void*)state->frame);
        }
    }
    g_ghostStrip = NULL;
    SetRectEmpty(&g_ghostClose);
}

static void DrawGhost(StripState* state, HDC dc, const RECT* client)
{
    int w = client->right - client->left;
    int h = client->bottom - client->top;
    if (w <= 0 || h <= 0 || !g_dragFrame)
        return;

    // Its own buffer rather than the strip's. EnsureSurface keeps exactly one DIB per strip and
    // rebuilds it whenever the size asked for changes, so borrowing it here - at a different size -
    // would throw the row's buffer away and rebuild it on every paint for the length of the drag.
    // This one is made and destroyed per paint, which is affordable because a ghost is painted when
    // it appears and not again: moving a window does not repaint its contents.
    BITMAPINFO info;
    memset(&info, 0, sizeof(info));
    info.bmiHeader.biSize        = sizeof(info.bmiHeader);
    info.bmiHeader.biWidth       = w;
    info.bmiHeader.biHeight      = -h;             // top-down, so row 0 is the top one
    info.bmiHeader.biPlanes      = 1;
    info.bmiHeader.biBitCount    = 32;
    info.bmiHeader.biCompression = BI_RGB;

    HDC mem = CreateCompatibleDC(dc);
    if (!mem)
        return;

    void*   bits = NULL;
    HBITMAP dib  = CreateDIBSection(mem, &info, DIB_RGB_COLORS, &bits, NULL, 0);
    if (!dib || !bits)
    {
        if (dib) DeleteObject(dib);
        DeleteDC(mem);
        return;
    }
    HBITMAP oldBitmap = (HBITMAP)SelectObject(mem, dib);

    Surface surface;
    surface.bits = (BYTE*)bits;
    surface.w    = w;
    surface.h    = h;
    surface.dc   = mem;
    SurfClip(&surface, NULL);

    // The strip's own background, so what is carried reads as a piece of the row rather than as a
    // card floating on nothing. It is also what the lifted card's shadow is composited against, and
    // the shadow is the only thing here that says the tab has been picked up.
    SurfFill(&surface, client, g_palette.back);

    HGDIOBJ oldFont = SelectObject(mem, state->font ? (HGDIOBJ)state->font
                                                    : GetStockObject(DEFAULT_GUI_FONT));
    SetBkMode(mem, TRANSPARENT);

    int spread = Scaled(LIFT_LOGICAL_SPREAD, state->dpi);

    RECT card;
    card.left   = client->left + spread;
    card.right  = client->right - spread;
    card.top    = client->top;
    card.bottom = client->bottom - spread;

    RECT close = g_ghostClose;
    if (!IsRectEmpty(&close))
        OffsetRect(&close, card.left, card.top);

    // Exactly the arguments the row uses for the tab it is carrying - selected, hot, lifted, and
    // `first` so no separator rule is drawn against an edge that has no neighbour. Index -1: the
    // name-was-cut record is a fact about a slot in the row, and this card is not in one.
    DrawOneTab(state, &surface, g_dragFrame, -1, card, close,
               TRUE, TRUE, TRUE, TRUE, FrameModified(g_dragFrame));

    SelectObject(mem, oldFont);

    GdiFlush();
    BitBlt(dc, client->left, client->top, w, h, mem, 0, 0, SRCCOPY);

    SelectObject(mem, oldBitmap);
    DeleteObject(dib);
    DeleteDC(mem);
}

static LRESULT CALLBACK GhostWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    StripState* state = (StripState*)GetWindowLongPtrW(hwnd, GWLP_USERDATA);

    switch (msg)
    {
    case WM_ERASEBKGND:
        return 1;

    case WM_PAINT:
        if (state)
        {
            PAINTSTRUCT ps;
            HDC dc = BeginPaint(hwnd, &ps);
            if (dc)
            {
                RECT client;
                GetClientRect(hwnd, &client);
                DrawGhost(state, dc, &client);
            }
            EndPaint(hwnd, &ps);
            return 0;
        }
        break;

    // Photographable, for the same reason the strip and the tooltip are: a check script takes its
    // pictures with PrintWindow, and DefWindowProc's answer to WM_PRINTCLIENT is an empty rectangle.
    case WM_PRINTCLIENT:
        if (state && wParam)
        {
            RECT client;
            GetClientRect(hwnd, &client);
            DrawGhost(state, (HDC)wParam, &client);
            return 0;
        }
        break;

    default:
        break;
    }

    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

static BOOL GhostEnsure(StripState* state)
{
    if (state->ghost && IsWindow(state->ghost))
        return TRUE;

    state->ghost = CreateWindowExW(WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE | WS_EX_TRANSPARENT,
                                   kGhostClass, L"", WS_POPUP,
                                   0, 0, 10, 10, state->frame, NULL, g_module, NULL);
    if (!state->ghost)
    {
        LogWrite(L"ghost  hwnd=0x%p  CreateWindowEx FAILED (lastError=%lu)",
                 (void*)state->frame, GetLastError());
        return FALSE;
    }

    SetWindowLongPtrW(state->ghost, GWLP_USERDATA, (LONG_PTR)state);
    return TRUE;
}

// Put the ghost under the pointer, creating and showing it the first time.
//
// `tab` and `close` are the rectangles the row would have drawn that tab with, in strip client
// coordinates, so the card carried away is the size the card in the row was - including the squeezed
// width of an overflowing row, which is the case where a fixed size would look most wrong.
static void GhostFollow(StripState* state, HWND strip, const RECT* tab, const RECT* close)
{
    if (!g_ghostEnabled || !state || !tab)
        return;

    POINT cursor;
    if (!GetCursorPos(&cursor))
        return;

    int spread = Scaled(LIFT_LOGICAL_SPREAD, state->dpi);
    int cardW  = tab->right - tab->left;
    int cardH  = tab->bottom - tab->top;
    if (cardW <= 0 || cardH <= 0)
        return;

    // Wider and taller than the card by the shadow's reach, or the lift is drawn into pixels the
    // window does not own and the card comes out square-edged at the bottom - which is the one part
    // of it that says it is off the row rather than in it.
    int w = cardW + spread * 2;
    int h = cardH + spread;

    BOOL fresh = (g_ghostStrip != strip);
    if (fresh)
    {
        if (!GhostEnsure(state))
            return;
        g_ghostStrip = strip;
    }

    if (close && !IsRectEmpty(close))
    {
        g_ghostClose = *close;
        OffsetRect(&g_ghostClose, -tab->left, -tab->top);
    }
    else
    {
        SetRectEmpty(&g_ghostClose);
    }

    // Carried from the point inside the tab that was grabbed - the same g_dragGrabDx the row uses,
    // and g_dragPressY for the other axis - so the card does not jump under the hand at the moment
    // the row lets go of it. That instant is already the one where the most changes at once.
    int x = cursor.x - g_dragGrabDx - spread;
    int y = cursor.y - g_dragPressY;

    SetWindowPos(state->ghost, HWND_TOP, x, y, w, h,
                 SWP_NOACTIVATE | (fresh ? SWP_SHOWWINDOW : 0));

    if (fresh)
    {
        // Once per gesture, not per movement. The position is deliberately absent: it changes with
        // every mouse message and the log is read by eye.
        LogWrite(L"ghost  hwnd=0x%p  carrying 0x%p  (%dx%d, grabbed %dpx in)",
                 (void*)state->frame, (void*)g_dragFrame, w, h, g_dragGrabDx);
    }
}

// ---------------------------------------------------------------------------------------------
// The context menu, owner-drawn.
//
// A Win32 popup menu is a system menu. It takes its colours from the system, and on Windows 11 the
// system's menu colours are light whatever the app around them looks like - so on a dark Word the
// tab menu was a white rectangle. That was recorded as a gap when the menu shipped
// (RESULT-menu.md), with the two possible routes named: owner-draw it, or use the undocumented
// uxtheme ordinals that would restyle menus for the whole of Word. This is the first route. The
// second was never a candidate: it changes a host application's rendering globally, from an add-in.
//
// Owner-drawing a menu draws its *items*. The window behind them, its margins, its border and its
// Windows 11 rounded corners stay the system's; SetMenuInfo with MIM_BACKGROUND is the only handle
// on the first of those and it is enough - photographed, the result is dark edge to edge.
//
// Four things were measured before any of this was written, because between them they decide whether
// the 32 checks in tools\check-menu.ps1 survive - and that suite is the only one that deliberately
// provokes Word's save prompt, so it is not one to break casually:
//
//   1. **An owner-drawn item keeps its string.** The documentation says GetMenuString has nothing to
//      return for an MFT_OWNERDRAW item. Measured, with the string supplied through MIIM_STRING in
//      the same call, it returns it: `GetMenuString=5 '&Save'`. check-menu reads every label that
//      way from outside the process and asserts all seven; nothing there had to change.
//   2. **Mnemonics still work.** Typing `n` over the owner-drawn menu returned CMD_NEW exactly as it
//      does over a system one, and WM_MENUCHAR was never sent - because the string is still there to
//      match against. No third case in the frame subclass.
//   3. **`MFT_SEPARATOR | MFT_OWNERDRAW` is a real combination**, undocumented though it is. The item
//      still reports MF_SEPARATOR to GetMenuState - which is what check-menu reads to expect a `-` -
//      *and* it is sent to us to draw. A plain MF_SEPARATOR is drawn by Windows as a bright rule
//      running the full width of the menu, which on a dark background is exactly wrong.
//   4. **WM_MEASUREITEM and WM_DRAWITEM arrive with CtlType == ODT_MENU and dwItemData intact**,
//      TPM_NONOTIFY notwithstanding - those are requests, not notifications.
//
// They arrive at the menu's *owner*, which is Word's frame and not the strip (see below for why it
// has to be), so they land in frames.cpp's subclass and are forwarded back here.
// ---------------------------------------------------------------------------------------------

#define MENU_MAGIC 0x57544144u   // 'WTAD', after 'WTAB' and 'WTAC' elsewhere in this add-in

struct MenuItemTag
{
    DWORD   magic;
    wchar_t label[40];      // empty for a separator
};

// Sized with headroom on purpose. The menu is nine items at its longest (Save, Move to New Window, a
// separator, four closes, a separator, New Document) and running out is not an assertion failure or a
// crash - the item is simply left off the owner-draw path and Windows draws it in the default style,
// which on a dark theme is one light grey row at the bottom of a dark menu. Every text and state
// assertion in check-menu still passes while that is on screen, so the cap has to be generous and the
// overflow has to say so. See MenuAddItem.
static MenuItemTag g_menuItems[12];
static int         g_menuItemCount = 0;
static HMENU       g_openMenu      = NULL;
static int         g_menuDpi       = 96;
static HWND        g_menuOwner     = NULL;

// Is this item one of ours?
//
// The range check is the point, and it is not paranoia. `itemData` is a value chosen by whoever
// built the menu, and menus owned by this frame are not all ours: Word's own Alt+Space menu is a
// real HMENU, shell extensions inject items, and Office Tab is installed on these machines and could
// be doing exactly what we are doing on the same window. Dereferencing a pointer another program
// chose, to read a magic number out of it, is a wild read inside Word's process - which is the whole
// class of thing the "nothing escapes" rule exists to prevent. So: prove it points into our own
// array, on an element boundary, and only then read it.
static MenuItemTag* MenuTagOf(ULONG_PTR itemData)
{
    if (!itemData)
        return NULL;

    const char* base = (const char*)&g_menuItems[0];
    const char* p    = (const char*)itemData;
    if (p < base || p >= base + sizeof(g_menuItems))
        return NULL;
    if ((size_t)(p - base) % sizeof(MenuItemTag) != 0)
        return NULL;

    MenuItemTag* tag = (MenuItemTag*)itemData;
    return (tag->magic == MENU_MAGIC) ? tag : NULL;
}

static void MenuBuildBegin(StripState* state)
{
    g_menuItemCount = 0;
    g_menuDpi       = state->dpi > 0 ? state->dpi : 96;
    g_menuOwner     = state->frame;
    memset(g_menuItems, 0, sizeof(g_menuItems));
}

static void MenuAddItem(HMENU menu, UINT id, const wchar_t* label, BOOL enabled)
{
    MENUITEMINFOW info;
    memset(&info, 0, sizeof(info));
    info.cbSize     = sizeof(info);
    info.fMask      = MIIM_ID | MIIM_STATE | MIIM_STRING;
    info.wID        = id;
    info.fState     = enabled ? MFS_ENABLED : MFS_GRAYED;
    info.dwTypeData = (LPWSTR)label;

    // MF_GRAYED has to stay on the item itself and not merely be drawn: it is the state
    // check-menu.ps1 reads through GetMenuState to assert that Close Others is unavailable with one
    // document open, and an owner-drawn item Windows does no graying for.
    if (g_lookEnabled && g_menuItemCount >= (int)(sizeof(g_menuItems) / sizeof(g_menuItems[0])))
    {
        // Not fatal, and that is exactly why it is logged: what the user sees is one default-styled
        // row in an otherwise themed menu, which reads as a rendering glitch rather than as a limit
        // somebody needs to raise.
        LogWrite(L"strip  menu item |%s| past the owner-draw limit of %d - Windows will draw it",
                 label, (int)(sizeof(g_menuItems) / sizeof(g_menuItems[0])));
    }

    if (g_lookEnabled && g_menuItemCount < (int)(sizeof(g_menuItems) / sizeof(g_menuItems[0])))
    {
        MenuItemTag* tag = &g_menuItems[g_menuItemCount++];
        tag->magic = MENU_MAGIC;
        wcsncpy(tag->label, label, 39);
        tag->label[39] = 0;

        info.fMask     |= MIIM_FTYPE | MIIM_DATA;
        info.fType      = MFT_OWNERDRAW;
        info.dwItemData = (ULONG_PTR)tag;
    }

    InsertMenuItemW(menu, GetMenuItemCount(menu), TRUE, &info);
}

static void MenuAddSeparator(HMENU menu)
{
    if (!g_lookEnabled || g_menuItemCount >= (int)(sizeof(g_menuItems) / sizeof(g_menuItems[0])))
    {
        AppendMenuW(menu, MF_SEPARATOR, 0, NULL);
        return;
    }

    MenuItemTag* tag = &g_menuItems[g_menuItemCount++];
    tag->magic    = MENU_MAGIC;
    tag->label[0] = 0;

    MENUITEMINFOW info;
    memset(&info, 0, sizeof(info));
    info.cbSize     = sizeof(info);
    info.fMask      = MIIM_FTYPE | MIIM_DATA;
    info.fType      = MFT_SEPARATOR | MFT_OWNERDRAW;
    info.dwItemData = (ULONG_PTR)tag;
    InsertMenuItemW(menu, GetMenuItemCount(menu), TRUE, &info);
}

static void MenuBuildEnd(HMENU menu)
{
    if (!g_lookEnabled || !g_menuBackBrush)
        return;

    // Everything outside the item rectangles - the gutter down the left, the margin at top and
    // bottom - belongs to the menu window, not to us. This is the whole of our say in it.
    MENUINFO mi;
    memset(&mi, 0, sizeof(mi));
    mi.cbSize  = sizeof(mi);
    mi.fMask   = MIM_BACKGROUND | MIM_APPLYTOSUBMENUS;
    mi.hbrBack = g_menuBackBrush;
    SetMenuInfo(menu, &mi);
}

// The two halves, called from frames.cpp with the frame the message arrived on. Both answer FALSE
// for anything that is not one of ours, and the caller must chain those on untouched.
//
// They run inside TrackPopupMenu's modal loop, on Word's UI thread. The rules there are the ones
// ShowTabMenu already lives by: look nothing up by a pointer captured earlier, allocate nothing per
// item, and never log - WM_DRAWITEM fires per item and again per item on every hover change, and a
// log line is a file write.
BOOL StripOnMenuMeasure(HWND frame, MEASUREITEMSTRUCT* item)
{
    if (!item || item->CtlType != ODT_MENU)
        return FALSE;

    MenuItemTag* tag = MenuTagOf(item->itemData);
    if (!tag)
        return FALSE;

    int dpi = g_menuDpi;
    if (tag->label[0] == 0)
    {
        item->itemWidth  = 0;
        item->itemHeight = (UINT)Scaled(7, dpi);
        return TRUE;
    }

    // Measured against the real font, because a menu sized from a guess is a menu with its longest
    // item clipped at some DPI and not at others.
    HDC dc = GetDC(frame);
    SIZE size;
    size.cx = Scaled(120, dpi);
    size.cy = Scaled(16, dpi);
    if (dc)
    {
        StripState* state = FindByFrame(frame);
        HGDIOBJ old = SelectObject(dc, (state && state->font) ? (HGDIOBJ)state->font
                                                              : GetStockObject(DEFAULT_GUI_FONT));
        GetTextExtentPoint32W(dc, tag->label, (int)wcslen(tag->label), &size);
        SelectObject(dc, old);
        ReleaseDC(frame, dc);
    }

    // The empty column on the left is where Windows would put a check mark or an icon. Leaving room
    // for it is most of what makes an owner-drawn menu still read as a menu.
    item->itemWidth  = (UINT)(size.cx + Scaled(28 + 28, dpi));
    item->itemHeight = (UINT)(size.cy + Scaled(9, dpi));
    if (item->itemHeight < (UINT)Scaled(24, dpi))
        item->itemHeight = (UINT)Scaled(24, dpi);
    return TRUE;
}

BOOL StripOnMenuDraw(HWND frame, DRAWITEMSTRUCT* item)
{
    if (!item || item->CtlType != ODT_MENU || !item->hDC)
        return FALSE;

    MenuItemTag* tag = MenuTagOf(item->itemData);
    if (!tag)
        return FALSE;

    // For a menu, hwndItem is the HMENU. Comparing it against the one we put up makes this exact
    // rather than merely well-guarded, and costs one comparison.
    if (g_openMenu && (HMENU)item->hwndItem != g_openMenu)
        return FALSE;

    HDC  dc  = item->hDC;
    RECT box = item->rcItem;
    int  dpi = g_menuDpi;

    HBRUSH back = CreateSolidBrush(g_palette.menuBack);
    if (back)
    {
        FillRect(dc, &box, back);
        DeleteObject(back);
    }

    if (tag->label[0] == 0)
    {
        RECT rule = box;
        rule.top    = (box.top + box.bottom) / 2;
        rule.bottom = rule.top + 1;
        rule.left  += Scaled(10, dpi);
        rule.right -= Scaled(10, dpi);
        HBRUSH line = CreateSolidBrush(g_palette.menuLine);
        if (line)
        {
            FillRect(dc, &rule, line);
            DeleteObject(line);
        }
        return TRUE;
    }

    BOOL grayed   = (item->itemState & (ODS_GRAYED | ODS_DISABLED)) != 0;
    BOOL selected = (item->itemState & ODS_SELECTED) != 0;

    // A disabled item does not highlight. Windows would not have highlighted it either, and an
    // unavailable command that lights up under the pointer is an invitation to click it.
    if (selected && !grayed)
    {
        RECT hot = box;
        hot.left   += Scaled(3, dpi);
        hot.right  -= Scaled(3, dpi);
        hot.top    += Scaled(1, dpi);
        hot.bottom -= Scaled(1, dpi);
        HBRUSH brush = CreateSolidBrush(g_palette.menuHot);
        if (brush)
        {
            FillRect(dc, &hot, brush);
            DeleteObject(brush);
        }
    }

    StripState* state = FindByFrame(frame);
    HGDIOBJ oldFont = SelectObject(dc, (state && state->font) ? (HGDIOBJ)state->font
                                                             : GetStockObject(DEFAULT_GUI_FONT));
    int oldMode = SetBkMode(dc, TRANSPARENT);
    SetTextColor(dc, grayed ? g_palette.menuTextDim : g_palette.menuText);

    RECT text = box;
    text.left += Scaled(28, dpi);

    // ODS_NOACCEL means the user has not pressed Alt, so the mnemonic underlines are not being
    // shown anywhere else either. Honouring it is the difference between a menu that matches the
    // rest of Windows and one that always looks like Alt is held down.
    UINT format = DT_SINGLELINE | DT_VCENTER | DT_LEFT;
    if (item->itemState & ODS_NOACCEL)
        format |= DT_HIDEPREFIX;
    DrawTextW(dc, tag->label, -1, &text, format);

    SetBkMode(dc, oldMode);
    SelectObject(dc, oldFont);
    return TRUE;
}

// The context menu.
//
// A plain Win32 popup, built and thrown away each time it is shown, because everything on it depends
// on what is true at that instant: how many tabs there are, and whether the pointer was over one.
// There is no menu to keep in sync with the tab row if there is no menu between right-clicks.
//
// TPM_RETURNCMD is what makes this small. The chosen id comes back as the return value, so there is
// no WM_COMMAND to route, no id space to keep clear of Word's own thousands of command ids, and no
// window that has to still exist by the time a command arrives.
//
// **Nothing is done from inside here.** TrackPopupMenu runs a modal loop - Word's timers fire, the
// janitor runs, documents can open and close - so by the time it returns, the state this function
// started with may be gone, including `state` itself and the strip window. The command is posted and
// acted on in a fresh message, where everything is looked up again.
static void ShowTabMenu(HWND hwnd, POINT client, HWND target)
{
    StripState* state = (StripState*)GetWindowLongPtrW(hwnd, GWLP_USERDATA);
    if (!state)
        return;

    HMENU menu = CreatePopupMenu();
    if (!menu)
        return;

    int count = StackTabs(state->frame, NULL, MAX_TABS, NULL);

    MenuBuildBegin(state);

    if (target)
    {
        MenuAddItem(menu, CMD_SAVE, L"&Save", TRUE);

        // Grouped with Save rather than with the closes, because it is the other item here that does
        // not destroy anything - and it deliberately does not go next to "Close All". Greyed on a
        // single tab by the same rule as Close Others: an item that appears and disappears makes the
        // menu a different shape each time and moves everything below it.
        //
        // Offered only when the stack can actually do it, so TabTearOff=0 removes the item rather
        // than leaving a permanently grey one that says the feature exists and refuses to work.
        if (StackCanTearOff())
            MenuAddItem(menu, CMD_TEAROFF, L"&Move to New Window", count > 1);

        // ...and the way back, offered only on a window that is actually out of a stack, which is the
        // only place it means anything. Unlike the item above this one appears and disappears, and
        // that is the right way round: "Move to New Window" is a thing every tab could do, so it
        // stays put and greys; this is a thing only a window standing on its own can do, and on every
        // other tab it would be a permanently grey line saying nothing.
        //
        // It exists because the gesture alone is not discoverable. A user who took this window out
        // through the menu will look in the menu to put it back, and the drag is not something they
        // can be expected to guess.
        if (StackCanRejoin(target))
            MenuAddItem(menu, CMD_JOIN, L"Move &Back to the Tab Row", TRUE);

        MenuAddSeparator(menu);
        MenuAddItem(menu, CMD_CLOSE, L"&Close", TRUE);
        MenuAddItem(menu, CMD_CLOSE_OTHERS, L"Close &Others", count > 1);

        // Greyed on the last tab rather than hidden. An item that comes and goes depending on which
        // tab was right-clicked makes the menu a different shape each time and moves everything
        // below it - and "Close All" moving under the pointer is the one thing on here worth being
        // careful about.
        MenuAddItem(menu, CMD_CLOSE_RIGHT, L"Close Tabs to the &Right", StackTabsRightOf(target) > 0);

        MenuAddItem(menu, CMD_CLOSE_ALL, L"Close &All", TRUE);
        MenuAddSeparator(menu);
    }

    // On the empty part of the strip this is the whole menu. A right-click that produces nothing at
    // all reads as a dead area rather than as a deliberate one, and this is the command that has
    // nothing to do with any particular tab.
    MenuAddItem(menu, CMD_NEW, L"&New Document", TRUE);

    MenuBuildEnd(menu);

    // The tab stays lit for as long as the menu is up. A tab can be narrow enough that its name is
    // an ellipsis, and "Close All" arriving from a menu the user is no longer sure they aimed
    // correctly is not a comfortable thing to click.
    state->menuFrame = target;
    InvalidateRect(hwnd, NULL, FALSE);
    UpdateWindow(hwnd);                 // painted before the modal loop, not after it

    POINT screen = client;
    ClientToScreen(hwnd, &screen);

    // The owner is Word's frame, not the strip: the strip is WS_EX_NOACTIVATE and can never be the
    // foreground window, and a popup menu whose owner is not foreground does not dismiss when the
    // user clicks away from it. The WM_NULL afterwards is the other half of that rule.
    SetForegroundWindow(state->frame);

    // Set before the modal loop and cleared after it, so a WM_DRAWITEM arriving at the frame can be
    // matched against the menu it belongs to rather than merely believed.
    g_openMenu = menu;

    int chosen = (int)TrackPopupMenu(menu,
                                     TPM_RETURNCMD | TPM_NONOTIFY | TPM_LEFTALIGN | TPM_TOPALIGN |
                                     TPM_RIGHTBUTTON,
                                     screen.x, screen.y, 0, state->frame, NULL);
    g_openMenu = NULL;
    g_menuItemCount = 0;
    DestroyMenu(menu);

    // Everything from before the modal loop is re-derived: the strip may have been detached and
    // destroyed while the menu was open, and the state array may have been compacted under us.
    if (!IsWindow(hwnd))
        return;
    state = (StripState*)GetWindowLongPtrW(hwnd, GWLP_USERDATA);
    if (state)
    {
        state->menuFrame = NULL;
        state->hotKind   = HIT_NONE;    // the pointer spent the last few seconds over a menu
        state->hotFrame  = NULL;
        InvalidateRect(hwnd, NULL, FALSE);
        PostMessageW(state->frame, WM_NULL, 0, 0);
    }

    LogWrite(L"strip  hwnd=0x%p  menu on 0x%p -> command %d", (void*)hwnd, (void*)target, chosen);

    if (chosen != 0)
        PostMessageW(hwnd, WM_WORDTAB_CMD, (WPARAM)chosen, (LPARAM)target);
}

// What a right-click is about: the tab under it, or nothing. A close button counts as its own tab -
// the two overlap, and a right-click is not aimed at a button.
static HWND MenuTargetAt(StripState* state, HWND hwnd, POINT point)
{
    StripHit hit = HitTestStrip(state, hwnd, point);
    if (hit.kind == HIT_TAB || hit.kind == HIT_CLOSE)
        return hit.frame;
    return NULL;
}

// ---------------------------------------------------------------------------------------------
// Dragging a tab: along the row to reorder it, or out of the row to take its document out of the
// stack.
//
// One gesture with two meanings and one press. Which one it is is decided continuously, by where the
// pointer is when the button comes up, and it can change its mind as many times as the user likes on
// the way - the two states are the same drag with the row either holding on to the tab or having let
// go of it. What separates them is a threshold much larger than the one that starts the drag, and the
// asymmetry is deliberate: a misaimed reorder is undone by dragging again, and a tear-off the user
// did not mean is a window standing somewhere they have to go and find.
//
// Everything about *what happens* to a torn-off window belongs to the stack and was built and driven
// before this: StackTearOffTab is what the Move to New Window menu item calls, and this drop raises
// the same command by the same route. There is no second answer here to where the window goes, what
// stops the janitor putting it straight back, or when it stops being torn off.
//
// The row rearranges *live*, as the tab is carried, rather than showing an insertion marker and
// rearranging on the drop. Both are defensible; live wins here because every window in the stack
// draws the same row, so a live reorder is a thing all of them already know how to show, whereas an
// insertion marker would be state the drawing code would have to be taught. It also means the drop
// itself has nothing to do: by the time the button is released the order is already what the user
// can see, and releasing is only letting go.
//
// The consequence is that cancelling has to undo, which is why the position the tab was picked up
// from is kept for the whole gesture.
// ---------------------------------------------------------------------------------------------

// The auto-scroll timer, defined with the rest of the drag below. Declared here because the two
// functions that end a gesture come first, and the timer has to stop when the gesture does - all
// three ways it can end.
static void DragScrollStart(HWND hwnd);
static void DragScrollStop(void);

// What the pointer says while a tab is being carried.
//
// The cursor is the only feedback this add-in has outside the strip's own 32 pixels. A tab dragged
// clear of the row is somewhere the strip cannot draw - the pointer is over Word's document, or off
// the window entirely - and a child window may not paint there. So the shape of the pointer is what
// distinguishes the two things a release can now mean, and IDC_SIZEALL is Windows' own "you are
// relocating this", not a cursor invented here.
//
// System cursors are shared and cached by the window manager: LoadCursorW hands back the same handle
// every time and there is nothing to destroy. Setting it while we hold the mouse capture is enough to
// keep it - WM_SETCURSOR is not sent to anybody while the mouse is captured, so nothing else is going
// to put the arrow back underneath us.
static void DragCursor(const wchar_t* which)
{
    SetCursor(LoadCursorW(NULL, which));
}

// Forget the gesture without touching the row: a drop that was agreed to, or a strip destroyed
// underneath one.
static void DragForget(void)
{
    DragScrollStop();
    GhostStop();
    if (g_dragTorn || g_dragJoin)
        DragCursor(IDC_ARROW);
    g_dragStrip = NULL;
    g_dragFrame = NULL;
    g_dragging  = FALSE;
    g_dragTorn  = FALSE;
    g_dragJoin  = FALSE;
    g_dragOnto  = NULL;
    g_dragOntoKnown = FALSE;
}

// Put the row back exactly as it was and stop carrying the tab - but keep the capture, because the
// left button is still down and letting the mouse go now would deliver its release to whatever
// happens to be underneath the pointer.
//
// There is deliberately no Escape. The strip is WS_EX_NOACTIVATE and never holds the keyboard focus,
// so a keypress never reaches it, and reading GetAsyncKeyState between two mouse movements is
// exactly the mistake the batch close was written twice to avoid: a state that can come and go
// inside a polling interval has to arrive as an event or it is not being observed at all. The two
// cancels that *are* events - the right button, and losing the capture - both come through here, and
// dragging the tab back where it came from is the third.
static void DragUndo(HWND hwnd, const wchar_t* why)
{
    HWND frame = g_dragFrame;
    BOOL was   = g_dragging;
    BOOL join  = g_dragJoin;
    int  from  = g_dragFrom;

    g_dragFrame = NULL;
    g_dragging  = FALSE;
    if (g_dragTorn || g_dragJoin)
        DragCursor(IDC_ARROW);
    g_dragTorn  = FALSE;
    g_dragJoin  = FALSE;
    g_dragOnto  = NULL;
    g_dragOntoKnown = FALSE;
    DragScrollStop();
    GhostStop();

    if (!was || !frame || !IsWindow(frame))
        return;

    // A cancelled rejoin has nothing to put back. Carrying a window towards a stack never moved
    // anything - the whole gesture is a question that is answered on release - so the undo is the
    // absence of the join, and calling StackMoveTab here would be asking the row to reorder a tab
    // that is not in it.
    if (join)
    {
        LogWrite(L"strip  hwnd=0x%p  rejoin cancelled (%s) - 0x%p stays on its own",
                 (void*)hwnd, why, (void*)frame);
        return;
    }

    StackMoveTab(frame, from);
    LogWrite(L"strip  hwnd=0x%p  drag cancelled (%s) - 0x%p back at tab %d",
             (void*)hwnd, why, (void*)frame, from);
    StripRefreshTabs();
}

// The pointer has moved with a tab held down. Below the slop this is still a click that has not
// finished; past it the tab is being carried, and the row rearranges under it.
static void DragMove(StripState* state, HWND hwnd, POINT point)
{
    // The document behind a carried tab can go away mid-gesture: Word hiding the window, a close
    // that was already in flight, the janitor dropping it from the stack. There is then nothing to
    // carry and nowhere to put it back, so the gesture is simply over.
    // A rejoin drag is carrying a window that is deliberately NOT in the row, so "no longer in the
    // row" cannot be its abandon test - only the window going away can be.
    BOOL lost = g_dragJoin ? !IsWindow(g_dragFrame)
                           : (!IsWindow(g_dragFrame) || StackTabIndex(g_dragFrame) < 0);
    if (lost)
    {
        BOOL was = g_dragging;
        g_dragFrame = NULL;
        g_dragging  = FALSE;
        if (g_dragTorn || g_dragJoin)
            DragCursor(IDC_ARROW);
        g_dragTorn  = FALSE;
        g_dragJoin  = FALSE;
        g_dragOnto  = NULL;
        g_dragOntoKnown = FALSE;
        DragScrollStop();
        if (was)
        {
            LogWrite(L"strip  hwnd=0x%p  drag abandoned - that tab is no longer in the row",
                     (void*)hwnd);
            StripRefreshTabs();
        }
        return;
    }

    // ---------------------------------------------------------------------------------------
    // The other gesture: a window standing on its own being carried INTO a stack.
    //
    // Handled before everything below and returning, because none of it applies. There is no row to
    // reorder - this window's strip draws exactly one tab, its own - and no threshold to cross,
    // since the tab is already outside every row that matters. The only question a mouse-move can
    // answer here is which window's row the pointer is over, and the only two things that can
    // happen on release are "join that one" and "nothing".
    // ---------------------------------------------------------------------------------------------
    if (g_dragJoin)
    {
        if (!g_dragging)
        {
            int startSlop = Scaled(DRAG_LOGICAL_SLOP, state->dpi);
            int dxStart   = point.x - g_dragPressX;
            int dyStart   = point.y - g_dragPressY;
            if (dxStart > -startSlop && dxStart < startSlop &&
                dyStart > -startSlop && dyStart < startSlop)
                return;

            g_dragging = TRUE;
            LogWrite(L"strip  hwnd=0x%p  drag started on 0x%p - it is on its own, so this is a rejoin",
                     (void*)hwnd, (void*)g_dragFrame);
        }

        // Which window's row the pointer is over. WindowFromPoint reports what is under the point
        // regardless of who holds the mouse capture, which is the whole reason this can work at all:
        // the strip being dragged FROM owns the mouse for the length of the gesture, so no other
        // window will ever be told the pointer is there.
        POINT screen = point;
        ClientToScreen(hwnd, &screen);

        HWND onto = NULL;
        StripState* target = FindByStrip(WindowFromPoint(screen));
        if (target && target->frame != g_dragFrame && StackTabIndex(target->frame) >= 0)
            onto = target->frame;

        // Re-asked every move rather than cached from the press: the row underneath can change while
        // the gesture is in the air - a document closing takes a window out of it - and a drop
        // aimed at a stack that is no longer there has to be refused, not honoured against a stale
        // answer.
        if (onto && !StackCanRejoin(g_dragFrame))
            onto = NULL;

        // The `known` flag rather than a comparison against NULL: at the start of the gesture the
        // answer is genuinely "nothing", and a bare `onto != g_dragOnto` would treat the first
        // evaluation as no change and say nothing at all - so the most common state this gesture is
        // ever in would be the one state that was never announced.
        if (!g_dragOntoKnown || onto != g_dragOnto)
        {
            g_dragOntoKnown = TRUE;
            g_dragOnto = onto;

            // What was actually under the pointer, named, when the answer is "nothing".
            //
            // Without it, a rejoin that would not take looks identical from outside to a pointer that
            // never arrived - and those have completely different causes. That cost a diagnosis: a
            // drag aimed at a confirmed-correct point produced seven seconds of silence, and nothing
            // anywhere could say whether the add-in had seen the wrong window under the pointer or
            // had never been told the pointer moved. Cheap because it is written on a *change* of
            // target, not per mouse-move.
            wchar_t under[64] = L"nothing";
            HWND    at = WindowFromPoint(screen);
            if (at)
            {
                wchar_t cls[48] = L"";
                GetClassNameW(at, cls, 48);
                _snwprintf(under, 64, L"0x%p %s", (void*)at, cls);
                under[63] = L'\0';
            }

            LogWrite(L"strip  hwnd=0x%p  rejoin drag is over %s  (pointer at %d,%d over %s)",
                     (void*)hwnd,
                     onto ? L"a row it can join - letting go now puts it back"
                          : L"nothing it can join - letting go now does nothing",
                     (int)screen.x, (int)screen.y, under);
        }

        // IDC_NO is the honest answer for most of this gesture's travel, and saying so is the point:
        // there is no window preview following the hand, so without it the user is carrying an
        // invisible thing with no way of knowing whether releasing will do anything.
        DragCursor(g_dragOnto ? IDC_SIZEALL : IDC_NO);
        return;
    }

    RECT client;
    if (!GetClientRect(hwnd, &client))
        return;

    HWND frames[MAX_STRIPS];
    StripLayout layout;
    LayoutOf(state, &client, frames, &layout);

    // The row this gesture is happening *in* can empty underneath it. The drag state is global and
    // the strip holding the capture is usually not the strip on screen, so `g_dragFrame` above can
    // still be a perfectly good tab in some other window's row while *this* window loses its last
    // document. There is then nowhere to draw the carried tab and no row to drop it into, and simply
    // returning would freeze the gesture: every strip would keep drawing the tab at the last position
    // this function computed, and the release would commit it there.
    if (layout.count <= 0)
    {
        DragUndo(hwnd, L"the row it was being carried in has no documents left");
        return;
    }

    // How far the pointer is outside the strip, in the only direction the strip has any room to be
    // outside in. The mouse is captured, so this keeps arriving - and keeps being meaningful - long
    // after the pointer has left the window.
    int outBy = 0;
    if (point.y < client.top)
        outBy = client.top - point.y;
    else if (point.y > client.bottom)
        outBy = point.y - client.bottom;

    // Whether leaving the row can mean anything here. Two reasons it might not: the switch is off, or
    // this is the only tab in the stack and there is nowhere for it to go - StackTearOffTab refuses
    // both, and a cursor that promised something the drop then declines is worse than no cursor.
    BOOL canTear = StackCanTearOff() && layout.count > 1;

    int slop    = Scaled(DRAG_LOGICAL_SLOP, state->dpi);
    int tearIn  = Scaled(TEAROFF_LOGICAL_SLOP, state->dpi);
    int tearOut = tearIn / 2;      // hysteresis; see below
    int dx      = point.x - g_dragPressX;

    if (!g_dragging)
    {
        // Horizontal travel starts a reorder, exactly as it always has. Vertical travel starts
        // nothing at all until it is a whole row clear of the strip - which is the tear-off threshold
        // itself, not a slop of its own.
        //
        // That asymmetry is the point rather than an oversight. A hand that slips downwards while
        // clicking a tab must still be a click, so vertical movement cannot have a small threshold;
        // but the one gesture that has no horizontal component whatsoever - pulling a tab straight
        // down out of the row - is precisely the one this exists for, and a purely horizontal test
        // could never see it.
        // Not called `far`: windef.h still #defines that to nothing for the sake of 16-bit code, so
        // `BOOL far = ...` compiles as `BOOL = ...` and the error names the line after it.
        BOOL travelled = (dx <= -slop || dx >= slop);
        if (!travelled && canTear && outBy > tearIn)
            travelled = TRUE;
        if (!travelled)
            return;

        g_dragging = TRUE;
        DragScrollStart(hwnd);
        LogWrite(L"strip  hwnd=0x%p  drag started on 0x%p (tab %d)",
                 (void*)hwnd, (void*)g_dragFrame, g_dragFrom);
    }

    // In the row or out of it. The threshold is lower on the way back in than on the way out, because
    // the boundary is somewhere a hand can come to rest: with one threshold a pointer sitting on the
    // line flips the gesture between reordering and tearing off several times a second, and the tab
    // jumps under the hand every time it does.
    BOOL torn = canTear && (g_dragTorn ? (outBy > tearOut) : (outBy > tearIn));

    if (torn != g_dragTorn)
    {
        g_dragTorn = torn;
        DragCursor(torn ? IDC_SIZEALL : IDC_ARROW);

        if (torn)
        {
            // The row lets go. The order goes back to what it was when the tab was picked up, and
            // the tab stops following the pointer - it sits still, in the slot it came from, while
            // the pointer walks away from it.
            //
            // A gesture that changes what it means has to show something at the moment it changes,
            // and this is the only change the strip can make that is visible whether or not the tab
            // has been moved yet. Following the pointer would go on saying "this is being carried in
            // the row", which is the one thing that has stopped being true.
            StackMoveTab(g_dragFrame, g_dragFrom);
            DragScrollStop();
            LogWrite(L"strip  hwnd=0x%p  drag left the row - letting go now tears 0x%p off"
                     L"  (%dpx outside a %dpx strip, threshold %d)",
                     (void*)hwnd, (void*)g_dragFrame, outBy, (int)client.bottom, tearIn);
        }
        else
        {
            // The row has taken it back, so the thing following the pointer would now be a second
            // copy of a tab that is in the row again and moving with the hand by itself.
            GhostStop();
            DragScrollStart(hwnd);
            LogWrite(L"strip  hwnd=0x%p  drag back in the row - letting go now reorders 0x%p"
                     L"  (%dpx outside a %dpx strip, threshold %d)",
                     (void*)hwnd, (void*)g_dragFrame, outBy, (int)client.bottom, tearOut);
        }

        // StackMoveTab above has rearranged the row underneath the layout computed a few lines up.
        LayoutOf(state, &client, frames, &layout);
    }

    if (g_dragTorn)
    {
        // Restated every move rather than only on the transition. The cursor is process-wide state
        // and this costs one call; a tear-off gesture that silently reverted to an arrow half way
        // through would be a gesture with no feedback at all.
        DragCursor(IDC_SIZEALL);

        // Repainted only when the tab is not already where it belongs, which out here is at most
        // once. A carried tab moves with the pointer and every strip in the stack draws it, so the
        // reorder path below pays for a full repaint of every window per mouse-move - but a tab that
        // has been let go of by the row does not move at all, and doing that work per move buys
        // nothing and puts the strip a visible fraction of a second behind the hand. Measured: with
        // the unconditional refresh here, the add-in was still working through the moves that brought
        // the pointer back into the row a second after the pointer had arrived.
        int home = -1;
        for (int i = 0; i < layout.count; i++)
            if (frames[i] == g_dragFrame)
                home = i;

        if (home >= 0 && g_dragLeft != layout.tab[home].left)
        {
            g_dragLeft = layout.tab[home].left;
            StripRefreshTabs();
        }

        // ...and the one thing out here that DOES move with the pointer. Every move, because that is
        // the whole of what it is for - and it is one SetWindowPos, which is why the paragraph above
        // can go on refusing to repaint the row.
        if (home >= 0)
            GhostFollow(state, hwnd, &layout.tab[home], &layout.close[home]);

        return;
    }

    // Where the tab is now: carried from the point inside it that was grabbed, so it does not jump
    // under the pointer when it is picked up, and never past either end of the *track* - a tab
    // carried out over the scroll buttons is one being aimed blind, and in an overflowing row the
    // end of the track is not the end of the row.
    int left = point.x - g_dragGrabDx;
    if (left > layout.track.right - layout.width)
        left = layout.track.right - layout.width;
    if (left < layout.track.left)
        left = layout.track.left;
    g_dragLeft = left;

    // Which slot it belongs in: the one containing the carried tab's own centre. Measured from the
    // tab rather than from the pointer, so the row swaps when the tab is visibly half way past its
    // neighbour - the pointer can be anywhere along it, and swapping on the pointer makes a tab
    // grabbed by its right-hand edge jump a place the instant it is picked up.
    //
    // Worked in the row's own coordinates rather than by scanning the rectangles on screen, because
    // in a scrolled row the slots either side of the visible ones are real places a tab can be put
    // and there is no rectangle to find them by.
    int centre  = left + layout.width / 2;
    int virtual_ = centre - layout.track.left + layout.scroll;
    int target  = (layout.width > 0) ? (virtual_ / layout.width) : 0;
    if (target < 0)                 target = 0;
    if (target > layout.count - 1)  target = layout.count - 1;

    StackMoveTab(g_dragFrame, target);   // free, and silent, while the tab is already there
    StripRefreshTabs();
}

// ---------------------------------------------------------------------------------------------
// Carrying a tab past the end of a scrolled row.
//
// Without this, a drag in an overflowing row can only rearrange the tabs that happen to be on
// screen: the pointer reaches the edge of the track and there is nowhere further to go. With it, the
// row moves under the carried tab while the hand holds still at the end - the same thing a file
// manager does when a drag reaches the edge of a list.
//
// Driven by a timer rather than by mouse movement, because the gesture that needs it is the one
// where the pointer has *stopped*. The timer runs for the length of the drag and each tick decides
// whether anything is needed, which is one place to start it and one place to stop it rather than
// four of each.
// ---------------------------------------------------------------------------------------------

static void DragScrollStop(void)
{
    if (g_dragScroll)
    {
        if (IsWindow(g_dragScroll))
            KillTimer(g_dragScroll, ID_DRAG_SCROLL);
        g_dragScroll = NULL;
    }
}

static void DragScrollStart(HWND hwnd)
{
    if (g_dragScroll == hwnd)
        return;
    DragScrollStop();
    if (SetTimer(hwnd, ID_DRAG_SCROLL, DRAG_SCROLL_MS, NULL))
        g_dragScroll = hwnd;
}

static void DragScrollTick(StripState* state, HWND hwnd)
{
    // g_dragTorn among the reasons to stop: a tab carried out of the row is not being aimed at a slot
    // in it, so scrolling the row under it would be moving something the gesture is no longer about.
    if (!g_dragging || g_dragTorn || g_dragStrip != hwnd || !g_dragFrame)
    {
        DragScrollStop();
        return;
    }

    RECT client;
    POINT cursor;
    if (!GetClientRect(hwnd, &client) || !GetCursorPos(&cursor) || !ScreenToClient(hwnd, &cursor))
        return;

    HWND frames[MAX_STRIPS];
    StripLayout layout;
    LayoutOf(state, &client, frames, &layout);
    if (layout.maxScroll <= 0)
        return;

    int edge = Scaled(DRAG_SCROLL_EDGE, state->dpi);
    int step = Scaled(DRAG_SCROLL_LOGICAL, state->dpi);
    int by   = 0;

    if (cursor.x <= layout.track.left + edge)
        by = -step;
    else if (cursor.x >= layout.track.right - edge)
        by = step;

    if (by == 0 || !ScrollBy(state, hwnd, by))
        return;

    // The row has moved, so where the carried tab belongs has moved with it. Re-run the same code
    // the pointer would have run had it twitched, rather than a second copy of it that could
    // disagree about which slot the tab is over.
    DragMove(state, hwnd, cursor);
}

// A scroll chevron held down.
//
// Each tick re-asks two questions rather than trusting what it was armed with. Is the pointer still on
// that button - which is what makes sliding off a held chevron pause the repeat and sliding back
// resume it, the same thing a scrollbar arrow does and the same rule the release already follows. And
// is there anywhere left to scroll, which stops the timer running at ninety milliseconds against a row
// that has reached its end.
static void ChevronRepeatTick(StripState* state, HWND hwnd)
{
    int kind = state->pressKind;
    if ((kind != HIT_PREV && kind != HIT_NEXT) || GetCapture() != hwnd)
    {
        KillTimer(hwnd, ID_CHEVRON_REPEAT);
        return;
    }

    POINT cursor;
    if (!GetCursorPos(&cursor) || !ScreenToClient(hwnd, &cursor))
        return;

    // Slid off: pause, and deliberately keep the timer so that sliding back on resumes it.
    StripHit hit = HitTestStrip(state, hwnd, cursor);
    if (hit.kind != kind)
        return;

    RECT client;
    HWND frames[MAX_STRIPS];
    StripLayout layout;
    if (!GetClientRect(hwnd, &client))
        return;
    LayoutOf(state, &client, frames, &layout);

    if (!ScrollBy(state, hwnd, (kind == HIT_NEXT) ? layout.width : -layout.width))
    {
        // Nothing moved, so the row is against that end. Stop rather than spin - the button is drawn
        // faint and is no longer a target, and the release will find nothing to do either.
        KillTimer(hwnd, ID_CHEVRON_REPEAT);
        return;
    }

    if (!state->pressRepeated)
    {
        state->pressRepeated = TRUE;
        LogWrite(L"strip  hwnd=0x%p  %s chevron held down - scrolling while it is held",
                 (void*)state->frame, (kind == HIT_NEXT) ? L"right" : L"left");
    }

    // The hold delay becoming a repeat rate, rather than a second timer: SetTimer with an id that
    // already exists replaces that timer's interval.
    SetTimer(hwnd, ID_CHEVRON_REPEAT, CHEVRON_REPEAT_MS, NULL);
}

// ---------------------------------------------------------------------------------------------
// The tooltip.
//
// Rest the pointer on a tab and, half a second later, a small panel below it says what the tab could
// not fit: the document's name in full, and underneath it the folder the document is in.
//
// **Why it exists is a measurement, not a nicety.** The icons slice ended with a finding that has
// been true of every slice since: the name is the only thing on a tab that identifies its document.
// And the name is the first thing the row spends when it runs short of room - Word's own default
// window fits five tabs, so a sixth document shortens every name in the row, permanently. Past that
// point a strip of `Qua...`, `Me...` and `Draf...` is a row of tabs that cannot be told apart, and
// there is nowhere else in the window to look.
//
// **The folder is the second line because it is the thing nothing else can tell you.** Two documents
// called `Report.docx` from two different folders draw two identical tabs, and no amount of width
// would separate them. It is read from Word on the hover that needs it rather than polled - see
// WordTabReadDocumentPath, which says why this one is not the dot.
//
// **A tooltip with nothing to add does not appear.** If the name fitted and there is no folder to
// name - an unsaved `Document1` on a wide tab - then everything the panel would say is already on
// screen, and showing it anyway would be half a second of the user's attention spent on a repetition.
// The two halves of that test come from different places and both are honest: "the name was cut" is
// recorded by the code that cut it, and "there is no folder" is Word's own answer.
//
// **It is a window of our own rather than a system tooltip**, for the reason every other surface here
// is: a system tooltip is painted in the *system's* colours, and Word's theme is not Windows' theme -
// this rig runs a dark Word on a light Windows. The panel uses the same three floating-surface
// colours the context menu does, so a tooltip and a menu are the same object seen twice rather than
// two guesses about what a panel over a document should look like.
// ---------------------------------------------------------------------------------------------

static ATOM g_tipClass = 0;

// The one tooltip that can be up, and what it says. Global rather than per strip for the same reason
// the drag and the scroll position are: there is one pointer, so there is one hover, and a per-strip
// copy of the text would be several answers to a question that has one. The *window* is per strip,
// because a popup is owned by a frame and an owner cannot be changed after it is created.
static HWND    g_tipStrip  = NULL;    // the strip whose hover has armed or shown it
static HWND    g_tipFrame  = NULL;    // ...and the tab it is about
static wchar_t g_tipName[256]      = L"";
static wchar_t g_tipFolder[MAX_PATH] = L"";

// Take down anything showing and disarm anything pending. Safe to call when neither is true, which
// is most of the times it is called.
static void TipStop(void)
{
    if (g_tipStrip && IsWindow(g_tipStrip))
    {
        KillTimer(g_tipStrip, ID_TIP_SHOW);
        KillTimer(g_tipStrip, ID_TIP_HIDE);

        StripState* state = FindByStrip(g_tipStrip);
        if (state && state->tip && IsWindow(state->tip) && IsWindowVisible(state->tip))
        {
            ShowWindow(state->tip, SW_HIDE);

            // The frame the tooltip was ABOUT, not the frame whose strip is carrying the panel.
            // Those are routinely different - every window in the stack draws every tab, so the
            // panel for tab 3 belongs to whichever window happens to be in front - and the first
            // version of this line logged the carrier, which made a hide look like it belonged to a
            // document nobody had hovered.
            LogWrite(L"tip: hwnd=0x%p  hidden", (void*)g_tipFrame);
        }
    }

    g_tipStrip = NULL;
    g_tipFrame = NULL;
    g_tipName[0]   = L'\0';
    g_tipFolder[0] = L'\0';
}

// What a change to the row wants, which is not the same thing as TipStop.
//
// A panel that is UP describes a row that has just changed underneath it, so it goes. A timer that
// is merely PENDING describes nothing yet: TipShow re-hit-tests the pointer, re-reads the tab's name
// and re-asks Word for the folder when it fires, so there is nothing about a pending arm that a row
// change can make stale.
//
// Killing it anyway was a real defect and not a tidy one. A pointer resting on a tab sends no
// further WM_MOUSEMOVE, so nothing re-arms: any row change landing inside that half second - a dot
// coming on in another document, a title Word rewrote, a window activating - meant the tooltip never
// appeared at all, for as long as the hand stayed still. Silently, because a tooltip that was never
// shown has nothing to log about being hidden.
static void TipRowChanged(void)
{
    StripState* state = g_tipStrip ? FindByStrip(g_tipStrip) : NULL;
    if (state && state->tip && IsWindow(state->tip) && IsWindowVisible(state->tip))
        TipStop();
}

// Start the clock on a tab the pointer has arrived at.
//
// The early return on "already this tab" is what makes the delay mean what it says. WM_MOUSEMOVE
// arrives on every pixel, and re-arming on each one would give a tooltip that never appears unless
// the hand is perfectly still for half a second - which is not the same gesture at all, and is
// notably hard to perform on a trackpad.
static void TipArm(HWND strip, HWND frame)
{
    if (!g_tipEnabled || !frame || !strip)
        return;
    if (g_tipStrip == strip && g_tipFrame == frame)
        return;

    TipStop();
    g_tipStrip = strip;
    g_tipFrame = frame;
    SetTimer(strip, ID_TIP_SHOW, GetDoubleClickTime(), NULL);
}

static void DrawTip(StripState* state, HDC dc, const RECT* client)
{
    HBRUSH back = CreateSolidBrush(g_palette.menuBack);
    if (back)
    {
        FillRect(dc, client, back);
        DeleteObject(back);
    }

    // A border, because this panel floats over a document rather than over the strip's own well, and
    // white text on a white page with no edge is not a panel.
    HBRUSH edge = CreateSolidBrush(g_palette.menuLine);
    if (edge)
    {
        FrameRect(dc, client, edge);
        DeleteObject(edge);
    }

    HGDIOBJ oldFont = SelectObject(dc, state->font ? (HGDIOBJ)state->font
                                                   : GetStockObject(DEFAULT_GUI_FONT));
    SetBkMode(dc, TRANSPARENT);

    int padx = Scaled(TIP_LOGICAL_PADX, state->dpi);
    int pady = Scaled(TIP_LOGICAL_PADY, state->dpi);
    int gap  = Scaled(TIP_LOGICAL_GAP, state->dpi);

    RECT line;
    line.left   = client->left + padx;
    line.right  = client->right - padx;
    line.top    = client->top + pady;
    line.bottom = client->bottom - pady;

    RECT measured = line;
    DrawTextW(dc, g_tipName, -1, &measured, DT_SINGLELINE | DT_LEFT | DT_NOPREFIX | DT_CALCRECT);
    int lineH = measured.bottom - measured.top;

    line.bottom = line.top + lineH;
    SetTextColor(dc, g_palette.menuText);
    DrawTextW(dc, g_tipName, -1, &line, DT_SINGLELINE | DT_LEFT | DT_NOPREFIX | DT_END_ELLIPSIS);

    if (g_tipFolder[0])
    {
        line.top    = line.bottom + gap;
        line.bottom = client->bottom - pady;

        // DT_PATH_ELLIPSIS, not DT_END_ELLIPSIS: a path cut at the end loses the folder the document
        // is actually in, which is the only part of it anybody reads. Cut in the middle it keeps both
        // ends, which is what every file dialog in Windows does with a path too long for its box.
        SetTextColor(dc, g_palette.menuTextDim);
        DrawTextW(dc, g_tipFolder, -1, &line,
                  DT_SINGLELINE | DT_LEFT | DT_NOPREFIX | DT_PATH_ELLIPSIS);
    }

    SelectObject(dc, oldFont);
}

static LRESULT CALLBACK TipWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    StripState* state = (StripState*)GetWindowLongPtrW(hwnd, GWLP_USERDATA);

    switch (msg)
    {
    case WM_ERASEBKGND:
        return 1;

    case WM_PAINT:
        if (state)
        {
            PAINTSTRUCT ps;
            HDC dc = BeginPaint(hwnd, &ps);
            if (dc)
            {
                RECT client;
                GetClientRect(hwnd, &client);
                DrawTip(state, dc, &client);
            }
            EndPaint(hwnd, &ps);
            return 0;
        }
        break;

    // Photographable, for the same reason the strip is: a PrintWindow of a panel DefWindowProc
    // cannot draw comes back blank, and a screenshot that silently omits the thing under test is
    // worse than no screenshot.
    case WM_PRINTCLIENT:
        if (state && wParam)
        {
            RECT client;
            GetClientRect(hwnd, &client);
            DrawTip(state, (HDC)wParam, &client);
            return 0;
        }
        break;

    // It never takes the focus and it never takes a click. WS_EX_TRANSPARENT means the pointer
    // reaches the document underneath, so moving onto the panel is moving off the strip - which is
    // what takes the panel away.
    case WM_MOUSEACTIVATE:
        return MA_NOACTIVATE;

    case WM_NCHITTEST:
        return HTTRANSPARENT;
    }

    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

static BOOL TipEnsure(StripState* state)
{
    if (state->tip && IsWindow(state->tip))
        return TRUE;

    // Owned by the frame, not a child of it. An owned popup is always above its owner without being
    // topmost over everything else, it goes away when the frame is minimised, and - the part that
    // matters most here - an owned WS_EX_TOOLWINDOW popup gets neither a taskbar button nor an
    // Alt+Tab entry. This project spent a slice making Word one taskbar button; a tooltip is not
    // allowed to add a second.
    state->tip = CreateWindowExW(WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE | WS_EX_TRANSPARENT,
                                 kTipClass, L"", WS_POPUP,
                                 0, 0, 10, 10, state->frame, NULL, g_module, NULL);
    if (!state->tip)
    {
        LogWrite(L"tip: hwnd=0x%p  CreateWindowEx FAILED (lastError=%lu)",
                 (void*)state->frame, GetLastError());
        return FALSE;
    }

    SetWindowLongPtrW(state->tip, GWLP_USERDATA, (LONG_PTR)state);
    return TRUE;
}

static void TipShow(StripState* state, HWND hwnd)
{
    KillTimer(hwnd, ID_TIP_SHOW);

    // Nothing is armed, so there is nothing to show.
    //
    // Not defensive tidiness - this fired twice in one day on the work rig, both times as a frame
    // was being destroyed under the pointer, and it produced a line naming the null window:
    //
    //     tip: hwnd=0x0000000000000000  ||  nothing to add (name fits; Word would not say where it is)
    //
    // The test below is `hit.frame != g_tipFrame`, and when the tab has gone both sides are NULL.
    // Two nothings comparing equal read as "the pointer is still on the tab we armed", so the whole
    // body ran on a window that does not exist - including a call into Word's object model asking
    // for the document path of NULL, which is the part that is worth not doing.
    if (!g_tipFrame)
    {
        LogWrite(L"tip: the wait ended with no tab armed - nothing to show");
        return;
    }

    // Where the pointer is *now*, not where it was when the timer was armed. Half a second is long
    // enough for the row to have scrolled under a still hand, for that document to have been closed,
    // or for the pointer to have been moved by something that sends no mouse message at all.
    //
    // Every exit below says so in the log, and that is the point of them rather than a courtesy. An
    // intermittent red `title` - one tab of seven producing no panel, once in a hundred - could not
    // be told apart from a hover that never happened at all, because a tooltip that is abandoned
    // before it is shown had nothing to say. "No tooltip and no log line" meant three different
    // things and named none of them.
    POINT cursor;
    POINT client;
    if (!GetCursorPos(&cursor))
    {
        LogWrite(L"tip: hwnd=0x%p  abandoned - Windows would not say where the pointer is",
                 (void*)g_tipFrame);
        TipStop();
        return;
    }
    client = cursor;
    if (!ScreenToClient(hwnd, &client))
    {
        LogWrite(L"tip: hwnd=0x%p  abandoned - the pointer could not be put in strip coordinates",
                 (void*)g_tipFrame);
        TipStop();
        return;
    }

    StripHit hit = HitTestStrip(state, hwnd, client);
    if (hit.frame != g_tipFrame)
    {
        // Half a second is long enough for the row to have scrolled under a still hand, for the tab
        // to have gone, or for the pointer to have been moved by something that sends no mouse
        // message at all. Dropping it silently was wrong twice: nothing recorded it, and nothing
        // brought it back either - a still pointer sends no WM_MOUSEMOVE, so the wait that would
        // have started the clock again never happens.
        //
        // Starting again on whatever tab IS under the pointer terminates: the next fire hit-tests
        // the same point against the same layout and matches. It only repeats while the row is
        // being rearranged twice a second, which the log is already full of when it is true.
        HWND armed = g_tipFrame;
        TipStop();
        if (hit.frame)
        {
            LogWrite(L"tip: hwnd=0x%p  the pointer is on 0x%p now - starting the wait again there",
                     (void*)armed, (void*)hit.frame);
            TipArm(hwnd, hit.frame);
        }
        else
        {
            LogWrite(L"tip: hwnd=0x%p  abandoned - the pointer is no longer on a tab", (void*)armed);
        }
        return;
    }

    wchar_t name[256];
    StripTabName(hit.frame, name, 256);

    wchar_t folder[MAX_PATH];
    BOOL answered = WordTabReadDocumentPath(hit.frame, folder, MAX_PATH);
    if (!answered)
        folder[0] = L'\0';

    BOOL cut = (hit.index >= 0 && hit.index < MAX_TABS) ? state->nameCut[hit.index] : FALSE;

    if (!cut && folder[0] == L'\0')
    {
        // Disarmed but not forgotten: g_tipFrame stays this tab, so the pointer wandering about on
        // it does not re-arm the timer and ask Word the same question every half second.
        LogWrite(L"tip: hwnd=0x%p  |%s|  nothing to add (name fits; %s)",
                 (void*)hit.frame, name,
                 answered ? L"no folder - never saved" : L"Word would not say where it is");
        return;
    }

    // Re-asserted, and this is not belt and braces.
    //
    // WordTabReadDocumentPath calls into Word, and a call into Word's object model pumps messages -
    // the rule this add-in already follows everywhere it posts a command to itself rather than
    // running one inside a window procedure. A janitor tick dispatched inside that pump reaches
    // StripRefreshTabs, which calls TipStop, which clears exactly the two variables that say a
    // tooltip is up. Without these two lines the panel below would be shown with nothing recording
    // that it exists, and the next TipStop would find nothing to hide: a tooltip that stays on the
    // screen until the pointer happens to arm another one.
    g_tipStrip = hwnd;
    g_tipFrame = hit.frame;

    wcsncpy(g_tipName, name, 255);
    g_tipName[255] = L'\0';
    wcsncpy(g_tipFolder, folder, MAX_PATH - 1);
    g_tipFolder[MAX_PATH - 1] = L'\0';

    if (!TipEnsure(state))
    {
        LogWrite(L"tip: hwnd=0x%p  abandoned - the panel window could not be created",
                 (void*)hit.frame);
        TipStop();
        return;
    }

    // The panel's size, measured with the strip's own font so the two agree about what a name is.
    int width  = 0;
    int height = 0;
    int lineH  = 0;
    {
        HDC dc = GetDC(hwnd);
        if (!dc)
        {
            LogWrite(L"tip: hwnd=0x%p  abandoned - no device context to measure the text with",
                     (void*)hit.frame);
            TipStop();
            return;
        }
        HGDIOBJ oldFont = SelectObject(dc, state->font ? (HGDIOBJ)state->font
                                                       : GetStockObject(DEFAULT_GUI_FONT));

        RECT m;
        SetRect(&m, 0, 0, 0, 0);
        DrawTextW(dc, g_tipName, -1, &m, DT_SINGLELINE | DT_LEFT | DT_NOPREFIX | DT_CALCRECT);
        int w1 = m.right - m.left;
        lineH  = m.bottom - m.top;

        int w2 = 0;
        if (g_tipFolder[0])
        {
            SetRect(&m, 0, 0, 0, 0);
            DrawTextW(dc, g_tipFolder, -1, &m,
                      DT_SINGLELINE | DT_LEFT | DT_NOPREFIX | DT_CALCRECT);
            w2 = m.right - m.left;
        }

        SelectObject(dc, oldFont);
        ReleaseDC(hwnd, dc);

        // DT_CALCRECT on a single line answers what the string wants and ignores the box it was
        // given, so the cap is applied here rather than trusted to the measurement.
        int wide = (w1 > w2) ? w1 : w2;
        int cap  = Scaled(TIP_LOGICAL_MAXW, state->dpi);
        if (wide > cap)
            wide = cap;

        width  = wide + Scaled(TIP_LOGICAL_PADX, state->dpi) * 2;
        height = lineH + Scaled(TIP_LOGICAL_PADY, state->dpi) * 2;
        if (g_tipFolder[0])
            height += lineH + Scaled(TIP_LOGICAL_GAP, state->dpi);
    }

    // Under the tab it describes rather than under the pointer. A tooltip that points at the thing
    // it is about needs no arrow, and the tab is a fixed target where the pointer is not.
    RECT anchor = hit.tab;
    MapWindowPoints(hwnd, NULL, (POINT*)&anchor, 2);

    int drop = Scaled(TIP_LOGICAL_DROP, state->dpi);
    int x = anchor.left;
    int y = anchor.bottom + drop;

    // Clamped to the monitor the pointer is on, not to the frame: a Word window can hang off the
    // right-hand edge of a screen, and a panel clamped to the window would then be drawn off it.
    MONITORINFO mi;
    mi.cbSize = sizeof(mi);
    if (GetMonitorInfoW(MonitorFromPoint(cursor, MONITOR_DEFAULTTONEAREST), &mi))
    {
        if (x + width > mi.rcWork.right)
            x = mi.rcWork.right - width;
        if (x < mi.rcWork.left)
            x = mi.rcWork.left;
        if (y + height > mi.rcWork.bottom)
            y = anchor.top - height - drop;      // above the tab instead of below it
        if (y < mi.rcWork.top)
            y = mi.rcWork.top;
    }

    // The whole panel as one string, on the window itself, and it is the only text this add-in draws
    // that anything outside the process can read. A tab name can only be asserted through the log
    // line that says what was computed; this can be asserted by asking the window what it says.
    wchar_t both[256 + MAX_PATH + 2];
    wcscpy(both, g_tipName);
    if (g_tipFolder[0])
    {
        wcscat(both, L"\n");
        wcscat(both, g_tipFolder);
    }
    SetWindowTextW(state->tip, both);

    SetWindowPos(state->tip, HWND_TOP, x, y, width, height, SWP_NOACTIVATE | SWP_SHOWWINDOW);
    InvalidateRect(state->tip, NULL, TRUE);

    // And it takes itself away, the way every tooltip in Windows does. Ten times the appearance
    // delay is the system's own ratio.
    SetTimer(hwnd, ID_TIP_HIDE, GetDoubleClickTime() * 10, NULL);

    LogWrite(L"tip: hwnd=0x%p  |%s|  folder |%s|  at %d,%d %dx%d  (%s%s)",
             (void*)hit.frame, g_tipName, g_tipFolder, x, y, width, height,
             cut ? L"name was cut" : L"name fits",
             answered ? L"" : L"; Word would not say where it is");
}

static LRESULT CALLBACK StripWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    StripState* state = (StripState*)GetWindowLongPtrW(hwnd, GWLP_USERDATA);

    // Any button anywhere in the strip takes the tooltip down: the user has stopped reading it and
    // started doing something. Here rather than in each of the three handlers, so a fourth button
    // cannot be added without it.
    if (msg == WM_LBUTTONDOWN || msg == WM_MBUTTONDOWN || msg == WM_RBUTTONDOWN)
        TipStop();

    switch (msg)
    {
    case WM_ERASEBKGND:
        return 1;              // WM_PAINT covers every pixel; erasing first only flickers

    case WM_PAINT:
        if (state)
        {
            PaintStrip(state, hwnd);
            return 0;
        }
        break;

    // How a strip is photographed. DefWindowProc cannot draw a window's client area for it, so
    // without this a PrintWindow of a Word frame comes back with a blank band where the tabs are -
    // and a screenshot that silently omits the thing under test is worse than no screenshot.
    case WM_PRINTCLIENT:
        if (state && wParam)
        {
            RECT client;
            GetClientRect(hwnd, &client);
            DrawStrip(state, (HDC)wParam, &client);
            return 0;
        }
        break;

    case WM_MOUSEMOVE:
        if (state)
        {
            // A held tab takes the whole message. Hover means nothing while the button is down - the
            // pointer is on the tab it is carrying - and the carried tab is its own feedback.
            if (g_dragStrip == hwnd)
            {
                if (g_dragFrame)
                    DragMove(state, hwnd, PointOf(lParam));
                return 0;
            }

            ArmLeaveTracking(state, hwnd);
            StripHit hit = HitTestStrip(state, hwnd, PointOf(lParam));
            SetHot(state, hwnd, hit.kind, hit.frame);

            // A tab under the pointer starts the clock; anything else - the +, a chevron, the bare
            // well - stops it. Keyed on the frame rather than on the hit kind, so sliding from a
            // tab onto its own close button is still the same tab and does not restart the wait.
            if (hit.frame)
                TipArm(hwnd, hit.frame);
            else
                TipStop();
            return 0;
        }
        break;

    case WM_MOUSELEAVE:
        if (state)
        {
            state->tracking = FALSE;
            SetHot(state, hwnd, HIT_NONE, NULL);
            TipStop();
            return 0;
        }
        break;

    // The wheel over the tab row scrolls it, which is how every other overflowing row of tabs
    // behaves. A vertical wheel doing a horizontal thing is the convention rather than a liberty -
    // a strip 32 pixels tall has nothing to scroll vertically, and the alternative is a row most
    // mice cannot move at all.
    //
    // Whether this message arrives here is Windows' decision, not ours: WM_MOUSEWHEEL goes to the
    // focused window, and the strip is WS_EX_NOACTIVATE and never has focus. What delivers it is
    // "scroll inactive windows when I hover over them", which is on by default and routes the wheel
    // to the window under the pointer. The scroll buttons exist partly because that is a setting and
    // settings can be off.
    case WM_MOUSEWHEEL:
    case WM_MOUSEHWHEEL:
        if (state)
        {
            int notches = GET_WHEEL_DELTA_WPARAM(wParam) / WHEEL_DELTA;
            if (notches == 0)
                break;

            RECT client;
            HWND frames[MAX_STRIPS];
            StripLayout layout;
            if (!GetClientRect(hwnd, &client))
                break;
            LayoutOf(state, &client, frames, &layout);
            if (layout.maxScroll <= 0)
                break;

            // A wheel forward is up, and up is left. A *horizontal* wheel is the other way round:
            // its positive direction is already right.
            int by = layout.width * notches;
            if (msg == WM_MOUSEWHEEL)
                by = -by;

            ScrollBy(state, hwnd, by);
            return 0;
        }
        break;

    case WM_TIMER:
        if (state && wParam == ID_DRAG_SCROLL)
        {
            DragScrollTick(state, hwnd);
            return 0;
        }
        if (state && wParam == ID_CHEVRON_REPEAT)
        {
            ChevronRepeatTick(state, hwnd);
            return 0;
        }
        if (state && wParam == ID_TIP_SHOW)
        {
            TipShow(state, hwnd);
            return 0;
        }
        if (state && wParam == ID_TIP_HIDE)
        {
            TipStop();
            return 0;
        }
        break;

    case WM_LBUTTONDOWN:
        if (state)
        {
            StripHit hit = HitTestStrip(state, hwnd, PointOf(lParam));

            if (hit.kind == HIT_CLOSE || hit.kind == HIT_PLUS ||
                hit.kind == HIT_PREV  || hit.kind == HIT_NEXT)
            {
                // Press and release on the same button, the way every other button in Windows
                // behaves: pressing one and sliding off cancels it. Acting on the press would mean
                // a mis-aimed click closes a document with no way to change your mind.
                state->pressKind     = hit.kind;
                state->pressFrame    = hit.frame;
                state->pressRepeated = FALSE;
                SetCapture(hwnd);
                InvalidateRect(hwnd, NULL, FALSE);

                // A held chevron keeps scrolling. Armed here rather than on the first repeat so
                // there is one place that starts it and one that stops it, and only for the two
                // buttons it means anything for - holding the close button down is not a request to
                // close several documents.
                if (hit.kind == HIT_PREV || hit.kind == HIT_NEXT)
                    SetTimer(hwnd, ID_CHEVRON_REPEAT, CHEVRON_HOLD_MS, NULL);
            }
            else if (hit.kind == HIT_TAB)
            {
                // Switching, though, happens on the press. It is instant, it is reversible by
                // clicking the tab you came from, and waiting for the release makes it feel slow.
                LogWrite(L"strip  hwnd=0x%p  tab clicked -> 0x%p",
                         (void*)state->frame, (void*)hit.frame);

                // Claimed as already revealed *before* it is activated. A tab at the end of a
                // scrolled row can be half off the track, and the reveal that follows an activation
                // would slide the row under a pointer that is still holding the button down - the
                // tab would move out from under the hand between the press and the drag, and the
                // carried tab would sit a scroll's width from the pointer for the rest of it. The
                // user can see the tab: they just clicked it.
                g_scrollShown = hit.frame;
                StackActivate(hit.frame);

                // ...and the same press may turn out to be a drag. Nothing is committed here: the
                // tab is only claimed, and it stays a plain click until the pointer travels far
                // enough to mean something else.
                //
                // The capture is taken *after* the activation, not before. Activating raises a
                // different window, and the one thing that has to be true when this returns is that
                // this strip owns the mouse. A tab that is not in a stack has no row to be reordered
                // within, so it is not picked up at all.
                //
                // Two gestures start the same way and the difference is decided once, here, from
                // whether this tab is in a row at all. A tab in the row can be reordered along it or
                // pulled out of it; a window standing on its own has one tab and nothing to reorder
                // it against, so the only thing carrying it can mean is putting it back.
                if (g_dragEnabled && StackTabIndex(hit.frame) >= 0)
                {
                    POINT point  = PointOf(lParam);
                    g_dragStrip  = hwnd;
                    g_dragFrame  = hit.frame;
                    g_dragPressX = point.x;
                    g_dragPressY = point.y;
                    g_dragGrabDx = point.x - hit.tab.left;
                    g_dragLeft   = hit.tab.left;
                    g_dragFrom   = hit.index;
                    g_dragging   = FALSE;
                    g_dragJoin   = FALSE;
                    SetCapture(hwnd);
                }
                else if (g_dragEnabled && StackCanRejoin(hit.frame))
                {
                    POINT point  = PointOf(lParam);
                    g_dragStrip  = hwnd;
                    g_dragFrame  = hit.frame;
                    g_dragPressX = point.x;
                    g_dragPressY = point.y;
                    g_dragGrabDx = point.x - hit.tab.left;
                    g_dragLeft   = hit.tab.left;
                    g_dragFrom   = hit.index;
                    g_dragging   = FALSE;
                    g_dragJoin   = TRUE;
                    g_dragOnto   = NULL;
                    g_dragOntoKnown = FALSE;
                    SetCapture(hwnd);
                }
            }
            return 0;
        }
        break;

    // Letting go of a tab. A drop back in the row has nothing to commit - the row rearranged as the
    // tab was carried - so that is bookkeeping and a log line. A drop *outside* the row is the one
    // that does something, and it is the only place in this file where releasing a button acts on
    // where the pointer is rather than on what is under it. It is also where a press that never
    // became a drag ends, which is a plain click and was already handled on the way down.
    case WM_LBUTTONUP:
        if (g_dragStrip == hwnd)
        {
            HWND frame = g_dragFrame;
            BOOL was   = g_dragging;
            BOOL torn  = g_dragTorn;
            BOOL join  = g_dragJoin;
            HWND onto  = g_dragOnto;
            int  from  = g_dragFrom;

            // Cleared *before* the capture goes back. ReleaseCapture sends this window a
            // WM_CAPTURECHANGED, and that handler's job is to put an interrupted drag back where it
            // started - which is the exact opposite of what a completed drop means.
            DragForget();
            if (GetCapture() == hwnd)
                ReleaseCapture();

            if (was && frame && join)
            {
                // Dropped on somebody's row, or dropped on nothing. Nothing is a perfectly good
                // outcome and is left silent apart from the log: the pointer said IDC_NO for the
                // whole approach, so a release there is a user who changed their mind, not a
                // gesture that failed.
                if (onto)
                {
                    LogWrite(L"strip  hwnd=0x%p  rejoin dropped on 0x%p's row - putting 0x%p back",
                             (void*)hwnd, (void*)onto, (void*)frame);
                    PostMessageW(hwnd, WM_WORDTAB_CMD, (WPARAM)CMD_JOIN, (LPARAM)frame);
                }
                else
                {
                    LogWrite(L"strip  hwnd=0x%p  rejoin let go over nothing - 0x%p stays on its own",
                             (void*)hwnd, (void*)frame);
                }
                return 0;
            }

            if (was && frame && torn)
            {
                // Posted, and posted as the same command the menu item raises. StackTearOffTab moves
                // windows, changes which one is in front and refits an interior - none of which
                // belongs inside a mouse handler that is still unwinding a capture - and routing it
                // through WM_WORDTAB_CMD means there is one implementation of "a tab was torn off"
                // rather than one per way of asking for it. That path already drops the command if
                // the window has gone in the meantime.
                // The release point is logged as well as the tab, because "the add-in never saw the
                // pointer come back" and "the pointer never came back" are the same evidence from
                // outside, and this is the one line that says where the add-in thought it was.
                POINT drop = PointOf(lParam);
                LogWrite(L"strip  hwnd=0x%p  drag dropped out of the row: tearing 0x%p off (tab %d)"
                         L"  (released at y=%d)", (void*)hwnd, (void*)frame, from, (int)drop.y);
                PostMessageW(hwnd, WM_WORDTAB_CMD, (WPARAM)CMD_TEAROFF, (LPARAM)frame);
            }
            else if (was && frame)
            {
                LogWrite(L"strip  hwnd=0x%p  drag ended: 0x%p is tab %d (was %d)",
                         (void*)hwnd, (void*)frame, StackTabIndex(frame), from);
                StripRefreshTabs();
            }
            return 0;
        }
        if (state && state->pressKind != HIT_NONE)
        {
            int  kind     = state->pressKind;
            HWND frame    = state->pressFrame;
            BOOL repeated = state->pressRepeated;

            state->pressKind     = HIT_NONE;
            state->pressFrame    = NULL;
            state->pressRepeated = FALSE;
            KillTimer(hwnd, ID_CHEVRON_REPEAT);
            if (GetCapture() == hwnd)
                ReleaseCapture();

            // A hold that already scrolled has done the work. Letting the release scroll once more
            // would make every held chevron overshoot by exactly one tab, which is the kind of error
            // that reads as the row being imprecise rather than as an extra step.
            if (repeated)
            {
                LogWrite(L"strip  hwnd=0x%p  chevron released after a hold - no extra step",
                         (void*)state->frame);
                InvalidateRect(hwnd, NULL, FALSE);
                return 0;
            }

            StripHit hit = HitTestStrip(state, hwnd, PointOf(lParam));
            if (hit.kind == kind && hit.frame == frame)
            {
                if (kind == HIT_CLOSE)
                {
                    LogWrite(L"strip  hwnd=0x%p  close clicked on 0x%p",
                             (void*)state->frame, (void*)frame);
                    StackCloseTab(frame);
                }
                else if (kind == HIT_PREV || kind == HIT_NEXT)
                {
                    // One tab per click. A page per click gets to the far end of a long row sooner
                    // and overshoots the tab you were looking for every time, and the wheel is
                    // already the way to cross a row quickly.
                    RECT client;
                    HWND frames[MAX_STRIPS];
                    StripLayout layout;
                    int step = Scaled(TAB_LOGICAL_MIN_W, state->dpi);
                    if (GetClientRect(hwnd, &client))
                    {
                        LayoutOf(state, &client, frames, &layout);
                        step = layout.width;
                    }
                    ScrollBy(state, hwnd, (kind == HIT_NEXT) ? step : -step);
                }
                else
                {
                    LogWrite(L"strip  hwnd=0x%p  new-document button clicked",
                             (void*)state->frame);
                    PostMessageW(hwnd, WM_WORDTAB_CMD, CMD_NEW, 0);
                }
            }
            else
            {
                LogWrite(L"strip  hwnd=0x%p  button released off target - cancelled",
                         (void*)state->frame);
            }

            InvalidateRect(hwnd, NULL, FALSE);
            return 0;
        }
        break;

    // Middle-click closes, which is what a middle click does to a tab everywhere else. Paired the
    // same way as the close button: the release has to land on the tab the press did.
    case WM_MBUTTONDOWN:
        if (g_dragStrip == hwnd)
            return 0;              // a tab is being held; a second button is not a second gesture
        if (state)
        {
            StripHit hit = HitTestStrip(state, hwnd, PointOf(lParam));
            if (hit.kind == HIT_TAB || hit.kind == HIT_CLOSE)
            {
                state->middleFrame = hit.frame;
                SetCapture(hwnd);
            }
            return 0;
        }
        break;

    case WM_MBUTTONUP:
        if (state && state->middleFrame)
        {
            HWND frame = state->middleFrame;
            state->middleFrame = NULL;
            if (GetCapture() == hwnd)
                ReleaseCapture();

            StripHit hit = HitTestStrip(state, hwnd, PointOf(lParam));
            if ((hit.kind == HIT_TAB || hit.kind == HIT_CLOSE) && hit.frame == frame)
            {
                LogWrite(L"strip  hwnd=0x%p  middle-clicked 0x%p", (void*)state->frame, (void*)frame);
                StackCloseTab(frame);
            }
            return 0;
        }
        break;

    // A right press claims a target the same way the other two buttons do, and the menu appears on
    // the release. Consistent with them on purpose: a press that lands on the wrong tab is still
    // taken back by sliding off it before letting go.
    case WM_RBUTTONDOWN:
        // The right button while a tab is held is the cancel, not a menu. The capture is kept: the
        // left button is still down, and the gesture is not over until it is let go of.
        if (g_dragStrip == hwnd)
        {
            DragUndo(hwnd, L"right button");
            return 0;
        }
        if (state && g_menuEnabled)
        {
            state->rightDown  = TRUE;
            state->rightFrame = MenuTargetAt(state, hwnd, PointOf(lParam));
            SetCapture(hwnd);
            return 0;
        }
        break;

    case WM_RBUTTONUP:
        // Swallowed while a tab is held, and not merely ignored. DefWindowProc turns a right release
        // into WM_CONTEXTMENU, and a child window's WM_CONTEXTMENU goes to its parent - so letting
        // this one through would end a cancelled drag by opening Word's own context menu.
        if (g_dragStrip == hwnd)
            return 0;
        if (state && state->rightDown)
        {
            POINT point  = PointOf(lParam);
            HWND  target = state->rightFrame;

            // Capture goes back *before* the menu opens: TrackPopupMenu takes capture itself, and
            // two owners of the mouse is one too many.
            state->rightDown  = FALSE;
            state->rightFrame = NULL;
            if (GetCapture() == hwnd)
                ReleaseCapture();

            if (MenuTargetAt(state, hwnd, point) == target)
                ShowTabMenu(hwnd, point, target);
            else
                LogWrite(L"strip  hwnd=0x%p  right button released off target - no menu",
                         (void*)state->frame);
            return 0;
        }
        break;

    // Capture can be taken away without a button release - a dialog appearing, Alt+Tab, Word
    // starting a modal loop. Anything half-pressed at that point is cancelled, not completed.
    case WM_CAPTURECHANGED:
        // Something took the mouse away mid-gesture: a dialog, Alt+Tab, Word starting a modal loop.
        // A drag that was interrupted was never agreed to, so the row goes back exactly as it was.
        // A drop clears the drag before releasing the capture, so it does not arrive here.
        if (g_dragStrip == hwnd)
        {
            DragUndo(hwnd, L"the mouse capture was taken away");
            DragForget();
        }
        if (state && (state->pressKind != HIT_NONE || state->middleFrame || state->rightDown))
        {
            state->pressKind     = HIT_NONE;
            state->pressFrame    = NULL;
            state->pressRepeated = FALSE;
            state->middleFrame   = NULL;
            state->rightDown     = FALSE;
            state->rightFrame    = NULL;
            KillTimer(hwnd, ID_CHEVRON_REPEAT);
            InvalidateRect(hwnd, NULL, FALSE);
        }
        break;

    // Every command the strip can raise, run here rather than where it was chosen: Documents.Add
    // makes a window and pumps messages while it does, Document.Save can put a dialog up, and a
    // close destroys the window whose procedure we would still be inside. All of them come back
    // round to this procedure, so all of them wait for it to have returned.
    case WM_WORDTAB_CMD:
    {
        HWND target = (HWND)lParam;
        if (target && !IsWindow(target))
        {
            LogWrite(L"strip  hwnd=0x%p  command %d dropped - that tab has gone",
                     (void*)hwnd, (int)wParam);
            return 0;
        }

        switch ((int)wParam)
        {
        case CMD_NEW:
            if (!WordTabNewDocument())
                LogWrite(L"strip  hwnd=0x%p  new document declined by Word", (void*)hwnd);
            break;

        case CMD_SAVE:
            // Activate here, save in the next message. Word updates which window is active while it
            // processes the activation, so asking it in this one would be asking before it knows -
            // and the check in WordTabSaveDocument would then refuse a save that was perfectly
            // legitimate. The activation is not optional either way: a Save As dialog owned by a
            // window underneath another at the same rectangle cannot be seen.
            StackActivate(target);
            PostMessageW(hwnd, WM_WORDTAB_SAVE, 0, (LPARAM)target);
            break;

        case CMD_TEAROFF:      StackTearOffTab(target);   break;
        case CMD_JOIN:         StackJoinTab(target);      break;
        case CMD_CLOSE:        StackCloseTab(target);     break;
        case CMD_CLOSE_OTHERS: StackCloseOthers(target);  break;
        case CMD_CLOSE_RIGHT:  StackCloseToRight(target); break;
        case CMD_CLOSE_ALL:    StackCloseAll(target);     break;
        default: break;
        }
        return 0;
    }

    case WM_WORDTAB_SAVE:
    {
        HWND target = (HWND)lParam;
        if (target && IsWindow(target))
            WordTabSaveDocument(target);
        return 0;
    }

    default:
        break;
    }

    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

static BOOL CreateStripWindow(StripState* state)
{
    if (state->strip && IsWindow(state->strip))
        return TRUE;

    state->strip = CreateWindowExW(WS_EX_NOACTIVATE, kStripClass, L"WordTab",
                                   WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS,
                                   0, 0, 10, state->stripH,
                                   state->frame, NULL, g_module, NULL);
    if (!state->strip)
    {
        LogWrite(L"strip  hwnd=0x%p  CreateWindowEx FAILED (lastError=%lu)",
                 (void*)state->frame, GetLastError());
        return FALSE;
    }

    SetWindowLongPtrW(state->strip, GWLP_USERDATA, (LONG_PTR)state);
    state->stripPlaced = FALSE;
    return TRUE;
}

// ---------------------------------------------------------------------------------------------
// Binding a frame to its `_WwF`.
// ---------------------------------------------------------------------------------------------

// ---------------------------------------------------------------------------------------------
// WHICH document frame. Not "the" one - Word does not promise there is only ever one.
//
// FindChildOfClass answers with the first `_WwF` the enumeration meets and stops, which was fine
// while the binding was made once at startup and only ever remade when the window it named was
// *destroyed*. It is not fine for the case queued as "a View change closes tabs": if Word retires a
// document frame by HIDING it and builds a second one beside it, then
//
//   - the old handle is still a window, so TryBind never comes back and the strip stays bound to it
//   - the old frame is empty, so StripHasDocument answers "no document" for a window that has one
//   - membership is decided from that answer, so the window is dropped out of the row and its strip
//     is hidden with it - permanently, because nothing re-tests the binding
//
// which is exactly the shape the user reported: the document is fine, the tab is gone, and it never
// comes back. **Not reproduced on this rig** - see tools\probe-view.ps1, which drives every View
// command through Word's own object model and finds one `_WwF` per frame throughout - so this is a
// guard against a mechanism, not a fix for a measured event. It costs nothing when there is one
// document frame, which is every case this machine can produce.
//
// The score is the whole of the choice: something inside beats nothing inside, and among equals a
// visible frame beats a hidden one. "Something inside" outranks "visible" because Backstage hides a
// perfectly good document frame and that must not hand the binding to an empty spare.
struct WwfHunt
{
    HWND best;
    int  bestScore;
};

static BOOL CALLBACK WwfHuntProc(HWND hwnd, LPARAM param)
{
    WwfHunt* hunt = (WwfHunt*)param;

    wchar_t cls[64] = L"";
    GetClassNameW(hwnd, cls, 64);
    if (_wcsicmp(cls, kWwfClass) != 0)
        return TRUE;

    int score = 0;
    if (GetWindow(hwnd, GW_CHILD) != NULL)
        score += 2;
    if (IsWindowVisible(hwnd))
        score += 1;

    if (score > hunt->bestScore)
    {
        hunt->bestScore = score;
        hunt->best      = hwnd;
    }
    return TRUE;
}

// The best document frame in this window, and whether it holds anything. NULL when the window has
// no `_WwF` at all, which is every Word window for the first moments of its life.
static HWND PickWwf(HWND frame, BOOL* holdsDocument)
{
    WwfHunt hunt;
    hunt.best      = NULL;
    hunt.bestScore = -1;

    if (frame && IsWindow(frame))
        EnumChildWindows(frame, WwfHuntProc, (LPARAM)&hunt);

    if (holdsDocument)
        *holdsDocument = (hunt.bestScore >= 2) ? TRUE : FALSE;
    return hunt.best;
}

static BOOL TitleEndsWith(const wchar_t* text, int len, const wchar_t* tail)
{
    int n = (int)wcslen(tail);
    return len >= n && wcscmp(text + len - n, tail) == 0;
}

// The name to put on a tab.
//
// Word's frame title is the document, then any state annotations, then the application. Measured on
// this build by tools\probe-titles.ps1, which prints codepoints as well as text because a separator
// that is an en dash or a non-breaking space is invisible in the second form and fatal to a match:
//
//     Quarterly report.docx                            - Word
//     Quarterly report.doc  -  Compatibility Mode      - Word
//     Locked report.docx  -  Read-Only                 - Word
//     Downloaded report.docx  -  Protected View        - Word
//     Document1                                        - Word
//
// Two measured facts do all the work here:
//
//  1. **The application suffix is single-spaced (" - Word"); the annotations are double-spaced
//     ("  -  X").** Two separators of different shapes, so the annotations can be recognised
//     without a table of English strings - which matters, because "Compatibility Mode" and
//     "Read-Only" are localised, and whatever Word adds next is in no list that could be written
//     today. All plain ASCII: no en dashes, no non-breaking spaces.
//  2. **Both sit at the end**, so both are removed from the end. The old rule cut at the *first*
//     " - Word" and a document named `Meeting - Wordsmith notes.docx` therefore drew as `Meeting`.
//     That was a real defect and it lost more of the name than the annotations ever did.
//
// **The extension is left exactly as Word gives it, and that is a measurement rather than a taste.**
// Word follows Explorer's "hide extensions for known file types": the same document reads
// `Quarterly report.docx - Word` with HideFileExt=0 and `Quarterly report - Word` with it set to 1.
// The user has already told Windows whether they want to see extensions and Word already listens, so
// stripping one here would override an answer they have given - and on a machine with them hidden
// there would be nothing to strip anyway.
//
// The bracketed form some older Word builds use ("[Read-Only]") is deliberately *not* handled: it
// could not be provoked on this build, so any code for it would be unmeasured, and a trailing
// "[...]" rule would eat the tail of a document genuinely named `Report [final].docx`.
void WordTabFrameTitle(HWND frame, wchar_t* out, int chars)
{
    if (!out || chars <= 0)
        return;
    out[0] = L'\0';
    if (!frame || !IsWindow(frame))
        return;

    GetWindowTextW(frame, out, chars);
    out[chars - 1] = L'\0';

    int len = (int)wcslen(out);

    // The application suffix, from the end. " - Microsoft Word" is what Word 2010 and earlier wrote;
    // this build writes " - Word". Longest first, or the short one matches inside the long one.
    static const wchar_t* const kApp[] = { L" - Microsoft Word", L" - Word" };
    for (int i = 0; i < 2; i++)
    {
        if (TitleEndsWith(out, len, kApp[i]))
        {
            len -= (int)wcslen(kApp[i]);
            out[len] = L'\0';
            break;
        }
    }

    // Then the state annotations, also from the end, keyed on the double-spaced separator.
    //
    // Bounded rather than while(TRUE): `a  -  b  -  c.docx` is a legal filename, and an unbounded
    // rule would leave such a document with a one-word tab. Three is more than Word has been seen to
    // append at once, and the guard is what keeps a filename that looks like an annotation from
    // being consumed to nothing.
    if (g_titleTrimEnabled)
    {
        for (int guard = 0; guard < 3; guard++)
        {
            wchar_t* at = NULL;
            for (wchar_t* p = out; *p; p++)
            {
                // Short-circuit &&, so nothing past the terminator is ever read: p[1] is only
                // examined once p[0] is a space, and so on down the run.
                if (p[0] == L' ' && p[1] == L' ' && p[2] == L'-' && p[3] == L' ' && p[4] == L' ')
                    at = p;      // the *last* one - the tail is what gets removed, not the head
            }
            if (!at || at == out)
                break;
            *at = L'\0';
        }
    }

    if (out[0] == L'\0')
        wcscpy(out, L"Word");
}

// What the tab is CALLED, which is not always what the window is called this instant.
//
// WordTabFrameTitle above is a pure function of the window title, and that is the right thing for it
// to be. This is the other question, and it needed asking separately: a frame whose document has
// just gone still has a tab in the row for as long as the stack is deciding whether to keep it, and
// during that half second its window title has already reverted to a bare "Word".
//
// Every reader goes through here - both painters, the tooltip, the row's log line, the close
// dialog's heading - and that is the point rather than tidiness. The report that queued this shows
// the row line printing |Word| for a tab that was about to disappear, and a log that disagrees
// with the screen about what a tab is called has cost this project two diagnoses.
//
// **Live except in the one case the last decided name exists for.** The first version of this
// returned `state->title` whenever there was one, which reads like the tidier rule and was wrong in
// a way the battery caught within the hour: a strip binds to a frame BEFORE Word has put a title on
// it, so the name it records at bind time is the placeholder, and serving that for the half second
// until the janitor's next tick made a document opening flash a tab called "Word" - a fresh instance
// of the defect this whole change is here to remove. check-onepage's log went from no row line
// carrying a |Word| to four of them.
//
// So the question is asked of the window first, and `state->title` is consulted only when the window
// has stopped offering a name at all. That also decides the case correctly at every stage without a
// flag to keep in step: before ReadTitle has ruled, the last real name is still in `state->title` and
// is what is drawn; once ReadTitle rules that the frame genuinely has no document, `state->title`
// becomes the placeholder itself and this falls through to the live answer, which is the same thing.
void StripTabName(HWND frame, wchar_t* out, int chars)
{
    if (!out || chars <= 0)
        return;

    wchar_t raw[256];
    raw[0] = L'\0';
    if (frame && IsWindow(frame))
    {
        GetWindowTextW(frame, raw, 256);
        raw[255] = L'\0';
    }

    // The same test ReadTitle makes, and on the same thing - the RAW title, so that a document
    // actually called Word ("Word - Word") is not mistaken for a window that has lost one.
    if (raw[0] == L'\0' || wcscmp(raw, L"Word") == 0)
    {
        StripState* state = FindByFrame(frame);
        if (state && state->title[0] != L'\0' && wcscmp(state->title, L"Word") != 0)
        {
            wcsncpy(out, state->title, (size_t)chars - 1);
            out[chars - 1] = L'\0';
            return;
        }
    }

    WordTabFrameTitle(frame, out, chars);
}

static void ReadTitle(StripState* state)
{
    wchar_t title[256];
    WordTabFrameTitle(state->frame, title, 256);

    if (wcscmp(title, state->title) == 0)
        return;

    wchar_t raw[256];
    GetWindowTextW(state->frame, raw, 256);
    raw[255] = L'\0';

    // A frame whose document has gone reverts to a bare "Word", or to no title at all, and Word
    // reuses that frame for the next document rather than destroying it. So this is not a rename -
    // it is a window on its way out of the row, or on its way to holding something else.
    //
    // Tested on the RAW title rather than the cleaned name, and that is the difference between a
    // rule and a guess: a document actually called Word gives "Word - Word" here and is not this.
    BOOL nameless = (raw[0] == L'\0' || wcscmp(raw, L"Word") == 0) ? TRUE : FALSE;

    if (nameless && state->title[0] != L'\0' && wcscmp(state->title, L"Word") != 0)
    {
        // Never believed on first sight, and after that only while the stack is not waiting.
        //
        // Two facts about the order of this file's janitor make this the shape it is. ReadTitle runs
        // in the per-strip loop ABOVE StackJanitor, so on the tick a document goes away the title
        // has already been read and drawn by the time the stack notices - measured at 11ms in the
        // report that queued this. And once the stack does have an answer, it is exactly the right
        // question to ask: while it is waiting to see whether the loss is a moment or a fact, a tab
        // that renames itself has decided it is a fact.
        //
        // So the first sight is deferred one tick, which puts the second sight after a StackJanitor
        // that has seen the same thing. Nothing is lost by waiting: a frame handed a new document
        // arrives with a real title and takes it here, and a frame that has genuinely gone leaves
        // the row rather than being renamed in it.
        //
        // A frame really sitting in the row with no document - which happens, and once had this row
        // showing a tab called "Word" for four slices - still gets its "Word". Half a second later
        // than before, and only once the stack has agreed it belongs there.
        if (!state->namelessSeen)
        {
            state->namelessSeen = TRUE;
            return;
        }
        if (StackIsWaitingFor(state->frame))
            return;
    }

    state->namelessSeen = FALSE;
    wcsncpy(state->title, title, 255);
    state->title[255] = L'\0';

    // A tab name is drawn and never stored in a control, so there is no way to read one from
    // outside the process. This line is that way. check-title.ps1 asserts the exact string
    // against it and then photographs the painted text as well, because a log line about what
    // was computed is not on its own evidence of what was drawn.
    LogWrite(L"strip  hwnd=0x%p  tab name |%s|  from window title |%s|",
             (void*)state->frame, state->title, raw);

    // Every strip shows every tab, so one window's title changing has to repaint all of them.
    StripRefreshTabs();
}

// Apply the shift for the first time. Everything after this is driven by WM_WINDOWPOSCHANGING;
// this is the one place that has to push, because Word will not spontaneously lay out a window just
// because we started caring about it.
static void ApplyInitial(StripState* state)
{
    RECT current;
    if (!ChildRect(state->frame, state->wwf, &current))
        return;

    // Idempotent, like everything else here. If the document frame is already exactly where we put
    // it then nothing has happened that needs re-deriving, and treating what we see as a fresh
    // natural rect would shift it a second time - the strip lands 32px lower and the document with
    // it. That is the accumulating-shift failure the whole design exists to avoid, and it does not
    // stop being possible just because this path is called "initial".
    if (state->hasApplied && SameRect(&current, &state->applied))
    {
        PlaceStrip(state, NULL);
        return;
    }

    RECT applied = current;
    applied.top = current.top + state->stripH;

    if ((applied.bottom - applied.top) < (2 * state->stripH))
    {
        LogWrite(L"strip  hwnd=0x%p  initial shift skipped: _WwF is only %ldpx tall",
                 (void*)state->frame, current.bottom - current.top);
        return;
    }

    // Set the state *before* moving, so the WM_WINDOWPOSCHANGING this call is about to trigger sees
    // its own proposal as already-ours and leaves it alone rather than shifting it a second time.
    state->natural    = current;
    state->applied    = applied;
    state->hasApplied = TRUE;

    SetWindowPos(state->wwf, NULL,
                 applied.left, applied.top,
                 applied.right - applied.left, applied.bottom - applied.top,
                 SWP_NOZORDER | SWP_NOACTIVATE);

    RememberClient(state);
    LogRelayout(state, L"initial");
    PlaceStrip(state, NULL);
    StackOnActiveLayout(state->frame, &state->natural);
}

// ---------------------------------------------------------------------------------------------
// What the stack needs from us.
//
// These exist because of the layout-oracle rule: Word lays out only the focused window, so the
// stack has to take that window's interior and hand it to the others itself. See stack.cpp.
// ---------------------------------------------------------------------------------------------

// TRUE while a press is being held on a tab: a gesture that has begun and has not been agreed to.
//
// Asked by the stack, which decides which tab is selected from which window is in front. That is
// the right question while the windows are sitting still and the wrong one while a tab is being
// carried: the card under the pointer is a top-level window of its own, and showing it was
// measured raising the window whose strip the press landed on - back over the tab that same press
// had just activated. The row followed it and undid the switch the user had just made, 92ms after
// the "tab clicked ->" line, in the reorder suite log. Nothing is concluded from where the windows
// have ended up until the gesture is over.
BOOL StripTabPressHeld(void)
{
    return g_dragStrip != NULL;
}

// Is there a document open in this window?
//
// The obvious test - "does the frame have a `_WwF`" - is wrong, and was wrong here for four slices.
// `_WwF` is the document *frame*, and Word keeps it for the life of the window: closing the last
// document destroys the `_WwB` and `_WwG` inside it and leaves `_WwF` behind, empty. So a Word window
// with nothing open passed that test, joined the stack, and was given a tab labelled "Word".
//
// Measured on 16.0.20228: `_WwF` holds `_WwB` -> `_WwG` in print layout, read mode, web layout,
// draft and outline, with Backstage open, and while minimised - and holds nothing whatsoever when
// there is no document. So the honest question is whether anything is *in* the document frame.
//
// `GetWindow(GW_CHILD)` rather than looking for `_WwB` by name, on two grounds: it is one call
// instead of a recursive enumeration on a half-second timer, and "the document frame is empty" is a
// weaker assumption about Word's internals than any particular class name inside it - a renamed
// child would silently break the name test and cannot break this one.
BOOL StripHasDocument(HWND frame)
{
    StripState* state = FindByFrame(frame);

    // The cheap answer, and the one that is right nearly always: one GetWindow on the document frame
    // this strip is bound to. Nothing below runs while a window has a document in the ordinary way.
    if (state && state->wwf && IsWindow(state->wwf) && GetWindow(state->wwf, GW_CHILD) != NULL)
        return TRUE;

    // Everything else lands here, and they are not the same thing: not bound yet (the janitor has
    // not come round, which is every window for its first half second), bound to a frame Word has
    // destroyed, or bound to one Word has emptied. The last of those is the interesting one - a
    // window whose document frame has been REPLACED reads exactly like a window with no document,
    // and answering FALSE from a stale binding is what would drop it out of the row for good.
    //
    // So the window itself is asked, and the answer is about the whole window rather than about one
    // handle we happen to be holding. This is the only path that pays an enumeration, and it is the
    // path where the alternative is a wrong answer. See PickWwf.
    BOOL holds = FALSE;
    PickWwf(frame, &holds);
    return holds;
}

// Every document frame in this window, for the log. See the declaration for why this exists.
struct WwfList
{
    wchar_t* out;
    int      chars;
    int      at;
    HWND     bound;
    HWND     frame;
};

static BOOL CALLBACK WwfListProc(HWND hwnd, LPARAM param)
{
    WwfList* list = (WwfList*)param;

    wchar_t cls[64] = L"";
    GetClassNameW(hwnd, cls, 64);
    if (_wcsicmp(cls, kWwfClass) != 0)
        return TRUE;

    int left = list->chars - list->at;
    if (left < 64)
        return FALSE;

    HWND inside = GetWindow(hwnd, GW_CHILD);
    wchar_t what[64] = L"empty";
    if (inside)
        GetClassNameW(inside, what, 64);

    // **Whose child it is, when the answer is not "the frame's".**
    //
    // Every rectangle in this file is computed in the FRAME's client coordinates and then handed to
    // SetWindowPos, which reads them as the PARENT's - and those are the same space only while the
    // document frame is a direct child of the frame. Office Tab injects a container of its own into
    // this window and is loading on the machine that reports the document area laying out wrong; if
    // it has taken `_WwF` in with it, every number here is off by that container's origin. Not fixed
    // and not guessed at - measured from wherever this line next turns up, because it cannot be
    // measured on a machine where the other add-in refuses to load.
    HWND owner = GetParent(hwnd);
    wchar_t reparented[48] = L"";
    if (owner != list->frame)
        _snwprintf(reparented, 48, L" PARENT=0x%p NOT THE FRAME", (void*)owner);

    int wrote = _snwprintf(list->out + list->at, (size_t)left,
                           L"  _WwF=0x%p%s visible=%d holds=%s%s",
                           (void*)hwnd, hwnd == list->bound ? L"(bound)" : L"",
                           (int)IsWindowVisible(hwnd), what, reparented);
    if (wrote > 0)
        list->at += wrote;
    return TRUE;
}

void StripDescribeDocumentFrames(HWND frame, wchar_t* out, int chars)
{
    if (!out || chars <= 0)
        return;

    StripState* state = FindByFrame(frame);

    WwfList list;
    list.out   = out;
    list.chars = chars;
    list.at    = 0;
    list.bound = (state && state->wwf) ? state->wwf : NULL;
    list.frame = frame;

    int wrote = _snwprintf(out, (size_t)chars, L"window visible=%d:", (int)IsWindowVisible(frame));
    list.at = (wrote > 0) ? wrote : 0;
    out[chars - 1] = L'\0';

    if (frame && IsWindow(frame))
        EnumChildWindows(frame, WwfListProc, (LPARAM)&list);

    if (list.at <= wrote)
    {
        _snwprintf(out + list.at, (size_t)(chars - list.at), L"  no _WwF at all");
    }
    out[chars - 1] = L'\0';
}

BOOL StripGetNatural(HWND frame, RECT* natural)
{
    StripState* state = FindByFrame(frame);
    if (!state || !state->hasApplied || !natural)
        return FALSE;
    *natural = state->natural;
    return TRUE;
}

// Give a window the *size and shape* of the document frame Word gave the focused one - and only
// those. Sound only because the stack has already made them the same size; the top edge is the one
// part of that rect which belongs to the receiving window and is kept, see below.
//
// **This is the one writer of `natural` that Word did not propose**, so it is the one that can put a
// number there which is right for some other window and wrong for this one. It therefore says so in
// the log, and `why` names the path that asked - a silent writer is why a strip 38px above its own
// ribbon could not be traced to the message that put it there.
void StripSetNatural(HWND frame, const RECT* natural, const wchar_t* why)
{
    StripState* state = FindByFrame(frame);
    if (!state || !state->enabled || !natural || !state->wwf || !IsWindow(state->wwf))
        return;

    // **The top edge is not broadcast.** The stack holds every window at one rectangle, so the left,
    // right and bottom edges of one window's document frame are true of all of them - but the top
    // edge is not a property of the stack. It is the height of that window's own chrome, and Word
    // does not give every window the same chrome.
    //
    // Measured, 2026-08-16, at 192 dpi: a **Protected View** window has a reduced ribbon and a
    // message bar, and its document frame legitimately sits at 318 where a normal window's sits at
    // 356. Opening a downloaded document alongside others made that window active, it pushed its own
    // - correct - interior onto four windows for which it was wrong, and every one of their strips
    // landed 38px up inside the ribbon with the `+` underneath `NetUIHWND` and unclickable. Nothing
    // ever corrected them: the same rule that makes this broadcast necessary - Word lays out only the
    // focused window - is what makes an error in it permanent.
    //
    // A window with no natural of its own has nothing better to go on and still takes what it is
    // given; `ApplyInitial` has not run for it yet.
    RECT want = *natural;
    if (state->hasApplied)
        want.top = state->natural.top;

    RECT applied = want;
    applied.top = want.top + state->stripH;
    if ((applied.bottom - applied.top) < (2 * state->stripH))
        return;

    if (state->hasApplied && SameRect(&want, &state->natural) && SameRect(&applied, &state->applied))
        return;

    // State first, then move: the WM_WINDOWPOSCHANGING this triggers then sees its own proposal as
    // already ours and leaves it alone, instead of shifting it a second time.
    state->natural    = want;
    state->applied    = applied;
    state->hasApplied = TRUE;
    RememberClient(state);
    LogRelayout(state, why);

    SetWindowPos(state->wwf, NULL,
                 applied.left, applied.top,
                 applied.right - applied.left, applied.bottom - applied.top,
                 SWP_NOZORDER | SWP_NOACTIVATE);

    PlaceStrip(state, NULL);
}

// Refit a window's document frame to the window's *own* current size. Needed when a window leaves
// the stack and goes back to the size it had before: Word will not lay out a window it is not
// focused on, so the insets it had are re-applied to the new client area by hand.
void StripRefit(HWND frame)
{
    StripState* state = FindByFrame(frame);
    if (!state || !state->enabled || !state->hasApplied)
        return;

    RECT client;
    if (!GetClientRect(state->frame, &client))
        return;
    if (state->clientAtNatural.cx <= 0 || state->clientAtNatural.cy <= 0)
        return;

    RECT natural;
    natural.left   = state->natural.left;
    natural.top    = state->natural.top;
    natural.right  = (client.right - client.left) - (state->clientAtNatural.cx - state->natural.right);
    natural.bottom = (client.bottom - client.top) - (state->clientAtNatural.cy - state->natural.bottom);

    if (natural.right <= natural.left || natural.bottom <= natural.top)
        return;

    StripSetNatural(frame, &natural, L"refit to the window's own size");
}

// One window's tab row is every window's tab row, so anything that changes it repaints them all.
void StripRefreshTabs(void)
{
    // And anything that changes the row invalidates a panel describing it. Every caller of this is
    // a real change - a title, a dot, the stack, the scroll position - so this is the one place that
    // has to know about the tooltip, rather than each of them. A tab that scrolls out from under its
    // own tooltip, or closes while it is up, would otherwise leave a panel naming a document that is
    // no longer there for as long as the auto-hide takes.
    //
    // TipRowChanged and not TipStop: this used to cancel a tooltip that had been ARMED and not yet
    // shown, which no row change can invalidate and which nothing re-arms while the hand is still.
    TipRowChanged();

    for (int i = 0; i < g_stripCount; i++)
    {
        StripState* state = &g_strips[i];
        if (state->strip && IsWindow(state->strip))
            InvalidateRect(state->strip, NULL, FALSE);
    }
}

static void TryBind(StripState* state)
{
    if (!g_stripEnabled || !state->enabled)
        return;
    if (state->wwf && IsWindow(state->wwf))
        return;
    if (!IsWindow(state->frame))
        return;

    HWND wwf = PickWwf(state->frame, NULL);
    if (!wwf)
        return;                 // not built yet - the janitor will come back

    state->wwf = wwf;
    state->hasApplied = FALSE;

    // Seeded from what is true now, not left at zero: otherwise the first janitor tick reads a
    // window that has been visible all along as having just appeared, and re-derives a layout that
    // did not need re-deriving.
    state->wasVisible = (IsWindowVisible(state->frame) && !IsIconic(state->frame)) ? TRUE : FALSE;
    ApplyMetrics(state);

    // refData carries the state pointer, so the procedure never has to search for it. The entries
    // live in a static array and are never moved, which is what makes that safe.
    if (!SetWindowSubclass(wwf, WwfSubclassProc, kWwfSubclassId, (DWORD_PTR)state))
    {
        LogWrite(L"strip  hwnd=0x%p  SetWindowSubclass on _WwF FAILED (lastError=%lu)",
                 (void*)state->frame, GetLastError());
        state->wwf = NULL;
        return;
    }

    if (!CreateStripWindow(state))
    {
        RemoveWindowSubclass(wwf, WwfSubclassProc, kWwfSubclassId);
        state->wwf = NULL;
        return;
    }

    ReadTitle(state);

    RECT rect;
    ChildRect(state->frame, wwf, &rect);

    // Whose child the document frame is, said out loud the moment it is not the frame's. Every rect
    // in this file is computed in the frame's client space and handed to SetWindowPos, which reads
    // it as the parent's; the two are the same space only while `_WwF` hangs directly off the frame.
    // See the note in WwfListProc - this is the one line that would catch another add-in having
    // taken the document frame in under a container of its own.
    HWND owner = GetParent(wwf);
    LogWrite(L"strip  hwnd=0x%p  bound _WwF=0x%p strip=0x%p  dpi=%d stripH=%d  "
             L"_WwF at (%ld,%ld %ldx%ld)%s",
             (void*)state->frame, (void*)wwf, (void*)state->strip,
             state->dpi, state->stripH,
             rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top,
             owner == state->frame ? L""
                                   : L"  *** ITS PARENT IS NOT THE FRAME - every rect here is in the "
                                     L"wrong space ***");

    ApplyInitial(state);
}

// Let go of the document frame this strip is bound to, keeping the strip itself.
//
// The strip is a child of the FRAME, not of `_WwF`, so it survives this and is re-used by the next
// TryBind - which matters: destroying and rebuilding it would lose the surface, the font and the
// hover state, and would flicker.
//
// The old frame is put back the size Word wanted it before the subclass comes off, for the same
// reason Restore does it. A retired document frame is one Word may yet show again, and one left
// 48px short with nobody watching it is a defect that would surface minutes later with nothing to
// connect it to.
static void Unbind(StripState* state, const wchar_t* why)
{
    if (!state->wwf)
        return;

    LogWrite(L"strip  hwnd=0x%p  letting go of _WwF=0x%p - %s",
             (void*)state->frame, (void*)state->wwf, why);

    if (IsWindow(state->wwf))
    {
        if (state->hasApplied)
        {
            SetWindowPos(state->wwf, NULL,
                         state->natural.left, state->natural.top,
                         state->natural.right - state->natural.left,
                         state->natural.bottom - state->natural.top,
                         SWP_NOZORDER | SWP_NOACTIVATE);
        }
        RemoveWindowSubclass(state->wwf, WwfSubclassProc, kWwfSubclassId);
    }

    state->wwf        = NULL;
    state->hasApplied = FALSE;
}

// Put Word's layout back exactly as we found it. Called on detach and on shutdown - an add-in that
// leaves the host's window layout altered after it is switched off is a bug report waiting to
// happen, and `enabled` has to go down first or the handler would simply shift it again.
static void Restore(StripState* state)
{
    state->enabled = FALSE;

    // A drag whose strip is being destroyed is over. Forgotten rather than undone: a window going
    // away says nothing about whether the user meant the moves they had already made, and
    // DestroyWindow releases the capture, which would otherwise arrive as a cancel.
    if (state->strip && g_dragStrip == state->strip)
        DragForget();

    // Before the strip goes: TipStop reaches its state through the strip handle, so a tooltip
    // disarmed afterwards would be disarmed through a window that no longer exists.
    if (state->strip && g_tipStrip == state->strip)
        TipStop();

    if (state->tip && IsWindow(state->tip))
    {
        DestroyWindow(state->tip);
        state->tip = NULL;
    }

    // Same order and the same reason as the tooltip above: GhostStop reaches its state through the
    // strip handle. DragForget a few lines up has already called it for the drag that was running,
    // but a ghost can be up on a strip that is not the one holding the capture - every window in the
    // stack has one - so this is not the same test.
    if (state->strip && g_ghostStrip == state->strip)
        GhostStop();

    if (state->ghost && IsWindow(state->ghost))
    {
        DestroyWindow(state->ghost);
        state->ghost = NULL;
    }

    if (state->strip && IsWindow(state->strip))
    {
        DestroyWindow(state->strip);
        state->strip = NULL;
    }

    // One writer of "let go of the document frame", rather than a second copy here that has to be
    // kept in step with it.
    Unbind(state, L"the strip is being taken down");

    if (state->font)
    {
        DeleteObject(state->font);
        state->font = NULL;
    }
    ReleaseSurface(state);
}

// ---------------------------------------------------------------------------------------------
// The janitor.
//
// The geometry itself is event-driven and needs no timer. This exists for the things that have no
// event: `_WwF` not existing yet when a frame is first subclassed, `_WwF` being replaced under us,
// and the document title changing (Word does not always send the frame a WM_SETTEXT we can see).
// Half a second is slow enough to be free and fast enough that nothing is visibly late.
// ---------------------------------------------------------------------------------------------

// ---------------------------------------------------------------------------------------------
// The unsaved-changes dot.
//
// This is the first thing the add-in reads about Word's *documents* rather than its windows, and it
// is a poll. The standing rule is that transient state must be heard rather than sampled - got wrong
// twice over Word's save prompt, at real cost - and it is worth being exact about why it does not
// bite here. The rule is about state that can appear and disappear *inside* one tick, which no
// sampler catches at any cadence. "This document has unsaved changes" is not that: it goes true on a
// keystroke and stays true until a save. A sampler that is one tick late is one tick late.
//
// There is also no event to hear. Word's Application events cover documents opening, closing and
// being saved; none of them is "the modified flag changed", and typing raises none of them. Hearing
// this properly would mean an IConnectionPoint sink for events that still would not answer the
// question. The honest mechanism is the one that can actually observe the state.
//
// The whole row in one pass rather than one lookup per tab - see WordTabReadModified.
//
// **AND THE CADENCE IS NOT FIXED ANY MORE. Measured on the user's work rig, in their own report:
// mean 96,000-155,000us and worst 518,726us with three to five SharePoint-backed documents open,
// against ~430us for one local document.** Half a second on Word's UI thread, twice a second: this
// is a poll that had only ever been measured against local files, and a document that lives on
// SharePoint answers three orders of magnitude slower. There is nothing to fix in the question - the
// add-in asks Word for Document.Saved, and how long Word takes to answer is Word's business - so
// what changes is how often, and how much is asked for.
//
// Two rules, and each of them is a measurement rather than a preference:
//
//   1. **The whole row is asked for only as often as the whole row costs.** RungFor turns a pass's
//      cost into a rung on a ladder that holds the poll to about 1% of the UI thread, whatever that
//      cost turns out to be. On this rig a full pass is a few hundred microseconds, RungFor returns
//      rung 0, and the cadence is exactly what it has always been - twice a second, unchanged.
//
//      **On theirs the first cut of this did not work, and the report that came back is why the
//      governor now looks the way it does.** Backing off is not enough on its own if coming back is
//      free: f333349 oscillated between 500ms and 20s about three times a minute for hours, spending
//      ~5.5% of Word's UI thread in half-second lumps. See PollModified for the measurement and for
//      what replaced it.
//
//   2. **In between, only the ACTIVE window is asked about**, at the same governed cadence and by
//      the same measurement. That is not a compromise on correctness so much as a statement of where
//      the flag can change: a document is edited in the window that has the keyboard, and saved from
//      the window that has the keyboard. What the periodic full pass is for is the rest - a
//      background document Word saves by itself, which is what AutoSave on a SharePoint document
//      does - and being a few seconds late on that is a dot that appears late, not a wrong one.
//
// The whole thing is off under `TabDot=0`, which is the switch the user was given as a mitigation.
// ---------------------------------------------------------------------------------------------

// The ladder the poll backs off along, in janitor ticks. Half a second a tick, so the ratio of time
// spent asking to time elapsed comes out at or under about 1% at every rung: a 200ms answer is asked
// for once a minute, a 15ms answer once every six seconds, and anything under 4ms is asked for on
// every tick, which is the behaviour this had before there was a governor.
//
// Rungs rather than bare intervals, because the governor now moves along them a step at a time in one
// direction and jumps in the other. See PollModified for why that asymmetry is the whole fix.
static const int kRungTicks[] = { 1, 4, 12, 40, 120 };   // 500ms, 2s, 6s, 20s, 60s
static const int kRungCount   = (int)(sizeof(kRungTicks) / sizeof(kRungTicks[0]));

// Which rung a pass of this cost belongs on.
static int RungFor(LONGLONG us)
{
    if (us > 200000) return 4;      // 60s
    if (us > 50000)  return 3;      // 20s
    if (us > 15000)  return 2;      // 6s
    if (us > 4000)   return 1;      // 2s
    return 0;                       // twice a second
}

// One ask, timed, with the flags and the log lines it produces. `count` frames in, TRUE if Word
// answered. `us` receives what it cost whether or not it did - a refusal that took a fifth of a
// second still cost a fifth of a second, and the governor has to see that or a Word that is slow to
// say no would be asked twice a second forever.
static BOOL AskModified(StripState** owners, HWND* frames, int count, LONGLONG* us)
{
    LARGE_INTEGER before, after, freq;
    QueryPerformanceFrequency(&freq);
    QueryPerformanceCounter(&before);

    BOOL modified[MAX_STRIPS];
    BOOL answered = WordTabReadModified(frames, count, modified);

    QueryPerformanceCounter(&after);
    *us = (freq.QuadPart > 0) ? ((after.QuadPart - before.QuadPart) * 1000000 / freq.QuadPart) : 0;

    if (!answered)
        return FALSE;               // every flag keeps its last value - see WordTabReadModified

    BOOL changed = FALSE;
    for (int i = 0; i < count; i++)
    {
        if (owners[i]->modified == modified[i])
            continue;

        owners[i]->modified = modified[i];
        changed = TRUE;

        // A dot is drawn and never stored in a control, so there is no reading one from outside the
        // process. This line is that way, and it is pipe-wrapped for the same reason the tab name's
        // is: check-dot.ps1 asserts the state against it and then photographs the button, because a
        // log line about what was computed is not evidence of what was drawn.
        LogWrite(L"strip  hwnd=0x%p  dot |%s|  document |%s|",
                 (void*)owners[i]->frame, owners[i]->modified ? L"on" : L"off", owners[i]->title);
    }

    if (changed)
        StripRefreshTabs();

    return TRUE;
}

static void PollModified(void)
{
    if (!g_dotEnabled)
        return;

    // A batch close is a run of WM_CLOSEs with Word's save prompt in between them, driven by us. It
    // is the one stretch where Word is repeatedly part-way through something we asked for, and the
    // dot is not worth a call into the object model during it. Nothing is lost: the flags keep what
    // they had, and the tabs they belong to are being closed.
    if (StackCloseInFlight())
        return;

    HWND        frames[MAX_STRIPS];
    StripState* owners[MAX_STRIPS];
    int count = 0;

    for (int i = 0; i < g_stripCount && count < MAX_STRIPS; i++)
    {
        StripState* state = &g_strips[i];
        if (!state->enabled || !state->frame || !IsWindow(state->frame))
            continue;
        if (!state->wwf || !IsWindow(state->wwf))
            continue;

        // A disabled frame is a frame with a modal dialog on it - that is what EnableWindow means
        // here, and it is the same signal the batch close listens for as an event. Word is asking
        // the user something; it is not the moment to ask Word something. The whole pass stands
        // down rather than skipping the one window, because the object model is Word's, not that
        // window's, and a tick missed is a flag unchanged.
        if (!IsWindowEnabled(state->frame))
            return;

        owners[count] = state;
        frames[count] = state->frame;
        count++;
    }

    if (count == 0)
        return;

    // Which pass is due, if either.
    //
    // Two counters rather than one, and they are independent on purpose: a row that has backed off
    // to once a minute must not take the window the user is typing in with it. Both run down every
    // tick, and each is re-set from what its own kind of pass last cost.
    static int fullIn   = 0;
    static int activeIn = 0;

    if (fullIn > 0)   fullIn--;
    if (activeIn > 0) activeIn--;

    BOOL full = (fullIn == 0) ? TRUE : FALSE;
    if (!full && activeIn > 0)
        return;

    // Which window, when this is not the full pass. GetForegroundWindow rather than the stack's idea
    // of the active tab: the question is which document can be being typed into, and that is the
    // window with the keyboard. When the answer is not one of ours - another application is in front
    // - nothing of ours can be changing, and the pass is skipped without asking Word anything.
    int only = -1;
    if (!full)
    {
        HWND front = GetForegroundWindow();
        for (int i = 0; i < count; i++)
            if (frames[i] == front)
                only = i;
        if (only < 0)
            return;
    }

    LONGLONG us    = 0;
    int      asked = full ? count : 1;

    if (full) AskModified(owners, frames, count, &us);
    else      AskModified(&owners[only], &frames[only], 1, &us);

    // The governor. Set from what this pass cost, including a pass Word refused to answer: a refusal
    // that took a fifth of a second cost a fifth of a second, and a Word that is slow to say no would
    // otherwise be asked twice a second forever.
    //
    // **It climbs in one step and comes down one rung at a time.** That is the reverse of what
    // f333349 shipped, and the work rig's own log is why it had to be turned round. That version took
    // the cadence from the *cheaper* of the last two passes and let a single cheap pass restore
    // twice-a-second, reasoning that being too eager costs a few milliseconds. On a rig whose
    // documents live on SharePoint it does not. Measured there a pass costs either ~300-600us with
    // the answer warm or 90,000-518,000us with it cold and nothing in between, so the cheaper of two
    // was nearly always the warm one and the cadence snapped back to 500ms after every cheap sample.
    // The log has that cycle repeating about every twenty-two seconds for hours:
    //
    //     10:20:18  took 268 us    (and 494927 us before)  -> now asking every 500 ms
    //     10:20:27  took 95044 us  (and 486406 us before)  -> now asking every 20000 ms
    //     10:20:46  took 344 us    (and 95044 us before)   -> now asking every 500 ms
    //     10:20:49  took 92637 us  (and 501877 us before)  -> now asking every 20000 ms
    //
    // ...costing, by its own summary line, mean 135302us over 480 passes in twenty minutes. That is
    // about 5.5% of Word's UI thread against the 1% this ladder was drawn for, and it is spent as
    // half-second freezes of the thread Word draws with, not as a smear.
    //
    // The reason it could not settle is worth stating plainly, because it is not a tuning problem:
    // **the governor's input is not independent of its output.** Asking often keeps SharePoint's
    // answer warm and therefore cheap; backing off lets it go cold and therefore expensive. A
    // governor reading a quantity its own cadence controls has two stable states and flips between
    // them, which is exactly what the log shows - thrashing at 500ms, or pinned at 60s for two hours
    // and twenty minutes with every single pass slow.
    //
    // So the basis is the **dearer** of the last two passes, one slow pass backs off at once, and
    // coming back takes a rung at a time. That is five cheap passes from once a minute to twice a
    // second, not four: because the basis is the dearer of two, the first cheap pass only clears the
    // slow sample out of `prev` and moves nothing, and the four rungs come after it. A single lucky
    // warm answer in the middle of trouble therefore changes the cadence not at all. Being too slow
    // now costs a background dot that is late; being too eager costs a visible stutter, and the
    // report says which of those is actually happening.
    //
    // What makes the staleness affordable is rule 2 above: the active window keeps its own counter
    // and its own rung, so the tab being typed into is still asked twice a second however far the row
    // has backed off. What goes late is a background document Word saved by itself.
    static int      fullRung   = 0;
    static int      activeRung = 0;
    static LONGLONG prevFull   = 0;
    static LONGLONG prevActive = 0;
    static BOOL     seenFull   = FALSE;
    static BOOL     seenActive = FALSE;

    int*      rung = full ? &fullRung : &activeRung;
    LONGLONG* prev = full ? &prevFull : &prevActive;
    BOOL*     seen = full ? &seenFull : &seenActive;

    // The first pass of a process is thrown away rather than measured. It is where Word builds its
    // automation machinery - 27547us against a steady state three orders of magnitude below it - and
    // under the old min it was harmless for free, because min(anything, 0) is 0. Taking the dearer of
    // two makes it the opposite of harmless: one cold start would put a healthy machine on a
    // six-second cadence and cost four more passes to walk back off it. It is not stored either, so
    // the pass after it is judged on itself rather than against the cold one.
    LONGLONG lastTime = *prev;
    if (!*seen)
    {
        *seen = TRUE;
    }
    else if (full)
    {
        *prev = us;
        LONGLONG basis = (us > lastTime) ? us : lastTime;

        int want = RungFor(basis);
        if (want > *rung)
            *rung = want;                                   // a slow Word stops being asked, at once
        else if (want < *rung)
            (*rung)--;                                      // and is trusted back a rung at a time
    }
    else
    {
        // **The active pass is judged on itself alone, and carries nothing between passes.** The
        // two-sample rule above is only meaningful when consecutive passes measure the same thing,
        // and the whole row is the same set every time. The active pass is not: switching tabs
        // changes which document is being asked about, so "the dearer of the last two" would let one
        // slow document set the cadence for every other tab the user then moves to. Memoryless is
        // the honest reading when the subject changes underneath the sampler.
        //
        // Nor does it need the stickiness. What made the row sticky is that its cost depends on its
        // own cadence, and one foreground window does not have that feedback: Word keeps the
        // document being looked at live. That is why the work rig's active pass never once exceeded
        // 4ms across five and a half hours while its row passes were running to half a second.
        //
        // **If that assumption is ever wrong, the log already says so in its own words** - "the
        // active window took ... - now asking every ..." - and that line is the one to look for in
        // the next report. It has never yet appeared after the first pass of a process.
        *prev = us;
        *rung = RungFor(us);
    }

    if (*rung < 0)           *rung = 0;
    if (*rung >= kRungCount) *rung = kRungCount - 1;

    int ticks = kRungTicks[*rung];
    if (full) fullIn   = ticks;
    else      activeIn = ticks;

    // Said when the cadence changes and never per pass. A line here every half second would be a
    // file write twice a second for a number that does not move; a line when it moves is the whole
    // diagnosis of "Word went slow and the add-in noticed".
    static int lastFull   = -1;
    static int lastActive = -1;
    int* last = full ? &lastFull : &lastActive;
    if (ticks != *last)
    {
        *last = ticks;
        LogWrite(L"strip  dot poll: %s took %lld us for %d window(s) (and %lld us the time before) "
                 L"- now asking every %d ms",
                 full ? L"the whole row" : L"the active window", us, asked, lastTime, ticks * 500);
    }

    // What it costs. The out-of-process measurement in tools\probe-saved.ps1 put a two-window pass
    // at ~2ms across a process boundary, and in-process on one STA thread this is a direct call with
    // no marshalling at all - but that is a reason to expect a small number, not a substitute for
    // having one. This runs inside Word's input loop; a claim about its cost should come off the rig
    // it runs on.
    //
    // Reported two ways, because the first version of this reported only new records and that turned
    // out to say almost nothing: the very first pass over a newly-arrived window cost 2727us against
    // a steady state of about 200, so the record line described a cold start and then fell silent
    // forever. A worst case with no typical case beside it is not a measurement of what something
    // costs.
    //
    // Kept apart for the two kinds of pass, like the rungs above and for the same reason: they
    // measure different things. Pooling them was a real fault in the instrument rather than an
    // untidiness. The field log has
    //
    //     dot poll: 1 window(s), 480 passes, mean 232103 us, worst 516165 us
    //
    // which reads as the ACTIVE window costing a quarter of a second a pass - and the comment fifty
    // lines above tells the next reader that a line saying so is the one to look for, because it
    // would mean the whole design's assumption had failed. It had not. "1 window(s)" is only
    // whichever kind of pass happened to run last before the counter came due, and the mean is over
    // both. The number that matters most in the report this instrument exists to produce could not
    // be read at all.
    static LONGLONG worstEver[2] = { -1, -1 };
    LONGLONG* record = &worstEver[full ? 1 : 0];
    if (us > *record)
    {
        *record = us;
        LogWrite(L"strip  dot poll: %s in %lld us (the most it has ever taken)",
                 full ? L"the whole row" : L"the active window", us);
    }

    // ...and a summary: once ten seconds in, and every four minutes after that.
    //
    // The early one is the point. The record line above is dominated by the very first pass of a
    // process - 27547us against a steady state three orders of magnitude below it, because the
    // janitor is usually the first thing in the process ever to touch Word's object model and that
    // first call is Word building its automation machinery. A cold outlier with nothing beside it
    // says almost nothing about what this costs. The four-minute cadence afterwards is rare enough
    // not to crowd a log that rolls at half a megabyte.
    static int      passes[2]   = { 0, 0 };
    static int      reportAt[2] = { 20, 20 };
    static LONGLONG total[2]    = { 0, 0 };
    static LONGLONG worst[2]    = { 0, 0 };

    const int kind = full ? 1 : 0;

    passes[kind]++;
    total[kind] += us;
    if (us > worst[kind])
        worst[kind] = us;

    if (passes[kind] >= reportAt[kind])
    {
        // The window count stays on the line - it is what a row pass actually asked, and it moves -
        // but it is no longer the only thing saying which kind of pass this was.
        LogWrite(L"strip  dot poll: %s, %d window(s), %d passes, mean %lld us, worst %lld us",
                 full ? L"the whole row" : L"the active window",
                 asked, passes[kind], total[kind] / passes[kind], worst[kind]);
        passes[kind]   = 0;
        total[kind]    = 0;
        worst[kind]    = 0;
        reportAt[kind] = 480;
    }
}

static void CALLBACK JanitorProc(HWND hwnd, UINT msg, UINT_PTR id, DWORD tick)
{
    (void)hwnd; (void)msg; (void)id; (void)tick;

    for (int i = 0; i < g_stripCount; i++)
    {
        StripState* state = &g_strips[i];
        if (!state->enabled || !IsWindow(state->frame))
            continue;

        // Only when a DPI override is configured, and configured means "was set before Word
        // started". A machine with no TabDpi value never reaches this line and never pays the
        // registry read, so production behaviour and production cost are both exactly unchanged.
        //
        // In production the trigger for a re-scale is WM_DPICHANGED, which is the real event and
        // arrives on its own. This is here because a rig with one monitor cannot produce that
        // event, and a synthetic one cannot be posted from another process: WM_DPICHANGED carries a
        // suggested RECT by pointer, and the frame proc chains it on to Word, which would
        // dereference an address from the wrong process. So the override is the way in, and it
        // reaches the same StripOnFrameDpiChanged that the real message reaches.
        if (g_dpiWatch && DpiOf(state->frame) != state->dpi)
            StripOnFrameDpiChanged(state->frame);

        if (!state->wwf || !IsWindow(state->wwf))
        {
            TryBind(state);
            continue;
        }

        // Has the document frame been REPLACED rather than emptied?
        //
        // The test is deliberately narrow: only a binding that has gone empty or invisible is
        // questioned at all, and it is only given up for a candidate that is both. So this never
        // fires while a window is behaving - one `_WwF`, a document in it - and it cannot start a
        // fight between two frames that both look usable. A window with genuinely nothing open
        // reaches the enumeration and finds nothing better, which costs one EnumChildWindows a tick
        // on windows showing Word's Start screen and nothing at all anywhere else.
        //
        // Queued as "a View change closes tabs", and NOT reproduced here - see PickWwf.
        if (GetWindow(state->wwf, GW_CHILD) == NULL || !IsWindowVisible(state->wwf))
        {
            BOOL holds = FALSE;
            HWND better = PickWwf(state->frame, &holds);
            if (holds && better && better != state->wwf)
            {
                LogWrite(L"strip  hwnd=0x%p  Word has BUILT A SECOND document frame: bound to "
                         L"0x%p (empty=%d visible=%d), rebinding to 0x%p which holds the document",
                         (void*)state->frame, (void*)state->wwf,
                         (int)(GetWindow(state->wwf, GW_CHILD) == NULL),
                         (int)IsWindowVisible(state->wwf), (void*)better);
                Unbind(state, L"it was replaced");
                TryBind(state);
                continue;
            }
        }

        // Hidden -> shown. Everything Word did to this window's layout while it was hidden was
        // ignored on purpose (see AdjustProposed), so what is there now is Word's own idea of the
        // layout and is exactly what "natural" means. Re-derive from it rather than carrying
        // forward a rect from before the window went away.
        // "On screen", not merely WS_VISIBLE: a minimised window keeps that style, and its layout
        // is frozen for the same reason a hidden one's is.
        BOOL visible = (IsWindowVisible(state->frame) && !IsIconic(state->frame)) ? TRUE : FALSE;
        if (visible && !state->wasVisible)
        {
            // ApplyInitial does nothing if the document frame is still where we put it, which is
            // the common case - Word usually hides and shows a window without touching its layout.
            // It only re-derives when Word actually did rearrange things while we were not looking.
            state->stripPlaced = FALSE;
            ApplyInitial(state);
        }
        state->wasVisible = visible;

        ReadTitle(state);

        // Cheap when nothing has moved - it compares against the strip's real rect and returns.
        // This is what makes any drift heal within half a second rather than staying wrong.
        if (visible)
            PlaceStrip(state, NULL);

        // The strip belongs with the document frame: shown when it is shown, hidden when Backstage
        // or a minimise takes it away.
        if (state->strip && IsWindow(state->strip))
        {
            BOOL wantVisible = IsWindowVisible(state->wwf) ? TRUE : FALSE;
            BOOL isVisible   = IsWindowVisible(state->strip) ? TRUE : FALSE;
            if (wantVisible != isVisible)
                ShowWindow(state->strip, wantVisible ? SW_SHOWNA : SW_HIDE);
        }
    }

    // Membership is decided from what is true right now - visible, has a document frame - so it is
    // re-decided on the same cadence rather than tracked through events that Word does not always
    // send.
    StackJanitor();

    // After StackJanitor, so that a window which has just left the stack is not asked about, and a
    // window which has just joined it is. Before the chrome sample, which is the cheaper of the two
    // and has waited two seconds already.
    PollModified();

    // Has Word changed colour underneath us?
    //
    // This is a poll, and this project's standing rule is that transient state must be heard rather
    // than sampled - got wrong twice over Word's save prompt, at some cost. The rule does not apply
    // here and it is worth being precise about why: it is about state that can appear and disappear
    // *inside* one tick, which a sampler cannot see at any cadence. A theme is not that. Somebody
    // changes it in File > Account, it stays changed, and a sampler that is one tick late is one
    // tick late rather than wrong.
    //
    // Sampling is the honest mechanism here for a second reason too: there is no event to hear. The
    // strip is a WS_CHILD, and WM_SETTINGCHANGE is broadcast to top-level windows only.
    //
    // Every fourth tick, once, off the first strip that can answer - not once per window. Cost is a
    // GetDC and twenty-four GetPixels every two seconds, and it stops at the first strip that gives
    // a usable answer.
    //
    // **The window in front is asked first, and its answer is the one reported.** Two reasons, and
    // neither is a preference. It is the only window whose pixels can be read at all - the stack
    // holds every window at one rectangle, so every other strip's ribbon is behind this one and
    // reads back as "something is covering Word's ribbon". And when nothing can be sampled, the
    // reason that reaches the log used to be whichever window happened to be last in an array, so
    // "the front window is showing a message bar" could be reported as "something is covering the
    // ribbon" - a true statement about a different window, which is exactly the three-steps-from-
    // the-cause shape that has cost this subsystem two diagnoses.
    if (g_sampleEnabled)
    {
        static int countdown = 0;
        if (--countdown <= 0)
        {
            countdown = 4;

            // The outcome is logged when it *changes*, never per tick. A sampler that cannot read
            // the ribbon looks exactly like a sampler that agrees with the fallback, because on this
            // rig the two answers are the same number - so silence here costs a whole diagnosis.
            static wchar_t lastWhy[192] = L"";
            const wchar_t* why      = L"no strip was ready to be asked";
            const wchar_t* frontWhy = NULL;
            const wchar_t* barWhy   = NULL;

            HWND front   = GetForegroundWindow();
            BOOL applied = FALSE;

            for (int pass = 0; pass < 2 && !applied; pass++)
            {
                for (int i = 0; i < g_stripCount; i++)
                {
                    StripState* state = &g_strips[i];

                    BOOL isFront = (state->frame == front);
                    if (pass == 0 ? !isFront : isFront)
                        continue;                       // pass 0 is the front window, pass 1 the rest

                    if (!state->enabled || !state->strip || !IsWindow(state->frame))
                        continue;
                    if (!IsWindowVisible(state->frame) || IsIconic(state->frame))
                        continue;

                    COLORREF chrome;
                    const wchar_t* mine = L"ok";
                    if (SampleChrome(state, &chrome, &mine))
                    {
                        // Which window answered is carried into the log line by AdoptSampledChrome:
                        // the palette is one object shared by every strip, so a line saying only
                        // "sampled from Word's ribbon" cannot distinguish the healthy case from this
                        // subsystem's one real bug - a window that should not have been asked
                        // speaking for all of them.
                        //
                        // `applied` means a window ANSWERED, not that the palette changed, which is
                        // why it is still set when the gate declines to adopt. It decides whose
                        // reason reaches the log, and the reason here is "ok" either way: the sample
                        // succeeded, and what was then done with it is a separate line in its own
                        // words.
                        AdoptSampledChrome(chrome, state->frame);
                        why     = mine;
                        applied = TRUE;
                        break;
                    }

                    why = mine;
                    if (isFront)
                        frontWhy = mine;
                    if (mine == kBandDoesNotReachTop)
                        barWhy = mine;
                }
            }

            // Precedence when nobody could answer: the window the user is looking at, then a
            // structural reason from any window, then whatever was tried last.
            //
            // The middle one is not tidiness. `front` is GetForegroundWindow(), which is often NOT
            // one of ours - another application, a Word dialog, a menu - and then frontWhy is never
            // set and the reason falls back to array order, which is the non-determinism this
            // rewrite exists to remove. "Something is docked above a strip" is a fact about this
            // Word that stays true whoever owns the desktop; "something is covering Word's ribbon"
            // is trivially true of every window whenever another application is in front, so it is
            // the reason that says least.
            if (!applied)
            {
                if (frontWhy)     why = frontWhy;
                else if (barWhy)  why = barWhy;
            }

            if (wcscmp(why, lastWhy) != 0)
            {
                wcsncpy(lastWhy, why, 191);
                lastWhy[191] = 0;
                LogWrite(L"strip  chrome sample: %s", why);
            }
        }
    }
}

// ---------------------------------------------------------------------------------------------
// Entry points, called from frames.cpp.
// ---------------------------------------------------------------------------------------------

void StripStart(void)
{
    g_stripEnabled = WordTabReadFlag(L"Strip", TRUE);

    if (!g_stripEnabled)
    {
        LogWrite(L"StripStart  disabled by HKCU\\Software\\WordTab\\Strip=0");
        return;
    }

    // The close and new-document buttons, switchable like every other piece. Off, the tabs are
    // exactly what they were before this slice - which is what makes them bisectable if something
    // about them ever misbehaves.
    g_buttonsEnabled = WordTabReadFlag(L"TabButtons", TRUE);

    // The context menu, on its own switch. Off, a right-click on the strip does nothing at all -
    // which is what it did before this slice.
    g_menuEnabled = WordTabReadFlag(L"TabMenu", TRUE);

    // Dragging a tab to reorder it. Off, a press on a tab switches to it and nothing else, which is
    // what it did before this slice - and the mouse capture is never taken, so the whole mechanism
    // is out of the way rather than merely inert.
    g_dragEnabled = WordTabReadFlag(L"TabDrag", TRUE);

    // The scrolling tab row. This one is not "off gives back what it did before", because what it
    // did before was draw the + on top of the last tab and hit-test the tab first - the user pressed
    // a + and closed a document. A switch whose only effect is to reproduce that would exist to
    // reproduce it.
    //
    // What it gives instead is the other honest answer to too many documents: no minimum width at
    // all, so the tabs divide the row between them however many there are and every one of them is
    // on screen. That is a worse row to read and a better one to be *sure* of, which makes it the
    // thing to reach for if the scrolling row ever strands somebody's document off the end of a
    // strip on a machine this has not run on.
    g_scrollEnabled = WordTabReadFlag(L"TabScroll", TRUE);

    // The look. Off, the strip is what it was before this slice: flat rectangles with a full border,
    // an aliased one-pixel x and +, a bare empty row, and a context menu drawn by the system in the
    // system's colours. The palette is still derived rather than hand-picked, because that is a
    // correction rather than a style - but nothing is composited and nothing is owner-drawn.
    g_lookEnabled = WordTabReadFlag(L"TabStyle", TRUE);

    // And the sampler on its own switch, because it is the one part of this that reads pixels out of
    // a window belonging to Word. Off, the palette comes from the Office theme registry value the
    // way it always did - which is the setting to try first if the strip is ever the wrong colour on
    // a machine this has not been run on.
    //
    // This line went missing when the tab-name slice inserted the block below it, and the switch was
    // dead from then until the dot slice found it: g_sampleEnabled stayed TRUE whatever the registry
    // said, and the summary line at the end of this function reported "on" every time - a log that
    // named a setting it could not observe. The escape hatch the notes call "the first thing to try
    // if the strip is ever the wrong colour on a machine this has not run on" did nothing at all.
    //
    // Restored exactly as it was, ternary included - `git show c6a0e11:src/native/strip.cpp` line
    // 3295. The first attempt at putting it back dropped the `g_lookEnabled ?` and would have made
    // TabStyle=0 sample where it used to fall back to the registry, which is a change of behaviour
    // wearing the clothes of a repair. Whether the sampler should be tied to the look at all is a
    // real question and it is not this slice's to answer.
    g_sampleEnabled = g_lookEnabled ? WordTabReadFlag(L"TabThemeSample", TRUE) : FALSE;

    // Trimming Word's state annotations off the tab name. Off, the tab reads exactly what Word's
    // title bar reads, minus the application suffix. This is a switch rather than nothing because it
    // is the one part of the title rule that is a *policy* - somebody may want to see "Read-Only" on
    // the tab, and a document genuinely named `a  -  b.docx` is the case where the rule is wrong.
    // Removing the suffix from the end rather than the first match is not on the switch: that was a
    // defect, and a setting that restored it would exist only to put the bug back.
    g_titleTrimEnabled = WordTabReadFlag(L"TabTitleTrim", TRUE);

    // The unsaved-changes dot. Off, a tab's close button is always an x and nothing reads Word's
    // object model on the janitor - the poll returns before it asks, so the switch takes the whole
    // mechanism out of the way rather than merely hiding what it draws. That is the point of it:
    // this is the first sustained use of the object model in the add-in, and if a Word somewhere
    // objects to being asked twice a second, this is the setting that stops the asking.
    g_dotEnabled = WordTabReadFlag(L"TabDot", TRUE);

    // The hover tooltip. Off, no timer is ever armed, no window is ever created and Word is never
    // asked where a document lives - the same shape of "off" as TabDot, and for a sharper version of
    // the same reason: this one reads the object model in response to the pointer moving, so a Word
    // that objects to being asked has to be able to stop the asking without giving up the tabs.
    g_tipEnabled = WordTabReadFlag(L"TabTip", TRUE);

    // The card carried under the pointer while a tab is out of the row. Off, the gesture is exactly
    // what it was before - the row lets go and the cursor changes - so this is the one switch here
    // whose `0` removes feedback rather than behaviour. It exists because it draws over Word's
    // document with a window this add-in owns, and that is the kind of thing a machine somewhere
    // will render badly; the tear-off itself should not have to be given up over it.
    g_ghostEnabled = WordTabReadFlag(L"TabGhost", TRUE);

    // Clamped rather than trusted. This is a DWORD anybody can type into regedit, it is fed to
    // CreateFontIndirect, and the two ends of the range fail differently: 0 means "let the mapper
    // choose", which is a font this file did not pick, and a large number draws a label taller than
    // the strip that clips instead of erroring. A value out of range is a typo, so the clamp keeps
    // the tabs legible and says so in the log rather than honouring it.
    {
        DWORD wanted = WordTabReadNumber(L"TabFontSize", (DWORD)TAB_LOGICAL_FONT);
        int   size   = (int)wanted;
        if (size < TAB_FONT_MIN) size = TAB_FONT_MIN;
        if (size > TAB_FONT_MAX) size = TAB_FONT_MAX;
        if ((DWORD)size != wanted)
            LogWrite(L"strip  TabFontSize=%lu is outside %d..%d - using %d",
                     (unsigned long)wanted, TAB_FONT_MIN, TAB_FONT_MAX, size);
        g_fontLogical = size;
    }

    if (!g_stripClass)
    {
        WNDCLASSEXW wc;
        memset(&wc, 0, sizeof(wc));
        wc.cbSize        = sizeof(wc);
        wc.style         = CS_HREDRAW | CS_VREDRAW;
        wc.lpfnWndProc   = StripWndProc;
        wc.hInstance     = g_module;
        wc.hCursor       = LoadCursorW(NULL, IDC_ARROW);
        wc.lpszClassName = kStripClass;
        g_stripClass = RegisterClassExW(&wc);
    }

    // CS_DROPSHADOW is what tells a panel from the page behind it when the two are nearly the same
    // colour, and it is the system's shadow rather than one of ours: a menu gets it for free and a
    // tooltip that did not have it would be the one floating surface here without one. CS_SAVEBITS
    // because it covers a document that would otherwise repaint every time it appears.
    if (!g_tipClass)
    {
        WNDCLASSEXW wc;
        memset(&wc, 0, sizeof(wc));
        wc.cbSize        = sizeof(wc);
        wc.style         = CS_DROPSHADOW | CS_SAVEBITS;
        wc.lpfnWndProc   = TipWndProc;
        wc.hInstance     = g_module;
        wc.hCursor       = LoadCursorW(NULL, IDC_ARROW);
        wc.lpszClassName = kTipClass;
        g_tipClass = RegisterClassExW(&wc);
    }

    // No CS_DROPSHADOW: the carried card draws its own lift, and a system shadow under it would be a
    // second shadow at a different offset. No CS_SAVEBITS either - that trades memory for the repaint
    // of what is underneath, and this window moves continuously, so what is underneath is never the
    // same twice.
    //
    // CS_HREDRAW | CS_VREDRAW, though, and it is not copied from the strip's class without thinking.
    // The card is followed by SetWindowPos on every mouse-move at the SAME size, which invalidates
    // nothing - so these cost nothing in the case that happens thousands of times. They matter in the
    // one that is rare and silent: the row can re-lay out mid-drag, a tab's width with it, and
    // without them the pixels the card gains are whatever was in the backing store.
    if (!g_ghostClass)
    {
        WNDCLASSEXW wc;
        memset(&wc, 0, sizeof(wc));
        wc.cbSize        = sizeof(wc);
        wc.style         = CS_HREDRAW | CS_VREDRAW;
        wc.lpfnWndProc   = GhostWndProc;
        wc.hInstance     = g_module;
        wc.hCursor       = LoadCursorW(NULL, IDC_ARROW);
        wc.lpszClassName = kGhostClass;
        g_ghostClass = RegisterClassExW(&wc);
    }

    // Something has to be on the palette before the first paint. The registry answer, which is the
    // one available this early: no window of Word's exists yet to take a colour off. The janitor
    // replaces it with the sampled one within two seconds, and ApplyPalette does nothing at all if
    // the two agree - which on this rig they do.
    ApplyFallbackPalette(L"from the Office theme setting");

    // A thread timer rather than a window timer: it needs no window of its own, and Word's message
    // loop dispatches it to the callback like any other.
    if (!g_janitor)
        g_janitor = SetTimer(NULL, 0, 500, JanitorProc);

    // The DPI override is reported separately and only when it is set. It is the one setting here
    // that makes the add-in disagree with the machine on purpose, so a log that mentioned it on
    // every ordinary startup would train the reader to skip the line that matters.
    int forcedDpi = DpiOverride();
    g_dpiWatch = forcedDpi ? TRUE : FALSE;
    if (forcedDpi)
        LogWrite(L"StripStart  DPI FORCED to %d by HKCU\\Software\\WordTab\\TabDpi - "
                 L"the strip is being built at a scale this machine is not running at, and the "
                 L"janitor is watching that value for changes", forcedDpi);

    LogWrite(L"StripStart  class=%s janitor=%s  stripH=%d logical px  theme=%s  tab buttons=%s  "
             L"tab menu=%s  tab drag=%s  tab scroll=%s  tab style=%s  theme sample=%s  "
             L"title trim=%s  dot=%s  tip=%s  ghost=%s",
             g_stripClass ? L"registered" : L"FAILED",
             g_janitor ? L"running" : L"FAILED", STRIP_LOGICAL_H,
             g_palette.dark ? L"dark" : L"light",
             g_buttonsEnabled ? L"on" : L"off (HKCU\\Software\\WordTab\\TabButtons=0)",
             g_menuEnabled ? L"on" : L"off (HKCU\\Software\\WordTab\\TabMenu=0)",
             g_dragEnabled ? L"on" : L"off (HKCU\\Software\\WordTab\\TabDrag=0)",
             g_scrollEnabled ? L"on" : L"squeeze (HKCU\\Software\\WordTab\\TabScroll=0)",
             g_lookEnabled ? L"on" : L"off (HKCU\\Software\\WordTab\\TabStyle=0)",
             g_sampleEnabled ? L"on" : L"off (HKCU\\Software\\WordTab\\TabThemeSample=0)",
             g_titleTrimEnabled ? L"on" : L"off (HKCU\\Software\\WordTab\\TabTitleTrim=0)",
             g_dotEnabled ? L"on" : L"off (HKCU\\Software\\WordTab\\TabDot=0)",
             g_tipEnabled ? L"on" : L"off (HKCU\\Software\\WordTab\\TabTip=0)",
             g_ghostEnabled ? L"on" : L"off (HKCU\\Software\\WordTab\\TabGhost=0)");
}

void StripAttachFrame(HWND frame)
{
    if (!g_stripEnabled || !frame)
        return;
    if (FindByFrame(frame))
        return;
    if (g_stripCount >= MAX_STRIPS)
        return;

    StripState* state = &g_strips[g_stripCount++];
    memset(state, 0, sizeof(*state));
    state->frame   = frame;
    state->enabled = TRUE;
    ApplyMetrics(state);

    TryBind(state);
}

void StripDetachFrame(HWND frame)
{
    StripState* state = FindByFrame(frame);
    if (!state)
        return;

    // If the frame itself is on its way out there is nothing to restore and its children are gone
    // or going; touching them is at best pointless and at worst a crash.
    if (IsWindow(frame))
        Restore(state);
    else
    {
        state->enabled = FALSE;
        if (state->strip && g_dragStrip == state->strip)
            DragForget();
        if (state->strip && g_tipStrip == state->strip)
            TipStop();
        state->wwf = NULL;
        state->strip = NULL;
        state->tip = NULL;
        if (state->font) { DeleteObject(state->font); state->font = NULL; }
        ReleaseSurface(state);
    }

    int index = (int)(state - g_strips);
    for (int i = index; i < g_stripCount - 1; i++)
        g_strips[i] = g_strips[i + 1];
    g_stripCount--;

    // The entries after the removed one have shifted, so every `_WwF` subclass still carrying a
    // refData pointer into this array now points at the wrong entry. Re-point them.
    for (int i = index; i < g_stripCount; i++)
    {
        if (g_strips[i].wwf && IsWindow(g_strips[i].wwf))
        {
            SetWindowSubclass(g_strips[i].wwf, WwfSubclassProc, kWwfSubclassId,
                              (DWORD_PTR)&g_strips[i]);
        }
        if (g_strips[i].strip && IsWindow(g_strips[i].strip))
            SetWindowLongPtrW(g_strips[i].strip, GWLP_USERDATA, (LONG_PTR)&g_strips[i]);

        // The tooltip window carries the same pointer for the same reason and has to be re-pointed
        // with it. It is easy to miss because it is usually not there: a frame whose tab has never
        // been hovered has no tooltip window at all, so a version of this loop that forgot it would
        // be right on every machine until somebody hovered a tab and then closed a different
        // document.
        if (g_strips[i].tip && IsWindow(g_strips[i].tip))
            SetWindowLongPtrW(g_strips[i].tip, GWLP_USERDATA, (LONG_PTR)&g_strips[i]);
    }
}

void StripOnFrameDpiChanged(HWND frame)
{
    StripState* state = FindByFrame(frame);
    if (!state || !state->enabled)
        return;

    int dpi = DpiOf(frame);
    if (dpi == state->dpi)
        return;

    LogWrite(L"strip  hwnd=0x%p  DPI %d -> %d, re-scaling the strip", (void*)frame, state->dpi, dpi);

    // A panel measured at the old DPI is the wrong size for the new one, and nothing re-measures it
    // while it is up. Taken away rather than rebuilt: the pointer has just crossed monitors, which
    // is not a hover.
    if (state->strip && g_tipStrip == state->strip)
        TipStop();

    // One call, and it is the same one StripAttachFrame and TryBind make. Everything derived from
    // DPI - the height, the font, the size of the back buffer - is rebuilt from the new value in one
    // place, which is what stops the next thing derived from DPI being rebuilt in only two of three.
    ApplyMetrics(state);

    // The old shift was computed at the old scale, so the current rect is not a natural one. Undo
    // it, then let the next layout - or the janitor - re-apply at the new height.
    if (state->wwf && IsWindow(state->wwf) && state->hasApplied)
    {
        state->hasApplied = FALSE;
        state->stripPlaced = FALSE;
        SetWindowPos(state->wwf, NULL,
                     state->natural.left, state->natural.top,
                     state->natural.right - state->natural.left,
                     state->natural.bottom - state->natural.top,
                     SWP_NOZORDER | SWP_NOACTIVATE);
        ApplyInitial(state);
    }
}

void StripStop(void)
{
    if (g_janitor)
    {
        KillTimer(NULL, g_janitor);
        g_janitor = 0;
    }

    // The scroll position is module state, not window state, so it has to be put back by hand.
    // Switching the add-in off and on again inside one Word session would otherwise start the new
    // row scrolled to where the old one happened to be.
    DragScrollStop();
    g_scroll      = 0;
    g_scrollShown = NULL;
    g_scrollMax   = -1;

    for (int i = 0; i < g_stripCount; i++)
        Restore(&g_strips[i]);
    g_stripCount = 0;

    // The palette's two handles. This used to free nothing at all, which was harmless while they
    // were created once and leaked once at process exit - but they are re-created now every time
    // Word changes theme, and a leak per theme change is a different thing.
    if (g_backBrush)     { DeleteObject(g_backBrush);     g_backBrush = NULL; }
    if (g_menuBackBrush) { DeleteObject(g_menuBackBrush); g_menuBackBrush = NULL; }
    g_paletteReady = FALSE;
    g_openMenu = NULL;
    g_menuItemCount = 0;

    // The window class is not unregistered here: frames.cpp's FramesStop does the same for its
    // coordinator class, and the module is pinned in the process anyway.
    LogWrite(L"StripStop  done");
}
