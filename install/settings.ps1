<#
.SYNOPSIS
  Show or change WordTab's settings, without regedit.

.DESCRIPTION
  Every part of WordTab can be switched off on its own. They live as DWORD values under
  HKCU\Software\WordTab - no admin rights, no files, nothing outside the current user - and until this
  script existed the only way to reach them was regedit, which is not a reasonable thing to ask of
  somebody whose tab row has gone wrong on a Monday morning.

  Absent means on. Every switch defaults to 1, so a value only ever needs to exist in order to turn
  something OFF, and -Reset removes them all rather than writing zeros.

  **This script decides nothing.** It writes registry values and shows you what the add-in reported
  the last time it started; the add-in is the authority on what those values did, and its startup
  lines are printed at the bottom for exactly that reason. If the table here and the log disagree,
  the log is right.

  **Word reads these once, when it starts.** Change one, close Word completely, start it again.

.PARAMETER Set
  One or more NAME=VALUE pairs, e.g. -Set TabDot=0 or -Set TabDrag=0,TabScroll=0

.PARAMETER Reset
  Remove every WordTab setting, putting all of them back to their defaults.

.PARAMETER Off
  Stop WordTab loading into Word at all, leaving it installed. Word looks and behaves exactly as it
  does on a machine that never had it.

  This writes the SAME registry value Word's own tick box writes - LoadBehavior under the add-in's
  key - rather than inventing a second idea of "off". There is already one: `-Set Stack=0` leaves the
  add-in loaded and inert. Both are legitimate, and they answer different questions: Stack=0 asks
  "what does Word look like without the tabs", -Off asks "is WordTab responsible for this at all".
  The second is the one worth having when something is wrong and nobody knows what is causing it.

.PARAMETER On
  Put it back. Reports Word's Disabled Items as well, because that list beats LoadBehavior: an add-in
  Word killed after a bad load stays dead at LoadBehavior=3, and turning it "on" without saying so
  would be a lie the log would have to correct later.

.PARAMETER Quiet
  Skip the explanations and print just the table.

.PARAMETER Report
  Write everything needed to diagnose WordTab on this machine to a single text file, and print where
  it went. Send that file. It is the answer to "it is not working" from a machine nobody else can
  reach - which is the normal case, because this add-in is installed by one person on their own
  work computer.

  It reads only. Nothing is installed, changed or removed, and no document is opened or touched.

.PARAMETER ReportPath
  Where to write it. Defaults to wordtab-report-<date>.txt on your Desktop, falling back to the
  current directory if there is no Desktop.

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File install\settings.ps1
  powershell -ExecutionPolicy Bypass -File install\settings.ps1 -Set TabDrag=0
  powershell -ExecutionPolicy Bypass -File install\settings.ps1 -Reset
  powershell -ExecutionPolicy Bypass -File install\settings.ps1 -Report
  powershell -ExecutionPolicy Bypass -File install\settings.ps1 -Off
  powershell -ExecutionPolicy Bypass -File install\settings.ps1 -On
#>
[CmdletBinding()]
param(
    [string[]]$Set,
    [switch]$Reset,
    [switch]$Quiet,
    [switch]$Report,
    [string]$ReportPath,
    [switch]$Off,
    [switch]$On
)

$ErrorActionPreference = 'Stop'
$Key = 'HKCU:\Software\WordTab'

# The registry questions install.ps1 asks too. See common.ps1 for why they are not asked twice.
$CommonFile = Join-Path $PSScriptRoot 'common.ps1'
if (-not (Test-Path $CommonFile)) {
    throw "install\common.ps1 is missing. It sits beside this script in both the repo and a package; copy the whole folder rather than settings.ps1 on its own."
}
. $CommonFile

# What each switch does, in the terms somebody turning it off would think in. The names and defaults
# are the add-in's, not this script's - see the note in the help above about which one is the
# authority. Order is roughly "biggest hammer first".
$Switches = @(
    @{ Name = 'Stack';          Does = 'Put several documents in one window at all. Off, WordTab does nothing visible.' }
    @{ Name = 'Strip';          Does = 'Draw the tab row. Off, no strip is created and Word looks untouched.' }
    @{ Name = 'Taskbar';        Does = 'Show one taskbar button for the stack instead of one per document.' }
    @{ Name = 'AltTab';         Does = 'Show one Alt+Tab entry for the stack instead of one per document.' }
    @{ Name = 'TabStyle';       Does = 'Draw tabs as rounded cards in Word''s colours. Off, plain rectangles.' }
    @{ Name = 'TabThemeSample'; Does = 'Take the colours off Word''s own ribbon. Off, use the Office theme setting instead. TRY THIS FIRST if the strip is the wrong colour.' }
    @{ Name = 'TabButtons';     Does = 'The x on each tab and the + at the end.' }
    @{ Name = 'TabMenu';        Does = 'The right-click menu on a tab.' }
    @{ Name = 'TabDrag';        Does = 'Dragging a tab - to reorder the row, to pull a document out into its own window, or to put one back.' }
    @{ Name = 'TabTearOff';     Does = 'Taking a document out of the stack into a window of its own. Off, neither the menu item nor the drag is offered.' }
    @{ Name = 'TabCloseStack';  Does = 'The window''s own close button (and Alt+F4) asks whether to close all the tabs or only the document in front. Off, it closes only the document in front with no question, which is what Word does on its own. Set it to 2 to close them all without asking.' }
    @{ Name = 'TabScroll';      Does = 'Scroll the row when there are more tabs than fit. Off, every tab is shown and they get narrower instead.' }
    @{ Name = 'TabTitleTrim';   Does = 'Trim "  -  Compatibility Mode" and the like off a tab''s name.' }
    @{ Name = 'TabDot';         Does = 'Show a dot instead of the x on a document with unsaved changes. Off, WordTab never asks Word about your documents.' }
    @{ Name = 'TabKeys';        Does = 'Ctrl+Tab and Ctrl+Shift+Tab step along the tab row. TURN THIS OFF if you need Ctrl+Tab to type a tab inside a table.' }
    @{ Name = 'TabTip';         Does = 'Resting the pointer on a tab shows the document''s full name and the folder it is in. Off, WordTab never asks Word where your documents live.' }
    @{ Name = 'TabGhost';       Does = 'While a tab is dragged clear of the row, carry a picture of it under the pointer. Off, the row letting go and the cursor are the only feedback.' }
    @{ Name = 'RowSize';        Does = 'The row comes back the size you last gave it, instead of whatever size Word restores its window to. Until you have given it one, a window Word opens three or more pages wide is narrowed to about one page. Off, Word decides - which on a wide screen is what puts several pages side by side.' }
    @{ Name = 'OnePage';        Does = 'A document that opens showing several pages side by side, or at a zoom too small to read, is put back to one page at 100%. Off, Word''s remembered column count and zoom stand.' }
    @{ Name = 'ShowLoadBanner'; Does = 'A dialog at Word startup confirming WordTab loaded. Off by default from the installer.' }
)

function Get-Current($name) {
    if (-not (Test-Path $Key)) { return $null }
    $p = Get-ItemProperty -Path $Key -Name $name -ErrorAction SilentlyContinue
    if ($p -and ($p.PSObject.Properties.Name -contains $name)) { return $p.$name }
    return $null
}

# A window left edge can be negative - the row on the rig this was written for sits at x=-10 - and
# a REG_DWORD carries that as two complement. What comes back here is the unsigned number, so the
# report said: Row remembered (4294967286,673 -4294965991x1419). Three of those four numbers are
# unreadable, and this is the one line in the report that says why the row is the size it is.
#
# Reinterpreted rather than clamped or hidden. The value in the registry is correct - the add-in
# reads it back as a LONG and the row comes up in the right place - so the only thing that was ever
# wrong is the reading of it here.
function Get-Signed32($value) {
    if ($null -eq $value) { return $null }
    $n = [int64]$value
    if ($n -gt 2147483647) { return $n - 4294967296 }
    return $n
}

# ---- -Report ------------------------------------------------------------------------------------
#
# One file, everything, no questions asked of the person running it.
#
# The shape is deliberate. Somebody whose tab row has gone wrong is not going to be walked through
# eight registry paths over email, and the answers that matter are exactly the ones they cannot be
# expected to know they need: which build this is, whether Word disabled it, whether something else
# is driving the same document frame, what DPI the screen is at, and what the add-in itself said
# when Word last started.
#
# **The add-in's own log is the most valuable thing in here and it goes in whole, at the end.** Every
# diagnosis this project has made from outside the machine came off those lines.

if ($Report) {
    if (-not $ReportPath) {
        $desktop = [Environment]::GetFolderPath('Desktop')
        if (-not $desktop -or -not (Test-Path $desktop)) { $desktop = (Get-Location).Path }
        $ReportPath = Join-Path $desktop ("wordtab-report-{0}.txt" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
    }

    $out = New-Object System.Collections.Generic.List[string]
    function Say($text) { $out.Add([string]$text) }
    function Head($text) { Say ''; Say $text; Say ('-' * $text.Length) }

    Say "WordTab report"
    Say ("written {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss K'))

    Head 'This machine'
    Say ("OS              : {0}" -f [Environment]::OSVersion.VersionString)
    Say ("64-bit OS       : {0}" -f [Environment]::Is64BitOperatingSystem)
    Say ("PowerShell      : {0}  ({1}-bit host)" -f $PSVersionTable.PSVersion, $(if ([Environment]::Is64BitProcess) { 64 } else { 32 }))
    Say ("User            : {0}" -f [Environment]::UserName)
    # DPI IS THE FIRST THING TO ASK ABOUT A MACHINE NOBODY HERE HAS SEEN.
    #
    # Every size the tab row draws - the height of the strip, how wide a tab is, the radius of its
    # corners, how far a tab has to be dragged before it comes out of the row - is a logical number
    # scaled by the DPI of the window it is drawn in. The development rig runs at 200%. A laptop
    # usually does not. "The tabs look wrong" and "the tabs look wrong AT 125%" are different reports
    # and only one of them can be acted on.
    # EVERY monitor, with its own DPI, not just the primary one.
    #
    # A laptop with an external monitor is the normal setup and the two are usually at DIFFERENT
    # scaling - and that is the case this add-in has never been able to test, because the machine it
    # was developed on has one screen. Dragging Word from one to the other fires WM_DPICHANGED and
    # every size in the tab row is rebuilt from the new number. **If the tabs look wrong on one
    # screen and right on the other, this section is the first thing to read.**
    try {
        Add-Type -TypeDefinition @'
using System; using System.Collections.Generic; using System.Runtime.InteropServices; using System.Text;
public static class WordTabDpi {
  delegate bool MonEnumProc(IntPtr mon, IntPtr dc, IntPtr rect, IntPtr data);
  [DllImport("user32.dll")] static extern bool EnumDisplayMonitors(IntPtr dc, IntPtr clip, MonEnumProc cb, IntPtr data);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern bool GetMonitorInfoW(IntPtr mon, ref MONITORINFOEX info);
  [DllImport("shcore.dll")] static extern int GetDpiForMonitor(IntPtr mon, int type, out uint x, out uint y);
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] static extern bool SetProcessDpiAwarenessContext(IntPtr ctx);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
  [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)]
  public struct MONITORINFOEX {
    public int cbSize; public RECT rcMonitor; public RECT rcWork; public uint dwFlags;
    [MarshalAs(UnmanagedType.ByValTStr, SizeConst=32)] public string szDevice;
  }
  public static string[] All() {
    // PER-MONITOR aware, not merely SetProcessDPIAware(), and the difference is this whole section.
    //
    // What GetDpiForMonitor reports is gated by the DPI awareness of the CALLING process.
    // Unaware -> 96 for everything. SYSTEM aware -> the system DPI for EVERY monitor. Only
    // per-monitor awareness gets each monitor's own number. So SetProcessDPIAware() - system
    // awareness - would have printed the SAME dpi against both screens of a two-monitor rig and the
    // "different scaling" line below could never have fired. That is the exact question this
    // section exists to answer, on the exact machine it exists for.
    //
    // WHAT WAS MEASURED HERE AND WHAT WAS NOT, because this rig has one screen and cannot show the
    // whole thing. MEASURED: the gating is real - the same monitor reports 96 from Windows
    // PowerShell 5.1 (DPI-unaware) and 192 from pwsh 7 (per-monitor aware in its manifest), and 5.1
    // is what the target machine has. NOT MEASURED HERE, and it is documented behaviour rather than
    // something this rig can demonstrate: that a SYSTEM-aware process gets the system DPI for every
    // monitor rather than each one's own. On one screen the two are the same number, so the old call
    // looked correct here and could only ever have been wrong on the machine it was written for.
    // **If you are reading this on a two-monitor rig, that is the claim to check.**
    //
    // PER_MONITOR_AWARE_V2 is (IntPtr)-4. It is Windows 10 1703+; SetProcessDPIAware is the fallback
    // for older, where there is one scaling for everything anyway and system awareness is right.
    // Either call fails harmlessly if the host already declared an awareness.
    try { if (!SetProcessDpiAwarenessContext(new IntPtr(-4))) SetProcessDPIAware(); }
    catch { try { SetProcessDPIAware(); } catch {} }

    List<string> lines = new List<string>();
    EnumDisplayMonitors(IntPtr.Zero, IntPtr.Zero, delegate(IntPtr mon, IntPtr dc, IntPtr r, IntPtr d) {
      MONITORINFOEX mi = new MONITORINFOEX();
      mi.cbSize = Marshal.SizeOf(typeof(MONITORINFOEX));
      string dpiText = "dpi unknown";
      uint x, y;
      // GetDpiForMonitor is Windows 8.1. Older than that has one scaling for everything anyway.
      try { if (GetDpiForMonitor(mon, 0 /* MDT_EFFECTIVE_DPI */, out x, out y) == 0)
              dpiText = string.Format("dpi {0} ({1}% scaling)", x, (int)(100 * x / 96)); }
      catch {}
      if (GetMonitorInfoW(mon, ref mi)) {
        lines.Add(string.Format("{0,-14} {1,5}x{2,-5} at ({3},{4})  {5}{6}",
          mi.szDevice,
          mi.rcMonitor.Right - mi.rcMonitor.Left, mi.rcMonitor.Bottom - mi.rcMonitor.Top,
          mi.rcMonitor.Left, mi.rcMonitor.Top,
          dpiText, ((mi.dwFlags & 1) != 0 ? "  PRIMARY" : "")));
      }
      return true;
    }, IntPtr.Zero);
    return lines.ToArray();
  }
}
'@ -ErrorAction Stop
        $monitors = @([WordTabDpi]::All())
        Say ("Monitors        : {0}" -f $monitors.Count)
        foreach ($m in $monitors) { Say ("  {0}" -f $m) }
        # Named explicitly rather than left to be spotted. Mixed scaling is the configuration most
        # likely to show up a bug in code that was only ever run against one screen.
        $scalings = @($monitors | ForEach-Object { if ($_ -match 'dpi (\d+)') { $Matches[1] } } | Sort-Object -Unique)
        if ($scalings.Count -gt 1) {
            Say "  ** These monitors are at DIFFERENT scaling. Say which screen Word was on."
        }
    } catch { Say "Monitors        : (could not be read: $($_.Exception.Message))" }

    # Physical pixels, and that is worth saying rather than leaving to be worked out. The DPI block
    # above calls SetProcessDPIAware, so these are the real numbers - the same ones the add-in works
    # in. Without it Windows would report a 2548-wide screen as 1274 at 200% scaling, and a rectangle
    # that disagrees with the log by exactly a factor of two is a confusing thing to be sent.
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        foreach ($screen in [System.Windows.Forms.Screen]::AllScreens) {
            Say ("Screen          : {0}  {1}  primary={2}  (physical pixels)" -f $screen.DeviceName, $screen.Bounds, $screen.Primary)
        }
    } catch { Say "Screen          : (could not be read: $($_.Exception.Message))" }

    Head 'Word'
    $wordExe = $null
    try { $wordExe = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\winword.exe' -ErrorAction Stop).'(default)' } catch {}
    if ($wordExe -and (Test-Path $wordExe)) {
        $v = (Get-Item $wordExe).VersionInfo
        Say ("Executable      : {0}" -f $wordExe)
        Say ("Version         : {0}" -f $v.FileVersion)
    } else {
        Say "Executable      : NOT FOUND via App Paths - Word may not be installed for this user"
    }
    $c2r = $null
    try { $c2r = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction Stop } catch {}
    if ($c2r) {
        Say ("Click-to-Run    : {0} v{1} {2}" -f $c2r.ProductReleaseIds, $c2r.VersionToReport, $c2r.Platform)
    } else {
        Say "Click-to-Run    : no ClickToRun key (an MSI install, or Office is absent)"
    }
    $running = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue)
    Say ("Running now     : {0} WINWORD process(es)" -f $running.Count)

    Head 'WordTab, as installed'
    $ourDll = Join-Path $env:LOCALAPPDATA 'Programs\WordTab\WordTab.dll'
    if (Test-Path $ourDll) {
        $f = Get-Item $ourDll
        Say ("DLL             : {0}" -f $ourDll)
        Say ("Size            : {0:N0} bytes" -f $f.Length)
        Say ("Modified        : {0:yyyy-MM-dd HH:mm:ss}" -f $f.LastWriteTime)
        Say ("SHA256          : {0}" -f (Get-FileHash $ourDll -Algorithm SHA256).Hash)
        # A mark-of-the-web on the installed copy is a policy refusal waiting to happen, and it looks
        # exactly like a bad registration from the outside.
        $zone = Get-Content -Path $ourDll -Stream Zone.Identifier -ErrorAction SilentlyContinue
        Say ("Mark-of-the-web : {0}" -f $(if ($zone) { 'PRESENT - this can stop Word loading it' } else { 'none' }))
    } else {
        Say "DLL             : NOT PRESENT at $ourDll  - it is not installed, or it was installed elsewhere"
    }
    $payload = Join-Path $PSScriptRoot '..\PAYLOAD.txt'
    if (Test-Path $payload) {
        Say 'Package         :'
        foreach ($line in (Get-Content $payload | Where-Object { $_ -match '^(Commit|Built|Sha256|Package)' })) { Say ("  {0}" -f $line.Trim()) }
    }
    # Can this machine get rid of WordTab? Two separate answers, and either can be no on its own. The
    # uninstaller is a file install.ps1 leaves beside the DLL; the listing is the registry entry that
    # points Settings > Apps at it. A missing file with a present listing is the worse of the two -
    # clicking Uninstall would fail with a path error - so name them apart rather than together.
    $ourUninstaller = Join-Path $env:LOCALAPPDATA 'Programs\WordTab\uninstall.ps1'
    Say ("Uninstaller     : {0}" -f $(if (Test-Path $ourUninstaller) { $ourUninstaller } else { "NOT PRESENT at $ourUninstaller" }))
    $arp = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\WordTab'
    if (Test-Path $arp) {
        $a = Get-ItemProperty -Path $arp
        Say ("Listed in Apps  : yes, as '{0}' version {1}" -f $a.DisplayName, $a.DisplayVersion)
        Say ("  Uninstall runs: {0}" -f $a.UninstallString)
    } else {
        Say "Listed in Apps  : no - Settings > Apps will not show WordTab. Installed before this entry existed, or removed by hand."
    }

    Head 'Registration'
    $clsidKey = 'HKCU:\Software\Classes\CLSID\{4BF75ED9-10EE-4866-BF4A-3D663A4149A1}\InprocServer32'
    if (Test-Path $clsidKey) {
        Say ("COM class       : {0}" -f (Get-ItemProperty $clsidKey).'(default)')
        Say ("Threading       : {0}" -f (Get-ItemProperty $clsidKey).ThreadingModel)
    } else {
        Say "COM class       : NOT REGISTERED under HKCU - Word cannot create it"
    }
    $ours = Get-AddinStatus 'WordTab.Connect'
    if ($null -eq $ours.Behavior) {
        # The case this report exists for. A null Behavior printed through the format string below
        # came out as "LoadBehavior= from " - an empty answer to the most important question here.
        Say  "Add-in entry    : NOT REGISTERED - Word has never been told this add-in exists"
        Say  "                  HKCU\Software\Microsoft\Office\Word\Addins\WordTab.Connect is absent."
        Say  "                  Re-run install\install.ps1."
    } else {
        Say ("Add-in entry    : LoadBehavior={0} from {1}" -f $ours.Behavior, $ours.Hive)
    }
    if ($ours.Behavior -eq 2) { Say "                  ** 2 means Word TRIED to load it and gave up. The log below says why." }
    if ($ours.Behavior -eq 0) { Say "                  ** 0 means it is switched off in File > Options > Add-ins." }
    Say ("Disabled by Word: {0}" -f $(if ($ours.Blocked) { 'YES - Disabled Items beats LoadBehavior. File > Options > Add-ins > Manage: Disabled Items > Go.' } else { 'no' }))

    Head 'Word''s windows, right now'
    #
    # **Run the report WHILE the thing that looks wrong is on screen and this section is the answer.**
    # Everything else in this file is about what is installed and configured; this is the only part
    # that says what Word is actually doing. Three questions, in the order they rule things out:
    # is the add-in even in there (is there a WordTabStrip at all), is it in EVERY window or only
    # some, and is the strip where it should be - immediately above the document frame, spanning it.
    #
    # A stack is several OpusApp windows at ONE rectangle. Several rectangles means the stacking
    # never happened, which is a completely different fault from a strip that is drawn in the wrong
    # place.
    if ($running.Count -eq 0) {
        Say '(Word is not running, so there is nothing to look at. If something looks wrong, start'
        Say ' Word, open two documents, leave it on screen and run this again.)'
    } else {
        try {
            Add-Type -TypeDefinition @'
using System; using System.Collections.Generic; using System.Runtime.InteropServices; using System.Text;
public static class WordTabWin {
  delegate bool EnumProc(IntPtr h, IntPtr p);
  [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr p);
  [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr h, EnumProc cb, IntPtr p);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetClassNameW(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowTextW(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
  static string ClassOf(IntPtr h) { StringBuilder s = new StringBuilder(64); GetClassNameW(h, s, 64); return s.ToString(); }
  static string TextOf(IntPtr h) { StringBuilder s = new StringBuilder(300); GetWindowTextW(h, s, 300); return s.ToString(); }
  static string RectOf(IntPtr h) {
    RECT r; if (!GetWindowRect(h, out r)) return "(no rect)";
    return string.Format("({0},{1} {2}x{3})", r.Left, r.Top, r.Right - r.Left, r.Bottom - r.Top);
  }
  public static string[] Describe() {
    List<string> lines = new List<string>();
    List<IntPtr> frames = new List<IntPtr>();
    EnumWindows(delegate(IntPtr h, IntPtr p) {
      if (ClassOf(h) == "OpusApp" && IsWindowVisible(h)) frames.Add(h);
      return true;
    }, IntPtr.Zero);
    lines.Add(string.Format("{0} visible OpusApp window(s)", frames.Count));
    foreach (IntPtr f in frames) {
      uint pid; GetWindowThreadProcessId(f, out pid);
      lines.Add("");
      lines.Add(string.Format("  frame 0x{0:X}  pid={1}  {2}", f.ToInt64(), pid, RectOf(f)));
      lines.Add(string.Format("    title: {0}", TextOf(f)));
      string strip = null, wwf = null;
      IntPtr ff = f;
      EnumChildWindows(f, delegate(IntPtr c, IntPtr p) {
        string cls = ClassOf(c);
        if (cls == "WordTabStrip" && strip == null) strip = RectOf(c) + (IsWindowVisible(c) ? " visible" : " HIDDEN");
        if (cls == "_WwF" && wwf == null) wwf = RectOf(c);
        return true;
      }, IntPtr.Zero);
      lines.Add(string.Format("    WordTabStrip: {0}", strip == null ? "ABSENT - the add-in has not attached to this window" : strip));
      lines.Add(string.Format("    _WwF document frame: {0}", wwf == null ? "absent" : wwf));
    }
    return lines.ToArray();
  }
}
'@ -ErrorAction Stop
            foreach ($line in [WordTabWin]::Describe()) { Say $line }
        } catch { Say "(could not be read: $($_.Exception.Message))" }
    }

    # HOW MANY PAGES ACROSS EACH DOCUMENT IS ACTUALLY SHOWING, AND AT WHAT ZOOM.
    #
    # The complaint this add-in was built to answer is documents opening several pages side by side
    # at 10% zoom, and the add-in corrects it on every join - but everything above this line is Win32,
    # so no report has ever carried the two numbers the correction is about. Six OnePage runs in the
    # 2026-09-04 report said nothing, and this section is the half of that answer that needs no new
    # build to reach the machine: it reads what Word is showing RIGHT NOW.
    #
    # 99 columns is not ninety-nine pages. It is Word's "as many as fit", and it is the HEALTHY value
    # - the thing to look for is a small number greater than one, with a crushed zoom beside it.
    #
    # Read through Word's object model, which is a different instrument from the window enumeration
    # above and fails in a different way: a Protected View document is in none of this Application's
    # collections at all. So the two counts are printed together deliberately - OpusApp windows seen
    # by Win32 against windows Word admits to - because a gap between them IS the Protected View
    # signature and has cost this project a diagnosis before.
    #
    # Nothing here can start Word, change a document or throw: GetActiveObject attaches only to an
    # instance that is already running, every property read is individually guarded, and the whole
    # block is wrapped. It is also the only part of this report that talks to Word rather than about
    # it, so it is last among the live sections.
    if ($running.Count -gt 0) {
        Head 'What each document is showing (pages across, and zoom)'
        try {
            # GetActiveObject is .NET Framework only - it is not implemented in .NET Core, so this
            # says so rather than throwing when the report is run under pwsh 7 instead of the
            # Windows PowerShell the .cmd uses.
            $word = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')

            $omCount = 0
            try { $omCount = $word.Windows.Count } catch { }
            Say ("Word's object model admits to {0} window(s)." -f $omCount)
            Say ''

            for ($i = 1; $i -le $omCount; $i++) {
                $caption = '(unreadable)'; $hwnd = '?'; $cols = '?'; $pct = '?'
                try { $w = $word.Windows.Item($i) } catch { $w = $null }
                if ($null -eq $w) { Say ("  window {0}: could not be read" -f $i); continue }

                try { $caption = $w.Caption } catch { }
                try { $hwnd = '0x{0:X}' -f [int]$w.Hwnd } catch { }
                # Each read separately: a window can give up its Zoom object and then refuse the
                # properties on it, which is a different fault from having no View at all.
                try { $cols = $w.View.Zoom.PageColumns } catch { }
                try { $pct  = $w.View.Zoom.Percentage } catch { }

                Say ("  {0,-12} {1} pages across at {2}%   {3}" -f $hwnd, $cols, $pct, $caption)
                if ($cols -is [int] -and $cols -gt 1 -and $cols -lt 99) {
                    Say '     ^ THIS IS THE FAULT: more than one page across, and not Word''s "as many as fit".'
                }
            }

            if ($omCount -lt $running.Count) {
                Say ''
                Say ("Note: {0} Word process(es) are running but the object model lists {1} window(s)." -f $running.Count, $omCount)
                Say 'A document Word will not list is usually one in Protected View, which is in none'
                Say 'of these collections and which the add-in cannot read or correct either.'
            }
        } catch {
            Say "(Word's object model could not be reached: $($_.Exception.Message))"
            Say 'That is not necessarily a fault - it is also what a busy Word, a modal dialog, or'
            Say 'running this report under PowerShell 7 rather than Windows PowerShell looks like.'
        }
    }

    Head 'Everything in Word''s Disabled Items'
    $disabled = @(Get-DisabledAddinPaths)
    if ($disabled.Count -eq 0) { Say '(empty)' } else { foreach ($d in $disabled) { Say "  $d" } }

    Head 'Other tabbed-Word add-ins'
    foreach ($r in (Get-RivalAddins)) {
        Say ("{0,-28} LoadBehavior={1,-4} hive={2,-6} blocked={3}" -f $r.ProgId, $r.Behavior, $r.Hive, $r.Blocked)
        if ($r.Dll) { Say ("  {0}" -f $r.Dll) }
    }
    Say ''
    Say 'WordTab has never been tested beside a WORKING one - they carve up the same document frame.'
    Say 'If any of the above shows LoadBehavior=3 with blocked=False, turn it off before judging WordTab.'

    Head 'Settings'
    foreach ($s in $Switches) {
        $current = Get-Current $s.Name
        Say ("{0,-16} {1}" -f $s.Name, $(if ($null -eq $current) { 'on (default, no value set)' } elseif ($current -eq 0) { 'OFF' } else { "on ($current)" }))
    }

    # The row's remembered rectangle, decoded. It is four separate DWORDs under the key and the
    # section below would print them raw, which is four numbers nobody can read as a window - and
    # this is the report that comes back from a machine nobody here can reach, so the one value
    # that says why the row is the size it is should be legible on sight.
    $rowL = Get-Signed32 (Get-Current 'RowLeft')
    $rowT = Get-Signed32 (Get-Current 'RowTop')
    $rowR = Get-Signed32 (Get-Current 'RowRight')
    $rowB = Get-Signed32 (Get-Current 'RowBottom')
    if (($null -ne $rowL) -and ($null -ne $rowT) -and ($null -ne $rowR) -and ($null -ne $rowB)) {
        Say ("{0,-16} ({1},{2} {3}x{4}){5}" -f 'Row remembered', $rowL, $rowT,
                                              ($rowR - $rowL), ($rowB - $rowT),
                                              $(if (Get-Current 'RowMaximized') { ' maximized' } else { '' }))
    } else {
        Say ("{0,-16} nothing yet - so a window Word opens three or more pages wide is" -f 'Row remembered')
        Say ("{0,-16} narrowed to about one page. Size the window and that wins instead." -f '')
    }

    # Anything under the key that is NOT one of the switches above. The list is a list of things a
    # user is offered, not a list of everything the add-in reads - TabDpi is read and deliberately
    # not offered - so a report that printed only the list would be silent about exactly the values
    # nobody expects to find set. This is the report that goes to a machine nobody here can reach,
    # and a strip built at the wrong scale must not be missing from it.
    if (Test-Path $Key) {
        $known = $Switches | ForEach-Object { $_.Name }
        # @() around the WHOLE pipeline. Filtering to exactly one name yields a bare string, and
        # `.Count` on a bare string is 1 in some hosts, $null in others and a hard error under
        # StrictMode - so the untouched version would have skipped this section silently in the one
        # case it exists for, a single non-switch value called TabDpi. This report is read on a
        # machine nobody here can reach, under whatever PowerShell it happens to have.
        $extra = @(@(Get-Item $Key).GetValueNames() | Where-Object { $_ -and ($known -notcontains $_) })
        if ($extra.Count -gt 0) {
            Say ''
            Say 'Values set here that are NOT ordinary switches:'
            foreach ($name in $extra) {
                Say ("{0,-16} {1}" -f $name, (Get-ItemProperty -Path $Key -Name $name).$name)
            }
            if ($extra -contains 'TabDpi') {
                Say ''
                Say 'TabDpi is set. The add-in is building the tab row at that scale INSTEAD of the'
                Say 'one this machine is running at. It is a test and diagnostic override, not a'
                Say 'setting - if you did not set it deliberately, remove it and restart Word:'
                Say '  Remove-ItemProperty HKCU:\Software\WordTab TabDpi'
            }
        }
    }

    Head 'The add-in''s log, in full'
    $logFile = Join-Path $env:LOCALAPPDATA 'WordTab\wordtab.log'
    if (Test-Path $logFile) {
        $log = @(Get-Content $logFile -ErrorAction SilentlyContinue)
        Say ("{0} ({1:N0} lines, {2:N0} bytes)" -f $logFile, $log.Count, (Get-Item $logFile).Length)
        Say ''

        # WHERE THIS FILE STARTS, SAID OUT LOUD, BECAUSE IT IS NOT ALWAYS THE BEGINNING.
        #
        # The add-in deletes its log whole once it passes 512KB and starts a new one with the very
        # next line it writes - src\native\log.cpp, RollIfLarge. Nothing marks the seam: the new file
        # opens on an ordinary mid-session line with an ordinary timestamp, so a rolled log and a
        # fresh install with a short history read exactly alike from here. Several diagnoses this
        # project has made turned on "the log says nothing about X", which is a different statement
        # when the lines that would have said it were deleted.
        #
        # A file that has never rolled opens with the '---- DllGetClassObject ----' line the add-in
        # writes the instant Word loads it: 40 of the 40 logs kept on the development rig begin with
        # it. Printed verbatim rather than reduced to a verdict, because this is read about a machine
        # nobody here can reach and the first line may be one this script has never seen.
        if ($log.Count -gt 0) {
            Say 'First line in the file - this is how you tell a rolled log from a whole one:'
            Say ("  {0}" -f $log[0])
            if ($log[0] -notlike '*DllGetClassObject*') {
                Say '  ^ that is not a Word-startup line, so this log HAS rolled: it passed 512KB and'
                Say '    everything written before it was deleted. It is not the whole history.'
            }
            Say ''
        }

        foreach ($line in $log) { Say $line }
    } else {
        Say "$logFile does not exist."
        Say 'Nothing has been written there, which means Word has never loaded the add-in at all.'
        Say 'Check the registration above before anything else.'
    }

    # Not Set-Content -Encoding UTF8: under Windows PowerShell 5.1 that means UTF-8 WITH a BOM and
    # under pwsh 7 it means without, so the same script would produce two different files depending
    # on which shell the machine happened to have. Driven, not assumed - the two were diffed.
    [System.IO.File]::WriteAllLines($ReportPath, $out, (New-Object System.Text.UTF8Encoding($false)))
    Write-Host ''
    Write-Host 'WordTab report written to:' -ForegroundColor Cyan
    Write-Host "  $ReportPath" -ForegroundColor Green
    Write-Host ''
    Write-Host 'It reads only - nothing was installed, changed or removed. Send that file.' -ForegroundColor Gray
    Write-Host 'It contains the names of documents you have had open (they appear in the log as tab names).' -ForegroundColor Yellow
    Write-Host ''
    return
}

# ---- -Off and -On -------------------------------------------------------------------------------
#
# The whole add-in, in or out, without uninstalling it. Everything else in this file turns a FEATURE
# off inside a WordTab that is still loading; this is the one that answers "is WordTab doing this at
# all", which is the question somebody actually has when Word starts behaving oddly and there is a
# new add-in on the machine.
#
# LoadBehavior, not a WordTab value of our own, because Word already owns this decision - the tick box
# in File > Options > Add-ins > COM Add-ins writes the same value name under the same key. A private
# WordTabEnabled flag would have been a second answer to a question that already has one, and the two
# would agree right up until somebody used the dialog.
#
# 0 rather than 2: 0 is Disconnected outright, it is the number install.ps1 already prints when it
# tells somebody how to stand a rival add-in down, and Word writes 2 itself when a load FAILS - so
# leaving 2 to mean "it broke" and 0 to mean "somebody asked" keeps the two apart in a log.

if ($Off -and $On) {
    throw "-Off and -On cannot both be given. Pick one."
}

if ($Off -or $On) {
    $ProgId   = 'WordTab.Connect'
    $AddinKey = "HKCU:\Software\Microsoft\Office\Word\Addins\$ProgId"

    if (-not (Test-Path $AddinKey)) {
        throw "WordTab is not registered for your account, so there is nothing to turn on or off. Run install.ps1 (or the one-file installer) first."
    }

    $want = if ($On) { 3 } else { 0 }
    Set-ItemProperty -Path $AddinKey -Name 'LoadBehavior' -Value $want -Type DWord

    # Read back through the same function install.ps1 and -Report use, rather than trusting the write.
    # This is the file that tells somebody whether WordTab is on; it should not be the one place that
    # believes its own last statement.
    $now = Get-AddinStatus $ProgId

    Write-Host ''
    if ($On) {
        Write-Host 'WordTab is ON.' -ForegroundColor Green
        Write-Host "  $AddinKey  LoadBehavior=$($now.Behavior)" -ForegroundColor DarkGray

        # The one case where LoadBehavior=3 is not the whole story, and it is not hypothetical - it is
        # the state the dev rig's Office Tab has been stuck in since somebody's bad load.
        if ($now.Blocked) {
            Write-Host ''
            Write-Host '  But Word has this add-in in its Disabled Items list, which beats LoadBehavior.' -ForegroundColor Yellow
            Write-Host '  It will not load until that is cleared:' -ForegroundColor Yellow
            Write-Host '    Word > File > Options > Add-ins > Manage: Disabled Items > Go > select WordTab > Enable' -ForegroundColor Gray
        }
    } else {
        Write-Host 'WordTab is OFF.' -ForegroundColor Yellow
        Write-Host "  $AddinKey  LoadBehavior=$($now.Behavior)" -ForegroundColor DarkGray
        Write-Host '  It is still installed. Nothing was removed, and no setting was changed.' -ForegroundColor Gray
        Write-Host '  Turn it back on with -On, or by ticking WordTab in Word:' -ForegroundColor Gray
        Write-Host '    File > Options > Add-ins > Manage: COM Add-ins > Go' -ForegroundColor Gray
    }
    Write-Host ''
    Write-Host 'Close Word completely and start it again for this to take effect.' -ForegroundColor Yellow
    Write-Host ''

    # Deliberately the end of the run. The table below is about which PARTS of a loaded WordTab are on,
    # and printing it under the words "WordTab is OFF" would read as a contradiction.
    return
}

# ---- -Reset -------------------------------------------------------------------------------------

if ($Reset) {
    if (Test-Path $Key) {
        # Every value under the key, not just the switches offered above. "Every WordTab setting
        # removed" has to be true when it is printed, and the switch list is what a user is offered
        # rather than everything the add-in reads - TabDpi is read and deliberately not offered. A
        # reset that left it behind would leave the add-in building the row at a scale the machine is
        # not running at, having just said it had put everything back.
        $known = $Switches | ForEach-Object { $_.Name }
        $extra = @(@(Get-Item $Key).GetValueNames() | Where-Object { $_ -and ($known -notcontains $_) })

        foreach ($s in $Switches) {
            Remove-ItemProperty -Path $Key -Name $s.Name -ErrorAction SilentlyContinue
        }
        foreach ($name in $extra) {
            Remove-ItemProperty -Path $Key -Name $name -ErrorAction SilentlyContinue
        }

        Write-Host 'Every WordTab setting removed - all of them are back to their defaults.' -ForegroundColor Green
        if ($extra.Count -gt 0) {
            Write-Host ("Including {0} value(s) that are not ordinary switches: {1}" -f $extra.Count, ($extra -join ', ')) -ForegroundColor Yellow
        }
    } else {
        Write-Host 'There were no WordTab settings to remove; everything is already at its default.' -ForegroundColor Green
    }
    Write-Host 'Close Word completely and start it again for this to take effect.' -ForegroundColor Yellow
    Write-Host ''
}

# ---- -Set ---------------------------------------------------------------------------------------

if ($Set) {
    if (-not (Test-Path $Key)) { New-Item -Path $Key -Force | Out-Null }

    foreach ($pair in $Set) {
        $bits = $pair -split '=', 2
        if ($bits.Count -ne 2) {
            Write-Host "Not a NAME=VALUE pair: `"$pair`"" -ForegroundColor Red
            continue
        }

        $name  = $bits[0].Trim()
        $value = $bits[1].Trim()

        # Matched against the known list rather than written blindly. A typo would otherwise become a
        # registry value that looks like a setting, is read by nothing, and is indistinguishable from
        # a switch that does not work.
        $known = @($Switches | Where-Object { $_.Name -ieq $name } | Select-Object -First 1)
        if (-not $known) {
            Write-Host "There is no WordTab setting called `"$name`". Run this script with no arguments to see the list." -ForegroundColor Red
            continue
        }

        $number = 0
        if (-not [int]::TryParse($value, [ref]$number)) {
            Write-Host "`"$value`" is not a number. These are 1 for on and 0 for off." -ForegroundColor Red
            continue
        }

        Set-ItemProperty -Path $Key -Name $known.Name -Value $number -Type DWord
        Write-Host ("{0} = {1}" -f $known.Name, $number) -ForegroundColor Green
    }
    Write-Host 'Close Word completely and start it again for this to take effect.' -ForegroundColor Yellow
    Write-Host ''
}

# ---- what things are now -------------------------------------------------------------------------

Write-Host "WordTab settings  ($Key)" -ForegroundColor Cyan
Write-Host ''
foreach ($s in $Switches) {
    $current = Get-Current $s.Name
    $shown   = if ($null -eq $current) { 'on (default)' } elseif ($current -eq 0) { 'OFF' } else { "on ($current)" }
    $colour  = if ($null -ne $current -and $current -eq 0) { 'Yellow' } else { 'Gray' }
    Write-Host ("  {0,-16} {1}" -f $s.Name, $shown) -ForegroundColor $colour
    if (-not $Quiet) { Write-Host ("  {0,-16}   {1}" -f '', $s.Does) -ForegroundColor DarkGray }
}

if (-not $Quiet) {
    Write-Host ''
    Write-Host 'To turn something off:  -Set TabDot=0        To put everything back:  -Reset' -ForegroundColor Gray
    Write-Host 'Word reads these once, at startup. Close it completely and start it again.' -ForegroundColor Yellow
}

# ---- and what the add-in itself last reported ------------------------------------------------------
#
# The table above is documentation and can go stale; these lines are the add-in saying what it
# actually did with these values the last time Word started. If the two disagree, believe these.

$log = Join-Path $env:LOCALAPPDATA 'WordTab\wordtab.log'
if (Test-Path $log) {
    $starts = @(Get-Content $log -ErrorAction SilentlyContinue |
                Where-Object { $_ -match 'StackStart|StripStart|FramesStart|TaskbarStart' })
    if ($starts.Count -gt 0) {
        Write-Host ''
        Write-Host 'What the add-in reported the last time Word started:' -ForegroundColor Cyan
        # The last four only. The log holds every start since it was last rolled, and there are four
        # of these lines per start - so printing all of them would show settings that have since been
        # changed, which is worse than showing none.
        foreach ($line in ($starts | Select-Object -Last 4)) {
            Write-Host ("  {0}" -f ($line -replace '^\d{4}-\d\d-\d\d ', '').Trim()) -ForegroundColor DarkGray
        }
    }
} else {
    Write-Host ''
    Write-Host "No add-in log yet ($log) - start Word once and run this again to see what it reports." -ForegroundColor DarkGray
}
