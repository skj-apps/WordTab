<#
.SYNOPSIS
  Measure where a per-document icon comes from, whether the shell's icons are actually different from
  each other, and whether one can be drawn onto the strip's surface without a black box behind it.

.DESCRIPTION
  An icon on a tab looks like a drawing job and is mostly not one. Three things decide the design and
  none of them is a matter of taste, so all three are measured here before a line is written:

    1. **Where the file type comes from.** The tab's text is Word's window title with the annotations
       trimmed off - and the title slice measured that Word follows Explorer's `HideFileExt`, so the
       title may carry no extension at all. `Document.Name` off the object model is the other
       candidate, on the same `Application.Windows` pass the dot already makes twice a second. This
       asks whether `Name` carries the extension when the title does not.

    2. **Whether the shell's icons differ.** A per-file-type icon is only worth the object-model read
       if `.docx`, `.doc`, `.rtf`, `.txt` and `.pdf` actually look different on this rig. If they come
       back identical, the honest answer is one Word glyph on every tab and no read at all. Compared
       by hashing the rendered pixels, not by eye.

    3. **Whether `DrawIconEx` composites onto the strip's own surface.** The strip draws into a
       top-down 32-bpp DIB it owns and blits it with `SRCCOPY`. A modern shell icon is 32-bit with an
       alpha channel; if `DrawIconEx` does not blend it, every tab gets a black square, which is the
       most visible bug this slice could ship. Measured by filling the DIB with a known colour,
       drawing the icon, and reading back a corner that should still be that colour.

  Also times the shell lookup, because it would otherwise sit on the janitor's half-second tick.

  Fixtures go in %TEMP%\wordtab-icons and Word is closed at both ends. Nothing here writes to the
  user's settings; `HideFileExt` is only *read*.

.EXAMPLE
  pwsh -File tools\probe-icons.ps1
#>
[CmdletBinding()]
param([switch]$KeepOpen)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$source = Get-Content -Raw -Path (Join-Path $PSScriptRoot 'WordLayout.cs')
Add-Type -TypeDefinition $source -Language CSharp -ReferencedAssemblies @(
    'System.Runtime', 'System.Collections', 'System.Threading.Thread', 'netstandard'
)
[WordLayout]::MakeDpiAware() | Out-Null

# The shell and GDI calls this probe needs. Deliberately its own type rather than an addition to
# WordLayout.cs: what the check suite ends up needing is not known until the design is settled, and a
# probe is evidence rather than infrastructure.
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;

public static class IconProbe
{
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    struct SHFILEINFOW
    {
        public IntPtr hIcon;
        public int    iIcon;
        public uint   dwAttributes;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] public string szDisplayName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]  public string szTypeName;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct BITMAPINFOHEADER
    {
        public uint biSize;
        public int  biWidth;
        public int  biHeight;
        public ushort biPlanes;
        public ushort biBitCount;
        public uint biCompression;
        public uint biSizeImage;
        public int  biXPelsPerMeter;
        public int  biYPelsPerMeter;
        public uint biClrUsed;
        public uint biClrImportant;
    }

    const uint SHGFI_ICON              = 0x000000100;
    const uint SHGFI_LARGEICON         = 0x000000000;
    const uint SHGFI_SMALLICON         = 0x000000001;
    const uint SHGFI_USEFILEATTRIBUTES = 0x000000010;
    const uint SHGFI_TYPENAME          = 0x000000400;
    const uint FILE_ATTRIBUTE_NORMAL   = 0x00000080;

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    static extern IntPtr SHGetFileInfoW(string path, uint attrs, ref SHFILEINFOW info, uint cb, uint flags);

    [DllImport("user32.dll")] static extern bool DestroyIcon(IntPtr icon);
    [DllImport("user32.dll")] static extern bool DrawIconEx(IntPtr dc, int x, int y, IntPtr icon,
                                                            int cx, int cy, uint step, IntPtr brush, uint flags);
    [DllImport("gdi32.dll")]  static extern IntPtr CreateCompatibleDC(IntPtr dc);
    [DllImport("gdi32.dll")]  static extern IntPtr CreateDIBSection(IntPtr dc, ref BITMAPINFOHEADER info,
                                                                    uint usage, out IntPtr bits, IntPtr section, uint offset);
    [DllImport("gdi32.dll")]  static extern IntPtr SelectObject(IntPtr dc, IntPtr obj);
    [DllImport("gdi32.dll")]  static extern bool DeleteObject(IntPtr obj);
    [DllImport("gdi32.dll")]  static extern bool DeleteDC(IntPtr dc);
    [DllImport("gdi32.dll")]  static extern bool GdiFlush();

    const uint DI_NORMAL = 0x0003;

    static IntPtr Lookup(string name, bool large, out string typeName)
    {
        SHFILEINFOW info = new SHFILEINFOW();
        uint flags = SHGFI_ICON | SHGFI_USEFILEATTRIBUTES | SHGFI_TYPENAME
                   | (large ? SHGFI_LARGEICON : SHGFI_SMALLICON);
        // USEFILEATTRIBUTES is the whole point: the shell answers from the name alone and never
        // touches the disk, so this costs nothing on a network path and works for a file that is not
        // there. The add-in will be asking about documents on corporate shares.
        SHGetFileInfoW(name, FILE_ATTRIBUTE_NORMAL, ref info, (uint)Marshal.SizeOf(typeof(SHFILEINFOW)), flags);
        typeName = info.szTypeName;
        return info.hIcon;
    }

    public static string TypeNameOf(string name)
    {
        string type;
        IntPtr icon = Lookup(name, true, out type);
        bool got = icon != IntPtr.Zero;
        if (got) DestroyIcon(icon);
        return (got ? "icon" : "NO ICON") + " type=|" + type + "|";
    }

    // Render an icon into a top-down 32-bpp DIB pre-filled with `bg`, exactly the way the strip's
    // surface is made, and report what came back. `corner` is a pixel the icon should not have
    // touched; `ink` is the middle of it. A corner that is not still `bg` means DrawIconEx did not
    // blend and the icon would arrive as a square.
    public static string Render(string name, int size, uint bg)
    {
        string type;
        IntPtr icon = Lookup(name, size > 16, out type);
        if (icon == IntPtr.Zero) return "NO ICON";

        BITMAPINFOHEADER h = new BITMAPINFOHEADER();
        h.biSize = (uint)Marshal.SizeOf(typeof(BITMAPINFOHEADER));
        h.biWidth = size; h.biHeight = -size;        // top-down, like the strip's
        h.biPlanes = 1; h.biBitCount = 32; h.biCompression = 0;

        IntPtr dc = CreateCompatibleDC(IntPtr.Zero);
        IntPtr bits;
        IntPtr dib = CreateDIBSection(dc, ref h, 0, out bits, IntPtr.Zero, 0);
        IntPtr old = SelectObject(dc, dib);

        int count = size * size;
        byte[] flat = new byte[count * 4];
        byte bb = (byte)(bg & 0xFF), gg = (byte)((bg >> 8) & 0xFF), rr = (byte)((bg >> 16) & 0xFF);
        for (int i = 0; i < count; i++)
        {
            flat[i * 4 + 0] = bb; flat[i * 4 + 1] = gg; flat[i * 4 + 2] = rr; flat[i * 4 + 3] = 255;
        }
        Marshal.Copy(flat, 0, bits, flat.Length);

        DrawIconEx(dc, 0, 0, icon, size, size, 0, IntPtr.Zero, DI_NORMAL);
        GdiFlush();
        Marshal.Copy(bits, flat, 0, flat.Length);

        SelectObject(dc, old);
        DeleteObject(dib);
        DeleteDC(dc);
        DestroyIcon(icon);

        int changed = 0;
        for (int i = 0; i < count; i++)
        {
            if (flat[i * 4 + 0] != bb || flat[i * 4 + 1] != gg || flat[i * 4 + 2] != rr) changed++;
        }

        // FNV-1a over the colour bytes only. The alpha byte is deliberately excluded: the strip blits
        // with SRCCOPY and never reads it, and DrawIconEx leaves it in a state that is not worth
        // asserting anything about.
        uint hash = 2166136261;
        for (int i = 0; i < count; i++)
            for (int c = 0; c < 3; c++) { hash ^= flat[i * 4 + c]; hash *= 16777619; }

        int mid = ((size / 2) * size + (size / 2)) * 4;
        return string.Format(
            "hash={0:X8} changed={1}/{2} corner=#{3:X2}{4:X2}{5:X2} centre=#{6:X2}{7:X2}{8:X2} type=|{9}|",
            hash, changed, count,
            flat[2], flat[1], flat[0],
            flat[mid + 2], flat[mid + 1], flat[mid + 0], type);
    }

    public static double TimeLookup(string name, int passes)
    {
        var sw = System.Diagnostics.Stopwatch.StartNew();
        for (int i = 0; i < passes; i++)
        {
            string type;
            IntPtr icon = Lookup(name, true, out type);
            if (icon != IntPtr.Zero) DestroyIcon(icon);
        }
        sw.Stop();
        return sw.Elapsed.TotalMilliseconds / passes;
    }
}
'@ -Language CSharp -ReferencedAssemblies @('System.Runtime', 'System.Drawing.Primitives', 'netstandard')

function Write-Step($text) { Write-Host "==> $text" -ForegroundColor Cyan }
function Write-Note($text) { Write-Host "    $text" -ForegroundColor DarkGray }
function Write-Fact($text) { Write-Host "    $text" -ForegroundColor Yellow }

function Get-WordPids { @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }) }
function Get-Frames {
    $found = @()
    foreach ($id in Get-WordPids) { $found += [WordLayout]::Frames($id) }
    return @($found)
}

function Close-Word {
    if (@(Get-WordPids).Count -eq 0) { return }
    for ($guard = 0; $guard -lt 24; $guard++) {
        $open = @(Get-Frames)
        if ($open.Count -eq 0) { break }
        [WordLayout]::Close($open[0])
        Start-Sleep -Milliseconds 800
    }
    foreach ($p in @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue)) { $p.CloseMainWindow() | Out-Null }
    $deadline = (Get-Date).AddSeconds(20)
    while ((Get-Date) -lt $deadline -and @(Get-WordPids).Count -gt 0) { Start-Sleep -Milliseconds 500 }
    Start-Sleep -Seconds 2
}

# ---- 1. Does the shell give different icons for different document types? ------------------------
#
# First, because if the answer is no then section 2 has nothing to decide and the object-model read
# is not worth making.

Write-Step 'The shell, asked about a name it never has to open'

$kinds = @('Report.docx', 'Report.doc', 'Report.rtf', 'Report.txt', 'Report.pdf',
           'Report.odt', 'Report.dotx', 'Report.docm', 'Report')

$rendered = @()
foreach ($k in $kinds) {
    $large = [IconProbe]::Render($k, 32, 0xE4E4E4)
    $rendered += [pscustomobject]@{ Name = $k; Large = $large }
    Write-Fact ("{0,-12} 32px  {1}" -f $k, $large)
}

Write-Step 'Which of those are actually distinct'
$byHash = @{}
foreach ($r in $rendered) {
    $h = ($r.Large -split ' ')[0]
    if (-not $byHash.ContainsKey($h)) { $byHash[$h] = @() }
    $byHash[$h] += $r.Name
}
foreach ($h in $byHash.Keys) {
    Write-Fact ("{0} -> {1}" -f $h, ($byHash[$h] -join ', '))
}
Write-Fact ("{0} distinct icons across {1} extensions" -f $byHash.Keys.Count, $kinds.Count)

# ---- 2. Does DrawIconEx blend onto the strip's kind of surface? ----------------------------------
#
# The corner pixel is the whole question. It was filled with the well colour and the icon's own corner
# is transparent, so if it comes back anything else - black, most likely - the icon is arriving as an
# opaque square and this slice needs to composite the bits by hand instead.

Write-Step 'DrawIconEx onto a top-down 32-bpp DIB, the way the strip draws'
foreach ($bg in @(0xE4E4E4, 0x0F0F0F)) {
    foreach ($size in @(16, 32)) {
        $out = [IconProbe]::Render('Report.docx', $size, $bg)
        Write-Fact ("background #{0:X6}  {1}px  {2}" -f $bg, $size, $out)
    }
}
Write-Note 'corner must still be the background colour; centre must not be.'

Write-Step 'What the lookup costs'
foreach ($k in @('Report.docx', 'Report.rtf')) {
    Write-Fact ("{0,-12} {1:N3} ms per SHGetFileInfo (icon created and destroyed each time)" -f
                $k, [IconProbe]::TimeLookup($k, 200))
}

# ---- 3. Where the extension comes from -----------------------------------------------------------
#
# The title may not have one. Document.Name is the candidate, on the pass the dot already makes.

Write-Step 'Explorer''s HideFileExt right now (read only, never written here)'
$hide = 1
try {
    $hide = (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name HideFileExt).HideFileExt
} catch { }
Write-Fact ("HideFileExt = {0}  ({1})" -f $hide, $(if ($hide -eq 1) { 'extensions hidden - the hard case' } else { 'extensions shown' }))

$dir = Join-Path $env:TEMP 'wordtab-icons'
Write-Step 'Word must be closed to start clean'
Close-Word
if (Test-Path $dir) { Remove-Item -Path $dir -Recurse -Force }
New-Item -ItemType Directory -Path $dir -Force | Out-Null

Write-Step 'Document.Name and the window title, side by side, per document type'

# Word's own SaveAs format numbers. Authored by Word itself so the files are real, the same way the
# title slice made its fixtures.
$formats = @(
    @{ Ext = 'docx'; Format = 16 },
    @{ Ext = 'doc';  Format = 0  },
    @{ Ext = 'rtf';  Format = 6  },
    @{ Ext = 'txt';  Format = 2  }
)

$word = New-Object -ComObject Word.Application
try {
    $word.Visible = $true
    $word.DisplayAlerts = 0
    Start-Sleep -Seconds 3

    $fresh = $word.Documents.Add()
    Start-Sleep -Seconds 2
    Write-Fact ("brand new, never saved:   Name=|{0}|  FullName=|{1}|  Path=|{2}|" -f
                $fresh.Name, $fresh.FullName, $fresh.Path)

    foreach ($f in $formats) {
        $path = Join-Path $dir ("Icon probe." + $f.Ext)
        $doc = $word.Documents.Add()
        Start-Sleep -Milliseconds 800
        $word.Selection.TypeText("WordTab icon probe.")
        $doc.SaveAs2($path, $f.Format)
        Start-Sleep -Seconds 1

        $title = ''
        foreach ($fr in @(Get-Frames)) {
            $t = [WordLayout]::TitleOf($fr)
            if ($t -like ("*Icon probe*")) { $title = $t }
        }

        Write-Fact ("{0,-5} Name=|{1}|  title=|{2}|" -f $f.Ext, $doc.Name, $title)
        Write-Note ("      FullName=|{0}|" -f $doc.FullName)
        $doc.Saved = $true
        $doc.Close(0)
        Start-Sleep -Milliseconds 800
    }

    $fresh.Saved = $true
    try { $fresh.Close(0) } catch { }
} finally {
    try { $word.Quit(0) } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

if (-not $KeepOpen) { Close-Word }

Write-Host ''
Write-Host 'Probe complete - the Yellow lines are the measurements.' -ForegroundColor Green
