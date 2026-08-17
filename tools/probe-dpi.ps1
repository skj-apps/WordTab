<#
    probe-dpi.ps1 - read, and optionally change, the desktop's display scaling.

    Why this exists: `StripOnFrameDpiChanged` has never executed. The add-in has only ever run at
    this rig's 192 dpi (200%), and the work rig is a laptop plus a 40" second monitor - almost
    certainly two different scale factors, which is the one arrangement that runs that code path.
    The project's notes recorded "changing the desktop's scaling needs an undocumented API" as a
    blocker. It is an undocumented API, but it is a reachable one - so the blocker was worth
    retiring, and this is what retired it.

    WHAT IT ANSWERED, 2026-08-17, AND IT IS A NO: this rig has one display source and the packet
    reports its relative range as 0..0. There is no second scale factor to switch to, so the desktop
    cannot be moved off 192 dpi here by this route or any other. That is why the add-in carries a
    TabDpi override instead - see DpiOverride in strip.cpp. This script's remaining job is to be the
    thing that says so, and to be re-run on a machine with two monitors, where the answer differs.

    AND A CAUTION IT MEASURED ABOUT ITSELF: the packet returns a SUCCESS code while answering 0/0/0,
    and on this rig that maps to "current 100%" on a monitor Windows independently reports at 192 dpi
    (200%). The relative indices are read straight from the packet and are sound; the percentages go
    through a hardcoded table and here they are wrong. The report cross-checks against
    GetDpiForMonitor and prints the disagreement rather than the number.

    DISPLAYCONFIG_DEVICE_INFO_GET_SOURCE_DPI_SCALE (-3) and _SET_SOURCE_DPI_SCALE (-4) are the two
    private DisplayConfig packets the Settings app itself uses. They are not in any SDK header. The
    scale is expressed RELATIVE to the monitor's recommended scaling, as an index into a fixed table
    of percentages, which is why the read has to come first: -2 means nothing until you know what
    the recommended value is.

    -Report        read every monitor's current / recommended / available scaling. Changes nothing.
    -Percent <n>   set the primary monitor's scaling to n percent.
    -Restore       put the scaling back to whatever -Report last saved.

    A set writes to the user's display settings. It is reversible and this script saves the previous
    value to $env:TEMP\wordtab-dpi-before.txt before touching anything, but it is the user's own
    desktop: every window on it will be resized. Do not run the set path unattended.
#>
[CmdletBinding()]
param(
    [switch] $Report,
    [int]    $Percent = 0,
    [switch] $Restore
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$code = @'
using System;
using System.Runtime.InteropServices;

public static class DisplayScale
{
    [StructLayout(LayoutKind.Sequential)]
    public struct LUID { public uint LowPart; public int HighPart; }

    [StructLayout(LayoutKind.Sequential)]
    public struct DEVICE_INFO_HEADER
    {
        public int  type;
        public uint size;
        public LUID adapterId;
        public uint id;
    }

    // The private GET packet. minScaleRel is negative (or zero) and its absolute value is the index
    // of the RECOMMENDED scaling in the table below; curScaleRel and maxScaleRel are offsets from it.
    [StructLayout(LayoutKind.Sequential)]
    public struct SOURCE_DPI_SCALE_GET
    {
        public DEVICE_INFO_HEADER header;
        public int minScaleRel;
        public int curScaleRel;
        public int maxScaleRel;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct SOURCE_DPI_SCALE_SET
    {
        public DEVICE_INFO_HEADER header;
        public int scaleRel;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct DISPLAYCONFIG_PATH_SOURCE_INFO
    {
        public LUID adapterId;
        public uint id;
        public uint modeInfoIdx;
        public uint statusFlags;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct DISPLAYCONFIG_PATH_TARGET_INFO
    {
        public LUID adapterId;
        public uint id;
        public uint modeInfoIdx;
        public uint outputTechnology;
        public uint rotation;
        public uint scaling;
        public uint refreshNumerator;
        public uint refreshDenominator;
        public uint scanLineOrdering;
        [MarshalAs(UnmanagedType.Bool)] public bool targetAvailable;
        public uint statusFlags;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct DISPLAYCONFIG_PATH_INFO
    {
        public DISPLAYCONFIG_PATH_SOURCE_INFO sourceInfo;
        public DISPLAYCONFIG_PATH_TARGET_INFO targetInfo;
        public uint flags;
    }

    // We only ever need the paths, but QueryDisplayConfig insists on a mode array too. The mode
    // union is 64 bytes plus the 20-byte header-ish prefix; we never read it, so a blob of the
    // right size is enough and saves declaring three more structures.
    [StructLayout(LayoutKind.Sequential)]
    public struct DISPLAYCONFIG_MODE_INFO_BLOB
    {
        public uint infoType;
        public uint id;
        public LUID adapterId;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 64)] public byte[] blob;
    }

    [DllImport("user32.dll")]
    public static extern int GetDisplayConfigBufferSizes(uint flags, out uint numPaths, out uint numModes);

    [DllImport("user32.dll")]
    public static extern int QueryDisplayConfig(uint flags, ref uint numPaths,
        [Out] DISPLAYCONFIG_PATH_INFO[] paths, ref uint numModes,
        [Out] DISPLAYCONFIG_MODE_INFO_BLOB[] modes, IntPtr currentTopologyId);

    [DllImport("user32.dll")]
    public static extern int DisplayConfigGetDeviceInfo(ref SOURCE_DPI_SCALE_GET packet);

    [DllImport("user32.dll")]
    public static extern int DisplayConfigSetDeviceInfo(ref SOURCE_DPI_SCALE_SET packet);

    public const uint QDC_ONLY_ACTIVE_PATHS = 0x00000002;
    public const int  GET_SOURCE_DPI_SCALE  = -3;
    public const int  SET_SOURCE_DPI_SCALE  = -4;

    // The fixed table the relative index addresses. Windows exposes no way to read it; it is the
    // same list the Settings drop-down shows.
    public static readonly int[] Percentages =
        new int[] { 100, 125, 150, 175, 200, 225, 250, 300, 350, 400, 450, 500 };

    // The control on everything above it. The private packet is undocumented and returns success
    // even when it has nothing to say, so its answer needs checking against a number Windows will
    // state on the record.
    //
    // AND THE CONTROL NEEDED A CONTROL. MDT_EFFECTIVE_DPI is "effective" in the literal sense: it is
    // reported through the DPI virtualisation of the CALLING process. Measured, after asserting the
    // opposite in this comment and being wrong: the same monitor answers 192 under pwsh 7, which
    // declares per-monitor awareness in its manifest, and 96 under Windows PowerShell 5.1, which
    // does not. 5.1 is the shell the machine this script exists for is guaranteed to have - so
    // unfixed, the cross-check would have reported "packet 100%, monitor 100%, they agree" on the
    // work rig while BOTH numbers were wrong, which is worse than not checking at all.
    //
    // So declare the awareness before asking. PER_MONITOR_AWARE_V2 (-4). It fails harmlessly when
    // the host already set one, which is the pwsh 7 case, and it is per-process and dies with the
    // script.
    [DllImport("user32.dll")] public static extern IntPtr GetDesktopWindow();
    [DllImport("user32.dll")] public static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint flags);
    [DllImport("user32.dll")] public static extern bool SetProcessDpiAwarenessContext(IntPtr ctx);
    [DllImport("shcore.dll")] public static extern int GetDpiForMonitor(IntPtr mon, int type,
                                                                       out uint dpiX, out uint dpiY);

    public static int PrimaryDpi()
    {
        try { SetProcessDpiAwarenessContext(new IntPtr(-4)); } catch { }

        IntPtr mon = MonitorFromWindow(GetDesktopWindow(), 1 /* MONITOR_DEFAULTTOPRIMARY */);
        uint x, y;
        if (GetDpiForMonitor(mon, 0 /* MDT_EFFECTIVE_DPI */, out x, out y) != 0) return 0;
        return (int)x;
    }

    public static DISPLAYCONFIG_PATH_INFO[] ActivePaths()
    {
        uint numPaths, numModes;
        int rc = GetDisplayConfigBufferSizes(QDC_ONLY_ACTIVE_PATHS, out numPaths, out numModes);
        if (rc != 0) throw new Exception("GetDisplayConfigBufferSizes failed: " + rc);

        DISPLAYCONFIG_PATH_INFO[] paths = new DISPLAYCONFIG_PATH_INFO[numPaths];
        DISPLAYCONFIG_MODE_INFO_BLOB[] modes = new DISPLAYCONFIG_MODE_INFO_BLOB[numModes];
        rc = QueryDisplayConfig(QDC_ONLY_ACTIVE_PATHS, ref numPaths, paths, ref numModes, modes, IntPtr.Zero);
        if (rc != 0) throw new Exception("QueryDisplayConfig failed: " + rc);

        Array.Resize(ref paths, (int)numPaths);
        return paths;
    }

    // Returns { recommended, current, minimum, maximum, curRel, minRel, maxRel } as percentages
    // (and raw relative indices), or null if the packet is refused for this source.
    public static int[] GetScaling(LUID adapterId, uint sourceId)
    {
        SOURCE_DPI_SCALE_GET packet = new SOURCE_DPI_SCALE_GET();
        packet.header.type      = GET_SOURCE_DPI_SCALE;
        packet.header.size      = (uint)Marshal.SizeOf(typeof(SOURCE_DPI_SCALE_GET));
        packet.header.adapterId = adapterId;
        packet.header.id        = sourceId;

        int rc = DisplayConfigGetDeviceInfo(ref packet);
        if (rc != 0) return null;

        int cur = packet.curScaleRel;
        if (cur < packet.minScaleRel) cur = packet.minScaleRel;
        if (cur > packet.maxScaleRel) cur = packet.maxScaleRel;

        int recIdx = Math.Abs(packet.minScaleRel);
        if (recIdx >= Percentages.Length) return null;

        int curIdx = recIdx + cur;
        int maxIdx = recIdx + packet.maxScaleRel;
        if (curIdx < 0 || curIdx >= Percentages.Length) return null;
        if (maxIdx < 0 || maxIdx >= Percentages.Length) maxIdx = Percentages.Length - 1;

        return new int[] {
            Percentages[recIdx], Percentages[curIdx], Percentages[0], Percentages[maxIdx],
            packet.curScaleRel, packet.minScaleRel, packet.maxScaleRel
        };
    }

    public static int SetScaling(LUID adapterId, uint sourceId, int relative)
    {
        SOURCE_DPI_SCALE_SET packet = new SOURCE_DPI_SCALE_SET();
        packet.header.type      = SET_SOURCE_DPI_SCALE;
        packet.header.size      = (uint)Marshal.SizeOf(typeof(SOURCE_DPI_SCALE_SET));
        packet.header.adapterId = adapterId;
        packet.header.id        = sourceId;
        packet.scaleRel         = relative;
        return DisplayConfigSetDeviceInfo(ref packet);
    }
}
'@

if (-not ('DisplayScale' -as [type])) {
    Add-Type -TypeDefinition $code -Language CSharp
}

$savePath = Join-Path $env:TEMP 'wordtab-dpi-before.txt'

function Get-Sources
{
    $paths = [DisplayScale]::ActivePaths()
    $seen  = @{}
    $out   = New-Object System.Collections.ArrayList

    foreach ($p in $paths) {
        $key = '{0}:{1}:{2}' -f $p.sourceInfo.adapterId.LowPart, $p.sourceInfo.adapterId.HighPart, $p.sourceInfo.id
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true

        $info = [DisplayScale]::GetScaling($p.sourceInfo.adapterId, $p.sourceInfo.id)

        # Built field by field rather than with `if` expressions inside the literal: this script has
        # to run under Windows PowerShell 5.1 as well as pwsh 7, and the two do not agree about what
        # is legal as a hashtable value.
        $row = [pscustomobject]@{
            Adapter     = $p.sourceInfo.adapterId
            SourceId    = $p.sourceInfo.id
            Recommended = $null
            Current     = $null
            Maximum     = $null
            CurRel      = $null
            MinRel      = $null
            MaxRel      = $null
        }
        if ($null -ne $info) {
            $row.Recommended = $info[0]
            $row.Current     = $info[1]
            $row.Maximum     = $info[3]
            $row.CurRel      = $info[4]
            $row.MinRel      = $info[5]
            $row.MaxRel      = $info[6]
        }
        $null = $out.Add($row)
    }
    return $out
}

if ($Report -or (-not $Restore -and $Percent -eq 0)) {
    # @() around the call, not around the ArrayList inside Get-Sources. A `return $out` unrolls the
    # list into the pipeline, so a rig with ONE display source hands back a bare object - and
    # `.Count` on that is a hard StrictMode error under Windows PowerShell 5.1, which is the shell
    # the machine this script exists to be run on is guaranteed to have. One monitor is the case
    # this rig is in and the case the script was written against.
    $sources = @(Get-Sources)
    Write-Host ''
    Write-Host ("Active display sources: {0}" -f $sources.Count)
    Write-Host ''
    foreach ($s in $sources) {
        if ($null -eq $s.Current) {
            Write-Host ("  source {0}  -- the scaling packet was refused for this source" -f $s.SourceId)
            continue
        }
        $steps = [DisplayScale]::Percentages |
                 Where-Object { $_ -le $s.Maximum } |
                 ForEach-Object { if ($_ -eq $s.Current) { "[$_]" } else { "$_" } }
        Write-Host ("  source {0}   current {1}%   recommended {2}%   max {3}%" -f `
                    $s.SourceId, $s.Current, $s.Recommended, $s.Maximum)
        Write-Host ("               available: {0}" -f ($steps -join '  '))
        Write-Host ("               relative index: cur {0}, range {1}..{2}" -f $s.CurRel, $s.MinRel, $s.MaxRel)
    }
    # The cross-check, and on this rig it is the whole finding. The packet is undocumented, and it
    # answers 0/0/0 with a success code rather than failing when it has nothing to report - which is
    # indistinguishable from "this monitor genuinely sits at the first entry in the table" unless
    # something else is asked. So ask something else.
    #
    # MEASURED HERE 2026-08-17: the packet says current 100%, and the primary monitor is running at
    # 192 dpi, which is 200%. Both cannot be true. The RELATIVE range 0..0 is still good - it is read
    # straight out of the packet rather than mapped through the percentage table, and it is what says
    # no other scale factor is on offer - but the ABSOLUTE percentages are not trustworthy here, and
    # a report that printed them unqualified would be stating a false number confidently.
    $sysDpi = [DisplayScale]::PrimaryDpi()
    if ($sysDpi -gt 0) {
        $sysPct = [int][Math]::Round(100.0 * $sysDpi / 96.0)
        Write-Host ("Windows reports the primary monitor at {0} dpi ({1}%)" -f $sysDpi, $sysPct)

        $primary = $sources | Where-Object { $null -ne $_.Current } | Select-Object -First 1
        if ($primary -and $primary.Current -ne $sysPct) {
            Write-Host ''
            Write-Host ("  ** The packet and Windows DISAGREE: packet {0}%, monitor {1}%." -f $primary.Current, $sysPct)
            Write-Host  '     Trust the relative range, not the percentages. A source that answers'
            Write-Host  '     0/0/0 with a success code offers no alternative scale either way, so'
            Write-Host  '     the conclusion "there is nothing to switch to here" still holds - but'
            Write-Host  '     do not quote the percentages off this machine as if they were read.'
        }
    } else {
        Write-Host 'GetDpiForMonitor would not answer, so the packet has nothing to be checked against.'
    }
    Write-Host ''

    # Save the current state so -Restore has something to put back. Written on every report so the
    # file always reflects the last state anyone looked at, rather than the first.
    $sources | Where-Object { $null -ne $_.Current } | ForEach-Object {
        '{0} {1} {2} {3}' -f $_.Adapter.LowPart, $_.Adapter.HighPart, $_.SourceId, $_.CurRel
    } | Set-Content -Path $savePath -Encoding ASCII
    Write-Host ("Saved current scaling to {0}" -f $savePath)
    return
}

if ($Restore) {
    if (-not (Test-Path $savePath)) { throw "No saved scaling at $savePath - run -Report first." }
    foreach ($line in (Get-Content $savePath)) {
        $bits = $line -split ' '
        if ($bits.Count -ne 4) { continue }
        $luid = New-Object DisplayScale+LUID
        $luid.LowPart  = [uint32]$bits[0]
        $luid.HighPart = [int]$bits[1]
        $rc = [DisplayScale]::SetScaling($luid, [uint32]$bits[2], [int]$bits[3])
        Write-Host ("restore source {0} -> relative {1}  (rc={2})" -f $bits[2], $bits[3], $rc)
    }
    return
}

if ($Percent -gt 0) {
    $sources = Get-Sources
    $primary = $sources | Where-Object { $null -ne $_.Current } | Select-Object -First 1
    if (-not $primary) { throw "No source would answer the scaling packet." }

    $idx = [array]::IndexOf([DisplayScale]::Percentages, $Percent)
    if ($idx -lt 0) { throw "$Percent% is not one of: $([DisplayScale]::Percentages -join ', ')" }

    $recIdx   = [Math]::Abs($primary.MinRel)
    $relative = $idx - $recIdx
    if ($relative -lt $primary.MinRel -or $relative -gt $primary.MaxRel) {
        throw ("$Percent% is outside this monitor's range ({0}%..{1}%)." -f 100, $primary.Maximum)
    }

    Write-Host ("source {0}: {1}% -> {2}%  (relative {3} -> {4})" -f `
                $primary.SourceId, $primary.Current, $Percent, $primary.CurRel, $relative)
    $rc = [DisplayScale]::SetScaling($primary.Adapter, $primary.SourceId, $relative)
    Write-Host ("DisplayConfigSetDeviceInfo rc=$rc")
    if ($rc -ne 0) { throw "DisplayConfigSetDeviceInfo refused the change (rc=$rc)." }
    return
}
