<#
.SYNOPSIS
  The questions install.ps1 and settings.ps1 both have to ask Word's registry.

.DESCRIPTION
  Dot-sourced by both. It exists because there is exactly one right answer to "has Word disabled this
  add-in" and "which tabbed-Word rivals will load", and this project has already paid three times for
  the same question being answered in two places - most expensively with three implementations of
  "is Word asking something", none of which was correct.

  Nothing here writes anything. Every function reads the registry and returns a value.

  **Everything in here answers about Word 16.0** (Office 2016 through Microsoft 365, which share that
  version number). The add-in registration itself is version-independent - Word reads
  HKCU\Software\Microsoft\Office\Word\Addins whatever the version - but Resiliency is under the
  version, so that one path is 16.0 and says so.
#>

# Word keeps a list of add-ins it killed after a crash or a failed load. An entry here beats
# LoadBehavior=3 outright and is the classic "I registered it and nothing happens" cause.
#
# The entries are NAMED rather than counted, and that is the point of this function. Each value is a
# fixed-width binary blob carrying the add-in's DLL path and the vendor's friendly name as
# NUL-separated UTF-16, so "is it US Word disabled, or somebody else" can be answered instead of
# raised - and those are different situations with different fixes.
function Get-DisabledAddinPaths {
    $key = 'HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency\DisabledItems'
    if (-not (Test-Path $key)) { return @() }
    $found = @()
    foreach ($name in @((Get-Item $key).Property)) {
        $bytes = (Get-ItemProperty -Path $key -Name $name -ErrorAction SilentlyContinue).$name
        if ($bytes -isnot [byte[]]) { continue }
        # Trim: a blob is fixed-width and padded, so the tail of a field is NULs. Only the PATH
        # fields are kept - each blob also carries the vendor's friendly name, and keeping those
        # would double the count and give the path matching a second thing to be confused by.
        foreach ($field in ([System.Text.Encoding]::Unicode.GetString($bytes) -split "`0")) {
            $field = $field.Trim()
            if ($field -match '\.dll$' -or $field -match '\.(vsto|xll|wll)$') { $found += $field }
        }
    }
    return @($found)
}

# A ProgId's DLL, resolved the way Word resolves it. Used to tell a disabled entry apart from ours by
# PATH rather than by a friendly name somebody else chose.
function Get-AddinDllPath($progId) {
    foreach ($hive in @('HKCU:', 'HKLM:')) {
        $classKey = "$hive\Software\Classes\$progId\CLSID"
        if (-not (Test-Path $classKey)) { continue }
        $guid = (Get-ItemProperty -Path $classKey -ErrorAction SilentlyContinue).'(default)'
        if (-not $guid) { continue }
        $serverKey = "$hive\Software\Classes\CLSID\$guid\InprocServer32"
        if (-not (Test-Path $serverKey)) { continue }
        $path = (Get-ItemProperty -Path $serverKey -ErrorAction SilentlyContinue).'(default)'
        if ($path) { return $path }
    }
    return $null
}

# What Word has been asked to do about a ProgId, resolved the way Word resolves it.
#
# Two rules decide it and getting either wrong reports an add-in that is not there:
#
#   - **HKCU beats HKLM for the same ProgId.** Word reads the per-user value and stops; a machine-wide
#     LoadBehavior=3 sitting under a per-user 0 or 2 is not an add-in that loads.
#   - **Disabled Items beats LoadBehavior.** An add-in Word killed after a failed load stays dead at
#     LoadBehavior=3.
#
# Returns Behavior ($null when the ProgId is not registered at all), the hive that answered, the DLL,
# and Blocked. **`Registered` deliberately means "asked to load at startup", not "will load"** - the
# registry says what Word was told, and the dev rig has a ProgId at LoadBehavior=3, absent from
# Disabled Items, whose DLL never turns up in WINWORD's module list.
function Get-AddinStatus($progId, $disabledPaths = $null) {
    if ($null -eq $disabledPaths) { $disabledPaths = @(Get-DisabledAddinPaths) }

    $behavior = $null; $fromHive = $null
    foreach ($hive in @('HKCU:', 'HKLM:')) {
        $key = "$hive\Software\Microsoft\Office\Word\Addins\$progId"
        if (-not (Test-Path $key)) { continue }
        $value = (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue).LoadBehavior
        if ($null -eq $value) { continue }
        $behavior = [int]$value; $fromHive = $hive
        break                       # per-user wins outright, so the first hive with a value decides
    }

    $dll = Get-AddinDllPath $progId
    $blocked = $false
    if ($dll) {
        $blocked = @($disabledPaths | Where-Object { $_ -and $_.ToLowerInvariant() -eq $dll.ToLowerInvariant() }).Count -gt 0
    }

    return [pscustomobject]@{
        ProgId     = $progId
        Behavior   = $behavior
        Hive       = $fromHive
        Dll        = $dll
        Blocked    = $blocked
        Registered = (($behavior -eq 3) -and (-not $blocked))
    }
}

# The other tabbed-Word add-ins this project knows about. Office Tab ships three ProgIds and any one
# of them loading is enough to matter, because they all drive the same document frame WordTab does.
$script:WordTabRivals = @('OfficeTab.TabsforWord2013', 'OfficeTabs.Connect', 'TabsforOfficeHelper.Helper')

function Get-RivalAddins {
    $disabled = @(Get-DisabledAddinPaths)
    return @($script:WordTabRivals | ForEach-Object { Get-AddinStatus $_ $disabled })
}
