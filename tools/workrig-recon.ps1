# workrig-recon.ps1 - read-only compatibility recon for the corporate target machine.
#
# Answers three questions before we design deployment:
#   1. Will an unsigned, per-user binary run there at all? (AppLocker / WDAC / SRP)
#   2. Does Office policy require add-ins be signed by a trusted publisher?
#   3. What runtimes already exist, so we know whether to ship self-contained?
#
# Writes nothing, installs nothing, needs no admin. Windows PowerShell 5.1 compatible.
# Run it, then paste the whole output back.

$r = [ordered]@{}

# --- 1. AppLocker ---------------------------------------------------------
# Rules are only enforced when the Application Identity service is running.
try {
    $svc = Get-CimInstance Win32_Service -Filter "Name='AppIDSvc'" -ErrorAction Stop
    if ($svc) { $r['AppLocker service'] = "State=$($svc.State) StartMode=$($svc.StartMode)" }
    else      { $r['AppLocker service'] = 'not present' }
} catch { $r['AppLocker service'] = "unreadable: $($_.Exception.Message)" }

# Which rule collections exist and whether each is enforcing or just auditing.
# The 'Dll' collection is the one that would affect an in-process add-in; it is
# off by default even on machines that enforce Exe rules.
try {
    $pol = Get-AppLockerPolicy -Effective -ErrorAction Stop
    $cols = $pol.RuleCollections | ForEach-Object {
        "$($_.RuleCollectionType)=$($_.EnforcementMode)/$($_.Count)rules"
    }
    if ($cols) { $r['AppLocker collections'] = ($cols -join ', ') }
    else       { $r['AppLocker collections'] = 'no rules defined' }
} catch {
    # The AppLocker module is absent on Home editions. Read the policy store directly.
    $srpV2 = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\SrpV2'
    if (Test-Path $srpV2) {
        $cols = Get-ChildItem $srpV2 -ErrorAction SilentlyContinue | ForEach-Object {
            $mode = (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).EnforcementMode
            "$($_.PSChildName)=$mode"
        }
        $r['AppLocker collections'] = 'via registry: ' + ($cols -join ', ')
    } else {
        $r['AppLocker collections'] = 'no AppLocker policy present'
    }
}

# --- 2. WDAC / Device Guard / Smart App Control ---------------------------
# CodeIntegrityPolicyEnforcementStatus: 0=off 1=audit 2=enforced
try {
    $dg = Get-CimInstance -Namespace root\Microsoft\Windows\DeviceGuard `
                          -ClassName Win32_DeviceGuard -ErrorAction Stop
    $r['WDAC kernel CI'] = $dg.CodeIntegrityPolicyEnforcementStatus
    $r['WDAC usermode CI'] = $dg.UsermodeCodeIntegrityPolicyEnforcementStatus
} catch { $r['WDAC'] = "unreadable: $($_.Exception.Message)" }

# Smart App Control (Win11): 0=off 1=enforced 2=evaluation. Blocks unsigned exes.
$ci = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\CI\Policy' -ErrorAction SilentlyContinue
if ($ci -and $null -ne $ci.VerifiedAndReputablePolicyState) {
    $r['Smart App Control'] = $ci.VerifiedAndReputablePolicyState
} else {
    $r['Smart App Control'] = 'not configured'
}

# Software Restriction Policies - the older mechanism, still seen in corp images.
$srp = Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Safer\CodeIdentifiers' -ErrorAction SilentlyContinue
if ($srp -and ($null -ne $srp.DefaultLevel -or $null -ne $srp.TransparentEnabled)) {
    $r['SRP'] = "DefaultLevel=$($srp.DefaultLevel) TransparentEnabled=$($srp.TransparentEnabled)"
} else {
    $r['SRP'] = 'not configured'
}

# --- 3. Office add-in policy ----------------------------------------------
# RequireAddinSig=1 would block an unsigned add-in even if nothing else does,
# and would explain why signed Office Tab loads there but ours might not.
$secPaths = @(
    'HKCU:\Software\Policies\Microsoft\Office\16.0\Word\Security',
    'HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Word\Security',
    'HKCU:\Software\Microsoft\Office\16.0\Word\Security',
    'HKCU:\Software\Policies\Microsoft\Office\16.0\Common\Security'
)
$found = @()
foreach ($p in $secPaths) {
    $k = Get-ItemProperty $p -ErrorAction SilentlyContinue
    if ($k) {
        $k.PSObject.Properties |
            Where-Object { $_.Name -match 'RequireAddinSig|DisableAllAddins|NoTBPromptUnsignedAddin|DisableCOMAddins' } |
            ForEach-Object { $found += "$($p.Split('\')[0])..$($_.Name)=$($_.Value)" }
    }
}
if ($found) { $r['Office add-in policy'] = ($found -join ', ') }
else        { $r['Office add-in policy'] = 'no signing/disable policy set (good)' }

# --- 4. Office build ------------------------------------------------------
$c2r = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
if ($c2r) { $r['Office'] = "$($c2r.ProductReleaseIds) v$($c2r.VersionToReport) $($c2r.Platform)" }
else      { $r['Office'] = 'no ClickToRun key - MSI or Store install?' }

# --- 5. Runtimes already present ------------------------------------------
$shared = 'C:\Program Files\dotnet\shared\Microsoft.NETCore.App'
if (Test-Path $shared) { $r['.NET runtimes'] = ((Get-ChildItem $shared -Directory).Name -join ', ') }
else                   { $r['.NET runtimes'] = 'none' }

$ndp = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -ErrorAction SilentlyContinue
if ($ndp) { $r['.NET Framework 4.x'] = "Release=$($ndp.Release)" } else { $r['.NET Framework 4.x'] = 'absent' }

# --- 6. Office Tab as the known-good deployment blueprint -----------------
foreach ($hive in @('HKCU:\Software\Microsoft\Office\Word\Addins',
                    'HKLM:\SOFTWARE\Microsoft\Office\Word\Addins')) {
    $names = Get-ChildItem $hive -ErrorAction SilentlyContinue | Select-Object -ExpandProperty PSChildName
    $label = 'Word add-ins ' + $hive.Split('\')[0]
    if ($names) { $r[$label] = ($names -join ', ') } else { $r[$label] = 'none' }
}

# Where does Office Tab's DLL actually live, and which hive registered its class?
foreach ($progId in @('OfficeTab.TabsforWord2013', 'OfficeTabs.Connect')) {
    foreach ($classes in @('HKCU:\Software\Classes', 'HKLM:\SOFTWARE\Classes')) {
        $clsid = (Get-ItemProperty "$classes\$progId\CLSID" -ErrorAction SilentlyContinue).'(default)'
        if ($clsid) {
            $dll = (Get-ItemProperty "$classes\CLSID\$clsid\InprocServer32" -ErrorAction SilentlyContinue).'(default)'
            $r["$progId @ $($classes.Split('\')[0])"] = "$clsid -> $dll"
        }
    }
}

# --- report ---------------------------------------------------------------
# Print it and put it on the clipboard, so it can be pasted straight back.
$txt = ($r.GetEnumerator() | ForEach-Object { '{0,-34} {1}' -f $_.Key, $_.Value }) -join "`r`n"
$txt
try { $txt | Set-Clipboard; '' ; '(also copied to your clipboard)' } catch { }
