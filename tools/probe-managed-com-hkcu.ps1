<#
.SYNOPSIS
  Answer one question in five seconds: can a *managed* (.NET) COM server be activated from a
  per-user HKCU registration on this machine?

.DESCRIPTION
  This decides WordTab's architecture, so it is a measurement rather than an assumption.

  WordTab must install without admin rights, which means registering under HKCU rather than HKLM
  (see the project memory for the no-admin constraint). The original plan was a .NET Framework
  COM add-in registered that way. On the dev rig that plan does not work, and this script is the
  reproduction.

  Three tests, each activating a COM class and reporting the raw HRESULT:

    1. NATIVE  from HKCU - a real Windows DLL under a made-up CLSID. Expected to reach the DLL
       and be refused by it (CLASS_E_CLASSNOTAVAILABLE, 0x80040111). That refusal is a PASS: it
       proves COM read HKCU, found the server and called into it.
    2. MANAGED from HKCU - a framework class (registered here by us under a made-up CLSID, with
       complete and correct values, pointing at mscoree.dll like regasm would).
    3. MANAGED from HKLM - the same framework class via the registration .NET itself installed.

  Test 3 is the control: it proves managed COM works at all on this machine. If 1 and 3 pass and
  2 fails, the conclusion is specific and actionable - the CLR's COM activation path does not see
  per-user registrations, so WordTab's COM entry point cannot be a managed DLL.

  Activation runs under Windows PowerShell on purpose: it is a native .NET Framework host, the
  same runtime a Framework add-in would load into.

  Writes only two temporary keys under HKCU and removes them again. Needs no admin rights, so it
  is safe to run on the work rig - which is the point, since the work rig is the target and may
  not behave like the dev rig.

.EXAMPLE
  pwsh -File tools\probe-managed-com-hkcu.ps1
  powershell -File tools\probe-managed-com-hkcu.ps1
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$NativeClsid  = '{11111111-2222-3333-4444-555555555555}'
$ManagedClsid = '{22222222-3333-4444-5555-666666666666}'
$WinPs        = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'

function Invoke-Activation($clsid) {
    $code = "try { [Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'$clsid',`$true)) | Out-Null; 'OK 0x00000000' } " +
            "catch { `$i = `$_.Exception.InnerException; if (`$i) { ('ERR 0x{0:X8}' -f `$i.HResult) } else { 'ERR (no hresult) ' + `$_.Exception.Message } }"
    (& $WinPs -NoProfile -NonInteractive -Command $code) -join ' '
}

function Remove-TestKey($clsid) {
    try { [Microsoft.Win32.Registry]::CurrentUser.DeleteSubKeyTree("Software\Classes\CLSID\$clsid", $false) } catch { }
}

function New-InprocKey($clsid, [hashtable]$values) {
    $key = "HKCU:\Software\Classes\CLSID\$clsid\InprocServer32"
    New-Item -Path $key -Force | Out-Null
    foreach ($name in $values.Keys) {
        Set-ItemProperty -Path $key -Name $name -Value $values[$name]
    }
}

Write-Host ''
Write-Host "Managed COM from HKCU - probe" -ForegroundColor Cyan
Write-Host "  host      : $env:COMPUTERNAME"
Write-Host "  64-bit ps : $([Environment]::Is64BitProcess)"
Write-Host ''

if (-not [Environment]::Is64BitProcess) {
    throw 'Run this from 64-bit PowerShell; a 32-bit host uses a different registry view than x64 Word.'
}

$results = [ordered]@{}

try {
    # -- 1. native server, per-user registration ------------------------------------------------
    # scrrun.dll is a real in-process COM server that ships with Windows. It will not recognise
    # our invented CLSID, and that is exactly what makes it a good probe: only a DLL that was
    # actually found, loaded and called can refuse us.
    New-InprocKey $NativeClsid @{
        '(default)'      = (Join-Path $env:SystemRoot 'System32\scrrun.dll')
        'ThreadingModel' = 'Both'
    }
    $results['native from HKCU'] = Invoke-Activation $NativeClsid

    # -- 2. managed server, per-user registration -----------------------------------------------
    # A complete, correct managed registration in regasm's shape. mscorlib lives in the GAC, so
    # no CodeBase is needed and assembly resolution cannot be blamed for a failure here.
    New-InprocKey $ManagedClsid @{
        '(default)'      = 'mscoree.dll'
        'ThreadingModel' = 'Both'
        'Class'          = 'System.Collections.SortedList'
        'Assembly'       = 'mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'
        'RuntimeVersion' = 'v4.0.30319'
    }
    $results['managed from HKCU'] = Invoke-Activation $ManagedClsid

    # -- 3. managed server, machine registration (the control) ----------------------------------
    # Discovered rather than hardcoded: .NET registers a large set of mscorlib types for COM when
    # it installs, but which CLSIDs exist varies by machine and framework version.
    $control = $null
    $root = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey('SOFTWARE\Classes\CLSID')
    if ($root) {
        foreach ($name in $root.GetSubKeyNames()) {
            try {
                $sub = $root.OpenSubKey("$name\InprocServer32")
                if ($sub) {
                    if (("$($sub.GetValue(''))" -like '*mscoree.dll') -and ("$($sub.GetValue('Assembly'))" -like 'mscorlib,*')) {
                        $control = $name
                        $sub.Close()
                        break
                    }
                    $sub.Close()
                }
            } catch { }
        }
        $root.Close()
    }
    $results['managed from HKLM (control)'] =
        if ($control) { Invoke-Activation $control } else { 'SKIPPED (no HKLM-registered managed class found)' }
}
finally {
    Remove-TestKey $NativeClsid
    Remove-TestKey $ManagedClsid
}

foreach ($name in $results.Keys) {
    $value = $results[$name]
    $colour = if ($value -like 'OK*' -or $value -like '*0x80040111*') { 'Green' } else { 'Yellow' }
    Write-Host ("  {0,-28} {1}" -f $name, $value) -ForegroundColor $colour
}

Write-Host ''
$nativeOk  = $results['native from HKCU'] -like '*0x80040111*'
$managedCu = $results['managed from HKCU'] -like 'OK*'
$managedLm = $results['managed from HKLM (control)'] -like 'OK*'

if ($nativeOk -and $managedLm -and -not $managedCu) {
    Write-Host 'VERDICT: managed COM does NOT activate from HKCU on this machine; native does.' -ForegroundColor Red
    Write-Host '         WordTab''s COM entry point must be a native DLL. This is the dev-rig result.' -ForegroundColor Red
} elseif ($managedCu) {
    Write-Host 'VERDICT: managed COM DOES activate from HKCU here. This machine differs from the dev rig -' -ForegroundColor Green
    Write-Host '         a .NET add-in registered per-user would load. Worth re-opening the architecture.' -ForegroundColor Green
} else {
    Write-Host 'VERDICT: inconclusive - read the three results above before drawing any conclusion.' -ForegroundColor Yellow
}
Write-Host ''
