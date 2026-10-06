<#
.SYNOPSIS
    Installs the ZebulonVSTO PowerPoint add-in for the current user (no admin).
    Re-run to update after extracting a newer package over the old one.

.DESCRIPTION
    Trusts the bundled self-signed certificate, copies the add-in to
    %LOCALAPPDATA%\ZebulonVSTO, and registers it under HKCU so PowerPoint loads
    it at startup. Also registers a per-user Apps & Features (Installed apps)
    entry with the Zebulon icon whose Uninstall button runs the installed copy
    of Uninstall.ps1. Verifies each step and reports the result in a dialog box so
    the outcome is clear even when the window closes on completion. Per-user
    only - no administrator rights required.

    NOTE: messages are intentionally ASCII/English - Windows PowerShell 5.1
    misreads a BOM-less UTF-8 .ps1 as the system code page, which would garble
    non-ASCII text. Korean guidance lives in README.txt.
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$packageRoot = $PSScriptRoot
$source      = Join-Path $packageRoot 'ZebulonVSTO'
$target      = Join-Path $env:LOCALAPPDATA 'ZebulonVSTO'
$cerPath     = Join-Path $source 'ZebulonVSTO.cer'
$vstoName    = 'ZebulonVSTO.vsto'
$addinKey    = 'HKCU:\Software\Microsoft\Office\PowerPoint\Addins\ZebulonVSTO'
$appsKey     = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\ZebulonVSTO'
$uninstaller = Join-Path $packageRoot 'Uninstall.ps1'
$required    = @('ZebulonVSTO.dll', 'ZebulonVSTO.vsto', 'ZebulonVSTO.dll.manifest')

function Show-Result([string] $title, [string] $message, [bool] $success) {
    $color = if ($success) { 'Green' } else { 'Red' }
    Write-Host ''
    Write-Host ('=' * 60) -ForegroundColor $color
    Write-Host $message -ForegroundColor $color
    Write-Host ('=' * 60) -ForegroundColor $color
    try {
        Add-Type -AssemblyName System.Windows.Forms
        $icon = if ($success) { [System.Windows.Forms.MessageBoxIcon]::Information } else { [System.Windows.Forms.MessageBoxIcon]::Error }
        [void][System.Windows.Forms.MessageBox]::Show($message, $title, [System.Windows.Forms.MessageBoxButtons]::OK, $icon)
    } catch { }
}

function Test-CertInStore([string] $storeName, [string] $thumb) {
    $s = New-Object System.Security.Cryptography.X509Certificates.X509Store($storeName, 'CurrentUser')
    $s.Open('ReadOnly')
    $found = ($s.Certificates | Where-Object { $_.Thumbprint -eq $thumb }).Count -gt 0
    $s.Close()
    return $found
}

try {
    Write-Host 'ZebulonVSTO installer (current user)' -ForegroundColor Cyan

    if (-not (Test-Path $source)) {
        throw "Cannot find the 'ZebulonVSTO' folder next to this script. Extract the whole package and run Install.ps1 from its root."
    }
    foreach ($f in $required) {
        if (-not (Test-Path (Join-Path $source $f))) { throw "Missing required file: $f" }
    }
    if (-not (Test-Path $uninstaller)) { throw 'Missing required file: Uninstall.ps1 (next to Install.ps1)' }
    if (Get-Process -Name POWERPNT -ErrorAction SilentlyContinue) {
        Write-Warning 'PowerPoint is running. Restart it after install so the add-in (re)loads.'
    }

    # --- VSTO 2010 runtime check (best effort; warn only) ---
    $vstoInstalled = $false
    foreach ($p in @(
        'HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R',
        'HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4',
        'HKLM:\SOFTWARE\Wow6432Node\Microsoft\VSTO Runtime Setup\v4R',
        'HKLM:\SOFTWARE\Wow6432Node\Microsoft\VSTO Runtime Setup\v4')) {
        if (Test-Path $p) { $vstoInstalled = $true; break }
    }
    if (-not $vstoInstalled) {
        Write-Warning 'VSTO 2010 Runtime not detected. If the add-in does not load, install "Microsoft Visual Studio 2010 Tools for Office Runtime" (microsoft.com/download id 48217).'
    }

    # --- trust the bundled self-signed certificate (CurrentUser) ---
    $certThumb = $null
    if (Test-Path $cerPath) {
        Write-Host '  Trusting signing certificate (CurrentUser Root + TrustedPublisher)...'
        Write-Host '  Windows may prompt to confirm trusting the certificate - click Yes.' -ForegroundColor Yellow
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 $cerPath
        $certThumb = $cert.Thumbprint
        foreach ($storeName in @('Root', 'TrustedPublisher')) {
            $store = New-Object System.Security.Cryptography.X509Certificates.X509Store($storeName, 'CurrentUser')
            $store.Open('ReadWrite'); $store.Add($cert); $store.Close()
        }
    } else {
        Write-Warning "Certificate $cerPath not found; skipping trust step."
    }

    # --- copy the add-in + the uninstaller (the .cer too, so the installed
    #     Uninstall.ps1 can find the cert to untrust when run from Settings) ---
    Write-Host "  Copying add-in to $target ..."
    New-Item -ItemType Directory -Path $target -Force | Out-Null
    Get-ChildItem -Path $source -File | ForEach-Object {
        Copy-Item -Path $_.FullName -Destination $target -Force
    }
    Copy-Item -Path $uninstaller -Destination $target -Force

    # --- register the add-in under HKCU ---
    Write-Host '  Registering the add-in (HKCU)...'
    New-Item -Path $addinKey -Force | Out-Null
    Set-ItemProperty -Path $addinKey -Name 'Manifest'     -Value ((Join-Path $target $vstoName) + '|vstolocal')
    Set-ItemProperty -Path $addinKey -Name 'LoadBehavior' -Value 3 -Type DWord
    Set-ItemProperty -Path $addinKey -Name 'FriendlyName' -Value 'ZebulonVSTO'
    Set-ItemProperty -Path $addinKey -Name 'Description'  -Value 'UDP slide synchronization add-in'

    # --- register the Apps & Features (Installed apps) entry (HKCU) ---
    Write-Host '  Registering in Settings > Apps (HKCU)...'
    $dllVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo((Join-Path $target 'ZebulonVSTO.dll')).FileVersion
    $sizeKB     = [int][Math]::Ceiling(((Get-ChildItem -Path $target -File | Measure-Object -Property Length -Sum).Sum) / 1KB)
    $psExe      = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    New-Item -Path $appsKey -Force | Out-Null
    Set-ItemProperty -Path $appsKey -Name 'DisplayName'     -Value 'ZebulonVSTO'
    Set-ItemProperty -Path $appsKey -Name 'DisplayVersion'  -Value $dllVersion
    Set-ItemProperty -Path $appsKey -Name 'Publisher'       -Value 'kimhono97'
    Set-ItemProperty -Path $appsKey -Name 'DisplayIcon'     -Value (Join-Path $target 'zebulon.ico')
    Set-ItemProperty -Path $appsKey -Name 'UninstallString' -Value ('"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}"' -f $psExe, (Join-Path $target 'Uninstall.ps1'))
    Set-ItemProperty -Path $appsKey -Name 'InstallLocation' -Value $target
    Set-ItemProperty -Path $appsKey -Name 'InstallDate'     -Value (Get-Date -Format 'yyyyMMdd')
    Set-ItemProperty -Path $appsKey -Name 'URLInfoAbout'    -Value 'https://github.com/kimhono97/zebulon-vsto'
    Set-ItemProperty -Path $appsKey -Name 'EstimatedSize'   -Value $sizeKB -Type DWord
    Set-ItemProperty -Path $appsKey -Name 'NoModify'        -Value 1 -Type DWord
    Set-ItemProperty -Path $appsKey -Name 'NoRepair'        -Value 1 -Type DWord

    # --- verify ---
    $checks = [ordered]@{
        'Registry entry (LoadBehavior=3)' = (Test-Path $addinKey) -and ((Get-ItemProperty $addinKey).LoadBehavior -eq 3)
        'Add-in DLL copied'               = Test-Path (Join-Path $target 'ZebulonVSTO.dll')
        'Manifest copied'                 = Test-Path (Join-Path $target $vstoName)
        'Uninstaller copied'              = Test-Path (Join-Path $target 'Uninstall.ps1')
        'Apps & Features entry'           = (Test-Path $appsKey) -and ((Get-ItemProperty $appsKey).DisplayVersion -eq $dllVersion)
    }
    if ($certThumb) {
        $checks['Cert trusted (Root)']            = Test-CertInStore 'Root' $certThumb
        $checks['Cert trusted (TrustedPublisher)'] = Test-CertInStore 'TrustedPublisher' $certThumb
    }

    Write-Host ''
    Write-Host 'Verification:'
    foreach ($k in $checks.Keys) {
        Write-Host ("  [{0}] {1}" -f $(if ($checks[$k]) { 'OK' } else { 'XX' }), $k) -ForegroundColor $(if ($checks[$k]) { 'Green' } else { 'Red' })
    }

    if ($checks.Values -notcontains $false) {
        Show-Result 'ZebulonVSTO - Install complete' "INSTALL COMPLETE.`n`nInstalled to: $target`nRegistered:   HKCU\...\PowerPoint\Addins\ZebulonVSTO`nRemove later: Settings > Apps > Installed apps > ZebulonVSTO`n`nStart PowerPoint - the 'Zebulon' tab appears under the Add-Ins ribbon." $true
    } else {
        $failed = ($checks.Keys | Where-Object { -not $checks[$_] }) -join ', '
        Show-Result 'ZebulonVSTO - Install not verified' "INSTALL NOT FULLY VERIFIED.`n`nFailed checks: $failed`n`n(If only the cert checks failed, you may have declined the Windows trust prompt - re-run and click Yes.)" $false
        exit 1
    }
} catch {
    Show-Result 'ZebulonVSTO - Install failed' "INSTALL FAILED.`n`n$($_.Exception.Message)" $false
    exit 1
}
