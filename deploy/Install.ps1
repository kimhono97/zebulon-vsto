<#
.SYNOPSIS
    Installs the ZebulonVSTO PowerPoint add-in for the current user (no admin).
    Re-run to update after extracting a newer package over the old one.

.DESCRIPTION
    Trusts the bundled self-signed certificate, copies the add-in to
    %LOCALAPPDATA%\ZebulonVSTO, and registers it under HKCU so PowerPoint loads
    it at startup. Per-user only — does not require administrator rights.
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

$required = @('ZebulonVSTO.dll', 'ZebulonVSTO.vsto', 'ZebulonVSTO.dll.manifest')

Write-Host 'ZebulonVSTO installer (current user)' -ForegroundColor Cyan

# --- sanity checks ---
if (-not (Test-Path $source)) {
    throw "Cannot find the 'ZebulonVSTO' folder next to this script. Extract the whole package and run Install.ps1 from its root."
}
foreach ($f in $required) {
    if (-not (Test-Path (Join-Path $source $f))) { throw "Missing required file: $f" }
}

if (Get-Process -Name POWERPNT -ErrorAction SilentlyContinue) {
    Write-Warning 'PowerPoint is running. Close it before/after installing so the add-in (re)loads cleanly.'
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
    Write-Warning 'The VSTO 2010 Runtime was not detected. If the add-in does not load, install "Microsoft Visual Studio 2010 Tools for Office Runtime" (microsoft.com/download, id 48217).'
}

# --- trust the bundled self-signed certificate (CurrentUser) ---
if (Test-Path $cerPath) {
    Write-Host '  Trusting signing certificate (CurrentUser Root + TrustedPublisher)...'
    Write-Host '  Windows may prompt to confirm trusting the certificate - click Yes.' -ForegroundColor Yellow
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 $cerPath
    foreach ($storeName in @('Root', 'TrustedPublisher')) {
        $store = New-Object System.Security.Cryptography.X509Certificates.X509Store($storeName, 'CurrentUser')
        $store.Open('ReadWrite')
        $store.Add($cert)
        $store.Close()
    }
} else {
    Write-Warning "Certificate $cerPath not found; skipping trust step. You may get an 'unknown publisher' prompt."
}

# --- copy the add-in to the install location (all bundled files except the
#     .cer, which is import-only) ---
Write-Host "  Copying add-in to $target ..."
New-Item -ItemType Directory -Path $target -Force | Out-Null
Get-ChildItem -Path $source -File | Where-Object { $_.Extension -ne '.cer' } | ForEach-Object {
    Copy-Item -Path $_.FullName -Destination $target -Force
}

# --- register the add-in under HKCU ---
Write-Host '  Registering the add-in (HKCU)...'
New-Item -Path $addinKey -Force | Out-Null
Set-ItemProperty -Path $addinKey -Name 'Manifest'     -Value ((Join-Path $target $vstoName) + '|vstolocal')
Set-ItemProperty -Path $addinKey -Name 'LoadBehavior' -Value 3 -Type DWord
Set-ItemProperty -Path $addinKey -Name 'FriendlyName' -Value 'ZebulonVSTO'
Set-ItemProperty -Path $addinKey -Name 'Description'  -Value 'UDP slide synchronization add-in'

Write-Host 'Done. Start PowerPoint - the "Zebulon" tab appears under the Add-Ins ribbon.' -ForegroundColor Green
