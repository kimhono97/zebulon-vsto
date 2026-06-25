<#
.SYNOPSIS
    Removes the ZebulonVSTO PowerPoint add-in for the current user.

.DESCRIPTION
    Unregisters the add-in (HKCU), deletes the installed files, and best-effort
    removes the bundled signing certificate from the current user's stores.
    Per-user only — no administrator rights required.
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$packageRoot = $PSScriptRoot
$target      = Join-Path $env:LOCALAPPDATA 'ZebulonVSTO'
$cerPath     = Join-Path (Join-Path $packageRoot 'ZebulonVSTO') 'ZebulonVSTO.cer'
$addinKey    = 'HKCU:\Software\Microsoft\Office\PowerPoint\Addins\ZebulonVSTO'

Write-Host 'ZebulonVSTO uninstaller (current user)' -ForegroundColor Cyan

if (Get-Process -Name POWERPNT -ErrorAction SilentlyContinue) {
    Write-Warning 'PowerPoint is running. Close it so the add-in unloads and files are not locked.'
}

# --- unregister ---
if (Test-Path $addinKey) {
    Remove-Item -Path $addinKey -Recurse -Force
    Write-Host '  Removed HKCU add-in registration.'
} else {
    Write-Host '  Add-in registration not found (already removed).'
}

# --- delete installed files ---
if (Test-Path $target) {
    try {
        Remove-Item -Path $target -Recurse -Force
        Write-Host "  Deleted $target."
    } catch {
        Write-Warning "Could not delete $target (PowerPoint may still hold the DLL). Close PowerPoint and re-run."
    }
} else {
    Write-Host '  Install folder not found (already removed).'
}

# --- best-effort: remove the bundled certificate from CurrentUser stores ---
if (Test-Path $cerPath) {
    Write-Host '  Removing signing certificate. Windows may prompt to confirm - click Yes to actually remove the trusted certificate.' -ForegroundColor Yellow
    try {
        $thumb = (New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 $cerPath).Thumbprint
        foreach ($storeName in @('Root', 'TrustedPublisher')) {
            $store = New-Object System.Security.Cryptography.X509Certificates.X509Store($storeName, 'CurrentUser')
            $store.Open('ReadWrite')
            $match = $store.Certificates | Where-Object { $_.Thumbprint -eq $thumb }
            foreach ($c in $match) { $store.Remove($c) }
            $store.Close()
        }
        Write-Host '  Removed signing certificate from CurrentUser stores.'
    } catch {
        Write-Warning "Could not remove the certificate (harmless if left): $($_.Exception.Message)"
    }
}

Write-Host 'Done.' -ForegroundColor Green
