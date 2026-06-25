<#
.SYNOPSIS
    Removes the ZebulonVSTO PowerPoint add-in for the current user.

.DESCRIPTION
    Unregisters the add-in (HKCU), deletes the installed files, and best-effort
    removes the bundled signing certificate from the current user's stores.
    Verifies removal and reports the result in a dialog box. Per-user only - no
    administrator rights required.

    NOTE: messages are intentionally ASCII/English (see Install.ps1). Korean
    guidance lives in README.txt.
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$packageRoot = $PSScriptRoot
$target      = Join-Path $env:LOCALAPPDATA 'ZebulonVSTO'
$cerPath     = Join-Path (Join-Path $packageRoot 'ZebulonVSTO') 'ZebulonVSTO.cer'
$addinKey    = 'HKCU:\Software\Microsoft\Office\PowerPoint\Addins\ZebulonVSTO'

function Show-Result([string] $title, [string] $message, [bool] $success) {
    $color = if ($success) { 'Green' } else { 'Red' }
    Write-Host ''
    Write-Host ('=' * 60) -ForegroundColor $color
    Write-Host $message -ForegroundColor $color
    Write-Host ('=' * 60) -ForegroundColor $color
    try {
        Add-Type -AssemblyName System.Windows.Forms
        $icon = if ($success) { [System.Windows.Forms.MessageBoxIcon]::Information } else { [System.Windows.Forms.MessageBoxIcon]::Warning }
        [void][System.Windows.Forms.MessageBox]::Show($message, $title, [System.Windows.Forms.MessageBoxButtons]::OK, $icon)
    } catch { }
}

try {
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
    $filesGone = $true
    if (Test-Path $target) {
        try {
            Remove-Item -Path $target -Recurse -Force
            Write-Host "  Deleted $target."
        } catch {
            $filesGone = $false
            Write-Warning "Could not delete $target (PowerPoint may still hold the DLL). Close PowerPoint and re-run."
        }
    } else {
        Write-Host '  Install folder not found (already removed).'
    }

    # --- best-effort: remove the bundled certificate from CurrentUser stores ---
    if (Test-Path $cerPath) {
        Write-Host '  Removing signing certificate. Windows may prompt - click Yes to actually remove the trusted certificate.' -ForegroundColor Yellow
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

    # --- verify ---
    $checks = [ordered]@{
        'Registry entry removed' = -not (Test-Path $addinKey)
        'Installed files removed' = $filesGone -and (-not (Test-Path $target))
    }
    Write-Host ''
    Write-Host 'Verification:'
    foreach ($k in $checks.Keys) {
        Write-Host ("  [{0}] {1}" -f $(if ($checks[$k]) { 'OK' } else { 'XX' }), $k) -ForegroundColor $(if ($checks[$k]) { 'Green' } else { 'Red' })
    }

    if ($checks.Values -notcontains $false) {
        Show-Result 'ZebulonVSTO - Uninstall complete' "UNINSTALL COMPLETE.`n`nThe add-in was unregistered and its files removed. Restart PowerPoint if it was open." $true
    } else {
        $failed = ($checks.Keys | Where-Object { -not $checks[$_] }) -join ', '
        Show-Result 'ZebulonVSTO - Uninstall incomplete' "UNINSTALL INCOMPLETE.`n`nRemaining: $failed`n`nClose PowerPoint and run Uninstall.ps1 again." $false
        exit 1
    }
} catch {
    Show-Result 'ZebulonVSTO - Uninstall failed' "UNINSTALL FAILED.`n`n$($_.Exception.Message)" $false
    exit 1
}
