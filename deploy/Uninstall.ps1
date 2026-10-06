<#
.SYNOPSIS
    Removes the ZebulonVSTO PowerPoint add-in for the current user.

.DESCRIPTION
    Unregisters the add-in (HKCU) and its Apps & Features entry, deletes the
    installed files and the saved preferences (%APPDATA%\ZebulonVSTO), and
    best-effort removes the bundled signing certificate from the current
    user's stores. Runs either from the extracted package or as the installed
    copy (Settings > Apps > Uninstall).
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
$prefsDir    = Join-Path $env:APPDATA 'ZebulonVSTO'   # preferences.json (remembered UI settings)
$addinKey    = 'HKCU:\Software\Microsoft\Office\PowerPoint\Addins\ZebulonVSTO'
$appsKey     = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\ZebulonVSTO'
# The cert sits in ZebulonVSTO\ when run from the package, or beside this
# script when run as the installed copy; fall back to the install folder.
$cerPath     = @(
    (Join-Path (Join-Path $packageRoot 'ZebulonVSTO') 'ZebulonVSTO.cer'),
    (Join-Path $packageRoot 'ZebulonVSTO.cer'),
    (Join-Path $target 'ZebulonVSTO.cer')
) | Where-Object { Test-Path $_ } | Select-Object -First 1

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

    # --- read the cert thumbprint now: the install folder may hold the only
    #     copy of the .cer and is deleted below ---
    $thumb = $null
    if ($cerPath) {
        try {
            $thumb = (New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 $cerPath).Thumbprint
        } catch {
            Write-Warning "Could not read the certificate (harmless if left): $($_.Exception.Message)"
        }
    }

    # --- unregister ---
    if (Test-Path $addinKey) {
        Remove-Item -Path $addinKey -Recurse -Force
        Write-Host '  Removed HKCU add-in registration.'
    } else {
        Write-Host '  Add-in registration not found (already removed).'
    }
    if (Test-Path $appsKey) {
        Remove-Item -Path $appsKey -Recurse -Force
        Write-Host '  Removed Settings > Apps entry.'
    }

    # --- delete installed files ---
    # Step out first: the installed copy run via Explorer starts with the
    # install folder as its working directory, which would block the delete.
    Set-Location -Path $env:TEMP
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

    # --- delete saved preferences (uninstall = full removal; updates keep them) ---
    $prefsGone = $true
    if (Test-Path $prefsDir) {
        try {
            Remove-Item -Path $prefsDir -Recurse -Force
            Write-Host "  Deleted saved preferences $prefsDir."
        } catch {
            $prefsGone = $false
            Write-Warning "Could not delete $prefsDir : $($_.Exception.Message)"
        }
    }

    # --- best-effort: remove the bundled certificate from CurrentUser stores ---
    if ($thumb) {
        Write-Host '  Removing signing certificate. Windows may prompt - click Yes to actually remove the trusted certificate.' -ForegroundColor Yellow
        try {
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
        'Apps entry removed'     = -not (Test-Path $appsKey)
        'Installed files removed' = $filesGone -and (-not (Test-Path $target))
        'Preferences removed'     = $prefsGone -and (-not (Test-Path $prefsDir))
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
