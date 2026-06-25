<#
.SYNOPSIS
    Builds the ZebulonVSTO add-in in Release and assembles a self-contained
    deployment zip under dist/.

.DESCRIPTION
    VSTO add-ins build only with MSBuild + the Office build targets (not the
    `dotnet` CLI, and not on hosted CI runners), so release packaging is a local
    step. This script: builds Release, exports the signing certificate's public
    .cer, and bundles the Release output + deploy/ templates + the Tools scripts
    (from tests/manual) into dist/ZebulonVSTO-<version>.zip.

.PARAMETER Configuration
    Build configuration. Default Release.
#>
[CmdletBinding()]
param(
    [string] $Configuration = 'Release'
)

$ErrorActionPreference = 'Stop'
$repo = $PSScriptRoot

# --- locate MSBuild (via vswhere; -find is unreliable so derive from installationPath) ---
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
$msbuild = $null
if (Test-Path $vswhere) {
    $inst = & $vswhere -latest -property installationPath
    if ($inst) {
        $candidate = Join-Path $inst 'MSBuild\Current\Bin\MSBuild.exe'
        if (Test-Path $candidate) { $msbuild = $candidate }
    }
}
if (-not $msbuild) { throw 'MSBuild not found. Install Visual Studio with the Office/VSTO workload.' }

# --- build Release (VisualStudioVersion=10.0 required for the VSTO targets) ---
Write-Host "Building $Configuration ..." -ForegroundColor Cyan
& $msbuild (Join-Path $repo 'ZebulonVSTO.sln') -t:Rebuild -p:Configuration=$Configuration -p:VisualStudioVersion=10.0 -v:minimal -nologo
if ($LASTEXITCODE -ne 0) { throw "Build failed (exit $LASTEXITCODE)." }

$bin = Join-Path $repo "ZebulonVSTO\bin\$Configuration"
# The signed deployment manifest requires EVERY dependency DLL to be present
# (e.g. Microsoft.Office.Tools.Common.v4.0.Utilities.dll), so bundle all
# assemblies from the build output, not just the main one.
$required = @('ZebulonVSTO.dll', 'ZebulonVSTO.vsto', 'ZebulonVSTO.dll.manifest', 'ZebulonVSTO.dll.config')
foreach ($f in $required) {
    if (-not (Test-Path (Join-Path $bin $f))) { throw "Expected build output missing: $f (in $bin)" }
}

$version = [System.Diagnostics.FileVersionInfo]::GetVersionInfo((Join-Path $bin 'ZebulonVSTO.dll')).FileVersion
Write-Host "  built version $version"

# --- export the signing certificate's PUBLIC .cer (for trust on target machines) ---
# Read the thumbprint from the csproj so this never drifts from what actually signs.
$csprojText = Get-Content (Join-Path $repo 'ZebulonVSTO\ZebulonVSTO.csproj') -Raw
if ($csprojText -notmatch '<ManifestCertificateThumbprint>([0-9A-Fa-f]+)</ManifestCertificateThumbprint>') {
    throw 'Could not read ManifestCertificateThumbprint from the csproj.'
}
$thumb = $matches[1]
$cert = Get-Item "Cert:\CurrentUser\My\$thumb" -ErrorAction SilentlyContinue
if (-not $cert) {
    $pfx = Join-Path $repo 'ZebulonVSTO\ZebulonVSTO_TemporaryKey.pfx'
    try {
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 $pfx, ''
    } catch {
        throw "Could not obtain the signing certificate (not in Cert:\CurrentUser\My\$thumb and the pfx needs a password). Open the solution in VS once to import the temp key, then retry."
    }
}

# --- stage the package ---
$dist  = Join-Path $repo 'dist'
$stage = Join-Path $dist "ZebulonVSTO-$version"
if (Test-Path $stage) { Remove-Item $stage -Recurse -Force }
$addinDir = Join-Path $stage 'ZebulonVSTO'
$toolsDir = Join-Path $stage 'Tools'
New-Item -ItemType Directory -Path $addinDir, $toolsDir -Force | Out-Null

Copy-Item (Join-Path $bin '*.dll') $addinDir -Force
foreach ($f in @('ZebulonVSTO.vsto', 'ZebulonVSTO.dll.manifest', 'ZebulonVSTO.dll.config')) {
    Copy-Item (Join-Path $bin $f) $addinDir -Force
}
[System.IO.File]::WriteAllBytes((Join-Path $addinDir 'ZebulonVSTO.cer'),
    $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert))

Copy-Item (Join-Path $repo 'tests\manual\Send-SyncCommand.ps1')  $toolsDir -Force
Copy-Item (Join-Path $repo 'tests\manual\Start-SyncSession.ps1') $toolsDir -Force

foreach ($f in @('Install.ps1', 'Uninstall.ps1', 'README.txt', 'AGENTS.md')) {
    Copy-Item (Join-Path $repo "deploy\$f") $stage -Force
}

# --- zip it ---
$zip = Join-Path $dist "ZebulonVSTO-$version.zip"
if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $zip

Write-Host "Packaged: $zip" -ForegroundColor Green
Get-ChildItem -Recurse $stage | ForEach-Object { '  ' + $_.FullName.Substring($stage.Length + 1) }
