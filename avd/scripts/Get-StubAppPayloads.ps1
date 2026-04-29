<#
.SYNOPSIS
    Downloads Microsoft Store Stub App payloads via winget for AVD Golden Image baking.

.DESCRIPTION
    Fixes the well-known "Stub App" provisioning issue on multi-session / shared
    Windows images where some inbox Store apps ship as stubs and never finish
    provisioning for new users. The workaround is to pre-stage the offline
    .msixbundle / .appxbundle + license files inside the image and side-load them
    during Packer image build (or first-boot script).

    This script is intended to be run LOCALLY on an interactive workstation,
    signed in to the Microsoft Store with an Entra ID account that has rights
    to acquire the listed packages. It is NOT a session-host runtime script.

    Workflow:
      1. Loads the app list from a JSON manifest (default: .\StubApps.json).
      2. For each app, runs `winget download --source msstore` into a per-app
         subfolder under -DownloadPath.
      3. Reports per-app success / failure and a final summary.
      4. The resulting folder is meant to be zipped and added to your Packer
         file provisioner (or Image Builder customizer) so the offline payloads
         travel with the image.

.PARAMETER DownloadPath
    Root folder for downloaded payloads. Default: C:\Temp\AVD_Stubs_Payload.

.PARAMETER ManifestPath
    Path to the JSON manifest describing the apps to download.
    Default: .\StubApps.json next to this script.

.PARAMETER Architecture
    Target architecture passed to winget (--architecture). Default: x64.
    Override per-image (e.g. arm64) if needed.

.PARAMETER Source
    winget source to query. Default: msstore. Override only if you have a
    private REST source mirroring Store packages.

.NOTES
    File:    avd/scripts/Get-StubAppPayloads.ps1
    Author:  Anton Romanyuk
    Version: 1.0.0
    Context: Run locally with Entra ID auth (interactive). Requires winget 1.6+.

.DISCLAIMER
    This script is provided "AS IS" with no warranties and confers no rights.
    It is not supported under any Microsoft standard support program or service.
    Use of this script is entirely at your own risk. The customer is solely
    responsible for testing and validating this script in their environment
    before deploying to production.

.EXAMPLE
    # Default run — downloads to C:\Temp\AVD_Stubs_Payload using StubApps.json
    .\Get-StubAppPayloads.ps1

.EXAMPLE
    # Custom download root + custom manifest
    .\Get-StubAppPayloads.ps1 -DownloadPath D:\ImageBuild\Stubs -ManifestPath .\StubApps.win11-24h2.json
#>

[CmdletBinding()]
param(
    [string]$DownloadPath  = 'C:\Temp\AVD_Stubs_Payload',
    [string]$ManifestPath  = (Join-Path $PSScriptRoot 'StubApps.json'),
    [string]$Architecture,
    [string]$Source
)

$ErrorActionPreference = 'Stop'

# -----------------------------------------------------------------------------
# HELPERS
# -----------------------------------------------------------------------------
function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS','HEADER')][string]$Level = 'INFO'
    )
    $ts = (Get-Date).ToString('HH:mm:ss')
    $color = switch ($Level) {
        'INFO'    { 'Gray' }
        'WARN'    { 'Yellow' }
        'ERROR'   { 'Red' }
        'SUCCESS' { 'Green' }
        'HEADER'  { 'Cyan' }
    }
    Write-Host "[$ts] [$Level] $Message" -ForegroundColor $color
}

function Test-Prerequisite {
    $winget = Get-Command winget.exe -ErrorAction SilentlyContinue
    if (-not $winget) {
        throw "winget.exe not found in PATH. Install App Installer from the Microsoft Store."
    }
    Write-Log "winget located: $($winget.Source)" -Level INFO
}

function Read-StubManifest {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Manifest not found: $Path"
    }
    try {
        $json = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    catch {
        throw "Failed to parse manifest '$Path': $($_.Exception.Message)"
    }
    if (-not $json.apps -or $json.apps.Count -eq 0) {
        throw "Manifest '$Path' contains no apps."
    }
    return $json
}

# -----------------------------------------------------------------------------
# MAIN
# -----------------------------------------------------------------------------
Write-Log "*** AVD Stub App Payload Downloader ***" -Level HEADER
Write-Log "Manifest      : $ManifestPath"
Write-Log "DownloadPath  : $DownloadPath"

Test-Prerequisite
$manifest = Read-StubManifest -Path $ManifestPath

# Resolve effective defaults: explicit param > manifest defaults > hard default
$effectiveSource = if ($Source)       { $Source }
                   elseif ($manifest.defaults.source)       { $manifest.defaults.source }
                   else                                     { 'msstore' }
$effectiveArch   = if ($Architecture) { $Architecture }
                   elseif ($manifest.defaults.architecture) { $manifest.defaults.architecture }
                   else                                     { 'x64' }

Write-Log "Source        : $effectiveSource"
Write-Log "Architecture  : $effectiveArch"
Write-Log "App count     : $($manifest.apps.Count)"

if (-not (Test-Path -LiteralPath $DownloadPath)) {
    New-Item -Path $DownloadPath -ItemType Directory -Force | Out-Null
}

$results = New-Object System.Collections.Generic.List[object]

foreach ($app in $manifest.apps) {
    Write-Log "--- Processing: $($app.Name) ($($app.Id)) ---" -Level HEADER

    $targetDir = Join-Path $DownloadPath $app.Name
    if (-not (Test-Path -LiteralPath $targetDir)) {
        New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
    }

    $wingetArgs = @(
        'download'
        '--id',                 $app.Id
        '--download-directory', $targetDir
        '--source',             $effectiveSource
        '--architecture',       $effectiveArch
        '--accept-package-agreements'
        '--accept-source-agreements'
        '--skip-license'
    )

    Write-Log "Running: winget $($wingetArgs -join ' ')"
    $proc = Start-Process -FilePath 'winget.exe' -ArgumentList $wingetArgs `
                          -Wait -NoNewWindow -PassThru

    if ($proc.ExitCode -eq 0) {
        Write-Log "Downloaded -> $targetDir" -Level SUCCESS
        $results.Add([pscustomobject]@{
            Name     = $app.Name
            Id       = $app.Id
            ExitCode = 0
            Status   = 'Success'
            Path     = $targetDir
        })
    }
    else {
        Write-Log "winget exited with code $($proc.ExitCode) for $($app.Name)" -Level ERROR
        $results.Add([pscustomobject]@{
            Name     = $app.Name
            Id       = $app.Id
            ExitCode = $proc.ExitCode
            Status   = 'Failed'
            Path     = $targetDir
        })
    }
}

# -----------------------------------------------------------------------------
# SUMMARY
# -----------------------------------------------------------------------------
$ok   = ($results | Where-Object Status -EQ 'Success').Count
$fail = ($results | Where-Object Status -EQ 'Failed').Count

Write-Log "*** DOWNLOAD COMPLETE ***" -Level HEADER
Write-Log "Succeeded: $ok / $($results.Count)" -Level $(if ($fail -eq 0) { 'SUCCESS' } else { 'WARN' })
if ($fail -gt 0) {
    Write-Log "Failed:    $fail" -Level ERROR
    $results | Where-Object Status -EQ 'Failed' |
        Format-Table Name, Id, ExitCode -AutoSize | Out-String | Write-Host
}

Write-Log "Zip '$DownloadPath' and add it to your Packer file provisioner." -Level INFO

# Emit results object for pipeline / programmatic callers
$results
