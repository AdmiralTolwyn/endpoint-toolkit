<#
.SYNOPSIS
    Downloads Microsoft Store Stub App payloads via winget for AVD Golden Image baking.

.DESCRIPTION
    Downloads full Store app payloads for offline provisioning on reference
    images, including images where inbox apps are initially provisioned as
    stubs. Stage the packages, dependencies, and licenses inside the image and
    provision them during image build. Validate per-user registration and launch
    on the target Windows build; downloading alone does not repair existing users.

    This script is intended to be run LOCALLY on an interactive workstation,
    with WinGet Entra ID authentication when retrieving offline licenses.
    The account needs a supported role (License Administrator, User
    Administrator, or Global Administrator). It is NOT a host runtime script.

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

.PARAMETER SkipLicense
    Omit offline license retrieval. Removes the license authorization requirement.
    Use only for apps that permit provisioning without an offline license.

.NOTES
    File:    avd/scripts/Get-StubAppPayloads.ps1
    Author:  Anton Romanyuk
    Version: 1.0.0
    Context: Run locally (interactive). Requires winget with download support.

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
    [ValidateNotNullOrEmpty()]
    [string]$DownloadPath  = 'C:\Temp\AVD_Stubs_Payload',
    [ValidateNotNullOrEmpty()]
    [string]$ManifestPath  = (Join-Path $PSScriptRoot 'StubApps.json'),
    [string]$Architecture,
    [string]$Source,
    [switch]$SkipLicense
)

$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $false

# -----------------------------------------------------------------------------
# HELPERS
# -----------------------------------------------------------------------------
function Write-Log {
<#
.SYNOPSIS
    Writes a colour-coded, timestamped, level-tagged line to the console.
.DESCRIPTION
    Lightweight console logger used throughout the script. Format:
        [HH:mm:ss] [LEVEL] message
    Level controls the foreground colour (INFO/grey, WARN/yellow, ERROR/red,
    SUCCESS/green, HEADER/cyan). No file output - this script is interactive.
.PARAMETER Message
    Free-form text to print.
.PARAMETER Level
    INFO | WARN | ERROR | SUCCESS | HEADER. Default INFO.
#>
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
<#
.SYNOPSIS
    Verifies winget.exe is installed and reachable on PATH.
.DESCRIPTION
    Throws a terminating error when winget is missing so the caller stops
    before attempting any download. Required because this script depends
    entirely on `winget download --source msstore`.
#>
    $winget = Get-Command winget.exe -CommandType Application -ErrorAction SilentlyContinue
    if (-not $winget) {
        throw "winget.exe not found in PATH. Install App Installer from the Microsoft Store."
    }
    Write-Log "winget located: $($winget.Source)" -Level INFO
    $helpText = (& $winget.Source download --help | Out-String)
    if ($LASTEXITCODE -ne 0 -or $helpText -notmatch '--download-directory' -or
        $helpText -notmatch '--skip-license') {
        throw 'Update App Installer: winget must support download and --skip-license.'
    }
    return $winget.Source
}

function Read-StubManifest {
<#
.SYNOPSIS
    Loads and validates the StubApps.json manifest.
.DESCRIPTION
    Reads the JSON file at $Path, parses it, and returns the resulting object.
    The manifest must contain an `apps` array with at least one entry; the
    optional `defaults` block lets callers set `source` / `architecture` once.
    Throws a terminating error on missing file, invalid JSON, or empty `apps`.
.PARAMETER Path
    Full path to the JSON manifest.
.OUTPUTS
    PSCustomObject parsed from the JSON.
#>
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
    if ($json.apps -isnot [array] -or $json.apps.Count -eq 0) {
        throw "Manifest '$Path' must contain a nonempty apps array."
    }
    $ids = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $names = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($app in $json.apps) {
        if ($app.Id -isnot [string] -or [string]::IsNullOrWhiteSpace($app.Id) -or $app.Id -match '^\s|\s$|["\r\n]') {
            throw 'Every manifest app must have a nonempty string Id without quotes or surrounding whitespace.'
        }
        if ($app.Name -isnot [string] -or [string]::IsNullOrWhiteSpace($app.Name) -or
            $app.Name.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0 -or
            $app.Name -match '(^\s|[. ]$|^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])($|\.))') {
            throw "Manifest app '$($app.Id)' must have a valid Windows folder name."
        }
        if (-not $ids.Add($app.Id) -or -not $names.Add($app.Name)) {
            throw "Duplicate manifest app Id or Name: $($app.Id) / $($app.Name)"
        }
    }
    return $json
}

# -----------------------------------------------------------------------------
# MAIN
# -----------------------------------------------------------------------------
Write-Log "*** AVD Stub App Payload Downloader ***" -Level HEADER
Write-Log "Manifest      : $ManifestPath"
Write-Log "DownloadPath  : $DownloadPath"

$manifest = Read-StubManifest -Path $ManifestPath

# Resolve effective defaults: explicit param > manifest defaults > hard default
$effectiveSource = if ($Source)       { $Source }
                   elseif ($manifest.defaults.source)       { $manifest.defaults.source }
                   else                                     { 'msstore' }
$effectiveArch   = if ($Architecture) { $Architecture }
                   elseif ($manifest.defaults.architecture) { $manifest.defaults.architecture }
                   else                                     { 'x64' }

if ($effectiveArch -notin @('x86', 'x64', 'arm', 'arm64')) {
    throw "Unsupported architecture: $effectiveArch"
}
if ($effectiveSource -isnot [string] -or [string]::IsNullOrWhiteSpace($effectiveSource)) {
    throw 'Source must be a nonempty string.'
}
$wingetPath = Test-Prerequisite

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

    $wingetArgs = @(
        'download'
        '--id',                 $app.Id
        '--exact'
        '--download-directory', $targetDir
        '--source',             $effectiveSource
        '--architecture',       $effectiveArch
        '--accept-package-agreements'
        '--accept-source-agreements'
    )
    if ($SkipLicense) { $wingetArgs += '--skip-license' }

    $exitCode = -1
    try {
        New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
        Write-Log "Running: winget $($wingetArgs -join ' ')"
        & $wingetPath @wingetArgs | Out-Host
        $exitCode = $LASTEXITCODE
    }
    catch {
        Write-Log "Download '$($app.Name)' failed: $($_.Exception.Message)" -Level ERROR
    }

    if ($exitCode -eq 0) {
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
        Write-Log "winget exited with code $exitCode for $($app.Name)" -Level ERROR
        $results.Add([pscustomobject]@{
            Name     = $app.Name
            Id       = $app.Id
            ExitCode = $exitCode
            Status   = 'Failed'
            Path     = $targetDir
        })
    }
}

# -----------------------------------------------------------------------------
# SUMMARY
# -----------------------------------------------------------------------------
$ok   = @($results | Where-Object Status -EQ 'Success').Count
$fail = @($results | Where-Object Status -EQ 'Failed').Count

Write-Log "*** DOWNLOAD COMPLETE ***" -Level HEADER
Write-Log "Succeeded: $ok / $($results.Count)" -Level $(if ($fail -eq 0) { 'SUCCESS' } else { 'WARN' })
if ($fail -gt 0) {
    Write-Log "Failed:    $fail" -Level ERROR
    $results | Where-Object Status -EQ 'Failed' |
        Format-Table Name, Id, ExitCode -AutoSize | Out-String | Write-Host
}

if ($fail -eq 0) {
    Write-Log "Zip '$DownloadPath' and add it to your Packer file provisioner." -Level INFO
}
else {
    Write-Log 'Resolve failed downloads before using this payload tree in an image.' -Level WARN
}

# Emit results object for pipeline / programmatic callers
$results
if ($fail -gt 0) { exit 1 } else { exit 0 }
