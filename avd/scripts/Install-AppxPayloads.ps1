<#
.SYNOPSIS
    Side-loads / re-provisions inbox AppX/MSIX packages from a local payload tree.

.DESCRIPTION
    Companion to Get-StubAppPayloads.ps1. Walks a payload root and provisions
    application packages for new user profiles via DISM
    (`Add-AppxProvisionedPackage`). Designed to fix two recurring image-build
    scenarios (validate provisioning and launch on your target image):

    1. STUB APPS  - Some inbox Microsoft Store apps (Photos, Clock,
         Phone Link, Xbox, Sticky Notes, ...) ship as stubs and never finish
         provisioning for new users on multi-session / shared images. Fix is
         to pre-stage the offline payloads (via the companion downloader) and
         side-load them during image bake.

      2. FoD / Language ISO APPX UPDATE  - Refresh built-in inbox apps from a
         mounted Features-on-Demand or language ISO so the image carries the
         latest signed versions. This is the original AVD scenario described
         at https://learn.microsoft.com/azure/virtual-desktop/language-packs

    Two modes:
      -Mode Install            (default)
        Provision every bundle found, regardless of whether the package is
        already present. Use this for the stub-app fix.

      -Mode UpdateProvisioned
        Only provision an app when its embedded Identity Name matches a
        DisplayName in Get-AppxProvisionedPackage. Use this when
        refreshing inbox apps from a FoD/Language ISO so you don't accidentally
        add Store apps that were never part of the base image.

    Layout assumptions (works for both winget-msstore downloads and FoD ISO trees):

      <SourcePath>\
        <AppName-or-arch>\
          *.msixbundle | *.appxbundle | *.msix | *.appx   <- main package
          <basename>.xml | <StoreId>_License.xml             <- license
          Dependencies\*.appx | *.msix                      <- frameworks

    The script:
    - Discovers bundles and classifies loose apps/frameworks from manifests
    - Passes frameworks beside/below each app via -DependencyPackagePath
    - Requests InstallFull and attaches the matching license
    - Requires explicit -SkipLicense when a main app has no license
      - Returns a result object per package and a summary at the end

.PARAMETER SourcePath
    Root folder containing the payload tree. For the stub-app workflow this is
    the folder produced by Get-StubAppPayloads.ps1 (e.g. C:\Temp\AVD_Stubs_Payload).
    For the FoD workflow this is the architecture folder on the mounted ISO
    (e.g. E:\LanguagesAndOptionalFeatures or D:\sources\<build>\amd64fre).

.PARAMETER Mode
    Install            -> provision every bundle (stub-app fix). Default.
    UpdateProvisioned  -> only refresh apps whose embedded Identity Name
                          already exists in Get-AppxProvisionedPackage.

.PARAMETER LogDirectory
    Directory for the log file. Default: $env:TEMP.

.PARAMETER SkipLicense
    Allow provisioning main apps without a license when none is found. Use only
    for apps that do not require an offline license on the target OS edition.
    A matching license is always preferred when present.

.NOTES
    File:    avd/scripts/Install-AppxPayloads.ps1
    Author:  Anton Romanyuk
    Version: 1.0.0
    Context: Run on a reference image / Image Builder VM with admin rights.
             Uses Add-AppxProvisionedPackage (DISM) so the install applies to
             every user profile created after this point, not just the current
             session.

.DISCLAIMER
    This script is provided "AS IS" with no warranties and confers no rights.
    It is not supported under any Microsoft standard support program or service.
    Use of this script is entirely at your own risk. The customer is solely
    responsible for testing and validating this script in their environment
    before deploying to production.

.EXAMPLE
    # Stub-app fix during Packer image bake
    .\Install-AppxPayloads.ps1 -SourcePath C:\BuildArtifacts\AVD_Stubs_Payload

.EXAMPLE
    # Refresh inbox apps from a mounted Features-on-Demand ISO (legacy AIB workflow)
    .\Install-AppxPayloads.ps1 `
        -SourcePath 'E:\sources\24H2\amd64fre' `
        -Mode UpdateProvisioned `
        -LogDirectory 'C:\BuildArtifacts\logs\appx'
#>

#Requires -RunAsAdministrator
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$SourcePath,

    [ValidateSet('Install','UpdateProvisioned')]
    [string]$Mode = 'Install',

    [ValidateNotNullOrEmpty()]
    [string]$LogDirectory = $env:TEMP,

    [switch]$SkipLicense
)

$ErrorActionPreference = 'Stop'

# -----------------------------------------------------------------------------
# LOGGING
# -----------------------------------------------------------------------------
$ScriptName = $MyInvocation.MyCommand.Name
$LogFile    = Join-Path $LogDirectory ("{0}_{1}.log" -f [IO.Path]::GetFileNameWithoutExtension($ScriptName), (Get-Date -Format 'yyyyMMdd_HHmmss'))

function Write-Log {
<#
.SYNOPSIS
    Writes a timestamped, level-tagged line to both the console and the log file.
.DESCRIPTION
    Uniform logger used by the rest of the script. Format on disk and on console:
        [yyyy-MM-dd HH:mm:ss] [LEVEL] message
    Console output is colour-coded by level. File writes use SilentlyContinue so a
    transient lock on the log file never aborts the cleanup pipeline.
.PARAMETER Message
    Free-form text to record.
.PARAMETER Level
    INFO | WARN | ERROR | SUCCESS | HEADER. Default INFO.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS','HEADER')][string]$Level = 'INFO'
    )
    $line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Add-Content -LiteralPath $LogFile -Value $line -ErrorAction SilentlyContinue
    $color = switch ($Level) {
        'WARN'    { 'Yellow' }
        'ERROR'   { 'Red' }
        'SUCCESS' { 'Green' }
        'HEADER'  { 'Cyan' }
        default   { 'Gray' }
    }
    Write-Host $line -ForegroundColor $color
}

# -----------------------------------------------------------------------------
# DISCOVERY
# -----------------------------------------------------------------------------

function Get-PackageMetadata {
    param([Parameter(Mandatory)][System.IO.FileInfo]$Package)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::OpenRead($Package.FullName)
    try {
        $entry = $archive.GetEntry('AppxManifest.xml')
        if (-not $entry) { $entry = $archive.GetEntry('AppxMetadata/AppxBundleManifest.xml') }
        if (-not $entry) { throw "Package manifest not found: $($Package.FullName)" }
        $stream = $entry.Open()
        try {
            $settings = [System.Xml.XmlReaderSettings]::new()
            $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
            $settings.XmlResolver = $null
            $reader = [System.Xml.XmlReader]::Create($stream, $settings)
            try {
                $document = [System.Xml.XmlDocument]::new()
                $document.XmlResolver = $null
                $document.Load($reader)
            }
            finally { $reader.Dispose() }
        }
        finally { $stream.Dispose() }
        $identity = $document.DocumentElement.SelectSingleNode('*[local-name()="Identity"]')
        if (-not $identity -or -not $identity.GetAttribute('Name')) {
            throw "Package identity not found: $($Package.FullName)"
        }
        $framework = $document.DocumentElement.SelectSingleNode('*[local-name()="Properties"]/*[local-name()="Framework"]')
        [pscustomobject]@{
            Name = $identity.GetAttribute('Name')
            IsFramework = $null -ne $framework -and $framework.InnerText -ieq 'true'
        }
    }
    finally { $archive.Dispose() }
}

function Get-Payload {
<#
.SYNOPSIS
    Recursively discovers AppX/MSIX bundles and dependency .appx files under a root.
.DESCRIPTION
    Walks $Root and partitions every file into two buckets:
    * Bundles      - bundles and loose .appx / .msix applications
    * Dependencies - .appx / .msix packages declaring Framework=true
    Throws if $Root does not exist.
.PARAMETER Root
    Folder to scan recursively.
.OUTPUTS
    PSCustomObject with Bundles and Dependencies arrays of FileInfo.
#>
    param(
        [Parameter(Mandatory)][string]$Root
    )
    if (-not (Test-Path -LiteralPath $Root -PathType Container)) {
        throw "SourcePath not found: $Root"
    }

    $all = Get-ChildItem -LiteralPath $Root -Recurse -File -ErrorAction Stop
    $bundles = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
    $dependencies = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
    foreach ($file in $all) {
        if ($file.Extension -in @('.msixbundle', '.appxbundle')) {
            $bundles.Add($file)
        }
        elseif ($file.Extension -in @('.msix', '.appx')) {
            if ((Get-PackageMetadata -Package $file).IsFramework) {
                $dependencies.Add($file)
            }
            else { $bundles.Add($file) }
        }
    }

    [pscustomobject]@{
        Bundles      = $bundles.ToArray()
        Dependencies = $dependencies.ToArray()
    }
}

function Resolve-LicensePath {
<#
.SYNOPSIS
    Returns the path to the license XML that pairs with a bundle, or $null.
.DESCRIPTION
    Prefer <BundleBaseName>.xml (FoD), otherwise use the single
    <StoreId>_License.xml (WinGet) in the same per-app directory.
    Multiple WinGet license files are ambiguous and cause an error.
.PARAMETER Bundle
    FileInfo for the .msixbundle / .appxbundle / .msix to look up.
.OUTPUTS
    [string] license file path, or $null when none is found.
#>
    param([Parameter(Mandatory)][System.IO.FileInfo]$Bundle)
    $candidate = Join-Path $Bundle.DirectoryName ("{0}.xml" -f $Bundle.BaseName)
    if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
    $licenses = @(Get-ChildItem -LiteralPath $Bundle.DirectoryName -Filter '*_License.xml' -File)
    if ($licenses.Count -eq 1) { return $licenses[0].FullName }
    if ($licenses.Count -gt 1) { throw "Ambiguous license files in '$($Bundle.DirectoryName)'. Keep one Store product per folder." }
    return $null
}

# -----------------------------------------------------------------------------
# INSTALL
# -----------------------------------------------------------------------------

function Install-Bundle {
<#
.SYNOPSIS
    Provisions an AppX/MSIX bundle, attaching its license XML when present.
.DESCRIPTION
    Calls Add-AppxProvisionedPackage -Online -PackagePath <bundle> with either
    -LicensePath <xml> (preferred) or an explicitly allowed -SkipLicense.
    All exceptions are converted to a Status='Failed' result object so
    the main loop can carry on and surface a single summary at the end.
.PARAMETER Bundle
    FileInfo for the .msixbundle / .appxbundle / .msix to provision.
.PARAMETER LicensePath
    Optional path to the matching license XML.
.PARAMETER DependencyPackagePath
    Framework packages located beside or below this application.
.PARAMETER SkipLicense
    Allow installation without a license XML for apps that permit it.
.OUTPUTS
    PSCustomObject (Name, Path, Kind='Bundle', LicensePath, Status, Error).
#>
    param(
        [Parameter(Mandatory)][System.IO.FileInfo]$Bundle,
        [string]$LicensePath,
        [string[]]$DependencyPackagePath,
        [switch]$SkipLicense
    )

    $base = @{
        Online      = $true
        PackagePath = $Bundle.FullName
        StubPackageOption = 'InstallFull'
        ErrorAction = 'Stop'
    }
    if ($DependencyPackagePath.Count -gt 0) { $base.DependencyPackagePath = $DependencyPackagePath }
    if ($LicensePath) {
        $base.LicensePath = $LicensePath
        Write-Log "Installing bundle:    $($Bundle.Name)  (license: $(Split-Path $LicensePath -Leaf))"
    }
    elseif ($SkipLicense) {
        $base.SkipLicense = $true
        Write-Log "Installing bundle:    $($Bundle.Name)  (no license file -> -SkipLicense)" -Level WARN
    }

    try {
        if (-not $LicensePath -and -not $SkipLicense) {
            throw 'No license found. Supply a license, or use -SkipLicense only for apps that permit it.'
        }
        Add-AppxProvisionedPackage @base | Out-Null
        return [pscustomobject]@{
            Name        = $Bundle.BaseName
            Path        = $Bundle.FullName
            Kind        = 'Bundle'
            LicensePath = $LicensePath
            Status      = 'Success'
            Error       = $null
        }
    }
    catch {
        $msg = ($_.Exception.Message -replace "[`r`n]+", ' ').Trim()
        Write-Log "Bundle '$($Bundle.Name)' failed: $msg" -Level ERROR
        return [pscustomobject]@{
            Name        = $Bundle.BaseName
            Path        = $Bundle.FullName
            Kind        = 'Bundle'
            LicensePath = $LicensePath
            Status      = 'Failed'
            Error       = $msg
        }
    }
}

function Test-ShouldUpdateProvisioned {
<#
.SYNOPSIS
    Decides whether a bundle should be (re)provisioned in -Mode UpdateProvisioned.
.DESCRIPTION
    In UpdateProvisioned mode we only refresh packages that are ALREADY part of
    the base image, so refreshing inbox apps from a FoD/Language ISO does not
    accidentally inject Store apps that were never present.

    Compares the embedded package Identity Name against the provisioned
    DisplayNames, case-insensitively. Download filenames are not identities.
.PARAMETER Bundle
    FileInfo for the bundle being considered.
.PARAMETER ProvisionedNames
    DisplayName values returned by Get-AppxProvisionedPackage on the live image.
.OUTPUTS
    [bool] $true when the bundle matches an already-provisioned package.
#>
    param(
        [Parameter(Mandatory)][System.IO.FileInfo]$Bundle,
        [Parameter(Mandatory)][AllowEmptyCollection()][string[]]$ProvisionedNames
    )
    $metadata = Get-PackageMetadata -Package $Bundle
    return $ProvisionedNames -contains $metadata.Name
}

# -----------------------------------------------------------------------------
# MAIN
# -----------------------------------------------------------------------------
New-Item -Path $LogDirectory -ItemType Directory -Force -ErrorAction Stop | Out-Null
Write-Log "=== $ScriptName starting (Mode=$Mode) ===" -Level HEADER
Write-Log "SourcePath  : $SourcePath"
Write-Log "Log file    : $LogFile"

$payload = Get-Payload -Root $SourcePath
Write-Log ("Discovered  : {0} bundle(s), {1} dependency package(s)" -f `
    $payload.Bundles.Count, $payload.Dependencies.Count)

if ($payload.Bundles.Count -eq 0) {
    Write-Log "No application payload found under '$SourcePath'. Check the download/export step." -Level ERROR
    exit 1
}

$results = New-Object System.Collections.Generic.List[object]

if ($Mode -eq 'UpdateProvisioned') {
    $provisioned = @(Get-AppxProvisionedPackage -Online | Select-Object -ExpandProperty DisplayName)
    Write-Log "Mode=UpdateProvisioned: $($provisioned.Count) provisioned package(s) currently on image."
}

if ($payload.Bundles.Count -gt 0) {
    Write-Log "--- Installing $($payload.Bundles.Count) bundle(s) ---" -Level HEADER
    foreach ($bundle in $payload.Bundles) {

        if ($Mode -eq 'UpdateProvisioned' -and -not (Test-ShouldUpdateProvisioned -Bundle $bundle -ProvisionedNames $provisioned)) {
            Write-Log "Skip (not provisioned on base image): $($bundle.Name)"
            $results.Add([pscustomobject]@{
                Name        = $bundle.BaseName
                Path        = $bundle.FullName
                Kind        = 'Bundle'
                LicensePath = $null
                Status      = 'Skipped'
                Error       = 'NotProvisioned'
            })
            continue
        }

        try {
            $licensePath = Resolve-LicensePath -Bundle $bundle
            $packageDirectory = $bundle.DirectoryName.TrimEnd([IO.Path]::DirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar
            $dependencyPaths = @($payload.Dependencies | Where-Object {
                $_.FullName.StartsWith($packageDirectory, [StringComparison]::OrdinalIgnoreCase)
            } | Select-Object -ExpandProperty FullName)
            $results.Add((Install-Bundle -Bundle $bundle -LicensePath $licensePath -DependencyPackagePath $dependencyPaths -SkipLicense:$SkipLicense))
        }
        catch {
            $message = ($_.Exception.Message -replace "[`r`n]+", ' ').Trim()
            Write-Log "Bundle '$($bundle.Name)' failed: $message" -Level ERROR
            $results.Add([pscustomobject]@{
                Name        = $bundle.BaseName
                Path        = $bundle.FullName
                Kind        = 'Bundle'
                LicensePath = $null
                Status      = 'Failed'
                Error       = $message
            })
        }
    }
}

# -----------------------------------------------------------------------------
# SUMMARY
# -----------------------------------------------------------------------------
$ok      = @($results | Where-Object Status -EQ 'Success').Count
$failed  = @($results | Where-Object Status -EQ 'Failed').Count
$skipped = @($results | Where-Object Status -EQ 'Skipped').Count

Write-Log "=== Summary ===" -Level HEADER
Write-Log ("Succeeded: {0} / {1}" -f $ok, $results.Count) -Level $(if ($failed -eq 0) { 'SUCCESS' } else { 'WARN' })
if ($skipped -gt 0) { Write-Log "Skipped  : $skipped" }
if ($failed  -gt 0) {
    Write-Log "Failed   : $failed" -Level ERROR
    $results | Where-Object Status -EQ 'Failed' |
        Format-Table Kind, Name, Error -AutoSize | Out-String | Write-Host
}

Write-Log "DISM details: %WinDir%\Logs\DISM\dism.log"
Write-Log "=== $ScriptName completed ===" -Level $(if ($failed -eq 0) { 'SUCCESS' } else { 'WARN' })

# Emit results object for pipeline / programmatic callers
$results

# Non-zero exit on any failure so Image Builder / Packer flags the step
if ($failed -gt 0) { exit 1 } else { exit 0 }
