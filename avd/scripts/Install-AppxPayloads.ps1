<#
.SYNOPSIS
    Side-loads / re-provisions inbox AppX/MSIX packages from a local payload tree.

.DESCRIPTION
    Companion to Get-StubAppPayloads.ps1. Walks a payload root and provisions
    every package it finds for all current and future users via DISM
    (`Add-AppxProvisionedPackage`). Designed to fix two recurring image-build
    problems:

      1. STUB APPS  - Some inbox Microsoft Store apps (Photos Legacy, Clock,
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
        Only provision a bundle when its package family name (or DisplayName
        prefix) already exists in Get-AppxProvisionedPackage. Use this when
        refreshing inbox apps from a FoD/Language ISO so you don't accidentally
        add Store apps that were never part of the base image.

    Layout assumptions (works for both winget-msstore downloads and FoD ISO trees):

      <SourcePath>\
        <AppName-or-arch>\
          *.msixbundle | *.appxbundle | *.msix | *.appx   <- main package
          *.xml                                            <- matching license
                                                              (basename.xml)
          *.appx                                           <- dependencies
                                                              (Microsoft.VCLibs,
                                                               Microsoft.UI.Xaml, ...)

    The script:
      - Recursively discovers all bundles and dependency .appx files
      - Installs dependency packages first (with -SkipLicense)
      - Installs each bundle, attaching <basename>.xml license when present
      - Returns a result object per package and a summary at the end

.PARAMETER SourcePath
    Root folder containing the payload tree. For the stub-app workflow this is
    the folder produced by Get-StubAppPayloads.ps1 (e.g. C:\Temp\AVD_Stubs_Payload).
    For the FoD workflow this is the architecture folder on the mounted ISO
    (e.g. E:\LanguagesAndOptionalFeatures or D:\sources\<build>\amd64fre).

.PARAMETER Mode
    Install            -> provision every bundle (stub-app fix). Default.
    UpdateProvisioned  -> only refresh bundles whose package family / display
                          name already exists in Get-AppxProvisionedPackage.

.PARAMETER LogDirectory
    Directory for the log file. Default: $env:TEMP.

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

    [string]$LogDirectory = $env:TEMP
)

$ErrorActionPreference = 'Stop'

# -----------------------------------------------------------------------------
# LOGGING
# -----------------------------------------------------------------------------
$ScriptName = $MyInvocation.MyCommand.Name
$LogFile    = Join-Path $LogDirectory ("{0}_{1}.log" -f [IO.Path]::GetFileNameWithoutExtension($ScriptName), (Get-Date -Format 'yyyyMMdd_HHmmss'))

function Write-Log {
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

# Bundle/main-package extensions, in priority order.
$BundleExtensions = @('.msixbundle', '.appxbundle', '.msix')

# A loose .appx that is NOT a main package (because no matching license) is
# treated as a dependency (e.g. Microsoft.VCLibs.x64.14.00.Desktop.appx).

function Get-Payload {
    param(
        [Parameter(Mandatory)][string]$Root
    )
    if (-not (Test-Path -LiteralPath $Root)) {
        throw "SourcePath not found: $Root"
    }

    $all = Get-ChildItem -LiteralPath $Root -Recurse -File -ErrorAction Stop

    $bundles = $all | Where-Object { $BundleExtensions -contains $_.Extension.ToLowerInvariant() }

    # Build the dependency set: standalone .appx files in the tree that are
    # NOT also picked up as a main package (rare, but defensive).
    $bundleFullNames = [System.Collections.Generic.HashSet[string]]::new(
        [string[]]($bundles.FullName), [System.StringComparer]::OrdinalIgnoreCase)
    $dependencies = $all | Where-Object {
        $_.Extension -ieq '.appx' -and -not $bundleFullNames.Contains($_.FullName)
    }

    [pscustomobject]@{
        Bundles      = @($bundles)
        Dependencies = @($dependencies)
    }
}

function Resolve-LicensePath {
    param([Parameter(Mandatory)][System.IO.FileInfo]$Bundle)
    # Convention: <bundle-basename>.xml in the same directory.
    $candidate = Join-Path $Bundle.DirectoryName ("{0}.xml" -f $Bundle.BaseName)
    if (Test-Path -LiteralPath $candidate) { return $candidate }
    return $null
}

# -----------------------------------------------------------------------------
# INSTALL
# -----------------------------------------------------------------------------

function Install-Dependency {
    param([Parameter(Mandatory)][System.IO.FileInfo]$Appx)
    Write-Log "Installing dependency: $($Appx.Name)"
    try {
        Add-AppxProvisionedPackage -Online -PackagePath $Appx.FullName -SkipLicense | Out-Null
        return [pscustomobject]@{
            Name   = $Appx.BaseName
            Path   = $Appx.FullName
            Kind   = 'Dependency'
            Status = 'Success'
            Error  = $null
        }
    }
    catch {
        $msg = ($_.Exception.Message -replace "[`r`n]+", ' ').Trim()
        Write-Log "Dependency '$($Appx.Name)' failed: $msg" -Level ERROR
        return [pscustomobject]@{
            Name   = $Appx.BaseName
            Path   = $Appx.FullName
            Kind   = 'Dependency'
            Status = 'Failed'
            Error  = $msg
        }
    }
}

function Install-Bundle {
    param(
        [Parameter(Mandatory)][System.IO.FileInfo]$Bundle,
        [string]$LicensePath
    )

    $base = @{
        Online      = $true
        PackagePath = $Bundle.FullName
    }
    if ($LicensePath) {
        $base.LicensePath = $LicensePath
        Write-Log "Installing bundle:    $($Bundle.Name)  (license: $(Split-Path $LicensePath -Leaf))"
    }
    else {
        $base.SkipLicense = $true
        Write-Log "Installing bundle:    $($Bundle.Name)  (no license file -> -SkipLicense)" -Level WARN
    }

    try {
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
      In -Mode UpdateProvisioned we only re-install a bundle if a package with
      a matching name is already provisioned on the image. Match heuristic:
      bundle BaseName starts with the provisioned DisplayName (handles the
      typical "Microsoft.WindowsCalculator_2024.1234.0_neutral_~_8wekyb3d8bbwe"
      vs "Microsoft.WindowsCalculator" comparison).
    #>
    param(
        [Parameter(Mandatory)][System.IO.FileInfo]$Bundle,
        [Parameter(Mandatory)][string[]]$ProvisionedNames
    )
    foreach ($name in $ProvisionedNames) {
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        if ($Bundle.BaseName.StartsWith($name, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }
    return $false
}

# -----------------------------------------------------------------------------
# MAIN
# -----------------------------------------------------------------------------
Write-Log "=== $ScriptName starting (Mode=$Mode) ===" -Level HEADER
Write-Log "SourcePath  : $SourcePath"
Write-Log "Log file    : $LogFile"

$payload = Get-Payload -Root $SourcePath
Write-Log ("Discovered  : {0} bundle(s), {1} dependency package(s)" -f `
    $payload.Bundles.Count, $payload.Dependencies.Count)

if ($payload.Bundles.Count -eq 0 -and $payload.Dependencies.Count -eq 0) {
    Write-Log "No payload found under '$SourcePath'. Nothing to do." -Level WARN
    exit 0
}

$results = New-Object System.Collections.Generic.List[object]

# --- Step 1: dependencies first (.appx without matching license) ----------
if ($payload.Dependencies.Count -gt 0) {
    Write-Log "--- Installing $($payload.Dependencies.Count) dependency package(s) ---" -Level HEADER
    foreach ($dep in $payload.Dependencies) {
        $results.Add((Install-Dependency -Appx $dep))
    }
}

# --- Step 2: bundles ------------------------------------------------------
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

        $licensePath = Resolve-LicensePath -Bundle $bundle
        $results.Add((Install-Bundle -Bundle $bundle -LicensePath $licensePath))
    }
}

# -----------------------------------------------------------------------------
# SUMMARY
# -----------------------------------------------------------------------------
$ok      = ($results | Where-Object Status -EQ 'Success').Count
$failed  = ($results | Where-Object Status -EQ 'Failed').Count
$skipped = ($results | Where-Object Status -EQ 'Skipped').Count

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
