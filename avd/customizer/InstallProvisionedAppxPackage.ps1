<#
.SYNOPSIS
    Side-loads a line-of-business UWP / MSIX app as a provisioned package on the
    running Windows image.

.DESCRIPTION
    Wraps `Add-AppxProvisionedPackage -Online` to install a single .appx /
    .appxbundle / .msix / .msixbundle as a provisioned package (available to all
    current and future user profiles). Designed for AVD / Windows 365 reference
    image bakes and one-off LOB deployments.

    Source folder layout (default - matches the legacy 2017 wrapper):

        <PackagePath>\
            MyApp_1.2.3.4_x64.appxbundle      (the bundle)
            MyApp_License.xml                 (optional - omit + use -SkipLicense)
            Dependencies\
                Microsoft.VCLibs....appx
                Microsoft.NET.Native.Framework....appx
                Microsoft.NET.Native.Runtime....appx

    -PackagePath may also point directly at a single .appxbundle / .msixbundle
    file, in which case -LicensePath and -DependenciesPath default to the file's
    sibling locations and can be overridden explicitly.

    Sideloading prerequisite:
      `HKLM\Software\Policies\Microsoft\Windows\Appx!AllowAllTrustedApps = 1`
    is required for non-Store-signed packages. The script captures the existing
    value, sets it to 1 for the install, and restores the original on exit
    (including on failure). Skip with -KeepSideloadingEnabled when the bake
    will install multiple LOB packages back-to-back.

    Runs in an ADMIN user context (NOT SYSTEM). The Appx provisioning APIs work
    in either, but the script does not assume access-token niceties only SYSTEM
    has.

.PARAMETER PackagePath
    Either a folder containing the bundle + (optional) .xml license + Dependencies
    subfolder, OR the full path to a single .appx / .appxbundle / .msix /
    .msixbundle file. Defaults to `$PSScriptRoot\AppxBundle`.

.PARAMETER LicensePath
    Full path to the .xml license file. When -PackagePath is a folder, this is
    auto-discovered (first *.xml in the folder). Mutually exclusive with
    -SkipLicense.

.PARAMETER DependenciesPath
    Folder containing dependency .appx / .msix files (recursive). When
    -PackagePath is a folder, defaults to `<PackagePath>\Dependencies`.

.PARAMETER SkipLicense
    Provision the bundle without a license file. Required when the package is
    not store-signed and you have no .xml license. Mutually exclusive with
    -LicensePath.

.PARAMETER KeepSideloadingEnabled
    Do NOT restore the original `AllowAllTrustedApps` value on exit. Use this
    when chaining multiple LOB installs in the same bake step.

.PARAMETER LogDirectory
    Directory for the log file. Default: $env:TEMP. The log is named
    Install-ProvisionedAppxPackage_yyyyMMdd_HHmmss.log.

.NOTES
    File:    avd/customizer/InstallProvisionedAppxPackage.ps1
    Author:  Anton Romanyuk
    Version: 1.0.0
    Context: Reference image bake (AIB / Packer / Run-Command) OR one-off ADMIN
             user session. Does NOT need to run as SYSTEM.
    Requires: Windows 10 / 11 / Server with Desktop Experience, PowerShell 5.1+,
              elevated (admin) session.

    Changes:
      1.0.0 - Port + hardening of the legacy 2017 Install-ProvisionedAppxPackage.ps1.
              Renamed to InstallProvisionedAppxPackage.ps1 to match the
              VerbNoun.ps1 convention used by the other customizers.
              Switched from dism.exe shell-out to Add-AppxProvisionedPackage,
              parameterised, idiomatic Write-Log, AllowAllTrustedApps captured
              and restored, 3010 mapped to success, package + dependencies
              auto-discovered, paths-with-spaces safe (array argument lists).

    Exit codes:
      0    - Success (or success + reboot required, exit 3010 mapped to 0 with WARN)
      1    - Hard failure (package not found, provisioning error, etc.)

.DISCLAIMER
    This script is provided "AS IS" with no warranties and confers no rights.
    It is not supported under any Microsoft standard support program or service.
    Use of this script is entirely at your own risk. Setting
    AllowAllTrustedApps = 1 lowers a sideloading guard; the script restores the
    original value on exit unless -KeepSideloadingEnabled is supplied.

.EXAMPLE
    # Default folder layout next to the script
    .\InstallProvisionedAppxPackage.ps1

.EXAMPLE
    # Explicit bundle, no license, custom dependency folder
    .\InstallProvisionedAppxPackage.ps1 `
        -PackagePath C:\Bake\MyApp\MyApp_1.2.3.4_x64.appxbundle `
        -DependenciesPath C:\Bake\MyApp\Deps `
        -SkipLicense

.EXAMPLE
    # Chain multiple LOB installs without flapping the sideloading policy
    .\InstallProvisionedAppxPackage.ps1 -PackagePath C:\Bake\App1 -KeepSideloadingEnabled
    .\InstallProvisionedAppxPackage.ps1 -PackagePath C:\Bake\App2 -KeepSideloadingEnabled
    .\InstallProvisionedAppxPackage.ps1 -PackagePath C:\Bake\App3
#>

#Requires -RunAsAdministrator
[CmdletBinding(DefaultParameterSetName = 'License')]
param(
    [string]$PackagePath = (Join-Path $PSScriptRoot 'AppxBundle'),

    [Parameter(ParameterSetName = 'License')]
    [string]$LicensePath,

    [string]$DependenciesPath,

    [Parameter(ParameterSetName = 'NoLicense', Mandatory)]
    [switch]$SkipLicense,

    [switch]$KeepSideloadingEnabled,

    [string]$LogDirectory = $env:TEMP
)

$ErrorActionPreference = 'Stop'

# -----------------------------------------------------------------------------
# LOGGING
# -----------------------------------------------------------------------------
$ScriptName = 'InstallProvisionedAppxPackage'
$LogFile    = Join-Path $LogDirectory ("{0}_{1}.log" -f $ScriptName, (Get-Date -Format 'yyyyMMdd_HHmmss'))

function Write-Log {
<#
.SYNOPSIS
    Writes a timestamped, level-tagged line to the console and the log file.
.PARAMETER Message
    Free-form text.
.PARAMETER Level
    INFO | WARN | ERROR | SUCCESS. Default INFO.
#>
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS')][string]$Level = 'INFO'
    )
    $ts   = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')
    $line = "[$ts] [$Level] [$ScriptName] $Message"
    Add-Content -LiteralPath $LogFile -Value $line -ErrorAction SilentlyContinue
    $color = switch ($Level) { 'WARN' { 'Yellow' } 'ERROR' { 'Red' } 'SUCCESS' { 'Green' } default { 'Gray' } }
    Write-Host $line -ForegroundColor $color
}

# -----------------------------------------------------------------------------
# SIDELOADING POLICY HELPERS
# -----------------------------------------------------------------------------
$AppxPolicyKey  = 'HKLM:\Software\Policies\Microsoft\Windows\Appx'
$AppxPolicyName = 'AllowAllTrustedApps'

function Get-AllowAllTrustedAppsState {
<#
.SYNOPSIS
    Snapshots the current AllowAllTrustedApps policy value (or 'absent') so we
    can restore it on exit.
.OUTPUTS
    PSCustomObject with KeyExisted (bool) and Value (int? — $null if absent).
#>
    if (-not (Test-Path -LiteralPath $AppxPolicyKey)) {
        return [pscustomobject]@{ KeyExisted = $false; Value = $null }
    }
    try {
        $v = Get-ItemProperty -LiteralPath $AppxPolicyKey -Name $AppxPolicyName -ErrorAction Stop
        return [pscustomobject]@{ KeyExisted = $true; Value = [int]$v.$AppxPolicyName }
    }
    catch {
        return [pscustomobject]@{ KeyExisted = $true; Value = $null }
    }
}

function Set-AllowAllTrustedApps {
<#
.SYNOPSIS
    Forces the AllowAllTrustedApps sideloading policy to a specific value.
.DESCRIPTION
    Creates the parent policy key if it does not yet exist, then writes the
    DWORD value. Used to enable sideloading (Value = 1) for the duration of
    a provisioning call.
.PARAMETER Value
    DWORD value to write (0 = disabled, 1 = enabled).
#>
    param([Parameter(Mandatory)][int]$Value)
    if (-not (Test-Path -LiteralPath $AppxPolicyKey)) {
        New-Item -Path $AppxPolicyKey -Force | Out-Null
    }
    New-ItemProperty -LiteralPath $AppxPolicyKey -Name $AppxPolicyName -Value $Value -PropertyType DWord -Force | Out-Null
}

function Restore-AllowAllTrustedApps {
<#
.SYNOPSIS
    Restores the AllowAllTrustedApps policy to the snapshot captured by
    Get-AllowAllTrustedAppsState.
.DESCRIPTION
    Handles three entry-state cases:
      1. Parent key did not exist  -> remove the value, and remove the key if
         we left it empty (no footprint).
      2. Key existed, value absent -> remove only the value.
      3. Value present             -> write the original DWORD back.
    Failures are swallowed and logged as WARN so they do not mask the real
    install error in a finally{} block.
.PARAMETER Original
    Snapshot object returned by Get-AllowAllTrustedAppsState
    (KeyExisted [bool], Value [int?]).
#>
    param([Parameter(Mandatory)][pscustomobject]$Original)
    try {
        if (-not $Original.KeyExisted) {
            # Original state was: parent key absent. Remove the value (and the key
            # iff we created it AND it's empty after) to leave no footprint.
            if (Test-Path -LiteralPath $AppxPolicyKey) {
                Remove-ItemProperty -LiteralPath $AppxPolicyKey -Name $AppxPolicyName -ErrorAction SilentlyContinue
                $remaining = Get-Item -LiteralPath $AppxPolicyKey -ErrorAction SilentlyContinue
                if ($remaining -and $remaining.Property.Count -eq 0 -and $remaining.SubKeyCount -eq 0) {
                    Remove-Item -LiteralPath $AppxPolicyKey -Force -ErrorAction SilentlyContinue
                }
            }
            Write-Log "Restored sideloading policy: removed (parent key did not exist on entry)"
            return
        }
        if ($null -eq $Original.Value) {
            # Key existed but value did not; remove only the value.
            Remove-ItemProperty -LiteralPath $AppxPolicyKey -Name $AppxPolicyName -ErrorAction SilentlyContinue
            Write-Log "Restored sideloading policy: $AppxPolicyName removed (was absent on entry)"
            return
        }
        New-ItemProperty -LiteralPath $AppxPolicyKey -Name $AppxPolicyName -Value $Original.Value -PropertyType DWord -Force | Out-Null
        Write-Log "Restored sideloading policy: $AppxPolicyName = $($Original.Value)"
    }
    catch {
        Write-Log "Failed to restore sideloading policy: $($_.Exception.Message)" -Level WARN
    }
}

# -----------------------------------------------------------------------------
# PACKAGE / DEPENDENCY DISCOVERY
# -----------------------------------------------------------------------------
function Resolve-PackageInputs {
<#
.SYNOPSIS
    Resolves -PackagePath into (BundleFile, LicenseFile, DependencyFiles[]).
.DESCRIPTION
    -PackagePath may be a folder OR a direct .appx*/.msix* file. Auto-discovers
    license (*.xml in the same folder) and dependencies (*.appx / *.msix under
    a sibling 'Dependencies' folder, recursive). Explicit -LicensePath /
    -DependenciesPath override discovery.
.OUTPUTS
    PSCustomObject with Bundle, License (string or $null), Dependencies (string[]).
#>
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$ExplicitLicense,
        [string]$ExplicitDeps,
        [bool]  $SkipLicenseFlag
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "PackagePath not found: $Path"
    }
    $item = Get-Item -LiteralPath $Path

    if ($item.PSIsContainer) {
        $folder = $item.FullName
        $bundle = Get-ChildItem -LiteralPath $folder -File -Include '*.appxbundle','*.msixbundle','*.appx','*.msix' -ErrorAction SilentlyContinue |
                  Select-Object -First 1
    } else {
        $folder = $item.DirectoryName
        $bundle = $item
    }

    if (-not $bundle) {
        throw "No .appx / .appxbundle / .msix / .msixbundle found under '$Path'"
    }

    # License
    $licenseFile = $null
    if (-not $SkipLicenseFlag) {
        if ($ExplicitLicense) {
            if (-not (Test-Path -LiteralPath $ExplicitLicense)) { throw "LicensePath not found: $ExplicitLicense" }
            $licenseFile = (Resolve-Path -LiteralPath $ExplicitLicense).Path
        } else {
            $licenseFile = Get-ChildItem -LiteralPath $folder -File -Filter '*.xml' -ErrorAction SilentlyContinue |
                           Select-Object -First 1 -ExpandProperty FullName
        }
    }

    # Dependencies
    $depFolder = if ($ExplicitDeps) { $ExplicitDeps } else { Join-Path $folder 'Dependencies' }
    $deps = @()
    if (Test-Path -LiteralPath $depFolder) {
        $deps = @(Get-ChildItem -LiteralPath $depFolder -File -Recurse -Include '*.appx','*.msix' -ErrorAction SilentlyContinue |
                  Select-Object -ExpandProperty FullName)
    }

    [pscustomobject]@{
        Bundle       = $bundle.FullName
        License      = $licenseFile
        Dependencies = $deps
    }
}

# -----------------------------------------------------------------------------
# MAIN
# -----------------------------------------------------------------------------
Write-Log "===== $ScriptName v1.0.0 starting =====" -Level SUCCESS
Write-Log "Log file        : $LogFile"
Write-Log "PackagePath     : $PackagePath"
if ($LicensePath)      { Write-Log "LicensePath     : $LicensePath" }
if ($DependenciesPath) { Write-Log "DependenciesPath: $DependenciesPath" }
if ($SkipLicense)      { Write-Log "SkipLicense     : true" -Level WARN }

$exitCode = 0
$originalSideloadState = $null

try {
    $resolved = Resolve-PackageInputs `
        -Path            $PackagePath `
        -ExplicitLicense $LicensePath `
        -ExplicitDeps    $DependenciesPath `
        -SkipLicenseFlag $SkipLicense.IsPresent

    Write-Log "Bundle          : $($resolved.Bundle)"
    if ($resolved.License) {
        Write-Log "License         : $($resolved.License)"
    } elseif ($SkipLicense) {
        Write-Log "License         : <none> (-SkipLicense)"
    } else {
        Write-Log "License         : <none found> - falling back to -SkipLicense" -Level WARN
        $SkipLicense = [switch]$true
    }
    Write-Log "Dependencies    : $($resolved.Dependencies.Count) file(s)"
    foreach ($d in $resolved.Dependencies) { Write-Log "  - $d" }

    # Snapshot + enable sideloading
    $originalSideloadState = Get-AllowAllTrustedAppsState
    Write-Log ("Sideloading policy on entry: KeyExisted={0}, Value={1}" -f $originalSideloadState.KeyExisted, $originalSideloadState.Value)
    Set-AllowAllTrustedApps -Value 1
    Write-Log "Sideloading policy: $AppxPolicyName = 1 (enabled for install)"

    # Build splat for Add-AppxProvisionedPackage
    $splat = @{
        Online      = $true
        PackagePath = $resolved.Bundle
        ErrorAction = 'Stop'
    }
    if ($resolved.Dependencies.Count -gt 0) {
        $splat.DependencyPackagePath = $resolved.Dependencies
    }
    if ($SkipLicense) {
        $splat.SkipLicense = $true
    } else {
        $splat.LicensePath = $resolved.License
    }

    Write-Log "Calling Add-AppxProvisionedPackage -Online ..."
    Add-AppxProvisionedPackage @splat | Out-Null
    Write-Log "Add-AppxProvisionedPackage completed successfully" -Level SUCCESS
}
catch [System.Runtime.InteropServices.COMException] {
    # 0x80073D02 == 0x00000BF6 (-2147009790) "Package operation pending reboot" - treat as 3010 (success + reboot)
    if ($_.Exception.HResult -eq -2147009790) {
        Write-Log "Provisioning succeeded but a reboot is required (HRESULT 0x80073D02) - exit 0 (mapped from 3010)" -Level WARN
        $exitCode = 0
    } else {
        Write-Log "COM exception during provisioning: 0x$('{0:X8}' -f $_.Exception.HResult) - $($_.Exception.Message)" -Level ERROR
        $exitCode = 1
    }
}
catch {
    Write-Log "Provisioning failed: $($_.Exception.Message)" -Level ERROR
    Write-Log $_.ScriptStackTrace -Level ERROR
    $exitCode = 1
}
finally {
    if ($originalSideloadState) {
        if ($KeepSideloadingEnabled) {
            Write-Log "-KeepSideloadingEnabled specified - leaving $AppxPolicyName = 1" -Level WARN
        } else {
            Restore-AllowAllTrustedApps -Original $originalSideloadState
        }
    }
}

Write-Log "===== $ScriptName completed (exit $exitCode) =====" -Level $(if ($exitCode -eq 0) { 'SUCCESS' } else { 'ERROR' })
exit $exitCode
