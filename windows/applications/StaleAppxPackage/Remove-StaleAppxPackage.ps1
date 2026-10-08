<#
.SYNOPSIS
    Removes per-user registrations of packaged apps (Appx/MSIX) whose version is
    below a minimum, e.g. inbox codec/image extensions flagged by vulnerability
    scanners.

.DESCRIPTION
    Packaged apps are registered per user. A user who has not signed in since an
    app update keeps the old registration, and the old package stays under
    C:\Program Files\WindowsApps where scanners report it (KB5011324).

    For every package name in -MinimumVersion the script enumerates all user
    registrations and removes each one whose package version is lower than the
    minimum, for that user only (Remove-AppxPackage -User <SID>). Users already
    on a current version are not touched.

    NOT done by this script:
      - Provisioned packages are not removed. An outdated provisioned copy is
        logged as a warning, because every new profile would register it again.
      - Well-known service accounts (SYSTEM, LOCAL SERVICE, NETWORK SERVICE)
        are skipped.
      - Packages with no user registration at all are not removed.

    Versions are compared as [version]. Get-AppxPackage returns Version as a
    STRING, so a plain -lt is a lexical compare and gets e.g. 1.0.9.0 vs
    1.0.10.0 wrong.

    FLOW:
      1. Preflight: elevation check and validation of every minimum version.
         Any failure ends the run with exit 1 before anything is touched.
      2. Per package name: warn on an outdated provisioned copy, enumerate
         outdated user registrations, remove each one (or report under
         -WhatIf). An enumeration error for one name does not stop the others.
      3. Verify: re-query every name and count outdated registrations that are
         still present for non-service accounts. Skipped under -WhatIf.

    SAFETY:
      - -WhatIf is a full dry run: enumeration, provisioned-package warnings
        and logging happen, no registration is removed.
      - -Confirm prompts once per user registration.
      - Without either switch the script runs unattended; it never prompts.

    LOG FORMAT, one line per event, also echoed to the host:
        <UTC ISO 8601> [INFO|WARN|ERROR] [PID:<pid>] <message>

.PARAMETER MinimumVersion
    Optional override of the built-in list: package Name -> lowest version that
    is considered current. Keys are exact package names (no wildcards); values
    are four-part version strings. Replaces the built-in list entirely, it is
    not merged. Example: @{ 'Microsoft.MSPaint' = '6.2203.1037.0' }

    Separate multiple entries with ';' or line breaks. Use [ordered]@{...} to
    keep processing order. From cmd.exe or a deployment tool, launch with
    powershell.exe -Command, not -File: -File passes the hashtable as a string
    and parameter binding fails.

    To maintain the built-in list, edit the ordered table at the top of the
    script body. Packages are processed in the listed order.

.PARAMETER LogPath
    Log file, appended with UTC timestamps. Defaults to the Intune Management
    Extension log folder so 'Collect diagnostics' picks it up. The folder is
    created if missing. A write failure warns once and does not change the
    result or exit code.

.INPUTS
    None.

.OUTPUTS
    One [pscustomobject] summary:
      Computer   Computer name.
      Mode       EXECUTE or WHATIF.
      Removed    Registrations removed.
      Failed     Registrations whose removal threw an error.
      Skipped    Registrations of well-known service accounts.
      WhatIf     Registrations that would be removed (-WhatIf or declined
                 -Confirm).
      Remaining  Outdated non-service registrations found by the verify pass;
                 0 under -WhatIf because verification is skipped.
      LogPath    Path of the log file.
      Details    One object per outdated registration: Name, PackageFullName,
                 Version, MinimumVersion, Sid, User, InstallState, Outcome
                 (Removed | Failed | Skipped | WhatIf) and Error (HRESULT and
                 message on failure, otherwise $null).

.EXAMPLE
    .\Remove-StaleAppxPackage.ps1 -WhatIf

    Lists every registration that would be removed. Changes nothing.

.EXAMPLE
    .\Remove-StaleAppxPackage.ps1

    Unattended removal using the built-in list (Intune / ConfigMgr / RMM).

.EXAMPLE
    .\Remove-StaleAppxPackage.ps1 -MinimumVersion @{ 'Microsoft.VP9VideoExtensions' = '1.0.52781.0' } -Confirm

    Checks a single package against a custom minimum and prompts before each
    removal.

.EXAMPLE
    .\Remove-StaleAppxPackage.ps1 -WhatIf -MinimumVersion ([ordered]@{
        'Microsoft.VP9VideoExtensions' = '1.0.52781.0'
        'Microsoft.WebMediaExtensions' = '1.0.62192.0'
        'Microsoft.HEVCVideoExtension' = '2.1.1803.0'
    })

    Dry run against three packages with custom minimums, in the listed order.
    Versions shown are placeholders.

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "& '.\Remove-StaleAppxPackage.ps1' -MinimumVersion @{ 'Microsoft.VP9VideoExtensions' = '1.0.52781.0'; 'Microsoft.WebMediaExtensions' = '1.0.62192.0' }; exit $LASTEXITCODE"

    Same from cmd.exe, ConfigMgr or an RMM tool. Must be -Command, not -File;
    'exit $LASTEXITCODE' passes the script's exit code through explicitly.

.EXAMPLE
    (.\Remove-StaleAppxPackage.ps1 -WhatIf).Details | Format-Table Name, Version, User, InstallState, Outcome

    Dry run, shown as a per-user table.

.NOTES
    Sample script, provided as-is and outside of Microsoft support scope.
    Review the built-in list against current scanner findings before use; it is
    not a complete or current list of vulnerable app versions.

    Requires Windows PowerShell 5.1, elevated or SYSTEM. In Intune, enable
    'Run script in 64 bit PowerShell Host'.

    Exit 0: no failures and no outdated user registrations left.
    Exit 1: preflight failed, a removal failed, or outdated registrations remain.
    A -WhatIf run exits 0 unless enumeration fails.

    Files can stay in WindowsApps after removal; Windows cleans them up
    asynchronously (KB5011324). An immediate rescan may still flag the folder.

    Prefer, in this order: let affected users sign in so the app updates; delete
    stale profiles (Group Policy 'Delete user profiles older than a specified
    number of days on system restart'); then this script; deprovisioning the app
    entirely is the last resort.

.LINK
    https://learn.microsoft.com/en-us/troubleshoot/windows-client/application-management/modern-apps-application-packages-reported-vulnerable

.LINK
    https://learn.microsoft.com/en-us/powershell/module/appx/remove-appxpackage
#>
#Requires -Version 5.1
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [System.Collections.IDictionary]$MinimumVersion,

    [string]$LogPath = (Join-Path $env:ProgramData 'Microsoft\IntuneManagementExtension\Logs\Remove-StaleAppxPackage.log')
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# Maintain here: package Name -> lowest version that is NOT outdated.
if (-not $MinimumVersion) {
    $MinimumVersion = [ordered]@{
        'Microsoft.HEIFImageExtension' = '1.0.43012.0'
        'Microsoft.Microsoft3DViewer'  = '7.2107.7012.0'
        'Microsoft.MSPaint'            = '6.2203.1037.0'
        'Microsoft.RawImageExtension'  = '2.1.30191.0'
        'Microsoft.VP9VideoExtensions' = '1.0.42791.0'
        'Microsoft.WebMediaExtensions' = '1.0.42192.0'
        'Microsoft.WebpImageExtension' = '1.0.42351.0'
    }
}

$SkipSids = @('S-1-5-18', 'S-1-5-19', 'S-1-5-20')
$script:LogWarned = $false

function Write-Log {
    <#
    .SYNOPSIS
        Writes a UTC-stamped line to the host and appends it to $LogPath.
    .PARAMETER Message
        Text to log.
    .PARAMETER Level
        INFO, WARN or ERROR. Default INFO.
    .NOTES
        Uses .NET file I/O so -WhatIf does not suppress logging. Writes UTF-8
        without BOM. A write failure warns once per run and is otherwise ignored.
    #>
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO'
    )
    $line = '{0} [{1}] [PID:{2}] {3}' -f [DateTime]::UtcNow.ToString('o'), $Level, $PID, $Message
    Write-Host $line
    try {
        [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($LogPath))
        [IO.File]::AppendAllText($LogPath, $line + [Environment]::NewLine, (New-Object System.Text.UTF8Encoding($false)))
    }
    catch {
        if (-not $script:LogWarned) {
            $script:LogWarned = $true
            Write-Warning "Cannot write log '$LogPath': $($_.Exception.Message)"
        }
    }
}

function Get-OutdatedRegistration {
    <#
    .SYNOPSIS
        Returns one object per user registration of a package below a minimum version.
    .PARAMETER Name
        Exact package name, e.g. Microsoft.MSPaint.
    .PARAMETER Minimum
        Lowest version considered current.
    .OUTPUTS
        [pscustomobject] with Name, PackageFullName, Version, MinimumVersion,
        Sid, User, InstallState. Service-account SIDs are included; callers
        decide whether to skip them.
    .NOTES
        Requires elevation (Get-AppxPackage -AllUsers). Errors propagate to the
        caller. A package with no user registrations yields nothing.
    #>
    param([Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][version]$Minimum)
    foreach ($pkg in @(Get-AppxPackage -AllUsers -Name $Name)) {
        $version = [version]$pkg.Version
        if ($version -ge $Minimum) { continue }
        foreach ($ui in @($pkg.PackageUserInformation)) {
            [pscustomobject]@{
                Name            = $Name
                PackageFullName = $pkg.PackageFullName
                Version         = $version
                MinimumVersion  = $Minimum
                Sid             = [string]$ui.UserSecurityId.Sid
                User            = [string]$ui.UserSecurityId.Username
                InstallState    = [string]$ui.InstallState
            }
        }
    }
}

#region Preflight

try {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    if (-not (New-Object Security.Principal.WindowsPrincipal($identity)).IsInRole(
            [Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw 'Must run elevated or as SYSTEM; Get-AppxPackage -AllUsers requires it.'
    }

    $targets = foreach ($key in @($MinimumVersion.Keys)) {
        $min = $null
        if (-not [version]::TryParse([string]$MinimumVersion[$key], [ref]$min)) {
            throw "Invalid minimum version '$($MinimumVersion[$key])' for '$key'."
        }
        [pscustomobject]@{ Name = [string]$key; MinimumVersion = $min }
    }
}
catch {
    Write-Log "Preflight failed: $($_.Exception.Message)" -Level ERROR
    exit 1
}

$mode = 'EXECUTE'
if ($WhatIfPreference) { $mode = 'WHATIF' }
Write-Log ("Start on {0} as {1}, mode {2}, {3} package name(s)." -f $env:COMPUTERNAME, $identity.Name, $mode, @($targets).Count)

$provisioned = @()
try {
    $provisioned = @(Get-AppxProvisionedPackage -Online)
}
catch {
    Write-Log "Cannot read provisioned packages: $($_.Exception.Message)" -Level WARN
}

#endregion Preflight

#region Removal

$details   = New-Object System.Collections.Generic.List[object]
$enumError = $false

foreach ($t in $targets) {
    foreach ($prov in @($provisioned | Where-Object { $_.DisplayName -eq $t.Name })) {
        $provVersion = $null
        if ([version]::TryParse([string]$prov.Version, [ref]$provVersion) -and $provVersion -lt $t.MinimumVersion) {
            Write-Log ("{0}: provisioned version {1} < {2}. New profiles will register the outdated version; update the provisioned package." -f $t.Name, $provVersion, $t.MinimumVersion) -Level WARN
        }
    }

    try {
        $outdated = @(Get-OutdatedRegistration -Name $t.Name -Minimum $t.MinimumVersion)
    }
    catch {
        Write-Log ("{0}: enumeration failed: {1}" -f $t.Name, $_.Exception.Message) -Level ERROR
        $enumError = $true
        continue
    }

    if ($outdated.Count -eq 0) {
        Write-Log ("{0}: no user registration below {1}." -f $t.Name, $t.MinimumVersion)
        continue
    }

    foreach ($reg in $outdated) {
        $who = $reg.Sid
        if ($reg.User) { $who = '{0} ({1})' -f $reg.User, $reg.Sid }
        $what = '{0} [{1}] for {2}' -f $reg.PackageFullName, $reg.InstallState, $who
        $outcome = 'WhatIf'
        $errorText = $null

        if ($SkipSids -contains $reg.Sid) {
            $outcome = 'Skipped'
            Write-Log "Skip service account: $what"
        }
        elseif ($PSCmdlet.ShouldProcess($what, 'Remove-AppxPackage -User')) {
            try {
                Remove-AppxPackage -Package $reg.PackageFullName -User $reg.Sid -ErrorAction Stop
                $outcome = 'Removed'
                Write-Log "Removed: $what"
            }
            catch {
                $outcome   = 'Failed'
                $errorText = '0x{0:X8} {1}' -f $_.Exception.HResult, $_.Exception.Message
                Write-Log "Failed: $what - $errorText" -Level ERROR
            }
        }
        else {
            Write-Log "WhatIf, not removed: $what"
        }

        $reg | Add-Member -NotePropertyName Outcome -NotePropertyValue $outcome
        $reg | Add-Member -NotePropertyName Error -NotePropertyValue $errorText
        $details.Add($reg)
    }
}

#endregion Removal

#region Verify

# Re-query rather than trusting cmdlet success; this is what a rescan will see.
$remaining = 0
if (-not $WhatIfPreference) {
    foreach ($t in $targets) {
        try {
            foreach ($reg in @(Get-OutdatedRegistration -Name $t.Name -Minimum $t.MinimumVersion)) {
                if ($SkipSids -contains $reg.Sid) { continue }
                $remaining++
                Write-Log ("Still outdated: {0} for {1}" -f $reg.PackageFullName, $reg.Sid) -Level WARN
            }
        }
        catch {
            Write-Log ("{0}: verification failed: {1}" -f $t.Name, $_.Exception.Message) -Level ERROR
            $enumError = $true
        }
    }
}

#endregion Verify

$count = @{}
foreach ($o in 'Removed', 'Failed', 'Skipped', 'WhatIf') {
    $count[$o] = @($details | Where-Object { $_.Outcome -eq $o }).Count
}

$summary = [pscustomobject]@{
    Computer  = $env:COMPUTERNAME
    Mode      = $mode
    Removed   = $count['Removed']
    Failed    = $count['Failed']
    Skipped   = $count['Skipped']
    WhatIf    = $count['WhatIf']
    Remaining = $remaining
    LogPath   = $LogPath
    Details   = $details.ToArray()
}

Write-Log ("End: removed {0}, failed {1}, skipped {2}, whatif {3}, remaining {4}." -f $summary.Removed, $summary.Failed, $summary.Skipped, $summary.WhatIf, $summary.Remaining)
$summary

if ($enumError -or $summary.Failed -gt 0 -or $remaining -gt 0) { exit 1 }
exit 0
