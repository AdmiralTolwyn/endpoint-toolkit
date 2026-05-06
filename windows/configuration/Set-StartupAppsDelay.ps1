<#
.SYNOPSIS
    Reverts the Windows 11 "wait-for-idle" startup-app delay so that
    Run / RunOnce / Startup-folder apps launch promptly after sign-in.

.DESCRIPTION
    On recent Windows 11 builds, Explorer no longer launches registered
    startup apps after a fixed delay. Instead it waits for the system to reach
    an "idle state" before kicking them off, gated by the registry value:

        HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Serialize
            WaitForIdleState (DWORD)   1 = wait for idle (new default)
                                       0 = legacy behaviour (no idle wait)
            StartupDelayInMSec (DWORD) legacy fixed delay in ms (default 10000)

    On busy devices that never go idle immediately after logon (Defender
    scan, OneDrive sync, GPO, Intune ESP, sensor agents, etc.), this can
    leave Outlook / Teams / Word / Excel apparently "not starting" for
    several minutes after the user signs in.

    This script reverts the new behaviour by writing:

        WaitForIdleState   = 0       (always - this is the actual fix)
        StartupDelayInMSec = 10000   (optional, restores the documented
                                      pre-Build 22621.1555 default of 10 s)

    to HKCU for the running user, and - when invoked with admin rights:
      * -IncludeDefaultUser also stamps the same values into the Default User
        hive (NTUSER.DAT) so brand-new user profiles created on the device
        after the script runs inherit the fix (image / shared / kiosk
        scenarios, or any device where new profiles still get created).
      * -IncludeLoadedUsers enumerates HKEY_USERS and patches every loaded
        user hive (HKU\<SID>) it finds, skipping built-in service accounts
        (SYSTEM / LOCAL SERVICE / NETWORK SERVICE), the .DEFAULT alias, and
        _Classes subkeys. This covers the currently signed-in user(s) when
        the script runs as SYSTEM.

    Designed to ship as an Intune platform script with no parameters in
    either user context (single-user fix) or SYSTEM context (image /
    multi-user fix). Behaviour adapts automatically to the run context:

      * Non-elevated  -> patch HKCU only.
      * Elevated/SYSTEM -> additionally patch the Default User hive AND
        every currently loaded user hive (HKU\<SID>) without any flags.

    No reboot required; the change takes effect at the user's next sign-in.

.PARAMETER ResetStartupDelay
    Also write StartupDelayInMSec = 10000 (the documented pre-Build
    22621.1555 default of 10 seconds). Defaults to ON automatically when
    the script runs elevated; off in non-elevated runs to preserve any
    custom value an admin / GPO / user has set in HKCU. Pass
    -ResetStartupDelay:$false from an elevated session to opt out.

.PARAMETER IncludeDefaultUser
    Force the Default User hive (NTUSER.DAT) to be patched. Defaults to ON
    automatically when the script runs elevated. Pass
    -IncludeDefaultUser:$false from an elevated session to opt out.

.PARAMETER IncludeLoadedUsers
    Force enumeration of HKEY_USERS to patch every loaded user hive
    (HKU\<SID>), skipping built-in service accounts and _Classes subkeys.
    Defaults to ON automatically when the script runs elevated. Pass
    -IncludeLoadedUsers:$false from an elevated session to opt out.

.PARAMETER Revert
    Restore the Windows 11 default behaviour by removing the values
    (WaitForIdleState and, when -ResetStartupDelay is also supplied,
    StartupDelayInMSec) from HKCU - and from the Default User hive and
    every loaded user hive when their respective include-* flags are
    effective (which is automatic under elevation). Use this to back the
    change out.

.PARAMETER LogDirectory
    Directory for the log file. Default: the Intune Management Extension
    logs folder (%ProgramData%\Microsoft\IntuneManagementExtension\Logs)
    when writable, otherwise %TEMP%. Created automatically if missing.
    Override with any folder.

.NOTES
    File:    windows/servicing/StartupAppsDelay/Set-StartupAppsDelay.ps1
    Author:  Anton Romanyuk
    Version: 1.0.0
    Requires: PowerShell 5.1+. User context patches HKCU only. Elevated /
              SYSTEM context additionally patches Default User and all
              loaded HKU\<SID> hives - automatically, no parameters needed.

    Deployment (Intune platform script - one assignment, no parameters):
        Run this script using the logged on credentials : choose context
        Enforce script signature check                  : No
        Run script in 64-bit PowerShell host            : Yes

        - 'No' (SYSTEM): patches Default User + every signed-in HKU\<SID>.
          Recommended for shared / multi-user devices.
        - 'Yes' (user): patches the signed-in user's HKCU only. Use for
          single-user AAD-joined laptops.

    Build context:
        - Applies to recent Windows 11 builds that gate startup apps on an
          idle-state check. On older builds that only honour
          StartupDelayInMSec, writing WaitForIdleState is a harmless no-op.

    Exit codes:
        0  - Success (or no-op / nothing to do under -WhatIf)
        1  - Unhandled failure (see log)

.DISCLAIMER
    This script is provided "AS IS" with no warranties and confers no rights.
    It is not supported under any Microsoft standard support program or
    service. Use of this script is entirely at your own risk. The customer is
    solely responsible for testing and validating this script in their
    environment before deploying to production.

.EXAMPLE
    .\Set-StartupAppsDelay.ps1
    No-arg run. Under user context: patches HKCU only. Under SYSTEM /
    elevated: also patches the Default User hive and every loaded
    HKU\<SID>. Ship this verbatim as an Intune platform script.

.EXAMPLE
    .\Set-StartupAppsDelay.ps1 -ResetStartupDelay
    Same as above, but ALSO writes StartupDelayInMSec = 10000 to restore
    the documented pre-22621.1555 default of 10 seconds.

.EXAMPLE
    .\Set-StartupAppsDelay.ps1 -IncludeLoadedUsers:$false
    Elevated run, but skip the HKU\<SID> sweep (e.g. only seed the Default
    User hive on a sysprep'd image where no users are loaded yet).

.EXAMPLE
    .\Set-StartupAppsDelay.ps1 -Revert
    Removes the WaitForIdleState value written by a previous run (and
    StartupDelayInMSec when -ResetStartupDelay is also supplied), restoring
    the Windows 11 default behaviour. Honours the same auto-elevation
    scope as the apply path.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [switch]$ResetStartupDelay,
    [switch]$IncludeDefaultUser,
    [switch]$IncludeLoadedUsers,
    [switch]$Revert,
    [string]$LogDirectory
)

$ErrorActionPreference = 'Stop'

# -----------------------------------------------------------------------------
# AUTO-DEFAULTS BASED ON CONTEXT
# -----------------------------------------------------------------------------
# Intune platform scripts cannot pass parameters - the .ps1 is uploaded as-is.
# So under elevation (SYSTEM / admin), default to patching the Default User
# hive AND every loaded HKU\<SID>. Caller can still opt out with
# -IncludeDefaultUser:$false / -IncludeLoadedUsers:$false in interactive runs.
$IsElevated = ([Security.Principal.WindowsPrincipal]::new(
                [Security.Principal.WindowsIdentity]::GetCurrent())
              ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if ($IsElevated) {
    if (-not $PSBoundParameters.ContainsKey('ResetStartupDelay'))   { $ResetStartupDelay   = $true }
    if (-not $PSBoundParameters.ContainsKey('IncludeDefaultUser'))  { $IncludeDefaultUser  = $true }
    if (-not $PSBoundParameters.ContainsKey('IncludeLoadedUsers'))  { $IncludeLoadedUsers  = $true }
}

# -----------------------------------------------------------------------------
# CONSTANTS
# -----------------------------------------------------------------------------
$SerializeSubKey   = 'Software\Microsoft\Windows\CurrentVersion\Explorer\Serialize'
$DefaultUserHive   = Join-Path $env:SystemDrive 'Users\Default\NTUSER.DAT'
$DefaultUserMount  = 'HKLM\TEMP_DefaultUser_StartupAppsDelay'

# -----------------------------------------------------------------------------
# LOGGING
# -----------------------------------------------------------------------------
# Default to the Intune Management Extension logs folder when writable
# (SYSTEM context always; user context only if the user has write access).
# Fall back to %TEMP% otherwise. -LogDirectory overrides both.
$ScriptName = 'Set-StartupAppsDelay'
if (-not $LogDirectory) {
    $imeLogs = Join-Path $env:ProgramData 'Microsoft\IntuneManagementExtension\Logs'
    try {
        if (-not (Test-Path -LiteralPath $imeLogs)) {
            [void](New-Item -Path $imeLogs -ItemType Directory -Force -ErrorAction Stop)
        }
        # Probe write access (a non-elevated user can read the folder but not write to it).
        $probe = Join-Path $imeLogs ("_writeprobe_{0}.tmp" -f ([guid]::NewGuid().ToString('N')))
        Set-Content -LiteralPath $probe -Value '' -ErrorAction Stop
        Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
        $LogDirectory = $imeLogs
    }
    catch {
        $LogDirectory = $env:TEMP
    }
}
if (-not (Test-Path -LiteralPath $LogDirectory)) {
    [void](New-Item -Path $LogDirectory -ItemType Directory -Force)
}
$LogFile = Join-Path $LogDirectory ("PS_{0}_{1}.log" -f $ScriptName, (Get-Date -Format 'yyyyMMdd_HHmmss'))

function Write-Log {
<#
.SYNOPSIS
    Writes a timestamped, level-tagged line to both the console and the log file.
.DESCRIPTION
    Uniform logger used by the rest of the script. Format on disk and on console:
        [yyyy-MM-dd HH:mm:ss] [LEVEL] message
    Console output is colour-coded by level. File writes use SilentlyContinue so
    a transient lock on the log file never aborts the pipeline.
.PARAMETER Message
    Free-form text to record.
.PARAMETER Level
    INFO | WARN | ERROR | SUCCESS. Default INFO.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS')][string]$Level = 'INFO'
    )
    $line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Add-Content -LiteralPath $LogFile -Value $line -ErrorAction SilentlyContinue
    $color = switch ($Level) {
        'WARN'    { 'Yellow' }
        'ERROR'   { 'Red' }
        'SUCCESS' { 'Green' }
        default   { 'Gray' }
    }
    Write-Host $line -ForegroundColor $color
}

# -----------------------------------------------------------------------------
# HELPERS
# -----------------------------------------------------------------------------
function Get-WindowsBuildString {
<#
.SYNOPSIS
    Returns "<MajorBuild>.<UBR>" for the running OS.
.DESCRIPTION
    Reads CurrentBuildNumber + UBR from
    HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion. Used purely for
    informational logging - the script writes the keys regardless of build,
    because they are harmless on builds that don't honour them.
.OUTPUTS
    [string]
#>
    try {
        $cv = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop
        '{0}.{1}' -f $cv.CurrentBuildNumber, $cv.UBR
    }
    catch {
        'unknown'
    }
}

function Set-SerializeValues {
<#
.SYNOPSIS
    Writes (or removes, when -RevertMode is set) the Serialize subkey values
    that control startup-app launch timing under the supplied registry hive
    root.
.DESCRIPTION
    Centralised so the same logic applies whether we are touching HKCU for
    the running user or the temporarily mounted Default User hive. Honours
    $WhatIfPreference via Set-ItemProperty / Remove-ItemProperty (both
    support ShouldProcess natively).

    When -RevertMode is supplied, removes WaitForIdleState (and, when
    -AlsoStartupDelay is also supplied, StartupDelayInMSec). Missing values
    are tolerated.

    Otherwise, sets WaitForIdleState=0 and (when -AlsoStartupDelay is
    supplied) StartupDelayInMSec=10000 (legacy default). The Serialize
    subkey is created on demand.
.PARAMETER HiveRoot
    PowerShell registry path of the hive root, e.g. 'HKCU:' or
    'Registry::HKEY_USERS\TEMP_DefaultUser_StartupAppsDelay'.
.PARAMETER Label
    Friendly label for log lines (e.g. 'HKCU' or 'Default User').
.PARAMETER AlsoStartupDelay
    Also touch StartupDelayInMSec.
.PARAMETER RevertMode
    Remove instead of write.
#>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$HiveRoot,
        [Parameter(Mandatory)][string]$Label,
        [switch]$AlsoStartupDelay,
        [switch]$RevertMode
    )

    $keyPath = Join-Path $HiveRoot $SerializeSubKey

    if ($RevertMode) {
        if (-not (Test-Path -LiteralPath $keyPath)) {
            Write-Log "[$Label] Serialize subkey does not exist - nothing to revert."
            return
        }
        foreach ($name in @('WaitForIdleState') + $(if ($AlsoStartupDelay) { 'StartupDelayInMSec' } else { @() })) {
            $existing = Get-ItemProperty -LiteralPath $keyPath -Name $name -ErrorAction SilentlyContinue
            if ($null -eq $existing) {
                Write-Log "[$Label] $name not present - skip."
                continue
            }
            if ($PSCmdlet.ShouldProcess("$keyPath\$name", 'Remove value')) {
                Remove-ItemProperty -LiteralPath $keyPath -Name $name -Force
                Write-Log "[$Label] Removed $name." -Level SUCCESS
            }
        }
        return
    }

    # Write mode
    if (-not (Test-Path -LiteralPath $keyPath)) {
        if ($PSCmdlet.ShouldProcess($keyPath, 'Create subkey')) {
            [void](New-Item -Path $keyPath -Force)
            Write-Log "[$Label] Created subkey $keyPath."
        }
    }

    $values = [ordered]@{ WaitForIdleState = 0 }
    if ($AlsoStartupDelay) { $values['StartupDelayInMSec'] = 10000 }

    foreach ($name in $values.Keys) {
        $desired = [int]$values[$name]
        $current = (Get-ItemProperty -LiteralPath $keyPath -Name $name -ErrorAction SilentlyContinue).$name
        if ($null -ne $current -and [int]$current -eq $desired) {
            Write-Log "[$Label] $name already $desired - no change."
            continue
        }
        if ($PSCmdlet.ShouldProcess("$keyPath\$name", "Set DWORD = $desired")) {
            Set-ItemProperty -LiteralPath $keyPath -Name $name -Value $desired -Type DWord -Force
            Write-Log "[$Label] Set $name = $desired (was: $(if ($null -eq $current) { '(absent)' } else { $current }))." -Level SUCCESS
        }
    }
}

function Get-LoadedUserHive {
<#
.SYNOPSIS
    Enumerates currently loaded interactive user hives under HKEY_USERS.
.DESCRIPTION
    Returns one PSCustomObject per loaded user hive (HKU\<SID>) that
    represents a real interactive user, suitable for writing per-user
    settings to. Filters out:
      - Built-in service accounts: SYSTEM (S-1-5-18), LOCAL SERVICE
        (S-1-5-19), NETWORK SERVICE (S-1-5-20).
      - The .DEFAULT alias (synonym for HKU\S-1-5-18).
      - _Classes companion subkeys (e.g. S-1-5-21-...-1001_Classes).
      - Any SID where Software\Microsoft\Windows\CurrentVersion\Explorer
        is missing (loaded but not a real user profile).

    The friendly account name is best-effort resolved via LookupAccountSid;
    resolution failure is non-fatal and the SID is returned as the label.
.OUTPUTS
    PSCustomObject with properties: Sid (string), Account (string),
    HiveRoot (PowerShell registry path 'Registry::HKEY_USERS\<SID>').
#>
    [CmdletBinding()]
    param()

    $skipSids = @('S-1-5-18','S-1-5-19','S-1-5-20','.DEFAULT')

    Get-ChildItem -LiteralPath 'Registry::HKEY_USERS' -ErrorAction SilentlyContinue | ForEach-Object {
        $sid = $_.PSChildName
        if ($skipSids -contains $sid) { return }
        if ($sid -like '*_Classes')  { return }

        # Sanity check: a real user hive has Software\...\Explorer.
        $probe = "Registry::HKEY_USERS\$sid\Software\Microsoft\Windows\CurrentVersion\Explorer"
        if (-not (Test-Path -LiteralPath $probe)) { return }

        $account = $sid
        try {
            $sidObj  = [Security.Principal.SecurityIdentifier]::new($sid)
            $account = $sidObj.Translate([Security.Principal.NTAccount]).Value
        }
        catch {
            # Orphan / unresolvable SID - still patch it, but log under the SID.
        }

        [pscustomobject]@{
            Sid      = $sid
            Account  = $account
            HiveRoot = "Registry::HKEY_USERS\$sid"
        }
    }
}

function Invoke-OnDefaultUserHive {
<#
.SYNOPSIS
    Loads %SystemDrive%\Users\Default\NTUSER.DAT, runs the supplied script
    block against it, then unloads it - guaranteeing the hive is unloaded
    even on failure.
.DESCRIPTION
    Uses reg.exe load / unload (Microsoft.Win32.RegistryHive does not work
    cleanly across runspaces from PowerShell). Calls [GC]::Collect() before
    unload to drop any RegistryKey handles PowerShell may still be holding,
    which otherwise blocks reg unload with "access denied".

    The callback receives the temporarily-mounted hive path as its single
    argument, e.g. 'Registry::HKEY_LOCAL_MACHINE\TEMP_DefaultUser_...'.
.PARAMETER ScriptBlock
    Code to run while the hive is mounted. Single argument: hive root path.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][scriptblock]$ScriptBlock
    )

    if (-not (Test-Path -LiteralPath $DefaultUserHive)) {
        Write-Log "Default User hive not found at '$DefaultUserHive' - skipping." -Level WARN
        return
    }

    Write-Log "Loading Default User hive '$DefaultUserHive' as '$DefaultUserMount' ..."
    $loadOut = & reg.exe load $DefaultUserMount $DefaultUserHive 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "reg load failed (exit $LASTEXITCODE): $loadOut"
    }

    $hiveRoot = "Registry::HKEY_LOCAL_MACHINE\$($DefaultUserMount.Substring(5))"  # strip 'HKLM\'
    try {
        & $ScriptBlock $hiveRoot
    }
    finally {
        # Drop any lingering handles before unload.
        [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
        Write-Log "Unloading Default User hive ..."
        $unloadOut = & reg.exe unload $DefaultUserMount 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Log "reg unload failed (exit $LASTEXITCODE): $unloadOut" -Level WARN
        }
    }
}

# -----------------------------------------------------------------------------
# MAIN
# -----------------------------------------------------------------------------
$exitCode = 0
try {
    Write-Log '=== Set-StartupAppsDelay v1.0.0 ===' -Level SUCCESS
    Write-Log ("OS build              : {0}" -f (Get-WindowsBuildString))
    Write-Log ("Running as            : {0} (Elevated={1})" -f [Environment]::UserName, $IsElevated)
    Write-Log ("Mode                  : {0}" -f $(if ($Revert) { 'REVERT' } else { 'APPLY' }))
    Write-Log ("Reset StartupDelay    : {0}" -f [bool]$ResetStartupDelay)
    Write-Log ("Include Default User  : {0}" -f [bool]$IncludeDefaultUser)
    Write-Log ("Include Loaded Users  : {0}" -f [bool]$IncludeLoadedUsers)
    Write-Log ("Log file              : {0}" -f $LogFile)

    # 1) Current user (HKCU) - works in any context, including SYSTEM
    #    (SYSTEM's HKCU is harmless to write to).
    Set-SerializeValues -HiveRoot 'HKCU:' -Label 'HKCU' `
                        -AlsoStartupDelay:$ResetStartupDelay `
                        -RevertMode:$Revert

    # 2) Default User hive - applies to NEW profiles created after this point.
    #    Auto-enabled under elevation; opt-out via -IncludeDefaultUser:$false.
    if ($IncludeDefaultUser) {
        if (-not $IsElevated) {
            Write-Log '-IncludeDefaultUser requires elevation - skipping Default User hive.' -Level WARN
        }
        else {
            Invoke-OnDefaultUserHive -ScriptBlock {
                param($HiveRoot)
                Set-SerializeValues -HiveRoot $HiveRoot -Label 'Default User' `
                                    -AlsoStartupDelay:$ResetStartupDelay `
                                    -RevertMode:$Revert
            }
        }
    }

    # 3) Loaded user hives - covers signed-in user(s) when the script runs
    #    as SYSTEM. From user context, HKCU above already covers it.
    #    Auto-enabled under elevation; opt-out via -IncludeLoadedUsers:$false.
    if ($IncludeLoadedUsers) {
        if (-not $IsElevated) {
            Write-Log '-IncludeLoadedUsers requires elevation - skipping HKEY_USERS scan.' -Level WARN
        }
        else {
            $loaded = @(Get-LoadedUserHive)
            Write-Log ("Found {0} loaded user hive(s) under HKEY_USERS" -f $loaded.Count)
            foreach ($u in $loaded) {
                Set-SerializeValues -HiveRoot $u.HiveRoot -Label ("HKU " + $u.Account) `
                                    -AlsoStartupDelay:$ResetStartupDelay `
                                    -RevertMode:$Revert
            }
        }
    }

    Write-Log '=== Done. Change takes effect at next sign-in (no reboot required). ===' -Level SUCCESS
}
catch {
    Write-Log "FATAL: $($_.Exception.Message)" -Level ERROR
    Write-Log $_.ScriptStackTrace -Level ERROR
    $exitCode = 1
}

exit $exitCode
