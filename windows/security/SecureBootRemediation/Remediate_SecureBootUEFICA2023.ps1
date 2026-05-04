<#
.SYNOPSIS
    Remediates non-compliant Windows UEFI CA 2023 Secure Boot certificate
    state by arming and/or triggering the Secure Boot Update scheduled task.

.DESCRIPTION
    Two-pronged remediation logic based on
    HKLM\SYSTEM\CurrentControlSet\Control\SecureBoot\AvailableUpdates:

      A. Value missing OR 0           -> Initial arm.
         - Set AvailableUpdates = 0x5944
         - Start scheduled task \Microsoft\Windows\PI\Secure-Boot-Update

      B. Value non-zero AND not 0x4000 -> Resume / progress.
         - Do NOT modify AvailableUpdates (preserve in-flight state such as
           0x4100 staged-for-reboot, 0x4104 KEK-pending, etc.)
         - Start the same scheduled task to drive the next step.

      C. Value == 0x4000              -> Already complete, no-op.

    Hard-blocker reporting markers (exit 1 with NO side effects) on devices
    that cannot be remediated from the OS. Exit 1 ensures Intune does NOT
    falsely mark the device as 'Remediation successful' / 'Fixed' on the
    PR dashboard - the device IS still non-compliant, we just chose not to
    make a system change. No registry writes, no scheduled-task triggers:
      * HighConfidenceOptOut = 1 -> NON-COMPLIANT-NOT-ACTIONABLE
        (admin-managed exclusion).
      * Event 1803 present -> NON-COMPLIANT-NOT-ACTIONABLE
        (missing KEK - OEM responsibility).
      * Event 1802 present -> NON-COMPLIANT-NOT-ACTIONABLE
        (known firmware issue / SkipReason: KI_<n>).
      * Event 1800 without 1808 -> NON-COMPLIANT-PENDING-REBOOT
        (update staged, awaiting reboot to fire 1808).

    Operational warnings (logged; remediation still attempts where possible):
      * Secure-Boot-Update scheduled task missing/disabled - aborts cleanly
        instead of writing the registry only to fail at task start.
      * CanAttemptUpdateAfter FILETIME in the future - firmware throttle
        active; trigger will be a no-op until the timestamp elapses.
      * AvailableUpdatesPolicy disagrees with AvailableUpdates - GPO/MDM
        is the source of truth; direct AvailableUpdates writes may be
        reverted on the next policy refresh.

    Logs are written in CMTrace format to:
        %ProgramData%\Microsoft\IntuneManagementExtension\Logs\PR_SecureBootUEFICA2023.log

    Intune Proactive Remediation settings:
      Run this script using the logged-on credentials : No
      Enforce script signature check                  : No
      Run script in 64-bit PowerShell                 : Yes

    A reboot is required after the scheduled task progresses past the staged
    state. Pair with a separate reboot policy / user notification mechanism.

    References:
      KB 5016061 - Secure Boot DBX update event reference
      KB 5072718 - Sample Secure Boot Inventory Data Collection script
      KB 5084567 - Sample Secure Boot E2E Automation Guide

.OUTPUTS
    System.Int32 exit code:
      0  = REMEDIATED or NO-OP (Intune marks device as 'Remediation successful')
      1  = Any other state - PREREQ / ABORT / FAIL / NON-COMPLIANT-* (Intune
           keeps the device in the work queue)

    Single-line STDOUT summary suitable for the Intune PR remediation column.
    Marker prefixes: REMEDIATED, NO-OP, NON-COMPLIANT-NOT-ACTIONABLE,
    NON-COMPLIANT-PENDING-REBOOT, PREREQ, ABORT, FAIL.

.NOTES
    Author:   Anton Romanyuk
    Version:  1.2
    Date:     2026-05-04
    Context:  Windows UEFI CA 2023 Secure Boot certificate rollout
    Requires: Windows PowerShell 5.1, 64-bit

    DISCLAIMER:
    THIS SCRIPT IS PROVIDED "AS-IS" WITHOUT WARRANTY OF ANY KIND.
    USE AT YOUR OWN RISK.
#>

#Requires -Version 5.1

# ─────────────────────────────────────────────────────────────────────────────
# Configuration
# ─────────────────────────────────────────────────────────────────────────────

# Log file is shared with Detect_SecureBootUEFICA2023.ps1 so detection and
# remediation entries appear in the same CMTrace timeline.
$Script:LogDir     = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs"
$Script:LogName    = 'PR_SecureBootUEFICA2023'
$Script:LogFile    = Join-Path $Script:LogDir "$Script:LogName.log"
$Script:LogMaxSize = 250KB                                   # rotation threshold
$Script:Component  = 'Remediate'                             # CMTrace component column

# Registry locations driving the Secure Boot servicing state machine.
$Script:RegPathRoot       = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot'
$Script:RegPathServicing  = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing'
$Script:RegPathDeviceAttr = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing\DeviceAttributes'
$Script:ArmValue          = 0x5944                            # Bitmask used to kick off the Secure Boot update
$Script:CompliantValue    = 0x4000                            # Terminal "complete" state

# Event log signals (provider Microsoft-Windows-TPM-WMI in the System log).
#   1808 - Update completed successfully.
#   1803 - Missing KEK update (OEM responsibility - hard blocker).
#   1802 - Known firmware issue, SkipReason: KI_<n> (hard blocker).
#   1800 - Reboot pending (suppress re-trigger).
$Script:EventProvider       = 'Microsoft-Windows-TPM-WMI'
$Script:HardBlockerEvents   = @(1802, 1803)
$Script:RebootPendingEvent  = 1800
$Script:AllSecureBootEvents = @(1800, 1802, 1803, 1808)
$Script:EventQueryMax       = 50

# Scheduled task that performs the actual servicing pass.
$Script:TaskPath = '\Microsoft\Windows\PI\'
$Script:TaskName = 'Secure-Boot-Update'

# ─────────────────────────────────────────────────────────────────────────────
#region Logging
# ─────────────────────────────────────────────────────────────────────────────

function Write-Log {
    <#
    .SYNOPSIS
        Writes a CMTrace-formatted log entry to the Intune ME log directory.

    .DESCRIPTION
        Emits a single-line CMTrace-compatible log record so the file can be
        opened with cmtrace.exe / Configuration Manager log viewer with full
        coloring and severity classification. Automatically creates the log
        directory if missing and rotates the active log when it exceeds
        $Script:LogMaxSize bytes.

    .PARAMETER Message
        The free-text message to record. Required.

    .PARAMETER Level
        Severity / category. One of DEBUG, INFO, WARN, ERROR, SUCCESS, STEP.
        Defaults to INFO. Maps to CMTrace numeric severity:
          DEBUG/INFO/SUCCESS/STEP -> 1, WARN -> 2, ERROR -> 3.
        Use DEBUG for verbose function entry/exit and intermediate values.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet('DEBUG', 'INFO', 'WARN', 'ERROR', 'SUCCESS', 'STEP')]
        [string]$Level = 'INFO'
    )

    # Build CMTrace timestamp/date (millisecond precision, MM-dd-yyyy).
    $ts = Get-Date -Format 'HH:mm:ss.fffffff'
    $dt = Get-Date -Format 'MM-dd-yyyy'

    # Map text level -> CMTrace severity integer.
    $severity = switch ($Level) { 'ERROR' { 3 } 'WARN' { 2 } default { 1 } }

    # Caller name shown in CMTrace "component" column. Falls back to the
    # script-scoped component label when invoked outside a named function.
    $caller = if ($MyInvocation.MyCommand.Name) { $MyInvocation.MyCommand.Name } else { $Script:Component }

    # CMTrace single-line entry format. Level is also embedded in the message
    # text for fast grep / minimap filtering since the CMTrace XML record
    # only carries numeric severity (1/2/3), not the friendly label.
    $tagged = '[{0}] {1}' -f $Level, $Message
    $entry = '<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="{4}" type="{5}" thread="{6}" file="{7}">' -f
        $tagged, $ts, $dt, $caller, $env:USERNAME, $severity, $PID, $Script:ScriptName

    # Ensure the log directory exists before attempting to write.
    if (-not (Test-Path $Script:LogDir)) {
        try { New-Item -Path $Script:LogDir -ItemType Directory -Force -ErrorAction Stop | Out-Null }
        catch { Write-Warning "Cannot create log directory: $_"; return }
    }

    # Append to the log file. Use default ANSI encoding so cmtrace renders
    # entries correctly without a BOM appearing mid-file on rotation.
    try { $entry | Out-File -FilePath $Script:LogFile -Append -NoClobber -Force -Encoding default -ErrorAction Stop }
    catch { Write-Warning "Log write failed: $_" }

    # Rotate when the active log exceeds the configured size threshold.
    # Archive copy is NTFS-compressed to save space on large fleets.
    if ((Test-Path $Script:LogFile) -and (Get-Item $Script:LogFile).Length -ge $Script:LogMaxSize) {
        $archive = Join-Path $Script:LogDir "$Script:LogName-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
        try {
            Copy-Item $Script:LogFile $archive -Force
            Remove-Item $Script:LogFile -Force
            Compact /C $archive 2>$null | Out-Null
        }
        catch { Write-Warning "Log rotation failed: $_" }
    }
}

#endregion

# ─────────────────────────────────────────────────────────────────────────────
#region Helpers
# ─────────────────────────────────────────────────────────────────────────────

function Test-Is64BitPS {
    <#
    .SYNOPSIS
        Returns $true when the current PowerShell host is a 64-bit process.

    .DESCRIPTION
        On 64-bit Windows the registry views differ between 32-bit and 64-bit
        processes. We must run 64-bit so the SecureBoot keys are written to
        the native HKLM hive (not the WOW6432Node redirected view).
    #>
    $is64 = [Environment]::Is64BitProcess
    Write-Log "Test-Is64BitPS -> $is64 (PSArch=$([IntPtr]::Size * 8)-bit, OSArch=$([Environment]::Is64BitOperatingSystem))" -Level DEBUG
    return $is64
}

function Get-RegValue {
    <#
    .SYNOPSIS
        Safely reads a single registry value, returning $null if absent.

    .DESCRIPTION
        Wraps Get-ItemProperty in try/catch + Test-Path so missing keys or
        missing values return $null instead of throwing. Lets the caller use
        a clean $null comparison without needing -ErrorAction handling.

    .PARAMETER Path
        The full registry path, e.g. 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot'.

    .PARAMETER Name
        The value name to read, e.g. 'AvailableUpdates'.

    .OUTPUTS
        The value (any type) or $null if the path / name does not exist.
    #>
    param(
        [Parameter(Mandatory = $true)] [string]$Path,
        [Parameter(Mandatory = $true)] [string]$Name
    )
    Write-Log "Get-RegValue ENTER (Path='$Path', Name='$Name')" -Level DEBUG
    try {
        if (-not (Test-Path $Path)) {
            Write-Log "Get-RegValue: path does not exist -> returning `$null" -Level DEBUG
            return $null
        }
        $val = (Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop).$Name
        $disp = if ($null -eq $val) { '<null>' } elseif ($val -is [int] -or $val -is [long]) { "{0} (0x{0:X})" -f $val } else { "'$val'" }
        Write-Log "Get-RegValue EXIT -> $disp (type=$(if ($null -ne $val) { $val.GetType().Name } else { 'null' }))" -Level DEBUG
        return $val
    }
    catch {
        Write-Log "Get-RegValue: exception reading '$Name' -> $($_.Exception.Message)" -Level DEBUG
        return $null
    }
}

function Set-RegDword {
    <#
    .SYNOPSIS
        Creates or overwrites a REG_DWORD value, creating the parent key if
        necessary.

    .PARAMETER Path
        The full registry path.

    .PARAMETER Name
        The value name to write.

    .PARAMETER Value
        The 32-bit signed integer to store as DWORD.

    .NOTES
        Throws on failure so the caller's try/catch can record a clean
        Branch-A registry-write error.
    #>
    param(
        [Parameter(Mandatory = $true)] [string]$Path,
        [Parameter(Mandatory = $true)] [string]$Name,
        [Parameter(Mandatory = $true)] [int]$Value
    )
    Write-Log ("Set-RegDword ENTER (Path='{0}', Name='{1}', Value={2} / 0x{2:X})" -f $Path, $Name, $Value) -Level DEBUG
    if (-not (Test-Path $Path)) {
        Write-Log "Set-RegDword: parent key missing, creating '$Path'" -Level DEBUG
        New-Item -Path $Path -Force -ErrorAction Stop | Out-Null
    }
    New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType DWord -Force -ErrorAction Stop | Out-Null
    Write-Log "Set-RegDword EXIT (write OK)" -Level DEBUG
}

function Get-SecureBootEventFlags {
    <#
    .SYNOPSIS
        Queries the System log once and returns hard-blocker / reboot-pending
        flags + the latest event message snippet (BucketId / SkipReason).

    .OUTPUTS
        Hashtable with keys: Has1808, Has1803, Has1802, Has1800,
        SkipReason (KI_<n>), Counts (id -> count).
    #>
    Write-Log ("Get-SecureBootEventFlags ENTER (IDs={0}, Max={1})" -f ($Script:AllSecureBootEvents -join ','), $Script:EventQueryMax) -Level DEBUG
    $r = @{
        Has1808    = $false
        Has1803    = $false
        Has1802    = $false
        Has1800    = $false
        SkipReason = $null
        Counts     = @{}
    }
    try {
        $events = @(Get-WinEvent -FilterHashtable @{
            LogName      = 'System'
            ProviderName = $Script:EventProvider
            Id           = $Script:AllSecureBootEvents
        } -MaxEvents $Script:EventQueryMax -ErrorAction SilentlyContinue)
    }
    catch {
        Write-Log "Get-SecureBootEventFlags: query exception -> $($_.Exception.Message)" -Level DEBUG
        return $r
    }
    foreach ($id in $Script:AllSecureBootEvents) { $r.Counts[$id] = @($events | Where-Object { $_.Id -eq $id }).Count }
    $r.Has1808 = ($r.Counts[1808] -gt 0)
    $r.Has1803 = ($r.Counts[1803] -gt 0)
    $r.Has1802 = ($r.Counts[1802] -gt 0)
    $r.Has1800 = ($r.Counts[1800] -gt 0)
    $latest1802 = @($events | Where-Object { $_.Id -eq 1802 } | Select-Object -First 1)
    if ($latest1802.Count -gt 0 -and $latest1802[0].Message -match 'SkipReason:\s*(KI_\d+)') {
        $r.SkipReason = $matches[1]
    }
    Write-Log ("Get-SecureBootEventFlags EXIT (1808={0}, 1803={1}, 1802={2}, 1800={3}, SkipReason='{4}')" -f
        $r.Counts[1808], $r.Counts[1803], $r.Counts[1802], $r.Counts[1800], $r.SkipReason) -Level DEBUG
    return $r
}

function Get-SecureBootTaskState {
    <#
    .SYNOPSIS
        Returns information about the Secure-Boot-Update scheduled task.

    .OUTPUTS
        Hashtable with keys:
          Exists  - $true if the task is present
          State   - 'Ready', 'Disabled', 'Running', 'Unknown', or 'Missing'
          Enabled - $true if State -in @('Ready','Running')
    #>
    Write-Log "Get-SecureBootTaskState ENTER (TaskPath='$($Script:TaskPath)', TaskName='$($Script:TaskName)')" -Level DEBUG
    $r = @{ Exists = $false; State = 'Missing'; Enabled = $false }
    try {
        $task = Get-ScheduledTask -TaskPath $Script:TaskPath -TaskName $Script:TaskName -ErrorAction Stop
        $r.Exists  = $true
        $r.State   = "$($task.State)"
        $r.Enabled = ($r.State -in @('Ready','Running'))
        Write-Log "Get-SecureBootTaskState EXIT (Exists=True, State='$($r.State)', Enabled=$($r.Enabled))" -Level DEBUG
    }
    catch {
        Write-Log "Get-SecureBootTaskState: not found / not accessible -> $($_.Exception.Message)" -Level DEBUG
    }
    return $r
}

function Get-CanAttemptUpdateAfter {
    <#
    .SYNOPSIS
        Returns the CanAttemptUpdateAfter throttle as a UTC DateTime, or
        $null if absent / unreadable.

    .DESCRIPTION
        Stored as REG_BINARY (8-byte FILETIME, little-endian) or REG_QWORD
        under SecureBoot\Servicing\DeviceAttributes. When set in the future,
        the firmware refuses to attempt the update again until that time -
        triggering the scheduled task during the throttle window is a no-op.
    #>
    Write-Log "Get-CanAttemptUpdateAfter ENTER (Path='$($Script:RegPathDeviceAttr)')" -Level DEBUG
    $raw = Get-RegValue -Path $Script:RegPathDeviceAttr -Name 'CanAttemptUpdateAfter'
    if ($null -eq $raw) {
        Write-Log "Get-CanAttemptUpdateAfter EXIT -> `$null (value not set)" -Level DEBUG
        return $null
    }
    try {
        # Coerce Object[] (PS wraps REG_BINARY as a generic object array) to a
        # real byte[]. Get-ItemProperty returns boxed bytes when the value was
        # round-tripped through PSObject, so an explicit [byte[]] cast is needed.
        if ($raw -is [array] -and $raw -isnot [byte[]]) {
            try { $raw = [byte[]]$raw }
            catch {
                Write-Log "Get-CanAttemptUpdateAfter: cannot coerce $($raw.GetType().FullName) to byte[] -> $($_.Exception.Message)" -Level DEBUG
                return $null
            }
        }
        if ($raw -is [byte[]]) {
            if ($raw.Length -lt 8) {
                Write-Log "Get-CanAttemptUpdateAfter: byte[] too short ($($raw.Length) bytes, need 8)" -Level DEBUG
                return $null
            }
            $ft = [BitConverter]::ToInt64($raw, 0)
            $dt = [DateTime]::FromFileTimeUtc($ft)
        }
        elseif ($raw -is [long] -or $raw -is [int]) {
            $dt = [DateTime]::FromFileTimeUtc([int64]$raw)
        }
        else {
            Write-Log "Get-CanAttemptUpdateAfter: unexpected type $($raw.GetType().FullName)" -Level DEBUG
            return $null
        }
        $local = $dt.ToLocalTime()
        Write-Log ("Get-CanAttemptUpdateAfter EXIT -> {0}Z (local {1})" -f $dt.ToString('s'), $local.ToString('s')) -Level DEBUG
        return $dt
    }
    catch {
        Write-Log "Get-CanAttemptUpdateAfter: parse failed -> $($_.Exception.Message)" -Level DEBUG
        return $null
    }
}

function Start-SecureBootUpdateTask {
    <#
    .SYNOPSIS
        Triggers the built-in \Microsoft\Windows\PI\Secure-Boot-Update task.

    .DESCRIPTION
        Wraps Start-ScheduledTask with logging. Returns $true on success,
        $false on failure so the caller can produce a clean STDOUT summary
        for the Intune PR UI without surfacing a raw exception.

    .OUTPUTS
        System.Boolean - $true if the task was started, $false otherwise.
    #>
    Write-Log "Start-SecureBootUpdateTask ENTER (TaskPath='$($Script:TaskPath)', TaskName='$($Script:TaskName)')" -Level DEBUG
    # Pre-flight: confirm the task exists so a failure to start is properly
    # attributed (missing task vs. permission/task-engine error).
    try {
        $task = Get-ScheduledTask -TaskPath $Script:TaskPath -TaskName $Script:TaskName -ErrorAction Stop
        Write-Log "Scheduled task found (State=$($task.State))" -Level DEBUG
    }
    catch {
        Write-Log "Scheduled task not found or not accessible: $($_.Exception.Message)" -Level ERROR
        return $false
    }
    try {
        Start-ScheduledTask -TaskPath $Script:TaskPath -TaskName $Script:TaskName -ErrorAction Stop
        Write-Log "Started scheduled task '$($Script:TaskPath)$($Script:TaskName)'." -Level SUCCESS
        # Read the task state shortly after to surface whether the engine
        # accepted the trigger (Running) or it has already returned to Ready.
        try {
            $post = Get-ScheduledTask -TaskPath $Script:TaskPath -TaskName $Script:TaskName -ErrorAction SilentlyContinue
            if ($post) { Write-Log "Post-trigger task state: $($post.State)" -Level DEBUG }
        } catch { }
        return $true
    }
    catch {
        Write-Log "Failed to start scheduled task '$($Script:TaskPath)$($Script:TaskName)': $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

#endregion

# ─────────────────────────────────────────────────────────────────────────────
#region Main
# ─────────────────────────────────────────────────────────────────────────────

# Capture the script file name once so Write-Log can include it in every
# CMTrace entry. $MyInvocation.MyCommand.Name is reliable at the script scope
# but inside functions it reports the function name instead.
$Script:ScriptName = $MyInvocation.MyCommand.Name
$exitCode = 0

Write-Log '--- Remediation started ---' -Level STEP

# -- Environment context (verbose) -----------------------------------------
# Labels are padded to a fixed width so the section reads as an aligned table.
Write-Log ("{0,-14} : {1}" -f 'Script',         $Script:ScriptName) -Level DEBUG
Write-Log ("{0,-14} : {1} (Edition={2})" -f 'PSVersion',  $PSVersionTable.PSVersion, $PSVersionTable.PSEdition) -Level DEBUG
Write-Log ("{0,-14} : {1}-bit" -f 'PSArch',     ([IntPtr]::Size * 8)) -Level DEBUG
Write-Log ("{0,-14} : {1} ({2})" -f 'OS',       [Environment]::OSVersion.VersionString, [Environment]::OSVersion.Platform) -Level DEBUG
Write-Log ("{0,-14} : {1}" -f 'Hostname',       $env:COMPUTERNAME) -Level DEBUG
Write-Log ("{0,-14} : {1}\{2} (PID={3})" -f 'User', $env:USERDOMAIN, $env:USERNAME, $PID) -Level DEBUG
Write-Log ("{0,-14} : {1}" -f 'WorkingDir',     (Get-Location).Path) -Level DEBUG
Write-Log ("{0,-14} : {1} (max={2}KB)" -f 'LogFile', $Script:LogFile, [math]::Round($Script:LogMaxSize/1KB)) -Level DEBUG
Write-Log ("{0,-14} : {1}" -f 'RegPathRoot',    $Script:RegPathRoot) -Level DEBUG
Write-Log ("{0,-14} : 0x{1:X}" -f 'ArmValue',   $Script:ArmValue) -Level DEBUG
Write-Log ("{0,-14} : 0x{1:X}" -f 'CompliantValue', $Script:CompliantValue) -Level DEBUG
Write-Log ("{0,-14} : {1}{2}" -f 'TaskPath/Name', $Script:TaskPath, $Script:TaskName) -Level DEBUG

# -- Prerequisite: 64-bit PowerShell ---------------------------------------
if (-not (Test-Is64BitPS)) {
    Write-Log 'Script is not running in 64-bit PowerShell.' -Level ERROR
    Write-Output 'PREREQ: Not running in 64-bit PowerShell.'
    exit 1
}

# -- Secure Boot must be enabled -------------------------------------------
# There is no point arming the update on a Legacy/BIOS system or on a device
# where the user disabled Secure Boot in firmware - the servicing component
# will not advance the bitmask. Abort cleanly so Intune does not loop.
Write-Log 'Calling Confirm-SecureBootUEFI ...' -Level DEBUG
try {
    $sbEnabled = Confirm-SecureBootUEFI -ErrorAction Stop
    Write-Log "Confirm-SecureBootUEFI returned: $sbEnabled" -Level DEBUG
}
catch {
    Write-Log "Confirm-SecureBootUEFI failed: $($_.Exception.Message)." -Level ERROR
    Write-Output 'ABORT: Secure Boot not supported.'
    exit 1
}
if (-not $sbEnabled) {
    Write-Log 'Secure Boot is disabled in firmware. Cannot remediate from OS.' -Level ERROR
    Write-Output 'ABORT: Secure Boot disabled.'
    exit 1
}

# -- Gather extended context (registry + events + task + throttle) --------
# Collected up front so short-circuits, branch logic, and post-trigger
# logging all reason about the same snapshot.
Write-Log 'Gathering extended Secure Boot context ...' -Level STEP
$avUpdates    = Get-RegValue -Path $Script:RegPathRoot -Name 'AvailableUpdates'
$avPolicy     = Get-RegValue -Path $Script:RegPathRoot -Name 'AvailableUpdatesPolicy'
$optOut       = Get-RegValue -Path $Script:RegPathRoot -Name 'HighConfidenceOptOut'
$svcStatus    = Get-RegValue -Path $Script:RegPathServicing -Name 'UEFICA2023Status'
$svcError     = Get-RegValue -Path $Script:RegPathServicing -Name 'UEFICA2023Error'
$svcErrEvent  = Get-RegValue -Path $Script:RegPathServicing -Name 'UEFICA2023ErrorEvent'
$canAttemptAt = Get-CanAttemptUpdateAfter
$taskState    = Get-SecureBootTaskState
$evt          = Get-SecureBootEventFlags

$avHex       = if ($null -ne $avUpdates) { '0x{0:X}' -f $avUpdates } else { '<missing>' }
$avPolicyHex = if ($null -ne $avPolicy)  { '0x{0:X}' -f $avPolicy }  else { '<not set>' }
Write-Log ("{0,-22} : {1}" -f 'AvailableUpdates',         $avHex) -Level INFO
Write-Log ("{0,-22} : {1}" -f 'AvailableUpdatesPolicy',   $avPolicyHex) -Level INFO
Write-Log ("{0,-22} : Status='{1}' Error='{2}' ErrEvent='{3}'" -f 'Servicing',
    $svcStatus, $svcError, $svcErrEvent) -Level INFO
Write-Log ("{0,-22} : {1} (Enabled={2})" -f 'Secure-Boot-Update task', $taskState.State, $taskState.Enabled) -Level INFO
Write-Log ("{0,-22} : 1808={1} 1803={2} 1802={3} 1800={4}" -f 'Event counts',
    $evt.Counts[1808], $evt.Counts[1803], $evt.Counts[1802], $evt.Counts[1800]) -Level INFO

# -- Hard-blocker reporting markers ----------------------------------------
# These conditions cannot be remediated from the OS. We deliberately exit 1
# (not 0) so Intune's PR dashboard does NOT mark the device as 'Remediation
# successful' / 'Fixed' - the device IS still non-compliant; we just chose
# not to make a system change. Side effects are still skipped (no registry
# writes, no scheduled-task triggers) so the PR cycle is essentially free.
# The disambiguating STDOUT prefix lets dashboards segment cohorts.
Write-Log 'Evaluating hard-blocker reporting markers ...' -Level STEP

if ($null -ne $optOut -and [int]$optOut -eq 1) {
    Write-Log 'HighConfidenceOptOut = 1 -> device intentionally excluded from CA 2023 rollout. Skipping remediation; reporting as non-compliant (not actionable).' -Level WARN
    Write-Output 'NON-COMPLIANT-NOT-ACTIONABLE: HighConfidenceOptOut=1 (admin-managed exclusion).'
    exit 1
}

if ($evt.Has1803) {
    Write-Log 'Event 1803 present -> matching KEK update not found. OEM must supply a PK-signed KEK; remediation cannot proceed.' -Level WARN
    Write-Output 'NON-COMPLIANT-NOT-ACTIONABLE: Event 1803 (missing KEK - OEM responsibility).'
    exit 1
}

if ($evt.Has1802) {
    $kiSuffix = if ($evt.SkipReason) { " ($($evt.SkipReason))" } else { '' }
    Write-Log "Event 1802 present -> known firmware issue is blocking the update$kiSuffix. OEM firmware update required." -Level WARN
    Write-Output ("NON-COMPLIANT-NOT-ACTIONABLE: Event 1802 (known firmware issue{0})." -f $kiSuffix)
    exit 1
}

# Reboot-pending: 1800 fired but 1808 has not yet -> the bitmask will
# advance after the next boot. Re-triggering the task does nothing useful
# while we wait for the user to reboot. Skip side effects but exit 1 so
# Intune does not falsely mark the device as 'Remediation successful'.
if ($evt.Has1800 -and -not $evt.Has1808) {
    Write-Log 'Event 1800 (reboot pending) is present and 1808 has not fired yet. Skipping task trigger; reporting as non-compliant (pending reboot).' -Level WARN
    Write-Output 'NON-COMPLIANT-PENDING-REBOOT: Update staged, awaiting reboot.'
    exit 1
}

# -- Operational warnings (don't change the decision) ----------------------
if ($null -ne $canAttemptAt) {
    $nowUtc = (Get-Date).ToUniversalTime()
    if ($canAttemptAt -gt $nowUtc) {
        $waitMin = [math]::Round(($canAttemptAt - $nowUtc).TotalMinutes, 1)
        Write-Log ("Operational warning: CanAttemptUpdateAfter is {0} minute(s) in the future ({1}Z, local {2}). Firmware throttle active; trigger may be a no-op." -f $waitMin, $canAttemptAt.ToString('s'), $canAttemptAt.ToLocalTime().ToString('s')) -Level WARN
    }
}

if ($null -ne $avPolicy -and $null -ne $avUpdates -and $avPolicy -ne $avUpdates) {
    Write-Log ("Operational warning: AvailableUpdatesPolicy ({0}) differs from AvailableUpdates ({1}). GPO/MDM may revert direct AvailableUpdates writes on the next policy refresh." -f $avPolicyHex, $avHex) -Level WARN
}

# -- Pre-flight: scheduled task must exist and be enabled ------------------
# Writing AvailableUpdates only to discover the task is disabled would leave
# the device armed but stuck. Abort early with a clear message instead.
if (-not $taskState.Exists) {
    Write-Log "Secure-Boot-Update scheduled task is missing. Cannot remediate." -Level ERROR
    Write-Output "FAIL: Scheduled task '$($Script:TaskPath)$($Script:TaskName)' is missing."
    exit 1
}
if (-not $taskState.Enabled) {
    Write-Log "Secure-Boot-Update scheduled task is in state '$($taskState.State)' (not Ready/Running). Cannot trigger." -Level ERROR
    Write-Output "FAIL: Scheduled task is '$($taskState.State)' (not Ready). Re-enable before remediation."
    exit 1
}

# -- Branch logic ----------------------------------------------------------
# C: Already complete -> no-op fast path.
if ($null -ne $avUpdates -and $avUpdates -eq $Script:CompliantValue) {
    Write-Log ("AvailableUpdates already at terminal complete state (0x{0:X}). No action taken." -f $Script:CompliantValue) -Level SUCCESS
    Write-Output 'NO-OP: Already at compliant terminal state (0x4000).'
    exit 0
}

# A: Initial arm - value missing or 0 means the rollout has never started.
#    Write the well-known kick-off bitmask, then trigger the task below.
if ($null -eq $avUpdates -or $avUpdates -eq 0) {
    Write-Log 'Branch A: Initial arm (AvailableUpdates missing or 0).' -Level STEP
    try {
        Set-RegDword -Path $Script:RegPathRoot -Name 'AvailableUpdates' -Value $Script:ArmValue
        Write-Log ("Set AvailableUpdates = 0x{0:X} (arm value)." -f $Script:ArmValue) -Level SUCCESS
    }
    catch {
        Write-Log "Registry write failed: $($_.Exception.Message)" -Level ERROR
        Write-Output "FAIL: Registry write error - $($_.Exception.Message)"
        exit 1
    }
}
else {
    # B: Resume - a non-zero, non-complete value means the state machine is
    #    already partway through (e.g. 0x4100 staged for reboot, 0x4104 KEK
    #    pending). DO NOT overwrite the bitmask; just nudge the task so it
    #    advances to the next step on the current boot or after reboot.
    Write-Log "Branch B: Resume (AvailableUpdates = $avHex). Triggering task without registry change." -Level STEP
}

# -- Trigger scheduled task ------------------------------------------------
if (-not (Start-SecureBootUpdateTask)) {
    Write-Output "FAIL: Could not start Secure-Boot-Update scheduled task."
    exit 1
}

# -- Read back new value (best effort) -------------------------------------
# The bitmask may not change immediately - most transitions only happen
# after the next reboot - but we log the post value so admins can correlate
# with detection runs.
$post = Get-RegValue -Path $Script:RegPathRoot -Name 'AvailableUpdates'
$postHex = if ($null -ne $post) { '0x{0:X}' -f $post } else { '<missing>' }
Write-Log "Post-trigger AvailableUpdates = $postHex (a reboot may be required to advance state)." -Level INFO

Write-Log "--- Remediation finished (ExitCode: $exitCode) ---" -Level STEP
Write-Output "REMEDIATED: AvailableUpdates pre=$avHex post=$postHex. Reboot required to finalize."
exit $exitCode

#endregion
