<#
.SYNOPSIS
    Detects whether the Windows UEFI CA 2023 Secure Boot certificate update
    has been fully applied on the device.

.DESCRIPTION
    Compliance gate (ALL must be true to be considered compliant):
      1. Secure Boot is enabled in firmware (Confirm-SecureBootUEFI).
      2. HKLM\...\SecureBoot\Servicing\UEFICA2023Status = 'Updated'.
      3. HKLM\...\SecureBoot\Servicing\UEFICA2023Error  = 0.
      4. HKLM\...\SecureBoot\AvailableUpdates           = 0x4000
         (terminal "complete" state of the Secure Boot Update bitmask).
      5. Event ID 1808 from provider Microsoft-Windows-TPM-WMI present
         in the System log (successful CA 2023 update completion).

    Short-circuits (exit 0 = "compliant / not actionable") that suppress
    the gate to prevent an Intune PR retry loop on devices that cannot
    be remediated from the OS:
      * HighConfidenceOptOut = 1 (admin-managed exclusion).
      * Event 1803 present (missing KEK - OEM must supply PK-signed KEK).
      * Event 1802 present (known firmware issue / SkipReason: KI_<n>).
      * Event 1800 present without 1808 (reboot pending - update will
        proceed after the next boot; remediation is a no-op).

    Operational warnings (logged but do NOT change the gate decision):
      * Secure-Boot-Update scheduled task missing or disabled - any
        remediation that triggers it will silently fail.
      * CanAttemptUpdateAfter FILETIME in the future - firmware throttle.
      * AvailableUpdatesPolicy disagrees with AvailableUpdates - GPO/MDM
        is overriding direct AvailableUpdates writes.
      * Latest 1795/1796 error code captured for triage.

    Diagnostic context surfaced in every run:
      * UEFICA2023ErrorEvent (last error event id from the registry).
      * BucketId / BucketConfidenceLevel / SkipReason parsed from the
        most recent 1801/1808/1802 event message.

    Any compliance failure that is NOT short-circuited -> exit 1
    (non-compliant) which triggers Remediate_SecureBootUEFICA2023.ps1
    via Intune Proactive Remediations.

    Logs are written in CMTrace format to:
        %ProgramData%\Microsoft\IntuneManagementExtension\Logs\PR_SecureBootUEFICA2023.log

    Intune Proactive Remediation settings:
      Run this script using the logged-on credentials : No
      Enforce script signature check                  : No
      Run script in 64-bit PowerShell                 : Yes

    References:
      KB 5016061 - Secure Boot DBX update event reference
      KB 5072718 - Sample Secure Boot Inventory Data Collection script
      KB 5084567 - Sample Secure Boot E2E Automation Guide

.OUTPUTS
    System.Int32 exit code:
      0  = Compliant, not-applicable, or pending-reboot (no remediation)
      1  = Non-compliant (remediation should run)

    Single-line STDOUT summary suitable for the Intune PR detection column.

.NOTES
    Author:   Anton Romanyuk
    Version:  1.1
    Date:     2026-04-30
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

# Log file is written to the Intune Management Extension log directory so it
# is automatically collected by Intune diagnostics / MDMDiagnosticsTool.
$Script:LogDir     = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs"
$Script:LogName    = 'PR_SecureBootUEFICA2023'
$Script:LogFile    = Join-Path $Script:LogDir "$Script:LogName.log"
$Script:LogMaxSize = 250KB                                   # rotation threshold
$Script:Component  = 'Detect'                                # CMTrace component column

# Registry locations driving the Secure Boot servicing state machine.
$Script:RegPathRoot        = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot'
$Script:RegPathServicing   = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing'
$Script:RegPathDeviceAttr  = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing\DeviceAttributes'
$Script:CompliantAvUpdates = 0x4000                          # Terminal "complete" bitmask state

# Event log signals (provider Microsoft-Windows-TPM-WMI in the System log).
#   1808 - Update completed successfully (gate condition).
#   1803 - Missing KEK update (OEM responsibility - hard blocker).
#   1802 - Known firmware issue, SkipReason: KI_<n> (hard blocker).
#   1800 - Reboot pending (suppress non-compliant; not an error).
#   1801 - Update initiated, reboot required (informational).
#   1795 - Firmware returned error (capture code for triage).
#   1796 - Generic error logged (capture code for triage).
$Script:EventProvider          = 'Microsoft-Windows-TPM-WMI'
$Script:EventID                = 1808
$Script:HardBlockerEvents      = @(1802, 1803)
$Script:RebootPendingEvent     = 1800
$Script:ErrorEvents            = @(1795, 1796)
$Script:AllSecureBootEvents    = @(1795, 1796, 1800, 1801, 1802, 1803, 1808)
$Script:EventQueryMax          = 100                         # cap for History scan

# Scheduled task that performs the actual servicing pass. Detection only
# checks its existence/state; the remediate script triggers it.
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
        processes. We must run 64-bit so the SecureBoot keys are read from
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

function Get-SecureBootEvents {
    <#
    .SYNOPSIS
        Queries all relevant Secure Boot events from the System log in one
        pass and returns a hashtable of parsed signals.

    .DESCRIPTION
        Performs a single Get-WinEvent call (capped at $Script:EventQueryMax)
        for every event ID in $Script:AllSecureBootEvents to keep cost low,
        then partitions the result and parses BucketId / BucketConfidenceLevel
        / SkipReason / error codes from message bodies via regex.

    .OUTPUTS
        Hashtable with the following keys:
          Latest                - newest matching event (or $null)
          ByID                  - hashtable: id -> array of events
          Counts                - hashtable: id -> count
          Has1808 / Has1803 / Has1802 / Has1800
          BucketId, BucketConfidenceLevel, SkipReason   (from latest msg)
          Event1795Code, Event1796Code                  (parsed error codes)
    #>
    Write-Log ("Get-SecureBootEvents ENTER (LogName='System', IDs={0}, Max={1})" -f ($Script:AllSecureBootEvents -join ','), $Script:EventQueryMax) -Level DEBUG
    $result = @{
        Latest                = $null
        ByID                  = @{}
        Counts                = @{}
        Has1808               = $false
        Has1803               = $false
        Has1802               = $false
        Has1800               = $false
        BucketId              = $null
        BucketConfidenceLevel = $null
        SkipReason            = $null
        Event1795Code         = $null
        Event1796Code         = $null
    }
    try {
        $events = @(Get-WinEvent -FilterHashtable @{
            LogName      = 'System'
            ProviderName = $Script:EventProvider
            Id           = $Script:AllSecureBootEvents
        } -MaxEvents $Script:EventQueryMax -ErrorAction SilentlyContinue)
    }
    catch {
        Write-Log "Get-SecureBootEvents: query exception -> $($_.Exception.Message)" -Level DEBUG
        return $result
    }
    Write-Log "Get-SecureBootEvents: retrieved $($events.Count) raw event(s) from provider $($Script:EventProvider)" -Level DEBUG

    # Partition by ID (note: $events is already newest-first from Get-WinEvent).
    foreach ($id in $Script:AllSecureBootEvents) {
        $bucket = @($events | Where-Object { $_.Id -eq $id })
        $result.ByID[$id]   = $bucket
        $result.Counts[$id] = $bucket.Count
    }
    if ($events.Count -gt 0) { $result.Latest = $events[0] }
    $result.Has1808 = ($result.Counts[1808] -gt 0)
    $result.Has1803 = ($result.Counts[1803] -gt 0)
    $result.Has1802 = ($result.Counts[1802] -gt 0)
    $result.Has1800 = ($result.Counts[1800] -gt 0)

    # Parse BucketId / Confidence / SkipReason from the latest 1801/1808/1802
    # event message - these fields appear in the structured EventData of the
    # Secure Boot servicing component.
    $msgEvent = @($events | Where-Object { $_.Id -in @(1801,1808,1802) } | Select-Object -First 1)
    if ($msgEvent.Count -gt 0 -and $null -ne $msgEvent[0].Message) {
        $m = $msgEvent[0].Message
        if ($m -match 'BucketId:\s*([^\r\n]+)')              { $result.BucketId              = $matches[1].Trim() }
        if ($m -match 'BucketConfidenceLevel:\s*([^\r\n]+)') { $result.BucketConfidenceLevel = $matches[1].Trim() }
        if ($m -match 'SkipReason:\s*(KI_\d+)')              { $result.SkipReason            = $matches[1] }
    }

    # Capture error codes from the latest 1795 / 1796 events for triage.
    # The KB sample uses (?:error|code|status)[:\s]*(?:0x)?([0-9A-Fa-f]+) -
    # match the same shape so the captured value matches what the inventory
    # script would record.
    foreach ($pair in @(@{Id=1795; Key='Event1795Code'}, @{Id=1796; Key='Event1796Code'})) {
        $latest = @($result.ByID[$pair.Id] | Select-Object -First 1)
        if ($latest.Count -gt 0 -and $latest[0].Message -match '(?:error|code|status)[:\s]*(?:0x)?([0-9A-Fa-f]{1,8})') {
            $result[$pair.Key] = $matches[1]
        }
    }

    Write-Log ("Get-SecureBootEvents EXIT (1808={0}, 1803={1}, 1802={2}, 1800={3}, 1801={4}, 1795={5}, 1796={6})" -f
        $result.Counts[1808], $result.Counts[1803], $result.Counts[1802], $result.Counts[1800],
        $result.Counts[1801], $result.Counts[1795], $result.Counts[1796]) -Level DEBUG
    return $result
}

function Get-SecureBootTaskState {
    <#
    .SYNOPSIS
        Returns information about the Secure-Boot-Update scheduled task.

    .DESCRIPTION
        The remediate script triggers \Microsoft\Windows\PI\Secure-Boot-Update
        to advance the AvailableUpdates bitmask. If the task is missing or
        disabled, the remediation will appear to run successfully but never
        change the device state - so detection must surface that condition.

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

    .OUTPUTS
        System.DateTime (UTC) or $null.
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

#endregion

# ─────────────────────────────────────────────────────────────────────────────
#region Main
# ─────────────────────────────────────────────────────────────────────────────

# Capture the script file name once so Write-Log can include it in every
# CMTrace entry. $MyInvocation.MyCommand.Name is reliable at the script scope
# but inside functions it reports the function name instead.
$Script:ScriptName = $MyInvocation.MyCommand.Name

# Aggregate non-compliance reasons across all 5 checks so the final STDOUT
# summary lists every failing condition in one line for the Intune UI.
$exitCode   = 0
$outputMsgs = @()

Write-Log '--- Detection started ---' -Level STEP

# -- Environment context (verbose) -----------------------------------------
# Dump the runtime context so log correlation across devices is trivial.
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
Write-Log ("{0,-14} : {1}" -f 'RegPathSvc',     $Script:RegPathServicing) -Level DEBUG
Write-Log ("{0,-14} : 0x{1:X}" -f 'CompliantValue', $Script:CompliantAvUpdates) -Level DEBUG
Write-Log ("{0,-14} : {1} (Id={2})" -f 'EventProvider', $Script:EventProvider, $Script:EventID) -Level DEBUG
Write-Log ("{0,-14} : {1}{2}" -f 'TaskPath/Name', $Script:TaskPath, $Script:TaskName) -Level DEBUG

# -- Prerequisite: 64-bit PowerShell ---------------------------------------
if (-not (Test-Is64BitPS)) {
    Write-Log 'Script is not running in 64-bit PowerShell.' -Level ERROR
    Write-Output 'PREREQ: Not running in 64-bit PowerShell.'
    exit 1
}
Write-Log 'Prerequisite passed (64-bit PS).' -Level SUCCESS

# -- Check 1: Secure Boot enabled ------------------------------------------
# Confirm-SecureBootUEFI throws on BIOS/Legacy systems where Secure Boot is
# unavailable. Treat both 'unsupported' and 'disabled' as non-compliant -
# the OS cannot apply the new CA without Secure Boot active.
Write-Log 'Check 1/5: Confirm-SecureBootUEFI ...' -Level DEBUG
try {
    $sbEnabled = Confirm-SecureBootUEFI -ErrorAction Stop
    Write-Log "Confirm-SecureBootUEFI returned: $sbEnabled" -Level DEBUG
}
catch {
    Write-Log "Confirm-SecureBootUEFI failed: $($_.Exception.Message). Device likely BIOS/Legacy or Secure Boot unavailable." -Level ERROR
    Write-Output 'NON-COMPLIANT: Secure Boot not supported.'
    exit 1
}
if (-not $sbEnabled) {
    Write-Log 'Secure Boot is disabled in firmware.' -Level ERROR
    Write-Output 'NON-COMPLIANT: Secure Boot disabled.'
    exit 1
}
Write-Log 'Secure Boot is ENABLED.' -Level SUCCESS

# -- Gather extended context (registry + events + task + throttle) --------
# All data is pulled up front in a single pass so subsequent checks and the
# final summary can reason about the same snapshot.
Write-Log 'Gathering extended Secure Boot context (registry / events / task / throttle) ...' -Level STEP

# Servicing-side telemetry (Status, Error, last error event id).
$status         = Get-RegValue -Path $Script:RegPathServicing -Name 'UEFICA2023Status'
$err            = Get-RegValue -Path $Script:RegPathServicing -Name 'UEFICA2023Error'
$errEvent       = Get-RegValue -Path $Script:RegPathServicing -Name 'UEFICA2023ErrorEvent'

# Servicing root: bitmask + GPO/MDM persistent value + admin opt-out.
$avUpdates      = Get-RegValue -Path $Script:RegPathRoot -Name 'AvailableUpdates'
$avPolicy       = Get-RegValue -Path $Script:RegPathRoot -Name 'AvailableUpdatesPolicy'
$optOut         = Get-RegValue -Path $Script:RegPathRoot -Name 'HighConfidenceOptOut'
$mgmtOptIn      = Get-RegValue -Path $Script:RegPathRoot -Name 'MicrosoftUpdateManagedOptIn'

# Firmware throttle: when set in the future, triggering the task is a no-op.
$canAttemptAt   = Get-CanAttemptUpdateAfter

# Scheduled task state - if missing/disabled, remediation is ineffective.
$taskState      = Get-SecureBootTaskState

# Event log signals (single query, partitioned by ID).
$evt            = Get-SecureBootEvents

# Render snapshot for the log so a single run captures the full picture.
$avHex          = if ($null -ne $avUpdates) { '0x{0:X}' -f $avUpdates } else { '<missing>' }
$avPolicyHex    = if ($null -ne $avPolicy)  { '0x{0:X}' -f $avPolicy }  else { '<not set>' }
$canAttemptStr  = if ($null -ne $canAttemptAt) { '{0}Z (local {1})' -f $canAttemptAt.ToString('s'), $canAttemptAt.ToLocalTime().ToString('s') } else { '<not set>' }
Write-Log ("{0,-22} : {1}" -f 'UEFICA2023Status',           ($status         | ForEach-Object { if ($_ -ne $null) { "'$_'" } else { '<null>' } })) -Level INFO
Write-Log ("{0,-22} : {1}" -f 'UEFICA2023Error',            ($err            | ForEach-Object { if ($_ -ne $null) { $_ }      else { '<null>' } })) -Level INFO
Write-Log ("{0,-22} : {1}" -f 'UEFICA2023ErrorEvent',       ($errEvent       | ForEach-Object { if ($_ -ne $null) { $_ }      else { '<null>' } })) -Level INFO
Write-Log ("{0,-22} : {1}" -f 'AvailableUpdates',           $avHex) -Level INFO
Write-Log ("{0,-22} : {1}" -f 'AvailableUpdatesPolicy',     $avPolicyHex) -Level INFO
Write-Log ("{0,-22} : {1}" -f 'HighConfidenceOptOut',       ($optOut         | ForEach-Object { if ($_ -ne $null) { $_ }      else { '<not set>' } })) -Level INFO
Write-Log ("{0,-22} : {1}" -f 'MicrosoftUpdateMgdOptIn',    ($mgmtOptIn      | ForEach-Object { if ($_ -ne $null) { $_ }      else { '<not set>' } })) -Level INFO
Write-Log ("{0,-22} : {1}" -f 'CanAttemptUpdateAfter',      $canAttemptStr) -Level INFO
Write-Log ("{0,-22} : {1} (Enabled={2})" -f 'Secure-Boot-Update task', $taskState.State, $taskState.Enabled) -Level INFO
Write-Log ("{0,-22} : 1808={1} 1803={2} 1802={3} 1801={4} 1800={5} 1795={6} 1796={7}" -f 'Event counts',
    $evt.Counts[1808], $evt.Counts[1803], $evt.Counts[1802], $evt.Counts[1801], $evt.Counts[1800], $evt.Counts[1795], $evt.Counts[1796]) -Level INFO
if ($evt.BucketId)              { Write-Log ("{0,-22} : {1}" -f 'BucketId',              $evt.BucketId)              -Level INFO }
if ($evt.BucketConfidenceLevel) { Write-Log ("{0,-22} : {1}" -f 'BucketConfidenceLevel', $evt.BucketConfidenceLevel) -Level INFO }
if ($evt.SkipReason)            { Write-Log ("{0,-22} : {1}" -f 'SkipReason',            $evt.SkipReason)            -Level INFO }
if ($evt.Event1795Code)         { Write-Log ("{0,-22} : {1}" -f 'Event 1795 ErrorCode',  $evt.Event1795Code)         -Level WARN }
if ($evt.Event1796Code)         { Write-Log ("{0,-22} : {1}" -f 'Event 1796 ErrorCode',  $evt.Event1796Code)         -Level WARN }

# -- Hard-blocker short-circuits -------------------------------------------
# These are conditions where the device cannot be remediated from the OS.
# Returning exit 0 prevents Intune from looping the PR forever. The reason
# is logged and surfaced in the STDOUT summary so the device shows up in
# reporting as "compliant for our purposes" without a failure noise spike.
Write-Log 'Evaluating hard-blocker short-circuits (HighConfidenceOptOut / 1803 / 1802) ...' -Level STEP

if ($null -ne $optOut -and [int]$optOut -eq 1) {
    Write-Log 'HighConfidenceOptOut = 1 -> device intentionally excluded from CA 2023 rollout. Treating as not-applicable.' -Level WARN
    Write-Output 'NOT-APPLICABLE: HighConfidenceOptOut=1 (admin-managed exclusion).'
    exit 0
}

if ($evt.Has1803) {
    Write-Log "Event 1803 present -> matching KEK update not found. OEM must supply a PK-signed KEK; this cannot be remediated from the OS." -Level WARN
    Write-Output 'NOT-APPLICABLE: Event 1803 (missing KEK - OEM responsibility).'
    exit 0
}

if ($evt.Has1802) {
    $kiSuffix = if ($evt.SkipReason) { " ($($evt.SkipReason))" } else { '' }
    Write-Log "Event 1802 present -> known firmware issue is blocking the update$kiSuffix. OEM firmware update required; cannot be remediated from the OS." -Level WARN
    Write-Output ("NOT-APPLICABLE: Event 1802 (known firmware issue{0})." -f $kiSuffix)
    exit 0
}

# -- Operational warnings (do NOT change the gate decision) ----------------
# Surface conditions that will make a future remediation pass ineffective so
# operators know to investigate before the bitmask state is interpreted.
if (-not $taskState.Exists) {
    Write-Log "Operational warning: Secure-Boot-Update scheduled task is missing. Remediation cannot trigger it." -Level WARN
}
elseif (-not $taskState.Enabled) {
    Write-Log "Operational warning: Secure-Boot-Update scheduled task is in state '$($taskState.State)' (not Ready/Running). Remediation may not advance the bitmask." -Level WARN
}

if ($null -ne $canAttemptAt) {
    $nowUtc = (Get-Date).ToUniversalTime()
    if ($canAttemptAt -gt $nowUtc) {
        $waitMin = [math]::Round(($canAttemptAt - $nowUtc).TotalMinutes, 1)
        Write-Log "Operational warning: CanAttemptUpdateAfter is $waitMin minute(s) in the future ($canAttemptStr). Firmware throttle active; triggering the task before this time is a no-op." -Level WARN
    } else {
        Write-Log "CanAttemptUpdateAfter is in the past; throttle window has elapsed." -Level DEBUG
    }
}

if ($null -ne $avPolicy -and $null -ne $avUpdates -and $avPolicy -ne $avUpdates) {
    Write-Log ("Operational warning: AvailableUpdatesPolicy ({0}) differs from AvailableUpdates ({1}). GPO/MDM is overriding direct writes; remediation that writes AvailableUpdates may be reverted." -f $avPolicyHex, $avHex) -Level WARN
}
elseif ($null -ne $avPolicy) {
    Write-Log "AvailableUpdatesPolicy is set ($avPolicyHex) and matches AvailableUpdates - GPO/MDM-driven deployment detected." -Level DEBUG
}

# -- Checks 2 & 3: Servicing registry (Status / Error) ---------------------
if ($null -eq $status) {
    Write-Log 'Check 2/5: UEFICA2023Status registry value not found (servicing has not run yet).' -Level WARN
    $outputMsgs += 'UEFICA2023Status missing'
    $exitCode = 1
}
elseif ($status -ne 'Updated') {
    # Common transient values: 'NotStarted', 'InProgress'. Anything other
    # than 'Updated' is non-compliant from the detection script's POV.
    Write-Log "Check 2/5: UEFICA2023Status = '$status' (expected 'Updated')." -Level WARN
    $outputMsgs += "Status='$status'"
    $exitCode = 1
}
else {
    Write-Log "Check 2/5: UEFICA2023Status = 'Updated'." -Level SUCCESS
}

if ($null -ne $err -and $err -ne 0) {
    # Non-zero error code surfaces firmware/KEK/KI failure during servicing.
    Write-Log "Check 3/5: UEFICA2023Error = $err (expected 0)." -Level ERROR
    $outputMsgs += "Error=$err"
    $exitCode = 1
}
elseif ($null -ne $err) {
    Write-Log 'Check 3/5: UEFICA2023Error = 0.' -Level SUCCESS
}
else {
    Write-Log 'Check 3/5: UEFICA2023Error not present (treated as 0).' -Level DEBUG
}

# -- Check 4: AvailableUpdates bitmask -------------------------------------
# 0x4000 is the terminal "complete" state. Any other value (missing, 0,
# 0x5944 armed, 0x4100 staged, 0x4104 KEK pending, etc.) means the rollout
# has not yet finished and the remediation script should re-trigger the
# Secure Boot Update scheduled task.
if ($null -eq $avUpdates) {
    Write-Log ("Check 4/5: AvailableUpdates registry value not present (expected 0x{0:X})." -f $Script:CompliantAvUpdates) -Level WARN
    $outputMsgs += 'AvailableUpdates missing'
    $exitCode = 1
}
elseif ($avUpdates -eq $Script:CompliantAvUpdates) {
    Write-Log "Check 4/5: AvailableUpdates = $avHex (compliant terminal state)." -Level SUCCESS
}
else {
    Write-Log ("Check 4/5: AvailableUpdates = {0} (expected 0x{1:X})." -f $avHex, $Script:CompliantAvUpdates) -Level WARN
    $outputMsgs += "AvailableUpdates=$avHex"
    $exitCode = 1
}

# -- Check 5: Event 1808 (TPM-WMI completion) ------------------------------
if ($evt.Has1808) {
    $latest1808 = @($evt.ByID[1808])[0]
    Write-Log "Check 5/5: Event 1808 found (TimeCreated=$($latest1808.TimeCreated.ToString('s')), Count=$($evt.Counts[1808]))." -Level SUCCESS
}
else {
    Write-Log "Check 5/5: Event ID $($Script:EventID) from $($Script:EventProvider) NOT found in System log." -Level WARN
    $outputMsgs += "Event $($Script:EventID) missing"
    $exitCode = 1
}

# -- Reboot-pending suppression --------------------------------------------
# If the gate failed but Event 1800 is present (and 1808 has NOT yet fired),
# the update has been staged and is waiting for the next reboot. Returning
# exit 1 here would cause the remediate script to re-trigger the task on
# every PR cycle while we wait for the user to reboot - noisy and pointless.
# Suppress to exit 0 with a clear PENDING-REBOOT marker; the next detect
# pass after reboot will either confirm compliance or restart remediation.
if ($exitCode -ne 0 -and $evt.Has1800 -and -not $evt.Has1808) {
    Write-Log 'Gate failed but Event 1800 (reboot pending) is present. Suppressing non-compliant signal until the device reboots.' -Level WARN
    Write-Output ('PENDING-REBOOT: Update staged, awaiting reboot. Gate would fail on: ' + ($outputMsgs -join ' | '))
    Write-Log "--- Detection finished (ExitCode: 0 - reboot-pending suppression) ---" -Level STEP
    exit 0
}

# -- Final summary + exit --------------------------------------------------
Write-Log "--- Detection finished (ExitCode: $exitCode) ---" -Level STEP
if ($exitCode -eq 0) {
    Write-Output 'COMPLIANT: Secure Boot UEFI CA 2023 update fully applied.'
} else {
    # Single-line aggregate summary that fits in the Intune PR detection column.
    Write-Output ('NON-COMPLIANT: ' + ($outputMsgs -join ' | '))
}
exit $exitCode

#endregion
