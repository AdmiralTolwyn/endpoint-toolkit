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

    Logs are written in CMTrace format to:
        %ProgramData%\Microsoft\IntuneManagementExtension\Logs\PR_SecureBootUEFICA2023.log

    Intune Proactive Remediation settings:
      Run this script using the logged-on credentials : No
      Enforce script signature check                  : No
      Run script in 64-bit PowerShell                 : Yes

    A reboot is required after the scheduled task progresses past the staged
    state. Pair with a separate reboot policy / user notification mechanism.

.OUTPUTS
    System.Int32 exit code:
      0  = Action taken successfully (or no-op when already compliant)
      1  = Failure (prerequisite, registry write, or scheduled task error)

    Single-line STDOUT summary suitable for the Intune PR remediation column.

.NOTES
    Author:   Anton Romanyuk
    Version:  1.0
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

# Log file is shared with Detect_SecureBootUEFICA2023.ps1 so detection and
# remediation entries appear in the same CMTrace timeline.
$Script:LogDir     = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs"
$Script:LogName    = 'PR_SecureBootUEFICA2023'
$Script:LogFile    = Join-Path $Script:LogDir "$Script:LogName.log"
$Script:LogMaxSize = 250KB                                   # rotation threshold
$Script:Component  = 'Remediate'                             # CMTrace component column

# Registry locations driving the Secure Boot servicing state machine.
$Script:RegPathRoot    = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot'
$Script:ArmValue       = 0x5944                               # Bitmask used to kick off the Secure Boot update
$Script:CompliantValue = 0x4000                               # Terminal "complete" state

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
# where the user disabled Secure Boot in firmware — the servicing component
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

# -- Read current state ----------------------------------------------------
Write-Log 'Reading current AvailableUpdates value ...' -Level DEBUG
$avUpdates = Get-RegValue -Path $Script:RegPathRoot -Name 'AvailableUpdates'
$avHex = if ($null -ne $avUpdates) { '0x{0:X}' -f $avUpdates } else { '<missing>' }
Write-Log "Current AvailableUpdates = $avHex" -Level INFO

# Also read the servicing-side telemetry so the log captures full state on
# every remediation pass, regardless of which branch is taken.
Write-Log 'Reading UEFICA2023Status / UEFICA2023Error for context ...' -Level DEBUG
$svcStatus = Get-RegValue -Path "$Script:RegPathRoot\Servicing" -Name 'UEFICA2023Status'
$svcError  = Get-RegValue -Path "$Script:RegPathRoot\Servicing" -Name 'UEFICA2023Error'
Write-Log "Servicing context: Status='$svcStatus' Error='$svcError'" -Level INFO

# -- Branch logic ----------------------------------------------------------
# C: Already complete -> no-op fast path.
if ($null -ne $avUpdates -and $avUpdates -eq $Script:CompliantValue) {
    Write-Log ("AvailableUpdates already at terminal complete state (0x{0:X}). No action taken." -f $Script:CompliantValue) -Level SUCCESS
    Write-Output 'NO-OP: Already at compliant terminal state (0x4000).'
    exit 0
}

# A: Initial arm — value missing or 0 means the rollout has never started.
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
    # B: Resume — a non-zero, non-complete value means the state machine is
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
# The bitmask may not change immediately — most transitions only happen
# after the next reboot — but we log the post value so admins can correlate
# with detection runs.
$post = Get-RegValue -Path $Script:RegPathRoot -Name 'AvailableUpdates'
$postHex = if ($null -ne $post) { '0x{0:X}' -f $post } else { '<missing>' }
Write-Log "Post-trigger AvailableUpdates = $postHex (a reboot may be required to advance state)." -Level INFO

Write-Log "--- Remediation finished (ExitCode: $exitCode) ---" -Level STEP
Write-Output "REMEDIATED: AvailableUpdates pre=$avHex post=$postHex. Reboot required to finalize."
exit $exitCode

#endregion
