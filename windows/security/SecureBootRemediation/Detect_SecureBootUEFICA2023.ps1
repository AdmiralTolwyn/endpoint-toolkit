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

    Any failure -> exit 1 (non-compliant) which triggers the paired
    Remediate_SecureBootUEFICA2023.ps1 script via Intune Proactive
    Remediations.

    Logs are written in CMTrace format to:
        %ProgramData%\Microsoft\IntuneManagementExtension\Logs\PR_SecureBootUEFICA2023.log

    Intune Proactive Remediation settings:
      Run this script using the logged-on credentials : No
      Enforce script signature check                  : No
      Run script in 64-bit PowerShell                 : Yes

.OUTPUTS
    System.Int32 exit code:
      0  = Compliant (all checks passed)
      1  = Non-compliant or prerequisite failure

    Single-line STDOUT summary suitable for the Intune PR detection column.

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
$Script:CompliantAvUpdates = 0x4000                          # Terminal "complete" bitmask state

# Event log signal indicating the Secure Boot Update finished successfully.
$Script:EventProvider = 'Microsoft-Windows-TPM-WMI'
$Script:EventID       = 1808

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

function Test-Event1808 {
    <#
    .SYNOPSIS
        Returns $true if at least one Event ID 1808 from
        Microsoft-Windows-TPM-WMI is present in the System log.

    .DESCRIPTION
        Event 1808 is logged by the Secure Boot servicing component when the
        UEFI CA 2023 update completes successfully. Its presence is the
        authoritative signal that the OS-side roll-out finished at least
        once on this device.
    #>
    Write-Log "Test-Event1808 ENTER (LogName='System', Provider='$($Script:EventProvider)', Id=$($Script:EventID))" -Level DEBUG
    try {
        $evt = Get-WinEvent -FilterHashtable @{
            LogName      = 'System'
            ProviderName = $Script:EventProvider
            Id           = $Script:EventID
        } -MaxEvents 1 -ErrorAction SilentlyContinue
        if ($null -ne $evt) {
            Write-Log "Test-Event1808 EXIT -> True (TimeCreated=$($evt.TimeCreated.ToString('s')), RecordId=$($evt.RecordId))" -Level DEBUG
            return $true
        }
        Write-Log 'Test-Event1808 EXIT -> False (no matching event)' -Level DEBUG
        return $false
    }
    catch {
        Write-Log "Test-Event1808: exception -> $($_.Exception.Message)" -Level DEBUG
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

# -- Prerequisite: 64-bit PowerShell ---------------------------------------
if (-not (Test-Is64BitPS)) {
    Write-Log 'Script is not running in 64-bit PowerShell.' -Level ERROR
    Write-Output 'PREREQ: Not running in 64-bit PowerShell.'
    exit 1
}
Write-Log 'Prerequisite passed (64-bit PS).' -Level SUCCESS

# -- Check 1: Secure Boot enabled ------------------------------------------
# Confirm-SecureBootUEFI throws on BIOS/Legacy systems where Secure Boot is
# unavailable. Treat both 'unsupported' and 'disabled' as non-compliant —
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

# -- Checks 2 & 3: Servicing registry (Status / Error) ---------------------
Write-Log 'Check 2/5: read UEFICA2023Status ...' -Level DEBUG
$status = Get-RegValue -Path $Script:RegPathServicing -Name 'UEFICA2023Status'
Write-Log 'Check 3/5: read UEFICA2023Error ...' -Level DEBUG
$err    = Get-RegValue -Path $Script:RegPathServicing -Name 'UEFICA2023Error'

if ($null -eq $status) {
    # Servicing key/value missing typically means the Secure Boot Update task
    # has never executed on this device.
    Write-Log 'UEFICA2023Status registry value not found (servicing has not run yet).' -Level WARN
    $outputMsgs += 'UEFICA2023Status missing'
    $exitCode = 1
}
elseif ($status -ne 'Updated') {
    # Common transient values: 'NotStarted', 'InProgress'. Anything other
    # than 'Updated' is non-compliant from the detection script's POV.
    Write-Log "UEFICA2023Status = '$status' (expected 'Updated')." -Level WARN
    $outputMsgs += "Status='$status'"
    $exitCode = 1
}
else {
    Write-Log "UEFICA2023Status = 'Updated'." -Level SUCCESS
}

if ($null -ne $err -and $err -ne 0) {
    # Non-zero error code surfaces firmware/KEK/KI failure during servicing.
    Write-Log "UEFICA2023Error = $err (expected 0)." -Level ERROR
    $outputMsgs += "Error=$err"
    $exitCode = 1
}
elseif ($null -ne $err) {
    Write-Log 'UEFICA2023Error = 0.' -Level SUCCESS
}

# -- Check 4: AvailableUpdates bitmask -------------------------------------
# 0x4000 is the terminal "complete" state. Any other value (missing, 0,
# 0x5944 armed, 0x4100 staged, 0x4104 KEK pending, etc.) means the rollout
# has not yet finished and the remediation script should re-trigger the
# Secure Boot Update scheduled task.
Write-Log 'Check 4/5: read AvailableUpdates bitmask ...' -Level DEBUG
$avUpdates = Get-RegValue -Path $Script:RegPathRoot -Name 'AvailableUpdates'
if ($null -eq $avUpdates) {
    Write-Log ("AvailableUpdates registry value not present (expected 0x{0:X})." -f $Script:CompliantAvUpdates) -Level WARN
    $outputMsgs += 'AvailableUpdates missing'
    $exitCode = 1
}
else {
    $avHex = '0x{0:X}' -f $avUpdates
    if ($avUpdates -eq $Script:CompliantAvUpdates) {
        Write-Log "AvailableUpdates = $avHex (compliant terminal state)." -Level SUCCESS
    }
    else {
        Write-Log ("AvailableUpdates = {0} (expected 0x{1:X})." -f $avHex, $Script:CompliantAvUpdates) -Level WARN
        $outputMsgs += "AvailableUpdates=$avHex"
        $exitCode = 1
    }
}

# -- Check 5: Event 1808 (TPM-WMI completion) ------------------------------
Write-Log 'Check 5/5: query System log for Event 1808 ...' -Level DEBUG
if (Test-Event1808) {
    Write-Log "Event ID $($Script:EventID) from $($Script:EventProvider) found in System log." -Level SUCCESS
}
else {
    Write-Log "Event ID $($Script:EventID) from $($Script:EventProvider) NOT found in System log." -Level WARN
    $outputMsgs += "Event $($Script:EventID) missing"
    $exitCode = 1
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
