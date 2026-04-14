<#
.SYNOPSIS
    Remediates BitLocker key escrow to Entra ID and optionally removes the MBAM agent.

.DESCRIPTION
    Unified remediation script for the MBAM-to-Entra ID BitLocker key migration.

    Phase 1 -- Enable protection (if disabled)
      For drives where protection is OFF, add TPM + RecoveryPassword protectors
      if missing, then turn on protection via manage-bde.

    Phase 2 -- Backup recovery keys to Entra ID
      For each encrypted drive, call BackupToAAD-BitLockerKeyProtector for every
      RecoveryPassword protector. Wait up to 30 minutes for Event 845 confirmation.
      On success, set a registry marker per drive.

    Phase 3 -- MBAM cleanup (conditional)
      If ALL drives are confirmed backed up AND the OS drive is marked, uninstall
      the MBAM agent via msiexec. If uninstall fails, disable the service as fallback.

    Intune Proactive Remediation settings:
      Run this script using the logged-on credentials : No
      Enforce script signature check                  : No
      Run script in 64-bit PowerShell                 : Yes

.NOTES
    Author:  Anton Romanyuk
    Version: 1.0
    Date:    2026-04-14
    Context: MBAM-to-Entra ID BitLocker key migration

    DISCLAIMER:
    THIS SCRIPT IS PROVIDED "AS-IS" WITHOUT WARRANTY OF ANY KIND.
    USE AT YOUR OWN RISK.
#>

# ─────────────────────────────────────────────────────────────────────────────
# Configuration
# ─────────────────────────────────────────────────────────────────────────────

$Script:LogDir      = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs"
$Script:LogName     = 'PR_BackupBDE2EntraID'
$Script:LogFile     = Join-Path $Script:LogDir "$Script:LogName.log"
$Script:LogMaxSize  = 250KB
$Script:Component   = 'Remediate'

$Script:ServiceName    = 'MBAMAgent'
$Script:CheckAllDrives = $true
$Script:MbamProductGuid = '{AEC5BCA3-A2C5-46D7-9873-7698E6D3CAA4}'

$Script:RegistryKey    = 'HKLM:\SOFTWARE\ZF\BitLocker'
$Script:RegistryName   = 'Drive_{0}_BitLockerBackupToAAD'   # {0} = drive letter without colon
$Script:RegistryValue  = 'True'

$Script:PackageKey     = 'HKLM:\SOFTWARE\ZF\SW-Distribution\Packages\ZF10001850'

$Script:EventProvider  = 'Microsoft-Windows-BitLocker-API'
$Script:EventMessage   = 'volume {0} was backed up successfully to your Azure AD.'
$Script:EventID        = 845

$Script:BackupTimeout  = 1800   # 30 minutes in seconds
$Script:BackupPoll     = 30     # poll every 30 seconds

# ─────────────────────────────────────────────────────────────────────────────
#region Logging
# ─────────────────────────────────────────────────────────────────────────────

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS', 'STEP')]
        [string]$Level = 'INFO'
    )

    $ts = Get-Date -Format 'HH:mm:ss.fffffff'
    $dt = Get-Date -Format 'MM-dd-yyyy'
    $severity = switch ($Level) { 'ERROR' { 3 } 'WARN' { 2 } default { 1 } }
    $caller = if ($MyInvocation.MyCommand.Name) { $MyInvocation.MyCommand.Name } else { $Script:Component }

    $entry = '<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="{4}" type="{5}" thread="{6}" file="{7}">' -f
        $Message, $ts, $dt, $caller, $env:USERNAME, $severity, $PID, $Script:ScriptName

    if (-not (Test-Path $Script:LogDir)) {
        try { New-Item -Path $Script:LogDir -ItemType Directory -Force -ErrorAction Stop | Out-Null }
        catch { Write-Warning "Cannot create log directory: $_"; return }
    }

    try { $entry | Out-File -FilePath $Script:LogFile -Append -NoClobber -Force -Encoding default -ErrorAction Stop }
    catch { Write-Warning "Log write failed: $_" }

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

function Test-IsSystemContext {
    return [Security.Principal.WindowsIdentity]::GetCurrent().IsSystem
}

function Test-Is64BitPS {
    return [Environment]::Is64BitProcess
}

function Invoke-StopService {
    <#
    .SYNOPSIS  Stops and disables a Windows service.
    #>
    param([Parameter(Mandatory)] [string]$Name)

    $svc = Get-Service -Name $Name -ErrorAction SilentlyContinue
    if (-not $svc) {
        Write-Log "Service '$Name' not found -- nothing to stop." -Level INFO
        return $true
    }
    if ($svc.Status -eq 'Stopped') {
        Write-Log "Service '$Name' is already stopped." -Level INFO
    }
    else {
        try {
            Stop-Service -Name $Name -Force -ErrorAction Stop
            Write-Log "Service '$Name' stopped." -Level SUCCESS
        }
        catch {
            Write-Log "Failed to stop service '$Name': $($_.Exception.Message)" -Level ERROR
            return $false
        }
    }
    try {
        Set-Service -Name $Name -StartupType Disabled -ErrorAction Stop
        Write-Log "Service '$Name' set to Disabled." -Level INFO
    }
    catch {
        Write-Log "Failed to disable service '$Name': $($_.Exception.Message)" -Level WARN
    }
    return $true
}

function Invoke-Process {
    <#
    .SYNOPSIS  Runs an external process and returns the exit code.
    #>
    param(
        [Parameter(Mandatory)] [string]$FilePath,
        [string]$Arguments = '',
        [int]$TimeoutSeconds = 300
    )

    Write-Log "Launching: $FilePath $Arguments" -Level INFO
    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName  = $FilePath
        $psi.Arguments = $Arguments
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError  = $true
        $psi.CreateNoWindow = $true

        $proc = [System.Diagnostics.Process]::Start($psi)
        $stdout = $proc.StandardOutput.ReadToEnd()
        $stderr = $proc.StandardError.ReadToEnd()

        if (-not $proc.WaitForExit($TimeoutSeconds * 1000)) {
            $proc.Kill()
            Write-Log "Process timed out after ${TimeoutSeconds}s -- killed." -Level ERROR
            return -1
        }

        if ($stdout) { Write-Log "STDOUT: $stdout" -Level INFO }
        if ($stderr) { Write-Log "STDERR: $stderr" -Level WARN }
        Write-Log "Exit code: $($proc.ExitCode)" -Level INFO
        return $proc.ExitCode
    }
    catch {
        Write-Log "Process launch failed: $($_.Exception.Message)" -Level ERROR
        return -1
    }
}

function Test-BackupEvent {
    <#
    .SYNOPSIS  Checks event log for successful BitLocker-to-AAD backup (ID 845) after a given time.
    #>
    param(
        [Parameter(Mandatory)] [string]$DriveLetter,
        [DateTime]$Since = [DateTime]::MinValue
    )

    $msg = $Script:EventMessage -f $DriveLetter
    try {
        $provider = (Get-WinEvent -ListProvider $Script:EventProvider -ErrorAction SilentlyContinue).Name
        if (-not $provider) { return $false }

        $hit = Get-WinEvent -ProviderName $Script:EventProvider -ErrorAction SilentlyContinue |
            Where-Object { $_.TimeCreated -gt $Since -and $_.Message -match [regex]::Escape($msg) -and $_.Id -eq $Script:EventID } |
            Select-Object -First 1

        return ($null -ne $hit)
    }
    catch {
        Write-Log "Event log query failed for ${DriveLetter}: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

function Set-BackupRegistry {
    <#
    .SYNOPSIS  Sets the registry marker for a successful backup.
    #>
    param([Parameter(Mandatory)] [string]$DriveLetter)

    $name = $Script:RegistryName -f ($DriveLetter -replace ':', '')
    try {
        if (-not (Test-Path $Script:RegistryKey)) {
            New-Item -Path $Script:RegistryKey -Force -ErrorAction Stop | Out-Null
        }
        Set-ItemProperty -Path $Script:RegistryKey -Name $name -Value $Script:RegistryValue -Type String -Force -ErrorAction Stop
        Write-Log "Registry marker set: $name = $($Script:RegistryValue)" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log "Failed to set registry marker: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

function Invoke-BitLockerBackupToAAD {
    <#
    .SYNOPSIS  Backs up all RecoveryPassword protectors for a drive, waits for Event 845.
    #>
    param([Parameter(Mandatory)] [string]$DriveLetter)

    $vol = Get-BitLockerVolume -MountPoint $DriveLetter -ErrorAction Stop
    $rpProtectors = $vol.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' }

    if (-not $rpProtectors) {
        Write-Log "Drive $DriveLetter has no RecoveryPassword protectors to back up." -Level WARN
        return $false
    }

    $backupTime = Get-Date
    foreach ($kp in $rpProtectors) {
        Write-Log "Backing up protector $($kp.KeyProtectorId) for drive $DriveLetter ..." -Level INFO
        try {
            BackupToAAD-BitLockerKeyProtector -MountPoint $DriveLetter -KeyProtectorId $kp.KeyProtectorId -ErrorAction Stop
            Write-Log "Backup API call succeeded for protector $($kp.KeyProtectorId)." -Level SUCCESS
        }
        catch {
            Write-Log "Backup API call failed for protector $($kp.KeyProtectorId): $($_.Exception.Message)" -Level ERROR
            return $false
        }
    }

    # Poll for Event 845 confirmation
    Write-Log "Waiting up to $($Script:BackupTimeout)s for Event 845 on drive $DriveLetter ..." -Level STEP
    $deadline = $backupTime.AddSeconds($Script:BackupTimeout)
    $attempt  = 0
    while ((Get-Date) -lt $deadline) {
        $attempt++
        Start-Sleep -Seconds $Script:BackupPoll
        if (Test-BackupEvent -DriveLetter $DriveLetter -Since $backupTime) {
            Write-Log "Event 845 confirmed for drive $DriveLetter after attempt $attempt." -Level SUCCESS
            return $true
        }
        Write-Log "Event 845 not yet found for $DriveLetter (attempt $attempt) ..." -Level INFO
    }

    Write-Log "Timed out waiting for Event 845 on drive $DriveLetter." -Level ERROR
    return $false
}

#endregion

# ─────────────────────────────────────────────────────────────────────────────
#region Main
# ─────────────────────────────────────────────────────────────────────────────

$Script:ScriptName = $MyInvocation.MyCommand.Name
$exitCode   = 0
$outputMsgs = @()
$allDrivesOK = $true
$osDriveOK   = $false

Write-Log '=== Remediation started ===' -Level STEP

# -- Prerequisites --
if (-not (Test-IsSystemContext)) {
    $msg = 'Script is not running in system context.'
    Write-Log $msg -Level ERROR
    Write-Output "PREREQ: $msg"
    exit 1
}
if (-not (Test-Is64BitPS)) {
    $msg = 'Script is not running in 64-bit PowerShell.'
    Write-Log $msg -Level ERROR
    Write-Output "PREREQ: $msg"
    exit 1
}
Write-Log 'Prerequisites passed (system context, 64-bit PS).' -Level SUCCESS

# -- Enumerate drives --
if ($Script:CheckAllDrives) {
    $drives = @((Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 }).DeviceID)
} else {
    $drives = @($env:SystemDrive)
}
Write-Log "Drives to remediate: $($drives -join ', ')" -Level INFO

# ─────────────────────────────────────────────────────────────────────────
# Phase 1 -- Enable protection on drives where it is OFF
# ─────────────────────────────────────────────────────────────────────────

Write-Log '--- Phase 1: Enable protection ---' -Level STEP

foreach ($drive in $drives) {
    $vol = Get-BitLockerVolume -MountPoint $drive -ErrorAction SilentlyContinue
    if (-not $vol) {
        Write-Log "Drive $($drive): Cannot query BitLocker volume. Skipped." -Level WARN
        continue
    }

    if ($vol.VolumeStatus -notmatch 'FullyEncrypted') {
        Write-Log "Drive $drive is not FullyEncrypted (Status: $($vol.VolumeStatus)). Skipped." -Level WARN
        continue
    }

    if ($vol.ProtectionStatus -eq 'On') {
        Write-Log "Drive $drive protection is already ON." -Level INFO
        continue
    }

    # Protection is OFF -- attempt to enable
    Write-Log "Drive $drive protection is OFF. Attempting to enable ..." -Level WARN

    # Ensure TPM protector exists on OS drive
    if ($drive -eq $env:SystemDrive) {
        $hasTpm = $vol.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'Tpm' }
        if (-not $hasTpm) {
            Write-Log "Adding TPM protector to $drive ..." -Level INFO
            try {
                Add-BitLockerKeyProtector -MountPoint $drive -TpmProtector -ErrorAction Stop
                Write-Log "TPM protector added to $drive." -Level SUCCESS
            }
            catch {
                Write-Log "Failed to add TPM protector: $($_.Exception.Message)" -Level ERROR
                $allDrivesOK = $false
                continue
            }
        }
    }

    # Ensure at least one RecoveryPassword protector exists
    $hasRP = $vol.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' }
    if (-not $hasRP) {
        Write-Log "Adding RecoveryPassword protector to $drive ..." -Level INFO
        try {
            Add-BitLockerKeyProtector -MountPoint $drive -RecoveryPasswordProtector -ErrorAction Stop
            Write-Log "RecoveryPassword protector added to $drive." -Level SUCCESS
        }
        catch {
            Write-Log "Failed to add RecoveryPassword protector: $($_.Exception.Message)" -Level ERROR
            $allDrivesOK = $false
            continue
        }
    }

    # Turn on protection
    $result = Invoke-Process -FilePath 'manage-bde.exe' -Arguments "-protectors -enable $drive"
    if ($result -eq 0) {
        Write-Log "Protection enabled on drive $drive." -Level SUCCESS
    }
    else {
        Write-Log "manage-bde returned $result for $drive -- protection may not be active." -Level ERROR
        $allDrivesOK = $false
    }
}

# ─────────────────────────────────────────────────────────────────────────
# Phase 2 -- Backup recovery keys to Entra ID
# ─────────────────────────────────────────────────────────────────────────

Write-Log '--- Phase 2: Backup to Entra ID ---' -Level STEP

foreach ($drive in $drives) {
    $vol = Get-BitLockerVolume -MountPoint $drive -ErrorAction SilentlyContinue
    if (-not $vol -or $vol.VolumeStatus -notmatch 'FullyEncrypted') { continue }

    if ($vol.ProtectionStatus -ne 'On') {
        Write-Log "Drive $drive protection still OFF after Phase 1 -- cannot backup. Skipped." -Level ERROR
        $allDrivesOK = $false
        continue
    }

    $backed = Invoke-BitLockerBackupToAAD -DriveLetter $drive
    if ($backed) {
        $null = Set-BackupRegistry -DriveLetter $drive
        $outputMsgs += "Drive $drive backed up to Entra ID."
        if ($drive -eq $env:SystemDrive) { $osDriveOK = $true }
    }
    else {
        Write-Log "Backup failed for drive $drive." -Level ERROR
        $allDrivesOK = $false
        $exitCode = 1
        $outputMsgs += "Drive $drive backup FAILED."
    }
}

# ─────────────────────────────────────────────────────────────────────────
# Phase 3 -- MBAM cleanup (only if all drives OK + OS drive confirmed)
# ─────────────────────────────────────────────────────────────────────────

Write-Log '--- Phase 3: MBAM cleanup ---' -Level STEP

$mbamSvc = Get-Service -Name $Script:ServiceName -ErrorAction SilentlyContinue
if (-not $mbamSvc) {
    Write-Log 'MBAMAgent service not found -- cleanup not needed.' -Level INFO
}
elseif (-not $allDrivesOK -or -not $osDriveOK) {
    Write-Log 'Skipping MBAM cleanup -- not all drives confirmed or OS drive backup failed.' -Level WARN
}
else {
    # Stop and disable the service before uninstall
    $null = Invoke-StopService -Name $Script:ServiceName

    # Attempt MSI uninstall
    Write-Log "Uninstalling MBAM agent ($($Script:MbamProductGuid)) ..." -Level INFO
    $msiResult = Invoke-Process -FilePath 'msiexec.exe' `
        -Arguments "/x $($Script:MbamProductGuid) /qn /norestart REBOOT=ReallySuppress" `
        -TimeoutSeconds 600

    if ($msiResult -eq 0) {
        Write-Log 'MBAM agent uninstalled successfully.' -Level SUCCESS
        $outputMsgs += 'MBAM agent removed.'

        # Set package tracking key
        try {
            if (-not (Test-Path $Script:PackageKey)) {
                New-Item -Path $Script:PackageKey -Force -ErrorAction Stop | Out-Null
            }
            Set-ItemProperty -Path $Script:PackageKey -Name 'Status' -Value 'Uninstalled' -Type String -Force
            Set-ItemProperty -Path $Script:PackageKey -Name 'Date'   -Value (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') -Type String -Force
            Write-Log 'Package tracking registry updated.' -Level INFO
        }
        catch { Write-Log "Package tracking update failed: $($_.Exception.Message)" -Level WARN }
    }
    else {
        Write-Log "MBAM MSI uninstall returned exit code $msiResult." -Level ERROR
        $outputMsgs += 'MBAM uninstall failed -- service disabled as fallback.'
        $exitCode = 1

        # Re-backup after failed uninstall (protector rotation risk)
        Write-Log 'Performing safety re-backup after failed uninstall ...' -Level WARN
        $null = Invoke-BitLockerBackupToAAD -DriveLetter $env:SystemDrive
    }
}

# -- Exit --
Write-Log "=== Remediation finished (ExitCode: $exitCode) ===" -Level STEP
if ($exitCode -eq 0) {
    Write-Output 'SUCCESS: All drives backed up to Entra ID. MBAM cleaned up if present.'
} else {
    Write-Output ($outputMsgs -join ' | ')
}
exit $exitCode

#endregion
