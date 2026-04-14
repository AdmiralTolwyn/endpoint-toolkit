<#
.SYNOPSIS
    Removes the MBAM agent after confirming BitLocker keys are safely escrowed.

.DESCRIPTION
    Standalone MBAM cleanup remediation script.

    Before removing the agent, the script performs a fresh BitLocker key backup
    to Entra ID for the OS drive and waits up to 5 minutes for Event 845
    confirmation. If the backup cannot be confirmed, the script aborts to
    prevent key loss.

    Removal sequence:
      1. Fresh backup of all RecoveryPassword protectors on the OS drive
      2. Wait for Event 845 (up to 5 minutes)
      3. Stop and disable MBAMAgent service
      4. MSI uninstall via msiexec /x
      5. Registry cleanup (package tracking key)
      6. On uninstall failure: disable service as fallback, re-backup keys

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
$Script:LogName     = 'PR_MBAMClientCleanup'
$Script:LogFile     = Join-Path $Script:LogDir "$Script:LogName.log"
$Script:LogMaxSize  = 250KB
$Script:Component   = 'Remediate'

$Script:ServiceName    = 'MBAMAgent'
$Script:MbamProductGuid = '{AEC5BCA3-A2C5-46D7-9873-7698E6D3CAA4}'

$Script:PackageKey     = 'HKLM:\SOFTWARE\ZF\SW-Distribution\Packages\ZF10001850'

$Script:EventProvider  = 'Microsoft-Windows-BitLocker-API'
$Script:EventMessage   = 'volume {0} was backed up successfully to your Azure AD.'
$Script:EventID        = 845

$Script:BackupTimeout  = 300    # 5 minutes
$Script:BackupPoll     = 15     # poll every 15 seconds

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

function Invoke-FreshBitLockerBackup {
    <#
    .SYNOPSIS  Performs a fresh backup of all RecoveryPassword protectors and waits for Event 845.
    .OUTPUTS   $true if Event 845 confirmed within timeout, $false otherwise.
    #>
    param([Parameter(Mandatory)] [string]$DriveLetter)

    $vol = Get-BitLockerVolume -MountPoint $DriveLetter -ErrorAction SilentlyContinue
    if (-not $vol) {
        Write-Log "Cannot query BitLocker volume for $DriveLetter." -Level ERROR
        return $false
    }

    $rpProtectors = $vol.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' }
    if (-not $rpProtectors) {
        Write-Log "No RecoveryPassword protectors found on $DriveLetter." -Level ERROR
        return $false
    }

    $backupTime = Get-Date
    foreach ($kp in $rpProtectors) {
        Write-Log "Backing up protector $($kp.KeyProtectorId) for $DriveLetter ..." -Level INFO
        try {
            BackupToAAD-BitLockerKeyProtector -MountPoint $DriveLetter -KeyProtectorId $kp.KeyProtectorId -ErrorAction Stop
            Write-Log "Backup API call succeeded for protector $($kp.KeyProtectorId)." -Level SUCCESS
        }
        catch {
            Write-Log "Backup API call failed for protector $($kp.KeyProtectorId): $($_.Exception.Message)" -Level ERROR
            return $false
        }
    }

    # Poll for Event 845
    Write-Log "Waiting up to $($Script:BackupTimeout)s for Event 845 on $DriveLetter ..." -Level STEP
    $deadline = $backupTime.AddSeconds($Script:BackupTimeout)
    $attempt  = 0
    while ((Get-Date) -lt $deadline) {
        $attempt++
        Start-Sleep -Seconds $Script:BackupPoll
        if (Test-BackupEvent -DriveLetter $DriveLetter -Since $backupTime) {
            Write-Log "Event 845 confirmed for $DriveLetter after attempt $attempt." -Level SUCCESS
            return $true
        }
        Write-Log "Event 845 not yet found for $DriveLetter (attempt $attempt) ..." -Level INFO
    }

    Write-Log "Timed out waiting for Event 845 on $DriveLetter." -Level ERROR
    return $false
}

#endregion

# ─────────────────────────────────────────────────────────────────────────────
#region Main
# ─────────────────────────────────────────────────────────────────────────────

$Script:ScriptName = $MyInvocation.MyCommand.Name
$exitCode   = 0
$outputMsgs = @()

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

# -- Verify MBAM is present --
$mbamSvc = Get-Service -Name $Script:ServiceName -ErrorAction SilentlyContinue
if (-not $mbamSvc) {
    Write-Log 'MBAMAgent service not found -- nothing to remove.' -Level SUCCESS
    Write-Output 'SUCCESS: MBAM agent is not present.'
    exit 0
}
Write-Log "MBAMAgent service found (Status: $($mbamSvc.Status))." -Level INFO

# -- Step 1: Fresh backup before touching MBAM --
Write-Log '--- Step 1: Fresh BitLocker backup ---' -Level STEP

$osDrive = $env:SystemDrive
$backupConfirmed = Invoke-FreshBitLockerBackup -DriveLetter $osDrive

if (-not $backupConfirmed) {
    $msg = "Fresh backup of $osDrive could not be confirmed -- aborting MBAM removal to prevent key loss."
    Write-Log $msg -Level ERROR
    Write-Output "ABORT: $msg"
    exit 1
}
Write-Log "Fresh backup confirmed for $osDrive." -Level SUCCESS

# -- Step 2: Stop and disable MBAM service --
Write-Log '--- Step 2: Stop MBAM service ---' -Level STEP
$null = Invoke-StopService -Name $Script:ServiceName

# -- Step 3: MSI uninstall --
Write-Log '--- Step 3: MSI uninstall ---' -Level STEP

$msiResult = Invoke-Process -FilePath 'msiexec.exe' `
    -Arguments "/x $($Script:MbamProductGuid) /qn /norestart REBOOT=ReallySuppress" `
    -TimeoutSeconds 600

if ($msiResult -eq 0) {
    Write-Log 'MBAM agent uninstalled successfully.' -Level SUCCESS
    $outputMsgs += 'MBAM agent removed.'

    # Package tracking registry
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
    $exitCode = 1
    $outputMsgs += "MBAM uninstall failed (exit $msiResult) -- service disabled as fallback."

    # Re-backup after failed uninstall (protector rotation risk)
    Write-Log 'Performing safety re-backup after failed uninstall ...' -Level WARN
    $null = Invoke-FreshBitLockerBackup -DriveLetter $osDrive
}

# -- Exit --
Write-Log "=== Remediation finished (ExitCode: $exitCode) ===" -Level STEP
if ($exitCode -eq 0) {
    Write-Output 'SUCCESS: MBAM agent removed. BitLocker keys confirmed in Entra ID.'
} else {
    Write-Output ($outputMsgs -join ' | ')
}
exit $exitCode

#endregion
