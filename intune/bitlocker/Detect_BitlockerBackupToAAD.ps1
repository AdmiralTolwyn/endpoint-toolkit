<#
.SYNOPSIS
    Detects whether BitLocker recovery keys are escrowed to Entra ID and protection is active.

.DESCRIPTION
    Unified detection script for the MBAM-to-Entra ID BitLocker key migration.

    For every fully-encrypted fixed drive the script checks:
      1. Is BitLocker protection turned ON?
      2. Has the recovery key been backed up to Entra ID?
         (Event 845 from Microsoft-Windows-BitLocker-API, or registry marker)

    If ANY drive fails either check, the script exits 1 (non-compliant) to trigger
    the paired remediation script.

    MBAM presence is logged as informational only -- it does not block detection.

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
$Script:Component   = 'Detect'

$Script:ServiceName    = 'MBAMAgent'
$Script:CheckAllDrives = $true

$Script:RegistryKey    = 'HKLM:\SOFTWARE\ZF\BitLocker'
$Script:RegistryName   = 'Drive_{0}_BitLockerBackupToAAD'   # {0} = drive letter without colon
$Script:RegistryValue  = 'True'

$Script:EventProvider  = 'Microsoft-Windows-BitLocker-API'
$Script:EventMessage   = 'volume {0} was backed up successfully to your Azure AD.'
$Script:EventID        = 845
$Script:EventSince     = [DateTime]'2022-01-01'

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

    # CMTrace-compatible format
    $entry = '<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="{4}" type="{5}" thread="{6}" file="{7}">' -f
        $Message, $ts, $dt, $caller, $env:USERNAME, $severity, $PID, $Script:ScriptName

    if (-not (Test-Path $Script:LogDir)) {
        try { New-Item -Path $Script:LogDir -ItemType Directory -Force -ErrorAction Stop | Out-Null }
        catch { Write-Warning "Cannot create log directory: $_"; return }
    }

    try { $entry | Out-File -FilePath $Script:LogFile -Append -NoClobber -Force -Encoding default -ErrorAction Stop }
    catch { Write-Warning "Log write failed: $_" }

    # Rotate when over max size
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

function Test-BackupEvent {
    <#
    .SYNOPSIS  Checks event log for successful BitLocker-to-AAD backup (ID 845).
    #>
    param([Parameter(Mandatory)] [string]$DriveLetter)

    $msg = $Script:EventMessage -f $DriveLetter
    try {
        $provider = (Get-WinEvent -ListProvider $Script:EventProvider -ErrorAction SilentlyContinue).Name
        if (-not $provider) { return $false }

        $hit = Get-WinEvent -ProviderName $Script:EventProvider -ErrorAction SilentlyContinue |
            Where-Object { $_.TimeCreated -gt $Script:EventSince -and $_.Message -match [regex]::Escape($msg) -and $_.Id -eq $Script:EventID } |
            Select-Object -First 1

        return ($null -ne $hit)
    }
    catch {
        Write-Log "Event log query failed for ${DriveLetter}: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

function Test-BackupRegistry {
    <#
    .SYNOPSIS  Checks the registry marker for a successful BitLocker backup.
    #>
    param([Parameter(Mandatory)] [string]$DriveLetter)

    $name = $Script:RegistryName -f ($DriveLetter -replace ':', '')
    try {
        if (-not (Test-Path $Script:RegistryKey)) { return $false }
        $val = Get-ItemProperty -Path $Script:RegistryKey -Name $name -ErrorAction SilentlyContinue
        return ($null -ne $val -and $val.$name -eq $Script:RegistryValue)
    }
    catch { return $false }
}

function Get-ProtectionStatus {
    param([Parameter(Mandatory)] [string]$DriveLetter)
    try { return (Get-BitLockerVolume -MountPoint $DriveLetter -ErrorAction Stop).ProtectionStatus }
    catch { return $null }
}

function Get-VolumeStatus {
    param([Parameter(Mandatory)] [string]$DriveLetter)
    try { return (Get-BitLockerVolume -MountPoint $DriveLetter -ErrorAction Stop).VolumeStatus }
    catch { return $null }
}

#endregion

# ─────────────────────────────────────────────────────────────────────────────
#region Main
# ─────────────────────────────────────────────────────────────────────────────

$Script:ScriptName = $MyInvocation.MyCommand.Name
$exitCode   = 0
$outputMsgs = @()

Write-Log '--- Detection started ---' -Level STEP

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

# -- MBAM presence (informational) --
$mbamSvc = Get-Service -Name $Script:ServiceName -ErrorAction SilentlyContinue
if ($mbamSvc) {
    Write-Log "MBAMAgent service detected (Status: $($mbamSvc.Status)). Informational only." -Level WARN
}

# -- Enumerate drives --
if ($Script:CheckAllDrives) {
    $drives = @((Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 }).DeviceID)
} else {
    $drives = @($env:SystemDrive)
}
Write-Log "Drives to check: $($drives -join ', ')" -Level INFO

# -- Per-drive checks --
foreach ($drive in $drives) {
    Write-Log "--- Checking drive $drive ---" -Level STEP

    $volStatus = Get-VolumeStatus -DriveLetter $drive
    if ($volStatus -notmatch 'FullyEncrypted') {
        Write-Log "Drive $drive not fully encrypted (Status: $volStatus). Skipped." -Level WARN
        continue
    }
    Write-Log "Drive $drive is FullyEncrypted." -Level INFO

    $protStatus = Get-ProtectionStatus -DriveLetter $drive
    if ($protStatus -ne 'On') {
        $msg = "Drive $drive protection is OFF."
        Write-Log $msg -Level ERROR
        $outputMsgs += $msg
        $exitCode = 1
        continue
    }
    Write-Log "Drive $drive protection is ON." -Level INFO

    $hasEvent    = Test-BackupEvent    -DriveLetter $drive
    $hasRegistry = Test-BackupRegistry -DriveLetter $drive
    if ($hasEvent -or $hasRegistry) {
        Write-Log "Drive $drive recovery key is in Entra ID (Event=$hasEvent, Registry=$hasRegistry)." -Level SUCCESS
    }
    else {
        $msg = "Drive $drive recovery key is NOT in Entra ID."
        Write-Log $msg -Level ERROR
        $outputMsgs += $msg
        $exitCode = 1
    }
}

# -- Exit --
Write-Log "--- Detection finished (ExitCode: $exitCode) ---" -Level STEP
if ($exitCode -eq 0) {
    Write-Output 'COMPLIANT: All drives encrypted, protected, and backed up to Entra ID.'
} else {
    Write-Output ($outputMsgs -join ' | ')
}
exit $exitCode

#endregion
