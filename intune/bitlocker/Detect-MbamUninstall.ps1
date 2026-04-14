<#
.SYNOPSIS
    Detects whether the MBAM agent is still present after BitLocker keys have been escrowed.

.DESCRIPTION
    Detection script for the standalone MBAM cleanup package.

    Checks for the MBAM agent via:
      1. MBAMAgent Windows service
      2. MSI product GUID in the Uninstall registry (native + WOW6432Node)

    Only flags non-compliant (exit 1) if MBAM is present AND there is evidence
    the OS drive has already been backed up to Entra ID (Event 845 or registry
    marker). This prevents premature removal before key escrow is confirmed.

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
$Script:Component   = 'Detect'

$Script:ServiceName    = 'MBAMAgent'
$Script:MbamProductGuid = '{AEC5BCA3-A2C5-46D7-9873-7698E6D3CAA4}'

$Script:RegistryKey    = 'HKLM:\SOFTWARE\ZF\BitLocker'
$Script:RegistryName   = 'Drive_{0}_BitLockerBackupToAAD'
$Script:RegistryValue  = 'True'

$Script:EventProvider  = 'Microsoft-Windows-BitLocker-API'
$Script:EventMessage   = 'volume {0} was backed up successfully to your Azure AD.'
$Script:EventID        = 845
$Script:EventSince     = [DateTime]'2022-01-01'

$Script:UninstallPaths = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
)

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

function Test-MbamPresent {
    <#
    .SYNOPSIS  Returns $true if MBAM agent is installed (service or MSI product).
    #>

    # Check service
    $svc = Get-Service -Name $Script:ServiceName -ErrorAction SilentlyContinue
    if ($svc) {
        Write-Log "MBAMAgent service found (Status: $($svc.Status), StartType: $($svc.StartType))." -Level INFO
        return $true
    }

    # Check MSI product in registry
    foreach ($path in $Script:UninstallPaths) {
        $key = Join-Path $path $Script:MbamProductGuid
        if (Test-Path $key) {
            $product = Get-ItemProperty -Path $key -ErrorAction SilentlyContinue
            Write-Log "MBAM MSI product found in registry: $($product.DisplayName) ($path)." -Level INFO
            return $true
        }
    }

    return $false
}

function Test-BackupEvent {
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
    param([Parameter(Mandatory)] [string]$DriveLetter)

    $name = $Script:RegistryName -f ($DriveLetter -replace ':', '')
    try {
        if (-not (Test-Path $Script:RegistryKey)) { return $false }
        $val = Get-ItemProperty -Path $Script:RegistryKey -Name $name -ErrorAction SilentlyContinue
        return ($null -ne $val -and $val.$name -eq $Script:RegistryValue)
    }
    catch { return $false }
}

#endregion

# ─────────────────────────────────────────────────────────────────────────────
#region Main
# ─────────────────────────────────────────────────────────────────────────────

$Script:ScriptName = $MyInvocation.MyCommand.Name
$exitCode = 0

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

# -- Check MBAM presence --
$mbamFound = Test-MbamPresent
if (-not $mbamFound) {
    Write-Log 'MBAM agent is not installed -- compliant.' -Level SUCCESS
    Write-Output 'COMPLIANT: MBAM agent is not present.'
    Write-Log "--- Detection finished (ExitCode: 0) ---" -Level STEP
    exit 0
}

# -- MBAM is present: check if OS drive is backed up to Entra ID --
$osDrive   = $env:SystemDrive
$hasEvent  = Test-BackupEvent    -DriveLetter $osDrive
$hasReg    = Test-BackupRegistry -DriveLetter $osDrive
$hasBackup = $hasEvent -or $hasReg

Write-Log "OS drive backup status: Event=$hasEvent, Registry=$hasReg" -Level INFO

if ($hasBackup) {
    $msg = "MBAM agent present AND OS drive backup confirmed -- remediation needed."
    Write-Log $msg -Level WARN
    Write-Output "NON-COMPLIANT: $msg"
    $exitCode = 1
}
else {
    $msg = "MBAM agent present but OS drive NOT yet backed up -- deferring removal."
    Write-Log $msg -Level WARN
    Write-Output "COMPLIANT (deferred): $msg"
    $exitCode = 0
}

Write-Log "--- Detection finished (ExitCode: $exitCode) ---" -Level STEP
exit $exitCode

#endregion
