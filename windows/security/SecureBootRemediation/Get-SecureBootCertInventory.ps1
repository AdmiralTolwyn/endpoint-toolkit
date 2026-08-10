<#
.SYNOPSIS
    Read-only Secure Boot 2023 certificate inventory for Intune Remediations (detection script).

.DESCRIPTION
    Collects the documented Secure Boot servicing signals and emits a single compact JSON line
    so it lands in "Pre-remediation detection output" in the Intune admin center.

    NOTHING IS WRITTEN. No certificate update is triggered, no DBX is applied, no boot order
    or BitLocker state is touched. AvailableUpdates / MicrosoftUpdateManagedOptIn /
    HighConfidenceOptOut are read only.

    Sources for every field:
      HKLM\SYSTEM\CurrentControlSet\Control\SecureBoot            AvailableUpdates, AvailableUpdatesPolicy,
                                                                  HighConfidenceOptOut, MicrosoftUpdateManagedOptIn
      HKLM\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing  UEFICA2023Status, UEFICA2023Error,
                                                                  UEFICA2023ErrorEvent, WindowsUEFICA2023Capable,
                                                                  BucketHash, ConfidenceLevel
      \Microsoft\Windows\PI\Secure-Boot-Update                    scheduled task state / last run
      HKLM\SOFTWARE\Policies\Microsoft\Windows\DataCollection     AllowTelemetry, DisableOneSettingsDownloads
      System event log (source TPM-WMI)                           1795/1801/1802/1803/1808 + 1032/1043/1044/1045/1799
      UEFI variables via Get-SecureBootUEFI                       KEK / db contents

    Intune truncates detection output at 2048 characters. Output is kept short deliberately;
    do not add fields without checking the length.

.NOTES
    PowerShell 5.1. Must run as SYSTEM in 64-bit PowerShell (Get-SecureBootUEFI requires both).
    Exit 0 = UEFICA2023Status is Updated, or Secure Boot is not enabled (nothing to do).
    Exit 1 = anything else, i.e. the device belongs in the follow-up population.
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'SilentlyContinue'

function Get-RegVal {
    param([string]$Path, [string]$Name)
    $k = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
    if ($null -eq $k) { return $null }
    return $k.$Name
}

# Latin-1 is a lossless byte<->char map, unlike ASCII which folds bytes >127 to '?'.
function Test-CertInVar {
    param([byte[]]$Bytes, [string]$Subject)
    if ($null -eq $Bytes -or $Bytes.Length -eq 0) { return $null }
    $s = [System.Text.Encoding]::GetEncoding(28591).GetString($Bytes)
    return $s.Contains($Subject)
}

$sbRoot = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot'
$sbSvc  = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing'
$dcPol  = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection'
$dcPref = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection'

$o = [ordered]@{}
$o.ts   = (Get-Date).ToUniversalTime().ToString('s') + 'Z'
$o.host = $env:COMPUTERNAME

# --- Secure Boot enablement -------------------------------------------------
$sbEnabled = $null
try { $sbEnabled = Confirm-SecureBootUEFI -ErrorAction Stop } catch { $sbEnabled = $null }
$o.sbOn = $sbEnabled

# --- Servicing state --------------------------------------------------------
$status = Get-RegVal $sbSvc 'UEFICA2023Status'
if ($null -eq $status) { $status = 'NoValue' }   # Servicing key absent = update never initiated
$o.st      = $status
$o.err     = Get-RegVal $sbSvc 'UEFICA2023Error'
$o.errEvt  = Get-RegVal $sbSvc 'UEFICA2023ErrorEvent'
$o.capable = Get-RegVal $sbSvc 'WindowsUEFICA2023Capable'   # reference only - never the primary signal
$o.conf    = Get-RegVal $sbSvc 'ConfidenceLevel'

$bucket = Get-RegVal $sbSvc 'BucketHash'
if ($bucket) { $o.bkt = $bucket.Substring(0, [Math]::Min(16, $bucket.Length)) } else { $o.bkt = $null }

$avail = Get-RegVal $sbRoot 'AvailableUpdates'
if ($null -ne $avail) { $o.avl = ('0x{0:X}' -f $avail) } else { $o.avl = $null }
$availPol = Get-RegVal $sbRoot 'AvailableUpdatesPolicy'
if ($null -ne $availPol) { $o.avlPol = ('0x{0:X}' -f $availPol) } else { $o.avlPol = $null }

$o.hcOut = Get-RegVal $sbRoot 'HighConfidenceOptOut'
$o.muOin = Get-RegVal $sbRoot 'MicrosoftUpdateManagedOptIn'

# --- Firmware variables -----------------------------------------------------
$kekBytes = $null
$dbBytes  = $null
try { $kekBytes = (Get-SecureBootUEFI -Name KEK -ErrorAction Stop).Bytes } catch { }
try { $dbBytes  = (Get-SecureBootUEFI -Name db  -ErrorAction Stop).Bytes } catch { }

$o.kek23 = Test-CertInVar $kekBytes 'Microsoft Corporation KEK 2K CA 2023'
$o.kek11 = Test-CertInVar $kekBytes 'Microsoft Corporation KEK CA 2011'
$o.dbWin23 = Test-CertInVar $dbBytes 'Windows UEFI CA 2023'
$o.dbOr23  = Test-CertInVar $dbBytes 'Microsoft Option ROM UEFI CA 2023'
$o.db3p23  = Test-CertInVar $dbBytes 'Microsoft UEFI CA 2023'
$o.db3p11  = Test-CertInVar $dbBytes 'Microsoft Corporation UEFI CA 2011'
$o.dbPca11 = Test-CertInVar $dbBytes 'Microsoft Windows Production PCA 2011'

# Trust configuration is what decides whether the Option ROM / 3P certs are even applicable.
# The 0x4000 modifier bit only applies them when Microsoft Corporation UEFI CA 2011 is in db.
if ($null -eq $dbBytes) { $o.trust = 'Unknown' }
elseif ($o.db3p11 -or $o.db3p23 -or $o.dbOr23) { $o.trust = 'MSAnd3P' }
else { $o.trust = 'MSOnly' }

# --- Secure-Boot-Update scheduled task --------------------------------------
$task = Get-ScheduledTask -TaskPath '\Microsoft\Windows\PI\' -TaskName 'Secure-Boot-Update' -ErrorAction SilentlyContinue
if ($task) {
    $o.task = [string]$task.State
    $info = $task | Get-ScheduledTaskInfo -ErrorAction SilentlyContinue
    if ($info -and $info.LastRunTime -gt [datetime]'1900-01-01') {
        $o.taskRun = $info.LastRunTime.ToUniversalTime().ToString('s') + 'Z'
    } else { $o.taskRun = $null }
} else {
    $o.task    = 'NotFound'
    $o.taskRun = $null
}

# --- Reporting prerequisites ------------------------------------------------
$tel = Get-RegVal $dcPol 'AllowTelemetry'
if ($null -eq $tel) { $tel = Get-RegVal $dcPref 'AllowTelemetry' }
$o.tel  = $tel
$o.oneS = Get-RegVal $dcPol 'DisableOneSettingsDownloads'

# --- Event log --------------------------------------------------------------
$evtIds = 1795, 1797, 1798, 1799, 1801, 1802, 1803, 1808, 1032, 1043, 1044, 1045
$filter = @{ LogName = 'System'; ProviderName = 'Microsoft-Windows-TPM-WMI'; Id = $evtIds }
$events = Get-WinEvent -FilterHashtable $filter -MaxEvents 200 -ErrorAction SilentlyContinue

if ($events) {
    $latest = $events | Sort-Object TimeCreated -Descending | Select-Object -First 1
    $o.lastEvt   = $latest.Id
    $o.lastEvtAt = $latest.TimeCreated.ToUniversalTime().ToString('s') + 'Z'

    # Only the failure/blocking IDs are worth carrying per-device; success is implied by st=Updated.
    $counts = @{}
    foreach ($id in 1795, 1801, 1802, 1803, 1032) {
        $n = @($events | Where-Object { $_.Id -eq $id }).Count
        if ($n -gt 0) { $counts["$id"] = $n }
    }
    if ($counts.Count -gt 0) { $o.evt = $counts } else { $o.evt = $null }

    # SkipReason (1802) and firmware error code (1795) are the actionable payloads.
    $e1802 = $events | Where-Object { $_.Id -eq 1802 } | Select-Object -First 1
    if ($e1802) {
        $m = [regex]::Match($e1802.Message, 'SkipReason:\s*(\S+)')
        if ($m.Success) { $o.skip = $m.Groups[1].Value }
    }
} else {
    $o.lastEvt   = $null
    $o.lastEvtAt = $null
    $o.evt       = $null
}

# --- Platform context -------------------------------------------------------
$cs   = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
$bios = Get-CimInstance Win32_BIOS -ErrorAction SilentlyContinue
$os   = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue

$o.mfr   = $cs.Manufacturer
$o.model = $cs.Model
$o.fw    = $bios.SMBIOSBIOSVersion
$o.osb   = $os.BuildNumber

# Coarse platform class - drives the segmentation split.
$plat = 'Physical'
if ($cs.Manufacturer -match 'VMware')                       { $plat = 'VMware' }
elseif ($cs.Model -match 'Virtual Machine' -and $cs.Manufacturer -match 'Microsoft') {
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows Azure') { $plat = 'Azure' } else { $plat = 'HyperV' }
}
elseif ($cs.Model -match 'Virtual|KVM|Xen')                 { $plat = 'OtherVM' }
$o.plat = $plat

# --- Emit -------------------------------------------------------------------
$json = ($o | ConvertTo-Json -Compress -Depth 4)
Write-Output $json

if ($sbEnabled -ne $true) { exit 0 }
if ($status -eq 'Updated') { exit 0 }
exit 1
