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
      UEFI variables via Get-SecureBootUEFI                       PK / KEK / db contents
      VMware guest context (when Manufacturer/Model is VMware)    vTPM, VMware Tools version, Capsule PnP status

    Intune truncates detection output at 2048 characters. Output is kept short deliberately;
    do not add fields without checking the length.

.EXAMPLE
    .\Get-SecureBootCertInventory.ps1

    Emits a single-line JSON inventory record and exits 0 when compliant or confirmed Secure Boot disabled.

.OUTPUTS
    System.String (single-line compressed JSON) and System.Int32 exit code.

.NOTES
    PowerShell 5.1. Must run as SYSTEM in 64-bit PowerShell (Get-SecureBootUEFI requires both).
    Exit 0 = Confirmed compliant (Status=Updated, Error=0, KEK 2023 + required 2023 DB certs in NVRAM),
             or Secure Boot confirmed disabled/unsupported.
    Exit 1 = Non-compliant or unknown/unreadable state (belongs in follow-up population).
#>
#Requires -Version 5.1

[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'SilentlyContinue'

function Get-RegVal {
    param([string]$Path, [string]$Name)
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    $k = Get-ItemProperty -LiteralPath $Path -Name $Name -ErrorAction SilentlyContinue
    if ($null -eq $k) { return $null }
    if ($k.PSObject.Properties.Name -contains $Name) { return $k.$Name }
    return $null
}

# Latin-1 is a lossless byte<->char map, unlike ASCII which folds bytes >127 to '?'.
function Test-CertInVar {
    param([byte[]]$Bytes, [string]$Subject)
    if ($null -eq $Bytes -or $Bytes.Length -eq 0) { return $null }
    $s = [System.Text.Encoding]::GetEncoding(28591).GetString($Bytes)
    return $s.Contains($Subject)
}

function Classify-PkBytes {
    param([byte[]]$Bytes)
    if ($null -eq $Bytes -or $Bytes.Length -eq 0) { return 'None' }
    if ((Test-CertInVar $Bytes 'NULLPK') -eq $true) { return 'NullPK' }
    if ($Bytes.Length -le 44) { return 'NullPK' }

    $nonZero = $false
    for ($i = 44; $i -lt $Bytes.Length; $i++) {
        if ($Bytes[$i] -ne 0) { $nonZero = $true; break }
    }
    if (-not $nonZero) { return 'NullPK' }
    if ((Test-CertInVar $Bytes 'Windows OEM Devices PK') -eq $true) { return 'WinOemPK' }

    try {
        $len = $Bytes.Length - 44
        [byte[]]$der = New-Object byte[] $len
        [Array]::Copy($Bytes, 44, $der, 0, $len)
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList @(,$der)
        if ($cert -and -not [string]::IsNullOrWhiteSpace($cert.Subject)) {
            if ($cert.Subject -match 'Windows OEM Devices PK') { return 'WinOemPK' }
            if ($cert.Subject -match 'NULLPK')                 { return 'NullPK' }
            return 'OtherCertPK'
        }
    } catch { }

    return 'UnrecognizedPK'
}

function Resolve-ServiceExe {
    param([string]$ImagePath)
    if ([string]::IsNullOrWhiteSpace($ImagePath)) { return $null }
    $exp = [Environment]::ExpandEnvironmentVariables($ImagePath.Trim())
    if ($exp.StartsWith('"')) {
        $q = $exp.IndexOf('"', 1)
        if ($q -gt 1) { return $exp.Substring(1, $q - 1) }
    }
    if ($exp -match '^([A-Za-z]:\\[^\r\n"]*?\.exe)\b') { return $matches[1] }
    return $exp.Trim('"')
}

$sbRoot  = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot'
$sbState = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\State'
$sbSvc   = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing'
$dcPol   = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection'
$dcPref  = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection'

$o = [ordered]@{}
$o.ts   = (Get-Date).ToUniversalTime().ToString('s') + 'Z'
$o.host = $env:COMPUTERNAME

# --- Secure Boot enablement (preserve unknown vs disabled) ------------------
$sbEnabled = $null
try {
    $sb = Confirm-SecureBootUEFI -ErrorAction Stop
    if ($sb -is [bool]) {
        $sbEnabled = $sb
    }
} catch {
    $ex = $_.Exception
    $regSb = Get-RegVal $sbState 'UEFISecureBootEnabled'
    $regInt = 0
    if ($null -ne $regSb -and [int]::TryParse("$regSb", [ref]$regInt) -and $regInt -in @(0, 1)) {
        $sbEnabled = ($regInt -eq 1)
    } elseif ($null -eq $regSb -and ($ex -is [System.PlatformNotSupportedException] -or $ex.Message -match 'not supported on this platform|Cmdlet not supported')) {
        $sbEnabled = $false
    } else {
        $sbEnabled = $null
    }
}
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

# --- Firmware variables (PK / KEK / db) -------------------------------------
$pkBytes  = $null
$kekBytes = $null
$dbBytes  = $null
$pkSt     = 'Unreadable'

if ($sbEnabled -eq $false) {
    $pkSt = 'N/A'
} else {
    try {
        $pkVar = Get-SecureBootUEFI -Name PK -ErrorAction Stop
        if ($null -ne $pkVar) { $pkBytes = $pkVar.Bytes }
        $pkSt = Classify-PkBytes $pkBytes
    } catch {
        if ($_.Exception.Message -match '0xC0000100|currently not defined') {
            $pkSt = 'None'
        } else {
            $pkSt = 'Unreadable'
        }
    }
    try { $kekBytes = (Get-SecureBootUEFI -Name KEK -ErrorAction Stop).Bytes } catch { }
    try { $dbBytes  = (Get-SecureBootUEFI -Name db  -ErrorAction Stop).Bytes } catch { }
}

$o.pkSt    = $pkSt
$o.kek23   = Test-CertInVar $kekBytes 'Microsoft Corporation KEK 2K CA 2023'
$o.kek11   = Test-CertInVar $kekBytes 'Microsoft Corporation KEK CA 2011'
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

# --- Platform & VMware guest context ----------------------------------------
$cs          = $null
$bios        = $null
$os          = $null
$lastBootUtc = $null
try { $cs   = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop } catch { }
try { $bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction Stop } catch { }
try {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    if ($os -and $null -ne $os.LastBootUpTime) {
        $lastBootUtc = ([datetime]$os.LastBootUpTime).ToUniversalTime()
    }
} catch { }

# --- Event log --------------------------------------------------------------
$evtIds               = 1795, 1796, 1797, 1798, 1799, 1800, 1801, 1802, 1803, 1808, 1032, 1033, 1043, 1044, 1045
$blockingIds          = 1032, 1033, 1795, 1796, 1797, 1798, 1802, 1803
$filter               = @{ LogName = 'System'; ProviderName = 'Microsoft-Windows-TPM-WMI'; Id = $evtIds }
$events               = @()
$evtQueryOk           = $true
$hasValidCompEvt      = $false
$rebootPendingCurrent = $false

try {
    $events = @(Get-WinEvent -FilterHashtable $filter -MaxEvents 200 -ErrorAction Stop)
} catch {
    if ($_.Exception.Message -notmatch 'No events were found|not find any events') {
        $evtQueryOk = $false
    }
}

if ((@($events | Where-Object { $_.Id -in @(1808, 1799) }).Count -eq 0)) {
    try {
        $compFb = @(Get-WinEvent -FilterHashtable @{ LogName = 'System'; ProviderName = 'Microsoft-Windows-TPM-WMI'; Id = @(1808, 1799) } -MaxEvents 5 -ErrorAction Stop)
        if ($compFb.Count -gt 0) { $events = @($events + $compFb) }
    } catch {
        if ($_.Exception.Message -notmatch 'No events were found|not find any events') { $evtQueryOk = $false }
    }
}

if ((@($events | Where-Object { $_.Id -in $blockingIds }).Count -eq 0)) {
    try {
        $blkFb = @(Get-WinEvent -FilterHashtable @{ LogName = 'System'; ProviderName = 'Microsoft-Windows-TPM-WMI'; Id = $blockingIds } -MaxEvents 5 -ErrorAction Stop)
        if ($blkFb.Count -gt 0) { $events = @($events + $blkFb) }
    } catch {
        if ($_.Exception.Message -notmatch 'No events were found|not find any events') { $evtQueryOk = $false }
    }
}

$op1799 = @()
try {
    $op1799 = @(Get-WinEvent -FilterHashtable @{ LogName = 'Microsoft-Windows-TPM-WMI/Operational'; Id = 1799 } -MaxEvents 5 -ErrorAction Stop)
} catch { }

if ($events.Count -gt 0 -or $op1799.Count -gt 0) {
    $sorted = @($events | Sort-Object TimeCreated -Descending)
    if ($sorted.Count -gt 0) {
        $o.lastEvt   = $sorted[0].Id
        $o.lastEvtAt = if ($null -ne $sorted[0].TimeCreated) { ([datetime]$sorted[0].TimeCreated).ToUniversalTime().ToString('s') + 'Z' } else { $null }
    } else {
        $o.lastEvt   = 1799
        $o.lastEvtAt = if ($null -ne $op1799[0].TimeCreated) { ([datetime]$op1799[0].TimeCreated).ToUniversalTime().ToString('s') + 'Z' } else { $null }
    }

    $counts = @{}
    foreach ($id in 1795, 1796, 1801, 1802, 1803, 1032, 1808) {
        $n = @($sorted | Where-Object { $_.Id -eq $id }).Count
        if ($n -gt 0) { $counts["$id"] = $n }
    }
    if ($counts.Count -gt 0) { $o.evt = $counts } else { $o.evt = $null }

    $e1802 = $sorted | Where-Object { $_.Id -eq 1802 } | Select-Object -First 1
    if ($e1802) {
        $m = [regex]::Match($e1802.Message, 'SkipReason:[ \t]*(\S+)')
        if ($m.Success) { $o.skip = $m.Groups[1].Value }
    }

    $compEvts   = @(@($sorted | Where-Object { $_.Id -in @(1808, 1799) }) + $op1799 | Where-Object { $null -ne $_.TimeCreated } | Sort-Object TimeCreated -Descending)
    $allBlk     = @($sorted | Where-Object { $_.Id -in $blockingIds })
    $undatedBlk = (@($allBlk | Where-Object { $null -eq $_.TimeCreated }).Count -gt 0)
    $blkEvts    = @($allBlk | Where-Object { $null -ne $_.TimeCreated } | Sort-Object TimeCreated -Descending)
    $compTime   = if ($compEvts.Count -gt 0) { ([datetime]$compEvts[0].TimeCreated).ToUniversalTime() } else { $null }

    $ev1800 = @($sorted | Where-Object { $_.Id -eq 1800 -and $null -ne $_.TimeCreated } | Sort-Object TimeCreated -Descending)
    if ($ev1800.Count -gt 0 -and $null -ne $lastBootUtc) {
        $t1800     = ([datetime]$ev1800[0].TimeCreated).ToUniversalTime()
        $afterBoot = ($t1800 -gt $lastBootUtc)
        $afterComp = ($null -eq $compTime) -or ($t1800 -gt $compTime)
        $afterBlk  = (-not $undatedBlk) -and ($blkEvts.Count -eq 0 -or $t1800 -ge ([datetime]$blkEvts[0].TimeCreated).ToUniversalTime())
        $rebootPendingCurrent = ($afterBoot -and $afterComp -and $afterBlk)
    }

    if ($null -ne $compTime -and -not $undatedBlk -and -not $rebootPendingCurrent) {
        if ($blkEvts.Count -eq 0 -or ([datetime]$blkEvts[0].TimeCreated).ToUniversalTime() -lt $compTime) {
            $hasValidCompEvt = $true
        }
    }
} else {
    $o.lastEvt   = $null
    $o.lastEvtAt = $null
    $o.evt       = $null
}
$o.compOk     = if ($evtQueryOk) { $hasValidCompEvt } else { $null }
$o.rebootPend = $rebootPendingCurrent

$o.mfr   = if ($cs)   { $cs.Manufacturer }      else { $null }
$o.model = if ($cs)   { $cs.Model }             else { $null }
$o.fw    = if ($bios) { $bios.SMBIOSBIOSVersion } else { $null }
$o.osb   = if ($os)   { $os.BuildNumber }       else { $null }

# Coarse platform class - drives the segmentation split; preserve Unknown when CIM fails.
$plat = 'Unknown'
if ($cs -and (-not [string]::IsNullOrWhiteSpace("$($cs.Manufacturer)") -or -not [string]::IsNullOrWhiteSpace("$($cs.Model)"))) {
    if ($cs.Manufacturer -match 'VMware' -or $cs.Model -match 'VMware') {
        $plat = 'VMware'
    } elseif ($cs.Model -match 'Virtual Machine' -and $cs.Manufacturer -match 'Microsoft') {
        if (Test-Path -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows Azure') { $plat = 'Azure' } else { $plat = 'HyperV' }
    } elseif ($cs.Model -match 'Virtual|KVM|Xen|Bochs|QEMU') {
        $plat = 'OtherVM'
    } else {
        $plat = 'Physical'
    }
}
$o.plat = $plat

# vTPM presence (VMware / virtual guests only; physical TPM is not vTPM)
$vtpm = $null
if ($plat -eq 'VMware') {
    try {
        $t = Get-Tpm -ErrorAction Stop
        if ($null -ne $t) { $vtpm = [bool]$t.TpmPresent }
    } catch {
        try {
            $ct = @(Get-CimInstance -Namespace 'root\CIMV2\Security\MicrosoftTpm' -ClassName Win32_Tpm -ErrorAction Stop)
            $vtpm = ($ct.Count -gt 0)
        } catch {
            $vtpm = $null
        }
    }
}
$o.vtpm = $vtpm

# VMware-specific fields (VMware Tools version + Capsule PK PnP device code)
if ($plat -eq 'VMware') {
    $vmt = 'NotInstalled'
    $imgPath = Get-RegVal 'HKLM:\SYSTEM\CurrentControlSet\Services\VMTools' 'ImagePath'
    if (-not [string]::IsNullOrWhiteSpace($imgPath)) {
        $exe = Resolve-ServiceExe $imgPath
        if ($exe -and (Test-Path -LiteralPath $exe)) {
            try {
                $rawVer = (Get-Item -LiteralPath $exe -ErrorAction Stop).VersionInfo.ProductVersion
                if ($rawVer -match '^(\d+\.\d+\.\d+)') { $vmt = $matches[1] }
                elseif (-not [string]::IsNullOrWhiteSpace($rawVer)) { $vmt = $rawVer.Trim() }
                else { $vmt = 'Unknown' }
            } catch { $vmt = 'Unknown' }
        } else {
            $vmt = 'Unknown'
        }
    }
    $o.vmt = $vmt

    $pkCap = 'Absent'
    try {
        $pkDev = @(Get-CimInstance -ClassName Win32_PnPEntity -Filter "Name LIKE '%Platform Key Firmware%'" -ErrorAction Stop | Select-Object -First 1)
        if ($pkDev.Count -gt 0 -and $null -ne $pkDev[0]) {
            $pkCap = "$($pkDev[0].ConfigManagerErrorCode)"
        }
    } catch {
        $pkCap = 'QueryError'
    }
    $o.pkCap = $pkCap
} else {
    $o.vmt   = $null
    $o.pkCap = $null
}

# --- Emit -------------------------------------------------------------------
$json = ($o | ConvertTo-Json -Compress -Depth 4)
Write-Output $json

if ($sbEnabled -eq $false -and $plat -ne 'Unknown') { exit 0 }

$dbAllOk    = ($o.dbWin23 -eq $true) -and ($o.trust -eq 'MSOnly' -or ($o.trust -eq 'MSAnd3P' -and $o.dbOr23 -eq $true -and $o.db3p23 -eq $true))
$errParsed  = 0L
$errOk      = ($null -eq $o.err) -or ([int64]::TryParse("$($o.err)", [ref]$errParsed) -and $errParsed -eq 0)
$pkReadOk   = ($o.pkSt -notin @('Unreadable', 'N/A'))
$capOk      = ($plat -ne 'VMware' -or $o.pkCap -in @('Absent', '0'))

if ($sbEnabled -eq $true -and $plat -ne 'Unknown' -and $pkReadOk -and $capOk -and
    $status -eq 'Updated' -and $errOk -and $o.avl -eq '0x4000' -and
    $o.kek23 -eq $true -and $dbAllOk -and $evtQueryOk -and $hasValidCompEvt -and -not $rebootPendingCurrent) {
    exit 0
}
exit 1
