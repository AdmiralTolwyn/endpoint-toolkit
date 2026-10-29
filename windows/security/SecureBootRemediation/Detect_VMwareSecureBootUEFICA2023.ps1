<#
.SYNOPSIS
    Read-only Intune Remediations detection script for VMware Virtual Machine
    Secure Boot UEFI CA 2023 and Platform Key (PK) readiness.

.DESCRIPTION
    Evaluates a Windows device against the combined Microsoft Secure Boot CA 2023
    servicing requirements (KB 5085046 / KB 5016061), Microsoft Windows Server
    Secure Boot guidance, and Broadcom KB 423893.

    STRICTLY READ-ONLY. No registry values, UEFI variables, scheduled tasks, or
    BitLocker protectors are modified. Diagnostic log entries are the only writes.

    Decision Tree:
      1. Collection / prerequisite checks:
         - Not running in 64-bit PowerShell -> Exit 1 (PREREQ).
         - Platform CIM query failed or Secure Boot state unreadable -> Exit 1
           (NON-COMPLIANT-COLLECTION-ERROR). Unknown states never exit 0.
      2. Confirmed out-of-scope devices:
         - Non-VMware platform -> Exit 0 (NOT-APPLICABLE).
         - VMware VM with Secure Boot confirmed disabled or Legacy BIOS -> Exit 0
           (NOT-APPLICABLE).
      3. UEFI NVRAM read check on Secure Boot-enabled VMware VMs:
         - If PK, KEK, or db cannot be read via Get-SecureBootUEFI -> Exit 1
           (NON-COMPLIANT-COLLECTION-ERROR).
      4. Capsule PK PnP pending restart check:
         - VMware 'Platform Key Firmware' PnP device present with
           ConfigManagerErrorCode = 14 (CM_PROB_NEED_RESTART) -> Exit 1
           (NON-COMPLIANT-VMWARE-CAPSULE-REBOOT).
      5. Strict compliance check (live NVRAM + servicing state + completion event):
         - KEK 2023 + Windows UEFI CA 2023 + applicable 3P 2023 DB certs (when
           Microsoft Corporation UEFI CA 2011 is in db) present in live NVRAM,
           AND UEFICA2023Status = 'Updated', UEFICA2023Error = 0,
           AvailableUpdates = 0x4000, AND Event 1808 (or 1799) present with no
           newer servicing error -> Exit 0 (COMPLIANT).
      6. Capsule PK PnP device error check:
         - 'Platform Key Firmware' PnP device has any non-zero error code other
           than 0 or 14 -> Exit 1 (NON-COMPLIANT-VMWARE-CAPSULE-ERROR).
      7. KEK 2023 is ALREADY present in live NVRAM (including ESXi 8.0 U2-U3i P08
         VMs that still carry VMW.NULLPK):
         - PK replacement is not required to complete Windows DB/Boot Manager
           servicing when KEK 2023 is already enrolled in NVRAM.
         - If AvailableUpdates = 0x4100 or a current-boot Event 1800 is pending ->
           Exit 1 (NON-COMPLIANT-PENDING-REBOOT).
         - Otherwise (missing 2023 DB certs such as Microsoft Option ROM UEFI CA
           2023, or OS Boot Manager servicing not yet completed) -> Exit 1
           (NON-COMPLIANT-OS-SERVICING-NEEDED).
      8. KEK 2023 is MISSING from NVRAM AND PK is NullPK or None:
         - OS-initiated KEK updates cannot be authorized without Windows OEM
           Devices PK in vUEFI PK (Broadcom KB 423893) -> Exit 1
           (NON-COMPLIANT-VMWARE-PK-BLOCKED).
      9. KEK 2023 is MISSING from NVRAM AND PK has 'Windows OEM Devices PK':
         - PK is ready for OS-driven KEK/DB servicing -> Exit 1
           (NON-COMPLIANT-PK-READY-NEEDS-KEK).
     10. KEK 2023 is MISSING from NVRAM AND PK is custom/unrecognized:
         - Exit 1 (NON-COMPLIANT-VMWARE-PK-UNRECOGNIZED).

    Logs are written in CMTrace format to:
        %ProgramData%\Microsoft\IntuneManagementExtension\Logs\PR_VMwareSecureBootCA2023.log

    Intune Proactive Remediation settings:
      Run this script using the logged-on credentials : No
      Enforce script signature check                  : No (or sign per org policy)
      Run script in 64-bit PowerShell                 : Yes

    References:
      Broadcom KB 423893 - Secure Boot Certificate Expirations and Update Failures in VMware VMs
      Microsoft KB 5085046 - Secure Boot troubleshooting guide
      Microsoft KB 5080921 - Monitoring Secure Boot certificate status with Intune remediations
      Microsoft KB 5016061 - Secure Boot DB and DBX variable update events

.INPUTS
    None.

.OUTPUTS
    System.Int32 exit code:
      0 = COMPLIANT or confirmed NOT-APPLICABLE
      1 = NON-COMPLIANT-* or PREREQ

    Single-line STDOUT summary (< 2048 chars) with marker prefix and key=value telemetry
    suitable for the Intune Remediations "Pre-remediation detection output" column.

.EXAMPLE
    .\Detect_VMwareSecureBootUEFICA2023.ps1

    Evaluates VMware Secure Boot CA 2023 and PK readiness without modifying the device.

.NOTES
    Author:   Anton Romanyuk
    Version:  1.1
    Date:     2026-10-29
    Context:  VMware vSphere Guest Secure Boot UEFI CA 2023 / Platform Key Detection
    Requires: Windows PowerShell 5.1, 64-bit, Administrator / SYSTEM

    DISCLAIMER:
    THIS SCRIPT IS PROVIDED "AS-IS" WITHOUT WARRANTY OF ANY KIND.
    USE AT YOUR OWN RISK.
#>

#Requires -Version 5.1
[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# ─────────────────────────────────────────────────────────────────────────────
# Configuration
# ─────────────────────────────────────────────────────────────────────────────

$Script:LogDir             = Join-Path $env:ProgramData 'Microsoft\IntuneManagementExtension\Logs'
$Script:LogName            = 'PR_VMwareSecureBootCA2023'
$Script:LogFile            = Join-Path $Script:LogDir "$Script:LogName.log"
$Script:LogMaxSize         = 250KB
$Script:Component          = 'DetectVMwareSB'
$Script:ScriptName         = if ($MyInvocation.MyCommand.Name) { $MyInvocation.MyCommand.Name } else { 'Detect_VMwareSecureBootUEFICA2023.ps1' }
$Script:LogWarningEmitted  = $false

$Script:RegPathRoot        = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot'
$Script:RegPathState       = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\State'
$Script:RegPathServicing   = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing'
$Script:RegPathDeviceAttr  = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing\DeviceAttributes'
$Script:CompliantAvUpdates = 0x4000

$Script:EventProvider        = 'Microsoft-Windows-TPM-WMI'
$Script:GoodEventIDs         = @(1034, 1036, 1037, 1042, 1043, 1044, 1045, 1799, 1800, 1808)
$Script:WarningEventIDs      = @(1801)
$Script:BlockingEventIDs     = @(1032, 1033, 1795, 1796, 1797, 1798, 1802, 1803)
$Script:AllSecureBootEvents  = $Script:GoodEventIDs + $Script:WarningEventIDs + $Script:BlockingEventIDs
$Script:EventQueryMax        = 100

$Script:TaskPath = '\Microsoft\Windows\PI\'
$Script:TaskName = 'Secure-Boot-Update'

# ─────────────────────────────────────────────────────────────────────────────
#region Logging
# ─────────────────────────────────────────────────────────────────────────────

function Write-Log {
    <#
    .SYNOPSIS
        Writes a CMTrace-formatted log entry to the Intune ME log directory.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet('DEBUG', 'INFO', 'WARN', 'ERROR', 'SUCCESS', 'STEP')]
        [string]$Level = 'INFO'
    )

    $ts = Get-Date -Format 'HH:mm:ss.fffffff'
    $dt = Get-Date -Format 'MM-dd-yyyy'
    $severity = switch ($Level) { 'ERROR' { 3 } 'WARN' { 2 } default { 1 } }
    $stack = @(Get-PSCallStack)
    $caller = if ($stack.Count -gt 1 -and $stack[1].FunctionName -and $stack[1].FunctionName -ne '<ScriptBlock>') {
        $stack[1].FunctionName
    } else {
        $Script:Component
    }

    $tagged = '[{0}] {1}' -f $Level, $Message
    $entry  = '<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="{4}" type="{5}" thread="{6}" file="{7}">' -f
        $tagged, $ts, $dt, $caller, $env:USERNAME, $severity, $PID, $Script:ScriptName

    if (-not (Test-Path -LiteralPath $Script:LogDir)) {
        try {
            New-Item -Path $Script:LogDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
        } catch {
            if (-not $Script:LogWarningEmitted) {
                $Script:LogWarningEmitted = $true
                Write-Warning "Cannot create log directory '$($Script:LogDir)': $($_.Exception.Message)"
            }
            return
        }
    }

    try {
        [System.IO.File]::AppendAllText($Script:LogFile, $entry + [Environment]::NewLine, [System.Text.Encoding]::Default)
    } catch {
        if (-not $Script:LogWarningEmitted) {
            $Script:LogWarningEmitted = $true
            Write-Warning "Cannot write to log file '$($Script:LogFile)': $($_.Exception.Message)"
        }
        return
    }

    try {
        if ((Test-Path -LiteralPath $Script:LogFile) -and (Get-Item -LiteralPath $Script:LogFile -ErrorAction Stop).Length -ge $Script:LogMaxSize) {
            $archive = Join-Path $Script:LogDir "$Script:LogName-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
            Copy-Item -LiteralPath $Script:LogFile -Destination $archive -Force -ErrorAction Stop
            Remove-Item -LiteralPath $Script:LogFile -Force -ErrorAction Stop
            Compact /C $archive 2>$null | Out-Null
        }
    } catch {
        if (-not $Script:LogWarningEmitted) {
            $Script:LogWarningEmitted = $true
            Write-Warning "Log rotation failed for '$($Script:LogFile)': $($_.Exception.Message)"
        }
    }
}

function Format-ReasonText {
    <#
    .SYNOPSIS
        Sanitizes free-text reason strings for single-line summary output.
    #>
    param([AllowNull()] $Value)
    if ($null -eq $Value) { return $null }
    $clean = ("$Value" -replace '[\r\n|]+', ' ' -replace '\s+', ' ').Trim()
    if ([string]::IsNullOrEmpty($clean)) { return $null }
    return $clean
}

function Format-TelemetryToken {
    <#
    .SYNOPSIS
        Sanitizes a value for space-delimited key=value telemetry output (no internal spaces).
    #>
    param([AllowNull()] $Value)
    if ($null -eq $Value) { return $null }
    $clean = ("$Value" -replace '[\r\n|=]+', ' ' -replace '\s+', ' ').Trim()
    if ([string]::IsNullOrEmpty($clean)) { return $null }
    return ($clean -replace ' ', '_')
}

function Complete-Detection {
    <#
    .SYNOPSIS
        Logs the final verdict, emits the single-line STDOUT contract for Intune, and exits.
    #>
    param(
        [Parameter(Mandatory = $true)] [int]$ExitCode,
        [Parameter(Mandatory = $true)] [string]$Marker,
        [Parameter(Mandatory = $true)] [string]$Reason,
        [Parameter(Mandatory = $true)] [System.Collections.IDictionary]$Telemetry
    )

    $cleanReason = Format-ReasonText -Value $Reason
    $pairs = foreach ($k in $Telemetry.Keys) {
        $v = Format-TelemetryToken -Value $Telemetry[$k]
        if ($null -ne $v) { '{0}={1}' -f $k, $v }
    }
    $summary = '{0}: {1} | {2}' -f $Marker, $cleanReason, ($pairs -join ' ')
    if ($summary.Length -gt 2000) {
        $summary = $summary.Substring(0, 1997) + '...'
    }

    $lvl = if ($ExitCode -eq 0) { 'SUCCESS' } else { 'WARN' }
    Write-Log "Verdict -> $summary" -Level $lvl
    Write-Log "--- Detection finished (ExitCode: $ExitCode) ---" -Level STEP
    Write-Output $summary
    exit $ExitCode
}

#endregion

# ─────────────────────────────────────────────────────────────────────────────
#region Helpers
# ─────────────────────────────────────────────────────────────────────────────

function Test-Is64BitPS {
    <#
    .SYNOPSIS
        Returns $true when the current PowerShell process is 64-bit.
    #>
    $is64 = [Environment]::Is64BitProcess
    Write-Log "Test-Is64BitPS -> $is64 (PSArch=$([IntPtr]::Size * 8)-bit, OSArch=$([Environment]::Is64BitOperatingSystem))" -Level DEBUG
    return $is64
}

function Get-RegValue {
    <#
    .SYNOPSIS
        Safely reads a single registry value, returning $null if absent.
    #>
    param(
        [Parameter(Mandatory = $true)] [string]$Path,
        [Parameter(Mandatory = $true)] [string]$Name
    )
    Write-Log "Get-RegValue ENTER (Path='$Path', Name='$Name')" -Level DEBUG
    try {
        if (-not (Test-Path -LiteralPath $Path)) {
            Write-Log "Get-RegValue: path '$Path' does not exist -> `$null" -Level DEBUG
            return $null
        }
        $item = Get-ItemProperty -LiteralPath $Path -Name $Name -ErrorAction Stop
        $val = $item.$Name
        $disp = if ($null -eq $val) { '<null>' } elseif ($val -is [int] -or $val -is [long]) { "{0} (0x{0:X})" -f $val } else { "'$val'" }
        Write-Log "Get-RegValue EXIT -> $disp" -Level DEBUG
        return $val
    } catch {
        Write-Log "Get-RegValue: '$Name' absent or unreadable -> $($_.Exception.Message)" -Level DEBUG
        return $null
    }
}

function Get-CanAttemptUpdateAfter {
    <#
    .SYNOPSIS
        Returns the CanAttemptUpdateAfter throttle as a UTC DateTime, or $null if absent.
    #>
    $raw = Get-RegValue -Path $Script:RegPathDeviceAttr -Name 'CanAttemptUpdateAfter'
    if ($null -eq $raw) { return $null }
    try {
        if ($raw -is [array] -and $raw -isnot [byte[]]) {
            $raw = [byte[]]$raw
        }
        if ($raw -is [byte[]]) {
            if ($raw.Length -lt 8) { return $null }
            $ft = [BitConverter]::ToInt64($raw, 0)
            if ($ft -le 0) { return $null }
            return [DateTime]::FromFileTimeUtc($ft)
        }
        if ($raw -is [long] -or $raw -is [int]) {
            if ([int64]$raw -le 0) { return $null }
            return [DateTime]::FromFileTimeUtc([int64]$raw)
        }
        return $null
    } catch {
        Write-Log "Get-CanAttemptUpdateAfter parse failed: $($_.Exception.Message)" -Level DEBUG
        return $null
    }
}

function Get-SecureBootEnablementState {
    <#
    .SYNOPSIS
        Determines Secure Boot enablement while preserving unknown/error states.
    .OUTPUTS
        OrderedDictionary with keys: State ('Enabled', 'Disabled', 'Unsupported', 'Unknown'),
        Enabled ($true, $false, or $null), ErrorMessage.
    #>
    $res = [ordered]@{
        State        = 'Unknown'
        Enabled      = $null
        ErrorMessage = $null
    }

    try {
        $sb = Confirm-SecureBootUEFI -ErrorAction Stop
        if ($sb -is [bool]) {
            if ($sb) {
                $res.State   = 'Enabled'
                $res.Enabled = $true
            } else {
                $res.State   = 'Disabled'
                $res.Enabled = $false
            }
            Write-Log "Confirm-SecureBootUEFI -> $($res.State)" -Level INFO
            return $res
        }
        $res.ErrorMessage = "Confirm-SecureBootUEFI returned non-boolean value '$sb'"
        Write-Log $res.ErrorMessage -Level ERROR
    } catch {
        $ex = $_.Exception
        $res.ErrorMessage = $ex.Message
        Write-Log "Confirm-SecureBootUEFI threw $($ex.GetType().FullName): $($ex.Message)" -Level DEBUG

        $regSb = Get-RegValue -Path $Script:RegPathState -Name 'UEFISecureBootEnabled'
        $regInt = 0
        if ($null -ne $regSb -and [int]::TryParse("$regSb", [ref]$regInt) -and $regInt -in @(0, 1)) {
            if ($regInt -eq 1) {
                $res.State   = 'Enabled'
                $res.Enabled = $true
            } else {
                $res.State   = 'Disabled'
                $res.Enabled = $false
            }
            Write-Log "Fallback UEFISecureBootEnabled registry value ($regInt) -> $($res.State)" -Level WARN
            return $res
        }

        if ($null -eq $regSb -and ($ex -is [System.PlatformNotSupportedException] -or $ex.Message -match 'not supported on this platform|Cmdlet not supported')) {
            $res.State   = 'Unsupported'
            $res.Enabled = $false
            Write-Log 'Secure Boot is not supported on this platform (Legacy BIOS).' -Level INFO
            return $res
        }

        Write-Log "Secure Boot enablement state is UNKNOWN: $($ex.Message)" -Level ERROR
        return $res
    }

    return $res
}

function Test-UefiCertPresent {
    <#
    .SYNOPSIS
        Checks whether a certificate subject string is present in raw UEFI variable bytes.
        Uses ISO-8859-1 (codepage 28591) for lossless 1:1 byte-to-char mapping.
    #>
    param(
        [AllowNull()] [byte[]]$Bytes,
        [Parameter(Mandatory = $true)] [string]$Subject
    )
    if ($null -eq $Bytes -or $Bytes.Length -eq 0) { return $null }
    $latin1 = [System.Text.Encoding]::GetEncoding(28591).GetString($Bytes)
    return $latin1.Contains($Subject)
}

function Classify-PlatformKeyBytes {
    <#
    .SYNOPSIS
        Classifies raw UEFI PK variable bytes without assuming arbitrary non-zero bytes
        are a valid authorization chain.
    .OUTPUTS
        String: 'None', 'NullPK', 'WinOemPK', 'OtherCertPK', or 'UnrecognizedPK'.
    #>
    param([AllowNull()] [byte[]]$Bytes)

    if ($null -eq $Bytes -or $Bytes.Length -eq 0) {
        return 'None'
    }

    if ((Test-UefiCertPresent -Bytes $Bytes -Subject 'NULLPK') -eq $true) {
        return 'NullPK'
    }

    # Broadcom KB 423893: EFI_SIGNATURE_LIST header (28 bytes) + SignatureOwner GUID (16 bytes) = 44 bytes.
    # VMW.NULLPK has <= 44 bytes or all-zero payload bytes starting at index 44 (e.g., 45 bytes with 0x00).
    if ($Bytes.Length -le 44) {
        return 'NullPK'
    }

    $hasNonZeroPayload = $false
    for ($i = 44; $i -lt $Bytes.Length; $i++) {
        if ($Bytes[$i] -ne 0) {
            $hasNonZeroPayload = $true
            break
        }
    }
    if (-not $hasNonZeroPayload) {
        return 'NullPK'
    }

    if ((Test-UefiCertPresent -Bytes $Bytes -Subject 'Windows OEM Devices PK') -eq $true) {
        return 'WinOemPK'
    }

    # Attempt X.509 DER parse on the signature payload starting at offset 44 to distinguish
    # a non-Windows OEM X.509 certificate from arbitrary/unrecognized bytes.
    try {
        $payloadLen = $Bytes.Length - 44
        [byte[]]$der = New-Object byte[] $payloadLen
        [Array]::Copy($Bytes, 44, $der, 0, $payloadLen)
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList @(,$der)
        if ($cert -and -not [string]::IsNullOrWhiteSpace($cert.Subject)) {
            if ($cert.Subject -match 'Windows OEM Devices PK') { return 'WinOemPK' }
            if ($cert.Subject -match 'NULLPK')                 { return 'NullPK' }
            return 'OtherCertPK'
        }
    } catch {
        Write-Log "PK payload starting at offset 44 is not a standalone X.509 DER cert: $($_.Exception.Message)" -Level DEBUG
    }

    return 'UnrecognizedPK'
}

function Get-UefiVariableState {
    <#
    .SYNOPSIS
        Reads UEFI PK, KEK, and db variables once and classifies PK + 2023/2011 cert presence.
        Preserves 'Unreadable' / $null states when Get-SecureBootUEFI fails.
    #>
    $state = [ordered]@{
        UefiReadable    = $false
        PkReadOk        = $false
        KekReadOk       = $false
        DbReadOk        = $false
        PkState         = 'Unreadable'
        Kek2023         = $null
        Kek2011         = $null
        DbWin2023       = $null
        DbOptRom2023    = $null
        Db3p2023        = $null
        Db3p2011        = $null
        DbPca2011       = $null
        TrustMode       = 'Unknown'
        AllRequiredDb23 = $null
        MissingDb23     = 'Unknown'
    }

    $pkBytes  = $null
    $kekBytes = $null
    $dbBytes  = $null

    try {
        $pkVar = Get-SecureBootUEFI -Name PK -ErrorAction Stop
        $state.PkReadOk = $true
        if ($null -ne $pkVar) { $pkBytes = $pkVar.Bytes }
        $state.PkState = Classify-PlatformKeyBytes -Bytes $pkBytes
    } catch {
        $msg = $_.Exception.Message
        if ($msg -match '0xC0000100|currently not defined') {
            $state.PkReadOk = $true
            $state.PkState  = 'None'
            Write-Log "Get-SecureBootUEFI PK indicates variable not defined (Setup Mode / None)." -Level WARN
        } else {
            Write-Log "Get-SecureBootUEFI PK failed: $msg" -Level ERROR
        }
    }

    try {
        $kekVar = Get-SecureBootUEFI -Name KEK -ErrorAction Stop
        if ($null -ne $kekVar -and $null -ne $kekVar.Bytes -and $kekVar.Bytes.Length -gt 0) {
            $state.KekReadOk = $true
            $kekBytes        = $kekVar.Bytes
            $state.Kek2023   = Test-UefiCertPresent -Bytes $kekBytes -Subject 'Microsoft Corporation KEK 2K CA 2023'
            $state.Kek2011   = Test-UefiCertPresent -Bytes $kekBytes -Subject 'Microsoft Corporation KEK CA 2011'
        } else {
            Write-Log 'Get-SecureBootUEFI KEK returned empty bytes.' -Level WARN
        }
    } catch {
        Write-Log "Get-SecureBootUEFI KEK failed: $($_.Exception.Message)" -Level ERROR
    }

    try {
        $dbVar = Get-SecureBootUEFI -Name db -ErrorAction Stop
        if ($null -ne $dbVar -and $null -ne $dbVar.Bytes -and $dbVar.Bytes.Length -gt 0) {
            $state.DbReadOk     = $true
            $dbBytes            = $dbVar.Bytes
            $state.DbWin2023    = Test-UefiCertPresent -Bytes $dbBytes -Subject 'Windows UEFI CA 2023'
            $state.DbOptRom2023 = Test-UefiCertPresent -Bytes $dbBytes -Subject 'Microsoft Option ROM UEFI CA 2023'
            $state.Db3p2023     = Test-UefiCertPresent -Bytes $dbBytes -Subject 'Microsoft UEFI CA 2023'
            $state.Db3p2011     = Test-UefiCertPresent -Bytes $dbBytes -Subject 'Microsoft Corporation UEFI CA 2011'
            $state.DbPca2011    = Test-UefiCertPresent -Bytes $dbBytes -Subject 'Microsoft Windows Production PCA 2011'

            # Per MS KB 5085046, the 0x4000 modifier bit applies Microsoft Option ROM UEFI CA 2023 (0x0800)
            # and Microsoft UEFI CA 2023 (0x1000) only when Microsoft Corporation UEFI CA 2011 is present in db.
            $missing = New-Object System.Collections.Generic.List[string]
            if (-not $state.DbWin2023) { [void]$missing.Add('WindowsUEFICA2023') }

            if ($state.Db3p2011) {
                $state.TrustMode = 'MSAnd3P'
                if (-not $state.DbOptRom2023) { [void]$missing.Add('MSOptionRomUEFICA2023') }
                if (-not $state.Db3p2023)     { [void]$missing.Add('MSUEFICA2023') }
            } elseif ($state.Db3p2023 -or $state.DbOptRom2023) {
                $state.TrustMode = 'MSAnd3P'
            } else {
                $state.TrustMode = 'MSOnly'
            }

            $state.AllRequiredDb23 = ($missing.Count -eq 0)
            $state.MissingDb23     = if ($missing.Count -eq 0) { 'None' } else { $missing -join ',' }
        } else {
            Write-Log 'Get-SecureBootUEFI db returned empty bytes.' -Level WARN
        }
    } catch {
        Write-Log "Get-SecureBootUEFI db failed: $($_.Exception.Message)" -Level ERROR
    }

    $state.UefiReadable = ($state.PkReadOk -and $state.KekReadOk -and $state.DbReadOk)

    Write-Log ("UEFI State -> Readable={0}, PK={1}, KEK23={2}, KEK11={3}, DbWin23={4}, DbOptRom23={5}, Db3P23={6}, Db3P11={7}, Trust={8}, AllReqDb23={9}, MissingDb23={10}" -f
        $state.UefiReadable, $state.PkState, $state.Kek2023, $state.Kek2011,
        $state.DbWin2023, $state.DbOptRom2023, $state.Db3p2023, $state.Db3p2011,
        $state.TrustMode, $state.AllRequiredDb23, $state.MissingDb23) -Level INFO
    return $state
}

function Resolve-ServiceExecutablePath {
    <#
    .SYNOPSIS
        Extracts the executable file path from a Windows service ImagePath string.
    #>
    param([AllowNull()] [string]$ImagePath)
    if ([string]::IsNullOrWhiteSpace($ImagePath)) { return $null }
    $expanded = [Environment]::ExpandEnvironmentVariables($ImagePath.Trim())
    if ($expanded.StartsWith('"')) {
        $endQuote = $expanded.IndexOf('"', 1)
        if ($endQuote -gt 1) {
            return $expanded.Substring(1, $endQuote - 1)
        }
    }
    if ($expanded -match '^([A-Za-z]:\\[^\r\n"]*?\.exe)\b') {
        return $matches[1]
    }
    return $expanded.Trim('"')
}

function Get-CapsuleOsEligibility {
    <#
    .SYNOPSIS
        Evaluates Guest OS build/product type against Broadcom KB 423893's
        Supported Windows Guest OS List for Capsule PK update in VMware Tools 13.1.5.
    .DESCRIPTION
        Per Broadcom KB 423893:
          - Supported in VMware Tools 13.1.5 (with July 2026+ CU):
              Windows Server 2025 (26100), Windows Server 2022 (20348), Windows Server 2019 (17763 Server),
              Windows 11 25H2 (26200), 24H2 (26100), 23H2 (22631), Windows 10 21H2 (19044).
          - Deferred to a future VMware Tools release (NOT in 13.1.5):
              Windows Server 2016 (14393), Windows 10 22H2 (19045), Windows 10 1809 (17763 Client).
          - CU install order:
              Server 2025 and Win11 24H2/25H2 (>= 26100) allow installing July 2026+ CU before or after
              upgrading to VMTools 13.1.5. Older supported builds require July 2026+ CU BEFORE upgrading
              VMTools to 13.1.5.
    #>
    param(
        [AllowNull()] [string]$BuildNumber,
        [AllowNull()] $ProductType
    )

    $info = [ordered]@{
        Status        = 'Unknown'
        CuBeforeTools = $true
        Note          = 'OS build unknown'
    }

    $build = 0
    if (-not [int]::TryParse("$BuildNumber", [ref]$build)) {
        return $info
    }

    $isServer = ($null -ne $ProductType -and [int]$ProductType -ne 1)

    if ($build -ge 26100) {
        $info.Status        = 'SupportedIn1315'
        $info.CuBeforeTools = $false
        $info.Note          = 'Supported in VMTools 13.1.5 (July 2026+ CU can be installed before or after Tools 13.1.5)'
    } elseif ($build -in @(22631, 20348, 19044) -or ($build -eq 17763 -and $isServer)) {
        $info.Status        = 'SupportedIn1315'
        $info.CuBeforeTools = $true
        $info.Note          = 'Supported in VMTools 13.1.5 (install July 2026+ CU prior to upgrading to Tools 13.1.5)'
    } elseif ($build -in @(14393, 19045) -or ($build -eq 17763 -and -not $isServer)) {
        $info.Status        = 'DeferredToolsRelease'
        $info.CuBeforeTools = $true
        $info.Note          = 'Not supported in VMTools 13.1.5 per Broadcom KB 423893 (Server 2016 / Win10 22H2 / Win10 1809 require a future VMTools release)'
    } else {
        $info.Status        = 'UnlistedBuild'
        $info.CuBeforeTools = $true
        $info.Note          = "OS build $build is not listed in Broadcom KB 423893 Capsule PK table"
    }

    return $info
}

function Get-VMwareGuestContext {
    <#
    .SYNOPSIS
        Collects platform manufacturer/model, vTPM presence, VMware Tools version,
        and Broadcom Capsule PK PnP firmware device status while preserving Unknown states.
    #>
    $ctx = [ordered]@{
        PlatformState     = 'Unknown'
        IsVMware          = $null
        Manufacturer      = 'Unknown'
        Model             = 'Unknown'
        BiosVersion       = 'Unknown'
        OsBuild           = 'Unknown'
        OsProductType     = $null
        LastBootUtc       = $null
        VtpmPresent       = 'N/A'
        VMToolsVersion    = 'N/A'
        VMTools1315Ready  = $false
        CapsuleOsSupport  = 'N/A'
        CapsuleOsNote     = $null
        CapsuleDeviceCode = 'N/A'
    }

    try {
        $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        if ($null -ne $cs) {
            $mfr   = "$($cs.Manufacturer)".Trim()
            $model = "$($cs.Model)".Trim()
            if (-not [string]::IsNullOrWhiteSpace($mfr))   { $ctx.Manufacturer = $mfr }
            if (-not [string]::IsNullOrWhiteSpace($model)) { $ctx.Model        = $model }

            if (-not [string]::IsNullOrWhiteSpace($mfr) -or -not [string]::IsNullOrWhiteSpace($model)) {
                if ($mfr -match 'VMware' -or $model -match 'VMware') {
                    $ctx.PlatformState = 'VMware'
                    $ctx.IsVMware      = $true
                } else {
                    $ctx.PlatformState = 'NonVMware'
                    $ctx.IsVMware      = $false
                }
            }
        }
    } catch {
        Write-Log "Win32_ComputerSystem query failed: $($_.Exception.Message)" -Level ERROR
    }

    try {
        $bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction Stop
        if ($null -ne $bios -and -not [string]::IsNullOrWhiteSpace("$($bios.SMBIOSBIOSVersion)")) {
            $ctx.BiosVersion = "$($bios.SMBIOSBIOSVersion)".Trim()
        }
    } catch {
        Write-Log "Win32_BIOS query failed: $($_.Exception.Message)" -Level WARN
    }

    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        if ($null -ne $os) {
            if (-not [string]::IsNullOrWhiteSpace("$($os.BuildNumber)")) {
                $ctx.OsBuild = "$($os.BuildNumber)".Trim()
            }
            $ctx.OsProductType = $os.ProductType
            if ($null -ne $os.LastBootUpTime) {
                $ctx.LastBootUtc = ([datetime]$os.LastBootUpTime).ToUniversalTime()
            }
        }
    } catch {
        Write-Log "Win32_OperatingSystem query failed: $($_.Exception.Message)" -Level WARN
    }

    if ($ctx.IsVMware -eq $true) {
        $ctx.VtpmPresent = 'Unknown'
        $tpmKnown = $false
        try {
            $tpm = Get-Tpm -ErrorAction Stop
            if ($null -ne $tpm) {
                $ctx.VtpmPresent = [bool]$tpm.TpmPresent
                $tpmKnown        = $true
            }
        } catch {
            Write-Log "Get-Tpm failed ($($_.Exception.Message)); trying Win32_Tpm CIM fallback." -Level DEBUG
        }

        if (-not $tpmKnown) {
            try {
                $cimTpm = @(Get-CimInstance -Namespace 'root\CIMV2\Security\MicrosoftTpm' -ClassName Win32_Tpm -ErrorAction Stop)
                $ctx.VtpmPresent = ($cimTpm.Count -gt 0)
            } catch {
                Write-Log "Win32_Tpm fallback failed: $($_.Exception.Message). Keeping VtpmPresent='Unknown'." -Level WARN
            }
        }

        $capElig = Get-CapsuleOsEligibility -BuildNumber $ctx.OsBuild -ProductType $ctx.OsProductType
        $ctx.CapsuleOsSupport = $capElig.Status
        $ctx.CapsuleOsNote    = $capElig.Note

        $ctx.VMToolsVersion   = 'NotInstalled'
        $imgPath = Get-RegValue -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\VMTools' -Name 'ImagePath'
        if (-not [string]::IsNullOrWhiteSpace($imgPath)) {
            $exePath = Resolve-ServiceExecutablePath -ImagePath $imgPath
            if ($exePath -and (Test-Path -LiteralPath $exePath)) {
                try {
                    $rawVer = (Get-Item -LiteralPath $exePath -ErrorAction Stop).VersionInfo.ProductVersion
                    if ($rawVer -match '^(\d+\.\d+\.\d+)') {
                        $ctx.VMToolsVersion = $matches[1]
                        $parsedVer = $null
                        if ([version]::TryParse($matches[1], [ref]$parsedVer)) {
                            $ctx.VMTools1315Ready = ($parsedVer -ge [version]'13.1.5')
                        }
                    } elseif (-not [string]::IsNullOrWhiteSpace($rawVer)) {
                        $ctx.VMToolsVersion = $rawVer.Trim()
                    } else {
                        $ctx.VMToolsVersion = 'Unknown'
                    }
                } catch {
                    $ctx.VMToolsVersion = 'Unknown'
                    Write-Log "Failed reading VMTools version from '$exePath': $($_.Exception.Message)" -Level WARN
                }
            } else {
                $ctx.VMToolsVersion = 'Unknown'
                Write-Log "VMTools service ImagePath '$imgPath' could not be resolved on disk." -Level WARN
            }
        }

        $ctx.CapsuleDeviceCode = 'Absent'
        try {
            $pkDev = @(Get-CimInstance -ClassName Win32_PnPEntity -Filter "Name LIKE '%Platform Key Firmware%'" -ErrorAction Stop |
                Select-Object -First 1)
            if ($pkDev.Count -gt 0 -and $null -ne $pkDev[0]) {
                $ctx.CapsuleDeviceCode = "$($pkDev[0].ConfigManagerErrorCode)"
            }
        } catch {
            $ctx.CapsuleDeviceCode = 'QueryError'
            Write-Log "Win32_PnPEntity query for Platform Key Firmware failed: $($_.Exception.Message)" -Level WARN
        }
    }

    Write-Log ("Platform Context -> State={0}, Mfr='{1}', Model='{2}', BIOS='{3}', OSBuild={4}, vTPM={5}, VMTools={6} (13.1.5+={7}), CapOS={8}, CapsulePnPCode={9}" -f
        $ctx.PlatformState, $ctx.Manufacturer, $ctx.Model, $ctx.BiosVersion, $ctx.OsBuild, $ctx.VtpmPresent,
        $ctx.VMToolsVersion, $ctx.VMTools1315Ready, $ctx.CapsuleOsSupport, $ctx.CapsuleDeviceCode) -Level INFO
    return $ctx
}

function Get-SecureBootEvents {
    <#
    .SYNOPSIS
        Queries System log (Microsoft-Windows-TPM-WMI) and Operational log for 1799,
        evaluating event freshness against completion events and last boot time.
    #>
    param([AllowNull()] [nullable[datetime]]$LastBootUtc)

    $res = [ordered]@{
        QueryOk               = $true
        LatestId              = 'None'
        LatestBlockingId      = $null
        Has1808               = $false
        Has1803               = $false
        Has1802               = $false
        Has1800               = $false
        Has1795               = $false
        Has1796               = $false
        Has1799               = $false
        RebootPendingCurrent  = $false
        Active1802            = $false
        Post1808BlockingError = $false
        Event1795Code         = $null
        Event1796Code         = $null
        SkipReason            = $null
        Confidence            = $null
        BucketId              = $null
    }

    $events = @()
    try {
        $events = @(Get-WinEvent -FilterHashtable @{
            LogName      = 'System'
            ProviderName = $Script:EventProvider
            Id           = $Script:AllSecureBootEvents
        } -MaxEvents $Script:EventQueryMax -ErrorAction Stop)
    } catch {
        if ($_.Exception.Message -notmatch 'No events were found|not find any events') {
            $res.QueryOk = $false
            Write-Log "Get-WinEvent System query failed: $($_.Exception.Message)" -Level WARN
        }
    }

    # Ensure completion (1808/1799), pending-restart (1800), and blocking events are never hidden
    # if noisy 1801 assessment events filled the top 100 window.
    $hasCompOr1800InBatch = ((@($events | Where-Object { $_.Id -in @(1808, 1799) }).Count -gt 0) -and (@($events | Where-Object { $_.Id -eq 1800 }).Count -gt 0))
    if (-not $hasCompOr1800InBatch) {
        try {
            $eCompFallback = @(Get-WinEvent -FilterHashtable @{
                LogName      = 'System'
                ProviderName = $Script:EventProvider
                Id           = @(1808, 1799, 1800)
            } -MaxEvents 10 -ErrorAction Stop)
            if ($eCompFallback.Count -gt 0) {
                $events = @($events + $eCompFallback)
            }
        } catch {
            if ($_.Exception.Message -notmatch 'No events were found|not find any events') {
                $res.QueryOk = $false
                Write-Log "Get-WinEvent completion/1800 fallback query failed: $($_.Exception.Message)" -Level WARN
            }
        }
    }

    $hasBlockerInBatch = (@($events | Where-Object { $_.Id -in $Script:BlockingEventIDs }).Count -gt 0)
    if (-not $hasBlockerInBatch) {
        try {
            $eBlockFallback = @(Get-WinEvent -FilterHashtable @{
                LogName      = 'System'
                ProviderName = $Script:EventProvider
                Id           = $Script:BlockingEventIDs
            } -MaxEvents 5 -ErrorAction Stop)
            if ($eBlockFallback.Count -gt 0) {
                $events = @($events + $eBlockFallback)
            }
        } catch {
            if ($_.Exception.Message -notmatch 'No events were found|not find any events') {
                $res.QueryOk = $false
                Write-Log "Get-WinEvent blocker fallback query failed: $($_.Exception.Message)" -Level WARN
            }
        }
    }

    # Check TPM-WMI/Operational for Event 1799 (bootloader swap) and merge into event list so timestamps are compared.
    $op1799Events = @()
    try {
        $evt1799Op = @(Get-WinEvent -FilterHashtable @{
            LogName = 'Microsoft-Windows-TPM-WMI/Operational'
            Id      = 1799
        } -MaxEvents 5 -ErrorAction Stop)
        if ($evt1799Op.Count -gt 0) {
            $op1799Events = $evt1799Op
        }
    } catch { }

    if ($events.Count -eq 0 -and $op1799Events.Count -eq 0) {
        return $res
    }

    $sorted = @($events | Sort-Object TimeCreated -Descending)
    if ($sorted.Count -gt 0) {
        $res.LatestId = "$($sorted[0].Id)"
    } elseif ($op1799Events.Count -gt 0) {
        $res.LatestId = '1799'
    }

    $ev1808  = @($sorted | Where-Object { $_.Id -eq 1808 -and $null -ne $_.TimeCreated })
    $ev1803  = @($sorted | Where-Object { $_.Id -eq 1803 })
    $ev1802  = @($sorted | Where-Object { $_.Id -eq 1802 })
    $ev1800  = @($sorted | Where-Object { $_.Id -eq 1800 })
    $ev1795  = @($sorted | Where-Object { $_.Id -eq 1795 })
    $ev1796  = @($sorted | Where-Object { $_.Id -eq 1796 })
    $ev1799  = @(@($sorted | Where-Object { $_.Id -eq 1799 }) + $op1799Events | Where-Object { $null -ne $_.TimeCreated } | Sort-Object TimeCreated -Descending)
    $evBlock = @($sorted | Where-Object { $_.Id -in $Script:BlockingEventIDs })

    $res.Has1808 = ($ev1808.Count -gt 0)
    $res.Has1803 = ($ev1803.Count -gt 0)
    $res.Has1802 = ($ev1802.Count -gt 0)
    $res.Has1800 = ($ev1800.Count -gt 0)
    $res.Has1795 = ($ev1795.Count -gt 0)
    $res.Has1796 = ($ev1796.Count -gt 0)
    $res.Has1799 = ($ev1799.Count -gt 0)
    if ($evBlock.Count -gt 0) { $res.LatestBlockingId = "$($evBlock[0].Id)" }

    # Use the newest completion timestamp across both 1808 and 1799 (System + Operational).
    $completionCandidates = @($ev1808 + $ev1799 | Sort-Object TimeCreated -Descending)
    $latestCompletionTime = if ($completionCandidates.Count -gt 0) {
        ([datetime]$completionCandidates[0].TimeCreated).ToUniversalTime()
    } else {
        $null
    }

    if ($evBlock.Count -gt 0) {
        $undatedBlocker = (@($evBlock | Where-Object { $null -eq $_.TimeCreated }).Count -gt 0)
        $datedBlockers  = @($evBlock | Where-Object { $null -ne $_.TimeCreated } | Sort-Object TimeCreated -Descending)
        if ($undatedBlocker -or $null -eq $latestCompletionTime) {
            $res.Post1808BlockingError = $true
        } elseif ($datedBlockers.Count -gt 0 -and ([datetime]$datedBlockers[0].TimeCreated).ToUniversalTime() -ge $latestCompletionTime) {
            $res.Post1808BlockingError = $true
        }
    }

    # Historical Event 1800 must not imply a current reboot requirement: require known LastBootUtc
    # and Event 1800 newer than LastBootUtc, newer than any completion event, and not superseded by a blocker.
    # If Event 1800 is newer than completion (or undated) and LastBootUtc is unknown, fail closed by setting QueryOk=$false.
    $undated1800 = (@($ev1800 | Where-Object { $null -eq $_.TimeCreated }).Count -gt 0)
    $dated1800   = @($ev1800 | Where-Object { $null -ne $_.TimeCreated } | Sort-Object TimeCreated -Descending)
    if ($undated1800) {
        $res.QueryOk = $false
    } elseif ($dated1800.Count -gt 0) {
        $t1800           = ([datetime]$dated1800[0].TimeCreated).ToUniversalTime()
        $afterCompletion = ($null -eq $latestCompletionTime) -or ($t1800 -gt $latestCompletionTime)
        $afterBlocker    = ($evBlock.Count -eq 0 -or $null -eq $evBlock[0].TimeCreated) -or ($t1800 -ge ([datetime]$evBlock[0].TimeCreated).ToUniversalTime())
        if ($afterCompletion -and $afterBlocker) {
            if ($null -eq $LastBootUtc) {
                $res.QueryOk = $false
            } elseif ($t1800 -gt $LastBootUtc) {
                $res.RebootPendingCurrent = $true
            }
        }
    }

    if ($ev1802.Count -gt 0) {
        if ($null -eq $latestCompletionTime -or ([datetime]$ev1802[0].TimeCreated).ToUniversalTime() -gt $latestCompletionTime) {
            $res.Active1802 = $true
        }
    }

    # Use [ \t]* instead of \s* so an empty field value does not cross the newline and capture the next label.
    $msgEvt = @($sorted | Where-Object { $_.Id -in @(1801, 1802, 1808) } | Select-Object -First 1)
    if ($msgEvt.Count -gt 0 -and $null -ne $msgEvt[0].Message) {
        $m = $msgEvt[0].Message
        if ($m -match 'BucketConfidenceLevel:[ \t]*([^\r\n]*)') {
            $val = $matches[1].Trim()
            if (-not [string]::IsNullOrEmpty($val)) { $res.Confidence = $val }
        }
        if ($m -match 'BucketId:[ \t]*([^\r\n]+)') {
            $res.BucketId = $matches[1].Trim()
        }
        if ($m -match 'SkipReason:[ \t]*(KI_\d+)') {
            $res.SkipReason = $matches[1].Trim()
        }
    }

    if ($ev1795.Count -gt 0 -and $ev1795[0].Message -match '(?:error|code|status)[:\s]*(?:0x)?([0-9A-Fa-f]{1,8})') {
        $res.Event1795Code = '0x' + $matches[1]
    }
    if ($ev1796.Count -gt 0 -and $ev1796[0].Message -match '(?:error|code|status)[:\s]*(?:0x)?([0-9A-Fa-f]{1,8})') {
        $res.Event1796Code = '0x' + $matches[1]
    }

    Write-Log ("Events -> Latest={0}, 1808={1}, 1799={2}, 1803={3}, 1802={4} (Active={5}, Skip={6}), 1800={7} (CurrentPending={8}), 1795={9} ({10}), 1796={11} ({12})" -f
        $res.LatestId, $res.Has1808, $res.Has1799, $res.Has1803, $res.Has1802, $res.Active1802, $res.SkipReason,
        $res.Has1800, $res.RebootPendingCurrent, $res.Has1795, $res.Event1795Code, $res.Has1796, $res.Event1796Code) -Level INFO
    return $res
}

function Get-SecureBootTaskState {
    <#
    .SYNOPSIS
        Checks existence and state of \Microsoft\Windows\PI\Secure-Boot-Update.
    #>
    try {
        $t = Get-ScheduledTask -TaskPath $Script:TaskPath -TaskName $Script:TaskName -ErrorAction Stop
        return [string]$t.State
    } catch {
        Write-Log "Get-ScheduledTask failed or task missing: $($_.Exception.Message)" -Level DEBUG
        return 'Missing'
    }
}

#endregion

# ─────────────────────────────────────────────────────────────────────────────
#region Main
# ─────────────────────────────────────────────────────────────────────────────

Write-Log '--- VMware Secure Boot CA 2023 Detection started ---' -Level STEP
Write-Log ("Script={0} | PSVersion={1} | PSArch={2}-bit | Host={3} | User={4}\{5} | PID={6}" -f
    $Script:ScriptName, $PSVersionTable.PSVersion, ([IntPtr]::Size * 8), $env:COMPUTERNAME, $env:USERDOMAIN, $env:USERNAME, $PID) -Level DEBUG

if (-not (Test-Is64BitPS)) {
    Write-Log 'Script is not running in 64-bit PowerShell.' -Level ERROR
    Complete-Detection -ExitCode 1 -Marker 'PREREQ' `
        -Reason 'Not running in 64-bit PowerShell (set Run script in 64-bit PowerShell = Yes)' `
        -Telemetry ([ordered]@{ Host = $env:COMPUTERNAME; PSArch = "$([IntPtr]::Size * 8)-bit" })
}

$vmCtx  = Get-VMwareGuestContext
$sbInfo = Get-SecureBootEnablementState

# Collect servicing registry state (absent UEFICA2023Error on a readable Servicing key means 0 / no error)
$status       = Get-RegValue -Path $Script:RegPathServicing -Name 'UEFICA2023Status'
if ([string]::IsNullOrEmpty($status)) { $status = 'NoValue' }
$errRaw       = Get-RegValue -Path $Script:RegPathServicing -Name 'UEFICA2023Error'
$errParsed    = 0L
$err          = if ($null -eq $errRaw) { 0L } elseif ([int64]::TryParse("$errRaw", [ref]$errParsed)) { $errParsed } else { $null }
$errEvt       = Get-RegValue -Path $Script:RegPathServicing -Name 'UEFICA2023ErrorEvent'
$kekErrReason = Get-RegValue -Path $Script:RegPathServicing -Name 'KEKLastUpdateErrorReason'
$capable      = Get-RegValue -Path $Script:RegPathServicing -Name 'WindowsUEFICA2023Capable'
$regConf      = Get-RegValue -Path $Script:RegPathServicing -Name 'ConfidenceLevel'
$avUpdates    = Get-RegValue -Path $Script:RegPathRoot      -Name 'AvailableUpdates'
$avPolicy     = Get-RegValue -Path $Script:RegPathRoot      -Name 'AvailableUpdatesPolicy'
$avHex        = if ($null -ne $avUpdates) { '0x{0:X}' -f $avUpdates } else { 'None' }
$avPolHex     = if ($null -ne $avPolicy)  { '0x{0:X}' -f $avPolicy }  else { $null }
$optOut       = Get-RegValue -Path $Script:RegPathRoot      -Name 'HighConfidenceOptOut'
$canAttemptAt = Get-CanAttemptUpdateAfter

$taskState = Get-SecureBootTaskState
$evt       = Get-SecureBootEvents -LastBootUtc $vmCtx.LastBootUtc

$uefi = if ($vmCtx.IsVMware -eq $true -and $sbInfo.State -eq 'Enabled') {
    Get-UefiVariableState
} else {
    [ordered]@{
        UefiReadable    = $false
        PkReadOk        = $false
        KekReadOk       = $false
        DbReadOk        = $false
        PkState         = 'N/A'
        Kek2023         = 'N/A'
        DbWin2023       = 'N/A'
        DbOptRom2023    = 'N/A'
        Db3p2023        = 'N/A'
        AllRequiredDb23 = 'N/A'
        MissingDb23     = 'N/A'
        TrustMode       = 'N/A'
    }
}

# Operational warnings for log correlation
if ($taskState -notin @('Ready', 'Running')) {
    Write-Log "Operational warning: Secure-Boot-Update scheduled task state is '$taskState'." -Level WARN
}
if ($null -ne $canAttemptAt -and $canAttemptAt -gt (Get-Date).ToUniversalTime()) {
    Write-Log ("Operational warning: CanAttemptUpdateAfter throttle active until {0}Z." -f $canAttemptAt.ToString('s')) -Level WARN
}
if ($null -ne $avPolicy -and $null -ne $avUpdates -and $avPolicy -ne $avUpdates) {
    Write-Log ("Operational warning: AvailableUpdatesPolicy ({0}) differs from AvailableUpdates ({1})." -f $avPolHex, $avHex) -Level WARN
}
if ($null -ne $optOut -and [int]$optOut -eq 1) {
    Write-Log 'Operational warning: HighConfidenceOptOut=1 is set in registry.' -Level WARN
}

$telemetry = [ordered]@{
    Host       = $env:COMPUTERNAME
    Plat       = if ($vmCtx.PlatformState -eq 'VMware') { 'VMware' } elseif ($vmCtx.PlatformState -eq 'NonVMware') { $vmCtx.Manufacturer } else { 'Unknown' }
    SB         = $sbInfo.State
    vTPM       = $vmCtx.VtpmPresent
    Status     = $status
    Err        = if ($null -eq $err) { 'Invalid' } elseif ($err -ne 0) { ('0x{0:X}' -f $err) } else { 0 }
    ErrEvt     = $errEvt
    KekReason  = $kekErrReason
    Avl        = $avHex
    AvlPol     = $avPolHex
    OptOut     = $optOut
    Capable    = $capable
    Conf       = if ($evt.Confidence) { $evt.Confidence } else { $regConf }
    PK         = $uefi.PkState
    KEK23      = $uefi.Kek2023
    DbWin23    = $uefi.DbWin2023
    DbOptRom23 = $uefi.DbOptRom2023
    Db3P23     = $uefi.Db3p2023
    DbAll23    = $uefi.AllRequiredDb23
    Trust      = $uefi.TrustMode
    VMTools    = $vmCtx.VMToolsVersion
    CapOS      = $vmCtx.CapsuleOsSupport
    CapPnP     = $vmCtx.CapsuleDeviceCode
    LastEvt    = $evt.LatestId
    Skip       = $evt.SkipReason
    Err1795    = $evt.Event1795Code
    Err1796    = $evt.Event1796Code
    Task       = $taskState
    OSBuild    = $vmCtx.OsBuild
}

# 1. Scope & collection health checks (unknown states never exit 0)
if ($vmCtx.PlatformState -eq 'Unknown') {
    Complete-Detection -ExitCode 1 -Marker 'NON-COMPLIANT-COLLECTION-ERROR' `
        -Reason 'Unable to determine platform manufacturer/model via CIM (Win32_ComputerSystem)' -Telemetry $telemetry
}

if ($vmCtx.PlatformState -eq 'NonVMware') {
    Complete-Detection -ExitCode 0 -Marker 'NOT-APPLICABLE' `
        -Reason "Device is not a VMware virtual machine (Manufacturer=$($vmCtx.Manufacturer), Model=$($vmCtx.Model))" -Telemetry $telemetry
}

if ($sbInfo.State -eq 'Disabled') {
    Complete-Detection -ExitCode 0 -Marker 'NOT-APPLICABLE' `
        -Reason 'Secure Boot is disabled on this VMware virtual machine' -Telemetry $telemetry
}

if ($sbInfo.State -eq 'Unsupported') {
    Complete-Detection -ExitCode 0 -Marker 'NOT-APPLICABLE' `
        -Reason 'Secure Boot is not supported on this VMware virtual machine (Legacy BIOS)' -Telemetry $telemetry
}

if ($sbInfo.State -ne 'Enabled') {
    Complete-Detection -ExitCode 1 -Marker 'NON-COMPLIANT-COLLECTION-ERROR' `
        -Reason "Unable to verify Secure Boot enablement state ($($sbInfo.ErrorMessage))" -Telemetry $telemetry
}

if (-not $uefi.UefiReadable) {
    Complete-Detection -ExitCode 1 -Marker 'NON-COMPLIANT-COLLECTION-ERROR' `
        -Reason "Unable to read UEFI Secure Boot variables (PKRead=$($uefi.PkReadOk), KEKRead=$($uefi.KekReadOk), DBRead=$($uefi.DbReadOk)); verify elevated Administrator or SYSTEM privileges" `
        -Telemetry $telemetry
}

if (-not $evt.QueryOk) {
    Complete-Detection -ExitCode 1 -Marker 'NON-COMPLIANT-COLLECTION-ERROR' `
        -Reason 'Unable to query System Microsoft-Windows-TPM-WMI event log or verify event chronology against OS boot time' `
        -Telemetry $telemetry
}

# 2. Capsule PK PnP device status check (evaluated before compliance so PnP/query errors never exit 0)
if ($vmCtx.CapsuleDeviceCode -eq 'QueryError') {
    Complete-Detection -ExitCode 1 -Marker 'NON-COMPLIANT-COLLECTION-ERROR' `
        -Reason 'Unable to query Win32_PnPEntity for VMware Platform Key Firmware capsule device status' `
        -Telemetry $telemetry
}

if ($vmCtx.CapsuleDeviceCode -eq '14') {
    Complete-Detection -ExitCode 1 -Marker 'NON-COMPLIANT-VMWARE-CAPSULE-REBOOT' `
        -Reason 'VMware Platform Key Firmware capsule device requires a guest restart to complete PK update (PnP error code 14)' `
        -Telemetry $telemetry
}

if ($vmCtx.CapsuleDeviceCode -notin @('Absent', '0', '14', 'N/A')) {
    Complete-Detection -ExitCode 1 -Marker 'NON-COMPLIANT-VMWARE-CAPSULE-ERROR' `
        -Reason "VMware Platform Key Firmware capsule device has PnP error code $($vmCtx.CapsuleDeviceCode) (not a pending restart; inspect Device Manager)" `
        -Telemetry $telemetry
}

# 3. Strict compliance check:
#    Requires live NVRAM proof (KEK 2023 + required 2023 DB certs) AND current servicing completion
#    (UEFICA2023Status='Updated', UEFICA2023Error=0 explicitly read, AvailableUpdates=0x4000,
#    event log query succeeded, Event 1808 or 1799 present with no newer blocking error and no current-boot pending reboot).
$nvramComplete     = ($uefi.Kek2023 -eq $true -and $uefi.DbWin2023 -eq $true -and $uefi.AllRequiredDb23 -eq $true)
$servicingComplete = ($status -eq 'Updated' -and $null -ne $err -and $err -eq 0 -and $avUpdates -eq $Script:CompliantAvUpdates)
$eventComplete     = ($evt.QueryOk -eq $true -and ($evt.Has1808 -or $evt.Has1799) -and -not $evt.Post1808BlockingError -and -not $evt.RebootPendingCurrent)

if ($nvramComplete -and $servicingComplete -and $eventComplete) {
    Complete-Detection -ExitCode 0 -Marker 'COMPLIANT' `
        -Reason 'VMware VM Secure Boot CA 2023 certificates verified in NVRAM and Boot Manager update complete' `
        -Telemetry $telemetry
}

if ($servicingComplete -and -not $nvramComplete) {
    Write-Log ("Stale servicing state detected: Registry reports Updated/0x4000, but live NVRAM has KEK23={0}, AllRequiredDb23={1} (Missing={2})." -f
        $uefi.Kek2023, $uefi.AllRequiredDb23, $uefi.MissingDb23) -Level WARN
}

# 4. KEK 2023 is ALREADY in NVRAM (e.g., ESXi 8.0 U2-U3i P08 VMs with VMW.NULLPK, or VMs where KEK 2023 was enrolled):
#    PK replacement is not required when KEK 2023 is already present; evaluate OS-side servicing / DB state.
if ($uefi.Kek2023 -eq $true) {
    if ($avUpdates -eq 0x4100 -or $evt.RebootPendingCurrent) {
        Complete-Detection -ExitCode 1 -Marker 'NON-COMPLIANT-PENDING-REBOOT' `
            -Reason 'KEK 2023 is present in NVRAM (PK update not required); guest reboot required before Secure Boot servicing can continue' `
            -Telemetry $telemetry
    }

    if ($nvramComplete -and $servicingComplete -and -not $eventComplete) {
        if (-not $evt.QueryOk) {
            Complete-Detection -ExitCode 1 -Marker 'NON-COMPLIANT-COLLECTION-ERROR' `
                -Reason 'KEK 2023 and required 2023 DB certs are in NVRAM with Status=Updated, but System TPM-WMI event log query failed' `
                -Telemetry $telemetry
        }
        $evtDetail = if ($evt.Post1808BlockingError) {
            "blocking event $($evt.LatestBlockingId) logged after completion event"
        } else {
            'Event 1808/1799 not found in event log'
        }
        Complete-Detection -ExitCode 1 -Marker 'NON-COMPLIANT-OS-SERVICING-NEEDED' `
            -Reason "KEK 2023 and required 2023 DB certs are in NVRAM with Status=Updated, but $evtDetail; re-verify or re-run Secure-Boot-Update task" `
            -Telemetry $telemetry
    }

    $contextNotes = New-Object System.Collections.Generic.List[string]
    if (-not $uefi.AllRequiredDb23) {
        [void]$contextNotes.Add("missing DB cert(s): $($uefi.MissingDb23)")
    } else {
        [void]$contextNotes.Add('all required 2023 DB certs present in NVRAM')
    }
    if ($servicingComplete) {
        [void]$contextNotes.Add('stale Status=Updated with incomplete NVRAM')
    }
    if ($null -ne $optOut -and [int]$optOut -eq 1) {
        [void]$contextNotes.Add('HighConfidenceOptOut=1')
    }
    if ($evt.Active1802 -and $evt.SkipReason) {
        [void]$contextNotes.Add("SkipReason=$($evt.SkipReason)")
    }
    if ($taskState -notin @('Ready', 'Running')) {
        [void]$contextNotes.Add("Task=$taskState")
    }

    Complete-Detection -ExitCode 1 -Marker 'NON-COMPLIANT-OS-SERVICING-NEEDED' `
        -Reason ("KEK 2023 is in NVRAM (PK update not required; {0}); complete Windows Secure-Boot-Update servicing (Status={1}, Avl={2})" -f
            ($contextNotes -join ', '), $status, $avHex) `
        -Telemetry $telemetry
}

# 6. KEK 2023 is MISSING from NVRAM AND PK is VMW.NULLPK (or None) -> Hypervisor-side PK remediation required (Broadcom KB 423893)
if ($uefi.Kek2023 -eq $false -and $uefi.PkState -in @('NullPK', 'None')) {
    $remedHint = if ($vmCtx.VtpmPresent -eq $true) {
        if ($vmCtx.CapsuleOsSupport -eq 'DeferredToolsRelease') {
            "vTPM=True, OSBuild=$($vmCtx.OsBuild): $($vmCtx.CapsuleOsNote); wait for upcoming VMTools/8.x automated release or use manual PK update with BitLocker/TPM suspended"
        } elseif ($vmCtx.VMTools1315Ready) {
            "vTPM=True, VMTools=$($vmCtx.VMToolsVersion) (>=13.1.5): requires ESX 9.1.1.0+, VM HW v14+ (cold power-on), and July 2026+ CU for Capsule PK update ($($vmCtx.CapsuleOsNote))"
        } else {
            "vTPM=True, VMTools=$($vmCtx.VMToolsVersion) (<13.1.5): requires ESX 9.1.1.0+, VM HW v14+, July 2026+ CU, and VMTools 13.1.5+ for Capsule PK update ($($vmCtx.CapsuleOsNote))"
        }
    } elseif ($vmCtx.VtpmPresent -eq $false) {
        'vTPM=False: requires VM HW v14+ and guest restart on ESXi 8.0 U3j (P09)+ or ESX 9.1.1.0+ to apply SilentPK (or manual PK update on older hosts)'
    } else {
        'vTPM=Unknown (TPM query failed): verify vTPM in vCenter before selecting Capsule PK (vTPM=True) vs SilentPK (vTPM=False) per Broadcom KB 423893'
    }

    Complete-Detection -ExitCode 1 -Marker 'NON-COMPLIANT-VMWARE-PK-BLOCKED' `
        -Reason "Missing KEK 2023 blocked by PK=$($uefi.PkState) (Broadcom KB 423893: $remedHint)" `
        -Telemetry $telemetry
}

# 7. KEK 2023 is MISSING from NVRAM, but 'Windows OEM Devices PK' is present -> Ready for OS-side servicing
if ($uefi.Kek2023 -eq $false -and $uefi.PkState -eq 'WinOemPK') {
    if ($evt.Has1803 -or $evt.Has1795) {
        Write-Log 'Historical Event 1803/1795 present, but Windows OEM Devices PK is now enrolled in NVRAM; VM can retry OS KEK servicing.' -Level INFO
    }
    Complete-Detection -ExitCode 1 -Marker 'NON-COMPLIANT-PK-READY-NEEDS-KEK' `
        -Reason "Windows OEM Devices PK is present in NVRAM but KEK 2023 is missing; run Windows Secure-Boot-Update servicing (Status=$status, Avl=$avHex, Task=$taskState)" `
        -Telemetry $telemetry
}

# 8. KEK 2023 is MISSING from NVRAM and PK is a custom certificate or unrecognized non-zero payload
Complete-Detection -ExitCode 1 -Marker 'NON-COMPLIANT-VMWARE-PK-UNRECOGNIZED' `
    -Reason "Missing KEK 2023 and PK state is $($uefi.PkState) (Windows OEM Devices PK not detected); verify PK per Broadcom KB 423893" `
    -Telemetry $telemetry

#endregion
