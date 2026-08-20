<#
.SYNOPSIS
    Ivanti BitLocker PCR Validation Profile Detection.

.DESCRIPTION
    Pure detection script intended for use as an Ivanti Custom Definition (or any
    Status/Reason/Expected/Found-style detect channel). DOES NOT modify the system.

    Compliance logic:

        Compliant = PCR validation profile of the TPM protector on the OS volume
                    is exactly "7, 11" or "4, 7, 11"

    Anything else is non-compliant. In particular the legacy profile "0, 2, 4, 11"
    -- which Windows falls back to when PCR 7 cannot be bound -- is treated as a
    finding, because Secure Boot servicing on those devices has been observed to
    end in a recovery prompt.

    The profile is read via Win32_EncryptableVolume WMI
    (GetKeyProtectorPlatformValidationProfile) where that works. On many TPM 2.0 /
    Secure Boot integrity validation devices that method returns E_INVALIDARG
    (0x80070057), so the script falls back to parsing manage-bde. The fallback
    anchors on the literal token "PCR" and then on a line of comma-separated
    integers, both of which survive localisation. The "found =" line reports which
    source produced the answer.

    Additional diagnostics surfaced in the "found =" line (informational only, they
    do NOT change the compliance verdict):
      - Protection and conversion status of the volume
      - Whether Secure Boot is on
      - Count of BitLocker-Driver event 24604 ("boot configuration options did not
        match expected values") and 24636 ("bootmgr failed to obtain the volume
        master key") in the System log

    Output contract (one Write-Host per line, exactly these keys):
        detected = true|false
        reason   = <single sentence>
        expected = Profile: 7,11 or 4,7,11
        found    = Profile: <p> | Source: <s> | Protection: <s> | Conversion: <c> | ...

    Diagnostic log (append-only, never written to stdout):
        C:\Windows\Temp\BitLockerPcrDetection-Ivanti.log

.PARAMETER MountPoint
    Volume to inspect. Defaults to the OS volume.

.NOTES
    Author:  Anton Romanyuk
    Version: 1.1
    Date:    2026-08-20
    Requires PowerShell 5.1. Must run elevated -- the MicrosoftVolumeEncryption
    namespace denies key-protector enumeration to standard users.
#>

[CmdletBinding()]
param (
    [string] $MountPoint = $env:SystemDrive
)

# -------------------------------------------------------------------------------------------------
# 0. Logging (append-only, file-based; never writes to stdout so the Ivanti contract stays clean)
# -------------------------------------------------------------------------------------------------
$LogPath = Join-Path $env:windir 'Temp\BitLockerPcrDetection-Ivanti.log'

function Write-DetectLog {
    param(
        [Parameter(Mandatory)] [string] $Message,
        [ValidateSet('INFO','WARN','ERROR','DEBUG')] [string] $Level = 'INFO'
    )
    $line = '[{0}] [{1}] [PID:{2}] {3}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'), $Level, $PID, $Message
    try {
        Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8 -ErrorAction Stop
    } catch {
        # If C:\Windows\Temp is unwritable (rare; non-admin), drop the line silently.
    }
}

Write-DetectLog '==================================================================='
Write-DetectLog ("Run start | User=$env:USERNAME | Computer=$env:COMPUTERNAME | OS={0} | PSVersion={1}" -f `
    ((Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue).Caption), $PSVersionTable.PSVersion)

# -------------------------------------------------------------------------------------------------
# 1. Setup
# -------------------------------------------------------------------------------------------------
$FveNamespace = 'Root\CIMv2\Security\MicrosoftVolumeEncryption'
$BlProvider   = 'Microsoft-Windows-BitLocker-Driver'

# Accepted profiles. Order-insensitive: the measured profile is sorted before comparison.
$GoodProfiles = @('7,11', '4,7,11')

# TPM-backed key protector types; only these expose a platform validation profile.
$TpmProtectorTypes = @(3, 4, 5, 6)

$ProtectionStatusText = @{ 0 = 'Off'; 1 = 'On'; 2 = 'Unknown' }
$ConversionStatusText = @{
    0 = 'FullyDecrypted'; 1 = 'FullyEncrypted'; 2 = 'EncryptionInProgress'
    3 = 'DecryptionInProgress'; 4 = 'EncryptionPaused'; 5 = 'DecryptionPaused'
}
$ProtectorTypeText = @{
    3 = 'TPM'; 4 = 'TPM+PIN'; 5 = 'TPM+StartupKey'; 6 = 'TPM+PIN+StartupKey'
}

$Drive = $MountPoint.TrimEnd('\')
if ($Drive -notmatch ':$') { $Drive = $Drive + ':' }
Write-DetectLog "Target volume: $Drive"

# -------------------------------------------------------------------------------------------------
# 2. Volume state
# -------------------------------------------------------------------------------------------------
$ProfileString   = 'N/A'
$ProfileSource   = 'None'
$ProtectionText  = 'N/A'
$ConversionText  = 'N/A'
$ProtectorText   = 'None'
$VolumeFound     = $false
$TpmProtectorFound = $false
$ReadError       = $null

# GetKeyProtectorPlatformValidationProfile returns E_INVALIDARG (0x80070057) on many TPM 2.0 /
# Secure Boot integrity validation devices, so manage-bde is the fallback source.
function Get-PcrProfileFromManageBde {
    param([Parameter(Mandatory)] [string] $Volume)

    $exe = Join-Path $env:windir 'System32\manage-bde.exe'
    if (-not (Test-Path -LiteralPath $exe)) {
        Write-DetectLog 'manage-bde.exe not found' 'WARN'
        return $null
    }

    $lines = @(& $exe -protectors -get $Volume 2>&1 | ForEach-Object { "$_" })
    if ($lines.Count -eq 0) {
        Write-DetectLog 'manage-bde returned no output' 'WARN'
        return $null
    }

    # Locale-independent: anchor on the literal token "PCR" (retained in every localised build),
    # then take the first line that is nothing but comma-separated integers.
    $start = 0
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match 'PCR') { $start = $i + 1; break }
    }

    for ($i = $start; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match '^\s*\d{1,2}(\s*,\s*\d{1,2})*\s*$') {
            $set = $lines[$i] -split ',' | ForEach-Object { [int]$_.Trim() } | Sort-Object -Unique
            return ($set -join ',')
        }
    }

    Write-DetectLog 'manage-bde output contained no PCR profile line' 'WARN'
    return $null
}

try {
    $vol = Get-WmiObject -Namespace $FveNamespace -Class Win32_EncryptableVolume `
                         -Filter "DriveLetter='$Drive'" -ErrorAction Stop
} catch {
    $vol = $null
    $ReadError = $_.Exception.Message
    Write-DetectLog ("Win32_EncryptableVolume query failed: $ReadError") 'ERROR'
}

if ($vol) {
    $VolumeFound = $true

    $ps = $vol.GetProtectionStatus()
    if ($ps.ReturnValue -eq 0 -and $ProtectionStatusText.ContainsKey([int]$ps.ProtectionStatus)) {
        $ProtectionText = $ProtectionStatusText[[int]$ps.ProtectionStatus]
    }

    $cs = $vol.GetConversionStatus()
    if ($cs.ReturnValue -eq 0 -and $ConversionStatusText.ContainsKey([int]$cs.ConversionStatus)) {
        $ConversionText = $ConversionStatusText[[int]$cs.ConversionStatus]
    }
    Write-DetectLog "Volume state: Protection=$ProtectionText | Conversion=$ConversionText"

    $kp = $vol.GetKeyProtectors(0)
    if ($kp.ReturnValue -ne 0) {
        $ReadError = ('GetKeyProtectors returned 0x{0:X8}' -f $kp.ReturnValue)
        Write-DetectLog $ReadError 'ERROR'
    } else {
        foreach ($id in @($kp.VolumeKeyProtectorID)) {
            $t = $vol.GetKeyProtectorType($id)
            if ($t.ReturnValue -ne 0) { continue }

            $type = [int]$t.KeyProtectorType
            if ($TpmProtectorTypes -notcontains $type) { continue }

            $TpmProtectorFound = $true
            $ProtectorText = $ProtectorTypeText[$type]

            $pv = $vol.GetKeyProtectorPlatformValidationProfile($id)
            if ($pv.ReturnValue -ne 0) {
                $ReadError = ('GetKeyProtectorPlatformValidationProfile returned 0x{0:X8}' -f $pv.ReturnValue)
                Write-DetectLog $ReadError 'WARN'
                break
            }

            $pcrs = @($pv.PlatformValidationProfile) | ForEach-Object { [int]$_ } | Sort-Object -Unique
            if ($pcrs.Count -gt 0) {
                $ProfileString = ($pcrs -join ',')
                $ProfileSource = 'WMI'
            }
            Write-DetectLog "Protector $id ($ProtectorText): profile = $ProfileString"
            break
        }
    }
} else {
    Write-DetectLog "No encryptable volume object for $Drive" 'WARN'
}

if ($ProfileString -eq 'N/A' -and $VolumeFound) {
    Write-DetectLog 'Falling back to manage-bde for the PCR validation profile'
    $fallback = Get-PcrProfileFromManageBde -Volume $Drive
    if ($fallback) {
        $ProfileString = $fallback
        $ProfileSource = 'manage-bde'
        $TpmProtectorFound = $true
        $ReadError = $null
        Write-DetectLog "manage-bde profile = $ProfileString"
    }
}

# -------------------------------------------------------------------------------------------------
# 3. Informational diagnostics (do NOT gate compliance on these)
# -------------------------------------------------------------------------------------------------
$SecureBootText = 'N/A'
try {
    $SecureBootText = "$(Confirm-SecureBootUEFI -ErrorAction Stop)"
} catch [System.UnauthorizedAccessException] {
    $SecureBootText = 'AccessDenied'
} catch {
    # Legacy BIOS, or the cmdlet is unavailable on this SKU.
    $SecureBootText = 'Unsupported'
}

$Evt24604Count = 0
$Evt24636Count = 0
try {
    $blEvents = @(Get-WinEvent -FilterHashtable @{
        LogName      = 'System'
        ProviderName = $BlProvider
        Id           = @(24604, 24636)
    } -MaxEvents 50 -ErrorAction SilentlyContinue)

    $Evt24604Count = @($blEvents | Where-Object { $_.Id -eq 24604 }).Count
    $Evt24636Count = @($blEvents | Where-Object { $_.Id -eq 24636 }).Count
} catch {
    Write-DetectLog ('BitLocker event sweep failed: ' + $_.Exception.Message) 'WARN'
}
Write-DetectLog "Diagnostics: SecureBoot=$SecureBootText | 24604=$Evt24604Count | 24636=$Evt24636Count"

# -------------------------------------------------------------------------------------------------
# 4. Output strings (Ivanti contract)
# -------------------------------------------------------------------------------------------------
$ExpectedString = 'Profile: ' + ($GoodProfiles -join ' or ')

$FoundParts = @(
    "Profile: $ProfileString"
    "Source: $ProfileSource"
    "Protector: $ProtectorText"
    "Protection: $ProtectionText"
    "Conversion: $ConversionText"
    "SecureBoot: $SecureBootText"
)
if ($Evt24604Count -gt 0) { $FoundParts += "Event24604: $Evt24604Count" }
if ($Evt24636Count -gt 0) { $FoundParts += "Event24636: $Evt24636Count" }
if ($ReadError)           { $FoundParts += "ReadError: $ReadError" }
$FoundString = $FoundParts -join ' | '

# -------------------------------------------------------------------------------------------------
# 5. Compliance verdict
# -------------------------------------------------------------------------------------------------
$IsCompliant = ($GoodProfiles -contains $ProfileString)

Write-DetectLog ("Expected: $ExpectedString")
Write-DetectLog ("Found:    $FoundString")
Write-DetectLog ("Compliance verdict: IsCompliant=$IsCompliant")

if (-not $IsCompliant) {
    $DetectedString = 'true'

    # Most specific signal first.
    if (-not $VolumeFound) {
        $ReasonString = "No encryptable volume found for $Drive. Run elevated and confirm the drive letter."
    }
    elseif ($ProtectionText -eq 'Off' -or $ConversionText -eq 'FullyDecrypted') {
        $ReasonString = "BitLocker is not protecting $Drive (Protection: $ProtectionText, Conversion: $ConversionText). No PCR profile to evaluate."
    }
    elseif (-not $TpmProtectorFound) {
        $ReasonString = "$Drive has no TPM key protector, so it is not bound to any PCR profile."
    }
    elseif ($ReadError) {
        $ReasonString = "PCR validation profile could not be read by WMI or manage-bde. $ReadError"
    }
    elseif ($ProfileString -eq '0,2,4,11') {
        $ReasonString = "Legacy PCR profile 0,2,4,11 in use -- PCR 7 could not be bound. Secure Boot servicing on this device risks a recovery prompt."
    }
    else {
        $ReasonString = "PCR profile is '$ProfileString' (expected 7,11 or 4,7,11)."
    }
}
else {
    $DetectedString = 'false'
    $ReasonString   = "PCR validation profile is '$ProfileString' (source: $ProfileSource)."
}

Write-DetectLog ("detected = $DetectedString")
Write-DetectLog ("reason   = $ReasonString")
Write-DetectLog 'Run end'

# Ivanti contract -- exactly four lines on stdout, in this order.
Write-Host "detected = $DetectedString"
Write-Host "reason = $ReasonString"
Write-Host "expected = $ExpectedString"
Write-Host "found = $FoundString"
