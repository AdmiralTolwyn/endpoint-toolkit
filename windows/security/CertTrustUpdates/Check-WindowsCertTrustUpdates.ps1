#Requires -Version 5.1
<#
.SYNOPSIS
    Detects whether a Windows device is vulnerable to the October 2026 eMudhra /
    AAA Certificate Services CA compromise (missing CTL updates, OOB KB, or locally trusted certs).

.DESCRIPTION
    Read-only detection script evaluating Microsoft's October 2026 certificate
    trust mitigations (MSRC 2026-10-26 / Microsoft Trusted Root Program).

    STRICTLY READ-ONLY. No registry values, certificate stores, or system files
    are modified. Diagnostic log entries are the only writes.

    Decision Tree:
      1. OS build / UBR and October 26, 2026 OOB / cumulative update check:
         - Evaluates both CurrentBuildNumber.UBR (>= minimum revision) and
           Get-HotFix for the October 26, 2026 out-of-band updates and subsequent
           cumulative updates across Windows 11, Windows 10, and Windows Server
           (KB5134134, KB5134133, KB5134130, KB5134137, KB5134140, KB5134125,
           KB5134124, KB5134123, KB5127936, KB5127653).
      2. Automatic CTL update policy check:
         - Reads HKLM:\SOFTWARE\Policies\Microsoft\SystemCertificates\AuthRoot
           (DisableRootAutoUpdate = 1 or EnableDisallowedCertAutoUpdate = 0)
           to identify policy-blocked CTL synchronization.
      3. Local Disallowed CTL verification (DisallowedCertEncodedCtl):
         - Decodes the PKCS#7 CTL in memory and verifies EffectiveUtc >=
           2026-10-19 01:53:30Z (SequenceNumber 01dd5f6ca45c9d86), which blocks
           the 8 compromised eMudhra G3B/G3C intermediate CA keys by SPKI.
         - Satisfied by either the synced Disallowed CTL or the Oct 26+ OOB/CU KB.
      4. Local AuthRoot CTL verification (EncodedCtl):
         - Decodes the PKCS#7 AuthRoot CTL in memory and verifies none of the
           7 removed root SHA-256 thumbprints remain in the active list
           (removed in SequenceNumber 1401dd6344d40f0a15 / 2026-10-23 23:18:35Z).
         - Satisfied by either the synced AuthRoot CTL or the Oct 26+ OOB/CU KB.
      5. Local certificate store audit (LocalMachine\Root, AuthRoot, TrustedPublisher, CA):
         - Computes SHA-256 over RawData for every certificate (because PowerShell's
           X509Certificate2.Thumbprint property is SHA-1, whereas MSRC publishes
           SHA-256 fingerprints) and checks whether any of the 7 removed roots
           or 16 revoked intermediates are still cached/trusted locally and absent
           from LocalMachine\Disallowed.

    Logs are written in CMTrace format to:
      %ProgramData%\Microsoft\IntuneManagementExtension\Logs\PR_WindowsCertTrustUpdates.log
      (falls back to %TEMP%\PR_WindowsCertTrustUpdates.log when non-elevated).

.PARAMETER TestOnline
    Optional. Performs an HTTP HEAD probe against ctldl.windowsupdate.com to
    verify outbound port 80 CTL reachability during interactive troubleshooting.
    Omitted by default so fleet executions remain 100% local and fast.

.PARAMETER LogPath
    Optional override for the CMTrace log file path.

.INPUTS
    None. Pipeline input is not supported.

.OUTPUTS
    System.Int32 exit code:
      0 = NOT-VULNERABLE
      1 = VULNERABLE (missing CTL/OOB mitigations, locally trusted removed cert, or collection error)

    Single-line STDOUT summary (< 2048 chars) with status prefix and key=value telemetry:
      NOT-VULNERABLE: <reason> | <telemetry>
      VULNERABLE: <reason> | <telemetry>

.EXAMPLE
    .\Check-WindowsCertTrustUpdates.ps1

    Evaluates local CTLs, OS build/UBR + hotfixes, and machine certificate stores
    without network calls or temporary files.

.EXAMPLE
    .\Check-WindowsCertTrustUpdates.ps1 -TestOnline

    Runs detection and additionally tests HTTP reachability of ctldl.windowsupdate.com.

.NOTES
    Author:   Anton Romanyuk
    Version:  2.2
    Date:     2026-11-04
    Context:  Windows Certificate Trust (eMudhra / Sectigo AAA Certificate Services) Detection
    Requires: Windows PowerShell 5.1, 64-bit recommended

    References:
      https://www.microsoft.com/en-us/msrc/blog/2026/10/microsoft-response-to-compromised-emudhra-certificate-authority-keys
      https://github.com/TrustedRootProgram/Program-Requirements/blob/main/Announcements.md#october-26-2026
      https://support.microsoft.com/en-us/servicing/os/windows-11/2026/10/kb5134133-windows-11-24h2-25h2-26h2-update

    DISCLAIMER:
    THIS SCRIPT IS PROVIDED "AS-IS" WITHOUT WARRANTY OF ANY KIND.
    USE AT YOUR OWN RISK.
#>

[CmdletBinding()]
param(
    [switch]$TestOnline,
    [string]$LogPath
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'

# Required for System.Security.Cryptography.Pkcs.SignedCms in Windows PowerShell 5.1
Add-Type -AssemblyName System.Security

# ─────────────────────────────────────────────────────────────────────────────
# Configuration
# ─────────────────────────────────────────────────────────────────────────────

$Script:LogDir      = Join-Path $env:ProgramData 'Microsoft\IntuneManagementExtension\Logs'
$Script:LogName     = 'PR_WindowsCertTrustUpdates'
$Script:LogFile     = if ($LogPath) { $LogPath } else { Join-Path $Script:LogDir "$Script:LogName.log" }
$Script:LogMaxSize  = 250KB
$Script:Component   = 'CheckCertTrust'
$Script:ScriptName  = if ($MyInvocation.MyCommand.Name) { $MyInvocation.MyCommand.Name } else { 'Check-WindowsCertTrustUpdates.ps1' }
$Script:LogWarned   = $false

$Script:RegPathCurrentVersion = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
$Script:RegPathAuthRootPolicy = 'HKLM:\SOFTWARE\Policies\Microsoft\SystemCertificates\AuthRoot'
$Script:RegPathAutoUpdate     = 'HKLM:\SOFTWARE\Microsoft\SystemCertificates\AuthRoot\AutoUpdate'
$Script:CtlCheckUrl           = 'http://ctldl.windowsupdate.com/msdownload/update/v3/static/trustedr/en/disallowedcertstl.cab'

# Minimum Disallowed CTL effective timestamp for the October 20, 2026 CTL release:
# SequenceNumber = 01dd5f6ca45c9d86 (FILETIME 0x01dd5f6ca45c9d86 = 2026-10-19 01:53:30 UTC)
$Script:MinDisallowedUtc = [DateTime]::SpecifyKind([DateTime]'2026-10-19T01:53:30', [DateTimeKind]::Utc)

# Known October 26, 2026 OOB KBs and October 28, 2026 Preview KBs carrying the updated root/CTL payload
$Script:KnownMitigationKBs = @(
    'KB5134134' # Windows 11 26H1 OOB (28000.3237)
    'KB5127653' # Windows 11 26H1 Oct 28 Preview (28000.3365)
    'KB5134133' # Windows 11 24H2 / 25H2 / 26H2 OOB (26100/26200/26300.9725)
    'KB5127936' # Windows 11 25H2 / 26H2 Oct 28 Preview (26200/26300.9818)
    'KB5134130' # Windows 11 23H2 OOB (22621/22631.7654)
    'KB5134137' # Windows Server 2025 OOB (26100.33674)
    'KB5134140' # Windows Server 2022 OOB (20348.5773)
    'KB5134125' # Windows 10 21H2 / 22H2 / LTSC 2021 OOB (19044/19045.7811)
    'KB5134124' # Windows 10 1809 LTSC 2019 / Windows Server 2019 OOB (17763.9378)
    'KB5134123' # Windows 10 1607 LTSB 2016 / Windows Server 2016 OOB (14393.9615)
)

# Seven roots removed from default Windows trust on October 26, 2026 (SHA-256 fingerprint -> Display Name)
# Source: Microsoft Trusted Root Program Announcement (2026-10-26)
$Script:RemovedRootCerts = @{
    '86A1ECBA089C4A8D3BBE2734C612BA341D813E043CF9E8A862CD5C57A36BBE6B' = 'emSign ECC Root CA - G3'
    'BBD8B3D745EFB79198D93B9C482B9D934537325729EA2F02C129261F91BE2E3D' = 'emSign Root Client Auth CA - G3'
    'F2EDAAFF3D9B701FB4D14C9F285F8C0BAD90FC3BCC1FA07D47F4718EDABA34A5' = 'emSign Root CS CA - G3'
    'F76A0D440C418DC5B362BE8E11E0A5D06023F4396093E4F455827BF55153D8A4' = 'emSign Root SMIME CA - G3'
    '7DD78D5F4F13459A83DFF9ABBBA62EDBAF6F2D102BF257FD712F4D9F2746ED8D' = 'emSign Root TLS CA - G3'
    '75D2CE75979CF2A43B172C9F89F733DDA2136031DB4BBF3698B053714C077155' = 'emSign Root TSA CA - G3'
    'D7A7A0FB5D7E2731D771E9484EBCDEF71D5F0C3E0A2948782BC83EE0EA699EF4' = 'AAA Certificate Services (Sectigo)'
}

# Sixteen revoked eMudhra G3B/G3C intermediate CA certificates (8 direct emSign + 8 Sectigo AAA cross-signs)
# Note: The 8 AAA cross-signs have Issuer='AAA Certificate Services' and Subject='CN=EM ...', so subject-only
# filtering for 'AAA Certificate Services' would miss them; matching by SHA-256 fingerprint is exact.
$Script:RevokedIntermediateCerts = @{
    '42C4ACCD12C4F59597925E1A08A7A90102608543809308D50EC427E118BA1493' = 'EM DV TLS CA - G3B (emSign)'
    '45A883F349AEDFD5F3A649B5498FFB121712BA8B448F3F2A91DBC5A766FCDDBE' = 'EM DV TLS CA - G3B (AAA cross-sign)'
    '0164D7CEAE5BB909C272AF151FF44B1177AE9B47D9424B94F1D9558C30934D19' = 'EM DV TLS CA - G3C (emSign)'
    '62B7F9868263F72F6911FC78CEAF72C1979158B7368BF648683BB2AA39B02E72' = 'EM DV TLS CA - G3C (AAA cross-sign)'
    'E7568FE6E12331F4FD7DE61DA4C7559D0F8EB192A218D4741A3BA75184333838' = 'EM EV TLS CA - G3B (emSign)'
    '579480AEA166290BD3C643589E795FC0D0E54E86505D180F1DFF99EFD3199512' = 'EM EV TLS CA - G3B (AAA cross-sign)'
    '432EC25F0D28DE9006D97BF752F17BBB2CD41F669933FCB072D9A1148D0261F3' = 'EM EV TLS CA - G3C (emSign)'
    'FD9D5C2DBBD4995AB69B8D50C439D9527B00529199C2FD01B2174EB58F07021F' = 'EM EV TLS CA - G3C (AAA cross-sign)'
    '98B896625C0C230ADEEC3EB2CF6AAD5001ECF9156A68644E82374CE5833BC25B' = 'EM OV TLS CA - G3B (emSign)'
    '5F1DED4DD590AF74631290CF073B122A2AAC6A1FA479778A0BD43579BD832699' = 'EM OV TLS CA - G3B (AAA cross-sign)'
    'CB6E15165AC802E9B3F993F16C25D41F4E81C21EBDE71AE1B012AC00BF9C7C3F' = 'EM OV TLS CA - G3C (emSign)'
    '7485B51C0ED192878FA1D2F948FB529AE1842CC5086C4B67F7BE5B2039DE9489' = 'EM OV TLS CA - G3C (AAA cross-sign)'
    '9E39DE2881273ACC225A149C72BA3E0F51424963669AB633EE1DC7DF5DB067B4' = 'EM OV CS CA - G3B (emSign)'
    'A496E23AC0CAB8155FC9B117B97EAE368EFCBB9041AEC8804D5A308D2D01C2D8' = 'EM OV CS CA - G3B (AAA cross-sign)'
    '4B2460A466F95FD40113BCDCCA2CEB29B3B093921DE45474DA7DF48874C630B1' = 'EM OV CS CA - G3C (emSign)'
    '84112972B8E59D14560AE7E7119654474AAD119990CF3722BFC1E3B91615E1FD' = 'EM OV CS CA - G3C (AAA cross-sign)'
}

# ─────────────────────────────────────────────────────────────────────────────
#region Logging
# ─────────────────────────────────────────────────────────────────────────────

function Write-Log {
    <#
    .SYNOPSIS
        Writes a CMTrace-formatted log entry to the IME log directory (with %TEMP% fallback).
    .DESCRIPTION
        Formats messages using the standard Configuration Manager / Intune CMTrace XML-like
        tag structure, resolves the calling function from the PowerShell call stack, and
        rotates the log file when it reaches $Script:LogMaxSize. Logging is best-effort and
        never throws or alters the detection verdict.
    .PARAMETER Message
        Diagnostic text to record.
    .PARAMETER Level
        Severity level mapped to CMTrace type (1 = Info/Step/Success, 2 = Warn, 3 = Error).
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS', 'STEP')]
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

    $dir = Split-Path -Path $Script:LogFile -Parent
    if (-not (Test-Path -LiteralPath $dir)) {
        try {
            New-Item -Path $dir -ItemType Directory -Force -ErrorAction Stop | Out-Null
        } catch {
            $Script:LogFile = Join-Path $env:TEMP "$Script:LogName.log"
        }
    }

    try {
        [System.IO.File]::AppendAllText($Script:LogFile, $entry + [Environment]::NewLine, [System.Text.Encoding]::UTF8)
    } catch {
        if (-not $Script:LogWarned) {
            $Script:LogWarned = $true
            try {
                $Script:LogFile = Join-Path $env:TEMP "$Script:LogName.log"
                [System.IO.File]::AppendAllText($Script:LogFile, $entry + [Environment]::NewLine, [System.Text.Encoding]::UTF8)
            } catch {
                Write-Verbose "Log write failed: $($_.Exception.Message)"
            }
        }
        return
    }

    # Rotate log when it reaches the size cap
    try {
        if ((Test-Path -LiteralPath $Script:LogFile) -and (Get-Item -LiteralPath $Script:LogFile -ErrorAction Stop).Length -ge $Script:LogMaxSize) {
            $archive = Join-Path (Split-Path -Path $Script:LogFile -Parent) "$Script:LogName-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
            Copy-Item -LiteralPath $Script:LogFile -Destination $archive -Force -ErrorAction Stop
            Remove-Item -LiteralPath $Script:LogFile -Force -ErrorAction Stop
            Compact /C $archive 2>$null | Out-Null
        }
    } catch { }
}

function Format-ReasonText {
    <#
    .SYNOPSIS
        Sanitizes free-text reason strings for single-line summary output.
    .PARAMETER Value
        Input reason string to normalize.
    .OUTPUTS
        System.String with line breaks and pipe characters collapsed, or $null if blank.
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
    .PARAMETER Value
        Input telemetry value to normalize.
    .OUTPUTS
        System.String suitable for a key=value token, or $null if blank.
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
        Logs the final verdict, emits the single-line STDOUT contract, and exits.
    .DESCRIPTION
        Constructs a bounded (< 2000 character) single-line output string prefixed with
        VULNERABLE or NOT-VULNERABLE, records the verdict to the CMTrace log, writes
        to STDOUT for the management agent, and exits with code 1 or 0.
    .PARAMETER Vulnerable
        $true when the device is vulnerable (exits 1); $false when protected (exits 0).
    .PARAMETER Reason
        Human-readable explanation of the verdict.
    .PARAMETER Telemetry
        Ordered dictionary of key/value diagnostic fields appended after the pipe delimiter.
    #>
    param(
        [Parameter(Mandatory = $true)] [bool]$Vulnerable,
        [Parameter(Mandatory = $true)] [string]$Reason,
        [Parameter(Mandatory = $true)] [System.Collections.IDictionary]$Telemetry
    )

    $status      = if ($Vulnerable) { 'VULNERABLE' } else { 'NOT-VULNERABLE' }
    $exitCode    = if ($Vulnerable) { 1 } else { 0 }
    $cleanReason = Format-ReasonText -Value $Reason

    $pairs = foreach ($k in $Telemetry.Keys) {
        $v = Format-TelemetryToken -Value $Telemetry[$k]
        if ($null -ne $v) { '{0}={1}' -f $k, $v }
    }

    $summary = '{0}: {1} | {2}' -f $status, $cleanReason, ($pairs -join ' ')
    if ($summary.Length -gt 2000) {
        $summary = $summary.Substring(0, 1997) + '...'
    }

    $level = if ($Vulnerable) { 'WARN' } else { 'SUCCESS' }
    Write-Log "Verdict -> $summary" -Level $level
    Write-Log "--- Detection finished (ExitCode: $exitCode) ---" -Level STEP
    Write-Output $summary
    exit $exitCode
}

#endregion

# ─────────────────────────────────────────────────────────────────────────────
#region Helpers
# ─────────────────────────────────────────────────────────────────────────────

function Get-RegValue {
    <#
    .SYNOPSIS
        Safely reads a single registry value, returning $null if the path or value is absent.
    .DESCRIPTION
        Preserves byte[] values as a single array object using the unary comma operator
        so large binary values (such as EncodedCtl) are not unrolled element-by-element
        across the PowerShell pipeline.
    .PARAMETER Path
        Registry key path (e.g. HKLM:\...).
    .PARAMETER Name
        Registry value name to read.
    .OUTPUTS
        The registry value, or $null if absent or unreadable.
    #>
    param(
        [Parameter(Mandatory = $true)] [string]$Path,
        [Parameter(Mandatory = $true)] [string]$Name
    )
    try {
        if (-not (Test-Path -LiteralPath $Path)) { return $null }
        $item = Get-ItemProperty -LiteralPath $Path -Name $Name -ErrorAction Stop
        $val  = $item.$Name
        if ($val -is [byte[]]) { return ,$val }
        return $val
    } catch {
        return $null
    }
}

function ConvertFrom-FileTimeBytes {
    <#
    .SYNOPSIS
        Converts an 8-byte Windows FILETIME REG_BINARY value into a UTC DateTime.
    .PARAMETER Bytes
        Raw 8-byte little-endian FILETIME array from AutoUpdate registry values.
    .OUTPUTS
        System.DateTime in UTC, or $null if absent/invalid.
    #>
    param([AllowNull()] [byte[]]$Bytes)
    if ($Bytes -is [byte[]] -and $Bytes.Length -ge 8) {
        try {
            $ft = [BitConverter]::ToInt64($Bytes, 0)
            if ($ft -gt 0) { return [DateTime]::FromFileTimeUtc($ft) }
        } catch { }
    }
    return $null
}

function Get-CtlMetadata {
    <#
    .SYNOPSIS
        Decodes a PKCS#7 Certificate Trust List (EncodedCtl / DisallowedCertEncodedCtl)
        in memory and extracts SequenceNumber, ThisUpdate (UTC), and hex content.
    .DESCRIPTION
        Uses System.Security.Cryptography.Pkcs.SignedCms to unwrap the PKCS#7 envelope
        without spawning certutil.exe or writing temporary .stl files. In a Microsoft
        CTL payload (OID 1.3.6.1.4.1.311.10.1), the ASN.1 SequenceNumber INTEGER (tag 0x02,
        length 8 or 9 bytes) immediately precedes the ThisUpdate UTCTime (tag 0x17, length
        0x0D = 13 ASCII bytes 'yyMMddHHmmssZ'). Also returns the full uppercase hex string
        of the inner CTL content so embedded SHA-256 root attributes can be checked directly.
    .PARAMETER CtlBytes
        Raw PKCS#7 CTL byte array from HKLM:\SOFTWARE\Microsoft\SystemCertificates\AuthRoot\AutoUpdate.
    .OUTPUTS
        PSCustomObject with SequenceNumber, EffectiveUtc, and ContentHex; or $null on failure.
    #>
    param([AllowNull()] [byte[]]$CtlBytes)
    if (-not $CtlBytes -or $CtlBytes.Length -lt 32) { return $null }
    try {
        $cms = New-Object System.Security.Cryptography.Pkcs.SignedCms
        $cms.Decode($CtlBytes)
        $content   = $cms.ContentInfo.Content
        $hex       = [BitConverter]::ToString($content).Replace('-', '')
        $headerHex = [BitConverter]::ToString($content[0..[Math]::Min(160, $content.Length - 1)]).Replace('-', '')

        $seqHex     = $null
        $thisUpdate = $null
        if ($headerHex -match '02(0[89])([0-9A-F]{16,18})170D([0-9A-F]{26})') {
            $seqHex = $matches[2].ToLowerInvariant()
            [byte[]]$utcBytes = for ($i = 0; $i -lt 26; $i += 2) {
                [Convert]::ToByte($matches[3].Substring($i, 2), 16)
            }
            $utcStr = [System.Text.Encoding]::ASCII.GetString($utcBytes)
            $thisUpdate = [DateTime]::ParseExact(
                $utcStr,
                'yyMMddHHmmss\Z',
                [Globalization.CultureInfo]::InvariantCulture,
                [Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal
            )
        }

        return [pscustomobject]@{
            SequenceNumber = $seqHex
            EffectiveUtc   = $thisUpdate
            ContentHex     = $hex
        }
    } catch {
        Write-Log "Failed to decode CTL bytes: $($_.Exception.Message)" -Level WARN
        return $null
    }
}

function Get-OobUpdateStatus {
    <#
    .SYNOPSIS
        Evaluates whether the OS has the October 26, 2026 OOB trusted-root update
        (or a newer cumulative update) via both UBR and Get-HotFix.
    .DESCRIPTION
        Maps the OS build and ProductType (client vs. server) to the corresponding
        October 26, 2026 out-of-band KB and minimum Update Build Revision (UBR), and
        also queries Get-HotFix for the expected KB (plus the October 28, 2026 Preview
        cumulative updates KB5127936 / KB5127653). Combining UBR and Get-HotFix ensures
        accurate detection both when the OOB KB is directly installed and when a newer
        cumulative update has superseded it.
    .PARAMETER Build
        OS build number (e.g. 28000, 26300, 26200, 26100, 22631, 20348, 19045, 17763, 14393).
    .PARAMETER Ubr
        Update Build Revision from HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\UBR.
    .PARAMETER ProductType
        Win32_OperatingSystem.ProductType (1 = Client workstation, 2/3 = Domain Controller/Server).
    .OUTPUTS
        PSCustomObject with ExpectedKb, MinUbr, UbrSatisfied, HotFixId, and Installed.
    #>
    param(
        [int]$Build,
        [int]$Ubr,
        [int]$ProductType
    )

    $expectedKb = $null
    $candidateKbs = @()
    $minUbr     = $null

    switch ($Build) {
        28000 { $expectedKb = 'KB5134134'; $candidateKbs = @('KB5134134', 'KB5127653'); $minUbr = 3237 }  # Windows 11 26H1
        26300 { $expectedKb = 'KB5134133'; $candidateKbs = @('KB5134133', 'KB5127936'); $minUbr = 9725 }  # Windows 11 26H2
        26200 { $expectedKb = 'KB5134133'; $candidateKbs = @('KB5134133', 'KB5127936'); $minUbr = 9725 }  # Windows 11 25H2
        26100 {
            if ($ProductType -eq 1) {
                $expectedKb = 'KB5134133'; $candidateKbs = @('KB5134133'); $minUbr = 9725                 # Windows 11 24H2
            } else {
                $expectedKb = 'KB5134137'; $candidateKbs = @('KB5134137'); $minUbr = 33674                # Windows Server 2025
            }
        }
        22631 { $expectedKb = 'KB5134130'; $candidateKbs = @('KB5134130'); $minUbr = 7654 }               # Windows 11 23H2
        22621 { $expectedKb = 'KB5134130'; $candidateKbs = @('KB5134130'); $minUbr = 7654 }               # Windows 11 22H2 / 23H2 base
        20348 { $expectedKb = 'KB5134140'; $candidateKbs = @('KB5134140'); $minUbr = 5773 }               # Windows Server 2022
        19045 { $expectedKb = 'KB5134125'; $candidateKbs = @('KB5134125'); $minUbr = 7811 }               # Windows 10 22H2
        19044 { $expectedKb = 'KB5134125'; $candidateKbs = @('KB5134125'); $minUbr = 7811 }               # Windows 10 21H2 / LTSC 2021
        17763 { $expectedKb = 'KB5134124'; $candidateKbs = @('KB5134124'); $minUbr = 9378 }               # Windows 10 1809 LTSC 2019 / Server 2019
        14393 { $expectedKb = 'KB5134123'; $candidateKbs = @('KB5134123'); $minUbr = 9615 }               # Windows 10 1607 LTSB 2016 / Server 2016
        default { $candidateKbs = $Script:KnownMitigationKBs }
    }

    $ubrSatisfied = ($null -ne $minUbr -and $Ubr -ge $minUbr)

    $hotFixId = $null
    if ($candidateKbs.Count -gt 0) {
        $hf = Get-HotFix -Id $candidateKbs -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($hf -and $hf.HotFixID) {
            $hotFixId = [string]$hf.HotFixID
        }
    }

    $installed = ($ubrSatisfied -or $null -ne $hotFixId)
    return [pscustomobject]@{
        ExpectedKb   = if ($expectedKb) { $expectedKb } else { 'UnknownBuild' }
        MinUbr       = $minUbr
        UbrSatisfied = $ubrSatisfied
        HotFixId     = $hotFixId
        Installed    = $installed
    }
}

#endregion

# ─────────────────────────────────────────────────────────────────────────────
#region Main
# ─────────────────────────────────────────────────────────────────────────────

Write-Log '--- Windows Certificate Trust (eMudhra / CTL) detection started ---' -Level STEP
Write-Log ("Script={0} | PSVersion={1} | PSArch={2}-bit | Host={3} | User={4} | PID={5}" -f
    $Script:ScriptName, $PSVersionTable.PSVersion, ([IntPtr]::Size * 8), $env:COMPUTERNAME, $env:USERNAME, $PID) -Level INFO

$telemetry = [ordered]@{
    Host = $env:COMPUTERNAME
}

try {
    # 1. OS Build / UBR and October 26, 2026 OOB / cumulative hotfix check
    #    Devices with the Oct 26, 2026 OOB update (or newer CU) carry the updated inbox CTL payload.
    $buildStr    = Get-RegValue -Path $Script:RegPathCurrentVersion -Name 'CurrentBuildNumber'
    $ubrVal      = Get-RegValue -Path $Script:RegPathCurrentVersion -Name 'UBR'
    $productName = Get-RegValue -Path $Script:RegPathCurrentVersion -Name 'ProductName'
    $osCim       = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
    $prodType    = if ($osCim) { [int]$osCim.ProductType } else { 1 }
    $caption     = if ($osCim -and $osCim.Caption) { $osCim.Caption } else { $productName }

    $build = if ($buildStr) { [int]$buildStr } else { 0 }
    $ubr   = if ($null -ne $ubrVal) { [int]$ubrVal } else { 0 }
    $oob   = Get-OobUpdateStatus -Build $build -Ubr $ubr -ProductType $prodType

    $telemetry['OS']     = "$build.$ubr"
    $telemetry['OobKB']  = "$($oob.ExpectedKb):$($oob.Installed)"
    $telemetry['HotFix'] = if ($oob.HotFixId) { $oob.HotFixId } else { 'None' }
    Write-Log ("OS: {0} (Build {1}.{2}, ProductType={3}) | ExpectedKB={4} (MinUBR={5}, UbrOk={6}, HotFix={7}) -> Installed={8}" -f
        $caption, $build, $ubr, $prodType, $oob.ExpectedKb, $oob.MinUbr, $oob.UbrSatisfied,
        $(if ($oob.HotFixId) { $oob.HotFixId } else { 'None' }), $oob.Installed) -Level INFO

    # 2. Automatic CTL update policy check
    #    DisableRootAutoUpdate=1 disables all automatic root/CTL syncs;
    #    EnableDisallowedCertAutoUpdate=0 specifically disables Disallowed CTL syncs.
    $disableRootAu      = Get-RegValue -Path $Script:RegPathAuthRootPolicy -Name 'DisableRootAutoUpdate'
    $enableDisallowedAu = Get-RegValue -Path $Script:RegPathAuthRootPolicy -Name 'EnableDisallowedCertAutoUpdate'

    $autoUpdateBlocked = ($disableRootAu -eq 1 -or $enableDisallowedAu -eq 0)
    $telemetry['AutoUpdateBlocked'] = $autoUpdateBlocked
    Write-Log "Policy ($Script:RegPathAuthRootPolicy): DisableRootAutoUpdate=$disableRootAu, EnableDisallowedCertAutoUpdate=$enableDisallowedAu -> AutoUpdateBlocked=$autoUpdateBlocked" -Level INFO

    # 3. Local AutoUpdate CTL state (DisallowedCertEncodedCtl + EncodedCtl)
    #    Note: LastSyncTime (AuthRoot) only updates on demand when CryptoAPI encounters an unknown root,
    #    so sync age alone is not a vulnerability indicator. Instead, inspect the actual CTL payloads.
    $rootSyncUtc  = ConvertFrom-FileTimeBytes (Get-RegValue -Path $Script:RegPathAutoUpdate -Name 'LastSyncTime')
    $disSyncUtc   = ConvertFrom-FileTimeBytes (Get-RegValue -Path $Script:RegPathAutoUpdate -Name 'DisallowedCertLastSyncTime')
    $disCtlBytes  = Get-RegValue -Path $Script:RegPathAutoUpdate -Name 'DisallowedCertEncodedCtl'
    $authCtlBytes = Get-RegValue -Path $Script:RegPathAutoUpdate -Name 'EncodedCtl'

    $disCtl  = Get-CtlMetadata -CtlBytes $disCtlBytes
    $authCtl = Get-CtlMetadata -CtlBytes $authCtlBytes

    $disSeq  = if ($disCtl -and $disCtl.SequenceNumber) { $disCtl.SequenceNumber } else { 'None' }
    $disDate = if ($disCtl -and $disCtl.EffectiveUtc) { $disCtl.EffectiveUtc.ToString('yyyy-MM-dd') } else { 'None' }
    $disCtlCurrent = ($null -ne $disCtl -and $null -ne $disCtl.EffectiveUtc -and $disCtl.EffectiveUtc -ge $Script:MinDisallowedUtc)

    $telemetry['DisallowedSeq']  = $disSeq
    $telemetry['DisallowedDate'] = $disDate
    $telemetry['DisallowedOk']   = ($disCtlCurrent -or $oob.Installed)

    Write-Log ("Disallowed CTL: Seq={0}, EffectiveUtc={1}, LastSyncUtc={2}, Current(>={3:yyyy-MM-dd HH:mm}Z)={4}" -f
        $disSeq,
        $(if ($disCtl -and $disCtl.EffectiveUtc) { $disCtl.EffectiveUtc.ToString('u') } else { 'None' }),
        $(if ($disSyncUtc) { $disSyncUtc.ToString('u') } else { 'None' }),
        $Script:MinDisallowedUtc,
        $disCtlCurrent) -Level INFO

    # Each root entry in Microsoft's AuthRoot CTL embeds its SHA-256 hash attribute
    # (CERT_AUTH_ROOT_SHA256_HASH_PROP_ID, OID 1.3.6.1.4.1.311.10.11.98) as raw 32 bytes.
    # Verify none of the 7 removed root SHA-256 hashes are present in EncodedCtl.
    $ctlRetainedRoots = New-Object System.Collections.Generic.List[string]
    if ($authCtl -and $authCtl.ContentHex) {
        foreach ($hash in $Script:RemovedRootCerts.Keys) {
            if ($authCtl.ContentHex.Contains($hash)) {
                [void]$ctlRetainedRoots.Add("$($Script:RemovedRootCerts[$hash]) ($($hash.Substring(0, 12))...)")
            }
        }
    }

    $authDate = if ($authCtl -and $authCtl.EffectiveUtc) { $authCtl.EffectiveUtc.ToString('yyyy-MM-dd') } else { 'None' }
    $authRootCtlOk = if ($null -ne $authCtl) { $ctlRetainedRoots.Count -eq 0 } else { $oob.Installed }

    $telemetry['AuthRootDate']     = $authDate
    $telemetry['CtlRetainedRoots'] = $ctlRetainedRoots.Count
    Write-Log ("AuthRoot CTL: Seq={0}, EffectiveUtc={1}, LastSyncUtc={2}, RetainedRemovedRoots={3}" -f
        $(if ($authCtl -and $authCtl.SequenceNumber) { $authCtl.SequenceNumber } else { 'None' }),
        $(if ($authCtl -and $authCtl.EffectiveUtc) { $authCtl.EffectiveUtc.ToString('u') } else { 'None' }),
        $(if ($rootSyncUtc) { $rootSyncUtc.ToString('u') } else { 'None' }),
        $ctlRetainedRoots.Count) -Level INFO

    if ($ctlRetainedRoots.Count -gt 0) {
        Write-Log "Stale AuthRoot CTL still contains removed root(s): $($ctlRetainedRoots -join '; ')" -Level WARN
    }

    # 4. Local Machine certificate store inspection (exact SHA-256 matching)
    #    Even with updated CTLs, a removed root previously cached in LocalMachine\AuthRoot or
    #    deployed via GPO/Intune/manual import into LocalMachine\Root or TrustedPublisher remains
    #    trusted unless removed (or cleaned by installing the Oct 26 OOB update) or placed in Disallowed.
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    $disallowedHashes = @{}
    foreach ($c in @(Get-ChildItem -LiteralPath 'Cert:\LocalMachine\Disallowed' -ErrorAction SilentlyContinue)) {
        $h = [BitConverter]::ToString($sha256.ComputeHash($c.RawData)).Replace('-', '')
        $disallowedHashes[$h] = $true
    }

    $activeBadCerts  = New-Object System.Collections.Generic.List[string]
    $cachedRevokedCa = New-Object System.Collections.Generic.List[string]

    foreach ($store in 'Root', 'AuthRoot', 'TrustedPublisher', 'CA') {
        $certs = @(Get-ChildItem -LiteralPath "Cert:\LocalMachine\$store" -ErrorAction SilentlyContinue)
        foreach ($c in $certs) {
            $h256 = [BitConverter]::ToString($sha256.ComputeHash($c.RawData)).Replace('-', '')
            if ($disallowedHashes.ContainsKey($h256)) { continue }

            if ($Script:RemovedRootCerts.ContainsKey($h256)) {
                $entry = "LocalMachine\$store`: $($Script:RemovedRootCerts[$h256]) [SHA256:$($h256.Substring(0, 12))]"
                [void]$activeBadCerts.Add($entry)
                Write-Log "Found removed root in active trust store -> $entry (SHA1=$($c.Thumbprint), SHA256=$h256)" -Level WARN
            } elseif ($Script:RevokedIntermediateCerts.ContainsKey($h256)) {
                $entry = "LocalMachine\$store`: $($Script:RevokedIntermediateCerts[$h256]) [SHA256:$($h256.Substring(0, 12))]"
                if ($store -eq 'CA') {
                    # Intermediate CAs cached in LocalMachine\CA are blocked by SPKI when Disallowed CTL is current,
                    # so record them as warning telemetry unless Disallowed CTL is missing.
                    [void]$cachedRevokedCa.Add($entry)
                    Write-Log "Found revoked intermediate cached in LocalMachine\CA -> $entry (SHA1=$($c.Thumbprint), SHA256=$h256)" -Level WARN
                } else {
                    [void]$activeBadCerts.Add($entry)
                    Write-Log "Found revoked intermediate in direct trust store -> $entry (SHA1=$($c.Thumbprint), SHA256=$h256)" -Level WARN
                }
            }
        }
    }
    $sha256.Dispose()

    $telemetry['TrustedRemovedCerts'] = $activeBadCerts.Count
    $telemetry['CachedRevokedCA']     = $cachedRevokedCa.Count

    # 5. Optional online reachability probe (-TestOnline)
    if ($TestOnline) {
        try {
            $resp = Invoke-WebRequest -Uri $Script:CtlCheckUrl -Method Head -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
            $telemetry['OnlineReachable'] = $true
            Write-Log "Online CTL check succeeded: HTTP $($resp.StatusCode) Last-Modified=$($resp.Headers['Last-Modified'])" -Level INFO
        } catch {
            $telemetry['OnlineReachable'] = $false
            Write-Log "Online CTL check failed ($Script:CtlCheckUrl): $($_.Exception.Message)" -Level WARN
        }
    }

    # 6. Verdict evaluation
    $reasons = New-Object System.Collections.Generic.List[string]

    if (-not $disCtlCurrent -and -not $oob.Installed) {
        if ($autoUpdateBlocked) {
            [void]$reasons.Add("Disallowed CTL is outdated/missing (Seq=$disSeq) and automatic CTL updates are disabled by policy; install $($oob.ExpectedKb)")
        } else {
            [void]$reasons.Add("Disallowed CTL is outdated/missing (Seq=$disSeq, need >=01dd5f6ca45c9d86) and OOB update $($oob.ExpectedKb) is not installed")
        }
    }

    if (-not $authRootCtlOk) {
        if ($ctlRetainedRoots.Count -gt 0) {
            [void]$reasons.Add("Local AuthRoot CTL still contains $($ctlRetainedRoots.Count) removed eMudhra/AAA root(s)")
        } else {
            [void]$reasons.Add("AuthRoot CTL is not synced and OOB update $($oob.ExpectedKb) is not installed")
        }
    }

    if ($activeBadCerts.Count -gt 0) {
        [void]$reasons.Add("Local certificate store still trusts $($activeBadCerts.Count) removed/revoked certificate(s): $($activeBadCerts -join '; ')")
    }

    if ($reasons.Count -gt 0) {
        Complete-Detection -Vulnerable $true -Reason ($reasons -join '; ') -Telemetry $telemetry
    }

    $okSource = if ($disCtlCurrent -and $oob.Installed) {
        "Disallowed/AuthRoot CTLs and OOB update $($oob.ExpectedKb) (Build $build.$ubr) applied; no removed roots trusted locally"
    } elseif ($disCtlCurrent) {
        "October 2026 Disallowed CTL (Seq=$disSeq, Date=$disDate) and AuthRoot CTL synced; no removed roots trusted locally"
    } else {
        "October 2026 OOB update $($oob.ExpectedKb) (Build $build.$ubr) installed; no removed roots trusted locally"
    }

    Complete-Detection -Vulnerable $false -Reason $okSource -Telemetry $telemetry
}
catch {
    $telemetry['Error'] = $_.Exception.Message
    Write-Log "Unhandled error during detection: $($_.Exception.Message)" -Level ERROR
    Complete-Detection -Vulnerable $true -Reason "Detection failed to verify certificate trust state: $($_.Exception.Message)" -Telemetry $telemetry
}

#endregion
