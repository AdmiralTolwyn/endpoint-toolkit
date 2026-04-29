<#
.SYNOPSIS
    Ivanti Secure Boot CA Update Detection (Enhanced).

.DESCRIPTION
    Pure detection script intended for use as an Ivanti Custom Definition (or any
    Status/Reason/Expected/Found-style detect channel). DOES NOT modify the system.

    Compliance logic (intentionally permissive — same as the original Ivanti
    detect script so existing baselines / tickets don't churn):

        Compliant = (UEFICA2023Status = "Updated") AND (UEFICA2023Error = 0)

    Additional diagnostics surfaced in the "found =" line and reason picker
    (informational only, do NOT change the compliance verdict):
      - Event 1801 confidence level (Microsoft-Windows-TPM-WMI / System)
      - Latest in-progress event id (1032, 1033, 1795-1798, 1801-1803)
      - Latest firmware error code from 1795 / 1796 (when present)
      - Known-issue id from 1802 (KI_xxxx)
      - Missing-KEK signal from 1803
      - Reboot-pending signal from 1800
      - WindowsUEFICA2023Capable raw value (with Server 2019 caveat)

    Output contract (one Write-Host per line, exactly these keys):
        detected = true|false
        reason   = <single sentence>
        expected = Status: Updated | Error: 0
        found    = Status: <s> | Error: <e> | Confidence: <c> | Capable: <cap> | ...

    For the active-remediation counterpart (writes registry, starts the
    Secure-Boot-Update scheduled task) see SecureBootCertRemediation.ps1.

.NOTES
    Author:  Anton Romanyuk
    Version: 1.0
    Date:    2026-04-29
    Context: Secure Boot UEFI CA 2023 Deployment — Ivanti Custom Definition

    Compliance gate is DELIBERATELY the old (Status + Error) logic. Do NOT
    tighten it here without coordinating with the Ivanti detect/baseline owner;
    the stricter (1808 + 1799 + Status + Error) gate lives in the remediation
    script for active deployment scenarios.
#>

[CmdletBinding()]
param ()

# -------------------------------------------------------------------------------------------------
# 1. Setup
# -------------------------------------------------------------------------------------------------
$RegPath        = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing'
$NameStatus     = 'UEFICA2023Status'
$NameError      = 'UEFICA2023Error'
$NameCapable    = 'WindowsUEFICA2023Capable'
$TpmWmiProvider = 'Microsoft-Windows-TPM-WMI'

# Target State (compliance gate — same as legacy Ivanti detect)
$TargetStatus = 'Updated'
$TargetError  = 0

# -------------------------------------------------------------------------------------------------
# 2. Registry Data
# -------------------------------------------------------------------------------------------------
$props          = Get-ItemProperty -Path $RegPath -ErrorAction SilentlyContinue
$CurrentStatus  = $props.$NameStatus
$CurrentError   = $props.$NameError
$CurrentCapable = $props.$NameCapable

# Defaults for missing values
if ([string]::IsNullOrEmpty($CurrentStatus)) { $CurrentStatus = 'NotStarted/NotFound' }
if ($null -eq $CurrentError)                 { $CurrentError  = 0 }
$CapableStr = if ($null -ne $CurrentCapable) { "$CurrentCapable" } else { 'N/A' }

# -------------------------------------------------------------------------------------------------
# 3. Event Log Diagnostics (informational only — do NOT gate compliance on these)
# -------------------------------------------------------------------------------------------------

# A. Confidence level from latest 1801
$ConfidenceMsg = 'N/A'
try {
    $Evt1801 = Get-WinEvent -FilterHashtable @{
        LogName      = 'System'
        ProviderName = $TpmWmiProvider
        ID           = 1801
    } -MaxEvents 1 -ErrorAction SilentlyContinue

    if ($Evt1801) {
        if ($Evt1801.Message -match '(High Confidence|Needs More Data|Unknown|Paused|Under Observation)') {
            $ConfidenceMsg = $matches[1]
        } else {
            $ConfidenceMsg = 'Format Error'
        }
    }
} catch {
    $ConfidenceMsg = 'LogAccessError'
}

# B. Sweep recent TPM-WMI events for in-progress / error / known-issue diagnostics.
#    Pulls a single batch (one Get-WinEvent call) and projects different views from it
#    to keep load on the System log low.
$LatestProgressEvent = 'N/A'
$Evt1795Code         = $null
$Evt1796Code         = $null
$KnownIssueId        = $null
$MissingKEK          = $false
$RebootPending       = $false

$ProgressIDs = @(1032, 1033, 1795, 1796, 1797, 1798, 1801, 1802, 1803)
$AllInterest = $ProgressIDs + @(1800)

try {
    $Recent = @(Get-WinEvent -FilterHashtable @{
        LogName      = 'System'
        ProviderName = $TpmWmiProvider
        ID           = $AllInterest
    } -MaxEvents 50 -ErrorAction SilentlyContinue)

    if ($Recent.Count -gt 0) {
        if ($CurrentStatus -eq 'InProgress') {
            $LatestProgress = $Recent | Where-Object { $_.Id -in $ProgressIDs } |
                              Sort-Object TimeCreated -Descending | Select-Object -First 1
            if ($LatestProgress) { $LatestProgressEvent = $LatestProgress.Id }
        }

        # 1795 — firmware error
        $e1795 = $Recent | Where-Object { $_.Id -eq 1795 } |
                 Sort-Object TimeCreated -Descending | Select-Object -First 1
        if ($e1795 -and $e1795.Message -match '(?:error|code|status)[:\s]*(?:0x)?([0-9A-Fa-f]{4,})') {
            $Evt1795Code = $matches[1]
        }

        # 1796 — error code logged during update
        $e1796 = $Recent | Where-Object { $_.Id -eq 1796 } |
                 Sort-Object TimeCreated -Descending | Select-Object -First 1
        if ($e1796 -and $e1796.Message -match '(?:error|code|status)[:\s]*(?:0x)?([0-9A-Fa-f]{4,})') {
            $Evt1796Code = $matches[1]
        }

        # 1802 — known firmware issue (KI_xxxx)
        $e1802 = $Recent | Where-Object { $_.Id -eq 1802 } |
                 Sort-Object TimeCreated -Descending | Select-Object -First 1
        if ($e1802 -and $e1802.Message -match 'SkipReason:\s*(KI_\d+)') {
            $KnownIssueId = $matches[1]
        }

        # 1803 — missing KEK
        if (@($Recent | Where-Object { $_.Id -eq 1803 }).Count -gt 0) { $MissingKEK = $true }

        # 1800 — reboot pending (good event, but actionable)
        if (@($Recent | Where-Object { $_.Id -eq 1800 }).Count -gt 0) { $RebootPending = $true }
    }
} catch {
    # Swallow log errors — diagnostics are best-effort, never fatal.
}

# -------------------------------------------------------------------------------------------------
# 4. Output strings (Ivanti contract)
# -------------------------------------------------------------------------------------------------
$ExpectedString = "Status: $TargetStatus | Error: $TargetError"

$FoundParts = @(
    "Status: $CurrentStatus"
    "Error: $CurrentError"
    "Confidence: $ConfidenceMsg"
    "Capable: $CapableStr"
)
if ($KnownIssueId)        { $FoundParts += "KnownIssue: $KnownIssueId" }
if ($MissingKEK)          { $FoundParts += 'MissingKEK: True' }
if ($RebootPending)       { $FoundParts += 'RebootPending: True' }
if ($Evt1795Code)         { $FoundParts += "FwError(1795): 0x$Evt1795Code" }
if ($Evt1796Code)         { $FoundParts += "LogError(1796): 0x$Evt1796Code" }
$FoundString = $FoundParts -join ' | '

# -------------------------------------------------------------------------------------------------
# 5. Compliance verdict (LEGACY GATE: Status + Error only)
# -------------------------------------------------------------------------------------------------
$IsCompliant = ($CurrentStatus -eq $TargetStatus -and $CurrentError -eq $TargetError)

if (-not $IsCompliant) {
    Write-Host 'detected = true'

    # Dynamic reason — most specific signal first.
    if ($KnownIssueId) {
        Write-Host "reason = Blocked by known firmware issue $KnownIssueId (Event 1802). OEM fix required."
    }
    elseif ($MissingKEK) {
        Write-Host 'reason = Missing KEK update (Event 1803). OEM must supply PK-signed KEK.'
    }
    elseif ($CurrentStatus -eq 'InProgress') {
        Write-Host "reason = Update InProgress. Latest Event ID: $LatestProgressEvent"
    }
    elseif ($CurrentError -ne 0) {
        $errDetail = "Error Code: $CurrentError"
        if ($Evt1795Code) { $errDetail += " | FW Error: 0x$Evt1795Code" }
        if ($Evt1796Code) { $errDetail += " | Log Error: 0x$Evt1796Code" }
        Write-Host "reason = Update Failed. $errDetail"
    }
    elseif ($RebootPending) {
        Write-Host 'reason = Reboot pending (Event 1800). Update will proceed after reboot.'
    }
    elseif ($null -ne $CurrentCapable -and $CurrentCapable -eq 0) {
        Write-Host 'reason = Device not capable (WindowsUEFICA2023Capable=0). NOTE: This value is unreliable on Server 2019.'
    }
    else {
        Write-Host "reason = Status is '$CurrentStatus' (Waiting for 'Updated')"
    }
}
else {
    Write-Host 'detected = false'
    Write-Host 'reason = Secure Boot Update successful.'
}

Write-Host "expected = $ExpectedString"
Write-Host "found = $FoundString"
