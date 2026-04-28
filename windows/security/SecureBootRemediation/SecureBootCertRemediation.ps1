<#
.SYNOPSIS
    Unified Secure Boot CA 2023 Analysis, Detection, and Remediation script.

.DESCRIPTION
    This script combines comprehensive detection logic with active remediation:
    1. DETECTION: 
       - Queries all Secure Boot Playbook registry keys (AvailableUpdates, AvailableUpdatesPolicy,
         HighConfidenceOptOut, MicrosoftUpdateManagedOptIn, Servicing keys).
       - Full event sweep: good (1034,1036,1037,1042-1045,1799,1800,1808)
         and warning/bad (1032,1033,1795-1798,1801,1802,1803) from System TPM-WMI log.
       - NOTE: Event 1801 is a status/assessment event that fires when the update
         has NOT yet completed or when issues are detected - classified as warning.
       - Event 1799 dual-log check (System + TPM-WMI/Operational).
       - Reports confidence, BucketId, error codes, and structured debug output.
    
    2. REMEDIATION (Smart Logic):
       - Initial Run: If AvUpdates is 0, sets 0x5944 and triggers task.
       - Post-Reboot: If AvUpdates is 0x4100, triggers task to finalize to 0x4000.
    
    3. MONITORING:
       - Loops for 30 seconds to track registry state changes in real-time.

.PARAMETER ForceRemediation
    Switch to force the registry trigger even if the state appears valid or indeterminate.

.NOTES
    Author:  Anton Romanyuk
    Version: 2.1
    Date:    2026-04-15
    Context: Secure Boot UEFI CA 2023 Deployment
    Changes: v2.1 - Reclassify 1801 as warning (not good); compliance = 1808 + Updated (PG rec);
             Server 2019 WindowsUEFICA2023Capable=0 bug caveat
#>

[CmdletBinding()]
param (
    [Switch]$ForceRemediation
)

# -------------------------------------------------------------------------------------------------
# 1. HELPER FUNCTIONS
# -------------------------------------------------------------------------------------------------

function Write-ColorLog {
    param (
        [Parameter(Mandatory=$true)] [string]$Message,
        [Parameter(Mandatory=$true)] [ValidateSet('Info', 'Success', 'Warning', 'Error', 'Verbose')] [string]$Level
    )
    $colorMap = @{ 'Info'='Cyan'; 'Success'='Green'; 'Warning'='Yellow'; 'Error'='Red'; 'Verbose'='Gray' }
    Write-Host "[$((Get-Date).ToString('HH:mm:ss'))] [$Level] $Message" -ForegroundColor $colorMap[$Level]
}

# Column widths for debug output alignment
$script:ColLabel = 30
$script:ColValue = 18

function Write-DebugField {
    param([string]$Label, [string]$Value, [string]$Color = 'White', [string]$Comment = '')
    $padLabel = $Label.PadRight($script:ColLabel)
    $padValue = $Value.PadRight($script:ColValue)
    Write-Host "  $padLabel" -NoNewline
    Write-Host " $padValue" -ForegroundColor $Color -NoNewline
    if ($Comment) { Write-Host " $Comment" -ForegroundColor DarkGray } else { Write-Host "" }
}

function Get-AvailableUpdatesDecoding {
    param ([int]$Value)
    $details = @()
    if ($Value -band 0x0040) { $details += "PENDING: Add Windows UEFI CA 2023 to DB (0x0040)" }
    if ($Value -band 0x0800) { $details += "PENDING: Apply Microsoft Option ROM UEFI CA 2023 (0x0800)" }
    if ($Value -band 0x1000) { $details += "PENDING: Apply Microsoft UEFI CA 2023 (0x1000)" }
    if ($Value -band 0x4000) { $details += "STATE: Modifier - Only apply if 2011 CA exists (0x4000)" }
    if ($Value -band 0x0004) { $details += "PENDING: Apply Key Exchange Key (KEK) Update (0x0004)" }
    if ($Value -band 0x0100) { $details += "PENDING: Apply Boot Manager signed by UEFI CA 2023 (0x0100)" }
    
    if ($details.Count -eq 0) {
        if ($Value -eq 0) { return "None (0x0)" }
        return "Unknown/Other ($('0x{0:X}' -f $Value))"
    }
    return $details -join "`n                                  "
}

# -------------------------------------------------------------------------------------------------
# 2. CORE DETECTION LOGIC
# -------------------------------------------------------------------------------------------------

function Get-SecureBootStatus {
    Write-ColorLog -Message "Starting Unified Secure Boot Analysis (v2.0)..." -Level "Info"

    $TpmWmiProvider = 'Microsoft-Windows-TPM-WMI'
    $regPathRoot      = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot'
    $regPathServicing = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing'

    # -------------------------------------------------------------------------
    # A. Registry: SecureBoot Root Keys
    # -------------------------------------------------------------------------
    $avUpdatesValue    = $null
    $avUpdatesHex      = "N/A"
    $avUpdatesPolicyValue = $null
    $avUpdatesPolicyHex   = "N/A"
    $highConfidenceOptOut       = $null
    $microsoftUpdateManagedOptIn = $null

    try {
        $val = Get-ItemProperty -Path $regPathRoot -Name 'AvailableUpdates' -ErrorAction Stop
        $avUpdatesValue = $val.AvailableUpdates
        if ($null -ne $avUpdatesValue) { $avUpdatesHex = '0x{0:X}' -f $avUpdatesValue }
    } catch { $avUpdatesValue = $null }

    try {
        $val = Get-ItemProperty -Path $regPathRoot -Name 'AvailableUpdatesPolicy' -ErrorAction Stop
        $avUpdatesPolicyValue = $val.AvailableUpdatesPolicy
        if ($null -ne $avUpdatesPolicyValue) { $avUpdatesPolicyHex = '0x{0:X}' -f $avUpdatesPolicyValue }
    } catch { $avUpdatesPolicyValue = $null }

    try {
        $val = Get-ItemProperty -Path $regPathRoot -Name 'HighConfidenceOptOut' -ErrorAction Stop
        $highConfidenceOptOut = $val.HighConfidenceOptOut
    } catch { $highConfidenceOptOut = $null }

    try {
        $val = Get-ItemProperty -Path $regPathRoot -Name 'MicrosoftUpdateManagedOptIn' -ErrorAction Stop
        $microsoftUpdateManagedOptIn = $val.MicrosoftUpdateManagedOptIn
    } catch { $microsoftUpdateManagedOptIn = $null }

    # -------------------------------------------------------------------------
    # B. Registry: Servicing Keys
    # -------------------------------------------------------------------------
    $servicingData = @{ Status = "NotStarted/NotFound"; Error = 0; Capable = $null }

    if (Test-Path $regPathServicing) {
        try {
            $props = Get-ItemProperty -Path $regPathServicing -ErrorAction SilentlyContinue
            if ($props.UEFICA2023Status)          { $servicingData.Status  = $props.UEFICA2023Status }
            if ($null -ne $props.UEFICA2023Error)  { $servicingData.Error   = $props.UEFICA2023Error }
            $servicingData.Capable = $props.WindowsUEFICA2023Capable   # may be $null
        } catch {
            Write-ColorLog -Message "Failed to read Servicing keys: $($_.Exception.Message)" -Level "Warning"
        }
    }

    # -------------------------------------------------------------------------
    # C. Event Log - Full Sweep (System log, TPM-WMI)
    # -------------------------------------------------------------------------
    $GoodEventIDs = @(1034, 1036, 1037, 1042, 1043, 1044, 1045, 1799, 1800, 1808)
    # 1801 is a STATUS/assessment event (fires under "Under Observation" too) — tracked as warning,
    # but NOT a real blocker. True blockers are firmware/KI/KEK errors.
    $WarningEventIDs = @(1032, 1033, 1801)
    $BlockingEventIDs = @(1795, 1796, 1797, 1798, 1802, 1803)
    $BadEventIDs  = $WarningEventIDs + $BlockingEventIDs
    $AllKnownIDs  = $GoodEventIDs + $BadEventIDs

    # --- C1. Event 1801 - Confidence, BucketId, UpdateType, DeviceAttributes ---
    $confidenceLevel    = "N/A"
    $bucketId           = "N/A"
    $updateType         = "N/A"
    $deviceAttributes   = "N/A"
    $latestStatusSummary = "N/A"
    $latestStatusSource  = $null

    try {
        $Evt1801 = Get-WinEvent -FilterHashtable @{LogName='System'; ProviderName=$TpmWmiProvider; ID=1801} -MaxEvents 1 -ErrorAction SilentlyContinue
        if ($Evt1801) {
            if ($Evt1801.Message -match 'BucketConfidenceLevel:\s*(.*)') {
                $v = $matches[1].Trim()
                $confidenceLevel = if ([string]::IsNullOrEmpty($v)) { 'Empty (not yet evaluated)' } else { $v }
            }
            if ($Evt1801.Message -match 'BucketId:\s*(.+)')      { $bucketId         = $matches[1].Trim() }
            if ($Evt1801.Message -match 'UpdateType:\s*(.*)')     {
                $v = $matches[1].Trim()
                $updateType = if ([string]::IsNullOrEmpty($v)) { 'Empty' } else { $v }
            }
            if ($Evt1801.Message -match 'DeviceAttributes:\s*(.+)') { $deviceAttributes = $matches[1].Trim() }
            $firstLine = ($Evt1801.Message -split "`n")[0].Trim()
            if (-not [string]::IsNullOrEmpty($firstLine)) {
                $latestStatusSummary = $firstLine
                $latestStatusSource  = 1801
            }
        }
    } catch { $confidenceLevel = "LogAccessError" }

    # --- C2. In-Progress bad-event check ---
    $latestProgressEvent = "N/A"
    if ($servicingData.Status -eq "InProgress") {
        try {
            $RecentLog = Get-WinEvent -FilterHashtable @{LogName='System'; ProviderName=$TpmWmiProvider; ID=$BadEventIDs} -MaxEvents 1 -ErrorAction SilentlyContinue
            if ($RecentLog) { $latestProgressEvent = $RecentLog.Id }
        } catch { $latestProgressEvent = "LogAccessError" }
    }

    # --- C3. Full Event Sweep ---
    $evt1808Count       = 0
    $evt1795ErrorCode   = $null
    $evt1796ErrorCode   = $null
    $rebootPending      = $false
    $knownIssueId       = $null
    $missingKEK         = $false
    $skipReasonKI       = $null
    $latestGoodId       = $null
    $latestBadId        = $null
    $bootloaderSwapped  = $false
    $blockingIssue      = $false
    $updateSuccess      = $false

    try {
        $AllEvents = @(Get-WinEvent -FilterHashtable @{LogName='System'; ProviderName=$TpmWmiProvider; ID=$AllKnownIDs} -MaxEvents 100 -ErrorAction SilentlyContinue)

        if ($AllEvents.Count -gt 0) {
            $evt1808Count = @($AllEvents | Where-Object { $_.Id -eq 1808 }).Count
            if ($evt1808Count -gt 0) { $updateSuccess = $true }

            $latestOverall = $AllEvents | Sort-Object TimeCreated -Descending | Select-Object -First 1

            # Latest good vs bad event
            $latestGoodEvt = $AllEvents | Where-Object { $_.Id -in $GoodEventIDs } | Sort-Object TimeCreated -Descending | Select-Object -First 1
            $latestBadEvt  = $AllEvents | Where-Object { $_.Id -in $BadEventIDs }  | Sort-Object TimeCreated -Descending | Select-Object -First 1
            if ($latestGoodEvt) { $latestGoodId = $latestGoodEvt.Id }
            if ($latestBadEvt)  { $latestBadId  = $latestBadEvt.Id }
            # Only flag as blocking if a true error event is present (not status/warning like 1801)
            $latestBlockingEvt = $AllEvents | Where-Object { $_.Id -in $BlockingEventIDs } | Sort-Object TimeCreated -Descending | Select-Object -First 1
            if ($latestBlockingEvt) { $blockingIssue = $true }

            # Event 1799 - bootloader swapped (System log)
            $bootloaderSwapped = (@($AllEvents | Where-Object { $_.Id -eq 1799 }).Count -gt 0)

            # SkipReason from latest event with BucketId
            if ($null -ne $latestOverall -and $latestOverall.Message -match 'SkipReason:\s*(KI_\d+)') {
                $skipReasonKI = $matches[1]
            }

            # Only parse error detail if update is NOT complete
            $updateComplete = ($latestOverall.Id -eq 1808) -or ($servicingData.Status -eq 'Updated')
            if (-not $updateComplete) {
                # 1795 - Firmware error
                $e1795 = @($AllEvents | Where-Object { $_.Id -eq 1795 })
                if ($e1795.Count -gt 0) {
                    $latest1795 = $e1795 | Sort-Object TimeCreated -Descending | Select-Object -First 1
                    if ($latest1795.Message -match '(?:error|code|status)[:\s]*(?:0x)?([0-9A-Fa-f]{4,})') {
                        $evt1795ErrorCode = $matches[1]
                    }
                }
                # 1796 - Error code logged
                $e1796 = @($AllEvents | Where-Object { $_.Id -eq 1796 })
                if ($e1796.Count -gt 0) {
                    $latest1796 = $e1796 | Sort-Object TimeCreated -Descending | Select-Object -First 1
                    if ($latest1796.Message -match '(?:error|code|status)[:\s]*(?:0x)?([0-9A-Fa-f]{4,})') {
                        $evt1796ErrorCode = $matches[1]
                    }
                }
                # 1800 - Reboot needed (good event, not an error)
                $rebootPending = (@($AllEvents | Where-Object { $_.Id -eq 1800 }).Count -gt 0)
                # 1802 - Known firmware issue
                $e1802 = @($AllEvents | Where-Object { $_.Id -eq 1802 })
                if ($e1802.Count -gt 0) {
                    $latest1802 = $e1802 | Sort-Object TimeCreated -Descending | Select-Object -First 1
                    if ($latest1802.Message -match 'SkipReason:\s*(KI_\d+)') {
                        $knownIssueId = $matches[1]
                    }
                }
                # 1803 - Missing KEK
                $missingKEK = (@($AllEvents | Where-Object { $_.Id -eq 1803 }).Count -gt 0)
            }
        }
    } catch {
        Write-ColorLog -Message "Event log query failed or returned no data." -Level "Verbose"
    }

    # --- C4. Event 1799 from TPM-WMI Operational log (bootloader swap verification) ---
    $evt1799Operational = $false
    try {
        $evt1799Op = Get-WinEvent -FilterHashtable @{
            LogName = 'Microsoft-Windows-TPM-WMI/Operational'
            Id      = 1799
        } -MaxEvents 1 -ErrorAction Stop
        if ($null -ne $evt1799Op) { $evt1799Operational = $true }
    } catch {
        if ($_.Exception.Message -notmatch 'No events were found') {
            $evt1799Operational = $null   # unknown (access denied, log not present)
        }
    }
    if ($evt1799Operational -eq $true) { $bootloaderSwapped = $true }

    # --- C5. Override summary with Event 1808 message if update completed ---
    if ($evt1808Count -gt 0) {
        try {
            $Evt1808 = Get-WinEvent -FilterHashtable @{LogName='System'; ProviderName=$TpmWmiProvider; ID=1808} -MaxEvents 1 -ErrorAction SilentlyContinue
            if ($Evt1808) {
                $firstLine = ($Evt1808.Message -split "`n")[0].Trim()
                if (-not [string]::IsNullOrEmpty($firstLine)) {
                    $latestStatusSummary = $firstLine
                    $latestStatusSource  = 1808
                }
            }
        } catch { <# keep 1801 summary as fallback #> }
    }

    # -------------------------------------------------------------------------
    # D. Structured Debug Output
    # -------------------------------------------------------------------------
    $w = 72
    $bar = '=' * $w

    Write-Host ""
    Write-Host $bar -ForegroundColor DarkCyan
    Write-Host "  Secure Boot CA 2023 - Detection Report" -ForegroundColor Cyan
    Write-Host $bar -ForegroundColor DarkCyan

    # Header row
    Write-Host ""
    $hdrLabel = "FIELD".PadRight($script:ColLabel)
    $hdrValue = "VALUE".PadRight($script:ColValue)
    Write-Host "  $hdrLabel $hdrValue DESCRIPTION" -ForegroundColor DarkGray
    Write-Host "  $('-' * $script:ColLabel) $('-' * $script:ColValue) $('-' * 20)" -ForegroundColor DarkGray

    # --- Registry: SecureBoot ---
    Write-Host ""
    Write-Host "  [Registry] HKLM\...\SecureBoot" -ForegroundColor DarkCyan

    $clr = if ($servicingData.Status -eq 'Updated') { 'Green' } elseif ($servicingData.Status -eq 'InProgress') { 'Yellow' } else { 'Red' }
    Write-DebugField 'UEFICA2023Status' $servicingData.Status $clr 'Target: Updated'

    $clr = if ($servicingData.Error -eq 0) { 'Green' } else { 'Red' }
    Write-DebugField 'UEFICA2023Error' "$($servicingData.Error)" $clr 'Target: 0'

    $capVal = if ($null -ne $servicingData.Capable) { "$($servicingData.Capable)" } else { 'N/A' }
    $clr = if ($null -eq $servicingData.Capable) { 'DarkGray' } elseif ($servicingData.Capable -ge 1) { 'Green' } else { 'Yellow' }
    Write-DebugField 'WindowsUEFICA2023Capable' $capVal $clr '0=not capable, 1+=capable (UNRELIABLE on Server 2019)'

    $clr = if ($null -eq $avUpdatesValue) { 'DarkGray' } elseif ($avUpdatesValue -eq 0) { 'Yellow' } else { 'Cyan' }
    Write-DebugField 'AvailableUpdates' $avUpdatesHex $clr 'Bitmask of pending updates'
    if ($null -ne $avUpdatesValue -and $avUpdatesValue -gt 0) {
        Write-ColorLog -Message "  $(Get-AvailableUpdatesDecoding -Value $avUpdatesValue)" -Level "Verbose"
    }

    $clr = if ($null -eq $avUpdatesPolicyValue) { 'DarkGray' } else { 'Cyan' }
    Write-DebugField 'AvailableUpdatesPolicy' $avUpdatesPolicyHex $clr 'Policy-set bitmask (GPO/Intune)'

    $hcoVal = if ($null -ne $highConfidenceOptOut) { "$highConfidenceOptOut" } else { 'NotSet' }
    $clr = if ($null -eq $highConfidenceOptOut) { 'DarkGray' } elseif ($highConfidenceOptOut -eq 0) { 'Green' } else { 'Yellow' }
    Write-DebugField 'HighConfidenceOptOut' $hcoVal $clr '1=opted out of auto-update'

    $muVal = if ($null -ne $microsoftUpdateManagedOptIn) { "$microsoftUpdateManagedOptIn" } else { 'NotSet' }
    $clr = if ($null -eq $microsoftUpdateManagedOptIn) { 'DarkGray' } elseif ($microsoftUpdateManagedOptIn -ge 1) { 'Green' } else { 'Yellow' }
    Write-DebugField 'MicrosoftUpdateManagedOptIn' $muVal $clr '1=org opted in via WSUS/WUfB'

    # --- Event Log Summary ---
    Write-Host ""
    Write-Host "  [Event Log] System - TPM-WMI" -ForegroundColor DarkCyan

    $clr = if ($confidenceLevel -match 'High Confidence') { 'Green' } elseif ($confidenceLevel -eq 'N/A') { 'DarkGray' } else { 'Yellow' }
    Write-DebugField 'Confidence' $confidenceLevel $clr 'Event 1801 bucket level'

    Write-DebugField 'BucketId' $bucketId $(if ($bucketId -eq 'N/A') { 'DarkGray' } else { 'Cyan' }) 'Device bucket from Event 1801'

    $clr = if ($updateType -eq 'N/A') { 'DarkGray' } else { 'Cyan' }
    Write-DebugField 'UpdateType' $updateType $clr 'From Event 1801'

    if ($deviceAttributes -ne 'N/A') {
        Write-DebugField 'DeviceAttributes' $deviceAttributes 'Cyan' 'FW/OEM info from 1801'
    }
    if ($latestStatusSummary -ne 'N/A') {
        $srcLabel = if ($latestStatusSource -eq 1808) { 'StatusSummary (1808)' } else { 'StatusSummary (1801)' }
        $clr      = if ($latestStatusSource -eq 1808) { 'Green' } else { 'Yellow' }
        Write-DebugField $srcLabel $latestStatusSummary $clr 'First line of latest status event'
    }

    $lgVal = if ($latestGoodId) { "$latestGoodId" } else { 'None' }
    $lbVal = if ($latestBadId)  { "$latestBadId"  } else { 'None' }
    Write-DebugField 'LatestGoodEvent' $lgVal $(if ($latestGoodId) { 'Green' } else { 'DarkGray' }) 'Most recent success event'
    Write-DebugField 'LatestBadEvent' $lbVal $(if ($latestBadId) { 'Red' } else { 'DarkGray' }) 'Most recent error event'

    $clr = if ($bootloaderSwapped) { 'Green' } else { 'DarkGray' }
    Write-DebugField 'BootloaderSwapped' "$bootloaderSwapped" $clr 'Event 1799 System/Operational'

    $clr = if ($evt1808Count -gt 0) { 'Green' } else { 'DarkGray' }
    Write-DebugField 'CompletedEvents (1808)' "$evt1808Count" $clr 'Successful completions'

    if ($rebootPending) {
        Write-DebugField 'RebootPending' 'True' 'Cyan' 'Event 1800 - reboot needed'
    }
    if ($evt1795ErrorCode) {
        Write-DebugField 'FirmwareError (1795)' "0x$evt1795ErrorCode" 'Red' 'Error from firmware'
    }
    if ($evt1796ErrorCode) {
        Write-DebugField 'ErrorLogged (1796)' "0x$evt1796ErrorCode" 'Red' 'Error during update'
    }
    if ($knownIssueId) {
        Write-DebugField 'KnownIssue (1802)' $knownIssueId 'Magenta' 'FW blocked by known issue'
    }
    if ($skipReasonKI) {
        Write-DebugField 'SkipReason' $skipReasonKI 'Magenta' 'KI from BucketId event'
    }
    if ($missingKEK) {
        Write-DebugField 'MissingKEK (1803)' 'True' 'Red' 'OEM must supply PK-signed KEK'
    }

    Write-Host ""
    Write-Host $bar -ForegroundColor DarkCyan

    # -------------------------------------------------------------------------
    # E. Detection Summary (Ivanti-format reason)
    # -------------------------------------------------------------------------
    # PG recommendation: compliance = Event 1808 present AND UEFICA2023Status=Updated
    $isCompliant = ($updateSuccess -and $servicingData.Status -eq 'Updated')

    $expectedStr = "Status: Updated | Event1808: True"
    $foundStr    = "Status: $($servicingData.Status) | Event1808: $updateSuccess | Error: $($servicingData.Error) | Confidence: $confidenceLevel | Capable: $capVal | AvUpdates: $avUpdatesHex | BootloaderSwapped: $bootloaderSwapped"

    Write-Host ""
    Write-ColorLog -Message "--- DETECTION SUMMARY ---" -Level "Info"

    if ($isCompliant) {
        Write-ColorLog -Message "Detected  : false (Compliant)" -Level "Success"
        Write-ColorLog -Message "Reason    : Secure Boot Update successful (Event 1808 + Status=Updated)." -Level "Success"
    } else {
        Write-ColorLog -Message "Detected  : true (Non-Compliant)" -Level "Warning"
        if ($knownIssueId) {
            Write-ColorLog -Message "Reason    : Blocked by known firmware issue $knownIssueId" -Level "Error"
        } elseif ($missingKEK) {
            Write-ColorLog -Message "Reason    : Missing KEK update (Event 1803). OEM must supply PK-signed KEK." -Level "Error"
        } elseif ($rebootPending) {
            Write-ColorLog -Message "Reason    : Reboot pending (Event 1800). Update will proceed after reboot." -Level "Warning"
        } elseif ($servicingData.Status -eq 'InProgress') {
            Write-ColorLog -Message "Reason    : Update InProgress. Latest Event ID: $latestProgressEvent" -Level "Warning"
        } elseif ($servicingData.Error -ne 0) {
            $errDetail = "Error Code: $($servicingData.Error)"
            if ($evt1795ErrorCode) { $errDetail += " | FW Error: 0x$evt1795ErrorCode" }
            if ($evt1796ErrorCode) { $errDetail += " | Log Error: 0x$evt1796ErrorCode" }
            Write-ColorLog -Message "Reason    : Update Failed. $errDetail" -Level "Error"
        } elseif ($null -ne $servicingData.Capable -and $servicingData.Capable -eq 0) {
            Write-ColorLog -Message "Reason    : Device not capable (WindowsUEFICA2023Capable=0). NOTE: This value is unreliable on Server 2019." -Level "Error"
        } else {
            Write-ColorLog -Message "Reason    : Status is '$($servicingData.Status)' (Waiting for 'Updated')" -Level "Warning"
        }
    }

    Write-ColorLog -Message "Expected  : $expectedStr" -Level "Verbose"
    Write-ColorLog -Message "Found     : $foundStr" -Level "Verbose"

    return @{
        Success            = $updateSuccess
        Blocking           = $blockingIssue
        AvUpdates          = if ($null -ne $avUpdatesValue) { $avUpdatesValue } else { 0 }
        ServicingStat      = $servicingData.Status
        ServicingErr       = $servicingData.Error
        Capable            = $servicingData.Capable
        Confidence         = $confidenceLevel
        Compliant          = $isCompliant
        RebootPending      = $rebootPending
        KnownIssueId       = $knownIssueId
        MissingKEK         = $missingKEK
        BootloaderSwapped  = $bootloaderSwapped
        LatestGoodId       = $latestGoodId
        LatestBadId        = $latestBadId
    }
}

# -------------------------------------------------------------------------------------------------
# 3. REMEDIATION & MONITORING
# -------------------------------------------------------------------------------------------------

function Invoke-Remediation {
    param ($StateObj)

    $regPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot'
    $taskName = "\Microsoft\Windows\PI\Secure-Boot-Update"
    $triggerRequired = $false
    $triggerReason = ""

    # --- LOGIC: When to trigger the task? ---
    
    # 1. Initial Deployment: Key is 0 and no success event.
    if (-not $StateObj.Success -and $StateObj.AvUpdates -eq 0) {
        $triggerRequired = $true
        $triggerReason = "Initial Deployment"
    }
    # 2. Post-Reboot Finalization: Key is 0x4100 (Boot Manager Staged).
    elseif ($StateObj.AvUpdates -eq 0x4100) {
        $triggerRequired = $true
        $triggerReason = "Post-Reboot Finalization"
    }
    # 3. Forced
    elseif ($ForceRemediation) {
        $triggerRequired = $true
        $triggerReason = "Forced by User"
    }

    # --- ABORT if Blocking Issues Found (unless forced) ---
    if ($StateObj.Blocking -and -not $ForceRemediation) {
        Write-Host ""
        Write-ColorLog -Message "--- REMEDIATION ABORTED ---" -Level "Error"
        if ($StateObj.KnownIssueId) {
            Write-ColorLog -Message "Known firmware issue $($StateObj.KnownIssueId) blocks update. OEM fix required." -Level "Warning"
        } elseif ($StateObj.MissingKEK) {
            Write-ColorLog -Message "Missing KEK update (Event 1803). OEM must supply PK-signed KEK." -Level "Warning"
        } else {
            Write-ColorLog -Message "Blocking issues (BitLocker/Firmware) detected. Resolve these first." -Level "Warning"
        }
        return
    }

    # --- INFO: Reboot pending (not a blocker) ---
    if ($StateObj.RebootPending -and -not $triggerRequired) {
        Write-Host ""
        Write-ColorLog -Message "Reboot pending (Event 1800). Update will proceed after reboot." -Level "Info"
    }

    if ($triggerRequired) {
        Write-Host ""
        Write-ColorLog -Message "--- INITIATING DEPLOYMENT ($triggerReason) ---" -Level "Info"
        
        try {
            # Step A: If Initial Deployment, Set Registry to 0x5944
            if ($StateObj.AvUpdates -eq 0 -or ($ForceRemediation -and $StateObj.AvUpdates -ne 0x4100)) {
                New-ItemProperty -Path $regPath -Name "AvailableUpdates" -Value 0x5944 -PropertyType DWord -Force | Out-Null
                Write-ColorLog -Message "Set AvailableUpdates to 0x5944." -Level "Success"
            } elseif ($StateObj.AvUpdates -eq 0x4100) {
                Write-ColorLog -Message "Existing State is 0x4100. Triggering task to finalize." -Level "Info"
            }

            # Step B: Run Scheduled Task
            Write-ColorLog -Message "Starting Scheduled Task: $taskName" -Level "Info"
            Start-ScheduledTask -TaskName $taskName | Out-Null

            # Step C: Monitor Loop (30 Seconds)
            Write-ColorLog -Message "Monitoring task progress (30s)..." -Level "Verbose"
            
            $finalVal = 0
            for ($i = 1; $i -le 6; $i++) {
                Start-Sleep -Seconds 5
                $finalVal = (Get-ItemProperty -Path $regPath -Name "AvailableUpdates" -ErrorAction SilentlyContinue).AvailableUpdates
                $hex = '0x{0:X}' -f $finalVal
                Write-ColorLog -Message "[$($i*5)s] AvailableUpdates: $hex" -Level "Verbose"
            }

            # Step D: Final Recommendation
            Write-Host ""
            Write-ColorLog -Message "--- RECOMMENDATION ---" -Level "Info"
            if ($finalVal -eq 0x4100) {
                Write-ColorLog -Message "State 0x4100 (Staged). ACTION: Reboot Manually." -Level "Success"
            } elseif ($finalVal -eq 0x4000) {
                Write-ColorLog -Message "State 0x4000 (Complete). ACTION: None. Update Finished." -Level "Success"
            } elseif ($finalVal -eq 0x4104) {
                Write-ColorLog -Message "State 0x4104 (KEK Pending). ACTION: Reboot." -Level "Warning"
            } elseif ($finalVal -eq 0x5944) {
                Write-ColorLog -Message "State 0x5944 (No Change). Task may be delayed. Wait or Reboot." -Level "Warning"
            } else {
                Write-ColorLog -Message "Progression detected. ACTION: Reboot to continue." -Level "Info"
            }

        } catch {
            Write-ColorLog -Message "Remediation Failed: $($_.Exception.Message)" -Level "Error"
        }
    } else {
        # No Action Needed
        if ($StateObj.Success -or $StateObj.AvUpdates -eq 0x4000) {
            Write-ColorLog -Message "System is fully updated (0x4000 / Event 1808)." -Level "Success"
        } else {
            Write-ColorLog -Message "Updates in progress or intermediate state ($('0x{0:X}' -f $StateObj.AvUpdates)). No action taken." -Level "Verbose"
        }
    }
}

# -------------------------------------------------------------------------------------------------
# EXECUTION ENTRY POINT
# -------------------------------------------------------------------------------------------------

$state = Get-SecureBootStatus
Invoke-Remediation -StateObj $state