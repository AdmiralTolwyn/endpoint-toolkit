<#
.SYNOPSIS
    Iteratively resets dot3svc migration status to ensure wired profile application on startup.

.DESCRIPTION
    Designed to run post-upgrade. This script resets the Wired AutoConfig (dot3svc) migration
    flag and restarts the service multiple times with configurable delays. This accounts for
    slower service initialization or hardware enumeration delays after OS upgrades.

    The script produces both a structured log file and color-coded console output.

.PARAMETER MaxRetries
    Number of reset iterations to perform. Default: 3.

.PARAMETER RetryDelaySeconds
    Seconds to wait between iterations. Default: 30.

.PARAMETER LogDirectory
    Directory for the log file. Default: $env:SystemDrive\Windows\Temp.

.PARAMETER RepairPolicies
    Enables post-upgrade policy repair logic. When set, the script will:
    1. Check for symlinks in the dot3svc Policies folder (unless -SkipSymlinkCheck is set)
    2. Search C:\Windows.old for the original 802.1x policy files
    3. Copy recovered policies to the active policies folder and restart the service
    This switch is DISABLED by default.

.PARAMETER SkipSymlinkCheck
    When -RepairPolicies is enabled, skip the symlink detection step.
    By default (without this switch), symlink checking IS performed.

.EXAMPLE
    .\Reset-Dot3SvcMigration.ps1
    Runs with default settings (3 retries, 30s delay). No policy repair.

.EXAMPLE
    .\Reset-Dot3SvcMigration.ps1 -MaxRetries 5 -RetryDelaySeconds 15
    Runs 5 iterations with 15-second delays.

.EXAMPLE
    .\Reset-Dot3SvcMigration.ps1 -RepairPolicies
    Runs migration reset AND checks for symlinks + recovers policies from Windows.old.

.EXAMPLE
    .\Reset-Dot3SvcMigration.ps1 -RepairPolicies -SkipSymlinkCheck
    Runs policy repair but skips symlink detection (only does Windows.old recovery).
#>

[CmdletBinding()]
param(
    [ValidateRange(1, 10)]
    [int]$MaxRetries = 3,

    [ValidateRange(5, 120)]
    [int]$RetryDelaySeconds = 30,

    [string]$LogDirectory = "$env:SystemDrive\Windows\Temp",

    [switch]$RepairPolicies,

    [switch]$SkipSymlinkCheck
)

#region Logging

$Script:LogFile = Join-Path $LogDirectory "Dot3Svc_Fix_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS', 'STEP')]
        [string]$Level = 'INFO'
    )

    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
    $entry = "[$ts] [$Level] $Message"

    # Write to log file
    try { $entry | Out-File -FilePath $Script:LogFile -Append -Encoding utf8 -ErrorAction Stop }
    catch { Write-Warning "Failed to write to log: $_" }

    # Console output with color coding
    $color = switch ($Level) {
        'INFO'    { 'Cyan' }
        'WARN'    { 'Yellow' }
        'ERROR'   { 'Red' }
        'SUCCESS' { 'Green' }
        'STEP'    { 'White' }
    }
    Write-Host "  $entry" -ForegroundColor $color
}

function Write-Banner {
    param([string]$Text, [string]$Char = '-', [int]$Width = 70)
    $line = $Char * $Width
    Write-Host ""
    Write-Host "  $line" -ForegroundColor DarkGray
    Write-Host "  $Text" -ForegroundColor White
    Write-Host "  $line" -ForegroundColor DarkGray
    # Log without ANSI
    "" | Out-File -FilePath $Script:LogFile -Append -Encoding utf8
    $line | Out-File -FilePath $Script:LogFile -Append -Encoding utf8
    $Text | Out-File -FilePath $Script:LogFile -Append -Encoding utf8
    $line | Out-File -FilePath $Script:LogFile -Append -Encoding utf8
}

#endregion

#region Helpers

function Test-IsAdmin {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]$identity
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-ServiceState {
    param([string]$Name)
    try {
        $svc = Get-Service -Name $Name -ErrorAction Stop
        return $svc.Status.ToString()
    }
    catch {
        return 'NotFound'
    }
}

function Wait-ServiceStatus {
    param(
        [string]$Name,
        [string]$DesiredStatus,
        [int]$TimeoutSeconds = 30
    )
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        $state = Get-ServiceState -Name $Name
        if ($state -eq $DesiredStatus) {
            $sw.Stop()
            return $true
        }
        Start-Sleep -Milliseconds 500
    }
    $sw.Stop()
    return $false
}

#endregion

#region Main

$regPath = 'HKLM:\SOFTWARE\Microsoft\dot3svc\MigrationData'
$regName = 'dot3svcMigrationDone'
$serviceName = 'dot3svc'
$exitCode = 0
$successCount = 0
$failCount = 0

# Ensure log directory exists
if (-not (Test-Path $LogDirectory)) {
    try {
        New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
    }
    catch {
        Write-Warning "Cannot create log directory '$LogDirectory'. Falling back to TEMP."
        $LogDirectory = $env:TEMP
        $Script:LogFile = Join-Path $LogDirectory "Dot3Svc_Fix_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    }
}

Write-Banner "Wired Profile Remediation  |  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

Write-Log "Script started" -Level INFO
Write-Log "Log file: $Script:LogFile" -Level INFO
Write-Log "Parameters: MaxRetries=$MaxRetries, RetryDelay=${RetryDelaySeconds}s, RepairPolicies=$RepairPolicies, SkipSymlinkCheck=$SkipSymlinkCheck" -Level INFO
Write-Log "Computer: $env:COMPUTERNAME | User: $env:USERNAME" -Level INFO

# Pre-flight: admin check
if (-not (Test-IsAdmin)) {
    Write-Log "Script is NOT running as Administrator. Registry and service operations may fail." -Level WARN
}

# Pre-flight: service existence
$initialState = Get-ServiceState -Name $serviceName
if ($initialState -eq 'NotFound') {
    Write-Log "Service '$serviceName' not found on this machine. Exiting." -Level ERROR
    exit 1
}
Write-Log "Service '$serviceName' initial state: $initialState" -Level INFO

Write-Banner "Starting $MaxRetries iteration(s) with ${RetryDelaySeconds}s delay"

for ($i = 1; $i -le $MaxRetries; $i++) {
    $iterationSuccess = $true

    Write-Host ""
    Write-Log "===== Iteration $i of $MaxRetries =====" -Level STEP

    # Step 1: Ensure registry path
    try {
        if (-not (Test-Path $regPath)) {
            Write-Log "Registry path not found - creating '$regPath'" -Level WARN
            New-Item -Path $regPath -Force -ErrorAction Stop | Out-Null
            Write-Log "Registry path created" -Level SUCCESS
        }
        else {
            Write-Log "Registry path exists" -Level INFO
        }
    }
    catch {
        Write-Log "Failed to create registry path: $($_.Exception.Message)" -Level ERROR
        $iterationSuccess = $false
    }

    # Step 2: Set registry value
    if ($iterationSuccess) {
        try {
            $currentValue = $null
            try {
                $currentValue = (Get-ItemProperty -Path $regPath -Name $regName -ErrorAction Stop).$regName
            }
            catch { }

            Write-Log "Current '$regName' value: $(if ($null -ne $currentValue) { $currentValue } else { '<not set>' })" -Level INFO
            New-ItemProperty -Path $regPath -Name $regName -Value 0 -PropertyType DWORD -Force -ErrorAction Stop | Out-Null
            Write-Log "Set '$regName' = 0" -Level SUCCESS
        }
        catch {
            Write-Log "Failed to set registry value: $($_.Exception.Message)" -Level ERROR
            $iterationSuccess = $false
        }
    }

    # Step 3: Restart service
    if ($iterationSuccess) {
        try {
            $preState = Get-ServiceState -Name $serviceName
            Write-Log "Service state before restart: $preState" -Level INFO

            if ($preState -eq 'Stopped') {
                Write-Log "Service is stopped - starting it" -Level WARN
                Start-Service -Name $serviceName -ErrorAction Stop
            }
            else {
                Write-Log "Restarting '$serviceName'..." -Level INFO
                Restart-Service -Name $serviceName -Force -ErrorAction Stop
            }

            # Wait for Running state
            $running = Wait-ServiceStatus -Name $serviceName -DesiredStatus 'Running' -TimeoutSeconds 30
            if ($running) {
                Write-Log "Service '$serviceName' is Running" -Level SUCCESS
            }
            else {
                $postState = Get-ServiceState -Name $serviceName
                Write-Log "Service did not reach Running state within 30s (current: $postState)" -Level WARN
                $iterationSuccess = $false
            }
        }
        catch {
            Write-Log "Failed to restart service: $($_.Exception.Message)" -Level ERROR
            $iterationSuccess = $false
        }
    }

    # Iteration summary
    if ($iterationSuccess) {
        $successCount++
        Write-Log "Iteration $i completed successfully" -Level SUCCESS
    }
    else {
        $failCount++
        Write-Log "Iteration $i completed with errors" -Level ERROR
    }

    # Delay between iterations
    if ($i -lt $MaxRetries) {
        Write-Log "Waiting ${RetryDelaySeconds}s before next iteration..." -Level INFO
        Start-Sleep -Seconds $RetryDelaySeconds
    }
}

#region Policy Repair (optional)

if ($RepairPolicies) {
    $policiesPath      = "$env:ProgramData\Microsoft\dot3svc\MigrationData\Policies"
    $windowsOldPolicies = "$env:SystemDrive\Windows.old\Windows\dot3svc\Policies"
    $targetPolicies     = "$env:SystemDrive\Windows\dot3svc\Policies"
    $policyRepairDone   = $false

    Write-Banner "Policy Repair"
    Write-Log "Policy repair enabled" -Level INFO
    Write-Log "Policies path:      $policiesPath" -Level INFO
    Write-Log "Windows.old source: $windowsOldPolicies" -Level INFO
    Write-Log "Target path:        $targetPolicies" -Level INFO

    # --- Phase 1: Symlink Detection ---
    if (-not $SkipSymlinkCheck) {
        Write-Log "--- Symlink Check ---" -Level STEP

        if (Test-Path $policiesPath) {
            $policyFiles = Get-ChildItem -Path $policiesPath -File -ErrorAction SilentlyContinue
            if ($policyFiles.Count -eq 0) {
                Write-Log "No files found in '$policiesPath'" -Level WARN
            }
            else {
                $symlinkCount = 0
                $realCount = 0
                foreach ($file in $policyFiles) {
                    $isSymlink = $false
                    # Check ReparsePoint attribute (covers symlinks and junctions)
                    if ($file.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
                        $isSymlink = $true
                    }

                    if ($isSymlink) {
                        $symlinkCount++
                        # Resolve target for logging
                        $linkTarget = $null
                        try {
                            $linkTarget = (Get-Item $file.FullName -ErrorAction Stop).Target
                            if ($linkTarget -is [System.Collections.IEnumerable] -and $linkTarget -isnot [string]) {
                                $linkTarget = $linkTarget[0]
                            }
                        }
                        catch { $linkTarget = '<unable to resolve>' }
                        Write-Log "SYMLINK: $($file.Name) -> $linkTarget" -Level WARN

                        # Check if the symlink target actually exists
                        if ($linkTarget -and $linkTarget -ne '<unable to resolve>') {
                            if (Test-Path $linkTarget) {
                                Write-Log "  Symlink target exists (file is accessible)" -Level INFO
                            }
                            else {
                                Write-Log "  Symlink target MISSING - policy file is broken" -Level ERROR
                            }
                        }
                    }
                    else {
                        $realCount++
                        $sizeKB = [math]::Round($file.Length / 1KB, 1)
                        Write-Log "OK: $($file.Name) ($sizeKB KB)" -Level SUCCESS
                    }
                }
                if ($symlinkCount -gt 0) { $symlinkLevel = 'WARN' } else { $symlinkLevel = 'SUCCESS' }
                Write-Log "Symlink check result: $realCount real file(s), $symlinkCount symlink(s) out of $($policyFiles.Count) total" -Level $symlinkLevel
            }
        }
        else {
            Write-Log "Policies folder not found: '$policiesPath'" -Level WARN
        }
    }
    else {
        Write-Log "Symlink check skipped (-SkipSymlinkCheck)" -Level INFO
    }

    # --- Phase 2: Windows.old Policy Recovery ---
    Write-Log "--- Windows.old Policy Recovery ---" -Level STEP

    if (Test-Path $windowsOldPolicies) {
        $oldPolicyFiles = Get-ChildItem -Path $windowsOldPolicies -Filter '*.tmp' -File -ErrorAction SilentlyContinue
        if ($oldPolicyFiles.Count -eq 0) {
            Write-Log "No .tmp policy files found in '$windowsOldPolicies'" -Level WARN
        }
        else {
            Write-Log "Found $($oldPolicyFiles.Count) policy file(s) in Windows.old:" -Level INFO
            foreach ($f in $oldPolicyFiles) {
                $sizeKB = [math]::Round($f.Length / 1KB, 1)
                Write-Log "  $($f.Name) ($sizeKB KB, Modified: $($f.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')))" -Level INFO
            }

            # Ensure target directory exists
            if (-not (Test-Path $targetPolicies)) {
                try {
                    New-Item -Path $targetPolicies -ItemType Directory -Force -ErrorAction Stop | Out-Null
                    Write-Log "Created target directory: '$targetPolicies'" -Level SUCCESS
                }
                catch {
                    Write-Log "Failed to create target directory '$targetPolicies': $($_.Exception.Message)" -Level ERROR
                }
            }

            if (Test-Path $targetPolicies) {
                $copiedCount = 0
                $copyFailCount = 0
                foreach ($f in $oldPolicyFiles) {
                    $destFile = Join-Path $targetPolicies $f.Name
                    try {
                        # If a symlink or file already exists at destination, remove it first
                        if (Test-Path $destFile) {
                            $existingItem = Get-Item $destFile -ErrorAction SilentlyContinue
                            if ($existingItem.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
                                Write-Log "Removing existing symlink at destination: $($f.Name)" -Level WARN
                            }
                            else {
                                Write-Log "Overwriting existing file at destination: $($f.Name)" -Level WARN
                            }
                            Remove-Item $destFile -Force -ErrorAction Stop
                        }
                        Copy-Item -Path $f.FullName -Destination $destFile -Force -ErrorAction Stop
                        $copiedCount++
                        Write-Log "Copied: $($f.Name) -> $destFile" -Level SUCCESS
                    }
                    catch {
                        $copyFailCount++
                        Write-Log "Failed to copy '$($f.Name)': $($_.Exception.Message)" -Level ERROR
                    }
                }
                if ($copyFailCount -eq 0) { $copyLevel = 'SUCCESS' } else { $copyLevel = 'WARN' }
                Write-Log "Copy result: $copiedCount succeeded, $copyFailCount failed" -Level $copyLevel

                if ($copiedCount -gt 0) {
                    $policyRepairDone = $true
                }
                if ($copyFailCount -gt 0) {
                    $exitCode = 1
                }
            }
        }
    }
    else {
        Write-Log "Windows.old policies folder not found: '$windowsOldPolicies' (no previous OS installation or already cleaned up)" -Level INFO
    }

    # --- Phase 3: Restart service after policy recovery ---
    if ($policyRepairDone) {
        Write-Log "--- Post-repair service restart ---" -Level STEP
        try {
            Restart-Service -Name $serviceName -Force -ErrorAction Stop
            $running = Wait-ServiceStatus -Name $serviceName -DesiredStatus 'Running' -TimeoutSeconds 30
            if ($running) {
                Write-Log "Service '$serviceName' restarted after policy repair" -Level SUCCESS
            }
            else {
                $postState = Get-ServiceState -Name $serviceName
                Write-Log "Service did not reach Running state after repair (current: $postState)" -Level WARN
                $exitCode = 1
            }
        }
        catch {
            Write-Log "Failed to restart service after policy repair: $($_.Exception.Message)" -Level ERROR
            $exitCode = 1
        }
    }
}

#endregion

# Final summary
Write-Banner "Summary"
if ($failCount -eq 0) { $iterLevel = 'SUCCESS' } else { $iterLevel = 'WARN' }
Write-Log "Iterations: $MaxRetries total, $successCount succeeded, $failCount failed" -Level $iterLevel
$finalState = Get-ServiceState -Name $serviceName
if ($finalState -eq 'Running') { $stateLevel = 'SUCCESS' } else { $stateLevel = 'WARN' }
Write-Log "Service '$serviceName' final state: $finalState" -Level $stateLevel
if ($RepairPolicies) {
    if ($policyRepairDone) { $repairMsg = 'policies recovered from Windows.old'; $repairLevel = 'SUCCESS' }
    else { $repairMsg = 'no recovery needed or no source found'; $repairLevel = 'INFO' }
    Write-Log "Policy repair: $repairMsg" -Level $repairLevel
}
Write-Log "Log file: $Script:LogFile" -Level INFO
Write-Log "Script finished" -Level INFO

if ($failCount -gt 0) { $exitCode = 1 }
exit $exitCode

#endregion
