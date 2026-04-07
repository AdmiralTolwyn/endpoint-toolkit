<#
.SYNOPSIS
    Removes residual Teams cache artifacts from FSLogix profile containers.

.DESCRIPTION
    When FSLogix Redirections.xml exclusions are added after profiles already
    exist, stale data remains inside the VHD(x). This script deletes the
    known Teams (classic + new) cache paths that should now live outside the
    profile container on the local disk.

    Intended to run as a scheduled task at user logon or logoff, or via GPO
    user logon/logoff script.

    Targeted paths (matching Exclude Copy="0" entries):
      - AppData\Roaming\Microsoft\Teams\media-stack
      - AppData\Local\Microsoft\Teams\meeting-addin\Cache
      - AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\Logs
      - AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\PerfLogs
      - AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\EBWebView\WV2Profile_tfw\WebStorage

.PARAMETER LogDir
    Directory for the cleanup log file. Defaults to $env:TEMP.

.PARAMETER MaxRetries
    Number of retry attempts per path after stopping Teams. Default: 3.

.PARAMETER RetryDelaySeconds
    Seconds to wait between retries. Default: 5.

.PARAMETER WhatIf
    When specified, logs what would be removed without deleting anything.

.NOTES
    File:     scripts/Remove-FSLogixTeamsArtifacts.ps1
    Context:  Runs as the logged-on user (user context)
    Trigger:  User logon or logoff scheduled task / GPO script

.EXAMPLE
    # Logon script (GPO or scheduled task, user context)
    powershell.exe -ExecutionPolicy Bypass -NoProfile -File "Remove-FSLogixTeamsArtifacts.ps1"

.EXAMPLE
    # Dry-run to see what would be cleaned up
    .\Remove-FSLogixTeamsArtifacts.ps1 -WhatIf
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$LogDir  = $env:TEMP,
    [int]$MaxRetries = 3,
    [int]$RetryDelaySeconds = 5
)

$ErrorActionPreference = "Stop"

# -----------------------------------------------------------------------------
# LOGGING
# -----------------------------------------------------------------------------
$LogFile = Join-Path $LogDir "FSLogix-TeamsCleanup_$(Get-Date -Format 'yyyyMMdd').log"

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS')]
        [string]$Level = 'INFO'
    )
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "[$ts] [$Level] $Message"
    Add-Content -Path $LogFile -Value $entry -ErrorAction SilentlyContinue
    switch ($Level) {
        'ERROR'   { Write-Host $entry -ForegroundColor Red }
        'WARN'    { Write-Host $entry -ForegroundColor Yellow }
        'SUCCESS' { Write-Host $entry -ForegroundColor Green }
        default   { Write-Host $entry -ForegroundColor Gray }
    }
}

# -----------------------------------------------------------------------------
# RELATIVE PATHS TO CLEAN (under $env:USERPROFILE)
# These match the Redirections.xml Exclude Copy="0" entries.
# -----------------------------------------------------------------------------
$RelativePaths = @(
    # Classic Teams
    'AppData\Roaming\Microsoft\Teams\media-stack'
    'AppData\Local\Microsoft\Teams\meeting-addin\Cache'
    # New Teams (MSIX)
    'AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\Logs'
    'AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\PerfLogs'
    'AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\EBWebView\WV2Profile_tfw\WebStorage'
)

# -----------------------------------------------------------------------------
# STOP TEAMS PROCESSES
# -----------------------------------------------------------------------------
function Stop-TeamsProcesses {
    # Classic Teams = "Teams", New Teams = "ms-teams"
    $teamsProcs = Get-Process -Name 'Teams', 'ms-teams' -ErrorAction SilentlyContinue |
                  Where-Object { $_.Path -match 'MSTeams|Teams' }

    if (-not $teamsProcs) { return $false }

    Write-Log "Teams processes detected ($($teamsProcs.Count)): $($teamsProcs.Name -join ', ')" -Level WARN

    foreach ($proc in $teamsProcs) {
        try {
            $proc.Kill()
            Write-Log "Stopped: $($proc.Name) (PID $($proc.Id))"
        }
        catch {
            Write-Log "Could not stop $($proc.Name) (PID $($proc.Id)): $($_.Exception.Message)" -Level WARN
        }
    }

    # Wait for processes to fully exit
    $timeout = [datetime]::Now.AddSeconds(15)
    while ([datetime]::Now -lt $timeout) {
        $remaining = Get-Process -Name 'Teams', 'ms-teams' -ErrorAction SilentlyContinue |
                     Where-Object { $_.Path -match 'MSTeams|Teams' }
        if (-not $remaining) { break }
        Start-Sleep -Milliseconds 500
    }

    return $true
}

# -----------------------------------------------------------------------------
# MAIN
# -----------------------------------------------------------------------------
Write-Log "=== FSLogix Teams artifact cleanup started ==="
Write-Log "User: $env:USERNAME | Profile: $env:USERPROFILE"

$TeamsWasRunning = Stop-TeamsProcesses
if ($TeamsWasRunning) {
    Write-Log "Teams terminated - waiting ${RetryDelaySeconds}s for file locks to release"
    Start-Sleep -Seconds $RetryDelaySeconds
}

$RemovedCount   = 0
$SkippedCount   = 0
$ErrorCount     = 0
$ReclaimedBytes = 0

foreach ($relPath in $RelativePaths) {
    $fullPath = Join-Path $env:USERPROFILE $relPath

    if (-not (Test-Path $fullPath)) {
        Write-Log "SKIP (not found): $relPath"
        $SkippedCount++
        continue
    }

    # Measure size before removal
    $dirSize = (Get-ChildItem -Path $fullPath -Recurse -Force -ErrorAction SilentlyContinue |
                Measure-Object -Property Length -Sum).Sum
    $sizeMB  = [math]::Round(($dirSize / 1MB), 2)

    $removeMsg = 'Remove directory ({0} MB)' -f $sizeMB
    if (-not $PSCmdlet.ShouldProcess($fullPath, $removeMsg)) {
        $whatifMsg = 'WHATIF: Would remove {0} ({1} MB)' -f $relPath, $sizeMB
        Write-Log $whatifMsg
        $SkippedCount++
        continue
    }

    $removed = $false
    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            Remove-Item -Path $fullPath -Recurse -Force -ErrorAction Stop
            $doneMsg = 'REMOVED: {0} ({1} MB)' -f $relPath, $sizeMB
            Write-Log $doneMsg -Level SUCCESS
            $RemovedCount++
            $ReclaimedBytes += $dirSize
            $removed = $true
            break
        }
        catch {
            if ($attempt -lt $MaxRetries) {
                Write-Log "Attempt $attempt/$MaxRetries failed for $relPath - $($_.Exception.Message). Retrying in ${RetryDelaySeconds}s..." -Level WARN
                Start-Sleep -Seconds $RetryDelaySeconds
            }
            else {
                Write-Log "FAILED after $MaxRetries attempts: $relPath - $($_.Exception.Message)" -Level ERROR
                $ErrorCount++
            }
        }
    }
}

$totalMB = [math]::Round(($ReclaimedBytes / 1MB), 2)
$summary = '=== Cleanup complete: Removed={0}, Skipped={1}, Errors={2}, Reclaimed={3} MB ===' -f $RemovedCount, $SkippedCount, $ErrorCount, $totalMB
Write-Log $summary

exit $ErrorCount
