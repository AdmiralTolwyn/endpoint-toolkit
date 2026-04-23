#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Expands the C: partition to consume all unallocated disk space after a Windows 365 resize.

.DESCRIPTION
    Intended to run as a System-context Intune platform script.
    Uses diskpart.exe instead of PowerShell Storage cmdlets to avoid potential BSOD issues
    with Resize-Partition / Remove-Partition on virtual disks.

    Workflow:
      1. Detect the disk and partition layout for C:
      2. If a Recovery (WinRE) partition exists immediately after C:, disable WinRE first, then delete the partition
      3. Extend C: into all available unallocated space via diskpart
      4. Validate and report

.NOTES
    Version : 1.0.1
    Author  : Anton Romanyuk

.DISCLAIMER
    This script is provided "AS IS" with no warranties and confers no rights.
    It is not supported under any Microsoft standard support program or service.
    Use of this script is entirely at your own risk. The customer is solely
    responsible for testing and validating this script in their environment
    before deploying to production. The author shall not be liable for any
    damage or data loss resulting from the use of this script.
#>

[CmdletBinding()]
param()

#region --- Configuration ---
$LogFolder   = Join-Path $env:ProgramData 'Microsoft\IntuneManagementExtension\Logs'
$LogFile     = Join-Path $LogFolder ("DiskResizer_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
$MinUnallocatedGB = 0.1   # Minimum unallocated space (GB) before attempting resize
#endregion

#region --- Helpers ---
function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR')][string]$Level = 'INFO'
    )
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "[$ts] [$Level] $Message"
    Add-Content -Path $LogFile -Value $entry -ErrorAction SilentlyContinue
    switch ($Level) {
        'ERROR' { Write-Error   "[$Level] $Message" }
        'WARN'  { Write-Warning "[$Level] $Message" }
        default { Write-Output  "[$Level] $Message" }
    }
}

function Invoke-DiskPart {
    <#
    .SYNOPSIS
        Runs a diskpart script and returns output + exit code.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$Commands
    )
    $scriptFile = Join-Path $env:TEMP "diskpart_$(New-Guid).txt"
    try {
        $Commands | Set-Content -Path $scriptFile -Encoding ASCII
        $result = & diskpart.exe /s $scriptFile 2>&1
        $exitCode = $LASTEXITCODE
        Write-Log "diskpart output:`n$($result | Out-String)" -Level INFO
        if ($exitCode -ne 0) {
            Write-Log "diskpart exited with code $exitCode" -Level ERROR
        }
        return [PSCustomObject]@{
            Output   = $result
            ExitCode = $exitCode
        }
    }
    finally {
        Remove-Item -Path $scriptFile -Force -ErrorAction SilentlyContinue
    }
}
#endregion

#region --- Main ---
try {
    # Ensure log directory exists
    if (-not (Test-Path $LogFolder)) {
        New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null
    }

    Start-Transcript -Path $LogFile -Append -ErrorAction SilentlyContinue
    Write-Log '=== DiskResizer v2 started ==='

    # --- Discover the OS partition on disk 0 via WMI ---
    # The Storage module (Get-Partition/Get-Volume/Get-Disk) is unreliable on W365 VMs
    # where MSFT_Partition/MSFT_Volume classes are empty. Win32_* classes always work.
    $diskNumber = 0
    $wmiDisk  = Get-CimInstance Win32_DiskDrive -Filter "Index = $diskNumber" -ErrorAction Stop
    $wmiParts = Get-CimInstance Win32_DiskPartition -Filter "DiskIndex = $diskNumber" -ErrorAction Stop

    Write-Log "Disk $diskNumber : $($wmiDisk.Caption) | Total: $([math]::Round($wmiDisk.Size / 1GB, 2)) GB"

    # Log partition layout
    foreach ($wp in ($wmiParts | Sort-Object Index)) {
        $sizeStr = if ($wp.Size -ge 1GB) { '{0:N2} GB' -f ($wp.Size / 1GB) } else { '{0:N0} MB' -f ($wp.Size / 1MB) }
        Write-Log ("  Partition #{0}: {1} | {2} | Boot: {3} | Primary: {4}" -f
            ($wp.Index + 1), $wp.Name, $sizeStr, $wp.BootPartition, $wp.PrimaryPartition)
    }

    # OS partition = the largest partition on disk 0
    $osWmiPart = $wmiParts | Sort-Object Size -Descending | Select-Object -First 1
    if (-not $osWmiPart) {
        Write-Log "No partitions found on disk $diskNumber. Cannot proceed." -Level ERROR
        Stop-Transcript -ErrorAction SilentlyContinue
        exit 1
    }
    # Win32_DiskPartition.Index is 0-based; diskpart partition numbers are 1-based
    $osPartNum = $osWmiPart.Index + 1
    $osSizeGB  = [math]::Round($osWmiPart.Size / 1GB, 2)
    Write-Log "OS partition is #$osPartNum ($($osWmiPart.Name), $osSizeGB GB)"

    # --- Check for Recovery partition via diskpart ---
    # WMI partition indices don't map 1:1 to diskpart partition numbers because
    # MSR (Microsoft Reserved) partitions are hidden from WMI but counted by diskpart.
    # Use 'diskpart list partition' to find the actual Recovery partition number.
    Write-Log 'Scanning for Recovery partition via diskpart...'
    $dpList = Invoke-DiskPart -Commands @(
        "select disk $diskNumber"
        "list partition"
    )
    $listText = $dpList.Output | Out-String

    # Parse "list partition" output to find a Recovery-type partition
    # Format: "  Partition N    Recovery    NNN MB   Offset"
    # Also handle locale variations and OEM type markers
    $recoveryMatch = [regex]::Match($listText, 'Partition\s+(\d+)\s+Recovery', 'IgnoreCase')

    if ($recoveryMatch.Success) {
        $recoveryPartNum = [int]$recoveryMatch.Groups[1].Value
        Write-Log "Recovery partition found at diskpart partition #$recoveryPartNum. Proceeding to remove."

        # Step 1: Disable WinRE FIRST
        Write-Log 'Disabling WinRE before removing Recovery partition...'
        try {
            # Temporarily silence ErrorActionPreference so stderr from native
            # commands doesn't surface as red NativeCommandError text in the console
            $prevEAP = $ErrorActionPreference
            $ErrorActionPreference = 'SilentlyContinue'
            $reagentResult = & reagentc.exe /disable 2>&1
            $reagentExit   = $LASTEXITCODE
            $ErrorActionPreference = $prevEAP
            $reagentText   = ($reagentResult | Out-String).Trim()
            Write-Log "reagentc /disable output (exit $reagentExit): $reagentText"

            # Exit code 2 with "already disabled" is fine — not a real failure
            if ($reagentExit -ne 0) {
                if ($reagentText -match 'already disabled') {
                    Write-Log 'WinRE was already disabled. Safe to continue.' -Level WARN
                }
                else {
                    Write-Log "reagentc /disable failed (exit code $reagentExit). Aborting partition deletion to avoid BCD corruption." -Level ERROR
                    Stop-Transcript -ErrorAction SilentlyContinue
                    exit 1
                }
            }
        }
        catch {
            Write-Log "reagentc /disable threw an exception: $($_.Exception.Message)" -Level ERROR
            Stop-Transcript -ErrorAction SilentlyContinue
            exit 1
        }

        # Step 2: Delete Recovery partition via diskpart
        Write-Log "Deleting Recovery partition #$recoveryPartNum via diskpart..."
        $dpResult = Invoke-DiskPart -Commands @(
            "select disk $diskNumber"
            "select partition $recoveryPartNum"
            "delete partition override"
        )
        if ($dpResult.ExitCode -ne 0) {
            Write-Log "Failed to delete Recovery partition via diskpart." -Level ERROR
            Stop-Transcript -ErrorAction SilentlyContinue
            exit 1
        }
        Write-Log 'Recovery partition deleted successfully.'
    }
    else {
        Write-Log "No Recovery partition found in diskpart listing. Nothing to remove."
    }

    # --- Calculate unallocated space via WMI ---
    # Re-query after potential deletion
    $wmiPartsNow = Get-CimInstance Win32_DiskPartition -Filter "DiskIndex = $diskNumber"
    $allocated   = ($wmiPartsNow | Measure-Object -Property Size -Sum).Sum
    $unallocatedBytes = $wmiDisk.Size - $allocated
    $unallocatedGB    = [math]::Round($unallocatedBytes / 1GB, 2)
    $diskTotalGB      = [math]::Round($wmiDisk.Size / 1GB, 2)

    Write-Log "Disk $diskNumber total: $diskTotalGB GB | Allocated: $([math]::Round($allocated / 1GB, 2)) GB | Unallocated: $unallocatedGB GB"

    if ($unallocatedGB -gt $MinUnallocatedGB) {
        Write-Log "Unallocated space ($unallocatedGB GB) exceeds threshold ($MinUnallocatedGB GB). Extending OS partition..."

        # Ensure defragsvc is startable (skip smphost — intentionally disabled per security policy)
        $svc = Get-Service -Name 'defragsvc' -ErrorAction SilentlyContinue
        if ($svc -and $svc.StartType -eq 'Disabled') {
            Set-Service -Name 'defragsvc' -StartupType Manual
            Write-Log "Set defragsvc startup to Manual (was Disabled)"
        }

        # Extend C: volume via diskpart
        # Use 'select volume C' (not 'select partition') — diskpart extend
        # requires a volume context, and partition selection fails with
        # E_INVALIDARG on some W365/Azure virtual disk configurations.
        $dpResult = Invoke-DiskPart -Commands @(
            "select volume C"
            "extend"
        )

        if ($dpResult.ExitCode -ne 0) {
            Write-Log "diskpart extend failed." -Level ERROR
            Stop-Transcript -ErrorAction SilentlyContinue
            exit 1
        }

        # --- Post-resize validation ---
        Start-Sleep -Seconds 2
        $wmiPartsPost  = Get-CimInstance Win32_DiskPartition -Filter "DiskIndex = $diskNumber"
        $allocatedPost = ($wmiPartsPost | Measure-Object -Property Size -Sum).Sum
        $remainingGB   = [math]::Round(($wmiDisk.Size - $allocatedPost) / 1GB, 2)
        # Re-discover OS partition as the largest one (indices may have shifted
        # after Recovery partition deletion)
        $newOsPart     = $wmiPartsPost | Sort-Object Size -Descending | Select-Object -First 1
        $newSizeGB     = if ($newOsPart) { [math]::Round($newOsPart.Size / 1GB, 2) } else { 0 }

        Write-Log "OS partition expanded to $newSizeGB GB (disk total: $diskTotalGB GB)"

        if ($remainingGB -gt $MinUnallocatedGB) {
            Write-Log "WARNING: $remainingGB GB still unallocated after extend." -Level WARN
        }
        else {
            Write-Log 'Resize verified successfully. No significant unallocated space remaining.'
        }
    }
    else {
        Write-Log "Unallocated space ($unallocatedGB GB) is below threshold ($MinUnallocatedGB GB). No resize needed."
    }

    Write-Log '=== DiskResizer v2 completed successfully ==='
    Stop-Transcript -ErrorAction SilentlyContinue
    exit 0
}
catch {
    Write-Log "Unhandled exception: $($_.Exception.Message)" -Level ERROR
    Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level ERROR
    Stop-Transcript -ErrorAction SilentlyContinue
    exit 1
}
#endregion
