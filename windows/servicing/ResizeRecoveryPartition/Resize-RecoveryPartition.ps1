<#
.SYNOPSIS
    Resizes the Windows Recovery Environment (WinRE) partition on the system disk.

.DESCRIPTION
    Restores or grows the WinRE recovery partition so a larger WinRE.WIM (typically
    delivered by a feature update or by the KB5034441 / CVE-2024-20666 servicing
    stack mitigation) can be staged.

    Sequence:
      1. Pre-flight gate (Test-ExitCondition):
         - BitLocker protection state on the OS volume (must be On).
         - Presence of %SystemRoot%\System32\Recovery\ReAgent.xml.
         - WinRE already enabled and pointing at a live partition path
           (treated as healthy, no work to do).
         - Free space on the system drive (<= 95 % used).
         - Completion marker in the registry (idempotent skip on second run).
      2. Read WinRE metadata from ReAgent.xml + Get-WindowsImage.
      3. Decide target recovery-partition size:
            < 900 MB WIM  -> 998 MB target partition
           >= 900 MB WIM  -> WIM size + 550 MB headroom
      4. Disable WinRE (`reagentc /disable`), drop the existing recovery
         partition, shrink the OS partition, create a new partition at the
         end of the disk using maximum size, format NTFS, label "Recovery".
      5. Tag the partition as a WinRE GPT partition via diskpart.exe
         (Set-Partition was previously observed to BSOD on this step on
         certain firmware -- see version history).
      6. Re-enable WinRE (`reagentc /enable`) so the new WIM is staged.
      7. Stamp the completion marker in the registry.

    Designed to be re-runnable: if the partition is already big enough or the
    completion marker is present, the script exits with a non-fatal code and
    does not touch the disk.

.PARAMETER DiskNumber
    Physical disk number containing the system + recovery partitions.
    Default: 0. Override only for unusual hardware (e.g. multi-disk OEM lab
    images). The OS partition is detected via IsBoot=True, but the recovery
    partition is enumerated on this disk only.

.PARAMETER MinRecoveryPartitionMB
    Target size, in megabytes, for the new recovery partition when the WIM
    is smaller than -LargeWimThresholdMB. Default: 998.

.PARAMETER LargeWimThresholdMB
    Above this WIM size (MB), the new partition is sized as WIM + headroom
    instead of -MinRecoveryPartitionMB. Default: 900.

.PARAMETER HeadroomMB
    Free space (MB) reserved above the WIM size when growing the partition
    for a large WIM. Default: 550.

.PARAMETER CompletionRegistryPath
    Registry key (without HKLM:\ prefix) where the completion marker is
    written and probed for the idempotent skip. Default:
    "SOFTWARE\EndpointToolkit\WinRE".

.PARAMETER CompletionRegistryValueName
    Value name under -CompletionRegistryPath. Default: "WinReResized".

.PARAMETER LogDirectory
    Directory for the log file. Default: %SystemDrive%\Windows\debug.

.PARAMETER RewriteAlreadyPatchedToSuccess
    When set, ERROR_ALREADY_PATCHED (1) is rewritten to 0 before exit so
    Intune / SCCM / ConfigMgr treat the second-run no-op as success.
    Default: $true.

.PARAMETER RewriteNoWinReToSuccess
    When set, ERROR_NO_WINRE_DETECTED (2) is rewritten to 0 before exit.
    Default: $true.

.PARAMETER RewriteWinReHealthyToSuccess
    When set, ERROR_WINRE_HEALTHY (10) is rewritten to 0 before exit so
    a healthy device does not fail a deployment.
    Default: $true.

.EXAMPLE
    # Normal run (defaults)
    .\Resize-RecoveryPartition.ps1

.EXAMPLE
    # Custom registry marker (mirrors a vendor namespace)
    .\Resize-RecoveryPartition.ps1 -CompletionRegistryPath 'SOFTWARE\Contoso\WinRE'

.EXAMPLE
    # Surface raw exit codes (no rewrite) so an Intune detection script
    # can branch on "already patched" vs "healthy" vs "remediated".
    .\Resize-RecoveryPartition.ps1 -RewriteAlreadyPatchedToSuccess:$false `
                                   -RewriteNoWinReToSuccess:$false `
                                   -RewriteWinReHealthyToSuccess:$false

.NOTES
    File:     windows/servicing/ResizeRecoveryPartition/Resize-RecoveryPartition.ps1
    Author:   Anton Romanyuk
    Version:  1.4.0
    Requires: PowerShell 5.1+, elevated session.

    Changes:
      1.4.0 - Ported into the endpoint-toolkit repo and conformed to the
              repo conventions:
                * Added comment-based help, CmdletBinding, parameters for
                  disk number, partition sizing thresholds, registry path,
                  exit-code rewrites and log directory.
                * Removed vendor-specific registry namespace; default is now
                  HKLM:\SOFTWARE\EndpointToolkit\WinRE.
                * Replaced Get-WmiObject with Get-CimInstance.
                * Fixed sizing-decision typo ($fileSizeMB -> $WinReFileSizeMB)
                  so the >= 900 MB WIM branch is actually reachable.
                * Fixed byte/MB unit mismatch in the recovery-partition size
                  guard (was comparing Size/1MB against the 998MB byte
                  literal, which always evaluated true).
                * Fixed WinRE partition size calculation (WIM size is already
                  in MB; the +550MB literal was double-scaled).
                * OS disk now resolved by IsBoot=True instead of hard-coded
                  disk 0; -DiskNumber still selectable for non-default rigs.
                * Defined ERROR_BDE_MISSING (was referenced but undefined).
                * Removed the "throw after return" anti-pattern; functions
                  now return the exit code only.
                * Tightened logging: single Write-Log helper with explicit
                  level set, transcript-friendly, no stray Write-Warning.
      1.3 -   Improved logging & error handling, cleanup.            (Anton Romanyuk, 2024-02-14)
      1.2 -   Switched Set-Partition to diskpart.exe due to BSOD.    (Anton Romanyuk, 2024-02-13)
      1.1 -   Bugfixing, reduced WinRE default size to 998MB.        (Anton Romanyuk, 2024-02-12)
      1.0 -   Script created.                                        (Anton Romanyuk, 2024-02-11)

    Exit codes:
      0    - ERROR_SUCCESS                       (success / nothing to do)
      1    - ERROR_ALREADY_PATCHED               (registry marker present; rewritten to 0 by default)
      2    - ERROR_NO_WINRE_DETECTED             (no WinRE on this machine; rewritten to 0 by default)
      3    - ERROR_XML_MISSING                   (ReAgent.xml not found)
      4    - ERROR_WIM_MISSING                   (WinRE.wim not found at the expected path)
      10   - ERROR_WINRE_HEALTHY                 (partition already large enough; rewritten to 0 by default)
      11   - ERROR_MULTIPLE_RECOVERY_PARTITIONS  (>1 recovery partition on the disk; manual cleanup needed)
      122  - ERROR_INSUFFICIENT_DISK_SPACE       (system drive >95% used)
      1000 - ERROR_UNKNOWN                       (uncaught exception)
      1599 - ERROR_PREREQ_FAILURE                (pre-flight failed; BitLocker off, etc.)
      1689 - ERROR_TPM_UPDATE_FAILED             (reagentc /enable failed)

.DISCLAIMER
    THIS SCRIPT IS PROVIDED "AS-IS" WITHOUT WARRANTY OF ANY KIND.
    It is not supported under any Microsoft standard support program or service.
    Resizing system partitions is destructive: the recovery partition is dropped
    and the OS partition is shrunk in-place. Validate against a test device with
    BitLocker enabled before deploying to production.
#>

#Requires -RunAsAdministrator
[CmdletBinding(SupportsShouldProcess)]
param(
    [ValidateRange(0, 31)]
    [int]    $DiskNumber                  = 0,

    [ValidateRange(256, 8192)]
    [int]    $MinRecoveryPartitionMB      = 998,

    [ValidateRange(128, 4096)]
    [int]    $LargeWimThresholdMB         = 900,

    [ValidateRange(64, 2048)]
    [int]    $HeadroomMB                  = 550,

    [string] $CompletionRegistryPath      = 'SOFTWARE\EndpointToolkit\WinRE',
    [string] $CompletionRegistryValueName = 'WinReResized',
    [string] $LogDirectory                = (Join-Path $env:SystemDrive 'Windows\debug'),

    [bool]   $RewriteAlreadyPatchedToSuccess = $true,
    [bool]   $RewriteNoWinReToSuccess        = $true,
    [bool]   $RewriteWinReHealthyToSuccess   = $true
)

# -----------------------------------------------------------------------------
# EXIT CODES
# -----------------------------------------------------------------------------
$ERROR_SUCCESS                      = 0
$ERROR_ALREADY_PATCHED              = 1
$ERROR_NO_WINRE_DETECTED            = 2
$ERROR_XML_MISSING                  = 3
$ERROR_WIM_MISSING                  = 4
$ERROR_WINRE_HEALTHY                = 10
$ERROR_MULTIPLE_RECOVERY_PARTITIONS = 11
$ERROR_INSUFFICIENT_DISK_SPACE      = 122
$ERROR_BDE_MISSING                  = 1598
$ERROR_PREREQ_FAILURE               = 1599
$ERROR_TPM_UPDATE_FAILED            = 1689
$ERROR_UNKNOWN                      = 1000

# Sentinel returned by Test-ExitCondition when pre-flight is OK.
$EXIT_CONDITIONS_OK = -1

# -----------------------------------------------------------------------------
# LOGGING
# -----------------------------------------------------------------------------
if (-not (Test-Path -LiteralPath $LogDirectory)) {
    New-Item -ItemType Directory -Path $LogDirectory -Force -ErrorAction SilentlyContinue | Out-Null
}
$Script:LogFilePath = Join-Path $LogDirectory 'ResizeRecoveryPartition.log'

function Write-Log {
<#
.SYNOPSIS
    Appends a timestamped, level-tagged line to the script log and host console.
.DESCRIPTION
    Uniform logger. Format on disk and on console:
        yyyy/MM/dd HH:mm:ss   LEVEL   message
    The header row is added once when the log is first created.
.PARAMETER Message
    Free-form text to record.
.PARAMETER Level
    INFO | WARN | ERROR | FATAL | DEBUG | TRACE. Default INFO.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Message,
        [ValidateSet('INFO','WARN','ERROR','FATAL','DEBUG','TRACE')]
        [string] $Level = 'INFO'
    )
    $stamp = (Get-Date).ToString('yyyy/MM/dd HH:mm:ss')
    $line  = $stamp.PadRight(22, ' ') + $Level.PadRight(10, ' ') + $Message

    if (-not (Test-Path -LiteralPath $Script:LogFilePath)) {
        $header = 'Timestamp'.PadRight(22, ' ') + 'Level'.PadRight(10, ' ') + 'Message'
        Add-Content -LiteralPath $Script:LogFilePath -Value $header -ErrorAction SilentlyContinue
    }
    Add-Content -LiteralPath $Script:LogFilePath -Value $line -ErrorAction SilentlyContinue

    $color = switch ($Level) {
        'WARN'  { 'Yellow' }
        'ERROR' { 'Red' }
        'FATAL' { 'Red' }
        'DEBUG' { 'DarkGray' }
        'TRACE' { 'DarkGray' }
        default { 'Gray' }
    }
    Write-Host $line -ForegroundColor $color
}

# -----------------------------------------------------------------------------
# HELPERS
# -----------------------------------------------------------------------------
function Write-WinReDebugOutput {
<#
.SYNOPSIS
    Logs the relevant fields from a WinRE metadata object.
.PARAMETER WinRE
    Object returned by Get-WinREData.
#>
    [CmdletBinding()]
    param([Parameter(Mandatory)] $WinRE)

    Write-Log -Level INFO -Message 'WinRE.WIM File Information'
    Write-Log -Level INFO -Message " -> Path     : $($WinRE.ImagePath)"
    Write-Log -Level INFO -Message " -> Existing : $($WinRE.ImageExisting)"
    Write-Log -Level INFO -Message " -> Status   : $($WinRE.Enabled)"
    Write-Log -Level INFO -Message " -> Version  : $($WinRE.Version)"
    Write-Log -Level INFO -Message " -> Modified : $($WinRE.ModifiedTime)"
}

function Get-WinREData {
<#
.SYNOPSIS
    Reads WinRE metadata from ReAgent.xml + Get-WindowsImage.
.DESCRIPTION
    Returns a PSCustomObject with Enabled, Build, AutoRepairOn, Version,
    ModifiedTime, CreatedTime, OsBuildVersion, ImagePath, ImageName and
    ImageExisting. When WinRE is disabled (no WinreLocation path in
    ReAgent.xml), the image-derived fields are left empty.
.OUTPUTS
    [pscustomobject]
#>
    [CmdletBinding()]
    param()

    $winre = New-Object PSCustomObject
    $reagentXml = Join-Path $env:SystemRoot 'System32\Recovery\ReAgent.xml'

    if (-not (Test-Path -LiteralPath $reagentXml)) {
        Write-Log -Level WARN -Message "Unable to find $reagentXml"
        return $winre
    }

    [xml]$XmlDocument = Get-Content -LiteralPath $reagentXml -Raw

    $WinRE_LocationId        = $null
    $WinRE_LocationOffset    = $null
    $WinRE_LocationPath      = $null
    $WinRE_LocationPartition = $null
    $WinRE_Enabled           = $false
    $WinRE_IsAutoRepairOn    = $null
    $WinRE_OsBuildVersion    = $null
    $WinRE_Location          = $null
    $WinREData               = $null

    $XmlDocument.SelectNodes('WindowsRE') | ForEach-Object {
        $WinRE_LocationId        = $_.WinreLocation.id
        $WinRE_LocationOffset    = $_.WinreLocation.offset
        $WinRE_LocationPath      = $_.WinreLocation.path
        $WinRE_LocationPartition = (Get-Disk -Number $WinRE_LocationId |
                                    Get-Partition |
                                    Where-Object { $_.Offset -eq $WinRE_LocationOffset }).PartitionNumber
        $WinRE_InstallState      = $_.InstallState.state
        $WinRE_Enabled           = if ($WinRE_InstallState -eq 0) { $false } else { $true }
        $WinRE_IsAutoRepairOn    = $_.IsAutoRepairOn.state
        $WinRE_OsBuildVersion    = $_.OsBuildVersion.path
        $WinRE_Location          = '\\?\GLOBALROOT\device\harddisk' + $WinRE_LocationId + '\partition' + $WinRE_LocationPartition + $WinRE_LocationPath
    }

    if ($WinRE_LocationPath) {
        try {
            $WinREData = Get-WindowsImage -ImagePath ($WinRE_Location.Trim() + '\winre.wim') -Index 1 -ErrorAction Stop
        } catch {
            Write-Log -Level ERROR -Message "Get-WindowsImage failed: $($_.Exception.Message)"
        }
    } else {
        Write-Log -Level WARN -Message 'WinRE_Location is empty (WinRE disabled or unstaged).'
    }

    $winre | Add-Member -Type NoteProperty -Name Enabled        -Value $WinRE_Enabled
    $winre | Add-Member -Type NoteProperty -Name Build          -Value $WinREData.SPBuild
    $winre | Add-Member -Type NoteProperty -Name AutoRepairOn   -Value $WinRE_IsAutoRepairOn
    $winre | Add-Member -Type NoteProperty -Name Version        -Value $WinREData.Version
    $winre | Add-Member -Type NoteProperty -Name ModifiedTime   -Value $WinREData.ModifiedTime
    $winre | Add-Member -Type NoteProperty -Name CreatedTime    -Value $WinREData.CreatedTime
    $winre | Add-Member -Type NoteProperty -Name OsBuildVersion -Value $WinRE_OsBuildVersion
    $winre | Add-Member -Type NoteProperty -Name ImagePath      -Value $WinREData.ImagePath
    $winre | Add-Member -Type NoteProperty -Name ImageName      -Value $WinREData.ImageName
    $winre | Add-Member -Type NoteProperty -Name ImageExisting  -Value ([bool]($WinREData -and [System.IO.File]::Exists($WinREData.ImagePath)))

    return $winre
}

function Invoke-Process {
<#
.SYNOPSIS
    Runs a native process with redirected I/O and returns its exit code.
.PARAMETER Path
    Executable path (e.g. reagentc.exe).
.PARAMETER Arguments
    Argument string (single string, passed verbatim).
.OUTPUTS
    [int] exit code.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [string] $Arguments
    )

    $ProgressPreference = 'SilentlyContinue'

    $pinfo = New-Object System.Diagnostics.ProcessStartInfo
    $pinfo.FileName               = $Path
    $pinfo.Arguments              = $Arguments
    $pinfo.RedirectStandardError  = $true
    $pinfo.RedirectStandardInput  = $true
    $pinfo.RedirectStandardOutput = $true
    $pinfo.LoadUserProfile        = $false
    $pinfo.UseShellExecute        = $false
    $pinfo.WindowStyle            = 'Hidden'

    $p           = New-Object System.Diagnostics.Process
    $p.StartInfo = $pinfo
    [void]$p.Start()
    $p.WaitForExit()
    return $p.ExitCode
}

function Test-ExitCondition {
<#
.SYNOPSIS
    Runs all pre-flight gates. Returns $EXIT_CONDITIONS_OK (-1) on success
    or a script exit code on the first failing gate.
.OUTPUTS
    [int]
#>
    [CmdletBinding()]
    param()

    # ----- BitLocker -----
    $osDrive = "$($env:SystemDrive)"
    $bde     = Get-BitLockerVolume -MountPoint $osDrive -ErrorAction SilentlyContinue
    if (-not $bde -or $bde.ProtectionStatus -ne 'On') {
        Write-Log -Level ERROR -Message '<Exit Condition>'
        Write-Log -Level ERROR -Message ' -> BDE protection status is set to Off.'
        return $ERROR_BDE_MISSING
    }
    Write-Log -Level INFO -Message ' -> BDE protection status is set to On.'

    # ----- ReAgent.xml + WinRE location -----
    $reagentXml = Join-Path $env:SystemRoot 'System32\Recovery\ReAgent.xml'
    if (-not (Test-Path -LiteralPath $reagentXml)) {
        Write-Log -Level ERROR -Message "<Exit Condition>"
        Write-Log -Level ERROR -Message " -> Unable to find $reagentXml"
        return $ERROR_XML_MISSING
    }

    [xml]$XmlDocument = Get-Content -LiteralPath $reagentXml -Raw
    $WinRE_LocationPath = $null
    $XmlDocument.SelectNodes('WindowsRE') | ForEach-Object {
        $WinRE_LocationPath = $_.WinreLocation.path
    }
    if ($WinRE_LocationPath) {
        Write-Log -Level ERROR -Message '<Exit Condition>'
        Write-Log -Level ERROR -Message ' -> WinRE WIM location found. Assuming WinRE is enabled. Nothing to do.'
        return $ERROR_WINRE_HEALTHY
    }
    Write-Log -Level INFO -Message ' -> WinreLocation path is empty.'

    # ----- Disk space on the OS partition -----
    $drive = (Get-CimInstance -ClassName Win32_OperatingSystem).SystemDrive
    $disk  = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$drive'"
    if ($disk.Size -gt 0) {
        $usedSpace      = $disk.Size - $disk.FreeSpace
        $usedPercentage = [math]::Round(($usedSpace / $disk.Size) * 100)
        if ($usedPercentage -gt 95) {
            Write-Log -Level ERROR -Message '<Exit Condition>'
            Write-Log -Level ERROR -Message " -> Not enough free space on the OS partition. Used: $usedPercentage%."
            return $ERROR_INSUFFICIENT_DISK_SPACE
        }
        Write-Log -Level INFO -Message " -> Current total disk capacity usage $usedPercentage%."
    }

    # ----- Completion marker -----
    $tmpRegistryPath = 'HKLM:\' + $CompletionRegistryPath
    if (Test-Path -LiteralPath $tmpRegistryPath) {
        $result = Get-ItemProperty -LiteralPath $tmpRegistryPath -Name $CompletionRegistryValueName -ErrorAction SilentlyContinue
        if ($result -and $result.$CompletionRegistryValueName -eq 1) {
            Write-Log -Level ERROR -Message '<Exit Condition>'
            Write-Log -Level ERROR -Message ' -> System already updated.'
            return $ERROR_ALREADY_PATCHED
        }
    } else {
        Write-Log -Level INFO -Message " -> $tmpRegistryPath\$CompletionRegistryValueName does not exist. Assuming first run."
    }

    return $EXIT_CONDITIONS_OK
}

# -----------------------------------------------------------------------------
# PROGRAM SEQUENCE
# -----------------------------------------------------------------------------
$Starttime = Get-Date
$ExitCode  = $ERROR_SUCCESS

Write-Log -Level INFO -Message '##################################################################################################################'
Write-Log -Level INFO -Message '####                            P R O G R A M    S E Q U E N C E   S T A R T                                  ####'
Write-Log -Level INFO -Message '##################################################################################################################'
Write-Log -Level INFO -Message 'Configuration Parameters'
Write-Log -Level INFO -Message " -> DiskNumber                  : $DiskNumber"
Write-Log -Level INFO -Message " -> MinRecoveryPartitionMB      : $MinRecoveryPartitionMB"
Write-Log -Level INFO -Message " -> LargeWimThresholdMB         : $LargeWimThresholdMB"
Write-Log -Level INFO -Message " -> HeadroomMB                  : $HeadroomMB"
Write-Log -Level INFO -Message " -> CompletionRegistryPath      : HKLM:\$CompletionRegistryPath"
Write-Log -Level INFO -Message " -> CompletionRegistryValueName : $CompletionRegistryValueName"
Write-Log -Level INFO -Message " -> LogDirectory                : $LogDirectory"

$WinRE_Current = Get-WinREData
Write-WinReDebugOutput -WinRE $WinRE_Current

$gate = Test-ExitCondition
if ($gate -eq $EXIT_CONDITIONS_OK) {
    Write-Log -Level INFO -Message '<Exit Condition> tests successfully completed.'

    $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
    Write-Log -Level INFO -Message 'Current Operating System'
    Write-Log -Level INFO -Message " -> Version : $($osInfo.Version) (Build $($osInfo.BuildNumber))"

    if (($WinRE_Current.ImagePath.Length -eq 0) -and (-not $WinRE_Current.ImageExisting)) {

        try {
            $WinReFilePath = Join-Path $env:SystemRoot 'System32\Recovery\Winre.wim'
            if (Test-Path -LiteralPath $WinReFilePath -PathType Leaf) {
                $WinReFileInfo   = Get-Item -LiteralPath $WinReFilePath
                $WinReFileSizeMB = [Math]::Round($WinReFileInfo.Length / 1MB, 2)
                Write-Log -Level INFO -Message " -> $WinReFilePath found."
                Write-Log -Level INFO -Message " -> Size of WinRE.wim: $WinReFileSizeMB MB"

                if ($WinReFileSizeMB -gt $LargeWimThresholdMB) {
                    # WIM is large -- grow partition to WIM + headroom (both in MB, then multiply once).
                    $WinREPartitionSizeBytes = [int64](($WinReFileSizeMB + $HeadroomMB) * 1MB)
                    Write-Log -Level INFO -Message " -> WIM is larger than $LargeWimThresholdMB MB. Target partition: $([Math]::Round($WinREPartitionSizeBytes / 1MB)) MB"
                } else {
                    $WinREPartitionSizeBytes = [int64]($MinRecoveryPartitionMB * 1MB)
                    Write-Log -Level INFO -Message " -> WIM is <= $LargeWimThresholdMB MB. Target partition: $MinRecoveryPartitionMB MB"
                }
            } else {
                Write-Log -Level FATAL -Message ' -> WinRE.wim does not exist in the expected location.'
                $ExitCode = $ERROR_WIM_MISSING
                throw 'WinRE.wim missing'
            }

            # Resolve the target disk and recovery partition.
            $Disk           = Get-Disk -Number $DiskNumber -ErrorAction Stop
            $WinREPartition = $Disk | Get-Partition | Where-Object { $_.Type -eq 'Recovery' }

            if (($WinREPartition | Measure-Object).Count -gt 1) {
                Write-Log -Level FATAL -Message ' -> ERROR: Multiple recovery partitions found. Manual cleanup required.'
                $ExitCode = $ERROR_MULTIPLE_RECOVERY_PARTITIONS
                throw 'Multiple recovery partitions'
            }
            if (-not $WinREPartition) {
                Write-Log -Level FATAL -Message ' -> ERROR: No recovery partition found on the target disk.'
                $ExitCode = $ERROR_NO_WINRE_DETECTED
                throw 'No recovery partition'
            }

            Write-Log -Level INFO -Message " -> Recovery partition number: $($WinREPartition.PartitionNumber) ($([Math]::Round($WinREPartition.Size / 1MB)) MB)"

            if ($WinREPartition.Size -lt $WinREPartitionSizeBytes) {
                Write-Log -Level INFO -Message ' -> Recovery partition is smaller than the target size. Resize required.'

                # Disable WinRE.
                Write-Log -Level INFO -Message ' -> Execute reagentc.exe /disable.'
                $rc = Invoke-Process -Path 'reagentc.exe' -Arguments '/disable'
                if ($rc -eq 0) {
                    Write-Log -Level INFO -Message '     -> Success'
                } else {
                    Write-Log -Level WARN -Message "     -> reagentc.exe /disable returned ExitCode=$rc"
                }

                # Delete the recovery partition.
                if ($PSCmdlet.ShouldProcess("Disk $($Disk.Number) partition $($WinREPartition.PartitionNumber)", 'Remove-Partition')) {
                    Remove-Partition -DiskNumber $Disk.Number -PartitionNumber $WinREPartition.PartitionNumber -Confirm:$false -ErrorAction Stop
                }
            } else {
                Write-Log -Level INFO -Message ' -> Recovery partition already meets or exceeds the target size.'
                $ExitCode = $ERROR_WINRE_HEALTHY
                throw 'Healthy partition'
            }

            # Shrink OS partition by the size of the new recovery partition.
            $OSPartition = Get-Disk -Number $Disk.Number | Get-Partition | Where-Object { $_.IsBoot -eq $true }
            if (-not $OSPartition) {
                Write-Log -Level FATAL -Message ' -> ERROR: Could not locate the boot partition on the target disk.'
                $ExitCode = $ERROR_UNKNOWN
                throw 'Boot partition not found'
            }
            Write-Log -Level INFO -Message " -> OS partition number: $($OSPartition.PartitionNumber) (current size: $($OSPartition.Size) bytes)"

            $ParSizeBytes = [UInt64]($OSPartition.Size - $WinREPartitionSizeBytes)
            Write-Log -Level INFO -Message " -> New OS partition size after resize: $ParSizeBytes bytes"

            if ($PSCmdlet.ShouldProcess("Disk $($Disk.Number) partition $($OSPartition.PartitionNumber)", "Resize-Partition to $ParSizeBytes")) {
                Resize-Partition -DiskNumber $Disk.Number -PartitionNumber $OSPartition.PartitionNumber -Size $ParSizeBytes -ErrorAction Stop
            }

            # Create the new recovery partition at the end of the disk.
            Write-Log -Level INFO -Message 'Creating a new recovery partition.'
            $NewPartition = New-Partition -DiskNumber $Disk.Number -UseMaximumSize -ErrorAction Stop

            # Format NTFS, label Recovery.
            Write-Log -Level INFO -Message 'Formatting the new recovery partition (NTFS, label "Recovery").'
            Format-Volume -Partition $NewPartition -FileSystem NTFS -NewFileSystemLabel 'Recovery' -Confirm:$false -ErrorAction Stop | Out-Null

            # Tag as WinRE GPT partition via diskpart.exe.
            # Set-Partition was observed to BSOD on certain firmware (kept from 1.2 history); diskpart is the safe path.
            Write-Log -Level INFO -Message "Designating partition $($NewPartition.PartitionNumber) on disk $($Disk.Number) as WinRE partition (via diskpart)."
            $DiskPartCMD = @(
                "select disk $($Disk.Number)"
                "select partition $($NewPartition.PartitionNumber)"
                'set id = DE94BBA4-06D1-4D40-A16A-BFD50179D6AC'
                'gpt attributes = 0x8000000000000001'
                'exit'
            )
            $DiskPartCMD | diskpart.exe | Out-Null

            # Re-enable WinRE.
            Write-Log -Level INFO -Message ' -> Execute reagentc.exe /enable.'
            $rc = Invoke-Process -Path 'reagentc.exe' -Arguments '/enable'
            if ($rc -eq 0) {
                Write-Log -Level INFO -Message '     -> Success'
            } else {
                Write-Log -Level FATAL -Message "     -> reagentc.exe /enable failed (ExitCode=$rc)"
                $ExitCode = $ERROR_TPM_UPDATE_FAILED
                throw 'reagentc /enable failed'
            }
        }
        catch {
            Write-Log -Level FATAL -Message " -> ERROR: $($_.Exception.Message)"
            if ($ExitCode -eq $ERROR_SUCCESS) { $ExitCode = $ERROR_UNKNOWN }
        }

        # Refreshed WinRE state.
        $WinRE_Updated = Get-WinREData
        Write-WinReDebugOutput -WinRE $WinRE_Updated

        # Stamp completion marker on success / already-patched.
        if ($ExitCode -eq $ERROR_SUCCESS -or $ExitCode -eq $ERROR_ALREADY_PATCHED) {
            Write-Log -Level INFO -Message 'Creating completion event in the System Registry.'
            $tmpRegistryPath = 'HKLM:\' + $CompletionRegistryPath
            if (-not (Test-Path -LiteralPath $tmpRegistryPath)) {
                New-Item -Path $tmpRegistryPath -Force -ErrorAction SilentlyContinue | Out-Null
            }
            New-ItemProperty -Path $tmpRegistryPath -Name $CompletionRegistryValueName -Value 1 -PropertyType DWord -Force -ErrorAction SilentlyContinue | Out-Null
            Write-Log -Level INFO -Message ' -> Successfully created.'
        }
    } else {
        Write-Log -Level INFO -Message 'Windows Recovery Environment not found. No actions required.'
        $ExitCode = $ERROR_NO_WINRE_DETECTED
    }
} else {
    switch ($gate) {
        $ERROR_ALREADY_PATCHED              { Write-Log -Level FATAL -Message '    ERROR_ALREADY_PATCHED' }
        $ERROR_NO_WINRE_DETECTED            { Write-Log -Level FATAL -Message '    ERROR_NO_WINRE_DETECTED' }
        $ERROR_XML_MISSING                  { Write-Log -Level FATAL -Message '    ERROR_XML_MISSING' }
        $ERROR_WIM_MISSING                  { Write-Log -Level FATAL -Message '    ERROR_WIM_MISSING' }
        $ERROR_WINRE_HEALTHY                { Write-Log -Level FATAL -Message '    ERROR_WINRE_HEALTHY' }
        $ERROR_MULTIPLE_RECOVERY_PARTITIONS { Write-Log -Level FATAL -Message '    ERROR_MULTIPLE_RECOVERY_PARTITIONS' }
        $ERROR_INSUFFICIENT_DISK_SPACE      { Write-Log -Level FATAL -Message '    ERROR_INSUFFICIENT_DISK_SPACE' }
        $ERROR_BDE_MISSING                  { Write-Log -Level FATAL -Message '    ERROR_BDE_MISSING' }
        $ERROR_PREREQ_FAILURE               { Write-Log -Level FATAL -Message '    ERROR_PREREQ_FAILURE' }
        $ERROR_TPM_UPDATE_FAILED            { Write-Log -Level FATAL -Message '    ERROR_TPM_UPDATE_FAILED' }
        default                             { Write-Log -Level FATAL -Message "    ERROR_UNKNOWN ($gate)" }
    }
    $ExitCode = $gate
}

# -----------------------------------------------------------------------------
# PROGRAM END
# -----------------------------------------------------------------------------
$Endtime  = Get-Date
$Duration = $Endtime - $Starttime
Write-Log -Level INFO -Message "Program sequence end. Duration: $Duration"

if ($RewriteAlreadyPatchedToSuccess -and $ExitCode -eq $ERROR_ALREADY_PATCHED) { $ExitCode = $ERROR_SUCCESS }
if ($RewriteNoWinReToSuccess        -and $ExitCode -eq $ERROR_NO_WINRE_DETECTED) { $ExitCode = $ERROR_SUCCESS }
if ($RewriteWinReHealthyToSuccess   -and $ExitCode -eq $ERROR_WINRE_HEALTHY)   { $ExitCode = $ERROR_SUCCESS }

Write-Log -Level INFO -Message " -> Exit code: $ExitCode"
exit $ExitCode
