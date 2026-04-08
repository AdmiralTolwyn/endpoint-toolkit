<#
    .SYNOPSIS
    Repairs dirty FSLogix Profile and O365 dynamically expanding disk(s).

    .DESCRIPTION
    FSLogix profile and O365 virtual hard disks can become marked as "dirty" when
    they are not cleanly dismounted. This can happen due to session host crashes,
    forced shutdowns, storage connectivity issues, or other abnormal terminations.

    A dirty volume flag causes Windows to run a consistency check the next time the
    disk is mounted, which can significantly delay user logon times. In some cases,
    file system corruption accompanies the dirty flag, leading to profile data loss.

    This script is designed to work at Enterprise scale to repair thousands of disks
    in the shortest time possible. It mounts each VHD/VHDX, checks for the dirty bit,
    runs chkdsk /f to repair file system errors and clear the dirty flag, then cleanly
    dismounts the disk.

    This script can be run from any machine in your environment - it does not need to
    be run from a file server hosting the disks. It does not need the Hyper-V role
    installed. Mount-DiskImage is used instead, which is available on all editions of
    Windows 10/11 and Server 2016+.

    PowerShell version 5.x, 7.x and above are supported. It must be run as
    administrator due to the requirement for mounting disks and running chkdsk.

    This tool is multi-threaded and will take advantage of multiple CPU cores on the
    machine from which you run the script. It is not advised to run more than 2x the
    threads of your available cores. You can use the ThrottleLimit parameter to
    throttle the load on your storage as well.

    On PowerShell 7+, ForEach-Object -Parallel is used for threading. On PowerShell
    5.x, a bundled Invoke-Parallel function using runspace pools provides equivalent
    multi-threading capability.

    The script will output a CSV in the following format:

    "Name","DiskState","FullName"
    "Profile_user1.vhdx","Repaired","\\Server\Share\Profile_user1.vhdx"
    "Profile_user2.vhdx","NotDirty","\\Server\Share\Profile_user2.vhdx"
    "Profile_user3.vhdx","Checked","\\Server\Share\Profile_user3.vhdx"

    Possible Information values for DiskState are as follows:
    Repaired            Disk had dirty bit set and was successfully repaired via chkdsk
    NotDirty            Disk was not dirty - no action needed
    Checked             Disk was not dirty but ForceRepair ran chkdsk anyway - no errors
    Skipped             Disk was skipped (dirty status could not be determined)

    Possible Error values for DiskState are as follows:
    FileIsNotDiskFormat Disk file extension was not vhd or vhdx
    DiskLocked          Disk could not be mounted due to being in use by an active session
    MountFailed         Failed to mount the disk (permissions, corrupt container, etc.)
    NoPartitionInfo     Could not get partition information from the mounted disk
    RepairFailed        Both chkdsk /f and Repair-Volume failed to complete successfully
    DismountFailed      Failed to cleanly dismount the disk after repair

    .PARAMETER Path
    The path to the folder/share containing the disks. You can also directly
    specify a single disk file. UNC paths are supported for remote file shares.

    .PARAMETER Recurse
    Gets the disks in the specified locations and in all child items of the locations.
    Use this when profile disks are organized in per-user subdirectories (the typical
    FSLogix folder structure).

    .PARAMETER LogFilePath
    All disk actions will be saved in a CSV file for admin reference. The default
    location for this CSV file is the user's temp directory. The default filename
    follows the format: FslRepairDisk yyyy-MM-dd HH-mm-ss.csv

    .PARAMETER PassThru
    Returns an object representing the item with which you are working. By default,
    this cmdlet does not generate any pipeline output. Use this when you want to
    capture results for further processing in a pipeline.

    .PARAMETER ThrottleLimit
    Specifies the number of disks that will be processed concurrently. Further disks
    in the queue will wait until a previous disk has finished. The default value is 8.
    The script will automatically cap this at 2x the number of logical processors if
    a higher value is specified.

    .PARAMETER ForceRepair
    When set, runs chkdsk /f on every disk regardless of dirty bit status. Useful
    for proactive file system integrity checks across an entire profile share.
    Without this switch, only disks with the dirty bit set (or those that fail the
    Repair-Volume -Scan check) will be repaired.

    .INPUTS
    You can pipe the path into the command which is recognised by type, you can also
    pipe any parameter by name. It will also take the path positionally.

    .OUTPUTS
    This script outputs a CSV file with the result of the disk processing.
    It will optionally produce a custom object with the same information when
    -PassThru is specified.

    .EXAMPLE
    C:\PS> .\Invoke-FslRepairDisk.ps1 -Path \\server\share -Recurse
    This checks and repairs all dirty disks in the specified share recursively.

    .EXAMPLE
    C:\PS> .\Invoke-FslRepairDisk.ps1 -Path c:\Profiles\Profile_user1.vhdx
    This repairs a single disk on the local file system.

    .EXAMPLE
    C:\PS> .\Invoke-FslRepairDisk.ps1 -Path \\server\share -Recurse -ForceRepair
    This runs chkdsk /f on all disks regardless of dirty bit status. Useful for
    proactive maintenance windows.

    .EXAMPLE
    C:\PS> .\Invoke-FslRepairDisk.ps1 -Path \\server\share -Recurse -PassThru -ThrottleLimit 16
    This repairs disks with 16 concurrent threads and outputs results to the pipeline.

    .EXAMPLE
    C:\PS> .\Invoke-FslRepairDisk.ps1 -Path \\server\share -Recurse -LogFilePath C:\Logs\repair.csv
    This repairs disks and saves the log to a custom location.

    .EXAMPLE
    C:\PS> .\Invoke-FslRepairDisk.ps1 -Path \\server\share -Recurse -PassThru |
        Where-Object DiskState -eq 'RepairFailed' | Select-Object FullName
    This repairs disks and filters the output to show only disks that failed repair,
    which may require manual intervention.

    .NOTES
    Based on the architecture of Invoke-FslShrinkDisk by Jim Moyle.
    https://github.com/FSLogix/Invoke-FslShrinkDisk

    Dirty bit detection uses fsutil dirty query as the primary method, with
    Repair-Volume -Scan as a fallback. Repair uses chkdsk /f /x as primary,
    with Repair-Volume -OfflineScanAndFix as a fallback.

    FSLogix diff disks (Merge.vhdx, RW.vhdx) are automatically excluded from
    processing as they are managed by the FSLogix agent.

    This script does not delete any disks. It is a read/repair-only operation
    on the file system inside each virtual disk.

    .LINK
    https://github.com/FSLogix/Invoke-FslShrinkDisk/
#>

[CmdletBinding()]

Param (

    [Parameter(
        Position = 1,
        ValuefromPipelineByPropertyName = $true,
        ValuefromPipeline = $true,
        Mandatory = $true
    )]
    [System.String]$Path,

    [Parameter(
        ValuefromPipelineByPropertyName = $true
    )]
    [Switch]$Recurse,

    [Parameter(
        ValuefromPipelineByPropertyName = $true
    )]
    [System.String]$LogFilePath = "$env:TEMP\FslRepairDisk $(Get-Date -Format yyyy-MM-dd` HH-mm-ss).csv",

    [Parameter(
        ValuefromPipelineByPropertyName = $true
    )]
    [switch]$PassThru,

    [Parameter(
        ValuefromPipelineByPropertyName = $true
    )]
    [int]$ThrottleLimit = 8,

    [Parameter(
        ValuefromPipelineByPropertyName = $true
    )]
    [switch]$ForceRepair
)

# ============================================================================
# BEGIN Block - One-time initialization: define helper functions, validate
# prerequisites (services, core count), and prepare the runspace environment.
# ============================================================================
BEGIN {
    Set-StrictMode -Version Latest
    #Requires -RunAsAdministrator

    #region Helper Functions
    # These functions are defined in the BEGIN block so they are available to
    # the PROCESS block and to Invoke-Parallel (PS 5.x) via -ImportFunctions.
    # For PowerShell 7+ ForEach-Object -Parallel, they are redefined inline
    # inside the scriptblock (see PROCESS block).

    Function Test-FslDependencies {
        <#
        .SYNOPSIS
        Validates that required Windows services are running.

        .DESCRIPTION
        Checks each named service and ensures it is in a Running state. If a service
        is Disabled, it will be set to Manual start and then started. This is required
        because chkdsk and disk operations depend on the Virtual Disk Service (vds)
        and Optimize Drives service (defragsvc).

        .PARAMETER Name
        An array of Windows service names to validate. Each service will be checked,
        enabled if disabled, and started if not running.
        #>
        [CmdletBinding()]
        Param (
            [Parameter(
                Mandatory = $true,
                Position = 0,
                ValueFromPipelineByPropertyName = $true,
                ValueFromPipeline = $true
            )]
            [System.String[]]$Name
        )
        BEGIN {
            Set-StrictMode -Version Latest
        }
        PROCESS {
            Foreach ($svc in $Name) {
                $svcObject = Get-Service -Name $svc

                # Service already running - nothing to do
                If ($svcObject.Status -eq "Running") { Return }

                # If the service is disabled, set it to manual so we can start it
                If ($svcObject.StartType -eq "Disabled") {
                    Write-Warning ("[{0}] Setting Service to Manual" -f $svcObject.DisplayName)
                    Set-Service -Name $svc -StartupType Manual | Out-Null
                }

                # Attempt to start the service
                Start-Service -Name $svc | Out-Null
                if ((Get-Service -Name $svc).Status -ne 'Running') {
                    Write-Error "Can not start $($svcObject.DisplayName)"
                }
            }
        }
        END { }
    }

    function Mount-FslDisk {
        <#
        .SYNOPSIS
        Mounts a VHD/VHDX disk to a temporary directory without requiring a drive letter.

        .DESCRIPTION
        Uses Mount-DiskImage (no Hyper-V dependency) to mount a virtual hard disk without
        assigning a drive letter. Instead, a GUID-named temp folder is created and the
        partition is mounted there via Add-PartitionAccessPath. This avoids the well-known
        problem of running out of drive letters when processing many disks in parallel.

        The function includes timeout-based retry loops for both disk number resolution
        and partition type detection, since the Windows Disk subsystem can be slow to
        surface information under heavy concurrent load.

        If any step fails, cleanup is attempted (dismount + remove temp dir) before
        returning an error.

        .PARAMETER Path
        Full path to the VHD or VHDX file to mount. Accepts the 'FullName' alias for
        pipeline compatibility with Get-ChildItem output.

        .PARAMETER TimeOut
        Number of seconds to wait for disk number and partition information to become
        available after mounting. Default is 3 seconds.

        .PARAMETER PassThru
        When specified, outputs a PSCustomObject with Path (mount directory), DiskNumber,
        ImagePath, and PartitionNumber properties. This object can be piped directly to
        Dismount-FslDisk.
        #>
        [CmdletBinding()]
        Param (
            [Parameter(
                Position = 1,
                ValuefromPipelineByPropertyName = $true,
                ValuefromPipeline = $true,
                Mandatory = $true
            )]
            [alias('FullName')]
            [System.String]$Path,

            [Parameter(
                ValuefromPipelineByPropertyName = $true,
                ValuefromPipeline = $true
            )]
            [Int]$TimeOut = 3,

            [Parameter(
                ValuefromPipelineByPropertyName = $true
            )]
            [Switch]$PassThru
        )

        BEGIN {
            Set-StrictMode -Version Latest
        }
        PROCESS {
            # Mount the disk without a drive letter - Mount-DiskImage avoids Hyper-V dependency
            try {
                $mountedDisk = Mount-DiskImage -ImagePath $Path -NoDriveLetter -PassThru -ErrorAction Stop
            }
            catch {
                $e = $error[0]
                Write-Error "Failed to mount disk - `"$e`""
                return
            }

            # Wait for the disk subsystem to assign a disk number (can be slow under load)
            $diskNumber = $false
            $timespan = (Get-Date).AddSeconds($TimeOut)
            while ($diskNumber -eq $false -and $timespan -gt (Get-Date)) {
                Start-Sleep 0.1
                try {
                    $mountedDisk = Get-DiskImage -ImagePath $Path
                    if ($mountedDisk.Number) {
                        $diskNumber = $true
                    }
                }
                catch {
                    $diskNumber = $false
                }
            }

            # If we couldn't get a disk number, clean up and bail
            if ($diskNumber -eq $false) {
                try { $mountedDisk | Dismount-DiskImage -ErrorAction SilentlyContinue }
                catch {
                    Write-Error 'Could not dismount Disk Due to no Disknumber'
                }
                Write-Error 'Cannot get mount information'
                return
            }

            # Wait for partition information - look for the 'Basic' data partition
            $partitionType = $false
            $timespan = (Get-Date).AddSeconds($TimeOut)
            while ($partitionType -eq $false -and $timespan -gt (Get-Date)) {
                try {
                    $allPartition = Get-Partition -DiskNumber $mountedDisk.Number -ErrorAction Stop
                    if ($allPartition.Type -contains 'Basic') {
                        $partitionType = $true
                        $partition = $allPartition | Where-Object -Property 'Type' -EQ -Value 'Basic'
                    }
                }
                catch {
                    # Fallback: if partition type detection fails, take the last partition
                    if (($allPartition | Measure-Object).Count -gt 0) {
                        $partition = $allPartition | Select-Object -Last 1
                        $partitionType = $true
                    }
                    else {
                        $partitionType = $false
                    }
                }
                Start-Sleep 0.1
            }

            # No partition found - clean up and return error
            if ($partitionType -eq $false) {
                try { $mountedDisk | Dismount-DiskImage -ErrorAction SilentlyContinue }
                catch {
                    Write-Error 'Could not dismount disk with no partition'
                }
                Write-Error 'Cannot get partition information'
                return
            }

            # Create a GUID-named temp directory for the mount point (avoids drive letter exhaustion)
            $tempGUID = [guid]::NewGuid().ToString()
            $mountPath = Join-Path $Env:Temp ('FSLogixMnt-' + $tempGUID)

            try {
                New-Item -Path $mountPath -ItemType Directory -ErrorAction Stop | Out-Null
            }
            catch {
                $e = $error[0]
                try { $mountedDisk | Dismount-DiskImage -ErrorAction SilentlyContinue }
                catch {
                    Write-Error "Could not dismount disk when no folder could be created - `"$e`""
                }
                Write-Error "Failed to create mounting directory - `"$e`""
                return
            }

            # Create a junction point from the temp directory to the partition
            try {
                $addPartitionAccessPathParams = @{
                    DiskNumber      = $mountedDisk.Number
                    PartitionNumber = $partition.PartitionNumber
                    AccessPath      = $mountPath
                    ErrorAction     = 'Stop'
                }
                Add-PartitionAccessPath @addPartitionAccessPathParams
            }
            catch {
                $e = $error[0]
                Remove-Item -Path $mountPath -Force -Recurse -ErrorAction SilentlyContinue
                try { $mountedDisk | Dismount-DiskImage -ErrorAction SilentlyContinue }
                catch {
                    Write-Error "Could not dismount disk when no junction point could be created - `"$e`""
                }
                Write-Error "Failed to create junction point to - `"$e`""
                return
            }

            # Output mount info for piping to Dismount-FslDisk
            if ($PassThru) {
                $output = [PSCustomObject]@{
                    Path            = $mountPath
                    DiskNumber      = $mountedDisk.Number
                    ImagePath       = $mountedDisk.ImagePath
                    PartitionNumber = $partition.PartitionNumber
                }
                Write-Output $output
            }
            Write-Verbose "Mounted $Path to $mountPath"
        }
        END { }
    }

    function Dismount-FslDisk {
        <#
        .SYNOPSIS
        Cleanly dismounts a VHD/VHDX disk and removes its temporary mount directory.

        .DESCRIPTION
        Reverses the operations performed by Mount-FslDisk. First removes the temp
        mount directory (junction point), then calls Dismount-DiskImage to detach the
        virtual disk. Both operations include retry loops because the Windows Disk
        Manager service can hold brief locks, especially under concurrent load.

        The dismount is verified by checking Get-DiskImage.Attached to ensure the disk
        is truly detached - Dismount-DiskImage alone is not always reliable.

        .PARAMETER Path
        The temp mount directory path created by Mount-FslDisk.

        .PARAMETER ImagePath
        The full path to the VHD/VHDX file to dismount.

        .PARAMETER PassThru
        When specified, outputs a PSCustomObject indicating whether the mount and
        directory were successfully removed.

        .PARAMETER Timeout
        Maximum number of seconds to retry dismounting before giving up. Default is
        120 seconds to handle stubborn disk manager locks.
        #>
        [CmdletBinding()]
        Param (
            [Parameter(
                Position = 1,
                ValuefromPipelineByPropertyName = $true,
                ValuefromPipeline = $true,
                Mandatory = $true
            )]
            [String]$Path,

            [Parameter(
                ValuefromPipelineByPropertyName = $true,
                Mandatory = $true
            )]
            [String]$ImagePath,

            [Parameter(
                ValuefromPipelineByPropertyName = $true
            )]
            [Switch]$PassThru,

            [Parameter(
                ValuefromPipelineByPropertyName = $true
            )]
            [Int]$Timeout = 120
        )

        BEGIN {
            Set-StrictMode -Version Latest
        }
        PROCESS {
            $mountRemoved = $false
            $directoryRemoved = $false

            # Step 1: Remove the temp mount directory (retry for up to 20 seconds)
            $timeStampDirectory = (Get-Date).AddSeconds(20)
            while ((Get-Date) -lt $timeStampDirectory -and $directoryRemoved -ne $true) {
                try {
                    Remove-Item -Path $Path -Force -Recurse -ErrorAction Stop | Out-Null
                    $directoryRemoved = $true
                }
                catch {
                    $directoryRemoved = $false
                }
            }
            if (Test-Path $Path) {
                Write-Warning "Failed to delete temp mount directory $Path"
            }

            # Step 2: Dismount the disk image and verify detachment
            # The disk manager service can be unreliable, so we triple-check via Get-DiskImage
            $timeStampDismount = (Get-Date).AddSeconds($Timeout)
            while ((Get-Date) -lt $timeStampDismount -and $mountRemoved -ne $true) {
                try {
                    Dismount-DiskImage -ImagePath $ImagePath -ErrorAction Stop | Out-Null
                    # Verify the disk is actually detached - Dismount-DiskImage can claim
                    # success while the disk is still attached
                    try {
                        $image = Get-DiskImage -ImagePath $ImagePath -ErrorAction Stop
                        switch ($image.Attached) {
                            $null  { $mountRemoved = $false ; Start-Sleep 0.1; break }
                            $true  { $mountRemoved = $false ; break }
                            $false { $mountRemoved = $true ; break }
                            Default { $mountRemoved = $false }
                        }
                    }
                    catch {
                        $mountRemoved = $false
                    }
                }
                catch {
                    $mountRemoved = $false
                }
            }
            if ($mountRemoved -ne $true) {
                Write-Error "Failed to dismount disk $ImagePath"
            }

            If ($PassThru) {
                $output = [PSCustomObject]@{
                    MountRemoved     = $mountRemoved
                    DirectoryRemoved = $directoryRemoved
                }
                Write-Output $output
            }
            if ($directoryRemoved -and $mountRemoved) {
                Write-Verbose "Dismounted $ImagePath"
            }
        }
        END { }
    }

    function Repair-OneDisk {
        <#
        .SYNOPSIS
        Processes a single VHD/VHDX disk: checks dirty bit, repairs if needed, logs result.

        .DESCRIPTION
        This is the core per-disk processing function. For each disk it:
        1. Dismisses any stale mounts from previous failed runs
        2. Validates the file is a VHD/VHDX
        3. Mounts the disk via Mount-FslDisk (no drive letter, temp directory)
        4. Retrieves partition info with a 120-second retry loop
        5. Checks the dirty bit via fsutil, with Repair-Volume -Scan as fallback
        6. Runs chkdsk /f /x to repair and clear the dirty flag
        7. Falls back to Repair-Volume -OfflineScanAndFix if chkdsk fails
        8. Cleanly dismounts the disk
        9. Logs the result to CSV and optionally to the pipeline

        Each step has error handling that ensures the disk is always dismounted,
        even on failure, to prevent leaked mounts.

        .PARAMETER Disk
        A System.IO.FileInfo object representing the VHD/VHDX file to process.
        Typically provided by Get-ChildItem in the outer loop.

        .PARAMETER MountTimeout
        Maximum seconds to wait for the disk to mount. Default: 30.

        .PARAMETER LogFilePath
        Path to the CSV log file for recording results.

        .PARAMETER Passthru
        When set, emits the result object to the pipeline.

        .PARAMETER ForceRepair
        When set, runs chkdsk regardless of whether the dirty bit is set.
        #>
        [CmdletBinding()]
        Param (
            [Parameter(
                ValuefromPipelineByPropertyName = $true,
                ValuefromPipeline = $true,
                Mandatory = $true
            )]
            [System.IO.FileInfo]$Disk,

            [Parameter(
                ValuefromPipelineByPropertyName = $true
            )]
            [int]$MountTimeout = 30,

            [Parameter(
                ValuefromPipelineByPropertyName = $true
            )]
            [string]$LogFilePath = "$env:TEMP\FslRepairDisk $(Get-Date -Format yyyy-MM-dd` HH-mm-ss).csv",

            [Parameter(
                ValuefromPipelineByPropertyName = $true
            )]
            [switch]$Passthru,

            [Parameter(
                ValuefromPipelineByPropertyName = $true
            )]
            [switch]$ForceRepair
        )

        BEGIN {
            Set-StrictMode -Version Latest
        }
        PROCESS {
            # Safety cleanup: dismiss any stale mount from a previous failed run.
            # This is a no-op if the disk isn't currently mounted.
            Dismount-DiskImage -ImagePath $Disk.FullName -ErrorAction SilentlyContinue

            # Record start time for elapsed time calculation in the CSV log
            $startTime = Get-Date

            $originalSize = $Disk.Length

            # Pre-populate Write-VhdOutput parameters via PSDefaultParameterValues.
            # This avoids repeating the same parameters at every call site below.
            # Only DiskState and EndTime need to be specified per-call.
            $PSDefaultParameterValues = @{
                "Write-VhdOutput:Path"         = $LogFilePath
                "Write-VhdOutput:StartTime"    = $startTime
                "Write-VhdOutput:Name"         = $Disk.Name
                "Write-VhdOutput:DiskState"    = $null
                "Write-VhdOutput:OriginalSize" = $originalSize
                "Write-VhdOutput:FinalSize"    = $originalSize
                "Write-VhdOutput:FullName"     = $Disk.FullName
                "Write-VhdOutput:Passthru"     = $Passthru
            }

            # Validate file extension before attempting to mount
            if ($Disk.Extension -ne '.vhd' -and $Disk.Extension -ne '.vhdx') {
                Write-VhdOutput -DiskState 'FileIsNotDiskFormat' -EndTime (Get-Date)
                return
            }

            # Mount the disk to a temp directory (no drive letter required)
            try {
                $mount = Mount-FslDisk -Path $Disk.FullName -TimeOut $MountTimeout -PassThru -ErrorAction Stop
            }
            catch {
                # Distinguish between "in use by active session" vs other mount failures
                $err = $error[0]
                if ($err -match 'disk is already in use' -or $err -match 'being used by another process' -or $err -match 'locked') {
                    Write-VhdOutput -DiskState 'DiskLocked' -EndTime (Get-Date)
                }
                else {
                    Write-VhdOutput -DiskState "MountFailed" -EndTime (Get-Date)
                }
                return
            }

            # Retrieve partition info with 120-second retry loop.
            # The Windows Disk subsystem can be slow to surface partition details
            # when many disks are being mounted concurrently across threads.
            $timespan = (Get-Date).AddSeconds(120)
            $partInfo = $null
            while (($partInfo | Measure-Object).Count -lt 1 -and $timespan -gt (Get-Date)) {
                try {
                    # Look for the data partition (Type = 'Basic')
                    $partInfo = Get-Partition -DiskNumber $mount.DiskNumber -ErrorAction Stop |
                        Where-Object -Property 'Type' -EQ -Value 'Basic' -ErrorAction Stop
                }
                catch {
                    # Fallback: grab whatever partition exists
                    $partInfo = Get-Partition -DiskNumber $mount.DiskNumber -ErrorAction SilentlyContinue |
                        Select-Object -Last 1
                }
                Start-Sleep 0.1
            }

            if (($partInfo | Measure-Object).Count -eq 0) {
                $mount | Dismount-FslDisk
                Write-VhdOutput -DiskState 'NoPartitionInfo' -EndTime (Get-Date)
                return
            }

            # Use the temp mount path as the volume path for fsutil and chkdsk
            $volumePath = $mount.Path

            # ---- Dirty Bit Detection ----
            # Primary: fsutil dirty query (works with mount paths on NTFS/ReFS)
            # Fallback: Repair-Volume -Scan (when fsutil output is ambiguous)
            $isDirty = $false
            try {
                $dirtyResult = & fsutil dirty query "$volumePath" 2>&1
                $dirtyOutput = $dirtyResult -join ' '
                if ($dirtyOutput -match 'is Dirty' -or $dirtyOutput -match 'dirty') {
                    $isDirty = $true
                }
                elseif ($dirtyOutput -match 'NOT Dirty' -or $dirtyOutput -match 'not dirty') {
                    $isDirty = $false
                }
                else {
                    # Fallback: Try Repair-Volume -Scan to check integrity
                    try {
                        $vol = Get-Volume -Partition $partInfo -ErrorAction Stop
                        $scanResult = Repair-Volume -InputObject $vol -Scan -ErrorAction Stop
                        if ($scanResult -eq 'ScanNeeded' -or $scanResult -eq 'SpotFixesNeeded') {
                            $isDirty = $true
                        }
                    }
                    catch {
                        # If we can't determine dirty status and ForceRepair is set, proceed
                        if ($ForceRepair) {
                            $isDirty = $true
                        }
                        else {
                            Write-Verbose "Could not determine dirty status for $($Disk.Name), skipping"
                            $mount | Dismount-FslDisk
                            Write-VhdOutput -DiskState 'Skipped' -EndTime (Get-Date)
                            return
                        }
                    }
                }
            }
            catch {
                if ($ForceRepair) {
                    $isDirty = $true
                }
                else {
                    Write-Verbose "Could not query dirty bit for $($Disk.Name), skipping"
                    $mount | Dismount-FslDisk
                    Write-VhdOutput -DiskState 'Skipped' -EndTime (Get-Date)
                    return
                }
            }

            # If not dirty and not forcing repair, skip
            if (-not $isDirty -and -not $ForceRepair) {
                $mount | Dismount-FslDisk
                Write-VhdOutput -DiskState 'NotDirty' -EndTime (Get-Date)
                return
            }

            # ---- Repair Phase ----
            # Primary: chkdsk /f /x (fixes errors, /x forces dismount of open handles)
            # Fallback: Repair-Volume -OfflineScanAndFix (if chkdsk fails)
            $repairSuccess = $false
            $chkdskOutput = $null
            try {
                # chkdsk requires a trailing backslash on mount-path volumes
                $chkdskPath = $volumePath.TrimEnd('\') + '\'
                $chkdskOutput = & chkdsk $chkdskPath /f /x 2>&1
                $chkdskText = $chkdskOutput -join "`n"

                # chkdsk returns 0 on success
                if ($LASTEXITCODE -eq 0) {
                    $repairSuccess = $true
                }
                else {
                    # Some non-zero exit codes still indicate work was done
                    # Check output text for key indicators
                    if ($chkdskText -match 'Windows has made corrections' -or
                        $chkdskText -match 'no problems' -or
                        $chkdskText -match 'Windows has checked the file system' -or
                        $chkdskText -match 'cleaning up') {
                        $repairSuccess = $true
                    }
                }
            }
            catch {
                $repairSuccess = $false
            }

            # Fallback: if chkdsk didn't work, try the Windows Storage cmdlet.
            # Repair-Volume -OfflineScanAndFix uses the same underlying engine but
            # can succeed in cases where chkdsk's text-based interface fails.
            if (-not $repairSuccess) {
                try {
                    $vol = Get-Volume -Partition $partInfo -ErrorAction Stop
                    $repairResult = Repair-Volume -InputObject $vol -OfflineScanAndFix -ErrorAction Stop
                    if ($repairResult -eq 'NoErrorsFound' -or $repairResult -eq 'Fixed') {
                        $repairSuccess = $true
                    }
                }
                catch {
                    $repairSuccess = $false
                }
            }

            # Always dismount cleanly, regardless of repair outcome
            $mount | Dismount-FslDisk

            # Log the final result: 'Repaired' if dirty bit was set, 'Checked' if ForceRepair
            # ran on a clean disk, or 'RepairFailed' if both repair methods failed
            if ($repairSuccess) {
                $state = if ($isDirty) { 'Repaired' } else { 'Checked' }
                Write-VhdOutput -DiskState $state -FinalSize (Get-ChildItem $Disk.FullName | Select-Object -ExpandProperty Length) -EndTime (Get-Date)
            }
            else {
                Write-VhdOutput -DiskState 'RepairFailed' -EndTime (Get-Date)
            }
        }
        END { }
    }

    function Write-VhdOutput {
        <#
        .SYNOPSIS
        Writes disk processing results to a CSV log file and optionally to the pipeline.

        .DESCRIPTION
        Creates a standardised PSCustomObject with disk processing metrics (name, state,
        timing, size) and appends it to the CSV log file. The CSV write includes a retry
        loop (up to 10 attempts with 1-second delay) to handle file contention when
        multiple threads write to the same log concurrently.

        This function is called with $PSDefaultParameterValues set in Repair-OneDisk,
        so most parameters are pre-populated. Only DiskState and EndTime need to be
        specified at each call site.

        .PARAMETER Path
        Path to the CSV log file.

        .PARAMETER Name
        The disk file name (e.g., Profile_user1.vhdx).

        .PARAMETER DiskState
        The result state string (Repaired, NotDirty, Checked, Skipped, RepairFailed, etc.).

        .PARAMETER OriginalSize
        The disk file size in bytes before processing.

        .PARAMETER FinalSize
        The disk file size in bytes after processing.

        .PARAMETER FullName
        The full path to the disk file.

        .PARAMETER StartTime
        Timestamp when processing began for this disk.

        .PARAMETER EndTime
        Timestamp when processing completed for this disk.

        .PARAMETER Passthru
        When set, emits the result object to the output pipeline.
        #>
        [CmdletBinding()]
        Param (
            [Parameter(Mandatory = $true)]
            [System.String]$Path,

            [Parameter(Mandatory = $true)]
            [System.String]$Name,

            [Parameter(Mandatory = $true)]
            [System.String]$DiskState,

            [Parameter(Mandatory = $true)]
            [System.String]$OriginalSize,

            [Parameter(Mandatory = $true)]
            [System.String]$FinalSize,

            [Parameter(Mandatory = $true)]
            [System.String]$FullName,

            [Parameter(Mandatory = $true)]
            [datetime]$StartTime,

            [Parameter(Mandatory = $true)]
            [datetime]$EndTime,

            [Parameter(Mandatory = $true)]
            [Switch]$Passthru
        )

        BEGIN {
            Set-StrictMode -Version Latest
        }
        PROCESS {
            # Build the standardised result object with human-readable values
            $output = [PSCustomObject]@{
                Name             = $Name
                StartTime        = $StartTime.ToLongTimeString()
                EndTime          = $EndTime.ToLongTimeString()
                'ElapsedTime(s)' = [math]::Round(($EndTime - $StartTime).TotalSeconds, 1)
                DiskState        = $DiskState
                OriginalSizeGB   = [math]::Round( $OriginalSize / 1GB, 2 )
                FinalSizeGB      = [math]::Round( $FinalSize / 1GB, 2 )
                FullName         = $FullName
            }

            # Emit to pipeline if -PassThru was specified
            if ($Passthru) {
                Write-Output $output
            }

            # Write to CSV with retry loop - multiple threads may contend on the same file
            $success = $False
            $retries = 0
            while ($retries -lt 10 -and $success -ne $true) {
                try {
                    $output | Export-Csv -Path $Path -NoClobber -Append -ErrorAction Stop -NoTypeInformation
                    $success = $true
                }
                catch {
                    $retries++
                }
                Start-Sleep 1
            }
        }
        END { }
    }

    #endregion Helper Functions

    #region Invoke-Parallel (PowerShell 5.x support)

    # Invoke-Parallel provides multi-threading via runspace pools for PowerShell 5.x.
    # On PowerShell 7+, the native ForEach-Object -Parallel is used instead.
    # Credit: Boe Prox (runspace pool engine), T Bryce Yehl (Quiet/NoCloseOnTimeout)
    # Source: https://github.com/RamblingCookieMonster/Invoke-Parallel

    function Invoke-Parallel {
        [cmdletbinding(DefaultParameterSetName = 'ScriptBlock')]
        Param (
            [Parameter(Mandatory = $false, position = 0, ParameterSetName = 'ScriptBlock')]
            [System.Management.Automation.ScriptBlock]$ScriptBlock,

            [Parameter(Mandatory = $false, ParameterSetName = 'ScriptFile')]
            [ValidateScript( { Test-Path $_ -pathtype leaf })]
            $ScriptFile,

            [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
            [Alias('CN', '__Server', 'IPAddress', 'Server', 'ComputerName')]
            [PSObject]$InputObject,

            [PSObject]$Parameter,

            [switch]$ImportVariables,
            [switch]$ImportModules,
            [switch]$ImportFunctions,

            [int]$Throttle = 20,
            [int]$SleepTimer = 200,
            [int]$RunspaceTimeout = 0,
            [switch]$NoCloseOnTimeout = $false,
            [int]$MaxQueue,

            [validatescript( { Test-Path (Split-Path $_ -parent) })]
            [switch] $AppendLog = $false,
            [string]$LogFile,

            [switch] $Quiet = $false
        )
        begin {
            if ( -not $PSBoundParameters.ContainsKey('MaxQueue') ) {
                if ($RunspaceTimeout -ne 0) { $script:MaxQueue = $Throttle }
                else { $script:MaxQueue = $Throttle * 3 }
            }
            else {
                $script:MaxQueue = $MaxQueue
            }
            $ProgressId = Get-Random
            Write-Verbose "Throttle: '$throttle' SleepTimer '$sleepTimer' runSpaceTimeout '$runspaceTimeout' maxQueue '$maxQueue' logFile '$logFile'"

            if ($ImportVariables -or $ImportModules -or $ImportFunctions) {
                $StandardUserEnv = [powershell]::Create().addscript( {
                    $Modules   = Get-Module | Select-Object -ExpandProperty Name
                    $Snapins   = Get-PSSnapin | Select-Object -ExpandProperty Name
                    $Functions = Get-ChildItem function:\ | Select-Object -ExpandProperty Name
                    $Variables = Get-Variable | Select-Object -ExpandProperty Name
                    @{
                        Variables = $Variables
                        Modules   = $Modules
                        Snapins   = $Snapins
                        Functions = $Functions
                    }
                }, $true).invoke()[0]

                if ($ImportVariables) {
                    Function _temp { [cmdletbinding(SupportsShouldProcess = $True)] param() }
                    $VariablesToExclude = @( (Get-Command _temp | Select-Object -ExpandProperty parameters).Keys + $PSBoundParameters.Keys + $StandardUserEnv.Variables )
                    $UserVariables = @( Get-Variable | Where-Object { -not ($VariablesToExclude -contains $_.Name) } )
                }
                if ($ImportModules) {
                    $UserModules = @( Get-Module | Where-Object { $StandardUserEnv.Modules -notcontains $_.Name -and (Test-Path $_.Path -ErrorAction SilentlyContinue) } | Select-Object -ExpandProperty Path )
                    $UserSnapins = @( Get-PSSnapin | Select-Object -ExpandProperty Name | Where-Object { $StandardUserEnv.Snapins -notcontains $_ } )
                }
                if ($ImportFunctions) {
                    $UserFunctions = @( Get-ChildItem function:\ | Where-Object { $StandardUserEnv.Functions -notcontains $_.Name } )
                }
            }

            Function Get-RunspaceData {
                [cmdletbinding()]
                param( [switch]$Wait )
                Do {
                    $more = $false
                    if (-not $Quiet) {
                        Write-Progress -Id $ProgressId -Activity "Running Query" -Status "Starting threads"`
                            -CurrentOperation "$startedCount threads defined - $totalCount input objects - $script:completedCount input objects processed"`
                            -PercentComplete $( Try { $script:completedCount / $totalCount * 100 } Catch { 0 } )
                    }
                    Foreach ($runspace in $runspaces) {
                        $currentdate = Get-Date
                        $runtime = $currentdate - $runspace.startTime
                        $runMin = [math]::Round( $runtime.totalminutes, 2 )

                        $log = "" | Select-Object Date, Action, Runtime, Status, Details
                        $log.Action = "Removing:'$($runspace.object)'"
                        $log.Date = $currentdate
                        $log.Runtime = "$runMin minutes"

                        If ($runspace.Runspace.isCompleted) {
                            $script:completedCount++
                            if ($runspace.powershell.Streams.Error.Count -gt 0) {
                                $log.status = "CompletedWithErrors"
                                Write-Verbose ($log | ConvertTo-Csv -Delimiter ";" -NoTypeInformation)[1]
                                foreach ($ErrorRecord in $runspace.powershell.Streams.Error) {
                                    Write-Error -ErrorRecord $ErrorRecord
                                }
                            }
                            else {
                                $log.status = "Completed"
                                Write-Verbose ($log | ConvertTo-Csv -Delimiter ";" -NoTypeInformation)[1]
                            }
                            $runspace.powershell.EndInvoke($runspace.Runspace)
                            $runspace.powershell.dispose()
                            $runspace.Runspace = $null
                            $runspace.powershell = $null
                        }
                        ElseIf ( $runspaceTimeout -ne 0 -and $runtime.totalseconds -gt $runspaceTimeout) {
                            $script:completedCount++
                            $timedOutTasks = $true
                            $log.status = "TimedOut"
                            Write-Verbose ($log | ConvertTo-Csv -Delimiter ";" -NoTypeInformation)[1]
                            Write-Error "Runspace timed out at $($runtime.totalseconds) seconds for the object:`n$($runspace.object | out-string)"
                            if (!$noCloseOnTimeout) { $runspace.powershell.dispose() }
                            $runspace.Runspace = $null
                            $runspace.powershell = $null
                            $completedCount++
                        }
                        ElseIf ($runspace.Runspace -ne $null) {
                            $log = $null
                            $more = $true
                        }
                        if ($logFile -and $log) {
                            ($log | ConvertTo-Csv -Delimiter ";" -NoTypeInformation)[1] | out-file $LogFile -append
                        }
                    }
                    $temphash = $runspaces.clone()
                    $temphash | Where-Object { $_.runspace -eq $Null } | ForEach-Object {
                        $Runspaces.remove($_)
                    }
                    if ($PSBoundParameters['Wait']) { Start-Sleep -milliseconds $SleepTimer }
                } while ($more -and $PSBoundParameters['Wait'])
            }

            if ($PSCmdlet.ParameterSetName -eq 'ScriptFile') {
                $ScriptBlock = [scriptblock]::Create( $(Get-Content $ScriptFile | out-string) )
            }
            elseif ($PSCmdlet.ParameterSetName -eq 'ScriptBlock') {
                [string[]]$ParamsToAdd = '$_'
                if ( $PSBoundParameters.ContainsKey('Parameter') ) {
                    $ParamsToAdd += '$Parameter'
                }

                $UsingVariableData = $Null

                if ($PSVersionTable.PSVersion.Major -gt 2) {
                    $UsingVariables = $ScriptBlock.ast.FindAll( { $args[0] -is [System.Management.Automation.Language.UsingExpressionAst] }, $True)
                    If ($UsingVariables) {
                        $List = New-Object 'System.Collections.Generic.List`1[System.Management.Automation.Language.VariableExpressionAst]'
                        ForEach ($Ast in $UsingVariables) {
                            [void]$list.Add($Ast.SubExpression)
                        }
                        $UsingVar = $UsingVariables | Group-Object -Property SubExpression | ForEach-Object { $_.Group | Select-Object -First 1 }
                        $UsingVariableData = ForEach ($Var in $UsingVar) {
                            try {
                                $Value = Get-Variable -Name $Var.SubExpression.VariablePath.UserPath -ErrorAction Stop
                                [pscustomobject]@{
                                    Name       = $Var.SubExpression.Extent.Text
                                    Value      = $Value.Value
                                    NewName    = ('$__using_{0}' -f $Var.SubExpression.VariablePath.UserPath)
                                    NewVarName = ('__using_{0}' -f $Var.SubExpression.VariablePath.UserPath)
                                }
                            }
                            catch {
                                Write-Error "$($Var.SubExpression.Extent.Text) is not a valid Using: variable!"
                            }
                        }
                        $ParamsToAdd += $UsingVariableData | Select-Object -ExpandProperty NewName -Unique

                        $NewParams = $UsingVariableData.NewName -join ', '
                        $Tuple = [Tuple]::Create($list, $NewParams)
                        $bindingFlags = [Reflection.BindingFlags]"Default,NonPublic,Instance"
                        $GetWithInputHandlingForInvokeCommandImpl = ($ScriptBlock.ast.gettype().GetMethod('GetWithInputHandlingForInvokeCommandImpl', $bindingFlags))
                        $StringScriptBlock = $GetWithInputHandlingForInvokeCommandImpl.Invoke($ScriptBlock.ast, @($Tuple))
                        $ScriptBlock = [scriptblock]::Create($StringScriptBlock)
                        Write-Verbose $StringScriptBlock
                    }
                }

                $ScriptBlock = $ExecutionContext.InvokeCommand.NewScriptBlock("param($($ParamsToAdd -Join ", "))`r`n" + $Scriptblock.ToString())
            }
            else {
                Throw "Must provide ScriptBlock or ScriptFile"; Break
            }

            Write-Debug "`$ScriptBlock: $($ScriptBlock | Out-String)"
            Write-Verbose "Creating runspace pool and session states"

            $sessionstate = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
            if ($ImportVariables -and $UserVariables.count -gt 0) {
                foreach ($Variable in $UserVariables) {
                    $sessionstate.Variables.Add((New-Object -TypeName System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $Variable.Name, $Variable.Value, $null))
                }
            }
            if ($ImportModules) {
                if ($UserModules.count -gt 0) {
                    foreach ($ModulePath in $UserModules) {
                        $sessionstate.ImportPSModule($ModulePath)
                    }
                }
                if ($UserSnapins.count -gt 0) {
                    foreach ($PSSnapin in $UserSnapins) {
                        [void]$sessionstate.ImportPSSnapIn($PSSnapin, [ref]$null)
                    }
                }
            }
            if ($ImportFunctions -and $UserFunctions.count -gt 0) {
                foreach ($FunctionDef in $UserFunctions) {
                    $sessionstate.Commands.Add((New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList $FunctionDef.Name, $FunctionDef.ScriptBlock))
                }
            }

            $runspacepool = [runspacefactory]::CreateRunspacePool(1, $Throttle, $sessionstate, $Host)
            $runspacepool.Open()

            Write-Verbose "Creating empty collection to hold runspace jobs"
            $Script:runspaces = New-Object System.Collections.ArrayList

            $bound = $PSBoundParameters.keys -contains "InputObject"
            if (-not $bound) {
                [System.Collections.ArrayList]$allObjects = @()
            }

            if ( $LogFile -and (-not (Test-Path $LogFile) -or $AppendLog -eq $false)) {
                New-Item -ItemType file -Path $logFile -Force | Out-Null
                ("" | Select-Object -Property Date, Action, Runtime, Status, Details | ConvertTo-Csv -NoTypeInformation -Delimiter ";")[0] | Out-File $LogFile
            }

            $log = "" | Select-Object -Property Date, Action, Runtime, Status, Details
            $log.Date = Get-Date
            $log.Action = "Batch processing started"
            $log.Runtime = $null
            $log.Status = "Started"
            $log.Details = $null
            if ($logFile) {
                ($log | convertto-csv -Delimiter ";" -NoTypeInformation)[1] | Out-File $LogFile -Append
            }
            $timedOutTasks = $false
        }
        process {
            if ($bound) {
                $allObjects = $InputObject
            }
            else {
                [void]$allObjects.add( $InputObject )
            }
        }
        end {
            try {
                $totalCount = $allObjects.count
                $script:completedCount = 0
                $startedCount = 0
                foreach ($object in $allObjects) {
                    $powershell = [powershell]::Create()

                    if ($VerbosePreference -eq 'Continue') {
                        [void]$PowerShell.AddScript( { $VerbosePreference = 'Continue' })
                    }

                    [void]$PowerShell.AddScript($ScriptBlock).AddArgument($object)

                    if ($parameter) {
                        [void]$PowerShell.AddArgument($parameter)
                    }

                    if ($UsingVariableData) {
                        Foreach ($UsingVariable in $UsingVariableData) {
                            Write-Verbose "Adding $($UsingVariable.Name) with value: $($UsingVariable.Value)"
                            [void]$PowerShell.AddArgument($UsingVariable.Value)
                        }
                    }

                    $powershell.RunspacePool = $runspacepool

                    $temp = "" | Select-Object PowerShell, StartTime, object, Runspace
                    $temp.PowerShell = $powershell
                    $temp.StartTime = Get-Date
                    $temp.object = $object
                    $temp.Runspace = $powershell.BeginInvoke()
                    $startedCount++

                    Write-Verbose ( "Adding {0} to collection at {1}" -f $temp.object, $temp.starttime.tostring() )
                    $runspaces.Add($temp) | Out-Null

                    Get-RunspaceData

                    $firstRun = $true
                    while ($runspaces.count -ge $Script:MaxQueue) {
                        if ($firstRun) {
                            Write-Verbose "$($runspaces.count) items running - exceeded $Script:MaxQueue limit."
                        }
                        $firstRun = $false
                        Get-RunspaceData
                        Start-Sleep -Milliseconds $sleepTimer
                    }
                }
                Write-Verbose ( "Finish processing the remaining runspace jobs: {0}" -f ( @($runspaces | Where-Object { $_.Runspace -ne $Null }).Count) )
                Get-RunspaceData -wait
                if (-not $quiet) {
                    Write-Progress -Id $ProgressId -Activity "Running Query" -Status "Starting threads" -Completed
                }
            }
            finally {
                if ( ($timedOutTasks -eq $false) -or ( ($timedOutTasks -eq $true) -and ($noCloseOnTimeout -eq $false) ) ) {
                    Write-Verbose "Closing the runspace pool"
                    $runspacepool.close()
                }
                [gc]::Collect()
            }
        }
    }

    #endregion Invoke-Parallel

    # ---- Prerequisite Checks ----

    # Validate required Windows services are running.
    # vds (Virtual Disk Service) is needed for disk mount/dismount operations.
    # defragsvc (Optimize Drives) is needed for volume operations.
    $servicesToTest = 'defragsvc', 'vds'
    try {
        $servicesToTest | Test-FslDependencies -ErrorAction Stop
    }
    catch {
        $err = $error[0]
        Write-Error $err
        return
    }

    # Cap the thread count at 2x the number of logical cores to prevent
    # overwhelming the disk subsystem (which is metadata-heavy, not IOPS-heavy)
    $numberOfCores = (Get-CimInstance Win32_ComputerSystem).NumberOfLogicalProcessors
    If (($ThrottleLimit / 2) -gt $numberOfCores) {
        $ThrottleLimit = $numberOfCores * 2
        Write-Warning "Number of threads set to double the number of cores - $ThrottleLimit"
    }

} # Begin

# ============================================================================
# PROCESS Block - Discover disks, build scriptblocks, dispatch parallel work.
# This block runs once per piped input (or once if -Path is specified directly).
# ============================================================================
PROCESS {

    # Check that the path is valid
    if (-not (Test-Path $Path)) {
        Write-Error "$Path not found"
        return
    }

    # Get a list of Virtual Hard Disk files depending on the -Recurse parameter
    if ($Recurse) {
        $diskList = Get-ChildItem -File -Filter *.vhd? -Path $Path -Recurse
    }
    else {
        $diskList = Get-ChildItem -File -Filter *.vhd? -Path $Path
    }

    # Exclude FSLogix diff/merge disks - these are internal to the FSLogix agent
    # and should never be repaired directly (Merge.vhdx = compaction target,
    # RW.vhdx = active differencing disk)
    $diskList = $diskList | Where-Object { $_.Name -ne "Merge.vhdx" -and $_.Name -ne "RW.vhdx" }

    if ( ($diskList | Measure-Object).count -eq 0 ) {
        Write-Warning "No files to process in $Path"
        return
    }

    # Scriptblock for ForEach-Object -Parallel (PowerShell 7+)
    # ForEach-Object -Parallel runs each iteration in an isolated runspace, so it
    # cannot see functions defined in the parent scope. All helper functions must be
    # redefined inside the scriptblock. These are compact copies of the functions
    # defined in the BEGIN block above - see those for full documentation.
    $scriptblockForEachObject = {

        #region Inline Helper Functions (required for ForEach-Object -Parallel)
        # These are identical to the functions in the BEGIN block but must be
        # duplicated here because -Parallel runspaces have no access to the
        # caller's function table.

        # Test-FslDependencies: Ensures required Windows services are running
        Function Test-FslDependencies {
            [CmdletBinding()]
            Param (
                [Parameter(Mandatory = $true, Position = 0, ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true)]
                [System.String[]]$Name
            )
            BEGIN { Set-StrictMode -Version Latest }
            PROCESS {
                Foreach ($svc in $Name) {
                    $svcObject = Get-Service -Name $svc
                    If ($svcObject.Status -eq "Running") { Return }
                    If ($svcObject.StartType -eq "Disabled") {
                        Write-Warning ("[{0}] Setting Service to Manual" -f $svcObject.DisplayName)
                        Set-Service -Name $svc -StartupType Manual | Out-Null
                    }
                    Start-Service -Name $svc | Out-Null
                    if ((Get-Service -Name $svc).Status -ne 'Running') {
                        Write-Error "Can not start $($svcObject.DisplayName)"
                    }
                }
            }
            END { }
        }

        # Mount-FslDisk: Mounts VHD/VHDX to a GUID-named temp dir (no drive letter)
        function Mount-FslDisk {
            [CmdletBinding()]
            Param (
                [Parameter(Position = 1, ValuefromPipelineByPropertyName = $true, ValuefromPipeline = $true, Mandatory = $true)]
                [alias('FullName')]
                [System.String]$Path,

                [Parameter(ValuefromPipelineByPropertyName = $true, ValuefromPipeline = $true)]
                [Int]$TimeOut = 3,

                [Parameter(ValuefromPipelineByPropertyName = $true)]
                [Switch]$PassThru
            )
            BEGIN { Set-StrictMode -Version Latest }
            PROCESS {
                try {
                    $mountedDisk = Mount-DiskImage -ImagePath $Path -NoDriveLetter -PassThru -ErrorAction Stop
                }
                catch {
                    $e = $error[0]
                    Write-Error "Failed to mount disk - `"$e`""
                    return
                }
                $diskNumber = $false
                $timespan = (Get-Date).AddSeconds($TimeOut)
                while ($diskNumber -eq $false -and $timespan -gt (Get-Date)) {
                    Start-Sleep 0.1
                    try {
                        $mountedDisk = Get-DiskImage -ImagePath $Path
                        if ($mountedDisk.Number) { $diskNumber = $true }
                    }
                    catch { $diskNumber = $false }
                }
                if ($diskNumber -eq $false) {
                    try { $mountedDisk | Dismount-DiskImage -ErrorAction SilentlyContinue } catch { Write-Error 'Could not dismount Disk Due to no Disknumber' }
                    Write-Error 'Cannot get mount information'
                    return
                }
                $partitionType = $false
                $timespan = (Get-Date).AddSeconds($TimeOut)
                while ($partitionType -eq $false -and $timespan -gt (Get-Date)) {
                    try {
                        $allPartition = Get-Partition -DiskNumber $mountedDisk.Number -ErrorAction Stop
                        if ($allPartition.Type -contains 'Basic') {
                            $partitionType = $true
                            $partition = $allPartition | Where-Object -Property 'Type' -EQ -Value 'Basic'
                        }
                    }
                    catch {
                        if (($allPartition | Measure-Object).Count -gt 0) {
                            $partition = $allPartition | Select-Object -Last 1
                            $partitionType = $true
                        }
                        else { $partitionType = $false }
                    }
                    Start-Sleep 0.1
                }
                if ($partitionType -eq $false) {
                    try { $mountedDisk | Dismount-DiskImage -ErrorAction SilentlyContinue } catch { Write-Error 'Could not dismount disk with no partition' }
                    Write-Error 'Cannot get partition information'
                    return
                }
                $tempGUID = [guid]::NewGuid().ToString()
                $mountPath = Join-Path $Env:Temp ('FSLogixMnt-' + $tempGUID)
                try {
                    New-Item -Path $mountPath -ItemType Directory -ErrorAction Stop | Out-Null
                }
                catch {
                    $e = $error[0]
                    try { $mountedDisk | Dismount-DiskImage -ErrorAction SilentlyContinue } catch { Write-Error "Could not dismount disk when no folder could be created - `"$e`"" }
                    Write-Error "Failed to create mounting directory - `"$e`""
                    return
                }
                try {
                    $addPartitionAccessPathParams = @{
                        DiskNumber      = $mountedDisk.Number
                        PartitionNumber = $partition.PartitionNumber
                        AccessPath      = $mountPath
                        ErrorAction     = 'Stop'
                    }
                    Add-PartitionAccessPath @addPartitionAccessPathParams
                }
                catch {
                    $e = $error[0]
                    Remove-Item -Path $mountPath -Force -Recurse -ErrorAction SilentlyContinue
                    try { $mountedDisk | Dismount-DiskImage -ErrorAction SilentlyContinue } catch { Write-Error "Could not dismount disk when no junction point could be created - `"$e`"" }
                    Write-Error "Failed to create junction point to - `"$e`""
                    return
                }
                if ($PassThru) {
                    $output = [PSCustomObject]@{
                        Path            = $mountPath
                        DiskNumber      = $mountedDisk.Number
                        ImagePath       = $mountedDisk.ImagePath
                        PartitionNumber = $partition.PartitionNumber
                    }
                    Write-Output $output
                }
                Write-Verbose "Mounted $Path to $mountPath"
            }
            END { }
        }

        # Dismount-FslDisk: Removes temp dir and detaches disk image with retry
        function Dismount-FslDisk {
            [CmdletBinding()]
            Param (
                [Parameter(Position = 1, ValuefromPipelineByPropertyName = $true, ValuefromPipeline = $true, Mandatory = $true)]
                [String]$Path,
                [Parameter(ValuefromPipelineByPropertyName = $true, Mandatory = $true)]
                [String]$ImagePath,
                [Parameter(ValuefromPipelineByPropertyName = $true)]
                [Switch]$PassThru,
                [Parameter(ValuefromPipelineByPropertyName = $true)]
                [Int]$Timeout = 120
            )
            BEGIN { Set-StrictMode -Version Latest }
            PROCESS {
                $mountRemoved = $false
                $directoryRemoved = $false
                $timeStampDirectory = (Get-Date).AddSeconds(20)
                while ((Get-Date) -lt $timeStampDirectory -and $directoryRemoved -ne $true) {
                    try { Remove-Item -Path $Path -Force -Recurse -ErrorAction Stop | Out-Null; $directoryRemoved = $true }
                    catch { $directoryRemoved = $false }
                }
                if (Test-Path $Path) { Write-Warning "Failed to delete temp mount directory $Path" }
                $timeStampDismount = (Get-Date).AddSeconds($Timeout)
                while ((Get-Date) -lt $timeStampDismount -and $mountRemoved -ne $true) {
                    try {
                        Dismount-DiskImage -ImagePath $ImagePath -ErrorAction Stop | Out-Null
                        try {
                            $image = Get-DiskImage -ImagePath $ImagePath -ErrorAction Stop
                            switch ($image.Attached) {
                                $null  { $mountRemoved = $false ; Start-Sleep 0.1; break }
                                $true  { $mountRemoved = $false ; break }
                                $false { $mountRemoved = $true ; break }
                                Default { $mountRemoved = $false }
                            }
                        }
                        catch { $mountRemoved = $false }
                    }
                    catch { $mountRemoved = $false }
                }
                if ($mountRemoved -ne $true) { Write-Error "Failed to dismount disk $ImagePath" }
                If ($PassThru) {
                    $output = [PSCustomObject]@{ MountRemoved = $mountRemoved; DirectoryRemoved = $directoryRemoved }
                    Write-Output $output
                }
                if ($directoryRemoved -and $mountRemoved) { Write-Verbose "Dismounted $ImagePath" }
            }
            END { }
        }

        # Repair-OneDisk: Core per-disk logic - dirty check, chkdsk, fallback, log
        function Repair-OneDisk {
            [CmdletBinding()]
            Param (
                [Parameter(ValuefromPipelineByPropertyName = $true, ValuefromPipeline = $true, Mandatory = $true)]
                [System.IO.FileInfo]$Disk,
                [Parameter(ValuefromPipelineByPropertyName = $true)]
                [int]$MountTimeout = 30,
                [Parameter(ValuefromPipelineByPropertyName = $true)]
                [string]$LogFilePath = "$env:TEMP\FslRepairDisk $(Get-Date -Format yyyy-MM-dd` HH-mm-ss).csv",
                [Parameter(ValuefromPipelineByPropertyName = $true)]
                [switch]$Passthru,
                [Parameter(ValuefromPipelineByPropertyName = $true)]
                [switch]$ForceRepair
            )
            BEGIN { Set-StrictMode -Version Latest }
            PROCESS {
                Dismount-DiskImage -ImagePath $Disk.FullName -ErrorAction SilentlyContinue
                $startTime = Get-Date
                $originalSize = $Disk.Length
                $PSDefaultParameterValues = @{
                    "Write-VhdOutput:Path"         = $LogFilePath
                    "Write-VhdOutput:StartTime"    = $startTime
                    "Write-VhdOutput:Name"         = $Disk.Name
                    "Write-VhdOutput:DiskState"    = $null
                    "Write-VhdOutput:OriginalSize" = $originalSize
                    "Write-VhdOutput:FinalSize"    = $originalSize
                    "Write-VhdOutput:FullName"     = $Disk.FullName
                    "Write-VhdOutput:Passthru"     = $Passthru
                }
                if ($Disk.Extension -ne '.vhd' -and $Disk.Extension -ne '.vhdx') {
                    Write-VhdOutput -DiskState 'FileIsNotDiskFormat' -EndTime (Get-Date)
                    return
                }
                try {
                    $mount = Mount-FslDisk -Path $Disk.FullName -TimeOut $MountTimeout -PassThru -ErrorAction Stop
                }
                catch {
                    $err = $error[0]
                    if ($err -match 'disk is already in use' -or $err -match 'being used by another process' -or $err -match 'locked') {
                        Write-VhdOutput -DiskState 'DiskLocked' -EndTime (Get-Date)
                    }
                    else {
                        Write-VhdOutput -DiskState "MountFailed" -EndTime (Get-Date)
                    }
                    return
                }
                $timespan = (Get-Date).AddSeconds(120)
                $partInfo = $null
                while (($partInfo | Measure-Object).Count -lt 1 -and $timespan -gt (Get-Date)) {
                    try {
                        $partInfo = Get-Partition -DiskNumber $mount.DiskNumber -ErrorAction Stop |
                            Where-Object -Property 'Type' -EQ -Value 'Basic' -ErrorAction Stop
                    }
                    catch {
                        $partInfo = Get-Partition -DiskNumber $mount.DiskNumber -ErrorAction SilentlyContinue |
                            Select-Object -Last 1
                    }
                    Start-Sleep 0.1
                }
                if (($partInfo | Measure-Object).Count -eq 0) {
                    $mount | Dismount-FslDisk
                    Write-VhdOutput -DiskState 'NoPartitionInfo' -EndTime (Get-Date)
                    return
                }
                $volumePath = $mount.Path
                $isDirty = $false
                try {
                    $dirtyResult = & fsutil dirty query "$volumePath" 2>&1
                    $dirtyOutput = $dirtyResult -join ' '
                    if ($dirtyOutput -match 'is Dirty' -or $dirtyOutput -match 'dirty') { $isDirty = $true }
                    elseif ($dirtyOutput -match 'NOT Dirty' -or $dirtyOutput -match 'not dirty') { $isDirty = $false }
                    else {
                        try {
                            $vol = Get-Volume -Partition $partInfo -ErrorAction Stop
                            $scanResult = Repair-Volume -InputObject $vol -Scan -ErrorAction Stop
                            if ($scanResult -eq 'ScanNeeded' -or $scanResult -eq 'SpotFixesNeeded') { $isDirty = $true }
                        }
                        catch {
                            if ($ForceRepair) { $isDirty = $true }
                            else {
                                $mount | Dismount-FslDisk
                                Write-VhdOutput -DiskState 'Skipped' -EndTime (Get-Date)
                                return
                            }
                        }
                    }
                }
                catch {
                    if ($ForceRepair) { $isDirty = $true }
                    else {
                        $mount | Dismount-FslDisk
                        Write-VhdOutput -DiskState 'Skipped' -EndTime (Get-Date)
                        return
                    }
                }
                if (-not $isDirty -and -not $ForceRepair) {
                    $mount | Dismount-FslDisk
                    Write-VhdOutput -DiskState 'NotDirty' -EndTime (Get-Date)
                    return
                }
                $repairSuccess = $false
                try {
                    $chkdskPath = $volumePath.TrimEnd('\') + '\'
                    $chkdskOutput = & chkdsk $chkdskPath /f /x 2>&1
                    $chkdskText = $chkdskOutput -join "`n"
                    if ($LASTEXITCODE -eq 0) { $repairSuccess = $true }
                    else {
                        if ($chkdskText -match 'Windows has made corrections' -or $chkdskText -match 'no problems' -or $chkdskText -match 'Windows has checked the file system' -or $chkdskText -match 'cleaning up') {
                            $repairSuccess = $true
                        }
                    }
                }
                catch { $repairSuccess = $false }
                if (-not $repairSuccess) {
                    try {
                        $vol = Get-Volume -Partition $partInfo -ErrorAction Stop
                        $repairResult = Repair-Volume -InputObject $vol -OfflineScanAndFix -ErrorAction Stop
                        if ($repairResult -eq 'NoErrorsFound' -or $repairResult -eq 'Fixed') { $repairSuccess = $true }
                    }
                    catch { $repairSuccess = $false }
                }
                $mount | Dismount-FslDisk
                if ($repairSuccess) {
                    $state = if ($isDirty) { 'Repaired' } else { 'Checked' }
                    Write-VhdOutput -DiskState $state -FinalSize (Get-ChildItem $Disk.FullName | Select-Object -ExpandProperty Length) -EndTime (Get-Date)
                }
                else {
                    Write-VhdOutput -DiskState 'RepairFailed' -EndTime (Get-Date)
                }
            }
            END { }
        }

        # Write-VhdOutput: Logs result to CSV with thread-safe retry
        function Write-VhdOutput {
            [CmdletBinding()]
            Param (
                [Parameter(Mandatory = $true)][System.String]$Path,
                [Parameter(Mandatory = $true)][System.String]$Name,
                [Parameter(Mandatory = $true)][System.String]$DiskState,
                [Parameter(Mandatory = $true)][System.String]$OriginalSize,
                [Parameter(Mandatory = $true)][System.String]$FinalSize,
                [Parameter(Mandatory = $true)][System.String]$FullName,
                [Parameter(Mandatory = $true)][datetime]$StartTime,
                [Parameter(Mandatory = $true)][datetime]$EndTime,
                [Parameter(Mandatory = $true)][Switch]$Passthru
            )
            BEGIN { Set-StrictMode -Version Latest }
            PROCESS {
                $output = [PSCustomObject]@{
                    Name             = $Name
                    StartTime        = $StartTime.ToLongTimeString()
                    EndTime          = $EndTime.ToLongTimeString()
                    'ElapsedTime(s)' = [math]::Round(($EndTime - $StartTime).TotalSeconds, 1)
                    DiskState        = $DiskState
                    OriginalSizeGB   = [math]::Round( $OriginalSize / 1GB, 2 )
                    FinalSizeGB      = [math]::Round( $FinalSize / 1GB, 2 )
                    FullName         = $FullName
                }
                if ($Passthru) { Write-Output $output }
                $success = $False
                $retries = 0
                while ($retries -lt 10 -and $success -ne $true) {
                    try {
                        $output | Export-Csv -Path $Path -NoClobber -Append -ErrorAction Stop -NoTypeInformation
                        $success = $true
                    }
                    catch { $retries++ }
                    Start-Sleep 1
                }
            }
            END { }
        }

        #endregion Inline Helper Functions

        $paramRepairOneDisk = @{
            Disk             = $_
            LogFilePath      = $using:LogFilePath
            PassThru         = $using:PassThru
            ForceRepair      = $using:ForceRepair
        }
        Repair-OneDisk @paramRepairOneDisk
    }

    # Scriptblock for Invoke-Parallel (PowerShell 5.x)
    # This scriptblock uses $Using: syntax is NOT needed here because Invoke-Parallel
    # imports variables from the parent scope via -ImportVariables.
    $scriptblockInvokeParallel = {

        $disk = $_

        $paramRepairOneDisk = @{
            Disk             = $disk
            LogFilePath      = $LogFilePath
            PassThru         = $PassThru
            ForceRepair      = $ForceRepair
        }
        Repair-OneDisk @paramRepairOneDisk
    }

    # Dispatch: use native ForEach-Object -Parallel on PS 7+, fall back to
    # Invoke-Parallel (runspace pools) on PS 5.x
    if ($PSVersionTable.PSVersion -ge [version]"7.0") {
        $diskList | ForEach-Object -Parallel $scriptblockForEachObject -ThrottleLimit $ThrottleLimit
    }
    else {
        $diskList | Invoke-Parallel -ScriptBlock $scriptblockInvokeParallel -Throttle $ThrottleLimit -ImportFunctions -ImportVariables -ImportModules
    }

} # Process

# ============================================================================
# END Block - No cleanup needed; runspace pools are closed in Invoke-Parallel's
# finally block, and ForEach-Object -Parallel handles its own cleanup.
# ============================================================================
END { } # End
