<#
.SYNOPSIS
    Decommissions outdated AVD session hosts from a host pool.

.DESCRIPTION
    Compares each drained session host's image version against the latest
    published version in the Azure Compute Gallery. Outdated hosts are fully
    removed: AVD registration, Entra ID device record (if applicable),
    VM, NIC, and OS disk.

    Safety mechanisms:
      - Only processes hosts that are already drained (AllowNewSession = false)
      - Grace period for hosts with active sessions (tag-based countdown)
      - Stale AVD records (VM deleted but registration remains) are cleaned up
      - Simulate mode for dry-run validation

.PARAMETER HostPoolName
    Name of the AVD host pool to scan.

.PARAMETER HostPoolRg
    Resource group containing the host pool object (core RG).

.PARAMETER ComputeRg
    Resource group containing the session host VMs (compute RG).

.PARAMETER GalleryRg
    Resource group of the Azure Compute Gallery.

.PARAMETER GalleryName
    Name of the Azure Compute Gallery.

.PARAMETER ImageDef
    Gallery image definition name to resolve the latest version from.

.PARAMETER Simulate
    When $true, logs what would be deleted without making changes (dry run).

.PARAMETER DrainGracePeriodHours
    Hours to wait before force-decommissioning a host that still has sessions.
    Default: 24 hours. Uses a 'PendingDrainTimestamp' tag on the VM.

.NOTES
    File:     scripts/Remove-AvdHosts.ps1
    Requires: Az.Compute, Az.DesktopVirtualization, Az.Resources modules
    Called by: avd-cleanup-vdpool-*.yml pipelines

.EXAMPLE
    .\Remove-AvdHosts.ps1 -HostPoolName "vdpool-mypool-001" `
                          -HostPoolRg "rg-avd-core" `
                          -ComputeRg "rg-avd-compute" `
                          -GalleryRg "rg-avd-gallery" `
                          -GalleryName "gal_avd_sim_ams" `
                          -ImageDef "win11-25h2-ms-generic" `
                          -Simulate $true
#>
param(
    [string]$HostPoolName,
    [string]$HostPoolRg,
    [string]$ComputeRg,
    [string]$GalleryRg,
    [string]$GalleryName,
    [string]$ImageDef,
    [bool]$Simulate = $false,
    [int]$DrainGracePeriodHours = 24
)

$ErrorActionPreference = "Stop"

# Counters for summary
$DecommissionCount = 0
$SkipCount = 0
$StaleCount = 0
$ErrorCount = 0

try {

# ─────────────────────────────────────────────────────────────────────────────
# 1. IDENTIFY "GOLD" STANDARD
#    Fetch all image versions from the gallery, exclude any marked
#    ExcludeFromLatest, and pick the most recently published one.
# ─────────────────────────────────────────────────────────────────────────────
Write-Host " [INIT] Fetching latest version for '$ImageDef'..." -ForegroundColor Cyan

try {
    $AllVersions = Get-AzGalleryImageVersion -ResourceGroupName $GalleryRg `
                                             -GalleryName $GalleryName `
                                             -GalleryImageDefinitionName $ImageDef `
                                             -ErrorAction Stop

    # Sort by PublishedDate (consistent with Get-AvdDetails.ps1), filter ExcludeFromLatest
    $LatestVersionObj = $AllVersions | 
                        Where-Object { $_.PublishingProfile.ExcludeFromLatest -ne $true } |
                        Sort-Object -Property @{Expression={$_.PublishingProfile.PublishedDate}; Descending=$true} | 
                        Select-Object -First 1
    $TargetVersion = $LatestVersionObj.Name
    Write-Host " [INIT] Target Version: $TargetVersion" -ForegroundColor Cyan
} catch {
    Write-Error "Could not determine target image version. Check Gallery/ImageDef names."
    exit 1
}

# ─────────────────────────────────────────────────────────────────────────────
# 2. SCAN HOST POOL
#    Iterate all session hosts and evaluate each for decommission eligibility.
# ─────────────────────────────────────────────────────────────────────────────
Write-Host " [SCAN] Scanning Host Pool: $HostPoolName"
$Hosts = Get-AzWvdSessionHost -ResourceGroupName $HostPoolRg -HostPoolName $HostPoolName

foreach ($SessionHost in $Hosts) {
    # AVD Name Format: "HostPoolName/VmName.domain.com"
    $AvdDbName = $SessionHost.Name 
    
    # Extract the leaf name for AVD removal commands (e.g. "avd0706.contoso.com")
    $SessionHostLeaf = $AvdDbName.Split('/')[-1]

    # Extract the short VM name for Azure resource lookups (e.g. "avd0706")
    $VmName = $SessionHostLeaf.Split('.')[0]

    # ── CHECK 1: IS DRAINED? ────────────────────────────────────────────────
    # Only process hosts that have been drained (AllowNewSession = false).
    # Active hosts are never touched by this script.
    if ($SessionHost.AllowNewSession -eq $true) {
        $SkipCount++
        continue
    }

    # ── CHECK 2: VM EXISTS? (Clean stale AVD records) ────────────────────────
    # If the VM has been deleted outside this pipeline, the AVD registration
    # becomes orphaned. Clean it up to avoid host pool clutter.
    $Vm = Get-AzVM -ResourceGroupName $ComputeRg -Name $VmName -ErrorAction SilentlyContinue

    if ($null -eq $Vm) {
        Write-Host "   [STALE] AVD Record '$VmName' exists, but Azure VM is missing. Cleaning up registration." -ForegroundColor DarkYellow
        try {
            Remove-AzWvdSessionHost -ResourceGroupName $HostPoolRg `
                                    -HostPoolName $HostPoolName `
                                    -Name $SessionHostLeaf `
                                    -Force -ErrorAction Stop | Out-Null
            Write-Host "   [OK] Stale AVD record removed for '$VmName'." -ForegroundColor Green
            $StaleCount++
        } catch {
            Write-Host "   [ERR] Failed to remove stale record: $($_.Exception.Message)" -ForegroundColor Red
            $ErrorCount++
        }
        continue
    }

    # ── CHECK 3: ACTIVE SESSIONS + GRACE PERIOD ─────────────────────────────
    # If users are still connected, set a drain timestamp tag on first
    # encounter. Subsequent runs check if the grace period has elapsed
    # before proceeding with forced decommission.
    if ($SessionHost.Sessions -gt 0) {
        $DrainTag = $Vm.Tags['PendingDrainTimestamp']
        if (-not $DrainTag) {
            # First encounter — stamp the VM with the current UTC time
            Write-Host "   [DRAIN] $VmName has $($SessionHost.Sessions) sessions. Setting drain timestamp." -ForegroundColor Yellow
            Update-AzTag -ResourceId $Vm.Id -Tag @{ PendingDrainTimestamp = (Get-Date -AsUTC).ToString('o') } -Operation Merge | Out-Null
            $SkipCount++
            continue
        }
        
        try {
            $GraceExpiry = [DateTime]::Parse($DrainTag).AddHours($DrainGracePeriodHours)
        } catch {
            # Invalid tag value — treat as freshly set
            $GraceExpiry = (Get-Date).AddHours($DrainGracePeriodHours)
        }

        if ((Get-Date) -lt $GraceExpiry) {
            Write-Host "   [WAIT] $VmName has $($SessionHost.Sessions) sessions. Grace period until $($GraceExpiry.ToString('u'))." -ForegroundColor Yellow
            $SkipCount++
            continue
        }
        Write-Host "   [EXPIRED] $VmName grace period exceeded. Proceeding with decommission despite $($SessionHost.Sessions) sessions." -ForegroundColor DarkYellow
    }

    # ── CHECK 4: OUTDATED IMAGE VERSION? ─────────────────────────────────────
    # Compare the VM's 'ImageVersion' tag against the gallery target.
    # Untagged or mismatched VMs are decommissioned.
    $CurrentTag = $Vm.Tags['ImageVersion']
    
    if ([string]::IsNullOrWhiteSpace($CurrentTag) -or $CurrentTag -ne $TargetVersion) {
        
        Write-Host "-------------------------------------------------------------"
        Write-Host " DECOMMISSIONING: $VmName" -ForegroundColor Magenta
        Write-Host "   Version: '$($CurrentTag)' -> Target: '$TargetVersion'"
        Write-Host "-------------------------------------------------------------"

        # ── DRY RUN ──────────────────────────────────────────────────────────
        if ($Simulate) {
            Write-Host "   [DRY RUN] Would delete resources." -ForegroundColor Cyan
            continue
        }

        # ── A. REMOVE ENTRA ID DEVICE RECORD ────────────────────────────────
        # Only applicable for Entra ID joined hosts (tagged Directory=EntraID).
        # Prevents orphaned device objects in Entra ID.
        $DirectoryTag = $Vm.Tags['Directory']
        if ($DirectoryTag -eq 'EntraID') {
            Write-Host "   -> Removing Entra ID device record..." -NoNewline
            try {
                $EntraDevice = Get-AzADDevice -DisplayName $VmName -ErrorAction SilentlyContinue
                if ($EntraDevice) {
                    Remove-AzADDevice -ObjectId $EntraDevice.Id -ErrorAction Stop | Out-Null
                    Write-Host " Done." -ForegroundColor Green
                } else {
                    Write-Host " Not found (OK)." -ForegroundColor DarkGray
                }
            } catch {
                Write-Host " Failed: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }

        # ── B. REMOVE AVD HOST REGISTRATION ─────────────────────────────────
        Write-Host "   -> Removing AVD Host Registration..." -NoNewline
        try {
            Remove-AzWvdSessionHost -ResourceGroupName $HostPoolRg `
                                    -HostPoolName $HostPoolName `
                                    -Name $SessionHostLeaf `
                                    -Force -ErrorAction Stop | Out-Null
            Write-Host " Done." -ForegroundColor Green
        } catch {
            Write-Host " Failed/Already Gone." -ForegroundColor Yellow
        }

        # ── C. DELETE AZURE RESOURCES (VM, NIC, OS Disk) ─────────────────────
        Write-Host "   -> Deleting VM..." -NoNewline
        Remove-AzVM -ResourceGroupName $ComputeRg -Name $VmName -Force -ErrorAction SilentlyContinue | Out-Null
        Write-Host " Done." -ForegroundColor Green

        # Capture resource IDs before deletion (from the VM object we already have)
        $NicId = $Vm.NetworkProfile.NetworkInterfaces[0].Id
        $OsDiskId = $Vm.StorageProfile.OsDisk.ManagedDisk.Id

        if ($NicId) {
            Write-Host "   -> Deleting NIC..." -NoNewline
            Remove-AzResource -ResourceId $NicId -Force -ErrorAction SilentlyContinue | Out-Null
            Write-Host " Done." -ForegroundColor Green
        }

        if ($OsDiskId) {
            Write-Host "   -> Deleting OS Disk..." -NoNewline
            Remove-AzResource -ResourceId $OsDiskId -Force -ErrorAction SilentlyContinue | Out-Null
            Write-Host " Done." -ForegroundColor Green
        }

        Write-Host " [SUCCESS] $VmName Retired." -ForegroundColor Green
        $DecommissionCount++
    } else {
        # VM is on the latest image — nothing to do
        $SkipCount++
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# 3. SUMMARY
# ─────────────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "       CLEANUP SUMMARY" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " Decommissioned : $DecommissionCount" -ForegroundColor Green
Write-Host " Stale Cleaned  : $StaleCount" -ForegroundColor DarkYellow
Write-Host " Skipped        : $SkipCount" -ForegroundColor Gray
Write-Host " Errors         : $ErrorCount" -ForegroundColor $(if ($ErrorCount -gt 0) { 'Red' } else { 'Gray' })
Write-Host "==========================================" -ForegroundColor Cyan

if ($ErrorCount -gt 0) {
    Write-Error "Cleanup completed with $ErrorCount error(s)."
    exit 1
}

} # end try
catch {
    Write-Host " [FATAL] $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}