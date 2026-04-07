<#
.SYNOPSIS
    AVD Update Orchestrator - Pre-Flight & Parameter Calculation (Extreme Verbosity)

.DESCRIPTION
    This script acts as the primary orchestrator for the AVD Blue/Green update factory.
    It performs the following high-level tasks:
    1. Identifies the "Gold Image" (latest version) from the Azure Compute Gallery.
    2. Scans the current Host Pool to identify hosts with outdated 'ImageVersion' tags.
    3. Calculates new hostnames using an ISO-8601 Week-based naming convention (avd-WW-XX).
    4. Manages the AVD Registration Token (retrieval or 24h generation).
    5. Exports all data as JSON strings for use in downstream Azure DevOps pipeline stages.

.PARAMETER HostPoolName
    The name of the target AVD Host Pool.
.PARAMETER ResourceGroupName
    The Resource Group containing the Host Pool and Session Hosts.
.PARAMETER GalleryResourceGroupName
    The Resource Group containing the Azure Compute Gallery.
.PARAMETER GalleryName
    The name of the Azure Compute Gallery.
.PARAMETER ImageDefinitionName
    The specific Image Definition to check (e.g., img-MyImageDef).
.PARAMETER BaseName
    The prefix for generated hostnames. Defaults to 'avd'.
.PARAMETER VmCountOverride
    Optional: Force the deployment of a specific number of VMs.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)] [string]$HostPoolName,
    [Parameter(Mandatory = $true)] [string]$ResourceGroupName,
    [Parameter(Mandatory = $true)] [string]$GalleryResourceGroupName,
    [Parameter(Mandatory = $true)] [string]$GalleryName,
    [Parameter(Mandatory = $true)] [string]$ImageDefinitionName,
    [string]$BaseName = "avd",
    [int]$VmCountOverride = 0,
    [string]$GallerySubscriptionId = "",
    [string]$ComputeResourceGroupName = ""
)

# --- SILENCE INFRASTRUCTURE NOISE ---
$ProgressPreference = 'SilentlyContinue' # Hides ASCII progress bars to keep logs clean
$ErrorActionPreference = 'Stop'

# --- HELPER: VISUAL FORMATTING ---
function Write-Header {
    param (
        [string]$Title,
        [hashtable]$Info
    )
    $Width = 85
    $Border = "=" * $Width
    
    Write-Host "`n$Border" -ForegroundColor Cyan
    Write-Host "   $Title" -ForegroundColor Cyan
    Write-Host "$Border" -ForegroundColor Cyan
    
    foreach ($Key in $Info.Keys) {
        $Label = "$Key".PadRight(30)
        $Val = $Info[$Key]
        Write-Host " $Label : $Val" -ForegroundColor White
    }
    Write-Host "$Border`n" -ForegroundColor Cyan
}

function Write-Log {
    param (
        [Parameter(Mandatory = $true)] [string]$Message,
        [ValidateSet("INFO", "SUCCESS", "WARN", "ERROR", "PLAN")] [string]$Level = "INFO"
    )
    $TimeStamp = (Get-Date).ToString("HH:mm:ss")
    
    switch ($Level) {
        "INFO"    { Write-Host " [$TimeStamp] [INFO]    $Message" -ForegroundColor Gray }
        "SUCCESS" { Write-Host " [$TimeStamp] [OK]      $Message" -ForegroundColor Green }
        "WARN"    { Write-Host " [$TimeStamp] [WARN]    $Message" -ForegroundColor Yellow }
        "ERROR"   { Write-Host " [$TimeStamp] [ERROR]   $Message" -ForegroundColor Red }
        "PLAN"    { Write-Host " [$TimeStamp] [PLAN]    $Message" -ForegroundColor Magenta }
    }
}

try {
    # -------------------------------------------------------------------------
    # STEP 0: CAPTURE CONTEXT & PRINT HEADER
    # -------------------------------------------------------------------------
    # Automatically capture details of the Service Principal and Subscription
    $Context = Get-AzContext
    $HeaderInfo = [ordered]@{
        "Azure Gallery"          = $GalleryName
        "Image Definition"       = $ImageDefinitionName
        "Resource Group"         = $ResourceGroupName
        "Host Pool Name"         = $HostPoolName
        "Account / SPN"          = "***" # Masked for log security
        "Tenant ID"              = $Context.Tenant.Id
        "Subscription"           = "$($Context.Subscription.Name) ($($Context.Subscription.Id))"
        "Environment"            = $Context.Environment.Name
    }

    Write-Header -Title "AVD UPDATE FACTORY: PRE-FLIGHT ANALYSIS" -Info $HeaderInfo

    # -------------------------------------------------------------------------
    # STEP 1: IDENTIFY TARGET IMAGE (GOLD IMAGE)
    # -------------------------------------------------------------------------
    Write-Log "Querying Azure Compute Gallery for latest image version..."
    
    # Cross-subscription support: switch context if gallery is in a different subscription
    $OriginalSubscription = (Get-AzContext).Subscription.Id
    if ($GallerySubscriptionId -and $GallerySubscriptionId -ne $OriginalSubscription) {
        Write-Log "Switching to gallery subscription: $GallerySubscriptionId" "INFO"
        Set-AzContext -SubscriptionId $GallerySubscriptionId | Out-Null
    }

    # Sort Image Versions by Publish Date descending, filter out excluded versions
    $LatestVerObj = Get-AzGalleryImageVersion -ResourceGroupName $GalleryResourceGroupName `
                    -GalleryName $GalleryName `
                    -GalleryImageDefinitionName $ImageDefinitionName | 
                    Where-Object { $_.PublishingProfile.ExcludeFromLatest -ne $true } |
                    Sort-Object -Property @{Expression={$_.PublishingProfile.PublishedDate}; Descending=$true} | 
                    Select-Object -First 1

    if (-not $LatestVerObj) { throw "CRITICAL: No image version found for definition '$ImageDefinitionName'" }

    $TargetVersion = $LatestVerObj.Name
    $TargetImageId = $LatestVerObj.Id
    
    # Switch back if we changed subscriptions
    if ($GallerySubscriptionId -and $GallerySubscriptionId -ne $OriginalSubscription) {
        Write-Log "Switching back to original subscription." "INFO"
        Set-AzContext -SubscriptionId $OriginalSubscription | Out-Null
    }

    Write-Log "Found Latest Image: $TargetVersion" "SUCCESS"

    # -------------------------------------------------------------------------
    # STEP 2: ANALYZE CURRENT SESSION HOSTS
    # -------------------------------------------------------------------------
    Write-Log "Analyzing session hosts in pool '$HostPoolName'..."
    
    $CurrentHosts = Get-AzWvdSessionHost -HostPoolName $HostPoolName -ResourceGroupName $ResourceGroupName
    $CurrentCount = $CurrentHosts.Count
    
    $OutdatedHosts = @()
    $UpToDateHosts = 0

    # Variable $SH used to avoid conflict with reserved $Host variable
    foreach ($SH in $CurrentHosts) {
        try {
            # Inspect VM tags to check for version drift
            $VM = Get-AzResource -ResourceId $SH.ResourceId -ErrorAction Stop
            $VMVersion = $VM.Tags['ImageVersion']

            if ([string]::IsNullOrWhiteSpace($VMVersion) -or ($VMVersion -ne $TargetVersion)) {
                $OutdatedHosts += $SH.Name
            } else {
                $UpToDateHosts++
            }
        }
        catch {
            Write-Log "Host $($SH.Name) is orphaned (no VM found). Marking for Drain." "WARN"
            $OutdatedHosts += $SH.Name
        }
    }

    $SummaryMsg = "Current State: {0} Healthy | {1} Outdated" -f $UpToDateHosts, $OutdatedHosts.Count
    if ($OutdatedHosts.Count -gt 0) { Write-Log $SummaryMsg "WARN" } else { Write-Log $SummaryMsg "SUCCESS" }

    # -------------------------------------------------------------------------
    # STEP 2b: DEPLOYMENT HEALTH CHECK (Failed & Long-Running)
    # -------------------------------------------------------------------------
    $RunningDeploymentVMNames = @()
    if ($ComputeResourceGroupName) {
        Write-Log "Checking deployment health in '$ComputeResourceGroupName'..."
        $Deployments = Get-AzResourceGroupDeployment -ResourceGroupName $ComputeResourceGroupName -ErrorAction SilentlyContinue |
                       Where-Object { $_.DeploymentName -like 'deploy-*' }

        $FailedDeployments = @($Deployments | Where-Object { $_.ProvisioningState -eq 'Failed' })
        if ($FailedDeployments.Count -gt 0) {
            throw "BLOCKED: Found $($FailedDeployments.Count) failed deployment(s) in '$ComputeResourceGroupName'. Clean up before proceeding."
        }

        $RunningDeployments = @($Deployments | Where-Object { $_.ProvisioningState -eq 'Running' })
        if ($RunningDeployments.Count -gt 0) {
            Write-Log "Found $($RunningDeployments.Count) running deployment(s)." "WARN"
            $LongRunning = @($RunningDeployments | Where-Object { $_.Timestamp -lt (Get-Date).AddHours(-2) })
            if ($LongRunning.Count -gt 0) {
                Write-Log "WARNING: $($LongRunning.Count) deployment(s) running for >2 hours. May block future deployments." "WARN"
            }
        }

        # Collect VM names from running deployments for collision avoidance
        foreach ($RD in $RunningDeployments) {
            try {
                $RDDetail = Get-AzResourceGroupDeployment -ResourceGroupName $ComputeResourceGroupName -Name $RD.DeploymentName -ErrorAction SilentlyContinue
                if ($RDDetail.Parameters.vmList) {
                    $RunningDeploymentVMNames += $RDDetail.Parameters.vmList.Value
                }
            } catch { }
        }
        Write-Log "Deployment health check passed." "SUCCESS"
    }

    # -------------------------------------------------------------------------
    # STEP 3: CALCULATE DEPLOYMENT TARGETS
    # -------------------------------------------------------------------------
    Write-Log "Calculating deployment targets..."

    # CIRCUIT BREAKER: If 100% of hosts are outdated in a multi-host pool, require explicit override
    if ($OutdatedHosts.Count -eq $CurrentCount -and $CurrentCount -gt 1 -and $VmCountOverride -eq 0) {
        Write-Log "CIRCUIT BREAKER: ALL $CurrentCount hosts are outdated. Full fleet replacement detected." "WARN"
        Write-Log "Pass -VmCountOverride $CurrentCount to confirm, or verify your image tags." "WARN"
        throw "Aborting: 100% of session hosts marked outdated. This requires explicit confirmation via -VmCountOverride."
    }

    # Determine final VM count based on overrides or existing count
    $TargetCount = if ($VmCountOverride -gt 0) { $VmCountOverride } elseif ($CurrentCount -eq 0) { 1 } else { $CurrentCount }

    # Generate names with collision avoidance (ISO-8601 Week-based)
    $WeekNum = (Get-Culture).Calendar.GetWeekOfYear((Get-Date), [System.Globalization.CalendarWeekRule]::FirstFourDayWeek, [DayOfWeek]::Monday)
    $WeekStr = "{0:D2}" -f $WeekNum
    $Prefix  = "$BaseName" + "$WeekStr"
    
    # Collect all existing VM names for collision avoidance
    $ExistingNames = @()
    foreach ($SH in $CurrentHosts) {
        $ShortName = ($SH.Name -split '/')[-1] -replace '\..*$', ''
        $ExistingNames += $ShortName
    }
    if ($RunningDeploymentVMNames) { $ExistingNames += $RunningDeploymentVMNames }

    # Generate unique names that avoid collisions
    $NewNames = @()
    $vmNumber = 1
    for ($i = 0; $i -lt $TargetCount; $i++) {
        while (("$Prefix{0:D2}" -f $vmNumber) -in $ExistingNames) { $vmNumber++ }
        # Validate computerName length (15 char Windows limit)
        $CandidateName = "$Prefix{0:D2}" -f $vmNumber
        if ($CandidateName.Length -gt 15) {
            throw "Generated hostname '$CandidateName' exceeds 15-character Windows limit. Shorten -BaseName."
        }
        $NewNames += $CandidateName
        $ExistingNames += $CandidateName
        $vmNumber++
    }

    # Split targets into Canary (1st host) and Blast (remaining hosts)
    $Canary = @($NewNames[0])
    $Blast  = if ($NewNames.Count -gt 1) { $NewNames | Select-Object -Skip 1 } else { @() }

    # Explicitly log the deployment plan in the DevOps console
    Write-Log "Planned Canary: $($Canary -join ', ')" "PLAN"
    if ($Blast) { Write-Log "Planned Blast : $($Blast -join ', ')" "PLAN" }

    # -------------------------------------------------------------------------
    # STEP 4: REGISTRATION TOKEN
    # -------------------------------------------------------------------------
    Write-Log "Ensuring valid Host Pool Registration Token..."
    $Now = (Get-Date).ToUniversalTime()
    $RegInfo = Get-AzWvdRegistrationInfo -ResourceGroupName $ResourceGroupName -HostPoolName $HostPoolName -ErrorAction SilentlyContinue
    
    $Token = $null
    # Reuse token only if it is valid for at least another 2 hours
    if ($RegInfo -and $RegInfo.Token -and [DateTime]$RegInfo.ExpirationTime -gt $Now.AddHours(2)) {
        Write-Log "Existing token is valid. Reusing." "SUCCESS"
        $Token = $RegInfo.Token
    } else {
        Write-Log "Generating new 24h registration token..." "INFO"
        $Expiry = $Now.AddHours(24).ToString('yyyy-MM-ddTHH:mm:ss.fffffffZ')
        $Token = (New-AzWvdRegistrationInfo -ResourceGroupName $ResourceGroupName -HostPoolName $HostPoolName -ExpirationTime $Expiry).Token
    }

    # -------------------------------------------------------------------------
    # STEP 5: OUTPUT TO CONSOLE & PIPELINE
    # -------------------------------------------------------------------------
    Write-Log "Finalizing Pipeline Variables..." "INFO"

    $CanaryJson   = ConvertTo-Json -InputObject @($Canary) -Compress
    $BlastJson    = ConvertTo-Json -InputObject @($Blast) -Compress
    $OutdatedJson = ConvertTo-Json -InputObject @($OutdatedHosts) -Compress

    $guid = (New-Guid).ToString()

    # --- VERBOSE CONSOLE OUTPUT (BEAUTIFIED) ---
    Write-Host "`n--- PIPELINE VARIABLE SUMMARY ---" -ForegroundColor Magenta
    Write-Log "TargetImageVersion  : $TargetVersion" "PLAN"
    Write-Log "TargetImageId       : $TargetImageId" "PLAN"
    Write-Log "CanaryList          : $CanaryJson" "PLAN"
    Write-Log "BlastList           : $BlastJson" "PLAN"
    Write-Log "OutdatedHostsList   : $OutdatedJson" "PLAN"
    Write-Log "HostPoolToken       : [REDACTED (SECRET)]" "PLAN"
    Write-Log "DeploymentGuid      : $guid" "PLAN"
    Write-Host "----------------------------------`n" -ForegroundColor Magenta

    # --- AZURE DEVOPS AGENT COMMANDS ---
    Write-Host "##vso[task.setvariable variable=TargetImageVersion;isOutput=true]$TargetVersion"
    Write-Host "##vso[task.setvariable variable=TargetImageId;isOutput=true]$TargetImageId"
    Write-Host "##vso[task.setvariable variable=CanaryList;isOutput=true]$CanaryJson"
    Write-Host "##vso[task.setvariable variable=BlastList;isOutput=true]$BlastJson"
    Write-Host "##vso[task.setvariable variable=OutdatedHostsList;isOutput=true]$OutdatedJson"
    Write-Host "##vso[task.setvariable variable=HostPoolToken;isOutput=true;isSecret=true]$Token"
    Write-Host "##vso[task.setvariable variable=DeploymentGuid;isOutput=true]$guid"
    
    Write-Log "Pre-Flight Complete. Readiness: GREEN." "SUCCESS"
}
catch {
    Write-Log "FATAL ERROR: $($_.Exception.Message)" "ERROR"
    exit 1
}