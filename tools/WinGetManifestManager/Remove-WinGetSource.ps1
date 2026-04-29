<#
.SYNOPSIS
    Complete teardown of a WinGet REST source deployment.

.DESCRIPTION
    Removes all Azure resources, Entra ID app registrations, soft-deleted resources,
    and local winget source registrations created by Deploy-WinGetSource.ps1.

    Handles these soft-delete scenarios:
      - API Management (48hr soft-delete retention)
      - Key Vault (soft-delete enabled by default, 7-90 day retention)
      - App Service / Function App (30-day soft-delete)
      - Entra ID app registrations (30-day soft-delete in recycle bin)

.PARAMETER Name
    Base name used during deployment (e.g. "corpwinget").

.PARAMETER ResourceGroup
    Resource group to delete. Default: "rg-winget-prod-001".

.PARAMETER SubscriptionName
    Azure subscription name. If not provided, uses current context.

.PARAMETER SourceDisplayName
    Local winget source display name. Default: "CorpWinGet".

.PARAMETER SkipSourceRemoval
    Skip removing the local winget source registration.

.PARAMETER SkipEntraCleanup
    Skip deleting Entra ID app registrations.

.PARAMETER SkipSoftDeletePurge
    Skip purging soft-deleted resources (APIM, KeyVault, App Service).

.PARAMETER PurgeOnly
    Only purge soft-deleted resources — don't delete the RG or Entra apps.
    Useful when a previous teardown left soft-deleted resources behind.

.EXAMPLE
    # Full teardown
    .\Remove-WinGetSource.ps1 -Name "corpwinget"

.EXAMPLE
    # Purge soft-deleted resources only (after a failed or partial teardown)
    .\Remove-WinGetSource.ps1 -Name "corpwinget" -PurgeOnly

.EXAMPLE
    # Full teardown with custom RG and no source removal
    .\Remove-WinGetSource.ps1 -Name "corpwinget" -ResourceGroup "rg-custom" -SkipSourceRemoval
#>

#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Az.Accounts'; ModuleVersion = '2.7.5' }

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory, Position = 0, HelpMessage = "Base name used during deployment")]
    [string]$Name,

    [string]$ResourceGroup = "rg-winget-prod-001",

    [string]$SubscriptionName,

    [string]$SourceDisplayName = "CorpWinGet",

    [switch]$SkipSourceRemoval,

    [switch]$SkipEntraCleanup,

    [switch]$SkipSoftDeletePurge,

    [switch]$PurgeOnly
)

$ErrorActionPreference = 'Stop'

# ═══════════════════════════════════════════════════════════════════════════════
#  HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

function Write-Step {
    param([string]$Message, [ValidateSet('INFO','SUCCESS','WARN','ERROR','SECTION')][string]$Level = 'INFO')
    $colors = @{ INFO = 'Cyan'; SUCCESS = 'Green'; WARN = 'Yellow'; ERROR = 'Red'; SECTION = 'Magenta' }
    $icons  = @{ INFO = '  '; SUCCESS = '✓ '; WARN = '⚠ '; ERROR = '✗ '; SECTION = '' }
    if ($Level -eq 'SECTION') {
        Write-Host ""
        Write-Host ("═" * 70) -ForegroundColor $colors[$Level]
        Write-Host "  $Message" -ForegroundColor $colors[$Level]
        Write-Host ("═" * 70) -ForegroundColor $colors[$Level]
    } else {
        Write-Host "$($icons[$Level])$Message" -ForegroundColor $colors[$Level]
    }
}

function Get-PlainToken {
    param([string]$ResourceUrl = 'https://management.azure.com')
    $tokenObj = Get-AzAccessToken -ResourceUrl $ResourceUrl -ErrorAction Stop
    if ($tokenObj.Token -is [securestring]) {
        return $tokenObj.Token | ConvertFrom-SecureString -AsPlainText
    }
    return $tokenObj.Token
}

function Get-ArmHeaders {
    $token = Get-PlainToken -ResourceUrl 'https://management.azure.com'
    return @{ Authorization = "Bearer $token"; 'Content-Type' = 'application/json' }
}

function Get-GraphHeaders {
    $token = Get-PlainToken -ResourceUrl 'https://graph.microsoft.com'
    return @{ Authorization = "Bearer $token"; 'Content-Type' = 'application/json' }
}

# ═══════════════════════════════════════════════════════════════════════════════
#  CONNECT TO AZURE
# ═══════════════════════════════════════════════════════════════════════════════

Write-Step "WINGET REST SOURCE TEARDOWN" -Level SECTION

$ctx = Get-AzContext
if (-not $ctx) {
    Write-Step "Not connected to Azure. Running Connect-AzAccount..." -Level WARN
    Connect-AzAccount
    $ctx = Get-AzContext
}

if ($SubscriptionName) {
    Set-AzContext -Subscription $SubscriptionName | Out-Null
    $ctx = Get-AzContext
}

$subId = $ctx.Subscription.Id
Write-Step "Subscription: $($ctx.Subscription.Name) ($subId)" -Level INFO
Write-Step "Name: $Name | RG: $ResourceGroup" -Level INFO

# Derived resource names (matching module's naming convention)
$funcAppName  = "func-$($Name -replace '[^a-zA-Z0-9-]', '')"
$apimName     = "apim-$($Name -replace '[^a-zA-Z0-9-]', '')"
$kvName       = "kv-$($Name -replace '[^a-zA-Z0-9-]', '')"

Write-Step "Expected resources: $funcAppName, $apimName, $kvName" -Level INFO

$summary = [ordered]@{
    SourceRemoved      = 'Skipped'
    EntraAppsDeleted   = 0
    ResourceGroup      = 'Skipped'
    SoftDeletedPurged  = @()
}

# ═══════════════════════════════════════════════════════════════════════════════
#  STEP 1: REMOVE LOCAL WINGET SOURCE
# ═══════════════════════════════════════════════════════════════════════════════

if (-not $PurgeOnly -and -not $SkipSourceRemoval) {
    Write-Step "REMOVE LOCAL WINGET SOURCE" -Level SECTION

    Write-Step "Checking for winget CLI..." -Level INFO
    $wingetCmd = Get-Command winget -ErrorAction SilentlyContinue
    if (-not $wingetCmd) {
        Write-Step "winget CLI not found — skipping source removal." -Level WARN
    } else {
        Write-Step "winget found at: $($wingetCmd.Source)" -Level INFO
        Write-Step "Checking if source '$SourceDisplayName' is registered..." -Level INFO
        $existingCheck = winget source list --name $SourceDisplayName 2>&1
        Write-Step "winget source list exit code: $LASTEXITCODE" -Level INFO
        if ($LASTEXITCODE -eq 0) {
            Write-Step "Source '$SourceDisplayName' is registered — proceeding with removal" -Level INFO
            if ($PSCmdlet.ShouldProcess($SourceDisplayName, "Remove winget source")) {
                $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
                    [Security.Principal.WindowsBuiltInRole]::Administrator)
                Write-Step "Running as admin: $isAdmin" -Level INFO
                if ($isAdmin) {
                    Write-Step "Executing: winget source remove --name $SourceDisplayName" -Level INFO
                    winget source remove --name $SourceDisplayName 2>&1 | Out-Null
                } else {
                    Write-Step "Elevating to admin for winget source remove..." -Level INFO
                    $tmpFile = [System.IO.Path]::GetTempFileName()
                    Start-Process pwsh -ArgumentList '-NoProfile', '-Command',
                        "winget source remove --name '$SourceDisplayName' 2>&1 | Out-File '$tmpFile'" `
                        -Verb RunAs -Wait -ErrorAction Stop
                    $result = Get-Content $tmpFile -ErrorAction SilentlyContinue
                    Remove-Item $tmpFile -ErrorAction SilentlyContinue
                    Write-Step "Result: $result" -Level INFO
                }
                $summary.SourceRemoved = 'Removed'
                Write-Step "Source '$SourceDisplayName' removed." -Level SUCCESS
            } else {
                Write-Step "ShouldProcess declined — skipping source removal" -Level INFO
            }
        } else {
            Write-Step "Source '$SourceDisplayName' not registered locally — nothing to remove." -Level INFO
            $summary.SourceRemoved = 'Not found'
        }
    }
} else {
    Write-Step "Source removal skipped (PurgeOnly=$PurgeOnly, SkipSourceRemoval=$SkipSourceRemoval)" -Level INFO
}

# ═══════════════════════════════════════════════════════════════════════════════
#  STEP 2: DELETE ENTRA ID APP REGISTRATIONS
# ═══════════════════════════════════════════════════════════════════════════════

if (-not $PurgeOnly -and -not $SkipEntraCleanup) {
    Write-Step "DELETE ENTRA ID APP REGISTRATIONS" -Level SECTION

    try {
        Write-Step "Acquiring Graph token..." -Level INFO
        $gh = Get-GraphHeaders
        Write-Step "Graph token acquired — querying apps matching '$Name'..." -Level INFO
        $graphApps = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/applications?`$filter=startswith(displayName,'$Name')" -Headers $gh
        Write-Step "Found $($graphApps.value.Count) Entra ID app(s) matching '$Name'." -Level INFO

        foreach ($app in $graphApps.value) {
            Write-Step "  $($app.displayName) | appId=$($app.appId) | id=$($app.id)" -Level INFO
            if ($PSCmdlet.ShouldProcess($app.displayName, "Delete Entra ID app")) {
                Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/applications/$($app.id)" -Headers $gh -Method Delete
                $summary.EntraAppsDeleted++
                Write-Step "  Deleted: $($app.displayName)" -Level SUCCESS
            }
        }

        # Also clean up service principals
        $sps = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=startswith(displayName,'$Name')" -Headers $gh
        foreach ($sp in $sps.value) {
            # Skip Microsoft-owned service principals
            if ($sp.appOwnerOrganizationId -and $sp.appOwnerOrganizationId -ne $ctx.Tenant.Id) { continue }
            Write-Step "  SP: $($sp.displayName) | id=$($sp.id)" -Level INFO
            if ($PSCmdlet.ShouldProcess($sp.displayName, "Delete Service Principal")) {
                try {
                    Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.id)" -Headers $gh -Method Delete
                    Write-Step "  Deleted SP: $($sp.displayName)" -Level SUCCESS
                } catch {
                    Write-Step "  Could not delete SP (may have been auto-removed): $($_.Exception.Message)" -Level WARN
                }
            }
        }
    } catch {
        Write-Step "Error querying Entra ID: $($_.Exception.Message)" -Level ERROR
    }
} else {
    Write-Step "Entra cleanup skipped (PurgeOnly=$PurgeOnly, SkipEntraCleanup=$SkipEntraCleanup)" -Level INFO
}

# ═══════════════════════════════════════════════════════════════════════════════
#  STEP 3: DELETE RESOURCE GROUP
# ═══════════════════════════════════════════════════════════════════════════════

if (-not $PurgeOnly) {
    Write-Step "DELETE RESOURCE GROUP" -Level SECTION

    Write-Step "Checking if resource group '$ResourceGroup' exists..." -Level INFO
    $rg = Get-AzResourceGroup -Name $ResourceGroup -ErrorAction SilentlyContinue
    if ($rg) {
        # List what's in it
        Write-Step "Resource group found — enumerating resources..." -Level INFO
        $resources = Get-AzResource -ResourceGroupName $ResourceGroup -ErrorAction SilentlyContinue
        Write-Step "Resource group '$ResourceGroup' contains $($resources.Count) resource(s):" -Level INFO
        $resources | ForEach-Object { Write-Step "  $($_.ResourceType.Split('/')[-1]): $($_.Name)" -Level INFO }

        if ($PSCmdlet.ShouldProcess($ResourceGroup, "Delete Resource Group (and all contents)")) {
            Write-Step "Deleting resource group '$ResourceGroup'... (this can take 5-15 minutes)" -Level INFO

            $deleteJob = Remove-AzResourceGroup -Name $ResourceGroup -Force -AsJob
            $spinner = @('⠋','⠙','⠹','⠸','⠼','⠴','⠦','⠧','⠇','⠏')
            $i = 0
            while ($deleteJob.State -eq 'Running') {
                $elapsed = ((Get-Date) - $deleteJob.PSBeginTime).ToString('mm\:ss')
                Write-Host "`r  $($spinner[$i % $spinner.Count]) Deleting... ($elapsed)" -NoNewline
                Start-Sleep -Seconds 5
                $i++
            }
            Write-Host ""

            if ($deleteJob.State -eq 'Completed') {
                Write-Step "Resource group '$ResourceGroup' deleted." -Level SUCCESS
                $summary.ResourceGroup = 'Deleted'
            } else {
                Write-Step "Resource group deletion failed:" -Level ERROR
                Receive-Job $deleteJob | ForEach-Object { Write-Step "  $_" -Level ERROR }
                $summary.ResourceGroup = 'Failed'
            }
            Remove-Job $deleteJob -Force -ErrorAction SilentlyContinue
        }
    } else {
        Write-Step "Resource group '$ResourceGroup' does not exist — nothing to delete." -Level INFO
        $summary.ResourceGroup = 'Not found'
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
#  STEP 4: PURGE SOFT-DELETED RESOURCES
# ═══════════════════════════════════════════════════════════════════════════════

if (-not $SkipSoftDeletePurge) {
    Write-Step "PURGE SOFT-DELETED RESOURCES" -Level SECTION

    Write-Step "Acquiring ARM token..." -Level INFO
    $h = Get-ArmHeaders
    Write-Step "ARM token acquired" -Level INFO

    # 4a: Soft-deleted API Management instances
    Write-Step "Checking soft-deleted API Management instances..." -Level INFO
    try {
        $apimDeleted = Invoke-RestMethod -Uri "https://management.azure.com/subscriptions/$subId/providers/Microsoft.ApiManagement/deletedservices?api-version=2022-08-01" -Headers $h
        $apimMatches = $apimDeleted.value | Where-Object { $_.properties.serviceName -like "*$Name*" }
        foreach ($apim in $apimMatches) {
            $svcName = $apim.properties.serviceName
            $location = $apim.location
            Write-Step "  Found: $svcName (deleted: $($apim.properties.deletionDate))" -Level WARN
            if ($PSCmdlet.ShouldProcess($svcName, "Purge soft-deleted APIM")) {
                $purgeUri = "https://management.azure.com/subscriptions/$subId/providers/Microsoft.ApiManagement/locations/$location/deletedservices/$svcName`?api-version=2022-08-01"
                Invoke-RestMethod -Uri $purgeUri -Headers $h -Method Delete | Out-Null
                $summary.SoftDeletedPurged += "APIM: $svcName"
                Write-Step "  Purged: $svcName" -Level SUCCESS
            }
        }
        if (-not $apimMatches) { Write-Step "  No soft-deleted APIM instances found." -Level INFO }
    } catch {
        Write-Step "  Error checking APIM: $($_.Exception.Message)" -Level WARN
    }

    # 4b: Soft-deleted Key Vaults
    Write-Step "Checking soft-deleted Key Vaults..." -Level INFO
    try {
        $kvDeleted = Get-AzKeyVault -InRemovedState -ErrorAction SilentlyContinue
        $kvMatches = $kvDeleted | Where-Object { $_.VaultName -like "*$Name*" }
        foreach ($kv in $kvMatches) {
            Write-Step "  Found: $($kv.VaultName) (deleted: $($kv.DeletionDate))" -Level WARN
            if ($PSCmdlet.ShouldProcess($kv.VaultName, "Purge soft-deleted Key Vault")) {
                try {
                    Remove-AzKeyVault -VaultName $kv.VaultName -Location $kv.Location -InRemovedState -Force -ErrorAction Stop
                    $summary.SoftDeletedPurged += "KeyVault: $($kv.VaultName)"
                    Write-Step "  Purged: $($kv.VaultName)" -Level SUCCESS
                } catch {
                    # Purge may be blocked by Azure Policy enforcing purge protection
                    Write-Step "  Cannot purge '$($kv.VaultName)': $($_.Exception.Message)" -Level WARN
                    Write-Step "  Purge protection may be enforced by Azure Policy. The vault will remain in soft-deleted state" -Level WARN
                    Write-Step "  and will be auto-recovered by Deploy-WinGetSource.ps1 on the next deployment." -Level WARN
                }
            }
        }
        if (-not $kvMatches) { Write-Step "  No soft-deleted Key Vaults found." -Level INFO }
    } catch {
        Write-Step "  Error checking Key Vaults: $($_.Exception.Message)" -Level WARN
    }

    # 4c: Soft-deleted Web Apps / Function Apps
    Write-Step "Checking soft-deleted Web/Function Apps..." -Level INFO
    try {
        $webDeleted = Invoke-RestMethod -Uri "https://management.azure.com/subscriptions/$subId/providers/Microsoft.Web/deletedSites?api-version=2023-01-01" -Headers $h
        $webMatches = $webDeleted.value | Where-Object { $_.properties.deletedSiteName -like "*$Name*" }
        foreach ($site in $webMatches) {
            $siteName = $site.properties.deletedSiteName
            $siteId   = $site.properties.deletedSiteId
            $slot     = $site.properties.slot
            Write-Step "  Found: $siteName (slot: $slot, deleted: $($site.properties.deletedTimestamp))" -Level WARN
            if ($PSCmdlet.ShouldProcess($siteName, "Purge soft-deleted Web/Function App")) {
                # The deletedSiteId from the list is used to permanently delete
                $deleteUri = "https://management.azure.com/subscriptions/$subId/providers/Microsoft.Web/deletedSites/$siteId`?api-version=2023-01-01"
                try {
                    Invoke-RestMethod -Uri $deleteUri -Headers $h -Method Delete | Out-Null
                    $summary.SoftDeletedPurged += "WebApp: $siteName"
                    Write-Step "  Purged: $siteName" -Level SUCCESS
                } catch {
                    # Some subscriptions don't support explicit purge — it auto-purges after 30 days
                    Write-Step "  Could not force-purge (may auto-purge): $($_.Exception.Message)" -Level WARN
                    $summary.SoftDeletedPurged += "WebApp: $siteName (auto-purge pending)"
                }
            }
        }
        if (-not $webMatches) { Write-Step "  No soft-deleted Web/Function Apps found." -Level INFO }
    } catch {
        Write-Step "  Error checking Web Apps: $($_.Exception.Message)" -Level WARN
    }

    # 4d: Soft-deleted App Configuration stores
    Write-Step "Checking soft-deleted App Configuration stores..." -Level INFO
    try {
        $appConfigDeleted = Invoke-RestMethod -Uri "https://management.azure.com/subscriptions/$subId/providers/Microsoft.AppConfiguration/deletedConfigurationStores?api-version=2023-03-01" -Headers $h
        $appConfigMatches = $appConfigDeleted.value | Where-Object { $_.name -like "*$Name*" }
        foreach ($ac in $appConfigMatches) {
            $acName = $ac.name
            $acLoc  = $ac.properties.location
            Write-Step "  Found: $acName (deleted: $($ac.properties.deletionDate))" -Level WARN
            if ($PSCmdlet.ShouldProcess($acName, "Purge soft-deleted App Configuration")) {
                $purgeUri = "https://management.azure.com/subscriptions/$subId/providers/Microsoft.AppConfiguration/locations/$acLoc/deletedConfigurationStores/$acName/purge?api-version=2023-03-01"
                Invoke-RestMethod -Uri $purgeUri -Headers $h -Method Post | Out-Null
                $summary.SoftDeletedPurged += "AppConfig: $acName"
                Write-Step "  Purged: $acName" -Level SUCCESS
            }
        }
        if (-not $appConfigMatches) { Write-Step "  No soft-deleted App Configuration stores found." -Level INFO }
    } catch {
        Write-Step "  Error checking App Configuration: $($_.Exception.Message)" -Level WARN
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
#  STEP 5: CLEAN UP LOCAL ARTIFACTS
# ═══════════════════════════════════════════════════════════════════════════════

if (-not $PurgeOnly) {
    Write-Step "CLEAN UP LOCAL ARTIFACTS" -Level SECTION

    # Remove generated parameter files
    $paramDir = Join-Path (Get-Location).Path 'Parameters'
    if (Test-Path $paramDir) {
        if ($PSCmdlet.ShouldProcess($paramDir, "Remove generated parameter files")) {
            Remove-Item $paramDir -Recurse -Force
            Write-Step "Removed: $paramDir" -Level SUCCESS
        }
    } else {
        Write-Step "No Parameters folder found." -Level INFO
    }

    # Remove temp files
    @(
        "$env:USERPROFILE\winget-source-result.txt",
        "$env:USERPROFILE\winget-remove-result.txt"
    ) | ForEach-Object {
        if (Test-Path $_) {
            Remove-Item $_ -Force -ErrorAction SilentlyContinue
            Write-Step "Removed temp file: $_" -Level INFO
        }
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
#  SUMMARY
# ═══════════════════════════════════════════════════════════════════════════════

Write-Step "TEARDOWN SUMMARY" -Level SECTION
Write-Step "Source registration : $($summary.SourceRemoved)" -Level INFO
Write-Step "Entra ID apps       : $($summary.EntraAppsDeleted) deleted" -Level INFO
Write-Step "Resource group      : $($summary.ResourceGroup)" -Level INFO
if ($summary.SoftDeletedPurged.Count -gt 0) {
    Write-Step "Soft-delete purged  : $($summary.SoftDeletedPurged.Count) resource(s)" -Level INFO
    $summary.SoftDeletedPurged | ForEach-Object { Write-Step "  - $_" -Level INFO }
} else {
    Write-Step "Soft-delete purged  : 0 resource(s)" -Level INFO
}

Write-Host ""
if ($summary.ResourceGroup -eq 'Failed') {
    Write-Step "Teardown completed with errors. Re-run or check Azure Portal." -Level WARN
    exit 1
} else {
    Write-Step "Teardown complete. Environment is clean for redeployment." -Level SUCCESS
    exit 0
}
