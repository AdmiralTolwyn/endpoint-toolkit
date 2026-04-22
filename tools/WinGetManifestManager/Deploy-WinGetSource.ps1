<#
.SYNOPSIS
    Deploys a private WinGet REST source to Azure and optionally seeds it with package manifests.

.DESCRIPTION
    Automates the end-to-end deployment of a Windows Package Manager REST source backed by
    Azure Functions + Cosmos DB. The script handles prerequisite validation, Azure authentication,
    resource deployment via the Microsoft.WinGet.RestSource module, manifest ingestion, and
    client source registration.

    All operations are logged to both the console (color-coded) and a timestamped log file.

.PARAMETER Name
    Base name for Azure resources (alphanumeric, 3-24 chars). Used as the Azure Function name.
    Example: "corpwinget"

.PARAMETER ResourceGroup
    Azure Resource Group to deploy into. Created if it does not exist.
    Default: "rg-winget-prod-001"

.PARAMETER Region
    Azure region for all resources.
    Default: "westeurope"

.PARAMETER SubscriptionId
    Azure Subscription ID. If omitted, uses the current Az context.

.PARAMETER PerformanceTier
    Azure resource sizing tier. Affects SKUs, redundancy, and cost.

    | Tier      | ASP  | Cosmos DB    | Storage | App Config | APIM      | Geo-Repl |
    |-----------|------|--------------|---------|------------|-----------|----------|
    | Developer | B1   | Free tier    | LRS     | Free       | Developer | No       |
    | Basic     | S1   | Pay-as-you-go| GRS     | Standard   | Basic     | Yes      |
    | Enhanced  | P1V2 | Pay-as-you-go| GRS     | Standard   | Standard  | Yes      |

    Note: Zone redundancy is always disabled. Only one Cosmos DB account per
    tenant can use the free tier (Developer).
    Default: "Developer"

.PARAMETER PublisherName
    Publisher name embedded in the REST source metadata.
    Default: Current user's display name or "WinGetRestSource"

.PARAMETER PublisherEmail
    Publisher email embedded in the REST source metadata.
    Default: Current user's Azure sign-in email

.PARAMETER Authentication
    REST source authentication mode. One of: None, MicrosoftEntraId.
    Default: "None"

.PARAMETER MicrosoftEntraIdResource
    The Application ID URI (IdentifierUri) for the Entra ID app registration.
    Required when Authentication is MicrosoftEntraId.
    Default: "api://<Name>"  (auto-generated from the Name parameter)

.PARAMETER MicrosoftEntraIdResourceScope
    The scope exposed by the Entra ID app registration.
    Default: "user_impersonation"

.PARAMETER ManifestPath
    Optional path to a folder containing WinGet YAML manifests to seed into the source
    after deployment. Supports nested folder structures.

.PARAMETER RegisterSource
    If specified, registers the deployed source on the local machine via 'winget source add'.

.PARAMETER SourceDisplayName
    The display name used when registering the source with the winget client.
    Default: "CorpWinGet"

.PARAMETER SkipAPIM
    When specified, removes the API Management instance created by New-WinGetSource
    after deployment completes. This saves ~$50-2,800/month depending on tier and
    simplifies the architecture by pointing clients directly at the Function App URL.
    The Function App works independently of APIM — all REST API functionality is
    preserved. Use IP access restrictions + Entra ID auth on the Function App for
    access control instead of APIM policies.

.PARAMETER CosmosDBZoneRedundant
    When specified, enables Availability Zone redundancy on Cosmos DB locations.
    By default the module deploys WITHOUT zone redundancy. Use this switch for
    production deployments in regions that support Availability Zones.
    Note: not all regions support zone-redundant Cosmos DB.

.PARAMETER CosmosDBRegion
    Deploy the Cosmos DB account to a different Azure region than the main resources.
    Useful when the primary region has Cosmos DB capacity constraints (e.g. West Europe
    currently rejects new accounts due to zonal redundancy capacity limits).
    If omitted, Cosmos DB deploys to the same region as all other resources.
    When specified, forces the split deployment workflow.
    Example: -CosmosDBRegion "northeurope"

.PARAMETER EnablePrivateEndpoints
    When specified, creates Azure Private Endpoints for the Function App, Cosmos DB,
    Key Vault, and Storage Account after deployment. This restricts network access to
    the VNet and disables public network access on each resource.
    Requires -VNetName and -SubnetName.

.PARAMETER VNetName
    Name of the existing Azure Virtual Network to attach private endpoints to.
    Required when -EnablePrivateEndpoints is specified.

.PARAMETER SubnetName
    Name of the subnet within the VNet to place private endpoint NICs.
    The subnet must have private endpoint network policies disabled.
    Default: "snet-privateendpoints"

.PARAMETER VNetResourceGroup
    Resource group containing the VNet, if different from the main resource group.
    Defaults to the -ResourceGroup value.

.PARAMETER RegisterPrivateDnsZones
    When specified alongside -EnablePrivateEndpoints, creates the required Azure
    Private DNS Zones and links them to the VNet:
      - privatelink.azurewebsites.net     (Function App)
      - privatelink.documents.azure.com   (Cosmos DB)
      - privatelink.vaultcore.azure.net   (Key Vault)
      - privatelink.blob.core.windows.net (Storage Account)
    Off by default — use this only if you need automatic DNS resolution within
    the VNet. If your environment already has DNS forwarding or centralized
    Private DNS Zones, skip this switch.

.PARAMETER RestSourcePath
    Path to the compiled Azure Function zip file (WinGet.RestSource.Functions.zip).
    If omitted, defaults to .\WinGet.RestSource.Functions.zip in the script directory.
    Use this to deploy a custom or patched build of the REST source Functions app.

.PARAMETER LogDirectory
    Directory for log files. Default: ".\logs"

.PARAMETER WhatIf
    Preview mode — validates everything but does not deploy.

.EXAMPLE
    # Lowest-cost deployment (Developer tier: free Cosmos DB, B1 ASP, LRS storage)
    .\Deploy-WinGetSource.ps1 -Name "corpwinget" -ResourceGroup "rg-winget-prod-001"

.EXAMPLE
    # Production deployment with zone-redundant Cosmos DB, Entra ID auth, and manifest seeding
    .\Deploy-WinGetSource.ps1 -Name "corpwinget" `
        -ResourceGroup "rg-winget-prod-001" `
        -Region "westeurope" `
        -PerformanceTier "Basic" `
        -CosmosDBZoneRedundant `
        -Authentication "MicrosoftEntraId" `
        -ManifestPath "C:\Manifests" `
        -RegisterSource

.EXAMPLE
    # Deploy with an existing Entra ID app registration
    .\Deploy-WinGetSource.ps1 -Name "corpwinget" `
        -ResourceGroup "rg-winget-prod-001" `
        -Authentication "MicrosoftEntraId" `
        -MicrosoftEntraIdResource "api://00000000-0000-0000-0000-000000000000" `
        -MicrosoftEntraIdResourceScope "user_impersonation"

.EXAMPLE
    # Deploy without APIM (direct Function App access, lower cost)
    .\Deploy-WinGetSource.ps1 -Name "corpwinget" `
        -ResourceGroup "rg-winget-prod-001" `
        -Authentication "MicrosoftEntraId" `
        -SkipAPIM

.EXAMPLE
    # Deploy Cosmos DB to a different region (e.g. when primary region has capacity issues)
    .\Deploy-WinGetSource.ps1 -Name "corpwinget" `
        -ResourceGroup "rg-winget-prod-001" `
        -Region "westeurope" `
        -CosmosDBRegion "northeurope" `
        -Authentication "MicrosoftEntraId" `
        -ManifestPath "C:\Manifests" `
        -RegisterSource

.EXAMPLE
    # Deploy with private endpoints (no auto DNS zones)
    .\Deploy-WinGetSource.ps1 -Name "corpwinget" `
        -ResourceGroup "rg-winget-prod-001" `
        -Authentication "MicrosoftEntraId" `
        -SkipAPIM `
        -EnablePrivateEndpoints `
        -VNetName "vnet-hub-westeurope-001" `
        -SubnetName "snet-privateendpoints" `
        -VNetResourceGroup "rg-network-prod-001"

.EXAMPLE
    # Deploy with private endpoints AND auto-created Private DNS Zones
    .\Deploy-WinGetSource.ps1 -Name "corpwinget" `
        -ResourceGroup "rg-winget-prod-001" `
        -Authentication "MicrosoftEntraId" `
        -SkipAPIM `
        -EnablePrivateEndpoints `
        -VNetName "vnet-hub-westeurope-001" `
        -SubnetName "snet-privateendpoints" `
        -RegisterPrivateDnsZones

.EXAMPLE
    # Dry-run to validate prerequisites
    .\Deploy-WinGetSource.ps1 -Name "corpwinget" -WhatIf
#>

#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Az.Accounts'; ModuleVersion = '2.7.5' }
#Requires -Modules Az.Network

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory, Position = 0, HelpMessage = "Base name for Azure resources (3-24 alphanumeric chars)")]
    [ValidatePattern('^[a-zA-Z][a-zA-Z0-9]{2,23}$')]
    [string]$Name,

    [ValidateNotNullOrEmpty()]
    [string]$ResourceGroup = "rg-winget-prod-001",

    [ValidateSet("eastus","eastus2","westus","westus2","westus3","centralus","northeurope","westeurope",
                 "uksouth","ukwest","southeastasia","eastasia","australiaeast","japaneast","canadacentral")]
    [string]$Region = "westeurope",

    [string]$SubscriptionId,

    [ValidateSet("Developer","Basic","Enhanced","BasicV2","StandardV2")]
    [string]$PerformanceTier = "Developer",

    [string]$PublisherName,

    [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$|^$')]
    [string]$PublisherEmail,

    [ValidateSet("None","MicrosoftEntraId")]
    [string]$Authentication = "None",

    [string]$MicrosoftEntraIdResource,

    [string]$MicrosoftEntraIdResourceScope = "user_impersonation",

    [ValidateScript({ if ($_ -and -not (Test-Path $_)) { throw "ManifestPath '$_' does not exist." }; $true })]
    [string]$ManifestPath,

    [switch]$RegisterSource,

    [switch]$SkipAPIM,

    [switch]$CosmosDBZoneRedundant,

    [ValidateSet("eastus","eastus2","westus","westus2","westus3","centralus","northeurope","westeurope",
                 "uksouth","ukwest","southeastasia","eastasia","australiaeast","japaneast","canadacentral")]
    [string]$CosmosDBRegion,

    [string]$SourceDisplayName = "CorpWinGet",

    [switch]$EnablePrivateEndpoints,

    [string]$VNetName,

    [string]$SubnetName = "snet-privateendpoints",

    [string]$VNetResourceGroup,

    [switch]$RegisterPrivateDnsZones,

    [string]$LogDirectory = ".\logs",

    [ValidateScript({ if ($_ -and -not (Test-Path $_)) { throw "RestSourcePath '$_' does not exist." }; $true })]
    [string]$RestSourcePath
)

# ═══════════════════════════════════════════════════════════════════════════════
#  GLOBALS
# ═══════════════════════════════════════════════════════════════════════════════

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'   # Suppress progress bars for speed

$Script:StartTime      = Get-Date
$Script:ExitCode       = 0
$Script:DeploymentUrl  = $null
$Script:Stats          = [ordered]@{
    PrereqChecks     = 'Pending'
    AzureAuth        = 'Pending'
    ResourceGroup    = 'Pending'
    ResourceAudit    = 'Pending'
    Deployment       = 'Pending'
    ManifestsLoaded  = 0
    ManifestsFailed  = 0
    SourceRegistered = 'Skipped'
    PrivateEndpoints = 'Skipped'
}
# Pre-flight inventory — populated by Step-ResourceAudit
$Script:ExistingResources = @{
    KeyVault       = $null   # $true if live in RG
    APIM           = $null
    CosmosDB       = $null
    FunctionApp    = $null
    StorageAccount = $null
    AppServicePlan = $null
    AppConfig      = $null
    AppInsights    = $null
}

# ═══════════════════════════════════════════════════════════════════════════════
#  LOGGING
# ═══════════════════════════════════════════════════════════════════════════════

# Ensure log directory exists
if (-not (Test-Path $LogDirectory)) {
    New-Item -ItemType Directory -Path $LogDirectory -Force -WhatIf:$false -Confirm:$false | Out-Null
}

$Script:LogFile = Join-Path $LogDirectory "Deploy-WinGetSource_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    <#
    .SYNOPSIS Writes a timestamped, color-coded log entry to console and file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [AllowEmptyString()]
        [string]$Message,

        [ValidateSet('INFO','WARN','ERROR','SUCCESS','DEBUG','SECTION')]
        [string]$Level = 'INFO',

        [switch]$NoNewline
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
    $prefix    = "[$timestamp] [$($Level.PadRight(7))]"
    $logLine   = "$prefix $Message"

    # Write to log file (always)
    Add-Content -Path $Script:LogFile -Value $logLine -Encoding UTF8 -WhatIf:$false -Confirm:$false

    # Console colors
    $colors = @{
        INFO    = 'Cyan'
        WARN    = 'Yellow'
        ERROR   = 'Red'
        SUCCESS = 'Green'
        DEBUG   = 'DarkGray'
        SECTION = 'Magenta'
    }

    $writeParams = @{
        ForegroundColor = $colors[$Level]
        NoNewline       = $NoNewline.IsPresent
    }

    if ($Level -eq 'SECTION') {
        Write-Host ""                                         @writeParams -NoNewline:$false
        Write-Host ("═" * 70)                                 @writeParams -NoNewline:$false
        Write-Host "  $Message"                               @writeParams -NoNewline:$false
        Write-Host ("═" * 70)                                 @writeParams -NoNewline:$false
    } elseif ($Level -eq 'DEBUG' -and $VerbosePreference -ne 'Continue') {
        # Debug messages only shown when -Verbose
        return
    } else {
        Write-Host $logLine @writeParams
    }
}

function Write-Banner {
    $banner = @"

    ╔══════════════════════════════════════════════════════════════╗
    ║          WinGet REST Source — Azure Deployment              ║
    ║          $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')                            ║
    ╚══════════════════════════════════════════════════════════════╝

"@
    Write-Host $banner -ForegroundColor Cyan
    Add-Content -Path $Script:LogFile -Value $banner -Encoding UTF8 -WhatIf:$false -Confirm:$false
}

function Write-Summary {
    $elapsed = (Get-Date) - $Script:StartTime
    $summary = @"

    ┌──────────────────────────────────────────────────────────────┐
    │                    DEPLOYMENT SUMMARY                        │
    ├──────────────────────────────────────────────────────────────┤
    │  Source Name      : $($Name.PadRight(38))│
    │  Resource Group   : $($ResourceGroup.PadRight(38))│
    │  Region           : $($Region.PadRight(38))│
    │  Performance      : $($PerformanceTier.PadRight(38))│
    │  CosmosDB Zone-HA : $($(if($CosmosDBZoneRedundant){'Enabled'}else{'Disabled'}).PadRight(38))│
    │  CosmosDB Region  : $($(if($CosmosDBRegion){$CosmosDBRegion}else{$Region}).PadRight(38))│
    │  Authentication   : $($Authentication.PadRight(38))│
    ├──────────────────────────────────────────────────────────────┤
    │  Prereq Checks    : $($Script:Stats.PrereqChecks.ToString().PadRight(38))│
    │  Azure Auth       : $($Script:Stats.AzureAuth.ToString().PadRight(38))│
    │  Resource Group   : $($Script:Stats.ResourceGroup.ToString().PadRight(38))│
    │  Resource Audit   : $($Script:Stats.ResourceAudit.ToString().PadRight(38))│
    │  Deployment       : $($Script:Stats.Deployment.ToString().PadRight(38))│
    │  Manifests OK     : $($Script:Stats.ManifestsLoaded.ToString().PadRight(38))│
    │  Manifests Failed : $($Script:Stats.ManifestsFailed.ToString().PadRight(38))│
    │  Source Registered: $($Script:Stats.SourceRegistered.ToString().PadRight(38))│
    ├──────────────────────────────────────────────────────────────┤
    │  REST Source URL  : $(if($Script:DeploymentUrl){$Script:DeploymentUrl.PadRight(38)}else{'N/A'.PadRight(38)})│
    │  Elapsed Time     : $("$($elapsed.Minutes)m $($elapsed.Seconds)s".PadRight(38))│
    │  Log File         : $(Split-Path $Script:LogFile -Leaf)  │
    │  Exit Code        : $($Script:ExitCode.ToString().PadRight(38))│
    └──────────────────────────────────────────────────────────────┘

"@
    Write-Host $summary -ForegroundColor $(if ($Script:ExitCode -eq 0) { 'Green' } else { 'Red' })
    Add-Content -Path $Script:LogFile -Value $summary -Encoding UTF8 -WhatIf:$false -Confirm:$false
}

# ═══════════════════════════════════════════════════════════════════════════════
#  HELPER FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════════

function Test-AdminElevation {
    $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]$identity
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-PlainToken {
    <#
    .SYNOPSIS Gets a plain-text Azure access token, handling SecureString in Az.Accounts 3+.
    #>
    [CmdletBinding()]
    param(
        [string]$ResourceUrl = 'https://management.azure.com'
    )
    $tokenObj = Get-AzAccessToken -ResourceUrl $ResourceUrl -ErrorAction Stop
    if ($tokenObj.Token -is [securestring]) {
        return $tokenObj.Token | ConvertFrom-SecureString -AsPlainText
    }
    return $tokenObj.Token
}

function Assert-ModuleAvailable {
    param([string]$ModuleName, [switch]$Install)

    $mod = Get-Module -ListAvailable -Name $ModuleName | Select-Object -First 1
    if (-not $mod) {
        if ($Install) {
            Write-Log "Module '$ModuleName' not found — installing..." -Level WARN
            try {
                Install-PSResource -Name $ModuleName -Scope CurrentUser -TrustRepository -ErrorAction Stop
                Write-Log "Module '$ModuleName' installed successfully." -Level SUCCESS
            } catch {
                # Fallback for older PS versions without Install-PSResource
                Install-Module -Name $ModuleName -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                Write-Log "Module '$ModuleName' installed (fallback Install-Module)." -Level SUCCESS
            }
        } else {
            throw "Required module '$ModuleName' is not installed. Run: Install-PSResource -Name $ModuleName"
        }
    } else {
        Write-Log "Module '$ModuleName' v$($mod.Version) found." -Level DEBUG
    }
}

function Watch-DeploymentProgress {
    <#
    .SYNOPSIS Polls ARM deployment and resource status while a blocking call runs.
    .DESCRIPTION
        Starts a background polling loop (ThreadJob) that prints live status
        every $IntervalSec seconds. Call Stop-Job on the returned job once the
        blocking operation finishes.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$StepName,
        [Parameter(Mandatory)][string]$RG,
        [int]$IntervalSec = 30
    )

    # Known resource type display-order (deploy sequence)
    $typeOrder = @(
        'Microsoft.OperationalInsights/workspaces'
        'Microsoft.Insights/components'
        'Microsoft.KeyVault/vaults'
        'Microsoft.AppConfiguration/configurationStores'
        'Microsoft.Storage/storageAccounts'
        'Microsoft.Web/serverFarms'
        'Microsoft.DocumentDB/databaseAccounts'
        'Microsoft.Web/sites'
        'Microsoft.ApiManagement/service'
    )

    # Pass subscription ID and access token so the ThreadJob can use REST API
    # instead of Az cmdlets (avoids .NET assembly conflicts in isolated runspaces).
    $pollSubId = (Get-AzContext).Subscription.Id
    $pollToken = Get-PlainToken

    $job = Start-ThreadJob -ArgumentList $StepName, $RG, $IntervalSec, $typeOrder, $Script:LogFile, $pollSubId, $pollToken -ScriptBlock {
        param($StepName, $RG, $IntervalSec, $typeOrder, $LogFile, $SubId, $Token)

        function Poll-Write {
            param([string]$Msg, [string]$Level = 'INFO')
            $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
            $prefix = "[$ts] [$($Level.PadRight(7))]"
            $color = switch ($Level) {
                'SUCCESS' { 'Green'  }
                'WARN'    { 'Yellow' }
                'ERROR'   { 'Red'    }
                default   { 'Gray'   }
            }
            Write-Host "$prefix $Msg" -ForegroundColor $color
            if ($LogFile) {
                Add-Content -Path $LogFile -Value "$prefix $Msg" -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }

        $sw       = [System.Diagnostics.Stopwatch]::StartNew()
        $lastHash = ''
        $headers  = @{ Authorization = "Bearer $Token"; 'Content-Type' = 'application/json' }

        while ($true) {
            Start-Sleep -Seconds $IntervalSec
            $elapsed = $sw.Elapsed

            try {
                # Poll resource provisioning states via ARM REST API
                # (avoids Az.Resources assembly conflicts in ThreadJob runspace)
                $resUri = "https://management.azure.com/subscriptions/$SubId/resourceGroups/$RG/resources?`$expand=provisioningState&api-version=2024-03-01"
                $resp = Invoke-RestMethod -Uri $resUri -Headers $headers -ErrorAction Stop

                $resources = $resp.value | ForEach-Object {
                    [PSCustomObject]@{
                        Name         = $_.name
                        ResourceType = $_.type
                        State        = if ($_.provisioningState) { $_.provisioningState }
                                       elseif ($_.properties.provisioningState) { $_.properties.provisioningState }
                                       else { '—' }
                    }
                }

                # Sort by deploy order
                $resources = $resources | Sort-Object {
                    $rt = $_.ResourceType
                    $idx = [array]::IndexOf($typeOrder, $rt)
                    if ($idx -lt 0) { 999 } else { $idx }
                }

                # Poll active ARM deployments via REST
                $depUri = "https://management.azure.com/subscriptions/$SubId/resourceGroups/$RG/providers/Microsoft.Resources/deployments?`$filter=provisioningState eq 'Running'&api-version=2024-03-01"
                $depResp = Invoke-RestMethod -Uri $depUri -Headers $headers -ErrorAction SilentlyContinue
                $running = $depResp.value

                # Build hash to detect changes
                $lines = @($resources | ForEach-Object { "$($_.Name)|$($_.State)" })
                $hash  = ($lines -join ';')

                $elapsed_str = "$([int]$elapsed.TotalMinutes)m$($elapsed.Seconds)s"

                if ($hash -ne $lastHash) {
                    Poll-Write "[$StepName] ${elapsed_str} — Resource status:" 'INFO'
                    foreach ($r in $resources) {
                        $icon = switch ($r.State) {
                            'Succeeded'  { '✓' }
                            'Creating'   { '⧖' }
                            'Activating' { '⧖' }
                            'Running'    { '⧖' }
                            'Failed'     { '✗' }
                            default      { '·' }
                        }
                        Poll-Write "  $icon $($r.Name.PadRight(42)) $($r.State)"
                    }
                    if ($running) {
                        $depNames = ($running | ForEach-Object { $_.name }) -join ', '
                        Poll-Write "  Active ARM deployments: $depNames" 'INFO'
                    }
                    $lastHash = $hash
                } else {
                    Poll-Write "[$StepName] ${elapsed_str} — still deploying..." 'INFO'
                }
            } catch {
                # Token may expire during long deploys — try to refresh
                if ($_.Exception.Message -match '401|Unauthorized|ExpiredToken') {
                    Poll-Write "[$StepName] ${elapsed_str} — token expired, poller stopping." 'WARN'
                    return
                }
                Poll-Write "[$StepName] ${elapsed_str} — poll error: $($_.Exception.Message)" 'WARN'
            }
        }
    }

    return $job
}

function Invoke-StepWithRetry {
    <#
    .SYNOPSIS Executes a script block with exponential backoff retry.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$StepName,

        [Parameter(Mandatory)]
        [scriptblock]$Action,

        [int]$MaxRetries = 3,

        [int]$BaseDelaySec = 10
    )

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            Write-Log "[$StepName] Attempt $attempt of $MaxRetries..." -Level DEBUG
            $result = & $Action
            # Some Azure cmdlets (e.g. New-WinGetSource) return $false on 
            # failure instead of throwing. Treat that as an error.
            if ($result -is [bool] -and $result -eq $false) {
                throw "$StepName returned `$false — operation did not succeed."
            }
            Write-Log "[$StepName] Completed successfully." -Level SUCCESS
            return $result
        } catch {
            Write-Log "[$StepName] Attempt $attempt failed: $($_.Exception.Message)" -Level WARN
            if ($attempt -eq $MaxRetries) {
                Write-Log "[$StepName] All $MaxRetries attempts exhausted." -Level ERROR
                throw
            }
            $delay = $BaseDelaySec * [math]::Pow(2, $attempt - 1)
            Write-Log "[$StepName] Retrying in ${delay}s..." -Level WARN
            Start-Sleep -Seconds $delay
        }
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
#  PHASE 1 — PREREQUISITE VALIDATION
# ═══════════════════════════════════════════════════════════════════════════════

function Step-ValidatePrerequisites {
    Write-Log "PHASE 1: PREREQUISITE VALIDATION" -Level SECTION

    # 1a. PowerShell version
    $psVer = $PSVersionTable.PSVersion
    Write-Log "PowerShell version: $psVer" -Level INFO
    if ($psVer.Major -lt 7 -or ($psVer.Major -eq 7 -and $psVer.Minor -lt 4)) {
        throw "PowerShell 7.4+ is required. Current: $psVer. Install from https://aka.ms/powershell"
    }
    Write-Log "PowerShell version requirement met." -Level SUCCESS

    # 1b. Admin elevation (required for winget source add)
    $isAdmin = Test-AdminElevation
    Write-Log "Running as administrator: $isAdmin" -Level $(if ($isAdmin) { 'SUCCESS' } else { 'WARN' })
    if (-not $isAdmin -and $RegisterSource) {
        Write-Log "WARNING: -RegisterSource requires elevation. Source registration will be attempted but may fail." -Level WARN
    }

    # 1c. Az modules — ensure compatible versions are loaded
    # Microsoft.WinGet.RestSource 1.10.0 calls New-AzADApplication (Graph-based
    # cmdlets) which requires Az.Resources ≥ 6.0. Older versions (e.g. 2.5.x)
    # use the legacy Azure AD Graph cmdlets where IdentifierUris is mandatory and
    # Update-AzADApplication lacks the -Api parameter. Az.Resources ≥ 6.0 in turn
    # requires Az.Accounts ≥ 2.7.5.
    #
    # The #Requires directive above auto-imports a compatible Az.Accounts in a
    # fresh session. If an older version was already loaded (e.g. from a prior
    # command or profile), .NET assembly conflicts prevent swapping — we detect
    # that here and fail fast with a clear message.
    Write-Log "Checking Az module versions (Az.Accounts ≥ 2.7.5, Az.Resources ≥ 6.0)..." -Level INFO

    $requiredModules = @(
        @{ Name = 'Az.Accounts'; MinVersion = [version]'2.7.5' }
        @{ Name = 'Az.Resources'; MinVersion = [version]'6.0.0' }
    )

    foreach ($req in $requiredModules) {
        $modName = $req.Name
        $minVer  = $req.MinVersion

        # Check if already loaded at a compatible version
        $loaded = Get-Module -Name $modName -ErrorAction SilentlyContinue
        if ($loaded -and $loaded.Version -ge $minVer) {
            Write-Log "$modName v$($loaded.Version) loaded (>= $minVer)." -Level SUCCESS
            continue
        }

        # If an incompatible version is already loaded, we cannot swap — .NET
        # assemblies are locked in the AppDomain for the lifetime of the process.
        if ($loaded) {
            $installed = Get-Module -ListAvailable -Name $modName |
                Where-Object { $_.Version -ge $minVer } |
                Sort-Object Version -Descending |
                Select-Object -First 1

            $suggestion = if ($installed) {
                "v$($installed.Version) is installed but cannot replace the already-loaded v$($loaded.Version) due to .NET assembly locking."
            } else {
                "No compatible version is installed."
            }

            throw @"
$modName v$($loaded.Version) is loaded but >= $minVer is required. $suggestion

To fix this, open a NEW PowerShell terminal and re-run the script.
The #Requires directive will auto-import the correct version in a fresh session.

If the problem persists, remove the old versions:
  Get-Module -ListAvailable $modName | Where-Object { `$_.Version -lt '$minVer' } | ForEach-Object { Remove-Item `$_.ModuleBase -Recurse -Force }
"@
        }

        # Not loaded yet — find and import the best available version
        $available = Get-Module -ListAvailable -Name $modName |
            Where-Object { $_.Version -ge $minVer } |
            Sort-Object Version -Descending |
            Select-Object -First 1

        if (-not $available) {
            throw "$modName >= $minVer is required but not installed. Installed versions: $(
                (Get-Module -ListAvailable -Name $modName | ForEach-Object { $_.Version }) -join ', '
            ). Run: Install-Module $modName -MinimumVersion $minVer -Scope CurrentUser -Force"
        }

        $psd1Path = Join-Path $available.ModuleBase "$modName.psd1"
        Import-Module $psd1Path -Force -ErrorAction Stop
        $nowLoaded = Get-Module -Name $modName
        Write-Log "$modName v$($nowLoaded.Version) loaded from $($nowLoaded.ModuleBase)" -Level SUCCESS
    }

    # 1d. WinGet REST source module
    Write-Log "Checking Microsoft.WinGet.RestSource module..." -Level INFO
    Assert-ModuleAvailable -ModuleName 'Microsoft.WinGet.RestSource' -Install

    # 1e. Import the module
    Write-Log "Importing Microsoft.WinGet.RestSource..." -Level INFO
    Import-Module Microsoft.WinGet.RestSource -Force -ErrorAction Stop
    Write-Log "Module imported successfully." -Level SUCCESS

    # 1f. Validate ManifestPath contents if provided
    if ($ManifestPath) {
        $yamlFiles = Get-ChildItem -Path $ManifestPath -Filter "*.yaml" -Recurse
        $yamlCount = ($yamlFiles | Measure-Object).Count
        if ($yamlCount -eq 0) {
            Write-Log "WARNING: ManifestPath '$ManifestPath' contains no .yaml files." -Level WARN
        } else {
            Write-Log "Found $yamlCount .yaml manifest files in '$ManifestPath'." -Level INFO
        }
    }

    # 1g. WinGet client (for source registration)
    if ($RegisterSource) {
        $wingetPath = Get-Command winget -ErrorAction SilentlyContinue
        if (-not $wingetPath) {
            Write-Log "WARNING: winget CLI not found on PATH. Source registration will be skipped." -Level WARN
        } else {
            $wingetVer = (winget --version 2>&1).Trim()
            Write-Log "winget CLI found: $wingetVer" -Level INFO
        }
    }

    $Script:Stats.PrereqChecks = 'Passed'
    Write-Log "All prerequisite checks passed." -Level SUCCESS
}

# ═══════════════════════════════════════════════════════════════════════════════
#  PHASE 2 — AZURE AUTHENTICATION
# ═══════════════════════════════════════════════════════════════════════════════

function Step-ConnectAzure {
    Write-Log "PHASE 2: AZURE AUTHENTICATION" -Level SECTION

    # Check existing context
    $ctx = Get-AzContext -ErrorAction SilentlyContinue
    if ($ctx -and $ctx.Account) {
        Write-Log "Already connected as: $($ctx.Account.Id)" -Level INFO
        Write-Log "Subscription: $($ctx.Subscription.Name) ($($ctx.Subscription.Id))" -Level INFO
    } else {
        Write-Log "No active Azure session. Initiating Connect-AzAccount..." -Level WARN
        Connect-AzAccount -ErrorAction Stop | Out-Null
        $ctx = Get-AzContext
        Write-Log "Authenticated as: $($ctx.Account.Id)" -Level SUCCESS
    }

    # Switch subscription if specified
    if ($SubscriptionId -and $ctx.Subscription.Id -ne $SubscriptionId) {
        Write-Log "Switching to subscription: $SubscriptionId" -Level INFO
        Set-AzContext -SubscriptionId $SubscriptionId -ErrorAction Stop | Out-Null
        $ctx = Get-AzContext
        Write-Log "Subscription set: $($ctx.Subscription.Name)" -Level SUCCESS
    }

    # Populate defaults for publisher info
    if (-not $PublisherName) {
        $Script:ResolvedPublisherName = $ctx.Account.Id.Split('@')[0]
        Write-Log "PublisherName defaulted to: $Script:ResolvedPublisherName" -Level DEBUG
    } else {
        $Script:ResolvedPublisherName = $PublisherName
    }

    if (-not $PublisherEmail) {
        $Script:ResolvedPublisherEmail = $ctx.Account.Id
        Write-Log "PublisherEmail defaulted to: $Script:ResolvedPublisherEmail" -Level DEBUG
    } else {
        $Script:ResolvedPublisherEmail = $PublisherEmail
    }

    $Script:Stats.AzureAuth = 'Connected'
    Write-Log "Azure authentication complete." -Level SUCCESS
}

# ═══════════════════════════════════════════════════════════════════════════════
#  PHASE 3 — RESOURCE GROUP
# ═══════════════════════════════════════════════════════════════════════════════

function Step-EnsureResourceGroup {
    Write-Log "PHASE 3: RESOURCE GROUP" -Level SECTION

    $rg = Get-AzResourceGroup -Name $ResourceGroup -ErrorAction SilentlyContinue
    if ($rg) {
        Write-Log "Resource Group '$ResourceGroup' already exists in '$($rg.Location)'." -Level INFO
        $Script:Stats.ResourceGroup = 'Exists'
    } else {
        if ($PSCmdlet.ShouldProcess($ResourceGroup, "Create Resource Group in $Region")) {
            Write-Log "Creating Resource Group '$ResourceGroup' in '$Region'..." -Level INFO
            New-AzResourceGroup -Name $ResourceGroup -Location $Region -ErrorAction Stop | Out-Null
            Write-Log "Resource Group '$ResourceGroup' created." -Level SUCCESS
            $Script:Stats.ResourceGroup = 'Created'
        } else {
            Write-Log "[WhatIf] Would create Resource Group '$ResourceGroup' in '$Region'." -Level WARN
            $Script:Stats.ResourceGroup = 'WhatIf'
        }
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
#  PHASE 3b — PRE-FLIGHT RESOURCE AUDIT
# ═══════════════════════════════════════════════════════════════════════════════

function Step-ResourceAudit {
    <#
    .SYNOPSIS Inventories existing resources and resolves soft-deletes before deployment.
    .DESCRIPTION
        1. Scans the target Resource Group for resources that already exist.
        2. Populates $Script:ExistingResources so the deploy step can reuse them.
        3. Checks for soft-deleted Key Vault / APIM and resolves:
           - Key Vault: always recovers (purge protection may be Azure-Policy mandated).
           - APIM: purges soft-deleted instances to free the name.
        This makes the deployment idempotent and safe for re-runs.
    #>
    Write-Log "PHASE 3b: PRE-FLIGHT RESOURCE AUDIT" -Level SECTION

    $subId   = (Get-AzContext).Subscription.Id
    $tok     = Get-PlainToken
    $headers = @{ Authorization = "Bearer $tok"; 'Content-Type' = 'application/json' }
    $cleanName = $Name -replace '[^a-zA-Z0-9-]', ''

    # ── 0. Resource name overrides (set via GUI environment variables) ──
    if ($env:WINGETMM_OVERRIDE_FUNCAPP)  { Write-Log "Resource override: Function App → $($env:WINGETMM_OVERRIDE_FUNCAPP)" -Level INFO }
    if ($env:WINGETMM_OVERRIDE_KEYVAULT) { Write-Log "Resource override: Key Vault → $($env:WINGETMM_OVERRIDE_KEYVAULT)" -Level INFO }
    if ($env:WINGETMM_OVERRIDE_STORAGE)  { Write-Log "Resource override: Storage → $($env:WINGETMM_OVERRIDE_STORAGE)" -Level INFO }
    if ($env:WINGETMM_OVERRIDE_APIM)     { Write-Log "Resource override: APIM → $($env:WINGETMM_OVERRIDE_APIM)" -Level INFO }
    if ($env:WINGETMM_OVERRIDE_ASP)      { Write-Log "Resource override: ASP → $($env:WINGETMM_OVERRIDE_ASP)" -Level INFO }

    # Parse deployment tags from environment variable
    $DeployTags = @{}
    if ($env:WINGETMM_DEPLOY_TAGS) {
        foreach ($Pair in ($env:WINGETMM_DEPLOY_TAGS -split ';')) {
            $KV = $Pair -split '=', 2
            if ($KV.Count -eq 2 -and $KV[0].Trim()) {
                $DeployTags[$KV[0].Trim()] = $KV[1].Trim()
            }
        }
        if ($DeployTags.Count -gt 0) {
            Write-Log "Deployment tags: $($DeployTags.Keys -join ', ')" -Level INFO
        }
    }

    # ── 1. INVENTORY EXISTING RESOURCES IN THE TARGET RG ──
    Write-Log "Scanning '$ResourceGroup' for existing resources..." -Level INFO
    $rgResources = Get-AzResource -ResourceGroupName $ResourceGroup -ErrorAction SilentlyContinue
    $existingCount = 0

    if ($rgResources) {
        # Map well-known resource types to our inventory
        $typeMap = @{
            'Microsoft.KeyVault/vaults'                              = 'KeyVault'
            'Microsoft.ApiManagement/service'                        = 'APIM'
            'Microsoft.DocumentDB/databaseAccounts'                  = 'CosmosDB'
            'Microsoft.Web/sites'                                    = 'FunctionApp'
            'Microsoft.Storage/storageAccounts'                      = 'StorageAccount'
            'Microsoft.Web/serverFarms'                              = 'AppServicePlan'
            'Microsoft.AppConfiguration/configurationStores'         = 'AppConfig'
            'Microsoft.Insights/components'                          = 'AppInsights'
        }

        foreach ($res in $rgResources) {
            $slot = $typeMap[$res.ResourceType]
            if ($slot) {
                $Script:ExistingResources[$slot] = $true
                $existingCount++
            }
        }

        # Print inventory table
        Write-Log "  ┌────────────────────┬──────────┬──────────────────────────────────┐" -Level INFO
        Write-Log "  │ Resource Type      │ Status   │ Name                             │" -Level INFO
        Write-Log "  ├────────────────────┼──────────┼──────────────────────────────────┤" -Level INFO
        $orderedTypes = @('KeyVault','APIM','CosmosDB','FunctionApp','StorageAccount','AppServicePlan','AppConfig','AppInsights')
        foreach ($t in $orderedTypes) {
            $match = $rgResources | Where-Object { $typeMap[$_.ResourceType] -eq $t } | Select-Object -First 1
            $status = if ($match) { 'EXISTS  ' } else { 'NEW     ' }
            $rName  = if ($match) { $match.Name } else { '-' }
            Write-Log ("  │ {0,-18} │ {1,-8} │ {2,-32} │" -f $t, $status, $rName.Substring(0, [Math]::Min(32, $rName.Length))) -Level INFO
        }
        Write-Log "  └────────────────────┴──────────┴──────────────────────────────────┘" -Level INFO

        if ($existingCount -gt 0) {
            Write-Log "$existingCount existing resource(s) found — deployment will update in-place (ARM Incremental mode)." -Level INFO
        }

        # ── 1b. Check APIM details (state, SKU) for reuse awareness ──
        if ($Script:ExistingResources.APIM) {
            $apimRes = $rgResources | Where-Object { $typeMap[$_.ResourceType] -eq 'APIM' } | Select-Object -First 1
            try {
                $apimDetail = Get-AzApiManagement -ResourceGroupName $ResourceGroup -Name $apimRes.Name -ErrorAction Stop
                Write-Log "  APIM '$($apimRes.Name)' state=$($apimDetail.ProvisioningState), SKU=$($apimDetail.Sku.Name) — will be reused." -Level SUCCESS
            } catch {
                Write-Log "  APIM info query failed: $($_.Exception.Message)" -Level DEBUG
            }
        }

        # ── 1c. Check Key Vault purge protection status ──
        if ($Script:ExistingResources.KeyVault) {
            $kvRes = $rgResources | Where-Object { $typeMap[$_.ResourceType] -eq 'KeyVault' } | Select-Object -First 1
            try {
                $kvDetail = Get-AzKeyVault -VaultName $kvRes.Name -ResourceGroupName $ResourceGroup -ErrorAction Stop
                if ($kvDetail.EnablePurgeProtection) {
                    Write-Log "  Key Vault '$($kvRes.Name)' has purge protection ENABLED (possibly Azure Policy enforced)." -Level INFO
                    Write-Log "  On teardown this vault will be soft-deleted but cannot be purged — redeploys will auto-recover it." -Level INFO
                }
            } catch {
                Write-Log "  Key Vault detail query failed: $($_.Exception.Message)" -Level DEBUG
            }
        }
    } else {
        Write-Log "  No existing resources found in '$ResourceGroup'. Fresh deployment." -Level INFO
    }

    # ── 2. RESOLVE SOFT-DELETED RESOURCES ──
    $resolved = 0

    # ── 2a. Key Vault soft-delete ──
    # Skip soft-delete check if KV already lives in the target RG — nothing to recover
    $kvName = if ($env:WINGETMM_OVERRIDE_KEYVAULT) { $env:WINGETMM_OVERRIDE_KEYVAULT } else { "kv-$cleanName" }
    if (-not $Script:ExistingResources.KeyVault) {
        Write-Log "Checking for soft-deleted Key Vault '$kvName'..." -Level INFO
        try {
            $kv = Get-AzKeyVault -VaultName $kvName -Location $Region -InRemovedState -ErrorAction SilentlyContinue
            if ($kv) {
                Write-Log "  Found soft-deleted Key Vault '$kvName'. Recovering into '$ResourceGroup'..." -Level WARN
                Write-Log "  (Recovery is the safe path — purge may be blocked by Azure Policy.)" -Level INFO
                $body = @{
                    location   = $Region
                    properties = @{
                        createMode = 'recover'
                        tenantId   = (Get-AzContext).Tenant.Id
                        sku        = @{ family = 'A'; name = 'standard' }
                    }
                } | ConvertTo-Json -Depth 5
                $kvUri = "https://management.azure.com/subscriptions/$subId/resourceGroups/$ResourceGroup/providers/Microsoft.KeyVault/vaults/$kvName`?api-version=2023-07-01"
                $resp = Invoke-RestMethod -Uri $kvUri -Method PUT -Headers $headers -Body $body -ErrorAction Stop
                Write-Log "  Key Vault recovered (provisioning: $($resp.properties.provisioningState))." -Level SUCCESS
                $Script:ExistingResources.KeyVault = $true
                $resolved++
            } else {
                Write-Log "  No soft-deleted Key Vault found. OK" -Level DEBUG
            }
        } catch {
            Write-Log "  Key Vault recovery check failed: $($_.Exception.Message)" -Level WARN
            Write-Log "  Deployment will attempt to proceed — ARM template may create a new one." -Level WARN
        }
    } else {
        Write-Log "Key Vault '$kvName' exists in RG — skipping soft-delete check." -Level DEBUG
    }

    # ── 2b. API Management soft-delete ──
    $apimName = if ($env:WINGETMM_OVERRIDE_APIM) { $env:WINGETMM_OVERRIDE_APIM } else { "apim-$cleanName" }
    if (-not $Script:ExistingResources.APIM) {
        Write-Log "Checking for soft-deleted APIM '$apimName'..." -Level INFO
        try {
            $apimDelUri = "https://management.azure.com/subscriptions/$subId/providers/Microsoft.ApiManagement/deletedservices?api-version=2022-08-01"
            $apimDel = Invoke-RestMethod -Uri $apimDelUri -Headers $headers -ErrorAction Stop
            $match = $apimDel.value | Where-Object { $_.name -eq $apimName }
            if ($match) {
                Write-Log "  Found soft-deleted APIM '$apimName'. Purging to free the name..." -Level WARN
                $purgeUri = "https://management.azure.com/subscriptions/$subId/providers/Microsoft.ApiManagement/locations/$Region/deletedservices/$apimName`?api-version=2022-08-01"
                Invoke-RestMethod -Uri $purgeUri -Method DELETE -Headers $headers -ErrorAction Stop
                Write-Log "  Waiting for APIM purge to complete..." -Level INFO
                $purgeWait = 0
                do {
                    Start-Sleep -Seconds 10
                    $purgeWait += 10
                    $tok     = Get-PlainToken
                    $headers = @{ Authorization = "Bearer $tok"; 'Content-Type' = 'application/json' }
                    $recheck = Invoke-RestMethod -Uri $apimDelUri -Headers $headers -ErrorAction Stop
                    $still = $recheck.value | Where-Object { $_.name -eq $apimName }
                } while ($still -and $purgeWait -lt 120)

                if ($still) {
                    Write-Log "  APIM purge still in progress after ${purgeWait}s — deployment will retry if needed." -Level WARN
                } else {
                    Write-Log "  APIM '$apimName' purged successfully." -Level SUCCESS
                    $resolved++
                }
            } else {
                Write-Log "  No soft-deleted APIM found. OK" -Level DEBUG
            }
        } catch {
            Write-Log "  APIM soft-delete check failed: $($_.Exception.Message)" -Level WARN
        }
    } else {
        Write-Log "APIM '$apimName' exists in RG — skipping soft-delete check." -Level DEBUG
    }

    # ── 2c. Function App (informational only — cannot be explicitly purged) ──
    $funcName = if ($env:WINGETMM_OVERRIDE_FUNCAPP) { $env:WINGETMM_OVERRIDE_FUNCAPP } else { "func-$cleanName" }
    if (-not $Script:ExistingResources.FunctionApp) {
        try {
            $webDelUri = "https://management.azure.com/subscriptions/$subId/providers/Microsoft.Web/deletedSites?api-version=2023-12-01"
            $webDel = Invoke-RestMethod -Uri $webDelUri -Headers $headers -ErrorAction Stop
            $webMatch = $webDel.value | Where-Object { $_.properties.deletedSiteName -eq $funcName }
            if ($webMatch) {
                Write-Log "  Note: Soft-deleted function app '$funcName' exists (auto-purges after 30 days). Should not block deployment." -Level INFO
            }
        } catch {
            # Non-critical — just informational
        }
    }

    # ── 3. SUMMARY ──
    $reuseList = ($Script:ExistingResources.GetEnumerator() | Where-Object { $_.Value -eq $true } | ForEach-Object { $_.Key }) -join ', '
    if ($reuseList) {
        Write-Log "Resources to reuse: $reuseList" -Level SUCCESS
    }
    if ($resolved -gt 0) {
        Write-Log "Recovered/purged $resolved soft-deleted resource(s)." -Level SUCCESS
    }
    if (-not $reuseList -and $resolved -eq 0) {
        Write-Log "Clean slate — all resources will be created fresh." -Level SUCCESS
    }

    $Script:Stats.ResourceAudit = if ($existingCount -gt 0) { "$existingCount reused" } else { 'Clean' }
}

# ═══════════════════════════════════════════════════════════════════════════════
#  PHASE 4 — DEPLOY WINGET REST SOURCE
# ═══════════════════════════════════════════════════════════════════════════════

function Step-DeployRestSource {
    Write-Log "PHASE 4: DEPLOY WINGET REST SOURCE" -Level SECTION

    Write-Log "Configuration:" -Level INFO
    Write-Log "  Name              : $Name" -Level INFO
    Write-Log "  ResourceGroup     : $ResourceGroup" -Level INFO
    Write-Log "  Region            : $Region" -Level INFO
    Write-Log "  PerformanceTier   : $PerformanceTier" -Level INFO
    Write-Log "  CosmosDB Zone-HA  : $(if($CosmosDBZoneRedundant){'Enabled'}else{'Disabled (default)'})" -Level INFO
    Write-Log "  CosmosDB Region   : $(if($CosmosDBRegion){$CosmosDBRegion}else{"$Region (same as primary)"})" -Level INFO
    Write-Log "  Authentication    : $Authentication" -Level INFO
    Write-Log "  Publisher         : $Script:ResolvedPublisherName <$Script:ResolvedPublisherEmail>" -Level INFO
    Write-Log "" -Level INFO

    # Time estimate based on existing resource inventory
    if ($Script:ExistingResources.APIM) {
        Write-Log "APIM already exists — skipping 30-40 min creation. Estimated time: 5-15 minutes (update-in-place)." -Level SUCCESS
    } else {
        Write-Log "This step typically takes 30-60 minutes (APIM creation is the bottleneck)." -Level WARN
    }

    # Resolve Entra ID settings
    $entraResource      = $MicrosoftEntraIdResource
    $entraResourceScope = $MicrosoftEntraIdResourceScope
    $createEntraApp     = $false

    if ($Authentication -eq 'MicrosoftEntraId') {
        Write-Log "Entra ID authentication enabled." -Level INFO

        if ($entraResource) {
            Write-Log "  Using existing app: $entraResource" -Level INFO
            Write-Log "  Resource Scope    : $entraResourceScope" -Level INFO
        } else {
            $createEntraApp = $true
            Write-Log "  Will create new Entra ID app registration." -Level INFO
        }
    }

    if ($PSCmdlet.ShouldProcess("Azure ($Region)", "Deploy WinGet REST source '$Name'")) {

        if (-not $CosmosDBZoneRedundant -and -not $CosmosDBRegion) {
            # ── DEFAULT PATH: delegate everything to New-WinGetSource ──
            $deployParams = @{
                Name                       = $Name
                ResourceGroup              = $ResourceGroup
                Region                     = $Region
                ImplementationPerformance  = $PerformanceTier
                PublisherName              = $Script:ResolvedPublisherName
                PublisherEmail             = $Script:ResolvedPublisherEmail
                ShowConnectionInstructions = $true
                InformationAction          = 'Continue'
                Verbose                    = $true
            }
            if ($RestSourcePath) {
                $deployParams['RestSourcePath'] = (Resolve-Path $RestSourcePath).Path
            }
            if ($Authentication -eq 'MicrosoftEntraId') {
                $deployParams['RestSourceAuthentication'] = 'MicrosoftEntraId'
                if ($entraResource) {
                    $deployParams['MicrosoftEntraIdResource']      = $entraResource
                    $deployParams['MicrosoftEntraIdResourceScope'] = $entraResourceScope
                } else {
                    $deployParams['CreateNewMicrosoftEntraIdAppRegistration'] = $true
                }
            }

            # Start background status poller
            $pollerJob = Watch-DeploymentProgress -StepName 'New-WinGetSource' -RG $ResourceGroup
            try {
                $deployResult = Invoke-StepWithRetry -StepName "New-WinGetSource" -MaxRetries 2 -BaseDelaySec 30 -Action {
                    New-WinGetSource @deployParams
                }
            } finally {
                Stop-Job $pollerJob -ErrorAction SilentlyContinue
                Remove-Job $pollerJob -Force -ErrorAction SilentlyContinue
            }

            # Try to derive APIM URL for source registration
            try {
                $apimName = "apim-$($Name -replace '[^a-zA-Z0-9-]', '')"
                $apimUrl = (Get-AzApiManagement -Name $apimName -ResourceGroupName $ResourceGroup -ErrorAction Stop).RuntimeUrl
                $Script:DeploymentUrl = "$apimUrl/winget/"
            } catch {
                Write-Log "Could not retrieve APIM URL, will use direct function URL." -Level WARN
                $Script:DeploymentUrl = "https://func-${Name}.azurewebsites.net/api/"
            }
        } else {
            # ── CUSTOM PATH: split workflow to patch Cosmos DB parameters ──
            $overrides = @()
            if ($CosmosDBZoneRedundant) { $overrides += 'zone redundancy' }
            if ($CosmosDBRegion)        { $overrides += "region override → $CosmosDBRegion" }
            Write-Log "Using split deployment workflow ($($overrides -join ', '))." -Level INFO

            $mod = Get-Module Microsoft.WinGet.RestSource
            $modBase = $mod.ModuleBase
            $templateFolder   = Join-Path $modBase 'Data\ARMTemplates'
            $restSourceZip    = if ($RestSourcePath) { (Resolve-Path $RestSourcePath).Path } else { Join-Path $modBase 'Data\WinGet.RestSource.Functions.zip' }
            $paramOutputPath  = Join-Path (Get-Location).Path 'Parameters'

            # Step 1: Create Entra ID app if needed
            if ($Authentication -eq 'MicrosoftEntraId' -and $createEntraApp) {
                Write-Log "Creating Entra ID app registration..." -Level INFO
                $entraResult = & $mod { param($n) New-MicrosoftEntraIdApp -Name $n } $Name
                if (-not $entraResult.Result) {
                    throw "Failed to create Entra ID app registration."
                }
                $entraResource      = $entraResult.Resource
                $entraResourceScope = $entraResult.ResourceScope
                Write-Log "  App created: $entraResource (scope: $entraResourceScope)" -Level SUCCESS
            }

            # Step 2: Generate ARM parameter files
            Write-Log "Generating ARM parameter files..." -Level INFO
            if (-not (Test-Path $paramOutputPath)) {
                New-Item -Path $paramOutputPath -ItemType Directory -Force | Out-Null
                Write-Log "  Created output directory: $paramOutputPath" -Level DEBUG
            }
            $armParams = @{
                ParameterFolderPath       = $paramOutputPath
                TemplateFolderPath        = $templateFolder
                Name                      = $Name
                Region                    = $Region
                ImplementationPerformance = $PerformanceTier
                PublisherName             = $Script:ResolvedPublisherName
                PublisherEmail            = $Script:ResolvedPublisherEmail
                RestSourceAuthentication  = $Authentication
            }
            if ($entraResource)      { $armParams['MicrosoftEntraIdResource']      = $entraResource }
            if ($entraResourceScope) { $armParams['MicrosoftEntraIdResourceScope'] = $entraResourceScope }

            $ARMObjects = & $mod { param($p) New-ARMParameterObjects @p } $armParams
            if (-not $ARMObjects) { throw 'Failed to create ARM parameter objects.' }

            # Step 3: Patch Cosmos DB parameter file (only the overrides requested)
            $cosmosParamFile = Join-Path $paramOutputPath 'cosmosdb.json'
            if (Test-Path $cosmosParamFile) {
                $cosmosJson = Get-Content $cosmosParamFile -Raw | ConvertFrom-Json -Depth 20
                $patched = $false

                # 3a: Zone redundancy override
                if ($CosmosDBZoneRedundant) {
                    Write-Log "Patching Cosmos DB: enabling zone redundancy..." -Level INFO
                    foreach ($loc in $cosmosJson.Parameters.locations.value) {
                        $loc.isZoneRedundant = $true
                        Write-Log "  $($loc.locationName): isZoneRedundant = true" -Level INFO
                    }
                    $patched = $true
                }

                # 3b: Region override
                if ($CosmosDBRegion) {
                    Write-Log "Patching Cosmos DB: relocating to $CosmosDBRegion..." -Level INFO
                    # Top-level location parameter
                    $cosmosJson.Parameters.location.value = $CosmosDBRegion
                    Write-Log "  location: $CosmosDBRegion" -Level INFO
                    # Locations array — update every entry's locationName
                    # Azure ARM expects the display name (e.g. "North Europe"), derive it
                    $regionDisplayName = (Get-AzLocation | Where-Object Location -eq $CosmosDBRegion).DisplayName
                    if (-not $regionDisplayName) { $regionDisplayName = $CosmosDBRegion }
                    foreach ($loc in $cosmosJson.Parameters.locations.value) {
                        $oldName = $loc.locationName
                        $loc.locationName = $regionDisplayName
                        Write-Log "  locations[]: $oldName → $regionDisplayName" -Level INFO
                    }
                    $patched = $true
                }

                if ($patched) {
                    $cosmosJson | ConvertTo-Json -Depth 20 | Set-Content $cosmosParamFile -Encoding UTF8
                    Write-Log "Cosmos DB parameter file updated." -Level SUCCESS
                }
            } else {
                Write-Log "Cosmos DB parameter file not found at $cosmosParamFile — skipping patch." -Level WARN
            }

            # Step 4: Create/verify Resource Group
            Write-Log "Ensuring resource group '$ResourceGroup' exists..." -Level INFO
            $rgResult = & $mod { param($n,$r) Add-AzureResourceGroup -Name $n -Region $r } $ResourceGroup $Region
            if (-not $rgResult) { throw "Failed to create resource group '$ResourceGroup' in '$Region'." }

            # Step 5: Validate ARM templates
            Write-Log "Validating ARM templates..." -Level INFO
            $valResult = & $mod { param($a,$rg) Test-ARMTemplates -ARMObjects $a -ResourceGroup $rg } $ARMObjects $ResourceGroup
            if ($valResult) {
                throw "ARM template validation failed. Check parameter files in '$paramOutputPath'."
            }
            Write-Log "ARM template validation passed." -Level SUCCESS

            # Step 6: Deploy ARM resources
            if ($Script:ExistingResources.APIM) {
                Write-Log "Deploying ARM resources (APIM reused — estimated 5-15 minutes)..." -Level INFO
            } else {
                Write-Log "Deploying ARM resources (APIM creation — estimated 30-40 minutes)..." -Level INFO
            }

            # Start background status poller
            $pollerJob = Watch-DeploymentProgress -StepName 'ARM-Deploy' -RG $ResourceGroup
            try {
                $deployOk = $false
                for ($attempt = 1; $attempt -le 5; $attempt++) {
                    $deployOk = & $mod { param($a,$z,$rg) New-ARMObjects -ARMObjects ([ref]$a) -RestSourcePath $z -ResourceGroup $rg } $ARMObjects $restSourceZip $ResourceGroup
                    if ($deployOk) { break }
                    if ($attempt -lt 5) {
                        Write-Log "Deployment attempt $attempt failed, retrying in 15s..." -Level WARN
                        Start-Sleep -Seconds 15
                    } else {
                        throw "Failed to deploy ARM resources after $attempt attempts."
                    }
                }
            } finally {
                Stop-Job $pollerJob -ErrorAction SilentlyContinue
                Remove-Job $pollerJob -Force -ErrorAction SilentlyContinue
            }

            # Step 7: Show connection instructions
            $apiMgmtName = $ARMObjects.Where({$_.ObjectType -eq 'ApiManagement'}).Parameters.Parameters.serviceName.value
            try {
                $apiMgmtUrl = (Get-AzApiManagement -Name $apiMgmtName -ResourceGroupName $ResourceGroup).RuntimeUrl
                $Script:DeploymentUrl = "$apiMgmtUrl/winget/"
                Write-Log "Connection command:" -Level SUCCESS
                Write-Log "  winget source add -n `"$Name`" -a `"$Script:DeploymentUrl`" -t `"Microsoft.Rest`"" -Level INFO
            } catch {
                Write-Log "Could not retrieve API Management URL: $($_.Exception.Message)" -Level WARN
                $Script:DeploymentUrl = "https://func-${Name}.azurewebsites.net/api/"
            }
        }

        # Derive the REST source URL (prefer APIM URL set above, else direct function URL)
        if (-not $Script:DeploymentUrl) {
            $Script:DeploymentUrl = "https://func-${Name}.azurewebsites.net/api/"
        }
        Write-Log "REST Source URL: $Script:DeploymentUrl" -Level SUCCESS

        # ── POST-DEPLOY: Remove APIM if -SkipAPIM was specified ──
        if ($SkipAPIM) {
            $apimCleanName = "apim-$($Name -replace '[^a-zA-Z0-9-]', '')"
            Write-Log "APIM REMOVAL (-SkipAPIM)" -Level SECTION
            Write-Log "Removing API Management '$apimCleanName' to reduce cost and simplify architecture..." -Level INFO
            try {
                $existingApim = Get-AzApiManagement -Name $apimCleanName -ResourceGroupName $ResourceGroup -ErrorAction Stop
                if ($existingApim) {
                    Remove-AzApiManagement -ResourceGroupName $ResourceGroup -Name $apimCleanName -ErrorAction Stop
                    Write-Log "APIM '$apimCleanName' removed successfully." -Level SUCCESS
                    Write-Log "Clients should use the direct Function App URL instead:" -Level INFO
                }
            } catch {
                $errMsg = $_.Exception.Message
                if ($errMsg -match 'not found|does not exist') {
                    Write-Log "APIM '$apimCleanName' not found — may not have been created. OK" -Level DEBUG
                } else {
                    Write-Log "Failed to remove APIM '$apimCleanName': $errMsg" -Level WARN
                    Write-Log "You can remove it manually: Remove-AzApiManagement -ResourceGroupName '$ResourceGroup' -Name '$apimCleanName'" -Level WARN
                }
            }
            # Force URL to direct Function App endpoint
            $Script:DeploymentUrl = "https://func-${Name}.azurewebsites.net/api/"
            Write-Log "REST Source URL (direct): $Script:DeploymentUrl" -Level SUCCESS
        }

        $Script:Stats.Deployment = 'Succeeded'

        # ── POST-DEPLOY: Apply resource tags if specified ──
        if ($DeployTags.Count -gt 0) {
            Write-Log "Applying $($DeployTags.Count) tag(s) to all resources in '$ResourceGroup'..." -Level INFO
            try {
                $allRes = Get-AzResource -ResourceGroupName $ResourceGroup -ErrorAction Stop
                foreach ($res in $allRes) {
                    $currentTags = if ($res.Tags) { [hashtable]$res.Tags } else { @{} }
                    $updated = $false
                    foreach ($key in $DeployTags.Keys) {
                        if ($currentTags[$key] -ne $DeployTags[$key]) {
                            $currentTags[$key] = $DeployTags[$key]
                            $updated = $true
                        }
                    }
                    if ($updated) {
                        Set-AzResource -ResourceId $res.ResourceId -Tag $currentTags -Force -ErrorAction SilentlyContinue | Out-Null
                    }
                }
                Write-Log "Tags applied to $($allRes.Count) resource(s)." -Level SUCCESS
            } catch {
                Write-Log "Failed to apply tags: $($_.Exception.Message)" -Level WARN
            }
        }
    } else {
        Write-Log "[WhatIf] Would deploy WinGet REST source '$Name' to '$ResourceGroup' in '$Region'." -Level WARN
        $Script:DeploymentUrl = "https://func-${Name}.azurewebsites.net/api/ (not deployed)"
        $Script:Stats.Deployment = 'WhatIf'
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
#  PHASE 4b — POST-DEPLOY HEALTH CHECK
# ═══════════════════════════════════════════════════════════════════════════════

function Step-PostDeployHealthCheck {
    if ($Script:Stats.Deployment -eq 'WhatIf') { return }

    Write-Log "POST-DEPLOY HEALTH CHECK" -Level SECTION

    $funcAppName = "func-$($Name -replace '[^a-zA-Z0-9-]', '')"
    Write-Log "Checking function app '$funcAppName' for loaded functions..." -Level INFO

    # Wait up to 2 min for functions to appear (cold start)
    $maxWait = 120; $waited = 0; $funcCount = 0
    while ($waited -lt $maxWait) {
        try {
            $tok = Get-PlainToken
            $hdr = @{ Authorization = "Bearer $tok"; 'Content-Type' = 'application/json' }
            $subId = (Get-AzContext).Subscription.Id
            $uri = "https://management.azure.com/subscriptions/$subId/resourceGroups/$ResourceGroup/providers/Microsoft.Web/sites/$funcAppName/functions?api-version=2023-12-01"
            $resp = Invoke-RestMethod -Uri $uri -Headers $hdr -Method Get
            $funcCount = $resp.value.Count
        } catch {
            Write-Log "  Could not query functions: $($_.Exception.Message)" -Level WARN
        }
        if ($funcCount -gt 0) { break }
        if ($waited -eq 0) {
            Write-Log "  No functions loaded yet, waiting for cold start..." -Level INFO
        }
        Start-Sleep -Seconds 15
        $waited += 15
    }

    if ($funcCount -gt 0) {
        Write-Log "  Function app has $funcCount functions loaded. ✓" -Level SUCCESS
        return
    }

    # Functions still not loaded — attempt manual ZipDeploy
    Write-Log "  WARNING: Function app has 0 functions after $maxWait seconds." -Level WARN
    Write-Log "  Attempting manual ZipDeploy of function code..." -Level INFO

    $mod = Get-Module Microsoft.WinGet.RestSource
    $zipPath = if ($RestSourcePath) { (Resolve-Path $RestSourcePath).Path } else { Join-Path $mod.ModuleBase 'Data\WinGet.RestSource.Functions.zip' }
    if (-not (Test-Path $zipPath)) {
        Write-Log "  Function zip not found at $zipPath — cannot recover." -Level ERROR
        return
    }

    try {
        # Get publishing credentials
        $pubCreds = Invoke-AzResourceAction -ResourceGroupName $ResourceGroup `
            -ResourceType 'Microsoft.Web/sites' -ResourceName $funcAppName `
            -Action 'publishxml' -ApiVersion '2023-01-01' -Force
        # Parse publishProfile XML for user/pass
        $xml = [xml]$pubCreds.InnerXml
        if (-not $xml) {
            $xml = [xml]$pubCreds
        }
        $profile = $xml.publishData.publishProfile | Where-Object publishMethod -eq "MSDeploy" | Select-Object -First 1
        $user = $profile.userName
        $pass = $profile.userPWD

        $pair = "$($user):$($pass)"
        $encodedCreds = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($pair))

        $zipBytes = [System.IO.File]::ReadAllBytes($zipPath)
        $zipDeployUri = "https://$funcAppName.scm.azurewebsites.net/api/zipdeploy"

        Write-Log "  Uploading $([math]::Round($zipBytes.Length / 1MB, 1)) MB to $zipDeployUri..." -Level INFO
        $resp = Invoke-WebRequest -Uri $zipDeployUri -Method Post `
            -Headers @{ Authorization = "Basic $encodedCreds" } `
            -ContentType "application/zip" -Body $zipBytes `
            -UseBasicParsing -TimeoutSec 600

        if ($resp.StatusCode -eq 200) {
            Write-Log "  ZipDeploy succeeded (HTTP 200). Waiting 30s for functions to load..." -Level SUCCESS
            Start-Sleep -Seconds 30

            # Re-check
            try {
                $tok = Get-PlainToken
                $hdr = @{ Authorization = "Bearer $tok"; 'Content-Type' = 'application/json' }
                $resp2 = Invoke-RestMethod -Uri $uri -Headers $hdr -Method Get
                Write-Log "  Function app now has $($resp2.value.Count) functions. ✓" -Level SUCCESS
            } catch {
                Write-Log "  Could not re-check function count: $($_.Exception.Message)" -Level WARN
            }
        } else {
            Write-Log "  ZipDeploy returned HTTP $($resp.StatusCode) — may need manual intervention." -Level ERROR
        }
    } catch {
        Write-Log "  Manual ZipDeploy failed: $($_.Exception.Message)" -Level ERROR
        Write-Log "  You can manually deploy: Invoke-WebRequest -Uri 'https://$funcAppName.scm.azurewebsites.net/api/zipdeploy' ..." -Level WARN
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
#  PHASE 4c — CREATE PRIVATE ENDPOINTS
# ═══════════════════════════════════════════════════════════════════════════════

function Step-CreatePrivateEndpoints {
    <#
    .SYNOPSIS Creates Azure Private Endpoints for all WinGet REST source resources.
    .DESCRIPTION
        Creates Private Endpoints for:
          - Function App   (groupId: sites)
          - Cosmos DB       (groupId: Sql)
          - Key Vault       (groupId: vault)
          - Storage Account (groupId: blob)
        Optionally creates Private DNS Zones and links them to the VNet.
        Disables public network access on each resource after PE creation.
    #>
    if (-not $EnablePrivateEndpoints) { return }
    if ($Script:Stats.Deployment -eq 'WhatIf') {
        Write-Log "[WhatIf] Would create private endpoints for deployed resources." -Level WARN
        $Script:Stats.PrivateEndpoints = 'WhatIf'
        return
    }

    Write-Log "PHASE 4c: CREATE PRIVATE ENDPOINTS" -Level SECTION

    # ── Validate PE parameters ──
    if (-not $VNetName) {
        Write-Log "-EnablePrivateEndpoints requires -VNetName. Skipping PE creation." -Level ERROR
        $Script:Stats.PrivateEndpoints = 'Failed (no VNet)'
        return
    }

    $peVNetRG = if ($VNetResourceGroup) { $VNetResourceGroup } else { $ResourceGroup }

    # ── Resolve VNet & Subnet ──
    Write-Log "Resolving VNet '$VNetName' in resource group '$peVNetRG'..." -Level INFO
    try {
        $vnet = Get-AzVirtualNetwork -Name $VNetName -ResourceGroupName $peVNetRG -ErrorAction Stop
        Write-Log "  VNet found: $($vnet.Id)" -Level SUCCESS
    } catch {
        Write-Log "VNet '$VNetName' not found in '$peVNetRG': $($_.Exception.Message)" -Level ERROR
        $Script:Stats.PrivateEndpoints = 'Failed (VNet not found)'
        return
    }

    $subnet = $vnet.Subnets | Where-Object Name -eq $SubnetName
    if (-not $subnet) {
        Write-Log "Subnet '$SubnetName' not found in VNet '$VNetName'." -Level ERROR
        Write-Log "Available subnets: $(($vnet.Subnets | ForEach-Object { $_.Name }) -join ', ')" -Level INFO
        $Script:Stats.PrivateEndpoints = 'Failed (subnet not found)'
        return
    }
    Write-Log "  Subnet: $($subnet.Id)" -Level SUCCESS

    # ── Discover deployed resources ──
    $cleanName  = $Name -replace '[^a-zA-Z0-9-]', ''
    $subId      = (Get-AzContext).Subscription.Id
    $tok        = Get-PlainToken
    $headers    = @{ Authorization = "Bearer $tok"; 'Content-Type' = 'application/json' }
    $rgRes      = Get-AzResource -ResourceGroupName $ResourceGroup -ErrorAction SilentlyContinue

    # Build resource map: DisplayName → @{ ResourceId, GroupId, DnsZone, ResourceType }
    $peTargets = [ordered]@{}

    # Function App
    $funcApp = $rgRes | Where-Object { $_.ResourceType -eq 'Microsoft.Web/sites' } | Select-Object -First 1
    if ($funcApp) {
        $peTargets['FunctionApp'] = @{
            ResourceId   = $funcApp.ResourceId
            GroupId      = 'sites'
            DnsZoneName  = 'privatelink.azurewebsites.net'
            ResourceName = $funcApp.Name
        }
    } else {
        Write-Log "  WARNING: Function App not found in '$ResourceGroup' — skipping PE." -Level WARN
    }

    # Cosmos DB
    $cosmosDb = $rgRes | Where-Object { $_.ResourceType -eq 'Microsoft.DocumentDB/databaseAccounts' } | Select-Object -First 1
    if ($cosmosDb) {
        $peTargets['CosmosDB'] = @{
            ResourceId   = $cosmosDb.ResourceId
            GroupId      = 'Sql'
            DnsZoneName  = 'privatelink.documents.azure.com'
            ResourceName = $cosmosDb.Name
        }
    } else {
        Write-Log "  WARNING: Cosmos DB not found in '$ResourceGroup' — skipping PE." -Level WARN
    }

    # Key Vault
    $keyVault = $rgRes | Where-Object { $_.ResourceType -eq 'Microsoft.KeyVault/vaults' } | Select-Object -First 1
    if ($keyVault) {
        $peTargets['KeyVault'] = @{
            ResourceId   = $keyVault.ResourceId
            GroupId      = 'vault'
            DnsZoneName  = 'privatelink.vaultcore.azure.net'
            ResourceName = $keyVault.Name
        }
    } else {
        Write-Log "  WARNING: Key Vault not found in '$ResourceGroup' — skipping PE." -Level WARN
    }

    # Storage Account
    $storageAcct = $rgRes | Where-Object { $_.ResourceType -eq 'Microsoft.Storage/storageAccounts' } | Select-Object -First 1
    if ($storageAcct) {
        $peTargets['StorageAccount'] = @{
            ResourceId   = $storageAcct.ResourceId
            GroupId      = 'blob'
            DnsZoneName  = 'privatelink.blob.core.windows.net'
            ResourceName = $storageAcct.Name
        }
    } else {
        Write-Log "  WARNING: Storage Account not found in '$ResourceGroup' — skipping PE." -Level WARN
    }

    if ($peTargets.Count -eq 0) {
        Write-Log "No resources found to create private endpoints for." -Level ERROR
        $Script:Stats.PrivateEndpoints = 'Failed (no resources)'
        return
    }

    Write-Log "Creating private endpoints for $($peTargets.Count) resource(s)..." -Level INFO
    $peCreated = 0
    $peFailed  = 0

    foreach ($entry in $peTargets.GetEnumerator()) {
        $label    = $entry.Key
        $target   = $entry.Value
        $peName   = "pe-$($target.ResourceName)"

        Write-Log "  ── ${label}: $($target.ResourceName) ──" -Level INFO

        # Check if PE already exists
        $existingPE = Get-AzPrivateEndpoint -Name $peName -ResourceGroupName $ResourceGroup -ErrorAction SilentlyContinue
        if ($existingPE) {
            Write-Log "    Private endpoint '$peName' already exists — skipping." -Level SUCCESS
            $peCreated++
            continue
        }

        try {
            # Create private link service connection
            $plsConn = New-AzPrivateLinkServiceConnection `
                -Name "plsc-$($target.ResourceName)" `
                -PrivateLinkServiceId $target.ResourceId `
                -GroupId $target.GroupId

            # Create private endpoint
            Write-Log "    Creating private endpoint '$peName'..." -Level INFO
            $pe = New-AzPrivateEndpoint `
                -Name $peName `
                -ResourceGroupName $ResourceGroup `
                -Location $Region `
                -Subnet $subnet `
                -PrivateLinkServiceConnection $plsConn `
                -ErrorAction Stop

            Write-Log "    Private endpoint created: $($pe.Id)" -Level SUCCESS

            # ── Optional: Private DNS Zone + link ──
            if ($RegisterPrivateDnsZones) {
                $dnsZoneName = $target.DnsZoneName
                Write-Log "    Creating/linking Private DNS Zone '$dnsZoneName'..." -Level INFO

                # Create DNS zone if it doesn't exist
                $dnsZone = Get-AzPrivateDnsZone -Name $dnsZoneName -ResourceGroupName $ResourceGroup -ErrorAction SilentlyContinue
                if (-not $dnsZone) {
                    $dnsZone = New-AzPrivateDnsZone `
                        -Name $dnsZoneName `
                        -ResourceGroupName $ResourceGroup `
                        -ErrorAction Stop
                    Write-Log "    DNS zone created." -Level SUCCESS
                } else {
                    Write-Log "    DNS zone already exists." -Level DEBUG
                }

                # Link DNS zone to VNet (if not already linked)
                $linkName = "vnetlink-$($VNetName)-$($dnsZoneName -replace '\.', '-')"
                $existingLink = Get-AzPrivateDnsVirtualNetworkLink -ZoneName $dnsZoneName `
                    -ResourceGroupName $ResourceGroup -Name $linkName -ErrorAction SilentlyContinue
                if (-not $existingLink) {
                    New-AzPrivateDnsVirtualNetworkLink `
                        -ZoneName $dnsZoneName `
                        -ResourceGroupName $ResourceGroup `
                        -Name $linkName `
                        -VirtualNetworkId $vnet.Id `
                        -EnableRegistration:$false `
                        -ErrorAction Stop | Out-Null
                    Write-Log "    VNet link '$linkName' created." -Level SUCCESS
                } else {
                    Write-Log "    VNet link already exists." -Level DEBUG
                }

                # Create DNS zone group on the PE to auto-register A records
                $dnsGroupName = "dnszg-$($target.ResourceName)"
                $dnsZoneConfig = New-AzPrivateDnsZoneConfig `
                    -Name $dnsZoneName `
                    -PrivateDnsZoneId $dnsZone.ResourceId

                # Use REST API for DNS zone group (Az.Network may not have the cmdlet in all versions)
                try {
                    $zoneGroupUri = "https://management.azure.com$($pe.Id)/privateDnsZoneGroups/$dnsGroupName`?api-version=2023-11-01"
                    $zoneGroupBody = @{
                        properties = @{
                            privateDnsZoneConfigs = @(
                                @{
                                    name       = ($dnsZoneName -replace '\.', '-')
                                    properties = @{
                                        privateDnsZoneId = $dnsZone.ResourceId
                                    }
                                }
                            )
                        }
                    } | ConvertTo-Json -Depth 5
                    $tok     = Get-PlainToken
                    $headers = @{ Authorization = "Bearer $tok"; 'Content-Type' = 'application/json' }
                    Invoke-RestMethod -Uri $zoneGroupUri -Method PUT -Headers $headers -Body $zoneGroupBody -ErrorAction Stop | Out-Null
                    Write-Log "    DNS zone group '$dnsGroupName' created — A records will auto-register." -Level SUCCESS
                } catch {
                    Write-Log "    DNS zone group creation failed: $($_.Exception.Message)" -Level WARN
                    Write-Log "    You may need to manually add A records to '$dnsZoneName'." -Level WARN
                }
            }

            $peCreated++
        } catch {
            Write-Log "    FAILED to create PE for $label`: $($_.Exception.Message)" -Level ERROR
            $peFailed++
        }
    }

    # ── Disable public network access ──
    Write-Log "" -Level INFO
    Write-Log "Disabling public network access on resources..." -Level INFO

    # Function App — set access restriction to deny all
    if ($peTargets.ContainsKey('FunctionApp')) {
        try {
            $funcName = $peTargets['FunctionApp'].ResourceName
            Write-Log "  Function App '$funcName': disabling public access..." -Level INFO
            $updateUri = "https://management.azure.com/subscriptions/$subId/resourceGroups/$ResourceGroup/providers/Microsoft.Web/sites/$funcName/config/web?api-version=2023-12-01"
            $tok     = Get-PlainToken
            $headers = @{ Authorization = "Bearer $tok"; 'Content-Type' = 'application/json' }
            $body = @{
                properties = @{
                    publicNetworkAccess = 'Disabled'
                }
            } | ConvertTo-Json -Depth 3
            Invoke-RestMethod -Uri $updateUri -Method PATCH -Headers $headers -Body $body -ErrorAction Stop | Out-Null
            Write-Log "  Function App: public access disabled. ✓" -Level SUCCESS
        } catch {
            Write-Log "  Function App: failed to disable public access: $($_.Exception.Message)" -Level WARN
        }
    }

    # Cosmos DB
    if ($peTargets.ContainsKey('CosmosDB')) {
        try {
            $cosmosName = $peTargets['CosmosDB'].ResourceName
            Write-Log "  Cosmos DB '$cosmosName': disabling public access..." -Level INFO
            $updateUri = "https://management.azure.com/subscriptions/$subId/resourceGroups/$ResourceGroup/providers/Microsoft.DocumentDB/databaseAccounts/$cosmosName`?api-version=2024-05-15"
            $tok     = Get-PlainToken
            $headers = @{ Authorization = "Bearer $tok"; 'Content-Type' = 'application/json' }
            # Cosmos DB requires full PUT with location — use PATCH via properties
            $currentCosmos = Invoke-RestMethod -Uri $updateUri -Headers $headers -Method GET -ErrorAction Stop
            $currentCosmos.properties | Add-Member -NotePropertyName 'publicNetworkAccess' -NotePropertyValue 'Disabled' -Force
            $patchBody = @{
                location   = $currentCosmos.location
                properties = @{
                    publicNetworkAccess = 'Disabled'
                    locations           = $currentCosmos.properties.locations
                    databaseAccountOfferType = $currentCosmos.properties.databaseAccountOfferType
                }
            } | ConvertTo-Json -Depth 10
            Invoke-RestMethod -Uri $updateUri -Method PUT -Headers $headers -Body $patchBody -ErrorAction Stop | Out-Null
            Write-Log "  Cosmos DB: public access disabled. ✓" -Level SUCCESS
        } catch {
            Write-Log "  Cosmos DB: failed to disable public access: $($_.Exception.Message)" -Level WARN
            Write-Log "  You can disable manually: Update-AzCosmosDBAccount -Name '$cosmosName' -ResourceGroupName '$ResourceGroup' -PublicNetworkAccess 'Disabled'" -Level WARN
        }
    }

    # Key Vault
    if ($peTargets.ContainsKey('KeyVault')) {
        try {
            $kvName = $peTargets['KeyVault'].ResourceName
            Write-Log "  Key Vault '$kvName': disabling public access..." -Level INFO
            Update-AzKeyVault -VaultName $kvName -ResourceGroupName $ResourceGroup `
                -PublicNetworkAccess 'Disabled' -ErrorAction Stop | Out-Null
            Write-Log "  Key Vault: public access disabled. ✓" -Level SUCCESS
        } catch {
            Write-Log "  Key Vault: failed to disable public access: $($_.Exception.Message)" -Level WARN
            Write-Log "  You can disable manually: Update-AzKeyVault -VaultName '$kvName' -ResourceGroupName '$ResourceGroup' -PublicNetworkAccess 'Disabled'" -Level WARN
        }
    }

    # Storage Account
    if ($peTargets.ContainsKey('StorageAccount')) {
        try {
            $saName = $peTargets['StorageAccount'].ResourceName
            Write-Log "  Storage Account '$saName': disabling public access..." -Level INFO
            Set-AzStorageAccount -ResourceGroupName $ResourceGroup -Name $saName `
                -PublicNetworkAccess 'Disabled' -ErrorAction Stop | Out-Null
            Write-Log "  Storage Account: public access disabled. ✓" -Level SUCCESS
        } catch {
            Write-Log "  Storage Account: failed to disable public access: $($_.Exception.Message)" -Level WARN
            Write-Log "  You can disable manually: Set-AzStorageAccount -ResourceGroupName '$ResourceGroup' -Name '$saName' -PublicNetworkAccess 'Disabled'" -Level WARN
        }
    }

    # ── Summary ──
    Write-Log "" -Level INFO
    if ($peFailed -eq 0) {
        Write-Log "Private endpoints: $peCreated/$($peTargets.Count) created successfully." -Level SUCCESS
        $Script:Stats.PrivateEndpoints = "$peCreated created"
    } else {
        Write-Log "Private endpoints: $peCreated created, $peFailed failed." -Level WARN
        $Script:Stats.PrivateEndpoints = "$peCreated ok / $peFailed failed"
    }

    if ($RegisterPrivateDnsZones) {
        Write-Log "Private DNS zones created and linked to VNet '$VNetName'." -Level SUCCESS
    } else {
        Write-Log "Private DNS zones were NOT created (-RegisterPrivateDnsZones not specified)." -Level INFO
        Write-Log "Ensure your DNS infrastructure can resolve privatelink.* zones to the PE IPs." -Level INFO
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
#  PHASE 5 — SEED MANIFESTS
# ═══════════════════════════════════════════════════════════════════════════════

function Step-SeedManifests {
    Write-Log "PHASE 5: SEED PACKAGE MANIFESTS" -Level SECTION

    if (-not $ManifestPath) {
        Write-Log "No ManifestPath specified — skipping manifest ingestion." -Level INFO
        return
    }

    if ($Script:Stats.Deployment -eq 'WhatIf') {
        Write-Log "[WhatIf] Would seed manifests from '$ManifestPath'." -Level WARN
        return
    }

    # Find all manifest folders (folders containing .yaml files)
    $manifestFolders = Get-ChildItem -Path $ManifestPath -Directory -Recurse |
        Where-Object { (Get-ChildItem $_.FullName -Filter "*.yaml" -ErrorAction SilentlyContinue).Count -gt 0 }

    if (-not $manifestFolders -or $manifestFolders.Count -eq 0) {
        Write-Log "No manifest folders with .yaml files found in '$ManifestPath'." -Level WARN
        return
    }

    Write-Log "Found $($manifestFolders.Count) manifest folder(s) to process." -Level INFO

    $funcName   = "func-$Name"
    $maxRetries = 3          # per-manifest retry attempts
    $retryable  = @(503, 429, 500, 502, 504)   # HTTP codes worth retrying

    # ── Warm-up: probe the Function App before seeding ──
    # After a fresh deploy or redeploy the function host may still be cold-starting,
    # which causes 503 Service Unavailable on the first few requests. We wait here
    # so individual manifest uploads don't burn retry budget on cold-start delays.
    # Warm-up: use the APIM URL (/winget/information doesn't require auth)
    # to avoid false 401s from the function app's auth-key requirement.
    # Fall back to function app URL if APIM URL isn't known.
    # If -SkipAPIM was used, APIM is already removed — go direct.
    $warmupUrl = if ($Script:DeploymentUrl) {
        "$($Script:DeploymentUrl.TrimEnd('/'))/../information" -replace '/\.\./','/' # → .../winget/information
    } else { $null }
    # Simpler: just derive APIM information URL directly
    if ($SkipAPIM) {
        $apimInfoUrl = "https://func-$($Name -replace '[^a-zA-Z0-9-]', '').azurewebsites.net/api/information"
    } else {
        $apimInfoUrl = "https://apim-$($Name -replace '[^a-zA-Z0-9-]', '').azure-api.net/winget/information"
    }

    Write-Log "Warming up Function App '$funcName' before seeding..." -Level INFO
    $warmupMax = 60   # seconds (reduced — 401 = warm)
    $warmupOk  = $false
    $warmupSw  = [System.Diagnostics.Stopwatch]::StartNew()

    while ($warmupSw.Elapsed.TotalSeconds -lt $warmupMax) {
        try {
            $probe = Invoke-WebRequest -Uri $apimInfoUrl -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
            if ($probe.StatusCode -eq 200) {
                Write-Log "  Function App is warm (HTTP 200 via APIM in $([int]$warmupSw.Elapsed.TotalSeconds)s)." -Level SUCCESS
                $warmupOk = $true
                break
            }
        } catch {
            $status = $null
            if ($_.Exception.Response) { $status = [int]$_.Exception.Response.StatusCode }
            # 401/403 = function host is running (just needs auth) → treat as warm
            if ($status -in @(401, 403)) {
                Write-Log "  Function App is warm (HTTP $status in $([int]$warmupSw.Elapsed.TotalSeconds)s — auth required but host is up)." -Level SUCCESS
                $warmupOk = $true
                break
            }
            Write-Log "  Warm-up probe: HTTP $status — waiting... ($([int]$warmupSw.Elapsed.TotalSeconds)s)" -Level DEBUG
        }
        Start-Sleep -Seconds 10
    }

    if (-not $warmupOk) {
        Write-Log "  Function App did not respond after ${warmupMax}s — proceeding anyway (retries will handle transient errors)." -Level WARN
    }

    # ── First pass: try every manifest ──
    $failedManifests = [System.Collections.Generic.List[object]]::new()

    foreach ($folder in $manifestFolders) {
        $relativePath = $folder.FullName.Replace($ManifestPath, '').TrimStart('\', '/')
        try {
            Write-Log "  Adding: $relativePath" -Level INFO
            if ($PSCmdlet.ShouldProcess($folder.FullName, "Add-WinGetManifest")) {
                Add-WinGetManifest -FunctionName $funcName -Path $folder.FullName -ErrorAction Stop
                $Script:Stats.ManifestsLoaded++
                Write-Log "  OK: $relativePath" -Level SUCCESS
            }
        } catch {
            $errMsg = $_.Exception.Message
            # Determine if this is a retryable error (HTTP 5xx / 429)
            $isRetryable = $false
            foreach ($code in $retryable) {
                if ($errMsg -match "\b$code\b") { $isRetryable = $true; break }
            }
            if ($isRetryable) {
                Write-Log "  FAIL (retryable): $relativePath — $errMsg" -Level WARN
                $failedManifests.Add(@{ Folder = $folder; RelativePath = $relativePath; LastError = $errMsg })
            } else {
                # Non-retryable (e.g. 400 Bad Request, schema error) — count it now
                $Script:Stats.ManifestsFailed++
                Write-Log "  FAIL: $relativePath — $errMsg" -Level ERROR
            }
        }
    }

    # ── Retry passes for transient failures ──
    for ($retry = 1; $retry -le $maxRetries -and $failedManifests.Count -gt 0; $retry++) {
        $delay = 15 * [math]::Pow(2, $retry - 1)   # 15s, 30s, 60s
        Write-Log "Retry $retry/$maxRetries — $($failedManifests.Count) manifest(s) to retry in ${delay}s..." -Level WARN
        Start-Sleep -Seconds $delay

        $stillFailing = [System.Collections.Generic.List[object]]::new()
        foreach ($item in $failedManifests) {
            $folder       = $item.Folder
            $relativePath = $item.RelativePath
            try {
                Write-Log "  Retrying: $relativePath (attempt $($retry + 1))" -Level INFO
                Add-WinGetManifest -FunctionName $funcName -Path $folder.FullName -ErrorAction Stop
                $Script:Stats.ManifestsLoaded++
                Write-Log "  OK: $relativePath" -Level SUCCESS
            } catch {
                $item.LastError = $_.Exception.Message
                Write-Log "  FAIL: $relativePath — $($_.Exception.Message)" -Level WARN
                $stillFailing.Add($item)
            }
        }
        $failedManifests = $stillFailing
    }

    # ── Count any remaining failures ──
    if ($failedManifests.Count -gt 0) {
        foreach ($item in $failedManifests) {
            $Script:Stats.ManifestsFailed++
            Write-Log "  ABANDONED: $($item.RelativePath) — $($item.LastError)" -Level ERROR
        }
    }

    Write-Log "Manifest seeding complete: $($Script:Stats.ManifestsLoaded) loaded, $($Script:Stats.ManifestsFailed) failed." -Level INFO

    # Treat all-manifests-failed as a non-fatal warning (deployment itself succeeded)
    if ($Script:Stats.ManifestsFailed -gt 0 -and $Script:Stats.ManifestsLoaded -eq 0) {
        Write-Log "WARNING: All manifests failed to load. The REST source is deployed but empty." -Level WARN
        Write-Log "You can retry manifest seeding later with: Add-WinGetManifest -FunctionName '$funcName' -Path <folder>" -Level WARN
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
#  PHASE 6 — REGISTER SOURCE ON LOCAL MACHINE
# ═══════════════════════════════════════════════════════════════════════════════

function Step-RegisterSource {
    Write-Log "PHASE 6: REGISTER WINGET SOURCE" -Level SECTION

    if (-not $RegisterSource) {
        Write-Log "Source registration not requested (-RegisterSource not specified)." -Level INFO
        $Script:Stats.SourceRegistered = 'Skipped'
        return
    }

    if ($Script:Stats.Deployment -eq 'WhatIf') {
        Write-Log "[WhatIf] Would register source '$SourceDisplayName' → $Script:DeploymentUrl" -Level WARN
        $Script:Stats.SourceRegistered = 'WhatIf'
        return
    }

    $wingetCmd = Get-Command winget -ErrorAction SilentlyContinue
    if (-not $wingetCmd) {
        Write-Log "winget CLI not found — cannot register source." -Level ERROR
        $Script:Stats.SourceRegistered = 'Failed (no winget)'
        return
    }

    # Check if running as admin (required for winget source add)
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)

    # Check if source already exists
    $existingCheck = winget source list --name $SourceDisplayName 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Log "Source '$SourceDisplayName' already registered. Removing old entry..." -Level WARN
        if ($isAdmin) {
            winget source remove --name $SourceDisplayName 2>&1 | Out-Null
        } else {
            Start-Process pwsh -ArgumentList '-NoProfile', '-Command',
                "winget source remove --name '$SourceDisplayName'" -Verb RunAs -Wait -ErrorAction SilentlyContinue
        }
    }

    Write-Log "Registering source '$SourceDisplayName' → $Script:DeploymentUrl" -Level INFO
    if (-not $isAdmin) {
        Write-Log "Elevating to admin for winget source add..." -Level INFO
    }
    try {
        if ($isAdmin) {
            $output = winget source add `
                --name $SourceDisplayName `
                --arg $Script:DeploymentUrl `
                --type "Microsoft.Rest" `
                --accept-source-agreements 2>&1
            $exitCode = $LASTEXITCODE
        } else {
            $tmpFile = [System.IO.Path]::GetTempFileName()
            $escapedUrl = $Script:DeploymentUrl -replace "'", "''"
            Start-Process pwsh -ArgumentList '-NoProfile', '-Command',
                "winget source add --name '$SourceDisplayName' --arg '$escapedUrl' --type 'Microsoft.Rest' --accept-source-agreements 2>&1 | Out-File '$tmpFile'" `
                -Verb RunAs -Wait -ErrorAction Stop
            $output = Get-Content $tmpFile -ErrorAction SilentlyContinue
            Remove-Item $tmpFile -ErrorAction SilentlyContinue
            # Verify source was added
            $verifyCheck = winget source list --name $SourceDisplayName 2>&1
            $exitCode = $LASTEXITCODE
        }

        if ($exitCode -eq 0) {
            Write-Log "Source '$SourceDisplayName' registered successfully." -Level SUCCESS
            $Script:Stats.SourceRegistered = 'Registered'
        } else {
            Write-Log "winget source add returned exit code $LASTEXITCODE" -Level ERROR
            Write-Log "Output: $output" -Level ERROR
            $Script:Stats.SourceRegistered = 'Failed'
        }
    } catch {
        Write-Log "Failed to register source: $($_.Exception.Message)" -Level WARN
        Write-Log "" -Level INFO
        Write-Log "To register manually, run the following in an elevated (Admin) PowerShell:" -Level INFO
        Write-Log "  winget source add --name '$SourceDisplayName' --arg '$($Script:DeploymentUrl)' --type 'Microsoft.Rest' --accept-source-agreements" -Level INFO
        Write-Log "" -Level INFO
        $Script:Stats.SourceRegistered = 'Manual'
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN EXECUTION
# ═══════════════════════════════════════════════════════════════════════════════

try {
    Write-Banner

    Write-Log "Script parameters:" -Level DEBUG
    Write-Log "  Name             = $Name" -Level DEBUG
    Write-Log "  ResourceGroup    = $ResourceGroup" -Level DEBUG
    Write-Log "  Region           = $Region" -Level DEBUG
    Write-Log "  SubscriptionId   = $(if($SubscriptionId){$SubscriptionId}else{'(current context)'})" -Level DEBUG
    Write-Log "  PerformanceTier  = $PerformanceTier" -Level DEBUG
    Write-Log "  CosmosDBZoneHA   = $CosmosDBZoneRedundant" -Level DEBUG
    Write-Log "  CosmosDBRegion   = $(if($CosmosDBRegion){$CosmosDBRegion}else{'(same as Region)'})" -Level DEBUG
    Write-Log "  Authentication   = $Authentication" -Level DEBUG
    Write-Log "  EntraIdResource  = $(if($MicrosoftEntraIdResource){$MicrosoftEntraIdResource}else{'(auto: api://<Name>)'})" -Level DEBUG
    Write-Log "  ManifestPath     = $(if($ManifestPath){$ManifestPath}else{'(none)'})" -Level DEBUG
    Write-Log "  RegisterSource   = $RegisterSource" -Level DEBUG
    Write-Log "  SkipAPIM         = $SkipAPIM" -Level DEBUG
    Write-Log "  PrivateEndpoints = $EnablePrivateEndpoints" -Level DEBUG
    if ($EnablePrivateEndpoints) {
        Write-Log "  VNetName         = $VNetName" -Level DEBUG
        Write-Log "  SubnetName       = $SubnetName" -Level DEBUG
        Write-Log "  VNetResourceGroup= $(if($VNetResourceGroup){$VNetResourceGroup}else{'(same as ResourceGroup)'})" -Level DEBUG
        Write-Log "  RegisterDnsZones = $RegisterPrivateDnsZones" -Level DEBUG
    }
    Write-Log "  LogFile          = $Script:LogFile" -Level DEBUG
    Write-Log "" -Level DEBUG

    # Phase 1
    Step-ValidatePrerequisites

    # Phase 2
    Step-ConnectAzure

    # Phase 3
    Step-EnsureResourceGroup

    # Phase 3b — inventory existing resources, resolve soft-deletes
    Step-ResourceAudit

    # Phase 4
    Step-DeployRestSource

    # Phase 4b — verify function code is deployed
    Step-PostDeployHealthCheck

    # Phase 4c — create private endpoints (if requested)
    Step-CreatePrivateEndpoints

    # Phase 5
    Step-SeedManifests

    # Phase 6
    Step-RegisterSource

    # Final status — set partial failure exit code if there were non-fatal issues
    if ($Script:Stats.ManifestsFailed -gt 0 -or $Script:Stats.SourceRegistered -eq 'Failed') {
        $Script:ExitCode = 2   # Partial success
        Write-Log "Deployment completed with warnings (exit code 2)." -Level WARN
    } else {
        Write-Log "Deployment pipeline completed successfully." -Level SUCCESS
    }

} catch {
    $Script:ExitCode = 1
    Write-Log "FATAL: $($_.Exception.Message)" -Level ERROR
    Write-Log "Stack trace:`n$($_.ScriptStackTrace)" -Level ERROR

    # Mark any pending steps as failed
    foreach ($key in $Script:Stats.Keys) {
        if ($Script:Stats[$key] -eq 'Pending') {
            $Script:Stats[$key] = 'Not reached'
        }
    }
} finally {
    Write-Summary
    Write-Log "Full log saved to: $($Script:LogFile)" -Level INFO

    if ($Script:ExitCode -ne 0) {
        Write-Log "Deployment failed. Review the log file for details." -Level ERROR
    }

    exit $Script:ExitCode
}
