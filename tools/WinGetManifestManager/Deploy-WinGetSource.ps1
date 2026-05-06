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
    Resolution order when omitted:
      1. .\WinGet.RestSource.Functions.zip in the script directory (the fork's bundled build)
      2. <Microsoft.WinGet.RestSource module>\Data\WinGet.RestSource.Functions.zip (upstream)
    Pass this explicitly to deploy a custom or patched build.

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
    [string]$RestSourcePath,

    # When set, runs Phase 1 (prerequisites + module hygiene + upstream SKU probe)
    # then exits without touching Azure. Use to verify the deploy will not blow up
    # 30 minutes in due to a stale upstream module shadowing the patched fork.
    [switch]$DryRun
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

    # Write to log file (always) — defensive: OneDrive/AV/file-lock interference can
    # produce 'Stream was not readable' or 'used by another process' errors. Retry
    # twice with a short backoff via raw .NET I/O (which does NOT keep a long-lived
    # handle), then silently skip if still failing. Logging must NEVER abort the script.
    $writeOk = $false
    for ($i = 0; $i -lt 3 -and -not $writeOk; $i++) {
        try {
            [System.IO.File]::AppendAllText($Script:LogFile, ($logLine + [Environment]::NewLine), [System.Text.Encoding]::UTF8)
            $writeOk = $true
        } catch {
            if ($i -lt 2) { Start-Sleep -Milliseconds 100 }
        }
    }

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
    $title = 'WinGet REST Source — Azure Deployment'
    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $inner = 62  # interior width (between the side borders)
    $top    = '╔' + ('═' * $inner) + '╗'
    $bot    = '╚' + ('═' * $inner) + '╝'
    $padTit = ' ' * [math]::Max(0, [math]::Floor(($inner - $title.Length) / 2))
    $padStp = ' ' * [math]::Max(0, [math]::Floor(($inner - $stamp.Length) / 2))
    $line1  = '║' + $padTit + $title + (' ' * ($inner - $padTit.Length - $title.Length)) + '║'
    $line2  = '║' + $padStp + $stamp + (' ' * ($inner - $padStp.Length - $stamp.Length)) + '║'

    $banner = @"

    $top
    $line1
    $line2
    $bot

"@
    Write-Host $banner -ForegroundColor Cyan
    Add-Content -Path $Script:LogFile -Value $banner -Encoding UTF8 -WhatIf:$false -Confirm:$false
}

function Write-Summary {
    $elapsed = (Get-Date) - $Script:StartTime
    # Box layout: total interior width = 76 chars; label column = 22 chars;
    # value column = 52 chars. Long values are truncated with an ellipsis so the
    # right border always aligns regardless of input length.
    $valW = 53
    $fmt = {
        param($label, $value)
        $v = if ($null -eq $value) { '' } else { "$value" }
        if ($v.Length -gt $valW) { $v = $v.Substring(0, $valW - 1) + [char]0x2026 }
        ('{0}  {1,-18} : {2,-' + $valW + '}{3}') -f [char]0x2502, $label, $v, [char]0x2502
    }
    $h = [char]0x2500
    $v = [char]0x2502
    $tl = [char]0x250C; $tr = [char]0x2510
    $bl = [char]0x2514; $br = [char]0x2518
    $ml = [char]0x251C; $mr = [char]0x2524
    $bar = ($h.ToString() * 76)
    $top    = "$tl$bar$tr"
    $sep    = "$ml$bar$mr"
    $bot    = "$bl$bar$br"
    $titleText = 'DEPLOYMENT SUMMARY'
    $padLeftAmt = [int](([double]76 - $titleText.Length) / 2)
    $titleInner = (' ' * $padLeftAmt) + $titleText
    $titleInner = $titleInner.PadRight(76)
    $title  = "$v$titleInner$v"

    $cosmosHa  = if ($CosmosDBZoneRedundant) { 'Enabled' } else { 'Disabled' }
    $cosmosReg = if ($CosmosDBRegion) { $CosmosDBRegion } else { $Region }
    $restUrl   = if ($Script:DeploymentUrl) { $Script:DeploymentUrl } else { 'N/A' }

    $rows = @(
        $top
        $title
        $sep
        (& $fmt 'Source Name'       $Name)
        (& $fmt 'Resource Group'    $ResourceGroup)
        (& $fmt 'Region'            $Region)
        (& $fmt 'Performance'       $PerformanceTier)
        (& $fmt 'CosmosDB Zone-HA'  $cosmosHa)
        (& $fmt 'CosmosDB Region'   $cosmosReg)
        (& $fmt 'Authentication'    $Authentication)
        $sep
        (& $fmt 'Prereq Checks'     $Script:Stats.PrereqChecks)
        (& $fmt 'Azure Auth'        $Script:Stats.AzureAuth)
        (& $fmt 'Resource Group'    $Script:Stats.ResourceGroup)
        (& $fmt 'Resource Audit'    $Script:Stats.ResourceAudit)
        (& $fmt 'Deployment'        $Script:Stats.Deployment)
        (& $fmt 'Manifests OK'      $Script:Stats.ManifestsLoaded)
        (& $fmt 'Manifests Failed'  $Script:Stats.ManifestsFailed)
        (& $fmt 'Source Registered' $Script:Stats.SourceRegistered)
        $sep
        (& $fmt 'REST Source URL'   $restUrl)
        (& $fmt 'Elapsed Time'      ("{0}m {1}s" -f $elapsed.Minutes, $elapsed.Seconds))
        (& $fmt 'Log File'          (Split-Path $Script:LogFile -Leaf))
        (& $fmt 'Exit Code'         $Script:ExitCode)
        $bot
    )
    $summary = [Environment]::NewLine + ('    ' + ($rows -join ("`r`n    "))) + [Environment]::NewLine
    Write-Host $summary -ForegroundColor $(if ($Script:ExitCode -eq 0) { 'Green' } else { 'Red' })
    try {
        [System.IO.File]::AppendAllText($Script:LogFile, $summary, [System.Text.Encoding]::UTF8)
    } catch {
        # log-file write failures must not abort the script
    }
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

function Resolve-RestSourceFunctionsZip {
    <#
    .SYNOPSIS
        Resolves the path to WinGet.RestSource.Functions.zip with this precedence:
          1. -RestSourcePath script parameter (explicit override)
          2. <script-directory>\WinGet.RestSource.Functions.zip  (this fork's bundled build)
          3. <Microsoft.WinGet.RestSource module>\Data\WinGet.RestSource.Functions.zip (upstream fallback)
    .DESCRIPTION
        Customer feedback (2026-05): the script was silently using the upstream module's
        bundled zip instead of the fork's bundled zip in the script directory. The README
        already promised the script-directory fallback, so this restores that contract.
    .OUTPUTS
        [string] absolute path, or $null if no zip is found.
    #>
    [CmdletBinding()]
    param(
        [string]$ExplicitPath
    )

    if ($ExplicitPath) {
        $resolved = (Resolve-Path -LiteralPath $ExplicitPath -ErrorAction Stop).Path
        Write-Log "Functions zip: explicit -RestSourcePath → $resolved" -Level INFO
        return $resolved
    }

    $scriptDirZip = if ($PSScriptRoot) { Join-Path $PSScriptRoot 'WinGet.RestSource.Functions.zip' } else { $null }
    if ($scriptDirZip -and (Test-Path -LiteralPath $scriptDirZip)) {
        $sz = [math]::Round((Get-Item -LiteralPath $scriptDirZip).Length / 1MB, 2)
        Write-Log "Functions zip: using fork-bundled build → $scriptDirZip ($sz MB)" -Level INFO
        return $scriptDirZip
    }

    $mod = Get-Module Microsoft.WinGet.RestSource
    if ($mod) {
        $upstreamZip = Join-Path $mod.ModuleBase 'Data\WinGet.RestSource.Functions.zip'
        if (Test-Path -LiteralPath $upstreamZip) {
            Write-Log "Functions zip: falling back to upstream module copy → $upstreamZip" -Level WARN
            Write-Log "  (Drop a fork build at '$scriptDirZip' to override.)" -Level WARN
            return $upstreamZip
        }
    }

    Write-Log "Functions zip: NOT FOUND in any of: -RestSourcePath, '$($scriptDirZip ?? '<no script dir>')', upstream module Data folder." -Level ERROR
    return $null
}

function Assert-ModuleAvailable {
    <#
    .SYNOPSIS
        Ensures a module is installed AND that the highest available copy on disk is the
        one that will load. Detects the dual-edition split (WindowsPowerShell\Modules vs
        PowerShell\Modules), stale already-loaded instances, and reports every copy found
        across PSModulePath so customers can spot a stale fork copy that's losing the
        version race.
    .NOTES
        Customer feedback (2026-05): a stale Microsoft.WinGet.RestSource lived in
        $HOME\Documents\PowerShell\Modules\... with the unpatched ValidateSet and
        silently won the import. This helper now refuses to silently ignore that case.
    #>
    param(
        [string]$ModuleName,
        [switch]$Install,
        [version]$MinimumVersion
    )

    # 1) Enumerate ALL copies across PSModulePath (both editions, all roots)
    $allCopies = @(Get-Module -ListAvailable -Name $ModuleName -ErrorAction SilentlyContinue |
                   Sort-Object Version -Descending)

    if ($allCopies.Count -gt 1) {
        Write-Log "Found $($allCopies.Count) installed copies of '$ModuleName':" -Level WARN
        foreach ($c in $allCopies) {
            Write-Log "  v$($c.Version) at $($c.ModuleBase)" -Level WARN
        }
        Write-Log "  → PowerShell will load v$($allCopies[0].Version) (highest version wins)." -Level WARN
    }

    $mod = $allCopies | Select-Object -First 1

    if (-not $mod) {
        if ($Install) {
            Write-Log "Module '$ModuleName' not found — installing..." -Level WARN
            try {
                Install-PSResource -Name $ModuleName -Scope CurrentUser -TrustRepository -Reinstall -ErrorAction Stop
                Write-Log "Module '$ModuleName' installed successfully." -Level SUCCESS
            } catch {
                Install-Module -Name $ModuleName -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                Write-Log "Module '$ModuleName' installed (fallback Install-Module)." -Level SUCCESS
            }
            $mod = Get-Module -ListAvailable -Name $ModuleName | Sort-Object Version -Descending | Select-Object -First 1
        } else {
            throw "Required module '$ModuleName' is not installed. Run: Install-PSResource -Name $ModuleName"
        }
    } else {
        Write-Log "Module '$ModuleName' v$($mod.Version) found at $($mod.ModuleBase)" -Level DEBUG
    }

    if ($MinimumVersion -and $mod.Version -lt $MinimumVersion) {
        throw "Module '$ModuleName' v$($mod.Version) is older than required v$MinimumVersion. " +
              "Run: Install-PSResource -Name $ModuleName -Reinstall"
    }

    # 2) Force a clean re-import: a previous session/profile may have a stale instance loaded
    Get-Module -Name $ModuleName -All -ErrorAction SilentlyContinue | Remove-Module -Force -ErrorAction SilentlyContinue

    return $mod
}

function Test-UpstreamSkuSupport {
    <#
    .SYNOPSIS
        Probes the loaded Microsoft.WinGet.RestSource module to confirm the requested
        APIM SKU is in the upstream ValidateSet for ImplementationPerformance.
    .DESCRIPTION
        Customer feedback (2026-05): when the user passed -PerformanceTier StandardV2 on
        a machine where the unpatched upstream module was installed, our wrapper accepted
        the value but the inner `New-WinGetSource` ValidateSet rejected it with:
            "The argument 'StandardV2' does not belong to the set
             'Developer,Basic,Standard,Premium,Consumption' ..."
        This probe fails fast BEFORE we start any ARM work.
    .OUTPUTS
        [string[]] of supported values (informational), or throws if the requested
        SKU is not in the set.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$RequestedTier
    )

    $cmd = Get-Command -Module Microsoft.WinGet.RestSource -Name New-WinGetSource -ErrorAction SilentlyContinue
    if (-not $cmd) {
        Write-Log "Microsoft.WinGet.RestSource\New-WinGetSource not visible — skipping SKU probe." -Level WARN
        return @()
    }
    $param = $cmd.Parameters['ImplementationPerformance']
    if (-not $param) {
        Write-Log "Upstream New-WinGetSource has no -ImplementationPerformance parameter — skipping probe." -Level WARN
        return @()
    }
    $vsAttr = $param.Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] } | Select-Object -First 1
    if (-not $vsAttr) {
        Write-Log "Upstream New-WinGetSource -ImplementationPerformance has no ValidateSet — assuming open." -Level DEBUG
        return @()
    }

    $supported = @($vsAttr.ValidValues)
    Write-Log "Upstream module accepts ImplementationPerformance values: $($supported -join ', ')" -Level INFO

    if ($supported -notcontains $RequestedTier) {
        Write-Log "Requested SKU '$RequestedTier' missing from upstream ValidateSet — attempting in-memory hot-patch." -Level WARN

        # Customer feedback (2026-05 follow-up): instead of failing, try to widen the
        # ValidateSet so the deploy actually succeeds on a vanilla module install.
        # Strategy: walk the chain New-WinGetSource -> New-ARMObjects -> any inner
        # cmdlet that exposes -ImplementationPerformance / -Sku, and rebuild the
        # ValidValues list with our requested tier appended.
        $unionValues = @($supported + $RequestedTier | Select-Object -Unique)
        $patched = 0
        foreach ($candidateName in @('New-WinGetSource', 'New-ARMObjects', 'New-ARMParameterObject')) {
            $c = Get-Command -Module Microsoft.WinGet.RestSource -Name $candidateName -ErrorAction SilentlyContinue
            if (-not $c) { continue }
            foreach ($pName in @('ImplementationPerformance', 'Sku')) {
                $p = $c.Parameters[$pName]
                if (-not $p) { continue }
                $vs = $p.Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] } | Select-Object -First 1
                if (-not $vs) { continue }
                if (@($vs.ValidValues) -contains $RequestedTier) { continue }
                try {
                    # ValidValues is a get-only IList<string>; replace it via reflection on the backing field.
                    $field = [System.Management.Automation.ValidateSetAttribute].GetField('validValues', [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic)
                    if ($field) {
                        $field.SetValue($vs, [string[]]$unionValues)
                        Write-Log "  Patched $candidateName -$pName ValidateSet -> $($unionValues -join ', ')" -Level SUCCESS
                        $patched++
                    } else {
                        Write-Log "  Could not locate ValidateSetAttribute backing field on $candidateName -$pName (PowerShell version mismatch)." -Level WARN
                    }
                } catch {
                    Write-Log "  Failed to patch $candidateName -${pName}: $($_.Exception.Message)" -Level WARN
                }
            }
        }

        # Also patch the ARM template's allowedValues if the JSON template ships in the module
        # (some versions of the module hard-code allowedValues for the apimSku parameter).
        try {
            $modBase = (Get-Module -Name Microsoft.WinGet.RestSource).ModuleBase
            if ($modBase) {
                $armTemplates = Get-ChildItem -Path $modBase -Recurse -Filter '*.json' -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -match 'azuredeploy|template|apim' -or (Select-String -Path $_.FullName -Pattern 'apimSku|ImplementationPerformance' -SimpleMatch -Quiet -ErrorAction SilentlyContinue) }
                foreach ($t in $armTemplates) {
                    try {
                        $raw = [System.IO.File]::ReadAllText($t.FullName)
                        # Only touch files whose allowedValues array is missing our requested tier.
                        if ($raw -match '"allowedValues"\s*:\s*\[[^\]]*\]' -and $raw -notmatch [regex]::Escape("`"$RequestedTier`"")) {
                            $patchedRaw = [regex]::Replace($raw, '("allowedValues"\s*:\s*\[)([^\]]*)(\])', {
                                param($m)
                                $arr = $m.Groups[2].Value
                                if ($arr -match [regex]::Escape("`"$RequestedTier`"")) { return $m.Value }
                                # Heuristic: only patch arrays that already contain SKU-like values
                                if ($arr -notmatch '"(Developer|Basic|Standard|Premium|Consumption)"') { return $m.Value }
                                $newArr = $arr.TrimEnd(", `n`r`t".ToCharArray()) + ", `"$RequestedTier`""
                                return $m.Groups[1].Value + $newArr + $m.Groups[3].Value
                            })
                            if ($patchedRaw -ne $raw) {
                                # Backup once, then overwrite in place (in module install dir).
                                if (-not (Test-Path "$($t.FullName).orig")) { Copy-Item $t.FullName "$($t.FullName).orig" -Force }
                                [System.IO.File]::WriteAllText($t.FullName, $patchedRaw)
                                Write-Log "  Patched ARM template allowedValues in $($t.Name) (backup: .orig)" -Level SUCCESS
                                $patched++
                            }
                        }
                    } catch {
                        Write-Log "  Skipped ARM template $($t.Name): $($_.Exception.Message)" -Level DEBUG
                    }
                }
            }
        } catch {
            Write-Log "ARM-template scan failed: $($_.Exception.Message)" -Level DEBUG
        }

        if ($patched -eq 0) {
            $msg = @"
Requested -PerformanceTier '$RequestedTier' is NOT supported by the installed
Microsoft.WinGet.RestSource module and the in-memory hot-patch could not find
any ValidateSet to widen. Upstream accepts: $($supported -join ', ')

Resolution options:
  1. Install the patched fork to a HIGHER version than the upstream copy:
        Install-PSResource -Name Microsoft.WinGet.RestSource -Repository <fork-feed> -Reinstall
  2. Manually edit the installed module's New-WinGetSource.ps1 ValidateSet.
  3. Pick a supported tier (e.g. -PerformanceTier Basic) and resize APIM after deploy:
        Update-AzApiManagement -ResourceGroupName '$ResourceGroup' -Name 'apim-$Name' -Sku $RequestedTier
"@
            throw $msg
        }

        Write-Log "Hot-patch applied to $patched location(s); upstream module will now accept '$RequestedTier'." -Level SUCCESS
    }
    return $supported
}

function Publish-FunctionZipOneDeploy {
    <#
    .SYNOPSIS
        Uploads a Functions zip via the Kudu OneDeploy REST endpoint.
    .DESCRIPTION
        Customer feedback (2026-05): `Publish-AzWebApp -ArchivePath` fails on
        Flex Consumption / .NET-isolated worker plans because it routes through the
        legacy MSDeploy pipeline. OneDeploy (`/api/publish?type=zip`) works for
        Consumption, Premium, Flex, AND isolated worker plans.

        Uses bearer auth from the current Az context — no Kudu basic-auth credentials
        required.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ResourceGroup,
        [Parameter(Mandatory)][string]$FunctionAppName,
        [Parameter(Mandatory)][string]$ZipPath,
        [int]$TimeoutSec = 600
    )

    if (-not (Test-Path $ZipPath)) {
        throw "Zip not found: $ZipPath"
    }
    $zipSize = [math]::Round((Get-Item $ZipPath).Length / 1MB, 2)
    Write-Log "OneDeploy: uploading $ZipPath ($zipSize MB) to '$FunctionAppName'..." -Level INFO

    $token = Get-PlainToken
    $uri = "https://$FunctionAppName.scm.azurewebsites.net/api/publish?type=zip&async=false&restart=true&clean=true"

    $headers = @{
        Authorization  = "Bearer $token"
        'Content-Type' = 'application/zip'
    }

    try {
        $resp = Invoke-WebRequest -Method POST -Uri $uri -Headers $headers `
                                   -InFile $ZipPath -TimeoutSec $TimeoutSec `
                                   -UseBasicParsing -ErrorAction Stop
        Write-Log "OneDeploy: HTTP $($resp.StatusCode) — upload accepted." -Level SUCCESS
        return $true
    } catch {
        $status = if ($_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { 0 }
        Write-Log "OneDeploy upload failed (HTTP $status): $($_.Exception.Message)" -Level ERROR
        throw
    }
}

function Assert-FunctionAppHealthy {
    <#
    .SYNOPSIS
        Validates that the Function App has the right runtime app-settings AND that the
        host is actually responsive after a deploy.
    .DESCRIPTION
        Customer feedback (2026-05): even after a successful zip upload via the portal,
        the host failed with an "isolated worker" startup error. That error surfaces
        when FUNCTIONS_WORKER_RUNTIME isn't 'dotnet-isolated' or when the Functions
        runtime extension version is wrong. We assert both, fix them if needed, and
        then ping a known endpoint to fail fast on host startup errors.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ResourceGroup,
        [Parameter(Mandatory)][string]$FunctionAppName,
        [string]$HealthPath = '/api/information',
        [int]$WarmupSec    = 180,
        [int]$ProbeTimeoutSec = 30
    )

    Write-Log "Verifying Function App '$FunctionAppName' app settings..." -Level INFO
    try {
        $app = Get-AzFunctionApp -ResourceGroupName $ResourceGroup -Name $FunctionAppName -ErrorAction Stop
    } catch {
        try {
            $app = Get-AzWebApp -ResourceGroupName $ResourceGroup -Name $FunctionAppName -ErrorAction Stop
        } catch {
            Write-Log "Could not fetch Function App '$FunctionAppName': $($_.Exception.Message)" -Level WARN
            return $false
        }
    }

    $required = @{
        'FUNCTIONS_WORKER_RUNTIME'   = 'dotnet-isolated'
        'FUNCTIONS_EXTENSION_VERSION'= '~4'
    }
    $current = @{}
    if ($app.ApplicationSettings) {
        foreach ($k in $app.ApplicationSettings.Keys) { $current[$k] = $app.ApplicationSettings[$k] }
    } elseif ($app.SiteConfig -and $app.SiteConfig.AppSettings) {
        foreach ($s in $app.SiteConfig.AppSettings) { $current[$s.Name] = $s.Value }
    }

    $patch = @{}
    foreach ($k in $required.Keys) {
        if ($current[$k] -ne $required[$k]) {
            Write-Log "  $k = '$($current[$k])' ≠ expected '$($required[$k])' — will patch." -Level WARN
            $patch[$k] = $required[$k]
        } else {
            Write-Log "  $k = '$($current[$k])' OK" -Level DEBUG
        }
    }

    if ($patch.Count -gt 0) {
        try {
            Update-AzFunctionAppSetting -ResourceGroupName $ResourceGroup -Name $FunctionAppName -AppSetting $patch -Force -ErrorAction Stop | Out-Null
            Write-Log "App settings patched: $($patch.Keys -join ', ')" -Level SUCCESS
            Restart-AzFunctionApp -ResourceGroupName $ResourceGroup -Name $FunctionAppName -Force -ErrorAction Stop
            Write-Log "Function App restarted to apply settings." -Level INFO
        } catch {
            Write-Log "Failed to patch app settings: $($_.Exception.Message)" -Level WARN
        }
    }

    Write-Log "Warming up host (waiting up to ${WarmupSec}s for startup, ${ProbeTimeoutSec}s per probe)..." -Level INFO
    $url = "https://$FunctionAppName.azurewebsites.net$HealthPath"
    $deadline = (Get-Date).AddSeconds($WarmupSec)
    $lastErr = $null
    while ((Get-Date) -lt $deadline) {
        try {
            $r = Invoke-WebRequest -Uri $url -Method GET -TimeoutSec $ProbeTimeoutSec -UseBasicParsing -ErrorAction Stop
            if ($r.StatusCode -lt 500) {
                Write-Log "Function host responded: HTTP $($r.StatusCode) at $HealthPath" -Level SUCCESS
                return $true
            }
            $lastErr = "HTTP $($r.StatusCode)"
        } catch {
            $msg = $_.Exception.Message
            # 401/403 from /api/information means the HTTP stack is fully up and the
            # host is routing requests — it just refuses us because we don't send a
            # function-host key. That's "healthy" for our purposes; treat as success.
            if ($msg -match '\b(401|403)\b|Unauthorized|Forbidden') {
                Write-Log "Function host responded: HTTP 401/403 (auth-protected) — host is up." -Level SUCCESS
                return $true
            }
            $lastErr = $msg
        }
        Start-Sleep -Seconds 5
    }
    Write-Log "Function host did NOT become healthy within ${WarmupSec}s. Last error: $lastErr" -Level WARN
    Write-Log "Hint: check Kudu eventlog and confirm the .NET-isolated worker is wired correctly." -Level WARN
    return $false
}

function Watch-DeploymentProgress {
    <#
    .SYNOPSIS Polls ARM deployment and resource status while a blocking call runs.
    .DESCRIPTION
        Starts a background polling loop (ThreadJob) that prints live status
        every $IntervalSec seconds. Call Stop-Job on the returned job once the
        blocking operation finishes.

        Customer feedback (2026-05): during the 30-40 min APIM creation the
        previous version sat silent for the first 30 s and then only showed
        coarse resource states. This version emits an immediate heartbeat,
        surfaces the active ARM deployment OPERATION (which resource ARM is
        currently working on), and shows elapsed / ETA so users know the
        deploy is alive.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$StepName,
        [Parameter(Mandatory)][string]$RG,
        [int]$IntervalSec = 20,
        # Estimated total minutes for ETA % display. Caller passes 35 for
        # fresh APIM, 10 for APIM-reuse paths.
        [int]$EstimateMinutes = 35
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

    $job = Start-ThreadJob -ArgumentList $StepName, $RG, $IntervalSec, $typeOrder, $Script:LogFile, $pollSubId, $pollToken, $EstimateMinutes -ScriptBlock {
        param($StepName, $RG, $IntervalSec, $typeOrder, $LogFile, $SubId, $Token, $EstimateMinutes)

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

        $sw         = [System.Diagnostics.Stopwatch]::StartNew()
        $lastHash   = ''
        $lastOpHash = ''
        $headers    = @{ Authorization = "Bearer $Token"; 'Content-Type' = 'application/json' }
        $firstTick  = $true

        while ($true) {
            # Initial heartbeat after 5s so users see the poller is alive immediately;
            # subsequent ticks use the configured interval.
            if ($firstTick) { Start-Sleep -Seconds 5; $firstTick = $false }
            else            { Start-Sleep -Seconds $IntervalSec }
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
                $running = @($depResp.value)

                # Drill into each running deployment's operations to surface the ACTIVE
                # sub-resource ARM is currently working on. Without this, ARM shows the
                # outer deployment as 'Running' for the entire 30-40 min APIM wait.
                $activeOps = @()
                foreach ($dep in $running) {
                    try {
                        $opsUri = "https://management.azure.com/subscriptions/$SubId/resourceGroups/$RG/providers/Microsoft.Resources/deployments/$($dep.name)/operations?api-version=2024-03-01"
                        $opsResp = Invoke-RestMethod -Uri $opsUri -Headers $headers -ErrorAction SilentlyContinue
                        foreach ($op in $opsResp.value) {
                            $opState = $op.properties.provisioningState
                            if ($opState -in 'Running','Creating','Accepted') {
                                $tgt = $op.properties.targetResource
                                if ($tgt -and $tgt.resourceType) {
                                    $shortType = ($tgt.resourceType -split '/')[-1]
                                    $activeOps += [PSCustomObject]@{
                                        Type  = $shortType
                                        Name  = $tgt.resourceName
                                        State = $opState
                                    }
                                }
                            }
                        }
                    } catch { }
                }

                # Build hash to detect changes (resources + active ops)
                $resLines = @($resources | ForEach-Object { "$($_.Name)|$($_.State)" })
                $opLines  = @($activeOps | ForEach-Object { "$($_.Type)|$($_.Name)|$($_.State)" })
                $hash     = ($resLines -join ';')
                $opHash   = ($opLines -join ';')

                $elapsed_str = "$([int]$elapsed.TotalMinutes)m$($elapsed.Seconds.ToString('00'))s"
                $pct = if ($EstimateMinutes -gt 0) { [math]::Min(99, [int](($elapsed.TotalMinutes / $EstimateMinutes) * 100)) } else { 0 }
                $progress = "${elapsed_str} / ~${EstimateMinutes}m (${pct}%)"

                if ($hash -ne $lastHash) {
                    Poll-Write "[$StepName] $progress — Resource status:" 'INFO'
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
                    $lastOpHash = ''  # force op block to print on next change too
                } else {
                    Poll-Write "[$StepName] $progress — still deploying..." 'INFO'
                }

                # Always re-emit the active operation block when it changes, even if
                # outer resource hash is unchanged (APIM creation reports many sub-ops).
                if ($activeOps.Count -gt 0 -and $opHash -ne $lastOpHash) {
                    Poll-Write "  ▶ Currently provisioning:" 'INFO'
                    foreach ($op in $activeOps) {
                        $tag = if ($op.Name) { "$($op.Type)/$($op.Name)" } else { $op.Type }
                        Poll-Write "      ⧖ $($tag.PadRight(48)) $($op.State)"
                    }
                    $lastOpHash = $opHash
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

    # 1d. WinGet REST source module — prefer the bundled fork shipped alongside this script
    # over any PSGallery copy. Customer feedback (2026-05): an upstream copy installed in
    # $HOME\Documents\(Windows)PowerShell\Modules\ silently shadowed the fork (because
    # Import-Module by name picks the highest version on PSModulePath, and the fork is
    # versioned 0.1.0). Importing by ABSOLUTE PATH bypasses name resolution entirely.
    $bundledModulePsd1 = Join-Path $PSScriptRoot 'Modules\Microsoft.WinGet.RestSource\Microsoft.WinGet.RestSource.psd1'
    if (Test-Path $bundledModulePsd1) {
        Write-Log "Using bundled fork module: $bundledModulePsd1" -Level INFO
        # Force-remove any already-loaded copy so Import-Module by path actually swaps it.
        Get-Module -Name 'Microsoft.WinGet.RestSource' -All -ErrorAction SilentlyContinue |
            Remove-Module -Force -ErrorAction SilentlyContinue
        Import-Module $bundledModulePsd1 -Force -ErrorAction Stop
        $loadedRest = Get-Module Microsoft.WinGet.RestSource
        Write-Log "Bundled fork imported: v$($loadedRest.Version) from $($loadedRest.ModuleBase)" -Level SUCCESS
    } else {
        Write-Log "Bundled fork not found at '$bundledModulePsd1' — falling back to PSGallery resolution." -Level WARN
        Write-Log "Checking Microsoft.WinGet.RestSource module..." -Level INFO
        Assert-ModuleAvailable -ModuleName 'Microsoft.WinGet.RestSource' -Install

        # 1e. Import the module
        Write-Log "Importing Microsoft.WinGet.RestSource..." -Level INFO
        Import-Module Microsoft.WinGet.RestSource -Force -ErrorAction Stop
        $loadedRest = Get-Module Microsoft.WinGet.RestSource
        Write-Log "Module imported: v$($loadedRest.Version) from $($loadedRest.ModuleBase)" -Level SUCCESS
    }

    # 1e-2. Verify the loaded module accepts the requested APIM SKU BEFORE doing
    # any ARM work. With the bundled fork this is a no-op (its ValidateSet already
    # contains BasicV2/StandardV2); with a PSGallery fallback it auto-hotpatches.
    Test-UpstreamSkuSupport -RequestedTier $PerformanceTier | Out-Null

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
        Write-Log "Tenant: $($ctx.Tenant.Id)" -Level DEBUG
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

    # Verify the cached context can actually access ARM in the current subscription's tenant.
    # Customer feedback (2026-05): a cached context from a different home tenant produces
    # 'credentials have not been set up or have expired' on the next ARM call (e.g.
    # New-AzResourceGroup) instead of failing here. Detect + recover gracefully.
    try {
        # A read-only call against the target subscription's tenant. If silent re-auth fails,
        # this throws with the same multi-tenant message we'd otherwise see in Phase 3.
        Get-AzResourceGroup -ErrorAction Stop -Name '__nonexistent__do_not_create__' -WarningAction SilentlyContinue | Out-Null
    } catch {
        $msg = $_.Exception.Message
        $isAuthGap = $msg -match 'credentials have not been set up|have expired|User interaction is required|conditional access|AADSTS|Authentication failed against tenant'
        # 'NotFound' on the bogus RG name is the SUCCESS signal — ARM reached.
        # Az/ARM has used several wordings over the years for the same 404; match all.
        if ($msg -notmatch 'NotFound|ResourceGroupNotFound|could not be found|does not exist|ResourceNotFound') {
            if ($isAuthGap) {
                $tenantId = $ctx.Tenant.Id
                if ($msg -match 'tenant\s+([0-9a-fA-F-]{36})') { $tenantId = $Matches[1] }
                Write-Log "Cached Azure context cannot access ARM in tenant '$tenantId' — silent re-auth failed." -Level WARN
                Write-Log "This typically happens when your cached login is for a different home tenant or MFA is required." -Level WARN
                Write-Log "Re-authenticating interactively (Connect-AzAccount -TenantId $tenantId)..." -Level INFO
                try {
                    Connect-AzAccount -TenantId $tenantId -ErrorAction Stop | Out-Null
                    if ($SubscriptionId) {
                        Set-AzContext -SubscriptionId $SubscriptionId -ErrorAction Stop | Out-Null
                    }
                    $ctx = Get-AzContext
                    Write-Log "Re-authenticated as: $($ctx.Account.Id) (tenant $($ctx.Tenant.Id))" -Level SUCCESS
                } catch {
                    throw "Interactive re-authentication failed: $($_.Exception.Message). " +
                          "Run manually: Connect-AzAccount -TenantId $tenantId" +
                          $(if ($SubscriptionId) { " ; Set-AzContext -SubscriptionId $SubscriptionId" } else { '' })
                }
            } else {
                throw
            }
        }
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

                # Detect SKU mismatch — ARM cannot perform certain cross-tier upgrades in-place
                # (notably classic Developer/Basic/Standard ↔ V2 SKUs). Customer-confirmed failure
                # mode: ARM returns 'Failed to connect to Management endpoint Port 3443 ... for the
                # Developer SKU service' for ~30 min then ValidationError. Abort early with a clear
                # remediation message instead of letting ARM retry-loop.
                $existingSku  = "$($apimDetail.Sku.Name)"
                $requestedSku = "$PerformanceTier"
                $skuMap = @{
                    'Developer'  = 'Developer'
                    'Basic'      = 'Basic'
                    'Enhanced'   = 'Standard'
                    'BasicV2'    = 'Basicv2'
                    'StandardV2' = 'Standardv2'
                }
                $expectedSku = $skuMap[$requestedSku]
                if ($expectedSku -and ($existingSku -ne $expectedSku)) {
                    $isClassicToV2 = ($existingSku -notmatch '(?i)v2$') -and ($expectedSku -match '(?i)v2$')
                    $isV2ToClassic = ($existingSku -match '(?i)v2$') -and ($expectedSku -notmatch '(?i)v2$')
                    if ($isClassicToV2 -or $isV2ToClassic) {
                        Write-Log "" -Level ERROR
                        Write-Log "═══════════════════════════════════════════════════════════════════════════" -Level ERROR
                        Write-Log "  APIM SKU MISMATCH — IN-PLACE UPGRADE NOT SUPPORTED" -Level ERROR
                        Write-Log "═══════════════════════════════════════════════════════════════════════════" -Level ERROR
                        Write-Log "  Existing APIM '$($apimRes.Name)' is SKU='$existingSku'." -Level ERROR
                        Write-Log "  Requested -PerformanceTier '$requestedSku' would deploy SKU='$expectedSku'." -Level ERROR
                        Write-Log "  Azure does not support in-place migration between classic and V2 APIM SKUs." -Level ERROR
                        Write-Log "  ARM would retry the deploy for ~30 minutes before failing." -Level ERROR
                        Write-Log "" -Level ERROR
                        Write-Log "  Resolution — delete the existing APIM first, then re-run this script:" -Level ERROR
                        Write-Log "    Remove-AzApiManagement -ResourceGroupName '$ResourceGroup' -Name '$($apimRes.Name)'" -Level ERROR
                        Write-Log "  (Deletion takes ~10-15 min. The script will then create a fresh '$expectedSku' APIM.)" -Level ERROR
                        Write-Log "  Alternatively re-run with -PerformanceTier matching the existing SKU." -Level ERROR
                        Write-Log "═══════════════════════════════════════════════════════════════════════════" -Level ERROR
                        throw "APIM SKU mismatch (existing='$existingSku', requested='$expectedSku') — see error block above."
                    } else {
                        Write-Log "  APIM SKU change detected ($existingSku → $expectedSku). ARM will attempt in-place update." -Level WARN
                    }
                }
            } catch {
                # Re-throw fatal SKU mismatch; otherwise log and continue (cmdlet may fail on V2 SKUs)
                if ("$($_.Exception.Message)" -match 'SKU MISMATCH|SKU mismatch') { throw }
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
                $null = Invoke-RestMethod -Uri $purgeUri -Method DELETE -Headers $headers -ErrorAction Stop
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

    # ── 2c. App Configuration soft-delete ──
    # Customer-confirmed failure: ARM 'appconfig' deployment fails with NameUnavailable
    # ('already in use by a soft-deleted configuration store') when a previous teardown
    # left a soft-deleted store. App Configuration soft-delete defaults to 7-day retention
    # and the name is reserved for the full window unless explicitly purged.
    $appConfigName = if ($env:WINGETMM_OVERRIDE_APPCONFIG) { $env:WINGETMM_OVERRIDE_APPCONFIG } else { "appcs-$cleanName" }
    if (-not $Script:ExistingResources.AppConfig) {
        Write-Log "Checking for soft-deleted App Configuration store '$appConfigName'..." -Level INFO
        try {
            # Listing endpoint returns all soft-deleted stores in the subscription
            $appCsDelUri = "https://management.azure.com/subscriptions/$subId/providers/Microsoft.AppConfiguration/deletedConfigurationStores?api-version=2023-03-01"
            $appCsDel = Invoke-RestMethod -Uri $appCsDelUri -Headers $headers -ErrorAction Stop
            $appCsMatch = $appCsDel.value | Where-Object { $_.name -eq $appConfigName }
            if ($appCsMatch) {
                # Listing payload has location under properties.location (not top-level)
                $appCsLocation = $appCsMatch.properties.location
                if (-not $appCsLocation) {
                    # Fallback: parse from the resource id (.../locations/<loc>/deletedConfigurationStores/...)
                    if ($appCsMatch.id -match '/locations/([^/]+)/deletedConfigurationStores/') { $appCsLocation = $Matches[1] }
                }
                Write-Log "  Found soft-deleted App Configuration '$appConfigName' in '$appCsLocation' (scheduled purge: $($appCsMatch.properties.scheduledPurgeDate)). Purging to free the name..." -Level WARN
                $purgeUri = "https://management.azure.com/subscriptions/$subId/providers/Microsoft.AppConfiguration/locations/$appCsLocation/deletedConfigurationStores/$appConfigName/purge?api-version=2023-03-01"
                $null = Invoke-RestMethod -Uri $purgeUri -Method POST -Headers $headers -ErrorAction Stop
                Write-Log "  Waiting for App Configuration purge to complete..." -Level INFO
                $purgeWait = 0
                do {
                    Start-Sleep -Seconds 10
                    $purgeWait += 10
                    $tok     = Get-PlainToken
                    $headers = @{ Authorization = "Bearer $tok"; 'Content-Type' = 'application/json' }
                    $recheck = Invoke-RestMethod -Uri $appCsDelUri -Headers $headers -ErrorAction Stop
                    $still = $recheck.value | Where-Object { $_.name -eq $appConfigName }
                } while ($still -and $purgeWait -lt 120)

                if ($still) {
                    Write-Log "  App Configuration purge still in progress after ${purgeWait}s — deployment will retry if needed." -Level WARN
                } else {
                    Write-Log "  App Configuration '$appConfigName' purged successfully." -Level SUCCESS
                    $resolved++
                }
            } else {
                Write-Log "  No soft-deleted App Configuration found. OK" -Level DEBUG
            }
        } catch {
            Write-Log "  App Configuration soft-delete check failed: $($_.Exception.Message)" -Level WARN
        }
    } else {
        Write-Log "App Configuration '$appConfigName' exists in RG — skipping soft-delete check." -Level DEBUG
    }

    # ── 2d. Function App (informational only — cannot be explicitly purged) ──
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
            # Always pass an explicit zip path — prefer fork-bundled build over upstream module copy.
            $resolvedZip = Resolve-RestSourceFunctionsZip -ExplicitPath $RestSourcePath
            if ($resolvedZip) {
                $deployParams['RestSourcePath'] = $resolvedZip
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

            # Start background status poller (ETA: 35 min fresh APIM, 10 min if reused)
            $pollEtaMin = if ($Script:ExistingResources.APIM) { 10 } else { 35 }
            $pollerJob = Watch-DeploymentProgress -StepName 'New-WinGetSource' -RG $ResourceGroup -EstimateMinutes $pollEtaMin
            try {
                $deployResult = Invoke-StepWithRetry -StepName "New-WinGetSource" -MaxRetries 2 -BaseDelaySec 30 -Action {
                    New-WinGetSource @deployParams
                }
            } finally {
                Stop-Job $pollerJob -ErrorAction SilentlyContinue
                Remove-Job $pollerJob -Force -ErrorAction SilentlyContinue
            }

            # Try to derive APIM URL for source registration
            $apimName = "apim-$($Name -replace '[^a-zA-Z0-9-]', '')"
            $apimUrl = $null
            try {
                $apimUrl = (Get-AzApiManagement -Name $apimName -ResourceGroupName $ResourceGroup -ErrorAction Stop).RuntimeUrl
            } catch {
                # V2 SKU — Get-AzApiManagement cmdlet enum can't deserialize. Fall back to ARM REST.
                try {
                    $subIdApim = (Get-AzContext).Subscription.Id
                    $apimRestUri = "/subscriptions/$subIdApim/resourceGroups/$ResourceGroup/providers/Microsoft.ApiManagement/service/$apimName`?api-version=2023-05-01-preview"
                    $apimRest = Invoke-AzRestMethod -Path $apimRestUri -Method GET -ErrorAction Stop
                    if ($apimRest.StatusCode -ge 200 -and $apimRest.StatusCode -lt 300) {
                        $apimUrl = ($apimRest.Content | ConvertFrom-Json -Depth 20).properties.gatewayUrl
                    }
                } catch {}
            }
            if ($apimUrl) {
                $Script:DeploymentUrl = "$apimUrl/winget/"
            } else {
                Write-Log "Could not retrieve APIM URL via cmdlet or REST, will use direct function URL." -Level WARN
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
            $restSourceZip    = Resolve-RestSourceFunctionsZip -ExplicitPath $RestSourcePath
            if (-not $restSourceZip) { throw 'Could not locate WinGet.RestSource.Functions.zip — pass -RestSourcePath explicitly.' }
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

                # 3c: Deduplicate locations[] by locationName.
                # The upstream template ships with two entries (failoverPriority 0 and 1)
                # that both default to the primary region — and our region override above
                # collapses them to the same name. Cosmos DB rejects this with
                # "Provided list of regions contains duplicate regions" (BadRequest).
                # Keep only the first occurrence of each region and renumber failoverPriority.
                $uniqueLocs = @()
                $seen = @{}
                foreach ($loc in $cosmosJson.Parameters.locations.value) {
                    $key = "$($loc.locationName)".Trim().ToLowerInvariant()
                    if (-not $seen.ContainsKey($key)) {
                        $seen[$key] = $true
                        $uniqueLocs += $loc
                    }
                }
                if ($uniqueLocs.Count -lt @($cosmosJson.Parameters.locations.value).Count) {
                    Write-Log "  Deduplicated locations[]: $(@($cosmosJson.Parameters.locations.value).Count) → $($uniqueLocs.Count) entry(ies)" -Level INFO
                    for ($i = 0; $i -lt $uniqueLocs.Count; $i++) {
                        $uniqueLocs[$i].failoverPriority = $i
                    }
                    $cosmosJson.Parameters.locations.value = @($uniqueLocs)
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

            # Start background status poller (ETA: 35 min fresh APIM, 10 min if reused)
            $pollEtaMin = if ($Script:ExistingResources.APIM) { 10 } else { 35 }
            $pollerJob = Watch-DeploymentProgress -StepName 'ARM-Deploy' -RG $ResourceGroup -EstimateMinutes $pollEtaMin
            try {
                $deployOk = $false
                $apimNameForRecovery = $ARMObjects.Where({$_.ObjectType -eq 'ApiManagement'}).Parameters.Parameters.serviceName.value
                # Derive KV name same way Step-ResourceAudit does, for the KV access-policy
                # race recovery branch below.
                $cleanNameForKv = $Name -replace '[^a-zA-Z0-9-]', ''
                $kvName = if ($env:WINGETMM_OVERRIDE_KEYVAULT) { $env:WINGETMM_OVERRIDE_KEYVAULT } else { "kv-$cleanNameForKv" }
                for ($attempt = 1; $attempt -le 5; $attempt++) {
                    # Per-attempt banner so the user sees "we're alive, this is attempt N" — the
                    # underlying ARM call can be silent for 5-15 minutes during APIM creation.
                    $attemptStart = Get-Date
                    Write-Log ("▶ ARM deploy attempt {0}/5 starting (typical: {1})..." -f $attempt, $(if($Script:ExistingResources.APIM){'5-15 min'}else{'10-20 min per attempt'})) -Level INFO

                    # Wrap the inner call in try/catch + ErrorAction Stop so cmdlet/script errors
                    # bubble up as a single catchable exception instead of being auto-rendered
                    # by PowerShell as a giant red block with line numbers + stack frames.
                    $deployErr = $null
                    try {
                        $deployOk = & $mod { param($a,$z,$rg) New-ARMObjects -ARMObjects ([ref]$a) -RestSourcePath $z -ResourceGroup $rg -ErrorAction Stop } $ARMObjects $restSourceZip $ResourceGroup
                    } catch {
                        $deployErr = $_.Exception.Message
                        $deployOk = $false
                    }
                    $attemptDur = ((Get-Date) - $attemptStart).ToString('mm\:ss')
                    if ($deployOk) {
                        Write-Log ("✓ ARM deploy attempt {0}/5 succeeded after {1}." -f $attempt, $attemptDur) -Level SUCCESS
                        break
                    }
                    # Always log a failure line — New-ARMObjects sometimes returns $false silently
                    # without throwing (e.g. when its own internal $DeployError was captured but
                    # the function chose to return-false instead of throw).
                    if ($deployErr) {
                        $cleanErr = $deployErr
                        if ($cleanErr -match 'Status Message:\s*(.+?)(?:\s*Please provide correlationId|\s*\(Code:|\r|\n|$)') {
                            $cleanErr = $Matches[1].Trim()
                        }
                        $errCode = if ($deployErr -match '\(Code:([^)]+)\)') { $Matches[1] } else { 'DeploymentFailed' }
                        Write-Log ("✗ ARM deploy attempt {0}/5 failed after {1}: [{2}] {3}" -f $attempt, $attemptDur, $errCode, $cleanErr) -Level WARN
                    } else {
                        Write-Log ("✗ ARM deploy attempt {0}/5 failed after {1} (no exception surfaced — inner cmdlet returned `$false; check log for ARM error details)." -f $attempt, $attemptDur) -Level WARN
                    }
                    if ($attempt -lt 5) {
                        # KV access-policy race recovery: ARM resolves named-value KV references
                        # immediately after APIM creation, but the freshly-issued APIM identity
                        # often hasn't propagated to KV yet. The error reads:
                        #   "does not have secrets get permission on key vault '<name>;location=...'"
                        # If APIM exists with an identity, grant 'Get' on the vault NOW so the next
                        # attempt can resolve named values immediately. Without this, we waste an
                        # entire attempt waiting for propagation.
                        $kvGranted = $false
                        if ($deployErr -and $deployErr -match 'does not have secrets get permission' -and $apimNameForRecovery) {
                            try {
                                $subIdKv = (Get-AzContext).Subscription.Id
                                $apimUriKv = "/subscriptions/$subIdKv/resourceGroups/$ResourceGroup/providers/Microsoft.ApiManagement/service/$apimNameForRecovery`?api-version=2023-05-01-preview"
                                $apimGetKv = Invoke-AzRestMethod -Path $apimUriKv -Method GET -ErrorAction SilentlyContinue
                                if ($apimGetKv -and $apimGetKv.StatusCode -ge 200 -and $apimGetKv.StatusCode -lt 300) {
                                    $apimObjKv = $apimGetKv.Content | ConvertFrom-Json -Depth 20
                                    $principalId = $apimObjKv.identity.principalId
                                    if ($principalId) {
                                        Write-Log "↻ KV access-policy race detected — granting 'Get/secrets' to APIM identity ($principalId) on '$kvName' before retry..." -Level WARN
                                        $null = Set-AzKeyVaultAccessPolicy -VaultName $kvName -ResourceGroupName $ResourceGroup -ObjectId $principalId -PermissionsToSecrets Get -BypassObjectIdValidation -ErrorAction SilentlyContinue -ErrorVariable kvErr
                                        if ($kvErr) {
                                            Write-Log "KV access-policy grant returned: $kvErr" -Level WARN
                                        } else {
                                            Write-Log "✓ KV access policy granted — next attempt should resolve named values successfully." -Level SUCCESS
                                            $kvGranted = $true
                                        }
                                    } else {
                                        Write-Log "APIM exists but has no system-assigned identity yet — cannot pre-grant KV; falling back to retry." -Level DEBUG
                                    }
                                }
                            } catch {
                                Write-Log "KV pre-grant check error: $($_.Exception.Message)" -Level DEBUG
                            }
                        }

                        # Check if APIM is in Failed provisioning state — if so it MUST be deleted
                        # before any retry can succeed (ARM returns ServiceInFailedProvisioningState
                        # 'Please delete the service before trying to re-create it with same name').
                        # Common after a transient ActivationFailed in the first attempt.
                        $apimRecovered = $false
                        if ($apimNameForRecovery) {
                            try {
                                $subIdRecov = (Get-AzContext).Subscription.Id
                                $apimUriRecov = "/subscriptions/$subIdRecov/resourceGroups/$ResourceGroup/providers/Microsoft.ApiManagement/service/$apimNameForRecovery`?api-version=2023-05-01-preview"
                                $apimGetRecov = Invoke-AzRestMethod -Path $apimUriRecov -Method GET -ErrorAction SilentlyContinue
                                if ($apimGetRecov -and $apimGetRecov.StatusCode -ge 200 -and $apimGetRecov.StatusCode -lt 300) {
                                    $apimObjRecov = $apimGetRecov.Content | ConvertFrom-Json -Depth 20
                                    $stateRecov = "$($apimObjRecov.properties.provisioningState)"
                                    if ($stateRecov -in @('Failed','Cancelled','Canceled')) {
                                        Write-Log "↻ APIM '$apimNameForRecovery' in '$stateRecov' state — deleting + purging soft-delete shadow before retry..." -Level WARN
                                        $delResp = Invoke-AzRestMethod -Path $apimUriRecov -Method DELETE -ErrorAction SilentlyContinue
                                        if ($delResp -and $delResp.StatusCode -ge 200 -and $delResp.StatusCode -lt 400) {
                                            # Poll for delete completion (typically 5-10 min for APIM).
                                            # Log progress every minute so the user sees we're alive.
                                            $delStart    = Get-Date
                                            $delDeadline = $delStart.AddMinutes(15)
                                            $lastLog     = $delStart
                                            do {
                                                Start-Sleep -Seconds 30
                                                $check = Invoke-AzRestMethod -Path $apimUriRecov -Method GET -ErrorAction SilentlyContinue
                                                if (-not $check -or $check.StatusCode -eq 404) { $apimRecovered = $true; break }
                                                # Surface progress every ~60s so the user sees we're alive.
                                                if (((Get-Date) - $lastLog).TotalSeconds -ge 60) {
                                                    $delEl = ((Get-Date) - $delStart).ToString('mm\:ss')
                                                    Write-Log ("  … still waiting for APIM delete ({0} elapsed, typical: 5-10 min)..." -f $delEl) -Level INFO
                                                    $lastLog = Get-Date
                                                }
                                            } while ((Get-Date) -lt $delDeadline)
                                            $delTotal = ((Get-Date) - $delStart).ToString('mm\:ss')
                                            if ($apimRecovered) {
                                                # Purge soft-deleted shadow so the next deploy can reuse the name
                                                $purgeUri = "/subscriptions/$subIdRecov/providers/Microsoft.ApiManagement/locations/$Region/deletedservices/$apimNameForRecovery`?api-version=2023-05-01-preview"
                                                $null = Invoke-AzRestMethod -Path $purgeUri -Method DELETE -ErrorAction SilentlyContinue
                                                Start-Sleep -Seconds 30
                                                Write-Log ("✓ APIM '{0}' recovered after {1} — next attempt starts with a fresh service name." -f $apimNameForRecovery, $delTotal) -Level SUCCESS
                                            } else {
                                                Write-Log ("APIM delete did not complete within 15 min ({0} elapsed) — next attempt may still hit ServiceInFailedProvisioningState." -f $delTotal) -Level WARN
                                            }
                                        } else {
                                            Write-Log "APIM DELETE request returned status $($delResp.StatusCode): $($delResp.Content)" -Level WARN
                                        }
                                    }
                                }
                            } catch {
                                Write-Log "Failed-APIM recovery check error: $($_.Exception.Message)" -Level DEBUG
                            }
                        }
                        $waitSec   = if ($apimRecovered -or $kvGranted) { 5 } else { 15 }
                        $reasonTxt = @()
                        if ($kvGranted)     { $reasonTxt += 'KV access policy pre-granted' }
                        if ($apimRecovered) { $reasonTxt += 'failed APIM purged' }
                        $reason = if ($reasonTxt) { ' (' + ($reasonTxt -join ', ') + ')' } else { ' (no recovery applied — transient retry)' }
                        Write-Log ("Retrying ARM deploy in {0}s — attempt {1}/5{2}..." -f $waitSec, ($attempt + 1), $reason) -Level INFO
                        Start-Sleep -Seconds $waitSec
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
            $apiMgmtUrl = $null
            try {
                $apiMgmtUrl = (Get-AzApiManagement -Name $apiMgmtName -ResourceGroupName $ResourceGroup -ErrorAction Stop).RuntimeUrl
            } catch {
                # Get-AzApiManagement cmdlet can't deserialize V2 SKUs back to its enum
                # ('Error mapping types. String -> PsApiManagementSku'). Fall back to ARM REST.
                Write-Log "Get-AzApiManagement cmdlet failed (likely V2 SKU enum issue) — falling back to ARM REST..." -Level DEBUG
                try {
                    $subIdApim = (Get-AzContext).Subscription.Id
                    $apimRestUri = "/subscriptions/$subIdApim/resourceGroups/$ResourceGroup/providers/Microsoft.ApiManagement/service/$apiMgmtName`?api-version=2023-05-01-preview"
                    $apimRest = Invoke-AzRestMethod -Path $apimRestUri -Method GET -ErrorAction Stop
                    if ($apimRest.StatusCode -ge 200 -and $apimRest.StatusCode -lt 300) {
                        $apimRestObj = $apimRest.Content | ConvertFrom-Json -Depth 20
                        $apiMgmtUrl = $apimRestObj.properties.gatewayUrl
                    }
                } catch {
                    Write-Log "ARM REST fallback for APIM URL also failed: $($_.Exception.Message)" -Level DEBUG
                }
            }
            if ($apiMgmtUrl) {
                $Script:DeploymentUrl = "$apiMgmtUrl/winget/"
                Write-Log "Connection command:" -Level SUCCESS
                Write-Log "  winget source add -n `"$Name`" -a `"$Script:DeploymentUrl`" -t `"Microsoft.Rest`"" -Level INFO
            } else {
                Write-Log "Could not retrieve API Management URL via cmdlet or REST — falling back to direct function URL." -Level WARN
                $Script:DeploymentUrl = "https://func-${Name}.azurewebsites.net/api/"
            }
        }

        # Derive the REST source URL (prefer APIM URL set above, else direct function URL)
        if (-not $Script:DeploymentUrl) {
            $Script:DeploymentUrl = "https://func-${Name}.azurewebsites.net/api/"
        }
        Write-Log "REST Source URL: $Script:DeploymentUrl" -Level SUCCESS

        # ── POST-DEPLOY: Verify Function App health + isolated-worker settings ──
        # Customer feedback (2026-05): host failed with "isolated worker" error
        # after upstream zip upload silently misconfigured FUNCTIONS_WORKER_RUNTIME.
        $funcAppName = "func-$($Name -replace '[^a-zA-Z0-9-]', '')"
        try {
            $healthy = Assert-FunctionAppHealthy -ResourceGroup $ResourceGroup -FunctionAppName $funcAppName
            if (-not $healthy) {
                Write-Log "Function App '$funcAppName' is not yet healthy — attempting OneDeploy fallback re-publish..." -Level WARN
                $zipForRepublish = Resolve-RestSourceFunctionsZip -ExplicitPath $RestSourcePath
                if ($zipForRepublish -and (Test-Path $zipForRepublish)) {
                    try {
                        Invoke-StepWithRetry -StepName 'OneDeploy-Republish' -MaxRetries 2 -BaseDelaySec 20 -Action {
                            Publish-FunctionZipOneDeploy -ResourceGroup $ResourceGroup -FunctionAppName $funcAppName -ZipPath $zipForRepublish
                        } | Out-Null
                        Assert-FunctionAppHealthy -ResourceGroup $ResourceGroup -FunctionAppName $funcAppName | Out-Null
                    } catch {
                        Write-Log "OneDeploy fallback also failed: $($_.Exception.Message)" -Level WARN
                    }
                } else {
                    Write-Log "Functions zip not found for re-publish: $zipForRepublish" -Level WARN
                }
            }
        } catch {
            Write-Log "Function App health check skipped: $($_.Exception.Message)" -Level WARN
        }

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

    $zipPath = Resolve-RestSourceFunctionsZip -ExplicitPath $RestSourcePath
    if (-not $zipPath -or -not (Test-Path $zipPath)) {
        Write-Log "  Function zip not found — cannot recover." -Level ERROR
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

    # Sanity check: confirm where the script is running from and whether the
    # bundled fork module is reachable on disk. This is the first thing logged
    # so future deploy logs make it trivial to confirm the bundled fork (vs a
    # PSGallery copy) is what will be loaded later in Phase 1.
    $bundledPsd1Check = Join-Path $PSScriptRoot 'Modules\Microsoft.WinGet.RestSource\Microsoft.WinGet.RestSource.psd1'
    $bundledArmCheck  = Join-Path $PSScriptRoot 'Modules\Microsoft.WinGet.RestSource\Data\ARMTemplates\azurefunction.json'
    $bundledDllCheck  = Join-Path $PSScriptRoot 'Modules\Microsoft.WinGet.RestSource\Library\WinGet.RestSource.PowershellSupport\Microsoft.WinGet.PowershellSupport.dll'
    Write-Log "Script directory  : $PSScriptRoot" -Level INFO
    Write-Log "Bundled fork psd1 : $bundledPsd1Check (exists=$(Test-Path $bundledPsd1Check))" -Level INFO
    Write-Log "Bundled ARM templ : $bundledArmCheck (exists=$(Test-Path $bundledArmCheck))" -Level INFO
    Write-Log "Bundled support   : $bundledDllCheck (exists=$(Test-Path $bundledDllCheck))" -Level INFO
    if (-not (Test-Path $bundledPsd1Check)) {
        Write-Log "Bundled fork module is missing — deploy will fall back to PSGallery resolution and the SKU hot-patch path." -Level WARN
    }

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

    if ($DryRun) {
        Write-Log "" -Level INFO
        Write-Log "DRY RUN COMPLETE" -Level SECTION
        Write-Log "All prerequisite checks passed. Module hygiene + upstream SKU probe OK." -Level SUCCESS
        Write-Log "Re-run without -DryRun to perform the actual Azure deployment." -Level INFO
        $Script:ExitCode = 0
        Write-Summary
        exit 0
    }

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
    $errMsg = $_.Exception.Message

    # Detect multi-tenant / expired-credential errors and surface an actionable hint
    # instead of a stack trace dump (customer feedback 2026-05).
    $isAuthGap = $errMsg -match 'credentials have not been set up|have expired|User interaction is required|conditional access|AADSTS|Authentication failed against tenant|Connect-AzAccount'
    $tenantHint = $null
    if ($errMsg -match 'tenant\s+([0-9a-fA-F-]{36})') { $tenantHint = $Matches[1] }
    elseif ($errMsg -match "-TenantId\s+'?([0-9a-fA-F-]{36})") { $tenantHint = $Matches[1] }

    if ($isAuthGap) {
        Write-Log "AZURE AUTHENTICATION REQUIRED" -Level SECTION
        Write-Log "The deployment could not access Azure Resource Manager with your current session." -Level ERROR
        Write-Log "Underlying error: $errMsg" -Level DEBUG
        Write-Log "" -Level INFO
        Write-Log "To resolve, run the following in this PowerShell session and re-launch the deployment:" -Level INFO
        if ($tenantHint) {
            Write-Log "    Connect-AzAccount -TenantId $tenantHint" -Level INFO
        } else {
            Write-Log "    Connect-AzAccount -TenantId <your-tenant-guid>" -Level INFO
        }
        if ($SubscriptionId) {
            Write-Log "    Set-AzContext -SubscriptionId $SubscriptionId" -Level INFO
        }
        Write-Log "" -Level INFO
        Write-Log "Tip: pass -SubscriptionId to this script to ensure the correct subscription/tenant is targeted." -Level INFO
    } else {
        Write-Log "FATAL: $errMsg" -Level ERROR
        Write-Log "Stack trace:`n$($_.ScriptStackTrace)" -Level ERROR
    }

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
