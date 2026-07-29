#Requires -Version 5.1
<#
.SYNOPSIS
    AVD Discovery - Automated Azure Virtual Desktop environment assessment.
.DESCRIPTION
    Standalone discovery script that connects to Azure, discovers all AVD
    resources, runs automated checks against CAF/WAF/LZA best practices,
    and exports a portable JSON file for import into AVD Assessor GUI.
    Can be run independently without the GUI tool.
.PARAMETER SubscriptionId
    Azure subscription ID(s) to assess. Accepts a single ID or array.
    If omitted, uses current Az context subscription.
.PARAMETER OutputPath
    Path to save the discovery JSON file. Defaults to
    AvdAssessor\assessments\discovery_<timestamp>.json
.PARAMETER SkipLogin
    Skip interactive login and use existing Az context.
.PARAMETER IncludeGuestChecks
    Opt-in in-guest FSLogix inspection. When set, the script runs a single consolidated
    PowerShell script (via Invoke-AzVMRunCommand) against up to 3 representative RUNNING
    session hosts per host pool to read HKLM\SOFTWARE\FSLogix\Profiles and the FSLogix agent
    version. Requires running VMs and the Microsoft.Compute/virtualMachines/runCommand action.
    Skipped by default (adds runtime + needs elevated permissions); the guest-derived checks
    (PROF-001/008/009/012/013/014/015) report N/A when this switch is absent.
.EXAMPLE
    .\Invoke-AvdDiscovery.ps1
    .\Invoke-AvdDiscovery.ps1 -SubscriptionId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
    .\Invoke-AvdDiscovery.ps1 -SubscriptionId @("sub1","sub2") -OutputPath "C:\temp\discovery.json"
    .\Invoke-AvdDiscovery.ps1 -IncludeGuestChecks
.NOTES
    Author : Anton Romanyuk
    Version: 0.6.0
    Date   : 2026-07-18
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string[]]$SubscriptionId,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [switch]$SkipLogin,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeGuestChecks
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Strip OneDrive module path to prevent old Az version conflicts
$env:PSModulePath = ($env:PSModulePath -split ';' |
    Where-Object { $_ -notlike '*OneDrive*' }) -join ';'

$ScriptRoot = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($ScriptRoot)) { $ScriptRoot = $PWD.Path }

$ScriptVersion = '0.6.0'

# ═══════════════════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════════════════

<#
.SYNOPSIS
    Writes a formatted status line to the console with icon, timestamp, and level-based coloring.
.DESCRIPTION
    Outputs colored status messages with level-specific icons. SECTION level renders a boxed header.
    Other levels (INFO, WARN, ERROR, SUCCESS, CHECK) render inline with timestamp.
.PARAMETER Message
    The status message text.
.PARAMETER Level
    Message level: INFO, WARN, ERROR, SUCCESS, CHECK, or SECTION.
#>
function Write-Status {
    param([string]$Message, [string]$Level = 'INFO')
    $ts = (Get-Date).ToString('HH:mm:ss')
    $Icon = switch ($Level) {
        'ERROR'   { '✗' }
        'WARN'    { '⚠' }
        'SUCCESS' { '✓' }
        'CHECK'   { '►' }
        'SECTION' { '─' }
        default   { '·' }
    }
    $Color = switch ($Level) {
        'ERROR'   { 'Red' }
        'WARN'    { 'Yellow' }
        'SUCCESS' { 'Green' }
        'CHECK'   { 'Cyan' }
        'SECTION' { 'DarkCyan' }
        default   { 'Gray' }
    }
    if ($Level -eq 'SECTION') {
        Write-Host ""
        Write-Host "  ┌── " -NoNewline -ForegroundColor DarkCyan
        Write-Host $Message -NoNewline -ForegroundColor Cyan
        Write-Host " $('─' * [math]::Max(1, 48 - $Message.Length))┐" -ForegroundColor DarkCyan
    } else {
        Write-Host "  " -NoNewline
        Write-Host $Icon -NoNewline -ForegroundColor $Color
        Write-Host " " -NoNewline
        Write-Host "[$ts] " -NoNewline -ForegroundColor DarkGray
        Write-Host $Message -ForegroundColor $(if ($Level -eq 'INFO') { 'White' } else { $Color })
    }
}

<#
.SYNOPSIS
    Writes a formatted key-value metric line for discovery summary output.
.PARAMETER Label
    The metric label text (padded to 20 chars).
.PARAMETER Value
    The numeric metric value.
.PARAMETER Icon
    Box-drawing character prefix. Defaults to vertical bar.
#>
function Write-Metric {
    param([string]$Label, [int]$Value, [string]$Icon = '│')
    Write-Host "  $Icon  " -NoNewline -ForegroundColor DarkCyan
    Write-Host $Label.PadRight(20) -NoNewline -ForegroundColor Gray
    Write-Host $Value -ForegroundColor White
}

<#
.SYNOPSIS
    Creates a standardized check result object for the discovery output.
.DESCRIPTION
    Constructs a PSCustomObject representing one automated check result with ID, category,
    status, severity, details, remediation recommendation, reference URL, and optional evidence.
    These objects are collected into the AllChecks ArrayList and exported in the discovery JSON.
.PARAMETER Id
    Check identifier matching checks.json (e.g. 'NET-001', 'SH-003').
.PARAMETER Category
    Assessment category (e.g. 'Networking', 'Session Hosts').
.PARAMETER Name
    Human-readable check name.
.PARAMETER Description
    What the check evaluates.
.PARAMETER Status
    Result: Pass, Fail, Warning, N/A, or Error.
.PARAMETER Severity
    Impact level: Critical, High, Medium, or Low.
.PARAMETER Details
    Detailed findings text.
.PARAMETER Recommendation
    Remediation guidance.
.PARAMETER Reference
    URL to documentation.
.PARAMETER Evidence
    Optional object containing supporting data.
.OUTPUTS
    PSCustomObject with all check result fields plus Timestamp and Source='Automated'.
#>
function New-CheckResult {
    param(
        [string]$Id,
        [string]$Category,
        [string]$Name,
        [string]$Description,
        [ValidateSet('Pass','Fail','Warning','N/A','Error')][string]$Status,
        [string]$Severity = 'Medium',
        [string]$Details = '',
        [string]$Recommendation = '',
        [string]$Reference = '',
        [object]$Evidence = $null
    )
    return [PSCustomObject]@{
        Id             = $Id
        Category       = $Category
        Name           = $Name
        Description    = $Description
        Status         = $Status
        Severity       = $Severity
        Details        = $Details
        Recommendation = $Recommendation
        Reference      = $Reference
        Evidence       = $Evidence
        Timestamp      = (Get-Date -Format 'o')
        Source         = 'Automated'
    }
}

<#
.SYNOPSIS
    Resolves vCPU count and memory (GB) for a VM size, caching one Get-AzVMSize call per region.
.DESCRIPTION
    Uses a small built-in lookup for common AVD families, falling back to Get-AzVMSize (cached per
    location in the supplied hashtable). Returns a PSCustomObject with VCPU and MemoryGB, or nulls
    when the size cannot be resolved.
.PARAMETER VMSize
    The VM size string (e.g. 'Standard_D4s_v5').
.PARAMETER Location
    Azure region of the VM (used for the Get-AzVMSize fallback).
.PARAMETER Cache
    Hashtable used to cache per-location Get-AzVMSize results across calls.
#>
function Get-VMSizeSpec {
    param([string]$VMSize, [string]$Location, [hashtable]$Cache)
    if (-not $VMSize) { return [PSCustomObject]@{ VCPU = $null; MemoryGB = $null } }
    # Small built-in lookup for the most common AVD families (avoids an API call in the common case).
    $Builtin = @{
        'Standard_D2s_v5' = @(2,8);   'Standard_D4s_v5' = @(4,16);  'Standard_D8s_v5' = @(8,32);  'Standard_D16s_v5' = @(16,64)
        'Standard_D2s_v4' = @(2,8);   'Standard_D4s_v4' = @(4,16);  'Standard_D8s_v4' = @(8,32);  'Standard_D16s_v4' = @(16,64)
        'Standard_D2s_v3' = @(2,8);   'Standard_D4s_v3' = @(4,16);  'Standard_D8s_v3' = @(8,32);  'Standard_D16s_v3' = @(16,64)
        'Standard_E4s_v5' = @(4,32);  'Standard_E8s_v5' = @(8,64);  'Standard_E16s_v5' = @(16,128)
        'Standard_E4s_v4' = @(4,32);  'Standard_E8s_v4' = @(8,64);  'Standard_E16s_v4' = @(16,128)
        'Standard_B2ms'   = @(2,8);   'Standard_B4ms'   = @(4,16);  'Standard_B8ms'   = @(8,32);  'Standard_B2s' = @(2,4)
        'Standard_D2as_v5'= @(2,8);   'Standard_D4as_v5'= @(4,16);  'Standard_D8as_v5'= @(8,32);  'Standard_D16as_v5' = @(16,64)
    }
    if ($Builtin.ContainsKey($VMSize)) {
        return [PSCustomObject]@{ VCPU = $Builtin[$VMSize][0]; MemoryGB = $Builtin[$VMSize][1] }
    }
    if ($Location -and $Cache -ne $null) {
        if (-not $Cache.ContainsKey($Location)) {
            $Cache[$Location] = @{}
            try {
                foreach ($S in @(Get-AzVMSize -Location $Location -ErrorAction Stop)) {
                    $Cache[$Location][$S.Name] = [PSCustomObject]@{ VCPU = $S.NumberOfCores; MemoryGB = [math]::Round($S.MemoryInMB / 1024, 0) }
                }
            } catch { }
        }
        if ($Cache[$Location].ContainsKey($VMSize)) { return $Cache[$Location][$VMSize] }
    }
    return [PSCustomObject]@{ VCPU = $null; MemoryGB = $null }
}

<#
.SYNOPSIS
    Tests whether an NSG security rule targets a given destination port, honoring singular ranges,
    the plural DestinationPortRanges collection, wildcard '*', and "start-end" ranges (audit C-1).
.PARAMETER Rule
    An NSG security rule object.
.PARAMETER Port
    The destination port number to test for (e.g. 3389).
#>
function Test-NsgRulePort {
    param($Rule, [int]$Port)
    $Candidates = @()
    if ($null -ne $Rule.DestinationPortRange)  { $Candidates += @($Rule.DestinationPortRange) }
    if ($null -ne $Rule.DestinationPortRanges) { $Candidates += @($Rule.DestinationPortRanges) }
    foreach ($C in $Candidates) {
        $Val = "$C".Trim()
        if ($Val -eq '*') { return $true }
        if ($Val -eq "$Port") { return $true }
        if ($Val -match '^(\d+)-(\d+)$') {
            if ($Port -ge [int]$Matches[1] -and $Port -le [int]$Matches[2]) { return $true }
        }
    }
    return $false
}

<#
.SYNOPSIS
    Returns true when an NSG rule's source is the public internet ('*', 'Internet', or 0.0.0.0/0).
#>
function Test-NsgInternetSource {
    param($Rule)
    return ($Rule.SourceAddressPrefix -eq '*' -or $Rule.SourceAddressPrefix -eq 'Internet' -or $Rule.SourceAddressPrefix -eq '0.0.0.0/0')
}

<#
.SYNOPSIS
    Acquires a Microsoft Graph bearer token from the existing Az login, returning a plain string.
.DESCRIPTION
    Uses Get-AzAccessToken -ResourceUrl https://graph.microsoft.com. Handles both the legacy plain-string
    .Token (Az.Accounts 2.x) and the SecureString .Token returned by Az.Accounts 5.x. Returns $null when
    a token cannot be obtained (caller degrades the identity checks to Status 'Error').
#>
function Get-GraphTokenString {
    try {
        $Tok = Get-AzAccessToken -ResourceUrl 'https://graph.microsoft.com' -ErrorAction Stop
    } catch {
        return $null
    }
    if (-not $Tok -or -not $Tok.Token) { return $null }
    if ($Tok.Token -is [System.Security.SecureString]) {
        try {
            $Bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Tok.Token)
            try { return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($Bstr) }
            finally { [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($Bstr) }
        } catch { return $null }
    }
    return "$($Tok.Token)"
}

<#
.SYNOPSIS
    Performs a paged Microsoft Graph GET, following @odata.nextLink, returning all .value items.
.PARAMETER Uri
    The initial Graph REST URI.
.PARAMETER Token
    A plain-string bearer token (from Get-GraphTokenString).
.OUTPUTS
    Array of result objects (the flattened .value collections across all pages). Throws on HTTP error
    so callers can distinguish permission failures (403) from empty results.
#>
function Invoke-GraphGet {
    param([string]$Uri, [string]$Token)
    $Headers = @{ Authorization = "Bearer $Token"; 'Content-Type' = 'application/json' }
    $Results = @()
    $Next = $Uri
    while ($Next) {
        $Resp = Invoke-RestMethod -Uri $Next -Headers $Headers -Method GET -ErrorAction Stop
        if ($null -ne $Resp.value) { $Results += @($Resp.value) } else { $Results += @($Resp) }
        $Next = $Resp.'@odata.nextLink'
    }
    return $Results
}

<#
.SYNOPSIS
    Runs a Log Analytics KQL query against a workspace given by its ARM resource ID.
.DESCRIPTION
    Optional dependency: Az.OperationalInsights (Invoke-AzOperationalInsightsQuery expects the
    workspace *customer ID* GUID, not the ARM resource ID, so the ARM ID from diagnostic settings
    is first resolved to its CustomerId via Get-AzOperationalInsightsWorkspace). Never throws:
    returns a hashtable @{ Ok=<bool>; Error=<string>; Rows=@(...) }. When the module is missing,
    Ok is $false and Error is 'Az.OperationalInsights not installed' so callers emit Status 'Error'.
.PARAMETER WorkspaceResourceId
    ARM resource ID of the Log Analytics workspace (as harvested from diagnostic settings).
.PARAMETER Query
    The KQL query text.
.PARAMETER TimespanDays
    Query window in days (default 7).
#>
function Invoke-AvdLaQuery {
    param([string]$WorkspaceResourceId, [string]$Query, [int]$TimespanDays = 7)
    if (-not (Get-Module -ListAvailable -Name Az.OperationalInsights -ErrorAction SilentlyContinue)) {
        return @{ Ok = $false; Error = 'Az.OperationalInsights not installed'; Rows = @() }
    }
    try {
        $Parts = $WorkspaceResourceId -split '/'
        $RgIdx = -1
        for ($i = 0; $i -lt $Parts.Length; $i++) { if ($Parts[$i] -ieq 'resourceGroups') { $RgIdx = $i; break } }
        if ($RgIdx -lt 0) { return @{ Ok = $false; Error = "Could not parse workspace resource ID: $WorkspaceResourceId"; Rows = @() } }
        $WsRg   = $Parts[$RgIdx + 1]
        $WsName = $Parts[-1]
        $Ws = Get-AzOperationalInsightsWorkspace -ResourceGroupName $WsRg -Name $WsName -ErrorAction Stop
        $CustId = if ($Ws.CustomerId -and $Ws.CustomerId.Guid) { $Ws.CustomerId.Guid } else { "$($Ws.CustomerId)" }
        if (-not $CustId) { return @{ Ok = $false; Error = "Workspace $WsName has no CustomerId"; Rows = @() } }
        $Res = Invoke-AzOperationalInsightsQuery -WorkspaceId $CustId -Query $Query -Timespan (New-TimeSpan -Days $TimespanDays) -ErrorAction Stop
        return @{ Ok = $true; Error = $null; Rows = @($Res.Results) }
    } catch {
        return @{ Ok = $false; Error = $_.Exception.Message; Rows = @() }
    }
}

# ═══════════════════════════════════════════════════════════════════════════
# PREREQUISITE CHECK
# ═══════════════════════════════════════════════════════════════════════════

Write-Host ""
Write-Host "      ___  _   ______     ___                                   " -ForegroundColor Cyan
Write-Host "     /   || | / / __ \   /   |  ___ ___  ___  ___ ___  ___  ____" -ForegroundColor Cyan
Write-Host "    / /| || |/ / / / /  / /| | / __/ __// -_)/ __/ __// _ \/ __/" -ForegroundColor Cyan
Write-Host "   / ___ ||   / /_/ /  / ___ |/__//__/ \__//__//__/ \___/_/   " -ForegroundColor Cyan
Write-Host "  /_/  |_|_/\_\____/  /_/  |_|                                " -ForegroundColor Cyan
Write-Host ""
Write-Host "  v$ScriptVersion" -NoNewline -ForegroundColor DarkGray
Write-Host "  ·  " -NoNewline -ForegroundColor DarkGray
Write-Host "CAF" -NoNewline -ForegroundColor Green
Write-Host "  ·  " -NoNewline -ForegroundColor DarkGray
Write-Host "WAF" -NoNewline -ForegroundColor Blue
Write-Host "  ·  " -NoNewline -ForegroundColor DarkGray
Write-Host "LZA" -NoNewline -ForegroundColor Yellow
Write-Host "  ·  " -NoNewline -ForegroundColor DarkGray
Write-Host "SEC" -NoNewline -ForegroundColor Red
Write-Host "  ·  " -NoNewline -ForegroundColor DarkGray
Write-Host "FSL" -ForegroundColor Magenta
Write-Host "  $('─' * 56)" -ForegroundColor DarkGray
Write-Host ""

$RequiredModules = @(
    @{ Name = 'Az.Accounts';              MinVersion = '2.7.5' }
    @{ Name = 'Az.DesktopVirtualization'; MinVersion = '4.0.0' }
    @{ Name = 'Az.Resources';             MinVersion = '6.0.0' }
    @{ Name = 'Az.Compute';               MinVersion = '5.0.0' }
    @{ Name = 'Az.Network';               MinVersion = '5.0.0' }
    @{ Name = 'Az.PrivateDns';            MinVersion = '1.0.0' }
    @{ Name = 'Az.Monitor';               MinVersion = '4.0.0' }
    @{ Name = 'Az.Storage';               MinVersion = '5.0.0' }
    @{ Name = 'Az.KeyVault';              MinVersion = '4.0.0' }
    @{ Name = 'Az.Security';              MinVersion = '1.0.0' }
)

# Pre-import modules with noisy warnings silently
foreach ($Noisy in @('Az.Network','Az.Monitor','Az.PrivateDns')) {
    Import-Module $Noisy -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
}

Write-Status "Prerequisites" -Level 'SECTION'
$Missing = @()
foreach ($Mod in $RequiredModules) {
    $Installed = Get-Module -ListAvailable -Name $Mod.Name -ErrorAction SilentlyContinue |
                 Sort-Object Version -Descending | Select-Object -First 1
    if (-not $Installed) {
        $Missing += $Mod.Name
        Write-Status "$($Mod.Name) >= $($Mod.MinVersion) - MISSING" -Level 'ERROR'
    } else {
        Write-Status "$($Mod.Name) v$($Installed.Version)" -Level 'SUCCESS'
    }
}

if ($Missing.Count -gt 0) {
    Write-Host ""
    Write-Status "Install missing modules:" -Level 'WARN'
    Write-Host "  Install-Module -Name '$($Missing -join "', '")' -Scope CurrentUser -Force" -ForegroundColor Yellow
    Write-Host ""
    exit 1
}

# OPTIONAL module — Az.OperationalInsights powers the AVD Insights / Log Analytics KQL checks
# (NET-008 latency, MON-002/008/009/010, PROF-010). Not a hard prerequisite: when absent, those
# checks individually emit Status 'Error' ("Az.OperationalInsights not installed") instead of
# blocking the whole run.
$OptionalModules = @(
    @{ Name = 'Az.OperationalInsights'; MinVersion = '3.0.0'; Reason = 'Log Analytics KQL checks (AVD Insights latency, Perf/Event, storage IOPS, profile load times)' }
)
foreach ($Opt in $OptionalModules) {
    $OptInstalled = Get-Module -ListAvailable -Name $Opt.Name -ErrorAction SilentlyContinue |
                    Sort-Object Version -Descending | Select-Object -First 1
    if ($OptInstalled) {
        Write-Status "$($Opt.Name) v$($OptInstalled.Version) (optional)" -Level 'SUCCESS'
    } else {
        Write-Status "$($Opt.Name) not installed (optional) - $($Opt.Reason) will report Error" -Level 'WARN'
    }
}

# ═══════════════════════════════════════════════════════════════════════════
# AUTHENTICATION
# ═══════════════════════════════════════════════════════════════════════════

Write-Status "Authentication" -Level 'SECTION'
if (-not $SkipLogin) {
    try {
        $Context = Get-AzContext -ErrorAction SilentlyContinue
        if ($Context -and $Context.Account) {
            Write-Status "$($Context.Account.Id)" -Level 'SUCCESS'
            Write-Status "Subscription: $($Context.Subscription.Name)" -Level 'INFO'
        } else {
            Write-Status "Launching interactive login..." -Level 'INFO'
            Connect-AzAccount -ErrorAction Stop | Out-Null
            $Context = Get-AzContext
            Write-Status "$($Context.Account.Id)" -Level 'SUCCESS'
        }
    } catch {
        Write-Status "Authentication failed: $($_.Exception.Message)" -Level 'ERROR'
        exit 1
    }
} else {
    $Context = Get-AzContext -ErrorAction SilentlyContinue
    if (-not $Context -or -not $Context.Account) {
        Write-Status "No existing Az context. Run Connect-AzAccount first or remove -SkipLogin." -Level 'ERROR'
        exit 1
    }
    Write-Status "$($Context.Account.Id)" -Level 'SUCCESS'
}

# Determine subscriptions to scan
if (-not $SubscriptionId -or $SubscriptionId.Count -eq 0) {
    # Show all available subscriptions and let user pick
    Write-Status "Subscriptions" -Level 'SECTION'
    $AllSubs = @(Get-AzSubscription -WarningAction SilentlyContinue -ErrorAction Stop |
        Where-Object { $_.State -eq 'Enabled' } | Sort-Object Name)

    if ($AllSubs.Count -eq 0) {
        Write-Status "No enabled subscriptions found" -Level 'ERROR'
        exit 1
    } elseif ($AllSubs.Count -eq 1) {
        $SubscriptionId = @($AllSubs[0].Id)
        Write-Status "Only one subscription available: $($AllSubs[0].Name)" -Level 'INFO'
    } else {
        Write-Host ""
        Write-Host "  Available subscriptions:" -ForegroundColor White
        Write-Host ""
        for ($i = 0; $i -lt $AllSubs.Count; $i++) {
            $Marker = if ($AllSubs[$i].Id -eq $Context.Subscription.Id) { ' *' } else { '  ' }
            $Idx = "$($i + 1)".PadLeft(3)
            Write-Host "  $Idx.$Marker " -NoNewline -ForegroundColor DarkCyan
            Write-Host $AllSubs[$i].Name -NoNewline -ForegroundColor White
            Write-Host " ($($AllSubs[$i].Id))" -ForegroundColor DarkGray
        }
        Write-Host ""
        Write-Host "  * = current context" -ForegroundColor DarkGray
        Write-Host ""
        $Selection = Read-Host "  Enter subscription number(s) (comma-separated, or 'all', default=current)"
        $Selection = $Selection.Trim()

        if (-not $Selection -or $Selection -eq '') {
            $SubscriptionId = @($Context.Subscription.Id)
            Write-Status "Using current: $($Context.Subscription.Name)" -Level 'INFO'
        } elseif ($Selection -eq 'all') {
            $SubscriptionId = @($AllSubs.Id)
            Write-Status "Scanning all $($AllSubs.Count) subscriptions" -Level 'INFO'
        } else {
            $SubscriptionId = @()
            foreach ($Num in ($Selection -split ',')) {
                $Idx = [int]$Num.Trim() - 1
                if ($Idx -ge 0 -and $Idx -lt $AllSubs.Count) {
                    $SubscriptionId += $AllSubs[$Idx].Id
                    Write-Status "Selected: $($AllSubs[$Idx].Name)" -Level 'SUCCESS'
                } else {
                    Write-Status "Invalid selection: $($Num.Trim())" -Level 'WARN'
                }
            }
            if ($SubscriptionId.Count -eq 0) {
                Write-Status "No valid subscriptions selected" -Level 'ERROR'
                exit 1
            }
        }
    }
}

# ═══════════════════════════════════════════════════════════════════════════
# DISCOVERY
# ═══════════════════════════════════════════════════════════════════════════

$Discovery = [PSCustomObject]@{
    SchemaVersion  = '1.0'
    ToolVersion    = $ScriptVersion
    Timestamp      = (Get-Date -Format 'o')
    AssessorId     = $Context.Account.Id
    Subscriptions  = @()
    Inventory      = [PSCustomObject]@{
        HostPools      = @()
        SessionHosts   = @()
        AppGroups      = @()
        Workspaces     = @()
        ScalingPlans   = @()
        VNets          = @()
        NSGs           = @()
        StorageAccounts = @()
        KeyVaults       = @()
        PrivateDnsZones = @()
        NetworkWatchers = @()
        OrphanedDisks   = @()
        OrphanedNICs    = @()
        PolicyAssignments = @()
        AlertRules      = @()
        GalleryImageVersions = @()
        Quotas          = @()
        CapacityReservations = @()
        Budgets         = @()
        Reservations    = @()
        Firewalls       = @()
        VPNGateways     = @()
        Subnets         = @()
        UDRs            = @()
        PrivateEndpoints = @()
    }
    CheckResults   = @()
    Errors         = @()
}

$AllChecks = [System.Collections.ArrayList]::new()

# Log Analytics workspace resource IDs harvested from host-pool diagnostic settings (feeds MON-SIEM / MON-012).
$LAWorkspaceIds = @{}

foreach ($SubId in $SubscriptionId) {
    Write-Status "Subscription: $SubId" -Level 'SECTION'

    try {
        Set-AzContext -SubscriptionId $SubId -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null
        $Sub = Get-AzContext
        $Discovery.Subscriptions += [PSCustomObject]@{
            Id   = $SubId
            Name = $Sub.Subscription.Name
        }
        # Short sub id used to keep singleton check IDs unique across subscriptions (A-1).
        $SubShort = ($SubId -split '-')[0]
        Write-Status "$($Sub.Subscription.Name)" -Level 'SUCCESS'
    } catch {
        Write-Status "Failed to set subscription context: $($_.Exception.Message)" -Level 'ERROR'
        $Discovery.Errors += "Failed to access subscription $SubId : $($_.Exception.Message)"
        continue
    }

    # ─── HOST POOLS ───────────────────────────────────────────────────────
    Write-Status "Host Pools" -Level 'SECTION'
    try {
        $HostPools = @(Get-AzWvdHostPool -ErrorAction Stop)
        Write-Status "  Found $($HostPools.Count) host pool(s)" -Level 'SUCCESS'

        foreach ($HP in $HostPools) {
            $HPObj = [PSCustomObject]@{
                SubscriptionId       = $SubId
                ResourceGroup        = ($HP.Id -split '/')[4]
                Name                 = $HP.Name
                Id                   = $HP.Id
                # Az 'Support' enum wrappers (HostPoolType/LoadBalancerType/PreferredAppGroupType) are
                # objects, not strings: they compare fine at runtime but serialize to {} in JSON, breaking
                # every '-eq' comparison after a save/reload. Coerce to string so they survive the round-trip.
                HostPoolType         = [string]$HP.HostPoolType
                LoadBalancerType     = [string]$HP.LoadBalancerType
                MaxSessionLimit      = $HP.MaxSessionLimit
                PreferredAppGroupType = [string]$HP.PreferredAppGroupType
                StartVMOnConnect     = $HP.StartVMOnConnect
                ValidationEnvironment = $HP.ValidationEnvironment
                Location             = $HP.Location
                Tags                 = $HP.Tag
                CustomRdpProperty    = $HP.CustomRdpProperty
            }
            $Discovery.Inventory.HostPools += $HPObj

            # ─── CHECK: Start VM on Connect ───
            [void]$AllChecks.Add((New-CheckResult -Id "GOV-001-$($HP.Name)" `
                -Category 'Governance & Cost' -Name 'Start VM on Connect' `
                -Description 'Start VM on Connect reduces costs by starting VMs only when users need them' `
                -Status $(if ($HP.StartVMOnConnect) { 'Pass' } else { 'Warning' }) `
                -Severity 'Medium' `
                -Details "StartVMOnConnect: $($HP.StartVMOnConnect)" `
                -Recommendation 'Enable Start VM on Connect to reduce compute costs during off-hours.' `
                -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/start-virtual-machine-connect' `
                -Evidence @{ HostPool = $HP.Name; Enabled = $HP.StartVMOnConnect }))

            # ─── CHECK: Max Session Limit ───
            # 999999 is Azure's "unlimited" sentinel (the never-configured default) - Fail.
            $SessionLimitDetail = "MaxSessionLimit: $($HP.MaxSessionLimit)"
            if ($HP.HostPoolType -ne 'Pooled') {
                $SessionLimitStatus = 'N/A'
            } elseif ($HP.MaxSessionLimit -le 0 -or $HP.MaxSessionLimit -ge 999999) {
                $SessionLimitStatus = 'Fail'
                $SessionLimitDetail = "MaxSessionLimit: $($HP.MaxSessionLimit) (default/unlimited - set an explicit workload-appropriate limit)"
            } elseif ($HP.MaxSessionLimit -gt 100) {
                $SessionLimitStatus = 'Warning'
                $SessionLimitDetail = "MaxSessionLimit: $($HP.MaxSessionLimit) (unusually high - verify VM sizing supports this density)"
            } else {
                $SessionLimitStatus = 'Pass'
            }
            [void]$AllChecks.Add((New-CheckResult -Id "SH-002-$($HP.Name)" `
                -Category 'Session Hosts' -Name 'Max Session Limit Configured' `
                -Description 'Pooled host pools should have an explicit max session limit' `
                -Status $SessionLimitStatus -Severity 'High' `
                -Details $SessionLimitDetail `
                -Recommendation 'Set MaxSessionLimit based on VM sizing and workload profile (typically 4-16 for multi-session).' `
                -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/configure-host-pool-load-balancing'))

            # ─── CHECK: Validation environment ───
            [void]$AllChecks.Add((New-CheckResult -Id "OPS-001-$($HP.Name)" `
                -Category 'Governance & Cost' -Name 'Validation Host Pool' `
                -Description 'At least one host pool should be marked as validation environment for safe update rollout' `
                -Status $(if ($HP.ValidationEnvironment) { 'Pass' } else { 'Warning' }) `
                -Severity 'Low' `
                -Details "ValidationEnvironment: $($HP.ValidationEnvironment)" `
                -Recommendation 'Mark at least one host pool as validation environment to receive service updates first.' `
                -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/configure-validation-environment'))

            # ─── CHECK: Comprehensive RDP Property Security Audit ───
            $RdpProps = $HP.CustomRdpProperty
            $ParsedRdp = @{}
            if ($RdpProps) {
                foreach ($RdpToken in ($RdpProps -split ';' | Where-Object { $_.Trim() })) {
                    $RdpParts = $RdpToken.Trim() -split ':', 3
                    if ($RdpParts.Count -ge 3) { $ParsedRdp[$RdpParts[0].ToLower()] = $RdpParts[2] }
                    elseif ($RdpParts.Count -eq 2) { $ParsedRdp[$RdpParts[0].ToLower()] = $RdpParts[1] }
                }
            }

            # Drive redirection (default when unset: Empty = no drives redirected)
            $DriveVal = $ParsedRdp['drivestoredirect']
            [void]$AllChecks.Add((New-CheckResult -Id "SEC-DRIVE-$($HP.Name)" `
                -Category 'Security' -Name 'RDP Drive Redirection' `
                -Description 'Drive/disk redirection should be restricted to prevent data exfiltration' `
                -Status $(if ($null -eq $DriveVal -or $DriveVal -eq '') { 'Pass' } else { 'Warning' }) `
                -Severity 'Medium' `
                -Details "drivestoredirect: $(if ($null -eq $DriveVal) { '(not set - default: no drives redirected)' } else { "'$DriveVal'" })" `
                -Recommendation 'Leave drivestoredirect unset or empty to block drive redirection; restrict to specific drives only if required.' `
                -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/rdp-properties#device-redirection'))

            # Secure-by-default redirection change (clipboard/printer/camera/mic off when unset) applies only to
            # host pools created after the mid-2025 rollout and is NOT retroactive. Unset therefore cannot be
            # assumed disabled - warn with a creation-date caveat.
            $RedirDefaultCaveat = 'default depends on host pool creation date - pools created before mid-2025 default to enabled; verify effective value'

            # Clipboard redirection
            $ClipVal = $ParsedRdp['redirectclipboard']
            [void]$AllChecks.Add((New-CheckResult -Id "SEC-CLIP-$($HP.Name)" `
                -Category 'Security' -Name 'RDP Clipboard Redirection' `
                -Description 'Clipboard redirection should be restricted for sensitive environments' `
                -Status $(if ($ClipVal -eq '0') { 'Pass' } else { 'Warning' }) `
                -Severity 'Medium' `
                -Details "redirectclipboard: $(if ($null -eq $ClipVal) { "(not set - $RedirDefaultCaveat)" } else { $ClipVal })" `
                -Recommendation 'Set redirectclipboard:i:0 explicitly to disable, or use clipboard transfer direction policies for granular control.' `
                -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/rdp-properties#device-redirection'))

            # Printer redirection
            $PrintVal = $ParsedRdp['redirectprinters']
            [void]$AllChecks.Add((New-CheckResult -Id "SEC-PRINT-$($HP.Name)" `
                -Category 'Security' -Name 'RDP Printer Redirection' `
                -Description 'Printer redirection should be evaluated - disable if not required' `
                -Status $(if ($PrintVal -eq '0') { 'Pass' } else { 'Warning' }) `
                -Severity 'Low' `
                -Details "redirectprinters: $(if ($null -eq $PrintVal) { "(not set - $RedirDefaultCaveat)" } else { $PrintVal })" `
                -Recommendation 'Set redirectprinters:i:0 explicitly to disable, or redirectprinters:i:1 only if printer redirection is required.' `
                -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/rdp-properties#device-redirection'))

            # USB redirection
            $UsbVal = $ParsedRdp['usbdevicestoredirect']
            [void]$AllChecks.Add((New-CheckResult -Id "SEC-USB-$($HP.Name)" `
                -Category 'Security' -Name 'RDP USB Redirection' `
                -Description 'USB device redirection should be blocked unless explicitly required' `
                -Status $(if ($null -eq $UsbVal -or $UsbVal -eq '') { 'Pass' } else { 'Warning' }) `
                -Severity 'Medium' `
                -Details "usbdevicestoredirect: $(if ($null -eq $UsbVal) { '(not set - default: none)' } else { "'$UsbVal'" })" `
                -Recommendation 'Remove usbdevicestoredirect or set to empty to block USB device redirection.' `
                -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/rdp-properties#device-redirection'))

            # COM port redirection
            $ComVal = $ParsedRdp['redirectcomports']
            if ($ComVal -and $ComVal -ne '0') {
                [void]$AllChecks.Add((New-CheckResult -Id "SEC-COM-$($HP.Name)" `
                    -Category 'Security' -Name 'RDP COM Port Redirection' `
                    -Description 'COM port redirection is rarely needed and should be disabled' `
                    -Status 'Warning' -Severity 'Medium' `
                    -Details "redirectcomports: $ComVal" `
                    -Recommendation 'Set redirectcomports:i:0 to disable COM port redirection.' `
                    -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/rdp-properties#device-redirection'))
            }

            # Camera redirection (webcam) - followed the same secure-defaults change.
            # camerastoredirect:s: (empty) = disabled; :s:* or device list = enabled.
            $CamVal = $ParsedRdp['camerastoredirect']
            [void]$AllChecks.Add((New-CheckResult -Id "SEC-CAM-$($HP.Name)" `
                -Category 'Security' -Name 'RDP Camera Redirection' `
                -Description 'Camera redirection status - required for Teams calls, evaluate for security-sensitive workloads' `
                -Status $(if ($null -eq $CamVal) { 'Warning' } elseif ($CamVal -eq '') { 'Pass' } else { 'Warning' }) `
                -Severity 'Low' `
                -Details "camerastoredirect: $(if ($null -eq $CamVal) { "(not set - $RedirDefaultCaveat)" } else { "'$CamVal'" })" `
                -Recommendation 'Set camerastoredirect:s: (empty) to disable webcam redirection for sensitive workloads; allow only if required for calls.' `
                -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/rdp-properties#device-redirection'))

            # Audio capture (microphone). audiocapturemode:i:0 = disabled, :i:1 = enabled.
            $AudioCapVal = $ParsedRdp['audiocapturemode']
            [void]$AllChecks.Add((New-CheckResult -Id "SEC-AUDIO-$($HP.Name)" `
                -Category 'Security' -Name 'RDP Audio Capture (Microphone)' `
                -Description 'Audio input capture - needed for calls, evaluate for other workloads' `
                -Status $(if ($null -eq $AudioCapVal) { 'Warning' } elseif ($AudioCapVal -eq '0') { 'Pass' } else { 'Warning' }) `
                -Severity 'Low' `
                -Details "audiocapturemode: $(if ($null -eq $AudioCapVal) { "(not set - $RedirDefaultCaveat)" } else { $AudioCapVal })" `
                -Recommendation 'Set audiocapturemode:i:0 to disable microphone redirection for sensitive workloads; allow only if required for calls.' `
                -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/rdp-properties#device-redirection'))

            # RDP property summary for evidence
            $RdpSecurityIssues = @()
            # Defaults when a property is unset are secure (drives Empty, clipboard 0, printers 0);
            # only flag redirection that has been explicitly enabled.
            if ($ParsedRdp['drivestoredirect'] -and $ParsedRdp['drivestoredirect'] -ne '') { $RdpSecurityIssues += 'Drives:Open' }
            if ($ParsedRdp['redirectclipboard'] -eq '1') { $RdpSecurityIssues += 'Clipboard:Open' }
            if ($ParsedRdp['redirectprinters'] -eq '1') { $RdpSecurityIssues += 'Printers:Open' }
            if ($ParsedRdp['usbdevicestoredirect'] -and $ParsedRdp['usbdevicestoredirect'] -ne '') { $RdpSecurityIssues += 'USB:Open' }
            if ($ParsedRdp['redirectcomports'] -eq '1') { $RdpSecurityIssues += 'COM:Open' }
            [void]$AllChecks.Add((New-CheckResult -Id "SEC-RDP-$($HP.Name)" `
                -Category 'Security' -Name 'RDP Properties Security Summary' `
                -Description 'Overall RDP property security posture' `
                -Status $(if ($RdpSecurityIssues.Count -eq 0) { 'Pass' } elseif ($RdpSecurityIssues.Count -le 2) { 'Warning' } else { 'Fail' }) `
                -Severity 'High' `
                -Details "Issues: $(if ($RdpSecurityIssues.Count -eq 0) { 'None - all redirections restricted' } else { $RdpSecurityIssues -join ', ' }). AllProps: $RdpProps" `
                -Recommendation 'Review and restrict all device redirections per security requirements.' `
                -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/rdp-properties' `
                -Evidence @{ HostPool = $HP.Name; ParsedProperties = $ParsedRdp; Issues = $RdpSecurityIssues }))

            # ─── CHECK: SSO via Entra ID (from RDP properties) ───
            $HasSSO = $ParsedRdp['enablerdsaadauth'] -eq '1'
            $IsEntraTarget = $ParsedRdp['targetisaadjoined'] -eq '1'
            [void]$AllChecks.Add((New-CheckResult -Id "IAM-SSO-$($HP.Name)" `
                -Category 'Identity & Access' -Name 'Single Sign-On (SSO)' `
                -Description 'Entra ID SSO should be enabled for seamless authentication' `
                -Status $(if ($HasSSO) { 'Pass' } elseif ($IsEntraTarget) { 'Fail' } else { 'Warning' }) `
                -Severity $(if ($IsEntraTarget -and -not $HasSSO) { 'High' } else { 'Medium' }) `
                -Details "SSO: $(if ($HasSSO) { 'Enabled' } else { 'Not configured' }), EntraTarget: $(if ($IsEntraTarget) { 'Yes' } else { 'Not set' })" `
                -Recommendation $(if ($IsEntraTarget -and -not $HasSSO) { 'SSO is strongly recommended for Entra ID joined hosts - enable enablerdsaadauth:i:1.' } else { 'Enable SSO for seamless authentication.' }) `
                -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/configure-single-sign-on'))

            # NOTE: Watermarking (SEC-WM) and Screen Capture Protection (SEC-SCP) are configured via
            # GPO/Intune registry policy (fEnableWatermarking / fEnableScreenCaptureProtect), never via
            # CustomRdpProperty. RDP-property greps produced permanent false negatives, so those emits
            # were removed and the checks reclassified Manual in checks.json (audit A-4).

            # ─── CHECK: Tag Quality ───
            $HasTags = $HP.Tag -and $HP.Tag.Count -gt 0
            $TagKeys = if ($HasTags) { @($HP.Tag.Keys) } else { @() }
            $RecommendedTags = @('Environment','Owner','CostCenter','Application','Department')
            $FoundRecommended = @($RecommendedTags | Where-Object { $TagKey = $_; $TagKeys | Where-Object { $_ -like "*$TagKey*" } })
            $TagScore = if ($TagKeys.Count -eq 0) { 0 } else { [math]::Round($FoundRecommended.Count / $RecommendedTags.Count * 100) }
            [void]$AllChecks.Add((New-CheckResult -Id "GOV-TAG-$($HP.Name)" `
                -Category 'Governance & Cost' -Name 'Resource Tag Quality' `
                -Description 'Resources should have tags for cost management and organization' `
                -Status $(if ($TagScore -ge 60) { 'Pass' } elseif ($HasTags) { 'Warning' } else { 'Fail' }) `
                -Severity 'Low' `
                -Details "Tags: $(if ($HasTags) { ($TagKeys -join ', ') } else { 'None' }). Score: $TagScore% ($($FoundRecommended.Count)/$($RecommendedTags.Count) recommended tags found)" `
                -Recommendation "Apply recommended tags: $($RecommendedTags -join ', ')." `
                -Reference 'https://learn.microsoft.com/en-us/azure/cloud-adoption-framework/ready/azure-best-practices/resource-tagging' `
                -Evidence @{ HostPool = $HP.Name; TagCount = $TagKeys.Count; Score = $TagScore; Missing = @($RecommendedTags | Where-Object { $_ -notin $FoundRecommended }) }))

            # NOTE: Load Balancing Algorithm check (SH-LB) is emitted later in the Scaling Plans section,
            # where scaling-plan assignment is known (autoscale overrides LB during scheduled hours).

            # ─── CHECK: Host Pool Private Link (SEC-024) ───
            try {
                $HPResource = Get-AzResource -ResourceId $HP.Id -ErrorAction SilentlyContinue
                $HPPrivateEndpoints = if ($HPResource -and $HPResource.Properties.privateEndpointConnections) {
                    @($HPResource.Properties.privateEndpointConnections)
                } else { @() }
                [void]$AllChecks.Add((New-CheckResult -Id "SEC-HPPL-$($HP.Name)" `
                    -Category 'Security' -Name 'Host Pool Private Link' `
                    -Description 'AVD host pools should use Private Link for control-plane traffic' `
                    -Status $(if ($HPPrivateEndpoints.Count -gt 0) { 'Pass' } else { 'Warning' }) `
                    -Severity 'Medium' `
                    -Details "PrivateEndpoints: $($HPPrivateEndpoints.Count)" `
                    -Recommendation 'Configure AVD Private Link to keep session brokering traffic off the public internet.' `
                    -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/private-link-overview' `
                    -Evidence @{ HostPool = $HP.Name; PECount = $HPPrivateEndpoints.Count }))
            } catch {
                [void]$AllChecks.Add((New-CheckResult -Id "SEC-HPPL-$($HP.Name)" `
                    -Category 'Security' -Name 'Host Pool Private Link' `
                    -Description 'AVD host pools should use Private Link for control-plane traffic' `
                    -Status 'Error' -Severity 'Medium' `
                    -Details "Could not assess host pool Private Link: $($_.Exception.Message)" `
                    -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/private-link-overview'))
            }

            # ─── CHECK: Private Link / Private Endpoints (NET-005) — reuses HP privateEndpointConnections ───
            try {
                $HPRes2 = Get-AzResource -ResourceId $HP.Id -ErrorAction SilentlyContinue
                $HPPE = if ($HPRes2 -and $HPRes2.Properties.privateEndpointConnections) { @($HPRes2.Properties.privateEndpointConnections) } else { @() }
                [void]$AllChecks.Add((New-CheckResult -Id "NET-PL-$($HP.Name)" `
                    -Category 'Networking' -Name 'Private Link / Private Endpoints' `
                    -Description 'AVD control-plane resources should use Private Link to keep management traffic off the public internet' `
                    -Status $(if ($HPPE.Count -gt 0) { 'Pass' } else { 'Warning' }) `
                    -Severity 'Medium' `
                    -Details "HostPool $($HP.Name): privateEndpointConnections: $($HPPE.Count)" `
                    -Recommendation 'Configure AVD Private Link (feed/broker/gateway) so control-plane traffic stays on the Microsoft backbone. RDP Shortpath over Private Link is supported.' `
                    -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/private-link-overview' `
                    -Evidence @{ HostPool = $HP.Name; PECount = $HPPE.Count }))
            } catch {
                [void]$AllChecks.Add((New-CheckResult -Id "NET-PL-$($HP.Name)" `
                    -Category 'Networking' -Name 'Private Link / Private Endpoints' `
                    -Description 'AVD control-plane resources should use Private Link' `
                    -Status 'Error' -Severity 'Medium' `
                    -Details "Could not assess Private Link: $($_.Exception.Message)" `
                    -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/private-link-overview'))
            }

            # ─── CHECK: Session Host Update feature (SH-020) — GA June 2026, automated management ───
            try {
                $ShuManaged = $false
                $ShuDetail  = ''
                # Prefer an explicit managementType on the host pool model when present.
                if ($HP.PSObject.Properties.Name -contains 'ManagementType' -and $HP.ManagementType) {
                    $ShuManaged = "$($HP.ManagementType)" -match 'Automated'
                    $ShuDetail  = "ManagementType: $($HP.ManagementType)"
                } else {
                    # Fall back to the sessionHostConfigurations child resource (session host update GA).
                    $ShuResp = Invoke-AzRestMethod -Path "$($HP.Id)/sessionHostConfigurations/default?api-version=2024-04-08-preview" -Method GET -ErrorAction SilentlyContinue
                    if ($ShuResp -and $ShuResp.StatusCode -eq 200) {
                        $ShuManaged = $true
                        $ShuDetail  = 'sessionHostConfigurations/default present (Automated session host management configured)'
                    } else {
                        $ShuManaged = $false
                        $ShuDetail  = "No session host configuration found (HTTP $(if ($ShuResp) { $ShuResp.StatusCode } else { 'n/a' })) - Standard (manual) management"
                    }
                }
                [void]$AllChecks.Add((New-CheckResult -Id "SH-SHU-$($HP.Name)" `
                    -Category 'Session Hosts' -Name 'Session Host Update Feature' `
                    -Description 'Session Host Update (GA June 2026) automates image rollout with health validation and rollback' `
                    -Status $(if ($ShuManaged) { 'Pass' } else { 'Warning' }) `
                    -Severity 'Medium' `
                    -Details "HostPool $($HP.Name): $ShuDetail" `
                    -Recommendation 'Session Host Update is GA — consider automated management (session host configuration + managed identity) to reduce golden-image lifecycle effort.' `
                    -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/session-host-update' `
                    -Evidence @{ HostPool = $HP.Name; Automated = $ShuManaged }))
            } catch {
                [void]$AllChecks.Add((New-CheckResult -Id "SH-SHU-$($HP.Name)" `
                    -Category 'Session Hosts' -Name 'Session Host Update Feature' `
                    -Description 'Session Host Update (GA June 2026) automates image rollout' `
                    -Status 'Error' -Severity 'Medium' `
                    -Details "Could not query session host configuration: $($_.Exception.Message)" `
                    -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/session-host-update'))
            }

            # ─── CHECK: App Attach - MSIX packages on host pool (APP-002) ───
            try {
                $MsixPkgs = @(Get-AzWvdMsixPackage -HostPoolName $HP.Name -ResourceGroupName (($HP.Id -split '/')[4]) -ErrorAction SilentlyContinue)
                if ($MsixPkgs.Count -gt 0) {
                    [void]$AllChecks.Add((New-CheckResult -Id "APP-ATTACH-MSIX-$($HP.Name)" `
                        -Category 'Application Delivery' -Name 'App Attach' `
                        -Description 'App Attach decouples application lifecycle from the golden image' `
                        -Status 'Pass' -Severity 'Low' `
                        -Details "HostPool $($HP.Name): $($MsixPkgs.Count) MSIX/App Attach package(s) ($(@($MsixPkgs | ForEach-Object { $_.DisplayName } | Select-Object -First 5) -join ', '))" `
                        -Recommendation 'Legacy MSIX App Attach was retired June 1 2025 — migrate to CIM-based App Attach.' `
                        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/app-attach-overview' `
                        -Evidence @{ HostPool = $HP.Name; PackageCount = $MsixPkgs.Count }))
                }
            } catch {
                Write-Status "    Could not enumerate MSIX packages for $($HP.Name): $($_.Exception.Message)" -Level 'WARN'
            }
        }
    } catch {
        Write-Status "  Error discovering host pools: $($_.Exception.Message)" -Level 'ERROR'
        $Discovery.Errors += "Host pool discovery failed: $($_.Exception.Message)"
    }

    # ─── SESSION HOSTS ────────────────────────────────────────────────────
    Write-Status "Session Hosts" -Level 'SECTION'
    # E-7 perf: cache Get-AzVM (model) and Get-AzVM -Status per resource group so we make one ARM
    # call per RG instead of two per session host. Keyed by "$RG/$VMName" (lower-case).
    $VMModelCache  = @{}
    $VMStatusCache = @{}
    $VMRGLoaded    = @{}
    $VMSizeCache   = @{}  # Get-AzVMSize per location (B-2)
    try {
        foreach ($HP in $HostPools) {
            $RG = ($HP.Id -split '/')[4]
            $SessionHosts = @(Get-AzWvdSessionHost -ResourceGroupName $RG -HostPoolName $HP.Name -ErrorAction Stop)
            Write-Status "  $($HP.Name): $($SessionHosts.Count) session host(s)" -Level 'INFO'

            foreach ($SH in $SessionHosts) {
                $VMName = ($SH.ResourceId -split '/')[-1]
                $VMRG   = ($SH.ResourceId -split '/')[4]

                # Batch-load all VMs in this RG (model + status) on first encounter.
                if (-not $VMRGLoaded.ContainsKey($VMRG)) {
                    $VMRGLoaded[$VMRG] = $true
                    try {
                        foreach ($V in @(Get-AzVM -ResourceGroupName $VMRG -ErrorAction Stop)) {
                            $VMModelCache["$VMRG/$($V.Name)".ToLower()] = $V
                        }
                    } catch {
                        Write-Status "    Could not batch VM models for RG $VMRG : $($_.Exception.Message)" -Level 'WARN'
                    }
                    try {
                        foreach ($V in @(Get-AzVM -ResourceGroupName $VMRG -Status -ErrorAction Stop)) {
                            $VMStatusCache["$VMRG/$($V.Name)".ToLower()] = $V
                        }
                    } catch {
                        Write-Status "    Could not batch VM status for RG $VMRG : $($_.Exception.Message)" -Level 'WARN'
                    }
                }

                # VM model (size, security, storage, zones) from cache
                $VMKey = "$VMRG/$VMName".ToLower()
                $VMModel = $VMModelCache[$VMKey]
                if (-not $VMModel) {
                    Write-Status "    Could not get VM model for $VMName (not in RG batch)" -Level 'WARN'
                }
                # VM instance view (power state, extensions) from cache
                $VMInstance = $VMStatusCache[$VMKey]

                # Derive extension-based properties from instance view
                $ExtList = @()
                $SHJoinType = 'Unknown'
                $HasTrustedLaunch = $false
                $HasAMAExt = $false
                $HasMDEExt = $false
                # Prefer instance-view extensions (running VMs); fall back to ARM model extensions (deallocated VMs)
                $RawExts = if ($VMInstance -and $VMInstance.Extensions) {
                    $VMInstance.Extensions
                } elseif ($VMModel -and $VMModel.Extensions) {
                    $VMModel.Extensions
                } else { $null }
                # Distinguish "no join extension found" from "could not read extension data at all" (C-8).
                $JoinDataAvailable = ($null -ne $VMModel) -or ($null -ne $VMInstance)
                if ($RawExts) {
                    $ExtList = @($RawExts | ForEach-Object {
                        $ExtType = if ($_.VirtualMachineExtensionType) { $_.VirtualMachineExtensionType } else { $_.Type }
                        $ExtType
                    } | Where-Object { $_ })
                    $HasAADExt = 'AADLoginForWindows' -in $ExtList
                    $HasDJExt  = 'JsonADDomainExtension' -in $ExtList
                    # The domain-join extension alone cannot distinguish pure AD DS from Hybrid (Hybrid = AD join
                    # plus Entra Connect sync, which is not visible from VM extensions), so report both (C-8).
                    $SHJoinType = if ($HasAADExt -and $HasDJExt) { 'Hybrid' } elseif ($HasAADExt) { 'Entra ID' } elseif ($HasDJExt) { 'AD DS or Hybrid' } else { 'Unknown' }
                    $HasAMAExt = 'AzureMonitorWindowsAgent' -in $ExtList
                    # MicrosoftMonitoringAgent (MMA) was retired Aug 2024 and is NOT MDE - only MDE.Windows counts (B-4).
                    $HasMDEExt = @($ExtList | Where-Object { $_ -eq 'MDE.Windows' }).Count -gt 0
                }
                if ($VMModel -and $VMModel.SecurityProfile) {
                    $HasTrustedLaunch = $VMModel.SecurityProfile.SecurityType -eq 'TrustedLaunch'
                }

                $SHObj = [PSCustomObject]@{
                    HostPoolName        = $HP.Name
                    Name                = $SH.Name
                    ResourceId          = $SH.ResourceId
                    ResourceGroup       = $VMRG
                    Location            = if ($VMModel) { $VMModel.Location } else { $HP.Location }
                    Status              = [string]$SH.Status
                    AllowNewSession     = $SH.AllowNewSession
                    Sessions            = $SH.Session
                    AgentVersion        = $SH.AgentVersion
                    LastHeartBeat       = $SH.LastHeartBeat
                    OSVersion           = $SH.OSVersion
                    OsType              = if ($VMModel) { "$($VMModel.StorageProfile.OsDisk.OsType)" } else { $null }
                    UpdateState         = [string]$SH.UpdateState
                    VMSize              = if ($VMModel) { $VMModel.HardwareProfile.VmSize } else { $null }
                    OSDiskType          = if ($VMModel -and $VMModel.StorageProfile.OsDisk.ManagedDisk.StorageAccountType) {
                        $VMModel.StorageProfile.OsDisk.ManagedDisk.StorageAccountType
                    } elseif ($VMModel -and $VMModel.StorageProfile.OsDisk.ManagedDisk.Id) {
                        try { (Get-AzDisk -ResourceGroupName $VMRG -DiskName ($VMModel.StorageProfile.OsDisk.ManagedDisk.Id -split '/')[-1] -ErrorAction Stop).Sku.Name } catch { $null }
                    } else { $null }
                    SecurityProfile     = if ($VMModel -and $VMModel.SecurityProfile) {
                        [PSCustomObject]@{
                            SecurityType = $VMModel.SecurityProfile.SecurityType
                            SecureBoot   = $VMModel.SecurityProfile.UefiSettings.SecureBootEnabled
                            VTpm         = $VMModel.SecurityProfile.UefiSettings.VTpmEnabled
                        }
                    } else { $null }
                    TrustedLaunch       = $HasTrustedLaunch
                    PowerState          = if ($VMInstance) { ($VMInstance.Statuses | Where-Object Code -like 'PowerState/*').DisplayStatus } else { $null }
                    AcceleratedNetworking = $null
                    AvailabilityZone    = if ($VMModel) { $VMModel.Zones } else { $null }
                    ImageReference      = if ($VMModel) { $VMModel.StorageProfile.ImageReference } else { $null }
                    JoinType            = $SHJoinType
                    JoinDataAvailable   = $JoinDataAvailable
                    AMAInstalled        = $HasAMAExt
                    MDEInstalled        = $HasMDEExt
                    Extensions          = $ExtList
                    Tags                = if ($VMModel) { $VMModel.Tags } else { $null }
                    # NIC facts cached here (E-7) so the networking pass need not re-fetch VM + NIC.
                    NicSubnetId         = $null
                    NicHasPublicIP      = $null
                }

                # Check NIC for accelerated networking; cache subnet + public IP for the networking pass (E-7).
                if ($VMModel -and $VMModel.NetworkProfile.NetworkInterfaces.Count -gt 0) {
                    try {
                        $NicId = $VMModel.NetworkProfile.NetworkInterfaces[0].Id
                        $Nic = Get-AzNetworkInterface -ResourceId $NicId -ErrorAction Stop
                        $SHObj.AcceleratedNetworking = $Nic.EnableAcceleratedNetworking
                        if ($Nic.IpConfigurations -and $Nic.IpConfigurations[0].Subnet.Id) {
                            $SHObj.NicSubnetId = $Nic.IpConfigurations[0].Subnet.Id
                        }
                        $SHObj.NicHasPublicIP = ($null -ne $Nic.IpConfigurations[0].PublicIpAddress)
                    } catch {
                        Write-Status "    Could not read NIC for $VMName : $($_.Exception.Message)" -Level 'WARN'
                    }
                }

                $Discovery.Inventory.SessionHosts += $SHObj

                # ─── CHECK: Trusted Launch ───
                if ($VMModel -and $VMModel.SecurityProfile) {
                    $SecureBoot = $VMModel.SecurityProfile.UefiSettings.SecureBootEnabled
                    $VTpm       = $VMModel.SecurityProfile.UefiSettings.VTpmEnabled
                    $TrustedLaunch = $VMModel.SecurityProfile.SecurityType -eq 'TrustedLaunch'
                    [void]$AllChecks.Add((New-CheckResult -Id "SEC-TL-$VMName" `
                        -Category 'Security & IAM' -Name 'Trusted Launch' `
                        -Description 'Session hosts should use Trusted Launch with Secure Boot and vTPM enabled' `
                        -Status $(if ($TrustedLaunch -and $SecureBoot -and $VTpm) { 'Pass' }
                                  elseif ($TrustedLaunch) { 'Warning' } else { 'Fail' }) `
                        -Severity 'High' `
                        -Details "SecurityType: $($VMModel.SecurityProfile.SecurityType), SecureBoot: $SecureBoot, vTPM: $VTpm" `
                        -Recommendation 'Enable Trusted Launch with Secure Boot and vTPM for enhanced boot integrity.' `
                        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-machines/trusted-launch' `
                        -Evidence @{ VM = $VMName; SecurityType = $VMModel.SecurityProfile.SecurityType }))
                } elseif ($VMModel) {
                    # Model loaded but no SecurityProfile => Standard security type (not Trusted Launch).
                    [void]$AllChecks.Add((New-CheckResult -Id "SEC-TL-$VMName" `
                        -Category 'Security & IAM' -Name 'Trusted Launch' `
                        -Description 'Session hosts should use Trusted Launch' `
                        -Status 'Fail' -Severity 'High' `
                        -Details 'No security profile on VM model - VM is using Standard security type, not Trusted Launch.' `
                        -Recommendation 'Enable Trusted Launch in-place on Gen2 VMs, or redeploy with Trusted Launch security type.' `
                        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-machines/trusted-launch'))
                } else {
                    # Model could not be read (API/permission error) - do not fabricate a conclusion (A-10).
                    [void]$AllChecks.Add((New-CheckResult -Id "SEC-TL-$VMName" `
                        -Category 'Security & IAM' -Name 'Trusted Launch' `
                        -Status 'Error' -Severity 'High' `
                        -Description 'Session hosts should use Trusted Launch' `
                        -Details "Could not read VM model for $VMName - security type undetermined." `
                        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-machines/trusted-launch'))
                }

                # ─── CHECK: Secure Boot (SH-025) ───
                if ($VMModel -and $VMModel.SecurityProfile) {
                    $SBEnabled = $VMModel.SecurityProfile.UefiSettings.SecureBootEnabled -eq $true
                    [void]$AllChecks.Add((New-CheckResult -Id "SH-SECBOOT-$VMName" `
                        -Category 'Session Hosts' -Name 'Secure Boot Enabled' `
                        -Description 'Trusted Launch VMs should have Secure Boot enabled to protect against boot-level malware' `
                        -Status $(if ($SBEnabled) { 'Pass' } else { 'Fail' }) `
                        -Severity 'High' `
                        -Details "SecureBoot: $SBEnabled" `
                        -Recommendation 'Enable Secure Boot in the VM security profile to protect the boot chain.' `
                        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-machines/trusted-launch#secure-boot' `
                        -Evidence @{ VM = $VMName; SecureBoot = $SBEnabled }))
                } elseif ($VMModel) {
                    [void]$AllChecks.Add((New-CheckResult -Id "SH-SECBOOT-$VMName" `
                        -Category 'Session Hosts' -Name 'Secure Boot Enabled' `
                        -Description 'Trusted Launch VMs should have Secure Boot enabled to protect against boot-level malware' `
                        -Status 'Fail' -Severity 'High' `
                        -Details 'No security profile - Secure Boot not enabled; VM is not Trusted Launch.' `
                        -Recommendation 'Enable Trusted Launch with Secure Boot to protect the boot chain.' `
                        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-machines/trusted-launch#secure-boot'))
                } else {
                    [void]$AllChecks.Add((New-CheckResult -Id "SH-SECBOOT-$VMName" `
                        -Category 'Session Hosts' -Name 'Secure Boot Enabled' `
                        -Status 'Error' -Severity 'High' `
                        -Description 'Trusted Launch VMs should have Secure Boot enabled' `
                        -Details "Could not read VM model for $VMName - Secure Boot state undetermined." `
                        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-machines/trusted-launch#secure-boot'))
                }

                # ─── CHECK: vTPM (SH-026) ───
                if ($VMModel -and $VMModel.SecurityProfile) {
                    $VTpmOn = $VMModel.SecurityProfile.UefiSettings.VTpmEnabled -eq $true
                    [void]$AllChecks.Add((New-CheckResult -Id "SH-VTPM-$VMName" `
                        -Category 'Session Hosts' -Name 'vTPM Enabled' `
                        -Description 'Trusted Launch VMs should have vTPM enabled for measured boot and key protection' `
                        -Status $(if ($VTpmOn) { 'Pass' } else { 'Fail' }) `
                        -Severity 'High' `
                        -Details "vTPM: $VTpmOn" `
                        -Recommendation 'Enable vTPM to support BitLocker, measured boot, and Windows Hello for Business.' `
                        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-machines/trusted-launch#vtpm' `
                        -Evidence @{ VM = $VMName; VTpm = $VTpmOn }))
                } elseif ($VMModel) {
                    [void]$AllChecks.Add((New-CheckResult -Id "SH-VTPM-$VMName" `
                        -Category 'Session Hosts' -Name 'vTPM Enabled' `
                        -Description 'Trusted Launch VMs should have vTPM enabled for measured boot and key protection' `
                        -Status 'Fail' -Severity 'High' `
                        -Details 'No security profile - vTPM not enabled; VM is not Trusted Launch.' `
                        -Recommendation 'Enable Trusted Launch with vTPM for measured boot and key protection.' `
                        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-machines/trusted-launch#vtpm'))
                } else {
                    [void]$AllChecks.Add((New-CheckResult -Id "SH-VTPM-$VMName" `
                        -Category 'Session Hosts' -Name 'vTPM Enabled' `
                        -Status 'Error' -Severity 'High' `
                        -Description 'Trusted Launch VMs should have vTPM enabled' `
                        -Details "Could not read VM model for $VMName - vTPM state undetermined." `
                        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-machines/trusted-launch#vtpm'))
                }

                # ─── CHECK: OS Disk Encryption - ADE or host-based (SEC-021) ───
                if ($VMModel) {
                    $HasADE = $false
                    $DiskEncType = 'None'
                    if ($SHObj.Extensions) {
                        $HasADE = $SHObj.Extensions | Where-Object { $_ -match 'AzureDiskEncryption' }
                    }
                    $OsDisk = $VMModel.StorageProfile.OsDisk
                    if ($OsDisk.ManagedDisk.SecurityProfile.DiskEncryptionSet) {
                        $DiskEncType = 'DiskEncryptionSet'
                    } elseif ($VMModel.SecurityProfile.EncryptionAtHost) {
                        $DiskEncType = 'EncryptionAtHost'
                    } elseif ($HasADE) {
                        $DiskEncType = 'ADE'
                    }
                    $IsEncrypted = $DiskEncType -ne 'None'
                    [void]$AllChecks.Add((New-CheckResult -Id "SEC-OSDISK-$VMName" `
                        -Category 'Security & IAM' -Name 'OS Disk Encryption' `
                        -Description 'OS disks should use ADE, host-based encryption, or customer-managed DES beyond platform default' `
                        -Status $(if ($IsEncrypted) { 'Pass' } else { 'Warning' }) `
                        -Severity 'High' `
                        -Details "EncryptionType: $DiskEncType" `
                        -Recommendation 'Enable Azure Disk Encryption or encryption at host for data-at-rest protection beyond platform-managed keys.' `
                        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/security-recommendations#azure-confidential-computing' `
                        -Evidence @{ VM = $VMName; EncryptionType = $DiskEncType }))
                }

                # ─── CHECK: Accelerated Networking ───
                # B-series and very small sizes (<2 vCPU) do not support accelerated networking - N/A, not Warning (C-7).
                $ANSpec = Get-VMSizeSpec -VMSize $SHObj.VMSize -Location $SHObj.Location -Cache $VMSizeCache
                $ANUnsupported = ($SHObj.VMSize -match '^Standard_B') -or ($ANSpec.VCPU -ne $null -and $ANSpec.VCPU -lt 2)
                if ($ANUnsupported) {
                    [void]$AllChecks.Add((New-CheckResult -Id "NET-AN-$VMName" `
                        -Category 'Networking' -Name 'Accelerated Networking' `
                        -Description 'Accelerated networking improves throughput and reduces latency' `
                        -Status 'N/A' -Severity 'Medium' `
                        -Details "VMSize: $($SHObj.VMSize) does not support accelerated networking (B-series / small size)." `
                        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-network/accelerated-networking-overview'))
                } elseif ($null -ne $SHObj.AcceleratedNetworking) {
                    [void]$AllChecks.Add((New-CheckResult -Id "NET-AN-$VMName" `
                        -Category 'Networking' -Name 'Accelerated Networking' `
                        -Description 'Accelerated networking improves throughput and reduces latency' `
                        -Status $(if ($SHObj.AcceleratedNetworking) { 'Pass' } else { 'Warning' }) `
                        -Severity 'Medium' `
                        -Details "AcceleratedNetworking: $($SHObj.AcceleratedNetworking)" `
                        -Recommendation 'Enable accelerated networking on session host NICs for better performance.' `
                        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-network/accelerated-networking-overview'))
                }

                # ─── CHECK: Managed disk encryption ───
                if ($SHObj.OSDiskType) {
                    [void]$AllChecks.Add((New-CheckResult -Id "SEC-DISK-$VMName" `
                        -Category 'Security & IAM' -Name 'Managed Disk Encryption' `
                        -Description 'OS disk should use managed encryption' `
                        -Status 'Pass' -Severity 'Medium' `
                        -Details "DiskType: $($SHObj.OSDiskType) (Azure managed disks are encrypted by default with platform-managed keys - CMK not evaluated)" `
                        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-machines/disk-encryption-overview'))
                }

                # NOTE: Availability Zone deployment (BCDR-AZ) is evaluated per host pool after the
                # session-host loop (zone SPREAD across hosts), not per VM (audit C-4).

                # ─── CHECK: OS Disk SSD ───
                if ($SHObj.OSDiskType) {
                    $IsSSD = $SHObj.OSDiskType -match 'SSD|Premium'
                    [void]$AllChecks.Add((New-CheckResult -Id "SH-SSD-$VMName" `
                        -Category 'Session Hosts' -Name 'OS Disk SSD Type' `
                        -Description 'Production session hosts should use SSD, not HDD' `
                        -Status $(if ($IsSSD) { 'Pass' } else { 'Warning' }) `
                        -Severity 'Medium' `
                        -Details "DiskType: $($SHObj.OSDiskType)" `
                        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-machines/disks-types'))
                }

                # ─── CHECK: VM Heartbeat / Token Risk ───
                if ($SH.LastHeartBeat) {
                    # LastHeartBeat is reported in UTC - compare against UTC now (C-9).
                    $DaysSinceHB = ((Get-Date).ToUniversalTime() - [DateTime]$SH.LastHeartBeat).Days
                    $HBStatus = if ($DaysSinceHB -gt 60) { 'Fail' } elseif ($DaysSinceHB -gt 30) { 'Warning' } else { 'Pass' }
                    [void]$AllChecks.Add((New-CheckResult -Id "OPS-HB-$VMName" `
                        -Category 'Governance & Cost' -Name 'VM Heartbeat / Token Risk' `
                        -Description 'VMs without heartbeat >60 days risk token expiration' `
                        -Status $HBStatus -Severity 'Medium' `
                        -Details "LastHeartBeat: $($SH.LastHeartBeat), DaysAgo: $DaysSinceHB" `
                        -Reference 'https://learn.microsoft.com/en-us/azure/well-architected/azure-virtual-desktop/operations'))
                }

                # ─── CHECK: Custom Image vs Marketplace ───
                if ($SHObj.ImageReference) {
                    $IsCustom = -not $SHObj.ImageReference.Publisher
                    [void]$AllChecks.Add((New-CheckResult -Id "SH-IMG-$VMName" `
                        -Category 'Session Hosts' -Name 'Custom Image Used' `
                        -Description 'Custom golden images recommended over marketplace for consistency' `
                        -Status $(if ($IsCustom) { 'Pass' } else { 'Warning' }) `
                        -Severity 'Low' `
                        -Details "Image: $(if ($IsCustom) { 'Custom/Gallery' } else { "$($SHObj.ImageReference.Publisher)/$($SHObj.ImageReference.Offer)" })" `
                        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/set-up-customize-master-image'))

                    # ─── CHECK: Gallery Image Version Freshness (SH-024) ───
                    if ($IsCustom -and $SHObj.ImageReference.Id) {
                        try {
                            $ImgId = $SHObj.ImageReference.Id
                            # Parse gallery image version ID: /subscriptions/.../galleries/.../images/.../versions/...
                            if ($ImgId -match '/galleries/(?<gal>[^/]+)/images/(?<img>[^/]+)/versions/(?<ver>[^/]+)') {
                                $GalRG = ($ImgId -split '/')[4]
                                $GalName = $Matches['gal']; $ImgName = $Matches['img']; $VerName = $Matches['ver']
                                $GalImgVer = Get-AzGalleryImageVersion -ResourceGroupName $GalRG -GalleryName $GalName `
                                    -GalleryImageDefinitionName $ImgName -Name $VerName -ErrorAction SilentlyContinue
                                if ($GalImgVer) {
                                    $PubDate = $GalImgVer.PublishingProfile.PublishedDate
                                    $AgeDays = if ($PubDate) { [math]::Round(((Get-Date) - $PubDate).TotalDays, 0) } else { -1 }
                                    $ImgEntry = [PSCustomObject]@{
                                        Gallery = $GalName; Image = $ImgName; Version = $VerName
                                        PublishedDate = $PubDate; AgeDays = $AgeDays; UsedBy = $VMName
                                    }
                                    # Avoid duplicate gallery entries
                                    if (-not ($Discovery.Inventory.GalleryImageVersions | Where-Object { $_.Gallery -eq $GalName -and $_.Version -eq $VerName })) {
                                        $Discovery.Inventory.GalleryImageVersions += $ImgEntry

                                        # ─── CHECK: Image Replication (BCDR-006) — replicated target regions ───
                                        $TargetRegions = @($GalImgVer.PublishingProfile.TargetRegions)
                                        $RegionCount = $TargetRegions.Count
                                        [void]$AllChecks.Add((New-CheckResult -Id "BCDR-IMGREP-$GalName-$VerName" `
                                            -Category 'BCDR' -Name 'Image Replication' `
                                            -Description 'Golden images should be replicated to secondary region(s) so DR host pools can deploy during an outage' `
                                            -Status $(if ($RegionCount -gt 1) { 'Pass' } else { 'Warning' }) `
                                            -Severity 'Medium' `
                                            -Details "Gallery image $GalName/$ImgName v$VerName replicated to $RegionCount region(s): $(@($TargetRegions | ForEach-Object { $_.Name }) -join ', ')" `
                                            -Recommendation 'Replicate golden images to at least one secondary region via Azure Compute Gallery so DR deployment is not blocked during a regional outage.' `
                                            -Reference 'https://learn.microsoft.com/en-us/azure/virtual-machines/azure-compute-gallery' `
                                            -Evidence @{ Gallery = $GalName; Version = $VerName; RegionCount = $RegionCount }))
                                    }
                                    $Stale = $AgeDays -gt 90
                                    [void]$AllChecks.Add((New-CheckResult -Id "SH-IMGFRESH-$VMName" `
                                        -Category 'Session Hosts' -Name 'Image Version Freshness' `
                                        -Description 'Gallery image versions should be less than 90 days old' `
                                        -Status $(if ($AgeDays -lt 0) { 'Warning' } elseif ($Stale) { 'Fail' } else { 'Pass' }) `
                                        -Severity 'High' `
                                        -Details "Image: $GalName/$ImgName v$VerName, Published: $PubDate, Age: ${AgeDays}d" `
                                        -Recommendation 'Update gallery image version via Azure Image Builder on a monthly cadence.' `
                                        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-machines/shared-image-galleries' `
                                        -Evidence @{ VM = $VMName; AgeDays = $AgeDays; Gallery = $GalName }))
                                }
                            }
                        } catch {
                            Write-Status "    Could not read gallery image version for $VMName : $($_.Exception.Message)" -Level 'WARN'
                        }
                    }
                }

                # ─── CHECK: VM Sizing (B-2 - emit for every session host) ───
                if ($SHObj.VMSize) {
                    $IsBSeries = $SHObj.VMSize -match '^Standard_B'
                    $SizeSpec  = Get-VMSizeSpec -VMSize $SHObj.VMSize -Location $SHObj.Location -Cache $VMSizeCache
                    $IsMultiSession = $HP.HostPoolType -eq 'Pooled'
                    $SizeStatus = 'Pass'
                    $SizeDetail = "VMSize: $($SHObj.VMSize)$(if ($SizeSpec.VCPU) { " ($($SizeSpec.VCPU) vCPU, $($SizeSpec.MemoryGB) GB)" })"
                    $SizeRec    = 'VM size meets Microsoft multi-session sizing guidance.'
                    if ($IsBSeries) {
                        $SizeStatus = 'Warning'
                        $SizeRec = 'B-series is burstable - use D-series or E-series for production pooled host pools.'
                        $SizeDetail += ' - burstable B-series (inconsistent performance for shared desktops)'
                    } elseif ($IsMultiSession -and $SizeSpec.VCPU -ne $null -and ($SizeSpec.VCPU -lt 4 -or $SizeSpec.MemoryGB -lt 16)) {
                        $SizeStatus = 'Warning'
                        $SizeRec = 'Below Microsoft multi-session minimum guidance (4+ vCPU, 16+ GB RAM). Increase VM size.'
                        $SizeDetail += ' - below multi-session minimum (4 vCPU / 16 GB)'
                    }
                    [void]$AllChecks.Add((New-CheckResult -Id "SH-BSERIES-$VMName" `
                        -Category 'Session Hosts' -Name 'VM Sizing' `
                        -Description 'Session host VM size should meet workload and multi-session guidance' `
                        -Status $SizeStatus -Severity 'Medium' `
                        -Details $SizeDetail `
                        -Recommendation $SizeRec `
                        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-machines/sizes' `
                        -Evidence @{ VM = $VMName; VMSize = $SHObj.VMSize; VCPU = $SizeSpec.VCPU; MemoryGB = $SizeSpec.MemoryGB }))
                }

                # ─── CHECK: Entra Join Type (from VM extensions) ───
                [void]$AllChecks.Add((New-CheckResult -Id "IAM-JOIN-$VMName" `
                    -Category 'Identity & Access' -Name 'Entra ID Join Type' `
                    -Description 'Session hosts should use Entra ID or Hybrid join' `
                    -Status $(if (-not $SHObj.JoinDataAvailable) { 'Error' } elseif ($SHObj.JoinType -ne 'Unknown') { 'Pass' } else { 'Warning' }) `
                    -Severity 'High' `
                    -Details "$(if (-not $SHObj.JoinDataAvailable) { 'Could not read VM extension data - join type undetermined.' } else { "JoinType: $($SHObj.JoinType)" })" `
                    -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/prerequisites#identity'))

                # ─── CHECK: Azure Monitor Agent (AMA) ───
                [void]$AllChecks.Add((New-CheckResult -Id "MON-AMA-$VMName" `
                    -Category 'Monitoring' -Name 'Azure Monitor Agent Installed' `
                    -Description 'AMA should be installed on session hosts for AVD Insights telemetry' `
                    -Status $(if ($SHObj.AMAInstalled) { 'Pass' } else { 'Fail' }) `
                    -Severity 'Medium' `
                    -Details "AMA Extension: $(if ($SHObj.AMAInstalled) { 'Installed' } else { 'Not found' })" `
                    -Recommendation 'Install Azure Monitor Agent for AVD Insights and performance monitoring.' `
                    -Reference 'https://learn.microsoft.com/en-us/azure/azure-monitor/agents/agents-overview'))

                # ─── CHECK: Endpoint Protection (MDE) ───
                [void]$AllChecks.Add((New-CheckResult -Id "SEC-MDE-$VMName" `
                    -Category 'Security' -Name 'Endpoint Protection (MDE)' `
                    -Description 'Microsoft Defender for Endpoint should be deployed on session hosts' `
                    -Status $(if ($SHObj.MDEInstalled) { 'Pass' } else { 'Warning' }) `
                    -Severity 'High' `
                    -Details "MDE Extension: $(if ($SHObj.MDEInstalled) { 'Installed' } else { 'Not found (may be deployed via Intune/GPO)' })" `
                    -Recommendation 'Deploy Microsoft Defender for Endpoint via VM extension, Intune, or Defender for Cloud auto-provisioning.' `
                    -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/security-recommendations'))

                # ─── CHECK: OS End-of-Support Risk (SH-029) ───
                if ($SHObj.ImageReference) {
                    $ImgOffer = "$($SHObj.ImageReference.Offer)".ToLower()
                    $ImgSku   = "$($SHObj.ImageReference.Sku)".ToLower()
                    $IsCustomImg = -not $SHObj.ImageReference.Publisher
                    $EolStatus = 'Warning'
                    $EolDetail = "Image: $ImgOffer/$ImgSku"
                    $EolRec    = 'Verify the OS build support lifecycle and plan an upgrade before end-of-support.'
                    if ($IsCustomImg) {
                        $EolStatus = 'Warning'
                        $EolDetail = "Custom/gallery image (offer/sku not marketplace) - verify OS build support status"
                    } elseif ($ImgOffer -match 'windows-11|windows11|win11' -or $ImgSku -match 'win11|windows-11') {
                        $EolStatus = 'Pass'
                        $EolDetail = "Windows 11 image: $ImgOffer/$ImgSku (supported)"
                    } elseif ($ImgOffer -match 'windowsserver' -or $ImgSku -match 'server|datacenter|core-') {
                        if ($ImgSku -match '2022|2025|23h2|2019|2016') {
                            if ($ImgSku -match '2022|2025|23h2') {
                                $EolStatus = 'Pass'; $EolDetail = "Windows Server image: $ImgSku (2022+ supported)"
                            } else {
                                $EolStatus = 'Warning'; $EolDetail = "Windows Server image: $ImgSku (verify support lifecycle; consider Server 2022/2025)"
                            }
                        } else {
                            $EolStatus = 'Warning'; $EolDetail = "Windows Server image: $ImgSku (verify OS build support status)"
                        }
                    } elseif ($ImgOffer -match 'windows-10|windows10|win10' -or $ImgSku -match 'win10|windows-10') {
                        if ($ImgSku -match 'esu') {
                            $EolStatus = 'Warning'
                            $EolDetail = "Windows 10 image with ESU: $ImgSku (Win10 reached end-of-support Oct 2025 - ESU in place, plan migration to Windows 11)"
                        } else {
                            $EolStatus = 'Fail'
                            $EolDetail = "Windows 10 image without ESU: $ImgOffer/$ImgSku (Windows 10 reached end-of-support Oct 2025 - no security updates)"
                            $EolRec    = 'Migrate to Windows 11 multi-session, or enroll in Extended Security Updates (ESU) as an interim measure.'
                        }
                    } else {
                        $EolStatus = 'Warning'
                        $EolDetail = "Unrecognized image: $ImgOffer/$ImgSku - verify OS build support status"
                    }
                    [void]$AllChecks.Add((New-CheckResult -Id "SH-OSEOL-$VMName" `
                        -Category 'Session Hosts' -Name 'OS End-of-Support Risk' `
                        -Description 'Session hosts should run an OS build that is within its support lifecycle (Windows 10 reached end-of-support Oct 2025)' `
                        -Status $EolStatus -Severity 'High' `
                        -Details $EolDetail `
                        -Recommendation $EolRec `
                        -Reference 'https://learn.microsoft.com/en-us/lifecycle/products/windows-10-enterprise-and-education' `
                        -Evidence @{ VM = $VMName; Offer = $ImgOffer; Sku = $ImgSku }))
                }

                # ─── CHECK: GPU Session Host Configuration (SH-030) ───
                if ($SHObj.VMSize -match '^Standard_(NV|NC|NG)') {
                    $HasGpuDriver = $false
                    if ($SHObj.Extensions) {
                        $HasGpuDriver = @($SHObj.Extensions | Where-Object { $_ -match 'NvidiaGpuDriverWindows|AmdGpuDriverWindows' }).Count -gt 0
                    }
                    [void]$AllChecks.Add((New-CheckResult -Id "SH-GPU-$VMName" `
                        -Category 'Session Hosts' -Name 'GPU Session Host Configuration' `
                        -Description 'GPU-enabled session hosts should have the GPU driver extension installed for hardware acceleration' `
                        -Status $(if ($HasGpuDriver) { 'Pass' } else { 'Warning' }) `
                        -Severity 'Medium' `
                        -Details "VMSize: $($SHObj.VMSize), GPUDriverExtension: $(if ($HasGpuDriver) { 'Installed' } else { 'Not found' })" `
                        -Recommendation 'Install the NVIDIA/AMD GPU driver extension and enable GPU acceleration policies for graphics-intensive workloads.' `
                        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/enable-gpu-acceleration' `
                        -Evidence @{ VM = $VMName; VMSize = $SHObj.VMSize; GpuDriver = $HasGpuDriver }))
                }

                # ─── CHECK: Ephemeral OS Disk (Pooled VMs) ───
                if ($HP.HostPoolType -eq 'Pooled' -and $VMModel) {
                    $IsEphemeral = $null -ne $VMModel.StorageProfile.OsDisk.DiffDiskSettings
                    [void]$AllChecks.Add((New-CheckResult -Id "SH-EPHEMERAL-$VMName" `
                        -Category 'Session Hosts' -Name 'Ephemeral OS Disk' `
                        -Description 'Pooled session hosts should use ephemeral OS disks for faster reimage and lower cost' `
                        -Status $(if ($IsEphemeral) { 'Pass' } else { 'Warning' }) `
                        -Severity 'Medium' `
                        -Details "EphemeralDisk: $(if ($IsEphemeral) { 'Yes' } else { 'No - uses persistent managed disk' })" `
                        -Recommendation 'Use ephemeral OS disks for pooled host pools to eliminate storage costs and improve reimage speed.' `
                        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-machines/ephemeral-os-disks'))
                }

                # ─── CHECK: Disk Type Cost Optimization ───
                if ($SHObj.OSDiskType -and $HP.HostPoolType -eq 'Pooled') {
                    $IsPremium = $SHObj.OSDiskType -match 'Premium'
                    $IsStdSSD  = $SHObj.OSDiskType -match 'StandardSSD'
                    [void]$AllChecks.Add((New-CheckResult -Id "GOV-DISKSKU-$VMName" `
                        -Category 'Governance & Cost' -Name 'Disk Type Cost Optimization' `
                        -Description 'Premium SSD on pooled hosts may be unnecessary cost - Standard SSD is often sufficient' `
                        -Status $(if ($IsPremium) { 'Warning' } elseif ($IsStdSSD) { 'Pass' } else { 'Warning' }) `
                        -Severity 'Low' `
                        -Details "DiskType: $($SHObj.OSDiskType), PoolType: $($HP.HostPoolType)" `
                        -Recommendation 'Standard SSD is typically sufficient for pooled hosts; Premium is only needed for heavy I/O workloads.' `
                        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-machines/disks-types'))
                }

                # ─── CHECK: Guest Attestation (Trusted Launch integrity monitoring) ───
                if ($SHObj.TrustedLaunch -and $SHObj.Extensions) {
                    $HasGuestAttest = 'GuestAttestation' -in $SHObj.Extensions
                    [void]$AllChecks.Add((New-CheckResult -Id "SEC-ATTEST-$VMName" `
                        -Category 'Security' -Name 'Guest Attestation Extension' `
                        -Description 'Trusted Launch VMs should have Guest Attestation extension for integrity monitoring' `
                        -Status $(if ($HasGuestAttest) { 'Pass' } else { 'Warning' }) `
                        -Severity 'Medium' `
                        -Details "TrustedLaunch: Yes, GuestAttestation: $(if ($HasGuestAttest) { 'Installed' } else { 'Not installed' })" `
                        -Recommendation 'Install the Guest Attestation extension to enable boot integrity monitoring via Defender for Cloud.' `
                        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-machines/trusted-launch#microsoft-defender-for-cloud-integration' `
                        -Evidence @{ VM = $VMName; GuestAttestation = [bool]$HasGuestAttest }))
                }
                # NOTE: Agent Version Currency (OPS-AGENT) is emitted after all subscriptions are
                # discovered so each host can also be compared against the fleet maximum (C-9).

                # ─── CHECK: Session Host Status ───
                if ($SH.Status) {
                    $SHStatusOK = $SH.Status -eq 'Available'
                    $IsDeallocated = ($SHObj.PowerState -and $SHObj.PowerState -match 'deallocated|stopped')
                    if (-not $SHStatusOK -and $IsDeallocated) {
                        # Deallocated hosts are commonly autoscale off-hours - not a health failure (C-10).
                        [void]$AllChecks.Add((New-CheckResult -Id "SH-STATUS-$VMName" `
                            -Category 'Session Hosts' -Name 'Session Host Health Status' `
                            -Description 'Session host should report Available status' `
                            -Status 'N/A' -Severity 'Medium' `
                            -Details "Status: $($SH.Status), PowerState: $($SHObj.PowerState) (deallocated - may be autoscale off-hours)" `
                            -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/troubleshoot-vm-connectivity'))
                    } else {
                        [void]$AllChecks.Add((New-CheckResult -Id "SH-STATUS-$VMName" `
                            -Category 'Session Hosts' -Name 'Session Host Health Status' `
                            -Description 'Session host should report Available status' `
                            -Status $(if ($SHStatusOK) { 'Pass' } elseif ($SH.Status -eq 'NeedsAssistance') { 'Fail' } else { 'Warning' }) `
                            -Severity $(if ($SH.Status -eq 'NeedsAssistance') { 'High' } else { 'Medium' }) `
                            -Details "Status: $($SH.Status), AllowNewSession: $($SH.AllowNewSession)" `
                            -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/troubleshoot-vm-connectivity'))
                    }
                }
            }

            # ─── CHECK: Availability Zone spread (per host pool, C-4) ───
            $HPHosts = @($Discovery.Inventory.SessionHosts | Where-Object { $_.HostPoolName -eq $HP.Name })
            if ($HPHosts.Count -gt 0) {
                $HPZones = @($HPHosts | ForEach-Object { $_.AvailabilityZone } | Where-Object { $_ } | Sort-Object -Unique)
                if ($HPZones.Count -ge 2) {
                    $AZStatus = 'Pass'; $AZDetail = "Hosts span $($HPZones.Count) zones: $($HPZones -join ',')"
                } elseif ($HPZones.Count -eq 1) {
                    $AZStatus = 'Warning'; $AZDetail = "All hosts pinned to a single zone ($($HPZones -join ',')) - no zone spread"
                } else {
                    $AZStatus = 'Warning'; $AZDetail = 'No availability zones assigned (note: some regions lack AZ support - verify region capability)'
                }
                [void]$AllChecks.Add((New-CheckResult -Id "BCDR-AZ-$($HP.Name)" `
                    -Category 'BCDR' -Name 'Availability Zone Deployment' `
                    -Description 'Session hosts should be spread across availability zones for resilience' `
                    -Status $AZStatus -Severity 'Medium' `
                    -Details $AZDetail `
                    -Recommendation 'Deploy session hosts across multiple availability zones for high availability where the region supports AZs.' `
                    -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/azure-virtual-desktop-fault-domain-mode' `
                    -Evidence @{ HostPool = $HP.Name; Zones = $HPZones; HostCount = $HPHosts.Count }))
            }
        }
    } catch {
        Write-Status "  Error discovering session hosts: $($_.Exception.Message)" -Level 'ERROR'
        $Discovery.Errors += "Session host discovery failed: $($_.Exception.Message)"
    }

    # ─── APPLICATION GROUPS ───────────────────────────────────────────────
    Write-Status "Application Groups" -Level 'SECTION'
    try {
        $AppGroups = @(Get-AzWvdApplicationGroup -ErrorAction Stop)
        Write-Status "  Found $($AppGroups.Count) app group(s)" -Level 'SUCCESS'

        # Group app-group types by host pool to detect Desktop+RemoteApp mixing (B-1).
        $AGTypesByHostPool = @{}
        foreach ($AG in $AppGroups) {
            $HpPath = "$($AG.HostPoolArmPath)".ToLower()
            if (-not $AGTypesByHostPool.ContainsKey($HpPath)) { $AGTypesByHostPool[$HpPath] = @() }
            $AGTypesByHostPool[$HpPath] += $AG.ApplicationGroupType
        }

        foreach ($AG in $AppGroups) {
            $Discovery.Inventory.AppGroups += [PSCustomObject]@{
                SubscriptionId       = $SubId
                ResourceGroup        = ($AG.Id -split '/')[4]
                Name                 = $AG.Name
                Id                   = $AG.Id
                ApplicationGroupType = $AG.ApplicationGroupType
                HostPoolArmPath      = $AG.HostPoolArmPath
                Location             = $AG.Location
                Tags                 = $AG.Tag
            }

            # ─── CHECK: App group configuration (B-1 - real evaluation) ───
            $AGStatus = 'Pass'
            $AGRec    = 'Application group is configured appropriately.'
            $HpTypes  = @($AGTypesByHostPool["$($AG.HostPoolArmPath)".ToLower()] | Sort-Object -Unique)
            $IsMixed  = (@($HpTypes | Where-Object { $_ -eq 'Desktop' }).Count -gt 0) -and (@($HpTypes | Where-Object { $_ -eq 'RemoteApp' }).Count -gt 0)
            $AppCount = $null
            if ($AG.ApplicationGroupType -eq 'RemoteApp') {
                try {
                    $Apps = @(Get-AzWvdApplication -ResourceGroupName (($AG.Id -split '/')[4]) -ApplicationGroupName $AG.Name -ErrorAction Stop)
                    $AppCount = $Apps.Count
                    if ($AppCount -eq 0) {
                        $AGStatus = 'Warning'
                        $AGRec = 'RemoteApp group has no published applications - add applications or remove the empty group.'
                    }
                } catch {
                    $AGStatus = 'Error'
                    $AGRec = "Could not enumerate applications: $($_.Exception.Message)"
                }
            }
            if ($IsMixed -and $AGStatus -eq 'Pass') {
                $AGStatus = 'Warning'
                $AGRec = 'Host pool mixes Desktop and RemoteApp application groups - separate them into distinct host pools.'
            }
            [void]$AllChecks.Add((New-CheckResult -Id "APP-CFG-$($AG.Name)" `
                -Category 'Application Delivery' -Name 'App Group Configuration' `
                -Description 'Application groups should be configured appropriately' `
                -Status $AGStatus -Severity 'Medium' `
                -Details "Type: $($AG.ApplicationGroupType), HostPool: $(($AG.HostPoolArmPath -split '/')[-1])$(if ($null -ne $AppCount) { ", Apps: $AppCount" }), HostPoolMixesTypes: $IsMixed" `
                -Recommendation $AGRec `
                -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/manage-app-groups' `
                -Evidence @{ AppGroup = $AG.Name; Type = $AG.ApplicationGroupType; AppCount = $AppCount; Mixed = $IsMixed }))
        }
    } catch {
        Write-Status "  Error: $($_.Exception.Message)" -Level 'ERROR'
        $Discovery.Errors += "App group discovery failed: $($_.Exception.Message)"
    }

    # ─── WORKSPACES ───────────────────────────────────────────────────────
    Write-Status "Workspaces" -Level 'SECTION'
    try {
        $Workspaces = @(Get-AzWvdWorkspace -ErrorAction Stop)
        Write-Status "  Found $($Workspaces.Count) workspace(s)" -Level 'SUCCESS'

        foreach ($WS in $Workspaces) {
            $Discovery.Inventory.Workspaces += [PSCustomObject]@{
                SubscriptionId       = $SubId
                ResourceGroup        = ($WS.Id -split '/')[4]
                Name                 = $WS.Name
                Id                   = $WS.Id
                ApplicationGroupReferences = $WS.ApplicationGroupReference
                Location             = $WS.Location
                Tags                 = $WS.Tag
            }
        }
    } catch {
        Write-Status "  Error: $($_.Exception.Message)" -Level 'ERROR'
        $Discovery.Errors += "Workspace discovery failed: $($_.Exception.Message)"
    }

    # ─── SCALING PLANS ────────────────────────────────────────────────────
    Write-Status "Scaling Plans" -Level 'SECTION'
    try {
        $ScalingPlans = @(Get-AzWvdScalingPlan -ErrorAction Stop)
        Write-Status "  Found $($ScalingPlans.Count) scaling plan(s)" -Level 'SUCCESS'

        foreach ($SP in $ScalingPlans) {
            $Discovery.Inventory.ScalingPlans += [PSCustomObject]@{
                SubscriptionId     = $SubId
                ResourceGroup      = ($SP.Id -split '/')[4]
                Name               = $SP.Name
                Id                 = $SP.Id
                HostPoolReferences = $SP.HostPoolReference
                Schedules          = $SP.Schedule
                Location           = $SP.Location
                Tags               = $SP.Tag
                TimeZone           = $SP.TimeZone
                HostPoolType       = $SP.HostPoolType
            }
        }

        # ─── CHECK: Scaling plan coverage + Load Balancing Algorithm (B-1) ───
        $HPsWithScaling = @($ScalingPlans | ForEach-Object { $_.HostPoolReference.HostPoolArmPath } | Where-Object { $_ })
        foreach ($HP in $HostPools) {
            $HasScaling = $HPsWithScaling -contains $HP.Id
            [void]$AllChecks.Add((New-CheckResult -Id "GOV-SCALE-$($HP.Name)" `
                -Category 'Governance & Cost' -Name 'Scaling Plan Assigned' `
                -Description 'Host pools should have a scaling plan for cost optimization' `
                -Status $(if ($HasScaling) { 'Pass' } else { 'Fail' }) `
                -Severity 'High' `
                -Details "ScalingPlan: $(if ($HasScaling) { 'Assigned' } else { 'Not assigned' })" `
                -Recommendation 'Create and assign a scaling plan to reduce compute costs during off-peak hours.' `
                -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/autoscale-scaling-plan' `
                -Evidence @{ HostPool = $HP.Name; HasScalingPlan = $HasScaling }))

            # ─── CHECK: Load Balancing Algorithm (moved here so scaling-plan assignment is known) ───
            if ($HasScaling) {
                $LBStatus = 'Pass'
                $LBDetail = "Algorithm: $($HP.LoadBalancerType), PoolType: $($HP.HostPoolType). A scaling plan is assigned - autoscale overrides the LB algorithm during scheduled hours."
                $LBRec    = 'Scaling plan governs ramp behavior; LB algorithm applies outside scheduled hours.'
            } elseif ($HP.LoadBalancerType -eq 'DepthFirst') {
                $LBStatus = 'Warning'
                $LBDetail = "Algorithm: DepthFirst, PoolType: $($HP.HostPoolType). Session packing without a scaling plan gives no autoscale ramp guidance."
                $LBRec    = 'Assign a scaling plan, or use BreadthFirst if you need even distribution during ramp-up.'
            } else {
                $LBStatus = 'Pass'
                $LBDetail = "Algorithm: $($HP.LoadBalancerType), PoolType: $($HP.HostPoolType)."
                $LBRec    = 'BreadthFirst distributes sessions evenly during ramp-up.'
            }
            [void]$AllChecks.Add((New-CheckResult -Id "SH-LB-$($HP.Name)" `
                -Category 'Session Hosts' -Name 'Load Balancing Algorithm' `
                -Description 'LB should align with workload (BreadthFirst for ramp-up, DepthFirst for cost)' `
                -Status $LBStatus -Severity 'Medium' `
                -Details $LBDetail `
                -Recommendation $LBRec `
                -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/host-pool-load-balancing' `
                -Evidence @{ HostPool = $HP.Name; LoadBalancer = $HP.LoadBalancerType; HasScalingPlan = $HasScaling }))
        }

        # CHECK: Scaling plan active - must be ENABLED on a host-pool reference, not just have schedules (C-5)
        foreach ($SP in $ScalingPlans) {
            $HasSchedules = $SP.Schedule -and $SP.Schedule.Count -gt 0
            $EnabledRefs  = @($SP.HostPoolReference | Where-Object { $_.ScalingPlanEnabled })
            $AnyEnabled   = $EnabledRefs.Count -gt 0
            if ($AnyEnabled) {
                $SPActiveStatus = 'Pass'
                $SPActiveDetail = "Enabled on $($EnabledRefs.Count) host pool reference(s), Schedules: $(if ($HasSchedules) { $SP.Schedule.Count } else { 0 })"
            } elseif ($HasSchedules) {
                $SPActiveStatus = 'Warning'
                $SPActiveDetail = "Schedules defined ($($SP.Schedule.Count)) but ScalingPlanEnabled is false on all host pool references - plan is inactive."
            } else {
                $SPActiveStatus = 'Warning'
                $SPActiveDetail = 'No schedules and not enabled on any host pool reference.'
            }
            [void]$AllChecks.Add((New-CheckResult -Id "GOV-SPACTIVE-$($SP.Name)" `
                -Category 'Governance & Cost' -Name 'Scaling Plan Active' `
                -Description 'Scaling plans should be enabled with active schedules' `
                -Status $SPActiveStatus -Severity 'High' `
                -Details $SPActiveDetail `
                -Recommendation 'Enable the scaling plan on its host pool references (ScalingPlanEnabled) and confirm schedules exist.' `
                -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/autoscale-scaling-plan' `
                -Evidence @{ ScalingPlan = $SP.Name; EnabledReferences = $EnabledRefs.Count; ScheduleCount = $(if ($HasSchedules) { $SP.Schedule.Count } else { 0 }) }))
        }
        # NOTE: BCDR-SPSCHED emit removed (audit M-3 - mis-mapped target; schedule data now feeds SH-LB / GOV-SPACTIVE).

        # ─── CHECK: Personal Pool Autoscale & Hibernate (SH-031) — Personal host pools only ───
        foreach ($HP in @($HostPools | Where-Object { $_.HostPoolType -eq 'Personal' })) {
            $PersonalPlan = @($ScalingPlans | Where-Object {
                "$($_.HostPoolType)" -eq 'Personal' -and
                (@($_.HostPoolReference | Where-Object { "$($_.HostPoolArmPath)".ToLower() -eq "$($HP.Id)".ToLower() }).Count -gt 0)
            })
            [void]$AllChecks.Add((New-CheckResult -Id "SH-PERSAUTO-$($HP.Name)" `
                -Category 'Session Hosts' -Name 'Personal Pool Autoscale & Hibernate' `
                -Description 'Personal host pools should use a Personal-type scaling plan to deallocate/hibernate idle hosts for cost savings' `
                -Status $(if ($PersonalPlan.Count -gt 0) { 'Pass' } else { 'Warning' }) `
                -Severity 'Medium' `
                -Details "Personal host pool $($HP.Name): Personal scaling plan $(if ($PersonalPlan.Count -gt 0) { "assigned ($(@($PersonalPlan | ForEach-Object { $_.Name }) -join ', '))" } else { 'not assigned' })" `
                -Recommendation 'Assign a Personal-type scaling plan (personal autoscale is GA); hibernate reduces cost further — note FSLogix/App Attach incompatibility with hibernate.' `
                -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/autoscale-scenarios' `
                -Evidence @{ HostPool = $HP.Name; PersonalPlan = $PersonalPlan.Count }))
        }

        # ─── CHECK: Scaling Plan Diagnostics (MON-017) — only when scaling plans exist ───
        foreach ($SP in $ScalingPlans) {
            try {
                $SPDiag = @(Get-AzDiagnosticSetting -ResourceId $SP.Id -ErrorAction Stop -WarningAction SilentlyContinue)
                [void]$AllChecks.Add((New-CheckResult -Id "MON-SPDIAG-$($SP.Name)" `
                    -Category 'Monitoring' -Name 'Scaling Plan Diagnostics' `
                    -Description 'Scaling plans should have diagnostic settings enabled for an autoscale audit trail' `
                    -Status $(if ($SPDiag.Count -gt 0) { 'Pass' } else { 'Warning' }) `
                    -Severity 'Low' `
                    -Details "ScalingPlan $($SP.Name): diagnostic settings: $($SPDiag.Count)" `
                    -Recommendation 'Enable diagnostic settings on scaling plans to capture autoscale evaluation and action logs.' `
                    -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/autoscale-diagnostics' `
                    -Evidence @{ ScalingPlan = $SP.Name; DiagCount = $SPDiag.Count }))
            } catch {
                [void]$AllChecks.Add((New-CheckResult -Id "MON-SPDIAG-$($SP.Name)" `
                    -Category 'Monitoring' -Name 'Scaling Plan Diagnostics' `
                    -Description 'Scaling plans should have diagnostic settings enabled' `
                    -Status 'Error' -Severity 'Low' `
                    -Details "Could not read scaling plan diagnostic settings: $($_.Exception.Message)" `
                    -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/autoscale-diagnostics'))
            }
        }
    } catch {
        Write-Status "  Error: $($_.Exception.Message)" -Level 'ERROR'
        $Discovery.Errors += "Scaling plan discovery failed: $($_.Exception.Message)"
    }
    # NOTE: BCDR-MULTIREGION is emitted once after all subscriptions (audit A-1) so host-pool regions
    # are aggregated across the whole estate rather than per subscription.

    # ─── NETWORKING ───────────────────────────────────────────────────────
    Write-Status "Networking" -Level 'SECTION'
    try {
        # Determine whether the whole estate is cloud-native Entra-joined (used by NET-DNS, C-7).
        $EstateHosts = @($Discovery.Inventory.SessionHosts | Where-Object { $_.JoinDataAvailable })
        $AllEntraJoined = ($EstateHosts.Count -gt 0) -and (@($EstateHosts | Where-Object { $_.JoinType -ne 'Entra ID' }).Count -eq 0)

        # Collect unique VNets from session host NICs. Reuse NIC facts cached during the session-host
        # pass (E-7) instead of re-fetching VM + NIC per host.
        $DiscoveredVNetIds = @{}
        foreach ($SH in $Discovery.Inventory.SessionHosts) {
            $VMName = if ($SH.ResourceId) { ($SH.ResourceId -split '/')[-1] } else { $SH.Name }
            $SubnetId = $SH.NicSubnetId
            if (-not $SubnetId) { continue }
            $VNetId = ($SubnetId -split '/subnets/')[0]
            if (-not $DiscoveredVNetIds.ContainsKey($VNetId)) {
                $DiscoveredVNetIds[$VNetId] = $true
            }
            # CHECK: Public IP on session host
            $HasPublicIP = [bool]$SH.NicHasPublicIP
            [void]$AllChecks.Add((New-CheckResult -Id "NET-PIP-$VMName" `
                -Category 'Networking' -Name 'No Public IP on Session Host' `
                -Description 'Session hosts should not have public IP addresses' `
                -Status $(if ($HasPublicIP) { 'Fail' } else { 'Pass' }) `
                -Severity 'Critical' `
                -Details "PublicIP: $(if ($HasPublicIP) { 'ASSIGNED' } else { 'None' })" `
                -Recommendation 'Remove public IPs from session hosts. Use Azure Bastion or JIT for management access.' `
                -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/security-guide'))
        }

        # Get VNet details
        foreach ($VNetId in $DiscoveredVNetIds.Keys) {
            try {
                $VNetRG   = ($VNetId -split '/')[4]
                $VNetName = ($VNetId -split '/')[-1]
                $VNet = Get-AzVirtualNetwork -ResourceGroupName $VNetRG -Name $VNetName -ErrorAction Stop

                $VNetObj = [PSCustomObject]@{
                    Name          = $VNet.Name
                    Id            = $VNet.Id
                    ResourceGroup = $VNetRG
                    AddressSpace  = $VNet.AddressSpace.AddressPrefixes
                    Subnets       = @($VNet.Subnets | ForEach-Object {
                        [PSCustomObject]@{
                            Name          = $_.Name
                            AddressPrefix = $_.AddressPrefix
                            NSG           = if ($_.NetworkSecurityGroup) { $_.NetworkSecurityGroup.Id } else { $null }
                            RouteTable    = if ($_.RouteTable) { $_.RouteTable.Id } else { $null }
                        }
                    })
                    Peerings      = @($VNet.VirtualNetworkPeerings | ForEach-Object {
                        [PSCustomObject]@{
                            Name           = $_.Name
                            RemoteVNet     = $_.RemoteVirtualNetwork.Id
                            PeeringState   = $_.PeeringState
                            AllowForwarded = $_.AllowForwardedTraffic
                            AllowGateway   = $_.AllowGatewayTransit
                        }
                    })
                    DnsServers    = $VNet.DhcpOptions.DnsServers
                    HasPeering    = $VNet.VirtualNetworkPeerings.Count -gt 0
                    Location      = $VNet.Location
                    Tags          = $VNet.Tag
                }
                $Discovery.Inventory.VNets += $VNetObj

                # CHECK: Custom DNS - cloud-native Entra-joined estates can safely use default Azure DNS (C-7)
                $HasCustomDns = $VNet.DhcpOptions -and $VNet.DhcpOptions.DnsServers -and $VNet.DhcpOptions.DnsServers.Count -gt 0
                if ($HasCustomDns) {
                    $DnsStatus = 'Pass'; $DnsDetail = "DNS: $($VNet.DhcpOptions.DnsServers -join ', ')"
                } elseif ($AllEntraJoined) {
                    $DnsStatus = 'Pass'; $DnsDetail = 'DNS: Azure Default (acceptable - all session hosts are Entra-joined, no AD DS/Hybrid resolution required)'
                } else {
                    $DnsStatus = 'Warning'; $DnsDetail = 'DNS: Azure Default (AD DS/Hybrid hosts need custom DNS pointing to domain controllers)'
                }
                [void]$AllChecks.Add((New-CheckResult -Id "NET-DNS-$VNetName" `
                    -Category 'Networking' -Name 'Custom DNS Configuration' `
                    -Description 'VNets with AD-joined session hosts should use custom DNS pointing to domain controllers' `
                    -Status $DnsStatus -Severity 'Medium' `
                    -Details $DnsDetail `
                    -Recommendation 'Configure custom DNS pointing to domain controllers for AD DS/Hybrid; default Azure DNS is fine for pure Entra join.'))

                # CHECK: NSG on AVD subnets - exclude reserved/gateway subnets from the denominator (C-3)
                $ReservedSubnets = @('GatewaySubnet','AzureFirewallSubnet','AzureFirewallManagementSubnet','AzureBastionSubnet','RouteServerSubnet')
                $EvalSubnets = @($VNet.Subnets | Where-Object { $_.Name -notin $ReservedSubnets })
                $SubnetsWithoutNSG = @($EvalSubnets | Where-Object { -not $_.NetworkSecurityGroup })
                $TotalSubnets = $EvalSubnets.Count
                $HasNSGCoverage = $SubnetsWithoutNSG.Count -eq 0 -and $TotalSubnets -gt 0
                [void]$AllChecks.Add((New-CheckResult -Id "NET-NSG-$VNetName" `
                    -Category 'Networking' -Name 'NSG on AVD Subnets' `
                    -Description 'All AVD subnets should have Network Security Groups applied' `
                    -Status $(if ($HasNSGCoverage) { 'Pass' } elseif ($TotalSubnets -eq 0) { 'N/A' } else { 'Warning' }) `
                    -Severity 'High' `
                    -Details "Subnets (excluding reserved): $TotalSubnets, $($SubnetsWithoutNSG.Count) without NSG$(if ($SubnetsWithoutNSG.Count -gt 0) { " ($( ($SubnetsWithoutNSG | ForEach-Object { $_.Name }) -join ', '))" })" `
                    -Recommendation 'Apply NSGs to all AVD subnets for network traffic filtering.' `
                    -Reference 'https://learn.microsoft.com/en-us/azure/virtual-network/network-security-groups-overview'))

                # Analyze route tables for a 0.0.0.0/0 hop to an NVA/firewall/gateway (C-7). Cache by RouteTable Id.
                $HasDefaultRouteToNVA = $false
                $RTNextHops = @()
                foreach ($SubnetEntry in ($EvalSubnets | Where-Object { $_.RouteTable })) {
                    $RTId = $SubnetEntry.RouteTable.Id
                    if (-not $RTId) { continue }
                    try {
                        $RT = Get-AzRouteTable -ResourceGroupName (($RTId -split '/')[4]) -Name (($RTId -split '/')[-1]) -ErrorAction Stop
                        foreach ($Route in @($RT.Routes)) {
                            if ($Route.AddressPrefix -eq '0.0.0.0/0' -and $Route.NextHopType -in @('VirtualAppliance','VirtualNetworkGateway')) {
                                $HasDefaultRouteToNVA = $true
                                $RTNextHops += $Route.NextHopType
                            }
                        }
                    } catch {
                        Write-Status "    Could not read route table $(($RTId -split '/')[-1]): $($_.Exception.Message)" -Level 'WARN'
                    }
                }

                # CHECK: Route table default route - Pass requires 0.0.0.0/0 to VirtualAppliance/Gateway (C-7)
                [void]$AllChecks.Add((New-CheckResult -Id "NET-UDR-$VNetName" `
                    -Category 'Networking' -Name 'Route Table on AVD Subnets' `
                    -Description 'UDR should force traffic through firewall/NVA for inspection' `
                    -Status $(if ($HasDefaultRouteToNVA) { 'Pass' } else { 'Warning' }) `
                    -Severity 'Medium' `
                    -Details "$(if ($HasDefaultRouteToNVA) { "0.0.0.0/0 route to $($RTNextHops -join ',') present" } else { 'No 0.0.0.0/0 route to a VirtualAppliance/Gateway found' })" `
                    -Recommendation 'Add a 0.0.0.0/0 user-defined route with next hop VirtualAppliance (firewall/NVA) to force egress inspection.' `
                    -Reference 'https://learn.microsoft.com/en-us/azure/well-architected/azure-virtual-desktop/networking'))

                # CHECK: NAT Gateway - soften to Pass when a firewall/NVA egress design exists (C-7)
                $SubnetsWithNAT = @($EvalSubnets | Where-Object { $_.NatGateway })
                if ($SubnetsWithNAT.Count -gt 0) {
                    $NatStatus = 'Pass'; $NatDetail = "SubnetsWithNATGW: $($SubnetsWithNAT.Count)/$TotalSubnets"
                } elseif ($HasDefaultRouteToNVA) {
                    $NatStatus = 'Pass'; $NatDetail = 'No NAT Gateway, but firewall/NVA egress design detected (0.0.0.0/0 UDR to VirtualAppliance/Gateway)'
                } else {
                    $NatStatus = 'Warning'; $NatDetail = "SubnetsWithNATGW: 0/$TotalSubnets (no explicit outbound method detected)"
                }
                [void]$AllChecks.Add((New-CheckResult -Id "NET-NATGW-$VNetName" `
                    -Category 'Networking' -Name 'NAT Gateway for Outbound' `
                    -Description 'Private subnets should use NAT Gateway for explicit outbound connectivity' `
                    -Status $NatStatus -Severity 'Medium' `
                    -Details $NatDetail))

                # CHECK: Subnet IP capacity
                foreach ($SubnetEntry in $VNet.Subnets) {
                    $SubPrefix = $SubnetEntry.AddressPrefix
                    if ($SubPrefix -is [array]) { $SubPrefix = $SubPrefix[0] }
                    if ($SubPrefix -match '/(\d+)$') {
                        $CidrBits = [int]$Matches[1]
                        $TotalAvailIPs = [math]::Pow(2, 32 - $CidrBits) - 5  # Azure reserves 5
                        $UsedIPs = ($SubnetEntry.IpConfigurations | Measure-Object).Count
                        $UtilPct = if ($TotalAvailIPs -gt 0) { [math]::Round($UsedIPs / $TotalAvailIPs * 100, 1) } else { 0 }
                        $CapStatus = if ($UtilPct -gt 80) { 'Fail' } elseif ($UtilPct -gt 70) { 'Warning' } else { 'Pass' }
                        [void]$AllChecks.Add((New-CheckResult -Id "NET-SUBCAP-$VNetName-$($SubnetEntry.Name)" `
                            -Category 'Networking' -Name 'Subnet IP Capacity' `
                            -Description 'AVD subnets should have sufficient IP address headroom for scaling' `
                            -Status $CapStatus -Severity 'High' `
                            -Details "Subnet: $($SubnetEntry.Name), CIDR: $SubPrefix, Used: $UsedIPs/$([int]$TotalAvailIPs) ($UtilPct%)" `
                            -Recommendation 'Ensure at least 30% IP headroom for scaling and maintenance. Consider expanding the subnet or adding additional subnets.' `
                            -Reference 'https://learn.microsoft.com/en-us/azure/virtual-network/virtual-networks-faq' `
                            -Evidence @{ VNet = $VNetName; Subnet = $SubnetEntry.Name; CIDR = $SubPrefix; Used = $UsedIPs; Total = [int]$TotalAvailIPs; Utilization = $UtilPct }))
                    }
                }

                # CHECK: VNet peering health
                foreach ($Peering in $VNet.VirtualNetworkPeerings) {
                    $PeerState = $Peering.PeeringState
                    $RemoteVNet = ($Peering.RemoteVirtualNetwork.Id -split '/')[-1]
                    [void]$AllChecks.Add((New-CheckResult -Id "NET-PEER-$($Peering.Name)" `
                        -Category 'Networking' -Name 'VNet Peering Health' `
                        -Description 'VNet peerings should be in Connected state for network connectivity' `
                        -Status $(if ($PeerState -eq 'Connected') { 'Pass' } else { 'Fail' }) `
                        -Severity 'High' `
                        -Details "Peering: $($Peering.Name), State: $PeerState, RemoteVNet: $RemoteVNet, AllowForwarded: $($Peering.AllowForwardedTraffic), AllowGateway: $($Peering.AllowGatewayTransit)" `
                        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-network/virtual-network-peering-overview'))
                }
            } catch {
                Write-Status "    Could not get full VNet details (Az.Network bug): $($_.Exception.Message)" -Level 'WARN'
                # Fallback: capture basic VNet info via ARM so checks still have something
                try {
                    $FallbackVNet = Get-AzResource -ResourceId $VNetId -ExpandProperties -ErrorAction Stop
                    $FbProps = $FallbackVNet.Properties
                    if (-not ($Discovery.Inventory.VNets | Where-Object { $_.Id -eq $VNetId })) {
                    $Discovery.Inventory.VNets += [PSCustomObject]@{
                        Name          = $FallbackVNet.Name
                        Id            = $FallbackVNet.ResourceId
                        ResourceGroup = $VNetRG
                        AddressSpace  = @($FbProps.addressSpace.addressPrefixes)
                        Subnets       = @($FbProps.subnets | ForEach-Object {
                            [PSCustomObject]@{
                                Name          = $_.name
                                AddressPrefix = $_.properties.addressPrefix
                                NSG           = if ($_.properties.networkSecurityGroup) { $_.properties.networkSecurityGroup.id } else { $null }
                                RouteTable    = if ($_.properties.routeTable) { $_.properties.routeTable.id } else { $null }
                            }
                        })
                        Peerings      = @($FbProps.virtualNetworkPeerings | ForEach-Object {
                            [PSCustomObject]@{
                                Name           = $_.name
                                RemoteVNet     = $_.properties.remoteVirtualNetwork.id
                                PeeringState   = $_.properties.peeringState
                                AllowForwarded = $_.properties.allowForwardedTraffic
                                AllowGateway   = $_.properties.allowGatewayTransit
                            }
                        })
                        DnsServers    = @($FbProps.dhcpOptions.dnsServers)
                        HasPeering    = @($FbProps.virtualNetworkPeerings).Count -gt 0
                        Location      = $FallbackVNet.Location
                    }
                    }
                    Write-Status "    Recovered VNet info via ARM fallback for $VNetName" -Level 'WARN'
                } catch {
                    Write-Status "    ARM fallback also failed for $VNetName`: $($_.Exception.Message)" -Level 'WARN'
                }
            }
        }

        # Get NSGs on AVD subnets
        $DiscoveredNSGs = @{}
        foreach ($VNetEntry in $Discovery.Inventory.VNets) {
            foreach ($Subnet in $VNetEntry.Subnets) {
                if ($Subnet.NSG -and -not $DiscoveredNSGs.ContainsKey($Subnet.NSG)) {
                    $DiscoveredNSGs[$Subnet.NSG] = $true
                    try {
                        $NSGRG   = ($Subnet.NSG -split '/')[4]
                        $NSGName = ($Subnet.NSG -split '/')[-1]
                        $NSG = Get-AzNetworkSecurityGroup -ResourceGroupName $NSGRG -Name $NSGName -ErrorAction Stop

                        $Discovery.Inventory.NSGs += [PSCustomObject]@{
                            Name          = $NSG.Name
                            Id            = $NSG.Id
                            ResourceGroup = $NSGRG
                            Rules         = @($NSG.SecurityRules | ForEach-Object {
                                [PSCustomObject]@{
                                    Name                   = $_.Name
                                    Priority               = $_.Priority
                                    Direction              = $_.Direction
                                    Access                 = $_.Access
                                    Protocol               = $_.Protocol
                                    SourcePortRange        = $_.SourcePortRange
                                    DestinationPortRange   = $_.DestinationPortRange
                                    DestinationPortRanges  = $_.DestinationPortRanges
                                    SourceAddressPrefix    = $_.SourceAddressPrefix
                                    DestinationAddressPrefix = $_.DestinationAddressPrefix
                                }
                            })
                        }

                        $InboundRules = @($NSG.SecurityRules | Where-Object { $_.Direction -eq 'Inbound' })

                        # Helper (inline): is an internet-source Allow on $Port overridden by a higher-priority
                        # (lower number) Deny covering the same port from the same/broader source? (C-1)
                        $EvalPortExposure = {
                            param([int]$Port)
                            $AllowExposed = @($InboundRules | Where-Object {
                                $_.Access -eq 'Allow' -and (Test-NsgRulePort -Rule $_ -Port $Port) -and (Test-NsgInternetSource -Rule $_)
                            })
                            $Effective = @($AllowExposed | Where-Object {
                                $AllowRule = $_
                                $OverridingDeny = @($InboundRules | Where-Object {
                                    $_.Access -eq 'Deny' -and $_.Priority -lt $AllowRule.Priority -and
                                    (Test-NsgRulePort -Rule $_ -Port $Port) -and (Test-NsgInternetSource -Rule $_)
                                })
                                $OverridingDeny.Count -eq 0
                            })
                            return $Effective
                        }

                        # CHECK: Port 3389 exposure
                        $ExposedToInternet = @(& $EvalPortExposure 3389)
                        [void]$AllChecks.Add((New-CheckResult -Id "NET-RDP-$NSGName" `
                            -Category 'Networking' -Name 'RDP Port 3389 Not Internet-Exposed' `
                            -Description 'Port 3389 should not be open to the internet on AVD subnets' `
                            -Status $(if ($ExposedToInternet.Count -gt 0) { 'Fail' } else { 'Pass' }) `
                            -Severity 'Critical' `
                            -Details "Internet-facing RDP rules (not overridden by a higher-priority Deny): $($ExposedToInternet.Count)$(if ($ExposedToInternet.Count -gt 0) { " ($( ($ExposedToInternet | ForEach-Object { $_.Name }) -join ', '))" })" `
                            -Recommendation 'Block inbound RDP from internet. Use Azure Bastion or JIT access for administration.' `
                            -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/security-guide' `
                            -Evidence @{ NSG = $NSGName; ExposedRules = @($ExposedToInternet | ForEach-Object { $_.Name }) }))

                        # CHECK: AVD required outbound connectivity
                        $OutboundRules = @($NSG.SecurityRules | Where-Object { $_.Direction -eq 'Outbound' })
                        $HasDenyAllOut = @($OutboundRules | Where-Object {
                            $_.Access -eq 'Deny' -and $_.DestinationAddressPrefix -eq '*' -and (Test-NsgRulePort -Rule $_ -Port 443)
                        }).Count -gt 0
                        if ($HasDenyAllOut) {
                            $HasWVDAllow = @($OutboundRules | Where-Object {
                                $_.Access -eq 'Allow' -and ($_.DestinationAddressPrefix -match 'WindowsVirtualDesktop|AzureCloud')
                            }).Count -gt 0
                            $HasAADAllow = @($OutboundRules | Where-Object {
                                $_.Access -eq 'Allow' -and ($_.DestinationAddressPrefix -match 'AzureActiveDirectory')
                            }).Count -gt 0
                            [void]$AllChecks.Add((New-CheckResult -Id "NET-AVDOUT-$NSGName" `
                                -Category 'Networking' -Name 'AVD Required Outbound Rules' `
                                -Description 'When default outbound is denied, NSG must allow WindowsVirtualDesktop and AzureAD service tags on 443' `
                                -Status $(if ($HasWVDAllow -and $HasAADAllow) { 'Pass' }
                                          elseif ($HasWVDAllow -or $HasAADAllow) { 'Warning' } else { 'Fail' }) `
                                -Severity 'Critical' `
                                -Details "DenyAllOutbound: Yes, WVDServiceTag: $HasWVDAllow, AzureAD: $HasAADAllow" `
                                -Recommendation 'Add allow rules for WindowsVirtualDesktop and AzureActiveDirectory service tags on port 443.' `
                                -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/required-fqdn-endpoint' `
                                -Evidence @{ NSG = $NSGName; DenyAllOutbound = $true; WVDAllow = $HasWVDAllow; AADAllow = $HasAADAllow }))
                        }

                        # CHECK: SSH port 22 exposure (common misconfiguration)
                        $SshExposed = @(& $EvalPortExposure 22)
                        if ($SshExposed.Count -gt 0) {
                            [void]$AllChecks.Add((New-CheckResult -Id "NET-SSH-$NSGName" `
                                -Category 'Networking' -Name 'SSH Port 22 Not Internet-Exposed' `
                                -Description 'Port 22 should not be open to the internet' `
                                -Status 'Fail' -Severity 'High' `
                                -Details "Internet-facing SSH rules (not overridden by a higher-priority Deny): $($SshExposed.Count)" `
                                -Recommendation 'Block inbound SSH from internet. Use Azure Bastion for management.' `
                                -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/security-guide'))
                        }
                    } catch {
                        [void]$AllChecks.Add((New-CheckResult -Id "NET-RDP-$($Subnet.NSG -split '/' | Select-Object -Last 1)" `
                            -Category 'Networking' -Name 'RDP Port 3389 Not Internet-Exposed' `
                            -Description 'Port 3389 should not be open to the internet on AVD subnets' `
                            -Status 'Error' -Severity 'Critical' `
                            -Details "Could not read NSG rules: $($_.Exception.Message)" `
                            -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/security-guide'))
                    }
                }
            }
        }
    } catch {
        Write-Status "  Error discovering network: $($_.Exception.Message)" -Level 'ERROR'
        $Discovery.Errors += "Network discovery failed: $($_.Exception.Message)"
    }

        # Populate top-level Subnets and UDRs arrays from VNet data
        foreach ($VNetEntry in $Discovery.Inventory.VNets) {
            foreach ($SubnetEntry in $VNetEntry.Subnets) {
                $Discovery.Inventory.Subnets += [PSCustomObject]@{
                    VNetName      = $VNetEntry.Name
                    Name          = $SubnetEntry.Name
                    AddressPrefix = $SubnetEntry.AddressPrefix
                    NSG           = $SubnetEntry.NSG
                    RouteTable    = $SubnetEntry.RouteTable
                }
                if ($SubnetEntry.RouteTable) {
                    $Discovery.Inventory.UDRs += [PSCustomObject]@{
                        VNetName  = $VNetEntry.Name
                        Subnet    = $SubnetEntry.Name
                        Id        = $SubnetEntry.RouteTable
                        Name      = ($SubnetEntry.RouteTable -split '/')[-1]
                    }
                }
            }
        }

    # ─── HUB NETWORK RESOURCES (Firewall, VPN/ER Gateway) ────────────────
    Write-Status "Hub Network Resources" -Level 'SECTION'
    try {
        # Discover Azure Firewalls in the subscription
        $AzFirewalls = @(Get-AzFirewall -ErrorAction SilentlyContinue)
        foreach ($Fw in $AzFirewalls) {
            $Discovery.Inventory.Firewalls += [PSCustomObject]@{
                Name          = $Fw.Name
                ResourceGroup = $Fw.ResourceGroup
                Location      = $Fw.Location
                Sku           = $Fw.Sku.Tier
                ThreatIntel   = $Fw.ThreatIntelMode
                VNetId        = if ($Fw.IpConfigurations -and $Fw.IpConfigurations[0].Subnet) {
                                    ($Fw.IpConfigurations[0].Subnet.Id -split '/subnets/')[0]
                                } else { $null }
            }
        }
        # Check if any AVD VNet peers to a VNet with a firewall
        $FwVNetIds = @($Discovery.Inventory.Firewalls | ForEach-Object { $_.VNetId } | Where-Object { $_ })
        $AvdVNetIds = @($Discovery.Inventory.VNets | ForEach-Object { $_.Id })
        $PeeredFw = $false
        foreach ($VNet in $Discovery.Inventory.VNets) {
            if ($VNet.Peerings) {
                foreach ($Peer in $VNet.Peerings) {
                    if ($Peer.RemoteVNet -in $FwVNetIds) { $PeeredFw = $true }
                }
            }
        }
        $DirectFw = @($AzFirewalls | Where-Object { ($_.IpConfigurations[0].Subnet.Id -split '/subnets/')[0] -in $AvdVNetIds }).Count -gt 0
        $HasFirewall = $PeeredFw -or $DirectFw -or ($AzFirewalls.Count -gt 0)
        Write-Status "  Azure Firewalls: $($AzFirewalls.Count), Peered to AVD: $PeeredFw" -Level $(if ($HasFirewall) { 'SUCCESS' } else { 'WARN' })
        [void]$AllChecks.Add((New-CheckResult -Id "NET-HUBFW-$SubShort" `
            -Category 'Networking' -Name 'Hub Firewall Present' `
            -Description 'Azure Firewall or NVA should exist in hub for centralized egress filtering' `
            -Status $(if ($HasFirewall) { 'Pass' } else { 'Warning' }) `
            -Severity 'Medium' `
            -Details "AzureFirewalls: $($AzFirewalls.Count)$(if ($AzFirewalls.Count -gt 0) { " ($( ($AzFirewalls | ForEach-Object { "$($_.Name) [$($_.Sku.Tier)]" }) -join ', '))" }), PeeredToAVD: $PeeredFw" `
            -Recommendation 'Deploy Azure Firewall in hub VNet for centralized egress filtering and threat intelligence.' `
            -Reference 'https://learn.microsoft.com/en-us/azure/architecture/networking/architecture/hub-spoke' `
            -Evidence @{ Count = $AzFirewalls.Count; PeeredToAVD = $PeeredFw }))

        # Discover VPN/ExpressRoute Gateways (Get-AzResource avoids mandatory -ResourceGroupName)
        $GwResources = @(Get-AzResource -ResourceType 'Microsoft.Network/virtualNetworkGateways' -ErrorAction SilentlyContinue)
        $VPNGateways = @()
        foreach ($GwRes in $GwResources) {
            try {
                $Gw = Get-AzVirtualNetworkGateway -ResourceGroupName $GwRes.ResourceGroupName -Name $GwRes.Name -ErrorAction Stop
                $VPNGateways += $Gw
                $Discovery.Inventory.VPNGateways += [PSCustomObject]@{
                    Name          = $Gw.Name
                    ResourceGroup = $Gw.ResourceGroupName
                    GatewayType   = $Gw.GatewayType   # Vpn or ExpressRoute
                    VpnType       = $Gw.VpnType
                    Sku           = $Gw.Sku.Name
                    Active        = $Gw.ActiveActive
                    Location      = $Gw.Location
                }
            } catch {
                Write-Status "    Could not get gateway $($GwRes.Name): $($_.Exception.Message)" -Level 'WARN'
            }
        }
        $HasGateway = $VPNGateways.Count -gt 0
        $GwTypes = @($VPNGateways | ForEach-Object { $_.GatewayType } | Sort-Object -Unique) -join ', '
        Write-Status "  VPN/ER Gateways: $($VPNGateways.Count) ($GwTypes)" -Level $(if ($HasGateway) { 'SUCCESS' } else { 'WARN' })
        [void]$AllChecks.Add((New-CheckResult -Id "NET-HUBGW-$SubShort" `
            -Category 'Networking' -Name 'VPN/ExpressRoute Gateway' `
            -Description 'Hub network should have VPN or ExpressRoute gateway for hybrid connectivity' `
            -Status $(if ($HasGateway) { 'Pass' } else { 'Warning' }) `
            -Severity 'Low' `
            -Details "Gateways: $($VPNGateways.Count)$(if ($VPNGateways.Count -gt 0) { " ($( ($VPNGateways | ForEach-Object { "$($_.Name) [$($_.GatewayType)/$($_.Sku.Name)]" }) -join ', '))" })" `
            -Recommendation 'Deploy VPN or ExpressRoute gateway for hybrid connectivity to on-premises AD DS and file shares.' `
            -Reference 'https://learn.microsoft.com/en-us/azure/cloud-adoption-framework/ready/azure-best-practices/define-an-azure-network-topology' `
            -Evidence @{ Count = $VPNGateways.Count; Types = $GwTypes }))
    } catch {
        Write-Status "  Hub network error: $($_.Exception.Message)" -Level 'WARN'
    }

    # ─── DIAGNOSTICS ──────────────────────────────────────────────────────
    Write-Status "Diagnostics" -Level 'SECTION'
    try {
        $RecommendedCategories = @('Checkpoint','Error','Management','Connection','HostRegistration','AgentHealthStatus')
        foreach ($HP in $HostPools) {
            try {
                $DiagSettings = @(Get-AzDiagnosticSetting -ResourceId $HP.Id -ErrorAction Stop -WarningAction SilentlyContinue)
                $HasDiag = $DiagSettings.Count -gt 0
                $HasLA   = @($DiagSettings | Where-Object { $_.WorkspaceId }).Count -gt 0
                # Harvest Log Analytics workspace resource IDs for the SIEM/Sentinel check (MON-012).
                foreach ($DS in @($DiagSettings | Where-Object { $_.WorkspaceId })) { $LAWorkspaceIds["$($DS.WorkspaceId)"] = $true }

                [void]$AllChecks.Add((New-CheckResult -Id "MON-DIAG-$($HP.Name)" `
                    -Category 'Monitoring' -Name 'Diagnostic Settings Enabled' `
                    -Description 'Host pools should have diagnostics enabled for monitoring and troubleshooting' `
                    -Status $(if ($HasLA) { 'Pass' } elseif ($HasDiag) { 'Warning' } else { 'Fail' }) `
                    -Severity 'High' `
                    -Details "DiagSettings: $($DiagSettings.Count), LogAnalytics: $HasLA" `
                    -Recommendation 'Enable diagnostic settings with a Log Analytics workspace for AVD Insights.' `
                    -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/diagnostics-log-analytics'))

                # ─── CHECK: Diagnostic Log Categories (MON-016) ───
                if ($HasDiag) {
                    # A diagnostic setting using a category GROUP (allLogs / audit) covers all recommended categories (C-6).
                    $UsesCategoryGroup = @($DiagSettings | ForEach-Object { $_.Log } | Where-Object { $_.Enabled -and $_.CategoryGroup -in @('allLogs','audit') }).Count -gt 0
                    $EnabledCategories = @($DiagSettings | ForEach-Object { $_.Log } | Where-Object { $_.Enabled } | ForEach-Object { $_.Category }) | Sort-Object -Unique
                    $MissingCategories = if ($UsesCategoryGroup) { @() } else { @($RecommendedCategories | Where-Object { $_ -notin $EnabledCategories }) }
                    [void]$AllChecks.Add((New-CheckResult -Id "MON-DIAGCAT-$($HP.Name)" `
                        -Category 'Monitoring' -Name 'Diagnostic Log Categories' `
                        -Description 'All recommended AVD diagnostic log categories should be enabled' `
                        -Status $(if ($MissingCategories.Count -eq 0) { 'Pass' } else { 'Warning' }) `
                        -Severity 'Medium' `
                        -Details "$(if ($UsesCategoryGroup) { 'CategoryGroup (allLogs/audit) enabled - covers all recommended categories' } else { "Enabled: $($EnabledCategories -join ', '). Missing: $(if ($MissingCategories.Count -gt 0) { $MissingCategories -join ', ' } else { 'None' })" })" `
                        -Recommendation "$(if ($MissingCategories.Count -gt 0) { "Enable missing categories: $($MissingCategories -join ', ')" } else { 'Diagnostic categories fully covered.' })" `
                        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/diagnostics-log-analytics' `
                        -Evidence @{ Enabled = $EnabledCategories; Missing = $MissingCategories; CategoryGroup = $UsesCategoryGroup }))
                }
            } catch {
                [void]$AllChecks.Add((New-CheckResult -Id "MON-DIAG-$($HP.Name)" `
                    -Category 'Monitoring' -Name 'Diagnostic Settings Enabled' `
                    -Description 'Host pools should have diagnostics enabled for monitoring and troubleshooting' `
                    -Status 'Error' -Severity 'High' `
                    -Details "Could not read diagnostic settings: $($_.Exception.Message)" `
                    -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/diagnostics-log-analytics'))
            }
        }
        $DiagHPCount = @($AllChecks | Where-Object { $_.Id -like 'MON-DIAG-*' }).Count
        $DiagPassCount = @($AllChecks | Where-Object { $_.Id -like 'MON-DIAG-*' -and $_.Status -eq 'Pass' }).Count
        Write-Status "  Host pools with diagnostics: $DiagPassCount/$DiagHPCount" -Level $(if ($DiagPassCount -eq $DiagHPCount -and $DiagHPCount -gt 0) { 'SUCCESS' } elseif ($DiagHPCount -eq 0) { 'WARN' } else { 'WARN' })
    } catch {
        Write-Status "  Error: $($_.Exception.Message)" -Level 'ERROR'
        $Discovery.Errors += "Diagnostics check failed: $($_.Exception.Message)"
    }

    # ─── RBAC ─────────────────────────────────────────────────────────────
    Write-Status "RBAC" -Level 'SECTION'
    try {
        $BroadRoles = @('Owner','Contributor','User Access Administrator')
        foreach ($HP in $HostPools) {
            try {
                # Assignments on the host pool AND its application groups - user assignments normally
                # live on app groups, not the host pool (A-3).
                $Assignments = @(Get-AzRoleAssignment -Scope $HP.Id -ErrorAction Stop)
                $HPAppGroups = @($AppGroups | Where-Object { "$($_.HostPoolArmPath)".ToLower() -eq "$($HP.Id)".ToLower() })
                foreach ($HPAG in $HPAppGroups) {
                    $Assignments += @(Get-AzRoleAssignment -Scope $HPAG.Id -ErrorAction Stop)
                }
                # Azure built-in AVD roles are named "Desktop Virtualization ..." (A-3 - wildcard was reversed).
                $AvdRoleCount = @($Assignments | Where-Object { $_.RoleDefinitionName -like '*Desktop Virtualization*' }).Count
                $BroadAtScope = @($Assignments | Where-Object {
                    $_.RoleDefinitionName -in $BroadRoles -and ($_.Scope -like "*$($HP.Name)*" -or $_.Scope -like '*applicationGroups*')
                })
                if ($BroadAtScope.Count -gt 0) {
                    $RbacStatus = 'Warning'
                    $RbacDetail = "Broad roles assigned directly at AVD scope: $(@($BroadAtScope | ForEach-Object { "$($_.RoleDefinitionName) ($($_.DisplayName))" } | Select-Object -First 5) -join ', '). AVDRoles: $AvdRoleCount"
                } elseif ($AvdRoleCount -gt 0) {
                    $RbacStatus = 'Pass'
                    $RbacDetail = "TotalAssignments: $($Assignments.Count), Desktop Virtualization roles: $AvdRoleCount, no broad roles at AVD scope"
                } else {
                    $RbacStatus = 'Warning'
                    $RbacDetail = "TotalAssignments: $($Assignments.Count), no Desktop Virtualization built-in roles found on host pool or app groups"
                }
                [void]$AllChecks.Add((New-CheckResult -Id "SEC-RBAC-$($HP.Name)" `
                    -Category 'Security & IAM' -Name 'AVD RBAC Roles Used' `
                    -Description 'Use built-in AVD roles for least-privilege access; avoid Owner/Contributor at AVD scopes' `
                    -Status $RbacStatus -Severity 'Medium' `
                    -Details $RbacDetail `
                    -Recommendation 'Use built-in AVD roles (Desktop Virtualization Contributor, Desktop Virtualization User, etc.); remove Owner/Contributor/User Access Administrator assignments at host pool and app group scopes.' `
                    -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/rbac' `
                    -Evidence @{ HostPool = $HP.Name; AvdRoles = $AvdRoleCount; BroadRoles = @($BroadAtScope | ForEach-Object { $_.RoleDefinitionName }) }))

                # ─── CHECK: App Group Assignment via Entra Groups (APP-004) — reuses $Assignments ───
                $UserAssignments = @($Assignments | Where-Object { $_.RoleDefinitionName -eq 'Desktop Virtualization User' })
                $GroupAssigned   = @($UserAssignments | Where-Object { $_.ObjectType -eq 'Group' })
                $DirectUserAssigned = @($UserAssignments | Where-Object { $_.ObjectType -eq 'User' })
                if ($UserAssignments.Count -gt 0) {
                    if ($DirectUserAssigned.Count -gt 0) {
                        $GrpStatus = 'Warning'
                        $GrpDetail = "Desktop Virtualization User: $($GroupAssigned.Count) group, $($DirectUserAssigned.Count) direct-user assignment(s) - assign via Entra groups"
                    } else {
                        $GrpStatus = 'Pass'
                        $GrpDetail = "Desktop Virtualization User assigned via $($GroupAssigned.Count) Entra group(s), no direct-user assignments"
                    }
                    [void]$AllChecks.Add((New-CheckResult -Id "APP-GRPASSIGN-$($HP.Name)" `
                        -Category 'Application Delivery' -Name 'App Group Assignment via Entra Groups' `
                        -Description 'Application group access should be granted via Entra ID security groups rather than direct user assignment' `
                        -Status $GrpStatus -Severity 'Medium' `
                        -Details "HostPool $($HP.Name): $GrpDetail" `
                        -Recommendation 'Assign the Desktop Virtualization User role to Entra ID security groups (self-service via access packages, simpler auditing) rather than to individual users.' `
                        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/manage-app-groups' `
                        -Evidence @{ HostPool = $HP.Name; GroupAssignments = $GroupAssigned.Count; UserAssignments = $DirectUserAssigned.Count }))
                }
            } catch {
                [void]$AllChecks.Add((New-CheckResult -Id "SEC-RBAC-$($HP.Name)" `
                    -Category 'Security & IAM' -Name 'AVD RBAC Roles Used' `
                    -Description 'Use built-in AVD roles for least-privilege access' `
                    -Status 'Error' -Severity 'Medium' `
                    -Details "Could not enumerate role assignments (authorization?): $($_.Exception.Message)" `
                    -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/rbac'))
            }
        }
        $RbacChecked = @($AllChecks | Where-Object { $_.Id -like 'SEC-RBAC-*' }).Count
        $RbacWithAvdRoles = @($AllChecks | Where-Object { $_.Id -like 'SEC-RBAC-*' -and $_.Status -eq 'Pass' }).Count
        Write-Status "  RBAC checked: $RbacChecked host pool(s), AVD roles present: $RbacWithAvdRoles" -Level $(if ($RbacWithAvdRoles -gt 0) { 'SUCCESS' } else { 'WARN' })
    } catch {
        Write-Status "  Error: $($_.Exception.Message)" -Level 'ERROR'
    }

    # ─── DIAGNOSTIC COVERAGE (workspaces + app groups) ────────────────────
    Write-Status "Diagnostic Coverage" -Level 'SECTION'
    try {
        # Check workspaces
        foreach ($WS in $Workspaces) {
            try {
                $WSDiag = @(Get-AzDiagnosticSetting -ResourceId $WS.Id -ErrorAction Stop -WarningAction SilentlyContinue)
                [void]$AllChecks.Add((New-CheckResult -Id "MON-WSDIAG-$($WS.Name)" `
                    -Category 'Monitoring' -Name 'Workspace Diagnostics' `
                    -Description 'AVD workspaces should have diagnostic settings enabled' `
                    -Status $(if ($WSDiag.Count -gt 0) { 'Pass' } else { 'Warning' }) `
                    -Severity 'Medium' `
                    -Details "DiagSettings: $($WSDiag.Count)" `
                    -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/diagnostics-log-analytics'))
            } catch {
                [void]$AllChecks.Add((New-CheckResult -Id "MON-WSDIAG-$($WS.Name)" `
                    -Category 'Monitoring' -Name 'Workspace Diagnostics' `
                    -Description 'AVD workspaces should have diagnostic settings enabled' `
                    -Status 'Error' -Severity 'Medium' `
                    -Details "Could not read diagnostic settings: $($_.Exception.Message)" `
                    -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/diagnostics-log-analytics'))
            }
        }
        # Check app groups
        foreach ($AG in $AppGroups) {
            try {
                $AGDiag = @(Get-AzDiagnosticSetting -ResourceId $AG.Id -ErrorAction Stop -WarningAction SilentlyContinue)
                [void]$AllChecks.Add((New-CheckResult -Id "MON-AGDIAG-$($AG.Name)" `
                    -Category 'Monitoring' -Name 'App Group Diagnostics' `
                    -Description 'App groups should have diagnostic settings enabled' `
                    -Status $(if ($AGDiag.Count -gt 0) { 'Pass' } else { 'Warning' }) `
                    -Severity 'Medium' `
                    -Details "DiagSettings: $($AGDiag.Count)" `
                    -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/diagnostics-log-analytics'))
            } catch {
                [void]$AllChecks.Add((New-CheckResult -Id "MON-AGDIAG-$($AG.Name)" `
                    -Category 'Monitoring' -Name 'App Group Diagnostics' `
                    -Description 'App groups should have diagnostic settings enabled' `
                    -Status 'Error' -Severity 'Medium' `
                    -Details "Could not read diagnostic settings: $($_.Exception.Message)" `
                    -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/diagnostics-log-analytics'))
            }
        }
        $WsDiagCount   = @($AllChecks | Where-Object { $_.Id -like 'MON-WSDIAG-*' }).Count
        $AgDiagCount   = @($AllChecks | Where-Object { $_.Id -like 'MON-AGDIAG-*' }).Count
        $WsDiagPass    = @($AllChecks | Where-Object { $_.Id -like 'MON-WSDIAG-*' -and $_.Status -eq 'Pass' }).Count
        $AgDiagPass    = @($AllChecks | Where-Object { $_.Id -like 'MON-AGDIAG-*' -and $_.Status -eq 'Pass' }).Count
        Write-Status "  Workspaces: $WsDiagPass/$WsDiagCount with diagnostics, App groups: $AgDiagPass/$AgDiagCount" -Level $(if ($WsDiagPass + $AgDiagPass -eq $WsDiagCount + $AgDiagCount -and $WsDiagCount + $AgDiagCount -gt 0) { 'SUCCESS' } else { 'WARN' })
    } catch {
        Write-Status "  Error: $($_.Exception.Message)" -Level 'WARN'
    }

    # ─── DEFENDER FOR CLOUD ───────────────────────────────────────────────
    try {
        $DefenderVMs = Get-AzSecurityPricing -Name 'VirtualMachines' -ErrorAction Stop
        if ($DefenderVMs) {
            [void]$AllChecks.Add((New-CheckResult -Id "MON-DEFENDER-$SubShort" `
                -Category 'Monitoring' -Name 'Defender for Cloud Enabled' `
                -Description 'Microsoft Defender for Cloud should be enabled for VMs' `
                -Status $(if ($DefenderVMs.PricingTier -eq 'Standard') { 'Pass' } else { 'Warning' }) `
                -Severity 'High' `
                -Details "PricingTier: $($DefenderVMs.PricingTier)" `
                -Reference 'https://learn.microsoft.com/en-us/azure/defender-for-cloud/enable-enhanced-security'))
        }
    } catch {
        [void]$AllChecks.Add((New-CheckResult -Id "MON-DEFENDER-$SubShort" `
            -Category 'Monitoring' -Name 'Defender for Cloud Enabled' `
            -Description 'Microsoft Defender for Cloud should be enabled for VMs' `
            -Status 'Error' -Severity 'High' `
            -Details "Could not read Defender for Cloud pricing (Az.Security missing or access denied): $($_.Exception.Message)" `
            -Reference 'https://learn.microsoft.com/en-us/azure/defender-for-cloud/enable-enhanced-security'))
    }

    # ─── STORAGE ──────────────────────────────────────────────────────────
    Write-Status "Storage (FSLogix)" -Level 'SECTION'
    try {
        $StorageAccounts = @(Get-AzStorageAccount -ErrorAction Stop)
        # Look for storage accounts that might be used for FSLogix (heuristic: has 'profiles' or 'fslogix' in name or file shares)
        foreach ($SA in $StorageAccounts) {
            # ─── FSLogix candidate classification (A-7): evidence-based, not name-substring only ───
            $FSLogixReasons = @()
            # Name signal: 'fslogix'/'profile' substrings, or 'avd' as a word-ish token (not e.g. 'mavdata').
            if ($SA.StorageAccountName -match 'fslogix|profile') { $FSLogixReasons += "name matches 'fslogix|profile'" }
            elseif ($SA.StorageAccountName -match '(^|[-_])avd')  { $FSLogixReasons += "name contains 'avd' token" }
            # Identity-based auth signal: AD DS or Entra Kerberos configured for Azure Files.
            $DirSvc = $null
            if ($SA.AzureFilesIdentityBasedAuth -and $SA.AzureFilesIdentityBasedAuth.DirectoryServiceOptions) {
                $DirSvc = "$($SA.AzureFilesIdentityBasedAuth.DirectoryServiceOptions)"
                if ($DirSvc -in @('AD','AADKERB','AADDS')) { $FSLogixReasons += "identity-based file auth: $DirSvc" }
            }
            # File share name signal: shares matching profile/fslogix/odfc.
            try {
                $Shares = @(Get-AzRmStorageShare -ResourceGroupName $SA.ResourceGroupName -StorageAccountName $SA.StorageAccountName -ErrorAction Stop)
                $ProfileShares = @($Shares | Where-Object { $_.Name -match 'profile|fslogix|odfc' })
                if ($ProfileShares.Count -gt 0) { $FSLogixReasons += "share name(s): $(@($ProfileShares | ForEach-Object { $_.Name }) -join ',')" }
            } catch {
                Write-Status "    Could not enumerate file shares on $($SA.StorageAccountName): $($_.Exception.Message)" -Level 'WARN'
            }
            $IsFSLogix = $FSLogixReasons.Count -gt 0
            # Get-AzStorageAccount doesn't reliably populate PrivateEndpointConnections — use ARM API
            $SAResource = Get-AzResource -ResourceId $SA.Id -ErrorAction SilentlyContinue
            $HasPrivateEndpoint = $SAResource -and $SAResource.Properties.privateEndpointConnections -and
                @($SAResource.Properties.privateEndpointConnections).Count -gt 0

            $SAObj = [PSCustomObject]@{
                Name              = $SA.StorageAccountName
                Id                = $SA.Id
                ResourceGroup     = $SA.ResourceGroupName
                Kind              = $SA.Kind
                SkuName           = $SA.Sku.Name
                Location          = $SA.PrimaryLocation
                AccessTier        = $SA.AccessTier
                MinTlsVersion     = $SA.MinimumTlsVersion
                HttpsOnly         = $SA.EnableHttpsTrafficOnly
                PrivateEndpoints  = $HasPrivateEndpoint
                LikelyFSLogix     = $IsFSLogix
                FSLogixEvidence   = $FSLogixReasons
                Replication       = $SA.Sku.Name  # LRS, ZRS, GRS, etc.
                LargeFileShares   = $SA.LargeFileSharesState
                Tags              = $SA.Tags
            }
            $Discovery.Inventory.StorageAccounts += $SAObj

            if ($IsFSLogix) {
                # Classification evidence is included so users can spot misclassification (A-7).
                $FSLogixWhy = "Classified as FSLogix candidate because: $($FSLogixReasons -join '; ')"

                # ─── CHECK: Entra Kerberos for Profile Storage (IAM-012) — evidence from DirectoryServiceOptions ───
                $KerbAuth = if ($DirSvc) { "$DirSvc" } else { 'None' }
                if ($KerbAuth -eq 'AADKERB') {
                    $KerbStat = 'Pass'; $KerbDet = "DirectoryServiceOptions: AADKERB (Entra Kerberos - cloud-native identity-based SMB auth)"
                } elseif ($KerbAuth -in @('AD','AADDS')) {
                    $KerbStat = 'Pass'; $KerbDet = "DirectoryServiceOptions: $KerbAuth (identity-based SMB auth configured; Entra Kerberos (AADKERB) is preferred for cloud-native estates)"
                } else {
                    $KerbStat = 'Warning'; $KerbDet = "DirectoryServiceOptions: None (identity-based SMB auth not configured)"
                }
                [void]$AllChecks.Add((New-CheckResult -Id "IAM-KERB-$($SA.StorageAccountName)" `
                    -Category 'Identity & Access' -Name 'Entra Kerberos for Profile Storage' `
                    -Description 'FSLogix profile storage should use identity-based SMB authentication (Entra Kerberos preferred for cloud-native, no on-prem AD DS required)' `
                    -Status $KerbStat -Severity 'Medium' `
                    -Details "$KerbDet. $FSLogixWhy" `
                    -Recommendation 'Enable Entra Kerberos (AADKERB) on the storage account for cloud-native identity-based access to FSLogix profile shares.' `
                    -Reference 'https://learn.microsoft.com/en-us/azure/storage/files/storage-files-identity-auth-hybrid-identities-enable' `
                    -Evidence @{ StorageAccount = $SA.StorageAccountName; DirectoryServiceOptions = $KerbAuth }))

                # CHECK: Private endpoint on FSLogix storage
                [void]$AllChecks.Add((New-CheckResult -Id "PROF-PE-$($SA.StorageAccountName)" `
                    -Category 'FSLogix & Profiles' -Name 'Private Endpoint on Profile Storage' `
                    -Description 'FSLogix storage should use private endpoints for security' `
                    -Status $(if ($HasPrivateEndpoint) { 'Pass' } else { 'Warning' }) `
                    -Severity 'High' `
                    -Details "PrivateEndpoint: $HasPrivateEndpoint. $FSLogixWhy" `
                    -Recommendation 'Configure private endpoints for FSLogix profile storage to keep traffic on the Microsoft network.' `
                    -Reference 'https://learn.microsoft.com/en-us/azure/storage/files/storage-files-networking-overview' `
                    -Evidence @{ StorageAccount = $SA.StorageAccountName; ClassificationEvidence = $FSLogixReasons }))

                # CHECK: HTTPS only
                [void]$AllChecks.Add((New-CheckResult -Id "PROF-HTTPS-$($SA.StorageAccountName)" `
                    -Category 'FSLogix & Profiles' -Name 'HTTPS Only Enabled' `
                    -Description 'Storage account should enforce HTTPS-only traffic' `
                    -Status $(if ($SA.EnableHttpsTrafficOnly) { 'Pass' } else { 'Fail' }) `
                    -Severity 'High' `
                    -Details "HttpsOnly: $($SA.EnableHttpsTrafficOnly). $FSLogixWhy"))

                # CHECK: TLS version - explicit membership test, not lexical compare (C-9)
                [void]$AllChecks.Add((New-CheckResult -Id "PROF-TLS-$($SA.StorageAccountName)" `
                    -Category 'FSLogix & Profiles' -Name 'Minimum TLS 1.2' `
                    -Description 'Storage account should enforce TLS 1.2 minimum' `
                    -Status $(if ("$($SA.MinimumTlsVersion)" -in @('TLS1_2','TLS1_3')) { 'Pass' } else { 'Fail' }) `
                    -Severity 'High' `
                    -Details "MinTLS: $($SA.MinimumTlsVersion)"))

                # CHECK: Storage replication for DR
                $RepType = $SA.Sku.Name
                [void]$AllChecks.Add((New-CheckResult -Id "BCDR-STOR-$($SA.StorageAccountName)" `
                    -Category 'BCDR' -Name 'Profile Storage Replication' `
                    -Description 'FSLogix storage should use ZRS or GRS for resilience' `
                    -Status $(if ($RepType -match 'ZRS|GRS|GZRS') { 'Pass' }
                              elseif ($RepType -match 'LRS') { 'Warning' } else { 'Warning' }) `
                    -Severity 'Medium' `
                    -Details "Replication: $RepType" `
                    -Recommendation 'Use ZRS for zone-level resilience or GRS for region-level DR.' `
                    -Reference 'https://learn.microsoft.com/en-us/azure/storage/common/storage-redundancy'))

                # CHECK: Premium tier for profiles
                [void]$AllChecks.Add((New-CheckResult -Id "PROF-TIER-$($SA.StorageAccountName)" `
                    -Category 'FSLogix & Profiles' -Name 'Premium Storage for Profiles' `
                    -Description 'Premium storage provides better IOPS for FSLogix profile containers' `
                    -Status $(if ($SA.Sku.Name -match 'Premium') { 'Pass' } else { 'Warning' }) `
                    -Severity 'Medium' `
                    -Details "SKU: $($SA.Sku.Name), Kind: $($SA.Kind)" `
                    -Recommendation 'Consider Premium FileStorage for better FSLogix performance (lower latency, higher IOPS).' `
                    -Reference 'https://learn.microsoft.com/en-us/azure/storage/files/storage-files-scale-targets'))

                # CHECK: Storage firewall
                $FwDefault = $SA.NetworkRuleSet.DefaultAction
                [void]$AllChecks.Add((New-CheckResult -Id "PROF-FW-$($SA.StorageAccountName)" `
                    -Category 'FSLogix & Profiles' -Name 'Storage Firewall Configured' `
                    -Description 'Storage default network action should be Deny' `
                    -Status $(if ($FwDefault -eq 'Deny') { 'Pass' } else { 'Warning' }) `
                    -Severity 'High' `
                    -Details "DefaultAction: $FwDefault" `
                    -Reference 'https://learn.microsoft.com/en-us/azure/storage/files/storage-files-networking-overview'))

                # CHECK: Soft delete
                try {
                    $FSP = Get-AzStorageFileServiceProperty -StorageAccountName $SA.StorageAccountName -ResourceGroupName $SA.ResourceGroupName -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                    $SDEnabled = if ($FSP) { $FSP.ShareDeleteRetentionPolicy.Enabled } else { $false }
                    [void]$AllChecks.Add((New-CheckResult -Id "PROF-SD-$($SA.StorageAccountName)" `
                        -Category 'FSLogix & Profiles' -Name 'Soft Delete Enabled' `
                        -Description 'File share soft delete protects against accidental deletion' `
                        -Status $(if ($SDEnabled) { 'Pass' } else { 'Warning' }) `
                        -Severity 'Medium' `
                        -Details "SoftDelete: $SDEnabled" `
                        -Reference 'https://learn.microsoft.com/en-us/azure/storage/files/storage-files-enable-soft-delete'))

                    # CHECK: SMB security settings (A-6). Null/unset settings are NOT compliant - Azure
                    # Files defaults still permit SMB 2.1 / RC4-HMAC / NTLMv2. Pass only when the
                    # property exists AND excludes the weak value. When ProtocolSetting.Smb is entirely
                    # null, still emit the same Warnings (do not skip).
                    $Smb = if ($FSP -and $FSP.ProtocolSetting) { $FSP.ProtocolSetting.Smb } else { $null }
                    $SmbUnsetDetail = 'SMB security not explicitly hardened; Azure Files defaults permit SMB 2.1 / RC4-HMAC / NTLMv2 - configure explicit versions/encryption/auth'

                    # SMB versions (PROF-020)
                    $Versions = if ($Smb) { $Smb.Versions } else { $null }
                    $SmbVerStatus = if ([string]::IsNullOrEmpty("$Versions")) { 'Warning' }
                                    elseif ("$Versions" -match 'SMB2\.1') { 'Warning' } else { 'Pass' }
                    [void]$AllChecks.Add((New-CheckResult -Id "PROF-SMBVER-$($SA.StorageAccountName)" `
                        -Category 'FSLogix & Profiles' -Name 'SMB Minimum Version' `
                        -Description 'SMB 2.1 should be disabled - require SMB 3.0+' `
                        -Status $SmbVerStatus -Severity 'High' `
                        -Details "$(if ([string]::IsNullOrEmpty("$Versions")) { "Versions: (not set) - $SmbUnsetDetail" } else { "Versions: $Versions" })" `
                        -Recommendation 'Explicitly restrict SMB versions to SMB3.0/SMB3.1.1 in the file service protocol settings.' `
                        -Reference 'https://learn.microsoft.com/en-us/azure/storage/files/files-smb-protocol#smb-security-settings'))

                    # Kerberos encryption (PROF-022)
                    $KerbEnc = if ($Smb) { $Smb.KerberosTicketEncryption } else { $null }
                    $KerbStatus = if ([string]::IsNullOrEmpty("$KerbEnc")) { 'Warning' }
                                  elseif ("$KerbEnc" -match 'RC4') { 'Warning' } else { 'Pass' }
                    [void]$AllChecks.Add((New-CheckResult -Id "PROF-KERB-$($SA.StorageAccountName)" `
                        -Category 'FSLogix & Profiles' -Name 'Kerberos Ticket Encryption' `
                        -Description 'RC4-HMAC should be disabled - use AES-256 only' `
                        -Status $KerbStatus -Severity 'Medium' `
                        -Details "$(if ([string]::IsNullOrEmpty("$KerbEnc")) { "KerberosEncryption: (not set) - $SmbUnsetDetail" } else { "KerberosEncryption: $KerbEnc" })" `
                        -Recommendation 'Explicitly set Kerberos ticket encryption to AES-256 (RC4-HMAC disabled).' `
                        -Reference 'https://learn.microsoft.com/en-us/azure/storage/files/files-smb-protocol#smb-security-settings'))

                    # Auth methods (PROF-023)
                    $AuthMethods = if ($Smb) { $Smb.AuthenticationMethods } else { $null }
                    $AuthStatus = if ([string]::IsNullOrEmpty("$AuthMethods")) { 'Warning' }
                                  elseif ("$AuthMethods" -match 'NTLMv2') { 'Warning' } else { 'Pass' }
                    [void]$AllChecks.Add((New-CheckResult -Id "PROF-AUTH-$($SA.StorageAccountName)" `
                        -Category 'FSLogix & Profiles' -Name 'Authentication Methods' `
                        -Description 'NTLMv2 should be disabled - use Kerberos only' `
                        -Status $AuthStatus -Severity 'Medium' `
                        -Details "$(if ([string]::IsNullOrEmpty("$AuthMethods")) { "AuthMethods: (not set) - $SmbUnsetDetail" } else { "AuthMethods: $AuthMethods" })" `
                        -Recommendation 'Explicitly restrict authentication methods to Kerberos (NTLMv2 disabled).' `
                        -Reference 'https://learn.microsoft.com/en-us/azure/storage/files/files-smb-protocol#smb-security-settings'))

                    # Channel encryption
                    $ChanEnc = if ($Smb) { $Smb.ChannelEncryption } else { $null }
                    $ChanStatus = if ([string]::IsNullOrEmpty("$ChanEnc")) { 'Warning' }
                                  elseif ("$ChanEnc" -match 'AES-256') { 'Pass' } else { 'Warning' }
                    [void]$AllChecks.Add((New-CheckResult -Id "PROF-SMBENC-$($SA.StorageAccountName)" `
                        -Category 'FSLogix & Profiles' -Name 'SMB Channel Encryption' `
                        -Description 'AES-256-GCM preferred for SMB channel encryption' `
                        -Status $ChanStatus -Severity 'Medium' `
                        -Details "$(if ([string]::IsNullOrEmpty("$ChanEnc")) { "ChannelEncryption: (not set) - $SmbUnsetDetail" } else { "ChannelEncryption: $ChanEnc" })" `
                        -Recommendation 'Explicitly configure SMB channel encryption to include AES-256-GCM.' `
                        -Reference 'https://learn.microsoft.com/en-us/azure/storage/files/files-smb-protocol#smb-security-settings'))
                } catch {
                    [void]$AllChecks.Add((New-CheckResult -Id "PROF-SD-$($SA.StorageAccountName)" `
                        -Category 'FSLogix & Profiles' -Name 'Soft Delete Enabled' `
                        -Description 'File share soft delete protects against accidental deletion' `
                        -Status 'Error' -Severity 'Medium' `
                        -Details "Could not read file service properties: $($_.Exception.Message)" `
                        -Reference 'https://learn.microsoft.com/en-us/azure/storage/files/storage-files-enable-soft-delete'))
                }
            }
        }
        $FSLogixAccts = @($Discovery.Inventory.StorageAccounts | Where-Object { $_.LikelyFSLogix })
        Write-Status "  Storage accounts: $($StorageAccounts.Count), FSLogix candidates: $($FSLogixAccts.Count)" -Level $(if ($FSLogixAccts.Count -gt 0) { 'SUCCESS' } else { 'WARN' })
    } catch {
        Write-Status "  Error: $($_.Exception.Message)" -Level 'ERROR'
        $Discovery.Errors += "Storage discovery failed: $($_.Exception.Message)"
    }
}

# ═══════════════════════════════════════════════════════════════════════════
# PER-SUBSCRIPTION SWEEP (A-1)
# Orphaned disks/NICs, Key Vaults, policy, alerts, quota, capacity reservations,
# budgets, RI, Network Watcher and Private DNS previously ran only against the
# LAST subscription's context. This sweep re-runs them per subscription and
# suffixes singleton check IDs with the sub short-id so results stay unique.
# ═══════════════════════════════════════════════════════════════════════════

foreach ($SubEntry in $Discovery.Subscriptions) {
    $SubId = $SubEntry.Id
    $SubShort = ($SubId -split '-')[0]
    Write-Status "Subscription sweep: $($SubEntry.Name)" -Level 'SECTION'
    try {
        Set-AzContext -SubscriptionId $SubId -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null
    } catch {
        Write-Status "  Could not switch context to $SubId : $($_.Exception.Message)" -Level 'ERROR'
        $Discovery.Errors += "Sweep failed for subscription $SubId : $($_.Exception.Message)"
        continue
    }

    # Scope RG/region lists to resources discovered in THIS subscription
    $SubPattern = "(?i)/subscriptions/$SubId/"
    $AvdResourceGroups = @(
        @($Discovery.Inventory.HostPools | Where-Object { $_.SubscriptionId -eq $SubId } | ForEach-Object { $_.ResourceGroup }) +
        @($Discovery.Inventory.SessionHosts | Where-Object { "$($_.ResourceId)" -match $SubPattern } | ForEach-Object { $_.ResourceGroup }) +
        @($Discovery.Inventory.StorageAccounts | Where-Object { "$($_.Id)" -match $SubPattern } | ForEach-Object { $_.ResourceGroup })
    ) | Sort-Object -Unique
    $AvdRegions = @($Discovery.Inventory.SessionHosts | Where-Object { "$($_.ResourceId)" -match $SubPattern } | ForEach-Object { $_.Location } | Sort-Object -Unique)

# ─── ORPHANED DISKS (GOV-011) ──────────────────────────────────────────
Write-Status "Orphaned Resources" -Level 'SECTION'
try {
    $OrphanedDisks = @(Get-AzDisk -ErrorAction SilentlyContinue |
        Where-Object { $_.ManagedBy -eq $null -and $_.DiskState -eq 'Unattached' -and $_.ResourceGroupName -in $AvdResourceGroups })
    Write-Status "  Orphaned disks: $($OrphanedDisks.Count)" -Level $(if ($OrphanedDisks.Count -gt 0) { 'WARN' } else { 'SUCCESS' })
    foreach ($Disk in $OrphanedDisks) {
        $Discovery.Inventory.OrphanedDisks += [PSCustomObject]@{
            Name          = $Disk.Name
            ResourceGroup = $Disk.ResourceGroupName
            SizeGB        = $Disk.DiskSizeGB
            Sku           = $Disk.Sku.Name
            Location      = $Disk.Location
        }
    }
    [void]$AllChecks.Add((New-CheckResult -Id "GOV-ORPHDISK-$SubShort" `
        -Category 'Governance & Cost' -Name 'Orphaned Disks Detected' `
        -Description 'No unattached managed disks should exist in AVD resource groups' `
        -Status $(if ($OrphanedDisks.Count -eq 0) { 'Pass' } else { 'Warning' }) `
        -Severity 'Medium' `
        -Details "OrphanedDisks: $($OrphanedDisks.Count)$(if ($OrphanedDisks.Count -gt 0) { " ($( ($OrphanedDisks | ForEach-Object { "$($_.Name) $($_.DiskSizeGB)GB" }) -join ', '))" })" `
        -Recommendation 'Review and delete orphaned disks to reduce costs and limit data exposure.' `
        -Reference 'https://learn.microsoft.com/en-us/azure/cost-management-billing/costs/cost-mgt-best-practices' `
        -Evidence @{ Count = $OrphanedDisks.Count }))
} catch {
    Write-Status "  Disk check error: $($_.Exception.Message)" -Level 'ERROR'
    $Discovery.Errors += "Orphaned disk check failed: $($_.Exception.Message)"
}

# ─── ORPHANED NICs (GOV-012) ───────────────────────────────────────────
try {
    $OrphanedNICs = @(Get-AzNetworkInterface -ErrorAction SilentlyContinue |
        Where-Object { $_.VirtualMachine -eq $null -and $_.PrivateEndpoint -eq $null -and $_.ResourceGroupName -in $AvdResourceGroups })
    Write-Status "  Orphaned NICs: $($OrphanedNICs.Count)" -Level $(if ($OrphanedNICs.Count -gt 0) { 'WARN' } else { 'SUCCESS' })
    foreach ($NIC in $OrphanedNICs) {
        $Discovery.Inventory.OrphanedNICs += [PSCustomObject]@{
            Name          = $NIC.Name
            ResourceGroup = $NIC.ResourceGroupName
            Location      = $NIC.Location
            HasPublicIP   = ($NIC.IpConfigurations | Where-Object { $_.PublicIpAddress }).Count -gt 0
        }
    }
    [void]$AllChecks.Add((New-CheckResult -Id "GOV-ORPHNIC-$SubShort" `
        -Category 'Governance & Cost' -Name 'Orphaned NICs Detected' `
        -Description 'No unattached network interfaces should exist in AVD resource groups' `
        -Status $(if ($OrphanedNICs.Count -eq 0) { 'Pass' } else { 'Warning' }) `
        -Severity 'Low' `
        -Details "OrphanedNICs: $($OrphanedNICs.Count)" `
        -Recommendation 'Delete orphaned NICs - those with public IPs still incur charges.' `
        -Reference 'https://learn.microsoft.com/en-us/azure/cost-management-billing/costs/cost-mgt-best-practices' `
        -Evidence @{ Count = $OrphanedNICs.Count }))
} catch {
    Write-Status "  NIC check error: $($_.Exception.Message)" -Level 'ERROR'
    $Discovery.Errors += "Orphaned NIC check failed: $($_.Exception.Message)"
}

# ─── KEY VAULTS (SEC-024) ──────────────────────────────────────────────
Write-Status "Key Vaults" -Level 'SECTION'
try {
    $KeyVaults = @()
    foreach ($RG in $AvdResourceGroups) {
        $KeyVaults += @(Get-AzKeyVault -ResourceGroupName $RG -ErrorAction SilentlyContinue -WarningAction SilentlyContinue)
    }
    Write-Status "  Found $($KeyVaults.Count) Key Vault(s) in AVD resource groups" -Level $(if ($KeyVaults.Count -gt 0) { 'SUCCESS' } else { 'WARN' })
    $KVsWithoutPE = 0
    foreach ($KV in $KeyVaults) {
        # Get detailed KV resource for PE info
        $KVDetail = Get-AzResource -ResourceId $KV.ResourceId -ErrorAction SilentlyContinue
        $KVHasPE = $false
        if ($KVDetail -and $KVDetail.Properties.privateEndpointConnections) {
            $KVHasPE = @($KVDetail.Properties.privateEndpointConnections).Count -gt 0
        }
        if (-not $KVHasPE) { $KVsWithoutPE++ }
        $Discovery.Inventory.KeyVaults += [PSCustomObject]@{
            Name              = $KV.VaultName
            ResourceGroup     = $KV.ResourceGroupName
            Location          = $KV.Location
            SoftDeleteEnabled = $KV.EnableSoftDelete
            PurgeProtection   = $KV.EnablePurgeProtection
            HasPrivateEndpoint = $KVHasPE
        }
    }
    [void]$AllChecks.Add((New-CheckResult -Id "SEC-KEYVAULT-$SubShort" `
        -Category 'Security & IAM' -Name 'Key Vault for Secrets' `
        -Description 'Azure Key Vault should exist in AVD resource groups for certificate and secret management' `
        -Status $(if ($KeyVaults.Count -gt 0) { 'Pass' } else { 'Warning' }) `
        -Severity 'Medium' `
        -Details "KeyVaults: $($KeyVaults.Count)$(if ($KeyVaults.Count -gt 0) { " ($( ($KeyVaults | ForEach-Object { $_.VaultName }) -join ', '))" })" `
        -Recommendation 'Deploy Azure Key Vault for centralized secret, certificate, and key management.' `
        -Reference 'https://learn.microsoft.com/en-us/azure/key-vault/general/overview' `
        -Evidence @{ Count = $KeyVaults.Count }))

    # ─── CHECK: Key Vault Private Endpoint (SEC-023) ───
    if ($KeyVaults.Count -gt 0) {
        [void]$AllChecks.Add((New-CheckResult -Id "SEC-KVPE-$SubShort" `
            -Category 'Security' -Name 'Key Vault Private Endpoint' `
            -Description 'Key Vaults should use private endpoints to prevent secret exposure over public networks' `
            -Status $(if ($KVsWithoutPE -eq 0) { 'Pass' } else { 'Warning' }) `
            -Severity 'High' `
            -Details "KeyVaults: $($KeyVaults.Count), Without PE: $KVsWithoutPE$(if ($KVsWithoutPE -gt 0) { " ($( ($Discovery.Inventory.KeyVaults | Where-Object { -not $_.HasPrivateEndpoint } | ForEach-Object { $_.Name }) -join ', '))" })" `
            -Recommendation 'Configure private endpoints for Key Vaults with privatelink.vaultcore.azure.net DNS zone.' `
            -Reference 'https://learn.microsoft.com/en-us/azure/key-vault/general/private-link-service' `
            -Evidence @{ Total = $KeyVaults.Count; WithoutPE = $KVsWithoutPE }))
    }
} catch {
    Write-Status "  Key Vault error: $($_.Exception.Message)" -Level 'ERROR'
    $Discovery.Errors += "Key Vault discovery failed: $($_.Exception.Message)"
}

# ─── NETWORK WATCHER (NET-019) ─────────────────────────────────────────
Write-Status "Network Watcher" -Level 'SECTION'
try {
    $AllWatchers = @(Get-AzNetworkWatcher -ErrorAction SilentlyContinue)
    $WatcherRegions = @($AllWatchers | ForEach-Object { $_.Location })
    $MissingRegions = @($AvdRegions | Where-Object { $_ -notin $WatcherRegions })

    foreach ($NW in $AllWatchers | Where-Object { $_.Location -in $AvdRegions }) {
        $Discovery.Inventory.NetworkWatchers += [PSCustomObject]@{
            Name              = $NW.Name
            Region            = $NW.Location
            ProvisioningState = $NW.ProvisioningState
        }
    }
    Write-Status "  Network Watchers in AVD regions: $($Discovery.Inventory.NetworkWatchers.Count), Missing: $($MissingRegions.Count)" -Level $(if ($MissingRegions.Count -eq 0) { 'SUCCESS' } else { 'WARN' })
    [void]$AllChecks.Add((New-CheckResult -Id "NET-NETWATCHER-$SubShort" `
        -Category 'Networking' -Name 'Network Watcher Enabled' `
        -Description 'Network Watcher should be enabled in each region where AVD session hosts are deployed' `
        -Status $(if ($MissingRegions.Count -eq 0) { 'Pass' } else { 'Warning' }) `
        -Severity 'Low' `
        -Details "AVD regions: $($AvdRegions -join ', '), Missing Network Watcher: $(if ($MissingRegions.Count -gt 0) { $MissingRegions -join ', ' } else { 'None' })" `
        -Recommendation 'Enable Network Watcher in all AVD regions for packet capture, flow logs, and connectivity troubleshooting.' `
        -Reference 'https://learn.microsoft.com/en-us/azure/network-watcher/network-watcher-overview' `
        -Evidence @{ MissingRegions = $MissingRegions }))
} catch {
    Write-Status "  Network Watcher error: $($_.Exception.Message)" -Level 'ERROR'
    $Discovery.Errors += "Network Watcher discovery failed: $($_.Exception.Message)"
}

# ─── PRIVATE DNS ZONES (NET-018) ──────────────────────────────────────
Write-Status "Private DNS Zones" -Level 'SECTION'
try {
    $PrivateDnsZones = @(Get-AzPrivateDnsZone -ErrorAction SilentlyContinue)
    $FileZone = $PrivateDnsZones | Where-Object { $_.Name -eq 'privatelink.file.core.windows.net' }
    $LinkedToAvdVNet = $false

    if ($FileZone) {
        $AvdVNetIds = @($Discovery.Inventory.VNets | ForEach-Object { $_.Id })
        foreach ($FZ in @($FileZone)) {
            $Links = @(Get-AzPrivateDnsVirtualNetworkLink -ZoneName $FZ.Name -ResourceGroupName $FZ.ResourceGroupName -ErrorAction SilentlyContinue)
            foreach ($Link in $Links) {
                if ($Link.VirtualNetworkId -in $AvdVNetIds) {
                    $LinkedToAvdVNet = $true
                }
            }
            $Discovery.Inventory.PrivateDnsZones += [PSCustomObject]@{
                Name          = $FZ.Name
                ResourceGroup = $FZ.ResourceGroupName
                VNetLinks     = @($Links | ForEach-Object { [PSCustomObject]@{ Name = $_.Name; VNetId = $_.VirtualNetworkId; Status = $_.VirtualNetworkLinkState } })
            }
        }
    }
    Write-Status "  privatelink.file.core.windows.net: $(if ($FileZone) { 'Found' } else { 'Not found' }), Linked to AVD VNet: $LinkedToAvdVNet" -Level $(if ($LinkedToAvdVNet) { 'SUCCESS' } else { 'WARN' })
    [void]$AllChecks.Add((New-CheckResult -Id "NET-PRIVDNS-$SubShort" `
        -Category 'Networking' -Name 'Private DNS Zone Linked' `
        -Description 'Private DNS zone (privatelink.file.core.windows.net) should exist and be linked to AVD VNets' `
        -Status $(if ($LinkedToAvdVNet) { 'Pass' } elseif ($FileZone) { 'Warning' } else { 'Fail' }) `
        -Severity 'Medium' `
        -Details "Zone exists: $(if ($FileZone) { 'Yes' } else { 'No' }), Linked to AVD VNet: $LinkedToAvdVNet" `
        -Recommendation 'Create privatelink.file.core.windows.net DNS zone and link to AVD VNets for private endpoint resolution.' `
        -Reference 'https://learn.microsoft.com/en-us/azure/private-link/private-endpoint-dns' `
        -Evidence @{ ZoneExists = [bool]$FileZone; LinkedToAvdVNet = $LinkedToAvdVNet }))
} catch {
    Write-Status "  Private DNS error: $($_.Exception.Message)" -Level 'ERROR'
    $Discovery.Errors += "Private DNS zone discovery failed: $($_.Exception.Message)"
}

# ─── AZURE POLICY ASSIGNMENTS (GOV-013) ───────────────────────────────
Write-Status "Azure Policy" -Level 'SECTION'
try {
    $PolicyAssignments = @()
    foreach ($RG in $AvdResourceGroups) {
        $Scope = "/subscriptions/$SubId/resourceGroups/$RG"
        $PolicyAssignments += @(Get-AzPolicyAssignment -Scope $Scope -ErrorAction SilentlyContinue -WarningAction SilentlyContinue)
    }
    # Deduplicate by assignment ID
    $PolicyAssignments = @($PolicyAssignments | Sort-Object -Property PolicyAssignmentId -Unique)
    Write-Status "  Policy assignments on AVD RGs: $($PolicyAssignments.Count)" -Level $(if ($PolicyAssignments.Count -gt 0) { 'SUCCESS' } else { 'WARN' })
    foreach ($PA in $PolicyAssignments) {
        $Discovery.Inventory.PolicyAssignments += [PSCustomObject]@{
            Name        = $PA.Name
            DisplayName = $PA.Properties.DisplayName
            Scope       = $PA.Properties.Scope
            PolicyDefId = $PA.Properties.PolicyDefinitionId
        }
    }
    [void]$AllChecks.Add((New-CheckResult -Id "GOV-POLICY-$SubShort" `
        -Category 'Governance & Cost' -Name 'Azure Policy Assignments' `
        -Description 'At least one Azure Policy assignment should exist on AVD resource groups for governance' `
        -Status $(if ($PolicyAssignments.Count -gt 0) { 'Pass' } else { 'Warning' }) `
        -Severity 'Medium' `
        -Details "PolicyAssignments: $($PolicyAssignments.Count)$(if ($PolicyAssignments.Count -gt 0) { " ($( ($PolicyAssignments | ForEach-Object { $_.Properties.DisplayName } | Select-Object -First 5) -join ', '))" })" `
        -Recommendation 'Assign Azure Policy to enforce governance guardrails (allowed SKUs, required tags, encryption, etc.).' `
        -Reference 'https://learn.microsoft.com/en-us/azure/governance/policy/overview' `
        -Evidence @{ Count = $PolicyAssignments.Count }))
} catch {
    Write-Status "  Policy error: $($_.Exception.Message)" -Level 'ERROR'
    $Discovery.Errors += "Policy assignment discovery failed: $($_.Exception.Message)"
}

# ─── ALERT RULES (MON-015) ────────────────────────────────────────────
Write-Status "Alert Rules" -Level 'SECTION'
try {
    $AlertRules = @()
    $LogAlertRules = @()
    foreach ($RG in $AvdResourceGroups) {
        $AlertRules += @(Get-AzMetricAlertRuleV2 -ResourceGroupName $RG -ErrorAction SilentlyContinue -WarningAction SilentlyContinue)
        # Also count scheduled-query (log) alerts - MON-015 says "metric or log" (C-6).
        try {
            $LogAlertRules += @(Get-AzScheduledQueryRule -ResourceGroupName $RG -ErrorAction SilentlyContinue -WarningAction SilentlyContinue)
        } catch {
            Write-Status "    Could not enumerate log alert rules in RG $RG : $($_.Exception.Message)" -Level 'WARN'
        }
    }
    $TotalAlertRules = $AlertRules.Count + $LogAlertRules.Count
    Write-Status "  Alert rules in AVD RGs: $($AlertRules.Count) metric, $($LogAlertRules.Count) log" -Level $(if ($TotalAlertRules -gt 0) { 'SUCCESS' } else { 'WARN' })
    foreach ($AR in $AlertRules) {
        $Discovery.Inventory.AlertRules += [PSCustomObject]@{
            Name           = $AR.Name
            Type           = 'Metric'
            ResourceGroup  = ($AR.Id -split '/')[4]
            Severity       = $AR.Severity
            Enabled        = $AR.Enabled
            TargetResource = ($AR.Scopes | Select-Object -First 1)
        }
    }
    foreach ($AR in $LogAlertRules) {
        $Discovery.Inventory.AlertRules += [PSCustomObject]@{
            Name           = $AR.Name
            Type           = 'Log'
            ResourceGroup  = ($AR.Id -split '/')[4]
            Severity       = $AR.Severity
            Enabled        = $AR.Enabled
            TargetResource = ($AR.Scope | Select-Object -First 1)
        }
    }
    [void]$AllChecks.Add((New-CheckResult -Id "MON-ALERTS-$SubShort" `
        -Category 'Monitoring' -Name 'Alert Rules Configured' `
        -Description 'Metric or log alert rules should exist targeting AVD resources' `
        -Status $(if ($TotalAlertRules -gt 0) { 'Pass' } else { 'Warning' }) `
        -Severity 'Medium' `
        -Details "AlertRules: $TotalAlertRules ($($AlertRules.Count) metric, $($LogAlertRules.Count) log)$(if ($TotalAlertRules -gt 0) { " ($( (@($AlertRules) + @($LogAlertRules) | ForEach-Object { $_.Name } | Select-Object -First 3) -join ', '))" })" `
        -Recommendation 'Configure metric alerts for CPU, memory, disk, and log alerts for AVD-specific health signals.' `
        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/set-up-service-alerts' `
        -Evidence @{ MetricCount = $AlertRules.Count; LogCount = $LogAlertRules.Count }))
} catch {
    Write-Status "  Alert rules error: $($_.Exception.Message)" -Level 'ERROR'
    $Discovery.Errors += "Alert rule discovery failed: $($_.Exception.Message)"
}

# ─── VM QUOTA USAGE (GOV-014) ─────────────────────────────────────────
Write-Status "VM Quota Usage" -Level 'SECTION'
try {
    $QuotaWarning = $false
    foreach ($Region in $AvdRegions) {
        $Usages = @(Get-AzVMUsage -Location $Region -ErrorAction SilentlyContinue)
        # Check vCPU family quotas used by AVD session hosts
        $VMSizes = @($Discovery.Inventory.SessionHosts | Where-Object { $_.Location -eq $Region } | ForEach-Object { $_.VMSize } | Sort-Object -Unique)
        foreach ($Usage in $Usages) {
            if ($Usage.Limit -gt 0) {
                $Pct = [math]::Round($Usage.CurrentValue / $Usage.Limit * 100, 0)
                if ($Pct -ge 70 -or $Usage.Name.LocalizedValue -eq 'Total Regional vCPUs') {
                    $Discovery.Inventory.Quotas += [PSCustomObject]@{
                        Region     = $Region
                        Name       = $Usage.Name.LocalizedValue
                        Current    = $Usage.CurrentValue
                        Limit      = $Usage.Limit
                        UsagePct   = $Pct
                    }
                    if ($Pct -ge 80) { $QuotaWarning = $true }
                }
            }
        }
    }
    $HighUsage = @($Discovery.Inventory.Quotas | Where-Object { $_.UsagePct -ge 80 })
    Write-Status "  Quota entries tracked: $($Discovery.Inventory.Quotas.Count), High usage (>80%): $($HighUsage.Count)" -Level $(if ($QuotaWarning) { 'WARN' } else { 'SUCCESS' })
    [void]$AllChecks.Add((New-CheckResult -Id "GOV-QUOTA-$SubShort" `
        -Category 'Governance & Cost' -Name 'VM Quota Headroom' `
        -Description 'vCPU quota per region should have headroom for scaling' `
        -Status $(if ($HighUsage.Count -gt 0) { 'Warning' } else { 'Pass' }) `
        -Severity 'High' `
        -Details "Regions checked: $($AvdRegions.Count), High usage quotas (>80%): $($HighUsage.Count)$(if ($HighUsage.Count -gt 0) { " ($( ($HighUsage | ForEach-Object { "$($_.Region)/$($_.Name): $($_.UsagePct)%" }) -join ', '))" })" `
        -Recommendation 'Request quota increase for VM families used in AVD before reaching limits.' `
        -Reference 'https://learn.microsoft.com/en-us/azure/quotas/per-vm-quota-requests' `
        -Evidence @{ HighUsageCount = $HighUsage.Count; Regions = $AvdRegions }))
} catch {
    Write-Status "  Quota check error: $($_.Exception.Message)" -Level 'WARN'
}

# ─── CAPACITY RESERVATIONS (GOV-015) ──────────────────────────────────
Write-Status "Capacity Reservations" -Level 'SECTION'
try {
    $CRGs = @()
    foreach ($RG in $AvdResourceGroups) {
        $CRGs += @(Get-AzCapacityReservationGroup -ResourceGroupName $RG -ErrorAction SilentlyContinue)
    }
    foreach ($CRG in $CRGs) {
        $Discovery.Inventory.CapacityReservations += [PSCustomObject]@{
            Name          = $CRG.Name
            ResourceGroup = ($CRG.Id -split '/')[4]
            Location      = $CRG.Location
            Zones         = $CRG.Zones
        }
    }
    Write-Status "  Capacity Reservation Groups: $($CRGs.Count)" -Level $(if ($CRGs.Count -gt 0) { 'SUCCESS' } else { 'WARN' })
    [void]$AllChecks.Add((New-CheckResult -Id "GOV-CAPRESERV-$SubShort" `
        -Category 'Governance & Cost' -Name 'Capacity Reservation' `
        -Description 'Capacity Reservation Groups guarantee VM availability for critical workloads' `
        -Status $(if ($CRGs.Count -gt 0) { 'Pass' } else { 'Warning' }) `
        -Severity 'Low' `
        -Details "CapacityReservationGroups: $($CRGs.Count)$(if ($CRGs.Count -gt 0) { " ($( ($CRGs | ForEach-Object { $_.Name }) -join ', '))" })" `
        -Recommendation 'Consider Capacity Reservation Groups for mission-critical AVD pools to prevent allocation failures.' `
        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-machines/capacity-reservation-overview' `
        -Evidence @{ Count = $CRGs.Count }))
} catch {
    Write-Status "  Capacity reservation error: $($_.Exception.Message)" -Level 'WARN'
}

# ─── BUDGETS (GOV-016) ────────────────────────────────────────────────
Write-Status "Cost Budgets" -Level 'SECTION'
try {
    # Use Invoke-AzRestMethod to check for budgets at subscription scope (avoids Az.Billing dependency)
    $SubScope = "/subscriptions/$SubId"
    $BudgetResponse = Invoke-AzRestMethod -Path "$SubScope/providers/Microsoft.Consumption/budgets?api-version=2023-05-01" -Method GET -ErrorAction SilentlyContinue
    $Budgets = @()
    if ($BudgetResponse -and $BudgetResponse.StatusCode -eq 200) {
        $BudgetData = $BudgetResponse.Content | ConvertFrom-Json -ErrorAction SilentlyContinue
        if ($BudgetData -and $BudgetData.value) {
            $Budgets = @($BudgetData.value)
        }
    }
    foreach ($B in $Budgets) {
        $Discovery.Inventory.Budgets += [PSCustomObject]@{
            Name       = $B.name
            Amount     = $B.properties.amount
            TimeGrain  = $B.properties.timeGrain
            Category   = $B.properties.category
        }
    }
    $HasAlertThresholds = @($Budgets | Where-Object { $_.properties.notifications }).Count -gt 0
    Write-Status "  Budgets: $($Budgets.Count), With alerts: $HasAlertThresholds" -Level $(if ($Budgets.Count -gt 0 -and $HasAlertThresholds) { 'SUCCESS' } else { 'WARN' })
    [void]$AllChecks.Add((New-CheckResult -Id "GOV-BUDGET-$SubShort" `
        -Category 'Governance & Cost' -Name 'Budget Alerts Configured' `
        -Description 'Azure cost budgets with alert thresholds should be configured' `
        -Status $(if ($Budgets.Count -gt 0 -and $HasAlertThresholds) { 'Pass' } elseif ($Budgets.Count -gt 0) { 'Warning' } else { 'Fail' }) `
        -Severity 'Medium' `
        -Details "Budgets: $($Budgets.Count)$(if ($Budgets.Count -gt 0) { " ($( ($Budgets | ForEach-Object { "$($_.name) `$$($_.properties.amount)" }) -join ', '))" }), AlertThresholds: $HasAlertThresholds" `
        -Recommendation 'Create cost budgets with alert thresholds at 80% and 100% to catch spend anomalies early.' `
        -Reference 'https://learn.microsoft.com/en-us/azure/cost-management-billing/costs/tutorial-acm-create-budgets' `
        -Evidence @{ Count = $Budgets.Count; HasAlerts = $HasAlertThresholds }))
} catch {
    Write-Status "  Budget check error: $($_.Exception.Message)" -Level 'WARN'
}

# ─── RESERVED INSTANCES (GOV-017) ─────────────────────────────────────
Write-Status "Reserved Instances" -Level 'SECTION'
try {
    $RIResponse = Invoke-AzRestMethod -Path "/providers/Microsoft.Capacity/reservationOrders?api-version=2022-11-01" -Method GET -ErrorAction SilentlyContinue
    $Reservations = @()
    if ($RIResponse -and $RIResponse.StatusCode -eq 200) {
        $RIData = $RIResponse.Content | ConvertFrom-Json -ErrorAction SilentlyContinue
        if ($RIData -and $RIData.value) {
            # Filter to VM reservations relevant to this subscription
            $Reservations = @($RIData.value | Where-Object {
                $_.properties.reservedResourceType -eq 'VirtualMachines'
            })
        }
    }
    foreach ($RI in $Reservations) {
        $Discovery.Inventory.Reservations += [PSCustomObject]@{
            Name        = $RI.name
            DisplayName = $RI.properties.displayName
            Term        = $RI.properties.term
            Quantity    = $RI.properties.quantity
        }
    }
    Write-Status "  VM Reservations/Orders: $($Reservations.Count)" -Level $(if ($Reservations.Count -gt 0) { 'SUCCESS' } else { 'WARN' })
    # Check if there are personal/always-on host pools that would benefit
    $AlwaysOnHPs = @($Discovery.Inventory.HostPools | Where-Object { $_.HostPoolType -eq 'Personal' })
    [void]$AllChecks.Add((New-CheckResult -Id "GOV-RI-$SubShort" `
        -Category 'Governance & Cost' -Name 'Reserved Instance Coverage' `
        -Description 'Evaluate RI or Savings Plans for always-on or personal host pools' `
        -Status $(if ($Reservations.Count -gt 0) { 'Pass' } elseif ($AlwaysOnHPs.Count -gt 0) { 'Warning' } else { 'Pass' }) `
        -Severity 'Low' `
        -Details "VMReservations: $($Reservations.Count), PersonalHostPools: $($AlwaysOnHPs.Count)" `
        -Recommendation 'Evaluate Azure Reserved Instances (1yr or 3yr) for personal host pools to save 40-72% on compute.' `
        -Reference 'https://learn.microsoft.com/en-us/azure/cost-management-billing/reservations/save-compute-costs-reservations' `
        -Evidence @{ Reservations = $Reservations.Count; PersonalPools = $AlwaysOnHPs.Count }))
} catch {
    Write-Status "  RI check error: $($_.Exception.Message)" -Level 'WARN'
}

# ─── RESOURCE LOCKS (GOV-007) ─────────────────────────────────────────
Write-Status "Resource Locks" -Level 'SECTION'
try {
    $Locks = @()
    foreach ($RG in $AvdResourceGroups) {
        $Locks += @(Get-AzResourceLock -ResourceGroupName $RG -ErrorAction SilentlyContinue)
    }
    Write-Status "  Resource locks on AVD RGs: $($Locks.Count)" -Level $(if ($Locks.Count -gt 0) { 'SUCCESS' } else { 'WARN' })
    [void]$AllChecks.Add((New-CheckResult -Id "GOV-LOCK-$SubShort" `
        -Category 'Governance & Cost' -Name 'Resource Locks Applied' `
        -Description 'CanNotDelete/ReadOnly locks should protect critical AVD resources from accidental deletion' `
        -Status $(if ($Locks.Count -gt 0) { 'Pass' } else { 'Warning' }) `
        -Severity 'Medium' `
        -Details "Locks on AVD resource groups: $($Locks.Count)$(if ($Locks.Count -gt 0) { " ($( @($Locks | ForEach-Object { "$($_.Name) [$($_.Properties.level)]" } | Select-Object -First 5) -join ', '))" })" `
        -Recommendation 'Apply CanNotDelete locks to host pools, workspaces, Log Analytics, and profile storage to prevent accidental deletion.' `
        -Reference 'https://learn.microsoft.com/en-us/azure/azure-resource-manager/management/lock-resources' `
        -Evidence @{ LockCount = $Locks.Count }))
} catch {
    Write-Status "  Resource lock error: $($_.Exception.Message)" -Level 'WARN'
}

# ─── SERVICE HEALTH ALERTS (MON-003) ──────────────────────────────────
Write-Status "Service Health Alerts" -Level 'SECTION'
try {
    $ActivityAlerts = @(Get-AzActivityLogAlert -ErrorAction SilentlyContinue)
    $SvcHealthAlerts = @($ActivityAlerts | Where-Object {
        $Cond = if ($_.Condition) { $_.Condition } elseif ($_.ConditionAllOf) { $_.ConditionAllOf } else { $_ }
        ("$($Cond | ConvertTo-Json -Depth 6 -Compress)" -match 'ServiceHealth')
    })
    Write-Status "  Service Health alert rules: $($SvcHealthAlerts.Count) of $($ActivityAlerts.Count) activity log alerts" -Level $(if ($SvcHealthAlerts.Count -gt 0) { 'SUCCESS' } else { 'WARN' })
    [void]$AllChecks.Add((New-CheckResult -Id "MON-SVCHEALTH-$SubShort" `
        -Category 'Monitoring' -Name 'Service Health Alerts' `
        -Description 'Azure Service Health alert rules should notify on AVD service issues, planned maintenance, and health advisories' `
        -Status $(if ($SvcHealthAlerts.Count -gt 0) { 'Pass' } else { 'Warning' }) `
        -Severity 'Medium' `
        -Details "ServiceHealth activity-log alerts: $($SvcHealthAlerts.Count) (total activity-log alerts: $($ActivityAlerts.Count))" `
        -Recommendation 'Create a Service Health alert (category=ServiceHealth) targeting the Azure Virtual Desktop service to receive maintenance and incident notifications.' `
        -Reference 'https://learn.microsoft.com/en-us/azure/service-health/alerts-activity-log-service-notifications-portal' `
        -Evidence @{ ServiceHealthAlerts = $SvcHealthAlerts.Count; TotalActivityAlerts = $ActivityAlerts.Count }))
} catch {
    Write-Status "  Service Health alert error: $($_.Exception.Message)" -Level 'WARN'
}

# ─── DEFENDER SECURE SCORE (SEC-005) ──────────────────────────────────
Write-Status "Secure Score" -Level 'SECTION'
try {
    $SecureScores = @(Get-AzSecuritySecureScore -ErrorAction Stop)
    $AscScore = $SecureScores | Where-Object { $_.Name -eq 'ascScore' } | Select-Object -First 1
    if (-not $AscScore) { $AscScore = $SecureScores | Select-Object -First 1 }
    $ScorePct = $null
    if ($AscScore) {
        if ($null -ne $AscScore.Percentage) { $ScorePct = [math]::Round([double]$AscScore.Percentage * 100, 0) }
        elseif ($AscScore.Score -and $null -ne $AscScore.Score.Percentage) { $ScorePct = [math]::Round([double]$AscScore.Score.Percentage * 100, 0) }
    }
    if ($null -eq $ScorePct) {
        [void]$AllChecks.Add((New-CheckResult -Id "SEC-SCORE-$SubShort" `
            -Category 'Security' -Name 'Secure Score Review' `
            -Description 'Microsoft Defender Secure Score for the AVD subscription should be reviewed and improved' `
            -Status 'Error' -Severity 'Medium' `
            -Details 'Secure Score returned no percentage (Defender for Cloud may not be enabled).' `
            -Reference 'https://learn.microsoft.com/en-us/azure/defender-for-cloud/secure-score-security-controls'))
    } else {
        Write-Status "  Defender Secure Score: $ScorePct%" -Level $(if ($ScorePct -ge 70) { 'SUCCESS' } else { 'WARN' })
        [void]$AllChecks.Add((New-CheckResult -Id "SEC-SCORE-$SubShort" `
            -Category 'Security' -Name 'Secure Score Review' `
            -Description 'Microsoft Defender Secure Score for the AVD subscription should be reviewed and improved' `
            -Status $(if ($ScorePct -ge 70) { 'Pass' } else { 'Warning' }) `
            -Severity 'Medium' `
            -Details "Defender Secure Score: $ScorePct%" `
            -Recommendation 'Review Defender for Cloud Secure Score recommendations for the AVD subscription and remediate high-impact controls.' `
            -Reference 'https://learn.microsoft.com/en-us/azure/defender-for-cloud/secure-score-security-controls' `
            -Evidence @{ SecureScorePercent = $ScorePct }))
    }
} catch {
    [void]$AllChecks.Add((New-CheckResult -Id "SEC-SCORE-$SubShort" `
        -Category 'Security' -Name 'Secure Score Review' `
        -Description 'Microsoft Defender Secure Score for the AVD subscription should be reviewed and improved' `
        -Status 'Error' -Severity 'Medium' `
        -Details "Could not read Secure Score (Defender for Cloud required): $($_.Exception.Message)" `
        -Reference 'https://learn.microsoft.com/en-us/azure/defender-for-cloud/secure-score-security-controls'))
}

# ─── JUST-IN-TIME VM ACCESS (SEC-021) ─────────────────────────────────
Write-Status "Just-in-Time VM Access" -Level 'SECTION'
try {
    $SubShHostIds = @($Discovery.Inventory.SessionHosts | Where-Object { "$($_.ResourceId)" -match $SubPattern } | ForEach-Object { "$($_.ResourceId)".ToLower() })
    # Exposure signal (reuse collected data): public IP on a session host, or an RDP-open NSG (NET-RDP Fail).
    $SubHostsPublicIP = @($Discovery.Inventory.SessionHosts | Where-Object { "$($_.ResourceId)" -match $SubPattern -and $_.NicHasPublicIP }).Count -gt 0
    $RdpOpen = @($AllChecks | Where-Object { $_.Id -like 'NET-RDP-*' -and $_.Status -eq 'Fail' }).Count -gt 0
    $Exposed = $SubHostsPublicIP -or $RdpOpen
    $JitPolicies = @(Get-AzJitNetworkAccessPolicy -ErrorAction SilentlyContinue)
    $JitCoveredVMs = @()
    foreach ($JP in $JitPolicies) {
        foreach ($JVM in @($JP.VirtualMachines)) {
            if ("$($JVM.Id)".ToLower() -in $SubShHostIds) { $JitCoveredVMs += "$($JVM.Id)" }
        }
    }
    $JitCovers = $JitCoveredVMs.Count -gt 0
    if ($JitCovers) {
        $JitStatus = 'Pass'; $JitDetail = "JIT policies cover $($JitCoveredVMs.Count) session host VM(s)"
    } elseif ($Exposed) {
        $JitStatus = 'Warning'; $JitDetail = "No JIT policy covers session hosts, and hosts are exposed (PublicIP: $SubHostsPublicIP, RDP-open NSG: $RdpOpen)"
    } else {
        $JitStatus = 'N/A'; $JitDetail = 'No JIT policy covering session hosts, but no exposed hosts detected (no public IPs / no internet-open RDP)'
    }
    Write-Status "  JIT policies: $($JitPolicies.Count), covering session hosts: $JitCovers" -Level $(if ($JitCovers) { 'SUCCESS' } else { 'WARN' })
    [void]$AllChecks.Add((New-CheckResult -Id "SEC-JIT-$SubShort" `
        -Category 'Security' -Name 'Just-in-Time VM Access' `
        -Description 'JIT VM access should keep RDP/SSH closed by default and open on-demand with time-limited, audited requests' `
        -Status $JitStatus -Severity 'Medium' `
        -Details $JitDetail `
        -Recommendation 'Enable Defender for Cloud Just-in-Time VM access on session hosts so management ports are closed by default.' `
        -Reference 'https://learn.microsoft.com/en-us/azure/defender-for-cloud/just-in-time-access-usage' `
        -Evidence @{ JitPolicies = $JitPolicies.Count; CoveredVMs = $JitCoveredVMs.Count; Exposed = $Exposed }))
} catch {
    Write-Status "  JIT check error: $($_.Exception.Message)" -Level 'WARN'
}

# ─── SECURITY BASELINE / REGULATORY COMPLIANCE (SEC-013) ──────────────
Write-Status "Security Baseline" -Level 'SECTION'
try {
    $ComplianceStandards = @(Get-AzSecurityRegulatoryComplianceStandard -ErrorAction Stop)
    $Benchmark = @($ComplianceStandards | Where-Object { $_.Name -match 'Azure-Security-Benchmark|Microsoft-cloud-security-benchmark' }) | Select-Object -First 1
    if ($Benchmark) {
        $Passed = [double]($Benchmark.PassedControls | ForEach-Object { $_ })
        $Failed = [double]($Benchmark.FailedControls | ForEach-Object { $_ })
        $Skipped = [double]($Benchmark.SkippedControls | ForEach-Object { $_ })
        $TotalCtl = $Passed + $Failed + $Skipped
        $CompliancePct = if ($TotalCtl -gt 0) { [math]::Round($Passed / $TotalCtl * 100, 0) } else { -1 }
        Write-Status "  $($Benchmark.Name): $CompliancePct% controls passing" -Level 'SUCCESS'
        [void]$AllChecks.Add((New-CheckResult -Id "SEC-BASELINE-$SubShort" `
            -Category 'Security' -Name 'Azure Security Baseline' `
            -Description 'The AVD subscription should be evaluated against the Microsoft cloud security benchmark' `
            -Status 'Pass' -Severity 'High' `
            -Details "$($Benchmark.Name) present$(if ($CompliancePct -ge 0) { " - $CompliancePct% controls passing ($([int]$Passed) passed / $([int]$Failed) failed)" })" `
            -Recommendation 'Review failed controls in the Microsoft cloud security benchmark for the AVD subscription and remediate.' `
            -Reference 'https://learn.microsoft.com/en-us/security/benchmark/azure/baselines/azure-virtual-desktop-security-baseline' `
            -Evidence @{ Standard = $Benchmark.Name; CompliancePercent = $CompliancePct }))
    } else {
        [void]$AllChecks.Add((New-CheckResult -Id "SEC-BASELINE-$SubShort" `
            -Category 'Security' -Name 'Azure Security Baseline' `
            -Description 'The AVD subscription should be evaluated against the Microsoft cloud security benchmark' `
            -Status 'Warning' -Severity 'High' `
            -Details 'Microsoft cloud security benchmark / Azure Security Benchmark not found in regulatory compliance standards.' `
            -Recommendation 'Enable the Microsoft cloud security benchmark standard in Defender for Cloud regulatory compliance for the AVD subscription.' `
            -Reference 'https://learn.microsoft.com/en-us/security/benchmark/azure/baselines/azure-virtual-desktop-security-baseline'))
    }
} catch {
    [void]$AllChecks.Add((New-CheckResult -Id "SEC-BASELINE-$SubShort" `
        -Category 'Security' -Name 'Azure Security Baseline' `
        -Description 'The AVD subscription should be evaluated against the Microsoft cloud security benchmark' `
        -Status 'Error' -Severity 'High' `
        -Details "Could not read regulatory compliance standards (Defender for Cloud required): $($_.Exception.Message)" `
        -Reference 'https://learn.microsoft.com/en-us/security/benchmark/azure/baselines/azure-virtual-desktop-security-baseline'))
}

# ─── APP ATTACH (CIM) PACKAGES AT SUBSCRIPTION SCOPE (APP-002) ────────
Write-Status "App Attach Packages" -Level 'SECTION'
try {
    $AppAttachResp = Invoke-AzRestMethod -Path "/subscriptions/$SubId/providers/Microsoft.DesktopVirtualization/appAttachPackages?api-version=2024-04-03" -Method GET -ErrorAction SilentlyContinue
    $AppAttachPkgs = @()
    if ($AppAttachResp -and $AppAttachResp.StatusCode -eq 200) {
        $AAData = $AppAttachResp.Content | ConvertFrom-Json -ErrorAction SilentlyContinue
        if ($AAData -and $AAData.value) { $AppAttachPkgs = @($AAData.value) }
    }
    if ($AppAttachPkgs.Count -gt 0) {
        Write-Status "  CIM App Attach packages: $($AppAttachPkgs.Count)" -Level 'SUCCESS'
        [void]$AllChecks.Add((New-CheckResult -Id "APP-ATTACH-CIM-$SubShort" `
            -Category 'Application Delivery' -Name 'App Attach' `
            -Description 'CIM-based App Attach decouples application lifecycle from the golden image' `
            -Status 'Pass' -Severity 'Low' `
            -Details "CIM App Attach packages in subscription: $($AppAttachPkgs.Count)" `
            -Recommendation 'CIM-based App Attach is GA - continue using it for admin-less app updates and per-group targeting.' `
            -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/app-attach-overview' `
            -Evidence @{ PackageCount = $AppAttachPkgs.Count }))
    }
    # No packages: N/A is emitted at estate level (APP-ATTACH-NONE) only if nothing found anywhere - handled post-loop.
} catch {
    Write-Status "  App Attach package error: $($_.Exception.Message)" -Level 'WARN'
}

# ─── DEFENDER FOR SERVERS / TVM (SEC-007) ──────────────────────────────
# Threat & Vulnerability Management ships with Defender for Servers. Plan P2 (subPlan) includes
# the integrated vulnerability assessment (MDVM); P1 is endpoint protection only.
Write-Status "Defender for Servers (TVM)" -Level 'SECTION'
try {
    $ServersPricing = Get-AzSecurityPricing -Name 'VirtualMachines' -ErrorAction Stop
    $Tier    = "$($ServersPricing.PricingTier)"
    $SubPlan = "$($ServersPricing.SubPlan)"
    if ($Tier -ne 'Standard') {
        $TvmStatus = 'Fail'; $TvmDetail = "Defender for Servers is off (PricingTier: $Tier) - no threat & vulnerability management."
    } elseif ($SubPlan -eq 'P2') {
        $TvmStatus = 'Pass'; $TvmDetail = "Defender for Servers Plan 2 (P2) enabled - integrated vulnerability management (MDVM) active."
    } elseif ($SubPlan -eq 'P1') {
        $TvmStatus = 'Warning'; $TvmDetail = "Defender for Servers Plan 1 (P1) enabled - endpoint protection only; P2 adds integrated vulnerability assessment."
    } else {
        $TvmStatus = 'Warning'; $TvmDetail = "Defender for Servers enabled (PricingTier: $Tier, SubPlan: $(if ($SubPlan) { $SubPlan } else { 'unspecified' })). Confirm Plan 2 for vulnerability management."
    }
    [void]$AllChecks.Add((New-CheckResult -Id "SEC-TVM-$SubShort" `
        -Category 'Security' -Name 'TVM Assessments Enabled' `
        -Description 'Threat & Vulnerability Management via Defender for Servers Plan 2 should be enabled for session hosts' `
        -Status $TvmStatus -Severity 'High' `
        -Details $TvmDetail `
        -Recommendation 'Enable Microsoft Defender for Servers Plan 2 to get integrated vulnerability management (MDVM) for AVD session hosts.' `
        -Reference 'https://learn.microsoft.com/en-us/azure/defender-for-cloud/plan-defender-for-servers-select-plan' `
        -Evidence @{ PricingTier = $Tier; SubPlan = $SubPlan }))
} catch {
    [void]$AllChecks.Add((New-CheckResult -Id "SEC-TVM-$SubShort" `
        -Category 'Security' -Name 'TVM Assessments Enabled' `
        -Description 'Threat & Vulnerability Management via Defender for Servers Plan 2 should be enabled for session hosts' `
        -Status 'Error' -Severity 'High' `
        -Details "Could not read Defender for Servers pricing (Az.Security missing or access denied): $($_.Exception.Message)" `
        -Reference 'https://learn.microsoft.com/en-us/azure/defender-for-cloud/plan-defender-for-servers-select-plan'))
}
}  # end per-subscription sweep (A-1)

# ═══════════════════════════════════════════════════════════════════════════
# MICROSOFT GRAPH — IDENTITY (Conditional Access, MFA, token protection, passwordless)
# Tenant-level (singleton) checks. Degrades to Status 'Error' when Graph permissions are missing.
# ═══════════════════════════════════════════════════════════════════════════
Write-Status "Microsoft Graph (Identity)" -Level 'SECTION'
$AvdAppIds = [ordered]@{
    'Azure Virtual Desktop'    = '9cdead84-a844-4324-93f2-b2e6bb768d07'
    'Windows Cloud Login'      = '270efc09-cd0d-444b-a71f-39af4910ec45'
    'Microsoft Remote Desktop' = 'a4a365df-50f1-4397-bc59-1a1564b8bb9c'
}
$AvdAppIdList = @($AvdAppIds.Values)
$WclAppId     = $AvdAppIds['Windows Cloud Login']
$MrdAppId     = $AvdAppIds['Microsoft Remote Desktop']

$GraphToken = Get-GraphTokenString
$GraphPermMsg = 'insufficient Graph permissions — grant Policy.Read.All / AuditLog.Read.All for identity checks'

# --- Conditional Access derived checks: IAM-002 (MFA), IAM-003 (CA), IAM-011 (WCL), IAM-010 (token protection) ---
$CaError = $null
$CaPolicies = @()
if (-not $GraphToken) {
    $CaError = 'Could not acquire a Microsoft Graph token from the current Az login.'
} else {
    try {
        $CaPolicies = @(Invoke-GraphGet -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies' -Token $GraphToken)
    } catch {
        $CaError = "$GraphPermMsg ($($_.Exception.Message))"
    }
}

if ($CaError) {
    foreach ($EId in @('IAM-MFA','IAM-CA','IAM-WCL','IAM-TOKPROT')) {
        [void]$AllChecks.Add((New-CheckResult -Id $EId `
            -Category 'Identity & Access' -Name 'Conditional Access (Graph)' `
            -Description 'Identity posture evaluated from Entra Conditional Access policies via Microsoft Graph' `
            -Status 'Error' -Severity 'High' `
            -Details $CaError `
            -Reference 'https://learn.microsoft.com/en-us/entra/identity/conditional-access/overview'))
    }
    Write-Status "  Conditional Access: $CaError" -Level 'WARN'
} else {
    # Classify each policy against the AVD app IDs.
    $EnabledMfa = @(); $ReportOnlyMfa = @()
    $EnabledAvdCa = @(); $EnabledMrdOnlyCa = @()
    $WclCovered = $false; $AvdCovered = $false
    $TokProtEnabled = @()
    foreach ($P in $CaPolicies) {
        $State = "$($P.state)"
        $IncApps = @($P.conditions.applications.includeApplications)
        $TargetsAll = ('All' -in $IncApps)
        $TargetsAvdApp = @($IncApps | Where-Object { $_ -in $AvdAppIdList }).Count -gt 0
        $TargetsWcl = ($WclAppId -in $IncApps)
        $TargetsMrdOnly = ($MrdAppId -in $IncApps) -and -not (@($IncApps | Where-Object { $_ -in @($AvdAppIds['Azure Virtual Desktop'], $WclAppId) }).Count -gt 0)
        $Grant = @($P.grantControls.builtInControls)
        $HasMfa = ('mfa' -in $Grant)
        $CoversAvd = $TargetsAll -or $TargetsAvdApp
        $CoversWcl = $TargetsAll -or $TargetsWcl

        if ($CoversAvd -and $HasMfa) {
            if ($State -eq 'enabled') { $EnabledMfa += $P.displayName }
            elseif ($State -eq 'enabledForReportingButNotEnforced') { $ReportOnlyMfa += $P.displayName }
        }
        if ($State -eq 'enabled' -and ($TargetsAll -or $TargetsAvdApp)) {
            if ($TargetsMrdOnly) { $EnabledMrdOnlyCa += $P.displayName } else { $EnabledAvdCa += $P.displayName }
        }
        if ($State -eq 'enabled' -and $CoversAvd) { $AvdCovered = $true }
        if ($State -eq 'enabled' -and $CoversWcl -and $HasMfa) { $WclCovered = $true }
        # Token protection: sessionControls.secureSignInSession / signInSessionControls
        $SessCtl = $P.sessionControls
        $HasTokProt = $false
        if ($SessCtl) {
            if ($SessCtl.secureSignInSession -and $SessCtl.secureSignInSession.isEnabled) { $HasTokProt = $true }
            if ($SessCtl.signInSessionControls) { $HasTokProt = $true }
        }
        if ($State -eq 'enabled' -and $HasTokProt) { $TokProtEnabled += $P.displayName }
    }

    # IAM-002 MFA
    if ($EnabledMfa.Count -gt 0) {
        $MfaStatus = 'Pass'; $MfaDetail = "Enabled MFA policies targeting AVD apps: $($EnabledMfa -join ', ')"
    } elseif ($ReportOnlyMfa.Count -gt 0) {
        $MfaStatus = 'Warning'; $MfaDetail = "MFA policies for AVD apps exist but are report-only: $($ReportOnlyMfa -join ', ')"
    } else {
        $MfaStatus = 'Fail'; $MfaDetail = 'No enabled Conditional Access policy enforces MFA for the AVD or Windows Cloud Login apps (or All apps).'
    }
    [void]$AllChecks.Add((New-CheckResult -Id 'IAM-MFA' `
        -Category 'Identity & Access' -Name 'MFA Enforcement' `
        -Description 'MFA must be enforced for AVD users via Conditional Access targeting the AVD / Windows Cloud Login apps' `
        -Status $MfaStatus -Severity 'Critical' `
        -Details $MfaDetail `
        -Recommendation 'Enforce MFA via a Conditional Access policy targeting Azure Virtual Desktop and Windows Cloud Login (grant control: require MFA).' `
        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/set-up-mfa' `
        -Evidence @{ EnabledMfaPolicies = $EnabledMfa; ReportOnly = $ReportOnlyMfa }))

    # IAM-003 CA
    if ($EnabledAvdCa.Count -gt 0) {
        $CaStatus = 'Pass'; $CaDetail = "Enabled CA policies targeting AVD apps: $($EnabledAvdCa -join ', ')"
    } elseif ($EnabledMrdOnlyCa.Count -gt 0) {
        $CaStatus = 'Warning'; $CaDetail = "CA policies target only the legacy Microsoft Remote Desktop app: $($EnabledMrdOnlyCa -join ', '). Retarget to Azure Virtual Desktop + Windows Cloud Login."
    } else {
        $CaStatus = 'Warning'; $CaDetail = 'No enabled Conditional Access policy targets the AVD cloud apps.'
    }
    [void]$AllChecks.Add((New-CheckResult -Id 'IAM-CA' `
        -Category 'Identity & Access' -Name 'Conditional Access Policies' `
        -Description 'Conditional Access should target the Azure Virtual Desktop and Windows Cloud Login cloud apps' `
        -Status $CaStatus -Severity 'High' `
        -Details $CaDetail `
        -Recommendation 'Target Conditional Access at Azure Virtual Desktop and Windows Cloud Login (Microsoft Remote Desktop is legacy). Consider the "Every time" sign-in frequency for high-sensitivity access.' `
        -Reference 'https://learn.microsoft.com/en-us/entra/identity/conditional-access/overview' `
        -Evidence @{ AvdPolicies = $EnabledAvdCa; LegacyOnlyPolicies = $EnabledMrdOnlyCa }))

    # IAM-011 Windows Cloud Login coverage
    if ($WclCovered) {
        $WclStatus = 'Pass'; $WclDetail = 'Windows Cloud Login app is covered by an enabled MFA/CA policy (explicitly or via All apps).'
    } elseif ($AvdCovered) {
        $WclStatus = 'Fail'; $WclDetail = 'Azure Virtual Desktop app is covered but Windows Cloud Login is NOT — SSO and current sign-in flows depend on Windows Cloud Login coverage.'
    } else {
        $WclStatus = 'Fail'; $WclDetail = 'Windows Cloud Login app is not covered by any enabled MFA/CA policy.'
    }
    [void]$AllChecks.Add((New-CheckResult -Id 'IAM-WCL' `
        -Category 'Identity & Access' -Name 'Windows Cloud Login CA Coverage' `
        -Description 'Conditional Access / MFA policies must cover the Windows Cloud Login app (the current AVD sign-in app)' `
        -Status $WclStatus -Severity 'High' `
        -Details $WclDetail `
        -Recommendation 'Add the Windows Cloud Login app (270efc09-cd0d-444b-a71f-39af4910ec45) to your AVD Conditional Access / MFA policies, or target All cloud apps.' `
        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/set-up-mfa' `
        -Evidence @{ WclCovered = $WclCovered; AvdCovered = $AvdCovered }))

    # IAM-010 Token protection
    [void]$AllChecks.Add((New-CheckResult -Id 'IAM-TOKPROT' `
        -Category 'Identity & Access' -Name 'Token Protection Policy' `
        -Description 'Conditional Access token protection (token binding) should bind sign-in sessions to the originating device' `
        -Status $(if ($TokProtEnabled.Count -gt 0) { 'Pass' } else { 'Warning' }) `
        -Severity 'Medium' `
        -Details "$(if ($TokProtEnabled.Count -gt 0) { "Enabled token-protection policies: $($TokProtEnabled -join ', ')" } else { 'No enabled Conditional Access policy configures token protection (secure sign-in session).' })" `
        -Recommendation 'Enable Conditional Access token protection (sign-in session token binding) for AVD sign-ins to prevent token replay.' `
        -Reference 'https://learn.microsoft.com/en-us/entra/identity/conditional-access/concept-token-protection' `
        -Evidence @{ TokenProtectionPolicies = $TokProtEnabled }))

    Write-Status "  CA policies: $($CaPolicies.Count), MFA=$MfaStatus, CA=$CaStatus, WCL=$WclStatus, TokProt=$(if ($TokProtEnabled.Count -gt 0) { 'Pass' } else { 'Warning' })" -Level 'SUCCESS'
}

# --- IAM-009 passwordless registration (authenticationMethods report; needs AuditLog.Read.All + premium) ---
if (-not $GraphToken) {
    [void]$AllChecks.Add((New-CheckResult -Id 'IAM-PWLESS' `
        -Category 'Identity & Access' -Name 'Passwordless Authentication' `
        -Description 'Users should be registered for passwordless methods (Windows Hello for Business, FIDO2, passkeys)' `
        -Status 'Error' -Severity 'Low' `
        -Details 'Could not acquire a Microsoft Graph token from the current Az login.' `
        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/authentication#in-session-passwordless-authentication'))
} else {
    try {
        $RegDetails = @(Invoke-GraphGet -Uri 'https://graph.microsoft.com/v1.0/reports/authenticationMethods/userRegistrationDetails?$top=999' -Token $GraphToken)
        $TotalUsers = $RegDetails.Count
        $Passwordless = @($RegDetails | Where-Object {
            @($_.methodsRegistered | Where-Object { $_ -match 'windowsHelloForBusiness|fido2|passKey' }).Count -gt 0
        })
        $PwlessPct = if ($TotalUsers -gt 0) { [math]::Round($Passwordless.Count / $TotalUsers * 100, 0) } else { 0 }
        $PwlessStatus = if ($PwlessPct -ge 50) { 'Pass' } elseif ($PwlessPct -gt 0) { 'Warning' } else { 'Fail' }
        Write-Status "  Passwordless registration: $PwlessPct% of $TotalUsers users" -Level $(if ($PwlessPct -ge 50) { 'SUCCESS' } else { 'WARN' })
        [void]$AllChecks.Add((New-CheckResult -Id 'IAM-PWLESS' `
            -Category 'Identity & Access' -Name 'Passwordless Authentication' `
            -Description 'Users should be registered for passwordless methods (Windows Hello for Business, FIDO2, passkeys)' `
            -Status $PwlessStatus -Severity 'Low' `
            -Details "$PwlessPct% of $TotalUsers users registered for a passwordless method (WHfB / FIDO2 / passkey)" `
            -Recommendation 'Drive passwordless adoption (Windows Hello for Business, FIDO2 keys, passkeys) to eliminate password-based attack vectors in-session.' `
            -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/authentication#in-session-passwordless-authentication' `
            -Evidence @{ TotalUsers = $TotalUsers; PasswordlessUsers = $Passwordless.Count; Percent = $PwlessPct }))
    } catch {
        [void]$AllChecks.Add((New-CheckResult -Id 'IAM-PWLESS' `
            -Category 'Identity & Access' -Name 'Passwordless Authentication' `
            -Description 'Users should be registered for passwordless methods (Windows Hello for Business, FIDO2, passkeys)' `
            -Status 'Error' -Severity 'Low' `
            -Details "$GraphPermMsg ($($_.Exception.Message))" `
            -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/authentication#in-session-passwordless-authentication'))
    }
}

# ═══════════════════════════════════════════════════════════════════════════
# MICROSOFT GRAPH — INTUNE / ENDPOINT MANAGER (tenant-level singleton checks)
# SH-028 baselines, SH-014 compliance/drift, SH-005 patch, SEC-001 app control,
# SEC-003 Credential Guard, SEC-004 VBS/HVCI, PROF-007 OneDrive KFM, PROF-019 AV exclusions.
# Needs DeviceManagementConfiguration.Read.All / DeviceManagementManagedDevices.Read.All.
# Each check degrades to Status 'Error' (never crashes) when its scope/token is missing.
# ═══════════════════════════════════════════════════════════════════════════
Write-Status "Microsoft Graph (Intune)" -Level 'SECTION'
$IntuneScopeMsg = 'insufficient Graph permissions - grant DeviceManagementConfiguration.Read.All / DeviceManagementManagedDevices.Read.All for Intune checks'

# Helper: emit an Error result for an Intune singleton check.
function Add-IntuneError {
    param([string]$Id, [string]$Name, [string]$Desc, [string]$Msg, [string]$Ref, [string]$Sev = 'Medium')
    [void]$AllChecks.Add((New-CheckResult -Id $Id -Category $(if ($Id -like 'SEC-*') { 'Security' } elseif ($Id -like 'PROF-*') { 'FSLogix & Profiles' } else { 'Session Hosts' }) `
        -Name $Name -Description $Desc -Status 'Error' -Severity $Sev -Details $Msg -Reference $Ref))
}

if (-not $GraphToken) {
    $NoTok = 'Could not acquire a Microsoft Graph token from the current Az login.'
    Add-IntuneError 'SH-BASELINE'   'OS Security Baselines'   'Intune security baselines should be assigned to session hosts' $NoTok 'https://learn.microsoft.com/en-us/mem/intune/protect/security-baselines' 'High'
    Add-IntuneError 'SH-DRIFT'      'Configuration Drift Detection' 'Intune device compliance policies detect configuration drift' $NoTok 'https://learn.microsoft.com/en-us/mem/intune/protect/device-compliance-get-started' 'Medium'
    Add-IntuneError 'SH-PATCH'      'Patch Management Strategy' 'Session host update compliance should be tracked in Intune' $NoTok 'https://learn.microsoft.com/en-us/mem/intune/protect/windows-update-for-business-configure' 'High'
    Add-IntuneError 'SEC-APPCTRL'   'Application Control (WDAC/AppLocker)' 'AppLocker/WDAC application control policies should be deployed' $NoTok 'https://learn.microsoft.com/en-us/windows/security/application-security/application-control/windows-defender-application-control/wdac' 'High'
    Add-IntuneError 'SEC-CREDGUARD' 'Credential Guard'        'Credential Guard should be enabled via device configuration' $NoTok 'https://learn.microsoft.com/en-us/windows/security/identity-protection/credential-guard/' 'High'
    Add-IntuneError 'SEC-VBS'       'VBS/HVCI Enabled'        'Virtualization-based security / HVCI should be enabled' $NoTok 'https://learn.microsoft.com/en-us/windows-hardware/design/device-experiences/oem-vbs' 'High'
    Add-IntuneError 'PROF-KFM'      'OneDrive KFM Enabled'    'OneDrive Known Folder Move should be silently enabled' $NoTok 'https://learn.microsoft.com/en-us/sharepoint/redirect-known-folders' 'Medium'
    Add-IntuneError 'PROF-AVEXCL'   'AV Exclusions Complete (Full List)' 'Antivirus policies should exclude FSLogix container paths' $NoTok 'https://learn.microsoft.com/en-us/fslogix/overview-prerequisites#configure-antivirus-exclusions' 'Medium'
    Write-Status "  Intune: $NoTok" -Level 'WARN'
} else {
    $GBeta = 'https://graph.microsoft.com/beta'
    $GV1   = 'https://graph.microsoft.com/v1.0'

    # --- SH-028: OS Security Baselines (deviceManagement/intents) ---
    try {
        $Intents = @(Invoke-GraphGet -Uri "$GBeta/deviceManagement/intents" -Token $GraphToken)
        $Baselines = @($Intents | Where-Object { "$($_.displayName)$($_.templateId)" -match '(?i)baseline|mdm security|defender for endpoint|windows security' })
        if ($Baselines.Count -eq 0) { $Baselines = $Intents }  # any intent is a managed configuration profile
        if ($Baselines.Count -gt 0) {
            $BlNames = @($Baselines | ForEach-Object { $_.displayName } | Where-Object { $_ } | Select-Object -First 8)
            [void]$AllChecks.Add((New-CheckResult -Id 'SH-BASELINE' -Category 'Session Hosts' -Name 'OS Security Baselines' `
                -Description 'Intune security baselines (or configuration intents) should be assigned to AVD session hosts' `
                -Status 'Pass' -Severity 'High' `
                -Details "Intune security baseline/intent profiles found: $($Baselines.Count) ($($BlNames -join ', ')). Verify each is assigned to the AVD device group." `
                -Recommendation 'Assign a Windows security baseline (or hardening configuration profiles) to the AVD session-host device group in Intune.' `
                -Reference 'https://learn.microsoft.com/en-us/mem/intune/protect/security-baselines' `
                -Evidence @{ BaselineCount = $Baselines.Count; Names = $BlNames }))
        } else {
            [void]$AllChecks.Add((New-CheckResult -Id 'SH-BASELINE' -Category 'Session Hosts' -Name 'OS Security Baselines' `
                -Description 'Intune security baselines (or configuration intents) should be assigned to AVD session hosts' `
                -Status 'Warning' -Severity 'High' `
                -Details 'No Intune security baselines or configuration intents found in the tenant.' `
                -Recommendation 'Deploy a Windows security baseline to the AVD session-host device group in Intune.' `
                -Reference 'https://learn.microsoft.com/en-us/mem/intune/protect/security-baselines'))
        }
    } catch {
        Add-IntuneError 'SH-BASELINE' 'OS Security Baselines' 'Intune security baselines should be assigned to session hosts' "$IntuneScopeMsg ($($_.Exception.Message))" 'https://learn.microsoft.com/en-us/mem/intune/protect/security-baselines' 'High'
    }

    # --- SH-014: Configuration Drift (deviceCompliancePolicies + assignments) ---
    try {
        $CompPolicies = @(Invoke-GraphGet -Uri "$GV1/deviceManagement/deviceCompliancePolicies?`$expand=assignments" -Token $GraphToken)
        $Assigned = @($CompPolicies | Where-Object { @($_.assignments).Count -gt 0 })
        if ($Assigned.Count -gt 0) {
            [void]$AllChecks.Add((New-CheckResult -Id 'SH-DRIFT' -Category 'Session Hosts' -Name 'Configuration Drift Detection' `
                -Description 'Intune device compliance policies should be assigned to detect configuration drift' `
                -Status 'Pass' -Severity 'Medium' `
                -Details "Assigned device compliance policies: $($Assigned.Count) of $($CompPolicies.Count) total." `
                -Recommendation 'Keep device compliance policies assigned to the AVD device group and pair with Conditional Access for drift enforcement.' `
                -Reference 'https://learn.microsoft.com/en-us/mem/intune/protect/device-compliance-get-started' `
                -Evidence @{ Total = $CompPolicies.Count; Assigned = $Assigned.Count }))
        } elseif ($CompPolicies.Count -gt 0) {
            [void]$AllChecks.Add((New-CheckResult -Id 'SH-DRIFT' -Category 'Session Hosts' -Name 'Configuration Drift Detection' `
                -Description 'Intune device compliance policies should be assigned to detect configuration drift' `
                -Status 'Warning' -Severity 'Medium' `
                -Details "$($CompPolicies.Count) device compliance policy(ies) exist but none are assigned." `
                -Recommendation 'Assign the device compliance policies to the AVD session-host device group.' `
                -Reference 'https://learn.microsoft.com/en-us/mem/intune/protect/device-compliance-get-started'))
        } else {
            [void]$AllChecks.Add((New-CheckResult -Id 'SH-DRIFT' -Category 'Session Hosts' -Name 'Configuration Drift Detection' `
                -Description 'Intune device compliance policies should be assigned to detect configuration drift' `
                -Status 'Warning' -Severity 'Medium' `
                -Details 'No Intune device compliance policies found in the tenant.' `
                -Recommendation 'Create and assign device compliance policies for the AVD session hosts.' `
                -Reference 'https://learn.microsoft.com/en-us/mem/intune/protect/device-compliance-get-started'))
        }
    } catch {
        Add-IntuneError 'SH-DRIFT' 'Configuration Drift Detection' 'Intune device compliance policies detect configuration drift' "$IntuneScopeMsg ($($_.Exception.Message))" 'https://learn.microsoft.com/en-us/mem/intune/protect/device-compliance-get-started' 'Medium'
    }

    # --- SH-005: Patch Management (softwareUpdateStatusSummary) ---
    try {
        $UpdSummary = @(Invoke-GraphGet -Uri "$GBeta/deviceManagement/softwareUpdateStatusSummary" -Token $GraphToken)
        $Sum = if ($UpdSummary.Count -gt 0) { $UpdSummary[0] } else { $null }
        if ($Sum) {
            $Compliant    = [int]$Sum.compliantDeviceCount
            $NonCompliant = [int]$Sum.nonCompliantDeviceCount
            $ErrorDev     = [int]$Sum.errorDeviceCount
            $Unknown      = [int]$Sum.unknownDeviceCount
            $ConflictDev  = [int]$Sum.conflictDeviceCount
            $TotalDev = $Compliant + $NonCompliant + $ErrorDev + $Unknown + $ConflictDev
            $Pct = if ($TotalDev -gt 0) { [math]::Round($Compliant / $TotalDev * 100, 0) } else { 0 }
            $PatchStatus = if ($TotalDev -eq 0) { 'Warning' } elseif ($Pct -ge 90) { 'Pass' } elseif ($Pct -ge 70) { 'Warning' } else { 'Fail' }
            [void]$AllChecks.Add((New-CheckResult -Id 'SH-PATCH' -Category 'Session Hosts' -Name 'Patch Management Strategy' `
                -Description 'Session hosts should report high software-update compliance in Intune' `
                -Status $PatchStatus -Severity 'High' `
                -Details "$(if ($TotalDev -eq 0) { 'No update-status devices reported by Intune yet.' } else { "$Pct% update-compliant ($Compliant of $TotalDev devices; $NonCompliant non-compliant, $ErrorDev error, $Unknown unknown)." })" `
                -Recommendation 'Drive update compliance to >=90% via Windows Update for Business / Update rings targeting the AVD device group.' `
                -Reference 'https://learn.microsoft.com/en-us/mem/intune/protect/windows-update-for-business-configure' `
                -Evidence @{ CompliantPct = $Pct; Total = $TotalDev; Compliant = $Compliant; NonCompliant = $NonCompliant }))
        } else {
            [void]$AllChecks.Add((New-CheckResult -Id 'SH-PATCH' -Category 'Session Hosts' -Name 'Patch Management Strategy' `
                -Description 'Session hosts should report high software-update compliance in Intune' `
                -Status 'Warning' -Severity 'High' `
                -Details 'Intune returned no software update status summary.' `
                -Recommendation 'Configure Windows Update for Business update rings for the AVD device group.' `
                -Reference 'https://learn.microsoft.com/en-us/mem/intune/protect/windows-update-for-business-configure'))
        }
    } catch {
        Add-IntuneError 'SH-PATCH' 'Patch Management Strategy' 'Session host update compliance should be tracked in Intune' "$IntuneScopeMsg ($($_.Exception.Message))" 'https://learn.microsoft.com/en-us/mem/intune/protect/windows-update-for-business-configure' 'High'
    }

    # --- Shared fetch: device configurations + settings-catalog policies (for SEC-001/003/004, PROF-007/019) ---
    $CfgFetchError = $null
    $DeviceConfigs = @(); $ConfigPolicies = @(); $GpConfigs = @()
    try {
        $DeviceConfigs = @(Invoke-GraphGet -Uri "$GBeta/deviceManagement/deviceConfigurations" -Token $GraphToken)
    } catch { $CfgFetchError = "$IntuneScopeMsg ($($_.Exception.Message))" }
    if (-not $CfgFetchError) {
        try { $ConfigPolicies = @(Invoke-GraphGet -Uri "$GBeta/deviceManagement/configurationPolicies" -Token $GraphToken) } catch { }
        try { $GpConfigs      = @(Invoke-GraphGet -Uri "$GBeta/deviceManagement/groupPolicyConfigurations" -Token $GraphToken) } catch { }
    }
    # Flatten a searchable text blob per profile (displayName + omaSettings + serialized settings).
    $CfgBlobs = @()
    if (-not $CfgFetchError) {
        foreach ($DC in $DeviceConfigs) {
            $Blob = "$($DC.'@odata.type') $($DC.displayName)"
            if ($DC.omaSettings) { $Blob += ' ' + (@($DC.omaSettings | ForEach-Object { "$($_.displayName) $($_.omaUri)" }) -join ' ') }
            try { $Blob += ' ' + ($DC | ConvertTo-Json -Depth 4 -Compress) } catch { }
            $CfgBlobs += $Blob
        }
        foreach ($CP in $ConfigPolicies) { $CfgBlobs += "$($CP.name) $($CP.description)" }
        foreach ($GP in $GpConfigs)      { $CfgBlobs += "$($GP.displayName)" }
    }

    # --- SEC-001: Application Control (AppLocker / WDAC) ---
    if ($CfgFetchError) {
        Add-IntuneError 'SEC-APPCTRL' 'Application Control (WDAC/AppLocker)' 'AppLocker/WDAC application control policies should be deployed' $CfgFetchError 'https://learn.microsoft.com/en-us/windows/security/application-security/application-control/windows-defender-application-control/wdac' 'High'
    } else {
        $AppCtl = @($CfgBlobs | Where-Object { $_ -match '(?i)applocker|applicationcontrol|\bwdac\b|application control|windowsDefenderApplicationControl' })
        [void]$AllChecks.Add((New-CheckResult -Id 'SEC-APPCTRL' -Category 'Security' -Name 'Application Control (WDAC/AppLocker)' `
            -Description 'AppLocker or WDAC application control should be deployed to AVD session hosts via Intune' `
            -Status $(if ($AppCtl.Count -gt 0) { 'Pass' } else { 'Warning' }) -Severity 'High' `
            -Details "$(if ($AppCtl.Count -gt 0) { "$($AppCtl.Count) profile(s) reference AppLocker/WDAC application control." } else { 'No AppLocker or WDAC application-control profiles found in Intune.' })" `
            -Recommendation 'Deploy WDAC (App Control for Business) or AppLocker policies to restrict which applications run on session hosts.' `
            -Reference 'https://learn.microsoft.com/en-us/windows/security/application-security/application-control/windows-defender-application-control/wdac' `
            -Evidence @{ MatchingProfiles = $AppCtl.Count }))
    }

    # --- SEC-003: Credential Guard + SEC-004: VBS/HVCI (deviceGuard settings) ---
    if ($CfgFetchError) {
        Add-IntuneError 'SEC-CREDGUARD' 'Credential Guard' 'Credential Guard should be enabled via device configuration' $CfgFetchError 'https://learn.microsoft.com/en-us/windows/security/identity-protection/credential-guard/' 'High'
        Add-IntuneError 'SEC-VBS' 'VBS/HVCI Enabled' 'Virtualization-based security / HVCI should be enabled' $CfgFetchError 'https://learn.microsoft.com/en-us/windows-hardware/design/device-experiences/oem-vbs' 'High'
    } else {
        $CredGuard = @($CfgBlobs | Where-Object { $_ -match '(?i)deviceGuardLocalSystemAuthorityCredentialGuard|credentialGuard|credential guard|lsaCfgFlags' })
        [void]$AllChecks.Add((New-CheckResult -Id 'SEC-CREDGUARD' -Category 'Security' -Name 'Credential Guard' `
            -Description 'Windows Defender Credential Guard should be enabled on session hosts via Intune device configuration' `
            -Status $(if ($CredGuard.Count -gt 0) { 'Pass' } else { 'Warning' }) -Severity 'High' `
            -Details "$(if ($CredGuard.Count -gt 0) { "$($CredGuard.Count) profile(s) configure Credential Guard." } else { 'No device configuration profile enabling Credential Guard was found.' })" `
            -Recommendation 'Enable Credential Guard via an Endpoint Protection / device-guard configuration profile assigned to session hosts (note: not supported on all pooled/GPU SKUs).' `
            -Reference 'https://learn.microsoft.com/en-us/windows/security/identity-protection/credential-guard/' `
            -Evidence @{ MatchingProfiles = $CredGuard.Count }))
        $Vbs = @($CfgBlobs | Where-Object { $_ -match '(?i)virtualizationBasedSecurity|virtualization based security|\bHVCI\b|hypervisorEnforcedCodeIntegrity|deviceGuardEnableVirtualizationBasedSecurity' })
        [void]$AllChecks.Add((New-CheckResult -Id 'SEC-VBS' -Category 'Security' -Name 'VBS/HVCI Enabled' `
            -Description 'Virtualization-based security (VBS) and HVCI should be enabled on session hosts via Intune' `
            -Status $(if ($Vbs.Count -gt 0) { 'Pass' } else { 'Warning' }) -Severity 'High' `
            -Details "$(if ($Vbs.Count -gt 0) { "$($Vbs.Count) profile(s) configure VBS/HVCI." } else { 'No device configuration profile enabling VBS/HVCI was found.' })" `
            -Recommendation 'Enable VBS with HVCI (memory integrity) via a device-guard configuration profile (verify nested-virtualization-capable SKUs).' `
            -Reference 'https://learn.microsoft.com/en-us/windows-hardware/design/device-experiences/oem-vbs' `
            -Evidence @{ MatchingProfiles = $Vbs.Count }))
    }

    # --- PROF-007: OneDrive KFM (KFMSilentOptIn) ---
    if ($CfgFetchError) {
        Add-IntuneError 'PROF-KFM' 'OneDrive KFM Enabled' 'OneDrive Known Folder Move should be silently enabled' $CfgFetchError 'https://learn.microsoft.com/en-us/sharepoint/redirect-known-folders' 'Medium'
    } else {
        $Kfm = @($CfgBlobs | Where-Object { $_ -match '(?i)KFMSilentOptIn|KnownFolderMove|known folder move' })
        [void]$AllChecks.Add((New-CheckResult -Id 'PROF-KFM' -Category 'FSLogix & Profiles' -Name 'OneDrive KFM Enabled' `
            -Description 'OneDrive Known Folder Move (silent opt-in) should redirect Desktop/Documents/Pictures for AVD users' `
            -Status $(if ($Kfm.Count -gt 0) { 'Pass' } else { 'Warning' }) -Severity 'Medium' `
            -Details "$(if ($Kfm.Count -gt 0) { "$($Kfm.Count) profile(s) configure OneDrive Known Folder Move." } else { 'No policy configuring OneDrive KFM (KFMSilentOptIn) was found.' })" `
            -Recommendation 'Configure OneDrive KFMSilentOptIn via an administrative-template / settings-catalog profile so known folders roam without bloating FSLogix profiles.' `
            -Reference 'https://learn.microsoft.com/en-us/sharepoint/redirect-known-folders' `
            -Evidence @{ MatchingProfiles = $Kfm.Count }))
    }

    # --- PROF-019: AV Exclusions for FSLogix (antivirus policies w/ FSLogix paths) ---
    if ($CfgFetchError) {
        Add-IntuneError 'PROF-AVEXCL' 'AV Exclusions Complete (Full List)' 'Antivirus policies should exclude FSLogix container paths' $CfgFetchError 'https://learn.microsoft.com/en-us/fslogix/overview-prerequisites#configure-antivirus-exclusions' 'Medium'
    } else {
        $AvProfiles  = @($CfgBlobs | Where-Object { $_ -match '(?i)antivirus|defender|windowsDefender|exclusion' })
        $AvWithFsl   = @($CfgBlobs | Where-Object { $_ -match '(?i)exclusion' -and $_ -match '(?i)fslogix|\.vhdx?|profiles|\.VHD' })
        if ($AvWithFsl.Count -gt 0) {
            [void]$AllChecks.Add((New-CheckResult -Id 'PROF-AVEXCL' -Category 'FSLogix & Profiles' -Name 'AV Exclusions Complete (Full List)' `
                -Description 'Antivirus policies should exclude FSLogix container paths (*.vhd(x), Profiles, ODFC)' `
                -Status 'Pass' -Severity 'Medium' `
                -Details "$($AvWithFsl.Count) antivirus profile(s) include FSLogix path exclusions." `
                -Recommendation 'Keep the full FSLogix antivirus exclusion list current per Microsoft guidance.' `
                -Reference 'https://learn.microsoft.com/en-us/fslogix/overview-prerequisites#configure-antivirus-exclusions' `
                -Evidence @{ AvProfilesWithFslExclusions = $AvWithFsl.Count }))
        } elseif ($AvProfiles.Count -gt 0) {
            [void]$AllChecks.Add((New-CheckResult -Id 'PROF-AVEXCL' -Category 'FSLogix & Profiles' -Name 'AV Exclusions Complete (Full List)' `
                -Description 'Antivirus policies should exclude FSLogix container paths (*.vhd(x), Profiles, ODFC)' `
                -Status 'Warning' -Severity 'Medium' `
                -Details "$($AvProfiles.Count) antivirus/Defender profile(s) found but none reference FSLogix path exclusions." `
                -Recommendation 'Add the FSLogix container exclusions (*.vhd, *.vhdx, Profiles/ODFC paths) to the antivirus policy.' `
                -Reference 'https://learn.microsoft.com/en-us/fslogix/overview-prerequisites#configure-antivirus-exclusions' `
                -Evidence @{ AvProfiles = $AvProfiles.Count }))
        } else {
            [void]$AllChecks.Add((New-CheckResult -Id 'PROF-AVEXCL' -Category 'FSLogix & Profiles' -Name 'AV Exclusions Complete (Full List)' `
                -Description 'Antivirus policies should exclude FSLogix container paths (*.vhd(x), Profiles, ODFC)' `
                -Status 'Warning' -Severity 'Medium' `
                -Details 'No antivirus/Defender configuration profiles found in Intune to verify FSLogix exclusions.' `
                -Recommendation 'Deploy a Microsoft Defender Antivirus policy that includes the FSLogix container exclusions.' `
                -Reference 'https://learn.microsoft.com/en-us/fslogix/overview-prerequisites#configure-antivirus-exclusions'))
        }
    }
    Write-Status "  Intune: baselines/compliance/patch/appcontrol/credguard/vbs/kfm/av-exclusions assessed" -Level 'SUCCESS'
}

# ═══════════════════════════════════════════════════════════════════════════
# SIEM / SENTINEL ONBOARDING (MON-012) — per Log Analytics workspace from diag settings
# ═══════════════════════════════════════════════════════════════════════════
Write-Status "SIEM / Sentinel" -Level 'SECTION'
$WsIds = @($LAWorkspaceIds.Keys)
if ($WsIds.Count -eq 0) {
    [void]$AllChecks.Add((New-CheckResult -Id 'MON-SIEM-NONE' `
        -Category 'Monitoring' -Name 'SIEM Integration' `
        -Description 'Microsoft Sentinel (or equivalent SIEM) should be connected to the AVD Log Analytics workspace' `
        -Status 'Warning' -Severity 'Medium' `
        -Details 'No Log Analytics workspace found via host-pool diagnostic settings — cannot assess Sentinel onboarding.' `
        -Recommendation 'Send AVD diagnostics to a Log Analytics workspace and onboard Microsoft Sentinel for security correlation.' `
        -Reference 'https://learn.microsoft.com/en-us/azure/sentinel/overview'))
} else {
    foreach ($WsId in $WsIds) {
        $WsName = ($WsId -split '/')[-1]
        try {
            $SentinelResp = Invoke-AzRestMethod -Path "$WsId/providers/Microsoft.SecurityInsights/onboardingStates/default?api-version=2024-03-01" -Method GET -ErrorAction Stop
            $Onboarded = $SentinelResp -and $SentinelResp.StatusCode -eq 200
            if ($SentinelResp -and $SentinelResp.StatusCode -notin @(200, 404)) {
                [void]$AllChecks.Add((New-CheckResult -Id "MON-SIEM-$WsName" `
                    -Category 'Monitoring' -Name 'SIEM Integration' `
                    -Description 'Microsoft Sentinel should be connected to the AVD Log Analytics workspace' `
                    -Status 'Error' -Severity 'Medium' `
                    -Details "Sentinel onboarding query returned HTTP $($SentinelResp.StatusCode) for workspace $WsName" `
                    -Reference 'https://learn.microsoft.com/en-us/azure/sentinel/overview'))
                continue
            }
            [void]$AllChecks.Add((New-CheckResult -Id "MON-SIEM-$WsName" `
                -Category 'Monitoring' -Name 'SIEM Integration' `
                -Description 'Microsoft Sentinel should be connected to the AVD Log Analytics workspace' `
                -Status $(if ($Onboarded) { 'Pass' } else { 'Warning' }) `
                -Severity 'Medium' `
                -Details "Workspace ${WsName}: Microsoft Sentinel $(if ($Onboarded) { 'onboarded' } else { 'not onboarded (HTTP 404)' })" `
                -Recommendation 'Onboard Microsoft Sentinel onto the AVD Log Analytics workspace for security event correlation and AVD analytics rules.' `
                -Reference 'https://learn.microsoft.com/en-us/azure/sentinel/overview' `
                -Evidence @{ Workspace = $WsName; Onboarded = $Onboarded }))
        } catch {
            [void]$AllChecks.Add((New-CheckResult -Id "MON-SIEM-$WsName" `
                -Category 'Monitoring' -Name 'SIEM Integration' `
                -Description 'Microsoft Sentinel should be connected to the AVD Log Analytics workspace' `
                -Status 'Error' -Severity 'Medium' `
                -Details "Could not query Sentinel onboarding for workspace $WsName : $($_.Exception.Message)" `
                -Reference 'https://learn.microsoft.com/en-us/azure/sentinel/overview'))
        }
    }
    Write-Status "  Log Analytics workspaces checked for Sentinel: $($WsIds.Count)" -Level 'SUCCESS'
}

# ═══════════════════════════════════════════════════════════════════════════
# AVD INSIGHTS / LOG ANALYTICS KQL (optional: Az.OperationalInsights)
# NET-008 latency (MON-LATENCY-*), MON-002 Insights data (MON-INSIGHTS-*),
# MON-008 Perf (MON-PERF-*), MON-009 Events (MON-EVENTS-*), MON-010 storage IOPS
# (MON-STORIOPS-*), PROF-010 profile load times (PROF-LOADTIME-*).
# Reuses the workspace resource IDs harvested from host-pool diagnostic settings.
# ═══════════════════════════════════════════════════════════════════════════
Write-Status "AVD Insights (Log Analytics KQL)" -Level 'SECTION'
$KqlWsIds = @($LAWorkspaceIds.Keys)
$OpInsightsPresent = [bool](Get-Module -ListAvailable -Name Az.OperationalInsights -ErrorAction SilentlyContinue)

# Data Collection Rule facts (supplementary for MON-008/009). Best-effort in the current context.
$DcrHasPerf = $false; $DcrHasEvent = $false
try {
    $DcrSubId = (Get-AzContext).Subscription.Id
    foreach ($Dcr in @(Get-AzDataCollectionRule -SubscriptionId $DcrSubId -ErrorAction Stop)) {
        $DcrJson = ''
        try { $DcrJson = ($Dcr | ConvertTo-Json -Depth 8 -Compress) } catch { }
        if ($DcrJson -match '(?i)performanceCounters|DataSourcePerformanceCounter') { $DcrHasPerf = $true }
        if ($DcrJson -match '(?i)windowsEventLogs|DataSourceWindowsEventLog') { $DcrHasEvent = $true }
    }
} catch { }

# FSLogix storage accounts (for MON-010 scoping).
$FslStorage = @($Discovery.Inventory.StorageAccounts | Where-Object { $_.LikelyFSLogix })

$LatRef  = 'https://learn.microsoft.com/en-us/azure/virtual-desktop/rdp-shortpath'
$InsRef  = 'https://learn.microsoft.com/en-us/azure/virtual-desktop/insights'
$PerfRef = 'https://learn.microsoft.com/en-us/azure/virtual-desktop/insights#session-host-data-settings'
$EvtRef  = 'https://learn.microsoft.com/en-us/azure/azure-monitor/agents/data-collection-rule-azure-monitor-agent'
$IopsRef = 'https://learn.microsoft.com/en-us/azure/storage/files/storage-files-monitoring'
$LoadRef = 'https://learn.microsoft.com/en-us/fslogix/tutorial-configure-logging'

if (-not $OpInsightsPresent) {
    # Optional module missing → emit Error for each KQL-backed check (never hard prereq-fail).
    $ModMsg = 'Az.OperationalInsights not installed - install it to run AVD Insights / Log Analytics KQL checks.'
    [void]$AllChecks.Add((New-CheckResult -Id 'MON-LATENCY-NOMODULE' -Category 'Networking' -Name 'Network Latency Requirements' -Description 'Round-trip latency from AVD Insights should be within target' -Status 'Error' -Severity 'Medium' -Details $ModMsg -Reference $LatRef))
    [void]$AllChecks.Add((New-CheckResult -Id 'MON-INSIGHTS-NOMODULE' -Category 'Monitoring' -Name 'AVD Insights Enabled' -Description 'AVD Insights connection data should be flowing to Log Analytics' -Status 'Error' -Severity 'Medium' -Details $ModMsg -Reference $InsRef))
    [void]$AllChecks.Add((New-CheckResult -Id 'MON-PERF-NOMODULE' -Category 'Monitoring' -Name 'Performance Counters Configured' -Description 'Session-host performance counters should be collected' -Status 'Error' -Severity 'Medium' -Details $ModMsg -Reference $PerfRef))
    [void]$AllChecks.Add((New-CheckResult -Id 'MON-EVENTS-NOMODULE' -Category 'Monitoring' -Name 'Windows Event Logs Collected' -Description 'Windows event logs should be collected from session hosts' -Status 'Error' -Severity 'Medium' -Details $ModMsg -Reference $EvtRef))
    [void]$AllChecks.Add((New-CheckResult -Id 'MON-STORIOPS-NOMODULE' -Category 'Monitoring' -Name 'Storage IOPS Monitoring' -Description 'FSLogix storage throttling/IOPS should be monitored' -Status 'Error' -Severity 'Medium' -Details $ModMsg -Reference $IopsRef))
    [void]$AllChecks.Add((New-CheckResult -Id 'PROF-LOADTIME-NOMODULE' -Category 'FSLogix & Profiles' -Name 'Profile Size Management' -Description 'FSLogix profile load times should be within target' -Status 'Error' -Severity 'Medium' -Details $ModMsg -Reference $LoadRef))
    Write-Status "  $ModMsg" -Level 'WARN'
} elseif ($KqlWsIds.Count -eq 0) {
    # No workspace discovered via diagnostic settings → data cannot be flowing.
    $NoWs = 'No Log Analytics workspace found via host-pool diagnostic settings - AVD Insights data is not flowing.'
    [void]$AllChecks.Add((New-CheckResult -Id 'MON-LATENCY-NODATA' -Category 'Networking' -Name 'Network Latency Requirements' -Description 'Round-trip latency from AVD Insights should be within target' -Status 'Warning' -Severity 'Medium' -Details $NoWs -Recommendation 'Send AVD diagnostics to a Log Analytics workspace and enable AVD Insights.' -Reference $LatRef))
    [void]$AllChecks.Add((New-CheckResult -Id 'MON-INSIGHTS-NODATA' -Category 'Monitoring' -Name 'AVD Insights Enabled' -Description 'AVD Insights connection data should be flowing to Log Analytics' -Status 'Warning' -Severity 'Medium' -Details $NoWs -Recommendation 'Configure AVD diagnostic settings to a Log Analytics workspace and enable AVD Insights.' -Reference $InsRef))
    [void]$AllChecks.Add((New-CheckResult -Id 'MON-PERF-NODATA' -Category 'Monitoring' -Name 'Performance Counters Configured' -Description 'Session-host performance counters should be collected' -Status 'Warning' -Severity 'Medium' -Details $NoWs -Recommendation 'Deploy a Data Collection Rule with performance counters and send to Log Analytics.' -Reference $PerfRef))
    [void]$AllChecks.Add((New-CheckResult -Id 'MON-EVENTS-NODATA' -Category 'Monitoring' -Name 'Windows Event Logs Collected' -Description 'Windows event logs should be collected from session hosts' -Status 'Warning' -Severity 'Medium' -Details $NoWs -Recommendation 'Deploy a Data Collection Rule with Windows event log data sources and send to Log Analytics.' -Reference $EvtRef))
    if ($FslStorage.Count -eq 0) {
        [void]$AllChecks.Add((New-CheckResult -Id 'MON-STORIOPS-NONE' -Category 'Monitoring' -Name 'Storage IOPS Monitoring' -Description 'FSLogix storage throttling/IOPS should be monitored' -Status 'N/A' -Severity 'Medium' -Details 'No FSLogix storage accounts classified - nothing to monitor for IOPS.' -Reference $IopsRef))
    } else {
        [void]$AllChecks.Add((New-CheckResult -Id 'MON-STORIOPS-NODATA' -Category 'Monitoring' -Name 'Storage IOPS Monitoring' -Description 'FSLogix storage throttling/IOPS should be monitored' -Status 'Warning' -Severity 'Medium' -Details "$($FslStorage.Count) FSLogix storage account(s) but no Log Analytics workspace to query StorageFileLogs throttling." -Recommendation 'Enable storage diagnostic logs and/or a metric alert on Transactions / Success E2E latency for FSLogix storage.' -Reference $IopsRef))
    }
    [void]$AllChecks.Add((New-CheckResult -Id 'PROF-LOADTIME-NODATA' -Category 'FSLogix & Profiles' -Name 'Profile Size Management' -Description 'FSLogix profile load times should be within target' -Status 'Warning' -Severity 'Medium' -Details "$NoWs Enable FSLogix event collection to measure profile load times." -Recommendation 'Collect FSLogix operational events / WVDCheckpoints in Log Analytics to measure profile load times (<30s target).' -Reference $LoadRef))
    Write-Status "  $NoWs" -Level 'WARN'
} else {
    foreach ($WsId in $KqlWsIds) {
        $WsName = ($WsId -split '/')[-1]

        # --- NET-008: latency (WVDConnectionNetworkData) ---
        $LatQ = 'WVDConnectionNetworkData | where TimeGenerated > ago(7d) | summarize AvgRtt = avg(EstRoundTripTimeInMs)'
        $LatR = Invoke-AvdLaQuery -WorkspaceResourceId $WsId -Query $LatQ -TimespanDays 7
        if (-not $LatR.Ok) {
            [void]$AllChecks.Add((New-CheckResult -Id "MON-LATENCY-$WsName" -Category 'Networking' -Name 'Network Latency Requirements' -Description 'Round-trip latency from AVD Insights should be within target' -Status 'Error' -Severity 'Medium' -Details "KQL latency query failed for workspace ${WsName}: $($LatR.Error)" -Reference $LatRef))
        } else {
            $AvgRtt = $null
            if ($LatR.Rows.Count -gt 0 -and $null -ne $LatR.Rows[0].AvgRtt -and "$($LatR.Rows[0].AvgRtt)" -ne '') { try { $AvgRtt = [double]$LatR.Rows[0].AvgRtt } catch { $AvgRtt = $null } }
            if ($null -eq $AvgRtt) {
                [void]$AllChecks.Add((New-CheckResult -Id "MON-LATENCY-$WsName" -Category 'Networking' -Name 'Network Latency Requirements' -Description 'Round-trip latency from AVD Insights should be within target' -Status 'Warning' -Severity 'Medium' -Details "Workspace ${WsName}: AVD Insights data not flowing (no WVDConnectionNetworkData in last 7d)." -Recommendation 'Enable AVD Insights so connection network data (round-trip time) is collected.' -Reference $LatRef))
            } else {
                $LatStatus = if ($AvgRtt -lt 100) { 'Pass' } elseif ($AvgRtt -le 150) { 'Warning' } else { 'Fail' }
                [void]$AllChecks.Add((New-CheckResult -Id "MON-LATENCY-$WsName" -Category 'Networking' -Name 'Network Latency Requirements' -Description 'Round-trip latency from AVD Insights should be within target' -Status $LatStatus -Severity 'Medium' -Details "Workspace ${WsName}: avg round-trip time $([math]::Round($AvgRtt,1))ms over 7d (Pass<100, Warn 100-150, Fail>150)." -Recommendation 'Reduce latency with RDP Shortpath, region proximity, and network optimization; investigate gateways/regions above target.' -Reference $LatRef -Evidence @{ Workspace = $WsName; AvgRttMs = [math]::Round($AvgRtt,1) }))
            }
        }

        # --- MON-002: AVD Insights data flowing (WVDConnections rows) ---
        $InsQ = 'WVDConnections | where TimeGenerated > ago(7d) | summarize Rows = count()'
        $InsR = Invoke-AvdLaQuery -WorkspaceResourceId $WsId -Query $InsQ -TimespanDays 7
        if (-not $InsR.Ok) {
            [void]$AllChecks.Add((New-CheckResult -Id "MON-INSIGHTS-$WsName" -Category 'Monitoring' -Name 'AVD Insights Enabled' -Description 'AVD Insights connection data should be flowing to Log Analytics' -Status 'Error' -Severity 'Medium' -Details "KQL WVDConnections query failed for workspace ${WsName}: $($InsR.Error)" -Reference $InsRef))
        } else {
            $InsRows = if ($InsR.Rows.Count -gt 0 -and $InsR.Rows[0].Rows) { [int64]$InsR.Rows[0].Rows } else { 0 }
            [void]$AllChecks.Add((New-CheckResult -Id "MON-INSIGHTS-$WsName" -Category 'Monitoring' -Name 'AVD Insights Enabled' -Description 'AVD Insights connection data should be flowing to Log Analytics' -Status $(if ($InsRows -gt 0) { 'Pass' } else { 'Warning' }) -Severity 'Medium' -Details "Workspace ${WsName}: $InsRows WVDConnections row(s) in last 7d $(if ($InsRows -gt 0) { '(data flowing)' } else { '(no data - diagnostic settings may exist but Insights is not flowing)' })." -Recommendation 'Enable AVD Insights and verify the diagnostic settings actually send Connection data to this workspace.' -Reference $InsRef -Evidence @{ Workspace = $WsName; Rows = $InsRows }))
        }

        # --- MON-008: performance counters (Perf) ---
        $PerfQ = 'Perf | where TimeGenerated > ago(7d) | where CounterName has "User Input Delay" or ObjectName has "User Input Delay per Session" | summarize Rows = count()'
        $PerfR = Invoke-AvdLaQuery -WorkspaceResourceId $WsId -Query $PerfQ -TimespanDays 7
        if (-not $PerfR.Ok) {
            [void]$AllChecks.Add((New-CheckResult -Id "MON-PERF-$WsName" -Category 'Monitoring' -Name 'Performance Counters Configured' -Description 'Session-host performance counters should be collected' -Status 'Error' -Severity 'Medium' -Details "KQL Perf query failed for workspace ${WsName}: $($PerfR.Error)" -Reference $PerfRef))
        } else {
            $PerfRows = if ($PerfR.Rows.Count -gt 0 -and $PerfR.Rows[0].Rows) { [int64]$PerfR.Rows[0].Rows } else { 0 }
            if ($PerfRows -gt 0) {
                [void]$AllChecks.Add((New-CheckResult -Id "MON-PERF-$WsName" -Category 'Monitoring' -Name 'Performance Counters Configured' -Description 'Session-host performance counters should be collected' -Status 'Pass' -Severity 'Medium' -Details "Workspace ${WsName}: session-host performance counters (e.g. User Input Delay) present in Perf ($PerfRows rows/7d). DCR perf sources: $DcrHasPerf." -Recommendation 'Keep the session-host performance counter Data Collection Rule assigned.' -Reference $PerfRef -Evidence @{ Workspace = $WsName; PerfRows = $PerfRows; DcrHasPerf = $DcrHasPerf }))
            } else {
                [void]$AllChecks.Add((New-CheckResult -Id "MON-PERF-$WsName" -Category 'Monitoring' -Name 'Performance Counters Configured' -Description 'Session-host performance counters should be collected' -Status 'Warning' -Severity 'Medium' -Details "Workspace ${WsName}: no session-host performance counters in Perf (last 7d). $(if ($DcrHasPerf) { 'A Data Collection Rule with performance counters exists but no data is arriving.' } else { 'No Data Collection Rule with performance counters found.' })" -Recommendation 'Deploy/assign a Data Collection Rule collecting AVD performance counters (User Input Delay per Session, Processor, Memory) to session hosts.' -Reference $PerfRef -Evidence @{ Workspace = $WsName; DcrHasPerf = $DcrHasPerf }))
            }
        }

        # --- MON-009: Windows event logs (Event) ---
        $EvtQ = 'Event | where TimeGenerated > ago(7d) | summarize Rows = count()'
        $EvtR = Invoke-AvdLaQuery -WorkspaceResourceId $WsId -Query $EvtQ -TimespanDays 7
        if (-not $EvtR.Ok) {
            [void]$AllChecks.Add((New-CheckResult -Id "MON-EVENTS-$WsName" -Category 'Monitoring' -Name 'Windows Event Logs Collected' -Description 'Windows event logs should be collected from session hosts' -Status 'Error' -Severity 'Medium' -Details "KQL Event query failed for workspace ${WsName}: $($EvtR.Error)" -Reference $EvtRef))
        } else {
            $EvtRows = if ($EvtR.Rows.Count -gt 0 -and $EvtR.Rows[0].Rows) { [int64]$EvtR.Rows[0].Rows } else { 0 }
            if ($EvtRows -gt 0) {
                [void]$AllChecks.Add((New-CheckResult -Id "MON-EVENTS-$WsName" -Category 'Monitoring' -Name 'Windows Event Logs Collected' -Description 'Windows event logs should be collected from session hosts' -Status 'Pass' -Severity 'Medium' -Details "Workspace ${WsName}: Windows event log data present in Event table ($EvtRows rows/7d). DCR event sources: $DcrHasEvent." -Recommendation 'Keep the Windows event log Data Collection Rule assigned to session hosts.' -Reference $EvtRef -Evidence @{ Workspace = $WsName; EventRows = $EvtRows; DcrHasEvent = $DcrHasEvent }))
            } else {
                [void]$AllChecks.Add((New-CheckResult -Id "MON-EVENTS-$WsName" -Category 'Monitoring' -Name 'Windows Event Logs Collected' -Description 'Windows event logs should be collected from session hosts' -Status 'Warning' -Severity 'Medium' -Details "Workspace ${WsName}: no Windows event log data in Event table (last 7d). $(if ($DcrHasEvent) { 'A Data Collection Rule with Windows event logs exists but no data is arriving.' } else { 'No Data Collection Rule with Windows event log sources found.' })" -Recommendation 'Deploy/assign a Data Collection Rule with Windows event log data sources to session hosts.' -Reference $EvtRef -Evidence @{ Workspace = $WsName; DcrHasEvent = $DcrHasEvent }))
            }
        }

        # --- MON-010: storage IOPS / throttling (StorageFileLogs) ---
        if ($FslStorage.Count -eq 0) {
            [void]$AllChecks.Add((New-CheckResult -Id "MON-STORIOPS-$WsName" -Category 'Monitoring' -Name 'Storage IOPS Monitoring' -Description 'FSLogix storage throttling/IOPS should be monitored' -Status 'N/A' -Severity 'Medium' -Details "Workspace ${WsName}: no FSLogix storage accounts classified - nothing to monitor for IOPS." -Reference $IopsRef))
        } else {
            $FslIds = @($FslStorage | ForEach-Object { "$($_.Id)".ToLower() })
            $IopsAlerts = @($Discovery.Inventory.AlertRules | Where-Object { $_.TargetResource -and ("$($_.TargetResource)".ToLower() -in $FslIds) })
            if ($IopsAlerts.Count -gt 0) {
                [void]$AllChecks.Add((New-CheckResult -Id "MON-STORIOPS-$WsName" -Category 'Monitoring' -Name 'Storage IOPS Monitoring' -Description 'FSLogix storage throttling/IOPS should be monitored' -Status 'Pass' -Severity 'Medium' -Details "$($IopsAlerts.Count) metric/log alert(s) target FSLogix storage account(s) (Transactions / E2E latency monitoring in place)." -Recommendation 'Keep alerts on FSLogix storage Transactions and Success E2E latency to catch throttling early.' -Reference $IopsRef -Evidence @{ Workspace = $WsName; AlertCount = $IopsAlerts.Count }))
            } else {
                $ThrQ = 'StorageFileLogs | where TimeGenerated > ago(7d) | where StatusCode == 429 or StatusText has "Throttl" or StatusText has "ServerBusy" | summarize Rows = count()'
                $ThrR = Invoke-AvdLaQuery -WorkspaceResourceId $WsId -Query $ThrQ -TimespanDays 7
                if (-not $ThrR.Ok) {
                    [void]$AllChecks.Add((New-CheckResult -Id "MON-STORIOPS-$WsName" -Category 'Monitoring' -Name 'Storage IOPS Monitoring' -Description 'FSLogix storage throttling/IOPS should be monitored' -Status 'Warning' -Severity 'Medium' -Details "Workspace ${WsName}: no metric alert on FSLogix storage and StorageFileLogs not queryable ($($ThrR.Error))." -Recommendation 'Add a metric alert on FSLogix storage Transactions / Success E2E latency, or enable StorageFileLogs diagnostics.' -Reference $IopsRef))
                } else {
                    $ThrRows = if ($ThrR.Rows.Count -gt 0 -and $ThrR.Rows[0].Rows) { [int64]$ThrR.Rows[0].Rows } else { 0 }
                    [void]$AllChecks.Add((New-CheckResult -Id "MON-STORIOPS-$WsName" -Category 'Monitoring' -Name 'Storage IOPS Monitoring' -Description 'FSLogix storage throttling/IOPS should be monitored' -Status 'Warning' -Severity 'Medium' -Details "Workspace ${WsName}: no metric alert on FSLogix storage. StorageFileLogs throttling events in last 7d: $ThrRows$(if ($ThrRows -gt 0) { ' (throttling observed - storage is undersized)' } else { '' })." -Recommendation 'Create a metric alert on FSLogix storage Transactions and Success E2E latency; upsize storage if throttling (429) is observed.' -Reference $IopsRef -Evidence @{ Workspace = $WsName; ThrottleRows = $ThrRows }))
                }
            }
        }

        # --- PROF-010: profile load times (WVDCheckpoints) ---
        $LoadQ = 'WVDCheckpoints | where TimeGenerated > ago(7d) | where Name has "Fslogix" or Name has "Profile" | extend Dur = todouble(column_ifexists("DurationMs", 0)) | summarize AvgSec = avg(Dur)/1000.0, Samples = count()'
        $LoadR = Invoke-AvdLaQuery -WorkspaceResourceId $WsId -Query $LoadQ -TimespanDays 7
        if (-not $LoadR.Ok) {
            # Degrade to Warning (enable FSLogix event collection) rather than a hard Error - the checkpoint
            # schema for FSLogix phases is not always present.
            [void]$AllChecks.Add((New-CheckResult -Id "PROF-LOADTIME-$WsName" -Category 'FSLogix & Profiles' -Name 'Profile Size Management' -Description 'FSLogix profile load times should be within target' -Status 'Warning' -Severity 'Medium' -Details "Workspace ${WsName}: could not measure profile load times ($($LoadR.Error)). Enable FSLogix event collection." -Recommendation 'Enable FSLogix operational event / WVDCheckpoints collection to measure profile load times (<30s target).' -Reference $LoadRef))
        } else {
            $AvgSec = $null; $Samples = 0
            if ($LoadR.Rows.Count -gt 0) {
                if ($LoadR.Rows[0].Samples) { $Samples = [int64]$LoadR.Rows[0].Samples }
                if ($null -ne $LoadR.Rows[0].AvgSec -and "$($LoadR.Rows[0].AvgSec)" -ne '') { try { $AvgSec = [double]$LoadR.Rows[0].AvgSec } catch { $AvgSec = $null } }
            }
            if ($Samples -eq 0 -or $null -eq $AvgSec) {
                [void]$AllChecks.Add((New-CheckResult -Id "PROF-LOADTIME-$WsName" -Category 'FSLogix & Profiles' -Name 'Profile Size Management' -Description 'FSLogix profile load times should be within target' -Status 'Warning' -Severity 'Medium' -Details "Workspace ${WsName}: no FSLogix profile load-time checkpoints in last 7d. Enable FSLogix event collection." -Recommendation 'Enable FSLogix operational event / WVDCheckpoints collection to measure profile load times (<30s target).' -Reference $LoadRef))
            } else {
                [void]$AllChecks.Add((New-CheckResult -Id "PROF-LOADTIME-$WsName" -Category 'FSLogix & Profiles' -Name 'Profile Size Management' -Description 'FSLogix profile load times should be within target' -Status $(if ($AvgSec -lt 30) { 'Pass' } else { 'Warning' }) -Severity 'Medium' -Details "Workspace ${WsName}: avg FSLogix profile load $([math]::Round($AvgSec,1))s over $Samples checkpoint(s) (target <30s)." -Recommendation 'If profile loads exceed 30s, reduce profile size, use Premium storage, and enable profile trimming / redirections.' -Reference $LoadRef -Evidence @{ Workspace = $WsName; AvgLoadSec = [math]::Round($AvgSec,1); Samples = $Samples }))
            }
        }
    }
    Write-Status "  KQL checks run across $($KqlWsIds.Count) workspace(s)" -Level 'SUCCESS'
}

# ═══════════════════════════════════════════════════════════════════════════
# REGION PROXIMITY (SH-018) — session hosts vs FSLogix storage regions, per host pool
# ═══════════════════════════════════════════════════════════════════════════
Write-Status "Region Proximity" -Level 'SECTION'
$FSLogixLocations = @($Discovery.Inventory.StorageAccounts | Where-Object { $_.LikelyFSLogix -and $_.Location } | ForEach-Object { "$($_.Location)".ToLower() } | Sort-Object -Unique)
foreach ($HP in $Discovery.Inventory.HostPools) {
    $HPHostLocs = @($Discovery.Inventory.SessionHosts | Where-Object { $_.HostPoolName -eq $HP.Name -and $_.Location } | ForEach-Object { "$($_.Location)".ToLower() } | Sort-Object -Unique)
    if ($HPHostLocs.Count -eq 0) { continue }
    if ($FSLogixLocations.Count -eq 0) {
        [void]$AllChecks.Add((New-CheckResult -Id "SH-PROX-$($HP.Name)" `
            -Category 'Session Hosts' -Name 'Region Proximity' `
            -Description 'Session hosts and their FSLogix profile storage should be in the same Azure region' `
            -Status 'N/A' -Severity 'Medium' `
            -Details "HostPool $($HP.Name) hosts in: $($HPHostLocs -join ', '); no FSLogix storage accounts classified to compare against." `
            -Reference 'https://learn.microsoft.com/en-us/azure/well-architected/azure-virtual-desktop/networking'))
        continue
    }
    $Mismatches = @($HPHostLocs | Where-Object { $_ -notin $FSLogixLocations })
    if ($Mismatches.Count -eq 0) {
        [void]$AllChecks.Add((New-CheckResult -Id "SH-PROX-$($HP.Name)" `
            -Category 'Session Hosts' -Name 'Region Proximity' `
            -Description 'Session hosts and their FSLogix profile storage should be in the same Azure region' `
            -Status 'Pass' -Severity 'Medium' `
            -Details "HostPool $($HP.Name): all hosts co-located with FSLogix storage region(s) ($($HPHostLocs -join ', '))" `
            -Recommendation 'Keep session hosts and FSLogix profile storage in the same region to avoid cross-region profile-mount latency.' `
            -Reference 'https://learn.microsoft.com/en-us/azure/well-architected/azure-virtual-desktop/networking' `
            -Evidence @{ HostPool = $HP.Name; HostRegions = $HPHostLocs; FSLogixRegions = $FSLogixLocations }))
    } else {
        [void]$AllChecks.Add((New-CheckResult -Id "SH-PROX-$($HP.Name)" `
            -Category 'Session Hosts' -Name 'Region Proximity' `
            -Description 'Session hosts and their FSLogix profile storage should be in the same Azure region' `
            -Status 'Warning' -Severity 'Medium' `
            -Details "HostPool $($HP.Name): host region(s) not matching FSLogix storage. Hosts: $($HPHostLocs -join ', '); FSLogix: $($FSLogixLocations -join ', '); mismatched: $($Mismatches -join ', ')" `
            -Recommendation 'Move session hosts or profile storage so both share a region; cross-region profile mounts add 50-200ms per file operation.' `
            -Reference 'https://learn.microsoft.com/en-us/azure/well-architected/azure-virtual-desktop/networking' `
            -Evidence @{ HostPool = $HP.Name; HostRegions = $HPHostLocs; FSLogixRegions = $FSLogixLocations; Mismatched = $Mismatches }))
    }
}

# ─── APP ATTACH: estate-level N/A when nothing found anywhere (APP-002) ───
if (@($AllChecks | Where-Object { $_.Id -like 'APP-ATTACH-*' -and $_.Id -notlike 'APP-ATTACH-NONE*' }).Count -eq 0) {
    [void]$AllChecks.Add((New-CheckResult -Id 'APP-ATTACH-NONE' `
        -Category 'Application Delivery' -Name 'App Attach' `
        -Description 'App Attach decouples application lifecycle from the golden image' `
        -Status 'N/A' -Severity 'Low' `
        -Details 'No MSIX or CIM App Attach packages found in any assessed host pool or subscription — App Attach is not in use.' `
        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/app-attach-overview'))
}

# ─── MULTI-REGION CHECK (aggregated across ALL subscriptions, A-1) ────
$HPLocations = @($Discovery.Inventory.HostPools | ForEach-Object { $_.Location } | Sort-Object -Unique)
if ($HPLocations.Count -gt 1) {
    [void]$AllChecks.Add((New-CheckResult -Id "BCDR-MULTIREGION" `
        -Category 'BCDR' -Name 'Multi-Region Host Pool' `
        -Description 'Host pools deployed in multiple regions for disaster recovery' `
        -Status 'Pass' -Severity 'Medium' `
        -Details "Regions (all subscriptions): $($HPLocations -join ', ')" `
        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/disaster-recovery'))
} else {
    [void]$AllChecks.Add((New-CheckResult -Id "BCDR-MULTIREGION" `
        -Category 'BCDR' -Name 'Multi-Region Host Pool' `
        -Description 'Host pools concentrated in single region - DR risk' `
        -Status 'Warning' -Severity 'Medium' `
        -Details "Regions (all subscriptions): $(if ($HPLocations) { $HPLocations -join ', ' } else { 'None' })"))
}

# ─── DR CAPACITY RESERVATION (BCDR-011) — estate-level, reuses per-sub reservation data ────
$AllCapResRegions = @($Discovery.Inventory.CapacityReservations | Where-Object { $_.Location } | ForEach-Object { "$($_.Location)".ToLower() } | Sort-Object -Unique)
$HPRegionsLower = @($HPLocations | ForEach-Object { "$_".ToLower() })
if ($HPLocations.Count -le 1) {
    [void]$AllChecks.Add((New-CheckResult -Id "BCDR-DRCAP" `
        -Category 'BCDR' -Name 'Capacity Reservation for DR' `
        -Description 'On-demand capacity reservations should exist in the secondary/failover region to guarantee VM availability during a regional outage' `
        -Status 'N/A' -Severity 'Medium' `
        -Details "Single-region estate ($(if ($HPLocations) { $HPLocations -join ', ' } else { 'none' })) - no secondary region to reserve capacity in." `
        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-machines/capacity-reservation-overview'))
} else {
    $PrimaryRegion = $HPRegionsLower[0]
    $DrReservations = @($AllCapResRegions | Where-Object { $_ -ne $PrimaryRegion })
    if ($DrReservations.Count -gt 0) {
        [void]$AllChecks.Add((New-CheckResult -Id "BCDR-DRCAP" `
            -Category 'BCDR' -Name 'Capacity Reservation for DR' `
            -Description 'On-demand capacity reservations should exist in the secondary/failover region to guarantee VM availability during a regional outage' `
            -Status 'Pass' -Severity 'Medium' `
            -Details "Capacity reservation group(s) present in non-primary region(s): $($DrReservations -join ', ') (primary host-pool region: $PrimaryRegion)" `
            -Recommendation 'Maintain capacity reservations in the failover region so DR host pools can allocate VMs during a regional outage.' `
            -Reference 'https://learn.microsoft.com/en-us/azure/virtual-machines/capacity-reservation-overview' `
            -Evidence @{ DrRegions = $DrReservations; PrimaryRegion = $PrimaryRegion }))
    } else {
        [void]$AllChecks.Add((New-CheckResult -Id "BCDR-DRCAP" `
            -Category 'BCDR' -Name 'Capacity Reservation for DR' `
            -Description 'On-demand capacity reservations should exist in the secondary/failover region to guarantee VM availability during a regional outage' `
            -Status 'Warning' -Severity 'Medium' `
            -Details "Multi-region estate ($($HPLocations -join ', ')) but no capacity reservation group found outside the primary region ($PrimaryRegion). Reservations found in: $(if ($AllCapResRegions.Count -gt 0) { $AllCapResRegions -join ', ' } else { 'none' })" `
            -Recommendation 'Create on-demand capacity reservations in the secondary/failover region to guarantee VM availability during a regional outage.' `
            -Reference 'https://learn.microsoft.com/en-us/azure/virtual-machines/capacity-reservation-overview' `
            -Evidence @{ EstateRegions = $HPLocations; ReservationRegions = $AllCapResRegions }))
    }
}

# ─── AGENT VERSION CURRENCY (fleet-relative, C-9) ─────────────────────
# Compare each host against a maintained floor AND against the fleet's maximum
# observed agent version (more than 2 distinct versions behind fleet max = Warning).
try {
    # UPDATE PERIODICALLY: 2026-era floor for the AVD agent version.
    $MinRecommendedAgent = [Version]'1.0.10000.0'
    $FleetVersions = @()
    foreach ($SH in $Discovery.Inventory.SessionHosts) {
        if ($SH.AgentVersion) {
            try { $FleetVersions += [Version]$SH.AgentVersion } catch { }
        }
    }
    $FleetSorted = @($FleetVersions | Sort-Object -Unique -Descending)
    $FleetMax = if ($FleetSorted.Count -gt 0) { $FleetSorted[0] } else { $null }
    foreach ($SH in $Discovery.Inventory.SessionHosts) {
        if (-not $SH.AgentVersion) { continue }
        $VMName = if ($SH.ResourceId) { ($SH.ResourceId -split '/')[-1] } else { $SH.Name }
        $ParsedAgent = $null
        try { $ParsedAgent = [Version]$SH.AgentVersion } catch { }
        $AgentStatus = 'Pass'
        $AgentDetail = "AgentVersion: $($SH.AgentVersion), MinRecommended: $MinRecommendedAgent, FleetMax: $FleetMax"
        if (-not $ParsedAgent) {
            $AgentStatus = 'Warning'
            $AgentDetail = "AgentVersion: '$($SH.AgentVersion)' could not be parsed"
        } elseif ($ParsedAgent -lt $MinRecommendedAgent) {
            $AgentStatus = 'Warning'
            $AgentDetail += ' - below recommended floor'
        } else {
            # Position within the fleet's distinct-version ladder (0 = newest).
            $VersionsBehind = [array]::IndexOf($FleetSorted, $ParsedAgent)
            if ($VersionsBehind -gt 2) {
                $AgentStatus = 'Warning'
                $AgentDetail += " - $VersionsBehind distinct versions behind fleet max"
            }
        }
        [void]$AllChecks.Add((New-CheckResult -Id "OPS-AGENT-$VMName" `
            -Category 'Operations' -Name 'Agent Version Currency' `
            -Description 'AVD Agent should be at current recommended version and close to fleet max' `
            -Status $AgentStatus -Severity 'Medium' `
            -Details $AgentDetail `
            -Recommendation 'AVD agent updates automatically when VM is running. Ensure VMs are powered on periodically.' `
            -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/agent-overview' `
            -Evidence @{ VM = $VMName; AgentVersion = "$($SH.AgentVersion)"; FleetMax = "$FleetMax" }))
    }
} catch {
    Write-Status "  Agent version check error: $($_.Exception.Message)" -Level 'WARN'
}

# ─── CROSS-RESOURCE TAG COMPLIANCE (GOV-018) ──────────────────────────
# Reuses tags captured during discovery (E-7) - no per-resource Get-AzResource calls.
Write-Status "Cross-Resource Tag Compliance" -Level 'SECTION'
try {
    $RecommendedTags = @('Environment','Owner','CostCenter','Application','Department')
    $TagScores = @()
    # Score VMs (session hosts) from cached VM model tags
    foreach ($SH in $Discovery.Inventory.SessionHosts) {
        $Tags = if ($SH.Tags) { @($SH.Tags.Keys) } else { @() }
        $Found = @($RecommendedTags | Where-Object { $T = $_; $Tags | Where-Object { $_ -like "*$T*" } })
        $TagScores += [PSCustomObject]@{ Type = 'VM'; Name = $SH.Name; Score = $Found.Count; Total = $RecommendedTags.Count }
    }
    # Score VNets from cached tags
    foreach ($VNet in $Discovery.Inventory.VNets) {
        $Tags = if ($VNet.Tags) { @($VNet.Tags.Keys) } else { @() }
        $Found = @($RecommendedTags | Where-Object { $T = $_; $Tags | Where-Object { $_ -like "*$T*" } })
        $TagScores += [PSCustomObject]@{ Type = 'VNet'; Name = $VNet.Name; Score = $Found.Count; Total = $RecommendedTags.Count }
    }
    # Score Storage Accounts from cached tags
    foreach ($SA in $Discovery.Inventory.StorageAccounts) {
        $Tags = if ($SA.Tags) { @($SA.Tags.Keys) } else { @() }
        $Found = @($RecommendedTags | Where-Object { $T = $_; $Tags | Where-Object { $_ -like "*$T*" } })
        $TagScores += [PSCustomObject]@{ Type = 'Storage'; Name = $SA.Name; Score = $Found.Count; Total = $RecommendedTags.Count }
    }
    $AvgScore = if ($TagScores.Count -gt 0) {
        [math]::Round(($TagScores | ForEach-Object { $_.Score / $_.Total * 100 } | Measure-Object -Average).Average, 0)
    } else { -1 }
    $PoorlyTagged = @($TagScores | Where-Object { ($_.Score / $_.Total) -lt 0.4 }).Count
    Write-Status "  Resources scored: $($TagScores.Count), Avg tag compliance: $AvgScore%, Poorly tagged: $PoorlyTagged" -Level $(if ($AvgScore -ge 60) { 'SUCCESS' } else { 'WARN' })
    [void]$AllChecks.Add((New-CheckResult -Id "GOV-TAGALL" `
        -Category 'Governance & Cost' -Name 'Cross-Resource Tag Compliance' `
        -Description 'All AVD resources should have consistent tagging for cost allocation and governance' `
        -Status $(if ($AvgScore -ge 60) { 'Pass' } elseif ($AvgScore -ge 30) { 'Warning' } else { 'Fail' }) `
        -Severity 'Medium' `
        -Details "Resources scored: $($TagScores.Count), Avg compliance: $AvgScore%, Poorly tagged (<40%): $PoorlyTagged" `
        -Recommendation "Apply recommended tags ($($RecommendedTags -join ', ')) consistently via Azure Policy across all AVD resources." `
        -Reference 'https://learn.microsoft.com/en-us/azure/cloud-adoption-framework/ready/azure-best-practices/resource-tagging' `
        -Evidence @{ ResourceCount = $TagScores.Count; AvgScore = $AvgScore; PoorlyTagged = $PoorlyTagged }))
} catch {
    Write-Status "  Tag compliance error: $($_.Exception.Message)" -Level 'WARN'
}

# ═══════════════════════════════════════════════════════════════════════════
# IN-GUEST FSLOGIX CHECKS (opt-in: -IncludeGuestChecks)
# PROF-001 (PROF-INSTALLED-*), PROF-008 (PROF-CCACHE-*), PROF-009 (PROF-ODFC-*),
# PROF-012 (PROF-VER-*), PROF-013 (PROF-VHDX-*), PROF-014 (PROF-FLIPFLOP-*),
# PROF-015 (PROF-DELLOCAL-*). Runs Invoke-AzVMRunCommand on up to 3 running hosts/pool.
# When the switch is absent these emit N/A (honest "not enabled") so they are never silent.
# ═══════════════════════════════════════════════════════════════════════════
Write-Status "In-Guest FSLogix (Run Command)" -Level 'SECTION'
$FslMinVersion = [Version]'2.9.8612.60056'  # UPDATE PERIODICALLY: 2024-era FSLogix agent floor.

# Emits the seven guest-derived checks as N/A with a shared reason (switch off / no eligible hosts).
function Add-GuestNA {
    param([string]$Suffix, [string]$Reason)
    $Defs = @(
        @{ Id = "PROF-INSTALLED-$Suffix"; Cat = 'FSLogix & Profiles'; Name = 'FSLogix Installed'; Desc = 'FSLogix agent should be installed on session hosts'; Ref = 'https://learn.microsoft.com/en-us/fslogix/install-ht' }
        @{ Id = "PROF-CCACHE-$Suffix";    Cat = 'FSLogix & Profiles'; Name = 'Cloud Cache Configuration'; Desc = 'FSLogix Cloud Cache configuration (CCDLocations) for resilient profiles'; Ref = 'https://learn.microsoft.com/en-us/fslogix/tutorial-cloud-cache-containers' }
        @{ Id = "PROF-ODFC-$Suffix";      Cat = 'FSLogix & Profiles'; Name = 'ODFC Container Separation'; Desc = 'Office data should use a separate ODFC container'; Ref = 'https://learn.microsoft.com/en-us/fslogix/concepts-office-container' }
        @{ Id = "PROF-VER-$Suffix";       Cat = 'FSLogix & Profiles'; Name = 'FSLogix Version Current'; Desc = 'FSLogix agent should be a current supported version'; Ref = 'https://learn.microsoft.com/en-us/fslogix/whats-new' }
        @{ Id = "PROF-VHDX-$Suffix";      Cat = 'FSLogix & Profiles'; Name = 'Container Type VHDX'; Desc = 'Profile containers should use the VHDX format'; Ref = 'https://learn.microsoft.com/en-us/fslogix/reference-configuration-settings' }
        @{ Id = "PROF-FLIPFLOP-$Suffix";  Cat = 'FSLogix & Profiles'; Name = 'FlipFlopProfileDirectoryName'; Desc = 'FlipFlopProfileDirectoryName should be enabled for readable profile folders'; Ref = 'https://learn.microsoft.com/en-us/fslogix/reference-configuration-settings' }
        @{ Id = "PROF-DELLOCAL-$Suffix";  Cat = 'FSLogix & Profiles'; Name = 'DeleteLocalProfileWhenVHDShouldApply'; Desc = 'DeleteLocalProfileWhenVHDShouldApply should be enabled to avoid local/roaming conflicts'; Ref = 'https://learn.microsoft.com/en-us/fslogix/reference-configuration-settings' }
    )
    foreach ($D in $Defs) {
        [void]$AllChecks.Add((New-CheckResult -Id $D.Id -Category $D.Cat -Name $D.Name -Description $D.Desc `
            -Status 'N/A' -Severity 'Medium' -Details $Reason -Reference $D.Ref))
    }
}

if (-not $IncludeGuestChecks) {
    Add-GuestNA -Suffix 'DISABLED' -Reason 'Guest checks not enabled (-IncludeGuestChecks). Re-run with -IncludeGuestChecks to inspect in-guest FSLogix configuration.'
    Write-Status "  Skipped (run with -IncludeGuestChecks to enable in-guest FSLogix inspection)" -Level 'INFO'
} else {
    # Consolidated in-guest script (single Run Command per host). Emits one JSON line.
    $GuestScript = @'
$ErrorActionPreference = 'SilentlyContinue'
$p = Get-ItemProperty -Path 'HKLM:\SOFTWARE\FSLogix\Profiles'
$o = Get-ItemProperty -Path 'HKLM:\SOFTWARE\FSLogix\ODFC'
$ver = $null
foreach ($cand in @('C:\Program Files\FSLogix\Apps\frxsvc.exe','C:\Program Files\FSLogix\Apps\frx.exe')) {
    if (Test-Path $cand) { $ver = (Get-Item $cand).VersionInfo.FileVersion; break }
}
$installed = ($null -ne $ver) -or (Test-Path 'HKLM:\SOFTWARE\FSLogix\Profiles')
$res = [ordered]@{
    Installed          = [bool]$installed
    AgentVersion       = $ver
    Enabled            = $p.Enabled
    VHDLocations       = ($p.VHDLocations -join ';')
    VolumeType         = "$($p.VolumeType)"
    FlipFlop           = $p.FlipFlopProfileDirectoryName
    DeleteLocalProfile = $p.DeleteLocalProfileWhenVHDShouldApply
    SizeInMBs          = $p.SizeInMBs
    CCDLocations       = ($p.CCDLocations -join ';')
    ODFCEnabled        = $o.Enabled
    ODFCVHDLocations   = ($o.VHDLocations -join ';')
}
'FSLOGIXJSON:' + ($res | ConvertTo-Json -Compress)
'@
    $GuestTmp = Join-Path $env:TEMP "avd_fslogix_guest_$(Get-Date -Format 'yyyyMMddHHmmss').ps1"
    Set-Content -Path $GuestTmp -Value $GuestScript -Encoding UTF8 -Force

    # Select up to 3 RUNNING hosts per host pool.
    $Sampled = @()
    foreach ($HP in $Discovery.Inventory.HostPools) {
        $Running = @($Discovery.Inventory.SessionHosts | Where-Object {
            $_.HostPoolName -eq $HP.Name -and "$($_.PowerState)" -match '(?i)running'
        } | Select-Object -First 3)
        foreach ($R in $Running) { $Sampled += $R }
    }

    if ($Sampled.Count -eq 0) {
        Add-GuestNA -Suffix 'NORUNNING' -Reason 'Guest checks enabled but no running session hosts were available to sample (Run Command requires a running VM).'
        Write-Status "  No running session hosts to sample" -Level 'WARN'
    } else {
        $CurrentGuestSub = $null
        foreach ($SH in $Sampled) {
            $VMName = ($SH.ResourceId -split '/')[-1]
            $VMRG   = $SH.ResourceGroup
            $VMSub  = ($SH.ResourceId -split '/')[2]
            $HostTag = ($SH.Name -split '/')[-1]
            if (-not $HostTag) { $HostTag = $VMName }
            $SafeTag = ($HostTag -replace '[^A-Za-z0-9]', '_')

            if ($VMSub -and $VMSub -ne $CurrentGuestSub) {
                try { Set-AzContext -SubscriptionId $VMSub -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null; $CurrentGuestSub = $VMSub }
                catch { Write-Status "    Could not switch context to $VMSub for $VMName" -Level 'WARN' }
            }

            $Parsed = $null; $RunError = $null
            try {
                $RunResult = Invoke-AzVMRunCommand -ResourceGroupName $VMRG -VMName $VMName -CommandId 'RunPowerShellScript' -ScriptPath $GuestTmp -ErrorAction Stop
                $OutText = ''
                if ($RunResult -and $RunResult.Value) { $OutText = ($RunResult.Value | ForEach-Object { $_.Message }) -join "`n" }
                $JsonLine = @($OutText -split "`n" | Where-Object { $_ -match 'FSLOGIXJSON:' } | Select-Object -First 1)
                if ($JsonLine.Count -gt 0) {
                    $JsonText = ($JsonLine[0] -replace '^.*FSLOGIXJSON:', '').Trim()
                    $Parsed = $JsonText | ConvertFrom-Json -ErrorAction Stop
                } else {
                    $RunError = 'Run Command returned no FSLogix JSON payload.'
                }
            } catch {
                $RunError = $_.Exception.Message
            }

            if (-not $Parsed) {
                # Whole-host failure → Error on the seven checks for this host (never crash).
                foreach ($Def in @(
                    @{ Id = "PROF-INSTALLED-$SafeTag"; Name = 'FSLogix Installed'; Cat = 'FSLogix & Profiles' }
                    @{ Id = "PROF-CCACHE-$SafeTag";    Name = 'Cloud Cache Configuration'; Cat = 'FSLogix & Profiles' }
                    @{ Id = "PROF-ODFC-$SafeTag";      Name = 'ODFC Container Separation'; Cat = 'FSLogix & Profiles' }
                    @{ Id = "PROF-VER-$SafeTag";       Name = 'FSLogix Version Current'; Cat = 'FSLogix & Profiles' }
                    @{ Id = "PROF-VHDX-$SafeTag";      Name = 'Container Type VHDX'; Cat = 'FSLogix & Profiles' }
                    @{ Id = "PROF-FLIPFLOP-$SafeTag";  Name = 'FlipFlopProfileDirectoryName'; Cat = 'FSLogix & Profiles' }
                    @{ Id = "PROF-DELLOCAL-$SafeTag";  Name = 'DeleteLocalProfileWhenVHDShouldApply'; Cat = 'FSLogix & Profiles' }
                )) {
                    [void]$AllChecks.Add((New-CheckResult -Id $Def.Id -Category $Def.Cat -Name $Def.Name `
                        -Description 'In-guest FSLogix configuration (Run Command)' -Status 'Error' -Severity 'Medium' `
                        -Details "Run Command failed on ${HostTag}: $RunError" `
                        -Reference 'https://learn.microsoft.com/en-us/fslogix/reference-configuration-settings'))
                }
                Write-Status "    ${HostTag}: Run Command failed - $RunError" -Level 'WARN'
                continue
            }

            $Ev = @{ Host = $HostTag; VM = $VMName }

            # PROF-001 Installed
            [void]$AllChecks.Add((New-CheckResult -Id "PROF-INSTALLED-$SafeTag" -Category 'FSLogix & Profiles' -Name 'FSLogix Installed' `
                -Description 'FSLogix agent should be installed on session hosts' `
                -Status $(if ($Parsed.Installed) { 'Pass' } else { 'Fail' }) -Severity 'High' `
                -Details "${HostTag}: FSLogix $(if ($Parsed.Installed) { "installed (agent $($Parsed.AgentVersion))" } else { 'NOT installed' })" `
                -Recommendation 'Install the FSLogix agent on all session hosts (bake into the golden image).' `
                -Reference 'https://learn.microsoft.com/en-us/fslogix/install-ht' -Evidence $Ev))

            # PROF-008 Cloud Cache (optional → N/A when unused)
            $CcSet = -not [string]::IsNullOrWhiteSpace("$($Parsed.CCDLocations)")
            [void]$AllChecks.Add((New-CheckResult -Id "PROF-CCACHE-$SafeTag" -Category 'FSLogix & Profiles' -Name 'Cloud Cache Configuration' `
                -Description 'FSLogix Cloud Cache (CCDLocations) for resilient/multi-region profiles' `
                -Status $(if ($CcSet) { 'Pass' } else { 'N/A' }) -Severity 'Medium' `
                -Details "${HostTag}: $(if ($CcSet) { "Cloud Cache configured (CCDLocations: $($Parsed.CCDLocations))" } else { 'Cloud Cache not configured (single-location VHD - optional feature)' })" `
                -Recommendation 'Use Cloud Cache (CCDLocations) only when profile resilience across multiple storage providers/regions is required.' `
                -Reference 'https://learn.microsoft.com/en-us/fslogix/tutorial-cloud-cache-containers' -Evidence $Ev))

            # PROF-009 ODFC separation
            $OdfcOn = ("$($Parsed.ODFCEnabled)" -eq '1') -and -not [string]::IsNullOrWhiteSpace("$($Parsed.ODFCVHDLocations)")
            [void]$AllChecks.Add((New-CheckResult -Id "PROF-ODFC-$SafeTag" -Category 'FSLogix & Profiles' -Name 'ODFC Container Separation' `
                -Description 'Office data should use a separate ODFC container' `
                -Status $(if ($OdfcOn) { 'Pass' } else { 'Warning' }) -Severity 'Medium' `
                -Details "${HostTag}: ODFC Enabled=$($Parsed.ODFCEnabled), VHDLocations=$(if ($Parsed.ODFCVHDLocations) { $Parsed.ODFCVHDLocations } else { '(none)' })" `
                -Recommendation 'Separate the Office cache into an ODFC container to reduce profile size and improve reliability.' `
                -Reference 'https://learn.microsoft.com/en-us/fslogix/concepts-office-container' -Evidence $Ev))

            # PROF-012 version
            $VerOk = $false; $VerDetail = 'agent version unknown'
            if ($Parsed.AgentVersion) {
                try { $VerOk = ([Version]$Parsed.AgentVersion -ge $FslMinVersion); $VerDetail = "agent $($Parsed.AgentVersion) (floor $FslMinVersion)" }
                catch { $VerDetail = "agent $($Parsed.AgentVersion) (unparseable)" }
            }
            [void]$AllChecks.Add((New-CheckResult -Id "PROF-VER-$SafeTag" -Category 'FSLogix & Profiles' -Name 'FSLogix Version Current' `
                -Description 'FSLogix agent should be a current supported version' `
                -Status $(if ($Parsed.AgentVersion -and $VerOk) { 'Pass' } else { 'Warning' }) -Severity 'Medium' `
                -Details "${HostTag}: $VerDetail" `
                -Recommendation 'Keep the FSLogix agent current (bake the latest supported release into the image).' `
                -Reference 'https://learn.microsoft.com/en-us/fslogix/whats-new' -Evidence $Ev))

            # PROF-013 VHDX
            $IsVhdx = ("$($Parsed.VolumeType)" -match '(?i)vhdx')
            [void]$AllChecks.Add((New-CheckResult -Id "PROF-VHDX-$SafeTag" -Category 'FSLogix & Profiles' -Name 'Container Type VHDX' `
                -Description 'Profile containers should use the VHDX format' `
                -Status $(if ($IsVhdx) { 'Pass' } else { 'Warning' }) -Severity 'Medium' `
                -Details "${HostTag}: VolumeType=$(if ($Parsed.VolumeType) { $Parsed.VolumeType } else { '(not set - defaults to VHD)' })" `
                -Recommendation 'Set VolumeType=vhdx (dynamic VHDX) for FSLogix profile containers.' `
                -Reference 'https://learn.microsoft.com/en-us/fslogix/reference-configuration-settings' -Evidence $Ev))

            # PROF-014 FlipFlop
            $FlipOn = ("$($Parsed.FlipFlop)" -eq '1')
            [void]$AllChecks.Add((New-CheckResult -Id "PROF-FLIPFLOP-$SafeTag" -Category 'FSLogix & Profiles' -Name 'FlipFlopProfileDirectoryName' `
                -Description 'FlipFlopProfileDirectoryName should be enabled for readable profile folder names' `
                -Status $(if ($FlipOn) { 'Pass' } else { 'Warning' }) -Severity 'Low' `
                -Details "${HostTag}: FlipFlopProfileDirectoryName=$(if ($null -ne $Parsed.FlipFlop) { $Parsed.FlipFlop } else { '(not set)' })" `
                -Recommendation 'Enable FlipFlopProfileDirectoryName=1 so profile folders are named %username%%sid% (easier to manage).' `
                -Reference 'https://learn.microsoft.com/en-us/fslogix/reference-configuration-settings' -Evidence $Ev))

            # PROF-015 DeleteLocalProfileWhenVHDShouldApply
            $DelOn = ("$($Parsed.DeleteLocalProfile)" -eq '1')
            [void]$AllChecks.Add((New-CheckResult -Id "PROF-DELLOCAL-$SafeTag" -Category 'FSLogix & Profiles' -Name 'DeleteLocalProfileWhenVHDShouldApply' `
                -Description 'DeleteLocalProfileWhenVHDShouldApply should be enabled to prevent local/roaming profile conflicts' `
                -Status $(if ($DelOn) { 'Pass' } else { 'Warning' }) -Severity 'Medium' `
                -Details "${HostTag}: DeleteLocalProfileWhenVHDShouldApply=$(if ($null -ne $Parsed.DeleteLocalProfile) { $Parsed.DeleteLocalProfile } else { '(not set)' })" `
                -Recommendation 'Enable DeleteLocalProfileWhenVHDShouldApply=1 to remove stale local profiles that block container mount.' `
                -Reference 'https://learn.microsoft.com/en-us/fslogix/reference-configuration-settings' -Evidence $Ev))

            Write-Status "    ${HostTag}: FSLogix installed=$($Parsed.Installed), ver=$($Parsed.AgentVersion), volumeType=$($Parsed.VolumeType)" -Level 'SUCCESS'
        }
    }
    Remove-Item -Path $GuestTmp -Force -ErrorAction SilentlyContinue
}

# ═══════════════════════════════════════════════════════════════════════════
# MATURITY SCORING ENGINE
# ═══════════════════════════════════════════════════════════════════════════

# Map check ID prefixes to maturity dimensions
<#
.SYNOPSIS
    Calculates maturity scores across six dimensions from automated check results.
.DESCRIPTION
    Maps check IDs to six maturity dimensions (Security, Operations, Networking, Resiliency,
    Profiles, Monitoring) by prefix matching, computes weighted scores per dimension, and
    derives a composite maturity level (Initial/Developing/Defined/Managed/Optimized).
.PARAMETER Checks
    ArrayList of check result objects from the discovery run.
.OUTPUTS
    PSCustomObject with Dimensions (ordered hashtable), CompositeScore (0-100), and MaturityLevel.
#>
function Get-MaturityScores {
    param([System.Collections.ArrayList]$Checks)

    $Dimensions = [ordered]@{
        Security    = @{ Prefixes = @('SEC-','IAM-');           Label = 'Security & Identity';   Icon = [char]0x26E8 }
        Operations  = @{ Prefixes = @('OPS-','GOV-','SH-LB','SH-002','SH-STATUS','SH-IMG','SH-BSERIES','SH-SSD')
                         Label = 'Operations & Cost';          Icon = [char]0x2699 }
        Networking  = @{ Prefixes = @('NET-');                  Label = 'Networking';             Icon = [char]0x2637 }
        Resiliency  = @{ Prefixes = @('BCDR-','SH-EPHEMERAL'); Label = 'Resiliency & BCDR';      Icon = [char]0x2694 }
        Profiles    = @{ Prefixes = @('PROF-');                 Label = 'FSLogix & Profiles';     Icon = [char]0x2750 }
        Monitoring  = @{ Prefixes = @('MON-');                  Label = 'Monitoring & Telemetry'; Icon = [char]0x2261 }
    }

    $Results = [ordered]@{}
    foreach ($Dim in $Dimensions.GetEnumerator()) {
        $DimChecks = @($Checks | Where-Object {
            $Id = $_.Id
            $Dim.Value.Prefixes | Where-Object { $Id -like "$_*" }
        })
        $Scoreable = @($DimChecks | Where-Object { $_.Status -in @('Pass','Fail','Warning') })
        if ($Scoreable.Count -eq 0) {
            $Results[$Dim.Key] = [PSCustomObject]@{
                Label = $Dim.Value.Label; Score = -1
                Pass = 0; Warn = 0; Fail = 0; Total = 0
                Icon = $Dim.Value.Icon
            }
            continue
        }
        $WSum = 0; $WMax = 0
        foreach ($C in $Scoreable) {
            $W = 3  # default weight
            # Try extracting weight from matching checks.json - use simple heuristic
            $Pts = switch ($C.Status) { 'Pass' { 100 } 'Warning' { 50 } 'Fail' { 0 } default { 0 } }
            $WSum += $Pts * $W
            $WMax += 100 * $W
        }
        $Score = if ($WMax -gt 0) { [math]::Round($WSum / $WMax * 100, 0) } else { -1 }
        $Results[$Dim.Key] = [PSCustomObject]@{
            Label = $Dim.Value.Label; Score = $Score
            Pass = @($Scoreable | Where-Object Status -eq 'Pass').Count
            Warn = @($Scoreable | Where-Object Status -eq 'Warning').Count
            Fail = @($Scoreable | Where-Object Status -eq 'Fail').Count
            Total = $Scoreable.Count
            Icon = $Dim.Value.Icon
        }
    }

    # Composite maturity score - weighted average of dimensions
    $ValidDims = @($Results.Values | Where-Object { $_.Score -ge 0 })
    $CompositeScore = if ($ValidDims.Count -gt 0) {
        [math]::Round(($ValidDims | Measure-Object -Property Score -Average).Average, 0)
    } else { -1 }

    # Maturity level
    $MaturityLevel = switch ($true) {
        ($CompositeScore -ge 90) { 'Optimized' }
        ($CompositeScore -ge 75) { 'Managed' }
        ($CompositeScore -ge 55) { 'Defined' }
        ($CompositeScore -ge 35) { 'Developing' }
        ($CompositeScore -ge 0)  { 'Initial' }
        default                  { 'Not Scored' }
    }

    return [PSCustomObject]@{
        Dimensions     = $Results
        CompositeScore = $CompositeScore
        MaturityLevel  = $MaturityLevel
    }
}

$MaturityResult = Get-MaturityScores -Checks $AllChecks

# Add maturity to discovery output
$Discovery | Add-Member -NotePropertyName 'Maturity' -NotePropertyValue ([PSCustomObject]@{
    CompositeScore = $MaturityResult.CompositeScore
    MaturityLevel  = $MaturityResult.MaturityLevel
    Dimensions     = $MaturityResult.Dimensions
}) -Force

# ═══════════════════════════════════════════════════════════════════════════
# FINALIZE
# ═══════════════════════════════════════════════════════════════════════════

$Discovery.CheckResults = $AllChecks.ToArray()

# Save JSON
if (-not $OutputPath) {
    $OutputPath = Join-Path $ScriptRoot "assessments\discovery_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
}
$OutputDir = Split-Path $OutputPath -Parent
if (-not (Test-Path $OutputDir)) {
    New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
}
$Discovery | ConvertTo-Json -Depth 10 | Set-Content $OutputPath -Encoding UTF8 -Force

# Summary
$PassCount    = @($AllChecks | Where-Object { $_.Status -eq 'Pass' }).Count
$FailCount    = @($AllChecks | Where-Object { $_.Status -eq 'Fail' }).Count
$WarnCount    = @($AllChecks | Where-Object { $_.Status -eq 'Warning' }).Count
$NACount      = @($AllChecks | Where-Object { $_.Status -eq 'N/A' }).Count
$TotalChecks  = $AllChecks.Count
$ScorePercent = if (($PassCount + $WarnCount + $FailCount) -gt 0) {
    [math]::Round(($PassCount * 100 + $WarnCount * 50) / ($PassCount + $WarnCount + $FailCount), 0)
} else { 0 }
$ScoreColor   = if ($ScorePercent -ge 80) { 'Green' } elseif ($ScorePercent -ge 50) { 'Yellow' } else { 'Red' }
$FileSize     = [math]::Round((Get-Item $OutputPath).Length / 1KB, 1)

# Box helper - fixed inner width of 54 chars
$BW = 54
function Write-BoxLine { param([string]$Text, [string]$Color = 'White', [string]$Prefix = '  ')
    $Pad = $BW - $Text.Length
    Write-Host "${Prefix}║" -NoNewline -ForegroundColor DarkCyan
    Write-Host $Text -NoNewline -ForegroundColor $Color
    Write-Host "$(' ' * [math]::Max(0,$Pad))║" -ForegroundColor DarkCyan
}
function Write-BoxKV { param([string]$Label, [string]$Value, [string]$LabelColor = 'Gray', [string]$ValueColor = 'White')
    $LblPad = $Label.PadRight(22)
    $ValPad = $Value.PadLeft(5)
    $Inner  = "    $LblPad$ValPad"
    Write-BoxLine $Inner $LabelColor
}
$BoxTop    = "  ╔$('═' * $BW)╗"
$BoxMid    = "  ╠$('═' * $BW)╣"
$BoxBot    = "  ╚$('═' * $BW)╝"
$BoxEmpty  = ' ' * $BW

Write-Host ""
Write-Host $BoxTop -ForegroundColor DarkCyan
$Title = 'Discovery Complete'
$TitlePad = [math]::Floor(($BW - $Title.Length) / 2)
$TitleLine = "$(' ' * $TitlePad)$Title$(' ' * ($BW - $TitlePad - $Title.Length))"
Write-Host "  ║" -NoNewline -ForegroundColor DarkCyan
Write-Host $TitleLine -NoNewline -ForegroundColor Cyan
Write-Host "║" -ForegroundColor DarkCyan
Write-Host $BoxMid -ForegroundColor DarkCyan
Write-BoxLine $BoxEmpty
Write-BoxLine '  RESOURCES DISCOVERED' 'White'
Write-BoxLine $BoxEmpty

$Resources = @(
    @{ L = 'Host Pools';       V = $Discovery.Inventory.HostPools.Count }
    @{ L = 'Session Hosts';    V = $Discovery.Inventory.SessionHosts.Count }
    @{ L = 'App Groups';       V = $Discovery.Inventory.AppGroups.Count }
    @{ L = 'Workspaces';       V = $Discovery.Inventory.Workspaces.Count }
    @{ L = 'Scaling Plans';    V = $Discovery.Inventory.ScalingPlans.Count }
    @{ L = 'Virtual Networks'; V = $Discovery.Inventory.VNets.Count }
    @{ L = 'Storage Accounts'; V = $Discovery.Inventory.StorageAccounts.Count }
    @{ L = 'Key Vaults';       V = $Discovery.Inventory.KeyVaults.Count }
    @{ L = 'Firewalls';        V = $Discovery.Inventory.Firewalls.Count }
    @{ L = 'VPN/ER Gateways';  V = $Discovery.Inventory.VPNGateways.Count }
)
foreach ($R in $Resources) { Write-BoxKV $R.L "$($R.V)" }

Write-BoxLine $BoxEmpty
Write-Host $BoxMid -ForegroundColor DarkCyan
Write-BoxLine $BoxEmpty

# Score line
$ScoreStr = "Score: $ScorePercent%"
$CheckLabel = "  AUTOMATED CHECKS"
$ScoreInner = "$CheckLabel$(' ' * ($BW - $CheckLabel.Length - $ScoreStr.Length - 2))$ScoreStr"
Write-Host "  ║" -NoNewline -ForegroundColor DarkCyan
Write-Host $CheckLabel -NoNewline -ForegroundColor White
$GapLen = $BW - $CheckLabel.Length - $ScoreStr.Length
Write-Host "$(' ' * [math]::Max(1,$GapLen))" -NoNewline
Write-Host $ScoreStr -NoNewline -ForegroundColor $ScoreColor
$Remainder = $BW - $CheckLabel.Length - [math]::Max(1,$GapLen) - $ScoreStr.Length
if ($Remainder -gt 0) { Write-Host "$(' ' * $Remainder)" -NoNewline }
Write-Host "║" -ForegroundColor DarkCyan

# Score bar
$BarLen = $BW - 8
$Filled = [math]::Round($BarLen * $ScorePercent / 100)
$Empty  = $BarLen - $Filled
$BarStr = ("$([char]0x2588)" * $Filled) + ("$([char]0x2591)" * $Empty)
Write-BoxLine "    $BarStr" $ScoreColor

Write-BoxLine $BoxEmpty

# Check counts
$CountItems = @(
    @{ L = "$([char]0x2713) Pass";    V = "$PassCount"; C = 'Green' }
    @{ L = "$([char]0x26A0) Warning"; V = "$WarnCount"; C = 'Yellow' }
    @{ L = "$([char]0x2717) Fail";    V = "$FailCount"; C = 'Red' }
    @{ L = "$([char]0x2500) N/A";     V = "$NACount";   C = 'DarkGray' }
)
foreach ($Ct in $CountItems) {
    $Lbl = $Ct.L.PadRight(14)
    $Val = $Ct.V.PadLeft(5)
    $Inner = "    $Lbl$Val"
    Write-BoxLine $Inner $Ct.C
}

Write-BoxLine $BoxEmpty

# Maturity dimensions
Write-Host $BoxMid -ForegroundColor DarkCyan
Write-BoxLine $BoxEmpty
$MTitle = "  MATURITY: $($MaturityResult.MaturityLevel.ToUpper())"
$MScore = "$($MaturityResult.CompositeScore)%"
Write-Host "  ║" -NoNewline -ForegroundColor DarkCyan
Write-Host $MTitle -NoNewline -ForegroundColor White
$MGap = $BW - $MTitle.Length - $MScore.Length
Write-Host "$(' ' * [math]::Max(1,$MGap))" -NoNewline
$MColor = if ($MaturityResult.CompositeScore -ge 80) { 'Green' } elseif ($MaturityResult.CompositeScore -ge 50) { 'Yellow' } else { 'Red' }
Write-Host $MScore -NoNewline -ForegroundColor $MColor
$MRem = $BW - $MTitle.Length - [math]::Max(1,$MGap) - $MScore.Length
if ($MRem -gt 0) { Write-Host "$(' ' * $MRem)" -NoNewline }
Write-Host "║" -ForegroundColor DarkCyan
Write-BoxLine $BoxEmpty

foreach ($Dim in $MaturityResult.Dimensions.GetEnumerator()) {
    $DScore = $Dim.Value.Score
    $DColor = if ($DScore -ge 80) { 'Green' } elseif ($DScore -ge 50) { 'Yellow' } elseif ($DScore -ge 0) { 'Red' } else { 'DarkGray' }
    $DLabel = $Dim.Value.Label.PadRight(24)
    $DBar = ''
    if ($DScore -ge 0) {
        $DBarLen = 16
        $DFill = [math]::Round($DBarLen * $DScore / 100)
        $DBar = "$("$([char]0x2588)" * $DFill)$("$([char]0x2591)" * ($DBarLen - $DFill))"
        $DScoreStr = "$DScore%".PadLeft(4)
    } else {
        $DBar = "$([char]0x2591)" * 16
        $DScoreStr = '  - '
    }
    $Inner = "    $DLabel$DBar $DScoreStr"
    Write-BoxLine $Inner $DColor
}

Write-BoxLine $BoxEmpty

# Errors
if ($Discovery.Errors.Count -gt 0) {
    Write-Host $BoxMid -ForegroundColor DarkCyan
    Write-BoxLine "  ERRORS: $($Discovery.Errors.Count)" 'Red'
    foreach ($Err in $Discovery.Errors) {
        $ErrShort = if ($Err.Length -gt ($BW - 6)) { $Err.Substring(0, $BW - 9) + '...' } else { $Err }
        Write-BoxLine "    $ErrShort" 'Yellow'
    }
    Write-BoxLine $BoxEmpty
}

# Output
Write-Host $BoxMid -ForegroundColor DarkCyan
Write-BoxLine '  OUTPUT' 'White'
$OutFile = Split-Path $OutputPath -Leaf
$OutTrunc = if ($OutFile.Length -gt ($BW - 6)) { $OutFile.Substring(0, $BW - 9) + '...' } else { $OutFile }
Write-BoxLine "    $OutTrunc" 'Green'
Write-BoxLine "    ${FileSize} KB" 'Gray'
Write-BoxLine $BoxEmpty
Write-Host $BoxBot -ForegroundColor DarkCyan
Write-Host ""
Write-Host "  Import into " -NoNewline -ForegroundColor Gray
Write-Host "AVD Assessor" -NoNewline -ForegroundColor Cyan
Write-Host " GUI for interactive review and reporting." -ForegroundColor Gray
Write-Host ""

