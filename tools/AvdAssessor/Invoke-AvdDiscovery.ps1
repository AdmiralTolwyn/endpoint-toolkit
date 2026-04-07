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
.EXAMPLE
    .\Invoke-AvdDiscovery.ps1
    .\Invoke-AvdDiscovery.ps1 -SubscriptionId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
    .\Invoke-AvdDiscovery.ps1 -SubscriptionId @("sub1","sub2") -OutputPath "C:\temp\discovery.json"
.NOTES
    Author : Anton Romanyuk
    Version: 0.2.0
    Date   : 2026-03-27
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string[]]$SubscriptionId,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [switch]$SkipLogin
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Strip OneDrive module path to prevent old Az version conflicts
$env:PSModulePath = ($env:PSModulePath -split ';' |
    Where-Object { $_ -notlike '*OneDrive*' }) -join ';'

$ScriptRoot = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($ScriptRoot)) { $ScriptRoot = $PWD.Path }

$ScriptVersion = '0.2.0'

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

# Pre-import modules with noisy warnings silently
foreach ($Noisy in @('Az.Network','Az.Monitor','Az.PrivateDns')) {
    Import-Module $Noisy -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
}
    @{ Name = 'Az.Monitor';               MinVersion = '4.0.0' }
    @{ Name = 'Az.Storage';               MinVersion = '5.0.0' }
)

Write-Status "Prerequisites" -Level 'SECTION'
$Missing = @()
foreach ($Mod in $RequiredModules) {
    $Installed = Get-Module -ListAvailable -Name $Mod.Name -ErrorAction SilentlyContinue |
                 Sort-Object Version -Descending | Select-Object -First 1
    if (-not $Installed) {
        $Missing += $Mod.Name
        Write-Status "$($Mod.Name) >= $($Mod.MinVersion) — MISSING" -Level 'ERROR'
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
    }
    CheckResults   = @()
    Errors         = @()
}

$AllChecks = [System.Collections.ArrayList]::new()

foreach ($SubId in $SubscriptionId) {
    Write-Status "Subscription: $SubId" -Level 'SECTION'

    try {
        Set-AzContext -SubscriptionId $SubId -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null
        $Sub = Get-AzContext
        $Discovery.Subscriptions += [PSCustomObject]@{
            Id   = $SubId
            Name = $Sub.Subscription.Name
        }
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
                HostPoolType         = $HP.HostPoolType
                LoadBalancerType     = $HP.LoadBalancerType
                MaxSessionLimit      = $HP.MaxSessionLimit
                PreferredAppGroupType = $HP.PreferredAppGroupType
                StartVMOnConnect     = $HP.StartVMOnConnect
                ValidationEnvironment = $HP.ValidationEnvironment
                Location             = $HP.Location
                Tags                 = $HP.Tag
                CustomRdpProperty    = $HP.CustomRdpProperty
            }
            $Discovery.Inventory.HostPools += $HPObj

            # ─── CHECK: Host pool type alignment ───
            [void]$AllChecks.Add((New-CheckResult -Id "SH-001-$($HP.Name)" `
                -Category 'Session Hosts' -Name 'Host Pool Type' `
                -Description 'Verify host pool type aligns with workload requirements' `
                -Status $(if ($HP.HostPoolType -eq 'Pooled') { 'Pass' } else { 'Warning' }) `
                -Severity 'Low' `
                -Details "Type: $($HP.HostPoolType), LB: $($HP.LoadBalancerType), MaxSessions: $($HP.MaxSessionLimit)" `
                -Recommendation 'Use Pooled for shared workloads (cost-effective). Personal for users needing admin rights or persistent state.' `
                -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/host-pool-load-balancing' `
                -Evidence @{ HostPool = $HP.Name; Type = $HP.HostPoolType }))

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
            $SessionLimitStatus = if ($HP.HostPoolType -eq 'Pooled' -and $HP.MaxSessionLimit -gt 0 -and $HP.MaxSessionLimit -le 999999) { 'Pass' }
                                  elseif ($HP.HostPoolType -eq 'Pooled' -and ($HP.MaxSessionLimit -le 0 -or $HP.MaxSessionLimit -gt 999999)) { 'Fail' }
                                  else { 'N/A' }
            [void]$AllChecks.Add((New-CheckResult -Id "SH-002-$($HP.Name)" `
                -Category 'Session Hosts' -Name 'Max Session Limit Configured' `
                -Description 'Pooled host pools should have an explicit max session limit' `
                -Status $SessionLimitStatus -Severity 'High' `
                -Details "MaxSessionLimit: $($HP.MaxSessionLimit)" `
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

            # Drive redirection
            $DriveVal = $ParsedRdp['drivestoredirect']
            $DrivesRestricted = ($null -ne $DriveVal -and $DriveVal -ne '*')
            [void]$AllChecks.Add((New-CheckResult -Id "SEC-DRIVE-$($HP.Name)" `
                -Category 'Security' -Name 'RDP Drive Redirection' `
                -Description 'Drive/disk redirection should be restricted to prevent data exfiltration' `
                -Status $(if ($DriveVal -eq '') { 'Pass' } elseif ($DrivesRestricted) { 'Warning' } else { 'Warning' }) `
                -Severity 'Medium' `
                -Details "drivestoredirect: $(if ($null -eq $DriveVal) { '(not set — default: all drives)' } else { "'$DriveVal'" })" `
                -Recommendation 'Set drivestoredirect:s: (empty) to block all drive redirection, or restrict to specific drives.' `
                -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/rdp-properties#device-redirection'))

            # Clipboard redirection
            $ClipVal = $ParsedRdp['redirectclipboard']
            [void]$AllChecks.Add((New-CheckResult -Id "SEC-CLIP-$($HP.Name)" `
                -Category 'Security' -Name 'RDP Clipboard Redirection' `
                -Description 'Clipboard redirection should be restricted for sensitive environments' `
                -Status $(if ($ClipVal -eq '0') { 'Pass' } else { 'Warning' }) `
                -Severity 'Medium' `
                -Details "redirectclipboard: $(if ($null -eq $ClipVal) { '(not set — default: enabled)' } else { $ClipVal })" `
                -Recommendation 'Set redirectclipboard:i:0 to disable, or use clipboard transfer direction policies for granular control.' `
                -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/rdp-properties#device-redirection'))

            # Printer redirection
            $PrintVal = $ParsedRdp['redirectprinters']
            [void]$AllChecks.Add((New-CheckResult -Id "SEC-PRINT-$($HP.Name)" `
                -Category 'Security' -Name 'RDP Printer Redirection' `
                -Description 'Printer redirection should be evaluated — disable if not required' `
                -Status $(if ($PrintVal -eq '0') { 'Pass' } else { 'Warning' }) `
                -Severity 'Low' `
                -Details "redirectprinters: $(if ($null -eq $PrintVal) { '(not set — default: enabled)' } else { $PrintVal })" `
                -Recommendation 'Set redirectprinters:i:0 if printer redirection is not needed.' `
                -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/rdp-properties#device-redirection'))

            # USB redirection
            $UsbVal = $ParsedRdp['usbdevicestoredirect']
            [void]$AllChecks.Add((New-CheckResult -Id "SEC-USB-$($HP.Name)" `
                -Category 'Security' -Name 'RDP USB Redirection' `
                -Description 'USB device redirection should be blocked unless explicitly required' `
                -Status $(if ($null -eq $UsbVal -or $UsbVal -eq '') { 'Pass' } else { 'Warning' }) `
                -Severity 'Medium' `
                -Details "usbdevicestoredirect: $(if ($null -eq $UsbVal) { '(not set — default: none)' } else { "'$UsbVal'" })" `
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

            # Camera redirection (informational — often needed for Teams)
            $CamVal = $ParsedRdp['camerastoredirect']
            [void]$AllChecks.Add((New-CheckResult -Id "SEC-CAM-$($HP.Name)" `
                -Category 'Security' -Name 'RDP Camera Redirection' `
                -Description 'Camera redirection status — required for Teams calls, evaluate for security-sensitive workloads' `
                -Status $(if ($null -eq $CamVal -or $CamVal -eq '*') { 'Pass' } else { 'Pass' }) `
                -Severity 'Low' `
                -Details "camerastoredirect: $(if ($null -eq $CamVal) { '(not set)' } else { "'$CamVal'" })" `
                -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/rdp-properties#device-redirection'))

            # Audio capture (microphone)
            $AudioCapVal = $ParsedRdp['audiocapturemode']
            [void]$AllChecks.Add((New-CheckResult -Id "SEC-AUDIO-$($HP.Name)" `
                -Category 'Security' -Name 'RDP Audio Capture (Microphone)' `
                -Description 'Audio input capture — needed for calls, evaluate for other workloads' `
                -Status 'Pass' -Severity 'Low' `
                -Details "audiocapturemode: $(if ($null -eq $AudioCapVal) { '(not set — default: disabled)' } else { $AudioCapVal })" `
                -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/rdp-properties#device-redirection'))

            # RDP property summary for evidence
            $RdpSecurityIssues = @()
            if ($null -eq $ParsedRdp['drivestoredirect'] -or $ParsedRdp['drivestoredirect'] -eq '*') { $RdpSecurityIssues += 'Drives:Open' }
            if ($null -eq $ParsedRdp['redirectclipboard'] -or $ParsedRdp['redirectclipboard'] -ne '0') { $RdpSecurityIssues += 'Clipboard:Open' }
            if ($null -eq $ParsedRdp['redirectprinters'] -or $ParsedRdp['redirectprinters'] -ne '0') { $RdpSecurityIssues += 'Printers:Open' }
            if ($ParsedRdp['usbdevicestoredirect'] -and $ParsedRdp['usbdevicestoredirect'] -ne '') { $RdpSecurityIssues += 'USB:Open' }
            if ($ParsedRdp['redirectcomports'] -eq '1') { $RdpSecurityIssues += 'COM:Open' }
            [void]$AllChecks.Add((New-CheckResult -Id "SEC-RDP-$($HP.Name)" `
                -Category 'Security' -Name 'RDP Properties Security Summary' `
                -Description 'Overall RDP property security posture' `
                -Status $(if ($RdpSecurityIssues.Count -eq 0) { 'Pass' } elseif ($RdpSecurityIssues.Count -le 2) { 'Warning' } else { 'Fail' }) `
                -Severity 'High' `
                -Details "Issues: $(if ($RdpSecurityIssues.Count -eq 0) { 'None — all redirections restricted' } else { $RdpSecurityIssues -join ', ' }). AllProps: $RdpProps" `
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
                -Recommendation $(if ($IsEntraTarget -and -not $HasSSO) { 'SSO is strongly recommended for Entra ID joined hosts — enable enablerdsaadauth:i:1.' } else { 'Enable SSO for seamless authentication.' }) `
                -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/configure-single-sign-on'))

            # ─── CHECK: Watermarking (from RDP properties) ───
            $HasWatermark = $RdpProps -match 'watermark'
            [void]$AllChecks.Add((New-CheckResult -Id "SEC-WM-$($HP.Name)" `
                -Category 'Security' -Name 'Watermarking Enabled' `
                -Description 'QR code watermarking deters screen capture and enables session tracing' `
                -Status $(if ($HasWatermark) { 'Pass' } else { 'Warning' }) `
                -Severity 'Medium' `
                -Details "Watermarking: $(if ($HasWatermark) { 'Enabled' } else { 'Not configured' })" `
                -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/watermarking'))

            # ─── CHECK: Screen Capture Protection (from RDP properties) ───
            $HasScreenCapture = $RdpProps -match 'screen capture protection'
            [void]$AllChecks.Add((New-CheckResult -Id "SEC-SCP-$($HP.Name)" `
                -Category 'Security' -Name 'Screen Capture Protection' `
                -Description 'Screen capture protection blocks screenshots and screen sharing of remote content' `
                -Status $(if ($HasScreenCapture) { 'Pass' } else { 'Warning' }) `
                -Severity 'Medium' `
                -Details "ScreenCaptureProtection: $(if ($HasScreenCapture) { 'Enabled' } else { 'Not configured' })" `
                -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/screen-capture-protection'))

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

            # ─── CHECK: Load Balancing Algorithm ───
            [void]$AllChecks.Add((New-CheckResult -Id "SH-LB-$($HP.Name)" `
                -Category 'Session Hosts' -Name 'Load Balancing Algorithm' `
                -Description 'LB should align with workload (BreadthFirst for ramp-up, DepthFirst for cost)' `
                -Status 'Pass' -Severity 'Medium' `
                -Details "Algorithm: $($HP.LoadBalancerType), PoolType: $($HP.HostPoolType)" `
                -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/host-pool-load-balancing'))

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
            } catch { }
        }
    } catch {
        Write-Status "  Error discovering host pools: $($_.Exception.Message)" -Level 'ERROR'
        $Discovery.Errors += "Host pool discovery failed: $($_.Exception.Message)"
    }

    # ─── SESSION HOSTS ────────────────────────────────────────────────────
    Write-Status "Session Hosts" -Level 'SECTION'
    try {
        foreach ($HP in $HostPools) {
            $RG = ($HP.Id -split '/')[4]
            $SessionHosts = @(Get-AzWvdSessionHost -ResourceGroupName $RG -HostPoolName $HP.Name -ErrorAction Stop)
            Write-Status "  $($HP.Name): $($SessionHosts.Count) session host(s)" -Level 'INFO'

            foreach ($SH in $SessionHosts) {
                $VMName = ($SH.ResourceId -split '/')[-1]
                $VMRG   = ($SH.ResourceId -split '/')[4]

                # Get VM details
                $VMObj = $null
                try {
                    $VMObj = Get-AzVM -ResourceGroupName $VMRG -Name $VMName -Status -ErrorAction Stop
                } catch {
                    Write-Status "    Could not get VM details for $VMName : $($_.Exception.Message)" -Level 'WARN'
                }

                $SHObj = [PSCustomObject]@{
                    HostPoolName        = $HP.Name
                    Name                = $SH.Name
                    ResourceId          = $SH.ResourceId
                    Status              = $SH.Status
                    AllowNewSession     = $SH.AllowNewSession
                    Sessions            = $SH.Session
                    AgentVersion        = $SH.AgentVersion
                    LastHeartBeat       = $SH.LastHeartBeat
                    OSVersion           = $SH.OSVersion
                    UpdateState         = $SH.UpdateState
                    VMSize              = if ($VMObj) { $VMObj.HardwareProfile.VmSize } else { 'Unknown' }
                    OSDiskType          = if ($VMObj) { $VMObj.StorageProfile.OsDisk.ManagedDisk.StorageAccountType } else { 'Unknown' }
                    SecurityProfile     = if ($VMObj -and $VMObj.SecurityProfile) {
                        [PSCustomObject]@{
                            SecurityType = $VMObj.SecurityProfile.SecurityType
                            SecureBoot   = $VMObj.SecurityProfile.UefiSettings.SecureBootEnabled
                            VTpm         = $VMObj.SecurityProfile.UefiSettings.VTpmEnabled
                        }
                    } else { $null }
                    PowerState          = if ($VMObj) { ($VMObj.Statuses | Where-Object Code -like 'PowerState/*').DisplayStatus } else { 'Unknown' }
                    AcceleratedNetworking = $null
                    AvailabilityZone    = if ($VMObj) { $VMObj.Zones } else { $null }
                    ImageReference      = if ($VMObj) { $VMObj.StorageProfile.ImageReference } else { $null }
                }

                # Check NIC for accelerated networking
                if ($VMObj -and $VMObj.NetworkProfile.NetworkInterfaces.Count -gt 0) {
                    try {
                        $NicId = $VMObj.NetworkProfile.NetworkInterfaces[0].Id
                        $Nic = Get-AzNetworkInterface -ResourceId $NicId -ErrorAction Stop
                        $SHObj.AcceleratedNetworking = $Nic.EnableAcceleratedNetworking
                    } catch { }
                }

                $Discovery.Inventory.SessionHosts += $SHObj

                # ─── CHECK: Drain mode ───
                [void]$AllChecks.Add((New-CheckResult -Id "SH-DRAIN-$VMName" `
                    -Category 'Session Hosts' -Name 'Drain Mode' `
                    -Description 'Session host should allow new sessions unless under maintenance' `
                    -Status $(if ($SH.AllowNewSession) { 'Pass' } else { 'Warning' }) `
                    -Severity 'Low' `
                    -Details "AllowNewSession: $($SH.AllowNewSession), Status: $($SH.Status)" `
                    -Evidence @{ VM = $VMName; AllowNewSession = $SH.AllowNewSession }))

                # ─── CHECK: Trusted Launch ───
                if ($VMObj -and $VMObj.SecurityProfile) {
                    $SecureBoot = $VMObj.SecurityProfile.UefiSettings.SecureBootEnabled
                    $VTpm       = $VMObj.SecurityProfile.UefiSettings.VTpmEnabled
                    $TrustedLaunch = $VMObj.SecurityProfile.SecurityType -eq 'TrustedLaunch'
                    [void]$AllChecks.Add((New-CheckResult -Id "SEC-TL-$VMName" `
                        -Category 'Security & IAM' -Name 'Trusted Launch' `
                        -Description 'Session hosts should use Trusted Launch with Secure Boot and vTPM enabled' `
                        -Status $(if ($TrustedLaunch -and $SecureBoot -and $VTpm) { 'Pass' }
                                  elseif ($TrustedLaunch) { 'Warning' } else { 'Fail' }) `
                        -Severity 'High' `
                        -Details "SecurityType: $($VMObj.SecurityProfile.SecurityType), SecureBoot: $SecureBoot, vTPM: $VTpm" `
                        -Recommendation 'Enable Trusted Launch with Secure Boot and vTPM for enhanced boot integrity.' `
                        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-machines/trusted-launch' `
                        -Evidence @{ VM = $VMName; SecurityType = $VMObj.SecurityProfile.SecurityType }))
                } else {
                    [void]$AllChecks.Add((New-CheckResult -Id "SEC-TL-$VMName" `
                        -Category 'Security & IAM' -Name 'Trusted Launch' `
                        -Description 'Session hosts should use Trusted Launch' `
                        -Status 'Fail' -Severity 'High' `
                        -Details 'No security profile detected — VM is likely using Standard security type.' `
                        -Recommendation 'Redeploy session hosts with Trusted Launch security type.' `
                        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-machines/trusted-launch'))
                }

                # ─── CHECK: Secure Boot (SH-022) ───
                if ($VMObj -and $VMObj.SecurityProfile) {
                    $SBEnabled = $VMObj.SecurityProfile.UefiSettings.SecureBootEnabled -eq $true
                    [void]$AllChecks.Add((New-CheckResult -Id "SH-SECBOOT-$VMName" `
                        -Category 'Session Hosts' -Name 'Secure Boot Enabled' `
                        -Description 'Trusted Launch VMs should have Secure Boot enabled to protect against boot-level malware' `
                        -Status $(if ($SBEnabled) { 'Pass' } else { 'Fail' }) `
                        -Severity 'High' `
                        -Details "SecureBoot: $SBEnabled" `
                        -Recommendation 'Enable Secure Boot in the VM security profile to protect the boot chain.' `
                        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-machines/trusted-launch#secure-boot' `
                        -Evidence @{ VM = $VMName; SecureBoot = $SBEnabled }))
                }

                # ─── CHECK: vTPM (SH-023) ───
                if ($VMObj -and $VMObj.SecurityProfile) {
                    $VTpmOn = $VMObj.SecurityProfile.UefiSettings.VTpmEnabled -eq $true
                    [void]$AllChecks.Add((New-CheckResult -Id "SH-VTPM-$VMName" `
                        -Category 'Session Hosts' -Name 'vTPM Enabled' `
                        -Description 'Trusted Launch VMs should have vTPM enabled for measured boot and key protection' `
                        -Status $(if ($VTpmOn) { 'Pass' } else { 'Fail' }) `
                        -Severity 'High' `
                        -Details "vTPM: $VTpmOn" `
                        -Recommendation 'Enable vTPM to support BitLocker, measured boot, and Windows Hello for Business.' `
                        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-machines/trusted-launch#vtpm' `
                        -Evidence @{ VM = $VMName; VTpm = $VTpmOn }))
                }

                # ─── CHECK: OS Disk Encryption — ADE or host-based (SEC-021) ───
                if ($VMObj) {
                    $HasADE = $false
                    $DiskEncType = 'None'
                    if ($SHObj.Extensions) {
                        $HasADE = $SHObj.Extensions | Where-Object { $_ -match 'AzureDiskEncryption' }
                    }
                    $OsDisk = $VMObj.StorageProfile.OsDisk
                    if ($OsDisk.ManagedDisk.SecurityProfile.DiskEncryptionSet) {
                        $DiskEncType = 'DiskEncryptionSet'
                    } elseif ($VMObj.SecurityProfile.EncryptionAtHost) {
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
                if ($null -ne $SHObj.AcceleratedNetworking) {
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
                if ($SHObj.OSDiskType -and $SHObj.OSDiskType -ne 'Unknown') {
                    [void]$AllChecks.Add((New-CheckResult -Id "SEC-DISK-$VMName" `
                        -Category 'Security & IAM' -Name 'Managed Disk Encryption' `
                        -Description 'OS disk should use managed encryption' `
                        -Status 'Pass' -Severity 'Medium' `
                        -Details "DiskType: $($SHObj.OSDiskType) (Azure managed disks are encrypted by default with platform-managed keys)" `
                        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-machines/disk-encryption-overview'))
                }

                # ─── CHECK: Availability Zones ───
                if ($VMObj) {
                    [void]$AllChecks.Add((New-CheckResult -Id "BCDR-AZ-$VMName" `
                        -Category 'BCDR' -Name 'Availability Zone Deployment' `
                        -Description 'Session hosts should be spread across availability zones for resilience' `
                        -Status $(if ($SHObj.AvailabilityZone -and $SHObj.AvailabilityZone.Count -gt 0) { 'Pass' } else { 'Warning' }) `
                        -Severity 'Medium' `
                        -Details "Zones: $(if ($SHObj.AvailabilityZone) { $SHObj.AvailabilityZone -join ',' } else { 'None' })" `
                        -Recommendation 'Deploy session hosts across multiple availability zones for high availability.' `
                        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/azure-virtual-desktop-fault-domain-mode'))
                }

                # ─── CHECK: OS Disk SSD ───
                if ($SHObj.OSDiskType -and $SHObj.OSDiskType -ne 'Unknown') {
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
                    $DaysSinceHB = ((Get-Date) - [DateTime]$SH.LastHeartBeat).Days
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
                        } catch { }
                    }
                }

                # ─── CHECK: VM Sizing (B-series flagging) ───
                if ($SHObj.VMSize -and $SHObj.VMSize -ne 'Unknown') {
                    $IsBSeries = $SHObj.VMSize -match '^Standard_B'
                    if ($IsBSeries) {
                        [void]$AllChecks.Add((New-CheckResult -Id "SH-BSERIES-$VMName" `
                            -Category 'Session Hosts' -Name 'VM Sizing (B-series)' `
                            -Description 'B-series VMs are burstable and may cause inconsistent performance for pooled desktops' `
                            -Status 'Warning' -Severity 'Medium' `
                            -Details "VMSize: $($SHObj.VMSize)" `
                            -Recommendation 'Use D-series or E-series for production pooled host pools.' `
                            -Reference 'https://learn.microsoft.com/en-us/azure/virtual-machines/sizes'))
                    }
                }

                # ─── CHECK: Entra Join Type (from VM extensions) ───
                if ($VMObj) {
                    $HasAADExt = $VMObj.Extensions | Where-Object { $_.Type -eq 'AADLoginForWindows' -or $_.VirtualMachineExtensionType -eq 'AADLoginForWindows' }
                    $HasDJExt = $VMObj.Extensions | Where-Object { $_.Type -eq 'JsonADDomainExtension' -or $_.VirtualMachineExtensionType -eq 'JsonADDomainExtension' }
                    $JoinType = if ($HasAADExt -and $HasDJExt) { 'Hybrid' } elseif ($HasAADExt) { 'Entra ID' } elseif ($HasDJExt) { 'AD DS' } else { 'Unknown' }
                    [void]$AllChecks.Add((New-CheckResult -Id "IAM-JOIN-$VMName" `
                        -Category 'Identity & Access' -Name 'Entra ID Join Type' `
                        -Description 'Session hosts should use Entra ID or Hybrid join' `
                        -Status $(if ($JoinType -ne 'Unknown') { 'Pass' } else { 'Warning' }) `
                        -Severity 'High' `
                        -Details "JoinType: $JoinType" `
                        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/prerequisites#identity'))
                }

                # ─── CHECK: Azure Monitor Agent (AMA) ───
                if ($VMObj) {
                    $HasAMA = $VMObj.Extensions | Where-Object {
                        $_.VirtualMachineExtensionType -eq 'AzureMonitorWindowsAgent' -or
                        $_.Type -eq 'AzureMonitorWindowsAgent'
                    }
                    [void]$AllChecks.Add((New-CheckResult -Id "MON-AMA-$VMName" `
                        -Category 'Monitoring' -Name 'Azure Monitor Agent Installed' `
                        -Description 'AMA should be installed on session hosts for AVD Insights telemetry' `
                        -Status $(if ($HasAMA) { 'Pass' } else { 'Fail' }) `
                        -Severity 'Medium' `
                        -Details "AMA Extension: $(if ($HasAMA) { 'Installed' } else { 'Not found' })" `
                        -Recommendation 'Install Azure Monitor Agent for AVD Insights and performance monitoring.' `
                        -Reference 'https://learn.microsoft.com/en-us/azure/azure-monitor/agents/agents-overview'))

                    # ─── CHECK: Endpoint Protection (MDE) ───
                    $HasMDE = $VMObj.Extensions | Where-Object {
                        $_.VirtualMachineExtensionType -in @('MDE.Windows','MicrosoftMonitoringAgent') -or
                        $_.Type -in @('MDE.Windows','MicrosoftMonitoringAgent')
                    }
                    [void]$AllChecks.Add((New-CheckResult -Id "SEC-MDE-$VMName" `
                        -Category 'Security' -Name 'Endpoint Protection (MDE)' `
                        -Description 'Microsoft Defender for Endpoint should be deployed on session hosts' `
                        -Status $(if ($HasMDE) { 'Pass' } else { 'Warning' }) `
                        -Severity 'High' `
                        -Details "MDE Extension: $(if ($HasMDE) { 'Installed' } else { 'Not found (may be deployed via Intune/GPO)' })" `
                        -Recommendation 'Deploy Microsoft Defender for Endpoint via VM extension, Intune, or Defender for Cloud auto-provisioning.' `
                        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/security-recommendations'))

                    # ─── CHECK: Ephemeral OS Disk (Pooled VMs) ───
                    if ($HP.HostPoolType -eq 'Pooled') {
                        $IsEphemeral = $null -ne $VMObj.StorageProfile.OsDisk.DiffDiskSettings
                        [void]$AllChecks.Add((New-CheckResult -Id "SH-EPHEMERAL-$VMName" `
                            -Category 'Session Hosts' -Name 'Ephemeral OS Disk' `
                            -Description 'Pooled session hosts should use ephemeral OS disks for faster reimage and lower cost' `
                            -Status $(if ($IsEphemeral) { 'Pass' } else { 'Warning' }) `
                            -Severity 'Medium' `
                            -Details "EphemeralDisk: $(if ($IsEphemeral) { 'Yes' } else { 'No — uses persistent managed disk' })" `
                            -Recommendation 'Use ephemeral OS disks for pooled host pools to eliminate storage costs and improve reimage speed.' `
                            -Reference 'https://learn.microsoft.com/en-us/azure/virtual-machines/ephemeral-os-disks'))
                    }

                    # ─── CHECK: Disk Type Cost Optimization ───
                    if ($SHObj.OSDiskType -and $SHObj.OSDiskType -ne 'Unknown') {
                        $IsPremium = $SHObj.OSDiskType -match 'Premium'
                        if ($IsPremium -and $HP.HostPoolType -eq 'Pooled') {
                            [void]$AllChecks.Add((New-CheckResult -Id "GOV-DISKSKU-$VMName" `
                                -Category 'Governance & Cost' -Name 'Disk Type Cost Optimization' `
                                -Description 'Premium SSD on pooled hosts may be unnecessary cost — Standard SSD is often sufficient' `
                                -Status 'Warning' -Severity 'Low' `
                                -Details "DiskType: $($SHObj.OSDiskType), PoolType: $($HP.HostPoolType)" `
                                -Recommendation 'Evaluate Standard SSD for pooled hosts to reduce storage costs. Premium is typically only needed for heavy I/O workloads.' `
                                -Reference 'https://learn.microsoft.com/en-us/azure/virtual-machines/disks-types'))
                        }
                    }

                    # ─── CHECK: Guest Attestation (Trusted Launch integrity monitoring) ───
                    $HasGuestAttest = $VMObj.Extensions | Where-Object {
                        $_.VirtualMachineExtensionType -eq 'GuestAttestation' -or
                        $_.Type -eq 'GuestAttestation'
                    }
                    if ($VMObj.SecurityProfile -and $VMObj.SecurityProfile.SecurityType -eq 'TrustedLaunch' -and -not $HasGuestAttest) {
                        [void]$AllChecks.Add((New-CheckResult -Id "SEC-ATTEST-$VMName" `
                            -Category 'Security' -Name 'Guest Attestation Extension' `
                            -Description 'Trusted Launch VMs should have Guest Attestation extension for integrity monitoring' `
                            -Status 'Warning' -Severity 'Medium' `
                            -Details "TrustedLaunch: Yes, GuestAttestation: Not installed" `
                            -Recommendation 'Install the Guest Attestation extension to enable boot integrity monitoring via Defender for Cloud.' `
                            -Reference 'https://learn.microsoft.com/en-us/azure/virtual-machines/trusted-launch#microsoft-defender-for-cloud-integration'))
                    }

                    # Collect extension inventory for evidence
                    $ExtList = @($VMObj.Extensions | ForEach-Object {
                        $ExtType = if ($_.VirtualMachineExtensionType) { $_.VirtualMachineExtensionType } else { $_.Type }
                        $ExtType
                    } | Where-Object { $_ })
                    $SHObj | Add-Member -NotePropertyName 'Extensions' -NotePropertyValue $ExtList -Force
                }

                # ─── CHECK: Agent Version Currency ───
                if ($SH.AgentVersion) {
                    $MinRecommendedAgent = [Version]'1.0.8431.0'
                    try {
                        $CurrentVer = [Version]$SH.AgentVersion
                        $IsAgentCurrent = $CurrentVer -ge $MinRecommendedAgent
                    } catch {
                        $IsAgentCurrent = $false
                    }
                    [void]$AllChecks.Add((New-CheckResult -Id "OPS-AGENT-$VMName" `
                        -Category 'Operations' -Name 'Agent Version Currency' `
                        -Description 'AVD Agent should be at current recommended version' `
                        -Status $(if ($IsAgentCurrent) { 'Pass' } else { 'Warning' }) `
                        -Severity 'Medium' `
                        -Details "AgentVersion: $($SH.AgentVersion), MinRecommended: $MinRecommendedAgent" `
                        -Recommendation 'AVD agent updates automatically when VM is running. Ensure VMs are powered on periodically.' `
                        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/agent-overview'))
                }

                # ─── CHECK: Power State ───
                if ($SHObj.PowerState -and $SHObj.PowerState -ne 'Unknown') {
                    $IsRunning = $SHObj.PowerState -eq 'VM running'
                    $IsDeallocated = $SHObj.PowerState -match 'deallocated'
                    [void]$AllChecks.Add((New-CheckResult -Id "SH-POWER-$VMName" `
                        -Category 'Session Hosts' -Name 'Power State' `
                        -Description 'Session host power state — deallocated VMs cannot serve users and may have stale agents' `
                        -Status $(if ($IsRunning) { 'Pass' } elseif ($IsDeallocated) { 'Warning' } else { 'Warning' }) `
                        -Severity 'Low' `
                        -Details "PowerState: $($SHObj.PowerState)" `
                        -Evidence @{ VM = $VMName; PowerState = $SHObj.PowerState }))
                }

                # ─── CHECK: Session Host Status ───
                if ($SH.Status) {
                    $SHStatusOK = $SH.Status -eq 'Available'
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
    } catch {
        Write-Status "  Error discovering session hosts: $($_.Exception.Message)" -Level 'ERROR'
        $Discovery.Errors += "Session host discovery failed: $($_.Exception.Message)"
    }

    # ─── APPLICATION GROUPS ───────────────────────────────────────────────
    Write-Status "Application Groups" -Level 'SECTION'
    try {
        $AppGroups = @(Get-AzWvdApplicationGroup -ErrorAction Stop)
        Write-Status "  Found $($AppGroups.Count) app group(s)" -Level 'SUCCESS'

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

            # ─── CHECK: App group type validation ───
            [void]$AllChecks.Add((New-CheckResult -Id "APP-CFG-$($AG.Name)" `
                -Category 'Application Delivery' -Name 'App Group Configuration' `
                -Description 'Application groups should be configured appropriately' `
                -Status 'Pass' -Severity 'Medium' `
                -Details "Type: $($AG.ApplicationGroupType), HostPool: $(($AG.HostPoolArmPath -split '/')[-1])" `
                -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/manage-app-groups'))
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

        # ─── CHECK: Scaling plan coverage ───
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
        }

        # CHECK: Scaling plan active (not just assigned but has schedules)
        foreach ($SP in $ScalingPlans) {
            $HasSchedules = $SP.Schedule -and $SP.Schedule.Count -gt 0
            [void]$AllChecks.Add((New-CheckResult -Id "GOV-SPACTIVE-$($SP.Name)" `
                -Category 'Governance & Cost' -Name 'Scaling Plan Active' `
                -Description 'Scaling plans should be enabled with active schedules' `
                -Status $(if ($HasSchedules) { 'Pass' } else { 'Warning' }) `
                -Severity 'High' `
                -Details "Schedules: $(if ($HasSchedules) { $SP.Schedule.Count } else { 0 })" `
                -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/autoscale-scaling-plan'))
        }


            [void]$AllChecks.Add((New-CheckResult -Id "BCDR-SPSCHED-$($SP.Name)" `
                -Category 'BCDR' -Name 'Scaling Plan Schedule Defined' `
                -Description 'Scaling plan should have schedules with peak and off-peak ramp configurations' `
                -Status $(if ($HasPeakOffPeak) { 'Pass' } else { 'Fail' }) `
                -Severity 'Medium' `
                -Details "Schedules: $(if ($Schedules) { $Schedules.Count } else { 0 }), Days covered: $(if ($Schedules) { ($Schedules | ForEach-Object { $_.DaysOfWeek } | Select-Object -Unique).Count } else { 0 })" `
                -Recommendation 'Configure scaling plan schedules with ramp-up, peak, ramp-down, and off-peak phases.' `
                -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/autoscale-scaling-plan' `
                -Evidence @{ ScalingPlan = $SP.Name; ScheduleCount = $(if ($Schedules) { $Schedules.Count } else { 0 }) }))
        }
    } catch {
        Write-Status "  Error: $($_.Exception.Message)" -Level 'ERROR'
        $Discovery.Errors += "Scaling plan discovery failed: $($_.Exception.Message)"
    }

    # ─── MULTI-REGION CHECK ───
    $HPLocations = @($HostPools | ForEach-Object { $_.Location } | Sort-Object -Unique)
    if ($HPLocations.Count -gt 1) {
        [void]$AllChecks.Add((New-CheckResult -Id "BCDR-MULTIREGION" `
            -Category 'BCDR' -Name 'Multi-Region Host Pool' `
            -Description 'Host pools deployed in multiple regions for disaster recovery' `
            -Status 'Pass' -Severity 'Medium' `
            -Details "Regions: $($HPLocations -join ', ')" `
            -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/disaster-recovery'))
    } else {
        [void]$AllChecks.Add((New-CheckResult -Id "BCDR-MULTIREGION" `
            -Category 'BCDR' -Name 'Multi-Region Host Pool' `
            -Description 'Host pools concentrated in single region — DR risk' `
            -Status 'Warning' -Severity 'Medium' `
            -Details "Regions: $(if ($HPLocations) { $HPLocations -join ', ' } else { 'None' })"))
    }

    # ─── NETWORKING ───────────────────────────────────────────────────────
    Write-Status "Networking" -Level 'SECTION'
    try {
        # Collect unique VNets from session host NICs
        $DiscoveredVNetIds = @{}
        foreach ($SH in $Discovery.Inventory.SessionHosts) {
            if ($SH.ResourceId) {
                try {
                    $VMRG   = ($SH.ResourceId -split '/')[4]
                    $VMName = ($SH.ResourceId -split '/')[-1]
                    $VM = Get-AzVM -ResourceGroupName $VMRG -Name $VMName -ErrorAction SilentlyContinue
                    if ($VM -and $VM.NetworkProfile.NetworkInterfaces.Count -gt 0) {
                        $NicId = $VM.NetworkProfile.NetworkInterfaces[0].Id
                        $Nic = Get-AzNetworkInterface -ResourceId $NicId -ErrorAction SilentlyContinue
                        if ($Nic -and $Nic.IpConfigurations[0].Subnet.Id) {
                            $SubnetId = $Nic.IpConfigurations[0].Subnet.Id
                            $VNetId = ($SubnetId -split '/subnets/')[0]
                            if (-not $DiscoveredVNetIds.ContainsKey($VNetId)) {
                                $DiscoveredVNetIds[$VNetId] = $true
                            }

                            # CHECK: Public IP on session host
                            $HasPublicIP = $null -ne $Nic.IpConfigurations[0].PublicIpAddress
                            [void]$AllChecks.Add((New-CheckResult -Id "NET-PIP-$VMName" `
                                -Category 'Networking' -Name 'No Public IP on Session Host' `
                                -Description 'Session hosts should not have public IP addresses' `
                                -Status $(if ($HasPublicIP) { 'Fail' } else { 'Pass' }) `
                                -Severity 'Critical' `
                                -Details "PublicIP: $(if ($HasPublicIP) { 'ASSIGNED' } else { 'None' })" `
                                -Recommendation 'Remove public IPs from session hosts. Use Azure Bastion or JIT for management access.' `
                                -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/security-guide'))
                        }
                    }
                } catch { }
            }
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
                    Location      = $VNet.Location
                }
                $Discovery.Inventory.VNets += $VNetObj

                # CHECK: Custom DNS
                $HasCustomDns = $VNet.DhcpOptions -and $VNet.DhcpOptions.DnsServers -and $VNet.DhcpOptions.DnsServers.Count -gt 0
                [void]$AllChecks.Add((New-CheckResult -Id "NET-DNS-$VNetName" `
                    -Category 'Networking' -Name 'Custom DNS Configuration' `
                    -Description 'VNets with AD-joined session hosts should use custom DNS pointing to domain controllers' `
                    -Status $(if ($HasCustomDns) { 'Pass' } else { 'Warning' }) `
                    -Severity 'Medium' `
                    -Details "DNS: $(if ($HasCustomDns) { $VNet.DhcpOptions.DnsServers -join ', ' } else { 'Azure Default' })" `
                    -Recommendation 'Configure custom DNS servers pointing to domain controllers for AD-joined environments.'))

                # CHECK: NSG on AVD subnets
                $SubnetsWithoutNSG = @($VNet.Subnets | Where-Object { -not $_.NetworkSecurityGroup })
                $TotalSubnets = $VNet.Subnets.Count
                $HasNSGCoverage = $SubnetsWithoutNSG.Count -eq 0 -and $TotalSubnets -gt 0
                [void]$AllChecks.Add((New-CheckResult -Id "NET-NSG-$VNetName" `
                    -Category 'Networking' -Name 'NSG on AVD Subnets' `
                    -Description 'All AVD subnets should have Network Security Groups applied' `
                    -Status $(if ($HasNSGCoverage) { 'Pass' } elseif ($TotalSubnets -eq 0) { 'N/A' } else { 'Warning' }) `
                    -Severity 'High' `
                    -Details "Subnets: $TotalSubnets total, $($SubnetsWithoutNSG.Count) without NSG" `
                    -Recommendation 'Apply NSGs to all AVD subnets for network traffic filtering.' `
                    -Reference 'https://learn.microsoft.com/en-us/azure/virtual-network/network-security-groups-overview'))

                # CHECK: Route tables on subnets
                $SubnetsWithRT = @($VNet.Subnets | Where-Object { $_.RouteTable })
                [void]$AllChecks.Add((New-CheckResult -Id "NET-UDR-$VNetName" `
                    -Category 'Networking' -Name 'Route Table on AVD Subnets' `
                    -Description 'UDR should force traffic through firewall/NVA for inspection' `
                    -Status $(if ($SubnetsWithRT.Count -gt 0) { 'Pass' } else { 'Warning' }) `
                    -Severity 'Medium' `
                    -Details "SubnetsWithRouteTable: $($SubnetsWithRT.Count)/$TotalSubnets" `
                    -Reference 'https://learn.microsoft.com/en-us/azure/well-architected/azure-virtual-desktop/networking'))

                # CHECK: NAT Gateway
                $SubnetsWithNAT = @($VNet.Subnets | Where-Object { $_.NatGateway })
                [void]$AllChecks.Add((New-CheckResult -Id "NET-NATGW-$VNetName" `
                    -Category 'Networking' -Name 'NAT Gateway for Outbound' `
                    -Description 'Private subnets should use NAT Gateway for explicit outbound connectivity' `
                    -Status $(if ($SubnetsWithNAT.Count -gt 0) { 'Pass' } else { 'Warning' }) `
                    -Severity 'Medium' `
                    -Details "SubnetsWithNATGW: $($SubnetsWithNAT.Count)/$TotalSubnets"))

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
                        Location      = $FallbackVNet.Location
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
                                    SourceAddressPrefix    = $_.SourceAddressPrefix
                                    DestinationAddressPrefix = $_.DestinationAddressPrefix
                                }
                            })
                        }

                        # CHECK: Port 3389 exposure
                        $RdpRules = @($NSG.SecurityRules | Where-Object {
                            $_.Direction -eq 'Inbound' -and $_.Access -eq 'Allow' -and
                            ($_.DestinationPortRange -contains '3389' -or $_.DestinationPortRange -contains '*')
                        })
                        $ExposedToInternet = @($RdpRules | Where-Object {
                            $_.SourceAddressPrefix -eq '*' -or $_.SourceAddressPrefix -eq 'Internet' -or $_.SourceAddressPrefix -eq '0.0.0.0/0'
                        })
                        [void]$AllChecks.Add((New-CheckResult -Id "NET-RDP-$NSGName" `
                            -Category 'Networking' -Name 'RDP Port 3389 Not Internet-Exposed' `
                            -Description 'Port 3389 should not be open to the internet on AVD subnets' `
                            -Status $(if ($ExposedToInternet.Count -gt 0) { 'Fail' } else { 'Pass' }) `
                            -Severity 'Critical' `
                            -Details "Internet-facing RDP rules: $($ExposedToInternet.Count)" `
                            -Recommendation 'Block inbound RDP from internet. Use Azure Bastion or JIT access for administration.' `
                            -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/security-guide' `
                            -Evidence @{ NSG = $NSGName; ExposedRules = $ExposedToInternet.Name }))

                        # CHECK: AVD required outbound connectivity
                        $OutboundRules = @($NSG.SecurityRules | Where-Object { $_.Direction -eq 'Outbound' })
                        $HasDenyAllOut = @($OutboundRules | Where-Object {
                            $_.Access -eq 'Deny' -and $_.DestinationAddressPrefix -eq '*' -and $_.DestinationPortRange -eq '*'
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
                        $SshRules = @($NSG.SecurityRules | Where-Object {
                            $_.Direction -eq 'Inbound' -and $_.Access -eq 'Allow' -and
                            ($_.DestinationPortRange -contains '22' -or $_.DestinationPortRange -contains '*')
                        })
                        $SshExposed = @($SshRules | Where-Object {
                            $_.SourceAddressPrefix -eq '*' -or $_.SourceAddressPrefix -eq 'Internet' -or $_.SourceAddressPrefix -eq '0.0.0.0/0'
                        })
                        if ($SshExposed.Count -gt 0) {
                            [void]$AllChecks.Add((New-CheckResult -Id "NET-SSH-$NSGName" `
                                -Category 'Networking' -Name 'SSH Port 22 Not Internet-Exposed' `
                                -Description 'Port 22 should not be open to the internet' `
                                -Status 'Fail' -Severity 'High' `
                                -Details "Internet-facing SSH rules: $($SshExposed.Count)" `
                                -Recommendation 'Block inbound SSH from internet. Use Azure Bastion for management.' `
                                -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/security-guide'))
                        }
                    } catch { }
                }
            }
        }
    } catch {
        Write-Status "  Error discovering network: $($_.Exception.Message)" -Level 'ERROR'
        $Discovery.Errors += "Network discovery failed: $($_.Exception.Message)"
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
                    if ($Peer.RemoteVNetId -in $FwVNetIds) { $PeeredFw = $true }
                }
            }
        }
        $DirectFw = @($AzFirewalls | Where-Object { ($_.IpConfigurations[0].Subnet.Id -split '/subnets/')[0] -in $AvdVNetIds }).Count -gt 0
        $HasFirewall = $PeeredFw -or $DirectFw -or ($AzFirewalls.Count -gt 0)
        Write-Status "  Azure Firewalls: $($AzFirewalls.Count), Peered to AVD: $PeeredFw" -Level $(if ($HasFirewall) { 'SUCCESS' } else { 'WARN' })
        [void]$AllChecks.Add((New-CheckResult -Id "NET-HUBFW" `
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
        [void]$AllChecks.Add((New-CheckResult -Id "NET-HUBGW" `
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
                    $EnabledCategories = @($DiagSettings | ForEach-Object { $_.Log } | Where-Object { $_.Enabled } | ForEach-Object { $_.Category }) | Sort-Object -Unique
                    $MissingCategories = @($RecommendedCategories | Where-Object { $_ -notin $EnabledCategories })
                    [void]$AllChecks.Add((New-CheckResult -Id "MON-DIAGCAT-$($HP.Name)" `
                        -Category 'Monitoring' -Name 'Diagnostic Log Categories' `
                        -Description 'All recommended AVD diagnostic log categories should be enabled' `
                        -Status $(if ($MissingCategories.Count -eq 0) { 'Pass' } else { 'Warning' }) `
                        -Severity 'Medium' `
                        -Details "Enabled: $($EnabledCategories -join ', '). Missing: $(if ($MissingCategories.Count -gt 0) { $MissingCategories -join ', ' } else { 'None' })" `
                        -Recommendation "Enable missing categories: $($MissingCategories -join ', ')" `
                        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/diagnostics-log-analytics' `
                        -Evidence @{ Enabled = $EnabledCategories; Missing = $MissingCategories }))
                }
            } catch { }
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
        foreach ($HP in $HostPools) {
            $RG = ($HP.Id -split '/')[4]
            try {
                $Assignments = @(Get-AzRoleAssignment -Scope $HP.Id -ErrorAction Stop)
                $HasCustomRoles = @($Assignments | Where-Object { $_.RoleDefinitionName -like '*Virtual Desktop*' }).Count -gt 0

                [void]$AllChecks.Add((New-CheckResult -Id "SEC-RBAC-$($HP.Name)" `
                    -Category 'Security & IAM' -Name 'AVD RBAC Roles Used' `
                    -Description 'Use built-in AVD roles for least-privilege access' `
                    -Status $(if ($HasCustomRoles) { 'Pass' } else { 'Warning' }) `
                    -Severity 'Medium' `
                    -Details "TotalAssignments: $($Assignments.Count), AVDRoles: $HasCustomRoles" `
                    -Recommendation 'Use built-in AVD roles (Desktop Virtualization Contributor, User, etc.) instead of broad roles like Contributor.' `
                    -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/rbac'))
            } catch { }
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
            } catch { }
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
            } catch { }
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
        $DefenderVMs = Get-AzSecurityPricing -Name 'VirtualMachines' -ErrorAction SilentlyContinue
        if ($DefenderVMs) {
            [void]$AllChecks.Add((New-CheckResult -Id "MON-DEFENDER" `
                -Category 'Monitoring' -Name 'Defender for Cloud Enabled' `
                -Description 'Microsoft Defender for Cloud should be enabled for VMs' `
                -Status $(if ($DefenderVMs.PricingTier -eq 'Standard') { 'Pass' } else { 'Warning' }) `
                -Severity 'High' `
                -Details "PricingTier: $($DefenderVMs.PricingTier)" `
                -Reference 'https://learn.microsoft.com/en-us/azure/defender-for-cloud/enable-enhanced-security'))
        }
    } catch { }

    # ─── STORAGE ──────────────────────────────────────────────────────────
    Write-Status "Storage (FSLogix)" -Level 'SECTION'
    try {
        $StorageAccounts = @(Get-AzStorageAccount -ErrorAction Stop)
        # Look for storage accounts that might be used for FSLogix (heuristic: has 'profiles' or 'fslogix' in name or file shares)
        foreach ($SA in $StorageAccounts) {
            $IsFSLogix = $SA.StorageAccountName -match 'fslogix|profile|avd'
            $HasPrivateEndpoint = $SA.PrivateEndpointConnections -and $SA.PrivateEndpointConnections.Count -gt 0

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
                Replication       = $SA.Sku.Name  # LRS, ZRS, GRS, etc.
                LargeFileShares   = $SA.LargeFileSharesState
            }
            $Discovery.Inventory.StorageAccounts += $SAObj

            if ($IsFSLogix) {
                # CHECK: Private endpoint on FSLogix storage
                [void]$AllChecks.Add((New-CheckResult -Id "PROF-PE-$($SA.StorageAccountName)" `
                    -Category 'FSLogix & Profiles' -Name 'Private Endpoint on Profile Storage' `
                    -Description 'FSLogix storage should use private endpoints for security' `
                    -Status $(if ($HasPrivateEndpoint) { 'Pass' } else { 'Warning' }) `
                    -Severity 'High' `
                    -Details "PrivateEndpoint: $HasPrivateEndpoint" `
                    -Recommendation 'Configure private endpoints for FSLogix profile storage to keep traffic on the Microsoft network.' `
                    -Reference 'https://learn.microsoft.com/en-us/azure/storage/files/storage-files-networking-overview'))

                # CHECK: HTTPS only
                [void]$AllChecks.Add((New-CheckResult -Id "PROF-HTTPS-$($SA.StorageAccountName)" `
                    -Category 'FSLogix & Profiles' -Name 'HTTPS Only Enabled' `
                    -Description 'Storage account should enforce HTTPS-only traffic' `
                    -Status $(if ($SA.EnableHttpsTrafficOnly) { 'Pass' } else { 'Fail' }) `
                    -Severity 'High' `
                    -Details "HttpsOnly: $($SA.EnableHttpsTrafficOnly)"))

                # CHECK: TLS version
                [void]$AllChecks.Add((New-CheckResult -Id "PROF-TLS-$($SA.StorageAccountName)" `
                    -Category 'FSLogix & Profiles' -Name 'Minimum TLS 1.2' `
                    -Description 'Storage account should enforce TLS 1.2 minimum' `
                    -Status $(if ($SA.MinimumTlsVersion -ge 'TLS1_2') { 'Pass' } else { 'Fail' }) `
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

                    # CHECK: SMB security settings (while we have file service props)
                    if ($FSP -and $FSP.ProtocolSetting -and $FSP.ProtocolSetting.Smb) {
                        $Smb = $FSP.ProtocolSetting.Smb
                        # SMB versions
                        $Versions = $Smb.Versions
                        $HasSMB21 = $Versions -match 'SMB2\.1'
                        [void]$AllChecks.Add((New-CheckResult -Id "PROF-SMBVER-$($SA.StorageAccountName)" `
                            -Category 'FSLogix & Profiles' -Name 'SMB Minimum Version' `
                            -Description 'SMB 2.1 should be disabled — require SMB 3.0+' `
                            -Status $(if ($HasSMB21) { 'Warning' } else { 'Pass' }) `
                            -Severity 'High' `
                            -Details "Versions: $Versions" `
                            -Reference 'https://learn.microsoft.com/en-us/azure/storage/files/files-smb-protocol#smb-security-settings'))

                        # Kerberos encryption
                        $KerbEnc = $Smb.KerberosTicketEncryption
                        $HasRC4 = $KerbEnc -match 'RC4'
                        [void]$AllChecks.Add((New-CheckResult -Id "PROF-KERB-$($SA.StorageAccountName)" `
                            -Category 'FSLogix & Profiles' -Name 'Kerberos Ticket Encryption' `
                            -Description 'RC4-HMAC should be disabled — use AES-256 only' `
                            -Status $(if ($HasRC4) { 'Warning' } else { 'Pass' }) `
                            -Severity 'Medium' `
                            -Details "KerberosEncryption: $KerbEnc" `
                            -Reference 'https://learn.microsoft.com/en-us/azure/storage/files/files-smb-protocol#smb-security-settings'))

                        # Auth methods
                        $AuthMethods = $Smb.AuthenticationMethods
                        $HasNTLM = $AuthMethods -match 'NTLMv2'
                        [void]$AllChecks.Add((New-CheckResult -Id "PROF-AUTH-$($SA.StorageAccountName)" `
                            -Category 'FSLogix & Profiles' -Name 'Authentication Methods' `
                            -Description 'NTLMv2 should be disabled — use Kerberos only' `
                            -Status $(if ($HasNTLM) { 'Warning' } else { 'Pass' }) `
                            -Severity 'Medium' `
                            -Details "AuthMethods: $AuthMethods" `
                            -Reference 'https://learn.microsoft.com/en-us/azure/storage/files/files-smb-protocol#smb-security-settings'))

                        # Channel encryption
                        $ChanEnc = $Smb.ChannelEncryption
                        $HasAES256 = $ChanEnc -match 'AES-256'
                        [void]$AllChecks.Add((New-CheckResult -Id "PROF-SMBENC-$($SA.StorageAccountName)" `
                            -Category 'FSLogix & Profiles' -Name 'SMB Channel Encryption' `
                            -Description 'AES-256-GCM preferred for SMB channel encryption' `
                            -Status $(if ($HasAES256) { 'Pass' } else { 'Warning' }) `
                            -Severity 'Medium' `
                            -Details "ChannelEncryption: $ChanEnc" `
                            -Reference 'https://learn.microsoft.com/en-us/azure/storage/files/files-smb-protocol#smb-security-settings'))
                    }
                } catch { }
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
# CROSS-SUBSCRIPTION DISCOVERY (runs after all subscription loops)
# ═══════════════════════════════════════════════════════════════════════════

# Collect unique RG names from discovered AVD resources for scoped queries
$AvdResourceGroups = @(
    @($Discovery.Inventory.HostPools | ForEach-Object { $_.ResourceGroup }) +
    @($Discovery.Inventory.SessionHosts | ForEach-Object { $_.ResourceGroup }) +
    @($Discovery.Inventory.StorageAccounts | ForEach-Object { $_.ResourceGroup })
) | Sort-Object -Unique

$AvdRegions = @($Discovery.Inventory.SessionHosts | ForEach-Object { $_.Location } | Sort-Object -Unique)

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
    [void]$AllChecks.Add((New-CheckResult -Id "GOV-ORPHDISK" `
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
        Where-Object { $_.VirtualMachine -eq $null -and $_.ResourceGroupName -in $AvdResourceGroups })
    Write-Status "  Orphaned NICs: $($OrphanedNICs.Count)" -Level $(if ($OrphanedNICs.Count -gt 0) { 'WARN' } else { 'SUCCESS' })
    foreach ($NIC in $OrphanedNICs) {
        $Discovery.Inventory.OrphanedNICs += [PSCustomObject]@{
            Name          = $NIC.Name
            ResourceGroup = $NIC.ResourceGroupName
            Location      = $NIC.Location
            HasPublicIP   = ($NIC.IpConfigurations | Where-Object { $_.PublicIpAddress }).Count -gt 0
        }
    }
    [void]$AllChecks.Add((New-CheckResult -Id "GOV-ORPHNIC" `
        -Category 'Governance & Cost' -Name 'Orphaned NICs Detected' `
        -Description 'No unattached network interfaces should exist in AVD resource groups' `
        -Status $(if ($OrphanedNICs.Count -eq 0) { 'Pass' } else { 'Warning' }) `
        -Severity 'Low' `
        -Details "OrphanedNICs: $($OrphanedNICs.Count)" `
        -Recommendation 'Delete orphaned NICs — those with public IPs still incur charges.' `
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
    [void]$AllChecks.Add((New-CheckResult -Id "SEC-KEYVAULT" `
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
        [void]$AllChecks.Add((New-CheckResult -Id "SEC-KVPE" `
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
    [void]$AllChecks.Add((New-CheckResult -Id "NET-NETWATCHER" `
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
    [void]$AllChecks.Add((New-CheckResult -Id "NET-PRIVDNS" `
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
        $Scope = "/subscriptions/$($Discovery.Subscriptions[-1])/resourceGroups/$RG"
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
    [void]$AllChecks.Add((New-CheckResult -Id "GOV-POLICY" `
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
    foreach ($RG in $AvdResourceGroups) {
        $AlertRules += @(Get-AzMetricAlertRuleV2 -ResourceGroupName $RG -ErrorAction SilentlyContinue -WarningAction SilentlyContinue)
    }
    Write-Status "  Metric alert rules in AVD RGs: $($AlertRules.Count)" -Level $(if ($AlertRules.Count -gt 0) { 'SUCCESS' } else { 'WARN' })
    foreach ($AR in $AlertRules) {
        $Discovery.Inventory.AlertRules += [PSCustomObject]@{
            Name           = $AR.Name
            ResourceGroup  = ($AR.Id -split '/')[4]
            Severity       = $AR.Severity
            Enabled        = $AR.Enabled
            TargetResource = ($AR.Scopes | Select-Object -First 1)
        }
    }
    [void]$AllChecks.Add((New-CheckResult -Id "MON-ALERTS" `
        -Category 'Monitoring' -Name 'Alert Rules Configured' `
        -Description 'Metric or log alert rules should exist targeting AVD resources' `
        -Status $(if ($AlertRules.Count -gt 0) { 'Pass' } else { 'Warning' }) `
        -Severity 'Medium' `
        -Details "AlertRules: $($AlertRules.Count)$(if ($AlertRules.Count -gt 0) { " ($( ($AlertRules | ForEach-Object { $_.Name } | Select-Object -First 3) -join ', '))" })" `
        -Recommendation 'Configure metric alerts for CPU, memory, disk, and AVD-specific health signals.' `
        -Reference 'https://learn.microsoft.com/en-us/azure/virtual-desktop/set-up-service-alerts' `
        -Evidence @{ Count = $AlertRules.Count }))
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
    [void]$AllChecks.Add((New-CheckResult -Id "GOV-QUOTA" `
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
    [void]$AllChecks.Add((New-CheckResult -Id "GOV-CAPRESERV" `
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
    $SubScope = "/subscriptions/$($Discovery.Subscriptions[-1])"
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
    [void]$AllChecks.Add((New-CheckResult -Id "GOV-BUDGET" `
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
    [void]$AllChecks.Add((New-CheckResult -Id "GOV-RI" `
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

# ─── CROSS-RESOURCE TAG COMPLIANCE (GOV-018) ──────────────────────────
Write-Status "Cross-Resource Tag Compliance" -Level 'SECTION'
try {
    $RecommendedTags = @('Environment','Owner','CostCenter','Application','Department')
    $TagScores = @()
    # Score VMs (session hosts)
    foreach ($SH in $Discovery.Inventory.SessionHosts) {
        if ($SH.ResourceId) {
            try {
                $Res = Get-AzResource -ResourceId $SH.ResourceId -ErrorAction SilentlyContinue
                $Tags = if ($Res -and $Res.Tags) { @($Res.Tags.Keys) } else { @() }
                $Found = @($RecommendedTags | Where-Object { $T = $_; $Tags | Where-Object { $_ -like "*$T*" } })
                $TagScores += [PSCustomObject]@{ Type = 'VM'; Name = ($SH.ResourceId -split '/')[-1]; Score = $Found.Count; Total = $RecommendedTags.Count }
            } catch { }
        }
    }
    # Score VNets
    foreach ($VNet in $Discovery.Inventory.VNets) {
        if ($VNet.Id) {
            try {
                $Res = Get-AzResource -ResourceId $VNet.Id -ErrorAction SilentlyContinue
                $Tags = if ($Res -and $Res.Tags) { @($Res.Tags.Keys) } else { @() }
                $Found = @($RecommendedTags | Where-Object { $T = $_; $Tags | Where-Object { $_ -like "*$T*" } })
                $TagScores += [PSCustomObject]@{ Type = 'VNet'; Name = $VNet.Name; Score = $Found.Count; Total = $RecommendedTags.Count }
            } catch { }
        }
    }
    # Score Storage Accounts
    foreach ($SA in $Discovery.Inventory.StorageAccounts) {
        if ($SA.Id) {
            try {
                $Res = Get-AzResource -ResourceId $SA.Id -ErrorAction SilentlyContinue
                $Tags = if ($Res -and $Res.Tags) { @($Res.Tags.Keys) } else { @() }
                $Found = @($RecommendedTags | Where-Object { $T = $_; $Tags | Where-Object { $_ -like "*$T*" } })
                $TagScores += [PSCustomObject]@{ Type = 'Storage'; Name = $SA.Name; Score = $Found.Count; Total = $RecommendedTags.Count }
            } catch { }
        }
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
        Operations  = @{ Prefixes = @('OPS-','GOV-','SH-LB','SH-002','SH-DRAIN','SH-POWER','SH-STATUS','SH-IMG','SH-BSERIES','SH-SSD')
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
            # Try extracting weight from matching checks.json — use simple heuristic
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

    # Composite maturity score — weighted average of dimensions
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

# Box helper — fixed inner width of 54 chars
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
        $DScoreStr = '  — '
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

