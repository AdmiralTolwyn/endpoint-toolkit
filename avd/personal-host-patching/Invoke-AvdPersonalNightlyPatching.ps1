<#
.SYNOPSIS
    Orchestrates hibernate-safe, Hotpatch-aware nightly Windows Update and Defender
    servicing across AVD Personal Host Pools without wiping hibernated user sessions.
.DESCRIPTION
    Evaluates and services Windows 10/11 Enterprise AVD Personal session hosts where
    Azure Update Manager is unsupported on client OS SKUs and AVD Autoscale is
    configured to hibernate personal VMs after disconnect/inactivity:
    1. Fail-Closed Scaling Plan Exclusion Gate: Resolves active Personal Scaling Plans
       linked to each target Host Pool and fails closed when no enabled plan is visible
       (unless -ExclusionTagName or -AllowNoScalingPlan is explicitly set) or when
       properties.exclusionTag is empty/mismatched, preventing 60-minute mid-patch
       Autoscale hibernation races.
    2. Authoritative Dual-Gate Session Protection: Combines pre/post-wake AVD Broker
       userSessions checks with in-guest Unicode Win32 WTSEnumerateSessionsW and loaded
       ProfileList hive checks (AD S-1-5-21-* and Entra S-1-12-1-*) so hibernated user
       sessions are never misclassified as 0-session hosts.
    3. Hotpatch-First Servicing: Installs rebootless updates (InstallationRebootBehavior
       == 0 or Windows 11 24H2+ Monthly B Hotpatch updates on VBS + AllowRebootlessUpdates
       enabled hosts, plus Defender signatures) without restarting when a Disconnected
       session is preserved in RAM, defers reboot-requiring Baseline LCUs within
       -DisconnectedSessionGraceDays, and triggers Intune OMA-DM sync before restoring
       HibernationState/Hibernated.
    4. Self-Healing Sliding Window & Safety Nets: Services hosts in a paced sliding
       window (-MaxConcurrentVMs, -MaxStartsPerMinute) with deterministic SHA-256 weekday
       cohorting (default 5-cohort Mon-Fri sharding or smart catch-up skipping recently
       serviced hosts), Managed Run Command quota pruning (<25 limit), self-healing
       expired-lock recovery, read-only Triage, and an independent MorningUnlock mode.
.PARAMETER SubscriptionId
    Optional Azure Subscription ID used with -ResourceGroupName and -HostPoolName.
.PARAMETER ResourceGroupName
    Resource Group containing the target AVD Personal Host Pool.
.PARAMETER HostPoolName
    Name of the target AVD Personal Host Pool.
.PARAMETER HostPoolResourceId
    Full ARM Resource ID of one or more AVD Personal Host Pools.
.PARAMETER ScalingPlanResourceId
    Optional ARM Resource ID(s) of Scaling Plans when subscription-wide scalingPlans/read
    is not granted to the caller (supports Resource-Group-only RBAC).
.PARAMETER ExclusionTagName
    Optional expected Scaling Plan exclusionTag name; validated against any active
    Scaling Plan or used directly when -AllowNoScalingPlan is specified.
.PARAMETER AllowNoScalingPlan
    Explicitly permits Execute mode on a Personal Host Pool that has no enabled AVD
    Scaling Plan linked (prevents silent fail-open when RBAC hides Scaling Plan RGs).
.PARAMETER Mode
    Audit (default, 100% read-only), Triage (100% read-only inspection of active/expired
    AvdPatchLock tags and orphaned RunCommands), MorningUnlock (restores allowNewSession
    and clears expired idle locks before business hours), LocalGuestProbe (runs the
    in-guest WTS/VBS/Hotpatch/WUA probe locally in read-only mode with zero Azure calls),
    or Execute.
.PARAMETER ForceUnlockActive
    In MorningUnlock mode, forces drain/tag recovery even if AvdPatchLock has not yet
    expired (use only after confirming the nightly pipeline run was terminated).
.PARAMETER CohortCount
    Number of deterministic SHA-256 host cohorts (1 = all hosts). Default: 1 (pipeline
    recommends 5 for Mon-Fri 20% daily shards).
.PARAMETER CohortIndex
    Zero-based cohort index (0..CohortCount-1), or -1 for automatic weekday selection
    (Mon..Fri = 0..4 when CohortCount == 5; catch-up window on unmapped days). Default: 0.
.PARAMETER CatchUpSkipRecentHours
    During automatic catch-up windows (e.g. CohortCount=4 on Friday, or Sat/Sun), skips
    waking hosts whose AvdPatchState records a clean servicing run within this many hours
    (0 disables skipping). Default: 96.
.PARAMETER MaxConcurrentVMs
    Maximum session hosts actively woken/patching concurrently in the sliding window. Default: 60.
.PARAMETER MaxStartsPerMinute
    Maximum VM start/resume operations initiated per minute to protect regional capacity. Default: 15.
.PARAMETER DisconnectedSessionGraceDays
    Days to defer reboot-requiring Baseline LCUs when a Disconnected user session is in RAM. Default: 3.
.PARAMETER PostRebootSettleSeconds
    Seconds after a verified OS reboot before calling deallocate?hibernate=true. Default: 300.
.PARAMETER PostRebootZeroSessionAction
    Power state target after a 0-session host reboots and originally started Hibernated:
    Hibernate (default, prewarmed after settle) or Deallocate.
.PARAMETER StopNewWakeAtLocalHour
    Local hour (0..23 in -MaintenanceTimeZoneId, or -1 to disable) after which new VM
    wakes are skipped so in-flight hosts finish before business hours. Default: -1.
.PARAMETER MaintenanceTimeZoneId
    Windows TimeZoneInfo ID used for StopNewWakeAtLocalHour. Validated before execution.
    Default: W. Europe Standard Time.
.PARAMETER RunCommandTimeoutSeconds
    Managed Run Command execution timeout in seconds. Default: 2700.
.PARAMETER LockTtlMinutes
    Minutes before an AvdPatchLock tag is considered expired and eligible for recovery. Default: 180.
.PARAMETER MinFreeDiskSpaceGB
    Minimum free space on C: (GiB) required in-guest before downloading OS updates. Default: 10.
.PARAMETER MaxFailurePercent
    Circuit-breaker failure percentage threshold across processed hosts. Default: 20.
.PARAMETER MaxFailureCount
    Minimum absolute host failures before MaxFailurePercent aborts new launches. Default: 5.
.PARAMETER IncludeDrivers
    Include Driver updates in WUA search/install when specified.
.PARAMETER IncludePreviewUpdates
    Allow non-security Preview quality updates when specified.
.PARAMETER IncludeQuarantined
    Allow patching VMs tagged AvdLifecycleState=Quarantined when explicitly specified.
.PARAMETER TargetSessionHostName
    Optional filter restricting evaluation to specific short hostnames.
.PARAMETER DesktopVirtApiVersion
    ARM API version for Microsoft.DesktopVirtualization. Default: 2024-04-03.
.PARAMETER ComputeApiVersion
    ARM API version for Microsoft.Compute. Default: 2024-07-01.
.PARAMETER TagsApiVersion
    ARM API version for Microsoft.Resources/tags. Default: 2021-04-01.
.PARAMETER ManagedIdentityClientId
    Optional Client ID of a User-Assigned Managed Identity.
.PARAMETER ArmAccessToken
    Optional pre-acquired ARM bearer token (primarily for testing or custom wrappers).
.PARAMETER AuthMode
    Auto (default), ManagedIdentity, or AzContext.
.OUTPUTS
    PSCustomObject[]
.EXAMPLE
    .\Invoke-AvdPersonalNightlyPatching.ps1 -ResourceGroupName 'rg-avd-core' -HostPoolName 'vdpool-pers-01' -Mode Audit
.EXAMPLE
    .\Invoke-AvdPersonalNightlyPatching.ps1 `
        -HostPoolResourceId '/subscriptions/00000000-0000-0000-0000-000000000000/resourceGroups/rg-avd-core/providers/Microsoft.DesktopVirtualization/hostPools/vdpool-pers-01' `
        -ScalingPlanResourceId '/subscriptions/00000000-0000-0000-0000-000000000000/resourceGroups/rg-avd-core/providers/Microsoft.DesktopVirtualization/scalingPlans/sp-pers-01' `
        -Mode Execute -CohortCount 5 -CohortIndex -1 -StopNewWakeAtLocalHour 5 -Confirm:$false
.EXAMPLE
    .\Invoke-AvdPersonalNightlyPatching.ps1 -Mode LocalGuestProbe | Format-List *
.NOTES
    - PowerShell Baseline: Windows PowerShell 5.1+ (compatible with PowerShell 7+).
    - Long-Running Pipeline Auth (>55 min): When invoked via AzurePowerShell@5 with
      Workload Identity Federation (OIDC), back the Azure DevOps service connection with
      a User-Assigned Managed Identity (24h token lifetime) rather than an App Registration
      federated credential (~60m assertion lifetime, AADSTS700024), or size cohorts <55m.
    - Managed Run Command Quota: Azure enforces <=25 Managed Run Commands per VM. This
      script deletes each command immediately after completion and prunes stale Avd*
      RunCommands during lock recovery.
    - Sibling Interoperability: Honors AvdLegalHold (any non-empty value except
      false/0/no) and AvdLifecycleState=Quarantined shared with
      Remove-InactiveAvdPersonalHosts.ps1, which reciprocally honors active AvdPatchLock.
#>

#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $false)] [string]   $SubscriptionId,
    [Parameter(Mandatory = $false)] [string]   $ResourceGroupName,
    [Parameter(Mandatory = $false)] [string]   $HostPoolName,
    [Parameter(Mandatory = $false)] [string[]] $HostPoolResourceId,
    [Parameter(Mandatory = $false)] [string[]] $ScalingPlanResourceId,
    [Parameter(Mandatory = $false)] [string]   $ExclusionTagName,
    [Parameter(Mandatory = $false)] [switch]   $AllowNoScalingPlan,
    [Parameter(Mandatory = $false)] [ValidateSet('Audit', 'Triage', 'MorningUnlock', 'LocalGuestProbe', 'Execute')] [string] $Mode = 'Audit',
    [Parameter(Mandatory = $false)] [switch]   $ForceUnlockActive,
    [Parameter(Mandatory = $false)] [ValidateRange(1, 31)]    [int] $CohortCount = 1,
    [Parameter(Mandatory = $false)] [ValidateRange(-1, 30)]   [int] $CohortIndex = 0,
    [Parameter(Mandatory = $false)] [ValidateRange(0, 336)]   [int] $CatchUpSkipRecentHours = 96,
    [Parameter(Mandatory = $false)] [ValidateRange(1, 300)]   [int] $MaxConcurrentVMs = 60,
    [Parameter(Mandatory = $false)] [ValidateRange(1, 120)]   [int] $MaxStartsPerMinute = 15,
    [Parameter(Mandatory = $false)] [ValidateRange(0, 30)]    [int] $DisconnectedSessionGraceDays = 3,
    [Parameter(Mandatory = $false)] [ValidateRange(0, 1800)]  [int] $PostRebootSettleSeconds = 300,
    [Parameter(Mandatory = $false)] [ValidateSet('Hibernate', 'Deallocate')] [string] $PostRebootZeroSessionAction = 'Hibernate',
    [Parameter(Mandatory = $false)] [ValidateRange(-1, 23)]   [int] $StopNewWakeAtLocalHour = -1,
    [Parameter(Mandatory = $false)] [string]   $MaintenanceTimeZoneId = 'W. Europe Standard Time',
    [Parameter(Mandatory = $false)] [ValidateRange(300, 5400)] [int] $RunCommandTimeoutSeconds = 2700,
    [Parameter(Mandatory = $false)] [ValidateRange(30, 720)]  [int] $LockTtlMinutes = 180,
    [Parameter(Mandatory = $false)] [ValidateRange(4, 100)]   [int] $MinFreeDiskSpaceGB = 10,
    [Parameter(Mandatory = $false)] [ValidateRange(1, 100)]   [int] $MaxFailurePercent = 20,
    [Parameter(Mandatory = $false)] [ValidateRange(1, 500)]   [int] $MaxFailureCount = 5,
    [Parameter(Mandatory = $false)] [switch]   $IncludeDrivers,
    [Parameter(Mandatory = $false)] [switch]   $IncludePreviewUpdates,
    [Parameter(Mandatory = $false)] [switch]   $IncludeQuarantined,
    [Parameter(Mandatory = $false)] [string[]] $TargetSessionHostName,
    [Parameter(Mandatory = $false)] [string]   $DesktopVirtApiVersion = '2024-04-03',
    [Parameter(Mandatory = $false)] [string]   $ComputeApiVersion = '2024-07-01',
    [Parameter(Mandatory = $false)] [string]   $TagsApiVersion = '2021-04-01',
    [Parameter(Mandatory = $false)] [string]   $ManagedIdentityClientId,
    [Parameter(Mandatory = $false)] [string]   $ArmAccessToken,
    [Parameter(Mandatory = $false)] [ValidateSet('Auto', 'ManagedIdentity', 'AzContext')] [string] $AuthMode = 'Auto'
)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

function Get-AvdPatchProp {
    <#
    .SYNOPSIS
        Reads a property or dictionary key case-insensitively under Set-StrictMode 2.0.
    .DESCRIPTION
        Returns the value of Name from InputObject or $null when InputObject or Name is absent.
    .PARAMETER InputObject
        Object or IDictionary to inspect.
    .PARAMETER Name
        Property or key name to retrieve.
    .OUTPUTS
        System.Object
    #>
    param(
        [Parameter(Mandatory = $false)] $InputObject,
        [Parameter(Mandatory = $true)]  [string] $Name
    )
    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [System.Collections.IDictionary]) {
        foreach ($k in $InputObject.Keys) {
            if ([string]::Equals([string]$k, $Name, [System.StringComparison]::OrdinalIgnoreCase)) {
                return $InputObject[$k]
            }
        }
        return $null
    }
    foreach ($p in $InputObject.PSObject.Properties) {
        if ([string]::Equals($p.Name, $Name, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $p.Value
        }
    }
    return $null
}

function Get-AvdPatchNowUtc {
    <#
    .SYNOPSIS
        Returns the current UTC timestamp as a DateTimeOffset.
    .DESCRIPTION
        Seam around DateTimeOffset::UtcNow so offline unit tests can control simulated time.
    .OUTPUTS
        System.DateTimeOffset
    #>
    param()
    return [DateTimeOffset]::UtcNow
}

function Start-AvdPatchPollSleep {
    <#
    .SYNOPSIS
        Sleeps between sliding-window polling iterations.
    .DESCRIPTION
        Wraps Start-Sleep so offline regression tests can override tick delays deterministically.
    .PARAMETER Seconds
        Seconds to pause between orchestrator ticks.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    param(
        [Parameter(Mandatory = $true)] [int] $Seconds
    )
    if ($Seconds -gt 0) { Start-Sleep -Seconds $Seconds }
}

function Get-AvdPatchArmToken {
    <#
    .SYNOPSIS
        Acquires or refreshes an Azure Resource Manager OAuth bearer token.
    .DESCRIPTION
        Reuses a cached token in TokenState while at least 10 minutes of validity remain;
        otherwise acquires a fresh ARM token via Azure Automation IDENTITY_ENDPOINT,
        Azure VM IMDS, or Az.Accounts (Get-AzAccessToken), decoding JWT 'exp' when needed.
    .PARAMETER TokenState
        Hashtable holding Token, ExpiresUtc, AuthMode, ClientId, and StaticToken.
    .OUTPUTS
        System.String
    #>
    param(
        [Parameter(Mandatory = $true)] [hashtable] $TokenState
    )
    $now = Get-AvdPatchNowUtc
    if (-not [string]::IsNullOrWhiteSpace([string]$TokenState['StaticToken'])) {
        return [string]$TokenState['StaticToken']
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$TokenState['Token']) -and
        $null -ne $TokenState['ExpiresUtc'] -and
        ([DateTimeOffset]$TokenState['ExpiresUtc']) -gt $now.AddMinutes(10)) {
        return [string]$TokenState['Token']
    }

    $mode = [string]$TokenState['AuthMode']
    $clientId = [string]$TokenState['ClientId']
    $resource = 'https://management.azure.com/'
    $acquired = $null
    $expUtc = $null

    if ($mode -in @('Auto', 'ManagedIdentity')) {
        $miTargets = @()
        if (-not [string]::IsNullOrWhiteSpace($env:IDENTITY_ENDPOINT) -and -not [string]::IsNullOrWhiteSpace($env:IDENTITY_HEADER)) {
            $q = "resource=$([uri]::EscapeDataString($resource))&api-version=2019-08-01" + $(if (-not [string]::IsNullOrWhiteSpace($clientId)) { "&client_id=$([uri]::EscapeDataString($clientId))" } else { '' })
            $sep = if ($env:IDENTITY_ENDPOINT.Contains('?')) { '&' } else { '?' }
            $miTargets += @{ Uri = "$($env:IDENTITY_ENDPOINT)$sep$q"; Headers = @{ 'X-IDENTITY-HEADER' = $env:IDENTITY_HEADER; Metadata = 'true' }; Timeout = 30 }
        }
        $imdsQ = "api-version=2018-02-01&resource=$([uri]::EscapeDataString($resource))" + $(if (-not [string]::IsNullOrWhiteSpace($clientId)) { "&client_id=$([uri]::EscapeDataString($clientId))" } else { '' })
        $miTargets += @{ Uri = "http://169.254.169.254/metadata/identity/oauth2/token?$imdsQ"; Headers = @{ Metadata = 'true' }; Timeout = 10 }

        $lastMiErr = $null
        foreach ($ep in $miTargets) {
            if (-not [string]::IsNullOrWhiteSpace($acquired)) { break }
            try {
                $resp = Invoke-WebRequest -UseBasicParsing -Method Get -Uri $ep.Uri -Headers $ep.Headers -TimeoutSec $ep.Timeout
                $obj = $resp.Content | ConvertFrom-Json
                $acquired = [string](Get-AvdPatchProp -InputObject $obj -Name 'access_token')
                $expOn = Get-AvdPatchProp -InputObject $obj -Name 'expires_on'
                if ($null -ne $expOn -and "$expOn" -match '^\d+$') { $expUtc = [DateTimeOffset]::FromUnixTimeSeconds([long]$expOn) }
            } catch {
                $lastMiErr = $_
            }
        }
        if ([string]::IsNullOrWhiteSpace($acquired) -and $mode -eq 'ManagedIdentity') {
            $errTxt = if ($null -ne $lastMiErr) { $lastMiErr.Exception.Message } else { 'No Managed Identity endpoint responded.' }
            throw "Failed to acquire Managed Identity token for ARM: $errTxt"
        }
    }

    if ([string]::IsNullOrWhiteSpace($acquired) -and $mode -in @('Auto', 'AzContext')) {
        if (-not (Get-Command -Name 'Get-AzAccessToken' -ErrorAction SilentlyContinue)) {
            throw 'No Managed Identity endpoint responded and Az.Accounts (Get-AzAccessToken) is unavailable.'
        }
        $azTok = Get-AzAccessToken -ResourceUrl $resource
        $rawTok = Get-AvdPatchProp -InputObject $azTok -Name 'Token'
        if ($rawTok -is [System.Security.SecureString]) {
            $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($rawTok)
            try { $acquired = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) } finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
        } else {
            $acquired = [string]$rawTok
        }
        $azExp = Get-AvdPatchProp -InputObject $azTok -Name 'ExpiresOn'
        if ($null -ne $azExp) { $expUtc = [DateTimeOffset]$azExp }
    }

    if ([string]::IsNullOrWhiteSpace($acquired)) {
        throw "Unable to acquire ARM access token under AuthMode='$mode'."
    }
    if ($null -eq $expUtc) {
        $expUtc = $now.AddMinutes(50)
        $parts = $acquired.Split('.')
        if ($parts.Count -ge 2) {
            try {
                $b64 = $parts[1].Replace('-', '+').Replace('_', '/')
                switch ($b64.Length % 4) { 2 { $b64 += '==' } 3 { $b64 += '=' } }
                $jwtObj = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($b64)) | ConvertFrom-Json
                $jwtExp = Get-AvdPatchProp -InputObject $jwtObj -Name 'exp'
                if ($null -ne $jwtExp -and [long]$jwtExp -gt 0) { $expUtc = [DateTimeOffset]::FromUnixTimeSeconds([long]$jwtExp) }
            } catch { $expUtc = $now.AddMinutes(50) }
        }
    }
    $TokenState['Token'] = $acquired
    $TokenState['ExpiresUtc'] = $expUtc
    return $acquired
}

function Invoke-AvdPatchRestRequest {
    <#
    .SYNOPSIS
        Executes an HTTPS ARM REST call with origin validation, 401 refresh, and retry backoff.
    .DESCRIPTION
        Verifies the target URI targets an HTTPS Azure management endpoint before sending
        bearer credentials, refreshes expired tokens on HTTP 401, and retries transient
        HTTP 429/500/502/503/504 responses and socket/DNS failures ($code == 0).
    .PARAMETER Method
        HTTP method (Get, Post, Put, Patch, Delete).
    .PARAMETER Uri
        Full HTTPS ARM endpoint URI.
    .PARAMETER TokenState
        Token state hashtable passed to Get-AvdPatchArmToken.
    .PARAMETER Body
        Optional hashtable/PSObject or JSON string request payload.
    .PARAMETER MaxRetries
        Maximum transient retry attempts. Default: 5.
    .PARAMETER AllowNotFound
        Returns StatusCode=404 and Body=$null instead of throwing on HTTP 404.
    .OUTPUTS
        PSCustomObject
    #>
    param(
        [Parameter(Mandatory = $true)]  [ValidateSet('Get', 'Post', 'Put', 'Patch', 'Delete')] [string] $Method,
        [Parameter(Mandatory = $true)]  [string] $Uri,
        [Parameter(Mandatory = $true)]  [hashtable] $TokenState,
        [Parameter(Mandatory = $false)] $Body = $null,
        [Parameter(Mandatory = $false)] [int] $MaxRetries = 5,
        [Parameter(Mandatory = $false)] [switch] $AllowNotFound
    )
    $parsedUri = [System.Uri]$Uri
    $allowedArmHosts = @('management.azure.com', 'management.usgovcloudapi.net', 'management.chinacloudapi.cn')
    if ($parsedUri.Scheme -ne 'https' -or ($allowedArmHosts -notcontains $parsedUri.Host.ToLowerInvariant())) {
        throw "Refusing to send ARM bearer token to untrusted URI '$Uri'."
    }

    $jsonBody = $null
    if ($null -ne $Body) {
        $jsonBody = if ($Body -is [string]) { $Body } else { $Body | ConvertTo-Json -Depth 15 -Compress }
    }

    $attempt = 0
    while ($true) {
        $attempt++
        $token = Get-AvdPatchArmToken -TokenState $TokenState
        $headers = @{
            Authorization            = ('Bearer ' + $token)
            Accept                   = 'application/json'
            'x-ms-client-request-id' = [guid]::NewGuid().ToString()
        }
        $iwrParams = @{
            UseBasicParsing = $true
            Method          = $Method
            Uri             = $Uri
            Headers         = $headers
            TimeoutSec      = 120
            ErrorAction     = 'Stop'
        }
        if ($null -ne $jsonBody) {
            $iwrParams['ContentType'] = 'application/json; charset=utf-8'
            $iwrParams['Body'] = [System.Text.Encoding]::UTF8.GetBytes($jsonBody)
        }

        try {
            $resp = Invoke-WebRequest @iwrParams
            $parsed = $null
            if (-not [string]::IsNullOrWhiteSpace($resp.Content)) {
                $parsed = $resp.Content | ConvertFrom-Json
            }
            return [PSCustomObject]@{
                StatusCode = [int]$resp.StatusCode
                Headers    = $resp.Headers
                Body       = $parsed
                RawContent = $resp.Content
            }
        } catch {
            $errRecord = $_
            $exResp = Get-AvdPatchProp -InputObject $errRecord.Exception -Name 'Response'
            $code = if ($null -ne $exResp -and $null -ne (Get-AvdPatchProp -InputObject $exResp -Name 'StatusCode')) { [int]$exResp.StatusCode } else { 0 }
            $errBody = ''
            if ($null -ne $errRecord.ErrorDetails -and -not [string]::IsNullOrWhiteSpace([string]$errRecord.ErrorDetails.Message)) {
                $errBody = [string]$errRecord.ErrorDetails.Message
            }
            $retryAfterSec = 0
            if ($null -ne $exResp) {
                try {
                    $raHeader = $null
                    $raVals = $null
                    if ($null -ne $exResp.Headers -and $exResp.Headers.PSObject.Methods['TryGetValues']) {
                        if ($exResp.Headers.TryGetValues('Retry-After', [ref]$raVals)) { $raHeader = [string](@($raVals)[0]) }
                    } elseif ($null -ne $exResp.Headers) {
                        $raHeader = [string]$exResp.Headers['Retry-After']
                    }
                    if (-not [string]::IsNullOrWhiteSpace($raHeader) -and $raHeader.Trim() -match '^\d+$') {
                        $retryAfterSec = [Math]::Min(120, [int]$raHeader.Trim())
                    }
                } catch { $retryAfterSec = 0 }

                if ([string]::IsNullOrWhiteSpace($errBody)) {
                    try {
                        if ($exResp.PSObject.Methods['GetResponseStream']) {
                            $stream = $exResp.GetResponseStream()
                            if ($null -ne $stream) {
                                $reader = New-Object System.IO.StreamReader($stream)
                                $errBody = $reader.ReadToEnd()
                                $reader.Close()
                            }
                        } elseif ($null -ne (Get-AvdPatchProp -InputObject $exResp -Name 'Content') -and $exResp.Content.PSObject.Methods['ReadAsStringAsync']) {
                            $errBody = [string]$exResp.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                        }
                    } catch { $errBody = '' }
                }
            }

            if ($code -eq 404 -and $AllowNotFound) {
                return [PSCustomObject]@{ StatusCode = 404; Headers = @{}; Body = $null; RawContent = $errBody }
            }
            if ($code -eq 401 -and [string]::IsNullOrWhiteSpace([string]$TokenState['StaticToken']) -and $attempt -le 2) {
                $TokenState['Token'] = $null
                continue
            }
            if (($code -eq 0 -or $code -in @(429, 500, 502, 503, 504)) -and $attempt -le $MaxRetries) {
                $waitSec = if ($retryAfterSec -gt 0) { $retryAfterSec } else { [Math]::Min(60, [int][Math]::Pow(2, $attempt) + (Get-Random -Minimum 1 -Maximum 4)) }
                Start-AvdPatchPollSleep -Seconds $waitSec
                continue
            }
            $ex = New-Object System.InvalidOperationException("ARM $Method $Uri failed (HTTP $code): $($errRecord.Exception.Message) $errBody", $errRecord.Exception)
            $ex.Data['HttpStatusCode'] = $code
            throw $ex
        }
    }
}

function Get-AvdPatchAsyncOperationStatus {
    <#
    .SYNOPSIS
        Extracts or polls an ARM Azure-AsyncOperation / Location LRO status endpoint.
    .DESCRIPTION
        When ExtractFromResponse is provided, extracts the LRO polling URL portably across
        Windows PowerShell 5.1 and PowerShell 7. When OperationUri is provided, polls the
        LRO endpoint and degrades gracefully to Status='Unavailable' on HTTP 403/404 so
        Resource-Group-only RBAC (without subscription locations/operations/read) falls
        back cleanly to instanceView polling.
    .PARAMETER OperationUri
        Full HTTPS Azure-AsyncOperation URI to poll.
    .PARAMETER TokenState
        ARM authentication token state hashtable.
    .PARAMETER ExtractFromResponse
        Optional response object from Invoke-AvdPatchRestRequest when extracting the URL.
    .OUTPUTS
        PSCustomObject
    #>
    param(
        [Parameter(Mandatory = $false)] [string] $OperationUri,
        [Parameter(Mandatory = $false)] [hashtable] $TokenState,
        [Parameter(Mandatory = $false)] $ExtractFromResponse
    )
    if ($null -ne $ExtractFromResponse) {
        $foundUrl = ''
        if ($null -ne $ExtractFromResponse.Headers) {
            foreach ($hName in @('Azure-AsyncOperation', 'Location')) {
                $val = $null
                try {
                    if ($ExtractFromResponse.Headers -is [System.Collections.IDictionary]) {
                        foreach ($k in $ExtractFromResponse.Headers.Keys) {
                            if ([string]::Equals([string]$k, $hName, [System.StringComparison]::OrdinalIgnoreCase)) {
                                $raw = $ExtractFromResponse.Headers[$k]
                                $val = if (($raw -is [System.Array]) -or ($raw -is [System.Collections.IEnumerable] -and $raw -isnot [string])) { [string](@($raw)[0]) } else { [string]$raw }
                                break
                            }
                        }
                    } elseif ($ExtractFromResponse.Headers.PSObject.Methods['TryGetValues']) {
                        $outVals = $null
                        if ($ExtractFromResponse.Headers.TryGetValues($hName, [ref]$outVals)) { $val = [string](@($outVals)[0]) }
                    }
                } catch { $val = $null }
                if (-not [string]::IsNullOrWhiteSpace($val)) { $foundUrl = $val.Trim(); break }
            }
        }
        return [PSCustomObject]@{ OperationUri = $foundUrl; Status = ''; ErrorCode = ''; ErrorMessage = '' }
    }

    if ([string]::IsNullOrWhiteSpace($OperationUri) -or $null -eq $TokenState) {
        return [PSCustomObject]@{ OperationUri = ''; Status = 'Unavailable'; ErrorCode = ''; ErrorMessage = '' }
    }
    try {
        $res = Invoke-AvdPatchRestRequest -Method Get -Uri $OperationUri -TokenState $TokenState -AllowNotFound
    } catch {
        $httpCode = if ($null -ne $_.Exception.Data -and $_.Exception.Data.Contains('HttpStatusCode')) { [int]$_.Exception.Data['HttpStatusCode'] } else { 0 }
        if ($httpCode -eq 403) {
            return [PSCustomObject]@{ OperationUri = ''; Status = 'Unavailable'; ErrorCode = ''; ErrorMessage = '' }
        }
        throw
    }
    if ($res.StatusCode -eq 404 -or $null -eq $res.Body) {
        return [PSCustomObject]@{ OperationUri = $OperationUri; Status = 'Unavailable'; ErrorCode = ''; ErrorMessage = '' }
    }
    $errObj = Get-AvdPatchProp -InputObject $res.Body -Name 'error'
    return [PSCustomObject]@{
        OperationUri = $OperationUri
        Status       = [string](Get-AvdPatchProp -InputObject $res.Body -Name 'status')
        ErrorCode    = [string](Get-AvdPatchProp -InputObject $errObj -Name 'code')
        ErrorMessage = [string](Get-AvdPatchProp -InputObject $errObj -Name 'message')
    }
}

function Invoke-AvdPatchArmPagedGet {
    <#
    .SYNOPSIS
        Retrieves all pages of an ARM collection endpoint following same-origin nextLink.
    .DESCRIPTION
        Verifies HTTPS same-origin authority on every continuation link and returns a flat
        array of all value[] items.
    .PARAMETER Uri
        Initial collection GET URI.
    .PARAMETER TokenState
        Token state hashtable for authentication.
    .OUTPUTS
        System.Object[]
    #>
    param(
        [Parameter(Mandatory = $true)] [string] $Uri,
        [Parameter(Mandatory = $true)] [hashtable] $TokenState
    )
    $origin = [System.Uri]$Uri
    $items = New-Object 'System.Collections.Generic.List[object]'
    $next = $Uri
    $guard = 0
    while (-not [string]::IsNullOrWhiteSpace($next)) {
        $guard++
        if ($guard -gt 200) { throw "Pagination loop guard exceeded on $Uri" }
        $nextUri = [System.Uri]$next
        if ($nextUri.Scheme -ne 'https' -or -not $nextUri.Authority.Equals($origin.Authority, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Refusing to follow off-origin ARM nextLink '$next'."
        }
        $res = Invoke-AvdPatchRestRequest -Method Get -Uri $next -TokenState $TokenState
        $val = Get-AvdPatchProp -InputObject $res.Body -Name 'value'
        if ($null -ne $val) {
            foreach ($entry in @($val)) {
                if ($null -ne $entry) { [void]$items.Add($entry) }
            }
        }
        $next = [string](Get-AvdPatchProp -InputObject $res.Body -Name 'nextLink')
    }
    return $items.ToArray()
}

function Get-AvdPatchVmPowerSummary {
    <#
    .SYNOPSIS
        Classifies a VM's exact power, hibernation, agent, and resume state from InstanceView.
    .DESCRIPTION
        Performs case-insensitive matching on instanceView.statuses[] so
        PowerState/deallocated + HibernationState/Hibernated is distinguished from
        Deallocated, StoppedAllocated, Running, or InTransition, and inspects
        AzureHibernateExtension substatuses for platform cold-boot fallback.
    .PARAMETER VmResource
        Parsed ARM VirtualMachine object (expanded with instanceView).
    .OUTPUTS
        PSCustomObject
    #>
    param(
        [Parameter(Mandatory = $true)] $VmResource
    )
    $props = Get-AvdPatchProp -InputObject $VmResource -Name 'properties'
    $addCap = Get-AvdPatchProp -InputObject $props -Name 'additionalCapabilities'
    $hibCap = [bool](Get-AvdPatchProp -InputObject $addCap -Name 'hibernationEnabled')

    $iv = Get-AvdPatchProp -InputObject $props -Name 'instanceView'
    if ($null -eq $iv) { $iv = Get-AvdPatchProp -InputObject $VmResource -Name 'instanceView' }
    $statuses = @(Get-AvdPatchProp -InputObject $iv -Name 'statuses')

    $powerCode = ''
    $isHibernated = $false
    foreach ($st in $statuses) {
        if ($null -eq $st) { continue }
        $code = [string](Get-AvdPatchProp -InputObject $st -Name 'code')
        if ($code.StartsWith('PowerState/', [System.StringComparison]::OrdinalIgnoreCase)) {
            $powerCode = $code.Substring(11).ToLowerInvariant()
        }
        if ([string]::Equals($code, 'HibernationState/Hibernated', [System.StringComparison]::OrdinalIgnoreCase)) {
            $isHibernated = $true
        }
    }

    $effectiveState = switch ($powerCode) {
        'running'      { 'Running' }
        'deallocated'  { if ($isHibernated) { 'Hibernated' } else { 'Deallocated' } }
        'stopped'      { 'StoppedAllocated' }
        'starting'     { 'InTransition' }
        'stopping'     { 'InTransition' }
        'deallocating' { 'InTransition' }
        default        { 'Unknown' }
    }

    $vmAgent = Get-AvdPatchProp -InputObject $iv -Name 'vmAgent'
    $agentReady = $false
    foreach ($as in @(Get-AvdPatchProp -InputObject $vmAgent -Name 'statuses')) {
        if ($null -eq $as) { continue }
        $disp = [string](Get-AvdPatchProp -InputObject $as -Name 'displayStatus')
        if ([string]::Equals($disp, 'Ready', [System.StringComparison]::OrdinalIgnoreCase)) {
            $agentReady = $true
        }
    }

    $bootSubstatus = ''
    foreach ($ext in @(Get-AvdPatchProp -InputObject $iv -Name 'extensions')) {
        if ($null -eq $ext) { continue }
        $extType = [string](Get-AvdPatchProp -InputObject $ext -Name 'type')
        $extName = [string](Get-AvdPatchProp -InputObject $ext -Name 'name')
        if ($extType -like '*WindowsHibernateExtension*' -or $extName -eq 'AzureHibernateExtension') {
            foreach ($sub in @(Get-AvdPatchProp -InputObject $ext -Name 'substatuses')) {
                if ($null -eq $sub) { continue }
                $sc = [string](Get-AvdPatchProp -InputObject $sub -Name 'code')
                if ($sc -like 'ComponentStatus/VMBootState/*') { $bootSubstatus = $sc }
            }
        }
    }

    return [PSCustomObject]@{
        EffectiveState     = $effectiveState
        PowerCode          = $powerCode
        IsHibernated       = $isHibernated
        HibernationEnabled = $hibCap
        VmAgentReady       = $agentReady
        BootSubstatus      = $bootSubstatus
    }
}

function ConvertTo-AvdPatchLockValue {
    <#
    .SYNOPSIS
        Serializes transient per-VM lock metadata into a <=240-char ARM tag value.
    .DESCRIPTION
        Encodes RunId, VmId, PriorState, PriorAllowNewSession, HadPriorExclusionTag,
        ExclusionTag name, and ExpiresUtc into a compact v2 pipe-delimited string so
        crash recovery and MorningUnlock can clean up tags without querying Scaling Plans.
    .PARAMETER RunId
        Orchestrator run GUID string.
    .PARAMETER VmId
        Immutable Compute properties.vmId GUID string.
    .PARAMETER PriorState
        Prior power state (Hibernated, Deallocated, Running, StoppedAllocated).
    .PARAMETER PriorAllowNewSession
        Prior sessionHost allowNewSession boolean.
    .PARAMETER HadPriorExclusionTag
        True if the VM already carried the Scaling Plan exclusionTag before this run.
    .PARAMETER ExclusionTag
        Scaling Plan exclusionTag key name stamped by this run (when HadPriorExclusionTag is false).
    .PARAMETER ExpiresUtc
        Lock expiration timestamp in UTC.
    .OUTPUTS
        System.String
    #>
    param(
        [Parameter(Mandatory = $true)]  [string] $RunId,
        [Parameter(Mandatory = $true)]  [string] $VmId,
        [Parameter(Mandatory = $true)]  [string] $PriorState,
        [Parameter(Mandatory = $true)]  [bool]   $PriorAllowNewSession,
        [Parameter(Mandatory = $true)]  [bool]   $HadPriorExclusionTag,
        [Parameter(Mandatory = $false)] [string] $ExclusionTag = 'excludeFromScaling',
        [Parameter(Mandatory = $true)]  [DateTimeOffset] $ExpiresUtc
    )
    $ans = if ($PriorAllowNewSession) { '1' } else { '0' }
    $pex = if ($HadPriorExclusionTag) { '1' } else { '0' }
    $safeTag = (($ExclusionTag -replace '[\|\r\n=]', '').Trim())
    if ($safeTag.Length -gt 64) { $safeTag = $safeTag.Substring(0, 64) }
    $exp = $ExpiresUtc.ToUnixTimeSeconds()
    return "v2|run=$($RunId.Trim())|vm=$($VmId.Trim())|ps=$PriorState|ans=$ans|pex=$pex|tag=$safeTag|exp=$exp"
}

function ConvertFrom-AvdPatchLockValue {
    <#
    .SYNOPSIS
        Parses an AvdPatchLock v2 tag string back into a structured object.
    .DESCRIPTION
        Validates the v2 prefix and expiration epoch and returns lock metadata or $null.
    .PARAMETER Value
        Raw AvdPatchLock tag string.
    .OUTPUTS
        PSCustomObject
    #>
    param(
        [Parameter(Mandatory = $false)] [string] $Value
    )
    if ([string]::IsNullOrWhiteSpace($Value) -or -not $Value.StartsWith('v2|', [System.StringComparison]::Ordinal)) {
        return $null
    }
    $map = @{}
    foreach ($seg in $Value.Split('|')) {
        $eq = $seg.IndexOf('=')
        if ($eq -gt 0) { $map[$seg.Substring(0, $eq)] = $seg.Substring($eq + 1) }
    }
    if (-not $map.ContainsKey('exp') -or $map['exp'] -notmatch '^\d+$') { return $null }
    $tagName = if ($map.ContainsKey('tag') -and -not [string]::IsNullOrWhiteSpace($map['tag'])) { [string]$map['tag'] } else { 'excludeFromScaling' }
    return [PSCustomObject]@{
        RunId                = [string]$map['run']
        VmId                 = [string]$map['vm']
        PriorState           = [string]$map['ps']
        PriorAllowNewSession = ($map['ans'] -eq '1')
        HadPriorExclusionTag = ($map['pex'] -eq '1')
        ExclusionTag         = $tagName
        ExpiresUtc           = [DateTimeOffset]::FromUnixTimeSeconds([long]$map['exp'])
    }
}

function ConvertTo-AvdPatchStateValue {
    <#
    .SYNOPSIS
        Serializes persistent per-VM servicing state into a <=240-char AvdPatchState tag.
    .DESCRIPTION
        Stores last successful UTC timestamp, OS UBR, outcome code, and Disconnected
        reboot-deferral start epoch across runs.
    .PARAMETER LastOkUtc
        Timestamp of the run in UTC.
    .PARAMETER Ubr
        OS Update Build Revision integer.
    .PARAMETER Outcome
        Short alphanumeric outcome string.
    .PARAMETER DeferSinceEpoch
        Unix epoch seconds when a reboot-requiring Baseline LCU was first deferred (0 if none).
    .OUTPUTS
        System.String
    #>
    param(
        [Parameter(Mandatory = $true)] [DateTimeOffset] $LastOkUtc,
        [Parameter(Mandatory = $true)] [int]    $Ubr,
        [Parameter(Mandatory = $true)] [string] $Outcome,
        [Parameter(Mandatory = $true)] [long]   $DeferSinceEpoch
    )
    $safeOutcome = ($Outcome -replace '[^A-Za-z0-9_\-]', '')
    if ($safeOutcome.Length -gt 48) { $safeOutcome = $safeOutcome.Substring(0, 48) }
    return "v2|ok=$($LastOkUtc.ToUniversalTime().ToString('yyyyMMddTHHmmZ', [Globalization.CultureInfo]::InvariantCulture))|ubr=$Ubr|res=$safeOutcome|def=$DeferSinceEpoch"
}

function ConvertFrom-AvdPatchStateValue {
    <#
    .SYNOPSIS
        Parses an AvdPatchState v2 tag string into a structured object.
    .DESCRIPTION
        Extracts LastOkUtc, UBR, Outcome, and DeferSinceEpoch so multi-day Disconnected
        grace tracking and catch-up window skipping work deterministically across runs.
    .PARAMETER Value
        Raw AvdPatchState tag string.
    .OUTPUTS
        PSCustomObject
    #>
    param(
        [Parameter(Mandatory = $false)] [string] $Value
    )
    if ([string]::IsNullOrWhiteSpace($Value) -or -not $Value.StartsWith('v2|', [System.StringComparison]::Ordinal)) {
        return [PSCustomObject]@{ LastOkCompact = ''; LastOkUtc = $null; Ubr = 0; Outcome = ''; DeferSinceEpoch = [long]0 }
    }
    $map = @{}
    foreach ($seg in $Value.Split('|')) {
        $eq = $seg.IndexOf('=')
        if ($eq -gt 0) { $map[$seg.Substring(0, $eq)] = $seg.Substring($eq + 1) }
    }
    $ubr = if ($map.ContainsKey('ubr') -and $map['ubr'] -match '^\d+$') { [int]$map['ubr'] } else { 0 }
    $def = if ($map.ContainsKey('def') -and $map['def'] -match '^\d+$') { [long]$map['def'] } else { [long]0 }
    $okRaw = if ($map.ContainsKey('ok')) { [string]$map['ok'] } else { '' }
    $parsedOk = $null
    if (-not [string]::IsNullOrWhiteSpace($okRaw)) {
        $dto = [DateTimeOffset]::MinValue
        if ([DateTimeOffset]::TryParseExact($okRaw, 'yyyyMMddTHHmm\Z', [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$dto)) {
            $parsedOk = $dto
        }
    }
    return [PSCustomObject]@{
        LastOkCompact   = $okRaw
        LastOkUtc       = $parsedOk
        Ubr             = $ubr
        Outcome         = [string]$map['res']
        DeferSinceEpoch = $def
    }
}

function Test-AvdPatchHostInCohort {
    <#
    .SYNOPSIS
        Determines whether a session host belongs to the active deterministic SHA-256 cohort.
    .DESCRIPTION
        Hashes the lowercase short hostname with SHA-256 so cohort membership is uniform
        and stateless. When CohortIndex == -1:
        - For CohortCount <= 5: maps Mon=0, Tue=1, Wed=2, Thu=3, Fri=4 when index < CohortCount;
          any weekday >= CohortCount (e.g. Friday when CohortCount == 4) or weekend (Sat/Sun)
          acts as a catch-up window returning $true.
        - For CohortCount > 5: rotates deterministically across all 0..CohortCount-1 shards
          by UTC epoch day so every shard is visited evenly.
    .PARAMETER ShortHostName
        Short session host VM name.
    .PARAMETER CohortCount
        Total number of cohorts (1 = always true).
    .PARAMETER CohortIndex
        Target 0-based cohort index, or -1 for automatic schedule selection.
    .PARAMETER NowUtc
        Reference timestamp used when CohortIndex == -1.
    .OUTPUTS
        System.Boolean
    #>
    param(
        [Parameter(Mandatory = $true)] [string] $ShortHostName,
        [Parameter(Mandatory = $true)] [int]    $CohortCount,
        [Parameter(Mandatory = $true)] [int]    $CohortIndex,
        [Parameter(Mandatory = $true)] [DateTimeOffset] $NowUtc
    )
    if ($CohortCount -le 1) { return $true }
    $effectiveIdx = $CohortIndex
    if ($effectiveIdx -lt 0) {
        if ($CohortCount -le 5) {
            $dow = [int]$NowUtc.DayOfWeek # Sun=0, Mon=1 .. Fri=5, Sat=6
            if ($dow -ge 1 -and $dow -le 5 -and ($dow - 1) -lt $CohortCount) {
                $effectiveIdx = ($dow - 1)
            } else {
                return $true
            }
        } else {
            $epochDay = [long][Math]::Floor($NowUtc.ToUnixTimeSeconds() / 86400)
            $effectiveIdx = [int]($epochDay % [long]$CohortCount)
        }
    }
    if ($effectiveIdx -ge $CohortCount) {
        throw "CohortIndex ($effectiveIdx) must be less than CohortCount ($CohortCount)."
    }
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($ShortHostName.Trim().ToLowerInvariant()))
        $val = [BitConverter]::ToUInt32($bytes, 0)
        return (($val % [uint32]$CohortCount) -eq [uint32]$effectiveIdx)
    } finally {
        $sha.Dispose()
    }
}

function Resolve-AvdPatchPoolExclusionTag {
    <#
    .SYNOPSIS
        Resolves and validates the Scaling Plan exclusionTag for a Personal Host Pool.
    .DESCRIPTION
        Queries Scaling Plans via explicit ScalingPlanArmId(s) (supporting RG-only RBAC)
        or subscription-level discovery, locates any enabled plan referencing HostPoolArmId,
        and enforces fail-closed validation:
        - Fails closed if an active Scaling Plan has an empty exclusionTag or mismatches
          ExpectedExclusionTag.
        - Fails closed when 0 active Scaling Plans are visible unless AllowNoScalingPlan
          is set or ExpectedExclusionTag is explicitly supplied (preventing silent
          'excludeFromScaling' mismatches when RBAC hides a Scaling Plan RG).
    .PARAMETER HostPoolArmId
        Full ARM resource ID of the Personal Host Pool.
    .PARAMETER ScalingPlanArmId
        Optional explicit Scaling Plan ARM ID(s) when subscription listing is restricted.
    .PARAMETER ExpectedExclusionTag
        Optional caller-supplied exclusionTag name to cross-validate or use explicitly.
    .PARAMETER AllowNoScalingPlan
        True when the pool intentionally runs without an active AVD Scaling Plan.
    .PARAMETER ApiVersion
        Microsoft.DesktopVirtualization API version.
    .PARAMETER TokenState
        ARM authentication token state hashtable.
    .OUTPUTS
        PSCustomObject
    #>
    param(
        [Parameter(Mandatory = $true)]  [string]   $HostPoolArmId,
        [Parameter(Mandatory = $false)] [string[]] $ScalingPlanArmId,
        [Parameter(Mandatory = $false)] [string]   $ExpectedExclusionTag,
        [Parameter(Mandatory = $false)] [bool]     $AllowNoScalingPlan = $false,
        [Parameter(Mandatory = $true)]  [string]   $ApiVersion,
        [Parameter(Mandatory = $true)]  [hashtable] $TokenState
    )
    if ($HostPoolArmId -notmatch '^/subscriptions/([^/]+)/resourceGroups/([^/]+)/providers/Microsoft\.DesktopVirtualization/hostPools/([^/]+)$') {
        throw "Invalid HostPool Resource ID format: '$HostPoolArmId'"
    }
    $subId = $Matches[1]
    $plans = New-Object 'System.Collections.Generic.List[object]'
    $explicitPlans = @($ScalingPlanArmId | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($explicitPlans.Count -gt 0) {
        foreach ($spId in $explicitPlans) {
            $res = Invoke-AvdPatchRestRequest -Method Get -Uri "https://management.azure.com$($spId.Trim())?api-version=$ApiVersion" -TokenState $TokenState
            if ($null -ne $res.Body) { [void]$plans.Add($res.Body) }
        }
    } else {
        foreach ($spItem in @(Invoke-AvdPatchArmPagedGet -Uri "https://management.azure.com/subscriptions/$subId/providers/Microsoft.DesktopVirtualization/scalingPlans?api-version=$ApiVersion" -TokenState $TokenState)) {
            [void]$plans.Add($spItem)
        }
    }

    $matchedPlans = New-Object 'System.Collections.Generic.List[object]'
    foreach ($sp in $plans) {
        $props = Get-AvdPatchProp -InputObject $sp -Name 'properties'
        foreach ($r in @(Get-AvdPatchProp -InputObject $props -Name 'hostPoolReferences')) {
            if ($null -eq $r) { continue }
            $refPath = [string](Get-AvdPatchProp -InputObject $r -Name 'hostPoolArmPath')
            $enabled = [bool](Get-AvdPatchProp -InputObject $r -Name 'scalingPlanEnabled')
            if ($enabled -and [string]::Equals($refPath.TrimEnd('/'), $HostPoolArmId.TrimEnd('/'), [System.StringComparison]::OrdinalIgnoreCase)) {
                [void]$matchedPlans.Add($sp)
            }
        }
    }

    if ($matchedPlans.Count -eq 0) {
        if (-not [string]::IsNullOrWhiteSpace($ExpectedExclusionTag)) {
            return [PSCustomObject]@{ IsValid = $true; HasActivePlan = $false; ScalingPlanName = ''; ExclusionTag = $ExpectedExclusionTag.Trim(); FailureReason = '' }
        }
        if ($AllowNoScalingPlan) {
            return [PSCustomObject]@{ IsValid = $true; HasActivePlan = $false; ScalingPlanName = ''; ExclusionTag = 'excludeFromScaling'; FailureReason = '' }
        }
        return [PSCustomObject]@{
            IsValid         = $false
            HasActivePlan   = $false
            ScalingPlanName = ''
            ExclusionTag    = ''
            FailureReason   = "No enabled Scaling Plan found for '$HostPoolArmId'. Grant subscription scalingPlans/read, pass -ScalingPlanResourceId, specify -ExclusionTagName, or set -AllowNoScalingPlan."
        }
    }

    $activePlan = $matchedPlans[0]
    $planName = [string](Get-AvdPatchProp -InputObject $activePlan -Name 'name')
    $planProps = Get-AvdPatchProp -InputObject $activePlan -Name 'properties'
    $planExcl = [string](Get-AvdPatchProp -InputObject $planProps -Name 'exclusionTag')
    if ([string]::IsNullOrWhiteSpace($planExcl)) {
        return [PSCustomObject]@{
            IsValid         = $false
            HasActivePlan   = $true
            ScalingPlanName = $planName
            ExclusionTag    = ''
            FailureReason   = "Active Scaling Plan '$planName' has no exclusionTag configured; aborting to prevent mid-patch 60-min hibernation."
        }
    }
    $trimmedPlanExcl = $planExcl.Trim()
    if (-not [string]::IsNullOrWhiteSpace($ExpectedExclusionTag) -and
        -not [string]::Equals($trimmedPlanExcl, $ExpectedExclusionTag.Trim(), [System.StringComparison]::OrdinalIgnoreCase)) {
        return [PSCustomObject]@{
            IsValid         = $false
            HasActivePlan   = $true
            ScalingPlanName = $planName
            ExclusionTag    = $trimmedPlanExcl
            FailureReason   = "ExclusionTagName '$ExpectedExclusionTag' does not match Scaling Plan '$planName' exclusionTag '$trimmedPlanExcl'."
        }
    }

    return [PSCustomObject]@{
        IsValid         = $true
        HasActivePlan   = $true
        ScalingPlanName = $planName
        ExclusionTag    = $trimmedPlanExcl
        FailureReason   = ''
    }
}

function Update-AvdPatchVmTags {
    <#
    .SYNOPSIS
        Merges or conditionally deletes tags on an Azure VM via Tags - Update At Scope.
    .DESCRIPTION
        Enforces ARM's 50-tag resource ceiling before Merge operations and performs
        ownership-conditional, case-insensitive key compare-and-delete on Delete: a tag
        is deleted only if its current value on the VM matches the expected value in
        TagsToApply, preventing accidental removal of another operator's or run's tag.
    .PARAMETER VmResourceId
        Full ARM resource ID of the virtual machine.
    .PARAMETER Operation
        Merge or Delete.
    .PARAMETER TagsToApply
        Hashtable of tag key-value pairs to merge, or expected owned key-value pairs to delete.
    .PARAMETER ExistingTags
        Current VM tags object/hashtable used to validate the 50-tag ceiling on Merge
        and verify exact ownership on Delete.
    .PARAMETER TagsApiVersion
        Microsoft.Resources/tags API version.
    .PARAMETER TokenState
        ARM authentication token state hashtable.
    .OUTPUTS
        PSCustomObject
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    param(
        [Parameter(Mandatory = $true)]  [string] $VmResourceId,
        [Parameter(Mandatory = $true)]  [ValidateSet('Merge', 'Delete')] [string] $Operation,
        [Parameter(Mandatory = $true)]  [hashtable] $TagsToApply,
        [Parameter(Mandatory = $false)] $ExistingTags = $null,
        [Parameter(Mandatory = $true)]  [string] $TagsApiVersion,
        [Parameter(Mandatory = $true)]  [hashtable] $TokenState
    )
    if ($Operation -eq 'Delete' -and $null -eq $ExistingTags) {
        throw "ExistingTags must be supplied for ownership-conditional tag Delete on '$VmResourceId'."
    }
    $currentMap = @{}
    if ($null -ne $ExistingTags) {
        if ($ExistingTags -is [System.Collections.IDictionary]) {
            foreach ($k in $ExistingTags.Keys) { $currentMap[[string]$k] = [string]$ExistingTags[$k] }
        } else {
            foreach ($p in $ExistingTags.PSObject.Properties) { $currentMap[$p.Name] = [string]$p.Value }
        }
    }

    $payloadTags = @{}
    if ($Operation -eq 'Merge') {
        $unionKeys = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($k in $currentMap.Keys) { [void]$unionKeys.Add($k) }
        foreach ($k in $TagsToApply.Keys) {
            $keyStr = [string]$k
            [void]$unionKeys.Add($keyStr)
            $val = [string]$TagsToApply[$k]
            if ($val.Length -gt 256) { throw "Tag '$keyStr' value length ($($val.Length)) exceeds ARM 256-char limit." }
            $matchedKey = $keyStr
            foreach ($ck in $currentMap.Keys) {
                if ([string]::Equals($ck, $keyStr, [System.StringComparison]::OrdinalIgnoreCase)) { $matchedKey = $ck; break }
            }
            $payloadTags[$matchedKey] = $val
        }
        if ($unionKeys.Count -gt 50) {
            throw "Merging tags on '$VmResourceId' would result in $($unionKeys.Count) tags, exceeding ARM's 50-tag limit."
        }
    } else {
        foreach ($k in $TagsToApply.Keys) {
            $keyStr = [string]$k
            $expectedVal = [string]$TagsToApply[$k]
            $matchedExistingKey = $null
            foreach ($ck in $currentMap.Keys) {
                if ([string]::Equals($ck, $keyStr, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $matchedExistingKey = $ck
                    break
                }
            }
            if ($null -eq $matchedExistingKey) { continue }
            $actualVal = $currentMap[$matchedExistingKey]
            if ([string]::Equals($actualVal, $expectedVal, [System.StringComparison]::Ordinal)) {
                $payloadTags[$matchedExistingKey] = $actualVal
            }
        }
        if ($payloadTags.Count -eq 0) {
            return [PSCustomObject]@{ StatusCode = 200; Skipped = $true }
        }
    }

    $uri = "https://management.azure.com$($VmResourceId.TrimEnd('/'))/providers/Microsoft.Resources/tags/default?api-version=$TagsApiVersion"
    $body = @{
        operation  = $Operation
        properties = @{
            tags = $payloadTags
        }
    }
    return Invoke-AvdPatchRestRequest -Method Patch -Uri $uri -TokenState $TokenState -Body $body
}

function New-AvdPatchGuestScript {
    <#
    .SYNOPSIS
        Generates the unified PowerShell 5.1 in-guest servicing, reboot, or verify script.
    .DESCRIPTION
        Produces a self-contained script executed locally as NT AUTHORITY\SYSTEM via
        Azure Managed Run Command across all four guest stages:
        - Service / ReadOnlyProbe: Compiles Unicode Win32 WTSEnumerateSessionsW +
          WTSQuerySessionInformationW (tolerating vanished tearing-down sessions), checks
          ProfileList hives (S-1-5-21-* and Entra S-1-12-1-*), refreshes Defender
          signatures, evaluates WUA updates via Select-AvdGuestUpdates (blocking Upgrades
          25H2/26H2 and Feature Packs, allowing RebootBehavior == 0 and Monthly B Hotpatch
          updates on VBS + AllowRebootlessUpdates hosts while deferring reboot-requiring
          Baseline LCUs when a Disconnected session is within grace), triggers Intune
          PushLaunch sync, and emits a single-line AVDPATCH_JSON sentinel.
        - ScheduleReboot: Re-checks sessions via WTSEnumerateSessionsW + ProfileList and
          registers an expiring, non-StartWhenAvailable task that re-verifies zero active
          sessions at T+20s execution time right before calling shutdown.exe /r /t 5.
        - VerifyPostReboot: Verifies LastBootUpTime advanced past BaselineBootUtc before
          unregistering AvdNightlyPatchOneShotReboot, checks UBR and CBS/WU RebootPending,
          and triggers post-boot Intune OMA-DM sync.
    .PARAMETER Stage
        Guest stage: Service, ReadOnlyProbe, ScheduleReboot, or VerifyPostReboot.
    .PARAMETER RunId
        Orchestrator run GUID embedded in the output JSON for replay protection.
    .PARAMETER PreWakeHadSession
        True if AVD Broker reported an existing session prior to or immediately after wake.
    .PARAMETER AllowRebootingUpdatesOnSession
        True if DisconnectedSessionGraceDays has expired and Baseline LCUs/reboots may proceed.
    .PARAMETER MinFreeDiskGB
        Minimum free C: disk space (GiB) required before WUA download/install.
    .PARAMETER IncludeDrivers
        True to permit Driver category updates in WUA.
    .PARAMETER IncludePreviewUpdates
        True to permit Preview-titled quality updates in WUA.
    .PARAMETER BaselineBootUtc
        Pre-reboot bootUtc ISO-8601 timestamp string used in VerifyPostReboot.
    .OUTPUTS
        System.String
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    param(
        [Parameter(Mandatory = $false)] [ValidateSet('Service', 'ReadOnlyProbe', 'ScheduleReboot', 'VerifyPostReboot')] [string] $Stage = 'Service',
        [Parameter(Mandatory = $true)]  [string] $RunId,
        [Parameter(Mandatory = $false)] [bool]   $PreWakeHadSession = $false,
        [Parameter(Mandatory = $false)] [bool]   $AllowRebootingUpdatesOnSession = $false,
        [Parameter(Mandatory = $false)] [int]    $MinFreeDiskGB = 10,
        [Parameter(Mandatory = $false)] [bool]   $IncludeDrivers = $false,
        [Parameter(Mandatory = $false)] [bool]   $IncludePreviewUpdates = $false,
        [Parameter(Mandatory = $false)] [string] $BaselineBootUtc = ''
    )
    $safeRunId = ($RunId -replace '[^A-Za-z0-9\-]', '')
    $safeBaseBoot = ($BaselineBootUtc -replace '[^A-Za-z0-9:\-\.]', '')
    $preWakeLit = if ($PreWakeHadSession) { '$true' } else { '$false' }
    $allowRbtLit = if ($AllowRebootingUpdatesOnSession) { '$true' } else { '$false' }
    $drvLit = if ($IncludeDrivers) { '$true' } else { '$false' }
    $prevLit = if ($IncludePreviewUpdates) { '$true' } else { '$false' }

    $template = @'
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$stage = '__STAGE__'
$runId = '__RUN_ID__'
$preWakeHadSession = __PRE_WAKE_SESSION__
$allowRebootOnSession = __ALLOW_REBOOT_SESSION__
$minFreeGB = __MIN_FREE_GB__
$incDrivers = __INC_DRIVERS__
$incPreviews = __INC_PREVIEWS__
$baseBootStr = '__BASE_BOOT_UTC__'

function Get-GuestBootUtcIso {
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        $lbt = $os.LastBootUpTime
        $bootLocal = if ($lbt -is [datetime]) { $lbt } else { [Management.ManagementDateTimeConverter]::ToDateTime([string]$lbt) }
        return $bootLocal.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    } catch {
        $os = Get-WmiObject -Class Win32_OperatingSystem
        return $os.ConvertToDateTime($os.LastBootUpTime).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    }
}

function Test-ServicingRebootPending {
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { return $true }
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\PackagesPending') { return $true }
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { return $true }
    try {
        $vol = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Updates' -Name 'UpdateExeVolatile' -ErrorAction SilentlyContinue).UpdateExeVolatile
        if ($null -ne $vol -and [int]$vol -ne 0) { return $true }
    } catch { $null = $_ }
    return $false
}

function Get-LocalUserSessionSummary {
    $wtsCount = 0
    $activeCount = 0
    try {
        if (-not ('AvdWtsNative' -as [type])) {
            Add-Type -TypeDefinition @"
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
public static class AvdWtsNative {
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct WTS_SESSION_INFO {
        public int SessionID;
        [MarshalAs(UnmanagedType.LPWStr)] public string pWinStationName;
        public int State;
    }
    [DllImport("wtsapi32.dll", EntryPoint = "WTSEnumerateSessionsW", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool WTSEnumerateSessionsW(IntPtr hServer, int Reserved, int Version, out IntPtr ppSessionInfo, out int pCount);
    [DllImport("wtsapi32.dll")]
    private static extern void WTSFreeMemory(IntPtr pMemory);
    [DllImport("wtsapi32.dll", EntryPoint = "WTSQuerySessionInformationW", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool WTSQuerySessionInformationW(IntPtr hServer, int sessionId, int wtsInfoClass, out IntPtr ppBuffer, out int pBytesReturned);
    public static int[] CountUserSessions() {
        IntPtr ppSessionInfo = IntPtr.Zero;
        int count = 0;
        int totalUsers = 0;
        int activeUsers = 0;
        if (!WTSEnumerateSessionsW(IntPtr.Zero, 0, 1, out ppSessionInfo, out count)) {
            throw new Win32Exception(Marshal.GetLastWin32Error());
        }
        try {
            int dataSize = Marshal.SizeOf(typeof(WTS_SESSION_INFO));
            long current = ppSessionInfo.ToInt64();
            for (int i = 0; i < count; i++) {
                WTS_SESSION_INFO si = (WTS_SESSION_INFO)Marshal.PtrToStructure(new IntPtr(current), typeof(WTS_SESSION_INFO));
                current += dataSize;
                if (si.SessionID <= 0) continue;
                if (si.State == 0 || si.State == 1 || si.State == 4) {
                    IntPtr userPtr = IntPtr.Zero;
                    int bytes = 0;
                    if (!WTSQuerySessionInformationW(IntPtr.Zero, si.SessionID, 5, out userPtr, out bytes) || userPtr == IntPtr.Zero) {
                        int err = Marshal.GetLastWin32Error();
                        if (err == 2 || err == 7022) continue;
                        throw new Win32Exception(err);
                    }
                    string user = "";
                    try { user = Marshal.PtrToStringUni(userPtr) ?? ""; } finally { WTSFreeMemory(userPtr); }
                    user = user.Trim();
                    if (!string.IsNullOrEmpty(user) &&
                        !user.StartsWith("DWM-", StringComparison.OrdinalIgnoreCase) &&
                        !user.StartsWith("UMFD-", StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(user, "SYSTEM", StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(user, "LOCAL SERVICE", StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(user, "NETWORK SERVICE", StringComparison.OrdinalIgnoreCase)) {
                        totalUsers++;
                        if (si.State == 0 || si.State == 1) activeUsers++;
                    }
                }
            }
        } finally {
            if (ppSessionInfo != IntPtr.Zero) WTSFreeMemory(ppSessionInfo);
        }
        return new int[] { totalUsers, activeUsers };
    }
}
"@ -ErrorAction Stop
        }
        $counts = [AvdWtsNative]::CountUserSessions()
        $wtsCount = $counts[0]
        $activeCount = $counts[1]
    } catch {
        $wtsCount = -1
    }

    $loadedSids = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    try {
        foreach ($hu in @(Get-ChildItem 'Registry::HKEY_USERS' -ErrorAction SilentlyContinue)) {
            if ($hu.PSChildName -match '^(S-1-5-21-|S-1-12-1-)\d+(-\d+)+$') {
                [void]$loadedSids.Add($hu.PSChildName)
            }
        }
        foreach ($p in @(Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList' -ErrorAction SilentlyContinue)) {
            if ($p.PSChildName -match '^(S-1-5-21-|S-1-12-1-)\d+(-\d+)+$') {
                $rc = (Get-ItemProperty -Path $p.PSPath -Name 'RefCount' -ErrorAction SilentlyContinue).RefCount
                if ($null -ne $rc -and [int]$rc -gt 0) { [void]$loadedSids.Add($p.PSChildName) }
            }
        }
    } catch { $null = $_ }
    $hiveCount = $loadedSids.Count
    return @{ WtsTotal = $wtsCount; WtsActive = $activeCount; LoadedHives = $hiveCount }
}

function Invoke-IntuneOmaDmSync {
    try {
        $mdmTasks = @(Get-ScheduledTask -TaskPath '\Microsoft\Windows\EnterpriseMgmt\*\' -ErrorAction SilentlyContinue |
            Where-Object { $_.TaskName -eq 'PushLaunch' -or $_.TaskName -like 'Schedule #3*' })
        if ($mdmTasks.Count -eq 0) { return $false }
        $syncStartLocal = (Get-Date).AddSeconds(-2)
        Start-ScheduledTask -InputObject $mdmTasks[0] -ErrorAction SilentlyContinue
        $spawnDeadline = [DateTime]::UtcNow.AddSeconds(15)
        $sawOma = $false
        while ([DateTime]::UtcNow -lt $spawnDeadline) {
            if (@(Get-Process -Name 'omadmclient' -ErrorAction SilentlyContinue).Count -gt 0) { $sawOma = $true; break }
            Start-Sleep -Milliseconds 400
        }
        if ($sawOma) {
            $exitDeadline = [DateTime]::UtcNow.AddSeconds(75)
            while ([DateTime]::UtcNow -lt $exitDeadline) {
                if (@(Get-Process -Name 'omadmclient' -ErrorAction SilentlyContinue).Count -eq 0) { return $true }
                Start-Sleep -Seconds 2
            }
            return $false
        }
        $ti = Get-ScheduledTaskInfo -InputObject $mdmTasks[0] -ErrorAction SilentlyContinue
        if ($null -ne $ti -and $ti.LastRunTime -ge $syncStartLocal -and $ti.LastTaskResult -eq 0) { return $true }
    } catch { $null = $_ }
    return $false
}

function Get-WuaCollItem {
    param($Coll, [int]$Idx)
    if ($Coll -is [System.Array] -or $Coll -is [System.Collections.IList]) { return $Coll[$Idx] }
    return $Coll.Item($Idx)
}

function Select-AvdGuestUpdates {
    param($UpdatesCollection, [bool]$SessionPresent, [bool]$AllowRebootOnSess, [bool]$VbsOk, [bool]$HotpatchPolOk, [bool]$DriversOk, [bool]$PreviewsOk)
    $denyCatIds = @(
        '3689bdc8-b205-4af4-8d4a-a63924c5e9d5', # Upgrades (blocks 25H2/26H2 feature updates)
        'b54e7d24-7add-428f-8b75-90a396fa584f'  # Feature Packs
    )
    if (-not $DriversOk) { $denyCatIds += 'ebfc1fc5-71a4-4f7b-9aca-3b9a503104a0' }
    $allowCatIds = @(
        '0fa1201d-4330-4fa8-8ae9-b877473b6441', # Security Updates (includes Hotpatch Monthly B)
        'e6cf1350-c01b-414d-a61f-263d14d133b4', # Critical Updates
        '28bc880e-0592-4cbf-8f95-c79b17911d5f', # Update Rollups
        'e0789628-ce08-4437-be74-2495b842f43b', # Definition Updates
        'cd5ffd1e-e932-4e3a-bf74-18bf0b1bbd83'  # Updates
    )
    if ($DriversOk) { $allowCatIds += 'ebfc1fc5-71a4-4f7b-9aca-3b9a503104a0' }

    $selected = New-Object 'System.Collections.Generic.List[object]'
    $selectedLabels = New-Object 'System.Collections.Generic.List[string]'
    $deferredLabels = New-Object 'System.Collections.Generic.List[string]'
    $uCount = [int]$UpdatesCollection.Count
    for ($i = 0; $i -lt $uCount; $i++) {
        $u = Get-WuaCollItem -Coll $UpdatesCollection -Idx $i
        if ([bool]$u.InstallationBehavior.CanRequestUserInput) { continue }
        $title = [string]$u.Title
        if (-not $PreviewsOk -and $title -match '\bPreview\b') { continue }

        $hasDenied = $false
        $hasAllowed = $false
        $catCount = [int]$u.Categories.Count
        for ($c = 0; $c -lt $catCount; $c++) {
            $catObj = Get-WuaCollItem -Coll $u.Categories -Idx $c
            $cid = ([string]$catObj.CategoryID).Trim('{}').ToLowerInvariant()
            if ($denyCatIds -contains $cid) { $hasDenied = $true }
            if ($allowCatIds -contains $cid) { $hasAllowed = $true }
        }
        if ($hasDenied -or -not $hasAllowed) { continue }

        $kbCount = [int]$u.KBArticleIDs.Count
        $firstKb = if ($kbCount -gt 0) { Get-WuaCollItem -Coll $u.KBArticleIDs -Idx 0 } else { '' }
        $kbLabel = if (-not [string]::IsNullOrWhiteSpace([string]$firstKb)) { "KB$firstKb" } else { ([string]$u.Identity.UpdateID).Substring(0, 8) }

        # RebootBehavior: 0=irbNeverReboots; 2=irbCanRequestReboot (permitted on session when VBS + AllowRebootlessUpdates + Hotpatch title)
        $rbtBeh = [int]$u.InstallationBehavior.RebootBehavior
        $isRebootless = ($rbtBeh -eq 0 -or ($rbtBeh -eq 2 -and $VbsOk -and $HotpatchPolOk -and ($title -match '\bHotpatch\b')))
        if ($SessionPresent -and -not $AllowRebootOnSess -and -not $isRebootless) {
            if ($deferredLabels.Count -lt 10) { [void]$deferredLabels.Add($kbLabel) }
            continue
        }
        [void]$selected.Add($u)
        if ($selectedLabels.Count -lt 10) { [void]$selectedLabels.Add($kbLabel) }
    }
    return @{ Selected = $selected; SelectedLabels = @($selectedLabels.ToArray()); DeferredLabels = @($deferredLabels.ToArray()) }
}

if ($stage -eq 'ScheduleReboot') {
    $out = [ordered]@{ runId = $runId; scheduled = $false; outcome = 'Unknown'; err = '' }
    try {
        $sess = Get-LocalUserSessionSummary
        if ([int]$sess.WtsTotal -lt 0) {
            $out.outcome = 'AbortedWtsQueryFailed'
        } elseif ([int]$sess.WtsActive -gt 0) {
            $out.outcome = 'AbortedActiveSessionDetected'
        } elseif (-not $allowRebootOnSession -and ([int]$sess.WtsTotal -gt 0 -or [int]$sess.LoadedHives -gt 0)) {
            $out.outcome = 'AbortedUserSessionDetected'
        } else {
            $taskName = 'AvdNightlyPatchOneShotReboot'
            Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
            $runAt = (Get-Date).AddSeconds(20)
            $allowDiscLit = if ($allowRebootOnSession) { '$true' } else { '$false' }
            $wtsBody = ${function:Get-LocalUserSessionSummary}.ToString()
            $guardScript = 'function Get-LocalUserSessionSummary {' + "`r`n" + $wtsBody + "`r`n" + '}' + "`r`n" +
                '$s = Get-LocalUserSessionSummary; if ([int]$s.WtsTotal -lt 0 -or [int]$s.WtsActive -gt 0) { exit 0 }; ' +
                'if (-not ' + $allowDiscLit + ' -and ([int]$s.WtsTotal -gt 0 -or [int]$s.LoadedHives -gt 0)) { exit 0 }; ' +
                '& "$env:SystemRoot\System32\shutdown.exe" /r /t 5 /d p:2:17 /c "AVD Nightly Baseline Quality Update Restart"'
            $encodedGuard = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($guardScript))
            $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument ('-NoProfile -NonInteractive -WindowStyle Hidden -EncodedCommand ' + $encodedGuard)
            $trigger = New-ScheduledTaskTrigger -Once -At $runAt
            $trigger.EndBoundary = $runAt.AddMinutes(5).ToString('yyyy-MM-ddTHH:mm:ss')
            $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
            $settings.DeleteExpiredTaskAfter = 'PT10M'
            Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -User 'SYSTEM' -RunLevel Highest -Settings $settings -Force | Out-Null
            $out.scheduled = $true
            $out.outcome = 'RebootScheduled'
        }
    } catch {
        $out.err = ($_.Exception.Message -replace '[\r\n]+', ' ')
        $out.outcome = 'RebootScheduleFailed'
    }
    Write-Output ('AVDPATCH_JSON:' + ($out | ConvertTo-Json -Compress) + ':END_AVDPATCH')
    return
}

if ($stage -eq 'VerifyPostReboot') {
    $out = [ordered]@{ runId = $runId; bootUtc = ''; bootAdvanced = $false; ubr = 0; postRbt = $false; mdmSynced = $false; outcome = 'Verified' }
    try {
        $out.bootUtc = Get-GuestBootUtcIso
        if ([string]::IsNullOrWhiteSpace($baseBootStr)) {
            $out.bootAdvanced = $true
        } else {
            $baseDt = [DateTimeOffset]::Parse($baseBootStr, [Globalization.CultureInfo]::InvariantCulture)
            $curDt = [DateTimeOffset]::Parse($out.bootUtc, [Globalization.CultureInfo]::InvariantCulture)
            $out.bootAdvanced = ($curDt -gt $baseDt)
        }
        if (-not $out.bootAdvanced) {
            $out.outcome = 'BootNotYetAdvanced'
        } else {
            Unregister-ScheduledTask -TaskName 'AvdNightlyPatchOneShotReboot' -Confirm:$false -ErrorAction SilentlyContinue
            try { $out.ubr = [int](Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name 'UBR' -ErrorAction SilentlyContinue).UBR } catch { $null = $_ }
            $out.postRbt = Test-ServicingRebootPending
            $out.mdmSynced = Invoke-IntuneOmaDmSync
        }
    } catch {
        $out.outcome = 'VerifyError'
    }
    Write-Output ('AVDPATCH_JSON:' + ($out | ConvertTo-Json -Compress) + ':END_AVDPATCH')
    return
}

$readOnlyProbe = ($stage -eq 'ReadOnlyProbe')
$res = [ordered]@{
    runId          = $runId
    bootUtc        = ''
    ubr            = 0
    vbsRunning     = $false
    hotpatchPolicy = $false
    pausedQuality  = $false
    freeGB         = 0
    wtsUsers       = 0
    wtsActive      = 0
    hives          = 0
    sessionPresent = $false
    preRbt         = $false
    postRbt        = $false
    defOk          = $false
    defErr         = ''
    installedKbs   = @()
    deferredKbs    = @()
    rbtReq         = $false
    mdmSynced      = $false
    outcome        = 'Unknown'
    err            = ''
}

try {
    $res.bootUtc = Get-GuestBootUtcIso
    try { $res.ubr = [int](Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name 'UBR' -ErrorAction SilentlyContinue).UBR } catch { $null = $_ }
    try {
        $dg = Get-CimInstance -ClassName Win32_DeviceGuard -Namespace 'root\Microsoft\Windows\DeviceGuard' -ErrorAction SilentlyContinue
        if ($null -ne $dg -and $dg.VirtualizationBasedSecurityStatus -eq 2) { $res.vbsRunning = $true }
    } catch { $null = $_ }
    try {
        $hpPol = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Update' -Name 'AllowRebootlessUpdates' -ErrorAction SilentlyContinue).AllowRebootlessUpdates
        if ($null -ne $hpPol -and [int]$hpPol -eq 1) { $res.hotpatchPolicy = $true }
    } catch { $null = $_ }
    foreach ($pausePath in @('HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate', 'HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Update')) {
        try {
            $pq = (Get-ItemProperty -Path $pausePath -Name 'PauseQualityUpdates' -ErrorAction SilentlyContinue).PauseQualityUpdates
            if ($null -ne $pq -and [int]$pq -eq 1) { $res.pausedQuality = $true }
        } catch { $null = $_ }
    }

    $sysDrive = Get-PSDrive -Name 'C' -ErrorAction Stop
    $res.freeGB = [int][Math]::Floor($sysDrive.Free / 1GB)
    $sess = Get-LocalUserSessionSummary
    $res.wtsUsers = [int]$sess.WtsTotal
    $res.wtsActive = [int]$sess.WtsActive
    $res.hives = [int]$sess.LoadedHives
    $res.sessionPresent = ($preWakeHadSession -or $res.wtsUsers -ne 0 -or $res.hives -gt 0)
    $res.preRbt = Test-ServicingRebootPending

    if (-not $readOnlyProbe) {
        try {
            $wd = Get-Service -Name 'WinDefend' -ErrorAction SilentlyContinue
            if ($null -ne $wd -and $wd.Status -eq 'Running' -and (Get-Command -Name 'Update-MpSignature' -ErrorAction SilentlyContinue)) {
                Update-MpSignature -ErrorAction Stop | Out-Null
                $res.defOk = $true
            }
        } catch {
            $dMsg = ($_.Exception.Message -replace '[\r\n]+', ' ').Trim()
            if ($dMsg.Length -gt 80) { $dMsg = $dMsg.Substring(0, 80) }
            $res.defErr = $dMsg
        }
    }

    $hasMockWua = [bool](Get-Command -Name 'Get-AvdMockWuaSession' -ErrorAction SilentlyContinue)
    if ($readOnlyProbe) {
        try {
            $wuProbe = if ($hasMockWua) { Get-AvdMockWuaSession } else { New-Object -ComObject 'Microsoft.Update.Session' }
            $srch = $wuProbe.CreateUpdateSearcher()
            $srch.ServerSelection = 0
            $criteria = if ($incDrivers) { "IsInstalled=0 and IsHidden=0 and BrowseOnly=0" } else { "IsInstalled=0 and Type='Software' and IsHidden=0 and BrowseOnly=0" }
            $sr = $srch.Search($criteria)
            $sel = Select-AvdGuestUpdates -UpdatesCollection $sr.Updates -SessionPresent $res.sessionPresent -AllowRebootOnSess $allowRebootOnSession -VbsOk $res.vbsRunning -HotpatchPolOk $res.hotpatchPolicy -DriversOk $incDrivers -PreviewsOk $incPreviews
            $res.installedKbs = $sel.SelectedLabels
            $res.deferredKbs = $sel.DeferredLabels
            $res.outcome = 'LocalProbeCompleted'
        } catch {
            $res.err = ($_.Exception.Message -replace '[\r\n]+', ' ')
            $res.outcome = 'LocalProbeWuaOffline'
        }
    } elseif ($res.wtsActive -gt 0) {
        $res.outcome = 'SkippedActiveGuestSession'
    } elseif ($res.pausedQuality) {
        $res.outcome = 'SkippedAutopatchPaused'
    } elseif ($res.freeGB -lt $minFreeGB) {
        $res.outcome = 'InsufficientDiskSpace'
    } elseif ($res.preRbt -and -not $res.sessionPresent) {
        $res.rbtReq = $true
        $res.outcome = 'PreExistingRebootPending'
    } else {
        $wuSession = if ($hasMockWua) { Get-AvdMockWuaSession } else { New-Object -ComObject 'Microsoft.Update.Session' }
        $wuSession.ClientApplicationID = 'AvdPersonalNightlyPatch'
        $installerProbe = $wuSession.CreateUpdateInstaller()
        if ([bool]$installerProbe.IsBusy) {
            $res.outcome = 'SkippedInstallerBusy'
        } else {
            $searcher = $wuSession.CreateUpdateSearcher()
            $searcher.ServerSelection = 0
            $criteria = if ($incDrivers) { "IsInstalled=0 and IsHidden=0 and BrowseOnly=0" } else { "IsInstalled=0 and Type='Software' and IsHidden=0 and BrowseOnly=0" }
            $searchResult = $searcher.Search($criteria)
            $sel = Select-AvdGuestUpdates -UpdatesCollection $searchResult.Updates -SessionPresent $res.sessionPresent -AllowRebootOnSess $allowRebootOnSession -VbsOk $res.vbsRunning -HotpatchPolOk $res.hotpatchPolicy -DriversOk $incDrivers -PreviewsOk $incPreviews
            $res.deferredKbs = $sel.DeferredLabels

            if ($sel.Selected.Count -eq 0) {
                $res.outcome = if ($res.deferredKbs.Count -gt 0) { 'DeferredBaselineForUserSession' } else { 'UpToDate' }
            } else {
                $toInstall = if ($hasMockWua) { New-Object 'System.Collections.Generic.List[object]' } else { New-Object -ComObject 'Microsoft.Update.UpdateColl' }
                foreach ($u in $sel.Selected) {
                    if (-not [bool]$u.EulaAccepted) { $u.AcceptEula() }
                    [void]$toInstall.Add($u)
                }
                $downloader = $wuSession.CreateUpdateDownloader()
                $downloader.Updates = $toInstall
                $dlRes = $downloader.Download()

                $readyColl = if ($hasMockWua) { New-Object 'System.Collections.Generic.List[object]' } else { New-Object -ComObject 'Microsoft.Update.UpdateColl' }
                for ($j = 0; $j -lt $toInstall.Count; $j++) {
                    $candItem = Get-WuaCollItem -Coll $toInstall -Idx $j
                    if ([bool]$candItem.IsDownloaded) { [void]$readyColl.Add($candItem) }
                }
                if ($readyColl.Count -eq 0) {
                    $res.outcome = "DownloadFailed_$($dlRes.ResultCode)"
                } else {
                    $installer = $wuSession.CreateUpdateInstaller()
                    $installer.Updates = $readyColl
                    $instRes = $installer.Install()
                    $res.rbtReq = [bool]$instRes.RebootRequired
                    $failedUnitCount = ($toInstall.Count - $readyColl.Count)
                    for ($k = 0; $k -lt $readyColl.Count; $k++) {
                        $ur = $instRes.GetUpdateResult($k)
                        if ([int]$ur.ResultCode -eq 2) {
                            if ($res.installedKbs.Count -lt 10) {
                                $uItem = Get-WuaCollItem -Coll $readyColl -Idx $k
                                $uKbCnt = [int]$uItem.KBArticleIDs.Count
                                $uFirstKb = if ($uKbCnt -gt 0) { Get-WuaCollItem -Coll $uItem.KBArticleIDs -Idx 0 } else { '' }
                                $kb = if (-not [string]::IsNullOrWhiteSpace([string]$uFirstKb)) { "KB$uFirstKb" } else { ([string]$uItem.Identity.UpdateID).Substring(0, 8) }
                                $res.installedKbs += $kb
                            }
                        } else {
                            $failedUnitCount++
                        }
                    }
                    if ([int]$instRes.ResultCode -eq 2 -and $failedUnitCount -eq 0) {
                        if (-not $res.rbtReq) {
                            $res.outcome = if ($res.deferredKbs.Count -gt 0) { 'HotpatchOrDefInstalled_BaselineDeferred' } else { 'InstalledNoReboot' }
                        } else {
                            $res.outcome = 'InstalledRebootRequired'
                        }
                    } elseif ($res.installedKbs.Count -gt 0) {
                        $res.outcome = if ($res.rbtReq) { 'PartialInstallRebootRequired' } else { 'PartialInstallWarning' }
                    } else {
                        $res.outcome = "InstallFailed_$($instRes.ResultCode)"
                    }
                }
            }
        }
    }

    $res.postRbt = ($res.rbtReq -or (Test-ServicingRebootPending))
    if (-not $readOnlyProbe) {
        $res.mdmSynced = Invoke-IntuneOmaDmSync
    }
} catch {
    $msg = ($_.Exception.Message -replace '[\r\n]+', ' ').Trim()
    if ($msg.Length -gt 200) { $msg = $msg.Substring(0, 200) }
    $res.err = $msg
    $res.outcome = 'GuestScriptException'
}

$json = $res | ConvertTo-Json -Compress -Depth 4
Write-Output ("AVDPATCH_JSON:" + $json + ":END_AVDPATCH")
'@
    return $template.
        Replace('__STAGE__', $Stage).
        Replace('__RUN_ID__', $safeRunId).
        Replace('__PRE_WAKE_SESSION__', $preWakeLit).
        Replace('__ALLOW_REBOOT_SESSION__', $allowRbtLit).
        Replace('__MIN_FREE_GB__', [string]$MinFreeDiskGB).
        Replace('__INC_DRIVERS__', $drvLit).
        Replace('__INC_PREVIEWS__', $prevLit).
        Replace('__BASE_BOOT_UTC__', $safeBaseBoot)
}

function ConvertFrom-AvdPatchRunCommandOutput {
    <#
    .SYNOPSIS
        Extracts and validates the AVDPATCH_JSON sentinel payload from RunCommand stdout.
    .DESCRIPTION
        Handles Azure Run Command's last-4,096-byte stdout buffer with arbitrary noisy
        prefix output, validates the :END_AVDPATCH sentinel and ExpectedRunId, and returns
        the parsed PSCustomObject or $null on mismatch/truncation.
    .PARAMETER RawOutput
        Managed Run Command instanceView.output string.
    .PARAMETER ExpectedRunId
        Expected RunId GUID to prevent stale RunCommand replay.
    .OUTPUTS
        PSCustomObject
    #>
    param(
        [Parameter(Mandatory = $false)] [string] $RawOutput,
        [Parameter(Mandatory = $true)]  [string] $ExpectedRunId
    )
    if ([string]::IsNullOrWhiteSpace($RawOutput)) { return $null }
    $m = [regex]::Match($RawOutput, 'AVDPATCH_JSON:(\{.*?\}):END_AVDPATCH', [System.Text.RegularExpressions.RegexOptions]::Singleline)
    if (-not $m.Success) { return $null }
    try {
        $parsed = $m.Groups[1].Value | ConvertFrom-Json
        $rid = [string](Get-AvdPatchProp -InputObject $parsed -Name 'runId')
        if (-not [string]::Equals($rid, $ExpectedRunId, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $null
        }
        return $parsed
    } catch {
        return $null
    }
}

function Set-AvdPatchSessionHostDrain {
    <#
    .SYNOPSIS
        Updates allowNewSession on an AVD Session Host resource via ARM PATCH.
    .DESCRIPTION
        Calls Microsoft.DesktopVirtualization/hostPools/sessionHosts PATCH to toggle
        drain mode without altering assignedUser or other session host properties.
    .PARAMETER SessionHostArmId
        Full ARM resource ID of the sessionHost resource.
    .PARAMETER AllowNewSession
        Target boolean value for properties.allowNewSession.
    .PARAMETER ApiVersion
        Microsoft.DesktopVirtualization API version.
    .PARAMETER TokenState
        ARM authentication token state hashtable.
    .OUTPUTS
        PSCustomObject
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    param(
        [Parameter(Mandatory = $true)] [string]    $SessionHostArmId,
        [Parameter(Mandatory = $true)] [bool]      $AllowNewSession,
        [Parameter(Mandatory = $true)] [string]    $ApiVersion,
        [Parameter(Mandatory = $true)] [hashtable] $TokenState
    )
    $uri = "https://management.azure.com$($SessionHostArmId.TrimEnd('/'))?api-version=$ApiVersion"
    $body = @{ properties = @{ allowNewSession = $AllowNewSession } }
    return Invoke-AvdPatchRestRequest -Method Patch -Uri $uri -TokenState $TokenState -Body $body
}

function Get-AvdPatchBrokerSessionInfo {
    <#
    .SYNOPSIS
        Queries AVD userSessions for a session host and classifies Active vs Disconnected state.
    .DESCRIPTION
        Enumerates Microsoft.DesktopVirtualization userSessions and counts Active/Pending
        sessions separately from Disconnected sessions so active developers are never disturbed.
    .PARAMETER SessionHostArmId
        Full ARM resource ID of the sessionHost resource.
    .PARAMETER ApiVersion
        Microsoft.DesktopVirtualization API version.
    .PARAMETER TokenState
        ARM authentication token state hashtable.
    .OUTPUTS
        PSCustomObject
    #>
    param(
        [Parameter(Mandatory = $true)] [string]    $SessionHostArmId,
        [Parameter(Mandatory = $true)] [string]    $ApiVersion,
        [Parameter(Mandatory = $true)] [hashtable] $TokenState
    )
    $uri = "https://management.azure.com$($SessionHostArmId.TrimEnd('/'))/userSessions?api-version=$ApiVersion"
    $sessions = @(Invoke-AvdPatchArmPagedGet -Uri $uri -TokenState $TokenState)
    $activeCount = 0
    $discCount = 0
    foreach ($s in $sessions) {
        $p = Get-AvdPatchProp -InputObject $s -Name 'properties'
        $st = [string](Get-AvdPatchProp -InputObject $p -Name 'sessionState')
        if ($st -in @('Active', 'Pending', 'UserProfileDiskMounted')) {
            $activeCount++
        } elseif ([string]::Equals($st, 'Disconnected', [System.StringComparison]::OrdinalIgnoreCase)) {
            $discCount++
        }
    }
    return [PSCustomObject]@{
        TotalCount        = ($activeCount + $discCount)
        ActiveCount       = $activeCount
        DisconnectedCount = $discCount
    }
}

function Invoke-AvdPatchManagedRunCommand {
    <#
    .SYNOPSIS
        Manages Azure VM Managed Run Command lifecycle (Put, Get, Delete, InspectAvd, PruneStale).
    .DESCRIPTION
        Submits asynchronous SYSTEM RunCommands, polls instanceView executionState, deletes
        completed commands, and inspects or prunes orphaned AvdPatch-*/AvdReboot-*/AvdVerify-*
        resources so VMs never exhaust Azure's 25-Managed-Run-Command per-VM quota.
    .PARAMETER Operation
        Put, Get, Delete, InspectAvd (read-only), or PruneStale (deletes terminal Avd* commands).
    .PARAMETER VmResourceId
        Full ARM resource ID of the virtual machine.
    .PARAMETER Location
        Azure region of the virtual machine (required for Put).
    .PARAMETER RunCommandName
        Managed Run Command child resource name (required for Put, Get, Delete).
    .PARAMETER ScriptText
        PowerShell 5.1 script body to execute as SYSTEM (required for Put).
    .PARAMETER TimeoutSeconds
        Script execution timeout in seconds for Put. Default: 2700.
    .PARAMETER ComputeApiVersion
        Microsoft.Compute API version.
    .PARAMETER TokenState
        ARM authentication token state hashtable.
    .OUTPUTS
        PSCustomObject
    #>
    param(
        [Parameter(Mandatory = $true)]  [ValidateSet('Put', 'Get', 'Delete', 'InspectAvd', 'PruneStale')] [string] $Operation,
        [Parameter(Mandatory = $true)]  [string]    $VmResourceId,
        [Parameter(Mandatory = $false)] [string]    $Location = '',
        [Parameter(Mandatory = $false)] [string]    $RunCommandName = '',
        [Parameter(Mandatory = $false)] [string]    $ScriptText = '',
        [Parameter(Mandatory = $false)] [int]       $TimeoutSeconds = 2700,
        [Parameter(Mandatory = $true)]  [string]    $ComputeApiVersion,
        [Parameter(Mandatory = $true)]  [hashtable] $TokenState
    )
    $trimmedVm = $VmResourceId.TrimEnd('/')
    switch ($Operation) {
        'Put' {
            $uri = "https://management.azure.com$trimmedVm/runCommands/$RunCommandName`?api-version=$ComputeApiVersion"
            $body = @{
                location   = $Location
                properties = @{
                    asyncExecution                  = $true
                    treatFailureAsDeploymentFailure = $false
                    timeoutInSeconds                = $TimeoutSeconds
                    source                          = @{ script = $ScriptText }
                }
            }
            return Invoke-AvdPatchRestRequest -Method Put -Uri $uri -TokenState $TokenState -Body $body
        }
        'Get' {
            $uri = "https://management.azure.com$trimmedVm/runCommands/$RunCommandName`?`$expand=instanceView&api-version=$ComputeApiVersion"
            $res = Invoke-AvdPatchRestRequest -Method Get -Uri $uri -TokenState $TokenState -AllowNotFound
            if ($res.StatusCode -eq 404 -or $null -eq $res.Body) {
                return [PSCustomObject]@{ Exists = $false; ExecutionState = 'NotFound'; ProvisioningState = ''; ExitCode = -1; Output = ''; Error = '' }
            }
            $props = Get-AvdPatchProp -InputObject $res.Body -Name 'properties'
            $prov = [string](Get-AvdPatchProp -InputObject $props -Name 'provisioningState')
            $iv = Get-AvdPatchProp -InputObject $props -Name 'instanceView'
            $execState = [string](Get-AvdPatchProp -InputObject $iv -Name 'executionState')
            if ([string]::IsNullOrWhiteSpace($execState)) {
                $execState = if ($prov -in @('Creating', 'Updating', 'Running')) { 'Running' } elseif ($prov -eq 'Failed') { 'Failed' } else { 'Pending' }
            }
            $exitRaw = Get-AvdPatchProp -InputObject $iv -Name 'exitCode'
            return [PSCustomObject]@{
                Exists            = $true
                ExecutionState    = $execState
                ProvisioningState = $prov
                ExitCode          = if ($null -ne $exitRaw) { [int]$exitRaw } else { 0 }
                Output            = [string](Get-AvdPatchProp -InputObject $iv -Name 'output')
                Error             = [string](Get-AvdPatchProp -InputObject $iv -Name 'error')
            }
        }
        'Delete' {
            $uri = "https://management.azure.com$trimmedVm/runCommands/$RunCommandName`?api-version=$ComputeApiVersion"
            return Invoke-AvdPatchRestRequest -Method Delete -Uri $uri -TokenState $TokenState -AllowNotFound
        }
        default {
            $listUri = "https://management.azure.com$trimmedVm/runCommands?`$expand=instanceView&api-version=$ComputeApiVersion"
            $cmds = @()
            try {
                $cmds = @(Invoke-AvdPatchArmPagedGet -Uri $listUri -TokenState $TokenState)
            } catch {
                $httpCode = if ($null -ne $_.Exception.Data -and $_.Exception.Data.Contains('HttpStatusCode')) { [int]$_.Exception.Data['HttpStatusCode'] } else { 0 }
                if ($httpCode -eq 404) { $cmds = @() } else { throw }
            }
            $runningNames = New-Object 'System.Collections.Generic.List[string]'
            $staleNames = New-Object 'System.Collections.Generic.List[string]'
            foreach ($c in $cmds) {
                $cName = (([string](Get-AvdPatchProp -InputObject $c -Name 'name')) -split '/')[-1]
                if ($cName -notmatch '^Avd(Patch|Reboot|Verify)-') { continue }
                $cProps = Get-AvdPatchProp -InputObject $c -Name 'properties'
                $cProv = [string](Get-AvdPatchProp -InputObject $cProps -Name 'provisioningState')
                $cIv = Get-AvdPatchProp -InputObject $cProps -Name 'instanceView'
                $cExec = [string](Get-AvdPatchProp -InputObject $cIv -Name 'executionState')
                if ([string]::IsNullOrWhiteSpace($cExec)) {
                    $cExec = if ($cProv -in @('Creating', 'Updating', 'Running')) { 'Running' } else { 'Succeeded' }
                }
                if ($cExec -in @('Running', 'Pending')) {
                    [void]$runningNames.Add($cName)
                } else {
                    [void]$staleNames.Add($cName)
                    if ($Operation -eq 'PruneStale') {
                        $delUri = "https://management.azure.com$trimmedVm/runCommands/$cName`?api-version=$ComputeApiVersion"
                        $null = Invoke-AvdPatchRestRequest -Method Delete -Uri $delUri -TokenState $TokenState -AllowNotFound
                    }
                }
            }
            return [PSCustomObject]@{
                HasRunning   = ($runningNames.Count -gt 0)
                RunningNames = @($runningNames.ToArray())
                StaleNames   = @($staleNames.ToArray())
            }
        }
    }
}

function Test-AvdPatchWakeCutoffReached {
    <#
    .SYNOPSIS
        Checks whether the local maintenance-window cutoff hour has been reached.
    .DESCRIPTION
        Resolves TimeZoneId (throwing explicitly on invalid IDs rather than silently
        falling back to UTC) and returns $true when local time is within the cutoff window.
    .PARAMETER NowUtc
        Current UTC timestamp.
    .PARAMETER CutoffHour
        Local hour (0..23) to stop new wakes, or -1 when disabled.
    .PARAMETER TimeZoneId
        Windows TimeZoneInfo identifier.
    .OUTPUTS
        System.Boolean
    #>
    param(
        [Parameter(Mandatory = $true)] [DateTimeOffset] $NowUtc,
        [Parameter(Mandatory = $true)] [int]    $CutoffHour,
        [Parameter(Mandatory = $true)] [string] $TimeZoneId
    )
    if ($CutoffHour -lt 0) { return $false }
    $tz = [System.TimeZoneInfo]::FindSystemTimeZoneById($TimeZoneId)
    $local = [System.TimeZoneInfo]::ConvertTime($NowUtc, $tz)
    if ($CutoffHour -le 12) {
        return ($local.Hour -ge $CutoffHour -and $local.Hour -lt 18)
    }
    return ($local.Hour -ge $CutoffHour)
}

function New-AvdPatchHostResult {
    <#
    .SYNOPSIS
        Constructs a standardized output record for an evaluated Personal Session Host.
    .DESCRIPTION
        Emits consistent schema fields across Audit, Triage, MorningUnlock, LocalGuestProbe,
        and Execute runs for pipeline logging, KQL ingestion, or table formatting.
    .PARAMETER HostPoolName
        Target AVD Personal Host Pool name.
    .PARAMETER ShortHostName
        Short VM / session host name.
    .PARAMETER AssignedUser
        UPN assigned to the personal desktop, if any.
    .PARAMETER PriorPowerState
        Initial VM power state (Hibernated, Deallocated, Running, StoppedAllocated).
    .PARAMETER FinalPowerState
        Final VM power state at end of processing.
    .PARAMETER SessionDetected
        True if a user session was detected via Broker or Guest WTS/ProfileList.
    .PARAMETER HotpatchEligible
        True if VBS and AllowRebootlessUpdates were confirmed running in guest.
    .PARAMETER InstalledKbs
        Comma-joined list of installed KB identifiers.
    .PARAMETER DeferredKbs
        Comma-joined list of reboot-requiring KBs deferred due to user session.
    .PARAMETER RebootPerformed
        True if a coordinated OS restart occurred during this run.
    .PARAMETER Status
        Machine-readable terminal status code.
    .PARAMETER Detail
        Human-readable summary or diagnostic message.
    .OUTPUTS
        PSCustomObject
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    param(
        [Parameter(Mandatory = $true)]  [string] $HostPoolName,
        [Parameter(Mandatory = $true)]  [string] $ShortHostName,
        [Parameter(Mandatory = $false)] [AllowEmptyString()] [string] $AssignedUser = '',
        [Parameter(Mandatory = $false)] [string] $PriorPowerState = 'N/A',
        [Parameter(Mandatory = $false)] [string] $FinalPowerState = '',
        [Parameter(Mandatory = $false)] [bool]   $SessionDetected = $false,
        [Parameter(Mandatory = $false)] [bool]   $HotpatchEligible = $false,
        [Parameter(Mandatory = $false)] [AllowEmptyString()] [string] $InstalledKbs = '',
        [Parameter(Mandatory = $false)] [AllowEmptyString()] [string] $DeferredKbs = '',
        [Parameter(Mandatory = $false)] [bool]   $RebootPerformed = $false,
        [Parameter(Mandatory = $true)]  [string] $Status,
        [Parameter(Mandatory = $false)] [AllowEmptyString()] [string] $Detail = ''
    )
    return [PSCustomObject]@{
        HostPoolName     = $HostPoolName
        ShortHostName    = $ShortHostName
        AssignedUser     = $AssignedUser
        PriorPowerState  = $PriorPowerState
        FinalPowerState  = if ([string]::IsNullOrWhiteSpace($FinalPowerState)) { $PriorPowerState } else { $FinalPowerState }
        SessionDetected  = $SessionDetected
        HotpatchEligible = $HotpatchEligible
        InstalledKbs     = $InstalledKbs
        DeferredKbs      = $DeferredKbs
        RebootPerformed  = $RebootPerformed
        Status           = $Status
        Detail           = $Detail
    }
}

# ==============================================================================
# MAIN ORCHESTRATION ENTRYPOINT
# ==============================================================================
if ($StopNewWakeAtLocalHour -ge 0) {
    $null = [System.TimeZoneInfo]::FindSystemTimeZoneById($MaintenanceTimeZoneId)
}

if ($Mode -eq 'LocalGuestProbe') {
    $probeRun = [guid]::NewGuid().ToString().ToLowerInvariant()
    $probeScript = New-AvdPatchGuestScript -Stage 'ReadOnlyProbe' -RunId $probeRun -MinFreeDiskGB $MinFreeDiskSpaceGB -IncludeDrivers ([bool]$IncludeDrivers) -IncludePreviewUpdates ([bool]$IncludePreviewUpdates)
    $probeOut = (& ([scriptblock]::Create($probeScript)) | Out-String)
    $parsedProbe = ConvertFrom-AvdPatchRunCommandOutput -RawOutput $probeOut -ExpectedRunId $probeRun
    if ($null -eq $parsedProbe) {
        throw "LocalGuestProbe failed to parse output: $probeOut"
    }
    $instProbe = (@(Get-AvdPatchProp -InputObject $parsedProbe -Name 'installedKbs') -join ',')
    $defProbe = (@(Get-AvdPatchProp -InputObject $parsedProbe -Name 'deferredKbs') -join ',')
    $hpProbe = ([bool](Get-AvdPatchProp -InputObject $parsedProbe -Name 'vbsRunning') -and [bool](Get-AvdPatchProp -InputObject $parsedProbe -Name 'hotpatchPolicy'))
    $detProbe = "UBR=$(Get-AvdPatchProp -InputObject $parsedProbe -Name 'ubr') FreeGB=$(Get-AvdPatchProp -InputObject $parsedProbe -Name 'freeGB') WtsUsers=$(Get-AvdPatchProp -InputObject $parsedProbe -Name 'wtsUsers') Hives=$(Get-AvdPatchProp -InputObject $parsedProbe -Name 'hives') VBS=$(Get-AvdPatchProp -InputObject $parsedProbe -Name 'vbsRunning') HotpatchPolicy=$(Get-AvdPatchProp -InputObject $parsedProbe -Name 'hotpatchPolicy') PreRbt=$(Get-AvdPatchProp -InputObject $parsedProbe -Name 'preRbt')"
    return @((New-AvdPatchHostResult -HostPoolName 'LOCAL' -ShortHostName $env:COMPUTERNAME -AssignedUser '' -PriorPowerState 'Running' -FinalPowerState 'Running' -SessionDetected ([bool](Get-AvdPatchProp -InputObject $parsedProbe -Name 'sessionPresent')) -HotpatchEligible $hpProbe -InstalledKbs $instProbe -DeferredKbs $defProbe -Status ([string](Get-AvdPatchProp -InputObject $parsedProbe -Name 'outcome')) -Detail $detProbe))
}

$targetPools = New-Object 'System.Collections.Generic.List[string]'
if ($null -ne $HostPoolResourceId -and @($HostPoolResourceId).Count -gt 0) {
    foreach ($id in $HostPoolResourceId) {
        if (-not [string]::IsNullOrWhiteSpace($id)) { [void]$targetPools.Add($id.Trim()) }
    }
} elseif (-not [string]::IsNullOrWhiteSpace($ResourceGroupName) -and -not [string]::IsNullOrWhiteSpace($HostPoolName)) {
    $resolvedSub = $SubscriptionId
    if ([string]::IsNullOrWhiteSpace($resolvedSub) -and (Get-Command -Name 'Get-AzContext' -ErrorAction SilentlyContinue)) {
        $ctx = Get-AzContext -ErrorAction SilentlyContinue
        if ($null -ne $ctx -and $null -ne $ctx.Subscription) { $resolvedSub = [string]$ctx.Subscription.Id }
    }
    if ([string]::IsNullOrWhiteSpace($resolvedSub)) {
        throw 'Specify -HostPoolResourceId, or provide -SubscriptionId together with -ResourceGroupName and -HostPoolName.'
    }
    [void]$targetPools.Add("/subscriptions/$($resolvedSub.Trim())/resourceGroups/$($ResourceGroupName.Trim())/providers/Microsoft.DesktopVirtualization/hostPools/$($HostPoolName.Trim())")
} else {
    throw 'Provide either -HostPoolResourceId or (-ResourceGroupName and -HostPoolName).'
}

$tokenState = @{
    AuthMode    = $AuthMode
    ClientId    = $ManagedIdentityClientId
    StaticToken = $ArmAccessToken
    Token       = $null
    ExpiresUtc  = $null
}

$runGuid = [guid]::NewGuid().ToString().ToLowerInvariant()
$shortRun = ($runGuid -replace '-', '').Substring(0, 8)
$allResults = New-Object 'System.Collections.Generic.List[object]'
$targetFilterSet = $null
if ($null -ne $TargetSessionHostName -and @($TargetSessionHostName).Count -gt 0) {
    $targetFilterSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($th in $TargetSessionHostName) {
        if (-not [string]::IsNullOrWhiteSpace($th)) {
            [void]$targetFilterSet.Add(($th.Trim() -split '\.')[0])
        }
    }
}

foreach ($poolArmId in $targetPools) {
    $poolResp = Invoke-AvdPatchRestRequest -Method Get -Uri "https://management.azure.com$poolArmId`?api-version=$DesktopVirtApiVersion" -TokenState $tokenState
    $poolName = [string](Get-AvdPatchProp -InputObject $poolResp.Body -Name 'name')
    $poolProps = Get-AvdPatchProp -InputObject $poolResp.Body -Name 'properties'
    $poolType = [string](Get-AvdPatchProp -InputObject $poolProps -Name 'hostPoolType')
    if (-not [string]::Equals($poolType, 'Personal', [System.StringComparison]::OrdinalIgnoreCase)) {
        [void]$allResults.Add((New-AvdPatchHostResult -HostPoolName $poolName -ShortHostName '*' -AssignedUser '' -PriorPowerState 'N/A' -FinalPowerState 'N/A' -Status 'SkippedNonPersonalPool' -Detail "HostPoolType is '$poolType'; script only targets Personal pools."))
        continue
    }

    # Skip ScalingPlan lookup in MorningUnlock & Triage so a 403 on scalingPlans never blocks emergency recovery
    $exclCheck = [PSCustomObject]@{ IsValid = $true; HasActivePlan = $false; ScalingPlanName = ''; ExclusionTag = 'excludeFromScaling'; FailureReason = '' }
    if ($Mode -in @('Audit', 'Execute')) {
        try {
            $exclCheck = Resolve-AvdPatchPoolExclusionTag -HostPoolArmId $poolArmId -ScalingPlanArmId $ScalingPlanResourceId -ExpectedExclusionTag $ExclusionTagName -AllowNoScalingPlan ([bool]$AllowNoScalingPlan) -ApiVersion $DesktopVirtApiVersion -TokenState $tokenState
        } catch {
            $exclCheck = [PSCustomObject]@{ IsValid = $false; HasActivePlan = $false; ScalingPlanName = ''; ExclusionTag = ''; FailureReason = "Scaling Plan lookup failed: $($_.Exception.Message)" }
        }
        if (-not $exclCheck.IsValid -and $Mode -eq 'Execute') {
            [void]$allResults.Add((New-AvdPatchHostResult -HostPoolName $poolName -ShortHostName '*' -AssignedUser '' -PriorPowerState 'N/A' -FinalPowerState 'N/A' -Status 'ScalingPlanExclusionTagMissingOrMismatch' -Detail $exclCheck.FailureReason))
            continue
        }
    }
    $exclusionTagKey = if (-not [string]::IsNullOrWhiteSpace($exclCheck.ExclusionTag)) { $exclCheck.ExclusionTag } else { 'excludeFromScaling' }

    $sessionHosts = @(Invoke-AvdPatchArmPagedGet -Uri "https://management.azure.com$poolArmId/sessionHosts?api-version=$DesktopVirtApiVersion" -TokenState $tokenState)
    $workQueue = New-Object 'System.Collections.Generic.Queue[hashtable]'

    foreach ($sh in $sessionHosts) {
        $shFullName = [string](Get-AvdPatchProp -InputObject $sh -Name 'name')
        $leafName = ($shFullName -split '/')[-1]
        $shortHost = ($leafName -split '\.')[0]
        if ($null -ne $targetFilterSet -and -not $targetFilterSet.Contains($shortHost)) { continue }

        $shId = [string](Get-AvdPatchProp -InputObject $sh -Name 'id')
        $shProps = Get-AvdPatchProp -InputObject $sh -Name 'properties'
        $assignedUser = [string](Get-AvdPatchProp -InputObject $shProps -Name 'assignedUser')
        $allowNewSess = [bool](Get-AvdPatchProp -InputObject $shProps -Name 'allowNewSession')
        $vmArmId = [string](Get-AvdPatchProp -InputObject $shProps -Name 'resourceId')

        if ([string]::IsNullOrWhiteSpace($vmArmId)) {
            [void]$allResults.Add((New-AvdPatchHostResult -HostPoolName $poolName -ShortHostName $shortHost -AssignedUser $assignedUser -PriorPowerState 'MissingVmId' -FinalPowerState 'MissingVmId' -Status 'SkippedOrphanedSessionHost' -Detail 'SessionHost has no resourceId.'))
            continue
        }

        $vmResp = $null
        try {
            $vmResp = Invoke-AvdPatchRestRequest -Method Get -Uri "https://management.azure.com$vmArmId`?`$expand=instanceView&api-version=$ComputeApiVersion" -TokenState $tokenState -AllowNotFound
        } catch {
            if ($Mode -eq 'MorningUnlock') {
                [void]$allResults.Add((New-AvdPatchHostResult -HostPoolName $poolName -ShortHostName $shortHost -AssignedUser $assignedUser -PriorPowerState 'Unknown' -FinalPowerState 'Unknown' -Status 'MorningUnlockError' -Detail $_.Exception.Message))
                continue
            }
            throw
        }
        if ($vmResp.StatusCode -eq 404 -or $null -eq $vmResp.Body) {
            [void]$allResults.Add((New-AvdPatchHostResult -HostPoolName $poolName -ShortHostName $shortHost -AssignedUser $assignedUser -PriorPowerState 'VmNotFound' -FinalPowerState 'VmNotFound' -Status 'SkippedOrphanedSessionHost' -Detail "Backing VM '$vmArmId' returned 404."))
            continue
        }

        $vmObj = $vmResp.Body
        $vmLocation = [string](Get-AvdPatchProp -InputObject $vmObj -Name 'location')
        $vmProps = Get-AvdPatchProp -InputObject $vmObj -Name 'properties'
        $vmGuid = [string](Get-AvdPatchProp -InputObject $vmProps -Name 'vmId')
        $vmTags = Get-AvdPatchProp -InputObject $vmObj -Name 'tags'
        $pwr = Get-AvdPatchVmPowerSummary -VmResource $vmObj
        $nowUtc = Get-AvdPatchNowUtc

        # Sibling safety tags (aligned identically with Remove-InactiveAvdPersonalHosts.ps1)
        $legalHoldVal = [string](Get-AvdPatchProp -InputObject $vmTags -Name 'AvdLegalHold')
        if (-not [string]::IsNullOrWhiteSpace($legalHoldVal) -and $legalHoldVal.Trim().ToLowerInvariant() -notin @('false', '0', 'no')) {
            [void]$allResults.Add((New-AvdPatchHostResult -HostPoolName $poolName -ShortHostName $shortHost -AssignedUser $assignedUser -PriorPowerState $pwr.EffectiveState -FinalPowerState $pwr.EffectiveState -Status 'SkippedLegalHold' -Detail "VM has AvdLegalHold='$($legalHoldVal.Trim())'."))
            continue
        }
        $lifecycleVal = ([string](Get-AvdPatchProp -InputObject $vmTags -Name 'AvdLifecycleState')).Trim()
        $isQuarantined = [string]::Equals($lifecycleVal, 'Quarantined', [System.StringComparison]::OrdinalIgnoreCase)
        if ($isQuarantined -and -not $IncludeQuarantined) {
            [void]$allResults.Add((New-AvdPatchHostResult -HostPoolName $poolName -ShortHostName $shortHost -AssignedUser $assignedUser -PriorPowerState $pwr.EffectiveState -FinalPowerState $pwr.EffectiveState -Status 'SkippedQuarantined' -Detail 'VM is Quarantined by inactivity lifecycle.'))
            continue
        }

        $existingLockRaw = [string](Get-AvdPatchProp -InputObject $vmTags -Name 'AvdPatchLock')
        $existingLock = ConvertFrom-AvdPatchLockValue -Value $existingLockRaw
        $existingStateRaw = [string](Get-AvdPatchProp -InputObject $vmTags -Name 'AvdPatchState')
        $existingState = ConvertFrom-AvdPatchStateValue -Value $existingStateRaw

        # ------------------------------------------------------------------
        # MorningUnlock Mode: isolated per-host recovery of drain + stale locks
        # ------------------------------------------------------------------
        if ($Mode -eq 'MorningUnlock') {
            try {
                if ($null -eq $existingLock) {
                    [void]$allResults.Add((New-AvdPatchHostResult -HostPoolName $poolName -ShortHostName $shortHost -AssignedUser $assignedUser -PriorPowerState $pwr.EffectiveState -FinalPowerState $pwr.EffectiveState -Status 'MorningHealthy' -Detail "allowNewSession=$allowNewSess"))
                    continue
                }
                if ($existingLock.ExpiresUtc -gt $nowUtc -and -not $ForceUnlockActive) {
                    [void]$allResults.Add((New-AvdPatchHostResult -HostPoolName $poolName -ShortHostName $shortHost -AssignedUser $assignedUser -PriorPowerState $pwr.EffectiveState -FinalPowerState $pwr.EffectiveState -Status 'SkippedActivePatchLock' -Detail "Active lock owned by run $($existingLock.RunId) until $($existingLock.ExpiresUtc.ToString('u'))."))
                    continue
                }
                $didChange = $false
                if (-not $allowNewSess -and $existingLock.PriorAllowNewSession -and -not $isQuarantined) {
                    if ($PSCmdlet.ShouldProcess($shortHost, 'Restore allowNewSession=true (MorningUnlock)')) {
                        $null = Set-AvdPatchSessionHostDrain -SessionHostArmId $shId -AllowNewSession $true -ApiVersion $DesktopVirtApiVersion -TokenState $tokenState
                        $didChange = $true
                    }
                }
                $rcCheck = Invoke-AvdPatchManagedRunCommand -Operation $(if ($PSCmdlet.ShouldProcess($shortHost, 'Prune stale Avd RunCommands')) { 'PruneStale' } else { 'InspectAvd' }) -VmResourceId $vmArmId -ComputeApiVersion $ComputeApiVersion -TokenState $tokenState
                if (-not $rcCheck.HasRunning -and $PSCmdlet.ShouldProcess($shortHost, 'Clear expired AvdPatchLock and exclusionTag')) {
                    $delMap = @{ AvdPatchLock = $existingLockRaw }
                    if (-not $existingLock.HadPriorExclusionTag -and -not [string]::IsNullOrWhiteSpace($existingLock.ExclusionTag)) {
                        $delMap[$existingLock.ExclusionTag] = 'AvdNightlyPatch'
                    }
                    $null = Update-AvdPatchVmTags -VmResourceId $vmArmId -Operation 'Delete' -TagsToApply $delMap -ExistingTags $vmTags -TagsApiVersion $TagsApiVersion -TokenState $tokenState
                    $didChange = $true
                }
                $mStatus = if ($rcCheck.HasRunning) { 'MorningBusyRunCommand' } elseif ($didChange) { 'MorningUnlocked' } else { 'MorningHealthy' }
                [void]$allResults.Add((New-AvdPatchHostResult -HostPoolName $poolName -ShortHostName $shortHost -AssignedUser $assignedUser -PriorPowerState $pwr.EffectiveState -FinalPowerState $pwr.EffectiveState -Status $mStatus -Detail "Processed lock from run $($existingLock.RunId) (PriorAllowNew=$($existingLock.PriorAllowNewSession))."))
            } catch {
                [void]$allResults.Add((New-AvdPatchHostResult -HostPoolName $poolName -ShortHostName $shortHost -AssignedUser $assignedUser -PriorPowerState $pwr.EffectiveState -FinalPowerState $pwr.EffectiveState -Status 'MorningUnlockError' -Detail $_.Exception.Message))
            }
            continue
        }

        # ------------------------------------------------------------------
        # Triage Mode: strictly 100% read-only inspection of locks & RunCommands
        # ------------------------------------------------------------------
        if ($Mode -eq 'Triage') {
            if ($null -eq $existingLock) {
                [void]$allResults.Add((New-AvdPatchHostResult -HostPoolName $poolName -ShortHostName $shortHost -AssignedUser $assignedUser -PriorPowerState $pwr.EffectiveState -FinalPowerState $pwr.EffectiveState -Status 'NoStaleLock' -Detail 'Clean'))
            } elseif ($existingLock.ExpiresUtc -gt $nowUtc) {
                [void]$allResults.Add((New-AvdPatchHostResult -HostPoolName $poolName -ShortHostName $shortHost -AssignedUser $assignedUser -PriorPowerState $pwr.EffectiveState -FinalPowerState $pwr.EffectiveState -Status 'ActivePatchLock' -Detail "Active lock $($existingLock.RunId) until $($existingLock.ExpiresUtc.ToString('u'))."))
            } else {
                $rcInspect = Invoke-AvdPatchManagedRunCommand -Operation 'InspectAvd' -VmResourceId $vmArmId -ComputeApiVersion $ComputeApiVersion -TokenState $tokenState
                $tStat = if ($rcInspect.HasRunning) { 'StaleLockBusyRunCommand' } else { 'StaleLockIdleRecoverable' }
                [void]$allResults.Add((New-AvdPatchHostResult -HostPoolName $poolName -ShortHostName $shortHost -AssignedUser $assignedUser -PriorPowerState $pwr.EffectiveState -FinalPowerState $pwr.EffectiveState -Status $tStat -Detail "Expired lock $($existingLock.RunId); RunningCmds=$($rcInspect.RunningNames -join ','); StaleCmds=$($rcInspect.StaleNames -join ',')."))
            }
            continue
        }

        # ------------------------------------------------------------------
        # Active / Expired Lock Handling for Audit & Execute
        # ------------------------------------------------------------------
        $effectivePriorState = $pwr.EffectiveState
        $effectivePriorAllow = $allowNewSess
        $hasPriorExclOverride = $false
        $priorExclOverrideVal = $false

        if ($null -ne $existingLock) {
            if ($existingLock.ExpiresUtc -gt $nowUtc) {
                [void]$allResults.Add((New-AvdPatchHostResult -HostPoolName $poolName -ShortHostName $shortHost -AssignedUser $assignedUser -PriorPowerState $pwr.EffectiveState -FinalPowerState $pwr.EffectiveState -Status 'SkippedActivePatchLock' -Detail "Active lock owned by run $($existingLock.RunId) until $($existingLock.ExpiresUtc.ToString('u'))."))
                continue
            }
            $rcInspect = Invoke-AvdPatchManagedRunCommand -Operation 'InspectAvd' -VmResourceId $vmArmId -ComputeApiVersion $ComputeApiVersion -TokenState $tokenState
            if ($rcInspect.HasRunning) {
                if ($Mode -eq 'Execute' -and -not $allowNewSess -and $existingLock.PriorAllowNewSession -and -not $isQuarantined) {
                    if ($PSCmdlet.ShouldProcess($shortHost, 'Restore allowNewSession=true on busy expired AvdPatchLock')) {
                        $null = Set-AvdPatchSessionHostDrain -SessionHostArmId $shId -AllowNewSession $true -ApiVersion $DesktopVirtApiVersion -TokenState $tokenState
                    }
                }
                [void]$allResults.Add((New-AvdPatchHostResult -HostPoolName $poolName -ShortHostName $shortHost -AssignedUser $assignedUser -PriorPowerState $pwr.EffectiveState -FinalPowerState $pwr.EffectiveState -Status 'NeedsOperatorTriage' -Detail "Expired AvdPatchLock ($($existingLock.RunId)) still has active RunCommand ($($rcInspect.RunningNames -join ',')); left untouched."))
                continue
            }
            # Idle expired lock: carry forward original pre-crash state so Execute self-heals cleanly
            if ($existingLock.PriorState -in @('Hibernated', 'Deallocated', 'Running', 'StoppedAllocated')) {
                $effectivePriorState = $existingLock.PriorState
            }
            $effectivePriorAllow = $existingLock.PriorAllowNewSession
            $hasPriorExclOverride = $true
            $priorExclOverrideVal = $existingLock.HadPriorExclusionTag
        }

        # Cohort & Smart Catch-Up Filter
        if (-not (Test-AvdPatchHostInCohort -ShortHostName $shortHost -CohortCount $CohortCount -CohortIndex $CohortIndex -NowUtc $nowUtc)) {
            [void]$allResults.Add((New-AvdPatchHostResult -HostPoolName $poolName -ShortHostName $shortHost -AssignedUser $assignedUser -PriorPowerState $pwr.EffectiveState -FinalPowerState $pwr.EffectiveState -Status 'SkippedCohortFilter' -Detail "Outside cohort $CohortIndex/$CohortCount."))
            continue
        }
        $dowNow = [int]$nowUtc.DayOfWeek
        $isCatchUpWindow = ($CohortCount -gt 1 -and $CohortIndex -lt 0 -and $CohortCount -le 5 -and ($dowNow -eq 0 -or $dowNow -eq 6 -or ($dowNow - 1) -ge $CohortCount))
        if ($isCatchUpWindow -and $CatchUpSkipRecentHours -gt 0 -and $null -eq $existingLock -and $null -ne $existingState.LastOkUtc) {
            $ageHours = ($nowUtc - [DateTimeOffset]$existingState.LastOkUtc).TotalHours
            if ($ageHours -ge 0 -and $ageHours -le $CatchUpSkipRecentHours -and $existingState.Outcome -in @('UpToDate', 'InstalledNoReboot', 'InstalledAndRebooted')) {
                [void]$allResults.Add((New-AvdPatchHostResult -HostPoolName $poolName -ShortHostName $shortHost -AssignedUser $assignedUser -PriorPowerState $pwr.EffectiveState -FinalPowerState $pwr.EffectiveState -Status 'SkippedCatchUpAlreadyHealthy' -Detail "Serviced $([int]$ageHours)h ago ($($existingState.Outcome), UBR=$($existingState.Ubr))."))
                continue
            }
        }

        $brokerSess = Get-AvdPatchBrokerSessionInfo -SessionHostArmId $shId -ApiVersion $DesktopVirtApiVersion -TokenState $tokenState
        if ($brokerSess.ActiveCount -gt 0) {
            [void]$allResults.Add((New-AvdPatchHostResult -HostPoolName $poolName -ShortHostName $shortHost -AssignedUser $assignedUser -PriorPowerState $pwr.EffectiveState -FinalPowerState $pwr.EffectiveState -SessionDetected $true -Status 'SkippedActiveSession' -Detail 'Active user session connected in AVD Broker.'))
            continue
        }
        if ($pwr.EffectiveState -in @('InTransition', 'Unknown')) {
            [void]$allResults.Add((New-AvdPatchHostResult -HostPoolName $poolName -ShortHostName $shortHost -AssignedUser $assignedUser -PriorPowerState $pwr.EffectiveState -FinalPowerState $pwr.EffectiveState -SessionDetected ($brokerSess.DisconnectedCount -gt 0) -Status 'SkippedVmInTransition' -Detail "VM powerCode='$($pwr.PowerCode)' is mid-transition."))
            continue
        }

        if ($Mode -eq 'Audit') {
            $planDetail = if ($exclCheck.IsValid) { "Eligible (ScalingExcl='$exclusionTagKey', UBR=$($existingState.Ubr))" } else { "WARN: $($exclCheck.FailureReason)" }
            [void]$allResults.Add((New-AvdPatchHostResult -HostPoolName $poolName -ShortHostName $shortHost -AssignedUser $assignedUser -PriorPowerState $pwr.EffectiveState -FinalPowerState $pwr.EffectiveState -SessionDetected ($brokerSess.DisconnectedCount -gt 0) -Status 'AuditEligible' -Detail $planDetail))
            continue
        }

        if (-not $PSCmdlet.ShouldProcess($shortHost, "Patch AVD Personal Session Host (PriorState=$effectivePriorState)")) {
            [void]$allResults.Add((New-AvdPatchHostResult -HostPoolName $poolName -ShortHostName $shortHost -AssignedUser $assignedUser -PriorPowerState $pwr.EffectiveState -FinalPowerState $effectivePriorState -SessionDetected ($brokerSess.DisconnectedCount -gt 0) -Status 'WhatIfPreview' -Detail "Would exclude ('$exclusionTagKey'), patch, and restore $effectivePriorState."))
            continue
        }

        # Case-insensitive check for pre-existing admin exclusionTag (unless recovering expired lock)
        $hadPriorExcl = $false
        if ($hasPriorExclOverride) {
            $hadPriorExcl = $priorExclOverrideVal
        } elseif ($null -ne $vmTags) {
            $existingExclVal = Get-AvdPatchProp -InputObject $vmTags -Name $exclusionTagKey
            if ($null -ne $existingExclVal) { $hadPriorExcl = $true }
        }

        $workQueue.Enqueue(@{
            PoolName               = $poolName
            ShortHost              = $shortHost
            AssignedUser           = $assignedUser
            SessionHostArmId       = $shId
            VmArmId                = $vmArmId
            VmGuid                 = $vmGuid
            Location               = $vmLocation
            InitialTags            = $vmTags
            InitialEffectiveState  = $pwr.EffectiveState
            HasExpiredLockRecovery = ($null -ne $existingLock)
            PriorState             = $effectivePriorState
            HibernationEnabled     = $pwr.HibernationEnabled
            PriorAllowNewSession   = $effectivePriorAllow
            HadPriorExclusionTag   = $hadPriorExcl
            IsQuarantined          = $isQuarantined
            PreWakeBrokerSession   = ($brokerSess.TotalCount -gt 0)
            PriorPatchState        = $existingState
            Phase                  = 'PendingStart'
            LockValue              = ''
            AsyncOpUri             = ''
            ActiveRunCmdName       = ''
            GuestStarted           = $false
            GuestResult            = $null
            BaselineBootUtc        = ''
            RebootOverallDeadline  = $null
            RebootPerformed        = $false
            HibernateRetries       = 0
            UnlockAttempts         = 0
            TargetFinalState       = $effectivePriorState
            PhaseDeadlineUtc       = $null
            FinalStatus            = ''
            FinalDetail            = ''
        })
    }

    if ($workQueue.Count -eq 0) { continue }

    # ==========================================================================
    # SLIDING-WINDOW STATE MACHINE
    # Transitions:
    #   PendingStart -> (WaitWake ->) LaunchPatchRunCmd -> WaitPatchRunCmd
    #     ├─ (no reboot needed) ──────────────────────────► RestorePower
    #     └─ (reboot needed) ─► WaitRebootScheduleCmd ─► WaitRebootDelay
    #                            ─► WaitPostRebootAgent ─► WaitVerifyRunCmd
    #                            ─► (WaitPostBootSettle ─►) RestorePower
    #   RestorePower -> (WaitPowerOffComplete ->) RestoreDrainAndUnlock -> UnlockOnly -> Completed
    # Safety Invariant: Never calls deallocate without hibernate=true on a Hibernated target,
    # and never powers off a VM whose RunCommand timed out while potentially running TiWorker.
    # ==========================================================================
    $activeItems = New-Object 'System.Collections.Generic.List[hashtable]'
    $startTimestamps = New-Object 'System.Collections.Generic.Queue[DateTimeOffset]'
    $processedCount = 0
    $failedCount = 0
    $circuitBroken = $false

    while ($workQueue.Count -gt 0 -or $activeItems.Count -gt 0) {
        $now = Get-AvdPatchNowUtc
        while ($startTimestamps.Count -gt 0 -and $startTimestamps.Peek() -lt $now.AddSeconds(-60)) {
            $null = $startTimestamps.Dequeue()
        }

        while ($workQueue.Count -gt 0 -and $activeItems.Count -lt $MaxConcurrentVMs) {
            if ($circuitBroken) {
                $skippedItem = $workQueue.Dequeue()
                [void]$allResults.Add((New-AvdPatchHostResult -HostPoolName $skippedItem.PoolName -ShortHostName $skippedItem.ShortHost -AssignedUser $skippedItem.AssignedUser -PriorPowerState $skippedItem.PriorState -FinalPowerState $skippedItem.PriorState -SessionDetected $skippedItem.PreWakeBrokerSession -Status 'AbortedByCircuitBreaker' -Detail "Circuit breaker tripped ($failedCount failures)."))
                continue
            }
            if (Test-AvdPatchWakeCutoffReached -NowUtc $now -CutoffHour $StopNewWakeAtLocalHour -TimeZoneId $MaintenanceTimeZoneId) {
                $cutItem = $workQueue.Dequeue()
                [void]$allResults.Add((New-AvdPatchHostResult -HostPoolName $cutItem.PoolName -ShortHostName $cutItem.ShortHost -AssignedUser $cutItem.AssignedUser -PriorPowerState $cutItem.PriorState -FinalPowerState $cutItem.PriorState -SessionDetected $cutItem.PreWakeBrokerSession -Status 'SkippedMaintenanceWindowCutoff' -Detail "Reached StopNewWakeAtLocalHour=$StopNewWakeAtLocalHour."))
                continue
            }
            $peekItem = $workQueue.Peek()
            if ($peekItem.InitialEffectiveState -ne 'Running' -and $startTimestamps.Count -ge $MaxStartsPerMinute) {
                break
            }
            $item = $workQueue.Dequeue()
            if ($item.InitialEffectiveState -ne 'Running') {
                $startTimestamps.Enqueue($now)
            }
            [void]$activeItems.Add($item)
        }

        for ($idx = $activeItems.Count - 1; $idx -ge 0; $idx--) {
            $it = $activeItems[$idx]
            $tickNow = Get-AvdPatchNowUtc

            try {
                switch ($it.Phase) {
                    'PendingStart' {
                        # Refresh live Broker sessions & VM state at admission so queued snapshots are never stale
                        $admSess = Get-AvdPatchBrokerSessionInfo -SessionHostArmId $it.SessionHostArmId -ApiVersion $DesktopVirtApiVersion -TokenState $tokenState
                        if ($admSess.ActiveCount -gt 0) {
                            $it.FinalStatus = 'SkippedActiveSession'
                            $it.FinalDetail = 'Active user session connected before lock acquisition.'
                            $it.Phase = 'Completed'
                            break
                        }
                        $it.PreWakeBrokerSession = ($it.PreWakeBrokerSession -or $admSess.TotalCount -gt 0)

                        $admVm = (Invoke-AvdPatchRestRequest -Method Get -Uri "https://management.azure.com$($it.VmArmId)?`$expand=instanceView&api-version=$ComputeApiVersion" -TokenState $tokenState).Body
                        $admPwr = Get-AvdPatchVmPowerSummary -VmResource $admVm
                        if ($admPwr.EffectiveState -in @('InTransition', 'Unknown')) {
                            $it.FinalStatus = 'SkippedVmInTransition'
                            $it.FinalDetail = "VM entered transitional state '$($admPwr.PowerCode)' before lock acquisition."
                            $it.Phase = 'Completed'
                            break
                        }
                        $admTags = Get-AvdPatchProp -InputObject $admVm -Name 'tags'
                        $it.InitialTags = $admTags
                        if (-not $it.HasExpiredLockRecovery) {
                            $it.PriorState = $admPwr.EffectiveState
                            $it.TargetFinalState = $admPwr.EffectiveState
                            $it.HadPriorExclusionTag = ($null -ne (Get-AvdPatchProp -InputObject $admTags -Name $exclusionTagKey))
                        } else {
                            $null = Invoke-AvdPatchManagedRunCommand -Operation 'PruneStale' -VmResourceId $it.VmArmId -ComputeApiVersion $ComputeApiVersion -TokenState $tokenState
                            $it.FinalDetail += 'Self-healed expired idle AvdPatchLock. '
                        }

                        $lockExp = $tickNow.AddMinutes($LockTtlMinutes)
                        $lockVal = ConvertTo-AvdPatchLockValue -RunId $runGuid -VmId $it.VmGuid -PriorState $it.PriorState -PriorAllowNewSession $it.PriorAllowNewSession -HadPriorExclusionTag $it.HadPriorExclusionTag -ExclusionTag $exclusionTagKey -ExpiresUtc $lockExp
                        $mergeTags = @{ AvdPatchLock = $lockVal }
                        if (-not $it.HadPriorExclusionTag) {
                            $mergeTags[$exclusionTagKey] = 'AvdNightlyPatch'
                        }

                        # Reserve 1 extra slot if AvdPatchState doesn't exist yet so final state persistence never hits 51 tags
                        $hasStateAlready = ($null -ne (Get-AvdPatchProp -InputObject $admTags -Name 'AvdPatchState'))
                        $uKeys = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
                        if ($null -ne $admTags) {
                            if ($admTags -is [System.Collections.IDictionary]) { foreach ($k in $admTags.Keys) { [void]$uKeys.Add([string]$k) } }
                            else { foreach ($p in $admTags.PSObject.Properties) { [void]$uKeys.Add($p.Name) } }
                        }
                        foreach ($mk in $mergeTags.Keys) { [void]$uKeys.Add([string]$mk) }
                        if (-not $hasStateAlready) { [void]$uKeys.Add('AvdPatchState') }
                        if ($uKeys.Count -gt 50) {
                            $it.FinalStatus = 'SkippedTagCeilingReached'
                            $it.FinalDetail = "VM has too many tags ($($uKeys.Count) needed > 50 ARM limit)."
                            $it.Phase = 'Completed'
                            break
                        }

                        $null = Update-AvdPatchVmTags -VmResourceId $it.VmArmId -Operation 'Merge' -TagsToApply $mergeTags -ExistingTags $admTags -TagsApiVersion $TagsApiVersion -TokenState $tokenState
                        $it.LockValue = $lockVal

                        # Re-check VM power state right after tag merge to catch in-flight Autoscale boundary transitions
                        $recheckVm = (Invoke-AvdPatchRestRequest -Method Get -Uri "https://management.azure.com$($it.VmArmId)?`$expand=instanceView&api-version=$ComputeApiVersion" -TokenState $tokenState).Body
                        $recheckPwr = Get-AvdPatchVmPowerSummary -VmResource $recheckVm
                        if ($recheckPwr.EffectiveState -eq 'InTransition') {
                            $it.FinalStatus = 'SkippedAutoscaleBoundaryRace'
                            $it.FinalDetail += 'VM entered transitional power state during lock acquisition; releasing cleanly.'
                            $it.TargetFinalState = $it.PriorState
                            $it.Phase = 'UnlockOnly'
                            break
                        }

                        if ($it.PriorAllowNewSession) {
                            $null = Set-AvdPatchSessionHostDrain -SessionHostArmId $it.SessionHostArmId -AllowNewSession $false -ApiVersion $DesktopVirtApiVersion -TokenState $tokenState
                        }

                        if ($recheckPwr.EffectiveState -eq 'Running' -and $recheckPwr.VmAgentReady) {
                            $it.Phase = 'LaunchPatchRunCmd'
                        } else {
                            try {
                                $startResp = Invoke-AvdPatchRestRequest -Method Post -Uri "https://management.azure.com$($it.VmArmId)/start?api-version=$ComputeApiVersion" -TokenState $tokenState
                                $it.AsyncOpUri = (Get-AvdPatchAsyncOperationStatus -ExtractFromResponse $startResp).OperationUri
                                $it.Phase = 'WaitWake'
                                $it.PhaseDeadlineUtc = $tickNow.AddMinutes(12)
                            } catch {
                                $errMsg = $_.Exception.Message
                                if ($errMsg -match 'AllocationFailed|ZonalAllocationFailed|Overconstrained') {
                                    $it.FinalStatus = 'WakeAllocationFailed'
                                    $it.FinalDetail += $errMsg
                                    $it.Phase = 'RestoreDrainAndUnlock'
                                } else {
                                    throw
                                }
                            }
                        }
                    }

                    'WaitWake' {
                        if (-not [string]::IsNullOrWhiteSpace($it.AsyncOpUri)) {
                            $opSt = Get-AvdPatchAsyncOperationStatus -OperationUri $it.AsyncOpUri -TokenState $tokenState
                            if ($opSt.Status -eq 'Unavailable') { $it.AsyncOpUri = '' }
                            elseif ([string]::Equals($opSt.Status, 'Failed', [System.StringComparison]::OrdinalIgnoreCase)) {
                                $combinedErr = "$($opSt.ErrorCode): $($opSt.ErrorMessage)"
                                $it.AsyncOpUri = ''
                                if ($combinedErr -match 'AllocationFailed|ZonalAllocationFailed|Overconstrained') {
                                    $it.FinalStatus = 'WakeAllocationFailed'
                                    $it.FinalDetail += $combinedErr
                                    $it.Phase = 'RestoreDrainAndUnlock'
                                    break
                                }
                                $it.FinalStatus = 'WakeFailed'
                                $it.FinalDetail += $combinedErr
                                $failedCount++
                                $it.Phase = 'RestoreDrainAndUnlock'
                                break
                            }
                        }
                        $vmNow = (Invoke-AvdPatchRestRequest -Method Get -Uri "https://management.azure.com$($it.VmArmId)?`$expand=instanceView&api-version=$ComputeApiVersion" -TokenState $tokenState).Body
                        $pNow = Get-AvdPatchVmPowerSummary -VmResource $vmNow
                        if ($pNow.EffectiveState -eq 'Running' -and $pNow.VmAgentReady) {
                            $it.AsyncOpUri = ''
                            if ($it.PriorState -eq 'Hibernated' -and $pNow.BootSubstatus -like '*Start/succeeded*') {
                                $it.FinalDetail += 'WARN: Azure cold-booted hibernated VM on resume (VMBootState/Start/succeeded). '
                            }
                            $it.Phase = 'LaunchPatchRunCmd'
                        } elseif ($tickNow -gt $it.PhaseDeadlineUtc) {
                            $it.FinalStatus = 'WakeTimeout'
                            $it.FinalDetail += 'Timed out waiting for PowerState/running and VmAgent Ready.'
                            $failedCount++
                            $it.Phase = 'RestorePower'
                        }
                    }

                    'LaunchPatchRunCmd' {
                        $postWakeSess = Get-AvdPatchBrokerSessionInfo -SessionHostArmId $it.SessionHostArmId -ApiVersion $DesktopVirtApiVersion -TokenState $tokenState
                        if ($postWakeSess.ActiveCount -gt 0) {
                            $it.FinalStatus = 'SkippedActiveSessionPostWake'
                            $it.FinalDetail += 'Active user session detected right after wake; leaving VM running.'
                            $it.TargetFinalState = 'Running'
                            $it.Phase = 'RestoreDrainAndUnlock'
                            break
                        }
                        # Prune any leftover terminal Avd* RunCommands so VM stays well below the 25-command quota
                        $null = Invoke-AvdPatchManagedRunCommand -Operation 'PruneStale' -VmResourceId $it.VmArmId -ComputeApiVersion $ComputeApiVersion -TokenState $tokenState

                        $hadAnySess = ($it.PreWakeBrokerSession -or $postWakeSess.TotalCount -gt 0)
                        $defEpoch = [long]$it.PriorPatchState.DeferSinceEpoch
                        $graceExpired = ($DisconnectedSessionGraceDays -eq 0 -or ($defEpoch -gt 0 -and ($tickNow.ToUnixTimeSeconds() - $defEpoch) -ge ($DisconnectedSessionGraceDays * 86400)))

                        $cmdName = "AvdPatch-$shortRun"
                        $scriptBody = New-AvdPatchGuestScript -Stage 'Service' -RunId $runGuid -PreWakeHadSession $hadAnySess -AllowRebootingUpdatesOnSession $graceExpired -MinFreeDiskGB $MinFreeDiskSpaceGB -IncludeDrivers ([bool]$IncludeDrivers) -IncludePreviewUpdates ([bool]$IncludePreviewUpdates)
                        $it.GuestStarted = $true
                        $null = Invoke-AvdPatchManagedRunCommand -Operation 'Put' -VmResourceId $it.VmArmId -Location $it.Location -RunCommandName $cmdName -ScriptText $scriptBody -TimeoutSeconds $RunCommandTimeoutSeconds -ComputeApiVersion $ComputeApiVersion -TokenState $tokenState
                        $it.ActiveRunCmdName = $cmdName
                        $it.Phase = 'WaitPatchRunCmd'
                        $it.PhaseDeadlineUtc = $tickNow.AddSeconds($RunCommandTimeoutSeconds + 300)
                    }

                    'WaitPatchRunCmd' {
                        $rc = Invoke-AvdPatchManagedRunCommand -Operation 'Get' -VmResourceId $it.VmArmId -RunCommandName $it.ActiveRunCmdName -ComputeApiVersion $ComputeApiVersion -TokenState $tokenState
                        if ($rc.ExecutionState -in @('Running', 'Pending')) {
                            if ($tickNow -gt $it.PhaseDeadlineUtc) {
                                $it.FinalStatus = 'NeedsOperatorTriage'
                                $it.FinalDetail += "RunCommand $($it.ActiveRunCmdName) exceeded orchestrator deadline while still Running; left VM Running with exclusionTag intact."
                                $it.TargetFinalState = 'Running'
                                $failedCount++
                                if ($it.PriorAllowNewSession -and -not $it.IsQuarantined) {
                                    $null = Set-AvdPatchSessionHostDrain -SessionHostArmId $it.SessionHostArmId -AllowNewSession $true -ApiVersion $DesktopVirtApiVersion -TokenState $tokenState
                                }
                                $it.Phase = 'Completed'
                            }
                            break
                        }

                        # Never hibernate if RunCommand TimedOut (PowerShell wrapper died while TiWorker may still be committing an LCU!)
                        if ([string]::Equals($rc.ExecutionState, 'TimedOut', [System.StringComparison]::OrdinalIgnoreCase)) {
                            $null = Invoke-AvdPatchManagedRunCommand -Operation 'Delete' -VmResourceId $it.VmArmId -RunCommandName $it.ActiveRunCmdName -ComputeApiVersion $ComputeApiVersion -TokenState $tokenState
                            $it.ActiveRunCmdName = ''
                            $it.FinalStatus = 'NeedsOperatorTriage'
                            $it.FinalDetail += "RunCommand timed out after ${RunCommandTimeoutSeconds}s; left VM Running to avoid interrupting TiWorker."
                            $it.TargetFinalState = 'Running'
                            $failedCount++
                            if ($it.PriorAllowNewSession -and -not $it.IsQuarantined) {
                                $null = Set-AvdPatchSessionHostDrain -SessionHostArmId $it.SessionHostArmId -AllowNewSession $true -ApiVersion $DesktopVirtApiVersion -TokenState $tokenState
                            }
                            $it.Phase = 'Completed'
                            break
                        }

                        $null = Invoke-AvdPatchManagedRunCommand -Operation 'Delete' -VmResourceId $it.VmArmId -RunCommandName $it.ActiveRunCmdName -ComputeApiVersion $ComputeApiVersion -TokenState $tokenState
                        $it.ActiveRunCmdName = ''
                        $parsed = ConvertFrom-AvdPatchRunCommandOutput -RawOutput $rc.Output -ExpectedRunId $runGuid
                        if ($null -eq $parsed) {
                            $tailPreview = if (-not [string]::IsNullOrWhiteSpace($rc.Output) -and $rc.Output.Length -gt 160) { $rc.Output.Substring($rc.Output.Length - 160) } else { $rc.Output }
                            $it.FinalStatus = 'RunCommandOutputParseFailed'
                            $it.FinalDetail += "ExecState=$($rc.ExecutionState) Exit=$($rc.ExitCode) Err=$($rc.Error) OutTail=$($tailPreview -replace '[\r\n]+',' ')"
                            $failedCount++
                            $it.Phase = 'RestorePower'
                            break
                        }

                        $it.GuestResult = $parsed
                        $it.BaselineBootUtc = [string](Get-AvdPatchProp -InputObject $parsed -Name 'bootUtc')
                        $gOutcome = [string](Get-AvdPatchProp -InputObject $parsed -Name 'outcome')
                        $gSess = [bool](Get-AvdPatchProp -InputObject $parsed -Name 'sessionPresent')
                        $gRbt = ([bool](Get-AvdPatchProp -InputObject $parsed -Name 'rbtReq') -or [bool](Get-AvdPatchProp -InputObject $parsed -Name 'postRbt'))
                        $defErrMsg = [string](Get-AvdPatchProp -InputObject $parsed -Name 'defErr')
                        if (-not [string]::IsNullOrWhiteSpace($defErrMsg)) {
                            $it.FinalDetail += "WARN: Defender signature update failed ($defErrMsg). "
                        }

                        if ($gOutcome -in @('GuestScriptException', 'InsufficientDiskSpace', 'PartialInstallWarning', 'PartialInstallRebootRequired') -or $gOutcome -like 'DownloadFailed*' -or $gOutcome -like 'InstallFailed*') {
                            $failedCount++
                        }

                        $graceExpiredNow = ($DisconnectedSessionGraceDays -eq 0 -or ($it.PriorPatchState.DeferSinceEpoch -gt 0 -and ($tickNow.ToUnixTimeSeconds() - $it.PriorPatchState.DeferSinceEpoch) -ge ($DisconnectedSessionGraceDays * 86400)))
                        $defEpochNext = [long]0
                        $defKbs = @(Get-AvdPatchProp -InputObject $parsed -Name 'deferredKbs')
                        if ($defKbs.Count -gt 0 -or ($gSess -and $gRbt -and -not $graceExpiredNow)) {
                            $defEpochNext = if ($it.PriorPatchState.DeferSinceEpoch -gt 0) { [long]$it.PriorPatchState.DeferSinceEpoch } else { $tickNow.ToUnixTimeSeconds() }
                        }
                        $it['NextDeferEpoch'] = $defEpochNext

                        if ($gRbt -and (-not $gSess -or $graceExpiredNow)) {
                            $rbtCmd = "AvdReboot-$shortRun"
                            $rbtScript = New-AvdPatchGuestScript -Stage 'ScheduleReboot' -RunId $runGuid -AllowRebootingUpdatesOnSession ($gSess -and $graceExpiredNow)
                            $null = Invoke-AvdPatchManagedRunCommand -Operation 'Put' -VmResourceId $it.VmArmId -Location $it.Location -RunCommandName $rbtCmd -ScriptText $rbtScript -TimeoutSeconds 300 -ComputeApiVersion $ComputeApiVersion -TokenState $tokenState
                            $it.ActiveRunCmdName = $rbtCmd
                            $it.Phase = 'WaitRebootScheduleCmd'
                            $it.PhaseDeadlineUtc = $tickNow.AddMinutes(6)
                        } else {
                            $it.FinalStatus = $gOutcome
                            $it.Phase = 'RestorePower'
                        }
                    }

                    'WaitRebootScheduleCmd' {
                        $rc = Invoke-AvdPatchManagedRunCommand -Operation 'Get' -VmResourceId $it.VmArmId -RunCommandName $it.ActiveRunCmdName -ComputeApiVersion $ComputeApiVersion -TokenState $tokenState
                        if ($rc.ExecutionState -in @('Running', 'Pending')) {
                            if ($tickNow -gt $it.PhaseDeadlineUtc) {
                                $it.FinalStatus = 'NeedsOperatorTriage'
                                $it.FinalDetail += 'Reboot scheduling RunCommand timed out while still Running; left VM Running with exclusionTag intact.'
                                $it.TargetFinalState = 'Running'
                                $failedCount++
                                if ($it.PriorAllowNewSession -and -not $it.IsQuarantined) {
                                    $null = Set-AvdPatchSessionHostDrain -SessionHostArmId $it.SessionHostArmId -AllowNewSession $true -ApiVersion $DesktopVirtApiVersion -TokenState $tokenState
                                }
                                $it.Phase = 'Completed'
                            }
                            break
                        }
                        $null = Invoke-AvdPatchManagedRunCommand -Operation 'Delete' -VmResourceId $it.VmArmId -RunCommandName $it.ActiveRunCmdName -ComputeApiVersion $ComputeApiVersion -TokenState $tokenState
                        $it.ActiveRunCmdName = ''
                        $rbtParsed = ConvertFrom-AvdPatchRunCommandOutput -RawOutput $rc.Output -ExpectedRunId $runGuid
                        if ($null -eq $rbtParsed -or -not [bool](Get-AvdPatchProp -InputObject $rbtParsed -Name 'scheduled')) {
                            $rOut = if ($null -ne $rbtParsed) { [string](Get-AvdPatchProp -InputObject $rbtParsed -Name 'outcome') } else { 'RebootScheduleParseError' }
                            $it.FinalStatus = $rOut
                            $it.Phase = 'RestorePower'
                            break
                        }
                        $it.RebootOverallDeadline = $tickNow.AddMinutes(25)
                        $it.Phase = 'WaitRebootDelay'
                        $it.PhaseDeadlineUtc = $tickNow.AddSeconds(45)
                    }

                    'WaitRebootDelay' {
                        if ($tickNow -ge $it.PhaseDeadlineUtc) {
                            $it.Phase = 'WaitPostRebootAgent'
                            $it.PhaseDeadlineUtc = $it.RebootOverallDeadline
                        }
                    }

                    'WaitPostRebootAgent' {
                        $vmNow = (Invoke-AvdPatchRestRequest -Method Get -Uri "https://management.azure.com$($it.VmArmId)?`$expand=instanceView&api-version=$ComputeApiVersion" -TokenState $tokenState).Body
                        $pNow = Get-AvdPatchVmPowerSummary -VmResource $vmNow
                        if ($pNow.EffectiveState -eq 'Running' -and $pNow.VmAgentReady) {
                            $verAttempt = if ($it.ContainsKey('VerifySeq')) { [int]$it['VerifySeq'] + 1 } else { 1 }
                            $it['VerifySeq'] = $verAttempt
                            $verCmd = "AvdVerify-$shortRun-$verAttempt"
                            $verScript = New-AvdPatchGuestScript -Stage 'VerifyPostReboot' -RunId $runGuid -BaselineBootUtc $it.BaselineBootUtc
                            $null = Invoke-AvdPatchManagedRunCommand -Operation 'Put' -VmResourceId $it.VmArmId -Location $it.Location -RunCommandName $verCmd -ScriptText $verScript -TimeoutSeconds 600 -ComputeApiVersion $ComputeApiVersion -TokenState $tokenState
                            $it.ActiveRunCmdName = $verCmd
                            $it.Phase = 'WaitVerifyRunCmd'
                            $it.PhaseDeadlineUtc = $tickNow.AddMinutes(10)
                        } elseif ($tickNow -gt $it.RebootOverallDeadline) {
                            $it.FinalStatus = 'PostRebootAgentTimeout'
                            $it.TargetFinalState = 'Running'
                            $failedCount++
                            $it.Phase = 'RestoreDrainAndUnlock'
                        }
                    }

                    'WaitVerifyRunCmd' {
                        $rc = Invoke-AvdPatchManagedRunCommand -Operation 'Get' -VmResourceId $it.VmArmId -RunCommandName $it.ActiveRunCmdName -ComputeApiVersion $ComputeApiVersion -TokenState $tokenState
                        if ($rc.ExecutionState -in @('Running', 'Pending')) {
                            if ($tickNow -gt $it.PhaseDeadlineUtc) {
                                $it.FinalStatus = 'NeedsOperatorTriage'
                                $it.FinalDetail += 'Post-reboot verify RunCommand timed out while still Running; left VM Running with exclusionTag intact.'
                                $it.TargetFinalState = 'Running'
                                $failedCount++
                                if ($it.PriorAllowNewSession -and -not $it.IsQuarantined) {
                                    $null = Set-AvdPatchSessionHostDrain -SessionHostArmId $it.SessionHostArmId -AllowNewSession $true -ApiVersion $DesktopVirtApiVersion -TokenState $tokenState
                                }
                                $it.Phase = 'Completed'
                            }
                            break
                        }
                        $null = Invoke-AvdPatchManagedRunCommand -Operation 'Delete' -VmResourceId $it.VmArmId -RunCommandName $it.ActiveRunCmdName -ComputeApiVersion $ComputeApiVersion -TokenState $tokenState
                        $it.ActiveRunCmdName = ''
                        $verParsed = ConvertFrom-AvdPatchRunCommandOutput -RawOutput $rc.Output -ExpectedRunId $runGuid
                        $verOutcome = if ($null -ne $verParsed) { [string](Get-AvdPatchProp -InputObject $verParsed -Name 'outcome') } else { '' }
                        $bootAdv = ($null -ne $verParsed -and [bool](Get-AvdPatchProp -InputObject $verParsed -Name 'bootAdvanced'))

                        if (-not $bootAdv -or $verOutcome -eq 'BootNotYetAdvanced') {
                            if ($tickNow -lt $it.RebootOverallDeadline) {
                                $it.Phase = 'WaitRebootDelay'
                                $it.PhaseDeadlineUtc = $tickNow.AddSeconds(30)
                            } else {
                                $it.FinalStatus = 'PostRebootBootNotAdvanced'
                                $it.FinalDetail += "BootUtc did not advance past $($it.BaselineBootUtc) within 25 minutes; left VM Running."
                                $it.TargetFinalState = 'Running'
                                $failedCount++
                                $it.Phase = 'RestoreDrainAndUnlock'
                            }
                            break
                        }

                        if ($verOutcome -ne 'Verified') {
                            $it.FinalStatus = 'PostRebootVerifyFailed'
                            $it.FinalDetail += "Post-reboot verify returned outcome='$verOutcome'; left VM Running."
                            $it.TargetFinalState = 'Running'
                            $failedCount++
                            $it.Phase = 'RestoreDrainAndUnlock'
                            break
                        }

                        $it.RebootPerformed = $true
                        $newUbr = [int](Get-AvdPatchProp -InputObject $verParsed -Name 'ubr')
                        if ($newUbr -gt 0 -and $null -ne $it.GuestResult) { $it.GuestResult.ubr = $newUbr }
                        $stillPending = [bool](Get-AvdPatchProp -InputObject $verParsed -Name 'postRbt')
                        if (-not $stillPending) { $it['NextDeferEpoch'] = [long]0 }
                        $preOutcome = if ($null -ne $it.GuestResult) { [string](Get-AvdPatchProp -InputObject $it.GuestResult -Name 'outcome') } else { '' }
                        $it.FinalStatus = if ($stillPending) {
                            'RebootedPendingSecondPass'
                        } elseif ($preOutcome -eq 'PreExistingRebootPending') {
                            'RebootedClearedPendingState'
                        } else {
                            'InstalledAndRebooted'
                        }
                        if ($it.PriorState -eq 'Hibernated' -and $PostRebootZeroSessionAction -eq 'Deallocate') {
                            $it.TargetFinalState = 'Deallocated'
                        }
                        if ($it.TargetFinalState -eq 'Hibernated' -and $PostRebootSettleSeconds -gt 0) {
                            $it.Phase = 'WaitPostBootSettle'
                            $it.PhaseDeadlineUtc = $tickNow.AddSeconds($PostRebootSettleSeconds)
                        } else {
                            $it.Phase = 'RestorePower'
                        }
                    }

                    'WaitPostBootSettle' {
                        if ($tickNow -ge $it.PhaseDeadlineUtc) {
                            $it.Phase = 'RestorePower'
                        }
                    }

                    'RestorePower' {
                        $latestSess = Get-AvdPatchBrokerSessionInfo -SessionHostArmId $it.SessionHostArmId -ApiVersion $DesktopVirtApiVersion -TokenState $tokenState
                        if ($latestSess.ActiveCount -gt 0) {
                            $it.TargetFinalState = 'Running'
                            $it.FinalDetail += 'Active user session appeared; left VM Running. '
                            $it.Phase = 'RestoreDrainAndUnlock'
                            break
                        }

                        if ($it.TargetFinalState -eq 'Running') {
                            $it.Phase = 'RestoreDrainAndUnlock'
                        } elseif ($it.TargetFinalState -eq 'Hibernated') {
                            try {
                                $hibResp = Invoke-AvdPatchRestRequest -Method Post -Uri "https://management.azure.com$($it.VmArmId)/deallocate?hibernate=true&api-version=$ComputeApiVersion" -TokenState $tokenState
                                $it.AsyncOpUri = (Get-AvdPatchAsyncOperationStatus -ExtractFromResponse $hibResp).OperationUri
                                $it.Phase = 'WaitPowerOffComplete'
                                $it.PhaseDeadlineUtc = $tickNow.AddMinutes(10)
                            } catch {
                                $hErr = $_.Exception.Message
                                if ($hErr -match 'VMHibernateNotSupported|VMHibernateFailed' -and $it.HibernateRetries -lt 3) {
                                    $it.HibernateRetries++
                                    $it.Phase = 'WaitPostBootSettle'
                                    $it.PhaseDeadlineUtc = $tickNow.AddSeconds(180)
                                } else {
                                    # NEVER fall back to hibernate=false (would wipe preserved user RAM)!
                                    $it.FinalStatus = 'HibernateRestoreFailed_LeftRunning'
                                    $it.FinalDetail += "deallocate?hibernate=true failed ($hErr); left VM Running to protect RAM."
                                    $it.TargetFinalState = 'Running'
                                    $failedCount++
                                    $it.Phase = 'RestoreDrainAndUnlock'
                                }
                            }
                        } elseif ($it.TargetFinalState -eq 'Deallocated') {
                            $deallocResp = Invoke-AvdPatchRestRequest -Method Post -Uri "https://management.azure.com$($it.VmArmId)/deallocate?api-version=$ComputeApiVersion" -TokenState $tokenState
                            $it.AsyncOpUri = (Get-AvdPatchAsyncOperationStatus -ExtractFromResponse $deallocResp).OperationUri
                            $it.Phase = 'WaitPowerOffComplete'
                            $it.PhaseDeadlineUtc = $tickNow.AddMinutes(10)
                        } elseif ($it.TargetFinalState -eq 'StoppedAllocated') {
                            $stopResp = Invoke-AvdPatchRestRequest -Method Post -Uri "https://management.azure.com$($it.VmArmId)/powerOff?api-version=$ComputeApiVersion" -TokenState $tokenState
                            $it.AsyncOpUri = (Get-AvdPatchAsyncOperationStatus -ExtractFromResponse $stopResp).OperationUri
                            $it.Phase = 'WaitPowerOffComplete'
                            $it.PhaseDeadlineUtc = $tickNow.AddMinutes(10)
                        } else {
                            $it.Phase = 'RestoreDrainAndUnlock'
                        }
                    }

                    'WaitPowerOffComplete' {
                        if (-not [string]::IsNullOrWhiteSpace($it.AsyncOpUri)) {
                            $opSt = Get-AvdPatchAsyncOperationStatus -OperationUri $it.AsyncOpUri -TokenState $tokenState
                            if ($opSt.Status -eq 'Unavailable') { $it.AsyncOpUri = '' }
                            elseif ([string]::Equals($opSt.Status, 'Failed', [System.StringComparison]::OrdinalIgnoreCase)) {
                                $combinedErr = "$($opSt.ErrorCode): $($opSt.ErrorMessage)"
                                $it.AsyncOpUri = ''
                                if ($it.TargetFinalState -eq 'Hibernated') {
                                    if ($combinedErr -match 'VMHibernateNotSupported|VMHibernateFailed' -and $it.HibernateRetries -lt 3) {
                                        $it.HibernateRetries++
                                        $it.Phase = 'WaitPostBootSettle'
                                        $it.PhaseDeadlineUtc = $tickNow.AddSeconds(180)
                                        break
                                    }
                                    $it.FinalStatus = 'HibernateRestoreFailed_LeftRunning'
                                    $it.FinalDetail += "Async hibernate LRO failed ($combinedErr); left VM Running to protect RAM."
                                    $it.TargetFinalState = 'Running'
                                    $failedCount++
                                    $it.Phase = 'RestoreDrainAndUnlock'
                                    break
                                }
                                $it.FinalStatus = 'PowerRestoreFailed'
                                $it.FinalDetail += "Async power-off LRO failed ($combinedErr)."
                                $it.TargetFinalState = 'Running'
                                $failedCount++
                                $it.Phase = 'RestoreDrainAndUnlock'
                                break
                            }
                        }

                        $vmNow = (Invoke-AvdPatchRestRequest -Method Get -Uri "https://management.azure.com$($it.VmArmId)?`$expand=instanceView&api-version=$ComputeApiVersion" -TokenState $tokenState).Body
                        $pNow = Get-AvdPatchVmPowerSummary -VmResource $vmNow
                        $reached = ($it.TargetFinalState -eq 'Hibernated' -and $pNow.EffectiveState -eq 'Hibernated') -or
                                   ($it.TargetFinalState -eq 'Deallocated' -and $pNow.EffectiveState -eq 'Deallocated') -or
                                   ($it.TargetFinalState -eq 'StoppedAllocated' -and $pNow.EffectiveState -eq 'StoppedAllocated')
                        if ($reached) {
                            $it.AsyncOpUri = ''
                            $it.Phase = 'RestoreDrainAndUnlock'
                        } elseif ($tickNow -gt $it.PhaseDeadlineUtc) {
                            $it.FinalStatus = if ($it.TargetFinalState -eq 'Hibernated') { 'HibernateRestoreFailed_LeftRunning' } else { 'PowerRestoreTimeout' }
                            $it.FinalDetail += "Timed out waiting for $($it.TargetFinalState) (actual=$($pNow.EffectiveState)). "
                            $it.TargetFinalState = $pNow.EffectiveState
                            $failedCount++
                            $it.Phase = 'RestoreDrainAndUnlock'
                        }
                    }

                    'RestoreDrainAndUnlock' {
                        if ($it.PriorAllowNewSession -and -not $it.IsQuarantined) {
                            $null = Set-AvdPatchSessionHostDrain -SessionHostArmId $it.SessionHostArmId -AllowNewSession $true -ApiVersion $DesktopVirtApiVersion -TokenState $tokenState
                        }
                        $it.Phase = 'UnlockOnly'
                    }

                    'UnlockOnly' {
                        $freshVm = (Invoke-AvdPatchRestRequest -Method Get -Uri "https://management.azure.com$($it.VmArmId)?api-version=$ComputeApiVersion" -TokenState $tokenState).Body
                        $freshTags = Get-AvdPatchProp -InputObject $freshVm -Name 'tags'
                        # Delete transient tags FIRST (ownership-conditional) so 50-tag-boundary VMs never fail adding AvdPatchState
                        $delTags = @{}
                        if (-not [string]::IsNullOrWhiteSpace($it.LockValue)) {
                            $delTags['AvdPatchLock'] = $it.LockValue
                        }
                        if (-not $it.HadPriorExclusionTag) {
                            $delTags[$exclusionTagKey] = 'AvdNightlyPatch'
                        }
                        if ($delTags.Count -gt 0 -and $null -ne $freshTags) {
                            $null = Update-AvdPatchVmTags -VmResourceId $it.VmArmId -Operation 'Delete' -TagsToApply $delTags -ExistingTags $freshTags -TagsApiVersion $TagsApiVersion -TokenState $tokenState
                        }
                        if ($null -ne $it.GuestResult) {
                            try {
                                $ubrVal = [int](Get-AvdPatchProp -InputObject $it.GuestResult -Name 'ubr')
                                $defEp = if ($it.ContainsKey('NextDeferEpoch')) { [long]$it['NextDeferEpoch'] } else { [long]0 }
                                $stVal = ConvertTo-AvdPatchStateValue -LastOkUtc $tickNow -Ubr $ubrVal -Outcome $it.FinalStatus -DeferSinceEpoch $defEp
                                $postDelVm = (Invoke-AvdPatchRestRequest -Method Get -Uri "https://management.azure.com$($it.VmArmId)?api-version=$ComputeApiVersion" -TokenState $tokenState).Body
                                $postDelTags = Get-AvdPatchProp -InputObject $postDelVm -Name 'tags'
                                $null = Update-AvdPatchVmTags -VmResourceId $it.VmArmId -Operation 'Merge' -TagsToApply @{ AvdPatchState = $stVal } -ExistingTags $postDelTags -TagsApiVersion $TagsApiVersion -TokenState $tokenState
                            } catch {
                                $it.FinalDetail += "WARN: AvdPatchState write skipped ($($_.Exception.Message)). "
                            }
                        }
                        $it.Phase = 'Completed'
                    }
                }
            } catch {
                $it.FinalStatus = 'OrchestratorError'
                $it.FinalDetail += $_.Exception.Message
                $failedCount++
                try {
                    if ($it.PriorAllowNewSession -and -not $it.IsQuarantined) {
                        $null = Set-AvdPatchSessionHostDrain -SessionHostArmId $it.SessionHostArmId -AllowNewSession $true -ApiVersion $DesktopVirtApiVersion -TokenState $tokenState
                    }
                } catch { $null = $_ }

                $rcStillRunning = $false
                if (-not [string]::IsNullOrWhiteSpace($it.ActiveRunCmdName)) {
                    try {
                        $rcProbe = Invoke-AvdPatchManagedRunCommand -Operation 'Get' -VmResourceId $it.VmArmId -RunCommandName $it.ActiveRunCmdName -ComputeApiVersion $ComputeApiVersion -TokenState $tokenState
                        if ($rcProbe.ExecutionState -in @('Running', 'Pending')) {
                            $rcStillRunning = $true
                        } else {
                            $null = Invoke-AvdPatchManagedRunCommand -Operation 'Delete' -VmResourceId $it.VmArmId -RunCommandName $it.ActiveRunCmdName -ComputeApiVersion $ComputeApiVersion -TokenState $tokenState
                            $it.ActiveRunCmdName = ''
                        }
                    } catch {
                        $rcStillRunning = $true
                    }
                }
                if (-not [string]::IsNullOrWhiteSpace($it.LockValue) -and -not $rcStillRunning -and $it.UnlockAttempts -lt 1) {
                    $it.UnlockAttempts++
                    $it.TargetFinalState = if ($it.GuestStarted) { 'Running' } else { $it.PriorState }
                    $it.Phase = 'UnlockOnly'
                } else {
                    $it.Phase = 'Completed'
                }
            }

            if ($it.Phase -eq 'Completed') {
                $processedCount++
                $g = $it.GuestResult
                $sessDet = ($it.PreWakeBrokerSession -or ($null -ne $g -and [bool](Get-AvdPatchProp -InputObject $g -Name 'sessionPresent')))
                $hpElig = ($null -ne $g -and [bool](Get-AvdPatchProp -InputObject $g -Name 'vbsRunning') -and [bool](Get-AvdPatchProp -InputObject $g -Name 'hotpatchPolicy'))
                $instStr = if ($null -ne $g) { (@(Get-AvdPatchProp -InputObject $g -Name 'installedKbs') -join ',') } else { '' }
                $defStr = if ($null -ne $g) { (@(Get-AvdPatchProp -InputObject $g -Name 'deferredKbs') -join ',') } else { '' }
                $detailStr = $it.FinalDetail
                if ([string]::IsNullOrWhiteSpace($detailStr) -and $null -ne $g) {
                    $detailStr = "UBR=$(Get-AvdPatchProp -InputObject $g -Name 'ubr') MDM=$(Get-AvdPatchProp -InputObject $g -Name 'mdmSynced') VBS=$(Get-AvdPatchProp -InputObject $g -Name 'vbsRunning')"
                }
                [void]$allResults.Add((New-AvdPatchHostResult -HostPoolName $it.PoolName -ShortHostName $it.ShortHost -AssignedUser $it.AssignedUser -PriorPowerState $it.PriorState -FinalPowerState $it.TargetFinalState -SessionDetected $sessDet -HotpatchEligible $hpElig -InstalledKbs $instStr -DeferredKbs $defStr -RebootPerformed $it.RebootPerformed -Status $it.FinalStatus -Detail $detailStr))
                $activeItems.RemoveAt($idx)

                if ($failedCount -ge $MaxFailureCount -and $processedCount -gt 0) {
                    $pct = ($failedCount * 100.0) / $processedCount
                    if ($pct -ge $MaxFailurePercent) { $circuitBroken = $true }
                }
            }
        }

        if ($activeItems.Count -gt 0 -or $workQueue.Count -gt 0) {
            Start-AvdPatchPollSleep -Seconds 15
        }
    }
}

return $allResults.ToArray()
