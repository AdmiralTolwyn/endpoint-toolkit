<#
.SYNOPSIS
    Propagates AVD Personal Host Pool user assignments to Intune Primary User
    and optionally triggers an Intune device sync.
.DESCRIPTION
    When AVD session hosts are joined to Microsoft Entra ID and enrolled into
    Intune via the AADLoginForWindows VM extension (mdmId = 0000000a-0000-0000-c000-000000000000),
    enrollment occurs under the SYSTEM account with no user affinity, leaving the
    Intune Managed Device Primary User unset (None / empty GUID). AVD does not
    natively synchronize Personal desktop assignments (properties.assignedUser)
    to Intune.
    This script performs fail-closed Primary User propagation for explicit AVD
    Personal Host Pools:
    1. Queries the specified Personal Host Pool(s) and their session hosts in ARM.
    2. Correlates each assigned session host to its Windows Intune managedDevice
       record using batched OData filters, failing closed on duplicate Intune
       records, cross-pool duplicate hostnames, or stale Intune records enrolled
       more than 24 hours before the current AVD sessionHost creation time.
    3. Resolves sessionHost.properties.assignedUser in Microsoft Entra ID (with
       B2B guest mail fallback) to obtain the immutable Entra User ObjectId.
    4. Reads the authoritative Intune Primary User relationship
       (GET /beta/deviceManagement/managedDevices/{id}/users) and compares
       resolved Entra User ObjectIds. Fails closed (PrimaryUserConflict) if a
       different Primary User is already assigned.
    5. Sets the Intune Primary User via
       POST /beta/deviceManagement/managedDevices/{id}/users/$ref, and if
       -SyncIntuneDevice is specified, invokes
       POST /v1.0/deviceManagement/managedDevices/{id}/syncDevice.
    Required Permissions (Managed Identity / Service Principal):
    - Azure RBAC: Desktop Virtualization Reader on target Resource Groups.
    - Microsoft Graph Application Permissions:
      * DeviceManagementManagedDevices.ReadWrite.All (read devices/users, set users/$ref)
      * User.Read.All (resolve assignedUser UPN/mail to Entra user ObjectId)
      * DeviceManagementManagedDevices.PrivilegedOperations.All (OPTIONAL — only required when -SyncIntuneDevice is used)
.PARAMETER SubscriptionId
    Optional Azure Subscription ID used with -ResourceGroupName and -HostPoolName.
    If omitted, uses the active Get-AzContext subscription.
.PARAMETER ResourceGroupName
    Resource Group containing the target AVD Personal Host Pool (used with -HostPoolName).
.PARAMETER HostPoolName
    Name of the target AVD Personal Host Pool (used with -ResourceGroupName).
.PARAMETER HostPoolResourceId
    Full ARM Resource ID of one or more AVD Personal Host Pools to evaluate.
.PARAMETER TargetSessionHostName
    Optional session host name (short name or FQDN) to scope evaluation to a
    single host within the target Host Pool.
.PARAMETER ManagedIdentityClientId
    Optional Client ID of a User-Assigned Managed Identity. Omit to use the
    System-Assigned Managed Identity.
.PARAMETER GraphAccessToken
    Optional pre-acquired Microsoft Graph Bearer token for interactive workstation
    runs where the Azure PowerShell client application lacks Intune Graph scopes.
.PARAMETER AuthMode
    Authentication source for ARM and Microsoft Graph tokens:
    - Auto (default): Uses Managed Identity if IDENTITY_ENDPOINT is present,
      otherwise falls back to an active Az.Accounts session (Get-AzAccessToken).
    - ManagedIdentity: Forces Managed Identity (Automation / IMDS).
    - AzContext: Forces Get-AzAccessToken from the current Az PowerShell session.
.PARAMETER SyncIntuneDevice
    Optional. Invoke POST /v1.0/deviceManagement/managedDevices/{id}/syncDevice
    immediately after setting the Intune Primary User. Disabled by default so
    DeviceManagementManagedDevices.PrivilegedOperations.All is not required
    unless explicitly requested.
.OUTPUTS
    PSCustomObject[]
    Emits one structured result object per evaluated AVD session host with
    HostPoolName, SessionHostName, ShortHostName, AssignedUser, IntuneDeviceId,
    PreviousPrimaryUser, NewPrimaryUser, Action, SyncTriggered, Status and Detail.
.EXAMPLE
    .\Sync-AvdPersonalPrimaryUser.ps1 -ResourceGroupName 'rg-avd' -HostPoolName 'vdpool-pers-01' -WhatIf
    Audits a Personal Host Pool and reports which Intune devices need their
    Primary User populated without making Graph changes.
.EXAMPLE
    .\Sync-AvdPersonalPrimaryUser.ps1 -ResourceGroupName 'rg-avd' -HostPoolName 'vdpool-pers-01' -Confirm:$false
    Populates missing Intune Primary Users for the specified Personal Host Pool
    (requires only ReadWrite.All and User.Read.All).
.EXAMPLE
    .\Sync-AvdPersonalPrimaryUser.ps1 -ResourceGroupName 'rg-avd' -HostPoolName 'vdpool-pers-01' -SyncIntuneDevice -Confirm:$false
    Populates missing Intune Primary Users and triggers an immediate Intune
    syncDevice check-in (requires PrivilegedOperations.All).
#>

#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $false)]
    [string] $SubscriptionId,
    [Parameter(Mandatory = $false)]
    [string] $ResourceGroupName,
    [Parameter(Mandatory = $false)]
    [string] $HostPoolName,
    [Parameter(Mandatory = $false)]
    [string[]] $HostPoolResourceId,
    [Parameter(Mandatory = $false)]
    [string] $TargetSessionHostName,
    [Parameter(Mandatory = $false)]
    [string] $ManagedIdentityClientId,
    [Parameter(Mandatory = $false)]
    [string] $GraphAccessToken,
    [Parameter(Mandatory = $false)]
    [ValidateSet('Auto', 'ManagedIdentity', 'AzContext')]
    [string] $AuthMode = 'Auto',
    [Parameter(Mandatory = $false)]
    [switch] $SyncIntuneDevice
)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

function Get-AvdPrimaryUserValue {
    <#
    .SYNOPSIS
        Safely reads a named property or dictionary key under Set-StrictMode.
    .DESCRIPTION
        Returns the value of PropertyName from a PSCustomObject, PSObject, or
        IDictionary without throwing PropertyNotFoundException under StrictMode 2.0.
        Returns $null when InputObject is $null or the property does not exist.
    .PARAMETER InputObject
        The object or dictionary to inspect.
    .PARAMETER PropertyName
        The property or key name to retrieve.
    .OUTPUTS
        System.Object
    #>
    param(
        [Parameter(Mandatory = $false)] $InputObject,
        [Parameter(Mandatory = $true)] [string] $PropertyName
    )
    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [System.Collections.IDictionary]) {
        if ($InputObject.Contains($PropertyName)) { return $InputObject[$PropertyName] }
        return $null
    }
    $prop = $InputObject.PSObject.Properties[$PropertyName]
    if ($null -ne $prop) { return $prop.Value }
    return $null
}

function ConvertTo-NormalizedPrincipalId {
    <#
    .SYNOPSIS
        Normalizes an Entra or Intune GUID/identifier string and filters zero GUIDs.
    .DESCRIPTION
        Trims whitespace from a GUID or identifier string and converts empty,
        whitespace-only, or all-zero GUID values (00000000-0000-0000-0000-000000000000,
        which Intune frequently reports on SYSTEM-enrolled devices with no user
        affinity) to $null.
    .PARAMETER Value
        The raw identifier or GUID string to normalize.
    .OUTPUTS
        System.String
    #>
    param([Parameter(Mandatory = $false)] [string] $Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $trimmed = $Value.Trim()
    if ($trimmed -match '^0{8}-0{4}-0{4}-0{4}-0{12}$') { return $null }
    if ($trimmed -match '[/\\?#]' -or $trimmed.Contains('..')) { return $null }
    return $trimmed
}

function ConvertTo-UtcTicks {
    <#
    .SYNOPSIS
        Parses a DateTime, DateTimeOffset, or ISO-8601 string into UTC ticks.
    .DESCRIPTION
        Converts a timestamp from ARM or Microsoft Graph into UTC ticks ([int64])
        using InvariantCulture and AssumeUniversal semantics so chronological
        comparisons are immune to fractional-second lexical sorting bugs and
        local culture differences. Returns 0 when the input is null or invalid.
    .PARAMETER DateValue
        The DateTime, DateTimeOffset, or ISO-8601 string to convert.
    .OUTPUTS
        System.Int64
    #>
    param([Parameter(Mandatory = $false)] $DateValue)
    if ($null -eq $DateValue) { return [int64]0 }
    if ($DateValue -is [DateTimeOffset]) { return $DateValue.UtcTicks }
    if ($DateValue -is [DateTime]) {
        if ($DateValue.Kind -eq [DateTimeKind]::Unspecified) {
            return ([DateTime]::SpecifyKind($DateValue, [DateTimeKind]::Utc)).Ticks
        }
        return $DateValue.ToUniversalTime().Ticks
    }
    $raw = [string]$DateValue
    if ([string]::IsNullOrWhiteSpace($raw)) { return [int64]0 }
    $dto = [DateTimeOffset]::MinValue
    if ([DateTimeOffset]::TryParse($raw.Trim(), [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$dto)) {
        return $dto.UtcTicks
    }
    return [int64]0
}

function Invoke-AvdPrimaryUserSleep {
    <#
    .SYNOPSIS
        Pauses execution for the specified number of seconds during retry backoff.
    .DESCRIPTION
        Thin wrapper around Start-Sleep used by Invoke-AvdPrimaryUserRest when
        honoring HTTP 429 Retry-After headers or transient 5xx exponential backoff.
        Separated into its own function so offline tests can intercept retry waits.
    .PARAMETER Seconds
        Number of seconds to sleep. Values <= 0 are ignored.
    .OUTPUTS
        None
    #>
    param([Parameter(Mandatory = $true)] [int] $Seconds)
    if ($Seconds -gt 0) {
        Start-Sleep -Seconds $Seconds
    }
}

function Invoke-AvdPrimaryUserHttpRequest {
    <#
    .SYNOPSIS
        Executes a single HTTP REST request via Invoke-RestMethod.
    .DESCRIPTION
        Low-level transport wrapper around Invoke-RestMethod with UseBasicParsing
        and a 100-second timeout for Windows PowerShell 5.1 compatibility.
        Separated from Invoke-AvdPrimaryUserRest so unit tests can mock raw HTTP
        responses and exceptions (such as HTTP 429 or 401).
    .PARAMETER Method
        HTTP method (Get, Post, Delete, Patch, Put).
    .PARAMETER Uri
        Full request URI.
    .PARAMETER Headers
        Request header hashtable including Authorization.
    .PARAMETER Body
        Optional JSON request body string.
    .OUTPUTS
        System.Object
    #>
    param(
        [Parameter(Mandatory = $true)] [string] $Method,
        [Parameter(Mandatory = $true)] [string] $Uri,
        [Parameter(Mandatory = $true)] [hashtable] $Headers,
        [Parameter(Mandatory = $false)] [string] $Body
    )
    $params = @{
        Method          = $Method
        Uri             = $Uri
        Headers         = $Headers
        UseBasicParsing = $true
        TimeoutSec      = 100
    }
    if (-not [string]::IsNullOrEmpty($Body)) {
        $params['Body'] = $Body
    }
    return (Invoke-RestMethod @params)
}

function Invoke-AvdPrimaryUserRest {
    <#
    .SYNOPSIS
        Invokes an ARM or Microsoft Graph REST request with retry and token refresh.
    .DESCRIPTION
        Wraps Invoke-AvdPrimaryUserHttpRequest with resilience for production runs:
        - Automatically refreshes Managed Identity / AzContext tokens once on HTTP 401
          (while preserving explicitly supplied -GraphAccessToken values).
        - Retries HTTP 429 (TooManyRequests) and transient 502/503/504 errors up
          to MaxAttempts, honoring the server's Retry-After response header or
          exponential backoff.
        - Extracts detailed Graph/ARM error messages from ErrorDetails.Message and
          attaches HttpStatusCode to thrown exceptions.
    .PARAMETER Method
        HTTP method (Get, Post, Delete, Patch, Put).
    .PARAMETER Uri
        Full ARM or Microsoft Graph request URI.
    .PARAMETER Headers
        Request header hashtable. Updated in place if an HTTP 401 token refresh occurs.
    .PARAMETER Body
        Optional JSON request body string.
    .PARAMETER MaxAttempts
        Maximum total attempts for transient HTTP 429/502/503/504 failures. Defaults to 4.
    .OUTPUTS
        System.Object
    #>
    param(
        [Parameter(Mandatory = $true)] [ValidateSet('Get', 'Post', 'Delete', 'Patch', 'Put')] [string] $Method,
        [Parameter(Mandatory = $true)] [string] $Uri,
        [Parameter(Mandatory = $true)] [hashtable] $Headers,
        [Parameter(Mandatory = $false)] [string] $Body,
        [Parameter(Mandatory = $false)] [int] $MaxAttempts = 4
    )
    $attempt = 0
    $refreshAttempted = $false
    while ($true) {
        $attempt++
        try {
            return (Invoke-AvdPrimaryUserHttpRequest -Method $Method -Uri $Uri -Headers $Headers -Body $Body)
        } catch {
            $statusCode = 0
            $retryAfterSec = 0
            $resp = $null
            if ($null -ne $_.Exception -and $null -ne $_.Exception.PSObject.Properties['Response']) {
                $resp = $_.Exception.Response
            }
            if ($null -ne $resp) {
                try { $statusCode = [int]$resp.StatusCode } catch { $statusCode = 0 }
                try {
                    if ($null -ne $resp.Headers) {
                        $raRaw = $null
                        if ($resp.Headers -is [System.Net.WebHeaderCollection]) {
                            $raRaw = $resp.Headers['Retry-After']
                        } elseif ($null -ne $resp.Headers.PSObject.Properties['RetryAfter'] -and $null -ne $resp.Headers.RetryAfter -and $null -ne $resp.Headers.RetryAfter.Delta) {
                            $raRaw = [int][Math]::Ceiling($resp.Headers.RetryAfter.Delta.TotalSeconds)
                        }
                        if (-not [string]::IsNullOrWhiteSpace([string]$raRaw)) {
                            $parsedInt = 0
                            if ([int]::TryParse([string]$raRaw, [ref]$parsedInt) -and $parsedInt -gt 0) {
                                $retryAfterSec = [Math]::Min($parsedInt, 300)
                            }
                        }
                    }
                } catch { $retryAfterSec = 0 }
            }
            if ($statusCode -eq 0 -and $null -ne $_.Exception.Data -and $_.Exception.Data.Contains('HttpStatusCode')) {
                $statusCode = [int]$_.Exception.Data['HttpStatusCode']
            }
            if ($retryAfterSec -eq 0 -and $null -ne $_.Exception.Data -and $_.Exception.Data.Contains('RetryAfterSeconds')) {
                $retryAfterSec = [int]$_.Exception.Data['RetryAfterSeconds']
            }
            if ($statusCode -eq 401 -and -not $refreshAttempted -and $Uri -notmatch 'oauth2/token|IDENTITY_ENDPOINT') {
                $resAudience = if ($Uri -like 'https://management.azure.com/*') { 'https://management.azure.com/' } elseif ($Uri -like 'https://graph.microsoft.com/*') { 'https://graph.microsoft.com/' } else { '' }
                $canRefreshGraph = ($resAudience -ne 'https://graph.microsoft.com/') -or ($null -eq (Get-Variable -Name 'GraphAccessToken' -ErrorAction SilentlyContinue)) -or [string]::IsNullOrWhiteSpace($GraphAccessToken)
                $authModeVar = Get-Variable -Name 'AuthMode' -ErrorAction SilentlyContinue
                if (-not [string]::IsNullOrWhiteSpace($resAudience) -and $canRefreshGraph -and $null -ne $authModeVar) {
                    $refreshAttempted = $true
                    $miClientVar = Get-Variable -Name 'ManagedIdentityClientId' -ErrorAction SilentlyContinue
                    $miClientVal = if ($null -ne $miClientVar) { [string]$miClientVar.Value } else { '' }
                    try {
                        $refreshedToken = Get-AvdPrimaryUserAccessToken -ResourceUri $resAudience -SelectedAuthMode ([string]$authModeVar.Value) -ClientId $miClientVal
                        $Headers['Authorization'] = "Bearer $refreshedToken"
                        continue
                    } catch {
                        Write-Verbose ("Token refresh on HTTP 401 failed: {0}" -f $_.Exception.Message)
                    }
                }
            }
            $isTransient = ($statusCode -eq 429 -or $statusCode -eq 502 -or $statusCode -eq 503 -or $statusCode -eq 504)
            if ($isTransient -and $attempt -lt $MaxAttempts) {
                if ($retryAfterSec -le 0) {
                    $retryAfterSec = [int][Math]::Min(60, [Math]::Pow(2, $attempt))
                }
                Write-Verbose ("HTTP {0} on {1} {2}; retrying in {3}s (attempt {4}/{5})..." -f $statusCode, $Method, $Uri, $retryAfterSec, $attempt, $MaxAttempts)
                Invoke-AvdPrimaryUserSleep -Seconds $retryAfterSec
                continue
            }
            $errDetail = ''
            if ($null -ne $_.ErrorDetails -and -not [string]::IsNullOrWhiteSpace($_.ErrorDetails.Message)) {
                $errDetail = $_.ErrorDetails.Message.Trim()
            }
            $baseMsg = $_.Exception.Message
            $fullMsg = if (-not [string]::IsNullOrWhiteSpace($errDetail) -and $baseMsg -notlike "*$errDetail*") {
                if ($statusCode -gt 0) { "HTTP $statusCode`: $baseMsg ($errDetail)" } else { "$baseMsg ($errDetail)" }
            } elseif ($statusCode -gt 0 -and $baseMsg -notlike "*$statusCode*") {
                "HTTP $statusCode`: $baseMsg"
            } else {
                $baseMsg
            }
            $ex = New-Object System.InvalidOperationException($fullMsg, $_.Exception)
            if ($statusCode -gt 0) {
                $ex.Data['HttpStatusCode'] = $statusCode
            }
            throw $ex
        }
    }
}

function Get-AvdPrimaryUserAccessToken {
    <#
    .SYNOPSIS
        Acquires an OAuth2 Bearer access token for ARM or Microsoft Graph.
    .DESCRIPTION
        Supports Azure Automation / App Service Managed Identity (IDENTITY_ENDPOINT
        and IDENTITY_HEADER), Azure VM Instance Metadata Service (IMDS), or an
        active Az.Accounts PowerShell session (Get-AzAccessToken), including
        SecureString token conversion on newer Az.Accounts versions.
    .PARAMETER ResourceUri
        Target resource audience URI (e.g., https://management.azure.com/ or https://graph.microsoft.com/).
    .PARAMETER SelectedAuthMode
        Authentication mode: Auto, ManagedIdentity, or AzContext.
    .PARAMETER ClientId
        Optional Client ID of a User-Assigned Managed Identity.
    .OUTPUTS
        System.String
    #>
    param(
        [Parameter(Mandatory = $true)] [string] $ResourceUri,
        [Parameter(Mandatory = $true)] [string] $SelectedAuthMode,
        [Parameter(Mandatory = $false)] [string] $ClientId
    )
    $useMi = ($SelectedAuthMode -eq 'ManagedIdentity') -or ($SelectedAuthMode -eq 'Auto' -and -not [string]::IsNullOrWhiteSpace($env:IDENTITY_ENDPOINT))
    if ($useMi) {
        if (-not [string]::IsNullOrWhiteSpace($env:IDENTITY_ENDPOINT) -and -not [string]::IsNullOrWhiteSpace($env:IDENTITY_HEADER)) {
            $uri = "{0}?resource={1}&api-version=2019-08-01" -f $env:IDENTITY_ENDPOINT, [uri]::EscapeDataString($ResourceUri)
            if (-not [string]::IsNullOrWhiteSpace($ClientId)) {
                $uri = "{0}&client_id={1}" -f $uri, [uri]::EscapeDataString($ClientId)
            }
            $headers = @{ 'X-IDENTITY-HEADER' = $env:IDENTITY_HEADER; 'Metadata' = 'true' }
        } else {
            $uri = "http://169.254.169.254/metadata/identity/oauth2/token?api-version=2018-02-01&resource={0}" -f [uri]::EscapeDataString($ResourceUri)
            if (-not [string]::IsNullOrWhiteSpace($ClientId)) {
                $uri = "{0}&client_id={1}" -f $uri, [uri]::EscapeDataString($ClientId)
            }
            $headers = @{ 'Metadata' = 'true' }
        }
        $resp = Invoke-AvdPrimaryUserRest -Method Get -Uri $uri -Headers $headers
        $token = Get-AvdPrimaryUserValue -InputObject $resp -PropertyName 'access_token'
        if ([string]::IsNullOrWhiteSpace($token)) {
            throw "Managed Identity token response for '$ResourceUri' did not contain an access_token."
        }
        return [string]$token
    }
    $azCmd = Get-Command -Name 'Get-AzAccessToken' -ErrorAction SilentlyContinue
    if ($null -eq $azCmd) {
        throw "Get-AzAccessToken is not available and Managed Identity was not detected. Connect with Connect-AzAccount or run under Managed Identity."
    }
    $tokenObj = Get-AzAccessToken -ResourceUrl $ResourceUri
    $rawToken = Get-AvdPrimaryUserValue -InputObject $tokenObj -PropertyName 'Token'
    if ($rawToken -is [System.Security.SecureString]) {
        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($rawToken)
        try {
            $rawToken = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
        } finally {
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        }
    }
    if ([string]::IsNullOrWhiteSpace([string]$rawToken)) {
        throw "Get-AzAccessToken returned an empty token for '$ResourceUri'."
    }
    return [string]$rawToken
}

function Invoke-AvdPrimaryUserPagedGet {
    <#
    .SYNOPSIS
        Retrieves all pages of an ARM or Microsoft Graph collection endpoint.
    .DESCRIPTION
        Repeatedly invokes GET requests starting at Uri and follows both ARM
        nextLink and OData @odata.nextLink continuation URLs until all items
        in the value array have been collected. Enforces HTTPS and same-origin
        authority on every continuation link and detects cyclic pagination.
    .PARAMETER Uri
        Initial collection GET URI.
    .PARAMETER Headers
        Request header hashtable including Authorization.
    .OUTPUTS
        System.Object[]
    #>
    param(
        [Parameter(Mandatory = $true)] [string] $Uri,
        [Parameter(Mandatory = $true)] [hashtable] $Headers
    )
    $baseUri = $null
    if (-not [System.Uri]::TryCreate($Uri, [System.UriKind]::Absolute, [ref]$baseUri) -or $baseUri.Scheme -ne 'https' -or -not [string]::IsNullOrEmpty($baseUri.UserInfo)) {
        throw "Invalid initial collection URI '$Uri'. Expected an absolute https:// URI."
    }
    $items = New-Object 'System.Collections.Generic.List[object]'
    $visitedUris = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $nextUri = $Uri
    $pageCount = 0
    while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
        $pageCount++
        if ($pageCount -gt 1000) {
            throw "Pagination exceeded maximum page limit (1000) starting from '$Uri'."
        }
        $parsedNext = $null
        if (-not [System.Uri]::TryCreate($nextUri, [System.UriKind]::Absolute, [ref]$parsedNext) -or
            $parsedNext.Scheme -ne 'https' -or
            -not [string]::IsNullOrEmpty($parsedNext.UserInfo) -or
            -not $parsedNext.Authority.Equals($baseUri.Authority, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Refusing to follow untrusted pagination link '$nextUri' (expected https://$($baseUri.Authority))."
        }
        if (-not $visitedUris.Add($nextUri)) {
            throw "Detected cyclic pagination link '$nextUri'."
        }
        $resp = Invoke-AvdPrimaryUserRest -Method Get -Uri $nextUri -Headers $Headers
        $hasValueProp = ($null -ne $resp) -and (($resp -is [System.Collections.IDictionary] -and $resp.Contains('value')) -or ($null -ne $resp.PSObject.Properties['value']))
        if (-not $hasValueProp) {
            throw "Collection response from '$nextUri' did not contain a 'value' array."
        }
        $val = Get-AvdPrimaryUserValue -InputObject $resp -PropertyName 'value'
        if ($null -ne $val) {
            foreach ($entry in @($val)) {
                if ($null -ne $entry) { [void]$items.Add($entry) }
            }
        }
        $nextLink = Get-AvdPrimaryUserValue -InputObject $resp -PropertyName 'nextLink'
        if ([string]::IsNullOrWhiteSpace([string]$nextLink)) {
            $nextLink = Get-AvdPrimaryUserValue -InputObject $resp -PropertyName '@odata.nextLink'
        }
        $nextUri = [string]$nextLink
    }
    return $items.ToArray()
}

function Get-AvdShortHostName {
    <#
    .SYNOPSIS
        Extracts the uppercase NetBIOS/short computer name for an AVD session host.
    .DESCRIPTION
        Derives the uppercase short hostname from an AVD sessionHost resource name
        (<hostPoolName>/<fqdn-or-hostname>) by stripping the pool prefix and any
        DNS domain suffix. Falls back to the leaf VM name of VmResourceId when
        SessionHostResourceName is empty.
    .PARAMETER SessionHostResourceName
        The ARM sessionHost resource name or FQDN.
    .PARAMETER VmResourceId
        Optional ARM resource ID of the underlying Azure virtual machine.
    .OUTPUTS
        System.String
    #>
    param(
        [Parameter(Mandatory = $false)] [string] $SessionHostResourceName,
        [Parameter(Mandatory = $false)] [string] $VmResourceId
    )
    if (-not [string]::IsNullOrWhiteSpace($SessionHostResourceName)) {
        $leaf = ($SessionHostResourceName -split '/')[-1].Trim()
        $short = ($leaf -split '\.')[0].Trim()
        if (-not [string]::IsNullOrWhiteSpace($short)) {
            return $short.ToUpperInvariant()
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($VmResourceId)) {
        $vmLeaf = ($VmResourceId -split '/')[-1].Trim()
        if (-not [string]::IsNullOrWhiteSpace($vmLeaf)) {
            return ($vmLeaf -split '\.')[0].Trim().ToUpperInvariant()
        }
    }
    return ''
}

function Get-IntuneWindowsDeviceIndex {
    <#
    .SYNOPSIS
        Queries Intune Windows managedDevices and indexes them by uppercase hostname.
    .DESCRIPTION
        Queries Microsoft Graph /v1.0/deviceManagement/managedDevices for Windows
        devices. When FilterHostNames is supplied, batches OData deviceName eq '...'
        clauses in chunks of up to 15 to avoid downloading the entire tenant's
        Windows inventory. Deduplicates records by Intune device id per hostname.
    .PARAMETER GraphHeaders
        Microsoft Graph request header hashtable including Authorization.
    .PARAMETER FilterHostNames
        Optional list of short hostnames to query in batched OData filters.
    .OUTPUTS
        System.Collections.Hashtable
    #>
    param(
        [Parameter(Mandatory = $true)] [hashtable] $GraphHeaders,
        [Parameter(Mandatory = $false)] [string[]] $FilterHostNames
    )
    $selectFields = 'id,deviceName,azureADDeviceId,userPrincipalName,userId,enrolledDateTime'
    $cleanNames = New-Object 'System.Collections.Generic.List[string]'
    $seenNames = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    if ($null -ne $FilterHostNames) {
        foreach ($n in $FilterHostNames) {
            if (-not [string]::IsNullOrWhiteSpace($n)) {
                $trimmed = $n.Trim().ToUpperInvariant()
                if ($seenNames.Add($trimmed)) {
                    [void]$cleanNames.Add($trimmed)
                }
            }
        }
    }
    if ($cleanNames.Count -eq 0) {
        return @{}
    }
    $allDevices = New-Object 'System.Collections.Generic.List[object]'
    $chunkSize = 15
    for ($i = 0; $i -lt $cleanNames.Count; $i += $chunkSize) {
        $endIdx = [Math]::Min($i + $chunkSize - 1, $cleanNames.Count - 1)
        $clauses = New-Object 'System.Collections.Generic.List[string]'
        for ($j = $i; $j -le $endIdx; $j++) {
            $escapedHost = $cleanNames[$j].Replace("'", "''")
            [void]$clauses.Add("deviceName eq '$escapedHost'")
        }
        $rawFilter = if ($clauses.Count -eq 1) {
            "operatingSystem eq 'Windows' and $($clauses[0])"
        } else {
            "operatingSystem eq 'Windows' and ($($clauses -join ' or '))"
        }
        $filter = [uri]::EscapeDataString($rawFilter)
        $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=$filter&`$select=$selectFields"
        $chunkDevs = @(Invoke-AvdPrimaryUserPagedGet -Uri $uri -Headers $GraphHeaders)
        foreach ($d in $chunkDevs) {
            if ($null -ne $d) { [void]$allDevices.Add($d) }
        }
    }
    $byHost = @{}
    $seenDeviceIdsPerHost = @{}
    foreach ($dev in $allDevices) {
        $devName = [string](Get-AvdPrimaryUserValue -InputObject $dev -PropertyName 'deviceName')
        if ([string]::IsNullOrWhiteSpace($devName)) { continue }
        $key = $devName.Trim().ToUpperInvariant()
        if (-not $byHost.ContainsKey($key)) {
            $byHost[$key] = New-Object 'System.Collections.Generic.List[object]'
            $seenDeviceIdsPerHost[$key] = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        }
        $devId = [string](Get-AvdPrimaryUserValue -InputObject $dev -PropertyName 'id')
        if ([string]::IsNullOrWhiteSpace($devId) -or $seenDeviceIdsPerHost[$key].Add($devId)) {
            [void]$byHost[$key].Add($dev)
        }
    }
    return $byHost
}

function Get-AvdTargetSessionHosts {
    <#
    .SYNOPSIS
        Retrieves AVD session host records from the specified Personal Host Pools.
    .DESCRIPTION
        Queries each ARM Host Pool in TargetPoolIds, verifies that hostPoolType is
        Personal, and retrieves its session hosts. When SingleSessionHost is
        provided, attempts a direct GET first and falls back to listing the pool's
        session hosts and matching by short hostname if the ARM resource uses an FQDN.
    .PARAMETER ArmHeaders
        Azure Resource Manager request header hashtable including Authorization.
    .PARAMETER TargetPoolIds
        One or more full ARM Host Pool resource IDs to evaluate.
    .PARAMETER SingleSessionHost
        Optional single session host name (short name or FQDN) to filter within the pool.
    .OUTPUTS
        PSCustomObject[]
    #>
    param(
        [Parameter(Mandatory = $true)] [hashtable] $ArmHeaders,
        [Parameter(Mandatory = $true)] [string[]] $TargetPoolIds,
        [Parameter(Mandatory = $false)] [string] $SingleSessionHost
    )
    $pools = New-Object 'System.Collections.Generic.List[object]'
    $seenPoolIds = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $poolIdPattern = '^/subscriptions/[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}/resourceGroups/[A-Za-z0-9_.()\-]{1,90}/providers/Microsoft\.DesktopVirtualization/hostPools/[A-Za-z0-9_.\-]{1,64}$'
    foreach ($poolId in $TargetPoolIds) {
        if ([string]::IsNullOrWhiteSpace($poolId)) { continue }
        $trimmedPoolId = $poolId.Trim()
        if ($trimmedPoolId -notmatch $poolIdPattern -or $trimmedPoolId.Contains('..')) {
            throw "Invalid HostPoolResourceId '$trimmedPoolId'. Expected '/subscriptions/<subGuid>/resourceGroups/<rg>/providers/Microsoft.DesktopVirtualization/hostPools/<poolName>'."
        }
        if (-not $seenPoolIds.Add($trimmedPoolId)) { continue }
        $uri = "https://management.azure.com{0}?api-version=2024-04-03" -f $trimmedPoolId
        $pool = Invoke-AvdPrimaryUserRest -Method Get -Uri $uri -Headers $ArmHeaders
        $props = Get-AvdPrimaryUserValue -InputObject $pool -PropertyName 'properties'
        $poolType = [string](Get-AvdPrimaryUserValue -InputObject $props -PropertyName 'hostPoolType')
        if ($poolType -ieq 'Personal') {
            [void]$pools.Add($pool)
        } else {
            Write-Verbose ("Skipping host pool '{0}' because hostPoolType is '{1}' (not Personal)." -f $trimmedPoolId, $poolType)
        }
    }
    $records = New-Object 'System.Collections.Generic.List[object]'
    foreach ($pool in $pools) {
        $poolId = [string](Get-AvdPrimaryUserValue -InputObject $pool -PropertyName 'id')
        if ($poolId -notmatch $poolIdPattern -or $poolId.Contains('..')) {
            throw "ARM returned an invalid host pool resource ID '$poolId'."
        }
        $poolName = [string](Get-AvdPrimaryUserValue -InputObject $pool -PropertyName 'name')
        if ([string]::IsNullOrWhiteSpace($poolName) -and -not [string]::IsNullOrWhiteSpace($poolId)) {
            $poolName = ($poolId -split '/')[-1]
        }
        $resolvedDirectly = $false
        $requestedLeaf = ''
        $requestedHasDomain = $false
        if (-not [string]::IsNullOrWhiteSpace($SingleSessionHost)) {
            $requestedLeaf = ($SingleSessionHost -split '/')[-1].Trim()
            if ($requestedLeaf -match '[/?#\\]' -or $requestedLeaf.Contains('..')) {
                throw "Invalid TargetSessionHostName '$SingleSessionHost'."
            }
            $requestedHasDomain = $requestedLeaf.Contains('.')
            $shUri = "https://management.azure.com{0}/sessionHosts/{1}?api-version=2024-04-03" -f $poolId, [uri]::EscapeDataString($requestedLeaf)
            try {
                $shObj = Invoke-AvdPrimaryUserRest -Method Get -Uri $shUri -Headers $ArmHeaders
                if ($null -ne $shObj) {
                    [void]$records.Add([PSCustomObject]@{
                        HostPoolId   = $poolId
                        HostPoolName = $poolName
                        SessionHost  = $shObj
                    })
                    $resolvedDirectly = $true
                }
            } catch {
                $directErrCode = 0
                if ($null -ne $_.Exception.Data -and $_.Exception.Data.Contains('HttpStatusCode')) {
                    $directErrCode = [int]$_.Exception.Data['HttpStatusCode']
                }
                if ($directErrCode -ne 404 -and $directErrCode -ne 400 -and $_.Exception.Message -notmatch '\b404\b') {
                    throw
                }
                Write-Verbose ("Direct sessionHost GET for '{0}' returned HTTP {1}; falling back to pool sessionHosts list." -f $requestedLeaf, $directErrCode)
            }
        }
        if (-not $resolvedDirectly) {
            $targetShort = if (-not [string]::IsNullOrWhiteSpace($SingleSessionHost)) { Get-AvdShortHostName -SessionHostResourceName $SingleSessionHost } else { '' }
            $shListUri = "https://management.azure.com{0}/sessionHosts?api-version=2024-04-03" -f $poolId
            $shItems = @(Invoke-AvdPrimaryUserPagedGet -Uri $shListUri -Headers $ArmHeaders)
            foreach ($shObj in $shItems) {
                if (-not [string]::IsNullOrWhiteSpace($targetShort)) {
                    $shName = [string](Get-AvdPrimaryUserValue -InputObject $shObj -PropertyName 'name')
                    $shLeaf = ($shName -split '/')[-1].Trim()
                    if ($requestedHasDomain) {
                        if ($shLeaf -ine $requestedLeaf) { continue }
                    } else {
                        $shProps = Get-AvdPrimaryUserValue -InputObject $shObj -PropertyName 'properties'
                        $vmResId = [string](Get-AvdPrimaryUserValue -InputObject $shProps -PropertyName 'resourceId')
                        $candidateShort = Get-AvdShortHostName -SessionHostResourceName $shName -VmResourceId $vmResId
                        if ($candidateShort -ine $targetShort) { continue }
                    }
                }
                [void]$records.Add([PSCustomObject]@{
                    HostPoolId   = $poolId
                    HostPoolName = $poolName
                    SessionHost  = $shObj
                })
            }
        }
    }
    return $records.ToArray()
}

function Resolve-EntraUserObject {
    <#
    .SYNOPSIS
        Resolves an AVD assignedUser UPN or email to an Entra ID User ObjectId.
    .DESCRIPTION
        Queries Microsoft Graph /v1.0/users/{upn}?$select=id,userPrincipalName and
        caches resolved user objects in Cache. If direct lookup returns HTTP 404
        (common for B2B guest users whose UPN contains #EXT# while AVD stores their
        mail address), falls back to an OData filter on userPrincipalName or mail.
        Negative-caches genuine 404 NotFound results and throws HTTP 409 when
        multiple Entra users match the fallback filter.
    .PARAMETER UserPrincipalName
        The assignedUser UPN or mail string from the AVD session host.
    .PARAMETER GraphHeaders
        Microsoft Graph request header hashtable including Authorization.
    .PARAMETER Cache
        Per-run hashtable used to cache positive and 404 user lookup results.
    .OUTPUTS
        PSCustomObject
    #>
    param(
        [Parameter(Mandatory = $true)] [string] $UserPrincipalName,
        [Parameter(Mandatory = $true)] [hashtable] $GraphHeaders,
        [Parameter(Mandatory = $true)] [hashtable] $Cache
    )
    $trimmedUpn = $UserPrincipalName.Trim()
    $key = $trimmedUpn.ToLowerInvariant()
    if ($Cache.ContainsKey($key)) {
        $cached = $Cache[$key]
        if ($null -ne (Get-AvdPrimaryUserValue -InputObject $cached -PropertyName 'NotFound') -and $cached.NotFound) {
            $nfEx = New-Object System.InvalidOperationException($cached.ErrorMessage)
            $nfEx.Data['HttpStatusCode'] = 404
            throw $nfEx
        }
        return $cached
    }
    $escaped = [uri]::EscapeDataString($trimmedUpn)
    $uri = "https://graph.microsoft.com/v1.0/users/{0}?`$select=id,userPrincipalName" -f $escaped
    $userResp = $null
    try {
        $userResp = Invoke-AvdPrimaryUserRest -Method Get -Uri $uri -Headers $GraphHeaders
    } catch {
        $status = 0
        if ($null -ne $_.Exception.Data -and $_.Exception.Data.Contains('HttpStatusCode')) {
            $status = [int]$_.Exception.Data['HttpStatusCode']
        }
        if ($status -eq 404 -or $_.Exception.Message -match '\b404\b|Request_ResourceNotFound') {
            $escapedFilterVal = $trimmedUpn.Replace("'", "''")
            $filter = [uri]::EscapeDataString("userPrincipalName eq '$escapedFilterVal' or mail eq '$escapedFilterVal'")
            $fallbackUri = "https://graph.microsoft.com/v1.0/users?`$filter=$filter&`$select=id,userPrincipalName"
            $matchesList = @(Invoke-AvdPrimaryUserPagedGet -Uri $fallbackUri -Headers $GraphHeaders)
            if ($matchesList.Count -eq 1) {
                $userResp = $matchesList[0]
            } elseif ($matchesList.Count -gt 1) {
                $ambEx = New-Object System.InvalidOperationException("Multiple Entra users matched '$trimmedUpn' via UPN/mail filter.")
                $ambEx.Data['HttpStatusCode'] = 409
                throw $ambEx
            } else {
                $msg = "Entra user '$trimmedUpn' was not found (HTTP 404)."
                $Cache[$key] = [PSCustomObject]@{ NotFound = $true; ErrorMessage = $msg }
                $nfEx = New-Object System.InvalidOperationException($msg, $_.Exception)
                $nfEx.Data['HttpStatusCode'] = 404
                throw $nfEx
            }
        } else {
            throw
        }
    }
    $userId = ConvertTo-NormalizedPrincipalId -Value ([string](Get-AvdPrimaryUserValue -InputObject $userResp -PropertyName 'id'))
    if ([string]::IsNullOrWhiteSpace($userId)) {
        $msg = "Entra user lookup for '$trimmedUpn' returned no valid user id."
        $Cache[$key] = [PSCustomObject]@{ NotFound = $true; ErrorMessage = $msg }
        $nfEx = New-Object System.InvalidOperationException($msg)
        $nfEx.Data['HttpStatusCode'] = 404
        throw $nfEx
    }
    $resolvedUpn = [string](Get-AvdPrimaryUserValue -InputObject $userResp -PropertyName 'userPrincipalName')
    if ([string]::IsNullOrWhiteSpace($resolvedUpn)) { $resolvedUpn = $trimmedUpn }
    $entry = [PSCustomObject]@{
        Id                = $userId
        UserPrincipalName = $resolvedUpn
        NotFound          = $false
    }
    $Cache[$key] = $entry
    return $entry
}

function Get-IntuneDevicePrimaryUsers {
    <#
    .SYNOPSIS
        Reads the authoritative Primary User relationship for an Intune managedDevice.
    .DESCRIPTION
        Queries GET /beta/deviceManagement/managedDevices/{id}/users?$select=id,userPrincipalName
        to inspect the actual Primary User relationship modified by POST .../users/$ref.
        Normalizes zero GUIDs and falls back to the managedDevice scalar userId and
        userPrincipalName properties only when the relationship collection is empty.
    .PARAMETER IntuneDeviceId
        The Intune managedDevice GUID.
    .PARAMETER IntuneDeviceObject
        The Intune managedDevice object from device inventory.
    .PARAMETER GraphHeaders
        Microsoft Graph request header hashtable including Authorization.
    .OUTPUTS
        PSCustomObject[]
    #>
    param(
        [Parameter(Mandatory = $true)] [string] $IntuneDeviceId,
        [Parameter(Mandatory = $true)] [object] $IntuneDeviceObject,
        [Parameter(Mandatory = $true)] [hashtable] $GraphHeaders
    )
    $relUri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/{0}/users?`$select=id,userPrincipalName" -f [uri]::EscapeDataString($IntuneDeviceId)
    $relItems = @(Invoke-AvdPrimaryUserPagedGet -Uri $relUri -Headers $GraphHeaders)
    $users = New-Object 'System.Collections.Generic.List[object]'
    foreach ($u in $relItems) {
        $uid = ConvertTo-NormalizedPrincipalId -Value ([string](Get-AvdPrimaryUserValue -InputObject $u -PropertyName 'id'))
        $upn = [string](Get-AvdPrimaryUserValue -InputObject $u -PropertyName 'userPrincipalName')
        if (-not [string]::IsNullOrWhiteSpace($uid) -or -not [string]::IsNullOrWhiteSpace($upn)) {
            [void]$users.Add([PSCustomObject]@{
                Id                = $uid
                UserPrincipalName = if ([string]::IsNullOrWhiteSpace($upn)) { $null } else { $upn.Trim() }
            })
        }
    }
    if ($users.Count -eq 0) {
        $scalarId  = ConvertTo-NormalizedPrincipalId -Value ([string](Get-AvdPrimaryUserValue -InputObject $IntuneDeviceObject -PropertyName 'userId'))
        $scalarUpn = [string](Get-AvdPrimaryUserValue -InputObject $IntuneDeviceObject -PropertyName 'userPrincipalName')
        if ([string]::IsNullOrWhiteSpace($scalarUpn)) { $scalarUpn = $null } else { $scalarUpn = $scalarUpn.Trim() }
        if ($null -ne $scalarId -or $null -ne $scalarUpn) {
            [void]$users.Add([PSCustomObject]@{
                Id                = $scalarId
                UserPrincipalName = $scalarUpn
            })
        }
    }
    return $users.ToArray()
}

# --- MAIN EXECUTION ---
$poolIdValidationPattern = '^/subscriptions/[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}/resourceGroups/[A-Za-z0-9_.()\-]{1,90}/providers/Microsoft\.DesktopVirtualization/hostPools/[A-Za-z0-9_.\-]{1,64}$'
$effectivePoolIds = New-Object 'System.Collections.Generic.List[string]'
$seenEffectivePools = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
if ($null -ne $HostPoolResourceId) {
    foreach ($id in $HostPoolResourceId) {
        if (-not [string]::IsNullOrWhiteSpace($id)) {
            $cleanId = $id.Trim()
            if ($cleanId -notmatch $poolIdValidationPattern -or $cleanId.Contains('..')) {
                throw "Invalid HostPoolResourceId '$cleanId'. Expected '/subscriptions/<subGuid>/resourceGroups/<rg>/providers/Microsoft.DesktopVirtualization/hostPools/<poolName>'."
            }
            if ($seenEffectivePools.Add($cleanId)) {
                [void]$effectivePoolIds.Add($cleanId)
            }
        }
    }
}
if (-not [string]::IsNullOrWhiteSpace($HostPoolName) -or -not [string]::IsNullOrWhiteSpace($ResourceGroupName)) {
    if ([string]::IsNullOrWhiteSpace($HostPoolName) -or [string]::IsNullOrWhiteSpace($ResourceGroupName)) {
        throw "Both -HostPoolName and -ResourceGroupName must be provided together."
    }
    $resolvedSubId = if (-not [string]::IsNullOrWhiteSpace($SubscriptionId)) {
        $SubscriptionId.Trim()
    } else {
        $azCtxCmd = Get-Command -Name 'Get-AzContext' -ErrorAction SilentlyContinue
        if ($null -ne $azCtxCmd) {
            $ctx = Get-AzContext
            if ($null -ne $ctx -and $null -ne $ctx.Subscription) { [string]$ctx.Subscription.Id } else { '' }
        } else { '' }
    }
    if ([string]::IsNullOrWhiteSpace($resolvedSubId)) {
        throw "When specifying -HostPoolName and -ResourceGroupName, either pass -SubscriptionId or have an active Get-AzContext subscription."
    }
    $constructedPoolId = "/subscriptions/{0}/resourceGroups/{1}/providers/Microsoft.DesktopVirtualization/hostPools/{2}" -f $resolvedSubId, $ResourceGroupName.Trim(), $HostPoolName.Trim()
    if ($constructedPoolId -notmatch $poolIdValidationPattern -or $constructedPoolId.Contains('..')) {
        throw "Invalid SubscriptionId, ResourceGroupName, or HostPoolName in '$constructedPoolId'."
    }
    if ($seenEffectivePools.Add($constructedPoolId)) {
        [void]$effectivePoolIds.Add($constructedPoolId)
    }
}
if ($effectivePoolIds.Count -eq 0) {
    throw "Specify target Personal Host Pool(s) via -HostPoolResourceId or -ResourceGroupName and -HostPoolName."
}
$armToken   = Get-AvdPrimaryUserAccessToken -ResourceUri 'https://management.azure.com/' -SelectedAuthMode $AuthMode -ClientId $ManagedIdentityClientId
$graphToken = if (-not [string]::IsNullOrWhiteSpace($GraphAccessToken)) {
    $GraphAccessToken.Trim()
} else {
    Get-AvdPrimaryUserAccessToken -ResourceUri 'https://graph.microsoft.com/' -SelectedAuthMode $AuthMode -ClientId $ManagedIdentityClientId
}
$armHeaders   = @{ Authorization = "Bearer $armToken";   'Content-Type' = 'application/json' }
$graphHeaders = @{ Authorization = "Bearer $graphToken"; 'Content-Type' = 'application/json' }
$sessionHostRecords = @(Get-AvdTargetSessionHosts -ArmHeaders $armHeaders -TargetPoolIds $effectivePoolIds.ToArray() -SingleSessionHost $TargetSessionHostName)
if ($sessionHostRecords.Count -eq 0) {
    Write-Verbose 'No Personal Host Pool session hosts found to evaluate.'
    return @()
}
$shortHostCounts = @{}
foreach ($rec in $sessionHostRecords) {
    $shName = [string](Get-AvdPrimaryUserValue -InputObject $rec.SessionHost -PropertyName 'name')
    $shProps = Get-AvdPrimaryUserValue -InputObject $rec.SessionHost -PropertyName 'properties'
    $vmId = [string](Get-AvdPrimaryUserValue -InputObject $shProps -PropertyName 'resourceId')
    $short = Get-AvdShortHostName -SessionHostResourceName $shName -VmResourceId $vmId
    if (-not [string]::IsNullOrWhiteSpace($short)) {
        if (-not $shortHostCounts.ContainsKey($short)) {
            $shortHostCounts[$short] = 1
        } else {
            $shortHostCounts[$short] = $shortHostCounts[$short] + 1
        }
    }
}
$targetShortNames = New-Object 'System.Collections.Generic.List[string]'
foreach ($rec in $sessionHostRecords) {
    $shName = [string](Get-AvdPrimaryUserValue -InputObject $rec.SessionHost -PropertyName 'name')
    $shProps = Get-AvdPrimaryUserValue -InputObject $rec.SessionHost -PropertyName 'properties'
    $assigned = [string](Get-AvdPrimaryUserValue -InputObject $shProps -PropertyName 'assignedUser')
    if ([string]::IsNullOrWhiteSpace($assigned)) { continue }
    $vmId = [string](Get-AvdPrimaryUserValue -InputObject $shProps -PropertyName 'resourceId')
    $short = Get-AvdShortHostName -SessionHostResourceName $shName -VmResourceId $vmId
    if (-not [string]::IsNullOrWhiteSpace($short) -and $shortHostCounts[$short] -eq 1 -and -not $targetShortNames.Contains($short)) {
        [void]$targetShortNames.Add($short)
    }
}
$intuneIndex = if ($targetShortNames.Count -gt 0) {
    Get-IntuneWindowsDeviceIndex -GraphHeaders $graphHeaders -FilterHostNames $targetShortNames.ToArray()
} else {
    @{}
}
$userCache = @{}
$results = New-Object 'System.Collections.Generic.List[object]'
foreach ($rec in $sessionHostRecords) {
    $shObj = $rec.SessionHost
    $rawShName = [string](Get-AvdPrimaryUserValue -InputObject $shObj -PropertyName 'name')
    $shLeafName = ($rawShName -split '/')[-1]
    $shProps = Get-AvdPrimaryUserValue -InputObject $shObj -PropertyName 'properties'
    $rawAssignedUpn = [string](Get-AvdPrimaryUserValue -InputObject $shProps -PropertyName 'assignedUser')
    $assignedUpn = if ([string]::IsNullOrWhiteSpace($rawAssignedUpn)) { $null } else { $rawAssignedUpn.Trim() }
    $vmResourceId = [string](Get-AvdPrimaryUserValue -InputObject $shProps -PropertyName 'resourceId')
    $shortHost = Get-AvdShortHostName -SessionHostResourceName $rawShName -VmResourceId $vmResourceId
    if ([string]::IsNullOrWhiteSpace($shortHost)) {
        [void]$results.Add([PSCustomObject]@{
            HostPoolName        = $rec.HostPoolName
            SessionHostName     = $shLeafName
            ShortHostName       = ''
            AssignedUser        = $assignedUpn
            IntuneDeviceId      = $null
            PreviousPrimaryUser = $null
            NewPrimaryUser      = $null
            Action              = 'None'
            SyncTriggered       = $false
            Status              = 'InvalidSessionHostName'
            Detail              = 'Could not derive a short hostname from the AVD session host.'
        })
        continue
    }
    if ($shortHostCounts.ContainsKey($shortHost) -and $shortHostCounts[$shortHost] -gt 1) {
        [void]$results.Add([PSCustomObject]@{
            HostPoolName        = $rec.HostPoolName
            SessionHostName     = $shLeafName
            ShortHostName       = $shortHost
            AssignedUser        = $assignedUpn
            IntuneDeviceId      = $null
            PreviousPrimaryUser = $null
            NewPrimaryUser      = $null
            Action              = 'None'
            SyncTriggered       = $false
            Status              = 'AmbiguousAvdSessionHost'
            Detail              = ("Multiple evaluated AVD session hosts share the short hostname '{0}'. Skipping to prevent cross-pool Primary User collision." -f $shortHost)
        })
        continue
    }
    # Case 1: AVD Session Host has no AssignedUser -> skip (fail-closed)
    if ([string]::IsNullOrWhiteSpace($assignedUpn)) {
        [void]$results.Add([PSCustomObject]@{
            HostPoolName        = $rec.HostPoolName
            SessionHostName     = $shLeafName
            ShortHostName       = $shortHost
            AssignedUser        = $null
            IntuneDeviceId      = $null
            PreviousPrimaryUser = $null
            NewPrimaryUser      = $null
            Action              = 'None'
            SyncTriggered       = $false
            Status              = 'UnassignedInAvd'
            Detail              = 'No user assigned in AVD.'
        })
        continue
    }
    if (-not $intuneIndex.ContainsKey($shortHost)) {
        [void]$results.Add([PSCustomObject]@{
            HostPoolName        = $rec.HostPoolName
            SessionHostName     = $shLeafName
            ShortHostName       = $shortHost
            AssignedUser        = $assignedUpn
            IntuneDeviceId      = $null
            PreviousPrimaryUser = $null
            NewPrimaryUser      = $null
            Action              = 'None'
            SyncTriggered       = $false
            Status              = 'PendingIntuneEnrollment'
            Detail              = "Assigned to '$assignedUpn' in AVD, but device '$shortHost' is not enrolled in Intune yet."
        })
        continue
    }
    $candidates = $intuneIndex[$shortHost].ToArray()
    if ($candidates.Count -gt 1) {
        [void]$results.Add([PSCustomObject]@{
            HostPoolName        = $rec.HostPoolName
            SessionHostName     = $shLeafName
            ShortHostName       = $shortHost
            AssignedUser        = $assignedUpn
            IntuneDeviceId      = $null
            PreviousPrimaryUser = $null
            NewPrimaryUser      = $null
            Action              = 'None'
            SyncTriggered       = $false
            Status              = 'AmbiguousIntuneDevice'
            Detail              = ("Found {0} Intune managedDevice records named '{1}'. Remove stale Intune records before syncing." -f $candidates.Count, $shortHost)
        })
        continue
    }
    $intuneDevice   = $candidates[0]
    $intuneDeviceId = ConvertTo-NormalizedPrincipalId -Value ([string](Get-AvdPrimaryUserValue -InputObject $intuneDevice -PropertyName 'id'))
    $aadDeviceId    = ConvertTo-NormalizedPrincipalId -Value ([string](Get-AvdPrimaryUserValue -InputObject $intuneDevice -PropertyName 'azureADDeviceId'))
    # Guard against stale Intune device records from a previous VM lifecycle
    if ([string]::IsNullOrWhiteSpace($aadDeviceId)) {
        [void]$results.Add([PSCustomObject]@{
            HostPoolName        = $rec.HostPoolName
            SessionHostName     = $shLeafName
            ShortHostName       = $shortHost
            AssignedUser        = $assignedUpn
            IntuneDeviceId      = $intuneDeviceId
            PreviousPrimaryUser = $null
            NewPrimaryUser      = $null
            Action              = 'None'
            SyncTriggered       = $false
            Status              = 'StaleIntuneDeviceRecord'
            Detail              = "Intune device '$shortHost' ($intuneDeviceId) has an empty or zero azureADDeviceId."
        })
        continue
    }
    $shCreatedRaw   = Get-AvdPrimaryUserValue -InputObject (Get-AvdPrimaryUserValue -InputObject $shObj -PropertyName 'systemData') -PropertyName 'createdAt'
    $shCreatedTicks = ConvertTo-UtcTicks -DateValue $shCreatedRaw
    $enrollTicks    = ConvertTo-UtcTicks -DateValue (Get-AvdPrimaryUserValue -InputObject $intuneDevice -PropertyName 'enrolledDateTime')
    # Allow a 24-hour provisioning window because AADLoginForWindows enrolls the VM before DSC registers the AVD sessionHost
    if ($shCreatedTicks -gt 0 -and $enrollTicks -gt 0 -and $enrollTicks -lt ($shCreatedTicks - [TimeSpan]::FromHours(24).Ticks)) {
        [void]$results.Add([PSCustomObject]@{
            HostPoolName        = $rec.HostPoolName
            SessionHostName     = $shLeafName
            ShortHostName       = $shortHost
            AssignedUser        = $assignedUpn
            IntuneDeviceId      = $intuneDeviceId
            PreviousPrimaryUser = $null
            NewPrimaryUser      = $null
            Action              = 'None'
            SyncTriggered       = $false
            Status              = 'StaleIntuneDeviceRecord'
            Detail              = "Intune device '$shortHost' ($intuneDeviceId) enrolled more than 24 hours before the current AVD session host was created."
        })
        continue
    }
    # Resolve target user in Entra ID first so ownership comparison uses immutable ObjectId
    $resolvedUser = $null
    try {
        $resolvedUser = Resolve-EntraUserObject -UserPrincipalName $assignedUpn -GraphHeaders $graphHeaders -Cache $userCache
    } catch {
        $errCode = 0
        if ($null -ne $_.Exception.Data -and $_.Exception.Data.Contains('HttpStatusCode')) {
            $errCode = [int]$_.Exception.Data['HttpStatusCode']
        }
        $lookupStatus = if ($errCode -eq 404) { 'EntraUserNotFound' } elseif ($errCode -eq 409) { 'AmbiguousEntraUser' } else { 'GraphUserLookupFailed' }
        [void]$results.Add([PSCustomObject]@{
            HostPoolName        = $rec.HostPoolName
            SessionHostName     = $shLeafName
            ShortHostName       = $shortHost
            AssignedUser        = $assignedUpn
            IntuneDeviceId      = $intuneDeviceId
            PreviousPrimaryUser = $null
            NewPrimaryUser      = $null
            Action              = 'None'
            SyncTriggered       = $false
            Status              = $lookupStatus
            Detail              = $_.Exception.Message
        })
        continue
    }
    # Query authoritative Intune Primary User relationship (/users)
    $primaryUsers = @()
    try {
        $primaryUsers = @(Get-IntuneDevicePrimaryUsers -IntuneDeviceId $intuneDeviceId -IntuneDeviceObject $intuneDevice -GraphHeaders $graphHeaders)
    } catch {
        [void]$results.Add([PSCustomObject]@{
            HostPoolName        = $rec.HostPoolName
            SessionHostName     = $shLeafName
            ShortHostName       = $shortHost
            AssignedUser        = $assignedUpn
            IntuneDeviceId      = $intuneDeviceId
            PreviousPrimaryUser = $null
            NewPrimaryUser      = $null
            Action              = 'None'
            SyncTriggered       = $false
            Status              = 'GraphLookupFailed'
            Detail              = "Failed to read Primary User relationship for '$shortHost': $($_.Exception.Message)"
        })
        continue
    }
    if ($primaryUsers.Count -gt 1) {
        [void]$results.Add([PSCustomObject]@{
            HostPoolName        = $rec.HostPoolName
            SessionHostName     = $shLeafName
            ShortHostName       = $shortHost
            AssignedUser        = $assignedUpn
            IntuneDeviceId      = $intuneDeviceId
            PreviousPrimaryUser = ($primaryUsers | ForEach-Object { $_.UserPrincipalName }) -join ','
            NewPrimaryUser      = $null
            Action              = 'None'
            SyncTriggered       = $false
            Status              = 'PrimaryUserConflict'
            Detail              = "Intune device '$shortHost' has multiple primary users returned by Graph."
        })
        continue
    }
    $currentPrimary = if ($primaryUsers.Count -eq 1) { $primaryUsers[0] } else { $null }
    $currentUserId  = if ($null -ne $currentPrimary) { $currentPrimary.Id } else { $null }
    $currentUpn     = if ($null -ne $currentPrimary) { $currentPrimary.UserPrincipalName } else { $null }
    # Case 2: Already in sync by immutable Entra User ObjectId
    if ($null -ne $currentUserId -and $currentUserId.Equals($resolvedUser.Id, [System.StringComparison]::OrdinalIgnoreCase)) {
        [void]$results.Add([PSCustomObject]@{
            HostPoolName        = $rec.HostPoolName
            SessionHostName     = $shLeafName
            ShortHostName       = $shortHost
            AssignedUser        = $assignedUpn
            IntuneDeviceId      = $intuneDeviceId
            PreviousPrimaryUser = if ($null -ne $currentUpn) { $currentUpn } else { $resolvedUser.UserPrincipalName }
            NewPrimaryUser      = $resolvedUser.UserPrincipalName
            Action              = 'None'
            SyncTriggered       = $false
            Status              = 'AlreadyInSync'
            Detail              = "Intune Primary User ObjectId ($currentUserId) already matches '$assignedUpn'."
        })
        continue
    }
    # Case 3: Different Primary User already exists -> fail closed (PrimaryUserConflict)
    if ($null -ne $currentUserId -or $null -ne $currentUpn) {
        $existingDesc = if ($null -ne $currentUpn) { "$currentUpn (userId=$currentUserId)" } else { "userId=$currentUserId" }
        [void]$results.Add([PSCustomObject]@{
            HostPoolName        = $rec.HostPoolName
            SessionHostName     = $shLeafName
            ShortHostName       = $shortHost
            AssignedUser        = $assignedUpn
            IntuneDeviceId      = $intuneDeviceId
            PreviousPrimaryUser = $currentUpn
            NewPrimaryUser      = $currentUpn
            Action              = 'None'
            SyncTriggered       = $false
            Status              = 'PrimaryUserConflict'
            Detail              = "Intune Primary User is '$existingDesc', which differs from AVD AssignedUser '$assignedUpn' ($($resolvedUser.Id))."
        })
        continue
    }
    # Case 4: Populate missing Primary User in Intune and optionally trigger syncDevice
    $targetLabel = "Intune ManagedDevice '$shortHost' ($intuneDeviceId)"
    $actionLabel = "Set Primary User to '$($resolvedUser.UserPrincipalName)' ($($resolvedUser.Id))"
    if ($PSCmdlet.ShouldProcess($targetLabel, $actionLabel)) {
        $escapedDeviceId = [uri]::EscapeDataString($intuneDeviceId)
        $escapedUserId   = [uri]::EscapeDataString($resolvedUser.Id)
        $refUri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/{0}/users/`$ref" -f $escapedDeviceId
        $bodyJson = @{ '@odata.id' = ("https://graph.microsoft.com/beta/users/{0}" -f $escapedUserId) } | ConvertTo-Json -Compress
        try {
            $null = Invoke-AvdPrimaryUserRest -Method Post -Uri $refUri -Headers $graphHeaders -Body $bodyJson
            $syncTriggered = $false
            $finalStatus = 'Updated'
            $syncDetail = ''
            if ($SyncIntuneDevice) {
                try {
                    $syncUri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/{0}/syncDevice" -f $escapedDeviceId
                    $null = Invoke-AvdPrimaryUserRest -Method Post -Uri $syncUri -Headers $graphHeaders
                    $syncTriggered = $true
                    $syncDetail = ' Triggered Intune syncDevice.'
                } catch {
                    $finalStatus = 'PrimaryUserUpdatedSyncFailed'
                    $syncDetail = " Primary User updated, but syncDevice failed: $($_.Exception.Message)"
                }
            }
            [void]$results.Add([PSCustomObject]@{
                HostPoolName        = $rec.HostPoolName
                SessionHostName     = $shLeafName
                ShortHostName       = $shortHost
                AssignedUser        = $assignedUpn
                IntuneDeviceId      = $intuneDeviceId
                PreviousPrimaryUser = $null
                NewPrimaryUser      = $resolvedUser.UserPrincipalName
                Action              = 'SetPrimaryUser'
                SyncTriggered       = $syncTriggered
                Status              = $finalStatus
                Detail              = ("Set Intune Primary User to '{0}' ({1}).{2}" -f $resolvedUser.UserPrincipalName, $resolvedUser.Id, $syncDetail)
            })
        } catch {
            [void]$results.Add([PSCustomObject]@{
                HostPoolName        = $rec.HostPoolName
                SessionHostName     = $shLeafName
                ShortHostName       = $shortHost
                AssignedUser        = $assignedUpn
                IntuneDeviceId      = $intuneDeviceId
                PreviousPrimaryUser = $null
                NewPrimaryUser      = $null
                Action              = 'SetPrimaryUser'
                SyncTriggered       = $false
                Status              = 'GraphUpdateFailed'
                Detail              = $_.Exception.Message
            })
        }
    } else {
        [void]$results.Add([PSCustomObject]@{
            HostPoolName        = $rec.HostPoolName
            SessionHostName     = $shLeafName
            ShortHostName       = $shortHost
            AssignedUser        = $assignedUpn
            IntuneDeviceId      = $intuneDeviceId
            PreviousPrimaryUser = $null
            NewPrimaryUser      = $resolvedUser.UserPrincipalName
            Action              = 'WhatIfSetPrimaryUser'
            SyncTriggered       = $false
            Status              = 'WhatIf'
            Detail              = ("Would set Intune Primary User to '{0}' ({1})." -f $resolvedUser.UserPrincipalName, $resolvedUser.Id)
        })
    }
}
return $results.ToArray()
