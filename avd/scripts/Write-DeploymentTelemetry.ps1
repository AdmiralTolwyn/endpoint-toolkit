<#
.SYNOPSIS
    Sends deployment telemetry to a Log Analytics workspace using the HTTP Data Collector API.

.DESCRIPTION
    Records structured AVD session host replacement events for operational dashboards and auditing.
    Requires a Log Analytics Workspace ID and Shared Key (store in Key Vault).

    ┌──────────────────────────────────────────────────────────────────────────┐
    │ DEPRECATION NOTICE                                                        │
    │ The HTTP Data Collector API used by this script is DEPRECATED. Microsoft  │
    │ ends support on 2026-09-14. The migration target is the Logs Ingestion   │
    │ API (Data Collection Rule / Data Collection Endpoint + Entra ID auth).   │
    │ This script does NOT implement that migration yet — it only carries a    │
    │ deprecation notice plus env-var input support. Plan the DCR/DCE + Entra  │
    │ auth rework before 2026-09-14.                                           │
    └──────────────────────────────────────────────────────────────────────────┘

.PARAMETER WorkspaceId
    The Log Analytics Workspace ID (GUID).
    Falls back to $env:WORKSPACE_ID if not bound.

.PARAMETER SharedKey
    The Primary or Secondary Key for the Log Analytics workspace.
    Falls back to $env:SHARED_KEY if not bound.

.PARAMETER LogType
    The custom log table name. Defaults to 'AVDDeployment'.
    Data appears in Log Analytics as 'AVDDeployment_CL'.

.PARAMETER EventData
    A hashtable of event data to send. Example keys:
    - HostPoolName, Stage (Canary/Blast/Cleanup), Action (Deploy/Drain/Decommission)
    - VMCount, ImageVersion, Status (Success/Failed), Duration, PipelineRunId
    Falls back to $env:EVENT_DATA_JSON (a JSON object string) if not bound.

.EXAMPLE
    .\Write-DeploymentTelemetry.ps1 `
        -WorkspaceId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -SharedKey "base64key==" `
        -EventData @{
            HostPoolName  = "vdpool-mypool-001"
            Stage         = "Canary"
            Action        = "Deploy"
            VMCount       = 1
            ImageVersion  = "1.0.3"
            Status        = "Success"
            PipelineRunId = $env:BUILD_BUILDID
        }
#>
param(
    [Parameter(Mandatory = $false)]
    [string]$WorkspaceId,

    [Parameter(Mandatory = $false)]
    [string]$SharedKey,

    [string]$LogType = 'AVDDeployment',

    [Parameter(Mandatory = $false)]
    [hashtable]$EventData
)

$ErrorActionPreference = "Stop"

# ── ENV-VAR FALLBACKS ──────────────────────────────────────────────────────
# Params stay fully backward compatible: only fall back to environment
# variables when the caller didn't bind the corresponding parameter.
if (-not $PSBoundParameters.ContainsKey('WorkspaceId') -and -not [string]::IsNullOrWhiteSpace($env:WORKSPACE_ID)) {
    $WorkspaceId = $env:WORKSPACE_ID
}
if (-not $PSBoundParameters.ContainsKey('SharedKey') -and -not [string]::IsNullOrWhiteSpace($env:SHARED_KEY)) {
    $SharedKey = $env:SHARED_KEY
}
if (-not $PSBoundParameters.ContainsKey('EventData') -and -not [string]::IsNullOrWhiteSpace($env:EVENT_DATA_JSON)) {
    try {
        $ParsedEventData = $env:EVENT_DATA_JSON | ConvertFrom-Json
        $EventData = @{}
        foreach ($Prop in $ParsedEventData.PSObject.Properties) {
            $EventData[$Prop.Name] = $Prop.Value
        }
    } catch {
        # Telemetry failures should never block a deployment
        Write-Warning "Failed to parse EVENT_DATA_JSON: $($_.Exception.Message)"
    }
}

if ([string]::IsNullOrWhiteSpace($WorkspaceId) -or [string]::IsNullOrWhiteSpace($SharedKey) -or -not $EventData) {
    # Telemetry failures should never block a deployment
    Write-Warning "Telemetry skipped: missing WorkspaceId/SharedKey/EventData (params or WORKSPACE_ID/SHARED_KEY/EVENT_DATA_JSON env vars)."
    return
}

# Inject timestamp
if (-not $EventData.ContainsKey('TimeGenerated')) {
    $EventData['TimeGenerated'] = (Get-Date -AsUTC).ToString('o')
}

# Serialize payload
$body = ConvertTo-Json -InputObject @($EventData) -Depth 10
$contentLength = [System.Text.Encoding]::UTF8.GetByteCount($body)

# Build HMAC-SHA256 authorization header
$rfc1123date = [DateTime]::UtcNow.ToString("r")
$stringToHash  = "POST`n$contentLength`napplication/json`nx-ms-date:${rfc1123date}`n/api/logs"
$bytesToHash   = [System.Text.Encoding]::UTF8.GetBytes($stringToHash)
$keyBytes      = [Convert]::FromBase64String($SharedKey)

$sha256        = New-Object System.Security.Cryptography.HMACSHA256
$sha256.Key    = $keyBytes
$calculatedHash = $sha256.ComputeHash($bytesToHash)
$encodedHash   = [Convert]::ToBase64String($calculatedHash)
$authorization = 'SharedKey {0}:{1}' -f $WorkspaceId, $encodedHash

# Send to Data Collector API
$uri = "https://${WorkspaceId}.ods.opinsights.azure.com/api/logs?api-version=2016-04-01"

$headers = @{
    "Authorization"        = $authorization
    "Log-Type"             = $LogType
    "x-ms-date"            = $rfc1123date
    "time-generated-field" = "TimeGenerated"
}

try {
    $response = Invoke-RestMethod -Uri $uri `
                                  -Method Post `
                                  -ContentType 'application/json' `
                                  -Headers $headers `
                                  -Body ([System.Text.Encoding]::UTF8.GetBytes($body))
    Write-Host " [TELEMETRY] Event sent to Log Analytics ($LogType)." -ForegroundColor DarkGray
} catch {
    # Telemetry failures should never block a deployment
    Write-Warning "Telemetry send failed: $($_.Exception.Message)"
}
