<#
.SYNOPSIS
    Sends deployment telemetry to a Log Analytics workspace using the HTTP Data Collector API.

.DESCRIPTION
    Records structured AVD session host replacement events for operational dashboards and auditing.
    Requires a Log Analytics Workspace ID and Shared Key (store in Key Vault).

.PARAMETER WorkspaceId
    The Log Analytics Workspace ID (GUID).

.PARAMETER SharedKey
    The Primary or Secondary Key for the Log Analytics workspace.

.PARAMETER LogType
    The custom log table name. Defaults to 'AVDDeployment'.
    Data appears in Log Analytics as 'AVDDeployment_CL'.

.PARAMETER EventData
    A hashtable of event data to send. Example keys:
    - HostPoolName, Stage (Canary/Blast/Cleanup), Action (Deploy/Drain/Decommission)
    - VMCount, ImageVersion, Status (Success/Failed), Duration, PipelineRunId

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
    [Parameter(Mandatory)]
    [string]$WorkspaceId,

    [Parameter(Mandatory)]
    [string]$SharedKey,

    [string]$LogType = 'AVDDeployment',

    [Parameter(Mandatory)]
    [hashtable]$EventData
)

$ErrorActionPreference = "Stop"

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
