<#
.SYNOPSIS
    AVD Update Factory - Drain Orchestrator

.DESCRIPTION
    Sets AVD Session Hosts to Drain Mode (AllowNewSession = $false).
    
    This script is the "Maintenance Worker" of the AVD Update Factory. It accepts a JSON list
    of host names (identified as "Outdated" by the Pre-Flight script) and disables new 
    user sessions on them.

    KEY FEATURES:
    - **Environment Variable Input:** Prioritizes $env:HOST_LIST_JSON to avoid CLI parsing errors.
    - **High-Visibility Logging:** Uses a custom visual format for Azure DevOps logs.
    - **Silent Execution:** Suppresses default cmdlet output for cleaner logs.
    - **Idempotency:** Can run repeatedly without failure.

.PARAMETER HostPoolName
    The name of the target Azure Virtual Desktop Host Pool.

.PARAMETER ResourceGroupName
    The Resource Group containing the Host Pool.

.PARAMETER HostListJson
    (Optional) A JSON-formatted string array of host names. 
    **BEST PRACTICE:** Pass this via the 'env:' block in YAML as HOST_LIST_JSON.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)] [string]$HostPoolName,
    [Parameter(Mandatory = $true)] [string]$ResourceGroupName,
    [Parameter(Mandatory = $false)] [string]$HostListJson
)

# --- CONFIGURATION: SILENCE INFRASTRUCTURE NOISE ---
$ProgressPreference = 'SilentlyContinue' # Hides Azure ASCII progress bars
$ErrorActionPreference = 'Stop'          # Fail fast on critical errors

# --- HELPER: VISUAL FORMATTING ---
function Write-Header {
    param ([string]$Title, [hashtable]$Info)
    $Width = 85
    $Border = "=" * $Width
    
    Write-Host "`n$Border" -ForegroundColor Cyan
    Write-Host "   $Title" -ForegroundColor Cyan
    Write-Host "$Border" -ForegroundColor Cyan
    
    foreach ($Key in $Info.Keys) {
        $Label = "$Key".PadRight(30)
        Write-Host " $Label : $($Info[$Key])" -ForegroundColor White
    }
    Write-Host "$Border`n" -ForegroundColor Cyan
}

function Write-Log {
    param ([string]$Message, [ValidateSet("INFO", "SUCCESS", "WARN", "ERROR")] $Level = "INFO")
    $Time = (Get-Date).ToString("HH:mm:ss")
    
    switch ($Level) {
        "INFO"    { Write-Host " [$Time] [INFO]    $Message" -ForegroundColor Gray }
        "SUCCESS" { Write-Host " [$Time] [OK]      $Message" -ForegroundColor Green }
        "WARN"    { Write-Host " [$Time] [WARN]    $Message" -ForegroundColor Yellow }
        "ERROR"   { Write-Host " [$Time] [ERROR]   $Message" -ForegroundColor Red }
    }
}

try {
    # -------------------------------------------------------------------------
    # 1. INITIALIZATION & CONTEXT CAPTURE
    # -------------------------------------------------------------------------
    $Ctx = Get-AzContext
    $HeaderInfo = [ordered]@{
        "Action"          = "DRAIN SESSION HOSTS"
        "Host Pool"       = $HostPoolName
        "Resource Group"  = $ResourceGroupName
        "Subscription"    = "$($Ctx.Subscription.Name) ($($Ctx.Subscription.Id))"
        "Status"          = "Updating Maintenance State"
    }

    Write-Header -Title "AVD UPDATE FACTORY: MAINTENANCE STAGE" -Info $HeaderInfo

    # -------------------------------------------------------------------------
    # 2. INPUT RESCUE (Environment Variable Fallback)
    # -------------------------------------------------------------------------
    # YAML pipelines often mangle JSON passed as arguments. We prefer Env Vars.
    if ([string]::IsNullOrWhiteSpace($HostListJson)) {
        if (-not [string]::IsNullOrWhiteSpace($env:HOST_LIST_JSON)) {
            Write-Log "Picking up input from Environment Variable (HOST_LIST_JSON)." "INFO"
            $HostListJson = $env:HOST_LIST_JSON
        }
    }

    # DEBUG: Print exactly what we received (useful for trace, but kept quiet)
    Write-Log "Raw Input Received: $HostListJson" "INFO"

    # -------------------------------------------------------------------------
    # 3. INPUT CLEANING & PARSING
    # -------------------------------------------------------------------------
    # Remove potential surrounding quotes passed by YAML variables
    $CleanJson = $HostListJson.Trim("'").Trim('"')

    # Graceful exit if list is empty or "[]"
    if ([string]::IsNullOrWhiteSpace($CleanJson) -or $CleanJson -eq "[]") {
        Write-Log "No outdated hosts provided in the plan. Nothing to drain." "SUCCESS"
        exit 0
    }

    try {
        $Hosts = $CleanJson | ConvertFrom-Json
    }
    catch {
        Write-Log "JSON PARSING FAILED. Input was: $CleanJson" "ERROR"
        throw $_
    }

    Write-Log "Detected $($Hosts.Count) host(s) requiring Drain Mode." "INFO"

    # -------------------------------------------------------------------------
    # 4. EXECUTION LOOP
    # -------------------------------------------------------------------------
    foreach ($H in $Hosts) {
        # Handle "HostPoolName/SessionHostName" format often returned by Azure
        # We split by "/" and take the last part to get the pure VM Name
        $TargetName = if ($H -like "*/*") { $H.Split("/")[-1] } else { $H }

        Write-Log "Targeting Session Host: $TargetName" "INFO"
        
        try {
            # Get session host details for resource ID and session count
            $SessionHost = Get-AzWvdSessionHost -HostPoolName $HostPoolName `
                                                -ResourceGroupName $ResourceGroupName `
                                                -Name $TargetName `
                                                -ErrorAction SilentlyContinue

            # NOTIFY ACTIVE USERS before draining (Item 3)
            if ($SessionHost -and $SessionHost.Session -gt 0) {
                Write-Log "Notifying $($SessionHost.Session) active session(s) on $TargetName..." "INFO"
                $Sessions = Get-AzWvdUserSession -HostPoolName $HostPoolName `
                                                 -ResourceGroupName $ResourceGroupName `
                                                 -SessionHostName $TargetName -ErrorAction SilentlyContinue
                foreach ($UserSession in $Sessions) {
                    $SessionId = $UserSession.Name.Split('/')[-1]
                    try {
                        Send-AzWvdUserSessionMessage -HostPoolName $HostPoolName `
                            -ResourceGroupName $ResourceGroupName `
                            -SessionHostName $TargetName `
                            -UserSessionId $SessionId `
                            -MessageTitle "Scheduled Maintenance" `
                            -MessageBody "This desktop is being replaced with a new version. Please save your work and sign out. Your session will be available on a new host shortly." `
                            -ErrorAction SilentlyContinue | Out-Null
                    } catch { }
                }
            }

            # SET DRAIN MODE
            Update-AzWvdSessionHost -HostPoolName $HostPoolName `
                                    -ResourceGroupName $ResourceGroupName `
                                    -Name $TargetName `
                                    -AllowNewSession:$false `
                                    -ErrorAction Stop | Out-Null
            
            Write-Log "SUCCESS: $TargetName is now in Drain Mode." "SUCCESS"

            # SET TAGS: PendingDrainTimestamp + ScalingPlanExclusion (Items 2 & 4)
            if ($SessionHost -and $SessionHost.ResourceId) {
                $DrainTimestamp = (Get-Date -AsUTC).ToString('o')
                $TagsToSet = @{
                    'PendingDrainTimestamp' = $DrainTimestamp
                    'ScalingPlanExclusion'  = 'True'
                }
                Update-AzTag -ResourceId $SessionHost.ResourceId -Tag $TagsToSet -Operation Merge | Out-Null
                Write-Log "Tags set: PendingDrainTimestamp=$DrainTimestamp, ScalingPlanExclusion=True" "INFO"
            }
        }
        catch {
            # Log warning but continue processing other hosts
            Write-Log "FAILED to update $TargetName." "WARN"
            Write-Log "Detail: $($_.Exception.Message)" "ERROR"
        }
    }

    Write-Log "Maintenance task completed successfully." "SUCCESS"
}
catch {
    Write-Log "FATAL ERROR in Drain Script: $($_.Exception.Message)" "ERROR"
    exit 1
}