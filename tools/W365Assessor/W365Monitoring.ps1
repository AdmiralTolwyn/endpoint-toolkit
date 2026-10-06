<#
.SYNOPSIS
    Endpoint Analytics and update-summary evidence projection for the Windows 365 collector.
.DESCRIPTION
    Dot-sourced by Invoke-W365Discovery.ps1. Keeps typed score and counter values and records
    missing or invalid fields as Unknown. Device names and identifiers are not exported.
.NOTES
    Author    : Anton Romanyuk
    Requires  : Windows PowerShell 5.1 or PowerShell 7
    Disclaimer: This script is provided "AS IS" with no warranties and confers no rights.
#>
#Requires -Version 5.1


<#
.SYNOPSIS
    Converts Endpoint Analytics device-score rows to tenant score evidence.
.PARAMETER Rows
    Rows from userExperienceAnalyticsDeviceScores. At most 10000.
.OUTPUTS
    One PSCustomObject per row with ScoreEntryId, HealthStatus, Scores and FieldState.
#>
function ConvertTo-W365AnalyticsEvidence {
    param([object[]]$Rows)
    if ($Rows.Count -gt 10000) { throw 'Analytics evidence row limit exceeded' }
    foreach ($Row in $Rows) {
        $State = 'Complete'
        $EntryId = if ($Row.id -is [string] -and $Row.id -cmatch '^[A-Za-z0-9-]{1,128}$') { $Row.id } else { $null }
        $Health = if ($Row.healthStatus -cin @('unknown','insufficientData','needsAttention','meetingGoals')) { $Row.healthStatus } else { $null }
        if ($null -eq $EntryId -or $null -eq $Health) { $State = 'Partial' }
        $Scores = [ordered]@{}
        foreach ($Field in @('endpointAnalyticsScore','startupPerformanceScore','appReliabilityScore','workFromAnywhereScore','meanResourceSpikeTimeScore','batteryHealthScore')) {
            $Value = $Row.$Field
            $Numeric = $Value -is [double] -or $Value -is [single] -or $Value -is [decimal] -or $Value -is [int] -or $Value -is [long]
            if ($Numeric -and $Value -eq -1) { $Scores[$Field] = @{ State = 'Unavailable'; Value = $null } }
            elseif ($Numeric -and $Value -ge 0 -and $Value -le 100) { $Scores[$Field] = @{ State = 'Observed'; Value = $Value } }
            else { $Scores[$Field] = @{ State = 'Unknown'; Value = $null }; $State = 'Partial' }
        }
        [pscustomobject]@{
            ScoreEntryId = $EntryId
            HealthStatus = $Health
            Scores = $Scores
            FieldState = $State
            IdentityMatch = 'NotEvaluated'
            DataAge = 'NotAvailableInContract'
        }
    }
}

<#
.SYNOPSIS
    Converts the software-update status summary to typed counter evidence.
.PARAMETER Response
    Response from softwareUpdateStatusSummary, with or without a value wrapper.
.OUTPUTS
    PSCustomObject with Counts, UnknownFields and FieldState.
#>
function ConvertTo-W365UpdateSummaryEvidence {
    param([object]$Response)
    $Fields = @('compliantDeviceCount','nonCompliantDeviceCount','remediatedDeviceCount','errorDeviceCount','unknownDeviceCount','conflictDeviceCount','notApplicableDeviceCount')
    if ($Response -isnot [Collections.IDictionary] -and $Response -isnot [pscustomobject]) { throw 'Invalid update-summary object' }
    $HasWrapper = if ($Response -is [Collections.IDictionary]) { $Response.Contains('value') } else { $null -ne $Response.PSObject.Properties['value'] }
    $Summary = $Response
    $Shape = 'Object'
    if ($HasWrapper) {
        foreach ($Field in $Fields) {
            $HasField = if ($Response -is [Collections.IDictionary]) { $Response.Contains($Field) } else { $null -ne $Response.PSObject.Properties[$Field] }
            if ($HasField) { throw 'Ambiguous update-summary response' }
        }
        $Summary = $Response.value
        $Shape = 'ValueObject'
        if ($Summary -isnot [Collections.IDictionary] -and $Summary -isnot [pscustomobject]) { throw 'Invalid wrapped update-summary object' }
    }
    $Counts = [ordered]@{}
    $UnknownFields = [Collections.Generic.List[string]]::new()
    foreach ($Field in $Fields) {
        $Value = $Summary.$Field
        if (($Value -is [int] -or $Value -is [long]) -and $Value -ge 0 -and $Value -le [int]::MaxValue) { $Counts[$Field] = $Value }
        else { $Counts[$Field] = $null; $UnknownFields.Add($Field) }
    }
    return [pscustomobject]@{
        Counts = $Counts
        UnknownFields = @($UnknownFields.ToArray())
        FieldState = $(if ($UnknownFields.Count -eq 0) { 'Complete' } else { 'Partial' })
        ResponseShape = $Shape
        Scope = 'TenantSummaryNotCloudPcScoped'
        DataAge = 'NotAvailableInContract'
    }
}