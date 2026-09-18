#Requires -Version 5.1

function Get-W365UserExperienceSyncAssessment {
    param(
        [object]$Policy,
        [ValidateSet('Review','Enabled','Disabled')][string]$ExpectedState = 'Review'
    )
    $Configuration = $Policy.userSettingsPersistenceConfiguration
    $Enabled = if ($Configuration.userSettingsPersistenceEnabled -is [bool]) { $Configuration.userSettingsPersistenceEnabled } else { $null }
    $Storage = if ($Configuration.userSettingsPersistenceStorageSizeCategory -cin @('fourGB','eightGB','sixteenGB','thirtyTwoGB','sixtyFourGB')) { $Configuration.userSettingsPersistenceStorageSizeCategory } else { $null }
    $Result = [ordered]@{
        PolicyId = [string]$Policy.id
        ProvisioningType = [string]$Policy.provisioningType
        Enabled = $Enabled
        StorageSizeCategory = $Storage
        ExpectedState = $ExpectedState
        ApiVersion = 'beta'
        Status = 'Error'
        Details = 'Policy intent only; does not prove assignment, applied configuration, profile persistence or sufficient storage capacity.'
    }
    if ($Policy.managedBy -cne 'windows365') {
        $Result.Details = 'Windows 365 policy ownership is missing or unsupported. ' + $Result.Details
    } elseif ($Policy.provisioningType -cin @('dedicated','shared','sharedByUser','reserve')) {
        $Result.Status = 'N/A'
        $Result.Details = 'User settings persistence applies only to sharedByEntraGroup policies. ' + $Result.Details
    } elseif ($Policy.provisioningType -cne 'sharedByEntraGroup') {
        $Result.Details = 'Provisioning type is missing or unsupported. ' + $Result.Details
    } elseif ($null -eq $Enabled -or ($Enabled -and $null -eq $Storage)) {
        $Result.Details = 'Explicit Boolean enablement and a documented storage enum when enabled are required. ' + $Result.Details
    } elseif ($ExpectedState -eq 'Review') {
        $Result.Details = 'No customer enablement target selected. ' + $Result.Details
    } else {
        $Result.Status = if ($Enabled -eq ($ExpectedState -eq 'Enabled')) { 'Pass' } else { 'Warning' }
        $Result.Details = "Collected enabled=$Enabled; storage=$Storage; customer target=$ExpectedState. " + $Result.Details
    }
    return [pscustomobject]$Result
}

function Invoke-W365UserExperienceSyncRead {
    param(
        [ValidateSet('Review','Enabled','Disabled')][string]$ExpectedState = 'Review',
        [scriptblock]$Request = {
            param($RequestUri)
            Invoke-MgGraphRequest -Method GET -Uri $RequestUri -Headers @{ Prefer = 'include-unknown-enum-members' } -ErrorAction Stop
        }
    )
    Add-Type -AssemblyName System.Web
    $Endpoint = 'https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies'
    $Fields = 'id,managedBy,provisioningType,userSettingsPersistenceConfiguration'
    $Next = $Endpoint + '?$select=' + $Fields
    $Visited = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $Identifiers = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $Results = [Collections.Generic.List[object]]::new()
    while ($Next) {
        $Parsed = [uri]$Next
        if ($Parsed.Scheme -cne 'https' -or $Parsed.Host -cne 'graph.microsoft.com' -or $Parsed.Port -ne 443 -or $Parsed.UserInfo -or $Parsed.Fragment -or $Parsed.AbsolutePath -cne '/beta/deviceManagement/virtualEndpoint/provisioningPolicies') { throw 'Untrusted UX Sync continuation URI' }
        $Query = [Web.HttpUtility]::ParseQueryString($Parsed.Query)
        if ($Query['$select'] -cne $Fields) { throw 'UX Sync projection changed' }
        foreach ($Key in $Query.AllKeys) {
            if ($Key -cnotin @('$select','$skiptoken') -or $Query.GetValues($Key).Count -ne 1) { throw 'Unreviewed UX Sync query' }
        }
        if (-not $Visited.Add($Next) -or $Visited.Count -gt 100) { throw 'UX Sync page limit or repeated continuation' }
        $Response = & $Request $Next
        if ($null -eq $Response.value -or $Response.value -is [string] -or $Response.value -isnot [System.Collections.IList]) { throw 'Invalid UX Sync collection response' }
        foreach ($Policy in $Response.value) {
            if ($Policy.id -isnot [string] -or $Policy.id -cnotmatch '^[A-Za-z0-9-]{1,128}$' -or -not $Identifiers.Add($Policy.id)) { throw 'Missing, invalid or duplicate policy ID' }
            if ($Results.Count -ge 10000) { throw 'UX Sync row limit exceeded' }
            $Results.Add((Get-W365UserExperienceSyncAssessment -Policy $Policy -ExpectedState $ExpectedState))
        }
        $Next = $Response.'@odata.nextLink'
        if ($Next -and $Response.value.Count -eq 0) { throw 'Empty UX Sync page with continuation' }
    }
    return $Results.ToArray()
}