#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Invoke-IntuneDiscovery.ps1') -LibraryOnly
function Assert-True([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
$Base = 'https://graph.microsoft.com/v1.0/deviceManagement/managedDevices'
foreach ($Uri in @('http://graph.microsoft.com/v1.0/deviceManagement/managedDevices', 'https://example.com/v1.0/deviceManagement/managedDevices', 'https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/id/wipe', 'https://graph.microsoft.com/beta/deviceManagement/managedDevices', 'https://graph.microsoft.com/v1.0/users')) {
    Assert-True (-not (Test-IntuneUri $Uri)) "Unsafe URI accepted: $Uri"
}
Assert-True (Test-IntuneUri $Base) 'Managed devices endpoint rejected'
$script:Calls = 0
$Result = Get-IntuneCollection -Uri $Base -Module ManagedDevices -Request {
    param($Uri)
    $script:Calls++
    if ($script:Calls -eq 1) { return @{ StatusCode = 200; Body = @{ value = @(); '@odata.nextLink' = $Base + '?$skiptoken=opaque-page-2' } } }
    return @{ StatusCode = 200; Body = @{ value = @(@{ id = 'device-1'; deviceName = 'One'; isEncrypted = $false; activationLockBypassCode = 'SECRET'; notes = 'SECRET' }) } }
}
Assert-True ($Result.Status.State -eq 'Complete' -and $Result.Status.PagesRead -eq 2 -and $Result.Rows.Count -eq 1) 'Empty-page paging failed'
Assert-True (-not (($Result | ConvertTo-Json -Depth 10).Contains('SECRET'))) 'Secret fields leaked'
Assert-True ($Result.Rows[0]['isEncrypted'] -eq $false) 'Explicit false lost'
$Role = ConvertTo-IntuneSafeRow -Module RoleDefinitions -Row @{ id = 'role-1'; rolePermissions = @(@{ resourceActions = @(@{ allowedResourceActions = @('read'); notAllowedResourceActions = @('write'); password = 'SECRET' }); secret = 'SECRET' }) }
Assert-True ($Role['rolePermissions'][0].resourceActions[0].allowedResourceActions[0] -eq 'read') 'Nested role permission lost'
Assert-True (-not (($Role | ConvertTo-Json -Depth 10).Contains('SECRET'))) 'Nested role secret leaked'
foreach ($Code in @(401, 403, 500)) {
    $Result = Get-IntuneCollection -Uri $Base -Module ManagedDevices -Request { @{ StatusCode = $Code } }
    Assert-True ($Result.Status.State -eq 'Error' -and $Result.Rows.Count -eq 0) "HTTP $Code inferred absence"
}
$script:Calls = 0
$script:Delays = 0
$Result = Get-IntuneCollection -Uri $Base -Module ManagedDevices -Request { $script:Calls++; @{ StatusCode = 429; RetryAfter = 1 } } -Delay { param($Seconds) $script:Delays++ }
Assert-True ($script:Calls -eq 4 -and $script:Delays -eq 3 -and $Result.Status.ErrorCode -eq 'Throttled') 'Retry bound failed'
$Result = Get-IntuneCollection -Uri $Base -Module ManagedDevices -Request { @{ StatusCode = 200; Body = @{ value = @(); '@odata.nextLink' = $Base } } }
Assert-True ($Result.Status.ErrorCode -eq 'PaginationCycle') 'Pagination cycle accepted'
$Result = Get-IntuneCollection -Uri $Base -Module ManagedDevices -Request { @{ StatusCode = 200; Body = @{ value = @(); '@odata.nextLink' = 'https://example.com/token' } } }
Assert-True ($Result.Status.ErrorCode -eq 'BlockedEndpoint') 'Cross-host nextLink accepted'
foreach ($Body in @(@{}, @{ value = $null }, @{ value = 'bad' }, @{ value = @(); error = @{ code = 'partial' } })) {
    $Result = Get-IntuneCollection -Uri $Base -Module ManagedDevices -Request { @{ StatusCode = 200; Body = $Body } }
    Assert-True ($Result.Status.ErrorCode -eq 'InvalidResponse') 'Malformed response accepted'
}
$Result = Get-IntuneCollection -Uri $Base -Module ManagedDevices -Request { @{ StatusCode = 200; Body = @{ value = @(@{ id = 'same' }, @{ id = 'same' }) } } }
Assert-True ($Result.Status.ErrorCode -eq 'DuplicateIdentity') 'Duplicate ID accepted'
$Result = Get-IntuneCollection -Uri $Base -Module ManagedDevices -MaxRows 1 -Request { @{ StatusCode = 200; Body = @{ value = @(@{ id = 'one' }, @{ id = 'two' }) } } }
Assert-True ($Result.Status.State -eq 'Partial' -and $Result.Rows.Count -eq 1) 'Row cap lost partial state'
$script:Paths = [System.Collections.Generic.List[string]]::new()
$Document = Invoke-IntuneDiscoveryCore -SelectedTenant '22222222-2222-4222-8222-222222222222' -Rbac $true -Audit $false -Request {
    param($Uri)
    $script:Paths.Add($Uri)
    $Values = @()
    if ($Uri.EndsWith('/managedDevices')) { $Values = @(@{ id = 'device-1'; operatingSystem = 'Windows'; deviceName = 'Example'; managementAgent = 'mdm'; enrolledDateTime = '2026-09-01T00:00:00Z'; lastSyncDateTime = [datetime]::UtcNow.AddHours(-1).ToString('o') }) }
    if ($Uri.EndsWith('/deviceCompliancePolicies')) { $Values = @(@{ id = 'policy-1'; displayName = 'Example'; lastModifiedDateTime = '2026-09-01T00:00:00Z' }) }
    if ($Uri.EndsWith('/policy-1/deviceStatuses')) { $Values = @(@{ id = 'report-1'; status = 'nonCompliant'; lastReportedDateTime = [datetime]::UtcNow.AddHours(-1).ToString('o') }) }
    if ($Uri.EndsWith('/mobileApps')) { $Values = @(@{ id = 'app-1'; displayName = 'Example app' }) }
    return @{ StatusCode = 200; Body = @{ value = $Values } }
} -Requirements @{ Assessor = 'Offline fixture'; ScopeDescription = 'Visible policy observations and managed devices'; ScopeConfirmed = $true; MaxCollectionAgeHours = 24; MaxPolicyReportAgeHours = 48; MaxDeviceSyncAgeDays = 7 }
Assert-True ($Document.Observations.Count -eq 2) 'Observation serialization failed'
Assert-True ($Document.CollectionStatus.ComplianceStates.CompletedParentIds.Count -eq 1) 'Child coverage missing'
Assert-True ($script:Paths.Contains('https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps/app-1/assignments')) 'Independent app collection missing'
Assert-True ($Document.CollectionStatus.EnrollmentPolicies.State -eq 'NotRequested') 'Unimplemented module hidden'
Add-Type -AssemblyName System.Net.Http
if (-not ('IntuneOfflineHttpClient' -as [type])) {
    Add-Type -ReferencedAssemblies @('System.Net.Http', [System.Net.HttpStatusCode].Assembly.Location) -TypeDefinition @'
public class IntuneOfflineHttpClient {
    public System.Threading.Tasks.Task<System.Net.Http.HttpResponseMessage> GetAsync(string uri) {
        var response = new System.Net.Http.HttpResponseMessage(System.Net.HttpStatusCode.OK);
        response.Content = new System.Net.Http.StringContent("{\"value\":[]}");
        return System.Threading.Tasks.Task.FromResult(response);
    }
}
'@
}
$script:IntuneHttpClient = New-Object IntuneOfflineHttpClient
$Page = Get-IntuneHttpPage $Base
Assert-True ($Page.StatusCode -eq 200 -and (Get-IntuneValue $Page.Body 'value') -is [array]) 'HTTP response adapter contract failed'
$script:IntuneHttpClient = $null
if ($env:ASSAY_INTUNE_FIXTURE) {
    [System.IO.File]::WriteAllText($env:ASSAY_INTUNE_FIXTURE, ($Document | ConvertTo-Json -Depth 30), [System.Text.UTF8Encoding]::new($false))
}
Write-Output 'PASS: Intune collector offline transport, paging, retry, identity, sanitization, child coverage and export tests.'