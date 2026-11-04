<#
.SYNOPSIS
    Runs offline regression checks for Sync-AvdPersonalPrimaryUser.ps1.
.DESCRIPTION
    Parses and executes Sync-AvdPersonalPrimaryUser.ps1 through the PowerShell AST
    with mocked ARM and Microsoft Graph REST responses. Verifies:
    - AST comment-based help completeness for all 13 functions and parameters
    - Strict ARM HostPoolResourceId and pagination nextLink origin/HTTPS validation
    - Personal vs Pooled host pool filtering and cross-pool duplicate hostname detection
    - Primary User assignment via Graph beta users/$ref and opt-in -SyncIntuneDevice
    - Authoritative /users relationship ObjectId comparison and PrimaryUserConflict fail-closed protection
    - Stale Intune record detection (24h provisioning window) and duplicate Intune device detection
    - HTTP 429 Retry-After backoff, 429 -> 401 token refresh, and -WhatIf non-mutation guarantee
.EXAMPLE
    powershell.exe -NoProfile -File .\scripts\Test-AvdPersonalPrimaryUser.ps1
#>

#Requires -Version 5.1

[CmdletBinding()]
param()
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$source = Join-Path $PSScriptRoot 'Sync-AvdPersonalPrimaryUser.ps1'
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($source, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -gt 0) {
    throw ($parseErrors | Out-String)
}
foreach ($definition in $ast.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.FunctionDefinitionAst] }) {
    . ([scriptblock]::Create($definition.Extent.Text))
}
$script:assertions = 0

function Assert-AvdPrimaryUserTest {
    <#
    .SYNOPSIS
        Asserts a boolean test condition and increments the assertion counter.
    .DESCRIPTION
        Throws a terminating error with the test Name when Condition is $false;
        otherwise increments the script-scoped assertion counter.
    .PARAMETER Condition
        Boolean expression result to verify.
    .PARAMETER Name
        Human-readable description of the assertion being verified.
    .OUTPUTS
        None
    #>
    param([bool]$Condition, [string]$Name)
    if (-not $Condition) {
        throw "FAIL: $Name"
    }
    $script:assertions++
}

# 0. Verify every function in Sync-AvdPersonalPrimaryUser.ps1 has complete comment-based help
& {
    $fnDefs = @($ast.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.FunctionDefinitionAst] })
    Assert-AvdPrimaryUserTest ($fnDefs.Count -eq 13) 'All 13 helper functions are present in Sync-AvdPersonalPrimaryUser.ps1'
    foreach ($fn in $fnDefs) {
        $help = $fn.GetHelpContent()
        $hasSynopsisAndDesc = ($null -ne $help) -and (-not [string]::IsNullOrWhiteSpace($help.Synopsis)) -and (-not [string]::IsNullOrWhiteSpace($help.Description))
        Assert-AvdPrimaryUserTest $hasSynopsisAndDesc ("Function '{0}' has .SYNOPSIS and .DESCRIPTION help" -f $fn.Name)
        foreach ($p in $fn.Body.ParamBlock.Parameters) {
            $pName = $p.Name.VariablePath.UserPath.ToUpperInvariant()
            $hasParamDoc = ($null -ne $help) -and $help.Parameters.ContainsKey($pName) -and (-not [string]::IsNullOrWhiteSpace($help.Parameters[$pName]))
            Assert-AvdPrimaryUserTest $hasParamDoc ("Function '{0}' documents parameter -{1}" -f $fn.Name, $p.Name.VariablePath.UserPath)
        }
    }
}

# 1. Unit tests for helper functions, zero-GUID normalization, HTTP 429 retry, and @odata.nextLink pagination
& {
    $short1 = Get-AvdShortHostName -SessionHostResourceName 'vdpool-pers-01/avd-pers-001.contoso.local'
    Assert-AvdPrimaryUserTest ($short1 -ceq 'AVD-PERS-001') 'FQDN session host name resolves to uppercase short hostname'
    $short2 = Get-AvdShortHostName -SessionHostResourceName '' -VmResourceId '/subscriptions/sub1/resourceGroups/rg1/providers/Microsoft.Compute/virtualMachines/avd-pers-002'
    Assert-AvdPrimaryUserTest ($short2 -ceq 'AVD-PERS-002') 'VM ResourceId fallback resolves to uppercase short hostname'
    # Zero-GUID and path-traversal normalization
    Assert-AvdPrimaryUserTest ($null -eq (ConvertTo-NormalizedPrincipalId -Value '00000000-0000-0000-0000-000000000000')) 'Zero GUID normalizes to null'
    Assert-AvdPrimaryUserTest ($null -eq (ConvertTo-NormalizedPrincipalId -Value '../evil')) 'Path-traversal identifier normalizes to null'
    Assert-AvdPrimaryUserTest ('uid-123' -ceq (ConvertTo-NormalizedPrincipalId -Value ' uid-123 ')) 'Non-empty GUID trims and remains intact'
    # HTTP 429 -> 401 token refresh and @odata.nextLink pagination + untrusted nextLink rejection
    $script:httpCalls = 0
    $script:retrySleeps = New-Object 'System.Collections.Generic.List[int]'
    $AuthMode = 'ManagedIdentity'
    function Get-AvdPrimaryUserAccessToken { param($ResourceUri, $SelectedAuthMode, $ClientId) return 'refreshed-graph-token' }
    function Invoke-AvdPrimaryUserSleep { param([int]$Seconds) [void]$script:retrySleeps.Add($Seconds) }
    function Invoke-AvdPrimaryUserHttpRequest {
        param($Method, $Uri, $Headers, $Body)
        $script:httpCalls++
        if ($script:httpCalls -eq 1) {
            $ex = New-Object System.Exception('Too Many Requests')
            $ex.Data['HttpStatusCode'] = 429
            $ex.Data['RetryAfterSeconds'] = 150
            throw $ex
        }
        if ($script:httpCalls -eq 2) {
            $ex401 = New-Object System.Exception('Unauthorized')
            $ex401.Data['HttpStatusCode'] = 401
            throw $ex401
        }
        if ($Uri -eq 'https://graph.microsoft.com/v1.0/page1') {
            return [PSCustomObject]@{
                value             = @([PSCustomObject]@{ id = 'p1' })
                '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/page2'
            }
        }
        if ($Uri -eq 'https://graph.microsoft.com/v1.0/page2') {
            return [PSCustomObject]@{
                value = @([PSCustomObject]@{ id = 'p2' })
            }
        }
        if ($Uri -eq 'https://graph.microsoft.com/v1.0/bad-next') {
            return [PSCustomObject]@{
                value             = @([PSCustomObject]@{ id = 'p1' })
                '@odata.nextLink' = 'https://evil.example/steal-token'
            }
        }
        throw "Unexpected URI in pagination test: $Uri"
    }
    $hdrs = @{ Authorization = 'Bearer initial' }
    $pagedResult = @(Invoke-AvdPrimaryUserPagedGet -Uri 'https://graph.microsoft.com/v1.0/page1' -Headers $hdrs)
    Assert-AvdPrimaryUserTest ($pagedResult.Count -eq 2 -and $pagedResult[0].id -eq 'p1' -and $pagedResult[1].id -eq 'p2' -and $script:retrySleeps.Count -eq 1 -and $script:retrySleeps[0] -eq 150 -and $hdrs['Authorization'] -eq 'Bearer refreshed-graph-token') 'Invoke-AvdPrimaryUserRest honors uncapped Retry-After, refreshes token on 429 -> 401, and follows @odata.nextLink'
    $badNextThrown = $false
    try {
        $null = @(Invoke-AvdPrimaryUserPagedGet -Uri 'https://graph.microsoft.com/v1.0/bad-next' -Headers $hdrs)
    } catch {
        $badNextThrown = ($_.Exception.Message -like '*Refusing to follow untrusted pagination link*')
    }
    Assert-AvdPrimaryUserTest $badNextThrown 'Invoke-AvdPrimaryUserPagedGet rejects off-origin nextLink to prevent token exfiltration'
}

# 2. End-to-end orchestration tests using real AST ParamBlock and mocked REST layer
& {
    $subId = '11111111-1111-1111-1111-111111111111'
    $personalPoolId  = "/subscriptions/$subId/resourceGroups/rg-avd/providers/Microsoft.DesktopVirtualization/hostPools/vdpool-pers-01"
    $personalPool2Id = "/subscriptions/$subId/resourceGroups/rg-avd/providers/Microsoft.DesktopVirtualization/hostPools/vdpool-pers-02"
    $pooledPoolId    = "/subscriptions/$subId/resourceGroups/rg-avd/providers/Microsoft.DesktopVirtualization/hostPools/vdpool-pooled-01"
    $script:mockWrites = New-Object 'System.Collections.Generic.List[object]'
    $script:failSyncDevice = $false
    $entryText = @(
        "[CmdletBinding(SupportsShouldProcess = `$true, ConfirmImpact = 'None')]"
        $ast.ParamBlock.Extent.Text
        foreach ($stmt in $ast.EndBlock.Statements | Where-Object { $_ -isnot [System.Management.Automation.Language.FunctionDefinitionAst] }) {
            $stmt.Extent.Text
        }
    ) -join [Environment]::NewLine
    $runner = [scriptblock]::Create($entryText)
    function Get-AvdPrimaryUserAccessToken {
        param($ResourceUri, $SelectedAuthMode, $ClientId)
        return "mock-token-$ResourceUri"
    }
    function Invoke-AvdPrimaryUserRest {
        param($Method, $Uri, $Headers, $Body)
        if ($Method -ne 'Get') {
            if ($Method -eq 'Post' -and $Uri -like '*/syncDevice' -and $script:failSyncDevice) {
                $syncEx = New-Object System.InvalidOperationException('HTTP 403: Missing DeviceManagementManagedDevices.PrivilegedOperations.All')
                $syncEx.Data['HttpStatusCode'] = 403
                throw $syncEx
            }
            [void]$script:mockWrites.Add([PSCustomObject]@{
                Method = $Method
                Uri    = $Uri
                Body   = $Body
            })
            return [PSCustomObject]@{}
        }
        if ($Uri -eq "https://management.azure.com$personalPoolId`?api-version=2024-04-03") {
            return [PSCustomObject]@{ id = $personalPoolId; name = 'vdpool-pers-01'; properties = [PSCustomObject]@{ hostPoolType = 'Personal' } }
        }
        if ($Uri -eq "https://management.azure.com$personalPool2Id`?api-version=2024-04-03") {
            return [PSCustomObject]@{ id = $personalPool2Id; name = 'vdpool-pers-02'; properties = [PSCustomObject]@{ hostPoolType = 'Personal' } }
        }
        if ($Uri -eq "https://management.azure.com$pooledPoolId`?api-version=2024-04-03") {
            return [PSCustomObject]@{ id = $pooledPoolId; name = 'vdpool-pooled-01'; properties = [PSCustomObject]@{ hostPoolType = 'Pooled' } }
        }
        if ($Uri -eq "https://management.azure.com$personalPool2Id/sessionHosts?api-version=2024-04-03") {
            return [PSCustomObject]@{
                value = @(
                    [PSCustomObject]@{
                        name       = 'vdpool-pers-02/avd-sh-002'
                        properties = [PSCustomObject]@{ assignedUser = 'alice@contoso.com'; resourceId = "/subscriptions/$subId/resourceGroups/rg-avd-2/providers/Microsoft.Compute/virtualMachines/avd-sh-002" }
                    }
                )
            }
        }
        if ($Uri -eq "https://management.azure.com$personalPoolId/sessionHosts?api-version=2024-04-03") {
            return [PSCustomObject]@{
                value = @(
                    # Host 1: Assigned in AVD (enrolled 30 min before DSC sessionHost creation), Intune Primary User has zero GUID and empty /users relationship -> should update
                    [PSCustomObject]@{
                        name       = 'vdpool-pers-01/avd-sh-001.contoso.com'
                        systemData = [PSCustomObject]@{ createdAt = '2026-11-01T10:30:00Z' }
                        properties = [PSCustomObject]@{ assignedUser = 'alice@contoso.com'; resourceId = "/subscriptions/$subId/resourceGroups/rg-avd/providers/Microsoft.Compute/virtualMachines/avd-sh-001" }
                    }
                    # Host 2: Already in sync via /users relationship ObjectId -> should skip
                    [PSCustomObject]@{
                        name       = 'vdpool-pers-01/avd-sh-002'
                        systemData = [PSCustomObject]@{ createdAt = '2026-11-01T09:00:00Z' }
                        properties = [PSCustomObject]@{ assignedUser = 'bob@contoso.com'; resourceId = "/subscriptions/$subId/resourceGroups/rg-avd/providers/Microsoft.Compute/virtualMachines/avd-sh-002" }
                    }
                    # Host 3: Conflict (same UPN string on device scalar, but /users relationship has a different Entra User ObjectId) -> PrimaryUserConflict
                    [PSCustomObject]@{
                        name       = 'vdpool-pers-01/avd-sh-003'
                        systemData = [PSCustomObject]@{ createdAt = '2026-11-01T09:00:00Z' }
                        properties = [PSCustomObject]@{ assignedUser = 'dave@contoso.com'; resourceId = "/subscriptions/$subId/resourceGroups/rg-avd/providers/Microsoft.Compute/virtualMachines/avd-sh-003" }
                    }
                    # Host 4: Unassigned in AVD -> UnassignedInAvd
                    [PSCustomObject]@{
                        name       = 'vdpool-pers-01/avd-sh-004'
                        properties = [PSCustomObject]@{ assignedUser = ''; resourceId = "/subscriptions/$subId/resourceGroups/rg-avd/providers/Microsoft.Compute/virtualMachines/avd-sh-004" }
                    }
                    # Host 5: Duplicate records in Intune -> AmbiguousIntuneDevice
                    [PSCustomObject]@{
                        name       = 'vdpool-pers-01/avd-sh-005'
                        properties = [PSCustomObject]@{ assignedUser = 'alice@contoso.com'; resourceId = "/subscriptions/$subId/resourceGroups/rg-avd/providers/Microsoft.Compute/virtualMachines/avd-sh-005" }
                    }
                    # Host 6: Not enrolled in Intune yet -> PendingIntuneEnrollment
                    [PSCustomObject]@{
                        name       = 'vdpool-pers-01/avd-sh-006'
                        properties = [PSCustomObject]@{ assignedUser = 'alice@contoso.com'; resourceId = "/subscriptions/$subId/resourceGroups/rg-avd/providers/Microsoft.Compute/virtualMachines/avd-sh-006" }
                    }
                    # Host 7: B2B Guest user resolved via mail fallback, /users relationship matches uid-guest -> AlreadyInSync
                    [PSCustomObject]@{
                        name       = 'vdpool-pers-01/avd-sh-007'
                        properties = [PSCustomObject]@{ assignedUser = 'guest@partner.com'; resourceId = "/subscriptions/$subId/resourceGroups/rg-avd/providers/Microsoft.Compute/virtualMachines/avd-sh-007" }
                    }
                    # Host 8: Sole stale Intune record after VM rebuild (Intune enrolled 2026-10-01, AVD host created 2026-11-03) -> StaleIntuneDeviceRecord
                    [PSCustomObject]@{
                        name       = 'vdpool-pers-01/avd-sh-008'
                        systemData = [PSCustomObject]@{ createdAt = '2026-11-03T12:00:00Z' }
                        properties = [PSCustomObject]@{ assignedUser = 'alice@contoso.com'; resourceId = "/subscriptions/$subId/resourceGroups/rg-avd/providers/Microsoft.Compute/virtualMachines/avd-sh-008" }
                    }
                )
            }
        }
        if ($Uri -eq "https://management.azure.com$personalPoolId/sessionHosts/avd-sh-001?api-version=2024-04-03") {
            $nf = New-Object System.InvalidOperationException('HTTP 404: SessionHost avd-sh-001 not found')
            $nf.Data['HttpStatusCode'] = 404
            throw $nf
        }
        if ($Uri -like 'https://graph.microsoft.com/v1.0/deviceManagement/managedDevices*') {
            $devs = New-Object 'System.Collections.Generic.List[object]'
            [void]$devs.Add([PSCustomObject]@{ id = 'intune-001'; deviceName = 'avd-sh-001'; azureADDeviceId = 'aad-001'; userPrincipalName = ''; userId = '00000000-0000-0000-0000-000000000000'; enrolledDateTime = '2026-11-01T10:00:00Z' })
            [void]$devs.Add([PSCustomObject]@{ id = 'intune-002'; deviceName = 'AVD-SH-002'; azureADDeviceId = 'aad-002'; userPrincipalName = 'bob@contoso.com'; userId = 'uid-bob'; enrolledDateTime = '2026-11-01T10:00:00Z' })
            [void]$devs.Add([PSCustomObject]@{ id = 'intune-003'; deviceName = 'avd-sh-003'; azureADDeviceId = 'aad-003'; userPrincipalName = 'dave@contoso.com'; userId = 'uid-dave-old-deleted'; enrolledDateTime = '2026-11-01T10:00:00Z' })
            [void]$devs.Add([PSCustomObject]@{ id = 'intune-005-old'; deviceName = 'avd-sh-005'; azureADDeviceId = 'aad-005-old'; userPrincipalName = ''; userId = ''; enrolledDateTime = '2026-11-02T10:00:00Z' })
            [void]$devs.Add([PSCustomObject]@{ id = 'intune-005-new'; deviceName = 'avd-sh-005'; azureADDeviceId = 'aad-005-new'; userPrincipalName = ''; userId = ''; enrolledDateTime = '2026-11-02T10:00:00.5000000Z' })
            [void]$devs.Add([PSCustomObject]@{ id = 'intune-007'; deviceName = 'avd-sh-007'; azureADDeviceId = 'aad-007'; userPrincipalName = 'guest_partner.com#EXT#@contoso.onmicrosoft.com'; userId = 'uid-guest'; enrolledDateTime = '2026-11-01T10:00:00Z' })
            [void]$devs.Add([PSCustomObject]@{ id = 'intune-008-stale'; deviceName = 'avd-sh-008'; azureADDeviceId = 'aad-008-stale'; userPrincipalName = ''; userId = ''; enrolledDateTime = '2026-10-01T10:00:00Z' })
            return [PSCustomObject]@{ value = $devs.ToArray() }
        }
        if ($Uri -eq 'https://graph.microsoft.com/beta/deviceManagement/managedDevices/intune-001/users?$select=id,userPrincipalName') {
            return [PSCustomObject]@{ value = @() }
        }
        if ($Uri -eq 'https://graph.microsoft.com/beta/deviceManagement/managedDevices/intune-002/users?$select=id,userPrincipalName') {
            return [PSCustomObject]@{ value = @([PSCustomObject]@{ id = 'uid-bob'; userPrincipalName = 'bob@contoso.com' }) }
        }
        if ($Uri -eq 'https://graph.microsoft.com/beta/deviceManagement/managedDevices/intune-003/users?$select=id,userPrincipalName') {
            # Same UPN dave@contoso.com, but different ObjectId (recreated user or relationship mismatch)
            return [PSCustomObject]@{ value = @([PSCustomObject]@{ id = 'uid-dave-old-deleted'; userPrincipalName = 'dave@contoso.com' }) }
        }
        if ($Uri -eq 'https://graph.microsoft.com/beta/deviceManagement/managedDevices/intune-007/users?$select=id,userPrincipalName') {
            return [PSCustomObject]@{ value = @([PSCustomObject]@{ id = 'uid-guest'; userPrincipalName = 'guest_partner.com#EXT#@contoso.onmicrosoft.com' }) }
        }
        if ($Uri -like 'https://graph.microsoft.com/v1.0/users/alice%40contoso.com*') {
            return [PSCustomObject]@{ id = 'uid-alice'; userPrincipalName = 'alice@contoso.com' }
        }
        if ($Uri -like 'https://graph.microsoft.com/v1.0/users/bob%40contoso.com*') {
            return [PSCustomObject]@{ id = 'uid-bob'; userPrincipalName = 'bob@contoso.com' }
        }
        if ($Uri -like 'https://graph.microsoft.com/v1.0/users/dave%40contoso.com*') {
            return [PSCustomObject]@{ id = 'uid-dave-new'; userPrincipalName = 'dave@contoso.com' }
        }
        if ($Uri -like 'https://graph.microsoft.com/v1.0/users/guest%40partner.com*') {
            $nfGuest = New-Object System.InvalidOperationException('HTTP 404: User not found by direct UPN')
            $nfGuest.Data['HttpStatusCode'] = 404
            throw $nfGuest
        }
        if ($Uri -like 'https://graph.microsoft.com/v1.0/users?$filter=*guest%40partner.com*') {
            return [PSCustomObject]@{
                value = @([PSCustomObject]@{ id = 'uid-guest'; userPrincipalName = 'guest_partner.com#EXT#@contoso.onmicrosoft.com' })
            }
        }
        throw "Unexpected mock URI: $Method $Uri"
    }
    # Test Run 0: Missing target scope and hostile HostPoolResourceId throw explicitly before REST calls
    $scopeThrown = $false
    try {
        $null = @(& $runner -AuthMode ManagedIdentity)
    } catch {
        $scopeThrown = ($_.Exception.Message -like '*Specify target Personal Host Pool*')
    }
    Assert-AvdPrimaryUserTest $scopeThrown 'Throws explicitly when no target Host Pool is specified'
    $hostileThrown = $false
    try {
        $null = @(& $runner -AuthMode ManagedIdentity -HostPoolResourceId '@evil.example/pool')
    } catch {
        $hostileThrown = ($_.Exception.Message -like '*Invalid HostPoolResourceId*')
    }
    Assert-AvdPrimaryUserTest $hostileThrown 'Rejects hostile HostPoolResourceId before token acquisition or HTTP calls'
    # Test Run A: Fail-closed reconciliation on Personal Host Pool (and skips Pooled Host Pool)
    $script:mockWrites.Clear()
    $out = @(& $runner -AuthMode ManagedIdentity -HostPoolResourceId @($personalPoolId, $pooledPoolId))
    Assert-AvdPrimaryUserTest ($out.Count -eq 8) 'Evaluates all 8 session hosts in Personal pool and skips Pooled pool'
    $r1 = $out | Where-Object { $_.ShortHostName -eq 'AVD-SH-001' }
    Assert-AvdPrimaryUserTest ($r1.Status -eq 'Updated' -and $r1.SyncTriggered -eq $false -and $r1.IntuneDeviceId -eq 'intune-001') 'Host 1 sets Primary User without calling syncDevice by default (PrivilegedOperations.All not required)'
    $r2 = $out | Where-Object { $_.ShortHostName -eq 'AVD-SH-002' }
    Assert-AvdPrimaryUserTest ($r2.Status -eq 'AlreadyInSync' -and $r2.Action -eq 'None') 'Host 2 with matching /users relationship ObjectId is skipped'
    $r3 = $out | Where-Object { $_.ShortHostName -eq 'AVD-SH-003' }
    Assert-AvdPrimaryUserTest ($r3.Status -eq 'PrimaryUserConflict' -and $r3.Action -eq 'None') 'Host 3 with matching UPN but mismatched /users ObjectId fails closed as PrimaryUserConflict'
    $r4 = $out | Where-Object { $_.ShortHostName -eq 'AVD-SH-004' }
    Assert-AvdPrimaryUserTest ($r4.Status -eq 'UnassignedInAvd' -and $r4.Action -eq 'None') 'Host 4 unassigned in AVD is skipped immediately'
    $r5 = $out | Where-Object { $_.ShortHostName -eq 'AVD-SH-005' }
    Assert-AvdPrimaryUserTest ($r5.Status -eq 'AmbiguousIntuneDevice') 'Host 5 duplicate Intune records fail closed as AmbiguousIntuneDevice'
    $r6 = $out | Where-Object { $_.ShortHostName -eq 'AVD-SH-006' }
    Assert-AvdPrimaryUserTest ($r6.Status -eq 'PendingIntuneEnrollment') 'Host 6 not yet in Intune is reported as PendingIntuneEnrollment'
    $r7 = $out | Where-Object { $_.ShortHostName -eq 'AVD-SH-007' }
    Assert-AvdPrimaryUserTest ($r7.Status -eq 'AlreadyInSync' -and $r7.NewPrimaryUser -eq 'guest_partner.com#EXT#@contoso.onmicrosoft.com') 'Host 7 resolves B2B guest via mail fallback and matches /users ObjectId'
    $r8 = $out | Where-Object { $_.ShortHostName -eq 'AVD-SH-008' }
    Assert-AvdPrimaryUserTest ($r8.Status -eq 'StaleIntuneDeviceRecord' -and $r8.Action -eq 'None') 'Host 8 with sole stale Intune record enrolled before AVD sessionHost creation fails closed as StaleIntuneDeviceRecord'
    Assert-AvdPrimaryUserTest ($script:mockWrites.Count -eq 1 -and $script:mockWrites[0].Uri -eq 'https://graph.microsoft.com/beta/deviceManagement/managedDevices/intune-001/users/$ref') 'Default run performs only 1 users/$ref POST and 0 syncDevice calls'
    # Test Run B: Opt-in -SyncIntuneDevice success and failure reporting
    $script:mockWrites.Clear()
    $outSyncOk = @(& $runner -AuthMode ManagedIdentity -SubscriptionId $subId -ResourceGroupName 'rg-avd' -HostPoolName 'vdpool-pers-01' -TargetSessionHostName 'avd-sh-001' -SyncIntuneDevice)
    Assert-AvdPrimaryUserTest ($outSyncOk.Count -eq 1 -and $outSyncOk[0].Status -eq 'Updated' -and $outSyncOk[0].SyncTriggered -eq $true -and $script:mockWrites.Count -eq 2) 'Opt-in -SyncIntuneDevice triggers POST .../syncDevice after setting Primary User'
    $script:mockWrites.Clear()
    $script:failSyncDevice = $true
    $outShort = @(& $runner -AuthMode ManagedIdentity -SubscriptionId $subId -ResourceGroupName 'rg-avd' -HostPoolName 'vdpool-pers-01' -TargetSessionHostName 'avd-sh-001' -SyncIntuneDevice)
    $script:failSyncDevice = $false
    Assert-AvdPrimaryUserTest ($outShort.Count -eq 1 -and $outShort[0].ShortHostName -eq 'AVD-SH-001' -and $outShort[0].Status -eq 'PrimaryUserUpdatedSyncFailed' -and -not $outShort[0].SyncTriggered) 'ResourceGroupName + HostPoolName + short-name TargetSessionHostName with -SyncIntuneDevice reports PrimaryUserUpdatedSyncFailed when syncDevice fails'
    # Test Run C: Cross-pool duplicate short hostname fails closed as AmbiguousAvdSessionHost
    $script:mockWrites.Clear()
    $outCollision = @(& $runner -AuthMode ManagedIdentity -HostPoolResourceId @($personalPoolId, $personalPool2Id))
    $dupRows = @($outCollision | Where-Object { $_.ShortHostName -eq 'AVD-SH-002' -and $_.Status -eq 'AmbiguousAvdSessionHost' })
    Assert-AvdPrimaryUserTest ($dupRows.Count -eq 2) 'Cross-pool duplicate short hostname fails closed as AmbiguousAvdSessionHost on both pools'
    # Test Run D: WhatIf performs zero writes
    $script:mockWrites.Clear()
    $outD = @(& $runner -AuthMode ManagedIdentity -HostPoolResourceId $personalPoolId -WhatIf)
    $whatIfRows = @($outD | Where-Object { $_.Status -eq 'WhatIf' })
    Assert-AvdPrimaryUserTest ($script:mockWrites.Count -eq 0 -and $whatIfRows.Count -eq 1) 'WhatIf mode performs zero POST Graph requests'
}
Write-Host ("PASS: {0} assertions succeeded." -f $script:assertions) -ForegroundColor Green
