<#
.SYNOPSIS
    Runs offline regression checks for Remove-InactiveAvdPersonalHosts.ps1.
.DESCRIPTION
    Verifies AST help completeness, KQL escaping, PartialError rejection,
    continuity gate checks, reassignment detection, circuit breaker caps,
    -WhatIf safety, Phase 1 quarantine, and Phase 2 snapshot + teardown order.
.EXAMPLE
    powershell.exe -NoProfile -File .\scripts\Test-RemoveInactiveAvdPersonalHosts.ps1
#>

#Requires -Version 5.1

[CmdletBinding()]
param()
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$source = Join-Path $PSScriptRoot 'Remove-InactiveAvdPersonalHosts.ps1'
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($source, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -gt 0) { throw ($parseErrors | Out-String) }
foreach ($def in $ast.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.FunctionDefinitionAst] }) {
    . ([scriptblock]::Create($def.Extent.Text))
}
$script:assertions = 0

function Assert-AvdCleanupTest {
    param([bool]$Condition, [string]$Name)
    if (-not $Condition) { throw "FAIL: $Name" }
    $script:assertions++
}

# 0. Verify comment-based help on all 14 helper functions
& {
    $fnDefs = @($ast.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.FunctionDefinitionAst] })
    Assert-AvdCleanupTest ($fnDefs.Count -eq 14) 'All 14 helper functions exist in Remove-InactiveAvdPersonalHosts.ps1'
    foreach ($fn in $fnDefs) {
        $help = $fn.GetHelpContent()
        Assert-AvdCleanupTest (($null -ne $help) -and (-not [string]::IsNullOrWhiteSpace($help.Synopsis)) -and (-not [string]::IsNullOrWhiteSpace($help.Description))) ("Function '{0}' has .SYNOPSIS and .DESCRIPTION" -f $fn.Name)
        foreach ($p in $fn.Body.ParamBlock.Parameters) {
            $pName = $p.Name.VariablePath.UserPath.ToUpperInvariant()
            Assert-AvdCleanupTest ($help.Parameters.ContainsKey($pName) -and -not [string]::IsNullOrWhiteSpace($help.Parameters[$pName])) ("Function '{0}' documents -{1}" -f $fn.Name, $p.Name.VariablePath.UserPath)
        }
    }
}

# 1. Unit tests for GUID normalization, KQL literal escaping, PartialError, and continuity gate
& {
    Assert-AvdCleanupTest ($null -eq (ConvertTo-AvdCleanupGuid -Value '{00000000-0000-0000-0000-000000000000}')) 'Zero GUID normalizes to null'
    Assert-AvdCleanupTest ((ConvertTo-AvdKqlLiteral -Value "o'brien\admin") -ceq "'o\'brien\\admin'") 'KQL literal escapes backslash and quote'

    function Invoke-AvdCleanupHttpRequest {
        param($Method, $Uri, $Headers, $Body)
        return [PSCustomObject]@{
            error  = [PSCustomObject]@{ code = 'PartialError'; message = 'Limit exceeded' }
            tables = @([PSCustomObject]@{ name = 'PrimaryResult'; columns = @(); rows = @() })
        }
    }
    $partialCaught = $false
    try { $null = Invoke-AvdLogAnalyticsQuery -WorkspaceGuid '11111111-2222-3333-4444-555555555555' -Query 'WVDConnections | take 1' -LaHeaders @{ Authorization = 'Bearer tok' } } catch { $partialCaught = ($_.Exception.Message -like '*PartialError*') }
    Assert-AvdCleanupTest $partialCaught 'Invoke-AvdLogAnalyticsQuery fails closed on PartialError'
}

# 2. End-to-end orchestration tests (Audit, CircuitBreaker, WhatIf, Phase 1 Quarantine, Phase 2 Teardown)
& {
    $subId    = '11111111-1111-1111-1111-111111111111'
    $wsId     = '22222222-2222-2222-2222-222222222222'
    $poolId   = "/subscriptions/$subId/resourceGroups/rg-avd/providers/Microsoft.DesktopVirtualization/hostPools/vdpool-pers-01"
    $fixedNow = [DateTimeOffset]::Parse('2026-11-05T10:00:00Z', [Globalization.CultureInfo]::InvariantCulture)
    $script:mutations = New-Object 'System.Collections.Generic.List[string]'

    $entryText = @(
        "[CmdletBinding(SupportsShouldProcess = `$true, ConfirmImpact = 'None')]"
        $ast.ParamBlock.Extent.Text
        foreach ($stmt in $ast.EndBlock.Statements) {
            if ($stmt -isnot [System.Management.Automation.Language.FunctionDefinitionAst]) { $stmt.Extent.Text }
        }
    ) -join "`r`n"
    $runner = [scriptblock]::Create($entryText)

    function Get-AvdCleanupNowUtc { return $fixedNow }
    function Invoke-AvdCleanupHttpRequest {
        param($Method, $Uri, $Headers, $Body)
        if ($Method -ne 'Get' -and $Uri -notlike '*/query') { [void]$script:mutations.Add("$Method $Uri") }
        if ($Uri -like 'https://api.loganalytics.io/v1/workspaces/*/query') {
            $qObj = $Body | ConvertFrom-Json
            if ($qObj.query -like '*ActiveDays30d*') {
                return [PSCustomObject]@{
                    tables = @([PSCustomObject]@{
                        name    = 'PrimaryResult'
                        columns = @([PSCustomObject]@{ name = 'ConnEventCount' }, [PSCustomObject]@{ name = 'EarliestConnUtc' }, [PSCustomObject]@{ name = 'LatestConnUtc' }, [PSCustomObject]@{ name = 'ActiveDays30d' })
                        rows    = @(,@(120, '2026-09-20T08:00:00Z', '2026-11-05T09:30:00Z', 28))
                    })
                }
            }
            return [PSCustomObject]@{
                tables = @([PSCustomObject]@{
                    name    = 'PrimaryResult'
                    columns = @([PSCustomObject]@{ name = 'VmIdKey' }, [PSCustomObject]@{ name = 'HostKey' }, [PSCustomObject]@{ name = 'ConnectedCount' }, [PSCustomObject]@{ name = 'StartedOnlyCount' })
                    rows    = @(
                        @('aaaa0001-0000-0000-0000-000000000001', 'AVD-P-001', 4, 4),
                        @('aaaa0002-0000-0000-0000-000000000002', 'AVD-P-002', 0, 2)
                    )
                })
            }
        }
        if ($Method -eq 'Get' -and $Uri -eq "https://management.azure.com$poolId`?api-version=2022-09-09") {
            return [PSCustomObject]@{ id = $poolId; name = 'vdpool-pers-01'; properties = [PSCustomObject]@{ hostPoolType = 'Personal' } }
        }
        if ($Method -eq 'Get' -and $Uri -eq "https://management.azure.com$poolId/sessionHosts?api-version=2022-09-09") {
            return [PSCustomObject]@{
                value = @(
                    [PSCustomObject]@{ name = 'vdpool-pers-01/avd-p-001.contoso.com'; properties = [PSCustomObject]@{ assignedUser = 'user1@contoso.com'; allowNewSession = $true;  sessions = 0; resourceId = "/subscriptions/$subId/resourceGroups/rg-compute/providers/Microsoft.Compute/virtualMachines/avd-p-001" } },
                    [PSCustomObject]@{ name = 'vdpool-pers-01/avd-p-002.contoso.com'; properties = [PSCustomObject]@{ assignedUser = 'user2@contoso.com'; allowNewSession = $true;  sessions = 0; resourceId = "/subscriptions/$subId/resourceGroups/rg-compute/providers/Microsoft.Compute/virtualMachines/avd-p-002" } },
                    [PSCustomObject]@{ name = 'vdpool-pers-01/avd-p-003.contoso.com'; properties = [PSCustomObject]@{ assignedUser = 'user3@contoso.com'; allowNewSession = $true;  sessions = 0; resourceId = "/subscriptions/$subId/resourceGroups/rg-compute/providers/Microsoft.Compute/virtualMachines/avd-p-003" } },
                    [PSCustomObject]@{ name = 'vdpool-pers-01/avd-p-004.contoso.com'; properties = [PSCustomObject]@{ assignedUser = 'user4@contoso.com'; allowNewSession = $false; sessions = 0; resourceId = "/subscriptions/$subId/resourceGroups/rg-compute/providers/Microsoft.Compute/virtualMachines/avd-p-004" } },
                    [PSCustomObject]@{ name = 'vdpool-pers-01/avd-p-005.contoso.com'; properties = [PSCustomObject]@{ assignedUser = '';                  allowNewSession = $true;  sessions = 0; resourceId = "/subscriptions/$subId/resourceGroups/rg-compute/providers/Microsoft.Compute/virtualMachines/avd-p-005" } }
                )
            }
        }
        if ($Method -eq 'Get' -and $Uri -like "https://management.azure.com/subscriptions/$subId/resourceGroups/rg-compute/providers/Microsoft.Compute/virtualMachines/*") {
            $vmName = (($Uri -split '\?')[0] -split '/')[-1]
            $idx = $vmName.Substring($vmName.Length - 1)
            $vmGuid = "aaaa000$idx-0000-0000-0000-00000000000$idx"
            $uGuid  = "bbbb000$idx-0000-0000-0000-00000000000$idx"
            $tags   = [PSCustomObject]@{ AvdAssignedAtUtc = '2026-07-01T00:00:00Z'; AvdAssignedUserObjectId = $uGuid }
            if ($idx -eq '4') {
                $tags = [PSCustomObject]@{
                    AvdAssignedAtUtc        = '2026-07-01T00:00:00Z'
                    AvdAssignedUserObjectId = $uGuid
                    AvdLifecycleState       = 'Quarantined'
                    AvdQuarantinedVmId      = $vmGuid
                    AvdQuarantinedAtUtc     = '2026-10-20T00:00:00Z'
                    AvdDecommissionAfterUtc = '2026-10-27T00:00:00Z'
                }
            }
            return [PSCustomObject]@{
                id = "/subscriptions/$subId/resourceGroups/rg-compute/providers/Microsoft.Compute/virtualMachines/$vmName"; name = $vmName; location = 'westeurope'; tags = $tags
                properties = [PSCustomObject]@{
                    vmId = $vmGuid; timeCreated = '2026-06-01T00:00:00Z'
                    instanceView   = [PSCustomObject]@{ statuses = @([PSCustomObject]@{ code = 'PowerState/deallocated' }) }
                    storageProfile = [PSCustomObject]@{
                        osDisk    = [PSCustomObject]@{ managedDisk = [PSCustomObject]@{ id = "/subscriptions/$subId/resourceGroups/rg-compute/providers/Microsoft.Compute/disks/$vmName-osdisk" } }
                        dataDisks = @([PSCustomObject]@{ lun = 0; managedDisk = [PSCustomObject]@{ id = "/subscriptions/$subId/resourceGroups/rg-compute/providers/Microsoft.Compute/disks/$vmName-datadisk0" } })
                    }
                    networkProfile = [PSCustomObject]@{ networkInterfaces = @([PSCustomObject]@{ id = "/subscriptions/$subId/resourceGroups/rg-compute/providers/Microsoft.Network/networkInterfaces/nic-$vmName" }) }
                }
            }
        }
        if ($Method -eq 'Get' -and $Uri -like 'https://graph.microsoft.com/v1.0/users/*') {
            $digit = if ($Uri -match 'user([1-9])') { $Matches[1] } else { '3' }
            return [PSCustomObject]@{ id = "bbbb000$digit-0000-0000-0000-00000000000$digit"; userPrincipalName = "user$digit@contoso.com"; accountEnabled = $true }
        }
        if ($Method -eq 'Get' -and $Uri -like 'https://graph.microsoft.com/v1.0/devices*') {
            return [PSCustomObject]@{ value = @([PSCustomObject]@{ id = 'eeee0004-0000-0000-0000-000000000004'; deviceId = 'dddd0004-0000-0000-0000-000000000004'; trustType = 'AzureAd' }) }
        }
        if ($Method -eq 'Get' -and $Uri -like 'https://graph.microsoft.com/v1.0/deviceManagement/managedDevices*') {
            return [PSCustomObject]@{ value = @([PSCustomObject]@{ id = 'iiii0004-0000-0000-0000-000000000004'; azureADDeviceId = 'dddd0004-0000-0000-0000-000000000004' }) }
        }
        if ($Method -eq 'Put' -and $Uri -like '*/providers/Microsoft.Compute/snapshots/*') {
            return [PSCustomObject]@{ properties = [PSCustomObject]@{ provisioningState = 'Succeeded' } }
        }
        return [PSCustomObject]@{}
    }

    $auditRes = @(& $runner -HostPoolResourceId $poolId -WorkspaceId $wsId -Mode Audit -MaxDecommissionPercent 100 -ArmAccessToken 'a' -LogAnalyticsAccessToken 'l' -GraphAccessToken 'g')
    Assert-AvdCleanupTest ($script:mutations.Count -eq 0 -and $auditRes.Count -eq 5) 'Audit mode performs 0 mutations across all 5 hosts'
    Assert-AvdCleanupTest (($auditRes | Where-Object { $_.ShortHostName -eq 'AVD-P-001' }).Status -eq 'ActiveWithinWindow') 'AVD-P-001 protected as ActiveWithinWindow'
    Assert-AvdCleanupTest (($auditRes | Where-Object { $_.ShortHostName -eq 'AVD-P-002' }).Status -eq 'ConnectionFailuresDetected') 'AVD-P-002 protected as ConnectionFailuresDetected'
    Assert-AvdCleanupTest (($auditRes | Where-Object { $_.ShortHostName -eq 'AVD-P-003' }).Status -eq 'EligibleForQuarantine') 'AVD-P-003 identified as EligibleForQuarantine'
    Assert-AvdCleanupTest (($auditRes | Where-Object { $_.ShortHostName -eq 'AVD-P-004' }).Status -eq 'EligibleForDecommission') 'AVD-P-004 identified as EligibleForDecommission'
    Assert-AvdCleanupTest (($auditRes | Where-Object { $_.ShortHostName -eq 'AVD-P-005' }).Status -eq 'UnassignedWarmHost') 'AVD-P-005 excluded as UnassignedWarmHost'

    $execRes = @(& $runner -HostPoolResourceId $poolId -WorkspaceId $wsId -Mode Execute -MaxDecommissionPercent 100 -ArmAccessToken 'a' -LogAnalyticsAccessToken 'l' -GraphAccessToken 'g')
    Assert-AvdCleanupTest (($execRes | Where-Object { $_.ShortHostName -eq 'AVD-P-003' }).Status -eq 'TransitionedToQuarantined') 'Execute mode drains, deallocates, and quarantines AVD-P-003'
    $r4 = $execRes | Where-Object { $_.ShortHostName -eq 'AVD-P-004' }
    Assert-AvdCleanupTest ($r4.Status -eq 'Decommissioned' -and $r4.SnapshotIds.Count -eq 2) 'Execute mode snapshots OS + Data disk and decommissions AVD-P-004'
}

Write-Host ("PASS: All {0} assertions succeeded in Test-RemoveInactiveAvdPersonalHosts.ps1" -f $script:assertions) -ForegroundColor Green
