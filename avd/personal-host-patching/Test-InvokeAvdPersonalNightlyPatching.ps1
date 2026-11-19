<#
.SYNOPSIS
    Runs offline regression checks for Invoke-AvdPersonalNightlyPatching.ps1.
.DESCRIPTION
    Verifies PowerShell 5.1 AST validity, comment-based help completeness on every
    helper function, InstanceView HibernationState/Hibernated parsing, AvdPatchLock
    and AvdPatchState tag serialization, 50-tag limit enforcement, ownership-conditional
    case-insensitive tag deletion, fail-closed Scaling Plan exclusionTag resolution,
    5-day SHA-256 cohort sharding + smart catch-up skipping, live execution of the
    in-guest WUA Hotpatch/Baseline filter and Unicode WTSEnumerateSessionsW probe,
    TimedOut RunCommand TiWorker protection, self-healing lock recovery, and E2E
    sliding-window orchestration.
.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\avd\personal-host-patching\Test-InvokeAvdPersonalNightlyPatching.ps1
#>

#Requires -Version 5.1

[CmdletBinding()]
param()
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$source = Join-Path $PSScriptRoot 'Invoke-AvdPersonalNightlyPatching.ps1'
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($source, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -gt 0) { throw ($parseErrors | Out-String) }
foreach ($def in $ast.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.FunctionDefinitionAst] }) {
    . ([scriptblock]::Create($def.Extent.Text))
}
$script:assertions = 0

function Assert-AvdPatchTest {
    param([bool]$Condition, [string]$Name)
    if (-not $Condition) { throw "FAIL: $Name" }
    $script:assertions++
}

# 0. Verify comment-based help on every top-level helper function + script-level .NOTES/.EXAMPLE
& {
    $scriptHelp = $ast.GetHelpContent()
    Assert-AvdPatchTest ($null -ne $scriptHelp -and -not [string]::IsNullOrWhiteSpace($scriptHelp.Notes) -and @($scriptHelp.Examples).Count -ge 2) 'Script header includes .NOTES and multiple .EXAMPLE entries'

    $fnDefs = @($ast.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.FunctionDefinitionAst] })
    Assert-AvdPatchTest ($fnDefs.Count -eq 22) ("Expected 22 top-level helper functions (found $($fnDefs.Count))")
    foreach ($fn in $fnDefs) {
        $help = $fn.GetHelpContent()
        Assert-AvdPatchTest (($null -ne $help) -and (-not [string]::IsNullOrWhiteSpace($help.Synopsis)) -and (-not [string]::IsNullOrWhiteSpace($help.Description))) ("Function '{0}' has .SYNOPSIS and .DESCRIPTION" -f $fn.Name)
        if ($null -ne $fn.Body.ParamBlock -and $null -ne $fn.Body.ParamBlock.Parameters) {
            foreach ($p in $fn.Body.ParamBlock.Parameters) {
                $pName = $p.Name.VariablePath.UserPath.ToUpperInvariant()
                Assert-AvdPatchTest ($help.Parameters.ContainsKey($pName) -and -not [string]::IsNullOrWhiteSpace($help.Parameters[$pName])) ("Function '{0}' documents -{1}" -f $fn.Name, $p.Name.VariablePath.UserPath)
            }
        }
    }
}

# 1. Verify InstanceView HibernationState/Hibernated vs Deallocated vs StoppedAllocated vs InTransition
& {
    $hibVm = [PSCustomObject]@{
        properties = [PSCustomObject]@{
            additionalCapabilities = [PSCustomObject]@{ hibernationEnabled = $true }
            instanceView           = [PSCustomObject]@{
                statuses   = @(
                    [PSCustomObject]@{ code = 'ProvisioningState/succeeded' },
                    [PSCustomObject]@{ code = 'PowerState/deallocated' },
                    [PSCustomObject]@{ code = 'HibernationState/Hibernated'; displayStatus = 'VM hibernated' }
                )
                vmAgent    = [PSCustomObject]@{ statuses = @([PSCustomObject]@{ displayStatus = 'Ready' }) }
                extensions = @(
                    [PSCustomObject]@{
                        name        = 'AzureHibernateExtension'
                        type        = 'Microsoft.CPlat.Core.WindowsHibernateExtension'
                        substatuses = @([PSCustomObject]@{ code = 'ComponentStatus/VMBootState/Resume/succeeded' })
                    }
                )
            }
        }
    }
    $sumHib = Get-AvdPatchVmPowerSummary -VmResource $hibVm
    Assert-AvdPatchTest ($sumHib.EffectiveState -ceq 'Hibernated' -and $sumHib.IsHibernated -and $sumHib.HibernationEnabled -and $sumHib.VmAgentReady) 'Classifies PowerState/deallocated + HibernationState/Hibernated as Hibernated'
    Assert-AvdPatchTest ($sumHib.BootSubstatus -ceq 'ComponentStatus/VMBootState/Resume/succeeded') 'Extracts VMBootState/Resume/succeeded substatus'

    $deallocVm = [PSCustomObject]@{
        properties = [PSCustomObject]@{
            additionalCapabilities = [PSCustomObject]@{ hibernationEnabled = $true }
            instanceView           = [PSCustomObject]@{ statuses = @([PSCustomObject]@{ code = 'PowerState/deallocated' }) }
        }
    }
    $sumDealloc = Get-AvdPatchVmPowerSummary -VmResource $deallocVm
    Assert-AvdPatchTest ($sumDealloc.EffectiveState -ceq 'Deallocated' -and -not $sumDealloc.IsHibernated) 'Classifies PowerState/deallocated without HibernationState as Deallocated'

    $stoppedVm = [PSCustomObject]@{ properties = [PSCustomObject]@{ instanceView = [PSCustomObject]@{ statuses = @([PSCustomObject]@{ code = 'PowerState/stopped' }) } } }
    Assert-AvdPatchTest ((Get-AvdPatchVmPowerSummary -VmResource $stoppedVm).EffectiveState -ceq 'StoppedAllocated') 'Classifies PowerState/stopped as StoppedAllocated'

    $transVm = [PSCustomObject]@{ properties = [PSCustomObject]@{ instanceView = [PSCustomObject]@{ statuses = @([PSCustomObject]@{ code = 'PowerState/deallocating' }) } } }
    Assert-AvdPatchTest ((Get-AvdPatchVmPowerSummary -VmResource $transVm).EffectiveState -ceq 'InTransition') 'Classifies PowerState/deallocating as InTransition'
}

# 2. Verify compact tag codecs, 50-tag ceiling, and ownership-conditional case-insensitive Delete
& {
    $expUtc = [DateTimeOffset]::Parse('2026-11-18T04:00:00Z', [Globalization.CultureInfo]::InvariantCulture)
    $lockStr = ConvertTo-AvdPatchLockValue -RunId '11111111-2222-3333-4444-555555555555' -VmId 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee' -PriorState 'Hibernated' -PriorAllowNewSession $true -HadPriorExclusionTag $false -ExclusionTag 'exclude-sp-pers-01' -ExpiresUtc $expUtc
    Assert-AvdPatchTest ($lockStr.Length -le 240) ("AvdPatchLock length ($($lockStr.Length)) <= 240 chars")
    $parsedLock = ConvertFrom-AvdPatchLockValue -Value $lockStr
    Assert-AvdPatchTest ($parsedLock.PriorState -ceq 'Hibernated' -and $parsedLock.PriorAllowNewSession -and -not $parsedLock.HadPriorExclusionTag -and $parsedLock.ExclusionTag -ceq 'exclude-sp-pers-01' -and $parsedLock.ExpiresUtc -eq $expUtc) 'Round-trips AvdPatchLock fields including ExclusionTag'

    $stateStr = ConvertTo-AvdPatchStateValue -LastOkUtc $expUtc -Ubr 3323 -Outcome 'InstalledNoReboot' -DeferSinceEpoch 1763400000
    Assert-AvdPatchTest ($stateStr.Length -le 240) 'AvdPatchState length <= 240 chars'
    $parsedState = ConvertFrom-AvdPatchStateValue -Value $stateStr
    Assert-AvdPatchTest ($parsedState.Ubr -eq 3323 -and $parsedState.Outcome -ceq 'InstalledNoReboot' -and $parsedState.DeferSinceEpoch -eq 1763400000 -and $parsedState.LastOkUtc -eq $expUtc) 'Round-trips AvdPatchState fields including LastOkUtc'

    $fiftyTags = @{}
    for ($i = 1; $i -le 50; $i++) { $fiftyTags["Tag$i"] = "Val$i" }
    $caught50 = $false
    try {
        $null = Update-AvdPatchVmTags -VmResourceId '/subscriptions/11111111-1111-1111-1111-111111111111/resourceGroups/rg/providers/Microsoft.Compute/virtualMachines/vm1' -Operation 'Merge' -TagsToApply @{ NewTag = 'x' } -ExistingTags $fiftyTags -TagsApiVersion '2021-04-01' -TokenState @{ StaticToken = 'tok' }
    } catch {
        $caught50 = ($_.Exception.Message -like '*50-tag limit*')
    }
    Assert-AvdPatchTest $caught50 'Fails closed before exceeding ARM 50-tag limit on Merge'

    $script:capturedPatchBody = $null
    function Invoke-AvdPatchRestRequest {
        param($Method, $Uri, $TokenState, $Body, $MaxRetries, [switch]$AllowNotFound)
        $null = $Method; $null = $Uri; $null = $TokenState; $null = $MaxRetries; $null = $AllowNotFound
        $script:capturedPatchBody = $Body
        return [PSCustomObject]@{ StatusCode = 200; Headers = @{}; Body = $null; RawContent = '' }
    }

    # Case-insensitive key match + exact owned value match -> deletes using VM's exact key casing
    $null = Update-AvdPatchVmTags -VmResourceId '/subscriptions/11111111-1111-1111-1111-111111111111/resourceGroups/rg/providers/Microsoft.Compute/virtualMachines/vm1' -Operation 'Delete' -TagsToApply @{ 'exclude-sp-pers-01' = 'AvdNightlyPatch' } -ExistingTags @{ 'Exclude-Sp-Pers-01' = 'AvdNightlyPatch'; KeepMe = '1' } -TagsApiVersion '2021-04-01' -TokenState @{ StaticToken = 'tok' }
    Assert-AvdPatchTest ($script:capturedPatchBody.operation -ceq 'Delete' -and $script:capturedPatchBody.properties.tags['Exclude-Sp-Pers-01'] -ceq 'AvdNightlyPatch' -and -not $script:capturedPatchBody.properties.tags.ContainsKey('KeepMe')) 'Ownership-conditional Delete matches tag key case-insensitively when owned value matches'

    # Value changed mid-run by another operator -> must NOT delete!
    $script:capturedPatchBody = $null
    $skipRes = Update-AvdPatchVmTags -VmResourceId '/subscriptions/11111111-1111-1111-1111-111111111111/resourceGroups/rg/providers/Microsoft.Compute/virtualMachines/vm1' -Operation 'Delete' -TagsToApply @{ 'exclude-sp-pers-01' = 'AvdNightlyPatch' } -ExistingTags @{ 'exclude-sp-pers-01' = 'IncidentOperatorPin' } -TagsApiVersion '2021-04-01' -TokenState @{ StaticToken = 'tok' }
    Assert-AvdPatchTest ([bool]$skipRes.Skipped -and $null -eq $script:capturedPatchBody) 'Ownership-conditional Delete refuses to delete exclusionTag modified mid-run by an operator'
}

# 3. Verify REST origin guard, LRO 403 RG-only RBAC fallback, Scaling Plan fail-closed gate, and Cohort math
& {
    # Off-origin bearer token protection
    $offOriginBlocked = $false
    try {
        $null = Invoke-AvdPatchRestRequest -Method Get -Uri 'https://attacker.example.com/steal' -TokenState @{ StaticToken = 'tok' }
    } catch {
        $offOriginBlocked = ($_.Exception.Message -like '*untrusted URI*')
    }
    Assert-AvdPatchTest $offOriginBlocked 'Invoke-AvdPatchRestRequest blocks non-ARM URIs before sending bearer token'

    # LRO 403 graceful fallback for Resource-Group-only RBAC
    function Invoke-AvdPatchRestRequest {
        param($Method, $Uri, $TokenState, $Body, $MaxRetries, [switch]$AllowNotFound)
        $null = $Method; $null = $Uri; $null = $TokenState; $null = $Body; $null = $MaxRetries; $null = $AllowNotFound
        $ex = New-Object System.InvalidOperationException('Forbidden')
        $ex.Data['HttpStatusCode'] = 403
        throw $ex
    }
    $lro403 = Get-AvdPatchAsyncOperationStatus -OperationUri 'https://management.azure.com/subscriptions/11111111-1111-1111-1111-111111111111/providers/Microsoft.Compute/locations/westeurope/operations/op1' -TokenState @{ StaticToken = 'tok' }
    Assert-AvdPatchTest ($lro403.Status -ceq 'Unavailable') 'Get-AvdPatchAsyncOperationStatus degrades to Unavailable on HTTP 403 for RG-only RBAC'

    $subId = '11111111-1111-1111-1111-111111111111'
    $poolId = "/subscriptions/$subId/resourceGroups/rg-avd/providers/Microsoft.DesktopVirtualization/hostPools/vdpool-pers-01"
    $script:mockPlans = @()
    function Invoke-AvdPatchArmPagedGet {
        param($Uri, $TokenState)
        $null = $Uri; $null = $TokenState
        return $script:mockPlans
    }

    # Zero visible Scaling Plans -> MUST fail closed unless -ExclusionTagName or -AllowNoScalingPlan!
    $script:mockPlans = @()
    $r0 = Resolve-AvdPatchPoolExclusionTag -HostPoolArmId $poolId -ApiVersion '2024-04-03' -TokenState @{ StaticToken = 'tok' }
    Assert-AvdPatchTest (-not $r0.IsValid -and $r0.FailureReason -like '*No enabled Scaling Plan found*') 'Fails closed when 0 enabled Scaling Plans are visible and -AllowNoScalingPlan is false'

    $r0Allow = Resolve-AvdPatchPoolExclusionTag -HostPoolArmId $poolId -AllowNoScalingPlan $true -ApiVersion '2024-04-03' -TokenState @{ StaticToken = 'tok' }
    Assert-AvdPatchTest ($r0Allow.IsValid -and $r0Allow.ExclusionTag -ceq 'excludeFromScaling') 'Succeeds with -AllowNoScalingPlan when pool intentionally has no Scaling Plan'

    # Empty exclusionTag on enabled plan -> fails closed
    $script:mockPlans = @(
        [PSCustomObject]@{
            name       = 'sp-missing-tag'
            properties = [PSCustomObject]@{
                exclusionTag       = ''
                hostPoolReferences = @([PSCustomObject]@{ hostPoolArmPath = $poolId; scalingPlanEnabled = $true })
            }
        }
    )
    $r1 = Resolve-AvdPatchPoolExclusionTag -HostPoolArmId $poolId -ApiVersion '2024-04-03' -TokenState @{ StaticToken = 'tok' }
    Assert-AvdPatchTest (-not $r1.IsValid -and $r1.FailureReason -like '*no exclusionTag*') 'Fails closed when active Scaling Plan has empty exclusionTag'

    # Mismatched -ExclusionTagName -> fails closed
    $script:mockPlans = @(
        [PSCustomObject]@{
            name       = 'sp-pers-01'
            properties = [PSCustomObject]@{
                exclusionTag       = 'exclude-sp-pers-01'
                hostPoolReferences = @([PSCustomObject]@{ hostPoolArmPath = $poolId; scalingPlanEnabled = $true })
            }
        }
    )
    $r2 = Resolve-AvdPatchPoolExclusionTag -HostPoolArmId $poolId -ExpectedExclusionTag 'WrongTag' -ApiVersion '2024-04-03' -TokenState @{ StaticToken = 'tok' }
    Assert-AvdPatchTest (-not $r2.IsValid -and $r2.FailureReason -like '*does not match*') 'Fails closed when -ExclusionTagName mismatches Scaling Plan exclusionTag'

    $r3 = Resolve-AvdPatchPoolExclusionTag -HostPoolArmId $poolId -ExpectedExclusionTag 'exclude-sp-pers-01' -ApiVersion '2024-04-03' -TokenState @{ StaticToken = 'tok' }
    Assert-AvdPatchTest ($r3.IsValid -and $r3.ExclusionTag -ceq 'exclude-sp-pers-01') 'Succeeds when Scaling Plan exclusionTag matches Bicep naming convention'

    # 5-Cohort Mon-Fri auto-selection (-1): every host runs on exactly 1 weekday Mon..Fri!
    $mon = [DateTimeOffset]::Parse('2026-11-16T01:00:00Z', [Globalization.CultureInfo]::InvariantCulture)
    for ($h = 1; $h -le 20; $h++) {
        $name = ('AVD-PERS-{0:D3}' -f $h)
        $weekdayHits = 0
        for ($d = 0; $d -lt 5; $d++) {
            if (Test-AvdPatchHostInCohort -ShortHostName $name -CohortCount 5 -CohortIndex -1 -NowUtc $mon.AddDays($d)) {
                $weekdayHits++
            }
        }
        Assert-AvdPatchTest ($weekdayHits -eq 1) ("Host $name maps to exactly 1 weekday across Mon-Fri under CohortCount=5")
    }

    # 7-Cohort auto-selection (-1): visits every cohort across 7 consecutive UTC days
    for ($h = 1; $h -le 5; $h++) {
        $name = ('AVD-PERS-7C-{0:D3}' -f $h)
        $hits7 = 0
        for ($d = 0; $d -lt 7; $d++) {
            if (Test-AvdPatchHostInCohort -ShortHostName $name -CohortCount 7 -CohortIndex -1 -NowUtc $mon.AddDays($d)) { $hits7++ }
        }
        Assert-AvdPatchTest ($hits7 -eq 1) ("Host $name maps to exactly 1 day across 7-day rotation under CohortCount=7")
    }
}

# 4. Real Execution of Generated Guest Scripts (WUA Filter, Hotpatch RebootBehavior=2, Upgrades Block, Reboot Guard, Sentinel Parser)
& {
    $runGuid = '99999999-8888-7777-6666-555555555555'
    foreach ($st in @('Service', 'ReadOnlyProbe', 'ScheduleReboot', 'VerifyPostReboot')) {
        $gs = New-AvdPatchGuestScript -Stage $st -RunId $runGuid -PreWakeHadSession $true -AllowRebootingUpdatesOnSession $false -BaselineBootUtc '2026-11-17T00:00:00Z'
        $gTok = $null; $gErr = $null
        $null = [System.Management.Automation.Language.Parser]::ParseInput($gs, [ref]$gTok, [ref]$gErr)
        Assert-AvdPatchTest ($gErr.Count -eq 0) ("Generated guest script Stage=$st parses with 0 PowerShell 5.1 AST errors")
    }

    $rbtSrc = New-AvdPatchGuestScript -Stage 'ScheduleReboot' -RunId $runGuid
    Assert-AvdPatchTest ($rbtSrc -like '*WTSEnumerateSessionsW*' -and $rbtSrc -like '*DeleteExpiredTaskAfter*' -and $rbtSrc -notlike '*-StartWhenAvailable*' -and $rbtSrc -like '*-EncodedCommand*' -and $rbtSrc -like '*Registry::HKEY_USERS*') 'ScheduleReboot uses Unicode WTSEnumerateSessionsW, HKEY_USERS hive detection, DeleteExpiredTaskAfter without StartWhenAvailable, and locale-independent EncodedCommand session guard'

    # Execute Select-AvdGuestUpdates directly from the parsed guest AST against realistic WUA mock updates!
    $svcSrc = New-AvdPatchGuestScript -Stage 'Service' -RunId $runGuid -PreWakeHadSession $true -AllowRebootingUpdatesOnSession $false
    $svcTok = $null; $svcErr = $null
    $svcAst = [System.Management.Automation.Language.Parser]::ParseInput($svcSrc, [ref]$svcTok, [ref]$svcErr)
    foreach ($fnAst in $svcAst.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $_.Name -in @('Get-WuaCollItem', 'Select-AvdGuestUpdates') }) {
        . ([scriptblock]::Create($fnAst.Extent.Text))
    }

    function Get-MockWuaUpdate {
        param([string]$Kb, [string]$Title, [string]$CategoryGuid, [int]$RebootBehavior, [bool]$Interactive = $false)
        return [PSCustomObject]@{
            Title                = $Title
            KBArticleIDs         = @($Kb)
            Identity             = [PSCustomObject]@{ UpdateID = "$Kb-0000-0000" }
            Categories           = @([PSCustomObject]@{ CategoryID = "{$CategoryGuid}" })
            InstallationBehavior = [PSCustomObject]@{ RebootBehavior = $RebootBehavior; CanRequestUserInput = $Interactive }
        }
    }

    $mockCatalog = @(
        (Get-MockWuaUpdate -Kb '5099999' -Title 'Windows 11, version 26H2' -CategoryGuid '3689bdc8-b205-4af4-8d4a-a63924c5e9d5' -RebootBehavior 1),
        (Get-MockWuaUpdate -Kb '5088888' -Title 'Windows Feature Experience Pack' -CategoryGuid 'b54e7d24-7add-428f-8b75-90a396fa584f' -RebootBehavior 0),
        (Get-MockWuaUpdate -Kb '5077777' -Title '2026-11 Cumulative Update Preview for Windows 11' -CategoryGuid 'cd5ffd1e-e932-4e3a-bf74-18bf0b1bbd83' -RebootBehavior 2),
        (Get-MockWuaUpdate -Kb '2267602' -Title 'Security Intelligence Update for Microsoft Defender Antivirus' -CategoryGuid 'e0789628-ce08-4437-be74-2495b842f43b' -RebootBehavior 0),
        (Get-MockWuaUpdate -Kb '5132043' -Title '2026-11 Hotpatch Security Update for Windows 11 Version 24H2 (KB5132043)' -CategoryGuid '0fa1201d-4330-4fa8-8ae9-b877473b6441' -RebootBehavior 2),
        (Get-MockWuaUpdate -Kb '5101650' -Title '2026-10 Cumulative Update for Windows 11 Version 24H2 (KB5101650)' -CategoryGuid '0fa1201d-4330-4fa8-8ae9-b877473b6441' -RebootBehavior 2)
    )

    # Disconnected session within grace on Hotpatch-enabled host -> selects Defender (0) + Hotpatch (2 + Hotpatch title), defers Baseline LCU, blocks 26H2/FeaturePack/Preview!
    $selWithinGrace = Select-AvdGuestUpdates -UpdatesCollection $mockCatalog -SessionPresent $true -AllowRebootOnSess $false -VbsOk $true -HotpatchPolOk $true -DriversOk $false -PreviewsOk $false
    Assert-AvdPatchTest (($selWithinGrace.SelectedLabels -join ',') -ceq 'KB2267602,KB5132043') 'Select-AvdGuestUpdates installs Defender (RebootBehavior=0) and Monthly B Hotpatch (RebootBehavior=2 + VBS + Policy) while blocking 26H2, FeaturePacks, and Previews'
    Assert-AvdPatchTest (($selWithinGrace.DeferredLabels -join ',') -ceq 'KB5101650') 'Select-AvdGuestUpdates defers standard Baseline LCU when Disconnected session is within grace'

    # Non-VBS host with session -> defers RebootBehavior=2 update even if titled Hotpatch
    $selNoVbs = Select-AvdGuestUpdates -UpdatesCollection $mockCatalog -SessionPresent $true -AllowRebootOnSess $false -VbsOk $false -HotpatchPolOk $true -DriversOk $false -PreviewsOk $false
    Assert-AvdPatchTest (($selNoVbs.SelectedLabels -join ',') -ceq 'KB2267602' -and ($selNoVbs.DeferredLabels -join ',') -ceq 'KB5132043,KB5101650') 'Select-AvdGuestUpdates defers RebootBehavior=2 Hotpatch candidate if VBS is not running'

    # Grace expired -> selects Baseline LCU too
    $selGraceExp = Select-AvdGuestUpdates -UpdatesCollection $mockCatalog -SessionPresent $true -AllowRebootOnSess $true -VbsOk $true -HotpatchPolOk $true -DriversOk $false -PreviewsOk $false
    Assert-AvdPatchTest (($selGraceExp.SelectedLabels -join ',') -ceq 'KB2267602,KB5132043,KB5101650' -and $selGraceExp.DeferredLabels.Count -eq 0) 'Select-AvdGuestUpdates includes Baseline LCU once DisconnectedSessionGraceDays expires'

    # 4KB last-bytes stdout sentinel extraction + replay rejection
    $noisyPrefix = ('WARNING: Third-party WMI hook output ' * 120) + "`r`n"
    $samplePayload = 'AVDPATCH_JSON:{"runId":"' + $runGuid + '","ubr":3624,"vbsRunning":true,"hotpatchPolicy":true,"installedKbs":["KB5132043"],"deferredKbs":[],"rbtReq":false,"postRbt":false,"mdmSynced":true,"outcome":"InstalledNoReboot"}:END_AVDPATCH'
    $fullOut = $noisyPrefix + $samplePayload
    $tail4k = if ($fullOut.Length -gt 4096) { $fullOut.Substring($fullOut.Length - 4096) } else { $fullOut }
    $parsedOut = ConvertFrom-AvdPatchRunCommandOutput -RawOutput $tail4k -ExpectedRunId $runGuid
    Assert-AvdPatchTest ($null -ne $parsedOut -and $parsedOut.outcome -ceq 'InstalledNoReboot' -and -not $parsedOut.rbtReq) 'Extracts InstalledNoReboot JSON cleanly from Azure last-4096-byte stdout tail'
    Assert-AvdPatchTest ($null -eq (ConvertFrom-AvdPatchRunCommandOutput -RawOutput $tail4k -ExpectedRunId '00000000-0000-0000-0000-000000000000')) 'Rejects RunCommand output when runId mismatches'
    Assert-AvdPatchTest ($null -eq (ConvertFrom-AvdPatchRunCommandOutput -RawOutput 'AVDPATCH_JSON:{"runId":"truncated' -ExpectedRunId $runGuid)) 'Rejects truncated JSON missing :END_AVDPATCH sentinel'
}

# 5. End-to-End Orchestration Suite (Audit, WhatIf, Triage, MorningUnlock, Self-Healing Lock Recovery, TimedOut Guard, LegalHold Ticket ID, Friday Catch-Up Skip, Circuit Breaker, Cutoff)
& {
    $subId = '11111111-1111-1111-1111-111111111111'
    $poolId = "/subscriptions/$subId/resourceGroups/rg-avd/providers/Microsoft.DesktopVirtualization/hostPools/vdpool-pers-01"
    $script:simClock = [DateTimeOffset]::Parse('2026-11-17T01:00:00Z', [Globalization.CultureInfo]::InvariantCulture)
    $script:mutations = New-Object 'System.Collections.Generic.List[string]'

    $entryText = @(
        "[CmdletBinding(SupportsShouldProcess = `$true, ConfirmImpact = 'None')]"
        $ast.ParamBlock.Extent.Text
        foreach ($stmt in $ast.EndBlock.Statements) {
            if ($stmt -isnot [System.Management.Automation.Language.FunctionDefinitionAst]) { $stmt.Extent.Text }
        }
    ) -join "`r`n"
    $runner = [scriptblock]::Create($entryText)

    function Get-AvdPatchNowUtc { return $script:simClock }
    function Start-AvdPatchPollSleep {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
        param([int]$Seconds)
        $script:simClock = $script:simClock.AddSeconds([Math]::Max(15, $Seconds))
    }

    # Mock fleet of 9 Personal VMs (contoso.com):
    # - sh-001: Hibernated + hidden guest session -> Hotpatch installed, Baseline deferred, Hibernated restored
    # - sh-002: Hibernated + 0 sessions + pre-existing mixed-case 'Exclude-Sp-Pers-01'='OpsPin' -> Baseline LCU + reboot + settle + Hibernated + preserves OpsPin
    # - sh-003: Running + Active Broker session -> SkippedActiveSession
    # - sh-004: Quarantined by Remove-InactiveAvdPersonalHosts.ps1 -> SkippedQuarantined
    # - sh-005: Expired idle AvdPatchLock (PriorState=Hibernated, PriorAllowNewSession=true) + stale finished RunCommand -> Triage reports StaleLockIdleRecoverable (0 writes); MorningUnlock clears lock & exclusionTag; Execute self-heals, patches, and restores Hibernated!
    # - sh-006: Hibernated + Disconnected session, deallocate?hibernate=true fails -> NEVER calls deallocate without hibernate=true; leaves Running & unlocks!
    # - sh-007: Expired busy AvdPatchLock with still-Running RunCommand -> Triage=StaleLockBusyRunCommand; Execute=NeedsOperatorTriage (restores allowNewSession=true, leaves Running)
    # - sh-008: Ticket-ID Legal Hold (AvdLegalHold='CASE-2026-42') -> SkippedLegalHold (aligned with cleanup script!)
    # - sh-009: Expired idle AvdPatchLock where admin had drained host BEFORE run (PriorAllowNewSession=false) -> MorningUnlock clears stale lock tags WITHOUT forcing allowNewSession=true!
    function Reset-MockFleet {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
        param()
        $script:simClock = [DateTimeOffset]::Parse('2026-11-17T01:00:00Z', [Globalization.CultureInfo]::InvariantCulture)
        $expLock5 = ConvertTo-AvdPatchLockValue -RunId 'deadbeef-0000-0000-0000-000000000005' -VmId 'aaaa0005-0000-0000-0000-000000000005' -PriorState 'Hibernated' -PriorAllowNewSession $true -HadPriorExclusionTag $false -ExclusionTag 'exclude-sp-pers-01' -ExpiresUtc ($script:simClock.AddHours(-2))
        $expLock7 = ConvertTo-AvdPatchLockValue -RunId 'deadbeef-0000-0000-0000-000000000007' -VmId 'aaaa0007-0000-0000-0000-000000000007' -PriorState 'Hibernated' -PriorAllowNewSession $true -HadPriorExclusionTag $false -ExclusionTag 'exclude-sp-pers-01' -ExpiresUtc ($script:simClock.AddHours(-2))
        $expLock9 = ConvertTo-AvdPatchLockValue -RunId 'deadbeef-0000-0000-0000-000000000009' -VmId 'aaaa0009-0000-0000-0000-000000000009' -PriorState 'Deallocated' -PriorAllowNewSession $false -HadPriorExclusionTag $false -ExclusionTag 'exclude-sp-pers-01' -ExpiresUtc ($script:simClock.AddHours(-2))
        $script:vmStates = @{
            'sh-001' = @{ Power = 'deallocated'; Hib = $true;  Agent = $true; AllowNew = $true;  Tags = @{ Owner = 'dev1' }; ActiveSess = 0; DiscSess = 0; FailHibCount = 0; TimedOutCmd = $false; PreRbtOnly = $false; RunCmds = @{} }
            'sh-002' = @{ Power = 'deallocated'; Hib = $true;  Agent = $true; AllowNew = $true;  Tags = @{ 'Exclude-Sp-Pers-01' = 'OpsPin' }; ActiveSess = 0; DiscSess = 0; FailHibCount = 0; TimedOutCmd = $false; PreRbtOnly = $false; RunCmds = @{} }
            'sh-003' = @{ Power = 'running';     Hib = $false; Agent = $true; AllowNew = $true;  Tags = @{}; ActiveSess = 1; DiscSess = 0; FailHibCount = 0; TimedOutCmd = $false; PreRbtOnly = $false; RunCmds = @{} }
            'sh-004' = @{ Power = 'deallocated'; Hib = $false; Agent = $true; AllowNew = $false; Tags = @{ AvdLifecycleState = ' Quarantined ' }; ActiveSess = 0; DiscSess = 0; FailHibCount = 0; TimedOutCmd = $false; PreRbtOnly = $false; RunCmds = @{} }
            'sh-005' = @{ Power = 'running';     Hib = $false; Agent = $true; AllowNew = $false; Tags = @{ 'exclude-sp-pers-01' = 'AvdNightlyPatch'; AvdPatchLock = $expLock5 }; ActiveSess = 0; DiscSess = 0; FailHibCount = 0; TimedOutCmd = $false; PreRbtOnly = $false; RunCmds = @{ 'AvdPatch-oldstale' = @{ State = 'Succeeded'; Output = '' } } }
            'sh-006' = @{ Power = 'deallocated'; Hib = $true;  Agent = $true; AllowNew = $true;  Tags = @{}; ActiveSess = 0; DiscSess = 1; FailHibCount = 99; TimedOutCmd = $false; PreRbtOnly = $false; RunCmds = @{} }
            'sh-007' = @{ Power = 'running';     Hib = $false; Agent = $true; AllowNew = $false; Tags = @{ 'exclude-sp-pers-01' = 'AvdNightlyPatch'; AvdPatchLock = $expLock7 }; ActiveSess = 0; DiscSess = 0; FailHibCount = 0; TimedOutCmd = $false; PreRbtOnly = $false; RunCmds = @{ 'AvdPatch-stillbusy' = @{ State = 'Running'; Output = '' } } }
            'sh-008' = @{ Power = 'deallocated'; Hib = $true;  Agent = $true; AllowNew = $true;  Tags = @{ AvdLegalHold = 'CASE-2026-42' }; ActiveSess = 0; DiscSess = 0; FailHibCount = 0; TimedOutCmd = $false; PreRbtOnly = $false; RunCmds = @{} }
            'sh-009' = @{ Power = 'deallocated'; Hib = $false; Agent = $true; AllowNew = $false; Tags = @{ 'exclude-sp-pers-01' = 'AvdNightlyPatch'; AvdPatchLock = $expLock9 }; ActiveSess = 0; DiscSess = 0; FailHibCount = 0; TimedOutCmd = $false; PreRbtOnly = $false; RunCmds = @{} }
        }
        $script:mutations.Clear()
    }

    function Invoke-AvdPatchRestRequest {
        param($Method, $Uri, $TokenState, $Body, $MaxRetries, [switch]$AllowNotFound)
        $null = $TokenState; $null = $MaxRetries; $null = $AllowNotFound
        if ($Method -ne 'Get') { [void]$script:mutations.Add("$Method $Uri") }

        if ($Method -eq 'Get' -and $Uri -like "https://management.azure.com$poolId`?api-version=*") {
            return [PSCustomObject]@{ StatusCode = 200; Headers = @{}; Body = [PSCustomObject]@{ id = $poolId; name = 'vdpool-pers-01'; properties = [PSCustomObject]@{ hostPoolType = 'Personal' } } }
        }
        if ($Method -eq 'Get' -and $Uri -like "https://management.azure.com/subscriptions/$subId/providers/Microsoft.DesktopVirtualization/scalingPlans*") {
            return [PSCustomObject]@{
                StatusCode = 200; Headers = @{}
                Body       = [PSCustomObject]@{
                    value = @([PSCustomObject]@{
                        name       = 'sp-pers-01'
                        properties = [PSCustomObject]@{
                            exclusionTag       = 'exclude-sp-pers-01'
                            hostPoolReferences = @([PSCustomObject]@{ hostPoolArmPath = $poolId; scalingPlanEnabled = $true })
                        }
                    })
                }
            }
        }
        if ($Method -eq 'Get' -and $Uri -like "https://management.azure.com$poolId/sessionHosts?api-version=*") {
            $list = foreach ($k in ($script:vmStates.Keys | Sort-Object)) {
                $st = $script:vmStates[$k]
                [PSCustomObject]@{
                    id         = "$poolId/sessionHosts/$k.contoso.com"
                    name       = "vdpool-pers-01/$k.contoso.com"
                    properties = [PSCustomObject]@{
                        assignedUser    = "$k@contoso.com"
                        allowNewSession = $st.AllowNew
                        sessions        = ($st.ActiveSess + $st.DiscSess)
                        resourceId      = "/subscriptions/$subId/resourceGroups/rg-vms/providers/Microsoft.Compute/virtualMachines/$k"
                    }
                }
            }
            return [PSCustomObject]@{ StatusCode = 200; Headers = @{}; Body = [PSCustomObject]@{ value = @($list) } }
        }
        if ($Method -eq 'Patch' -and $Uri -like "https://management.azure.com$poolId/sessionHosts/*") {
            $shLeaf = ((($Uri -split '\?')[0] -split '/')[-1] -split '\.')[0]
            $bObj = if ($Body -is [string]) { $Body | ConvertFrom-Json } else { $Body }
            $script:vmStates[$shLeaf].AllowNew = [bool]$bObj.properties.allowNewSession
            return [PSCustomObject]@{ StatusCode = 200; Headers = @{}; Body = $null }
        }
        if ($Method -eq 'Get' -and $Uri -like "https://management.azure.com$poolId/sessionHosts/*/userSessions*") {
            $parts = ($Uri -split '\?')[0] -split '/'
            $shLeaf = ($parts[$parts.Length - 2] -split '\.')[0]
            $st = $script:vmStates[$shLeaf]
            $uList = New-Object 'System.Collections.Generic.List[object]'
            for ($a = 0; $a -lt $st.ActiveSess; $a++) { [void]$uList.Add([PSCustomObject]@{ properties = [PSCustomObject]@{ sessionState = 'Active' } }) }
            for ($d = 0; $d -lt $st.DiscSess; $d++) { [void]$uList.Add([PSCustomObject]@{ properties = [PSCustomObject]@{ sessionState = 'Disconnected' } }) }
            return [PSCustomObject]@{ StatusCode = 200; Headers = @{}; Body = [PSCustomObject]@{ value = $uList.ToArray() } }
        }
        if ($Method -eq 'Patch' -and $Uri -like '*/providers/Microsoft.Resources/tags/default*') {
            $vmName = (($Uri -split '/providers/Microsoft\.Resources/tags')[0] -split '/')[-1]
            $st = $script:vmStates[$vmName]
            $op = [string]$Body.operation
            $tMap = $Body.properties.tags
            if ($op -eq 'Merge') {
                foreach ($tk in $tMap.Keys) {
                    $existingMatch = $null
                    foreach ($ck in @($st.Tags.Keys)) { if ([string]::Equals([string]$ck, [string]$tk, [System.StringComparison]::OrdinalIgnoreCase)) { $existingMatch = $ck; break } }
                    $targetKey = if ($null -ne $existingMatch) { $existingMatch } else { [string]$tk }
                    $st.Tags[$targetKey] = [string]$tMap[$tk]
                }
            } elseif ($op -eq 'Delete') {
                foreach ($tk in $tMap.Keys) {
                    foreach ($ck in @($st.Tags.Keys)) {
                        if ([string]::Equals([string]$ck, [string]$tk, [System.StringComparison]::OrdinalIgnoreCase) -and $st.Tags[$ck] -ceq [string]$tMap[$tk]) {
                            $st.Tags.Remove($ck)
                        }
                    }
                }
            }
            return [PSCustomObject]@{ StatusCode = 200; Headers = @{}; Body = $null }
        }
        if ($Method -eq 'Post' -and $Uri -like '*/virtualMachines/*/start*') {
            $vmName = (($Uri -split '/start')[0] -split '/')[-1]
            $script:vmStates[$vmName].Power = 'running'
            $script:vmStates[$vmName].Hib = $false
            return [PSCustomObject]@{ StatusCode = 202; Headers = @{}; Body = $null }
        }
        if ($Method -eq 'Post' -and $Uri -like '*/virtualMachines/*/deallocate*') {
            $vmName = (($Uri -split '/deallocate')[0] -split '/')[-1]
            $wantHib = ($Uri -like '*hibernate=true*')
            if ($wantHib -and $script:vmStates[$vmName].FailHibCount -gt 0) {
                $script:vmStates[$vmName].FailHibCount--
                throw 'ARM Post failed (HTTP 409): VMHibernateFailed: Hibernating the VM failed due to an internal error.'
            }
            $script:vmStates[$vmName].Power = 'deallocated'
            $script:vmStates[$vmName].Hib = $wantHib
            return [PSCustomObject]@{ StatusCode = 202; Headers = @{}; Body = $null }
        }
        # RunCommands list (GET .../runCommands?$expand=instanceView)
        if ($Method -eq 'Get' -and (($Uri -split '\?')[0] -like '*/virtualMachines/*/runCommands')) {
            $vmName = (($Uri -split '/runCommands')[0] -split '/')[-1]
            $st = $script:vmStates[$vmName]
            $rcList = foreach ($rk in @($st.RunCmds.Keys)) {
                $rcEntry = $st.RunCmds[$rk]
                [PSCustomObject]@{
                    name       = $rk
                    properties = [PSCustomObject]@{
                        provisioningState = 'Succeeded'
                        instanceView      = [PSCustomObject]@{ executionState = $rcEntry.State; exitCode = 0; output = $rcEntry.Output; error = '' }
                    }
                }
            }
            return [PSCustomObject]@{ StatusCode = 200; Headers = @{}; Body = [PSCustomObject]@{ value = @($rcList) } }
        }
        # Individual RunCommand PUT / GET / DELETE
        if ($Uri -like '*/virtualMachines/*/runCommands/*') {
            $cleanPath = ($Uri -split '\?')[0]
            $segs = $cleanPath -split '/'
            $cmdName = $segs[-1]
            $vmName = $segs[-3]
            $st = $script:vmStates[$vmName]
            if ($Method -eq 'Put') {
                $scr = [string]$Body.properties.source.script
                $ridMatch = [regex]::Match($scr, "runId\s*=\s*'([^']+)'")
                $extractedRun = $ridMatch.Groups[1].Value
                if ($cmdName -like 'AvdPatch-*') {
                    if ($st.TimedOutCmd) {
                        $st.RunCmds[$cmdName] = @{ State = 'TimedOut'; Output = '' }
                    } elseif ($st.PreRbtOnly) {
                        $jsonOut = 'AVDPATCH_JSON:{"runId":"' + $extractedRun + '","bootUtc":"2026-11-01T08:00:00Z","ubr":3323,"vbsRunning":true,"hotpatchPolicy":true,"pausedQuality":false,"freeGB":60,"wtsUsers":0,"wtsActive":0,"hives":0,"sessionPresent":false,"preRbt":true,"postRbt":true,"defOk":true,"defErr":"","installedKbs":[],"deferredKbs":[],"rbtReq":true,"mdmSynced":false,"outcome":"PreExistingRebootPending","err":""}:END_AVDPATCH'
                        $st.RunCmds[$cmdName] = @{ State = 'Succeeded'; Output = $jsonOut }
                    } elseif ($vmName -in @('sh-001', 'sh-006')) {
                        $jsonOut = 'AVDPATCH_JSON:{"runId":"' + $extractedRun + '","bootUtc":"2026-11-10T08:00:00Z","ubr":3624,"vbsRunning":true,"hotpatchPolicy":true,"pausedQuality":false,"freeGB":48,"wtsUsers":1,"wtsActive":0,"hives":1,"sessionPresent":true,"preRbt":false,"postRbt":false,"defOk":true,"defErr":"","installedKbs":["KB5132043"],"deferredKbs":["KB5101650"],"rbtReq":false,"mdmSynced":true,"outcome":"HotpatchOrDefInstalled_BaselineDeferred","err":""}:END_AVDPATCH'
                        $st.RunCmds[$cmdName] = @{ State = 'Succeeded'; Output = $jsonOut }
                    } else {
                        $jsonOut = 'AVDPATCH_JSON:{"runId":"' + $extractedRun + '","bootUtc":"2026-11-01T08:00:00Z","ubr":3323,"vbsRunning":true,"hotpatchPolicy":true,"pausedQuality":false,"freeGB":60,"wtsUsers":0,"wtsActive":0,"hives":0,"sessionPresent":false,"preRbt":false,"postRbt":true,"defOk":true,"defErr":"","installedKbs":["KB5101650"],"deferredKbs":[],"rbtReq":true,"mdmSynced":true,"outcome":"InstalledRebootRequired","err":""}:END_AVDPATCH'
                        $st.RunCmds[$cmdName] = @{ State = 'Succeeded'; Output = $jsonOut }
                    }
                } elseif ($cmdName -like 'AvdReboot-*') {
                    $st.RunCmds[$cmdName] = @{ State = 'Succeeded'; Output = ('AVDPATCH_JSON:{"runId":"' + $extractedRun + '","scheduled":true,"outcome":"RebootScheduled","err":""}:END_AVDPATCH') }
                } elseif ($cmdName -like 'AvdVerify-*') {
                    if ($cmdName -like '*-1') {
                        $st.RunCmds[$cmdName] = @{ State = 'Succeeded'; Output = ('AVDPATCH_JSON:{"runId":"' + $extractedRun + '","bootUtc":"2026-11-01T08:00:00Z","bootAdvanced":false,"ubr":3323,"postRbt":true,"mdmSynced":false,"outcome":"BootNotYetAdvanced"}:END_AVDPATCH') }
                    } else {
                        $st.RunCmds[$cmdName] = @{ State = 'Succeeded'; Output = ('AVDPATCH_JSON:{"runId":"' + $extractedRun + '","bootUtc":"2026-11-17T01:15:00Z","bootAdvanced":true,"ubr":3910,"postRbt":false,"mdmSynced":true,"outcome":"Verified"}:END_AVDPATCH') }
                    }
                }
                return [PSCustomObject]@{ StatusCode = 201; Headers = @{}; Body = $null }
            }
            if ($Method -eq 'Get') {
                if (-not $st.RunCmds.ContainsKey($cmdName)) {
                    return [PSCustomObject]@{ StatusCode = 404; Headers = @{}; Body = $null }
                }
                $rcObj = $st.RunCmds[$cmdName]
                return [PSCustomObject]@{
                    StatusCode = 200; Headers = @{}
                    Body       = [PSCustomObject]@{
                        properties = [PSCustomObject]@{
                            provisioningState = 'Succeeded'
                            instanceView      = [PSCustomObject]@{ executionState = $rcObj.State; exitCode = 0; output = $rcObj.Output; error = '' }
                        }
                    }
                }
            }
            if ($Method -eq 'Delete') {
                $null = $st.RunCmds.Remove($cmdName)
                return [PSCustomObject]@{ StatusCode = 200; Headers = @{}; Body = $null }
            }
        }
        if ($Method -eq 'Get' -and $Uri -like "https://management.azure.com/subscriptions/$subId/resourceGroups/rg-vms/providers/Microsoft.Compute/virtualMachines/*") {
            $vmName = (($Uri -split '\?')[0] -split '/')[-1]
            $st = $script:vmStates[$vmName]
            $idxNum = $vmName.Substring($vmName.Length - 1)
            $statList = @([PSCustomObject]@{ code = 'ProvisioningState/succeeded' }, [PSCustomObject]@{ code = "PowerState/$($st.Power)" })
            if ($st.Hib) { $statList += [PSCustomObject]@{ code = 'HibernationState/Hibernated' } }
            $tagCopy = @{}
            foreach ($tk in $st.Tags.Keys) { $tagCopy[$tk] = $st.Tags[$tk] }
            return [PSCustomObject]@{
                StatusCode = 200; Headers = @{}
                Body       = [PSCustomObject]@{
                    id         = "/subscriptions/$subId/resourceGroups/rg-vms/providers/Microsoft.Compute/virtualMachines/$vmName"
                    name       = $vmName
                    location   = 'germanywestcentral'
                    tags       = $tagCopy
                    properties = [PSCustomObject]@{
                        vmId                   = "aaaa000$idxNum-0000-0000-0000-00000000000$idxNum"
                        additionalCapabilities = [PSCustomObject]@{ hibernationEnabled = $true }
                        instanceView           = [PSCustomObject]@{
                            statuses   = $statList
                            vmAgent    = [PSCustomObject]@{ statuses = @([PSCustomObject]@{ displayStatus = 'Ready' }) }
                            extensions = @([PSCustomObject]@{ name = 'AzureHibernateExtension'; substatuses = @([PSCustomObject]@{ code = 'ComponentStatus/VMBootState/Resume/succeeded' }) })
                        }
                    }
                }
            }
        }
        throw "Unhandled mock URI: $Method $Uri"
    }

    # 5a. Audit Mode: strictly 0 mutations, ticket-ID AvdLegalHold & trimmed Quarantined honored
    Reset-MockFleet
    $auditRes = @(& $runner -HostPoolResourceId $poolId -Mode Audit -ArmAccessToken 'static-test-tok')
    Assert-AvdPatchTest ($script:mutations.Count -eq 0) 'Audit mode performs zero ARM mutations'
    Assert-AvdPatchTest (($auditRes | Where-Object { $_.ShortHostName -eq 'sh-003' }).Status -ceq 'SkippedActiveSession') 'Audit mode skips host with Active user session'
    Assert-AvdPatchTest (($auditRes | Where-Object { $_.ShortHostName -eq 'sh-004' }).Status -ceq 'SkippedQuarantined') 'Audit mode skips whitespace-padded Quarantined host'
    Assert-AvdPatchTest (($auditRes | Where-Object { $_.ShortHostName -eq 'sh-008' }).Status -ceq 'SkippedLegalHold') 'Audit mode honors ticket-ID AvdLegalHold=CASE-2026-42 (matching Remove-InactiveAvdPersonalHosts.ps1)'

    # 5b. Triage Mode: strictly 100% read-only (0 mutations), distinguishes idle vs busy stale locks
    Reset-MockFleet
    $triageRes = @(& $runner -HostPoolResourceId $poolId -Mode Triage -ArmAccessToken 'static-test-tok')
    Assert-AvdPatchTest ($script:mutations.Count -eq 0) 'Triage mode performs zero ARM mutations'
    Assert-AvdPatchTest (($triageRes | Where-Object { $_.ShortHostName -eq 'sh-005' }).Status -ceq 'StaleLockIdleRecoverable') 'Triage mode identifies expired idle lock as StaleLockIdleRecoverable'
    Assert-AvdPatchTest (($triageRes | Where-Object { $_.ShortHostName -eq 'sh-007' }).Status -ceq 'StaleLockBusyRunCommand') 'Triage mode identifies expired lock with Running RunCommand as StaleLockBusyRunCommand'

    # 5c. Execute -WhatIf: strictly 0 mutations
    Reset-MockFleet
    $whatIfRes = @(& $runner -HostPoolResourceId $poolId -Mode Execute -ArmAccessToken 'static-test-tok' -WhatIf)
    Assert-AvdPatchTest ($script:mutations.Count -eq 0 -and ($whatIfRes | Where-Object { $_.Status -eq 'WhatIfPreview' }).Count -ge 4) '-WhatIf performs zero ARM mutations'

    # 5d. MorningUnlock: clears expired idle lock tags + orphaned RunCommand, honors PriorAllowNewSession=false, protects busy RunCommand & Quarantined
    Reset-MockFleet
    $unlockRes = @(& $runner -HostPoolResourceId $poolId -Mode MorningUnlock -ArmAccessToken 'static-test-tok' -Confirm:$false)
    Assert-AvdPatchTest ($script:vmStates['sh-005'].AllowNew -eq $true -and -not $script:vmStates['sh-005'].Tags.ContainsKey('AvdPatchLock') -and -not $script:vmStates['sh-005'].Tags.ContainsKey('exclude-sp-pers-01') -and $script:vmStates['sh-005'].RunCmds.Count -eq 0 -and ($unlockRes | Where-Object { $_.ShortHostName -eq 'sh-005' }).Status -ceq 'MorningUnlocked') 'MorningUnlock restores allowNewSession, prunes stale RunCommand, and clears expired AvdPatchLock + exclusionTag on sh-005'
    Assert-AvdPatchTest ($script:vmStates['sh-009'].AllowNew -eq $false -and -not $script:vmStates['sh-009'].Tags.ContainsKey('AvdPatchLock')) 'MorningUnlock clears expired lock tags on sh-009 WITHOUT overriding pre-existing admin drain (PriorAllowNewSession=false)'
    Assert-AvdPatchTest ($script:vmStates['sh-007'].AllowNew -eq $true -and $script:vmStates['sh-007'].Tags.ContainsKey('AvdPatchLock') -and ($unlockRes | Where-Object { $_.ShortHostName -eq 'sh-007' }).Status -ceq 'MorningBusyRunCommand') 'MorningUnlock restores allowNewSession on sh-007 but keeps lock/exclusionTag while RunCommand is Running'
    Assert-AvdPatchTest ($script:vmStates['sh-004'].AllowNew -eq $false) 'MorningUnlock never unlocks Quarantined session host'

    # 5e. Full Execute Run: Hotpatch RAM preserve, mixed-case OpsPin preserve, self-healing sh-005 recovery, busy sh-007 triage, never-fallback hibernate=false!
    Reset-MockFleet
    $execRes = @(& $runner -HostPoolResourceId $poolId -Mode Execute -PostRebootSettleSeconds 60 -ArmAccessToken 'static-test-tok' -Confirm:$false)

    $r1 = $execRes | Where-Object { $_.ShortHostName -eq 'sh-001' }
    Assert-AvdPatchTest ($r1.Status -ceq 'HotpatchOrDefInstalled_BaselineDeferred' -and -not $r1.RebootPerformed -and $r1.InstalledKbs -ceq 'KB5132043' -and $r1.DeferredKbs -ceq 'KB5101650' -and $r1.FinalPowerState -ceq 'Hibernated' -and $script:vmStates['sh-001'].Hib) 'sh-001 installs Hotpatch KB without reboot, defers Baseline LCU, and restores Hibernated'
    Assert-AvdPatchTest (-not $script:vmStates['sh-001'].Tags.ContainsKey('exclude-sp-pers-01') -and -not $script:vmStates['sh-001'].Tags.ContainsKey('AvdPatchLock') -and $script:vmStates['sh-001'].Tags.ContainsKey('AvdPatchState')) 'sh-001 cleans up exclusionTag and AvdPatchLock and writes persistent AvdPatchState'

    $r2 = $execRes | Where-Object { $_.ShortHostName -eq 'sh-002' }
    Assert-AvdPatchTest ($r2.Status -ceq 'InstalledAndRebooted' -and $r2.RebootPerformed -and $r2.FinalPowerState -ceq 'Hibernated' -and $script:vmStates['sh-002'].Hib) 'sh-002 installs Baseline LCU, performs verified reboot, settles, and restores Hibernated'
    Assert-AvdPatchTest ($script:vmStates['sh-002'].Tags['Exclude-Sp-Pers-01'] -ceq 'OpsPin') 'sh-002 preserves pre-existing mixed-case Exclude-Sp-Pers-01=OpsPin tag'

    $r5 = $execRes | Where-Object { $_.ShortHostName -eq 'sh-005' }
    Assert-AvdPatchTest ($r5.Status -ceq 'InstalledAndRebooted' -and $r5.FinalPowerState -ceq 'Hibernated' -and $script:vmStates['sh-005'].Hib -and $script:vmStates['sh-005'].AllowNew -eq $true -and -not $script:vmStates['sh-005'].Tags.ContainsKey('AvdPatchLock') -and $script:vmStates['sh-005'].RunCmds.Count -eq 0) 'sh-005 self-heals expired idle AvdPatchLock, prunes stale RunCommand, patches, and restores original Hibernated state'

    $r6 = $execRes | Where-Object { $_.ShortHostName -eq 'sh-006' }
    $anyNonHibDeallocOn6 = @($script:mutations | Where-Object { $_ -like '*virtualMachines/sh-006/deallocate?api-version=*' }).Count
    Assert-AvdPatchTest ($r6.Status -ceq 'HibernateRestoreFailed_LeftRunning' -and $anyNonHibDeallocOn6 -eq 0 -and $script:vmStates['sh-006'].Power -ceq 'running' -and $script:vmStates['sh-006'].AllowNew -eq $true -and -not $script:vmStates['sh-006'].Tags.ContainsKey('AvdPatchLock')) 'sh-006 NEVER falls back to deallocate (hibernate=false) when hibernate fails and cleans up AvdPatchLock'

    $r7 = $execRes | Where-Object { $_.ShortHostName -eq 'sh-007' }
    Assert-AvdPatchTest ($r7.Status -ceq 'NeedsOperatorTriage' -and $script:vmStates['sh-007'].AllowNew -eq $true -and $script:vmStates['sh-007'].Power -ceq 'running') 'sh-007 expired lock with still-Running RunCommand restores allowNewSession=true and reports NeedsOperatorTriage'

    # 5f. Verify TimedOut RunCommand NEVER hibernates mid-TiWorker, and PreExistingRebootPending reports RebootedClearedPendingState
    Reset-MockFleet
    $script:vmStates['sh-001'].TimedOutCmd = $true
    $script:vmStates['sh-002'].PreRbtOnly = $true
    $resTo = @(& $runner -HostPoolResourceId $poolId -TargetSessionHostName @('sh-001', 'sh-002') -Mode Execute -PostRebootSettleSeconds 0 -ArmAccessToken 'static-test-tok' -Confirm:$false)
    Assert-AvdPatchTest (($resTo | Where-Object { $_.ShortHostName -eq 'sh-001' }).Status -ceq 'NeedsOperatorTriage' -and $script:vmStates['sh-001'].Power -ceq 'running') 'TimedOut RunCommand leaves VM Running (NeedsOperatorTriage) instead of hibernating mid-TiWorker'
    Assert-AvdPatchTest (($resTo | Where-Object { $_.ShortHostName -eq 'sh-002' }).Status -ceq 'RebootedClearedPendingState') 'PreExistingRebootPending reports accurate RebootedClearedPendingState instead of InstalledAndRebooted'

    # 5g. Verify Smart Friday Catch-Up (-CohortCount 4 -CohortIndex -1) skips recently healthy hosts
    Reset-MockFleet
    $script:simClock = [DateTimeOffset]::Parse('2026-11-20T01:00:00Z', [Globalization.CultureInfo]::InvariantCulture) # Friday
    $script:vmStates['sh-001'].Tags['AvdPatchState'] = ConvertTo-AvdPatchStateValue -LastOkUtc ($script:simClock.AddHours(-24)) -Ubr 3910 -Outcome 'InstalledNoReboot' -DeferSinceEpoch 0
    $script:vmStates['sh-002'].Tags['AvdPatchState'] = ConvertTo-AvdPatchStateValue -LastOkUtc ($script:simClock.AddHours(-24)) -Ubr 3624 -Outcome 'DeferredBaselineForUserSession' -DeferSinceEpoch 1763400000
    $resFri = @(& $runner -HostPoolResourceId $poolId -TargetSessionHostName @('sh-001', 'sh-002') -Mode Audit -CohortCount 4 -CohortIndex -1 -ArmAccessToken 'static-test-tok')
    Assert-AvdPatchTest (($resFri | Where-Object { $_.ShortHostName -eq 'sh-001' }).Status -ceq 'SkippedCatchUpAlreadyHealthy') 'Friday catch-up skips host cleanly serviced 24h ago'
    Assert-AvdPatchTest (($resFri | Where-Object { $_.ShortHostName -eq 'sh-002' }).Status -ceq 'AuditEligible') 'Friday catch-up includes host with deferred baseline update'

    # 5h. Verify 47/48-tag ceiling, MaxStartsPerMinute pacing (without CPU spin), StopNewWakeAtLocalHour cutoff, and Circuit Breaker
    Reset-MockFleet
    for ($ti = 1; $ti -le 46; $ti++) { $script:vmStates['sh-001'].Tags["ExtraTag$ti"] = "v$ti" }
    $res47 = @(& $runner -HostPoolResourceId $poolId -TargetSessionHostName 'sh-001' -Mode Execute -ArmAccessToken 'static-test-tok' -Confirm:$false)
    Assert-AvdPatchTest ($res47[0].Status -ceq 'HotpatchOrDefInstalled_BaselineDeferred' -and -not $script:vmStates['sh-001'].Tags.ContainsKey('AvdPatchLock') -and $script:vmStates['sh-001'].Tags.ContainsKey('AvdPatchState')) '47-tag VM (50-tag exact ceiling) deletes transient tags before writing AvdPatchState'

    Reset-MockFleet
    for ($ti = 1; $ti -le 47; $ti++) { $script:vmStates['sh-001'].Tags["ExtraTag$ti"] = "v$ti" }
    $res48 = @(& $runner -HostPoolResourceId $poolId -TargetSessionHostName 'sh-001' -Mode Execute -ArmAccessToken 'static-test-tok' -Confirm:$false)
    Assert-AvdPatchTest ($res48[0].Status -ceq 'SkippedTagCeilingReached' -and $script:vmStates['sh-001'].Hib) '48-tag VM is skipped with SkippedTagCeilingReached before any lock write'

    # Cutoff hour reached -> skips queued wakes
    Reset-MockFleet
    $script:simClock = [DateTimeOffset]::Parse('2026-11-17T04:30:00Z', [Globalization.CultureInfo]::InvariantCulture) # 05:30 CET
    $resCut = @(& $runner -HostPoolResourceId $poolId -TargetSessionHostName 'sh-001' -Mode Execute -StopNewWakeAtLocalHour 5 -MaintenanceTimeZoneId 'W. Europe Standard Time' -ArmAccessToken 'static-test-tok' -Confirm:$false)
    Assert-AvdPatchTest ($resCut[0].Status -ceq 'SkippedMaintenanceWindowCutoff' -and $script:vmStates['sh-001'].Hib) 'StopNewWakeAtLocalHour skips new wakes once local Berlin cutoff hour is reached'

    # Circuit breaker + post-GuestStarted exception lock cleanup + Async 202 ZonalAllocationFailed
    Reset-MockFleet
    $origRestFn = ${function:Invoke-AvdPatchRestRequest}
    function Invoke-AvdPatchRestRequest {
        param($Method, $Uri, $TokenState, $Body, $MaxRetries, [switch]$AllowNotFound)
        if ($Method -eq 'Post' -and $Uri -like '*/virtualMachines/sh-001/start*') {
            return [PSCustomObject]@{ StatusCode = 202; Headers = @{ 'Azure-AsyncOperation' = "https://management.azure.com/subscriptions/$subId/providers/Microsoft.Compute/locations/germanywestcentral/operations/op-zonal-fail" }; Body = $null }
        }
        if ($Method -eq 'Get' -and $Uri -like '*/operations/op-zonal-fail*') {
            return [PSCustomObject]@{ StatusCode = 200; Headers = @{}; Body = [PSCustomObject]@{ status = 'Failed'; error = [PSCustomObject]@{ code = 'ZonalAllocationFailed'; message = 'Allocation failed.' } } }
        }
        if ($Method -eq 'Put' -and $Uri -like '*/virtualMachines/sh-002/runCommands/AvdReboot-*') {
            throw 'Simulated unexpected post-GuestStarted exception'
        }
        return (& $origRestFn -Method $Method -Uri $Uri -TokenState $TokenState -Body $Body -MaxRetries $MaxRetries -AllowNotFound:$AllowNotFound)
    }
    $resCb = @(& $runner -HostPoolResourceId $poolId -TargetSessionHostName @('sh-001', 'sh-002', 'sh-006') -Mode Execute -MaxConcurrentVMs 1 -MaxFailureCount 1 -MaxFailurePercent 50 -PostRebootSettleSeconds 0 -ArmAccessToken 'static-test-tok' -Confirm:$false)
    Assert-AvdPatchTest (($resCb | Where-Object { $_.ShortHostName -eq 'sh-001' }).Status -ceq 'WakeAllocationFailed' -and -not $script:vmStates['sh-001'].Tags.ContainsKey('AvdPatchLock')) 'Async 202 ZonalAllocationFailed preserves hibernated RAM and cleans up lock'
    Assert-AvdPatchTest (($resCb | Where-Object { $_.ShortHostName -eq 'sh-002' }).Status -ceq 'OrchestratorError' -and -not $script:vmStates['sh-002'].Tags.ContainsKey('AvdPatchLock') -and $script:vmStates['sh-002'].Tags['Exclude-Sp-Pers-01'] -ceq 'OpsPin') 'Post-GuestStarted exception cleans up AvdPatchLock once RunCommand is deleted and preserves OpsPin'
    Assert-AvdPatchTest (($resCb | Where-Object { $_.ShortHostName -eq 'sh-006' }).Status -ceq 'AbortedByCircuitBreaker') 'Circuit breaker aborts remaining queued VMs once MaxFailureCount and MaxFailurePercent are reached'
}

# 6. LIVE Local Execution of Unicode WTSEnumerateSessionsW + DeviceGuard + ProfileList Probe
& {
    $probeRunId = [guid]::NewGuid().ToString().ToLowerInvariant()
    function Get-AvdMockWuaSession {
        $sessObj = [PSCustomObject]@{ ClientApplicationID = '' }
        Add-Member -InputObject $sessObj -MemberType ScriptMethod -Name 'CreateUpdateSearcher' -Value {
            $s = [PSCustomObject]@{ ServerSelection = 0 }
            Add-Member -InputObject $s -MemberType ScriptMethod -Name 'Search' -Value { param($c) $null = $c; return [PSCustomObject]@{ Updates = @() } }
            return $s
        }
        return $sessObj
    }
    $liveProbeScript = New-AvdPatchGuestScript -Stage 'ReadOnlyProbe' -RunId $probeRunId -MinFreeDiskGB 1
    $liveRaw = (& ([scriptblock]::Create($liveProbeScript)) | Out-String)
    Remove-Item Function:\Get-AvdMockWuaSession -ErrorAction SilentlyContinue
    $liveParsed = ConvertFrom-AvdPatchRunCommandOutput -RawOutput $liveRaw -ExpectedRunId $probeRunId
    Assert-AvdPatchTest ($null -ne $liveParsed -and $liveParsed.outcome -ceq 'LocalProbeCompleted' -and $liveParsed.wtsUsers -ge 0 -and $liveParsed.freeGB -gt 0 -and -not [string]::IsNullOrWhiteSpace($liveParsed.bootUtc)) ("Live Unicode WTSEnumerateSessionsW + DeviceGuard + UBR probe executed cleanly (wtsUsers=$($liveParsed.wtsUsers), hives=$($liveParsed.hives), ubr=$($liveParsed.ubr))")
}

Write-Output ("PASS: All {0} assertions succeeded." -f $script:assertions)
