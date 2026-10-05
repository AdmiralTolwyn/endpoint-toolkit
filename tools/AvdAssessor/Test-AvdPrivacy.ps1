#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
$CollectorPath = Join-Path $PSScriptRoot 'Invoke-AvdDiscovery.ps1'
$Tokens = $null
$ParseErrors = $null
$Ast = [Management.Automation.Language.Parser]::ParseFile($CollectorPath, [ref]$Tokens, [ref]$ParseErrors)
if ($ParseErrors.Count) { throw 'Collector parse failure' }
function Assert-Privacy([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }

foreach ($Name in @('New-CheckResult','ConvertTo-AvdTagKeys','Add-AvdTagValues','ConvertTo-AvdRdpProperties','Get-AvdRoleAssignmentSummary','Invoke-AvdMdeHuntingQuery')) {
    $Definition = $Ast.Find({ param($Node) $Node -is [Management.Automation.Language.FunctionDefinitionAst] -and $Node.Name -ceq $Name }, $true)
    Assert-Privacy ($null -ne $Definition) ('Missing production function ' + $Name)
    . ([scriptblock]::Create($Definition.Extent.Text))
}
$Allowlist = $Ast.Find({ param($Node) $Node -is [Management.Automation.Language.AssignmentStatementAst] -and $Node.Left.Extent.Text -ceq '$Script:AvdRdpAllowlist' }, $true)
. ([scriptblock]::Create($Allowlist.Extent.Text))
. (Join-Path $PSScriptRoot 'CollectorPrivacy.ps1')
function Write-Status { param($Message, $Level) }
function New-TestContext([string]$Mode) {
    [pscustomobject]@{ Mode = $Mode; Key = $(if ($Mode -eq 'Pseudonymous') { [byte[]](1..32) } else { $null }); KeyId = 'testkey'; KeyPath = $null; Identities = [Collections.Generic.Dictionary[string,string]]::new() }
}
$Sentinels = 'pii\.sentinel|HOST_SENTINEL|CC_SENTINEL|PRINCIPAL_SENTINEL|203\.0\.113\.77|198\.51\.100\.53|ERROR_SENTINEL'

# Host-pool projection and RDP checks, from the production loop up to the SSO check.
$HostPoolLoop = @($Ast.FindAll({ param($Node) $Node -is [Management.Automation.Language.ForEachStatementAst] -and $Node.Variable.Extent.Text -ceq '$HP' -and $Node.Body.Extent.Text.Contains('$RdpSummary = ConvertTo-AvdRdpProperties') }, $true))
Assert-Privacy ($HostPoolLoop.Count -eq 1) 'Expected one host-pool projection loop'
$LoopStatements = @($HostPoolLoop[0].Body.Statements)
$LastIndex = [array]::FindIndex($LoopStatements, [Predicate[object]]{ param($Statement) $Statement.Extent.Text.Contains('-Id "IAM-SSO-') })
Assert-Privacy ($LastIndex -gt 0) 'SSO check not found in host-pool loop'
$HostPoolBlock = [scriptblock]::Create((($LoopStatements[0..$LastIndex] | ForEach-Object { $_.Extent.Text }) -join "`n"))
$HP = [pscustomobject]@{
    Name = 'hp1'; Id = '/subscriptions/s/resourceGroups/rg/providers/Microsoft.DesktopVirtualization/hostPools/hp1'
    HostPoolType = 'Pooled'; LoadBalancerType = 'BreadthFirst'; MaxSessionLimit = 10; PreferredAppGroupType = 'Desktop'
    StartVMOnConnect = $true; ValidationEnvironment = $false; Location = 'westeurope'
    Tag = @{ Owner = 'pii.sentinel@example.test'; CostCenter = 'CC_SENTINEL' }
    CustomRdpProperty = 'drivestoredirect:s:*;redirectclipboard:i:1;full address:s:HOST_SENTINEL;username:s:pii.sentinel@example.test;enablerdsaadauth:i:1;targetisaadjoined:i:1'
}
function Invoke-HostPoolProjection([string]$Mode, [bool]$ExportTags) {
    $script:PrivacyContext = New-TestContext $Mode
    $Script:ExportTagValues = $ExportTags
    $SubId = 'sub'
    $RedirDefaultCaveat = 'default'
    $Discovery = [pscustomobject]@{ Inventory = [pscustomobject]@{ HostPools = @() } }
    $AllChecks = [Collections.ArrayList]::new()
    . $HostPoolBlock
    [pscustomobject]@{ Discovery = $Discovery; Checks = $AllChecks.ToArray() }
}
$Run = Invoke-HostPoolProjection 'Pseudonymous' $false
$Pool = $Run.Discovery.Inventory.HostPools[0]
$PoolJson = $Run | ConvertTo-Json -Depth 10
Assert-Privacy ($PoolJson -notmatch $Sentinels) ('Host-pool export leaked: ' + [regex]::Match($PoolJson, $Sentinels).Value)
Assert-Privacy ((@($Pool.TagKeys) -join ',') -ceq 'CostCenter,Owner' -and $null -eq $Pool.PSObject.Properties['Tags'] -and $null -eq $Pool.PSObject.Properties['CustomRdpProperty']) 'Host-pool tags or raw RDP string exported'
Assert-Privacy ((@($Pool.RdpProperties.Keys) -join ',') -ceq 'drivestoredirect,redirectclipboard,enablerdsaadauth,targetisaadjoined' -and $Pool.RdpUnknownPropertyCount -eq 2) 'RDP allowlist incorrect'
$Checks = @{}
foreach ($Check in $Run.Checks) { $Checks[$Check.Id] = $Check }
Assert-Privacy ($Checks['SEC-RDP-hp1'].Details.Contains('Drives:Open') -and $Checks['SEC-RDP-hp1'].Details.Contains('Clipboard:Open') -and $Checks['SEC-RDP-hp1'].Evidence.OtherPropertyCount -eq 2) 'RDP summary lost evaluated properties'
Assert-Privacy ($Checks['SEC-DRIVE-hp1'].Status -eq 'Warning' -and $Checks['SEC-CLIP-hp1'].Status -ne 'Pass' -and $Checks['IAM-SSO-hp1'].Status -eq 'Pass') 'RDP checks changed after allowlisting'
$Run = Invoke-HostPoolProjection 'Identified' $true
Assert-Privacy ($Run.Discovery.Inventory.HostPools[0].Tags.Owner -ceq 'pii.sentinel@example.test') 'Identified tag values opt-in failed'
Assert-Privacy (($Run | ConvertTo-Json -Depth 10) -notmatch 'HOST_SENTINEL') 'Unknown RDP property exported in identified mode'
$Run = Invoke-HostPoolProjection 'Identified' $false
Assert-Privacy ($null -eq $Run.Discovery.Inventory.HostPools[0].PSObject.Properties['Tags']) 'Tag values exported without opt-in'

# VNet DNS and NSG prefix projection.
$VNetStatement = $Ast.Find({ param($Node) $Node -is [Management.Automation.Language.AssignmentStatementAst] -and $Node.Left.Extent.Text -ceq '$VNetObj' }, $true)
$NsgStatement = $Ast.Find({ param($Node) $Node -is [Management.Automation.Language.AssignmentStatementAst] -and $Node.Left.Extent.Text -ceq '$Discovery.Inventory.NSGs' }, $true)
Assert-Privacy ($VNetStatement -and $NsgStatement) 'Network projections not found'
$VNet = [pscustomobject]@{ Name = 'vnet'; Id = 'vnet-id'; Location = 'westeurope'; Tag = @{ Owner = 'pii.sentinel@example.test' }; AddressSpace = [pscustomobject]@{ AddressPrefixes = @('10.0.0.0/16') }; Subnets = @(); VirtualNetworkPeerings = @(); DhcpOptions = [pscustomobject]@{ DnsServers = @('10.0.0.4','198.51.100.53') } }
$NSG = [pscustomobject]@{ Name = 'nsg'; Id = 'nsg-id'; SecurityRules = @(
    [pscustomobject]@{ Name = 'r1'; Priority = 100; Direction = 'Inbound'; Access = 'Allow'; Protocol = 'Tcp'; SourcePortRange = '*'; DestinationPortRange = '3389'; SourceAddressPrefix = '203.0.113.77/32'; DestinationAddressPrefix = '10.0.1.0/24' },
    [pscustomobject]@{ Name = 'r2'; Priority = 200; Direction = 'Outbound'; Access = 'Allow'; Protocol = '*'; SourcePortRange = '*'; DestinationPortRange = '443'; SourceAddressPrefix = 'VirtualNetwork'; DestinationAddressPrefix = 'WindowsVirtualDesktop' },
    [pscustomobject]@{ Name = 'r3'; Priority = 300; Direction = 'Inbound'; Access = 'Deny'; Protocol = '*'; SourcePortRange = '*'; DestinationPortRange = '*'; SourceAddressPrefix = '0.0.0.0/0'; DestinationAddressPrefix = '*' }
) }
foreach ($Mode in @('Pseudonymous','Identified')) {
    $script:PrivacyContext = New-TestContext $Mode
    $Script:ExportTagValues = $false
    $VNetSub = 'sub'; $VNetRG = 'rg'; $NSGSub = 'sub'; $NSGRG = 'rg'
    $Discovery = [pscustomobject]@{ Inventory = [pscustomobject]@{ NSGs = @() } }
    . ([scriptblock]::Create($VNetStatement.Extent.Text))
    . ([scriptblock]::Create($NsgStatement.Extent.Text))
    $Rules = $Discovery.Inventory.NSGs[0].Rules
    if ($Mode -eq 'Pseudonymous') {
        Assert-Privacy ((@($VNetObj.DnsServers) -join ',') -ceq '10.0.0.4,Public' -and (@($VNetObj.TagKeys) -join ',') -ceq 'Owner' -and $null -eq $VNetObj.PSObject.Properties['Tags']) 'VNet DNS or tags not classified'
        Assert-Privacy ($Rules[0].SourceAddressPrefix -ceq 'Public/32' -and $Rules[0].DestinationAddressPrefix -ceq '10.0.1.0/24' -and $Rules[1].DestinationAddressPrefix -ceq 'WindowsVirtualDesktop' -and $Rules[2].SourceAddressPrefix -ceq '0.0.0.0/0' -and $Rules[2].DestinationAddressPrefix -ceq '*') 'NSG prefixes not classified'
    } else {
        Assert-Privacy ((@($VNetObj.DnsServers) -join ',') -ceq '10.0.0.4,198.51.100.53' -and $Rules[0].SourceAddressPrefix -ceq '203.0.113.77/32') 'Identified network values not retained'
    }
}

# RBAC labels and Defender device names.
$Assignments = @(
    [pscustomobject]@{ RoleDefinitionName = 'Owner'; ObjectType = 'User'; DisplayName = 'PRINCIPAL_SENTINEL one' },
    [pscustomobject]@{ RoleDefinitionName = 'Owner'; ObjectType = 'User'; DisplayName = 'PRINCIPAL_SENTINEL two' },
    [pscustomobject]@{ RoleDefinitionName = 'Contributor'; ObjectType = 'Group'; DisplayName = 'PRINCIPAL_SENTINEL group' }
)
$script:PrivacyContext = New-TestContext 'Pseudonymous'
Assert-Privacy ((@(Get-AvdRoleAssignmentSummary $Assignments) -join ', ') -ceq 'Contributor (Group x1), Owner (User x2)') 'RBAC summary exported principal names'
$script:PrivacyContext = New-TestContext 'Identified'
Assert-Privacy ((@(Get-AvdRoleAssignmentSummary $Assignments) -join ', ').Contains('PRINCIPAL_SENTINEL one')) 'Identified RBAC names lost'
$VmId = '/subscriptions/s/resourceGroups/rg/providers/Microsoft.Compute/virtualMachines/vm1'
function Invoke-RestMethod { [pscustomobject]@{ results = @([pscustomobject]@{ DeviceId = 'device-1'; DeviceName = 'HOST_SENTINEL'; AzureResourceId = $VmId; AzureVmId = 'vm-guid'; Timestamp = '2026-10-01T00:00:00Z'; OnboardingStatus = 'Onboarded'; SensorHealthState = 'Active' }) } }
$script:PrivacyContext = New-TestContext 'Pseudonymous'
$Result = Invoke-AvdMdeHuntingQuery -ResourceIds @($VmId) -Token 'token'
Assert-Privacy ($Result.Ok -and $Result.Rows[0].DeviceId -ceq 'device-1' -and $null -eq $Result.Rows[0].PSObject.Properties['DeviceName']) 'Defender device name exported'
$script:PrivacyContext = New-TestContext 'Identified'
$Result = Invoke-AvdMdeHuntingQuery -ResourceIds @($VmId) -Token 'token'
Assert-Privacy ($Result.Rows[0].DeviceName -ceq 'HOST_SENTINEL') 'Identified Defender device name lost'

# Guest temp script is removed on success, handled failure and escaping failure.
$GuestBranch = $Ast.Find({ param($Node) $Node -is [Management.Automation.Language.IfStatementAst] -and $Node.Clauses[0].Item1.Extent.Text -ceq '-not $IncludeGuestChecks' }, $true)
Assert-Privacy ($null -ne $GuestBranch) 'Guest branch not found'
$GuestDefinition = $Ast.Find({ param($Node) $Node -is [Management.Automation.Language.FunctionDefinitionAst] -and $Node.Name -ceq 'Add-GuestNA' }, $true)
. ([scriptblock]::Create($GuestDefinition.Extent.Text))
$Temp = Join-Path ([IO.Path]::GetTempPath()) ('avd-privacy-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($Temp)
$OriginalTemp = $env:TEMP
try {
    $env:TEMP = $Temp
    function Set-AzContext { param($SubscriptionId, $ErrorAction, $WarningAction) }
    function Invoke-AzVMRunCommand { param($ResourceGroupName, $VMName, $CommandId, $ScriptPath, $ErrorAction) $script:GuestScriptSeen = Test-Path -LiteralPath $ScriptPath; throw 'Run Command failed for pii.sentinel@example.test ERROR_SENTINEL' }
    $script:PrivacyContext = New-TestContext 'Pseudonymous'
    $IncludeGuestChecks = $true
    $FslMinVersion = [version]'2.9.8612.60056'
    $Discovery = [pscustomobject]@{ Inventory = [pscustomobject]@{
        HostPools = @([pscustomobject]@{ Name = 'hp1' })
        SessionHosts = @([pscustomobject]@{ HostPoolName = 'hp1'; PowerState = 'VM running'; ResourceId = $VmId; ResourceGroup = 'rg'; Name = 'hp1/vm1' })
    } }
    $AllChecks = [Collections.ArrayList]::new()
    $script:GuestScriptSeen = $false
    . ([scriptblock]::Create($GuestBranch.Extent.Text))
    Assert-Privacy ($script:GuestScriptSeen -and @(Get-ChildItem -LiteralPath $Temp -Filter 'avd_fslogix_guest_*.ps1').Count -eq 0) 'Guest script not removed after handled failure'
    Assert-Privacy ((@($AllChecks.ToArray()) | ConvertTo-Json -Depth 6) -notmatch $Sentinels) 'Guest error text leaked'
    function Write-Status { param($Message, $Level) if ($Message -match 'Run Command failed') { throw 'escaping failure' } }
    $Escaped = $false
    try { . ([scriptblock]::Create($GuestBranch.Extent.Text)) } catch { $Escaped = $true }
    Assert-Privacy ($Escaped -and @(Get-ChildItem -LiteralPath $Temp -Filter 'avd_fslogix_guest_*.ps1').Count -eq 0) 'Guest script not removed after escaping failure'
    function Write-Status { param($Message, $Level) }

    # Export manifest, protection and pre-collection gates.
    $ExportStatements = @($Ast.EndBlock.Statements | Where-Object { $_.Extent.Text.Contains('New-CollectorPrivacyManifest') -or $_.Extent.Text.Contains('Write-CollectorExport') })
    Assert-Privacy ($ExportStatements.Count -eq 2) 'Expected manifest and export statements'
    $Export = [scriptblock]::Create((($ExportStatements | ForEach-Object { $_.Extent.Text }) -join "`n"))
    $Script:ExportTagValues = $false
    $IncludeGuestChecks = [switch]$true
    $IncludeMdeDeviceChecks = [switch]$false
    $Discovery = [pscustomobject]@{ SchemaVersion = '1.0'; CollectionId = 'collection'; Assessor = $null; Privacy = $null; Inventory = (Invoke-HostPoolProjection 'Pseudonymous' $false).Discovery.Inventory; CheckResults = @(); Errors = @() }
    $script:PrivacyContext = New-CollectorPrivacyContext -Mode Pseudonymous -OutputPath (Join-Path $Temp 'avd.json')
    $OutputPath = Join-Path $Temp 'avd.json'
    $IdentityMapPath = $null
    . $Export
    $Written = Get-Content -LiteralPath $OutputPath -Raw
    $Parsed = $Written | ConvertFrom-Json
    Assert-Privacy ($Parsed.Privacy.Mode -ceq 'Pseudonymous' -and (@($Parsed.Privacy.OptIns) -join ',') -ceq 'IncludeGuestChecks' -and $null -eq $Parsed.PSObject.Properties['AssessorId']) 'AVD manifest incorrect'
    Assert-Privacy ($Written -notmatch $Sentinels -and (Get-Acl -LiteralPath $OutputPath).AreAccessRulesProtected) 'AVD export leaked or is not protected'
    $Rejected = $false
    try { . $Export } catch { $Rejected = $true }
    Assert-Privacy $Rejected 'Existing AVD export was overwritten'
    foreach ($Case in @(
        @{ Arguments = @{ PrivacyMode = 'Identified'; OutputPath = (Join-Path $Temp 'identified.json') }; Pattern = 'ConfirmIdentifiedExport' },
        @{ Arguments = @{ OutputPath = $OutputPath }; Pattern = 'already exists' }
    )) {
        $Arguments = $Case.Arguments
        $Message = $null
        try { & $CollectorPath @Arguments *> $null } catch { $Message = $_.Exception.Message }
        Assert-Privacy ($Message -match $Case.Pattern) ('Collector did not stop before collection: ' + $Case.Pattern)
    }
    Assert-Privacy (-not (Test-Path -LiteralPath (Join-Path $Temp 'identified.json'))) 'Rejected identified run wrote output'
} finally {
    $env:TEMP = $OriginalTemp
    Remove-Item -LiteralPath $Temp -Recurse -Force -ErrorAction SilentlyContinue
}
Write-Output 'PASS: AVD tag keys/opt-in values, RDP allowlist, DNS/NSG classification, RBAC and Defender name omission, guest script cleanup, sanitized guest errors, manifest and protected output.'
