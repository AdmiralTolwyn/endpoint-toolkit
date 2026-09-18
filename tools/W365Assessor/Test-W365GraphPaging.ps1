#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
$Tokens = $null
$ParseErrors = $null
$Ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'Invoke-W365Discovery.ps1'), [ref]$Tokens, [ref]$ParseErrors)
if ($ParseErrors.Count) { throw 'Collector parse failure' }
$Function = $Ast.Find({ param($Node) $Node -is [Management.Automation.Language.FunctionDefinitionAst] -and $Node.Name -ceq 'Invoke-GraphPaged' }, $true)
if (-not $Function) { throw 'Production pager missing' }
. ([scriptblock]::Create($Function.Extent.Text))
$script:RequestedUris = [Collections.Generic.List[string]]::new()
$script:Pages = [Collections.Generic.Queue[object]]::new()
function Invoke-MgGraphRequest {
    param($Method, $Uri)
    if ($Method -cne 'GET') { throw 'Unexpected method' }
    $script:RequestedUris.Add($Uri)
    if ($script:TestPagerClock) { $script:TestPagerClock.Elapsed.TotalSeconds = 1801 }
    if ($script:Pages.Count -eq 0) { throw 'Unexpected extra request' }
    $Page = $script:Pages.Dequeue()
    if ($Page -is [Exception]) { throw $Page }
    return $Page
}
$script:Pages.Enqueue(@{ value = @(@{ id = 'first' }); '@odata.nextLink' = 'https://example.com/collect' })
$Rejected = $false
try { Invoke-GraphPaged -Uri 'https://graph.microsoft.com/v1.0/deviceManagement/virtualEndpoint/cloudPCs' | Out-Null } catch { $Rejected = $true }
if (-not $Rejected -or $script:RequestedUris.Count -ne 1) { throw 'Untrusted continuation was followed instead of rejected before request' }
function Assert-Paging([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
function Reset-Paging { $script:RequestedUris.Clear(); $script:Pages.Clear() }
function Assert-PagingRejected([string]$InitialUri, [int]$RequestCount, [hashtable]$Options = @{}) {
    $Emitted = [Collections.Generic.List[object]]::new()
    $Rejected = $false
    try { Invoke-GraphPaged -Uri $InitialUri @Options | ForEach-Object { $Emitted.Add($_) } } catch { $Rejected = $true }
    Assert-Paging ($Rejected -and $script:RequestedUris.Count -eq $RequestCount -and $Emitted.Count -eq 0) ('Pager did not reject atomically; requests=' + $script:RequestedUris.Count + '; emitted=' + $Emitted.Count)
}
$Base = 'https://graph.microsoft.com/v1.0/deviceManagement/virtualEndpoint/cloudPCs'
$Scoped = $Base + '?$select=id&$filter=status%20eq%20%27provisioned%27&$top=50'
$OpaqueNext = $Base + '?%24top=50&%24filter=status+eq+%27provisioned%27&%24select=id&%24skiptoken=a%2Bb%2F%3D%2526'
Reset-Paging
$script:Pages.Enqueue(@{ value = @(); '@odata.nextLink' = $OpaqueNext })
$script:Pages.Enqueue(('{"value":[{"id":"second"}]}' | ConvertFrom-Json))
$Rows = @(Invoke-GraphPaged -Uri $Scoped)
Assert-Paging ($Rows.Count -eq 1 -and $Rows[0].id -eq 'second' -and $script:RequestedUris[1] -ceq $OpaqueNext) 'Opaque continuation or valid empty page changed'
foreach ($Suffix in @('?$skip=1','?$skipToken=opaque','?$skiptoken=opaque')) {
    Reset-Paging
    $script:Pages.Enqueue(@{ value = @(@{ id = 'first' }); '@odata.nextLink' = $Base + $Suffix })
    $script:Pages.Enqueue(@{ value = @(@{ id = 'second' }); '@odata.nextLink' = $null })
    Assert-Paging (@(Invoke-GraphPaged -Uri $Base).Count -eq 2) 'Supported paging token rejected'
}
foreach ($Invalid in @('http://graph.microsoft.com/v1.0/deviceManagement/virtualEndpoint/cloudPCs','https://graph.microsoft.com:444/v1.0/deviceManagement/virtualEndpoint/cloudPCs','https://user@graph.microsoft.com/v1.0/deviceManagement/virtualEndpoint/cloudPCs','https://graph.microsoft.com.evil.example/v1.0/deviceManagement/virtualEndpoint/cloudPCs','/v1.0/deviceManagement/virtualEndpoint/cloudPCs',($Base + '#fragment'),($Base + '?$expand=assignments&%24expand=other'))) {
    Reset-Paging
    Assert-PagingRejected $Invalid 0
}
foreach ($Next in @('http://graph.microsoft.com/v1.0/deviceManagement/virtualEndpoint/cloudPCs','https://example.com/collect','https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/cloudPCs','https://graph.microsoft.com/v1.0/users',($Base + '#fragment'),($Base + '?$expand=assignments'),($Base + '?$select=id'),($Base + '?$skip=-1'),($Base + '?$skip=1&$skiptoken=opaque'),($Base + '?$skiptoken=one&%24skiptoken=two'),'',0,$false,@{})) {
    Reset-Paging
    $script:Pages.Enqueue(@{ value = @(@{ id = 'first' }); '@odata.nextLink' = $Next })
    Assert-PagingRejected $Base 1
}
foreach ($Next in @(($Base + '?$skiptoken=opaque'), ($Scoped.Replace('provisioned','failed') + '&$skiptoken=opaque'))) {
    Reset-Paging
    $script:Pages.Enqueue(@{ value = @(@{ id = 'first' }); '@odata.nextLink' = $Next })
    Assert-PagingRejected $Scoped 1
}
foreach ($InvalidPage in @($null,@{},@{ id = 'singleton' },@{ value = $null },@{ value = 'invalid' },@{ value = @{ id = 'not-array' } },@{ value = @($null) },@{ value = @('scalar') },@{ value = @(@{}) },@{ value = @(@{ id = 1 }) },@{ value = @(@{ id = 'first' },@{ id = 'first' }) },@{ value = @(); error = @{ code = 'failed' } })) {
    Reset-Paging
    $script:Pages.Enqueue($InvalidPage)
    Assert-PagingRejected $Base 1
}
Reset-Paging
$script:Pages.Enqueue(@{ value = @(@{ id = 'first' }); '@odata.nextLink' = $Base })
Assert-PagingRejected $Base 1
Reset-Paging
$script:Pages.Enqueue(@{ value = @(@{ id = 'first' }); '@odata.nextLink' = $Base + '?$skiptoken=next' })
$script:Pages.Enqueue(@{ value = @(@{ id = 'first' }) })
Assert-PagingRejected $Base 2
Reset-Paging
$script:Pages.Enqueue(@{ value = @(@{ id = 'first' }); '@odata.nextLink' = $Base + '?$skiptoken=next' })
Assert-PagingRejected $Base 1 @{ MaxPages = 1 }
Reset-Paging
$script:Pages.Enqueue(@{ value = @(@{ id = 'first' },@{ id = 'second' }) })
Assert-PagingRejected $Base 1 @{ MaxRows = 1 }
Reset-Paging
$script:Pages.Enqueue(@{ value = @(@{ id = 'first' }); '@odata.nextLink' = $Base + '?$skiptoken=next' })
$script:Pages.Enqueue([InvalidOperationException]::new('Synthetic late-page failure'))
Assert-PagingRejected $Base 2
Reset-Paging
$script:Pages.Enqueue(@{ value = @() })
Assert-Paging (@(Invoke-GraphPaged -Uri $Base).Count -eq 0) 'Empty final collection rejected'

$ClockText = $Function.Extent.Text.Replace('[Diagnostics.Stopwatch]::StartNew()', '(Get-TestPagerClock)')
. ([scriptblock]::Create($ClockText))
function Get-TestPagerClock { [pscustomobject]@{ Elapsed = [pscustomobject]@{ TotalSeconds = 1801 } } }
Reset-Paging
Assert-PagingRejected $Base 0
$script:TestPagerClock = [pscustomobject]@{ Elapsed = [pscustomobject]@{ TotalSeconds = 0 } }
function Get-TestPagerClock { $script:TestPagerClock }
Reset-Paging
$script:Pages.Enqueue(@{ value = @(@{ id = 'late' }) })
Assert-PagingRejected $Base 1
$script:TestPagerClock = $null
. ([scriptblock]::Create($Function.Extent.Text))

function Write-Status { param($Message, $Level) }
foreach ($Name in @('New-CheckResult','Add-DiscoveryError')) {
    $Definition = $Ast.Find({ param($Node) $Node -is [Management.Automation.Language.FunctionDefinitionAst] -and $Node.Name -ceq $Name }, $true)
    . ([scriptblock]::Create($Definition.Extent.Text))
}
$Initializer = @($Ast.FindAll({ param($Node) $Node -is [Management.Automation.Language.AssignmentStatementAst] -and $Node.Left.Extent.Text -ceq '$Discovery' }, $true))
Assert-Paging ($Initializer.Count -eq 1) 'Expected one production discovery initializer'
$ScriptVersion = 'synthetic'
$Context = @{ Account = 'synthetic'; TenantId = '11111111-1111-4111-8111-111111111111' }
. ([scriptblock]::Create($Initializer[0].Extent.Text))
Assert-Paging ($Discovery.Inventory -is [Collections.IDictionary]) 'Production inventory cannot accept evidence adapter dictionary entries'
$Discovery.Inventory['SyntheticEvidence'] = @{ State = 'Observed' }
Assert-Paging ($Discovery.Inventory.SyntheticEvidence.State -eq 'Observed') 'Dynamic evidence entry failed'
$RoundTrip = $Discovery | ConvertTo-Json -Depth 10 | ConvertFrom-Json
Assert-Paging ($RoundTrip.Inventory.SyntheticEvidence.State -eq 'Observed' -and $RoundTrip.Inventory.CloudPCs.Count -eq 0) 'Inventory JSON object shape changed'
$CloudBlock = @($Ast.FindAll({ param($Node) $Node -is [Management.Automation.Language.TryStatementAst] -and $Node.Body.Extent.Text.Contains('$CloudPCs = @(Invoke-GraphPaged') }, $true))
Assert-Paging ($CloudBlock.Count -eq 1) 'Expected one Cloud PC production read'
$GraphBaseV1 = 'https://graph.microsoft.com/v1.0/deviceManagement/virtualEndpoint'
$AllChecks = [Collections.Generic.List[object]]::new()
$CloudPCs = @()
Reset-Paging
$script:Pages.Enqueue(@{ value = @(@{ id = 'first'; displayName = 'Synthetic' }); '@odata.nextLink' = $Base + '?$skiptoken=next' })
$script:Pages.Enqueue([InvalidOperationException]::new('Synthetic late-page failure'))
. ([scriptblock]::Create($CloudBlock[0].Extent.Text))
Assert-Paging ($CloudPCs.Count -eq 0 -and $Discovery.Inventory.CloudPCs.Count -eq 0 -and $Discovery.Errors.Count -eq 1) 'Production catch retained partial inventory or hid failure'
Assert-Paging ($AllChecks.Count -eq 1 -and $AllChecks[0].Status -eq 'Error') 'Production collection error row missing'
$Summary = @($Ast.EndBlock.Statements | Where-Object { $_.Extent.Text.Contains("-Id 'W365-INV-001'") })
Assert-Paging ($Summary.Count -eq 1) 'Expected one production inventory summary'
$ProvPols = @(); $UserSet = @(); $ANCs = @(); $DevImgs = @(); $GalImgs = @()
$StateSummary = 'none'; $ProvAssigned = 0; $AncHealthy = 0
& ([scriptblock]::Create($Summary[0].Extent.Text))
Assert-Paging ($AllChecks[-1].Status -eq 'Error') 'Production summary passed failed core collection'
if ($env:ASSAY_W365_PAGING_FIXTURE) {
    [IO.File]::WriteAllText($env:ASSAY_W365_PAGING_FIXTURE, (@{ CheckResults = @($AllChecks.ToArray()); Errors = @($Discovery.Errors) } | ConvertTo-Json -Depth 10), [Text.UTF8Encoding]::new($false))
}
$Discovery.Errors = @()
& ([scriptblock]::Create($Summary[0].Extent.Text))
Assert-Paging ($AllChecks[-1].Status -eq 'Pass') 'Successful inventory summary changed'
Write-Output 'PASS: trusted origin/path/query, opaque continuations and empty pages, collection shape, row IDs, finite budgets and atomic failure; no network calls.'