#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
$Tokens = $null
$ParseErrors = $null
$Ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'Invoke-W365Discovery.ps1'), [ref]$Tokens, [ref]$ParseErrors)
if ($ParseErrors.Count) { throw 'Collector parse failure' }
$ContextHelper = $Ast.Find({ param($Node) $Node -is [Management.Automation.Language.FunctionDefinitionAst] -and $Node.Name -ceq 'Get-W365GraphContextIssue' }, $true)
if ($ContextHelper) { . ([scriptblock]::Create($ContextHelper.Extent.Text)) }
$ScopeStart = $Ast.Find({ param($Node) $Node -is [Management.Automation.Language.AssignmentStatementAst] -and $Node.Left.Extent.Text -ceq '$Scopes' }, $true)
$DiscoveryStart = $Ast.Find({ param($Node) $Node -is [Management.Automation.Language.AssignmentStatementAst] -and $Node.Left.Extent.Text -ceq '$Discovery' }, $true)
if (-not $ScopeStart -or -not $DiscoveryStart) { throw 'Authentication boundaries missing' }
$AuthText = ($Ast.EndBlock.Statements | Where-Object { $_.Extent.StartOffset -ge $ScopeStart.Extent.StartOffset -and $_.Extent.EndOffset -le $DiscoveryStart.Extent.StartOffset } | ForEach-Object { $_.Extent.Text }) -join "`n"
$Production = [scriptblock]::Create($AuthText.Replace('exit 1', "throw 'Authentication stopped'"))
function Write-Status { param($Message, $Level) }
function Get-MgContext { $script:GraphContext }
function Connect-MgGraph { $script:ConnectCalls++; throw 'Test must not perform sign-in' }
$TenantId = '11111111-1111-4111-8111-111111111111'
$SkipLogin = $true
$script:ConnectCalls = 0
$script:GraphContext = [pscustomobject]@{
    TenantId = '22222222-2222-4222-8222-222222222222'
    Account = 'synthetic@example.test'
    AuthType = 'Delegated'
    Environment = 'Global'
    Scopes = @('CloudPC.Read.All','DeviceManagementConfiguration.Read.All','DeviceManagementManagedDevices.Read.All','Directory.Read.All')
}
$Rejected = $false
try { & $Production } catch { $Rejected = $true }
if (-not $Rejected -or $script:ConnectCalls -ne 0) { throw 'SkipLogin accepted the wrong tenant or attempted sign-in' }
$script:GraphContext.TenantId = $TenantId
$script:GraphContext.Scopes = @('CloudPC.Read.All')
$Rejected = $false
try { & $Production } catch { $Rejected = $true }
if (-not $Rejected -or $script:ConnectCalls -ne 0) { throw 'SkipLogin accepted missing core scopes or attempted sign-in' }
$Production = [scriptblock]::Create($AuthText.Replace('exit 1', "throw 'Authentication stopped'") + "`n" + '[pscustomobject]@{ ValidatedTenant = $Context.TenantId; RequestedScopes = $RequestScopes }')
function Assert-Auth([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
function New-TestGraphContext {
    [pscustomobject]@{
        TenantId = '11111111-1111-4111-8111-111111111111'
        Account = 'synthetic@example.test'
        AuthType = 'Delegated'
        Environment = 'Global'
        Scopes = @('CloudPC.Read.All','DeviceManagementConfiguration.Read.All','DeviceManagementManagedDevices.Read.All')
    }
}
function Get-MgContext { $script:ContextReads++; $script:GraphContext }
function Connect-MgGraph {
    [CmdletBinding()]
    param([string[]]$Scopes, [string]$TenantId, [string]$Environment, [string]$ContextScope, [switch]$NoWelcome)
    $script:ConnectCalls++
    $script:ConnectArguments = @{ Scopes = $Scopes; TenantId = $TenantId; TenantSpecified = $PSBoundParameters.ContainsKey('TenantId'); Environment = $Environment; ContextScope = $ContextScope; NoWelcome = [bool]$NoWelcome }
    if ($script:ConnectFails) { throw 'Synthetic sign-in failure' }
    $script:GraphContext = $script:PostConnectContext
}
function Invoke-TestAuthentication {
    param([object]$Existing, [object]$After = $null, [bool]$Skip = $true, [bool]$IncludeCa = $false, [string]$Tenant = '11111111-1111-4111-8111-111111111111', [bool]$ConnectFails = $false)
    $script:GraphContext = $Existing
    $script:PostConnectContext = $After
    $script:ConnectCalls = 0
    $script:ContextReads = 0
    $script:ConnectArguments = $null
    $script:ConnectFails = $ConnectFails
    $TenantId = $Tenant
    $SkipLogin = $Skip
    $IncludeConditionalAccess = $IncludeCa
    $Rejected = $false
    $Result = $null
    try { $Result = & $Production } catch { $Rejected = $true }
    [pscustomobject]@{ Rejected = $Rejected; Connects = $script:ConnectCalls; Reads = $script:ContextReads; Result = $Result; Arguments = $script:ConnectArguments }
}

$Valid = New-TestGraphContext
$Case = Invoke-TestAuthentication -Existing $Valid
Assert-Auth (-not $Case.Rejected -and $Case.Connects -eq 0 -and $Case.Reads -eq 1 -and $Case.Result.RequestedScopes.Count -eq 3) 'Valid core-only session was not reused'
Assert-Auth ($Case.Result.RequestedScopes -notcontains 'Directory.Read.All' -and $Case.Result.RequestedScopes -notcontains 'Policy.Read.All') 'Unneeded scope requested by default'
$Case = Invoke-TestAuthentication -Existing $null
Assert-Auth ($Case.Rejected -and $Case.Connects -eq 0) 'SkipLogin signed in without a session'
foreach ($Mutation in @(
    @{ Field = 'TenantId'; Value = '22222222-2222-4222-8222-222222222222' },
    @{ Field = 'TenantId'; Value = '' },
    @{ Field = 'TenantId'; Value = 'not-a-guid' },
    @{ Field = 'TenantId'; Value = [guid]::Empty.ToString() },
    @{ Field = 'AuthType'; Value = 'AppOnly' },
    @{ Field = 'AuthType'; Value = 'UserProvidedAccessToken' },
    @{ Field = 'Environment'; Value = 'USGov' },
    @{ Field = 'Environment'; Value = $null },
    @{ Field = 'Account'; Value = ' ' },
    @{ Field = 'Scopes'; Value = $null },
    @{ Field = 'Scopes'; Value = 'CloudPC.Read.All DeviceManagementConfiguration.Read.All DeviceManagementManagedDevices.Read.All' },
    @{ Field = 'Scopes'; Value = @('CloudPC.Read.All',1) }
)) {
    $Invalid = New-TestGraphContext
    $Invalid.($Mutation.Field) = $Mutation.Value
    $Case = Invoke-TestAuthentication -Existing $Invalid
    Assert-Auth ($Case.Rejected -and $Case.Connects -eq 0) ('Invalid reused context accepted: ' + $Mutation.Field)
    $Case = Invoke-TestAuthentication -Existing $null -After $Invalid -Skip $false
    Assert-Auth ($Case.Rejected -and $Case.Connects -eq 1 -and $Case.Reads -eq 2) ('Invalid post-connect context accepted: ' + $Mutation.Field)
}
foreach ($InvalidTenant in @('common','organizations','example.onmicrosoft.com','not-a-guid',[guid]::Empty.ToString(),' ')) {
    $Case = Invoke-TestAuthentication -Existing $Valid -Skip $false -Tenant $InvalidTenant
    Assert-Auth ($Case.Rejected -and $Case.Connects -eq 0 -and $Case.Reads -eq 0) 'Ambiguous tenant selector reached authentication'
}
foreach ($MissingScope in $Valid.Scopes) {
    $Invalid = New-TestGraphContext
    $Invalid.Scopes = @($Invalid.Scopes | Where-Object { $_ -ne $MissingScope })
    $Case = Invoke-TestAuthentication -Existing $Invalid
    Assert-Auth ($Case.Rejected -and $Case.Connects -eq 0) 'Missing core scope bypassed SkipLogin'
    $Case = Invoke-TestAuthentication -Existing $Invalid -After $Valid -Skip $false
    Assert-Auth (-not $Case.Rejected -and $Case.Connects -eq 1 -and $Case.Reads -eq 2) 'Valid refreshed core session rejected'
    Assert-Auth ($Case.Arguments.TenantId -eq $Valid.TenantId -and $Case.Arguments.Environment -ceq 'Global' -and $Case.Arguments.ContextScope -ceq 'Process' -and $Case.Arguments.NoWelcome -and $Case.Arguments.Scopes.Count -eq 3) 'Connection parameters did not preserve the tenant/scope boundary'
}
$WrongTenant = New-TestGraphContext
$WrongTenant.TenantId = '22222222-2222-4222-8222-222222222222'
$Case = Invoke-TestAuthentication -Existing $WrongTenant -After $Valid -Skip $false
Assert-Auth (-not $Case.Rejected -and $Case.Connects -eq 1 -and $Case.Arguments.TenantId -eq $Valid.TenantId) 'Explicit tenant was not used for reconnect'
$Case = Invoke-TestAuthentication -Existing $Valid -IncludeCa $true
Assert-Auth ($Case.Rejected -and $Case.Connects -eq 0) 'Requested CA scope was not required'
$WithCa = New-TestGraphContext
$WithCa.Scopes += 'Policy.Read.All'
$Case = Invoke-TestAuthentication -Existing $Valid -After $WithCa -IncludeCa $true -Skip $false
Assert-Auth (-not $Case.Rejected -and $Case.Connects -eq 1 -and $Case.Arguments.Scopes.Count -eq 4 -and $Case.Arguments.Scopes -contains 'Policy.Read.All') 'CA opt-in connection scope incorrect'
$Case = Invoke-TestAuthentication -Existing $WithCa -IncludeCa $false
Assert-Auth (-not $Case.Rejected -and $Case.Connects -eq 0 -and $Case.Result.RequestedScopes.Count -eq 3) 'Broader cached scopes changed the requested collection'
$OtherWithCa = New-TestGraphContext
$OtherWithCa.TenantId = $WrongTenant.TenantId
$OtherWithCa.Scopes += 'Policy.Read.All'
$Case = Invoke-TestAuthentication -Existing $WrongTenant -After $OtherWithCa -IncludeCa $true -Skip $false -Tenant ''
Assert-Auth (-not $Case.Rejected -and $Case.Arguments.TenantId -eq $WrongTenant.TenantId) 'Omitted selector changed an existing tenant during reauthorization'
$Case = Invoke-TestAuthentication -Existing $WrongTenant -After $WithCa -IncludeCa $true -Skip $false -Tenant ''
Assert-Auth $Case.Rejected 'Reauthorization silently changed existing tenant'
$Case = Invoke-TestAuthentication -Existing $null -After $WrongTenant -Skip $false -Tenant ''
Assert-Auth (-not $Case.Rejected -and -not $Case.Arguments.TenantSpecified -and $Case.Result.ValidatedTenant -eq $WrongTenant.TenantId) 'New sign-in tenant selection did not validate correctly'
$Case = Invoke-TestAuthentication -Existing $null -After $null -Skip $false
Assert-Auth ($Case.Rejected -and $Case.Connects -eq 1) 'Null post-connect context accepted'
$Case = Invoke-TestAuthentication -Existing $null -Skip $false -ConnectFails $true
Assert-Auth ($Case.Rejected -and $Case.Connects -eq 1 -and $Case.Reads -eq 1) 'Connection failure continued'
$Case = Invoke-TestAuthentication -Existing $WrongTenant -After $Valid -Skip $false
Assert-Auth (-not $Case.Rejected) 'Validated context unavailable for export test'
$Context = $script:GraphContext
$ScriptVersion = 'synthetic'
. ([scriptblock]::Create($DiscoveryStart.Extent.Text))
Assert-Auth ($Discovery.TenantId -eq $Valid.TenantId -and $Discovery.TenantId -eq $Case.Result.ValidatedTenant -and $Discovery.AssessorId -eq $Valid.Account) 'Export metadata did not use validated context'
Write-Output 'PASS: delegated Global tenant/scope binding, strict SkipLogin, process-scoped reconnect, post-connect checks and optional permission selection; all authentication mocked.'