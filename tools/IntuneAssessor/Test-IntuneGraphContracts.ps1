#Requires -Version 5.1
[CmdletBinding()]
param([Parameter(Mandatory)][string]$EvidenceDirectory, [switch]$CheckDocumentation)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Invoke-IntuneDiscovery.ps1') -LibraryOnly
$Contracts = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'GraphContracts.json') -Raw | ConvertFrom-Json
$Results = [Collections.Generic.List[object]]::new()
$Metadata = @{}
foreach ($Version in @('v1.0', 'beta')) {
    $Path = Join-Path $EvidenceDirectory ('intune-audit-' + $Version + '.xml')
    $Settings = [Xml.XmlReaderSettings]::new()
    $Settings.DtdProcessing = [Xml.DtdProcessing]::Prohibit
    $Settings.XmlResolver = $null
    $Reader = [Xml.XmlReader]::Create([IO.Path]::GetFullPath($Path), $Settings)
    $Document = [Xml.XmlDocument]::new()
    $Document.XmlResolver = $null
    try { $Document.Load($Reader) } finally { $Reader.Dispose() }
    $Types = @{}
    foreach ($Schema in $Document.SelectNodes('//*[local-name()="Schema"]')) {
        foreach ($Type in $Schema.SelectNodes('*[local-name()="EntityType" or local-name()="ComplexType"]')) {
            $Types[$Schema.GetAttribute('Namespace') + '.' + $Type.GetAttribute('Name')] = $Type
            if ($Schema.GetAttribute('Alias')) { $Types[$Schema.GetAttribute('Alias') + '.' + $Type.GetAttribute('Name')] = $Type }
        }
    }
    $Metadata[$Version] = @{ Types = $Types; Container = $Document.SelectSingleNode('//*[local-name()="EntityContainer"]'); Hash = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash }
}

function Get-ContractMembers {
    param($Schema, [string]$TypeName)
    $Members = @{}
    $Visited = [Collections.Generic.HashSet[string]]::new()
    while ($TypeName -and $Schema.Types.ContainsKey($TypeName) -and $Visited.Add($TypeName)) {
        $Type = $Schema.Types[$TypeName]
        foreach ($Member in $Type.SelectNodes('*[local-name()="Property" or local-name()="NavigationProperty"]')) { $Members[$Member.GetAttribute('Name')] = $Member }
        $TypeName = $Type.GetAttribute('BaseType')
    }
    return $Members
}
function Get-ContractTypeName([string]$TypeName) { return $TypeName -replace '^Collection\((.+)\)$', '$1' }
if (-not (Get-ContractMembers $Metadata['v1.0'] 'graph.managedDevice').ContainsKey('id') -or
    -not (Get-ContractMembers $Metadata['v1.0'] 'microsoft.graph.managedDevice').ContainsKey('id') -or
    -not (Get-ContractMembers $Metadata['v1.0'] 'graph.deviceManagement').ContainsKey('managedDevices') -or
    (Get-ContractMembers $Metadata['v1.0'] 'graph.managedDevice').ContainsKey('inventedAuditControl') -or
    (Get-ContractMembers $Metadata['v1.0'] 'graph.win32LobApp').ContainsKey('detectionRules')) {
    throw 'Schema verifier controls failed; no contract results are trustworthy.'
}
foreach ($Contract in $Contracts) {
    $Enabled = [bool](Get-IntuneValue $Contract 'enabled' $true)
    $Fields = if ($null -ne $Contract.PSObject.Properties['fields']) { @($Contract.fields) } else { @(Get-IntuneFieldList $Contract.module) }
    if ('target' -in $Fields) {
        $TargetMembers = Get-ContractMembers $Metadata[$Contract.version] 'graph.groupAssignmentTarget'
        $Probe = @{ id = 'audit'; target = @{ '@odata.type' = '#microsoft.graph.groupAssignmentTarget'; groupId = 'audit'; entraObjectId = 'audit'; targetType = 'audit'; deviceAndAppManagementAssignmentFilterId = 'audit'; deviceAndAppManagementAssignmentFilterType = 'include' } }
        $Projected = ConvertTo-IntuneSafeRow -Module $Contract.module -Row $Probe
        foreach ($Key in $Projected.target.Keys) {
            if ($Key -ne '@odata.type' -and -not $TargetMembers.ContainsKey($Key)) { throw ('Unsupported projected target field: ' + $Contract.module + '.' + $Key) }
        }
    }
    $Moniker = if ($Contract.version -eq 'v1.0') { '1.0' } else { $Contract.version }
    $Source = 'https://learn.microsoft.com/en-us/graph/api/' + $Contract.doc + '?view=graph-rest-' + $Moniker
    $Schema = $Metadata[$Contract.version]
    $TypeName = ''
    $RouteValid = $true
    foreach ($Segment in $Contract.path.Split('/')) {
        if ($Segment -eq '{id}') { continue }
        if (-not $TypeName) {
            $Member = $Schema.Container.SelectSingleNode('*[@Name="' + $Segment + '"]')
            if ($null -eq $Member) { $RouteValid = $false; break }
            $TypeName = Get-ContractTypeName ($Member.GetAttribute('Type') + $Member.GetAttribute('EntityType'))
        } else {
            $Members = Get-ContractMembers $Schema $TypeName
            if (-not $Members.ContainsKey($Segment)) { $RouteValid = $false; break }
            $TypeName = Get-ContractTypeName $Members[$Segment].GetAttribute('Type')
        }
    }
    $Family = [Collections.Generic.List[string]]::new()
    $Family.Add($TypeName)
    for ($Index = 0; $Index -lt $Family.Count; $Index++) {
        foreach ($Name in $Schema.Types.Keys) { if ($Schema.Types[$Name].GetAttribute('BaseType') -ceq $Family[$Index]) { $Family.Add($Name) } }
    }
    $FieldResults = @(foreach ($Field in $Fields) {
        $Owners = @()
        if ($Field -ne '@odata.type') { foreach ($Name in $Family) { if ((Get-ContractMembers $Schema $Name).ContainsKey($Field)) { $Owners += $Name } } }
        @{ Field = $Field; Status = $(if ($Field -eq '@odata.type') { 'ODataAnnotation' } elseif ($Owners.Count) { 'SchemaMember' } else { 'UNVERIFIED' }); DeclaredOn = $Owners }
    })
    $Documentation = 'NotChecked'
    $Scope = 'NotChecked'
    $DocHash = $null
    if (-not $Enabled) { $Documentation = 'Quarantined'; $Scope = 'NotEstablished' }
    elseif ($CheckDocumentation) {
        $DocPath = Join-Path $EvidenceDirectory ($Contract.version + '-' + $Contract.doc + '.md')
        try {
            if (-not (Test-Path -LiteralPath $DocPath)) { Invoke-WebRequest -UseBasicParsing -Uri ($Source + '&accept=text/markdown') -OutFile $DocPath -TimeoutSec 60 }
            $Doc = Get-Content -LiteralPath $DocPath -Raw
            if ($Doc -cnotmatch ('(?m)^source_path: api-reference/' + [regex]::Escape($Contract.version + '/api/' + $Contract.doc + '.md') + '\s*$')) { throw 'Documentation version/source mismatch' }
            $NormalizedPath = $Contract.path -replace '\{id\}', '{id}'
            $Requests = @([regex]::Matches($Doc, '(?m)^GET\s+(?:https://graph.microsoft.com/(?:v1\.0|beta))?/([^\s?]+)') | ForEach-Object { $_.Groups[1].Value -replace '\{[^}]+\}', '{id}' })
            $Documentation = if ($NormalizedPath -cin $Requests) { 'DocumentedGET' } else { 'UNVERIFIED_ROUTE' }
            $Scope = if ($Doc -cmatch ('Delegated \(work or school account\)[^\r\n]*' + [regex]::Escape($Contract.scope))) { 'DocumentedDelegatedRead' } else { 'UNVERIFIED_SCOPE' }
            $DocHash = (Get-FileHash -LiteralPath $DocPath -Algorithm SHA256).Hash
        } catch { $Documentation = 'UNAVAILABLE'; $Scope = 'UNVERIFIED_SCOPE' }
    }
    $Results.Add([ordered]@{ Module = $Contract.module; Enabled = $Enabled; Limitation = (Get-IntuneValue $Contract 'limitation'); Version = $Contract.version; Path = $Contract.path; Type = $TypeName; MetadataRoute = $RouteValid; MetadataSha256 = $Schema.Hash; Documentation = $Documentation; Scope = $Scope; RequiredScope = $Contract.scope; Source = $Source; DocumentationSha256 = $DocHash; Fields = $FieldResults })
}
$ReportPath = Join-Path $EvidenceDirectory 'intune-graph-contract-audit.json'
[IO.File]::WriteAllText($ReportPath, (@{ RetrievedAtUtc = [datetime]::UtcNow.ToString('o'); Contracts = @($Results.ToArray()) } | ConvertTo-Json -Depth 12), [Text.UTF8Encoding]::new($false))
$Failures = @($Results | Where-Object { $_.Enabled -and (-not $_.MetadataRoute -or @($_.Fields | Where-Object Status -eq 'UNVERIFIED').Count -or $_.Documentation -in @('UNAVAILABLE','UNVERIFIED_ROUTE') -or $_.Scope -eq 'UNVERIFIED_SCOPE') })
foreach ($Result in $Failures) { Write-Output ($Result.Module + ': route=' + $Result.MetadataRoute + '; doc=' + $Result.Documentation + '; scope=' + $Result.Scope + '; fields=' + ((@($Result.Fields | Where-Object Status -eq 'UNVERIFIED' | ForEach-Object Field)) -join ',')) }
Write-Output ('Audited {0} route contracts / {1} projected field entries; {2} contracts require review. Full evidence: {3}' -f $Results.Count, (@($Results | ForEach-Object { $_.Fields }).Count), $Failures.Count, $ReportPath)
Write-Output ('Enabled: {0}; quarantined: {1}; documentation checked: {2}. This checks top-level field schema, GET paths and named scopes, not nested values, recommendations, licensing or live compatibility.' -f @($Results | Where-Object Enabled).Count, @($Results | Where-Object { -not $_.Enabled }).Count, [bool]$CheckDocumentation)
if ($Failures.Count) { throw 'Collector contracts require review; do not label the audit passed.' }