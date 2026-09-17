function Read-IntunePolicyXml {
    param([string]$Text, [switch]$Fragment)
    if ($Text.Length -gt 1MB) { throw 'Policy XML limit' }
    $Settings = [Xml.XmlReaderSettings]::new()
    $Settings.DtdProcessing = [Xml.DtdProcessing]::Prohibit
    $Settings.XmlResolver = $null
    $Settings.MaxCharactersInDocument = 2MB
    $InputText = if ($Fragment) { '<root>' + $Text + '</root>' } else { $Text }
    $TextReader = [IO.StringReader]::new($InputText)
    $Reader = [Xml.XmlReader]::Create($TextReader, $Settings)
    try {
        $Document = [Xml.XmlDocument]::new()
        $Document.XmlResolver = $null
        $Document.Load($Reader)
        return ,$Document
    } finally { $Reader.Dispose(); $TextReader.Dispose() }
}

function ConvertTo-IntuneAdmxMetadata {
    param([string]$Text)
    $Document = Read-IntunePolicyXml $Text -Fragment
    $Enabled = @($Document.SelectNodes('/root/enabled'))
    $Disabled = @($Document.SelectNodes('/root/disabled'))
    if ($Enabled.Count + $Disabled.Count -ne 1) { throw 'Ambiguous ADMX enablement' }
    $Data = [ordered]@{}
    foreach ($Node in $Document.SelectNodes('/root/data')) {
        $Identifier = $Node.GetAttribute('id')
        $Number = [long]0
        if ($Identifier.Length -gt 128 -or -not $Identifier -or $Data.Contains($Identifier)) { throw 'Invalid ADMX data identity' }
        if ([long]::TryParse($Node.GetAttribute('value'), [ref]$Number)) { $Data[$Identifier] = $Number }
    }
    return @{ enabled = ($Enabled.Count -eq 1); data = $Data }
}

function ConvertTo-IntuneAppControlMetadata {
    param([string]$Text)
    $Document = Read-IntunePolicyXml $Text
    $Namespace = [Xml.XmlNamespaceManager]::new($Document.NameTable)
    $Namespace.AddNamespace('ci', 'urn:schemas-microsoft-com:sipolicy')
    $Root = $Document.SelectSingleNode('/ci:SiPolicy', $Namespace)
    if ($null -eq $Root) { throw 'Unsupported App Control XML namespace' }
    $IdentityNode = $Document.SelectSingleNode('/ci:SiPolicy/ci:PolicyID', $Namespace)
    $BaseNode = $Document.SelectSingleNode('/ci:SiPolicy/ci:BasePolicyID', $Namespace)
    $Identity = [guid]::Empty
    $Base = [guid]::Empty
    if ($null -eq $IdentityNode -or $null -eq $BaseNode -or -not [guid]::TryParse($IdentityNode.InnerText, [ref]$Identity) -or -not [guid]::TryParse($BaseNode.InnerText, [ref]$Base) -or $Identity -eq [guid]::Empty -or $Base -eq [guid]::Empty) { throw 'Missing App Control policy identity' }
    $Options = @($Document.SelectNodes('/ci:SiPolicy/ci:Rules/ci:Rule/ci:Option', $Namespace) | ForEach-Object { $_.InnerText } | Where-Object { $_ -in @('Enabled:Audit Mode', 'Enabled:Managed Installer', 'Enabled:Intelligent Security Graph Authorization', 'Enabled:Allow Supplemental Policies', 'Enabled:UMCI') })
    return @{ id = $Identity.ToString(); basePolicyId = $Base.ToString(); policyType = $Root.GetAttribute('PolicyType'); auditMode = ($Options -contains 'Enabled:Audit Mode'); managedInstaller = ($Options -contains 'Enabled:Managed Installer'); options = $Options; source = 'Explicit XML file; assignment and runtime not established' }
}

function Test-IntuneStructuredPath {
    param([string]$Path)
    return $Path -ceq 'Policy/Config/InternetExplorer/DisableInternetExplorerLaunchViaCOM' -or $Path -cmatch '^BitLocker/(SystemDrivesRequireStartupAuthentication|SystemDrivesRecoveryOptions|FixedDrivesRecoveryOptions)$'
}

function Test-IntuneExtendedPath {
    param([string]$Path)
    return (Test-IntuneStructuredPath $Path) -or $Path -cmatch '^Policy/Config/Defender/Excluded(Paths|Processes|Extensions)$' -or $Path -cmatch '^Firewall/MdmStore/FirewallRules/[^/|]+/(Protocol|Direction|LocalPortRanges|RemotePortRanges|EdgeTraversal|InterfaceTypes|Action/Type|Enabled)$'
}