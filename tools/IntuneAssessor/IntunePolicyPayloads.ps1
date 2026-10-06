<#
.SYNOPSIS
    Safe parsing of ADMX setting fragments and App Control policy XML.
.DESCRIPTION
    Dot-sourced by IntuneExpansion.ps1. XML is parsed with DTD processing prohibited, no resolver and
    size limits. Only reviewed data IDs and metadata are returned.
.NOTES
    Author    : Anton Romanyuk
    Requires  : Windows PowerShell 5.1 or PowerShell 7
    Disclaimer: This script is provided "AS IS" with no warranties and confers no rights.
#>


<#
.SYNOPSIS
    Parses policy XML with DTDs and external resolution disabled.
.PARAMETER Text
    XML text, at most 1 MB.
.PARAMETER Fragment
    Wrap the text in a root element before parsing.
.OUTPUTS
    System.Xml.XmlDocument
#>
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

<#
.SYNOPSIS
    Returns the reviewed ADMX data IDs and value types for a CSP path.
.PARAMETER CspPath
    CSP path of an ADMX-backed setting.
.OUTPUTS
    Hashtable of data ID to type. Unreviewed paths throw.
#>
function Get-IntuneAdmxDataContract {
    param([string]$CspPath)
    switch -CaseSensitive ($CspPath) {
        'Policy/Config/InternetExplorer/DisableInternetExplorerLaunchViaCOM' { return @{} }
        'BitLocker/SystemDrivesRequireStartupAuthentication' {
            return @{
                ConfigureNonTPMStartupKeyUsage_Name = 'Boolean'
                ConfigureTPMStartupKeyUsageDropDown_Name = 'Usage'
                ConfigurePINUsageDropDown_Name = 'Usage'
                ConfigureTPMPINKeyUsageDropDown_Name = 'Usage'
                ConfigureTPMUsageDropDown_Name = 'Usage'
            }
        }
        { $_ -cin @('BitLocker/SystemDrivesRecoveryOptions', 'BitLocker/FixedDrivesRecoveryOptions') } {
            $Prefix = if ($CspPath -ceq 'BitLocker/SystemDrivesRecoveryOptions') { 'OS' } else { 'FDV' }
            $Contract = @{}
            foreach ($Suffix in @('AllowDRA_Name', 'HideRecoveryPage_Name', 'ActiveDirectoryBackup_Name', 'RequireActiveDirectoryBackup_Name')) { $Contract[$Prefix + $Suffix] = 'Boolean' }
            foreach ($Suffix in @('RecoveryPasswordUsageDropDown_Name', 'RecoveryKeyUsageDropDown_Name')) { $Contract[$Prefix + $Suffix] = 'Usage' }
            $Contract[$Prefix + 'ActiveDirectoryBackupDropDown_Name'] = 'Backup'
            return $Contract
        }
        default { throw 'Unreviewed ADMX policy' }
    }
}

<#
.SYNOPSIS
    Decodes an ADMX setting fragment into its state and typed data values.
.PARAMETER Text
    ADMX fragment such as <enabled/><data id="..." value="..."/>.
.PARAMETER CspPath
    CSP path that selects the data contract.
.OUTPUTS
    Hashtable with enabled (bool) and data (ordered dictionary). Unknown IDs or values throw.
#>
function ConvertTo-IntuneAdmxMetadata {
    param([string]$Text, [string]$CspPath)
    $Document = Read-IntunePolicyXml $Text -Fragment
    $State = $null
    $Data = [ordered]@{}
    foreach ($Node in $Document.DocumentElement.ChildNodes) {
        if ($Node.NodeType -in @([Xml.XmlNodeType]::Whitespace, [Xml.XmlNodeType]::SignificantWhitespace, [Xml.XmlNodeType]::Comment)) { continue }
        if ($Node.NodeType -ne [Xml.XmlNodeType]::Element -or $Node.NamespaceURI -or $Node.HasChildNodes) { throw 'Unsupported ADMX fragment structure' }
        if ($Node.Name -cin @('enabled', 'Enabled', 'disabled', 'Disabled')) {
            if ($null -ne $State -or $Node.Attributes.Count -ne 0) { throw 'Ambiguous ADMX enablement' }
            $State = $Node.Name -ieq 'enabled'
        } elseif ($Node.Name -cin @('data', 'Data')) {
            if ($Node.Attributes.Count -ne 2 -or -not $Node.HasAttribute('id') -or -not $Node.HasAttribute('value')) { throw 'Unsupported ADMX data attributes' }
            $Identifier = $Node.GetAttribute('id')
            $Number = [long]0
            if ($Identifier.Length -gt 128 -or -not $Identifier -or $Identifier -match '[\x00-\x1F]' -or $Data.Contains($Identifier) -or $Data.Count -ge 100) { throw 'Invalid ADMX data identity' }
            $RawValue = $Node.GetAttribute('value')
            if ($RawValue -ceq 'true') { $Data[$Identifier] = $true }
            elseif ($RawValue -ceq 'false') { $Data[$Identifier] = $false }
            elseif ([long]::TryParse($RawValue, [Globalization.NumberStyles]::AllowLeadingSign, [Globalization.CultureInfo]::InvariantCulture, [ref]$Number)) { $Data[$Identifier] = $Number }
            else { throw 'Unreviewed ADMX value' }
        } else { throw 'Unsupported ADMX element' }
    }
    if ($null -eq $State -or (-not $State -and $Data.Count)) { throw 'Unsupported ADMX state/data combination' }
    if ($CspPath) {
        $Contract = Get-IntuneAdmxDataContract $CspPath
        if ($State) {
            if ($Data.Count -ne $Contract.Count) { throw 'Incomplete ADMX policy data' }
            foreach ($Key in $Data.Keys) {
                if ($Key -cnotin @($Contract.Keys)) { throw 'Unreviewed ADMX data ID' }
                $Value = $Data[$Key]
                switch ($Contract[$Key]) {
                    'Boolean' { if ($Value -isnot [bool]) { throw 'Expected Boolean ADMX value' } }
                    'Usage' { if ($Value -isnot [long] -or $Value -lt 0 -or $Value -gt 2) { throw 'Unknown ADMX usage enum' } }
                    'Backup' { if ($Value -isnot [long] -or $Value -notin @(1,2)) { throw 'Unknown ADMX backup enum' } }
                }
            }
        }
    }
    return @{ enabled = $State; data = $Data }
}

<#
.SYNOPSIS
    Extracts identity and option metadata from an App Control (WDAC) policy.
.PARAMETER Text
    SiPolicy XML text, at most 1 MB.
.OUTPUTS
    Hashtable with id, basePolicyId, policyType, auditMode and related option metadata; rule contents are not exported.
#>
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

<#
.SYNOPSIS
    Tests whether a CSP path holds an ADMX fragment decoded by ConvertTo-IntuneAdmxMetadata.
.PARAMETER Path
    CSP path.
.OUTPUTS
    System.Boolean
#>
function Test-IntuneStructuredPath {
    param([string]$Path)
    return $Path -ceq 'Policy/Config/InternetExplorer/DisableInternetExplorerLaunchViaCOM' -or $Path -cmatch '^BitLocker/(SystemDrivesRequireStartupAuthentication|SystemDrivesRecoveryOptions|FixedDrivesRecoveryOptions)$'
}

<#
.SYNOPSIS
    Tests whether a CSP path is in the extended setting set (structured, Defender exclusions, firewall rules).
.PARAMETER Path
    CSP path.
.OUTPUTS
    System.Boolean
#>
function Test-IntuneExtendedPath {
    param([string]$Path)
    return (Test-IntuneStructuredPath $Path) -or $Path -cmatch '^Policy/Config/Defender/Excluded(Paths|Processes|Extensions)$' -or $Path -cmatch '^Firewall/MdmStore/FirewallRules/[^/|]+/(Protocol|Direction|LocalPortRanges|RemotePortRanges|EdgeTraversal|InterfaceTypes|Action/Type|Enabled)$'
}