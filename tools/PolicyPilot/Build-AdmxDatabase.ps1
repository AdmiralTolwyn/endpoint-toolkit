#Requires -Version 5.1
<#
.SYNOPSIS
    Build-AdmxDatabase.ps1 — Parses ADMX/ADML files into a structured JSON database.
.DESCRIPTION
    Scans Windows PolicyDefinitions (ADMX + ADML) to build a searchable database of
    Group Policy settings with registry keys, categories, descriptions, and allowed values.
    Supports built-in Windows 11 templates plus optional SecGuide/MSS-Legacy ADMX from
    the Microsoft Security Compliance Toolkit.
.PARAMETER AdmxPath
    Path to PolicyDefinitions folder. Default: C:\Windows\PolicyDefinitions
.PARAMETER Language
    ADML language folder. Default: en-US
.PARAMETER OutputPath
    Output JSON path. Default: admx_metadata.json in script directory
.PARAMETER IncludeAll
    Parse ALL ADMX files (241+). Default: only security-relevant subset (~50 files)
.PARAMETER NameLanguages
    Extra ADML languages whose display names and category paths are stored (L10n) so
    gpresult rows from non-English devices can be matched by name. Default: de-DE, fr-FR.
.NOTES
    Author : Anton Romanyuk
    Version: 1.0.0
    Date   : 2026-03-31
#>
[CmdletBinding()]
param(
    [string]$AdmxPath = "$env:SystemRoot\PolicyDefinitions",
    [string]$Language = 'en-US',
    [string]$OutputPath,
    [switch]$IncludeAll,
    [string[]]$NameLanguages = @('de-DE', 'fr-FR')
)

$ErrorActionPreference = 'Continue'
$ScriptVersion = '1.0.0'

$Root = $PSScriptRoot
if (-not $Root) { $Root = $PWD.Path }
if (-not $OutputPath) {
    $OutputPath = Join-Path $Root 'admx_metadata.json'
}

$AdmlPath = Join-Path $AdmxPath $Language

Write-Host ''
Write-Host '  Build-AdmxDatabase v1.0.0' -ForegroundColor Cyan
Write-Host '  Parses ADMX/ADML into structured JSON' -ForegroundColor DarkGray
Write-Host ''

if (-not (Test-Path $AdmxPath)) {
    Write-Host "  ERROR: ADMX path not found: $AdmxPath" -ForegroundColor Red
    exit 1
}

# Security-relevant ADMX files (subset for focused parsing)
$SecurityAdmxFiles = @(
    'WindowsDefender', 'WindowsFirewall', 'DeviceGuard', 'BitLocker', 'FVE',
    'Kerberos', 'CredentialProviders', 'WinRM', 'RemoteDesktopServices',
    'TerminalServer', 'WindowsPowerShell', 'DnsClient', 'Netlogon',
    'LanmanServer', 'LanmanWorkstation', 'tcpip', 'tcpip6',
    'Search', 'WindowsExplorer', 'Biometrics', 'AppxPackageManager',
    'inetres', 'SmartScreen', 'WindowsUpdate', 'CloudContent',
    'DataCollection', 'DigitalLocker', 'EventLog', 'Globalization',
    'NetworkConnections', 'OfflineFiles', 'PeerToPeerCaching',
    'Printing', 'Programs', 'sam', 'Sensors', 'WindowsConnectNow',
    'WCM', 'wlansvc', 'ICM', 'Securitycenter', 'DeviceInstallation',
    'CredUI', 'CtrlAltDel', 'Logon', 'Power', 'UserProfiles',
    'ActiveXInstallService', 'AttachmentManager', 'AutoPlay',
    'MSS-legacy', 'SecGuide', 'PassportForWork'
)

# Determine which files to parse
$AllAdmx = @(Get-ChildItem $AdmxPath -Filter '*.admx' -ErrorAction SilentlyContinue)

# Also scan local templates folder (SecGuide, MSS-legacy from Security Baseline)
$LocalTemplates = Join-Path $Root 'templates'
if (Test-Path $LocalTemplates) {
    $extraAdmx = @(Get-ChildItem $LocalTemplates -Filter '*.admx' -ErrorAction SilentlyContinue)
    if ($extraAdmx.Count -gt 0) {
        Write-Host "  Found $($extraAdmx.Count) extra ADMX in templates/ (SecGuide, MSS-legacy, etc.)" -ForegroundColor Cyan
        $AllAdmx += $extraAdmx
    }
}

if ($IncludeAll) {
    $AdmxFiles = $AllAdmx
    Write-Host "  Parsing ALL $($AdmxFiles.Count) ADMX files" -ForegroundColor White
} else {
    $AdmxFiles = $AllAdmx | Where-Object { $_.BaseName -in $SecurityAdmxFiles }
    Write-Host "  Parsing $($AdmxFiles.Count) security-relevant ADMX files (of $($AllAdmx.Count) total)" -ForegroundColor White
    Write-Host "  Use -IncludeAll to parse everything" -ForegroundColor DarkGray
}
Write-Host ''

# String table cache for resolving $(string.XXX) references
$StringTables = @{}

function Load-StringTable {
    param([string]$AdmlFile)
    if ($StringTables.ContainsKey($AdmlFile)) { return $StringTables[$AdmlFile] }
    $strings = @{}
    if (-not (Test-Path $AdmlFile)) { return $strings }
    try {
        [xml]$adml = Get-Content $AdmlFile -Raw -Encoding UTF8
        $ns = @{ a = 'http://schemas.microsoft.com/GroupPolicy/2006/07/PolicyDefinitions' }
        # Try with namespace
        $stringTable = $adml.policyDefinitionResources.resources.stringTable
        if ($stringTable -and $stringTable.string) {
            foreach ($s in $stringTable.string) {
                if ($s.id) { $strings[$s.id] = $s.'#text' }
            }
        }
    } catch { }
    $StringTables[$AdmlFile] = $strings
    return $strings
}

function Resolve-String {
    param([string]$Ref, [hashtable]$Strings)
    if (-not $Ref) { return '' }
    if ($Ref -match '^\$\(string\.(.+)\)$') {
        $key = $Matches[1]
        if ($Strings.ContainsKey($key)) { return $Strings[$key] }
        return $key
    }
    return $Ref
}

# Category name resolution
# Parents often live in another file (e.g. windows:WindowsComponents), so index categories from all ADMX first
$CategoryNames = @{}
$GlobalCategories = @{}
$XmlCache = @{}

function Get-AdmlFile {
    param($AdmxFile, [string]$Lang = $Language)
    $local = Join-Path (Join-Path $AdmxFile.DirectoryName $Lang) "$($AdmxFile.BaseName).adml"
    if (Test-Path $local) { return $local }
    return (Join-Path (Join-Path $AdmxPath $Lang) "$($AdmxFile.BaseName).adml")
}

function Get-Namespaces {
    param($Admx)
    $map = @{}; $target = ''
    $pn = $Admx.policyDefinitions.policyNamespaces
    if ($pn.target) { $target = $pn.target.namespace; $map[$pn.target.prefix] = $target }
    foreach ($u in @($pn.using)) { if ($u -and $u.prefix) { $map[$u.prefix] = $u.namespace } }
    @{ Target = $target; Map = $map }
}

function Get-QualifiedRef {
    param([string]$Ref, $Ns)
    if (-not $Ref) { return '' }
    if ($Ref.Contains(':')) {
        $prefix, $name = $Ref.Split(':', 2)
        if ($Ns.Map.ContainsKey($prefix)) { return "$($Ns.Map[$prefix]):$name" }
        return $Ref
    }
    return "$($Ns.Target):$Ref"
}

# The XML adapter turns text-only <string> into System.String, so .string.InnerText is always empty
function Get-ValueText {
    param($ValueNode)
    if ($ValueNode -isnot [System.Xml.XmlElement]) { return $null }
    $node = $ValueNode.SelectSingleNode("*[local-name()='decimal' or local-name()='longDecimal']")
    if ($node) { return $node.GetAttribute('value') }
    $node = $ValueNode.SelectSingleNode("*[local-name()='string']")
    if ($node) { return $node.InnerText }
    if ($ValueNode.SelectSingleNode("*[local-name()='delete']")) { return '(value deleted)' }
    return $null
}

foreach ($f in $AllAdmx) {
    try {
        [xml]$x = Get-Content $f.FullName -Raw -Encoding UTF8
        $XmlCache[$f.FullName] = $x
        $ns = Get-Namespaces $x
        $catStrings = Load-StringTable (Get-AdmlFile $f)
        $langStrings = @{}
        foreach ($lang in $NameLanguages) { $langStrings[$lang] = Load-StringTable (Get-AdmlFile $f $lang) }
        foreach ($c in @($x.policyDefinitions.categories.category)) {
            if (-not $c -or -not $c.name) { continue }
            $localNames = @{}
            foreach ($lang in $NameLanguages) {
                if ($c.displayName -match '^\$\(string\.(.+)\)$' -and $langStrings[$lang].ContainsKey($Matches[1])) { $localNames[$lang] = $langStrings[$lang][$Matches[1]] }
            }
            $GlobalCategories[(Get-QualifiedRef $c.name $ns)] = @{
                Name   = Resolve-String $c.displayName $catStrings
                Names  = $localNames
                Parent = Get-QualifiedRef $c.parentCategory.ref $ns
            }
        }
    } catch { }
}

function Resolve-Category {
    param([string]$QualifiedRef, [string]$Lang = '', [string]$Separator = ' > ')
    if (-not $QualifiedRef) { return '' }
    $cacheKey = "$Lang|$QualifiedRef"
    if ($CategoryNames.ContainsKey($cacheKey)) { return $CategoryNames[$cacheKey] }

    # Build path: walk up parentCategory chain
    $path = @()
    $current = $QualifiedRef
    $maxDepth = 10
    while ($current -and $maxDepth -gt 0) {
        $maxDepth--
        $cat = $GlobalCategories[$current]
        $shortName = $current -replace '^.*:', ''
        if (-not $cat) { $path += $shortName; break }
        if ($Lang -and $cat.Names[$Lang]) { $path += $cat.Names[$Lang] }
        elseif ($cat.Name) { $path += $cat.Name } else { $path += $shortName }
        $current = $cat.Parent
    }
    [array]::Reverse($path)
    $fullPath = $path -join $Separator
    $CategoryNames[$cacheKey] = $fullPath
    return $fullPath
}

# Parse all ADMX files
$Database = [ordered]@{}
$TotalPolicies = 0
$ParsedFiles = 0
$Errors = 0

foreach ($admxFile in $AdmxFiles) {
    $baseName = $admxFile.BaseName
    $admlFile = Get-AdmlFile $admxFile

    Write-Host "  [$($ParsedFiles+1)/$($AdmxFiles.Count)] $baseName " -NoNewline

    try {
        $admx = $XmlCache[$admxFile.FullName]
        if (-not $admx) { [xml]$admx = Get-Content $admxFile.FullName -Raw -Encoding UTF8 }
        $strings = Load-StringTable $admlFile
        $namespaces = Get-Namespaces $admx

        $policies = $admx.policyDefinitions.policies.policy

        if (-not $policies -or $policies.Count -eq 0) {
            Write-Host "0 policies" -ForegroundColor DarkGray
            $ParsedFiles++
            continue
        }

        $count = 0
        foreach ($pol in $policies) {
            $name = $pol.name
            if (-not $name) { continue }

            $displayName = Resolve-String $pol.displayName $strings
            $description = Resolve-String $pol.explainText $strings
            if (-not $description) { $description = Resolve-String $pol.displayName $strings }

            $regKey = $pol.key
            $valueName = $pol.valueName
            $class = $pol.class  # Machine, User, Both

            # Category path
            $catRef = Get-QualifiedRef $pol.parentCategory.ref $namespaces
            $catPath = Resolve-Category $catRef

            # Localized name + category path in gpresult's 'A/B' form
            $l10n = [ordered]@{}
            if ($pol.displayName -match '^\$\(string\.(.+)\)$') {
                $nameId = $Matches[1]
                foreach ($lang in $NameLanguages) {
                    $ls = Load-StringTable (Get-AdmlFile $admxFile $lang)
                    if ($ls.ContainsKey($nameId) -and $ls[$nameId] -ne $displayName) {
                        $l10n[$lang] = @($ls[$nameId], (Resolve-Category $catRef -Lang $lang -Separator '/'))
                    }
                }
            }

            # Enabled/Disabled values
            $enabledValue = Get-ValueText $pol.enabledValue
            $disabledValue = Get-ValueText $pol.disabledValue

            # Elements (dropdown, text, numeric sub-settings)
            $elements = @()
            if ($pol.elements) {
                foreach ($elem in $pol.elements.ChildNodes) {
                    if ($elem.NodeType -ne 'Element') { continue }
                    $elemInfo = @{
                        Type      = $elem.LocalName  # decimal, text, enum, boolean, list
                        Id        = $elem.id
                        Key       = if ($elem.key) { $elem.key } else { $regKey }
                        ValueName = $elem.valueName
                    }
                    # For enum elements, extract items
                    if ($elem.LocalName -eq 'enum' -and $elem.item) {
                        $enumItems = [ordered]@{}
                        foreach ($item in $elem.item) {
                            $itemDisplay = Resolve-String $item.displayName $strings
                            $itemValue = Get-ValueText $item.value
                            if ($null -eq $itemValue) { $itemValue = '' }
                            $enumItems["$itemValue"] = $itemDisplay
                        }
                        $elemInfo['EnumValues'] = $enumItems
                    }
                    # For decimal elements, capture min/max
                    if ($elem.LocalName -eq 'decimal') {
                        if ($elem.minValue) { $elemInfo['MinValue'] = $elem.minValue }
                        if ($elem.maxValue) { $elemInfo['MaxValue'] = $elem.maxValue }
                    }
                    $elements += $elemInfo
                }
            }

            # SupportedOn
            $supportedOn = $pol.supportedOn.ref

            $key = "$baseName/$name"
            $entry = [ordered]@{
                Friendly     = $displayName
                Desc         = $description
                Class        = $class
                Category     = $catPath
                RegistryKey  = $regKey
                ValueName    = $valueName
                EnabledValue = $enabledValue
                DisabledValue = $disabledValue
                AdmxFile     = "$baseName.admx"
                SupportedOn  = $supportedOn
            }
            if ($elements.Count -gt 0) { $entry['Elements'] = $elements }
            if ($l10n.Count -gt 0) { $entry['L10n'] = $l10n }

            $Database[$key] = $entry
            $count++
            $TotalPolicies++
        }

        Write-Host "$count policies" -ForegroundColor Green
        $ParsedFiles++
    } catch {
        Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
        $Errors++
    }
}

Write-Host ''
Write-Host "  Parsed: $ParsedFiles files, $TotalPolicies policies, $Errors errors" -ForegroundColor White

# Safety check
if ($TotalPolicies -lt 50) {
    Write-Host "  ERROR: Only $TotalPolicies policies found — aborting (minimum 50 expected)" -ForegroundColor Red
    exit 1
}

# Write output
$Output = [ordered]@{
    _metadata = [ordered]@{
        version       = $ScriptVersion
        generatedAt   = (Get-Date).ToString('o')
        source        = $AdmxPath
        language      = $Language
        totalPolicies = $TotalPolicies
        filesProcessed = $ParsedFiles
        includeAll    = [bool]$IncludeAll
        nameLanguages = @($NameLanguages)
    }
    policies = $Database
}

# Backup existing
if (Test-Path $OutputPath) {
    $bakPath = $OutputPath + '.bak'
    Copy-Item $OutputPath $bakPath -Force
    Write-Host "  Backed up existing: $bakPath" -ForegroundColor DarkGray
}

$jsonStr = $Output | ConvertTo-Json -Depth 10
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($OutputPath, $jsonStr, $utf8Bom)

$sizeMB = [math]::Round((Get-Item $OutputPath).Length / 1MB, 1)
Write-Host ''
Write-Host "  Output: $OutputPath ($sizeMB MB, $TotalPolicies policies)" -ForegroundColor Green
Write-Host ''
