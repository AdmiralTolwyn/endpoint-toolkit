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
    [switch]$IncludeAll
)

$ErrorActionPreference = 'Continue'
$ScriptVersion = '1.0.0'

if (-not $OutputPath) {
    $Root = $PSScriptRoot
    if (-not $Root) { $Root = $PWD.Path }
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
$CategoryNames = @{}

function Resolve-Category {
    param([string]$CatRef, [hashtable]$Strings, $Categories)
    if (-not $CatRef) { return '' }
    if ($CategoryNames.ContainsKey($CatRef)) { return $CategoryNames[$CatRef] }

    # Build path: walk up parentCategory chain
    $path = @()
    $current = $CatRef
    $maxDepth = 10
    while ($current -and $maxDepth -gt 0) {
        $maxDepth--
        $cat = $Categories | Where-Object { $_.name -eq $current } | Select-Object -First 1
        if (-not $cat) { $path += $current; break }
        $displayName = Resolve-String $cat.displayName $Strings
        if ($displayName) { $path += $displayName } else { $path += $current }
        $current = $cat.parentCategory.ref
    }
    [array]::Reverse($path)
    $fullPath = $path -join ' > '
    $CategoryNames[$CatRef] = $fullPath
    return $fullPath
}

# Parse all ADMX files
$Database = [ordered]@{}
$TotalPolicies = 0
$ParsedFiles = 0
$Errors = 0

foreach ($admxFile in $AdmxFiles) {
    $baseName = $admxFile.BaseName
    # Look for ADML next to the ADMX first (local templates), then system PolicyDefinitions
    $admlFile = Join-Path (Join-Path $admxFile.DirectoryName $Language) "$baseName.adml"
    if (-not (Test-Path $admlFile)) {
        $admlFile = Join-Path $AdmlPath "$baseName.adml"
    }

    Write-Host "  [$($ParsedFiles+1)/$($AdmxFiles.Count)] $baseName " -NoNewline

    try {
        [xml]$admx = Get-Content $admxFile.FullName -Raw -Encoding UTF8
        $strings = Load-StringTable $admlFile

        $policies = $admx.policyDefinitions.policies.policy
        $categories = $admx.policyDefinitions.categories.category

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

            # Truncate description to 500 chars
            if ($description.Length -gt 500) { $description = $description.Substring(0, 497) + '...' }

            $regKey = $pol.key
            $valueName = $pol.valueName
            $class = $pol.class  # Machine, User, Both

            # Category path
            $catRef = $pol.parentCategory.ref
            $catPath = Resolve-Category $catRef $strings $categories

            # Enabled/Disabled values
            $enabledValue = $null; $disabledValue = $null
            if ($pol.enabledValue) {
                if ($pol.enabledValue.decimal) { $enabledValue = $pol.enabledValue.decimal.value }
                elseif ($pol.enabledValue.string) { $enabledValue = $pol.enabledValue.string.InnerText }
            }
            if ($pol.disabledValue) {
                if ($pol.disabledValue.decimal) { $disabledValue = $pol.disabledValue.decimal.value }
                elseif ($pol.disabledValue.string) { $disabledValue = $pol.disabledValue.string.InnerText }
            }

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
                        $enumItems = @{}
                        foreach ($item in $elem.item) {
                            $itemDisplay = Resolve-String $item.displayName $strings
                            $itemValue = if ($item.value.decimal) { $item.value.decimal.value }
                                        elseif ($item.value.string) { $item.value.string.InnerText }
                                        else { '' }
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
