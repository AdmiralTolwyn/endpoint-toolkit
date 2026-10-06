<#
.SYNOPSIS
    Shared privacy helpers for the Assay evidence collectors.
.DESCRIPTION
    Provides output-path resolution, protected CreateNew writes, keyed pseudonyms, path masking,
    network classification, error-text minimization and the privacy manifest used by the AVD,
    Windows 365, Intune and Baseline collectors and the WPF assessors.

    Canonical source: tools/Shared/CollectorPrivacy.ps1. The copies beside each tool and the region
    embedded in Invoke-BaselineCollection.ps1 must remain identical; Test-CollectorPrivacy.ps1
    verifies this. Dot-source the file; it has no parameters and performs no collection.
.NOTES
    Author    : Anton Romanyuk
    Version   : 1.0.0
    Date      : 2026-10-06
    Requires  : Windows PowerShell 5.1 or PowerShell 7
    Disclaimer: This script is provided "AS IS" with no warranties and confers no rights.
#>


<#
.SYNOPSIS
    Resolves a path to a full file-system path.
.DESCRIPTION
    Relative paths resolve against the current PowerShell location, not the process working
    directory, which can differ after Set-Location. The target does not need to exist.
.PARAMETER Path
    Absolute or relative path.
.OUTPUTS
    System.String
#>
function Resolve-CollectorFullPath {
    param([string]$Path)
    return $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
}

<#
.SYNOPSIS
    Returns the full path for a new collector export.
.DESCRIPTION
    Without an operator path, uses
    %LOCALAPPDATA%\AssayCollections\<collector>\<collector>_<collectionId>.json.
    Throws when the target already exists; exports are never overwritten.
.PARAMETER Collector
    Collector name used for the default folder and file name, for example 'Baseline'.
.PARAMETER OutputPath
    Operator-supplied path. Empty selects the default location.
.PARAMETER CollectionId
    Collection GUID used in the default file name.
.OUTPUTS
    System.String
#>
function Resolve-CollectorOutputPath {
    param([string]$Collector, [string]$OutputPath, [string]$CollectionId)
    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        if ([string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) { throw 'LOCALAPPDATA is unavailable; supply -OutputPath.' }
        $Name = $Collector.ToLowerInvariant()
        $OutputPath = Join-Path $env:LOCALAPPDATA ('AssayCollections\{0}\{0}_{1}.json' -f $Name, $CollectionId)
    }
    $FullPath = Resolve-CollectorFullPath $OutputPath
    if (Test-Path -LiteralPath $FullPath) { throw 'Output already exists; choose a new output path.' }
    return $FullPath
}

<#
.SYNOPSIS
    Tests whether a path is inside a OneDrive-synchronized folder.
.DESCRIPTION
    Checks the OneDrive, OneDriveCommercial and OneDriveConsumer environment roots. Callers warn;
    they do not block.
.PARAMETER Path
    File path to test.
.OUTPUTS
    System.Boolean
#>
function Test-CollectorSyncedPath {
    param([string]$Path)
    $FullPath = Resolve-CollectorFullPath $Path
    foreach ($Root in @($env:OneDrive, $env:OneDriveCommercial, $env:OneDriveConsumer)) {
        if ([string]::IsNullOrWhiteSpace($Root)) { continue }
        $Prefix = [IO.Path]::GetFullPath($Root).TrimEnd('\') + '\'
        if ($FullPath.StartsWith($Prefix, [StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

<#
.SYNOPSIS
    Creates a new file with a restricted ACL and writes bytes to it.
.DESCRIPTION
    The file is created with CreateNew and a protected ACL that grants full control only to the
    current user, SYSTEM and BUILTIN\Administrators; no inherited ACEs apply. An existing file
    causes an exception. Missing parent folders are created. Works on Windows PowerShell 5.1
    and PowerShell 7.
.PARAMETER Path
    New file path.
.PARAMETER Bytes
    File content.
.OUTPUTS
    System.String. The full path written.
#>
function Write-CollectorProtectedFile {
    param([string]$Path, [byte[]]$Bytes)
    if ($PSVersionTable.PSEdition -eq 'Core') { Add-Type -AssemblyName System.IO.FileSystem.AccessControl -ErrorAction SilentlyContinue }
    $FullPath = Resolve-CollectorFullPath $Path
    $Directory = [IO.Path]::GetDirectoryName($FullPath)
    if (-not [IO.Directory]::Exists($Directory)) { [void][IO.Directory]::CreateDirectory($Directory) }
    $Security = New-Object Security.AccessControl.FileSecurity
    $Security.SetAccessRuleProtection($true, $false)
    $Principals = @(
        [Security.Principal.WindowsIdentity]::GetCurrent().User,
        (New-Object Security.Principal.SecurityIdentifier 'S-1-5-18'),
        (New-Object Security.Principal.SecurityIdentifier 'S-1-5-32-544')
    ) | Sort-Object -Property Value -Unique
    foreach ($Principal in $Principals) {
        $Security.AddAccessRule((New-Object Security.AccessControl.FileSystemAccessRule $Principal, ([Security.AccessControl.FileSystemRights]::FullControl), ([Security.AccessControl.AccessControlType]::Allow)))
    }
    $AclExtensions = 'System.IO.FileSystemAclExtensions' -as [type]
    if ($AclExtensions) {
        $Stream = $AclExtensions::Create((New-Object IO.FileInfo $FullPath), [IO.FileMode]::CreateNew, [Security.AccessControl.FileSystemRights]::FullControl, [IO.FileShare]::None, 4096, [IO.FileOptions]::None, $Security)
    } else {
        $Stream = New-Object IO.FileStream $FullPath, ([IO.FileMode]::CreateNew), ([Security.AccessControl.FileSystemRights]::FullControl), ([IO.FileShare]::None), 4096, ([IO.FileOptions]::None), $Security
    }
    try { $Stream.Write($Bytes, 0, $Bytes.Length) } finally { $Stream.Dispose() }
    return $FullPath
}

<#
.SYNOPSIS
    Creates the privacy context used by the collector transforms.
.DESCRIPTION
    Pseudonymous mode loads the 32-byte Base64 key at KeyPath, or creates a protected key next to
    the output. Identified mode requires ConfirmIdentified and uses no key.
.PARAMETER Mode
    Pseudonymous (default) or Identified.
.PARAMETER ConfirmIdentified
    Must be true for Identified mode.
.PARAMETER KeyPath
    Existing or new pseudonym key file. Empty uses <output>.pseudonym-key.
.PARAMETER OutputPath
    Export path used to place the default key.
.OUTPUTS
    PSCustomObject with Mode, Key, KeyId, KeyPath and Identities.
#>
function New-CollectorPrivacyContext {
    param(
        [ValidateSet('Pseudonymous', 'Identified')][string]$Mode = 'Pseudonymous',
        [bool]$ConfirmIdentified,
        [string]$KeyPath,
        [string]$OutputPath
    )
    if ($Mode -eq 'Identified' -and -not $ConfirmIdentified) { throw 'Identified mode requires -ConfirmIdentifiedExport.' }
    $Key = $null
    $KeyId = $null
    $ResolvedKeyPath = $null
    if ($Mode -eq 'Pseudonymous') {
        $ResolvedKeyPath = if ($KeyPath) { Resolve-CollectorFullPath $KeyPath } else { [IO.Path]::ChangeExtension((Resolve-CollectorFullPath $OutputPath), '.pseudonym-key') }
        if (Test-Path -LiteralPath $ResolvedKeyPath) {
            try { $Key = [Convert]::FromBase64String(([IO.File]::ReadAllText($ResolvedKeyPath)).Trim()) } catch { throw 'Pseudonym key file is not valid Base64.' }
            if ($Key.Length -ne 32) { throw 'Pseudonym key must contain exactly 32 bytes.' }
        } else {
            $Key = New-Object byte[] 32
            $Random = [Security.Cryptography.RandomNumberGenerator]::Create()
            try { $Random.GetBytes($Key) } finally { $Random.Dispose() }
            [void](Write-CollectorProtectedFile -Path $ResolvedKeyPath -Bytes ([Text.Encoding]::ASCII.GetBytes([Convert]::ToBase64String($Key))))
        }
        $Hash = [Security.Cryptography.SHA256]::Create()
        try { $KeyId = -join ($Hash.ComputeHash($Key)[0..7] | ForEach-Object { $_.ToString('x2') }) } finally { $Hash.Dispose() }
    }
    return [pscustomobject]@{
        Mode = $Mode
        Key = $Key
        KeyId = $KeyId
        KeyPath = $ResolvedKeyPath
        Identities = New-Object 'Collections.Generic.Dictionary[string,string]'
    }
}

<#
.SYNOPSIS
    Returns a keyed pseudonym for a person, device, SID or group identifier.
.DESCRIPTION
    The pseudonym is the prefix, an underscore and the first 16 hex digits of
    HMAC-SHA256(key, trimmed lowercase value). The original is recorded in the context identity map.
    Identified mode, null and blank values are returned unchanged.
.PARAMETER Context
    Privacy context from New-CollectorPrivacyContext.
.PARAMETER Prefix
    usr, dev, sid or grp.
.PARAMETER Value
    Identifier to transform.
.OUTPUTS
    System.String
#>
function ConvertTo-CollectorIdentity {
    param($Context, [string]$Prefix, $Value)
    if ($null -eq $Value) { return $null }
    $Text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($Text)) { return $Text }
    if ($Context.Mode -eq 'Identified') { return $Text }
    $Normalized = $Text.Trim().ToLowerInvariant()
    $Hmac = New-Object Security.Cryptography.HMACSHA256 (, $Context.Key)
    try { $Digest = $Hmac.ComputeHash([Text.Encoding]::UTF8.GetBytes($Normalized)) } finally { $Hmac.Dispose() }
    $Pseudonym = $Prefix + '_' + (-join ($Digest[0..7] | ForEach-Object { $_.ToString('x2') }))
    if (-not $Context.Identities.ContainsKey($Pseudonym)) { $Context.Identities[$Pseudonym] = $Text.Trim() }
    return $Pseudonym
}

<#
.SYNOPSIS
    Masks user-profile and UNC host/share segments in a path.
.DESCRIPTION
    C:\Users\<name>\ becomes C:\Users\{profile}\ and \\host\share becomes \\{host}\{share}.
    Wildcards, the remainder of the path, \\?\ and \\.\ prefixes and empty values are preserved.
    Identified mode and non-string values are returned unchanged.
.PARAMETER Context
    Privacy context from New-CollectorPrivacyContext.
.PARAMETER Path
    Path or path pattern to mask.
.OUTPUTS
    System.String
#>
function ConvertTo-CollectorMaskedPath {
    param($Context, $Path)
    if ($null -eq $Path -or $Path -isnot [string] -or $Context.Mode -eq 'Identified') { return $Path }
    $Masked = $Path -replace '(?i)((?:^|[\\/])(?:users|documents and settings)[\\/])(?!(?:public|default|default user|all users)(?:[\\/]|$))(?![^\\/]*[*?])[^\\/]+', '${1}{profile}'
    $Prefix = [regex]::Match($Masked, '^(?:[\\/]{2}[?.][\\/](?:(?i:UNC)[\\/])?|[\\/]{2})')
    if (-not $Prefix.Success -or ($Prefix.Length -gt 2 -and $Prefix.Value -notmatch '(?i)UNC[\\/]$')) { return $Masked }
    $Rest = [regex]::Replace($Masked.Substring($Prefix.Length), '^([^\\/]+)(?:([\\/])([^\\/]*))?', {
        param($Match)
        $UncHost = if ($Match.Groups[1].Value -match '[*?]') { $Match.Groups[1].Value } else { '{host}' }
        if (-not $Match.Groups[2].Success) { return $UncHost }
        $Share = if ($Match.Groups[3].Value -eq '' -or $Match.Groups[3].Value -match '[*?]') { $Match.Groups[3].Value } else { '{share}' }
        $UncHost + $Match.Groups[2].Value + $Share
    })
    return $Prefix.Value + $Rest
}

<#
.SYNOPSIS
    Classifies public IP addresses and prefixes.
.DESCRIPTION
    Private, loopback, link-local, shared-address (100.64/10), unique-local and default-route
    values are kept. Public addresses become Public/<prefix length>. Service tags and other
    non-IP values, Identified mode and non-string values are returned unchanged.
.PARAMETER Context
    Privacy context from New-CollectorPrivacyContext.
.PARAMETER Value
    Address, CIDR prefix or service tag.
.OUTPUTS
    System.String
#>
function ConvertTo-CollectorNetworkValue {
    param($Context, $Value)
    if ($null -eq $Value -or $Value -isnot [string] -or $Context.Mode -eq 'Identified') { return $Value }
    $Parts = $Value.Trim().Split('/')
    $Address = $null
    if ($Parts.Count -gt 2 -or -not [Net.IPAddress]::TryParse($Parts[0], [ref]$Address)) { return $Value }
    $Length = if ($Parts.Count -eq 2) { $Parts[1] } elseif ($Address.AddressFamily -eq [Net.Sockets.AddressFamily]::InterNetwork) { '32' } else { '128' }
    $Bytes = $Address.GetAddressBytes()
    $Internal = if ($Address.AddressFamily -eq [Net.Sockets.AddressFamily]::InterNetwork) {
        $Bytes[0] -eq 10 -or $Bytes[0] -eq 127 -or ($Bytes[0] -eq 172 -and $Bytes[1] -ge 16 -and $Bytes[1] -le 31) -or
        ($Bytes[0] -eq 192 -and $Bytes[1] -eq 168) -or ($Bytes[0] -eq 169 -and $Bytes[1] -eq 254) -or
        ($Bytes[0] -eq 100 -and $Bytes[1] -ge 64 -and $Bytes[1] -le 127) -or ($Value.Trim() -eq '0.0.0.0/0')
    } else {
        $Address.IsIPv6LinkLocal -or $Address.Equals([Net.IPAddress]::IPv6Loopback) -or (($Bytes[0] -band 0xFE) -eq 0xFC) -or ($Value.Trim() -eq '::/0')
    }
    if ($Internal) { return $Value }
    return 'Public/' + $Length
}

<#
.SYNOPSIS
    Returns error text that is safe to export.
.DESCRIPTION
    Pseudonymous mode keeps only the exception type, HTTP status and service error code. Plain
    strings are kept only when they are short identifiers; otherwise 'CollectionError' is used.
    Identified mode returns the full message.
.PARAMETER Context
    Privacy context from New-CollectorPrivacyContext.
.PARAMETER ErrorInput
    ErrorRecord, Exception or string.
.OUTPUTS
    System.String
#>
function Get-CollectorErrorText {
    param($Context, $ErrorInput)
    $Exception = if ($ErrorInput -is [Management.Automation.ErrorRecord]) { $ErrorInput.Exception } elseif ($ErrorInput -is [Exception]) { $ErrorInput } else { $null }
    if ($Context.Mode -eq 'Identified') { if ($Exception) { return $Exception.Message } return [string]$ErrorInput }
    if (-not $Exception) {
        $Text = [string]$ErrorInput
        if ($Text -cmatch '^[A-Za-z0-9_.:-]{1,64}$') { return $Text }
        return 'CollectionError'
    }
    $Parts = @($Exception.GetType().Name)
    $Status = $null
    if ($Exception.PSObject.Properties['Response'] -and $Exception.Response -and $Exception.Response.PSObject.Properties['StatusCode']) { $Status = [int]$Exception.Response.StatusCode }
    elseif ($Exception.Message -match '(?i)(?:status code\D{0,40}|\bHTTP\s*:?\s*)(\d{3})\b') { $Status = [int]$Matches[1] }
    if ($Status) { $Parts += "HTTP $Status" }
    if ($Exception.Message -match '"code"\s*:\s*"([A-Za-z0-9_.]{1,64})"') { $Parts += 'code ' + $Matches[1] }
    elseif ($Exception.Message -match '(?i)\b(Authorization_RequestDenied|AuthorizationFailed|Forbidden|Unauthorized|ResourceNotFound|NotFound|BadRequest|TooManyRequests|InvalidAuthenticationToken|AccessDenied)\b') { $Parts += 'code ' + $Matches[1] }
    return ($Parts -join '; ')
}

<#
.SYNOPSIS
    Builds the top-level Privacy manifest for an export.
.PARAMETER Context
    Privacy context from New-CollectorPrivacyContext.
.PARAMETER OptIns
    Names of opt-in switches that widened collection.
.OUTPUTS
    System.Collections.Specialized.OrderedDictionary
#>
function New-CollectorPrivacyManifest {
    param($Context, [string[]]$OptIns = @())
    $Pseudonymous = $Context.Mode -eq 'Pseudonymous'
    return [ordered]@{
        SchemaVersion = '1.0'
        Mode = $Context.Mode
        PseudonymKeyId = $Context.KeyId
        Classification = 'Confidential'
        OptIns = @($OptIns | Where-Object { $_ } | Sort-Object -Unique)
        RemovedFieldClasses = $(if ($Pseudonymous) { @('Secret', 'Content', 'FreeText') } else { @('Secret', 'Content') })
        PseudonymizedFieldClasses = $(if ($Pseudonymous) { @('Person') } else { @() })
        ClassifiedFieldClasses = $(if ($Pseudonymous) { @('Network') } else { @() })
    }
}

<#
.SYNOPSIS
    Writes the export JSON and optional identity map as protected new files.
.PARAMETER Context
    Privacy context from New-CollectorPrivacyContext.
.PARAMETER Path
    New export path.
.PARAMETER Json
    Export content, written as UTF-8 without a BOM.
.PARAMETER IdentityMapPath
    Optional new pseudonym-to-original map. Written only in Pseudonymous mode.
.OUTPUTS
    System.String. The full export path.
#>
function Write-CollectorExport {
    param($Context, [string]$Path, [string]$Json, [string]$IdentityMapPath)
    $Written = Write-CollectorProtectedFile -Path $Path -Bytes ((New-Object Text.UTF8Encoding $false).GetBytes($Json))
    if ($IdentityMapPath -and $Context.Mode -eq 'Pseudonymous') {
        $Map = [ordered]@{ SchemaVersion = '1.0'; PseudonymKeyId = $Context.KeyId; Identities = $Context.Identities }
        [void](Write-CollectorProtectedFile -Path $IdentityMapPath -Bytes ((New-Object Text.UTF8Encoding $false).GetBytes(($Map | ConvertTo-Json -Depth 4))))
    }
    return $Written
}
