# Canonical source: tools/Shared/CollectorPrivacy.ps1. Collector copies must remain identical.

function Resolve-CollectorOutputPath {
    param([string]$Collector, [string]$OutputPath, [string]$CollectionId)
    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        if ([string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) { throw 'LOCALAPPDATA is unavailable; supply -OutputPath.' }
        $Name = $Collector.ToLowerInvariant()
        $OutputPath = Join-Path $env:LOCALAPPDATA ('AssayCollections\{0}\{0}_{1}.json' -f $Name, $CollectionId)
    }
    $FullPath = [IO.Path]::GetFullPath($OutputPath)
    if (Test-Path -LiteralPath $FullPath) { throw 'Output already exists; choose a new output path.' }
    return $FullPath
}

function Test-CollectorSyncedPath {
    param([string]$Path)
    $FullPath = [IO.Path]::GetFullPath($Path)
    foreach ($Root in @($env:OneDrive, $env:OneDriveCommercial, $env:OneDriveConsumer)) {
        if ([string]::IsNullOrWhiteSpace($Root)) { continue }
        $Prefix = [IO.Path]::GetFullPath($Root).TrimEnd('\') + '\'
        if ($FullPath.StartsWith($Prefix, [StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

function Write-CollectorProtectedFile {
    param([string]$Path, [byte[]]$Bytes)
    if ($PSVersionTable.PSEdition -eq 'Core') { Add-Type -AssemblyName System.IO.FileSystem.AccessControl -ErrorAction SilentlyContinue }
    $FullPath = [IO.Path]::GetFullPath($Path)
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
        $ResolvedKeyPath = if ($KeyPath) { [IO.Path]::GetFullPath($KeyPath) } else { [IO.Path]::ChangeExtension([IO.Path]::GetFullPath($OutputPath), '.pseudonym-key') }
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

function ConvertTo-CollectorMaskedPath {
    param($Context, $Path)
    if ($null -eq $Path -or $Path -isnot [string] -or $Context.Mode -eq 'Identified') { return $Path }
    $Masked = $Path -replace '(?i)((?:^|[\\/])(?:users|documents and settings)[\\/])(?!(?:public|default|default user|all users)(?:[\\/]|$))(?![^\\/]*[*?])[^\\/]+', '${1}{profile}'
    return [regex]::Replace($Masked, '^(\\\\|//)([^\\/]+)([\\/])([^\\/]+)', {
        param($Match)
        $UncHost = if ($Match.Groups[2].Value -match '[*?]') { $Match.Groups[2].Value } else { '{host}' }
        $Share = if ($Match.Groups[4].Value -match '[*?]') { $Match.Groups[4].Value } else { '{share}' }
        $Match.Groups[1].Value + $UncHost + $Match.Groups[3].Value + $Share
    })
}

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

function Write-CollectorExport {
    param($Context, [string]$Path, [string]$Json, [string]$IdentityMapPath)
    $Written = Write-CollectorProtectedFile -Path $Path -Bytes ((New-Object Text.UTF8Encoding $false).GetBytes($Json))
    if ($IdentityMapPath -and $Context.Mode -eq 'Pseudonymous') {
        $Map = [ordered]@{ SchemaVersion = '1.0'; PseudonymKeyId = $Context.KeyId; Identities = $Context.Identities }
        [void](Write-CollectorProtectedFile -Path $IdentityMapPath -Bytes ((New-Object Text.UTF8Encoding $false).GetBytes(($Map | ConvertTo-Json -Depth 4))))
    }
    return $Written
}
