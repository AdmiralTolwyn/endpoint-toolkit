#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
$Canonical = Join-Path $PSScriptRoot 'CollectorPrivacy.ps1'
. $Canonical
function Assert-Privacy([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }

$Expected = [IO.File]::ReadAllText($Canonical).Replace("`r`n", "`n").TrimEnd("`n")
foreach ($Copy in @('AvdAssessor', 'W365Assessor', 'IntuneAssessor', 'BaselineAssessor')) {
    $Text = [IO.File]::ReadAllText((Join-Path $PSScriptRoot "..\$Copy\CollectorPrivacy.ps1")).Replace("`r`n", "`n").TrimEnd("`n")
    Assert-Privacy ($Text -ceq $Expected) "$Copy privacy helper differs from the canonical copy"
}
$Baseline = [IO.File]::ReadAllText((Join-Path $PSScriptRoot '..\BaselineAssessor\Invoke-BaselineCollection.ps1')).Replace("`r`n", "`n")
$Region = [regex]::Match($Baseline, '(?s)# region CollectorPrivacy\n(.*?)\n# endregion CollectorPrivacy')
Assert-Privacy ($Region.Success -and $Region.Groups[1].Value -ceq $Expected) 'Embedded Baseline privacy helper differs from the canonical copy'
Write-Output 'PASS: collector helper copies and the embedded Baseline region match the canonical source.'

$Root = Join-Path ([IO.Path]::GetTempPath()) ('collector-privacy-' + [guid]::NewGuid())
try {
    $Output = Join-Path $Root 'nested\export.json'
    $Rejected = $false
    try { New-CollectorPrivacyContext -Mode Identified -ConfirmIdentified $false -OutputPath $Output | Out-Null } catch { $Rejected = $true }
    Assert-Privacy $Rejected 'Identified mode accepted without confirmation'
    $Identified = New-CollectorPrivacyContext -Mode Identified -ConfirmIdentified $true -OutputPath $Output
    Assert-Privacy ($null -eq $Identified.Key -and $null -eq $Identified.KeyId -and -not (Test-Path ([IO.Path]::ChangeExtension($Output, '.pseudonym-key')))) 'Identified mode created a pseudonym key'
    Assert-Privacy ((ConvertTo-CollectorIdentity $Identified 'usr' 'Alice@Contoso.com') -ceq 'Alice@Contoso.com') 'Identified mode altered identity'

    $Context = New-CollectorPrivacyContext -OutputPath $Output
    $KeyPath = [IO.Path]::ChangeExtension($Output, '.pseudonym-key')
    Assert-Privacy ((Test-Path $KeyPath) -and $Context.KeyPath -eq $KeyPath -and $Context.KeyId -cmatch '^[0-9a-f]{16}$') 'Default pseudonym key not created beside the output'
    $First = ConvertTo-CollectorIdentity $Context 'usr' ' Alice@Contoso.com '
    Assert-Privacy ($First -cmatch '^usr_[0-9a-f]{16}$' -and $First -ceq (ConvertTo-CollectorIdentity $Context 'usr' 'alice@contoso.com')) 'Pseudonym format or normalization incorrect'
    Assert-Privacy ($Context.Identities[$First] -ceq 'Alice@Contoso.com') 'Identity map lost the original value'
    $Reused = New-CollectorPrivacyContext -KeyPath $KeyPath -OutputPath (Join-Path $Root 'second.json')
    Assert-Privacy ($Reused.KeyId -ceq $Context.KeyId -and (ConvertTo-CollectorIdentity $Reused 'usr' 'alice@contoso.com') -ceq $First) 'Reused key produced different pseudonyms'
    [IO.File]::WriteAllText((Join-Path $Root 'bad.key'), 'not-base64')
    $Rejected = $false
    try { New-CollectorPrivacyContext -KeyPath (Join-Path $Root 'bad.key') -OutputPath $Output | Out-Null } catch { $Rejected = $true }
    Assert-Privacy $Rejected 'Invalid key accepted'
    [IO.File]::WriteAllText((Join-Path $Root 'short.key'), [Convert]::ToBase64String([byte[]](1..16)))
    $Rejected = $false
    try { New-CollectorPrivacyContext -KeyPath (Join-Path $Root 'short.key') -OutputPath $Output | Out-Null } catch { $Rejected = $true }
    Assert-Privacy $Rejected 'Short key accepted'
    Assert-Privacy ((ConvertTo-CollectorIdentity $Context 'usr' '') -ceq '' -and $null -eq (ConvertTo-CollectorIdentity $Context 'usr' $null)) 'Empty identities must be preserved'

    $Map = Join-Path $Root 'identities.json'
    $Written = Write-CollectorExport $Context $Output '{"ok":true}' $Map
    Assert-Privacy ($Written -eq $Output -and [IO.File]::ReadAllText($Output) -ceq '{"ok":true}') 'Export not written exactly'
    Assert-Privacy ([IO.File]::ReadAllBytes($Output)[0] -eq [byte][char]'{') 'Export must not contain a BOM'
    Assert-Privacy ((Get-Content $Map -Raw | ConvertFrom-Json).Identities.$First -ceq 'Alice@Contoso.com') 'Identity map content incorrect'
    foreach ($Protected in @($Output, $KeyPath, $Map)) {
        $Acl = Get-Acl -LiteralPath $Protected
        $Rules = @($Acl.GetAccessRules($true, $true, [Security.Principal.SecurityIdentifier]))
        $Allowed = @([Security.Principal.WindowsIdentity]::GetCurrent().User.Value, 'S-1-5-18', 'S-1-5-32-544')
        Assert-Privacy ($Acl.AreAccessRulesProtected -and @($Rules | Where-Object IsInherited).Count -eq 0) "Inherited ACEs remain on $Protected"
        Assert-Privacy (@($Rules | Where-Object { $_.IdentityReference.Value -notin $Allowed }).Count -eq 0) "Unexpected principal on $Protected"
    }
    $Rejected = $false
    try { Write-CollectorProtectedFile -Path $Output -Bytes ([byte[]](1)) | Out-Null } catch { $Rejected = $true }
    Assert-Privacy ($Rejected -and [IO.File]::ReadAllText($Output) -ceq '{"ok":true}') 'Existing export overwritten'
    $Rejected = $false
    try { Resolve-CollectorOutputPath 'W365' $Output 'id' | Out-Null } catch { $Rejected = $true }
    Assert-Privacy $Rejected 'Existing output path accepted'
    $SavedLocal = $env:LOCALAPPDATA
    try {
        $env:LOCALAPPDATA = $Root
        Assert-Privacy ((Resolve-CollectorOutputPath 'W365' '' 'abc') -eq [IO.Path]::GetFullPath((Join-Path $Root 'AssayCollections\w365\w365_abc.json'))) 'Default output path incorrect'
    } finally { $env:LOCALAPPDATA = $SavedLocal }
    $SavedOneDrive = $env:OneDrive
    try {
        $env:OneDrive = $Root
        Assert-Privacy ((Test-CollectorSyncedPath (Join-Path $Root 'x.json')) -and -not (Test-CollectorSyncedPath ($Root + 'x\y.json'))) 'OneDrive detection incorrect'
    } finally { $env:OneDrive = $SavedOneDrive }
    Write-Output 'PASS: confirmation gate, keyed pseudonyms, key reuse/validation, identity map, protected no-overwrite files and default paths.'

    $Paths = [ordered]@{
        'C:\Users\Alice\Source' = 'C:\Users\{profile}\Source'
        'C:\Users\Public\Tools' = 'C:\Users\Public\Tools'
        'C:\Users\*\AppData\Local\Temp' = 'C:\Users\*\AppData\Local\Temp'
        '%USERPROFILE%\Downloads' = '%USERPROFILE%\Downloads'
        '\\fileserver\finance\app.exe' = '\\{host}\{share}\app.exe'
        '\\*\share\x' = '\\*\{share}\x'
        'D:\Builds\*.exe' = 'D:\Builds\*.exe'
        'C:\Documents and Settings\Bob\x' = 'C:\Documents and Settings\{profile}\x'
        '\\?\C:\Users\Alice\tool.exe' = '\\?\C:\Users\{profile}\tool.exe'
        '\\.\PhysicalDrive0' = '\\.\PhysicalDrive0'
        '\\?\UNC\fileserver\finance\app.exe' = '\\?\UNC\{host}\{share}\app.exe'
        '\\fileserver' = '\\{host}'
        '//fileserver/share/x' = '//{host}/{share}/x'
        '' = ''
    }
    foreach ($Entry in $Paths.GetEnumerator()) {
        $Actual = ConvertTo-CollectorMaskedPath $Context $Entry.Key
        Assert-Privacy ($Actual -ceq $Entry.Value) "Path mask '$($Entry.Key)' gave '$Actual'"
    }
    Assert-Privacy ((ConvertTo-CollectorMaskedPath $Identified 'C:\Users\Alice') -ceq 'C:\Users\Alice' -and (ConvertTo-CollectorMaskedPath $Context 5) -eq 5) 'Path masking mode or type handling incorrect'
    $Networks = [ordered]@{
        '203.0.113.77' = 'Public/32'; '198.51.100.0/24' = 'Public/24'; '10.1.0.0/16' = '10.1.0.0/16'; '172.31.0.0/16' = '172.31.0.0/16'
        '192.168.1.5' = '192.168.1.5'; '100.64.0.1' = '100.64.0.1'; '0.0.0.0/0' = '0.0.0.0/0'; '*' = '*'; 'Internet' = 'Internet'
        'AzureCloud.WestEurope' = 'AzureCloud.WestEurope'; '2001:db8::1/128' = 'Public/128'; 'fd00::/8' = 'fd00::/8'; '::/0' = '::/0'; '172.32.0.0/16' = 'Public/16'
    }
    foreach ($Entry in $Networks.GetEnumerator()) {
        $Actual = ConvertTo-CollectorNetworkValue $Context $Entry.Key
        Assert-Privacy ($Actual -ceq $Entry.Value) "Network '$($Entry.Key)' gave '$Actual'"
    }
    Assert-Privacy ((ConvertTo-CollectorNetworkValue $Identified '203.0.113.77') -ceq '203.0.113.77') 'Identified network value altered'
    Write-Output 'PASS: profile/UNC path masking and public network classification preserve wildcards and private ranges.'

    try { throw 'User alice@contoso.com cannot read C:\Users\Alice. Response status code does not indicate success: 403 (Forbidden).' } catch { $Record = $_ }
    $Text = Get-CollectorErrorText $Context $Record
    Assert-Privacy ($Text -ceq 'RuntimeException; HTTP 403; code Forbidden') "Error text '$Text' not minimized"
    Assert-Privacy ((Get-CollectorErrorText $Identified $Record).Contains('alice@contoso.com')) 'Identified error text lost detail'
    Assert-Privacy ((Get-CollectorErrorText $Context 'TransportError') -ceq 'TransportError' -and (Get-CollectorErrorText $Context 'user alice@contoso.com failed') -ceq 'CollectionError') 'Plain error codes handled incorrectly'
    try { throw '{"error":{"code":"Authorization_RequestDenied","message":"alice@contoso.com"}}' } catch { $Record = $_ }
    Assert-Privacy ((Get-CollectorErrorText $Context $Record) -ceq 'RuntimeException; code Authorization_RequestDenied') 'Graph error code not extracted'
    $Manifest = New-CollectorPrivacyManifest $Context @('B', 'A', '', 'A')
    Assert-Privacy ($Manifest.Mode -ceq 'Pseudonymous' -and $Manifest.PseudonymKeyId -ceq $Context.KeyId -and ($Manifest.OptIns -join ',') -ceq 'A,B' -and ($Manifest.PseudonymizedFieldClasses -join ',') -ceq 'Person') 'Pseudonymous manifest incorrect'
    $Manifest = New-CollectorPrivacyManifest $Identified @()
    Assert-Privacy ($Manifest.Mode -ceq 'Identified' -and $null -eq $Manifest.PseudonymKeyId -and @($Manifest.PseudonymizedFieldClasses).Count -eq 0 -and ($Manifest.RemovedFieldClasses -join ',') -ceq 'Secret,Content') 'Identified manifest incorrect'
    Write-Output 'PASS: error text minimization and privacy manifest contents.'
} finally {
    Remove-Item -LiteralPath $Root -Recurse -Force -ErrorAction SilentlyContinue
}
