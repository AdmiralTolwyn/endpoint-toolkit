#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
$Tokens = $null
$ParseErrors = $null
$Ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'Invoke-BaselineCollection.ps1'), [ref]$Tokens, [ref]$ParseErrors)
if ($ParseErrors.Count) { throw 'Collector parse errors' }
foreach ($Function in $Ast.EndBlock.Statements | Where-Object { $_ -is [Management.Automation.Language.FunctionDefinitionAst] -and $_.Name -match 'Collector|^ConvertTo-Baseline|^Test-Baseline|^Get-BaselineEventData$' }) {
    . ([scriptblock]::Create($Function.Extent.Text))
}
function Get-Area([string]$Variable) {
    $Assignment = @($Ast.FindAll({ param($Node) $Node -is [Management.Automation.Language.AssignmentStatementAst] -and $Node.Left.Extent.Text -ceq $Variable -and $Node.Right.Extent.Text.StartsWith('Invoke-CollectionArea') }, $true))
    if ($Assignment.Count -ne 1) { throw ('Expected one production area: ' + $Variable) }
    $Block = @($Assignment[0].Right.PipelineElements[0].CommandElements | Where-Object { $_ -is [Management.Automation.Language.ScriptBlockExpressionAst] })
    $Text = $Block[0].ScriptBlock.Extent.Text.Trim()
    return [scriptblock]::Create($Text.Substring(1, $Text.Length - 2))
}
function Assert-Privacy([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
function Write-CollectorProgress { }
$Lolbins = @($Ast.FindAll({ param($Node) $Node -is [Management.Automation.Language.AssignmentStatementAst] -and $Node.Left.Extent.Text -ceq '$Script:LolbinNames' }, $true))
. ([scriptblock]::Create($Lolbins[0].Extent.Text))

$Days = ConvertTo-BaselineWorkDays 'Sun-Thu'
Assert-Privacy (($Days -join ',') -eq '0,1,2,3,4') 'Sun-Thu range parsed incorrectly'
Assert-Privacy (((ConvertTo-BaselineWorkDays 'fri-mon') -join ',') -eq '0,1,5,6') 'Wrapping range parsed incorrectly'
Assert-Privacy (((ConvertTo-BaselineWorkDays 'Mon, Wed,FRI') -join ',') -eq '1,3,5') 'List parsed incorrectly'
foreach ($Invalid in @('', 'Mon-Xyz', 'Monday', 'Mon;Tue')) {
    $Rejected = $false
    try { ConvertTo-BaselineWorkDays $Invalid | Out-Null } catch { $Rejected = $true }
    Assert-Privacy $Rejected "Invalid work days accepted: $Invalid"
}
$Monday = [datetime]'2026-10-05T00:00:00'
Assert-Privacy (-not (Test-BaselineOffHours $Monday.AddHours(6) @(1) 360 1320)) 'Business start must be inside hours'
Assert-Privacy (Test-BaselineOffHours $Monday.AddHours(22) @(1) 360 1320) 'Business end must be off-hours'
Assert-Privacy (Test-BaselineOffHours $Monday.AddHours(12) @(2) 360 1320) 'Non-work day must be off-hours'
Assert-Privacy (-not (Test-BaselineOffHours $Monday.AddHours(23) @(1) 1320 360)) 'Overnight shift must wrap midnight'
Assert-Privacy (Test-BaselineOffHours $Monday.AddHours(12) @(1) 1320 360) 'Overnight shift midday must be off-hours'
Write-Output 'PASS: configurable business hours and work days, including wrapped ranges, overnight windows and rejected input.'

function New-SyntheticEvent([int]$Id, [datetime]$Time, [hashtable]$Named, [string[]]$Values = @()) {
    $Data = foreach ($Key in $Named.Keys) { '<Data Name="{0}">{1}</Data>' -f $Key, [Security.SecurityElement]::Escape([string]$Named[$Key]) }
    $Data += foreach ($Value in $Values) { '<Data>{0}</Data>' -f [Security.SecurityElement]::Escape($Value) }
    $Xml = '<Event xmlns="http://schemas.microsoft.com/win/2004/08/events/event"><EventData>' + ($Data -join '') + '</EventData></Event>'
    $Record = [pscustomobject]@{ Id = $Id; TimeCreated = $Time; Message = 'RAW_MESSAGE_SENTINEL C:\Users\PII_SENTINEL' }
    $Record | Add-Member -MemberType ScriptMethod -Name ToXml -Value ([scriptblock]::Create("'" + $Xml.Replace("'", "''") + "'"))
    return $Record
}
$Saturday = (Get-Date).Date
while ($Saturday.DayOfWeek -ne [DayOfWeek]::Saturday) { $Saturday = $Saturday.AddDays(-1) }
$OffHours = $Saturday.AddHours(2)
if ($OffHours -gt (Get-Date)) { $OffHours = $OffHours.AddDays(-7) }
$Weekday = $OffHours.AddDays(-3).Date.AddHours(10)
$script:SyntheticEvents = @(
    (New-SyntheticEvent 4672 $OffHours @{ SubjectLogonId = '0x3e7abc'; SubjectUserName = 'PII_SENTINEL' }),
    (New-SyntheticEvent 4672 $Weekday @{ SubjectLogonId = '0x51'; SubjectUserName = 'PII_SENTINEL' }),
    (New-SyntheticEvent 4672 $OffHours @{ SubjectLogonId = '0x3e7'; SubjectUserName = 'SYSTEM' }),
    (New-SyntheticEvent 4624 $OffHours @{ LogonType = '10'; TargetLogonId = '0x3E7ABC'; TargetUserSid = 'S-1-5-21-111-222-333-1001'; TargetUserName = 'PII_SENTINEL'; TargetDomainName = 'CONTOSO'; IpAddress = '203.0.113.77'; WorkstationName = 'PII_WORKSTATION' }),
    (New-SyntheticEvent 4624 $Weekday @{ LogonType = '2'; TargetLogonId = '0x51'; TargetUserSid = 'S-1-5-21-111-222-333-1001'; TargetUserName = 'PII_SENTINEL'; TargetDomainName = 'CONTOSO' }),
    (New-SyntheticEvent 4624 $OffHours @{ LogonType = '5'; TargetLogonId = '0x3e7'; TargetUserSid = 'S-1-5-18'; TargetUserName = 'SYSTEM'; TargetDomainName = 'NT AUTHORITY' }),
    (New-SyntheticEvent 4625 $OffHours @{ LogonType = '10'; TargetUserName = 'PII_SENTINEL'; IpAddress = '203.0.113.77'; FailureReason = '%%2313'; SubStatus = '0xc000006a' }),
    (New-SyntheticEvent 4625 $OffHours @{ LogonType = '3'; TargetUserName = 'PII_SENTINEL'; IpAddress = '203.0.113.77' }),
    (New-SyntheticEvent 4740 $OffHours @{ TargetUserName = 'PII_SENTINEL'; TargetDomainName = 'CONTOSO'; CallerComputerName = 'PII_HOST' }),
    (New-SyntheticEvent 4740 $Weekday @{ TargetUserName = 'pii_sentinel'; TargetDomainName = 'contoso'; CallerComputerName = 'PII_HOST' }),
    (New-SyntheticEvent 4688 $OffHours @{ NewProcessName = 'C:\Windows\System32\CertUtil.exe'; CommandLine = 'certutil -urlcache SECRET_TOKEN=abc'; SubjectUserName = 'PII_SENTINEL' }),
    (New-SyntheticEvent 4688 $Weekday @{ NewProcessName = 'C:\Users\PII_SENTINEL\private.exe'; CommandLine = 'private.exe --password SECRET_TOKEN'; SubjectUserName = 'PII_SENTINEL' }),
    (New-SyntheticEvent 1000 $Weekday @{} @('C:\Users\PII_SENTINEL\tool.exe', '1.0.0.0')),
    (New-SyntheticEvent 41 $Weekday @{ BugcheckCode = '0' })
)
function Get-WinEvent {
    param($FilterHashtable, $MaxEvents, $ErrorAction)
    $Matched = @($script:SyntheticEvents | Where-Object { $_.Id -in @($FilterHashtable.Id) -and $_.TimeCreated -ge $FilterHashtable.StartTime })
    if ($Matched.Count -eq 0) { throw 'No events were found that match the specified selection criteria.' }
    $Matched | Sort-Object TimeCreated -Descending | Select-Object -First $MaxEvents
}
function Invoke-EventArea([string]$Mode, [bool]$Detailed, [bool]$Summary, [byte[]]$Key) {
    $script:PrivacyContext = New-CollectorPrivacyContext -Mode Identified -ConfirmIdentified $true
    if ($Mode -eq 'Pseudonymous') { $script:PrivacyContext = [pscustomobject]@{ Mode = 'Pseudonymous'; Key = $Key; KeyId = 'synthetic'; KeyPath = $null; Identities = New-Object 'Collections.Generic.Dictionary[string,string]' } }
    $script:LookbackDays = 30
    $script:MaxEventsPerQuery = 2000
    $script:IncludeSecurityEvents = $Detailed
    $script:EventSummaryOnly = $Summary
    $script:BusinessHours = '06:00-22:00'
    $script:WorkDays = 'Mon-Fri'
    $script:WorkDayNumbers = ConvertTo-BaselineWorkDays 'Mon-Fri'
    $script:BusinessStartMinute = 360
    $script:BusinessEndMinute = 1320
    return & (Get-Area '$eventData')
}
$Key = [byte[]](1..32)
$Minimal = Invoke-EventArea 'Pseudonymous' $false $false $Key
$MinimalJson = $Minimal | ConvertTo-Json -Depth 10
foreach ($Sentinel in @('PII_SENTINEL', 'pii_sentinel', '203.0.113.77', 'SECRET_TOKEN', 'PII_HOST', 'PII_WORKSTATION', 'RAW_MESSAGE_SENTINEL', 'CONTOSO')) {
    Assert-Privacy (-not $MinimalJson.Contains($Sentinel)) "Default export leaked $Sentinel"
}
foreach ($Removed in @('privilegeUse', 'policyChange', 'systemIntegrity', 'objectAccess', 'hardwareErrors')) { Assert-Privacy (-not $Minimal.Contains($Removed)) "Unused query $Removed still collected" }
$Interactive = @($Minimal.logonEvents | Where-Object { $_.id -eq 4624 -and $_.logonType -eq 10 })
Assert-Privacy ($Interactive.Count -eq 1 -and $Interactive[0].elevated -eq $true -and $Interactive[0].offHours -eq $true) 'Elevated off-hours logon not derived'
$Daytime = @($Minimal.logonEvents | Where-Object { $_.id -eq 4624 -and $_.logonType -eq 2 })
Assert-Privacy ($Daytime[0].elevated -eq $true -and $Daytime[0].offHours -eq $false) 'Business-hours elevated logon misclassified'
$Service = @($Minimal.logonEvents | Where-Object { $_.id -eq 4624 -and $_.logonType -eq 5 })
Assert-Privacy ($Service[0].elevated -eq $false) 'Service SYSTEM logon counted as elevated administrator'
Assert-Privacy (@($Minimal.logonEvents | Where-Object { $_.id -eq 4625 -and $_.logonType -eq 10 }).Count -eq 1) 'RDP failure logon type lost'
Assert-Privacy ($Minimal._queryMeta.elevatedCorrelation -eq 'Complete' -and $Minimal._queryMeta.recordSchema -eq 'minimal-1.0') 'Derivation metadata missing'
$Accounts = @($Minimal.accountLockout | ForEach-Object { $_.accountKey } | Sort-Object -Unique)
Assert-Privacy ($Accounts.Count -eq 1 -and $Accounts[0] -cmatch '^usr_[0-9a-f]{16}$') 'Lockout account pseudonym missing or not case-normalized'
$Again = Invoke-EventArea 'Pseudonymous' $false $false $Key
Assert-Privacy ($Again.accountLockout[0].accountKey -ceq $Minimal.accountLockout[0].accountKey) 'Pseudonym unstable with the same key'
$Other = Invoke-EventArea 'Pseudonymous' $false $false ([byte[]](32..63))
Assert-Privacy ($Other.accountLockout[0].accountKey -cne $Minimal.accountLockout[0].accountKey) 'Pseudonym did not depend on the key'
Assert-Privacy (@($Minimal.processCreation | Where-Object { $_.lolbin -ceq 'certutil.exe' }).Count -eq 1 -and @($Minimal.processCreation | Where-Object { $_.Contains('lolbin') }).Count -eq 1) 'LOLBin derivation incorrect'
Assert-Privacy ($Minimal.applicationCrashes[0].faultingApp -ceq 'tool.exe') 'Faulting application must be a file name only'
Assert-Privacy (@($Minimal.Values | Where-Object { $_ -is [array] } | ForEach-Object { $_ } | Where-Object { $_.time -notmatch '^\d{4}-\d{2}-\d{2}T\d{2}:00:00Z$' }).Count -eq 0) 'Event time not truncated to the UTC hour'

$Detailed = Invoke-EventArea 'Pseudonymous' $true $false $Key
$DetailedJson = $Detailed | ConvertTo-Json -Depth 10
Assert-Privacy ($DetailedJson.Contains('%%2313') -and -not $DetailedJson.Contains('PII_SENTINEL') -and -not $DetailedJson.Contains('203.0.113.77')) 'Pseudonymous diagnostics included identity fields'
$Identified = Invoke-EventArea 'Identified' $true $false $Key
$IdentifiedJson = $Identified | ConvertTo-Json -Depth 10
Assert-Privacy ($IdentifiedJson.Contains('PII_SENTINEL') -and $IdentifiedJson.Contains('203.0.113.77')) 'Identified diagnostics omitted identity fields'
foreach ($Never in @('SECRET_TOKEN', 'RAW_MESSAGE_SENTINEL', 'CommandLine', 'ObjectName')) { Assert-Privacy (-not $IdentifiedJson.Contains($Never)) "Identified export included prohibited $Never" }
Assert-Privacy ($Identified.accountLockout[0].accountKey -match '\\') 'Identified account key should retain the account'
$Summary = Invoke-EventArea 'Pseudonymous' $false $true $Key
Assert-Privacy (-not ($Summary | ConvertTo-Json -Depth 10).Contains('topUsers') -and $Summary.logonEvents.count -gt 0) 'Summary mode exported users or lost counts'
Write-Output 'PASS: minimal event records, derived flags, keyed pseudonyms, opt-in diagnostics and prohibited content.'

$script:PrivacyContext = [pscustomobject]@{ Mode = 'Pseudonymous'; Key = $Key; KeyId = 'synthetic'; KeyPath = $null; Identities = New-Object 'Collections.Generic.Dictionary[string,string]' }
function Get-CimInstance {
    param([string]$ClassName, $ErrorAction)
    if ($ClassName -eq 'Win32_PnPSignedDriver') { [pscustomobject]@{ DriverProviderName = 'Vendor'; DeviceName = "PII_SENTINEL's Headphones"; IsSigned = $false; DriverVersion = '1.0'; Manufacturer = 'Vendor'; DeviceClass = 'BLUETOOTH'; HardWareID = 'BTHENUM\{0000110b}_VID&0001004c_PID&2027' } }
    else { [pscustomobject]@{ Name = "PII_SENTINEL's Phone"; ConfigManagerErrorCode = 28; Status = 'Error'; Manufacturer = 'Vendor'; PNPClass = 'WPD'; HardwareID = @('USB\VID_05AC&PID_12A8&REV_0001&MI_00') } }
}
$Drivers = & (Get-Area '$drivers')
$DriversJson = $Drivers | ConvertTo-Json -Depth 6
Assert-Privacy (-not $DriversJson.Contains('PII_SENTINEL') -and $Drivers.problematic[0].Name -ceq 'WPD USB\VID_05AC&PID_12A8' -and $Drivers.unsigned[0].Name -like 'BLUETOOTH BTHENUM*') 'Driver names leaked or lost class evidence'
Remove-Item Function:\Get-CimInstance
function Get-ScheduledTask { [pscustomobject]@{ TaskName = 'OneDrive Reporting Task-S-1-5-21-111-222-333-1001'; TaskPath = '\'; State = 'Ready'; LastTaskResult = 1; Principal = [pscustomobject]@{ UserId = 'SYSTEM' } } }
$Tasks = & (Get-Area '$scheduledTasks')
Assert-Privacy (($Tasks | ConvertTo-Json -Depth 6) -cmatch 'Task-sid_[0-9a-f]{16}' -and -not ($Tasks | ConvertTo-Json -Depth 6).Contains('S-1-5-21-111')) 'Task SID not pseudonymized'
Remove-Item Function:\Get-ScheduledTask
function Get-MpPreference { [pscustomobject]@{ ExclusionPath = @('C:\Users\PII_SENTINEL\Source', 'D:\Builds\*', '\\PII_HOST\share\tools') } }
function Get-MpComputerStatus { $null }
$Defender = & (Get-Area '$defenderConfig')
Assert-Privacy (($Defender.ExclusionPath -join '|') -ceq 'C:\Users\{profile}\Source|D:\Builds\*|\\{host}\{share}\tools') 'Exclusion paths not masked'
Remove-Item Function:\Get-MpPreference, Function:\Get-MpComputerStatus
function Get-NetFirewallProfile { throw 'Access denied for C:\Users\PII_SENTINEL' }
$Firewall = & (Get-Area '$firewallProfiles')
Assert-Privacy ($Firewall._collectionFailed -and -not $Firewall._error.Contains('PII_SENTINEL')) 'Provider error text leaked'
Remove-Item Function:\Get-NetFirewallProfile
Write-Output 'PASS: driver names, task SIDs, exclusion paths and provider errors are minimized in default mode.'

if ($env:ASSAY_BASELINE_PRIVACY_FIXTURE) {
    $Fixture = [ordered]@{
        Privacy = New-CollectorPrivacyManifest $script:PrivacyContext @()
        systemInfo = @{ hostname = (ConvertTo-CollectorIdentity $script:PrivacyContext 'dev' 'PII_SENTINEL-PC'); isServer = $false; osBuild = '26200' }
        eventData = $Minimal
    }
    [IO.File]::WriteAllText($env:ASSAY_BASELINE_PRIVACY_FIXTURE, ($Fixture | ConvertTo-Json -Depth 10), [Text.UTF8Encoding]::new($false))
}
