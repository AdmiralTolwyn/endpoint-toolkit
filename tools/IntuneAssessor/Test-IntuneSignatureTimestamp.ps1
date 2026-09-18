#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Invoke-IntuneDiscovery.ps1') -LibraryOnly
. (Join-Path $PSScriptRoot 'Get-IntuneEndpointEvidence.ps1') -LibraryOnly
. (Join-Path $PSScriptRoot 'IntuneEndpointTimestamps.ps1')
$Unspecified = [datetime]::SpecifyKind([datetime]::new(2026,9,18,8,0,0), [DateTimeKind]::Unspecified)
$Safe = ConvertTo-IntuneEndpointModules @{ DefenderStatus = @{ State = 'Complete'; Rows = @(@{ AntivirusSignatureLastUpdated = $Unspecified; AntivirusEnabled = $true }) } }
if ($null -ne $Safe.DefenderStatus.Rows[0]['AntivirusSignatureLastUpdated']) { throw 'Unqualified signature time acquired an assumed timezone' }
if ($Safe.DefenderStatus.Rows[0]['AntivirusEnabled'] -ne $true) { throw 'Other Defender evidence was changed' }
$Cases = @(
	@{ Value = [datetime]::new(2026,9,18,8,0,0,[DateTimeKind]::Utc); Expected = '2026-09-18T08:00:00.0000000Z' },
	@{ Value = [datetimeoffset]::new(2026,9,18,10,0,0,[timespan]::FromHours(2)); Expected = '2026-09-18T08:00:00.0000000Z' },
	@{ Value = '2026-09-18T03:00:00-05:00'; Expected = '2026-09-18T08:00:00.0000000Z' },
	@{ Value = '2026-09-18T08:00:00.1234567Z'; Expected = '2026-09-18T08:00:00.1234567Z' },
	@{ Value = '2026-09-18T08:00:00Z'; Expected = '2026-09-18T08:00:00.0000000Z' },
	@{ Value = '2026-09-18T13:30:00+05:30'; Expected = '2026-09-18T08:00:00.0000000Z' },
	@{ Value = '2026-09-18T08:00:00.1+00:00'; Expected = '2026-09-18T08:00:00.1000000Z' }
)
$Local = [datetime]::new(2026,1,15,12,0,0,[DateTimeKind]::Local)
if (-not [TimeZoneInfo]::Local.IsAmbiguousTime($Local) -and -not [TimeZoneInfo]::Local.IsInvalidTime($Local)) {
	$Cases += @{ Value = $Local; Expected = $Local.ToUniversalTime().ToString('o') }
}
foreach ($Invalid in @($Unspecified,$null,[datetime]::MinValue,42,$true,'2026-09-18T08:00:00','2026-09-18T08:00:00-00:00','2026-02-30T08:00:00Z','2026-09-18T08:00:00+15:00','2026-09-18T08:00:60Z','2026-09-18T08:00:00.12345678Z',' 2026-09-18T08:00:00Z','2026-09-18t08:00:00z','09/18/2026 08:00:00','/Date(1789718400000)/')) {
	$Cases += @{ Value = $Invalid; Expected = $null }
}
$FixtureRows = [Collections.Generic.List[object]]::new()
foreach ($Case in $Cases) {
	$script:SignatureTime = $Case.Value
	$Result = ConvertTo-IntuneSignatureTimestamp $Case.Value
	if ($Result -cne $Case.Expected) { throw 'Signature timestamp normalization mismatch' }
	$Imported = ConvertTo-IntuneEndpointModules @{ DefenderStatus = @{ State = 'Complete'; Rows = @(@{ AntivirusSignatureLastUpdated = $Case.Value; AntivirusEnabled = $true }) } }
	if ($Imported.DefenderStatus.Rows[0]['AntivirusSignatureLastUpdated'] -cne $Case.Expected) { throw 'Importer assumed or lost a signature timezone' }
	if ($Case.Value -is [string]) {
		$RawJson = @{ Modules = @{ DefenderStatus = @{ State = 'Complete'; Rows = @(@{ AntivirusSignatureLastUpdated = $Case.Value; AntivirusEnabled = $true }) } } } | ConvertTo-Json -Depth 15
		$DecodedRaw = ConvertFrom-IntuneEndpointJson $RawJson
		$ImportedRaw = ConvertTo-IntuneEndpointModules $DecodedRaw.Modules
		if ($ImportedRaw.DefenderStatus.Rows[0]['AntivirusSignatureLastUpdated'] -cne $Case.Expected) { throw 'JSON parser changed a signature timestamp before validation' }
	}
	foreach ($Shape in @('dictionary','object')) {
		$script:SignatureShape = $Shape
		$Sample = New-IntuneEndpointEvidence -SelectedTenant '22222222-2222-4222-8222-222222222222' -DeviceId '33333333-3333-4333-8333-333333333333' -Read {
			param($Module)
			if ($Module -eq 'DefenderStatus') {
				$Row = @{ AntivirusSignatureLastUpdated = $script:SignatureTime; AntivirusEnabled = $true }
				if ($script:SignatureShape -eq 'object') { [pscustomobject]$Row } else { $Row }
			} else { throw 'Synthetic unavailable provider' }
		}
		if ($Sample.Modules.DefenderStatus.State -cne 'Complete') { throw 'Timestamp issue discarded status provider' }
		$Stored = $Sample.Modules.DefenderStatus.Rows[0].AntivirusSignatureLastUpdated
		if ($Stored -cne $Case.Expected -or ($null -ne $Stored -and $Stored -isnot [string])) { throw 'Companion did not emit canonical string before JSON' }
		$Decoded = ($Sample | ConvertTo-Json -Depth 15) | ConvertFrom-Json
		$Safe = ConvertTo-IntuneEndpointModules $Decoded.Modules
		if ($Safe.DefenderStatus.Rows[0]['AntivirusSignatureLastUpdated'] -cne $Case.Expected -or $Safe.DefenderStatus.Rows[0]['AntivirusEnabled'] -ne $true) { throw 'Signature time or adjacent evidence changed in JSON roundtrip' }
		$FixtureRows.Add(@{ expected = $Case.Expected; sample = @{ id = $Sample.DeviceId; collectedAtUtc = $Sample.CollectedAtUtc; modules = $Safe } })
	}
}
foreach ($LegacyJson in @(
	'{"modules":{"DefenderStatus":{"State":"Complete","Rows":[{"AntivirusSignatureLastUpdated":"2026-09-18T08:00:00-00:00","AntivirusEnabled":true}]}}}',
	'{"Modules":{"DefenderStatus":{"State":"Complete","Rows":[{"antivirussignaturelastupdated":"/Date(1789718400000)/","AntivirusEnabled":true}]}}}'
)) {
	$Legacy = ConvertFrom-IntuneEndpointJson $LegacyJson
	$Safe = ConvertTo-IntuneEndpointModules $Legacy.Modules
	if ($null -ne $Safe.DefenderStatus.Rows[0]['AntivirusSignatureLastUpdated']) { throw 'Noncanonical field casing bypassed raw timestamp validation' }
}
if ($env:ASSAY_INTUNE_SIGNATURE_TIME_FIXTURE) {
	[IO.File]::WriteAllText($env:ASSAY_INTUNE_SIGNATURE_TIME_FIXTURE, (ConvertTo-Json -InputObject $FixtureRows.ToArray() -Depth 15), [Text.UTF8Encoding]::new($false))
}
Write-Output 'PASS: explicit signature instants roundtrip without timezone guessing across companion/import boundaries; all reads injected.'