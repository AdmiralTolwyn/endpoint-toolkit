#Requires -Version 5.1

function ConvertTo-IntuneSignatureTimestamp {
    param($Value)
    if ($Value -is [datetimeoffset]) {
        $Instant = $Value
    } elseif ($Value -is [datetime]) {
        if ($Value.Kind -eq [DateTimeKind]::Unspecified) { return $null }
        if ($Value.Kind -eq [DateTimeKind]::Local -and
            ([TimeZoneInfo]::Local.IsAmbiguousTime($Value) -or [TimeZoneInfo]::Local.IsInvalidTime($Value))) { return $null }
        $Instant = [datetimeoffset]::new($Value)
    } elseif ($Value -is [string]) {
        if ($Value -cnotmatch '^[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}(?:\.[0-9]{1,7})?(?:Z|[+-][0-9]{2}:[0-9]{2})$' -or $Value.EndsWith('-00:00', [StringComparison]::Ordinal)) { return $null }
        $Instant = [datetimeoffset]::MinValue
        $Formats = [string[]]@("yyyy-MM-dd'T'HH:mm:ss'Z'", "yyyy-MM-dd'T'HH:mm:ss.FFFFFFF'Z'", "yyyy-MM-dd'T'HH:mm:sszzz", "yyyy-MM-dd'T'HH:mm:ss.FFFFFFFzzz")
        $Styles = [Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal
        if (-not [datetimeoffset]::TryParseExact($Value, $Formats, [Globalization.CultureInfo]::InvariantCulture, $Styles, [ref]$Instant)) { return $null }
    } else { return $null }
    if ($Instant.UtcDateTime -eq [datetime]::MinValue) { return $null }
    return $Instant.UtcDateTime.ToString('o', [Globalization.CultureInfo]::InvariantCulture)
}

function ConvertFrom-IntuneEndpointJson {
    param([string]$Json)
    $Parsed = ConvertFrom-Json -InputObject $Json -ErrorAction Stop
    if ($PSVersionTable.PSVersion.Major -le 5) { return $Parsed }
    $ModulesProperty = $Parsed.PSObject.Properties['Modules']
    if ($null -eq $ModulesProperty -or $null -eq $ModulesProperty.Value) { return $Parsed }
    $StatusProperty = $ModulesProperty.Value.PSObject.Properties['DefenderStatus']
    if ($null -eq $StatusProperty -or $null -eq $StatusProperty.Value) { return $Parsed }
    $RowsProperty = $StatusProperty.Value.PSObject.Properties['Rows']
    if ($null -eq $RowsProperty) { return $Parsed }
    $Rows = @($RowsProperty.Value)
    foreach ($Row in $Rows) {
        if ($null -ne $Row -and $null -ne $Row.PSObject.Properties['AntivirusSignatureLastUpdated']) {
            $Row.AntivirusSignatureLastUpdated = $null
        }
    }
    $Document = [System.Text.Json.JsonDocument]::Parse($Json)
    try {
        $RawModules = [System.Text.Json.JsonElement]::new()
        $RawStatus = [System.Text.Json.JsonElement]::new()
        $RawRows = [System.Text.Json.JsonElement]::new()
        if ($Document.RootElement.ValueKind -ne [System.Text.Json.JsonValueKind]::Object -or
            -not $Document.RootElement.TryGetProperty('Modules', [ref]$RawModules) -or
            $RawModules.ValueKind -ne [System.Text.Json.JsonValueKind]::Object -or
            -not $RawModules.TryGetProperty('DefenderStatus', [ref]$RawStatus) -or
            $RawStatus.ValueKind -ne [System.Text.Json.JsonValueKind]::Object -or
            -not $RawStatus.TryGetProperty('Rows', [ref]$RawRows) -or
            $RawRows.ValueKind -ne [System.Text.Json.JsonValueKind]::Array) { return $Parsed }
        if ($Rows.Count -ne $RawRows.GetArrayLength()) { throw 'Endpoint status row shape mismatch' }
        for ($Index = 0; $Index -lt $Rows.Count; $Index++) {
            $RawTime = [System.Text.Json.JsonElement]::new()
            if ($RawRows[$Index].ValueKind -eq [System.Text.Json.JsonValueKind]::Object -and
                $RawRows[$Index].TryGetProperty('AntivirusSignatureLastUpdated', [ref]$RawTime)) {
                $Rows[$Index].AntivirusSignatureLastUpdated = if ($RawTime.ValueKind -eq [System.Text.Json.JsonValueKind]::String) { $RawTime.GetString() } else { $null }
            }
        }
        return $Parsed
    } finally { $Document.Dispose() }
}