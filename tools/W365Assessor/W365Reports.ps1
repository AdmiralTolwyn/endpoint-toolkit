#Requires -Version 5.1

function Get-W365ReportContract {
    param([string]$Action)
    switch -CaseSensitive ($Action) {
        'retrieveConnectionQualityReports' {
            return @{ ReportName = 'regionalConnectionQualityTrendReport'; Columns = [ordered]@{ GatewayRegion = 'String'; DailyAvgRoundTripTimeInMs = 'Double'; DailyAvailableBandwidthInMbps = 'Double' } }
        }
        'retrieveCloudPcTenantMetricsReport' {
            return @{ ReportName = 'performanceTrendReport'; Columns = [ordered]@{ EventDateTime = 'DateTime'; SlowRoundTripTimeCloudPcCount = 'Int64'; LowUdpConnectionPercentageCount = 'Int64' } }
        }
        default { throw 'Report action is not enabled by the read-only adapter' }
    }
}

function ConvertTo-W365ReportPageEvidence {
    param([object]$Response, [string]$Action)
    $Contract = Get-W365ReportContract $Action
    $MaxBytes = 1MB
    if ($Response -is [IO.Stream]) {
        $Buffer = New-Object byte[] 8192
        $Content = [IO.MemoryStream]::new()
        try {
            while (($Read = $Response.Read($Buffer, 0, $Buffer.Length)) -gt 0) {
                if ($Content.Length + $Read -gt $MaxBytes) { throw 'Report body exceeds decoder limit' }
                $Content.Write($Buffer, 0, $Read)
            }
            $Response = $Content.ToArray()
        } finally { $Content.Dispose() }
    }
    if ($Response -is [byte[]]) {
        if ($Response.Length -gt $MaxBytes) { throw 'Report body exceeds decoder limit' }
        $Response = [Text.UTF8Encoding]::new($false, $true).GetString($Response)
    }
    if ($Response -is [string]) {
        if ($Response.Length -gt $MaxBytes -or [Text.Encoding]::UTF8.GetByteCount($Response) -gt $MaxBytes) { throw 'Report body exceeds decoder limit' }
        $Response = ConvertFrom-Json -InputObject $Response -ErrorAction Stop
    }
    if ($Response -isnot [Collections.IDictionary] -and $Response -isnot [pscustomobject]) { throw 'Invalid report object' }
    foreach ($Field in @('TotalRowCount','Schema','Values')) {
        $Present = if ($Response -is [Collections.IDictionary]) { $Response.Contains($Field) } else { $null -ne $Response.PSObject.Properties[$Field] }
        if (-not $Present) { throw 'Missing report envelope field' }
    }
    if ($null -ne $Response.error -or $Response.Schema -isnot [Collections.IList] -or $Response.Values -isnot [Collections.IList]) { throw 'Invalid report table envelope' }
    $Total = $Response.TotalRowCount
    if (($Total -isnot [int] -and $Total -isnot [long]) -or $Total -lt 0 -or $Total -gt [int]::MaxValue -or $Total -lt $Response.Values.Count -or $Response.Values.Count -gt 25) { throw 'Invalid report row count' }
    if ($Response.Schema.Count -ne $Contract.Columns.Count) { throw 'Unexpected report columns' }
    $Seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($Column in $Response.Schema) {
        if ($Column.Column -isnot [string] -or $Column.Column -cnotin @($Contract.Columns.Keys) -or -not $Seen.Add($Column.Column) -or $Column.PropertyType -cne $Contract.Columns[$Column.Column]) { throw 'Unreviewed report schema' }
    }
    foreach ($Row in $Response.Values) {
        if ($Row -isnot [Collections.IList] -or $Row.Count -ne $Response.Schema.Count) { throw 'Invalid report row width' }
        foreach ($Cell in $Row) {
            if ($null -eq $Cell) { continue }
            if ($Cell -is [string] -and $Cell.Length -le 4096) { continue }
            if ($Cell -is [datetime] -or $Cell -is [datetimeoffset]) { continue }
            if (($Cell -is [int] -or $Cell -is [long] -or $Cell -is [double] -or $Cell -is [decimal]) -and -not [double]::IsNaN([double]$Cell) -and -not [double]::IsInfinity([double]$Cell)) { continue }
            throw 'Unsupported report cell shape'
        }
    }
    return [pscustomobject]@{
        Action = $Action
        ReportName = $Contract.ReportName
        ApiVersion = 'beta'
        RowsReturned = $Response.Values.Count
        ReportedTotalRowCount = $Total
        RequestedTop = 25
        RequestedSkip = 0
        PageCoverage = $(if ($Total -eq $Response.Values.Count) { 'CompletePageSet' } else { 'PartialPageSet' })
        Columns = @($Contract.Columns.Keys)
        AssessmentState = 'ReportContentNotEvaluated'
    }
}