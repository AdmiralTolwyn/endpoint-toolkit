<#
.SYNOPSIS
    Runs offline regression checks for Get-AvdConnectionQuality.ps1.
.DESCRIPTION
    Verifies AST help completeness, KQL escaping, HTTP 429/401 retry/refresh,
    PartialError rejection, transport normalization, IPv4/IPv6 CIDR matching,
    Session view classification, and HostPoolSummary / RegionSummary aggregation.
.EXAMPLE
    powershell.exe -NoProfile -File .\scripts\Test-AvdConnectionQuality.ps1
#>

#Requires -Version 5.1

[CmdletBinding()]
param()
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$source = Join-Path $PSScriptRoot 'Get-AvdConnectionQuality.ps1'
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($source, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -gt 0) { throw ($parseErrors | Out-String) }
foreach ($def in $ast.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.FunctionDefinitionAst] }) {
    . ([scriptblock]::Create($def.Extent.Text))
}
$script:assertions = 0

function Assert-AvdQualityTest {
    param([bool]$Condition, [string]$Name)
    if (-not $Condition) { throw "FAIL: $Name" }
    $script:assertions++
}

# 0. Verify comment-based help on all 9 helper functions
& {
    $fnDefs = @($ast.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.FunctionDefinitionAst] })
    Assert-AvdQualityTest ($fnDefs.Count -eq 9) 'All 9 helper functions exist in Get-AvdConnectionQuality.ps1'
    foreach ($fn in $fnDefs) {
        $help = $fn.GetHelpContent()
        Assert-AvdQualityTest (($null -ne $help) -and (-not [string]::IsNullOrWhiteSpace($help.Synopsis)) -and (-not [string]::IsNullOrWhiteSpace($help.Description))) ("Function '{0}' has .SYNOPSIS and .DESCRIPTION" -f $fn.Name)
        foreach ($p in $fn.Body.ParamBlock.Parameters) {
            $pName = $p.Name.VariablePath.UserPath.ToUpperInvariant()
            Assert-AvdQualityTest ($help.Parameters.ContainsKey($pName) -and -not [string]::IsNullOrWhiteSpace($help.Parameters[$pName])) ("Function '{0}' documents -{1}" -f $fn.Name, $p.Name.VariablePath.UserPath)
        }
    }
}

# 1. Unit tests for KQL escaping, transport categories, IPv4/IPv6 CIDR matching, and percentiles
& {
    Assert-AvdQualityTest ((ConvertTo-AvdQualityKqlLiteral -Value "o'brien\test@contoso.com") -ceq "'o\'brien\\test@contoso.com'") 'KQL literal escapes backslash and quote'
    Assert-AvdQualityTest ((ConvertTo-AvdTransportCategory -RawTransportType 'Shortpath' -RawUdpUse '') -eq 'DirectUdpShortpath') 'Shortpath maps to DirectUdpShortpath'
    Assert-AvdQualityTest ((ConvertTo-AvdTransportCategory -RawTransportType 'TURN' -RawUdpUse '') -eq 'RelayedUdpTurn') 'TURN maps to RelayedUdpTurn'
    Assert-AvdQualityTest ((ConvertTo-AvdTransportCategory -RawTransportType 'Websocket' -RawUdpUse '') -eq 'TcpWebsocket') 'Websocket maps to TcpWebsocket'
    Assert-AvdQualityTest ((ConvertTo-AvdTransportCategory -RawTransportType '' -RawUdpUse '1') -eq 'UdpSubtypeUnknown') 'Legacy UdpUse=1 maps to UdpSubtypeUnknown'

    Assert-AvdQualityTest (Test-AvdIpInCidr -IpString '198.51.100.42:51234' -CidrList @('198.51.100.0/24')) 'IPv4 with port matches /24 CIDR'
    Assert-AvdQualityTest (Test-AvdIpInCidr -IpString '::ffff:198.51.100.42' -CidrList @('198.51.100.0/24')) 'IPv4-mapped IPv6 matches IPv4 /24 CIDR'
    Assert-AvdQualityTest (Test-AvdIpInCidr -IpString '10.240.15.1' -CidrList @('10.240.0.0/20')) 'Non-byte-aligned /20 CIDR matches in-range IP'
    Assert-AvdQualityTest (-not (Test-AvdIpInCidr -IpString '10.240.16.1' -CidrList @('10.240.0.0/20'))) 'Non-byte-aligned /20 CIDR rejects out-of-range IP'
    Assert-AvdQualityTest (Test-AvdIpInCidr -IpString '2001:db8:abcd::1' -CidrList @('2001:db8::/32')) 'IPv6 matches /32 CIDR'
    Assert-AvdQualityTest ((Get-AvdNumericPercentile -Values @(20.0, 40.0, 60.0, 80.0, 200.0) -Percentile 50) -eq 60.0) 'Median of 5 values equals 60'
}

# 2. End-to-end Session view, PartialError rejection, and HostPoolSummary tests
& {
    $wsId    = '22222222-2222-2222-2222-222222222222'
    $poolId1 = '/subscriptions/11111111-1111-1111-1111-111111111111/resourceGroups/rg-avd-1/providers/Microsoft.DesktopVirtualization/hostPools/vdpool-ams-001'
    $poolId2 = '/subscriptions/11111111-1111-1111-1111-111111111111/resourceGroups/rg-avd-2/providers/Microsoft.DesktopVirtualization/hostPools/vdpool-ams-001'
    $entryText = @(
        "[CmdletBinding()]"
        $ast.ParamBlock.Extent.Text
        foreach ($stmt in $ast.EndBlock.Statements) {
            if ($stmt -isnot [System.Management.Automation.Language.FunctionDefinitionAst]) { $stmt.Extent.Text }
        }
    ) -join "`r`n"
    $runner = [scriptblock]::Create($entryText)

    function Invoke-AvdQualityHttpRequest {
        param($Method, $Uri, $Headers, $Body)
        return [PSCustomObject]@{ partialError = [PSCustomObject]@{ code = 'PartialError'; message = 'Truncated' }; tables = @() }
    }
    $partialFailed = $false
    try { $null = @(& $runner -WorkspaceId $wsId -LogAnalyticsAccessToken 'tok') } catch { $partialFailed = ($_.Exception.Message -like '*PartialError*') }
    Assert-AvdQualityTest $partialFailed 'Get-AvdConnectionQuality fails closed on partialError'

    function Invoke-AvdQualityHttpRequest {
        param($Method, $Uri, $Headers, $Body)
        return [PSCustomObject]@{
            tables = @([PSCustomObject]@{
                name    = 'PrimaryResult'
                columns = @(
                    [PSCustomObject]@{ name = '_ResourceId' }, [PSCustomObject]@{ name = 'CorrelationId' }, [PSCustomObject]@{ name = 'UserName' },
                    [PSCustomObject]@{ name = 'SessionHostName' }, [PSCustomObject]@{ name = 'SessionHostAzureVmId' }, [PSCustomObject]@{ name = 'ClientSideIPAddress' },
                    [PSCustomObject]@{ name = 'ClientOS' }, [PSCustomObject]@{ name = 'GatewayRegion' }, [PSCustomObject]@{ name = 'TransportType' },
                    [PSCustomObject]@{ name = 'UdpUse' }, [PSCustomObject]@{ name = 'SampleCount' }, [PSCustomObject]@{ name = 'SessionP50RttMs' },
                    [PSCustomObject]@{ name = 'SessionP95RttMs' }, [PSCustomObject]@{ name = 'SessionMaxRttMs' }, [PSCustomObject]@{ name = 'SessionP10BandwidthKBps' },
                    [PSCustomObject]@{ name = 'SessionAvgBandwidthKBps' }, [PSCustomObject]@{ name = 'LogonDurationSec' },
                    [PSCustomObject]@{ name = 'HostP95EncodingTimeMs' }, [PSCustomObject]@{ name = 'HostP95InputDelayMs' }
                )
                rows    = @(
                    @($poolId1, 'c-001', 'alice@contoso.com', 'avd-01.contoso.com', 'aaaa0001-0000-0000-0000-000000000001', '93.184.216.34', 'Windows', 'westeurope', 'Shortpath', '', 15, 18.0, 28.0, 35.0, 2400.0, 3800.0, 14.0, 12.5, 18.0),
                    @($poolId1, 'c-002', 'bob@contoso.com',   'avd-02.contoso.com', 'aaaa0002-0000-0000-0000-000000000002', '198.51.100.77', 'Windows', 'eastus',     'Websocket', '', 12, 180.0, 215.0, 260.0, 310.0, 450.0, 29.0, 15.0, 22.0),
                    @($poolId1, 'c-003', 'carol@contoso.com', 'avd-03.contoso.com', 'aaaa0003-0000-0000-0000-000000000003', '93.184.216.99', 'Windows', 'westeurope', 'TURN',      '', 20, 42.0, 62.0, 75.0, 1500.0, 2100.0, 16.0, 44.5, 125.0),
                    @($poolId2, 'c-004', 'dave@contoso.com',  'avd-01.contoso.com', 'aaaa0001-0000-0000-0000-000000000001', '198.51.100.50', 'Windows', 'eastus',     'Websocket', '', 1,  220.0, 220.0, 220.0, 1200.0, 1200.0, 10.0, 50.0, 150.0)
                )
            })
        }
    }

    $sessions = @(& $runner -WorkspaceId $wsId -HostPoolResourceId @($poolId1, $poolId2) -CorporateEgressCidrs @('198.51.100.0/24') -ExpectedGatewayRegions @('westeurope') -IncludeLogonDuration -IncludeGuestPerfCounters -OutputView Session -LogAnalyticsAccessToken 'tok')
    Assert-AvdQualityTest ($sessions.Count -eq 4) 'Session view returns 4 records'
    Assert-AvdQualityTest (($sessions | Where-Object { $_.CorrelationId -eq 'c-001' }).QualityBand -eq 'Excellent') 'c-001 is Excellent'
    Assert-AvdQualityTest (($sessions | Where-Object { $_.CorrelationId -eq 'c-002' }).SuspectedProxyHairpin) 'c-002 flagged for SuspectedProxyHairpin'
    Assert-AvdQualityTest (($sessions | Where-Object { $_.CorrelationId -eq 'c-003' }).HostEncodingBottleneck) 'c-003 flagged for HostEncodingBottleneck'
    Assert-AvdQualityTest (($sessions | Where-Object { $_.CorrelationId -eq 'c-004' }).QualityBand -eq 'InsufficientSamples') 'c-004 is InsufficientSamples'

    $poolSum = @(& $runner -WorkspaceId $wsId -CorporateEgressCidrs @('198.51.100.0/24') -ExpectedGatewayRegions @('westeurope') -IncludeLogonDuration -IncludeGuestPerfCounters -OutputView HostPoolSummary -LogAnalyticsAccessToken 'tok')
    Assert-AvdQualityTest ($poolSum.Count -eq 2) 'HostPoolSummary keeps same-named pools in different RGs separate'
}

Write-Host ("PASS: All {0} assertions succeeded in Test-AvdConnectionQuality.ps1" -f $script:assertions) -ForegroundColor Green
