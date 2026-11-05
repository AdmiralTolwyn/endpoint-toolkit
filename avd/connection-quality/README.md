# AVD Connection Quality & Latency Evaluator (Log Analytics)

[Get-AvdConnectionQuality.ps1](Get-AvdConnectionQuality.ps1) evaluates Azure Virtual Desktop (AVD) session round-trip time (RTT), available bandwidth, transport protocol (RDP Shortpath UDP vs. TURN relay vs. WebSocket TCP), corporate proxy/VPN hairpinning, and session-host encoding bottlenecks using Azure Monitor Log Analytics.

---

## 1. What It Evaluates & Why Naïve KQL Joins Fail

1. **Deduplicated `WVDConnections` Join:**
   * `WVDConnections` emits multiple state rows per `CorrelationId` (`Started`, `Connected`, `Completed`), and fields such as `TransportType`, `ClientSideIPAddress`, `GatewayRegion`, and `SessionHostName` are often populated only on the `Connected` row.
   * Joining `WVDConnectionNetworkData` directly to `WVDConnections` with KQL's default `innerunique` join drops network samples, while an un-deduplicated `inner` join multiplies every network sample 2–3×.
   * This script collapses `WVDConnections` to one metadata record per `(CorrelationId, _ResourceId)` using `take_anyif(..., isnotempty(...))` (with a 24-hour pre-window lookback for long-running sessions) before joining 2-minute `WVDConnectionNetworkData` averages.
2. **Schema-Resilient Transport Classification:**
   * Normalizes both current `TransportType` (`Shortpath`, `TURN`, `Websocket`) and legacy `UdpUse` columns into `DirectUdpShortpath`, `RelayedUdpTurn`, `TcpWebsocket`, `UdpSubtypeUnknown`, or `Unknown`.
3. **SWG / VPN Proxy Hairpin Detection (`-CorporateEgressCidrs` & `-ExpectedGatewayRegions`):**
   * Matches `ClientSideIPAddress` against IPv4/IPv6 corporate egress CIDR ranges (including IPv4-mapped IPv6 `::ffff:a.b.c.d`) and compares `GatewayRegion` against expected gateway regions to flag `SuspectedProxyHairpin` when non-direct transport or gateway mismatch coincides with `SessionP95RttMs >= 150 ms`.
4. **Host vs. Network Bottleneck Separation (`-IncludeGuestPerfCounters` & `-IncludeLogonDuration`):**
   * Optionally joins AMA `Perf` counters (`RemoteFX Graphics` `Average Encoding Time > 33 ms` and `User Input Delay per Session` `Max Input Delay > 100 ms`) and `WVDCheckpoints` logon milestones so host CPU/GPU starvation is separated from network latency.
5. **Session-Weighted Summaries (`-OutputView`):**
   * Computes `HostPoolSummary` (grouped by `HostPoolResourceId`), `UserSummary`, and `RegionSummary` across per-session metrics so a single 10-hour session does not outweigh twenty 15-minute sessions.

---

## 2. Prerequisites

* **Host Pool Diagnostic Settings** configured to send to the target Log Analytics Workspace:
  * `Connection` (`WVDConnections`)
  * `NetworkData` (`WVDConnectionNetworkData`)
  * `Checkpoint` (`WVDCheckpoints` — optional, used when `-IncludeLogonDuration` is specified)
* **Azure Monitor Agent (AMA) / AVD Insights DCR** *(optional)* — populates the `Perf` table when `-IncludeGuestPerfCounters` is specified.
* **Permissions:** `Log Analytics Reader` on the target Log Analytics Workspace.

---

## 3. Arguments

| Parameter | Type | Default | Description |
| --- | --- | --- | --- |
| `-WorkspaceId` | `string` | *(Mandatory)* | Log Analytics Workspace GUID (`CustomerId`). |
| `-LookbackHours` | `int` | `168` | Time window in hours to evaluate (`1..2160`, default 7 days). |
| `-HostPoolResourceId` | `string[]` | All pools | Optional ARM Resource ID(s) of AVD Host Pools to filter. |
| `-UserPrincipalName` | `string` | All users | Optional user UPN filter. |
| `-MinSampleCount` | `int` | `3` | Minimum 2-minute `WVDConnectionNetworkData` samples required to classify a session (`InsufficientSamples` below threshold). |
| `-CorporateEgressCidrs` | `string[]` | — | Optional IPv4/IPv6 CIDR ranges of corporate SWG/VPN egress IPs (e.g. `198.51.100.0/24`). |
| `-ExpectedGatewayRegions` | `string[]` | — | Optional expected AVD Gateway region(s) (e.g. `westeurope`, `northeurope`). |
| `-IncludeLogonDuration` | `switch` | `$false` | Join `WVDCheckpoints` logon duration (`LogonDurationSec`). |
| `-IncludeGuestPerfCounters` | `switch` | `$false` | Join AMA `Perf` encoding time (`HostP95EncodingTimeMs`) and input delay (`HostP95InputDelayMs`). |
| `-OutputView` | `string` | `Session` | `Session`, `HostPoolSummary`, `UserSummary`, or `RegionSummary`. |
| `-AuthMode` | `string` | `Auto` | `Auto`, `ManagedIdentity`, or `AzContext`. |

### QualityBand Thresholds (`SampleCount >= MinSampleCount`)

| QualityBand | Criteria |
| --- | --- |
| `Excellent` | `SessionP95RttMs < 50 ms` and `SessionRttSpreadMs < 25 ms` |
| `Good` | `SessionP95RttMs < 100 ms` and `SessionRttSpreadMs < 50 ms` |
| `Fair` | `SessionP95RttMs < 150 ms` |
| `Degraded` | `SessionP95RttMs < 200 ms` |
| `Poor` | `SessionP95RttMs >= 200 ms` |
| `InsufficientSamples` | `SampleCount < MinSampleCount` (e.g. `< 6 minutes` of network telemetry) |

---

## 4. Usage Examples

### Host Pool Summary Over Last 7 Days
```powershell
.\avd\connection-quality\Get-AvdConnectionQuality.ps1 `
    -WorkspaceId '11111111-2222-3333-4444-555555555555' `
    -OutputView HostPoolSummary |
    Format-Table HostPoolName, QualifiedSessions, MedianOfSessionP95RttMs, DirectUdpShortpathPct, TcpWebsocketPct, DegradedOrPoorSessionCount -AutoSize
```

### Detailed Session Evaluation with Proxy Hairpin & Host Encoding Checks
```powershell
.\avd\connection-quality\Get-AvdConnectionQuality.ps1 `
    -WorkspaceId '11111111-2222-3333-4444-555555555555' `
    -LookbackHours 24 `
    -CorporateEgressCidrs @('198.51.100.0/24') `
    -ExpectedGatewayRegions @('westeurope') `
    -IncludeLogonDuration `
    -IncludeGuestPerfCounters |
    Format-Table UserName, ShortHostName, TransportCategory, SessionP95RttMs, SuspectedProxyHairpin, HostEncodingBottleneck, QualityBand -AutoSize
```

---

## 5. Offline Regression Tests

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\avd\connection-quality\Test-AvdConnectionQuality.ps1
```
