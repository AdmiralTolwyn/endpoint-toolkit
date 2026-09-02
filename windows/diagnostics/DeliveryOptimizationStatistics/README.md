# Get-DeliveryOptimizationStatistics.ps1

Reports local Delivery Optimization configuration, month-to-date traffic, efficiency, uploads, cache state, connections, and active jobs. Output can be a human-readable table, a PowerShell object, or single-line JSON for Grafana / Loki / Telegraf ingestion.

The script is read-only and PowerShell 5.1 compatible.

## Why the script reads CIM directly

`Get-DeliveryOptimizationPerfSnapThisMonth` exposes CDN, Connected Cache, LAN-peer, and Internet-peer bytes, but its PowerShell wrapper omits two counters that exist in the underlying provider:

- `MonthlyGroupBytes`
- `MonthlyLinkLocalBytes`

The script reads `root/Microsoft/Windows/DeliveryOptimization` directly so those sources are not silently dropped. It also collects current transfers and peer outcomes from that provider, effective settings from `Get-DOConfig -Verbose`, MDM and Group Policy registry values, and Entra identity from Cloud Domain Join metadata.

## Output modes

| Command | Output |
|---------|--------|
| `.\Get-DeliveryOptimizationStatistics.ps1` | Formatted console report |
| `.\Get-DeliveryOptimizationStatistics.ps1 -AsObject` | Structured `PSCustomObject` |
| `.\Get-DeliveryOptimizationStatistics.ps1 -AsJson` | Compact single-line JSON |
| `.\Get-DeliveryOptimizationStatistics.ps1 -AsJson -Pretty` | Indented JSON for inspection |
| `.\Get-DeliveryOptimizationStatistics.ps1 -OutputPath <path>` | Selected console output plus compact JSON appended to a file |

`-OutputPath` uses UTF-8 without a BOM and appends one Unix-terminated JSON object per run, suitable for line-oriented collectors.

## Metrics

The JSON/object payload contains:

- Collection identity: `timestamp`, `hostname`, `os_build`, and `display_version`
- Entra identity: device ID, tenant ID/name, and join state when locally available
- Measurement window: UTC `period_start` / `period_end`, plus local-time equivalents and `period_timezone`
- Effective configuration: download mode/provider, bandwidth limits, upload limits, RAM/disk/battery gates, VPN behavior, working directory, and business-hour limits
- Raw MDM and Group Policy Delivery Optimization values, kept in separate source maps
- Month-to-date inclusive HTTP bytes plus mutually exclusive direct-CDN, Connected Cache, LAN, Group, Internet, and Link-Local bytes
- Month-to-date upload bytes to each peer class
- Strict WUfB-compatible and complete all-source efficiency calculations
- Current cache size, disk capacity, peer count, source connections, active/pending jobs, CPU, and memory use
- Current transfer details and peer outcome aggregates when transfers are active
- A `coverage` object that identifies report concepts unavailable from the local provider

Byte fields remain integers in object and JSON output so Grafana can aggregate them without converting formatted sizes. The console table adds GiB values for readability.

The console report uses local time because the counters reset at local calendar-month boundaries. JSON retains UTC timestamps for ingestion and adds `period_start_local`, `period_end_local`, and `period_timezone` to make that boundary explicit.

## Calculations

The local provider's `MonthlyCdnBytes` counter is exposed by Microsoft's PowerShell wrapper as `DownloadHttpBytes`. Microsoft documents this as the inclusive HTTP total, including Connected Cache. The script therefore derives direct CDN bytes before applying the mutually exclusive WUfB source formulas:

```text
HTTPBytes = MonthlyCdnBytes
CDN = max(HTTPBytes - ConnectedCache, 0)
WUfBPeerBytes = LAN + Group
WUfBTotalBytes = HTTPBytes + LAN + Group
WUfBLocalSourceBytes = ConnectedCache + LAN + Group

BandwidthSavingsPct = 100 * WUfBLocalSourceBytes / WUfBTotalBytes
P2PEfficiencyPct = 100 * WUfBPeerBytes / WUfBTotalBytes
ConnectedCacheEfficiencyPct = 100 * ConnectedCache / WUfBTotalBytes
```

The complete local totals additionally include Internet and Link-Local peers:

```text
TotalPeerBytes = LAN + Group + Internet + LinkLocal
TotalBytes = HTTPBytes + TotalPeerBytes
AllSourcesBandwidthSavingsPct = 100 * (ConnectedCache + TotalPeerBytes) / TotalBytes
AllSourcesP2PEfficiencyPct = 100 * TotalPeerBytes / TotalBytes
```

When a denominator is zero, its percentages are `0`. `BandwidthSavingsStatus` follows the report thresholds: `ERROR` below 10%, `WARNING` from 10% through 60%, `OK` above 60%, and `NO_DATA` when no bytes exist.

## Local data versus WUfB reports

This output uses WUfB field names where the local meaning is genuinely equivalent, but the data sources and time windows differ:

| Local script | Windows Update for Business reports |
|--------------|-------------------------------------|
| One device | Device and tenant cloud telemetry |
| Current local calendar month | Rolling 7-day and 28-day views |
| Delivery Optimization CIM counters | Delivery Optimization telemetry events |
| Includes explicit Internet and Link-Local counters | Uses the published `UCDOStatus` source fields and aggregation rules |
| Cache-host bytes have no ISP classification | ISP-hosted Connected Cache bytes are filtered out of `BytesFromCache` |

Do not compare the local percentage directly with `BWOptPercent28Days` unless the date window and included source classes are aligned.

## WUfB schema coverage

### `UCDOStatus`

| WUfB field | Local coverage | Script field / limitation |
|------------|----------------|---------------------------|
| `AzureADDeviceId` | Captured when Entra joined | `identity.AzureADDeviceId` |
| `AzureADTenantId`, `TenantId` | Captured when Entra joined | `identity.AzureADTenantId` |
| `BWOptPercent7Days`, `BWOptPercent28Days` | Not locally available | Rolling windows are produced from cloud telemetry |
| `BytesFromCache` | Captured month-to-date | `download.BytesFromCache`; no ISP-level MCC filtering |
| `BytesFromCDN` | Derived month-to-date | `download.BytesFromCDN` is direct CDN only: inclusive HTTP minus Connected Cache |
| `BytesFromPeers` | Captured month-to-date | `download.BytesFromPeers` means LAN peers, matching the report calculation page |
| `BytesFromGroupPeers` | Captured month-to-date | `download.BytesFromGroupPeers` |
| `BytesFromIntPeers` | Captured month-to-date | `download.BytesFromIntPeers`; excluded from strict WUfB formulas |
| `City`, `Country`, `ISP` | Not locally available | Cloud-side IP geolocation/ISP attribution |
| `ContentDownloadMode` | Current transfers only | `transfers[].DownloadModeId`; no month-to-date content grouping |
| `ContentType` | Not locally available historically | `transfers[].CallerApplication` is a current hint, not a WUfB content category |
| `DeviceName` | Captured | `identity.DeviceName` / `hostname` |
| `DOStatusDescription` | Current snapshot | `current.DOStatusDescription` and `transfers[].Status` |
| `DownloadMode` | Captured | `configuration.DownloadMode` and `DownloadModeId` |
| `DownloadModeSrc` | Captured | `configuration.DownloadModeProvider` |
| `GlobalDeviceId` | Not locally available | Microsoft-internal telemetry identifier |
| `GroupID` | Partial | Explicit `DOGroupID` and its WUfB SHA-256 hash are captured; dynamically derived effective Group IDs are not exposed locally without parsing verbose logs |
| `LastCensusSeenTime` | Not locally available | Cloud census telemetry |
| `NoPeersCount` | Current snapshot only | `current.PeerCount`, transfer peer counts, and peer-type aggregates are not the historical WUfB counter |
| `OSVersion` | Captured | `display_version`, `os_build` |
| `PeerEligibleTransfers` | Current transfers only | `transfers[].PeerEligible`; no historical counter |
| `PeeringStatus` | Derived exactly from configured mode | `configuration.PeeringStatus` |
| `PeersCannotConnectCount`, `PeersSuccessCount` | Current transfers only | `PeerConnectionsFailed` / `PeerConnectionsEstablished`; no historical counters |
| `PeersUnknownCount` | Not locally equivalent | Local peer detail exposes type and connection state, not the WUfB historical relation |
| `TimeGenerated` | Captured for this collection | `timestamp` |
| `TotalTimeForDownload` | Current transfers only | `transfers[].DownloadDurationSeconds` |
| `TotalTransfers` | Current snapshot only | `current.TransferCount`; not the historical WUfB count |
| `Type` | Local equivalent | `type = LocalDeliveryOptimizationStatistics`, deliberately not `UCDOStatus` |

### `UCDOAggregatedStatus` and report-only terms

The source-byte fields and formulas are captured for this device. `DeviceCount`, P2P device count, MCC device count, total active devices, top-ten groups, and tenant/content-type aggregations require records from multiple devices and are therefore unavailable to a local one-device script. Local booleans `PeerConfigured`, `PeerUsedThisMonth`, and `MCCUsedThisMonth` allow Grafana to calculate those counts after ingesting results from a fleet.

`download.BytesFromHTTP` preserves the provider's inclusive HTTP counter. The same distinction is present for current transfers: `transfers[].BytesFromHTTP` includes `BytesFromCache`, while `transfers[].BytesFromCDN` is the derived direct-CDN remainder.

The report's content categories are Apps, Driver Updates, Edge Updates, Feature Updates, Intune Apps, Office, Other, Quality Updates, and Teams Updates. The local month-to-date CIM counters do not retain that category dimension.

References:

- [Delivery Optimization data in reports](https://learn.microsoft.com/windows/deployment/update/wufb-reports-do)
- [UCDOStatus data schema](https://learn.microsoft.com/windows/deployment/update/wufb-reports-schema-ucdostatus)
- [UCDOAggregatedStatus data schema](https://learn.microsoft.com/windows/deployment/update/wufb-reports-schema-ucdoaggregatedstatus)

## Requirements

- Windows 10, Windows 11, or Windows Server with Delivery Optimization
- PowerShell 5.1 or later
- In-box `DeliveryOptimization` PowerShell module
- Read access to the Delivery Optimization CIM provider

The script does not require elevation on a normal Windows client. Access restrictions imposed by endpoint security controls can still cause collection to fail; the error path returns exit code `4`.
