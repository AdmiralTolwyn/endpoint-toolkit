# NTLM Usage Detection For Ivanti

`Get-NtlmUsageDetection-Ivanti.ps1` collects local NTLM Operational and Security event evidence and reports audit visibility. It is a read-only detector, not a remediation tool: it does not enable auditing, block NTLM, alter policy or clear event logs. It appends a local diagnostic log by default; JSON export is optional.

Run locally in **64-bit Windows PowerShell 5.1** as LocalSystem or an elevated administrator. Standard-user runs can lack access to the Security log and audit policy.

```powershell
.\Get-NtlmUsageDetection-Ivanti.ps1
.\Get-NtlmUsageDetection-Ivanti.ps1 -Detailed -Days 7
.\Get-NtlmUsageDetection-Ivanti.ps1 -Days 30 -ReportPath C:\Temp\NtlmEvidence.json
.\Get-NtlmUsageDetection-Ivanti.ps1 -CmdLine -LogPath ''
Get-Help .\Get-NtlmUsageDetection-Ivanti.ps1 -Full
```

The default Ivanti output is exactly four `Write-Host` lines: `detected`, `reason`, `expected`, and `found`. Do not use `-Detailed` or its alias `-CmdLine` in an Ivanti definition expecting that contract. `detected=true` means matching NTLM or NTLMv1-derived credential records were observed; audit gaps, access failures and file errors alone do not set it. A `false` result **does not certify that NTLM is unused or safe to disable**. Check `Collection diagnostics` in the output, the append-only log, or the optional JSON report before interpreting a negative result.

| Parameter | Default | Purpose |
|-----------|---------|---------|
| `Days` | 14 | Lookback window; retained history may be shorter. |
| `MaxEventsPerSource` | 10000 | Per-query record cap; truncation is reported. |
| `MaxEvidence` | 200 | Maximum recent records retained in the report; detailed console output shows at most three. |
| `Detailed` / `CmdLine` | Off | Replace Ivanti output with a compact console report. |
| `ReportPath` | None | Optional JSON report; parent directory must exist and an existing file is overwritten. |
| `LogPath` | `%windir%\Temp\Get-NtlmUsageDetection-Ivanti.log` | Append diagnostic log; set to `''` to disable it. |

The detector reads NTLM Operational audit, block and enhanced events, Security logon/credential-validation records, audit-policy and retention state. Records are evidence of attempts, blocks, or logons according to their fields, not proof that every authentication succeeded. The report may contain account names, hosts, addresses, and process paths; restrict access to its output and configure log retention. The script makes no network requests. The exact Ivanti execution host, output parser and limits require a pilot before deployment.

Run the offline synthetic-event suite from this directory (no authentication settings or real events are changed):

```powershell
.\Test-NtlmUsageDetection.ps1
```

The suite verifies event classification, audit/collection gaps, four-line output, optional JSON and logging behavior on Windows PowerShell 5.1. A real elevated or LocalSystem pilot is still needed to validate event availability in your environment.