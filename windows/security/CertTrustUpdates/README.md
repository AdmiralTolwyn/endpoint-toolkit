# Windows Certificate Trust Updates Detection (eMudhra / Sectigo AAA)

## Overview

Read-only detection script ([Check-WindowsCertTrustUpdates.ps1](Check-WindowsCertTrustUpdates.ps1)) that evaluates whether a Windows endpoint or server is protected against the October 2026 **eMudhra intermediate CA private key compromise** and associated **Sectigo `AAA Certificate Services` cross-signed paths**.

In October 2026, eight eMudhra G3B/G3C subordinate CA private keys (six TLS, two Code Signing) were compromised. Microsoft responded in two stages:

1. **October 20, 2026 — Windows Disallowed CTL update:** Blocked the eight compromised intermediate CA keys by Subject Public Key Information (SPKI) across both direct eMudhra paths and Sectigo `AAA Certificate Services` cross-signed paths (`SequenceNumber = 01dd5f6ca45c9d86`, effective `2026-10-19 01:53:30 UTC`).
2. **October 26, 2026 — Windows Trusted Root removal & Out-of-Band (OOB) updates:** Removed six `emSign ... Root ... CA - G3` roots and the Sectigo `AAA Certificate Services` root from the default Windows trust list (`SequenceNumber = 1401dd6344d40f0a15`, effective `2026-10-23 23:18:35 UTC`), and shipped OOB cumulative updates (`KB5134134`, `KB5134133`, `KB5134130`, `KB5134137`, `KB5134140`, `KB5134125`, `KB5134124`, `KB5134123`) for disconnected or policy-restricted devices that do not receive automatic CTL updates.

References:
- [Microsoft’s response to compromised eMudhra certificate authority keys (MSRC Blog, Oct 26, 2026)](https://www.microsoft.com/en-us/msrc/blog/2026/10/microsoft-response-to-compromised-emudhra-certificate-authority-keys)
- [Microsoft Trusted Root Certificate Program — October 26, 2026 Announcement](https://github.com/TrustedRootProgram/Program-Requirements/blob/main/Announcements.md#october-26-2026)
- [Configure trusted roots and disallowed certificates in Windows (Microsoft Learn)](https://learn.microsoft.com/en-us/windows-server/identity/ad-cs/configure-trusted-roots-disallowed-certificates)

---

## What the Script Evaluates

The script is strictly read-only, runs 100% locally by default (no temp files, no child processes, no outbound network traffic unless `-TestOnline` is specified), and evaluates five signals:

1. **OS Build / UBR & Hotfix state (`Get-OobUpdateStatus`):**  
   Checks both `CurrentBuildNumber.UBR` against the minimum revision for the October 26, 2026 OOB updates and `Get-HotFix` for the OOB and subsequent October 28, 2026 Preview cumulative updates:

   | OS Version | Build(s) | OOB KB (Oct 26, 2026) | Preview KB (Oct 28, 2026) | Minimum UBR |
   |---|---|---|---|---|
   | Windows 11, version 26H1 | `28000` | [KB5134134](https://support.microsoft.com/en-au/servicing/os/windows-11/2026/10/kb5134134-windows-11-26h1-update) (`28000.3237`) | [KB5127653](https://support.microsoft.com/en-au/servicing/os/windows-11/2026/10/kb5127653-windows-11-26h1-update) (`28000.3365`) | `3237` |
   | Windows 11, versions 25H2 & 26H2 | `26200`, `26300` | [KB5134133](https://support.microsoft.com/en-us/servicing/os/windows-11/2026/10/kb5134133-windows-11-24h2-25h2-26h2-update) (`.9725`) | [KB5127936](https://support.microsoft.com/en-au/servicing/os/windows-11/2026/10/kb5127936-windows-11-25h2-26h2-update) (`.9818`) | `9725` |
   | Windows 11, version 24H2 | `26100` (Client) | [KB5134133](https://support.microsoft.com/en-us/servicing/os/windows-11/2026/10/kb5134133-windows-11-24h2-25h2-26h2-update) (`26100.9725`) | — | `9725` |
   | Windows 11, version 23H2 | `22621`, `22631` | [KB5134130](https://support.microsoft.com/en-au/servicing/os/windows-11/2026/10/kb5134130-windows-11-23h2-update) (`.7654`) | — | `7654` |
   | Windows Server 2025 | `26100` (Server) | [KB5134137](https://support.microsoft.com/en-us/servicing/os/windows-server/2026/10/kb5134137-windows-server-2025-update) (`26100.33674`) | — | `33674` |
   | Windows Server 2022 | `20348` | [KB5134140](https://support.microsoft.com/en-au/servicing/os/windows-server/2026/10/kb5134140-windows-server-2022-update) (`20348.5773`) | — | `5773` |
   | Windows 10 21H2 / 22H2 / LTSC 2021 | `19044`, `19045` | [KB5134125](https://support.microsoft.com/en-au/servicing/os/windows-10/2026/10/kb5134125-windows-10-21h2-22h2-update) (`.7811`) | — | `7811` |
   | Windows 10 1809 LTSC 2019 / Server 2019 | `17763` | [KB5134124](https://support.microsoft.com/en-us/servicing/os/windows-10/2026/10/kb5134124-windows-10-1809-update) (`17763.9378`) | — | `9378` |
   | Windows 10 1607 LTSB 2016 / Server 2016 | `14393` | [KB5134123](https://support.microsoft.com/en-us/servicing/os/windows-10/2026/10/kb5134123-windows-10-1607-update) (`14393.9615`) | — | `9615` |

2. **Automatic CTL Update Policy:**  
   Reads `HKLM:\SOFTWARE\Policies\Microsoft\SystemCertificates\AuthRoot` (`DisableRootAutoUpdate = 1` or `EnableDisallowedCertAutoUpdate = 0`) to flag devices where automatic CTL synchronization is disabled by policy.
3. **Local Disallowed CTL (`DisallowedCertEncodedCtl`):**  
   Decodes the PKCS#7 CTL in memory using `System.Security.Cryptography.Pkcs.SignedCms` and verifies `EffectiveUtc >= 2026-10-19 01:53:30Z` (`SequenceNumber = 01dd5f6ca45c9d86`). Satisfied when either the synced Disallowed CTL is current or the October 26+ OOB/CU update is installed.
4. **Local AuthRoot CTL (`EncodedCtl`):**  
   Decodes the PKCS#7 AuthRoot CTL in memory and verifies none of the 7 removed root SHA-256 fingerprints remain in the active list (`CERT_AUTH_ROOT_SHA256_HASH_PROP_ID`). Satisfied when either the synced AuthRoot CTL is updated or the October 26+ OOB/CU update is installed.
5. **Local Certificate Store Audit (`LocalMachine\Root`, `AuthRoot`, `TrustedPublisher`, `CA`):**  
   Computes SHA-256 over `RawData` for every certificate in the local machine stores (because PowerShell's `X509Certificate2.Thumbprint` property is SHA-1, whereas MSRC publishes SHA-256 fingerprints). Flags the device as `VULNERABLE` if any of the 7 removed roots or 16 revoked intermediates remain cached in `Root`, `AuthRoot`, or `TrustedPublisher` without a matching entry in `LocalMachine\Disallowed`.

---

## Removed Root Certificates Checked (SHA-256 & SHA-1)

| Certificate | SHA-256 Thumbprint | SHA-1 Thumbprint |
|---|---|---|
| `emSign ECC Root CA - G3` | `86A1ECBA089C4A8D3BBE2734C612BA341D813E043CF9E8A862CD5C57A36BBE6B` | `3043FA4FF257DCA0C380EE2E58EA78B23FE6BBC1` |
| `emSign Root Client Auth CA - G3` | `BBD8B3D745EFB79198D93B9C482B9D934537325729EA2F02C129261F91BE2E3D` | `3E48813E1F94CE3ACB726F7E897FD9D868F4EC15` |
| `emSign Root CS CA - G3` | `F2EDAAFF3D9B701FB4D14C9F285F8C0BAD90FC3BCC1FA07D47F4718EDABA34A5` | `1F74B6010A219FD9CA058D07AA053D15EFD9300D` |
| `emSign Root SMIME CA - G3` | `F76A0D440C418DC5B362BE8E11E0A5D06023F4396093E4F455827BF55153D8A4` | `9A839AFC3549E4C1DFA89863D28E536C8EADC608` |
| `emSign Root TLS CA - G3` | `7DD78D5F4F13459A83DFF9ABBBA62EDBAF6F2D102BF257FD712F4D9F2746ED8D` | `39D3236E4A609AB6190CBFA1AD02B3CC61DFC2C3` |
| `emSign Root TSA CA - G3` | `75D2CE75979CF2A43B172C9F89F733DDA2136031DB4BBF3698B053714C077155` | `49A2493755861F9AE1AFB9438E3519BA9135A75F` |
| `AAA Certificate Services (Sectigo)` | `D7A7A0FB5D7E2731D771E9484EBCDEF71D5F0C3E0A2948782BC83EE0EA699EF4` | `D1EB23A46D17D68FD92564C2F1F1601764D8E349` |

---

## Usage & Output Contract

```powershell
# Default local detection (no network requests)
.\Check-WindowsCertTrustUpdates.ps1

# Include live HTTP HEAD probe to ctldl.windowsupdate.com
.\Check-WindowsCertTrustUpdates.ps1 -TestOnline
```

### Exit Codes & STDOUT

- **Exit `0` (`NOT-VULNERABLE`):**
  ```text
  NOT-VULNERABLE: October 2026 Disallowed CTL (Seq=01dd5f6ca45c9d86, Date=2026-10-19) and AuthRoot CTL synced; no removed roots trusted locally | Host=CLIENT01 OS=26200.9711 OobKB=KB5134133:False HotFix=None AutoUpdateBlocked=False DisallowedSeq=01dd5f6ca45c9d86 DisallowedDate=2026-10-19 DisallowedOk=True AuthRootDate=2026-10-23 CtlRetainedRoots=0 TrustedRemovedCerts=0 CachedRevokedCA=0
  ```
- **Exit `1` (`VULNERABLE`):**
  ```text
  VULNERABLE: Disallowed CTL is outdated/missing (Seq=None) and automatic CTL updates are disabled by policy; install KB5134133 | Host=CLIENT02 OS=26200.9711 OobKB=KB5134133:False HotFix=None AutoUpdateBlocked=True DisallowedSeq=None DisallowedDate=None DisallowedOk=False AuthRootDate=None CtlRetainedRoots=0 TrustedRemovedCerts=1 CachedRevokedCA=0
  ```

### Logging

Writes CMTrace-formatted logs to:
- `%ProgramData%\Microsoft\IntuneManagementExtension\Logs\PR_WindowsCertTrustUpdates.log`
- Falls back automatically to `%TEMP%\PR_WindowsCertTrustUpdates.log` when executed without administrative privileges.
