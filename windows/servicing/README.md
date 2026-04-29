# Windows Servicing

Scripts that prepare a device (or reference image) for Windows feature updates and component-store maintenance.

## Scripts

| Script | Purpose |
|--------|---------|
| [Invoke-PreUpgradeCleanup.ps1](Invoke-PreUpgradeCleanup.ps1) | Reclaims disk space prior to a Windows feature update or after a reference image build — purges per-user `%TEMP%`, `%WinDir%\Temp`, `SoftwareDistribution`, runs `cleanmgr /sagerun`, and `dism /StartComponentCleanup /ResetBase`. |

## Invoke-PreUpgradeCleanup.ps1

### Behavior

1. Skips work entirely if `SystemDrive` already has more than `-MinFreeGB` (default `20`) free, unless `-Force` is supplied.
2. Optionally cleans:
   - per-user `AppData\Local\Temp` (`-IncludeUserTemp`)
   - `%WinDir%\Temp` (`-IncludeWindowsTemp`)
   - `C:\Windows\SoftwareDistribution` after stopping `wuauserv` (`-IncludeSoftwareDistribution`)
3. Configures every known `cleanmgr.exe` `VolumeCaches` handler with `StateFlags<SageId>=2` and runs `cleanmgr /sagerun:<SageId>`. The handler set can be narrowed with `-IncludeOnlyHandler` or `-ExcludeHandler` (case-insensitive, exact subkey-name match; unknown names log a warning).
4. Runs `dism /Online /Cleanup-Image /StartComponentCleanup /ResetBase` to shrink WinSxS.
5. Re-checks free space and refuses to proceed on battery power.

> **`/ResetBase` warning:** This permanently removes the ability to uninstall previously installed Windows updates. Do not run on production endpoints that need rollback capability.

> **`-IncludeSoftwareDistribution` warning:** This is a **Windows Update reset**, not routine cleanup. It forces a full WU metadata re-sync on the next scan, discards any partially staged update payloads (the next pass re-downloads multi-GB ESDs), wipes the Update history UI, and on WSUS-managed devices breaks reporting until the client re-handshakes. Use it only when WU is broken (0x8024xxxx, stuck scans, corrupt BITS queue), when prepping a sysprep'd reference image, or when the disk is critically full and `cleanmgr` + `dism` cleanup did not free enough space.

### Exit codes

| Code | Meaning |
|------|---------|
| `0`    | Success / nothing to do |
| `1602` | Free space still below `-MinFreeGB` after cleanup |
| `1603` | Device on battery power (override with `-IgnoreBattery`) |

### Examples

```powershell
# Full pre-upgrade cleanup
.\Invoke-PreUpgradeCleanup.ps1 -IncludeUserTemp -IncludeWindowsTemp -IncludeSoftwareDistribution

# Reference image cleanup — force run, ignore battery (e.g. inside Image Builder VM)
.\Invoke-PreUpgradeCleanup.ps1 -IncludeUserTemp -IncludeWindowsTemp -IncludeSoftwareDistribution -Force -IgnoreBattery

# Conservative — keep last 30 days of temp files, larger free-space target
.\Invoke-PreUpgradeCleanup.ps1 -IncludeUserTemp -IncludeWindowsTemp -MaxAgeDays 30 -MinFreeGB 30

# Skip the Downloads folder + Previous Installations rollback for a feature-update prep
.\Invoke-PreUpgradeCleanup.ps1 -IncludeUserTemp -IncludeWindowsTemp -ExcludeHandler 'DownloadsFolder','Previous Installations'

# Surgical run: only Update Cleanup + WER, force regardless of free space
.\Invoke-PreUpgradeCleanup.ps1 -IncludeOnlyHandler 'Update Cleanup','Windows Error Reporting Files' -Force
```

### Logs

Written to `$env:TEMP\Invoke-PreUpgradeCleanup_<timestamp>.log` (override with `-LogDirectory`).

## Requirements

- Windows 10 / 11 (or Windows Server with Desktop Experience for `cleanmgr.exe`)
- PowerShell 5.1+
- Elevated session (`#Requires -RunAsAdministrator`)

## References

Authoritative Microsoft Learn documentation for the underlying tools and registry surface this script drives:

| Topic | Link |
|---|---|
| `cleanmgr.exe` syntax, `/sageset:n`, `/sagerun:n` | [learn.microsoft.com/.../windows-commands/cleanmgr](https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/cleanmgr) |
| `DISM /Cleanup-Image /StartComponentCleanup [/ResetBase]` | [learn.microsoft.com/.../dism-operating-system-package-servicing-command-line-options](https://learn.microsoft.com/en-us/windows-hardware/manufacture/desktop/dism-operating-system-package-servicing-command-line-options) |
| Determine the actual size of WinSxS / when to run `/ResetBase` | [learn.microsoft.com/.../determine-the-actual-size-of-the-winsxs-folder](https://learn.microsoft.com/en-us/windows-hardware/manufacture/desktop/determine-the-actual-size-of-the-winsxs-folder) |
| Reduce the size of the Component Store (Windows Server) | [learn.microsoft.com/.../reduce-size-of-component-store-in-windows](https://learn.microsoft.com/en-us/troubleshoot/windows-server/deployment/reduce-size-of-component-store-in-windows) |
| Free up drive space in Windows | [support.microsoft.com/.../free-up-drive-space-in-windows](https://support.microsoft.com/en-us/windows/free-up-drive-space-in-windows-85529ccb-c365-490d-b548-831022bc9b32) |
| Windows Update troubleshooting (when `-IncludeSoftwareDistribution` is justified) | [learn.microsoft.com/.../windows-update-issues-troubleshooting](https://learn.microsoft.com/en-us/troubleshoot/windows-client/installing-updates-features-roles/windows-update-issues-troubleshooting) |

> **Tip:** To check when `/ResetBase` last ran on a device, read `LastResetBase_UTC` under `HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing` (per the DISM doc above).
