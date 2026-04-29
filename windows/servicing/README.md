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
   - per-user `AppData\Local\Temp` (`-UserTmp`)
   - `%WinDir%\Temp` (`-WindowsTmp`)
   - `C:\Windows\SoftwareDistribution` after stopping `wuauserv` (`-SoftwareDistribution`)
3. Configures every known `cleanmgr.exe` `VolumeCaches` handler with `StateFlags<SageId>=2` and runs `cleanmgr /sagerun:<SageId>`.
4. Runs `dism /Online /Cleanup-Image /StartComponentCleanup /ResetBase` to shrink WinSxS.
5. Re-checks free space and refuses to proceed on battery power.

> **`/ResetBase` warning:** This permanently removes the ability to uninstall previously installed Windows updates. Do not run on production endpoints that need rollback capability.

### Exit codes

| Code | Meaning |
|------|---------|
| `0`    | Success / nothing to do |
| `1602` | Free space still below `-MinFreeGB` after cleanup |
| `1603` | Device on battery power (override with `-IgnoreBattery`) |

### Examples

```powershell
# Full pre-upgrade cleanup
.\Invoke-PreUpgradeCleanup.ps1 -UserTmp -WindowsTmp -SoftwareDistribution

# Reference image cleanup — force run, ignore battery (e.g. inside Image Builder VM)
.\Invoke-PreUpgradeCleanup.ps1 -UserTmp -WindowsTmp -SoftwareDistribution -Force -IgnoreBattery

# Conservative — keep last 30 days of temp files, larger free-space target
.\Invoke-PreUpgradeCleanup.ps1 -UserTmp -WindowsTmp -MaxAgeDays 30 -MinFreeGB 30
```

### Logs

Written to `$env:TEMP\Invoke-PreUpgradeCleanup_<timestamp>.log` (override with `-LogDir`).

## Requirements

- Windows 10 / 11 (or Windows Server with Desktop Experience for `cleanmgr.exe`)
- PowerShell 5.1+
- Elevated session (`#Requires -RunAsAdministrator`)
