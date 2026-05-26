# Resize-RecoveryPartition.ps1

Resizes the Windows Recovery Environment (WinRE) partition on the system disk so a larger `WinRE.WIM` can be staged. Originally written to remediate undersized recovery partitions that block the [KB5034441](https://support.microsoft.com/en-us/topic/kb5034441-windows-recovery-environment-update-for-windows-10-version-21h2-and-22h2-january-9-2024-62c04204-aaa5-4fee-a02a-2fdea17075a8) / CVE-2024-20666 servicing-stack update.

## Behavior

1. Pre-flight gate. Returns an exit code without touching disk if:
   - BitLocker on the OS volume is **off** (`1598`).
   - `%SystemRoot%\System32\Recovery\ReAgent.xml` is missing (`3`).
   - WinRE already points at a live partition (`10`, treated as healthy).
   - OS partition is more than 95 % full (`122`).
   - The completion marker is already stamped (`1`, idempotent skip).
2. Reads WinRE metadata from `ReAgent.xml` + `Get-WindowsImage` to size the target partition:
   - **WIM < `-LargeWimThresholdMB` (default `900` MB)** &rarr; partition sized to `-MinRecoveryPartitionMB` (default `998` MB).
   - **WIM &ge; `-LargeWimThresholdMB`** &rarr; partition sized to `WIM + -HeadroomMB` (default `550` MB).
3. Disables WinRE (`reagentc /disable`), drops the existing recovery partition, shrinks the OS partition by the target size, creates a new partition at the end of the disk using maximum size, and formats it NTFS with the label `Recovery`.
4. Tags the new partition as a WinRE GPT partition (`DE94BBA4-06D1-4D40-A16A-BFD50179D6AC`, attribute `0x8000000000000001`) via `diskpart.exe`. `Set-Partition -GptType` is **not** used: it was observed to BSOD on certain firmware (kept from the legacy script's 1.2 history).
5. Re-enables WinRE (`reagentc /enable`) so the staged WIM is registered.
6. Stamps `HKLM:\SOFTWARE\EndpointToolkit\WinRE!WinReResized = 1 (DWord)` (configurable via `-CompletionRegistryPath` / `-CompletionRegistryValueName`).

> **This script is destructive.** It drops the recovery partition and shrinks the OS partition in place. BitLocker must remain enabled throughout (the script bails out if it's off). Validate against a test device before deploying to production.

> **Single-disk assumption.** The recovery partition is searched on `-DiskNumber` only (default `0`). The OS partition is detected via `IsBoot=True` on the same disk. Multi-disk OEM layouts with the recovery partition on a different disk are not supported and will exit with `2` (ERROR_NO_WINRE_DETECTED).

## Exit codes

| Code   | Constant                              | Meaning                                                                |
|--------|---------------------------------------|------------------------------------------------------------------------|
| `0`    | `ERROR_SUCCESS`                       | Success / nothing to do                                                |
| `1`    | `ERROR_ALREADY_PATCHED`               | Registry marker present (rewritten to `0` by default)                  |
| `2`    | `ERROR_NO_WINRE_DETECTED`             | No WinRE on this machine (rewritten to `0` by default)                 |
| `3`    | `ERROR_XML_MISSING`                   | `ReAgent.xml` not found                                                |
| `4`    | `ERROR_WIM_MISSING`                   | `WinRE.wim` not found at the expected path                             |
| `10`   | `ERROR_WINRE_HEALTHY`                 | Partition already meets the target size (rewritten to `0` by default)  |
| `11`   | `ERROR_MULTIPLE_RECOVERY_PARTITIONS`  | More than one recovery partition on the disk; manual cleanup required  |
| `122`  | `ERROR_INSUFFICIENT_DISK_SPACE`       | System drive &gt; 95 % used                                            |
| `1598` | `ERROR_BDE_MISSING`                   | BitLocker protection is off on the OS volume                           |
| `1599` | `ERROR_PREREQ_FAILURE`                | Generic pre-flight failure                                             |
| `1689` | `ERROR_TPM_UPDATE_FAILED`             | `reagentc /enable` failed                                              |
| `1000` | `ERROR_UNKNOWN`                       | Uncaught exception                                                     |

Exit-code rewrites (`-RewriteAlreadyPatchedToSuccess`, `-RewriteNoWinReToSuccess`, `-RewriteWinReHealthyToSuccess`) default to `$true` so Intune / SCCM / ConfigMgr do not mark idempotent skips as failures. Pass `:$false` to surface the raw code (useful for detection scripts).

## Examples

```powershell
# Normal run, default sizing, default registry namespace
.\Resize-RecoveryPartition.ps1

# Custom registry namespace (mirror a vendor convention)
.\Resize-RecoveryPartition.ps1 -CompletionRegistryPath 'SOFTWARE\Contoso\WinRE'

# Headroom override for a known-large WinRE.WIM (e.g. 23H2 + WinPE drivers)
.\Resize-RecoveryPartition.ps1 -LargeWimThresholdMB 900 -HeadroomMB 800

# Intune detection script: surface raw exit codes
.\Resize-RecoveryPartition.ps1 -RewriteAlreadyPatchedToSuccess:$false `
                               -RewriteNoWinReToSuccess:$false `
                               -RewriteWinReHealthyToSuccess:$false

# Non-default disk (rare; only for multi-disk OEM rigs)
.\Resize-RecoveryPartition.ps1 -DiskNumber 1
```

## Parameter quick reference

| Parameter                          | Default                                        | Notes                                                                                  |
|-----------------------------------:|-----------------------------------------------:|----------------------------------------------------------------------------------------|
| `-DiskNumber`                      | `0`                                            | Physical disk holding the system + recovery partitions.                                |
| `-MinRecoveryPartitionMB`          | `998`                                          | Target size when the WIM is small.                                                     |
| `-LargeWimThresholdMB`             | `900`                                          | WIM size at which the partition is grown to WIM + headroom.                            |
| `-HeadroomMB`                      | `550`                                          | Free space above WIM size when growing for a large WIM.                                |
| `-CompletionRegistryPath`          | `SOFTWARE\EndpointToolkit\WinRE`               | Without the `HKLM:\` prefix.                                                           |
| `-CompletionRegistryValueName`     | `WinReResized`                                 | DWord, value `1` on completion.                                                        |
| `-LogDirectory`                    | `%SystemDrive%\Windows\debug`                  | Log file: `ResizeRecoveryPartition.log`.                                               |
| `-RewriteAlreadyPatchedToSuccess`  | `$true`                                        | Rewrite exit `1` &rarr; `0`.                                                           |
| `-RewriteNoWinReToSuccess`         | `$true`                                        | Rewrite exit `2` &rarr; `0`.                                                           |
| `-RewriteWinReHealthyToSuccess`    | `$true`                                        | Rewrite exit `10` &rarr; `0`.                                                          |

## Requirements

- PowerShell 5.1+
- Elevated session (the script declares `#Requires -RunAsAdministrator`).
- `reagentc.exe`, `diskpart.exe`, `Get-Disk`, `Get-Partition`, `Resize-Partition`, `New-Partition`, `Format-Volume`, `Get-WindowsImage`, `Get-BitLockerVolume`.
- BitLocker must be **on** for the OS volume.
