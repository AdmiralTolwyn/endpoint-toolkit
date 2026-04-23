# W365 Disk Resizer

Expands the C: partition to consume all unallocated disk space after a Windows 365 Cloud PC resize.

## Problem

When a Windows 365 Cloud PC is resized to a larger disk, the additional space appears as unallocated. Windows does not automatically extend the OS partition, and the built-in Disk Management snap-in is not available to end users. The Recovery partition, positioned immediately after the OS partition, blocks a simple `extend` operation.

## How It Works

1. Detects disk and partition layout via WMI (`Win32_DiskDrive`, `Win32_DiskPartition`)
2. If a Recovery (WinRE) partition exists after the OS partition, disables WinRE via `reagentc /disable` and deletes the partition
3. Extends the C: volume into all available unallocated space using `diskpart`
4. Validates the resize and logs results

> **Note:** The script uses `diskpart.exe` instead of PowerShell Storage cmdlets (`Resize-Partition`, `Remove-Partition`) to avoid potential BSOD issues observed on W365 virtual disks.

## Deployment

Deploy as an **Intune platform script** (Devices → Scripts and remediations → Platform scripts):

| Setting | Value |
|---------|-------|
| Run this script using the logged on credentials | **No** (runs as SYSTEM) |
| Enforce script signature check | No |
| Run script in 64​-bit PowerShell host | **Yes** |

## Logs

Log files are written to the Intune Management Extension log folder:

```
%ProgramData%\Microsoft\IntuneManagementExtension\Logs\DiskResizer_<timestamp>.log
```

## Requirements

- Windows 10/11 (Windows 365 Cloud PC)
- PowerShell 5.1+
- SYSTEM context (via Intune platform script)
