# Processor Boost

[Configure-ProcessorBoost.ps1](Configure-ProcessorBoost.ps1) queries or configures the Windows processor performance boost mode (`PERFBOOSTMODE`) on the active power plan.

Running without parameters is read-only. Changes require an elevated PowerShell session and apply the same value to both AC and battery power.

## Requirements

- Windows 10 or Windows 11 with `powercfg.exe` and the `PERFBOOSTMODE` setting available.
- PowerShell 5.1 or later.
- Administrator rights for `-Disable`, `-Enable`, or `-Configure`. The default query does not enforce elevation.

## Usage

Run from this directory. Query the current AC and battery values and the available modes:

```powershell
.\Configure-ProcessorBoost.ps1
```

Disable processor boost for AC and battery power:

```powershell
.\Configure-ProcessorBoost.ps1 -Disable
```

Enable processor boost using mode 1:

```powershell
.\Configure-ProcessorBoost.ps1 -Enable
```

Select a specific mode, for example Aggressive (2):

```powershell
.\Configure-ProcessorBoost.ps1 -Configure 2
```

## Parameters

| Parameter | Effect |
|-----------|--------|
| None | Query the active plan without making changes. |
| `-Disable` | Set mode 0 (Disabled). |
| `-Enable` | Set mode 1 (Enabled). This does not restore the previous mode. |
| `-Configure <0-6>` | Set the specified mode, including 0. |

The three change parameters are mutually exclusive. Values outside 0 through 6 are rejected before any commands run.

## Modes

| Value | Name |
|-------|------|
| 0 | Disabled |
| 1 | Enabled |
| 2 | Aggressive |
| 3 | Efficient Enabled |
| 4 | Efficient Aggressive |
| 5 | Aggressive At Guaranteed |
| 6 | Efficient Aggressive At Guaranteed |

The query displays the modes exposed by Windows on the target device. Actual performance behavior depends on the processor and its power-management interface; mode names do not guarantee a particular clock speed or performance gain.

## Scope And Failure Handling

- Queries the hidden setting with `powercfg /qh SCHEME_CURRENT SUB_PROCESSOR PERFBOOSTMODE`; it does not unhide it in the power-plan UI.
- Changes only `PERFBOOSTMODE` for AC and battery power on the current plan, then reapplies that plan and displays the settings.
- Does not create plans, update other plans, or enforce the setting continuously. Selecting another plan can select different boost values.
- Checks the exit code after each `powercfg` command and stops on failure. Changes are sequential, not transactional: an AC change can remain applied if the battery change fails. There is no automatic rollback.
- Record the original AC and battery indices before changing them. `-Enable` always writes 1; it is not an undo operation. Different original AC and battery values must be restored individually with `powercfg`.
- Output is native `powercfg` text, not JSON or structured PowerShell objects. The final query displays the result but does not programmatically compare it with the requested value or test hardware boost behavior.

## References

- [Microsoft: PERFBOOSTMODE](https://learn.microsoft.com/en-us/windows-hardware/customize/power-settings/options-for-perf-state-engine-perfboostmode)
- [Microsoft: Powercfg command-line options](https://learn.microsoft.com/en-us/windows-hardware/design/device-experiences/powercfg-command-line-options)