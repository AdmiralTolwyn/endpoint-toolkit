# Windows Servicing

Scripts that prepare a device (or reference image) for Windows feature updates and component-store maintenance, or analyse the result afterwards.

## Scripts

| Script | Purpose |
|--------|---------|
| [PreUpgradeCleanup/](PreUpgradeCleanup/README.md) | Reclaims disk space prior to a Windows feature update or after a reference image build (cleanmgr handlers, DISM `/StartComponentCleanup`, optional WU reset). |
| [SetupTimeline/](SetupTimeline/README.md) | Reconstructs the phase-by-phase timeline of a completed (or rolled-back) Windows in-place upgrade from `setupact.log`, separating active upgrade time from idle (powered-off / standby / sleep) windows. |
| [StartupAppsDelay/](StartupAppsDelay/README.md) | Reverts the Windows 11 "wait-for-idle" startup-app delay (`HKCU\...\Explorer\Serialize\WaitForIdleState=0`) so Outlook / Teams / Word / Excel launch promptly after sign-in on busy devices. |