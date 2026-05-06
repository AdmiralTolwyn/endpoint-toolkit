# Windows 365

Scripts for Windows 365 Cloud PC provisioning and day-to-day operations, deployed via Intune platform scripts.

## Scripts

| Script | Purpose |
|--------|---------|
| [ExpandOSDisk/](ExpandOSDisk/README.md) | Expands the C: partition to consume all unallocated disk space after a Windows 365 Cloud PC resize (`diskpart`-based, avoids BSOD with Storage cmdlets). |
| [KeyboardLayout/](KeyboardLayout/README.md) | Adds a keyboard layout (default: Swiss German) and sets it as the default input method for the current user, Default User profile, and system (welcome screen) via `intl.cpl` InputPreferences XML. |
