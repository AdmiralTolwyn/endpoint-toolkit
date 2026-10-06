# mdmresult

gpresult for Intune and hybrid-managed Windows devices. One script and two metadata files;
copy the folder to a device, run it, get an HTML report.

The report is the same as PolicyPilot's **Export HTML**: Group Policy result and local MDM
policy state with English ADMX/CSP names and descriptions, CSP gap analysis, GPO conflicts,
Intune apps, certificates, LAPS, provisioning packages and compliance summary.

## Package

| File | Purpose |
|---|---|
| `mdmresult.ps1` | The tool (PowerShell 5.1, no modules, no GUI) |
| `admx_metadata.json` | ADMX/ADML metadata: Windows 11 26H2 + SecGuide / MSS-legacy (3,726 policies) |
| `csp_metadata.json` | Policy CSP metadata from Microsoft Learn (1,617 settings) |

`build/` is only needed to regenerate the package; don't ship it.

## Usage

```powershell
# Combined Group Policy + MDM (default); run elevated
.\mdmresult.ps1

# Intune only, fixed path, overwrite, open when done
.\mdmresult.ps1 -Mode Intune -Path C:\Temp\policy.html -Force -Open
```

| Parameter | Default | Description |
|---|---|---|
| `-Mode` | `Combined` | `Local` (Group Policy), `Intune` (MDM), `Combined` (hybrid / co-managed) |
| `-Path` (`-H`) | `.\mdmresult_<COMPUTER>_<timestamp>.html` | Report file |
| `-Force` (`-F`) | off | Overwrite an existing report |
| `-Open` | off | Open the report when finished |

A gpresult `/r`-style summary is printed to the console; `-Verbose` shows the full scan log.

Without elevation, `gpresult /scope computer` is denied (user scope is used instead), and the
MDM WMI bridge and Win32 app registry are unreadable. The report is still produced.

## Rebuilding

`mdmresult.ps1` is generated from [PolicyPilot](../PolicyPilot/). Don't edit it directly.

```powershell
.\build\Build-MdmResult.ps1
```

The build takes PolicyPilot's scan block, post-scan enrichment and gap analysis, conflict
detection and HTML report code from `PolicyPilot.ps1`, fills them into
`build/mdmresult.template.ps1`, and copies both metadata files from PolicyPilot. Run it after
any change to `PolicyPilot.ps1` or the metadata. It stops with an error if PolicyPilot's code
no longer has the expected shape, rather than generating a report that differs from the GUI.
