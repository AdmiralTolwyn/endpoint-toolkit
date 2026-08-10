<#
.SYNOPSIS
    Segments a Secure Boot 2023 fleet export into actionable populations.

.DESCRIPTION
    Consumes either:
      - the CSV exported from Intune > Devices > Remediations > <package> > Monitor > Device status
        (requires the "Pre-remediation detection output" column to be added before export), or
      - a folder/file of raw JSON lines produced by Get-SecureBootCertInventory.ps1.

    Splits the fleet into the buckets that actually drive different remediation work, rather than
    the report's device-level Up to date / Not up to date / Unknown split which mixes
    "certificates missing" with "we have no data".

    Segments, evaluated in order (first match wins):

      SecureBootOff         Secure Boot disabled. Out of scope - certificates cannot be written.
      ReportingBlocked      Device has data locally but cannot report it: Secure-Boot-Update task
                            not Ready, AllowTelemetry < 1, or DisableOneSettingsDownloads = 1.
                            This is the population behind most of the "Unknown" bucket.
      Updated               UEFICA2023Status = Updated. No action.
      OptionRomNotApplicable Microsoft-only trust and the only "missing" certs are the
                            non-Microsoft ones (Option ROM / Microsoft UEFI CA 2023). These are
                            false positives against a hand-rolled cert-presence check - the
                            0x4000 modifier bit only applies them when Microsoft Corporation
                            UEFI CA 2011 is present in db.
      VirtualVMware         VMware guest not updated. Remediation is hypervisor-side; there is no
                            Microsoft-side patch listed for VMware. Track against Broadcom KB 423893.
      VirtualOtherHV        Hyper-V / Azure / other hypervisor guest not updated. Hyper-V KEK
                            failures (1795) were fixed in Mar 2026 (Apr 2026 for Server 2025) and
                            need the fix on BOTH host and guest. Azure Trusted Launch 1795 on KEK
                            is a live known issue with no customer action.
      FirmwareBlocked       Physical device stalled by firmware/platform: event 1795/1802/1803, or
                            confidence Temporarily Paused / Not Supported. Needs an OEM firmware
                            update, or is a documented exception.
      PendingRestart        InProgress and AvailableUpdates has cleared down to 0x4100 - only the
                            boot manager step is left, which lands on the next natural restart.
      NotTargeted           Servicing key absent / NotStarted with healthy prerequisites. Nothing
                            is wrong; the device simply has not been targeted or picked up by a
                            high-confidence bucket yet.
      NeedsInvestigation    Anything else.

.PARAMETER Path
    CSV exported from Intune, or a .json/.txt file of one JSON object per line, or a folder of those.

.PARAMETER OutputCsv
    Optional path to write the segmented result.

.EXAMPLE
    .\Split-SecureBootPopulation.ps1 -Path .\DeviceStatus.csv -OutputCsv .\segmented.csv

.NOTES
    PowerShell 5.1. Read-only against the export; touches no devices.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Path,

    [string]$OutputCsv
)

$ErrorActionPreference = 'Stop'

function Get-DetectionRecords {
    param([string]$InputPath)

    $items = @()
    if (Test-Path $InputPath -PathType Container) {
        $files = Get-ChildItem -Path $InputPath -File -Include *.csv, *.json, *.txt -Recurse
    }
    else {
        $files = @(Get-Item -Path $InputPath)
    }

    foreach ($f in $files) {
        if ($f.Extension -eq '.csv') {
            $rows = Import-Csv -Path $f.FullName
            foreach ($row in $rows) {
                # Intune's column name has varied; accept the common spellings.
                $raw = $null
                foreach ($col in 'Pre-remediation detection output', 'PreRemediationDetectionScriptOutput', 'DetectionOutput') {
                    if ($row.PSObject.Properties.Name -contains $col -and $row.$col) { $raw = $row.$col; break }
                }
                if (-not $raw) { continue }
                $obj = $null
                try { $obj = $raw | ConvertFrom-Json } catch { continue }
                if ($row.PSObject.Properties.Name -contains 'Device name' -and $row.'Device name') {
                    Add-Member -InputObject $obj -NotePropertyName 'IntuneDeviceName' -NotePropertyValue $row.'Device name' -Force
                }
                $items += $obj
            }
        }
        else {
            foreach ($line in (Get-Content -Path $f.FullName)) {
                if ([string]::IsNullOrWhiteSpace($line)) { continue }
                try { $items += ($line | ConvertFrom-Json) } catch { continue }
            }
        }
    }
    return $items
}

function Get-Prop {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    if ($Object.PSObject.Properties.Name -contains $Name) { return $Object.$Name }
    return $null
}

function Get-Segment {
    param($R)

    $sbOn  = Get-Prop $R 'sbOn'
    $st    = Get-Prop $R 'st'
    $task  = Get-Prop $R 'task'
    $tel   = Get-Prop $R 'tel'
    $oneS  = Get-Prop $R 'oneS'
    $trust = Get-Prop $R 'trust'
    $plat  = Get-Prop $R 'plat'
    $conf  = Get-Prop $R 'conf'
    $avl   = Get-Prop $R 'avl'
    $evt   = Get-Prop $R 'evt'

    if ($sbOn -ne $true) { return 'SecureBootOff' }

    $prereqBad = ($task -ne 'Ready') -or ($null -eq $tel) -or ($tel -lt 1) -or ($oneS -eq 1)
    if ($prereqBad) { return 'ReportingBlocked' }

    if ($st -eq 'Updated') { return 'Updated' }

    # Microsoft-only trust: the Option ROM and Microsoft UEFI CA 2023 entries are not applicable.
    # If the Windows UEFI CA 2023 and the KEK are present, the device is effectively done.
    if ($trust -eq 'MSOnly' -and (Get-Prop $R 'dbWin23') -eq $true -and (Get-Prop $R 'kek23') -eq $true) {
        return 'OptionRomNotApplicable'
    }

    $hasFwEvent = $false
    if ($evt) {
        foreach ($id in '1795', '1802', '1803', '1032') {
            if ($evt.PSObject.Properties.Name -contains $id) { $hasFwEvent = $true }
        }
    }
    $confBlocked = ($conf -match 'Temporarily Paused|Not Supported')

    if ($plat -eq 'VMware') { return 'VirtualVMware' }
    if ($plat -eq 'HyperV' -or $plat -eq 'Azure' -or $plat -eq 'OtherVM') { return 'VirtualOtherHV' }

    if ($hasFwEvent -or $confBlocked) { return 'FirmwareBlocked' }

    if ($st -eq 'InProgress' -and $avl -eq '0x4100') { return 'PendingRestart' }
    if ($st -eq 'NoValue' -or $st -eq 'NotStarted') { return 'NotTargeted' }

    return 'NeedsInvestigation'
}

$records = Get-DetectionRecords -InputPath $Path
if (-not $records -or $records.Count -eq 0) {
    throw "No parseable detection records found in '$Path'."
}

$result = foreach ($r in $records) {
    [PSCustomObject][ordered]@{
        Device     = $(if (Get-Prop $r 'IntuneDeviceName') { Get-Prop $r 'IntuneDeviceName' } else { Get-Prop $r 'host' })
        Segment    = Get-Segment $r
        Status     = Get-Prop $r 'st'
        SecureBoot = Get-Prop $r 'sbOn'
        Trust      = Get-Prop $r 'trust'
        Platform   = Get-Prop $r 'plat'
        Kek2023    = Get-Prop $r 'kek23'
        DbWin2023  = Get-Prop $r 'dbWin23'
        DbOptRom23 = Get-Prop $r 'dbOr23'
        Available  = Get-Prop $r 'avl'
        Confidence = Get-Prop $r 'conf'
        Task       = Get-Prop $r 'task'
        Telemetry  = Get-Prop $r 'tel'
        OneSettings= Get-Prop $r 'oneS'
        LastEvent  = Get-Prop $r 'lastEvt'
        SkipReason = Get-Prop $r 'skip'
        Mfr        = Get-Prop $r 'mfr'
        Model      = Get-Prop $r 'model'
        Firmware   = Get-Prop $r 'fw'
    }
}

$result | Group-Object Segment | Sort-Object Count -Descending |
    Select-Object @{ n = 'Segment'; e = { $_.Name } }, Count | Format-Table -AutoSize

if ($OutputCsv) {
    $result | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
    Write-Host "Wrote $($result.Count) rows to $OutputCsv"
}
else {
    $result
}
