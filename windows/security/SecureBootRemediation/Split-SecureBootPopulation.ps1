<#
.SYNOPSIS
    Segments a Secure Boot 2023 fleet export into actionable populations.

.DESCRIPTION
    Consumes either:
      - the CSV exported from Intune > Devices > Remediations > <package> > Monitor > Device status
        (requires the "Pre-remediation detection output" column to be added before export), or
      - a folder/file of raw JSON lines produced by Get-SecureBootCertInventory.ps1 or
        key=value summary lines produced by Detect_VMwareSecureBootUEFICA2023.ps1.

    Splits the fleet into the buckets that actually drive different remediation work, rather than
    the report's device-level Up to date / Not up to date / Unknown split which mixes
    "certificates missing" with "we have no data".

    Segments, evaluated in order (first match wins):

      CollectionError         Secure Boot state, platform, or UEFI variables could not be read.
      SecureBootOff           Secure Boot confirmed disabled or unsupported. Out of scope.
      Updated                 UEFICA2023Status = Updated, Error = 0, and live NVRAM has KEK 2023 +
                              required 2023 DB certs. No action.
      OptionRomNotApplicable  Microsoft-only trust and the only "missing" certs are the
                              non-Microsoft ones (Option ROM / Microsoft UEFI CA 2023).
      VMwareCapsuleReboot     VMware guest with Platform Key Firmware PnP error code 14
                              (awaiting guest restart to complete Capsule PK update).
      VMwareCapsuleError      VMware guest with Platform Key Firmware PnP error code other than 0/14.
      VMwareOsServicing       VMware guest where KEK 2023 is ALREADY in NVRAM (even if PK is NullPK,
                              per Microsoft AMA); PK update is not required, only OS servicing.
      VMwarePkReadyNeedsKek   VMware guest missing KEK 2023, but Windows OEM Devices PK is present;
                              ready for OS-side Secure-Boot-Update servicing.
      VMwarePkBlockedVtpm     VMware guest missing KEK 2023 with NullPK/None and vTPM present
                              (Capsule PK on ESX 9.1.1.0+ / VMTools 13.1.5+ / July 2026+ CU).
      VMwarePkBlockedNoVtpm   VMware guest missing KEK 2023 with NullPK/None and no vTPM
                              (SilentPK guest restart on ESXi 8.0 U3j P09+ or ESX 9.1.1.0+).
      VMwarePkBlocked         VMware guest missing KEK 2023 with NullPK/None and unknown vTPM state.
      VMwarePkUnrecognized    VMware guest missing KEK 2023 with custom/unrecognized PK.
      VirtualVMware           VMware guest without detailed PK/KEK NVRAM fields (legacy export).
      ReportingBlocked        Device has data locally but cannot report it: Secure-Boot-Update task
                              not Ready/Running, AllowTelemetry < 1, or DisableOneSettingsDownloads = 1.
      VirtualOtherHV          Hyper-V / Azure / other hypervisor guest not updated.
      FirmwareBlocked         Physical device stalled by firmware/platform: event 1795/1802/1803, or
                              confidence Temporarily Paused / Not Supported.
      PendingRestart          InProgress and AvailableUpdates has cleared down to 0x4100.
      NotTargeted             Servicing key absent / NotStarted with healthy prerequisites.
      NeedsInvestigation      Anything else.

.PARAMETER Path
    CSV exported from Intune, or a .json/.txt file of one detection output line per row, or a folder of those.

.PARAMETER OutputCsv
    Optional path to write the segmented result.

.EXAMPLE
    .\Split-SecureBootPopulation.ps1 -Path .\DeviceStatus.csv -OutputCsv .\segmented.csv

.OUTPUTS
    PSCustomObject[] when -OutputCsv is omitted.

.NOTES
    PowerShell 5.1. Read-only against the export; touches no devices.
#>
#Requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Path,

    [string]$OutputCsv
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function ConvertTo-NullableBool {
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [bool]) { return $Value }
    $s = "$Value".Trim()
    if ($s -ieq 'true')  { return $true }
    if ($s -ieq 'false') { return $false }
    if ($s -in @('Enabled')) { return $true }
    if ($s -in @('Disabled', 'Unsupported')) { return $false }
    return $null
}

function ConvertFrom-DetectionLine {
    param([string]$Line)
    if ([string]::IsNullOrWhiteSpace($Line)) { return $null }
    $trimmed = $Line.Trim()

    if ($trimmed.StartsWith('{')) {
        try { return ($trimmed | ConvertFrom-Json) } catch { return $null }
    }

    # Support Detect_VMwareSecureBootUEFICA2023.ps1 summary lines: "MARKER: Reason | Key=Val Key2=Val2"
    if ($trimmed -match '^([A-Z0-9-]+):\s*([^|]+)\|\s*(.+)$') {
        $marker = $matches[1].Trim()
        $kvPart = $matches[3].Trim()
        $map = @{}
        foreach ($tok in ($kvPart -split '\s+')) {
            $eq = $tok.IndexOf('=')
            if ($eq -gt 0) {
                $k = $tok.Substring(0, $eq)
                $v = $tok.Substring($eq + 1)
                $map[$k] = $v
            }
        }
        $platVal = if ($map['Plat'] -eq 'VMware') { 'VMware' } elseif ($map['Plat'] -eq 'Unknown' -or [string]::IsNullOrWhiteSpace($map['Plat'])) { 'Unknown' } else { 'Physical' }
        $confVal = if ($null -ne $map['Conf']) { $map['Conf'] -replace '_', ' ' } else { $null }
        $mfrVal  = if ($null -ne $map['Plat']) { $map['Plat'] -replace '_', ' ' } else { $null }
        $errVal  = if ($null -eq $map['Err'] -or $map['Err'] -eq 'NoValue') { $null } elseif ($map['Err'] -eq '0') { 0 } else { $map['Err'] }
        return [PSCustomObject][ordered]@{
            host       = $map['Host']
            marker     = $marker
            sbOn       = if ($map['SB'] -eq 'Enabled' -or $map['SB'] -ieq 'True') { $true } elseif ($map['SB'] -in @('Disabled', 'Unsupported', 'False')) { $false } else { $null }
            st         = $map['Status']
            err        = $errVal
            avl        = if ($map['Avl'] -eq 'None') { $null } else { $map['Avl'] }
            conf       = $confVal
            task       = $map['Task']
            tel        = $null
            oneS       = $null
            fromPrLine = $true
            trust      = if ($map['Trust'] -eq 'N/A') { 'Unknown' } else { $map['Trust'] }
            plat       = $platVal
            pkSt       = $map['PK']
            kek23      = ConvertTo-NullableBool $map['KEK23']
            dbWin23    = ConvertTo-NullableBool $map['DbWin23']
            dbOr23     = ConvertTo-NullableBool $map['DbOptRom23']
            db3p23     = ConvertTo-NullableBool $map['Db3P23']
            dbAll23    = ConvertTo-NullableBool $map['DbAll23']
            vtpm       = ConvertTo-NullableBool $map['vTPM']
            vmt        = $map['VMTools']
            pkCap      = $map['CapPnP']
            lastEvt    = $map['LastEvt']
            skip       = $map['Skip']
            mfr        = $mfrVal
            model      = $null
            fw         = $null
            osb        = $map['OSBuild']
        }
    }

    return $null
}

function Get-DetectionRecords {
    param([string]$InputPath)

    $items = New-Object System.Collections.Generic.List[object]
    if (Test-Path -LiteralPath $InputPath -PathType Container) {
        $files = @(Get-ChildItem -Path $InputPath -File -Include *.csv, *.json, *.txt -Recurse)
    }
    else {
        $files = @(Get-Item -LiteralPath $InputPath)
    }

    foreach ($f in $files) {
        if ($f.Extension -eq '.csv') {
            $rows = @(Import-Csv -LiteralPath $f.FullName)
            foreach ($row in $rows) {
                $raw = $null
                foreach ($col in 'Pre-remediation detection output', 'PreRemediationDetectionScriptOutput', 'DetectionOutput') {
                    if ($row.PSObject.Properties.Name -contains $col -and $row.$col) { $raw = $row.$col; break }
                }
                if (-not $raw) { continue }
                $obj = ConvertFrom-DetectionLine -Line $raw
                if ($null -eq $obj) { continue }
                if ($row.PSObject.Properties.Name -contains 'Device name' -and $row.'Device name') {
                    Add-Member -InputObject $obj -NotePropertyName 'IntuneDeviceName' -NotePropertyValue $row.'Device name' -Force
                }
                [void]$items.Add($obj)
            }
        }
        else {
            foreach ($line in @(Get-Content -LiteralPath $f.FullName)) {
                $obj = ConvertFrom-DetectionLine -Line $line
                if ($null -ne $obj) { [void]$items.Add($obj) }
            }
        }
    }
    return $items.ToArray()
}

function Get-Prop {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    if ($Object.PSObject.Properties.Name -contains $Name) { return $Object.$Name }
    return $null
}

function Test-HasProp {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $false }
    return ($Object.PSObject.Properties.Name -contains $Name)
}

function Get-Segment {
    param($R)

    $marker     = Get-Prop $R 'marker'
    $sbOn       = ConvertTo-NullableBool (Get-Prop $R 'sbOn')
    $st         = Get-Prop $R 'st'
    $err        = Get-Prop $R 'err'
    $task       = Get-Prop $R 'task'
    $tel        = Get-Prop $R 'tel'
    $oneS       = Get-Prop $R 'oneS'
    $trust      = Get-Prop $R 'trust'
    $plat       = Get-Prop $R 'plat'
    $conf       = Get-Prop $R 'conf'
    $avl        = Get-Prop $R 'avl'
    $evt        = Get-Prop $R 'evt'
    $lastEvt    = Get-Prop $R 'lastEvt'
    $pkSt       = Get-Prop $R 'pkSt'
    $kek23      = ConvertTo-NullableBool (Get-Prop $R 'kek23')
    $dbWin23    = ConvertTo-NullableBool (Get-Prop $R 'dbWin23')
    $dbOr23     = ConvertTo-NullableBool (Get-Prop $R 'dbOr23')
    $db3p23     = ConvertTo-NullableBool (Get-Prop $R 'db3p23')
    $db3p11     = ConvertTo-NullableBool (Get-Prop $R 'db3p11')
    $vtpm       = ConvertTo-NullableBool (Get-Prop $R 'vtpm')
    $pkCap      = Get-Prop $R 'pkCap'
    $rebootPend = ConvertTo-NullableBool (Get-Prop $R 'rebootPend')
    $fromPr     = (Get-Prop $R 'fromPrLine') -eq $true

    # Honor explicit Detect_VMwareSecureBootUEFICA2023.ps1 verdict markers first so non-compliant
    # or collection-error verdicts are never upgraded to Updated.
    if (-not [string]::IsNullOrWhiteSpace($marker)) {
        switch ($marker) {
            'PREREQ'                               { return 'CollectionError' }
            'NON-COMPLIANT-COLLECTION-ERROR'       { return 'CollectionError' }
            'NON-COMPLIANT-VMWARE-CAPSULE-REBOOT'  { return 'VMwareCapsuleReboot' }
            'NON-COMPLIANT-VMWARE-CAPSULE-ERROR'   { return 'VMwareCapsuleError' }
            'NON-COMPLIANT-PENDING-REBOOT'         { return 'PendingRestart' }
            'NON-COMPLIANT-OS-SERVICING-NEEDED'    { return 'VMwareOsServicing' }
            'NON-COMPLIANT-PK-READY-NEEDS-KEK'     { return 'VMwarePkReadyNeedsKek' }
            'NON-COMPLIANT-VMWARE-PK-UNRECOGNIZED' { return 'VMwarePkUnrecognized' }
            'NON-COMPLIANT-VMWARE-PK-BLOCKED'      {
                if ($vtpm -eq $true)  { return 'VMwarePkBlockedVtpm' }
                if ($vtpm -eq $false) { return 'VMwarePkBlockedNoVtpm' }
                return 'VMwarePkBlocked'
            }
            'NOT-APPLICABLE' {
                if ($sbOn -eq $false) { return 'SecureBootOff' }
                if ($null -eq $sbOn)  { return 'CollectionError' }
                return 'OutOfScopeNonVMware'
            }
        }
    }

    if ($null -eq $sbOn -or $plat -eq 'Unknown') {
        return 'CollectionError'
    }

    if ($sbOn -eq $false) { return 'SecureBootOff' }

    if ($pkSt -eq 'Unreadable' -or $pkCap -eq 'QueryError' -or $null -eq $kek23 -or $null -eq $dbWin23 -or $trust -notin @('MSOnly', 'MSAnd3P')) {
        return 'CollectionError'
    }

    # If the record includes the explicit compOk property (from Get-SecureBootCertInventory.ps1) and it is $null,
    # the event log query failed -> treat as CollectionError rather than falling back to lastEvt.
    $hasCompOkProp = Test-HasProp $R 'compOk'
    $compPropVal   = if ($hasCompOkProp) { ConvertTo-NullableBool (Get-Prop $R 'compOk') } else { $null }
    if ($hasCompOkProp -and $null -eq $compPropVal) {
        return 'CollectionError'
    }

    $errParsed = 0L
    $errOk     = ($null -eq $err) -or ("$err" -eq '0x0') -or ([int64]::TryParse("$err", [ref]$errParsed) -and $errParsed -eq 0)
    $need3pDb  = if (Test-HasProp $R 'db3p11') { $db3p11 -eq $true } else { $trust -eq 'MSAnd3P' }
    $db3pOk    = if (Test-HasProp $R 'dbAll23') {
        (ConvertTo-NullableBool (Get-Prop $R 'dbAll23')) -eq $true
    } else {
        (-not $need3pDb) -or ($dbOr23 -eq $true -and $db3p23 -eq $true)
    }
    $nvramOk   = ($kek23 -eq $true -and $dbWin23 -eq $true -and $db3pOk)
    $capOk     = ($null -eq $pkCap -or "$pkCap" -in @('', 'Absent', '0', 'N/A'))
    $compEvtOk = if ($hasCompOkProp) {
        $compPropVal -eq $true
    } elseif ($marker -eq 'COMPLIANT') {
        $true
    } else {
        ($null -ne $lastEvt -and "$lastEvt" -in @('1808', '1799')) -or ($null -ne $evt -and $evt.PSObject.Properties.Name -contains '1808' -and "$lastEvt" -notin @('1032', '1033', '1795', '1796', '1797', '1798', '1802', '1803'))
    }

    if (($null -eq $marker -or $marker -eq 'COMPLIANT') -and $sbOn -eq $true -and $st -eq 'Updated' -and $errOk -and $avl -eq '0x4000' -and $nvramOk -and $capOk -and $compEvtOk -and $rebootPend -ne $true) {
        return 'Updated'
    }

    # Microsoft-only trust: the Option ROM and Microsoft UEFI CA 2023 entries are not applicable.
    if ($trust -eq 'MSOnly' -and $dbWin23 -eq $true -and $kek23 -eq $true -and $plat -ne 'VMware' -and $st -eq 'Updated' -and $errOk -and $rebootPend -ne $true) {
        return 'OptionRomNotApplicable'
    }

    # VMware guest segmentation (evaluates KEK 2023 presence before PK state per MS AMA + Broadcom KB 423893)
    if ($plat -eq 'VMware') {
        if ("$pkCap" -eq '14') { return 'VMwareCapsuleReboot' }
        if ($null -ne $pkCap -and "$pkCap" -notin @('', 'Absent', '0', '14', 'N/A')) {
            return 'VMwareCapsuleError'
        }

        if ($kek23 -eq $true) {
            if ($avl -eq '0x4100' -or $rebootPend -eq $true) { return 'PendingRestart' }
            return 'VMwareOsServicing'
        }

        if ($kek23 -eq $false) {
            if ($pkSt -eq 'WinOemPK') { return 'VMwarePkReadyNeedsKek' }
            if ($pkSt -in @('NullPK', 'None')) {
                if ($vtpm -eq $true)  { return 'VMwarePkBlockedVtpm' }
                if ($vtpm -eq $false) { return 'VMwarePkBlockedNoVtpm' }
                return 'VMwarePkBlocked'
            }
            if ($pkSt -in @('OtherCertPK', 'UnrecognizedPK')) {
                return 'VMwarePkUnrecognized'
            }
        }

        return 'VirtualVMware'
    }

    $telBad    = (-not $fromPr -and $null -eq $tel) -or ($null -ne $tel -and [int]$tel -lt 1)
    $prereqBad = ($task -notin @('Ready', 'Running')) -or $telBad -or ($oneS -eq 1)
    if ($prereqBad) { return 'ReportingBlocked' }

    $hasFwEvent = $false
    if ($evt) {
        foreach ($id in '1795', '1802', '1803', '1032') {
            if ($evt.PSObject.Properties.Name -contains $id) { $hasFwEvent = $true }
        }
    }
    $confBlocked = ($conf -match 'Temporarily Paused|Not Supported')

    if ($plat -eq 'HyperV' -or $plat -eq 'Azure' -or $plat -eq 'OtherVM') { return 'VirtualOtherHV' }

    if ($hasFwEvent -or $confBlocked) { return 'FirmwareBlocked' }

    if (($st -eq 'InProgress' -and $avl -eq '0x4100') -or $rebootPend -eq $true) { return 'PendingRestart' }
    if ($st -eq 'NoValue' -or $st -eq 'NotStarted') { return 'NotTargeted' }

    return 'NeedsInvestigation'
}

$records = @(Get-DetectionRecords -InputPath $Path)
if ($records.Count -eq 0) {
    throw "No parseable detection records found in '$Path'."
}

$result = @(foreach ($r in $records) {
    [PSCustomObject][ordered]@{
        Device     = $(if (Get-Prop $r 'IntuneDeviceName') { Get-Prop $r 'IntuneDeviceName' } else { Get-Prop $r 'host' })
        Segment    = Get-Segment $r
        Status     = Get-Prop $r 'st'
        SecureBoot = Get-Prop $r 'sbOn'
        Trust      = Get-Prop $r 'trust'
        Platform   = Get-Prop $r 'plat'
        PkState    = Get-Prop $r 'pkSt'
        Kek2023    = Get-Prop $r 'kek23'
        DbWin2023  = Get-Prop $r 'dbWin23'
        DbOptRom23 = Get-Prop $r 'dbOr23'
        Db3p2023   = Get-Prop $r 'db3p23'
        Vtpm       = Get-Prop $r 'vtpm'
        VMTools    = Get-Prop $r 'vmt'
        PkCapsule  = Get-Prop $r 'pkCap'
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
})

$result | Group-Object Segment | Sort-Object Count -Descending |
    Select-Object @{ n = 'Segment'; e = { $_.Name } }, Count | Format-Table -AutoSize | Out-Host

if ($OutputCsv) {
    $result | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
    Write-Host "Wrote $($result.Count) rows to $OutputCsv"
}
else {
    $result
}
