#requires -Version 5.1
<#
.SYNOPSIS
    Reports the effective Windows Location policy state and identifies every
    author that can force / lock the Settings > Privacy > Location toggle.

.DESCRIPTION
    Read-only. Inspects all known vectors that control the Location master
    switch and per-app access, then prints a summary table and a plain-English
    verdict. Designed for triaging "location won't turn on / is greyed out"
    on Intune-managed (Entra-only and Hybrid) clients.

    Vectors checked:
      1.  MDM AllowLocation (effective)  HKLM\...\PolicyManager\current\device\System\AllowLocation
      2.  MDM AllowLocation (providers)  HKLM\...\PolicyManager\providers\<GUID>\...\System\AllowLocation
      2b. MDM LetAppsAccessLocation      HKLM\...\PolicyManager\current\device\Privacy\LetAppsAccessLocation
      3.  Legacy Location GP             HKLM\SOFTWARE\Policies\Microsoft\Windows\LocationAndSensors\DisableLocation
                                         (+ companions: DisableLocationScripting / DisableWindowsLocationProvider / DisableSensors)
      4.  App-privacy GP                 HKLM\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy\LetAppsAccessLocation
                                         (+ per-app lists: _ForceDenyTheseApps / _ForceAllowTheseApps / _UserInControlOfTheseApps)
      5.  CAM consent (HKLM/HKCU)        ...\CapabilityAccessManager\ConsentStore\location\Value (+ key owner/ACL)
      5b. DeviceAccess capability gate   ...\DeviceAccess\Global\{BFA794E4-F964-4FDB-90F6-51056BFE4B44}\Value (HKLM/HKCU)
      6.  Settings page hide             ...\Policies\Explorer\SettingsPageVisibility
      7.  Geolocation service            lfsvc (Start value + running state)

.NOTES
    Author: Anton Romanyuk
    Read-only. No changes are made to the system.

    Sources / confidence:
      - AllowLocation & LetAppsAccessLocation CSP value meanings are grounded in
        Microsoft Learn:
        https://learn.microsoft.com/en-us/windows/client-management/mdm/policy-csp-system#allowlocation
      - The DeviceAccess\Global capability-broker gate (5b) and the ConsentStore
        ACL/owner check (5) come from internal/community research, NOT a formal
        Learn CSP doc. They are cheap read-only checks and the leading suspects
        when ConsentStore edits have no effect - but treat their exact behaviour
        as "to be confirmed on the affected client", not documented fact.

    PowerShell 5.1 compatible.
#>

[CmdletBinding()]
param(
    [switch]$AsObject,        # emit the result object instead of the formatted report
    [switch]$Collect,         # gather a deep artifact bundle (reg exports + ACLs + report) and zip it
    [string]$CollectPath      # optional folder for the bundle; default: %TEMP%\LocationDiag_<host>_<stamp>
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Get-RegValue {
    param([string]$Path, [string]$Name)
    try {
        $item = Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop
        return [pscustomobject]@{ Found = $true; Value = $item.$Name }
    } catch {
        return [pscustomobject]@{ Found = $false; Value = $null }
    }
}

# Returns a summary of Deny access-control entries on a registry key (vector #7:
# a user write-deny lets the toggle appear but silently snaps the value back).
function Get-KeyAclSummary {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return $null }
    try {
        $acl  = Get-Acl -Path $Path -ErrorAction Stop
        $deny = @($acl.Access | Where-Object { $_.AccessControlType -eq 'Deny' })
        [pscustomobject]@{
            Owner     = $acl.Owner
            DenyCount = $deny.Count
            DenyAces  = (($deny | ForEach-Object { "$($_.IdentityReference)=$($_.RegistryRights)" }) -join '; ')
        }
    } catch {
        [pscustomobject]@{ Owner = '(ACL read failed)'; DenyCount = 0; DenyAces = '' }
    }
}

function Convert-AllowLocation {
    param($Value)
    switch ([int]$Value) {
        0 { 'Force OFF (0) - toggles greyed off, no app access' }
        1 { 'User Control (1) - default; does NOT force on' }
        2 { 'Force ON (2) - toggles greyed on' }
        default { "Unknown ($Value)" }
    }
}

function Convert-LetApps {
    param($Value)
    switch ([int]$Value) {
        0 { 'User in control (0) - default' }
        1 { 'Force Allow (1) - locked on' }
        2 { 'Force Deny (2) - locked off' }
        default { "Unknown ($Value)" }
    }
}

$results = New-Object System.Collections.Generic.List[object]
function Add-Result {
    param([string]$Area, [string]$Path, [string]$Name, [bool]$Present, $Value, [string]$Meaning)
    $results.Add([pscustomobject]@{
        Area    = $Area
        Present = $Present
        Value   = $Value
        Meaning = $Meaning
        Path    = ($Path + '\' + $Name)
    })
}

# ── 1. MDM winner (merged/effective CSP value) ─────────────────────────────
$mdmCurPath = 'HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\System'
$mdmCur = Get-RegValue -Path $mdmCurPath -Name 'AllowLocation'
Add-Result -Area 'MDM AllowLocation (effective)' -Path $mdmCurPath -Name 'AllowLocation' `
    -Present $mdmCur.Found -Value $mdmCur.Value `
    -Meaning $(if ($mdmCur.Found) { Convert-AllowLocation $mdmCur.Value } else { 'NOT APPLIED - no MDM AllowLocation policy on this device' })

# ── 2. MDM providers (which enrollment authored it) ────────────────────────
$provRoot = 'HKLM:\SOFTWARE\Microsoft\PolicyManager\providers'
$provHits = @()
if (Test-Path $provRoot) {
    foreach ($prov in (Get-ChildItem $provRoot -ErrorAction SilentlyContinue)) {
        $pPath = Join-Path $prov.PSPath 'default\Device\System'
        $pVal = Get-RegValue -Path $pPath -Name 'AllowLocation'
        if ($pVal.Found) {
            $provHits += [pscustomobject]@{ Provider = $prov.PSChildName; Value = $pVal.Value }
        }
    }
}
if ($provHits.Count -gt 0) {
    foreach ($h in $provHits) {
        Add-Result -Area 'MDM AllowLocation (provider)' -Path "$provRoot\$($h.Provider)\default\Device\System" -Name 'AllowLocation' `
            -Present $true -Value $h.Value -Meaning "Authored by provider $($h.Provider): $(Convert-AllowLocation $h.Value)"
    }
} else {
    Add-Result -Area 'MDM AllowLocation (provider)' -Path $provRoot -Name '(any)' `
        -Present $false -Value $null -Meaning 'No enrollment provider is authoring AllowLocation'
}

# ── 2b. MDM effective LetAppsAccessLocation (Privacy) ──────────────────────
$mdmPrivCur = Get-RegValue -Path 'HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Privacy' -Name 'LetAppsAccessLocation'
Add-Result -Area 'MDM LetAppsAccessLocation (effective)' -Path 'HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Privacy' -Name 'LetAppsAccessLocation' `
    -Present $mdmPrivCur.Found -Value $mdmPrivCur.Value `
    -Meaning $(if ($mdmPrivCur.Found) { Convert-LetApps $mdmPrivCur.Value } else { 'Not applied via MDM' })

# ── 3. Legacy Location GP (+ companions) ───────────────────────────────────
$locGpPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\LocationAndSensors'
$locGp = Get-RegValue -Path $locGpPath -Name 'DisableLocation'
Add-Result -Area 'GP DisableLocation' -Path $locGpPath -Name 'DisableLocation' `
    -Present $locGp.Found -Value $locGp.Value `
    -Meaning $(if ($locGp.Found) { if ([int]$locGp.Value -eq 1) { '1 = location turned OFF and locked (produces "managed by your organization")' } else { "$($locGp.Value) = not forcing off" } } else { 'Not configured' })
foreach ($comp in 'DisableLocationScripting','DisableWindowsLocationProvider','DisableSensors') {
    $c = Get-RegValue -Path $locGpPath -Name $comp
    if ($c.Found) {
        Add-Result -Area "GP $comp" -Path $locGpPath -Name $comp -Present $true -Value $c.Value `
            -Meaning $(if ([int]$c.Value -eq 1) { 'Enabled - locks a related location sub-feature (does not gate the master toggle)' } else { "$($c.Value)" })
    }
}

# ── 4. App-privacy GP (+ per-app force lists) ──────────────────────────────
$appPrivPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy'
$appPriv = Get-RegValue -Path $appPrivPath -Name 'LetAppsAccessLocation'
Add-Result -Area 'GP LetAppsAccessLocation' -Path $appPrivPath -Name 'LetAppsAccessLocation' `
    -Present $appPriv.Found -Value $appPriv.Value `
    -Meaning $(if ($appPriv.Found) { Convert-LetApps $appPriv.Value } else { 'Not configured (user in control)' })
foreach ($comp in 'LetAppsAccessLocation_ForceDenyTheseApps','LetAppsAccessLocation_ForceAllowTheseApps','LetAppsAccessLocation_UserInControlOfTheseApps') {
    $c = Get-RegValue -Path $appPrivPath -Name $comp
    if ($c.Found) {
        $list = if ($c.Value -is [array]) { ($c.Value -join '; ') } else { "$($c.Value)" }
        Add-Result -Area "GP $($comp -replace 'LetAppsAccessLocation_','')" -Path $appPrivPath -Name $comp -Present $true -Value $list `
            -Meaning 'Per-app exception list (package family names) - locks those app rows'
    }
}

# ── 5. Capability Access Manager consent store (HKLM + HKCU) + writeability ─
$aclDenyHits = @()
foreach ($hive in @('HKLM','HKCU')) {
    $camPath = "${hive}:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location"
    $cam = Get-RegValue -Path $camPath -Name 'Value'
    $writeNote = ''
    $acl = Get-KeyAclSummary -Path $camPath
    if ($acl) {
        $writeNote = " [owner: $($acl.Owner)" + $(if ($acl.DenyCount -gt 0) { "; DENY ACEs: $($acl.DenyAces)" } else { '' }) + ']'
        if ($acl.DenyCount -gt 0) { $aclDenyHits += "CAM consent ($hive): $($acl.DenyAces)" }
    }
    Add-Result -Area "CAM consent ($hive)" -Path $camPath -Name 'Value' `
        -Present $cam.Found -Value $cam.Value `
        -Meaning $(if ($cam.Found) { "$($cam.Value)$(if ($cam.Value -eq 'Deny') { ' - blocks location' } elseif ($cam.Value -eq 'Allow') { ' - allows location' } else { '' })$writeNote" } else { "Not set at this scope$writeNote" })
}

# ── 5b. DeviceAccess\Global location capability GUID (HKLM + HKCU) ──────────
# The Settings toggle also honours this capability-broker key; a 'Deny' here
# gates location and survives ConsentStore edits.
$locCapGuid = '{BFA794E4-F964-4FDB-90F6-51056BFE4B44}'
foreach ($hive in @('HKLM','HKCU')) {
    $daPath = "${hive}:\SOFTWARE\Microsoft\Windows\CurrentVersion\DeviceAccess\Global\$locCapGuid"
    $da = Get-RegValue -Path $daPath -Name 'Value'
    $daAcl = Get-KeyAclSummary -Path $daPath
    $daNote = if ($daAcl -and $daAcl.DenyCount -gt 0) { " [DENY ACEs: $($daAcl.DenyAces)]" } else { '' }
    if ($daAcl -and $daAcl.DenyCount -gt 0) { $aclDenyHits += "DeviceAccess Global ($hive): $($daAcl.DenyAces)" }
    Add-Result -Area "DeviceAccess Global ($hive)" -Path $daPath -Name 'Value' `
        -Present $da.Found -Value $da.Value `
        -Meaning $(if ($da.Found) { "$($da.Value)$(if ($da.Value -eq 'Deny') { ' - capability broker BLOCKS location (survives ConsentStore edits)' } elseif ($da.Value -eq 'Allow') { ' - allows' } else { '' })$daNote" } else { "Not set at this scope$daNote" })
}

# ── 6. Settings page visibility (page hidden entirely?) ────────────────────
$spvPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer'
$spv = Get-RegValue -Path $spvPath -Name 'SettingsPageVisibility'
$spvHidesLoc = $false
if ($spv.Found -and $spv.Value -match 'privacy-location') { $spvHidesLoc = $true }
Add-Result -Area 'SettingsPageVisibility' -Path $spvPath -Name 'SettingsPageVisibility' `
    -Present $spv.Found -Value $spv.Value `
    -Meaning $(if ($spvHidesLoc) { 'Location page is HIDDEN from Settings' } elseif ($spv.Found) { 'Set, but does not hide the location page' } else { 'Not configured' })

# ── 7. Geolocation service (lfsvc) ─────────────────────────────────────────
$svcStartPath = 'HKLM:\SYSTEM\CurrentControlSet\Services\lfsvc'
$svcStart = Get-RegValue -Path $svcStartPath -Name 'Start'
$svcState = 'unknown'
try { $svcState = (Get-Service -Name 'lfsvc' -ErrorAction Stop).Status.ToString() } catch { $svcState = 'not found' }
$startMeaning = 'unknown'
if ($svcStart.Found) {
    switch ([int]$svcStart.Value) {
        2 { $startMeaning = 'Automatic' }
        3 { $startMeaning = 'Manual (default)' }
        4 { $startMeaning = 'DISABLED - location cannot start' }
        default { $startMeaning = "Start=$($svcStart.Value)" }
    }
}
Add-Result -Area 'Geolocation service (lfsvc)' -Path $svcStartPath -Name 'Start' `
    -Present $svcStart.Found -Value $svcStart.Value `
    -Meaning "$startMeaning; current state: $svcState"

# ── Verdict ────────────────────────────────────────────────────────────────
$verdict = New-Object System.Collections.Generic.List[string]

$forcedOff = ($locGp.Found -and [int]$locGp.Value -eq 1) -or ($mdmCur.Found -and [int]$mdmCur.Value -eq 0)
$forcedOn  =  $mdmCur.Found -and [int]$mdmCur.Value -eq 2
$svcDisabled = $svcStart.Found -and [int]$svcStart.Value -eq 4
$daBlock = $results | Where-Object { $_.Area -like 'DeviceAccess Global*' -and $_.Value -eq 'Deny' }

if (-not $mdmCur.Found -and $provHits.Count -eq 0) {
    $verdict.Add('AllowLocation MDM policy is NOT present on this device - the Intune profile did not apply here (assignment/targeting or sync issue), it is not a conflict. Collect an MDM diagnostic report and confirm the profile is assigned to this device''s group.')
}
if ($forcedOn)  { $verdict.Add('AllowLocation = 2 (Force ON) is applied - location should be greyed ON. If it is not, check lfsvc.') }
if ($forcedOff) { $verdict.Add('A policy is FORCING location OFF (GP DisableLocation=1 or MDM AllowLocation=0). This wins over "user control" and produces the "managed by your organization" banner. Find and re-scope/remove it (check SCCM/GPO as well as Intune).') }
if ($appPriv.Found -and [int]$appPriv.Value -eq 2) { $verdict.Add('LetAppsAccessLocation = 2 (Force Deny) blocks all app access even if the master switch is on.') }
if ($daBlock) { $verdict.Add('DeviceAccess\Global location capability = Deny - the capability broker is blocking location. This gate survives ConsentStore\location\Value edits, which explains why changing that value had no effect.') }
if ($aclDenyHits.Count -gt 0) { $verdict.Add('A registry ACL DENY was found on a location key (' + ($aclDenyHits -join ' | ') + '). A write-deny lets the toggle appear but silently reverts the value - a non-GP/non-MDM hardening lock. Inspect and remove the Deny ACE.') }
if ($spvHidesLoc) { $verdict.Add('SettingsPageVisibility hides the Location page entirely.') }
if ($svcDisabled) { $verdict.Add('lfsvc is DISABLED (Start=4) - the master switch cannot turn on until the service is enabled.') }
if ($verdict.Count -eq 0) { $verdict.Add('No forcing/locking policy detected in the queried vectors. If the toggle is still greyed, re-run with -Collect to gather a full artifact bundle (recursive reg exports + ACLs of ConsentStore\location and the entire DeviceAccess\Global tree, both hives) for offline analysis, and consider arming registry auditing to catch the writer (see the bundle README).') }

# ── Deep collection bundle (read-only; writes only to the output folder) ───
function Invoke-DeepCollect {
    param($Results, $Verdict, [string]$OutDir)

    New-Item -ItemType Directory -Path $OutDir -Force | Out-Null

    # 1. The console report + verdict as text
    $report = New-Object System.Collections.Generic.List[string]
    $report.Add("Windows Location Policy State - deep collection")
    $report.Add("Computer: $env:COMPUTERNAME   User: $env:USERNAME   $(Get-Date -Format o)")
    $report.Add('')
    $report.Add(($Results | Select-Object Area,
        @{n='Present';e={ if ($_.Present) { 'Yes' } else { 'no' } }},
        @{n='Value';e={ if ($null -ne $_.Value) { $_.Value } else { '-' } }},
        Meaning | Format-Table -AutoSize -Wrap | Out-String))
    $report.Add('Verdict:')
    $Verdict | ForEach-Object { $report.Add("  - $_") }
    Set-Content -Path (Join-Path $OutDir 'LocationState.txt') -Value ($report -join [Environment]::NewLine) -Encoding UTF8

    # 2. Recursive registry exports of every relevant key (reg.exe, read-only)
    $exports = @(
        @{ n='HKLM_ConsentStore_location';  k='HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location' },
        @{ n='HKCU_ConsentStore_location';  k='HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location' },
        @{ n='HKLM_DeviceAccess_Global';    k='HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\DeviceAccess\Global' },
        @{ n='HKCU_DeviceAccess_Global';    k='HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\DeviceAccess\Global' },
        @{ n='HKLM_LocationAndSensors_pol'; k='HKLM\SOFTWARE\Policies\Microsoft\Windows\LocationAndSensors' },
        @{ n='HKLM_AppPrivacy_pol';         k='HKLM\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy' },
        @{ n='HKLM_PolicyManager_System';   k='HKLM\SOFTWARE\Microsoft\PolicyManager\current\device\System' },
        @{ n='HKLM_PolicyManager_Privacy';  k='HKLM\SOFTWARE\Microsoft\PolicyManager\current\device\Privacy' },
        @{ n='HKLM_lfsvc';                  k='HKLM\SYSTEM\CurrentControlSet\Services\lfsvc' }
    )
    foreach ($e in $exports) {
        $file = Join-Path $OutDir ("$($e.n).reg")
        # reg.exe writes to stderr for a missing key, which becomes a terminating
        # error under $ErrorActionPreference='Stop'. Skip absent keys and isolate
        # the native call so a missing key never aborts the collection.
        $psKey = ($e.k -replace '^HKLM\\', 'HKLM:\') -replace '^HKCU\\', 'HKCU:\'
        if (-not (Test-Path -LiteralPath $psKey)) { continue }
        try {
            cmd.exe /c "reg.exe export `"$($e.k)`" `"$file`" /y" 2>$null 1>$null
        } catch { }
    }

    # 3. ACL dumps for the consent + capability keys (finds a user write-deny)
    $aclText = New-Object System.Collections.Generic.List[string]
    $aclKeys = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location',
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\DeviceAccess\Global\$locCapGuid",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\DeviceAccess\Global\$locCapGuid"
    )
    foreach ($k in $aclKeys) {
        $aclText.Add("===== $k =====")
        if (Test-Path $k) {
            try {
                $a = Get-Acl -Path $k
                $aclText.Add("Owner: $($a.Owner)")
                foreach ($ace in $a.Access) {
                    $aclText.Add(("  {0,-6} {1,-30} {2}" -f $ace.AccessControlType, $ace.IdentityReference, $ace.RegistryRights))
                }
            } catch { $aclText.Add("  (ACL read failed: $($_.Exception.Message))") }
        } else { $aclText.Add('  (key not present)') }
        $aclText.Add('')
    }
    Set-Content -Path (Join-Path $OutDir 'ACLs.txt') -Value ($aclText -join [Environment]::NewLine) -Encoding UTF8

    # 4. Full DeviceAccess\Global capability map (Value per GUID, both hives)
    $daText = New-Object System.Collections.Generic.List[string]
    foreach ($hive in @('HKLM','HKCU')) {
        $root = "${hive}:\SOFTWARE\Microsoft\Windows\CurrentVersion\DeviceAccess\Global"
        $daText.Add("===== $root =====")
        if (Test-Path $root) {
            Get-ChildItem $root -ErrorAction SilentlyContinue | ForEach-Object {
                $v = (Get-ItemProperty $_.PSPath -Name 'Value' -ErrorAction SilentlyContinue).Value
                $daText.Add(("  {0}  Value={1}" -f $_.PSChildName, $(if ($null -ne $v) { $v } else { '(none)' })))
            }
        } else { $daText.Add('  (not present)') }
        $daText.Add('')
    }
    Set-Content -Path (Join-Path $OutDir 'DeviceAccess_Global_map.txt') -Value ($daText -join [Environment]::NewLine) -Encoding UTF8

    # 5. MDM diagnostic report (what Intune actually delivered) - best effort.
    #    Needs elevation; skips cleanly if unavailable so it never breaks the bundle.
    $mdmTool = Join-Path $env:windir 'System32\MdmDiagnosticsTool.exe'
    if (Test-Path $mdmTool) {
        $mdmZip = Join-Path $OutDir 'MDMDiagReport.zip'
        try {
            $p = Start-Process -FilePath $mdmTool `
                -ArgumentList @('-area', 'DeviceProvisioning', '-zip', "`"$mdmZip`"") `
                -Wait -PassThru -WindowStyle Hidden
            if ($p.ExitCode -ne 0 -or -not (Test-Path $mdmZip)) {
                Set-Content -Path (Join-Path $OutDir 'MDMDiagReport_SKIPPED.txt') `
                    -Value "MdmDiagnosticsTool exit code $($p.ExitCode). Re-run this script elevated (as admin) to capture the MDM report." -Encoding UTF8
            }
        } catch {
            Set-Content -Path (Join-Path $OutDir 'MDMDiagReport_SKIPPED.txt') `
                -Value "MdmDiagnosticsTool failed: $($_.Exception.Message). Re-run elevated (as admin)." -Encoding UTF8
        }
    }

    # 6. Zip the bundle next to the folder
    $zip = "$OutDir.zip"
    if (Test-Path $zip) { Remove-Item $zip -Force }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::CreateFromDirectory($OutDir, $zip)
    return $zip
}

$zipPath = $null
if ($Collect) {
    if (-not $CollectPath) {
        $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        $CollectPath = Join-Path $env:TEMP "LocationDiag_${env:COMPUTERNAME}_$stamp"
    }
    $zipPath = Invoke-DeepCollect -Results $results -Verdict $verdict -OutDir $CollectPath
}

if ($AsObject) {
    [pscustomobject]@{
        Checks    = $results
        Providers = $provHits
        Verdict   = $verdict
        ZipPath   = $zipPath
    }
    return
}

# ── Formatted output ───────────────────────────────────────────────────────
Write-Host ''
Write-Host '=== Windows Location Policy State ===' -ForegroundColor Cyan
Write-Host ("Computer: {0}   User: {1}   {2}" -f $env:COMPUTERNAME, $env:USERNAME, (Get-Date)) -ForegroundColor DarkGray
Write-Host ''

$results |
    Select-Object Area,
        @{n='Present';e={ if ($_.Present) { 'Yes' } else { 'no' } }},
        @{n='Value';e={ if ($null -ne $_.Value) { $_.Value } else { '-' } }},
        Meaning |
    Format-Table -AutoSize -Wrap

Write-Host 'Effective source paths:' -ForegroundColor DarkGray
$results | ForEach-Object { Write-Host ('  ' + $_.Path) -ForegroundColor DarkGray }

Write-Host ''
Write-Host '=== Verdict ===' -ForegroundColor Cyan
foreach ($v in $verdict) {
    $color = 'Yellow'
    if ($v -like 'No forcing*') { $color = 'Green' }
    if ($v -like '*NOT present*' -or $v -like '*FORCING location OFF*' -or $v -like '*DISABLED*') { $color = 'Red' }
    Write-Host ('  - ' + $v) -ForegroundColor $color
}

Write-Host ''
Write-Host 'Next step if AllowLocation is missing: capture an MDM report and search it for AllowLocation:' -ForegroundColor DarkGray
Write-Host '  MdmDiagnosticsTool.exe -area DeviceProvisioning -zip C:\Temp\MDMDiag.zip' -ForegroundColor DarkGray

if ($Collect) {
    Write-Host ''
    Write-Host '=== Deep collection ===' -ForegroundColor Cyan
    Write-Host "  Folder: $CollectPath" -ForegroundColor Green
    Write-Host "  Zip:    $zipPath" -ForegroundColor Green
    Write-Host '  Contents: LocationState.txt, ACLs.txt, DeviceAccess_Global_map.txt, *.reg exports, MDMDiagReport.zip' -ForegroundColor DarkGray
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) { Write-Host '  NOTE: not elevated - the MDM report (and some ACLs) may be skipped. Re-run as admin for a complete bundle.' -ForegroundColor Yellow }
    Write-Host '  Send the .zip back for offline analysis.' -ForegroundColor DarkGray
}
else {
    Write-Host ''
    Write-Host 'If nothing above explains a greyed-out toggle, re-run with -Collect to gather a full artifact bundle (reg exports + ACLs + capability map) as a zip.' -ForegroundColor DarkGray
}
