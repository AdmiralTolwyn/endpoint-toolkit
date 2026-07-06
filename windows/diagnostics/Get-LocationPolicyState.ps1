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
    [switch]$AsObject   # emit the result object instead of the formatted report
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
foreach ($hive in @('HKLM','HKCU')) {
    $camPath = "${hive}:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location"
    $cam = Get-RegValue -Path $camPath -Name 'Value'
    $writeNote = ''
    if (Test-Path $camPath) {
        try {
            $acl = Get-Acl -Path $camPath
            $owner = $acl.Owner
            $writeNote = " [owner: $owner]"
        } catch { $writeNote = ' [ACL read failed]' }
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
    Add-Result -Area "DeviceAccess Global ($hive)" -Path $daPath -Name 'Value' `
        -Present $da.Found -Value $da.Value `
        -Meaning $(if ($da.Found) { "$($da.Value)$(if ($da.Value -eq 'Deny') { ' - capability broker BLOCKS location (survives ConsentStore edits)' } elseif ($da.Value -eq 'Allow') { ' - allows' } else { '' })" } else { 'Not set at this scope' })
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
if ($spvHidesLoc) { $verdict.Add('SettingsPageVisibility hides the Location page entirely.') }
if ($svcDisabled) { $verdict.Add('lfsvc is DISABLED (Start=4) - the master switch cannot turn on until the service is enabled.') }
if ($verdict.Count -eq 0) { $verdict.Add('No forcing/locking policy detected in the queried vectors. If the toggle is still greyed, capture a fresh reg export of both ConsentStore\location and DeviceAccess\Global under HKLM and HKCU, and check the key ACLs for a user write-deny.') }

if ($AsObject) {
    [pscustomobject]@{
        Checks   = $results
        Providers= $provHits
        Verdict  = $verdict
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
