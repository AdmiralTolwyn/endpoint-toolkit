<#
.SYNOPSIS
    Scrapes Microsoft CSP documentation to build a comprehensive metadata database for PolicyPilot.

.DESCRIPTION
    Fetches each Policy CSP area page from learn.microsoft.com/en-us/windows/client-management/mdm/
    and parses: setting name, description, default value, allowed values, scope (Device/User),
    supported editions, minimum Windows version, and Group Policy registry mapping.

    Uses reliable HTML comment markers (<!-- SettingName-Description-Begin --> etc.) for extraction.
    Output: csp_metadata.json — loaded by PolicyPilot at startup.

.NOTES
    Run once to generate. Re-run to refresh when MS updates docs.
    Requires internet access. Takes ~5-10 minutes for all areas.

    UPDATE FREQUENCY: MS CSP documentation updates with major Windows releases
    (roughly quarterly). PolicyPilot warns in the log if the JSON is >90 days old.
    Recommended: re-run every 3-6 months or after a major Windows feature update.

.EXAMPLE
    .\Build-CspDatabase.ps1
    .\Build-CspDatabase.ps1 -Areas 'Defender','Update','DeviceLock'
#>
[CmdletBinding()]
param(
    [string[]]$Areas,
    [string]$OutputPath = (Join-Path $PSScriptRoot 'csp_metadata.json')
)

$ErrorActionPreference = 'Stop'
$baseUrl = 'https://learn.microsoft.com/en-us/windows/client-management/mdm/policy-csp-'

# Non-ADMX areas (most relevant for Intune/MDM policies in PolicyManager registry)
$allAreas = @(
    'AboveLock','Accounts','ActiveXControls','ApplicationManagement','AppRuntime',
    'AttachmentManager','Audit','Authentication','Autoplay','Bitlocker','BITS',
    'Bluetooth','Browser','Camera','Cellular','CloudDesktop','Connectivity',
    'ControlPolicyConflict','CredentialProviders','CredentialsDelegation','CredentialsUI',
    'Cryptography','DataProtection','DataUsage','Defender','DeliveryOptimization',
    'Desktop','DesktopAppInstaller','DeviceGuard','DeviceHealthMonitoring',
    'DeviceInstallation','DeviceLock','Display','DmaGuard','Eap','Education',
    'EnterpriseCloudPrint','ErrorReporting','EventLogService','Experience',
    'ExploitGuard','FederatedAuthentication','FileExplorer','FileSystem','Games',
    'Handwriting','HumanPresence','InternetExplorer','Kerberos','KioskBrowser',
    'LanmanServer','LanmanWorkstation','Licensing','LocalPoliciesSecurityOptions',
    'LocalSecurityAuthority','LocalUsersAndGroups','LockDown','Maps','MemoryDump',
    'Messaging','MixedReality','MSSecurityGuide','MSSLegacy','Multitasking',
    'NetworkIsolation','NetworkListManager','NewsAndInterests','Notifications',
    'Power','Printers','Privacy','RemoteAssistance','RemoteDesktop',
    'RemoteDesktopServices','RemoteManagement','RemoteProcedureCall','RemoteShell',
    'RestrictedGroups','Search','SecureBoot','Security','ServiceControlManager',
    'Settings','SettingsSync','SmartScreen','Speech','Start','Stickers','Storage',
    'Sudo','System','SystemServices','TaskManager','TaskScheduler',
    'TenantDefinedTelemetry','TenantRestrictions','TextInput','TimeLanguageSettings',
    'Troubleshooting','Update','UserRights','VirtualizationBasedTechnology',
    'WebThreatDefense','Wifi','WindowsAI','WindowsAutopilot',
    'WindowsConnectionManager','WindowsDefenderSecurityCenter','WindowsInkWorkspace',
    'WindowsLogon','WindowsPowerShell','WindowsSandbox','WirelessDisplay',
    'AppDeviceInventory','ApplicationDefaults','AppVirtualization'
)

if ($Areas) { $targetAreas = $Areas } else { $targetAreas = $allAreas }

$database = [ordered]@{}
$totalSettings = 0
$failedAreas = @()

Write-Host "CSP Metadata Builder - scraping $($targetAreas.Count) area(s) from MS Learn..." -ForegroundColor Cyan

foreach ($area in $targetAreas) {
    $url = "$baseUrl$($area.ToLower())"
    Write-Host "  [$($targetAreas.IndexOf($area)+1)/$($targetAreas.Count)] $area ... " -NoNewline

    try {
        $response = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 30
        $html = $response.Content
    }
    catch {
        Write-Host "FAILED ($($_.Exception.Message))" -ForegroundColor Red
        $failedAreas += $area
        continue
    }

    # Parse settings using HTML comment markers: <!-- SettingName-Begin --> ... <!-- SettingName-End -->
    $singleLine = [System.Text.RegularExpressions.RegexOptions]::Singleline
    $settingBlocks = [regex]::Matches($html, '<!-- (\w+)-Begin -->\s*<h[23][^>]*>(.+?)</h[23]>(.+?)<!-- \1-End -->', $singleLine)

    $areaSettings = 0
    foreach ($block in $settingBlocks) {
        $settingName = $block.Groups[1].Value
        $content     = $block.Groups[3].Value
        $key         = "$area/$settingName"

        # --- Scope, editions, min version from Applicability sub-block ---
        $scope = 'Device'; $editions = @(); $minVersion = ''
        $appM = [regex]::Match($content, 'Applicability-Begin\s*-->(.+?)<!--\s*\w+-Applicability-End', $singleLine)
        if ($appM.Success) {
            $a = $appM.Groups[1].Value
            if ($a -match '✅ Device[^✅❌]*✅ User')   { $scope = 'Both' }
            elseif ($a -match '❌ Device[^✅❌]*✅ User') { $scope = 'User' }
            if ($a -match '✅ Pro')        { $editions += 'Pro' }
            if ($a -match '✅ Enterprise')  { $editions += 'Enterprise' }
            if ($a -match '✅ Education')   { $editions += 'Education' }
            if ($a -match '✅ IoT')         { $editions += 'IoT Enterprise' }
            if ($a -match '✅ Windows\s+(\d+),?\s*version\s+(\d+[A-Za-z]*)\s*\[([^\]]+)\]') {
                $minVersion = "Windows $($Matches[1]) $($Matches[2]) [$($Matches[3])]"
            } elseif ($a -match '✅ Windows\s+(\d+),?\s*version\s+(\d+)') {
                $minVersion = "Windows $($Matches[1]) $($Matches[2])"
            }
        }

        # --- Description from Description sub-block ---
        $desc = ''
        $descM = [regex]::Match($content, 'Description-Begin\s*-->(.+?)<!--\s*\w+-Description-End', $singleLine)
        if ($descM.Success) {
            $raw = $descM.Groups[1].Value
            $raw = $raw -replace '<!-- Description-Source-\w+ -->',''
            $raw = $raw -replace '<[^>]+>',''
            $raw = $raw -replace '&amp;','&' -replace '&lt;','<' -replace '&gt;','>' -replace '&#xA0;',' ' -replace '&quot;','"'
            $raw = $raw -replace '\s+',' '
            $desc = $raw.Trim()
            if ($desc.Length -gt 500) { $desc = $desc.Substring(0, 497) + '...' }
        }

        # --- Format & default from DFProperties sub-block ---
        $defaultValue = ''; $format = ''
        $dfM = [regex]::Match($content, 'DFProperties-Begin\s*-->(.+?)<!--\s*\w+-DFProperties-End', $singleLine)
        if ($dfM.Success) {
            $df = $dfM.Groups[1].Value
            if ($df -match 'Default Value\s*</td>\s*<td[^>]*>\s*(.+?)\s*</td>') {
                $defaultValue = ($Matches[1] -replace '<[^>]+>','').Trim()
            }
            if ($df -match '>Format\s*</td>\s*<td[^>]*>\s*(.+?)\s*</td>') {
                $format = ($Matches[1] -replace '<[^>]+>','').Trim()
            }
        }

        # --- Allowed values from AllowedValues sub-block ---
        $allowedValues = @{}
        $avM = [regex]::Match($content, 'AllowedValues-Begin\s*-->(.+?)<!--\s*\w+-AllowedValues-End', $singleLine)
        if ($avM.Success) {
            $avH = $avM.Groups[1].Value
            $avRows = [regex]::Matches($avH, '<td[^>]*>\s*(.+?)\s*</td>\s*<td[^>]*>\s*(.+?)\s*</td>', $singleLine)
            foreach ($row in $avRows) {
                $val = ($row.Groups[1].Value -replace '<[^>]+>','').Trim()
                $meaning = ($row.Groups[2].Value -replace '<[^>]+>','').Trim()
                $val = ($val -replace '\s*\(Default\)\s*','').Trim()
                if ($val -eq 'Value' -and $meaning -eq 'Description') { continue }
                if ($val -and $meaning) { $allowedValues[$val] = $meaning }
            }
        }
        # Fallback: check DFProperties for inline Allowed Values (Range/List)
        if ($allowedValues.Count -eq 0 -and $dfM.Success) {
            if ($dfM.Groups[1].Value -match 'Allowed Values\s*</td>\s*<td[^>]*>\s*(.+?)\s*</td>') {
                $avInline = ($Matches[1] -replace '<[^>]+>','').Trim()
                if ($avInline) { $allowedValues['_format'] = $avInline }
            }
        }

        # --- GP mapping from GpMapping sub-block ---
        $gpMapping = @{}
        $gpM = [regex]::Match($content, 'GpMapping-Begin\s*-->(.+?)<!--\s*\w+-GpMapping-End', $singleLine)
        if ($gpM.Success) {
            $gp = $gpM.Groups[1].Value
            foreach ($field in @('Name','Friendly Name','Element Name','Location','Path','Registry Key Name','Registry Value Name','ADMX File Name')) {
                $fm = [regex]::Match($gp, ">\s*$([regex]::Escape($field))\s*</td>\s*<td[^>]*>\s*(.+?)\s*</td>", $singleLine)
                if ($fm.Success) {
                    $gpMapping[$field] = ($fm.Groups[1].Value -replace '<[^>]+>','').Trim() -replace '&gt;','>' -replace '&amp;','&'
                }
            }
        }

        # --- Derive category ---
        $category = switch -Wildcard ($area) {
            'Defender'          { 'Endpoint Security: Antivirus' }
            'ExploitGuard'      { 'Endpoint Security: Attack Surface Reduction' }
            'Bitlocker'         { 'Endpoint Security: Disk Encryption' }
            'DeviceGuard'       { 'Device Security' }
            'DeviceLock'        { 'Device Security: Password' }
            'Update'            { 'Windows Update' }
            'DeliveryOptimization' { 'Windows Update: Delivery Optimization' }
            'Browser'           { 'Microsoft Edge (Legacy)' }
            'Wifi'              { 'Connectivity: WiFi' }
            'Bluetooth'         { 'Connectivity: Bluetooth' }
            'Connectivity'      { 'Connectivity' }
            'Privacy'           { 'Privacy' }
            'System'            { 'System' }
            'Experience'        { 'User Experience' }
            'Start'             { 'User Experience: Start Menu' }
            'Search'            { 'User Experience: Search' }
            'Power'             { 'Power Management' }
            'Camera'            { 'Hardware: Camera' }
            'Display'           { 'Hardware: Display' }
            'SmartScreen'       { 'Security: SmartScreen' }
            'ApplicationManagement' { 'App Management' }
            'DesktopAppInstaller'   { 'App Management: WinGet' }
            'Authentication'    { 'Authentication' }
            'Cryptography'      { 'Security: Cryptography' }
            'Security'          { 'Security' }
            'WindowsDefenderSecurityCenter' { 'Endpoint Security: Security Center' }
            'Notifications'     { 'User Experience: Notifications' }
            'Storage'           { 'Storage' }
            'RemoteDesktop'     { 'Remote Access: RDP' }
            'RemoteDesktopServices' { 'Remote Access: RDS' }
            'WindowsLogon'      { 'Authentication: Logon' }
            'LocalPoliciesSecurityOptions' { 'Security: Local Policies' }
            'Kerberos'          { 'Authentication: Kerberos' }
            'Accounts'          { 'Accounts' }
            'TextInput'         { 'User Experience: Input' }
            'WindowsAI'        { 'Windows AI' }
            'WindowsAutopilot'  { 'Deployment: Autopilot' }
            'Printers'          { 'Hardware: Printers' }
            'VirtualizationBasedTechnology' { 'Device Security: VBS' }
            'DmaGuard'          { 'Device Security: DMA' }
            'Cellular'          { 'Connectivity: Cellular' }
            'NetworkIsolation'  { 'Network: Isolation' }
            'EventLogService'   { 'System: Event Log' }
            'LanmanServer'      { 'Network: SMB Server' }
            'LanmanWorkstation' { 'Network: SMB Client' }
            'Audit'             { 'Security: Audit' }
            'UserRights'        { 'Security: User Rights' }
            'RestrictedGroups'  { 'Security: Restricted Groups' }
            'SecureBoot'        { 'Device Security: Secure Boot' }
            'MSSecurityGuide'   { 'Security: MS Security Guide' }
            'MSSLegacy'         { 'Security: MSS Legacy' }
            'TaskManager'       { 'System: Task Manager' }
            'TaskScheduler'     { 'System: Task Scheduler' }
            'WindowsPowerShell' { 'System: PowerShell' }
            'FileSystem'        { 'System: File System' }
            'Sudo'              { 'System: Sudo' }
            'WindowsSandbox'    { 'Security: Windows Sandbox' }
            'Speech'            { 'Privacy: Speech' }
            'Maps'              { 'System: Maps' }
            default             { $area -creplace '([a-z])([A-Z])','$1 $2' }
        }

        # --- Build entry ---
        $entry = [ordered]@{
            Friendly      = $settingName -creplace '([a-z])([A-Z])', '$1 $2'
            Desc          = $desc
            Cat           = $category
            Def           = $defaultValue
            Scope         = $scope
            Editions      = $editions -join ', '
            MinVersion    = $minVersion
            Format        = $format
            AllowedValues = $allowedValues
        }
        if ($gpMapping.Count -gt 0) {
            $entry['GPMapping'] = $gpMapping
        }

        $database[$key] = $entry
        $areaSettings++
        $totalSettings++
    }

    Write-Host "$areaSettings settings" -ForegroundColor Green
}

# --- Safe-replace: backup existing JSON, write to temp, validate, then swap ---
$json = $database | ConvertTo-Json -Depth 5 -Compress:$false
$minExpectedSettings = 100

if ($totalSettings -lt $minExpectedSettings) {
    Write-Warning "Scrape produced only $totalSettings settings (expected >=$minExpectedSettings). Aborting — existing database NOT overwritten."
    exit 1
}
if ($failedAreas.Count -gt ($targetAreas.Count / 2)) {
    Write-Warning "Too many failed areas ($($failedAreas.Count)/$($targetAreas.Count)). Aborting — existing database NOT overwritten."
    exit 1
}

# Backup existing file if present
if (Test-Path $OutputPath) {
    $backupPath = $OutputPath -replace '\.json$', '.bak.json'
    Copy-Item $OutputPath $backupPath -Force
    Write-Host "Backed up existing database to: $backupPath" -ForegroundColor DarkGray
}

# Write to temp file first, then move into place
$tempPath = "$OutputPath.tmp"
[System.IO.File]::WriteAllText($tempPath, $json, [System.Text.Encoding]::UTF8)
$tempSize = (Get-Item $tempPath).Length
if ($tempSize -lt 1KB) {
    Remove-Item $tempPath -Force
    Write-Warning "Generated JSON is suspiciously small ($tempSize bytes). Aborting — existing database NOT overwritten."
    exit 1
}
Move-Item $tempPath $OutputPath -Force

Write-Host "`n=== CSP Metadata Build Complete ===" -ForegroundColor Cyan
Write-Host "  Total settings: $totalSettings"
Write-Host "  Areas scraped:  $($targetAreas.Count - $failedAreas.Count) / $($targetAreas.Count)"
if ($failedAreas.Count -gt 0) {
    Write-Host "  Failed areas:   $($failedAreas -join ', ')" -ForegroundColor Yellow
}
Write-Host "  Output:         $OutputPath"
Write-Host "  File size:      $([math]::Round((Get-Item $OutputPath).Length / 1KB, 1)) KB"