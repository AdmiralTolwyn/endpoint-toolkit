<#
.SYNOPSIS
    Collects a full audio-stack configuration snapshot and diffs two snapshots.

.DESCRIPTION
    Captures a targeted audio configuration snapshot from documented APIs and
    observable registry state so two machines or image versions can be compared.
    It is not an exhaustive inventory of every component in the audio path.

    Collected areas:
      - OS build, UBR and the installed update history.
      - Windows audio services and their binary versions, including any
        third-party audio service registered as a dependent of Audiosrv.
      - Every MMDevices render and capture endpoint, its raw property values
        and its stored effects (FxProperties) configuration.
      - Audio Processing Objects (APOs) referenced by those endpoints, resolved
        through HKLM\SOFTWARE\Classes\CLSID to a DLL path, file version and
        Authenticode signer. Unresolved, invalid, unsigned and non-Microsoft
        modules are all flagged separately.
      - Audio endpoint drivers present on the machine, including virtual
        endpoints published by remoting stacks.
      - audiodg / Windows Audio reliability history (crashes and restarts).
      - Citrix products, Citrix audio policy registry state and the HDX audio
        driver version.
            - The legacy global USB DisableSelectiveSuspend registry value. This is
                not the effective per-device USB power policy.
      - The active sound scheme and the WAV files bound to notification events,
        with the measured peak level of each WAV.
      - Installed headset vendor software.
      - Collaboration-client VDI optimization state: new Teams'
        vdi_connection_info.json monitoring file(s), installed MSTeams Appx
        packages, the Citrix HDXMediaStream and MsTeamsPlugin registry keys,
        and Webex/Cisco products and services.

    APOs are software DSP components and are relevant comparison points. APO
    discovery here is heuristic: it resolves CLSID strings found in endpoint
    FxProperties and can miss binary or differently registered effects.

    The script is read-only. It writes nothing outside -OutputPath.

.PARAMETER OutputPath
    Directory to write the snapshot JSON to. The file is named
    AudioStackInventory-<COMPUTERNAME>-<yyyyMMdd-HHmmss>.json. When omitted the
    snapshot object is returned on the pipeline and nothing is written.

.PARAMETER CompareWith
    Path to a previously captured snapshot JSON. When supplied the script does
    not collect from the local machine; it compares the two snapshots and emits
    one object per difference. Requires -Baseline.

    The comparison is index-based: MMDevices endpoints, APOs and similar
    collections are matched by their position in the flattened attribute list,
    not by a stable identity. This is only meaningful for the same machine
    compared to itself over time, or for identical image clones. Comparing two
    different machines will surface endpoint-GUID differences (EndpointId,
    SoftwareId and similar per-install identifiers) as noise alongside any
    real configuration delta.

.PARAMETER Baseline
    Path to the snapshot JSON treated as the known-good side of the comparison.
    Used only with -CompareWith. See -CompareWith for the scope of what this
    diff can and cannot tell you.

.PARAMETER UpdateHistoryCount
    Number of most recent installed updates to record. Defaults to 15.

.PARAMETER MeasureSoundAssets
    Measure the peak level of every WAV bound to a system sound event. Off by
    default because an exact peak requires reading every sample of every asset,
    which takes noticeably longer than the rest of the collection combined.
    Worth enabling once per image rather than on every host.

.PARAMETER LobProcessNamePattern
    Regular expression matched against running process names (no extension)
    recorded in the LineOfBusiness section with path, version, start time and
    command line. Defaults to the Java launchers. Add the line-of-business
    application's own executable once its name is known, for example
    '^(java|javaw|myapp)$'.

.PARAMETER LogPath
    Optional path to a run log. When supplied, the script appends a line for
    the run start, each collection step and its duration, the output file
    (if any), every warning, and any terminating error before it is rethrown.
    Omitted by default; no log is written.

.EXAMPLE
    .\Get-AudioStackInventory.ps1 -OutputPath C:\temp\audio

    Capture the local machine's audio stack snapshot to a JSON file.

.EXAMPLE
    .\Get-AudioStackInventory.ps1 -Baseline .\good.json -CompareWith .\affected.json

    Diff an affected machine against a known-good baseline and list every
    attribute that differs.

.EXAMPLE
    Invoke-Command -ComputerName $vdaList -FilePath .\Get-AudioStackInventory.ps1

    Collect from a set of session hosts and return the objects for comparison.

.NOTES
    Author:   Anton Romanyuk
    Version:  1.0.0
    Requires: PowerShell 5.1. Run elevated for the broadest registry and file
              access. Missing or inaccessible data is retained as unresolved.

    The MMDevices property key GUIDs are emitted verbatim rather than resolved
    to friendly PKEY names. Endpoint friendly names are taken from the PnP
    subsystem instead, so no property-key mapping is assumed anywhere.

    Expect roughly a minute per machine. Authenticode verification of the APO
    modules dominates that time and is the reason the script exists, so it is
    not optional. Run -Verbose to see the per-section breakdown, and collect
    across a fleet with Invoke-Command rather than sequentially.

    Collaboration-client optimization state (CollaborationClients) is per
    user and per session, not a fixed host property: a Teams or Webex/Cisco
    engine's own state file only reflects whichever user and session last
    wrote it. Capture at incident time from the affected session, not once
    per host.
#>

#Requires -Version 5.1

[CmdletBinding(DefaultParameterSetName = 'Collect')]
param(
    [Parameter(ParameterSetName = 'Collect')]
    [string] $OutputPath,

    [Parameter(ParameterSetName = 'Collect')]
    [ValidateRange(0, 200)]
    [int] $UpdateHistoryCount = 15,

    [Parameter(ParameterSetName = 'Collect')]
    [switch] $MeasureSoundAssets,

    [Parameter(ParameterSetName = 'Collect')]
    [string] $LobProcessNamePattern = '^(java|javaw|javaws|jp2launcher)$',

    [Parameter(ParameterSetName = 'Compare', Mandatory)]
    [string] $CompareWith,

    [Parameter(ParameterSetName = 'Compare', Mandatory)]
    [string] $Baseline,

    [string] $LogPath
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$MMDEVICES_ROOT = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio'
$CLSID_ROOT = 'HKLM:\SOFTWARE\Classes\CLSID'
$CURRENT_VERSION = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
$GUID_PATTERN = '\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}'
$AUDIO_SERVICES = @('Audiosrv', 'AudioEndpointBuilder')
$SERVICES_ROOT = 'HKLM:\SYSTEM\CurrentControlSet\Services'
$DRIVER_NAME_PATTERN = 'audio|ctxad|portcls|hdaudio|usbaud|^ks$|vaudio|rdpaud'
$PROFILE_LIST_ROOT = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
$TEAMS_VDI_JSON_SUFFIX = 'AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\tfw\vdi_connection_info.json'

# Authenticode verification dominates collection time when the same binary is
# reached from several endpoints, so both lookups are memoised for the run.
$Script:FileFactsCache = @{}
$Script:ComServerCache = @{}

<#
.SYNOPSIS
    Reads a single registry value without throwing when it is absent.

.PARAMETER Path
    Registry key path to read from.

.PARAMETER Name
    Value name to read.

.OUTPUTS
    System.Object. The value data, or $null when the key or value is missing.
#>
function Get-RegValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [string] $Name
    )

    try {
        $item = Get-ItemProperty -LiteralPath $Path -Name $Name -ErrorAction Stop
    } catch {
        return $null
    }

    if ($null -eq $item) { return $null }
    if (-not $item.PSObject.Properties.Match($Name).Count) { return $null }

    return $item.$Name
}

<#
.SYNOPSIS
    Returns every value under a registry key as name/data pairs.

.DESCRIPTION
    PowerShell metadata properties (PSPath, PSParentPath and friends) are
    excluded so the result can be diffed directly between two machines.

.PARAMETER Path
    Registry key path to enumerate.

.OUTPUTS
    System.Collections.Specialized.OrderedDictionary
#>
function Get-RegValueMap {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Path
    )

    $map = [ordered] @{}

    try {
        $item = Get-ItemProperty -LiteralPath $Path -ErrorAction Stop
    } catch {
        return $map
    }

    if ($null -eq $item) { return $map }

    foreach ($property in $item.PSObject.Properties) {
        if ($property.Name -like 'PS*') { continue }

        $value = $property.Value
        if ($value -is [byte[]]) {
            $value = ($value | ForEach-Object { $_.ToString('x2') }) -join ''
        }

        $map[$property.Name] = $value
    }

    return $map
}

<#
.SYNOPSIS
    Resolves a file to its version and Authenticode signer.

.PARAMETER Path
    Full path to the file to inspect.

.OUTPUTS
    System.Management.Automation.PSCustomObject
#>
function Get-FileFacts {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [AllowEmptyString()] [string] $Path
    )

    $facts = [pscustomobject] @{
        Path                   = $Path
        Exists                 = $false
        FileVersion            = $null
        ProductVersion         = $null
        SignerSubject          = $null
        SignatureStatus        = $null
        SignatureStatusMessage = $null
        IsMicrosoft            = $null
        TrustClass             = 'Unresolved'
    }

    if ([string]::IsNullOrWhiteSpace($Path)) { return $facts }

    $expanded = [System.Environment]::ExpandEnvironmentVariables($Path.Trim('"'))

    if ($Script:FileFactsCache.ContainsKey($expanded)) {
        return $Script:FileFactsCache[$expanded]
    }

    if (-not (Test-Path -LiteralPath $expanded -PathType Leaf)) {
        $facts.Path = $expanded
        $Script:FileFactsCache[$expanded] = $facts
        return $facts
    }

    $facts.Path = $expanded
    $facts.Exists = $true

    try {
        $info = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($expanded)
        $facts.FileVersion = $info.FileVersion
        $facts.ProductVersion = $info.ProductVersion
    } catch {
        Write-Verbose "Version info unavailable for $expanded : $($_.Exception.Message)"
    }

    try {
        $signature = Get-AuthenticodeSignature -LiteralPath $expanded -ErrorAction Stop
        $facts.SignatureStatus = [string] $signature.Status
        if ($signature.Status -ne 'Valid') {
            $facts.SignatureStatusMessage = $signature.StatusMessage
        }
        if ($null -ne $signature.SignerCertificate) {
            $facts.SignerSubject = $signature.SignerCertificate.Subject
            $facts.IsMicrosoft = ($signature.Status -eq 'Valid' -and $signature.SignerCertificate.Subject -match 'O=Microsoft Corporation')
            if ($signature.Status -ne 'Valid') {
                # UnknownError/NotTrusted with a Microsoft-subject signer is the
                # offline/no-CRL shape: the chain cannot be verified from this
                # host, not necessarily a bad signature. Named separately so an
                # analyst knows which of those two cases they are looking at.
                if (($signature.Status -eq 'UnknownError' -or $signature.Status -eq 'NotTrusted') -and $signature.SignerCertificate.Subject -match 'O=Microsoft Corporation') {
                    $facts.TrustClass = 'SignedChainNotVerifiable'
                } else {
                    $facts.TrustClass = 'InvalidOrUntrustedSignature'
                }
            } elseif ($facts.IsMicrosoft) {
                $facts.TrustClass = 'MicrosoftValid'
            } else {
                $facts.TrustClass = 'ThirdPartyValid'
            }
        } else {
            $facts.IsMicrosoft = $false
            $facts.TrustClass = 'Unsigned'
        }
    } catch {
        Write-Verbose "Signature unavailable for $expanded : $($_.Exception.Message)"
    }

    $Script:FileFactsCache[$expanded] = $facts
    return $facts
}

<#
.SYNOPSIS
    Reads a property from a parsed-JSON object graph by dotted path.

.DESCRIPTION
    ConvertFrom-Json produces PSCustomObject graphs, and under
    Set-StrictMode -Version 2.0 dotted access to a missing property throws.
    This walks the path one segment at a time via PSObject.Properties.Match
    so a missing segment anywhere in the path returns $null instead of
    throwing, without disabling StrictMode for the caller.

.PARAMETER InputObject
    Parsed JSON object (or any object) to read from. $null is tolerated.

.PARAMETER Path
    Dotted property path, e.g. 'vdiConnectedState.timestamp'.

.OUTPUTS
    System.Object. The property value, or $null when any segment is absent.
#>
function Get-JsonProperty {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [AllowNull()] $InputObject,
        [Parameter(Mandatory)] [string] $Path
    )

    $current = $InputObject

    foreach ($segment in ($Path -split '\.')) {
        if ($null -eq $current) { return $null }

        $match = $current.PSObject.Properties.Match($segment)
        if ($match.Count -eq 0) { return $null }

        $current = $match[0].Value
    }

    return $current
}

<#
.SYNOPSIS
    Returns the profile root directory of every local user profile.

.DESCRIPTION
    Reads ProfileImagePath from each subkey of the ProfileList registry key
    rather than listing C:\Users, so a profile redirected to another volume
    or path is still found. Each candidate path is verified to exist before
    being returned; a profile this identity cannot enumerate or read is
    silently skipped rather than failing the whole collection.

.OUTPUTS
    System.String[]
#>
function Get-UserProfileRoots {
    [CmdletBinding()]
    param()

    $results = New-Object System.Collections.Generic.List[string]

    try {
        foreach ($key in @(Get-ChildItem -LiteralPath $PROFILE_LIST_ROOT -ErrorAction Stop)) {
            $imagePath = Get-RegValue -Path $key.PSPath -Name 'ProfileImagePath'
            if ([string]::IsNullOrWhiteSpace($imagePath)) { continue }
            if (-not (Test-Path -LiteralPath $imagePath -PathType Container)) { continue }
            if (-not $results.Contains($imagePath)) { [void] $results.Add($imagePath) }
        }
    } catch {
        Write-Verbose "ProfileList enumeration unavailable: $($_.Exception.Message)"
    }

    return $results.ToArray()
}

<#
.SYNOPSIS
    Resolves a COM CLSID to its in-process server module.

.PARAMETER Clsid
    CLSID in registry brace form.

.OUTPUTS
    System.Management.Automation.PSCustomObject
#>
function Resolve-ComServer {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Clsid
    )

    if ($Script:ComServerCache.ContainsKey($Clsid)) {
        return $Script:ComServerCache[$Clsid]
    }

    $keyPath = Join-Path $CLSID_ROOT $Clsid
    $friendly = Get-RegValue -Path $keyPath -Name '(default)'
    $module = Get-RegValue -Path (Join-Path $keyPath 'InprocServer32') -Name '(default)'

    $result = [pscustomobject] @{
        Clsid        = $Clsid
        Registered   = (Test-Path -LiteralPath $keyPath)
        FriendlyName = $friendly
        Module       = $null
    }

    if (-not [string]::IsNullOrWhiteSpace($module)) {
        $result.Module = Get-FileFacts -Path $module
    }

    $Script:ComServerCache[$Clsid] = $result
    return $result
}

<#
.SYNOPSIS
    Collects OS build identity and recent update history.

.OUTPUTS
    System.Management.Automation.PSCustomObject
#>
function Get-OsFacts {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [int] $HistoryCount
    )

    $build = Get-RegValue -Path $CURRENT_VERSION -Name 'CurrentBuildNumber'
    $ubr = Get-RegValue -Path $CURRENT_VERSION -Name 'UBR'

    $fullBuild = $null
    if ($null -ne $build -and $null -ne $ubr) {
        $fullBuild = '{0}.{1}' -f $build, $ubr
    }

    $updates = @()
    if ($HistoryCount -gt 0) {
        try {
            $updates = @(
                Get-HotFix -ErrorAction Stop |
                    Sort-Object -Property InstalledOn -Descending |
                    Select-Object -First $HistoryCount |
                    ForEach-Object {
                        [pscustomobject] @{
                            HotFixId    = $_.HotFixId
                            Description = $_.Description
                            InstalledOn = $_.InstalledOn
                        }
                    }
            )
        } catch {
            Write-Verbose "Update history unavailable: $($_.Exception.Message)"
        }
    }

    return [pscustomobject] @{
        ProductName      = Get-RegValue -Path $CURRENT_VERSION -Name 'ProductName'
        DisplayVersion   = Get-RegValue -Path $CURRENT_VERSION -Name 'DisplayVersion'
        EditionId        = Get-RegValue -Path $CURRENT_VERSION -Name 'EditionID'
        CurrentBuild     = $build
        Ubr              = $ubr
        FullBuild        = $fullBuild
        InstallationType = Get-RegValue -Path $CURRENT_VERSION -Name 'InstallationType'
        RecentUpdates    = $updates
    }
}

<#
.SYNOPSIS
    Collects the state of the Windows audio services and their dependents.

.DESCRIPTION
    Any service registered as a dependent of Audiosrv sits in the audio
    control path. On a remoting host this is where the vendor audio service
    appears, and its version is part of the image delta that must be compared
    alongside the Windows build.

.OUTPUTS
    System.Management.Automation.PSCustomObject[]
#>
function Get-AudioServiceFacts {
    [CmdletBinding()]
    param()

    $names = New-Object System.Collections.Generic.List[string]
    foreach ($name in $AUDIO_SERVICES) { [void] $names.Add($name) }

    foreach ($name in $AUDIO_SERVICES) {
        try {
            $service = Get-Service -Name $name -ErrorAction Stop
        } catch {
            continue
        }

        foreach ($dependent in @($service.DependentServices)) {
            if (-not $names.Contains($dependent.Name)) { [void] $names.Add($dependent.Name) }
        }
    }

    $results = New-Object System.Collections.Generic.List[object]

    foreach ($name in $names) {
        $servicePath = "HKLM:\SYSTEM\CurrentControlSet\Services\$name"
        $imagePath = Get-RegValue -Path $servicePath -Name 'ImagePath'
        $serviceDll = Get-RegValue -Path (Join-Path $servicePath 'Parameters') -Name 'ServiceDll'

        $binary = $serviceDll
        if ([string]::IsNullOrWhiteSpace($binary)) { $binary = $imagePath }

        $status = $null
        $startType = $null
        try {
            $service = Get-Service -Name $name -ErrorAction Stop
            $status = [string] $service.Status
            $startType = [string] (Get-RegValue -Path $servicePath -Name 'Start')
        } catch {
            Write-Verbose "Service $name not queryable: $($_.Exception.Message)"
        }

        [void] $results.Add([pscustomobject] @{
            Name          = $name
            Status        = $status
            StartValue    = $startType
            ImagePath     = $imagePath
            ServiceDll    = $serviceDll
            BinaryFacts   = Get-FileFacts -Path $binary
            IsDependent   = (-not ($AUDIO_SERVICES -contains $name))
        })
    }

    return $results.ToArray()
}

<#
.SYNOPSIS
    Collects every MMDevices endpoint with its raw properties and effects.

.DESCRIPTION
    Property key GUIDs are returned verbatim. Endpoint names are resolved from
    the PnP subsystem where possible rather than by mapping property keys, so
    no assumption is made about which key holds the friendly name.

.OUTPUTS
    System.Management.Automation.PSCustomObject[]
#>
function Get-EndpointFacts {
    [CmdletBinding()]
    param()

    $pnpNames = @{}
    try {
        foreach ($device in @(Get-PnpDevice -Class 'AudioEndpoint' -ErrorAction Stop)) {
            if ([string]::IsNullOrWhiteSpace($device.InstanceId)) { continue }
            $match = [regex]::Match($device.InstanceId, $GUID_PATTERN)
            if ($match.Success) { $pnpNames[$match.Value.ToLowerInvariant()] = $device.FriendlyName }
        }
    } catch {
        Write-Verbose "PnP audio endpoint enumeration unavailable: $($_.Exception.Message)"
    }

    $results = New-Object System.Collections.Generic.List[object]

    foreach ($flow in @('Render', 'Capture')) {
        $flowRoot = Join-Path $MMDEVICES_ROOT $flow
        if (-not (Test-Path -LiteralPath $flowRoot)) { continue }

        foreach ($endpoint in @(Get-ChildItem -LiteralPath $flowRoot -ErrorAction SilentlyContinue)) {
            $id = $endpoint.PSChildName
            $name = $null
            if ($pnpNames.ContainsKey($id.ToLowerInvariant())) { $name = $pnpNames[$id.ToLowerInvariant()] }

            $properties = Get-RegValueMap -Path (Join-Path $endpoint.PSPath 'Properties')
            $fxProperties = Get-RegValueMap -Path (Join-Path $endpoint.PSPath 'FxProperties')

            # The registry value carries undocumented high-order bits beyond
            # DEVICE_STATEMASK_ALL; only the low nibble is the documented state.
            $rawState = Get-RegValue -Path $endpoint.PSPath -Name 'DeviceState'
            $stateName = $null
            if ($null -ne $rawState) {
                switch ([int] $rawState -band 0xF) {
                    1 { $stateName = 'Active' }
                    2 { $stateName = 'Disabled' }
                    4 { $stateName = 'NotPresent' }
                    8 { $stateName = 'Unplugged' }
                    default { $stateName = 'Unknown' }
                }
            }

            [void] $results.Add([pscustomobject] @{
                DataFlow        = $flow
                EndpointId      = $id
                FriendlyName    = $name
                DeviceState     = $rawState
                DeviceStateName = $stateName
                Properties      = $properties
                FxProperties    = $fxProperties
            })
        }
    }

    return $results.ToArray()
}

<#
.SYNOPSIS
    Extracts and resolves every APO referenced by the endpoint effects data.

.DESCRIPTION
    Scans the FxProperties of each endpoint for CLSID-shaped values and
    resolves each one to its implementing module. APOs run inside audiodg.exe
    and are able to change sample amplitude, so a non-Microsoft signer here is
    a direct candidate for an amplitude artefact.

.PARAMETER Endpoint
    Endpoint objects produced by Get-EndpointFacts.

.OUTPUTS
    System.Management.Automation.PSCustomObject[]
#>
function Get-ApoFacts {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [object[]] $Endpoint
    )

    $seen = @{}
    $results = New-Object System.Collections.Generic.List[object]

    foreach ($item in $Endpoint) {
        if ($null -eq $item.FxProperties) { continue }

        foreach ($key in $item.FxProperties.Keys) {
            $data = [string] $item.FxProperties[$key]
            if ([string]::IsNullOrWhiteSpace($data)) { continue }

            foreach ($match in [regex]::Matches($data, $GUID_PATTERN)) {
                $clsid = $match.Value
                # GUID_NULL marks an empty effect slot, not an APO.
                if ($clsid -eq '{00000000-0000-0000-0000-000000000000}') { continue }
                $cacheKey = '{0}|{1}' -f $item.EndpointId, $clsid
                if ($seen.ContainsKey($cacheKey)) { continue }
                $seen[$cacheKey] = $true

                $server = Resolve-ComServer -Clsid $clsid
                $isMicrosoft = $null
                $modulePath = $null
                $moduleVersion = $null
                $signer = $null
                $signatureStatus = $null
                $trustClass = 'UnresolvedRegistration'
                if ($null -ne $server.Module) {
                    $isMicrosoft = $server.Module.IsMicrosoft
                    $modulePath = $server.Module.Path
                    $moduleVersion = $server.Module.FileVersion
                    $signer = $server.Module.SignerSubject
                    $signatureStatus = $server.Module.SignatureStatus
                    $trustClass = $server.Module.TrustClass
                }

                [void] $results.Add([pscustomobject] @{
                    EndpointId    = $item.EndpointId
                    DataFlow      = $item.DataFlow
                    FriendlyName  = $item.FriendlyName
                    PropertyKey   = $key
                    Clsid         = $clsid
                    Registered    = $server.Registered
                    ComName       = $server.FriendlyName
                    ModulePath    = $modulePath
                    ModuleVersion = $moduleVersion
                    Signer        = $signer
                    SignatureStatus = $signatureStatus
                    TrustClass     = $trustClass
                    IsMicrosoft   = $isMicrosoft
                    NeedsReview    = ($server.Registered -and $trustClass -ne 'MicrosoftValid')
                })
            }
        }
    }

    return $results.ToArray()
}

<#
.SYNOPSIS
    Collects audio-class and virtual audio drivers present on the machine.

.DESCRIPTION
    Enumerated from the service registry rather than through CIM. The CIM driver
    class takes roughly fifteen seconds to materialise on a normal machine,
    which dominates the whole collection; the registry holds the same fields and
    returns immediately.

.OUTPUTS
    System.Management.Automation.PSCustomObject[]
#>
function Get-AudioDriverFacts {
    [CmdletBinding()]
    param()

    $results = New-Object System.Collections.Generic.List[object]

    foreach ($key in @(Get-ChildItem -LiteralPath $SERVICES_ROOT -ErrorAction SilentlyContinue)) {
        $name = $key.PSChildName
        if ($name -notmatch $DRIVER_NAME_PATTERN) { continue }

        $type = Get-RegValue -Path $key.PSPath -Name 'Type'
        if ($type -ne 1 -and $type -ne 2) { continue }

        $imagePath = [string] (Get-RegValue -Path $key.PSPath -Name 'ImagePath')
        $normalised = $imagePath
        if (-not [string]::IsNullOrWhiteSpace($normalised)) {
            $normalised = $normalised -replace '^\\\?\?\\', ''
            $normalised = $normalised -replace '^\\SystemRoot\\', '%SystemRoot%\'
            if ($normalised -match '^[Ss]ystem32\\') { $normalised = '%SystemRoot%\' + $normalised }
        }

        [void] $results.Add([pscustomobject] @{
            Name        = $name
            DisplayName = Get-RegValue -Path $key.PSPath -Name 'DisplayName'
            StartValue  = Get-RegValue -Path $key.PSPath -Name 'Start'
            TypeValue   = $type
            ImagePath   = $imagePath
            BinaryFacts = Get-FileFacts -Path $normalised
        })
    }

    return $results.ToArray()
}

<#
.SYNOPSIS
    Collects audiodg crash records and Windows Audio service restarts.

.DESCRIPTION
    A restart of the audio device graph re-initialises every endpoint. That is
    one of the few documented events able to produce a level discontinuity on
    an already-open render stream, and it is cheap to rule in or out.

.OUTPUTS
    System.Management.Automation.PSCustomObject
#>
function Get-AudioReliabilityFacts {
    [CmdletBinding()]
    param()

    $crashes = @()
    $restarts = @()
    $applicationOldest = $null
    $systemOldest = $null

    try {
        # audiodg is the render/effects host process; its binary name is not
        # localized so this text match is safe on Message. ToXml() is tested as
        # well in case Message is unavailable (missing provider manifest on a
        # foreign-language image).
        $applicationEvents = @(Get-WinEvent -FilterHashtable @{ LogName = 'Application'; Id = @(1000, 1001) } -MaxEvents 5000 -ErrorAction Stop)
        if ($applicationEvents.Count -gt 0) {
            $applicationOldest = ($applicationEvents | Select-Object -Last 1).TimeCreated
        }

        $crashes = @(
            $applicationEvents |
                Where-Object { $_.Message -match 'audiodg' -or $_.ToXml() -match 'audiodg' } |
                Select-Object -First 25 |
                ForEach-Object {
                    [pscustomobject] @{
                        TimeCreated = $_.TimeCreated
                        Id          = $_.Id
                        Provider    = $_.ProviderName
                        Message     = ($_.Message -split "`r?`n" | Select-Object -First 3) -join ' '
                    }
                }
        )
    } catch {
        Write-Verbose "Application log query failed: $($_.Exception.Message)"
    }

    try {
        # Filtered by event Id and provider rather than by matching rendered
        # Message text, because Message is localized and param0 (the service
        # display name) varies by build. ToXml() is the primary test since it
        # carries the raw EventData regardless of locale; the Properties[0]
        # check is a fallback for hosts where ToXml() text still differs.
        $systemEvents = @(Get-WinEvent -FilterHashtable @{ LogName = 'System'; ProviderName = 'Service Control Manager'; Id = @(7031, 7034, 7036, 7040, 7000, 7001, 7009, 7011, 7023, 7024) } -MaxEvents 5000 -ErrorAction Stop)
        if ($systemEvents.Count -gt 0) {
            $systemOldest = ($systemEvents | Select-Object -Last 1).TimeCreated
        }

        $restarts = @(
            $systemEvents |
                Where-Object {
                    $xmlMatch = $_.ToXml() -match 'Audiosrv|AudioEndpointBuilder|CtxAudioSvc'
                    $param0Match = $false
                    if ($_.Properties.Count -gt 0) {
                        $param0Match = $_.Properties[0].Value -match '^(Audiosrv|AudioEndpointBuilder|Windows Audio|Windows Audio Endpoint Builder|CtxAudioSvc|Citrix Audio)'
                    }
                    $xmlMatch -or $param0Match
                } |
                Select-Object -First 25 |
                ForEach-Object {
                    [pscustomobject] @{
                        TimeCreated = $_.TimeCreated
                        Id          = $_.Id
                        Message     = ($_.Message -split "`r?`n" | Select-Object -First 2) -join ' '
                    }
                }
        )
    } catch {
        Write-Verbose "System log query failed: $($_.Exception.Message)"
    }

    return [pscustomobject] @{
        AudiodgFaults               = $crashes
        AudioServiceMessages        = $restarts
        ApplicationLogOldestScanned = $applicationOldest
        SystemLogOldestScanned      = $systemOldest
        Note                        = 'Absence of a record within the scanned window is not proof the component behaved correctly.'
    }
}

<#
.SYNOPSIS
    Collects remoting-stack products, audio policy state and USB power settings.

.OUTPUTS
    System.Management.Automation.PSCustomObject
#>
function Get-RemotingAndUsbFacts {
    [CmdletBinding()]
    param()

    $uninstallRoots = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    )

    $products = New-Object System.Collections.Generic.List[object]

    foreach ($root in $uninstallRoots) {
        if (-not (Test-Path -LiteralPath $root)) { continue }

        foreach ($key in @(Get-ChildItem -LiteralPath $root -ErrorAction SilentlyContinue)) {
            $displayName = Get-RegValue -Path $key.PSPath -Name 'DisplayName'
            if ([string]::IsNullOrWhiteSpace($displayName)) { continue }
            if ($displayName -notmatch 'citrix|jabra|poly|plantronics|epos|sennheiser|dell|teams|audio|webex|cisco') { continue }

            [void] $products.Add([pscustomobject] @{
                DisplayName    = $displayName
                DisplayVersion = Get-RegValue -Path $key.PSPath -Name 'DisplayVersion'
                Publisher      = Get-RegValue -Path $key.PSPath -Name 'Publisher'
                InstallDate    = Get-RegValue -Path $key.PSPath -Name 'InstallDate'
            })
        }
    }

    # Keyed by a root-qualified relative path (e.g. 'Policies\Citrix\...' vs
    # 'Citrix\...') rather than the bare relative path, because both policy
    # roots can contain a key of the same relative name (for example an HDX
    # policy key and an unrelated product key both named "Audio"), which would
    # otherwise silently collide in the same dictionary. The prefix is derived
    # from the root itself so it stays correct if the root list changes.
    $citrixPolicy = [ordered] @{}
    foreach ($policyRoot in @('HKLM:\SOFTWARE\Policies\Citrix', 'HKLM:\SOFTWARE\Citrix')) {
        if (-not (Test-Path -LiteralPath $policyRoot)) { continue }

        $rootPrefix = $policyRoot -replace '^HKLM:\\SOFTWARE\\', ''

        foreach ($key in @(Get-ChildItem -LiteralPath $policyRoot -Recurse -ErrorAction SilentlyContinue)) {
            if ($key.Name -notmatch 'audio|multimedia|sound') { continue }
            $values = Get-RegValueMap -Path $key.PSPath
            if ($values.Count -gt 0) {
                $relativePath = $key.Name.Substring($policyRoot.Replace('HKLM:', 'HKEY_LOCAL_MACHINE').Length).TrimStart('\')
                $qualifiedPath = if ([string]::IsNullOrEmpty($relativePath)) { $rootPrefix } else { Join-Path $rootPrefix $relativePath }
                $citrixPolicy[$qualifiedPath] = $values
            }
        }
    }

    return [pscustomobject] @{
        RelevantProducts        = $products.ToArray()
        CitrixAudioPolicyKeys   = $citrixPolicy
        LegacyUsbDisableSelectiveSuspend = Get-RegValue -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\USB' -Name 'DisableSelectiveSuspend'
    }
}

<#
.SYNOPSIS
    Records classic Outlook alert settings and Java-hosted line-of-business
    processes, two common application-layer producers of notification audio.

.DESCRIPTION
    A classic Outlook alert and a Java-hosted business application are audio
    producers that a lab image without the production application set does
    not reproduce. Both sit above the Windows audio engine:
    classic Outlook's desktop alert is a toast plus, when "Play a sound" is on,
    the AppEvents MailBeep binding (reported in SoundScheme); a Java
    application renders through javax.sound.sampled or a bundled engine and
    appears in the state monitor's session CSV under its process name.

    Outlook value names under HKCU\Software\Microsoft\Office\16.0\Outlook are
    observed implementation locations; a value that is absent means Outlook is
    using its default (desktop alert on, sound on), and a policy value under
    HKCU\Software\Policies overrides it. Reads only; nothing is changed.

.PARAMETER ProcessNamePattern
    Regular expression matched against running process names (without
    extension) to capture as line-of-business processes. Defaults to Java
    launchers; add the application's own executable name when known.

.OUTPUTS
    System.Management.Automation.PSCustomObject
#>
function Get-LineOfBusinessFacts {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $ProcessNamePattern
    )

    $outlookRoot = 'HKCU:\Software\Microsoft\Office\16.0\Outlook'
    $outlookPolicyRoot = 'HKCU:\Software\Policies\Microsoft\Office\16.0\Outlook'
    $alertValuePattern = '^(NewmailDesktopAlerts|PlaySound|ShowEnvelope|ChangeCursor|UseNewOutlook|NewOutlookMigration.*|DesktopAlert.*)$'
    $reminderValuePattern = '^(PlaySound|ReminderSoundFile|Beep.*|Reminder.*)$'

    $outlookPath = Get-RegValue -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE' -Name '(default)'
    $outlookFileVersion = $null
    if (-not [string]::IsNullOrWhiteSpace($outlookPath) -and (Test-Path -LiteralPath $outlookPath -PathType Leaf)) {
        $outlookFileVersion = (Get-Item -LiteralPath $outlookPath).VersionInfo.FileVersion
    }

    $filteredPrefs = [ordered] @{}
    $prefs = Get-RegValueMap -Path (Join-Path $outlookRoot 'Preferences')
    foreach ($name in @($prefs.Keys)) { if ($name -match $alertValuePattern) { $filteredPrefs[$name] = $prefs[$name] } }

    $filteredReminders = [ordered] @{}
    $reminders = Get-RegValueMap -Path (Join-Path $outlookRoot 'Options\Reminders')
    foreach ($name in @($reminders.Keys)) { if ($name -match $reminderValuePattern) { $filteredReminders[$name] = $reminders[$name] } }

    $policyPrefs = Get-RegValueMap -Path (Join-Path $outlookPolicyRoot 'Preferences')
    $policyReminders = Get-RegValueMap -Path (Join-Path $outlookPolicyRoot 'Options\Reminders')

    $c2r = Get-RegValueMap -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    $c2rFacts = [ordered] @{}
    foreach ($name in @('VersionToReport', 'UpdateChannel', 'CDNBaseUrl', 'Platform', 'ProductReleaseIds')) {
        if ($c2r.Contains($name)) { $c2rFacts[$name] = $c2r[$name] }
    }

    $outlookProcesses = @(Get-Process -Name 'OUTLOOK', 'olk' -ErrorAction SilentlyContinue | ForEach-Object {
        [pscustomobject] @{ Name = $_.ProcessName; Id = $_.Id; Path = $_.Path }
    })

    $uninstallRoots = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall'
    )
    $javaProducts = New-Object System.Collections.Generic.List[object]
    foreach ($root in $uninstallRoots) {
        if (-not (Test-Path -LiteralPath $root)) { continue }
        foreach ($key in @(Get-ChildItem -LiteralPath $root -ErrorAction SilentlyContinue)) {
            $displayName = Get-RegValue -Path $key.PSPath -Name 'DisplayName'
            if ([string]::IsNullOrWhiteSpace($displayName)) { continue }
            if ($displayName -notmatch 'java|jdk|jre|openjdk|temurin|zulu|corretto|semeru|liberica|graalvm') { continue }
            [void] $javaProducts.Add([pscustomobject] @{
                DisplayName     = $displayName
                DisplayVersion  = Get-RegValue -Path $key.PSPath -Name 'DisplayVersion'
                Publisher       = Get-RegValue -Path $key.PSPath -Name 'Publisher'
                InstallLocation = Get-RegValue -Path $key.PSPath -Name 'InstallLocation'
            })
        }
    }

    $commandLines = @{}
    try {
        foreach ($proc in @(Get-CimInstance -ClassName Win32_Process -ErrorAction Stop)) {
            $commandLines[[int] $proc.ProcessId] = $proc.CommandLine
        }
    } catch {
        Write-Verbose "Win32_Process command lines unavailable: $($_.Exception.Message)"
    }

    $lobProcesses = @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -match $ProcessNamePattern } | ForEach-Object {
        $info = $null
        try { $info = $_.MainModule.FileVersionInfo } catch { }
        $startTime = $null
        try { $startTime = $_.StartTime.ToUniversalTime().ToString('o') } catch { }
        [pscustomobject] @{
            Name         = $_.ProcessName
            Id           = $_.Id
            Path         = $_.Path
            FileVersion  = if ($null -ne $info) { $info.FileVersion } else { $null }
            ProductName  = if ($null -ne $info) { $info.ProductName } else { $null }
            CompanyName  = if ($null -ne $info) { $info.CompanyName } else { $null }
            StartedUtc   = $startTime
            CommandLine  = if ($commandLines.ContainsKey([int] $_.Id)) { $commandLines[[int] $_.Id] } else { $null }
        }
    })

    return [pscustomobject] @{
        Outlook = [pscustomobject] @{
            ExecutablePath       = $outlookPath
            FileVersion          = $outlookFileVersion
            ClickToRun           = $c2rFacts
            AlertPreferences     = $filteredPrefs
            ReminderOptions      = $filteredReminders
            PolicyPreferences    = $policyPrefs
            PolicyReminderOptions = $policyReminders
            RunningProcesses     = $outlookProcesses
            Note                 = 'Absent values mean Outlook defaults (desktop alert on, play a sound on). New-mail sound is the AppEvents MailBeep binding in SoundScheme; the desktop alert itself is a toast. Value names are observed locations, not a documented contract.'
        }
        Java = [pscustomobject] @{
            InstalledRuntimes = $javaProducts.ToArray()
            JavaHomeMachine   = [System.Environment]::GetEnvironmentVariable('JAVA_HOME', 'Machine')
            JavaHomeUser      = [System.Environment]::GetEnvironmentVariable('JAVA_HOME', 'User')
            ProcessNamePattern = $ProcessNamePattern
            RunningProcesses  = $lobProcesses
            Note              = 'A Java application renders through javax.sound.sampled (DirectSound shared mode on Windows) or a bundled media engine and appears in the state monitor session CSV under its process name. Command lines are read for the collecting user; other users need elevation.'
        }
    }
}

<#
.SYNOPSIS
    Collects new Teams' vdi_connection_info.json monitoring file(s).

.DESCRIPTION
    Per Microsoft (learn.microsoft.com/microsoftteams/vdi-2, "Monitoring
    API"), new Teams writes this file on the VDA at
    %LOCALAPPDATA%\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\tfw\
    for the connected user. Only vdiConnectedState is populated when WebRTC
    optimization is active; when no optimization is available the file is
    not updated at all, so its LastWriteTimeUtc can predate the session.

    The current user's copy is always checked via $env:LOCALAPPDATA. When
    the collecting identity is elevated, every local profile is also checked
    so a session host can be inspected without impersonating the affected
    user; a profile this identity cannot read is skipped, not failed.

.PARAMETER IsElevated
    Whether the collecting identity holds local Administrator rights.

.OUTPUTS
    System.Management.Automation.PSCustomObject[]
#>
function Get-TeamsVdiConnectionInfoFacts {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [bool] $IsElevated
    )

    $note = 'Per Microsoft (learn.microsoft.com/microsoftteams/vdi-2, "Monitoring API"): under WebRTC optimization only vdiConnectedState is populated; when no optimization is available the file is not updated.'

    $paths = New-Object System.Collections.Generic.List[string]

    if (-not [string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) {
        $currentUserPath = Join-Path $env:LOCALAPPDATA 'Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\tfw\vdi_connection_info.json'
        [void] $paths.Add($currentUserPath)
    }

    if ($IsElevated) {
        foreach ($profileRoot in @(Get-UserProfileRoots)) {
            $candidate = Join-Path $profileRoot $TEAMS_VDI_JSON_SUFFIX
            if (-not $paths.Contains($candidate)) { [void] $paths.Add($candidate) }
        }
    }

    $results = New-Object System.Collections.Generic.List[object]

    foreach ($path in $paths) {
        $entry = [pscustomobject] @{
            ProfilePath           = $path
            Exists                = $false
            LastWriteTimeUtc      = $null
            Length                = $null
            Timestamp             = $null
            VdiMode               = $null
            ConnectedStack        = $null
            RemoteSlimCoreVersion = $null
            BridgeVersion         = $null
            PluginVersion         = $null
            TeamsVersion          = $null
            ClientPlatform        = $null
            RdClientVersion       = $null
            VmVersion             = $null
            SelectedSpeaker       = $null
            SecondaryRinger       = $null
            RawJson               = $null
            ParseError            = $null
            Note                  = $note
        }

        try {
            if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
                [void] $results.Add($entry)
                continue
            }

            $entry.Exists = $true

            $fileInfo = Get-Item -LiteralPath $path -ErrorAction Stop
            $entry.LastWriteTimeUtc = $fileInfo.LastWriteTimeUtc
            $entry.Length = $fileInfo.Length

            $text = Get-Content -LiteralPath $path -Raw -ErrorAction Stop
            if ($text.Length -gt 16384) {
                $entry.RawJson = $text.Substring(0, 16384)
            } else {
                $entry.RawJson = $text
            }

            try {
                $json = $text | ConvertFrom-Json -ErrorAction Stop

                $entry.Timestamp             = Get-JsonProperty -InputObject $json -Path 'vdiConnectedState.timestamp'
                $entry.VdiMode               = Get-JsonProperty -InputObject $json -Path 'vdiConnectedState.vdiMode'
                $entry.ConnectedStack        = Get-JsonProperty -InputObject $json -Path 'vdiConnectedState.connectedStack'
                $entry.RemoteSlimCoreVersion = Get-JsonProperty -InputObject $json -Path 'remoteSlimCoreVersion'
                $entry.BridgeVersion         = Get-JsonProperty -InputObject $json -Path 'bridgeVersion'
                $entry.PluginVersion         = Get-JsonProperty -InputObject $json -Path 'pluginVersion'
                $entry.TeamsVersion          = Get-JsonProperty -InputObject $json -Path 'vdiVersionInfo.teamsVersion'
                $entry.ClientPlatform        = Get-JsonProperty -InputObject $json -Path 'vdiVersionInfo.clientPlatform'
                $entry.RdClientVersion       = Get-JsonProperty -InputObject $json -Path 'vdiVersionInfo.rdClientVersion'
                $entry.VmVersion             = Get-JsonProperty -InputObject $json -Path 'vdiVersionInfo.vmVersion'
                $entry.SelectedSpeaker       = Get-JsonProperty -InputObject $json -Path 'devices.speakers.selected'
                $entry.SecondaryRinger       = Get-JsonProperty -InputObject $json -Path 'devices.secondaryRinger'
            } catch {
                $entry.ParseError = $_.Exception.Message
            }
        } catch {
            $entry.ParseError = $_.Exception.Message
        }

        [void] $results.Add($entry)
    }

    return $results.ToArray()
}

<#
.SYNOPSIS
    Collects installed MSTeams Appx package facts.

.PARAMETER IsElevated
    Whether the collecting identity holds local Administrator rights. Only
    when true is the -AllUsers package registration also queried.

.OUTPUTS
    System.Management.Automation.PSCustomObject[]
#>
function Get-TeamsPackageFacts {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [bool] $IsElevated
    )

    $results = New-Object System.Collections.Generic.List[object]

    try {
        foreach ($pkg in @(Get-AppxPackage -Name 'MSTeams' -ErrorAction Stop)) {
            [void] $results.Add([pscustomobject] @{
                Name            = $pkg.Name
                Version         = [string] $pkg.Version
                InstallLocation = $pkg.InstallLocation
                Scope           = 'CurrentUser'
                UserSids        = @()
            })
        }
    } catch {
        Write-Verbose "Get-AppxPackage (current user) unavailable: $($_.Exception.Message)"
    }

    if ($IsElevated) {
        try {
            foreach ($pkg in @(Get-AppxPackage -Name 'MSTeams' -AllUsers -ErrorAction Stop)) {
                $sids = @()
                if ($null -ne $pkg.PackageUserInformation) {
                    $sids = @(
                        $pkg.PackageUserInformation | ForEach-Object {
                            if ($null -ne $_.UserSecurityId) { [string] $_.UserSecurityId }
                        }
                    )
                }

                [void] $results.Add([pscustomobject] @{
                    Name            = $pkg.Name
                    Version         = [string] $pkg.Version
                    InstallLocation = $pkg.InstallLocation
                    Scope           = 'AllUsers'
                    UserSids        = $sids
                })
            }
        } catch {
            Write-Verbose "Get-AppxPackage -AllUsers unavailable: $($_.Exception.Message)"
        }
    }

    return $results.ToArray()
}

<#
.SYNOPSIS
    Returns uninstall registry entries whose DisplayName matches a pattern.

.DESCRIPTION
    Scans the same two per-machine Uninstall roots enumerated in
    Get-RemotingAndUsbFacts. Kept as its own scan rather than extending that
    function's result set, so that function's established output shape is
    not disturbed by an unrelated product filter.

.PARAMETER Pattern
    Regular expression tested against DisplayName.

.OUTPUTS
    System.Management.Automation.PSCustomObject[]
#>
function Get-UninstallEntriesMatching {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Pattern
    )

    $uninstallRoots = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    )

    $results = New-Object System.Collections.Generic.List[object]

    foreach ($root in $uninstallRoots) {
        if (-not (Test-Path -LiteralPath $root)) { continue }

        foreach ($key in @(Get-ChildItem -LiteralPath $root -ErrorAction SilentlyContinue)) {
            $displayName = Get-RegValue -Path $key.PSPath -Name 'DisplayName'
            if ([string]::IsNullOrWhiteSpace($displayName)) { continue }
            if ($displayName -notmatch $Pattern) { continue }

            [void] $results.Add([pscustomobject] @{
                DisplayName    = $displayName
                DisplayVersion = Get-RegValue -Path $key.PSPath -Name 'DisplayVersion'
                Publisher      = Get-RegValue -Path $key.PSPath -Name 'Publisher'
                InstallDate    = Get-RegValue -Path $key.PSPath -Name 'InstallDate'
            })
        }
    }

    return $results.ToArray()
}

<#
.SYNOPSIS
    Collects Webex/Cisco services present on the machine.

.DESCRIPTION
    Enumerated from the service registry rather than through CIM, for the
    same reason as Get-AudioDriverFacts: the registry holds the same fields
    and returns immediately.

.OUTPUTS
    System.Management.Automation.PSCustomObject[]
#>
function Get-WebexCiscoServiceFacts {
    [CmdletBinding()]
    param()

    $results = New-Object System.Collections.Generic.List[object]

    foreach ($key in @(Get-ChildItem -LiteralPath $SERVICES_ROOT -ErrorAction SilentlyContinue)) {
        $name = $key.PSChildName
        $displayName = Get-RegValue -Path $key.PSPath -Name 'DisplayName'

        $nameMatches = ($name -match 'webex|cisco')
        $displayMatches = (-not [string]::IsNullOrWhiteSpace($displayName) -and $displayName -match 'webex|cisco')
        if (-not $nameMatches -and -not $displayMatches) { continue }

        $imagePath = Get-RegValue -Path $key.PSPath -Name 'ImagePath'
        $status = $null
        try {
            $service = Get-Service -Name $name -ErrorAction Stop
            $status = [string] $service.Status
        } catch {
            Write-Verbose "Service $name not queryable: $($_.Exception.Message)"
        }

        [void] $results.Add([pscustomobject] @{
            Name        = $name
            DisplayName = $displayName
            Status      = $status
            ImagePath   = $imagePath
            BinaryFacts = Get-FileFacts -Path $imagePath
        })
    }

    return $results.ToArray()
}

<#
.SYNOPSIS
    Collects collaboration-client (Teams, Webex/Cisco) VDI-optimization facts.

.DESCRIPTION
    Surfaces the state new Teams and known third-party collaboration clients
    report about VDI/HDX media optimization, alongside the Citrix and
    Microsoft registry keys that describe the plugin side of that
    negotiation. This is read-only: it does not determine whether
    optimization is active for any given session, only what the installed
    software and its last-written state file(s) currently record.

    This section's facts are per user and per session, not a fixed host
    property. See the returned object's Note fields for the caveats specific
    to each sub-section.

.OUTPUTS
    System.Management.Automation.PSCustomObject
#>
function Get-CollaborationClientFacts {
    [CmdletBinding()]
    param()

    $isElevated = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    return [pscustomobject] @{
        TeamsMonitoringFiles  = @(Get-TeamsVdiConnectionInfoFacts -IsElevated $isElevated)
        TeamsPackages         = @(Get-TeamsPackageFacts -IsElevated $isElevated)
        CitrixHdxMediaStream  = [pscustomobject] @{
            Hklm        = Get-RegValueMap -Path 'HKLM:\SOFTWARE\Citrix\HDXMediaStream'
            HklmWow6432 = Get-RegValueMap -Path 'HKLM:\SOFTWARE\WOW6432Node\Citrix\HDXMediaStream'
            Hkcu        = Get-RegValueMap -Path 'HKCU:\Software\Citrix\HDXMediaStream'
            Note        = 'HKCU is the collecting identity''s hive, not necessarily the hive of the user who experiences the artefact.'
        }
        MsTeamsPluginRegistry = [pscustomobject] @{
            HklmWow6432 = Get-RegValueMap -Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Teams\MsTeamsPlugin'
            Hklm        = Get-RegValueMap -Path 'HKLM:\SOFTWARE\Microsoft\Teams\MsTeamsPlugin'
        }
        WebexAndCisco         = [pscustomobject] @{
            Products = @(Get-UninstallEntriesMatching -Pattern 'webex|cisco')
            Services = @(Get-WebexCiscoServiceFacts)
            Note     = 'Citrix states it does not instrument the Webex VDI plugin or Cisco Media Engine; engine state must come from Cisco.'
        }
        Note                  = 'Optimization state is per user and per session; capture at incident time, not once per host. Presence of a sound in the VDA loopback proves the HDX audio path; sounds rendered by the endpoint Teams or Webex engines never reach the VDA render endpoint.'
    }
}

<#
.SYNOPSIS
    Collects the sound scheme and measures the peak level of each bound WAV.

.DESCRIPTION
    Reads the peak sample of every WAV file wired to a system sound event. A
    notification asset mastered close to full scale sounds far louder than
    level-controlled speech even when the whole stack behaves correctly, so the
    asset levels have to be measured before a defect is assumed.

    Only uncompressed PCM WAV files are measured. Anything else is reported
    with a null peak.

.PARAMETER MeasurePeak
    Measure the peak level of each bound WAV. Skipped when absent.

.NOTES
    Reads HKCU:\AppEvents\Schemes, which is per-user. When this script runs
    under Invoke-Command or in an elevated session, HKCU resolves to the
    collecting account's own hive, not necessarily the hive of the user who
    experiences the artefact. The returned SoundSchemeHive field records the
    collecting identity's SID for that reason.

.OUTPUTS
    System.Management.Automation.PSCustomObject
#>
function Get-SoundSchemeFacts {
    [CmdletBinding()]
    param(
        [switch] $MeasurePeak
    )

    $schemesRoot = 'HKCU:\AppEvents\Schemes'
    $current = Get-RegValue -Path $schemesRoot -Name '(default)'
    $sounds = New-Object System.Collections.Generic.List[object]

    $appsRoot = Join-Path $schemesRoot 'Apps'
    if (Test-Path -LiteralPath $appsRoot) {
        foreach ($app in @(Get-ChildItem -LiteralPath $appsRoot -ErrorAction SilentlyContinue)) {
            foreach ($soundEvent in @(Get-ChildItem -LiteralPath $app.PSPath -ErrorAction SilentlyContinue)) {
                $currentKey = Join-Path $soundEvent.PSPath '.Current'
                $file = Get-RegValue -Path $currentKey -Name '(default)'
                if ([string]::IsNullOrWhiteSpace($file)) { continue }

                $expanded = [System.Environment]::ExpandEnvironmentVariables($file)

                $peak = $null
                if ($MeasurePeak) { $peak = Measure-WavPeakDbfs -Path $expanded }

                [void] $sounds.Add([pscustomobject] @{
                    App        = $app.PSChildName
                    Event      = $soundEvent.PSChildName
                    File       = $expanded
                    Exists     = (Test-Path -LiteralPath $expanded -PathType Leaf)
                    PeakDbfs   = $peak
                })
            }
        }
    }

    return [pscustomobject] @{
        CurrentScheme   = $current
        Sounds          = $sounds.ToArray()
        SoundSchemeHive = 'HKCU of the collecting identity (UserSid)'
    }
}

<#
.SYNOPSIS
    Measures the peak level of a 16-bit PCM WAV file in dBFS.

.PARAMETER Path
    Full path to the WAV file.

.OUTPUTS
    System.Nullable[System.Double]. Null when the file is absent or not 16-bit PCM.
#>
function Measure-WavPeakDbfs {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [AllowEmptyString()] [string] $Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $null }

    try {
        $bytes = [System.IO.File]::ReadAllBytes($Path)
    } catch {
        return $null
    }

    if ($bytes.Length -lt 44) { return $null }
    if ([System.Text.Encoding]::ASCII.GetString($bytes, 0, 4) -ne 'RIFF') { return $null }

    $offset = 12
    $bitsPerSample = 0
    $formatTag = 0
    $dataOffset = -1
    $dataLength = 0

    while ($offset + 8 -le $bytes.Length) {
        $chunkId = [System.Text.Encoding]::ASCII.GetString($bytes, $offset, 4)
        $chunkSize = [System.BitConverter]::ToUInt32($bytes, $offset + 4)

        if ($chunkId -eq 'fmt ') {
            $formatTag = [System.BitConverter]::ToUInt16($bytes, $offset + 8)
            $bitsPerSample = [System.BitConverter]::ToUInt16($bytes, $offset + 22)
        } elseif ($chunkId -eq 'data') {
            $dataOffset = $offset + 8
            $dataLength = [int] [Math]::Min([int64] $chunkSize, [int64] ($bytes.Length - $dataOffset))
            break
        }

        $offset += 8 + $chunkSize
        if ($chunkSize % 2 -eq 1) { $offset++ }
    }

    if ($dataOffset -lt 0 -or $dataLength -le 1) { return $null }
    if ($formatTag -ne 1 -or $bitsPerSample -ne 16) { return $null }

    $peak = 0
    for ($i = $dataOffset; $i -lt ($dataOffset + $dataLength - 1); $i += 2) {
        $sample = [Math]::Abs([int] [System.BitConverter]::ToInt16($bytes, $i))
        if ($sample -gt $peak) { $peak = $sample }
    }

    if ($peak -le 0) { return $null }

    return [Math]::Round(20.0 * [Math]::Log10($peak / 32768.0), 2)
}

<#
.SYNOPSIS
    Flattens a snapshot object into dotted attribute paths for comparison.

.PARAMETER InputObject
    Object to flatten.

.PARAMETER Prefix
    Attribute path prefix used during recursion.

.PARAMETER Result
    Dictionary accumulating the flattened attributes.
#>
function Expand-ObjectGraph {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [AllowNull()] $InputObject,
        [Parameter(Mandatory)] [string] $Prefix,
        [Parameter(Mandatory)] [System.Collections.Generic.Dictionary[string, string]] $Result
    )

    if ($null -eq $InputObject) {
        $Result[$Prefix] = '<null>'
        return
    }

    if ($InputObject -is [string] -or $InputObject.GetType().IsPrimitive -or $InputObject -is [datetime] -or $InputObject -is [decimal]) {
        $Result[$Prefix] = [string] $InputObject
        return
    }

    if ($InputObject -is [System.Collections.IDictionary]) {
        foreach ($key in $InputObject.Keys) {
            Expand-ObjectGraph -InputObject $InputObject[$key] -Prefix ('{0}.{1}' -f $Prefix, $key) -Result $Result
        }
        return
    }

    if ($InputObject -is [System.Collections.IEnumerable]) {
        $index = 0
        foreach ($element in $InputObject) {
            Expand-ObjectGraph -InputObject $element -Prefix ('{0}[{1}]' -f $Prefix, $index) -Result $Result
            $index++
        }
        return
    }

    foreach ($property in $InputObject.PSObject.Properties) {
        Expand-ObjectGraph -InputObject $property.Value -Prefix ('{0}.{1}' -f $Prefix, $property.Name) -Result $Result
    }
}

<#
.SYNOPSIS
    Compares two snapshot files and emits one object per differing attribute.

.PARAMETER BaselinePath
    Snapshot treated as known good.

.PARAMETER CandidatePath
    Snapshot to test against the baseline.

.OUTPUTS
    System.Management.Automation.PSCustomObject[]
#>
function Compare-Snapshot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $BaselinePath,
        [Parameter(Mandatory)] [string] $CandidatePath
    )

    $baselineObject = Get-Content -LiteralPath $BaselinePath -Raw | ConvertFrom-Json
    $candidateObject = Get-Content -LiteralPath $CandidatePath -Raw | ConvertFrom-Json

    if ($baselineObject.ComputerName -ne $candidateObject.ComputerName) {
        Write-Warning ("Comparing snapshots from different computers ('{0}' vs '{1}'). This diff is index-based and only meaningful for the same machine over time or identical image clones; comparing different machines will surface endpoint-GUID differences as noise." -f $baselineObject.ComputerName, $candidateObject.ComputerName)
    }

    $baselineMap = New-Object 'System.Collections.Generic.Dictionary[string,string]'
    $candidateMap = New-Object 'System.Collections.Generic.Dictionary[string,string]'

    Expand-ObjectGraph -InputObject $baselineObject -Prefix 'root' -Result $baselineMap
    Expand-ObjectGraph -InputObject $candidateObject -Prefix 'root' -Result $candidateMap

    $keys = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($key in $baselineMap.Keys) { [void] $keys.Add($key) }
    foreach ($key in $candidateMap.Keys) { [void] $keys.Add($key) }

    $differences = New-Object System.Collections.Generic.List[object]

    foreach ($key in ($keys | Sort-Object)) {
        if ($key -eq 'root.CollectedUtc' -or $key -eq 'root.ComputerName') { continue }

        $inBaseline = $baselineMap.ContainsKey($key)
        $inCandidate = $candidateMap.ContainsKey($key)

        $baselineValue = $null
        $candidateValue = $null
        if ($inBaseline) { $baselineValue = $baselineMap[$key] }
        if ($inCandidate) { $candidateValue = $candidateMap[$key] }

        if ($inBaseline -and $inCandidate -and $baselineValue -eq $candidateValue) { continue }

        $change = 'Changed'
        if (-not $inBaseline) { $change = 'AddedOnCandidate' }
        elseif (-not $inCandidate) { $change = 'MissingOnCandidate' }

        [void] $differences.Add([pscustomobject] @{
            Attribute = $key
            Change    = $change
            Baseline  = $baselineValue
            Candidate = $candidateValue
        })
    }

    return $differences.ToArray()
}

<#
.SYNOPSIS
    Runs one collection step and reports how long it took.

.PARAMETER Name
    Step name for the verbose record.

.PARAMETER Action
    Script block performing the collection.

.OUTPUTS
    System.Object. Whatever the script block returns.
#>
function Invoke-CollectionStep {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Name,
        [Parameter(Mandatory)] [scriptblock] $Action
    )

    Write-RunLog -Message "Starting: $Name"
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $result = & $Action
    $stopwatch.Stop()
    Write-Verbose ('{0} completed in {1:F1}s' -f $Name, $stopwatch.Elapsed.TotalSeconds)
    Write-RunLog -Message ('Completed: {0} in {1:F1}s' -f $Name, $stopwatch.Elapsed.TotalSeconds)

    return $result
}

<#
.SYNOPSIS
    Appends a timestamped line to -LogPath, if one was supplied.

.PARAMETER Message
    Text to record.
#>
function Write-RunLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Message
    )

    if ([string]::IsNullOrWhiteSpace($Script:LogPath)) { return }

    $line = '{0:yyyy-MM-dd HH:mm:ss.fff}  {1}' -f (Get-Date), $Message
    Add-Content -LiteralPath $Script:LogPath -Value $line -Encoding UTF8
}

Write-RunLog -Message "Run started. ParameterSetName=$($PSCmdlet.ParameterSetName)"

try {
    if ($PSCmdlet.ParameterSetName -eq 'Compare') {
        $differences = @(Compare-Snapshot -BaselinePath $Baseline -CandidatePath $CompareWith)
        Write-RunLog -Message ('Compare completed. Baseline={0} CompareWith={1} Differences={2}' -f $Baseline, $CompareWith, $differences.Count)
        $differences
        return
    }

    $os = Invoke-CollectionStep -Name 'Operating system' -Action { Get-OsFacts -HistoryCount $UpdateHistoryCount }
    $endpoints = @(Invoke-CollectionStep -Name 'Endpoints' -Action { Get-EndpointFacts } | Sort-Object DataFlow, EndpointId)
    $apos = @(Invoke-CollectionStep -Name 'Processing objects' -Action { Get-ApoFacts -Endpoint $endpoints } | Sort-Object EndpointId, Clsid, PropertyKey)
    $services = @(Invoke-CollectionStep -Name 'Services' -Action { Get-AudioServiceFacts } | Sort-Object Name)
    $drivers = @(Invoke-CollectionStep -Name 'Drivers' -Action { Get-AudioDriverFacts } | Sort-Object Name)
    $reliability = Invoke-CollectionStep -Name 'Reliability' -Action { Get-AudioReliabilityFacts }
    $remoting = Invoke-CollectionStep -Name 'Remoting and USB' -Action { Get-RemotingAndUsbFacts }
    $collaboration = Invoke-CollectionStep -Name 'Collaboration clients' -Action { Get-CollaborationClientFacts }
    $soundScheme = Invoke-CollectionStep -Name 'Sound scheme' -Action { Get-SoundSchemeFacts -MeasurePeak:$MeasureSoundAssets }
    $lineOfBusiness = Invoke-CollectionStep -Name 'Line-of-business producers' -Action { Get-LineOfBusinessFacts -ProcessNamePattern $LobProcessNamePattern }

    $snapshot = [pscustomobject] @{
        SchemaVersion     = '1.0'
        CollectedUtc      = (Get-Date).ToUniversalTime().ToString('o')
        ComputerName      = $env:COMPUTERNAME
        UserSid           = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
        IsElevated        = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        OperatingSystem   = $os
        AudioServices     = $services
        AudioDrivers      = $drivers
        Endpoints         = $endpoints
        ProcessingObjects = $apos
        Reliability       = $reliability
        RemotingAndUsb    = $remoting
        CollaborationClients = $collaboration
        SoundScheme       = $soundScheme
        LineOfBusiness    = $lineOfBusiness
    }

    $needsReview = @($apos | Where-Object { $_.NeedsReview })
    if ($needsReview.Count -gt 0) {
        Write-Warning "$($needsReview.Count) APO reference(s) require review:"
        Write-RunLog -Message "$($needsReview.Count) APO reference(s) require review:"
        foreach ($apo in $needsReview) {
            Write-Warning "  $($apo.DataFlow) $($apo.Clsid) -> $($apo.ModulePath) [$($apo.TrustClass)]"
            Write-RunLog -Message "  $($apo.DataFlow) $($apo.Clsid) -> $($apo.ModulePath) [$($apo.TrustClass)]"
        }
    }

    if ($PSBoundParameters.ContainsKey('OutputPath') -and -not [string]::IsNullOrWhiteSpace($OutputPath)) {
        if (-not (Test-Path -LiteralPath $OutputPath -PathType Container)) {
            [void] (New-Item -Path $OutputPath -ItemType Directory -Force)
        }

        $fileName = 'AudioStackInventory-{0}-{1}.json' -f $env:COMPUTERNAME, (Get-Date -Format 'yyyyMMdd-HHmmss')
        $target = Join-Path $OutputPath $fileName
        $json = $snapshot | ConvertTo-Json -Depth 12
        [System.IO.File]::WriteAllText($target, $json, [System.Text.UTF8Encoding]::new($true))
        Write-Verbose "Snapshot written to $target"
        Write-RunLog -Message "Snapshot written to $target"
    }

    Write-RunLog -Message 'Run completed successfully.'
    $snapshot
} catch {
    Write-RunLog -Message "ERROR: $($_.Exception.Message)"
    throw
}
