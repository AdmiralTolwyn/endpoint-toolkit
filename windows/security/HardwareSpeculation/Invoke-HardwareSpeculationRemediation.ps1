<#
.SYNOPSIS
    Detection + remediation for Speculative Store Bypass (SSB / ADV180012) and
    Microarchitectural Data Sampling (MDS / ADV190013) hardware vulnerabilities.

.DESCRIPTION
    Companion remediation to Get-HardwareSpeculationStatus.ps1.

    1. DETECTION:
       - Queries CPU speculation control flags via ntdll!NtQuerySystemInformation
         (class 201), the same low-level interface used by the Microsoft
         SpeculationControl module.
       - Reads the current FeatureSettingsOverride / FeatureSettingsOverrideMask
         registry values and decodes which mitigations are already forced on/off.

    2. REMEDIATION (only with -ForceRemediation):
       - Writes the single FeatureSettingsOverride value that satisfies BOTH the
         SSB and MDS scanner checks, plus FeatureSettingsOverrideMask = 3:
            * Default (HT kept):      FeatureSettingsOverride = 72   (0x48)
            * -DisableHyperThreading: FeatureSettingsOverride = 8264 (0x2048)
         Both values enable the full documented mitigation set (MDS, TAA, Spectre,
         Meltdown, SSBD, L1TF) per KB4073119; 8264 additionally disables
         Hyper-Threading for complete L1TF/MDS coverage at a performance cost.
       - On Hyper-V hosts, also sets MinVmVersionForCpuBasedMitigations = "1.0"
         under ...\CurrentVersion\Virtualization (required by QID 91537).
       - A reboot is required. On Hyper-V hosts, fully shut down VMs before reboot
         so the firmware mitigation is applied to the host first.

    SCANNER NOTES - these QIDs check the REGISTRY CONFIG, not the live CPU/ntdll flag:
       - Qualys QID 91462 (SSB / ADV180012): wants FeatureSettingsOverride present
         with an accepted value (8, 72, 8264, 8388616, 8388680, 8396872) AND
         FeatureSettingsOverrideMask = 3.
       - Qualys QID 91537 (MDS/TAA / ADV190013, Intel servers): EXACT match -
         FeatureSettingsOverride = 72 OR 8264, FeatureSettingsOverrideMask = 3,
         and (Hyper-V hosts) MinVmVersionForCpuBasedMitigations = "1.0".
       The ONLY values satisfying BOTH QIDs are 72 and 8264, which is why this
       script targets those (not the minimal 0x8). A device can report "not
       required" at runtime yet still be flagged purely on the missing keys.

.PARAMETER ForceRemediation
    ENABLES the registry write. Without this switch the script is detection-only
    and does NOT modify the system, making it safe to run as a compliance detect
    script (Ivanti, Intune Proactive Remediations detect phase, ConfigMgr CI).

.PARAMETER DisableHyperThreading
    Use the HT-disabled mitigation value (FeatureSettingsOverride = 8264 / 0x2048)
    instead of the default (72 / 0x48). This gives complete L1TF/MDS coverage but
    DISABLES Hyper-Threading/SMT, which can significantly reduce performance. Only
    use where your security policy explicitly requires Hyper-Threading disabled.

.PARAMETER ForceExactValue
    Overwrite FeatureSettingsOverride to EXACTLY 72 (or 8264) and the mask to 3,
    discarding any other bits currently set. By default the script is ADDITIVE
    (it preserves existing higher-order mitigation bits such as BHI 0x800000 and
    only forces the Spectre V2/Meltdown disable bits off + sets SSBD/MDS bits on).
    QID 91537 (MDS) requires an EXACT 72/8264 match, so if a device already carries
    extra mitigation bits the additive result will not satisfy that QID; use this
    switch to force the exact value. WARNING: this can REMOVE other mitigations
    (e.g. BHI/MMIO) - only use it when you are sure those are not needed.

.NOTES
    Author:  Anton Romanyuk
    Version: 1.0
    Context: ADV180012 (SSB) / ADV190013 (MDS) - KB4073119
    Source:  https://support.microsoft.com/en-us/topic/kb4073119-windows-client-guidance-for-it-pros-to-protect-against-silicon-based-microarchitectural-and-speculative-execution-side-channel-vulnerabilities-35820a8a-ae13-1299-88cc-357f104f5b11

    DISCLAIMER:
    THIS SCRIPT IS PROVIDED "AS-IS" WITHOUT WARRANTY OF ANY KIND. USE AT YOUR
    OWN RISK. Enabling mitigations that are off by default can affect device
    performance. Test before broad deployment.
#>

[CmdletBinding()]
param (
    [switch]$ForceRemediation,
    [switch]$DisableHyperThreading,
    [switch]$ForceExactValue
)

# -------------------------------------------------------------------------------------------------
# 1. HELPER FUNCTIONS
# -------------------------------------------------------------------------------------------------

function Write-ColorLog {
    param (
        [Parameter(Mandatory = $true)] [string]$Message,
        [Parameter(Mandatory = $true)] [ValidateSet('Info', 'Success', 'Warning', 'Error', 'Verbose')] [string]$Level
    )
    $colorMap = @{ 'Info' = 'Cyan'; 'Success' = 'Green'; 'Warning' = 'Yellow'; 'Error' = 'Red'; 'Verbose' = 'Gray' }
    Write-Host "[$((Get-Date).ToString('HH:mm:ss'))] [$Level] $Message" -ForegroundColor $colorMap[$Level]
}

# FeatureSettingsOverride bit semantics (KB4073119).
# NOTE: bits 0/1 use INVERTED logic (set = DISABLE the mitigation); bit 3 is normal (set = ENABLE SSBD).
$script:BIT_DISABLE_SPECTRE_V2 = 0x1     # set => Spectre V2 (CVE-2017-5715) DISABLED
$script:BIT_DISABLE_MELTDOWN   = 0x2     # set => Meltdown (CVE-2017-5754) DISABLED
$script:BIT_ENABLE_SSBD        = 0x8     # set => SSBD (CVE-2018-3639) ENABLED
$script:BIT_ENABLE_MDS_L1TF    = 0x40    # set => MDS/TAA/L1TF mitigation override ENABLED
$script:BIT_DISABLE_HT         = 0x2000  # set => Hyper-Threading DISABLED (full L1TF/MDS)
$script:MASK_SPECTRE_MELTDOWN  = 0x3     # mask bits the OS must honor for Spectre/Meltdown override

# Scanner-accepted target values (KB4073119 combined-mitigation values). Both contain
# SSBD (0x8) + the MDS/TAA/L1TF override bit (0x40); 8264 also sets the HT-disable bit
# (0x2000). These are the only values that satisfy BOTH Qualys QID 91462 (SSB) and
# QID 91537 (MDS exact-match).
$script:TARGET_OVERRIDE_NO_HT       = 72    # 0x48   - full mitigation set, HT NOT disabled
$script:TARGET_OVERRIDE_HT_DISABLED = 8264  # 0x2048 - full mitigation set, HT disabled (L1TF/MDS complete)
$script:TARGET_MASK                 = 3     # 0x3

# Hyper-V host CPU-mitigation key required by QID 91537.
$script:HyperVRegPath  = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Virtualization'
$script:HyperVRegName  = 'MinVmVersionForCpuBasedMitigations'
$script:HyperVRegValue = '1.0'

# Test whether the registry already satisfies the SSBD config that scanners (Qualys
# QID 91462) require: SSBD bit (0x8) set, neither Spectre V2 nor Meltdown disabled
# (bits 0x1/0x2 clear), and the override mask honoring those bits (Mask & 0x3 == 0x3).
function Test-SsbdRegCompliant {
    param ($Reg)
    return (
        ($null -ne $Reg.Override) -and (($Reg.Override -band $script:BIT_ENABLE_SSBD) -ne 0) -and
        (($Reg.Override -band ($script:BIT_DISABLE_SPECTRE_V2 -bor $script:BIT_DISABLE_MELTDOWN)) -eq 0) -and
        ($null -ne $Reg.Mask) -and (($Reg.Mask -band $script:MASK_SPECTRE_MELTDOWN) -eq $script:MASK_SPECTRE_MELTDOWN)
    )
}

# Test whether the registry satisfies the MDS/TAA config that Qualys QID 91537 requires:
# FeatureSettingsOverride EXACTLY 72 or 8264, with the mask honoring bits 0/1
# (Mask & 0x3 == 0x3). Intel servers; clients are enabled by default.
function Test-MdsRegCompliant {
    param ($Reg)
    return (
        ($null -ne $Reg.Override) -and
        (($Reg.Override -eq $script:TARGET_OVERRIDE_NO_HT) -or ($Reg.Override -eq $script:TARGET_OVERRIDE_HT_DISABLED)) -and
        ($null -ne $Reg.Mask) -and (($Reg.Mask -band $script:MASK_SPECTRE_MELTDOWN) -eq $script:MASK_SPECTRE_MELTDOWN)
    )
}

function Get-FeatureSettingsDecoding {
    param ([Nullable[int]]$Override, [Nullable[int]]$Mask)
    $lines = @()
    if ($null -eq $Override) {
        $lines += "FeatureSettingsOverride: NotSet (OS defaults apply)"
    }
    else {
        $lines += ("FeatureSettingsOverride: 0x{0:X}" -f $Override)
        if ($Override -band $script:BIT_DISABLE_SPECTRE_V2) { $lines += "  bit0 (0x1) SET -> Spectre V2 mitigation DISABLED" }
        if ($Override -band $script:BIT_DISABLE_MELTDOWN)   { $lines += "  bit1 (0x2) SET -> Meltdown mitigation DISABLED" }
        if ($Override -band $script:BIT_ENABLE_SSBD)        { $lines += "  bit3 (0x8) SET -> SSBD ENABLED" }
        if ($Override -band 0x40)                           { $lines += "  bit6 (0x40) SET -> MDS/TAA/L1TF mitigation override enabled" }
        if ($Override -band 0x2000)                         { $lines += "  bit13 (0x2000) SET -> Hyper-Threading DISABLED" }
        if ($Override -band 0x800000)                       { $lines += "  bit23 (0x800000) SET -> BHI mitigation forced on" }
    }
    if ($null -eq $Mask) {
        $lines += "FeatureSettingsOverrideMask: NotSet"
    }
    else {
        $lines += ("FeatureSettingsOverrideMask: 0x{0:X}" -f $Mask)
    }
    return $lines
}

# -------------------------------------------------------------------------------------------------
# 2. CORE DETECTION LOGIC
# -------------------------------------------------------------------------------------------------

function Get-SpeculationStatus {
    Write-ColorLog -Message "Querying CPU speculation control flags (ntdll class 201)..." -Level "Info"

    # Speculation-control flags (ntdll class 201). See ADV180002 / KB4073119.
    # SsbdRequired is the STATIC hardware-vulnerability flag (never clears); SsbdSystemWide is
    # the flag that flips once SSBD is actually active (registry + reboot + CPU microcode).
    $scfMdsHardwareProtected = 0x1000000
    $scfSsbdAvailable        = 0x100     # OS support for SSBD present
    $scfSsbdSupported        = 0x200     # CPU microcode supports SSBD
    $scfSsbdSystemWide       = 0x400     # SSBD ENABLED system-wide
    $scfSsbdRequired         = 0x1000    # hardware is VULNERABLE / requires SSBD (STATIC)

    $results = @{
        MdsVulnerable    = $false
        SsbVulnerable    = $false
        SsbdSystemWide   = $false
        SsbdAvailable    = $false
        SsbdSupported    = $false
        Success          = $false
        RawFlags         = $null
        CpuManufacturer  = $null
    }

    $NtQSIDefinition = @'
    [DllImport("ntdll.dll")]
    public static extern int NtQuerySystemInformation(uint systemInformationClass, IntPtr systemInformation, uint systemInformationLength, IntPtr returnLength);
'@

    try {
        $ntdll = Add-Type -MemberDefinition $NtQSIDefinition -Name 'ntdllSpec' -Namespace 'Win32Spec' -PassThru -ErrorAction Stop
    }
    catch {
        Write-ColorLog -Message "Failed to load ntdll definition: $($_.Exception.Message)" -Level "Error"
        return $results
    }

    $len    = 8
    [IntPtr]$ptr    = [System.Runtime.InteropServices.Marshal]::AllocHGlobal($len)
    [IntPtr]$retPtr = [System.Runtime.InteropServices.Marshal]::AllocHGlobal(4)

    try {
        $cpu = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue
        $results.CpuManufacturer = $cpu.Manufacturer

        $retval = $ntdll::NtQuerySystemInformation(201, $ptr, $len, $retPtr)
        if ($retval -ne 0) {
            Write-ColorLog -Message "NtQuerySystemInformation failed (status 0x$($retval.ToString('X')))." -Level "Error"
            return $results
        }

        $flags = [uint32][System.Runtime.InteropServices.Marshal]::ReadInt32($ptr)
        $results.RawFlags = $flags

        # MDS (ADV190013): hardware-protected bit, with AMD/ARM treated as protected.
        $mdsProtected = (($flags -band $scfMdsHardwareProtected) -ne 0)
        if ($cpu.Manufacturer -eq "AuthenticAMD" -or $cpu.Architecture -eq 5 -or $cpu.Architecture -eq 12) {
            $mdsProtected = $true
        }
        $results.MdsVulnerable = (-not $mdsProtected)

        # SSB (ADV180012): vulnerable ONLY if the hardware requires SSBD AND it is not enabled
        # system-wide. SsbdRequired alone is the static silicon flag - it stays TRUE after the
        # mitigation is enabled, so using it by itself reports a remediated device as vulnerable.
        $ssbdRequired   = (($flags -band $scfSsbdRequired)   -ne 0)
        $ssbdSystemWide = (($flags -band $scfSsbdSystemWide) -ne 0)
        $results.SsbdSystemWide = $ssbdSystemWide
        $results.SsbdAvailable  = (($flags -band $scfSsbdAvailable) -ne 0)
        $results.SsbdSupported  = (($flags -band $scfSsbdSupported) -ne 0)
        $results.SsbVulnerable  = ($ssbdRequired -and -not $ssbdSystemWide)

        $results.Success = $true
    }
    catch {
        Write-ColorLog -Message "Exception during speculation query: $($_.Exception.Message)" -Level "Error"
    }
    finally {
        if ($ptr -ne [IntPtr]::Zero)    { [System.Runtime.InteropServices.Marshal]::FreeHGlobal($ptr) }
        if ($retPtr -ne [IntPtr]::Zero) { [System.Runtime.InteropServices.Marshal]::FreeHGlobal($retPtr) }
    }

    return $results
}

function Get-FeatureSettingsState {
    $regPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management'
    $state = @{ Override = $null; Mask = $null; Path = $regPath }
    try {
        $props = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
        if ($null -ne $props.FeatureSettingsOverride)     { $state.Override = [int]$props.FeatureSettingsOverride }
        if ($null -ne $props.FeatureSettingsOverrideMask)  { $state.Mask     = [int]$props.FeatureSettingsOverrideMask }
    }
    catch {
        Write-ColorLog -Message "Failed to read FeatureSettings registry: $($_.Exception.Message)" -Level "Warning"
    }
    return $state
}

function Get-HyperVMitigationState {
    # QID 91537 requires MinVmVersionForCpuBasedMitigations = "1.0" on Hyper-V hosts.
    # Detect the Hyper-V host role via the vmms (Virtual Machine Management) service.
    $present = $null -ne (Get-Service -Name 'vmms' -ErrorAction SilentlyContinue)
    $val = $null
    try {
        $p = Get-ItemProperty -Path $script:HyperVRegPath -Name $script:HyperVRegName -ErrorAction SilentlyContinue
        if ($null -ne $p) { $val = $p.$($script:HyperVRegName) }
    }
    catch { }
    $keyOk = ($val -eq $script:HyperVRegValue)
    return @{
        Present = $present
        Path    = $script:HyperVRegPath
        Name    = $script:HyperVRegName
        Value   = $val
        KeyOk   = $keyOk
    }
}

# -------------------------------------------------------------------------------------------------
# 3. REMEDIATION
# -------------------------------------------------------------------------------------------------

function Invoke-Remediation {
    param ($Status, $Reg, $HyperV, $IsIntel)

    Write-Host ""
    Write-ColorLog -Message "--- REMEDIATION ---" -Level "Info"

    # Target value: 72 (HT kept) or 8264 (HT disabled). Both satisfy QID 91462 (SSB)
    # AND QID 91537 (MDS), plus Mask = 3.
    $targetOverride = if ($DisableHyperThreading) { $script:TARGET_OVERRIDE_HT_DISABLED } else { $script:TARGET_OVERRIDE_NO_HT }
    $targetMask     = $script:TARGET_MASK

    $ssbOk = Test-SsbdRegCompliant -Reg $Reg
    $mdsOk = Test-MdsRegCompliant  -Reg $Reg
    $mdsRelevant = $IsIntel   # QID 91537 is Intel TSX/TAA specific
    # The Hyper-V CPU-mitigation key is only checked by QID 91537 (Intel hosts).
    $hvRelevant = $HyperV.Present -and $IsIntel
    $hvOk = (-not $hvRelevant) -or $HyperV.KeyOk

    # Report current scanner posture.
    Write-ColorLog -Message ("SSB (QID 91462): {0}" -f $(if ($ssbOk) { 'COMPLIANT' } else { 'NON-COMPLIANT' })) -Level $(if ($ssbOk) { 'Success' } else { 'Warning' })
    if ($mdsRelevant) {
        Write-ColorLog -Message ("MDS (QID 91537): {0}" -f $(if ($mdsOk) { 'COMPLIANT' } else { 'NON-COMPLIANT' })) -Level $(if ($mdsOk) { 'Success' } else { 'Warning' })
    }
    else {
        Write-ColorLog -Message "MDS (QID 91537): N/A (non-Intel CPU; Intel TSX/TAA specific)" -Level "Verbose"
    }
    if ($hvRelevant) {
        Write-ColorLog -Message ("Hyper-V key    : {0} ({1} = '{2}')" -f $(if ($HyperV.KeyOk) { 'COMPLIANT' } else { 'NON-COMPLIANT' }), $HyperV.Name, $(if ($HyperV.Value) { $HyperV.Value } else { 'NotSet' })) -Level $(if ($HyperV.KeyOk) { 'Success' } else { 'Warning' })
    }

    $mdsSatisfied     = (-not $mdsRelevant) -or $mdsOk
    $alreadyCompliant = $ssbOk -and $mdsSatisfied -and $hvOk
    if ($alreadyCompliant) {
        Write-ColorLog -Message "Registry already satisfies the applicable scanner checks. No write needed." -Level "Success"
        return
    }

    if (-not $ForceRemediation) {
        Write-ColorLog -Message ("Detection-only mode. Re-run with -ForceRemediation to set FeatureSettingsOverride = {0} (0x{0:X}), Mask = {1}." -f $targetOverride, $targetMask) -Level "Info"
        if ($hvRelevant -and -not $HyperV.KeyOk) {
            Write-ColorLog -Message ("  Will also set {0} = '1.0' (Hyper-V host)." -f $HyperV.Name) -Level "Info"
        }
        return
    }

    # ---- Compute the value to write ----
    # DEFAULT (do-no-harm / ADDITIVE): preserve every existing higher-order mitigation bit,
    # FORCE the inverted-logic disable bits (0x1 Spectre V2, 0x2 Meltdown) CLEAR so those
    # mitigations stay enabled, then SET SSBD (0x8) + MDS/TAA/L1TF override (0x40), plus
    # HT-disable (0x2000) with -DisableHyperThreading. On a clean/missing device this yields
    # EXACTLY 72 (or 8264) -> clears BOTH QIDs with no regression. On a device that already
    # carries extra bits (e.g. BHI 0x800000) it produces a SUPERSET (stronger, never weaker),
    # which still will NOT equal the exact 72/8264 that QID 91537 demands.
    #
    # -ForceExactValue: overwrite to EXACTLY 72/8264 + Mask 3 so the QID 91537 exact-match
    # clears, at the cost of DROPPING any extra mitigation bits (regression risk; opt-in).
    $curOverride = if ($null -ne $Reg.Override) { $Reg.Override } else { 0 }
    $curMask     = if ($null -ne $Reg.Mask)     { $Reg.Mask }     else { 0 }

    if ($ForceExactValue) {
        $newOverride = $targetOverride
        $newMask     = $targetMask
    }
    else {
        $bitsToSet = $script:BIT_ENABLE_SSBD -bor $script:BIT_ENABLE_MDS_L1TF
        if ($DisableHyperThreading) { $bitsToSet = $bitsToSet -bor $script:BIT_DISABLE_HT }
        $newOverride = (($curOverride -band (-bnot ($script:BIT_DISABLE_SPECTRE_V2 -bor $script:BIT_DISABLE_MELTDOWN))) -bor $bitsToSet)
        $newMask     = ($curMask -bor $script:MASK_SPECTRE_MELTDOWN)
    }

    # Bits beyond the ones we manage (i.e. mitigations set by other advisories).
    $managedBits = $script:BIT_DISABLE_SPECTRE_V2 -bor $script:BIT_DISABLE_MELTDOWN -bor $script:BIT_ENABLE_SSBD -bor $script:BIT_ENABLE_MDS_L1TF -bor $script:BIT_DISABLE_HT
    $extraBits   = $curOverride -band (-bnot $managedBits)

    if ($ForceExactValue -and $extraBits -ne 0) {
        Write-ColorLog -Message ("-ForceExactValue will OVERWRITE 0x{0:X} -> 0x{1:X}, DROPPING extra mitigation bits 0x{2:X} (e.g. BHI/MMIO). Confirm those are not required." -f $curOverride, $newOverride, $extraBits) -Level "Warning"
    }
    elseif (-not $ForceExactValue -and $mdsRelevant -and ($newOverride -ne $script:TARGET_OVERRIDE_NO_HT) -and ($newOverride -ne $script:TARGET_OVERRIDE_HT_DISABLED)) {
        Write-ColorLog -Message ("Additive result 0x{0:X} PRESERVES existing mitigations but is not the exact 72/8264 QID 91537 wants; that MDS QID may stay flagged. Use -ForceExactValue to force exactly 72/8264 (drops extra bits)." -f $newOverride) -Level "Warning"
    }

    if (($newOverride -eq $curOverride) -and ($newMask -eq $curMask) -and ($null -ne $Reg.Override)) {
        Write-ColorLog -Message ("FeatureSettingsOverride already 0x{0:X} / Mask 0x{1:X}; no override change needed." -f $curOverride, $curMask) -Level "Info"
        $skipOverrideWrite = $true
    }
    else {
        $skipOverrideWrite = $false
    }

    try {
        if (-not $skipOverrideWrite) {
            New-ItemProperty -Path $Reg.Path -Name 'FeatureSettingsOverride'     -Value $newOverride -PropertyType DWord -Force | Out-Null
            New-ItemProperty -Path $Reg.Path -Name 'FeatureSettingsOverrideMask' -Value $newMask     -PropertyType DWord -Force | Out-Null
            Write-ColorLog -Message ("Set FeatureSettingsOverride = {0} (0x{0:X}), FeatureSettingsOverrideMask = {1} (0x{1:X})." -f $newOverride, $newMask) -Level "Success"
        }

        if ($hvRelevant -and -not $HyperV.KeyOk) {
            if (-not (Test-Path $HyperV.Path)) { New-Item -Path $HyperV.Path -Force | Out-Null }
            New-ItemProperty -Path $HyperV.Path -Name $HyperV.Name -Value $script:HyperVRegValue -PropertyType String -Force | Out-Null
            Write-ColorLog -Message ("Set {0} = '{1}' (Hyper-V host)." -f $HyperV.Name, $script:HyperVRegValue) -Level "Success"
        }

        Write-ColorLog -Message "Registry written. A REBOOT is required to take effect." -Level "Success"
        if ($hvRelevant) {
            Write-ColorLog -Message "Hyper-V host: fully shut down all VMs before rebooting so the firmware mitigation applies to the host first." -Level "Warning"
        }
        if ($DisableHyperThreading) {
            Write-ColorLog -Message "NOTE: the HT-disabled value DISABLES Hyper-Threading/SMT - expect reduced multi-threaded performance." -Level "Warning"
        }
    }
    catch {
        Write-ColorLog -Message "Failed to write registry: $($_.Exception.Message)" -Level "Error"
    }
}

# -------------------------------------------------------------------------------------------------
# 4. EXECUTION ENTRY POINT
# -------------------------------------------------------------------------------------------------

$status = Get-SpeculationStatus
$reg    = Get-FeatureSettingsState
$hyperV = Get-HyperVMitigationState

$isIntel  = ($status.CpuManufacturer -eq 'GenuineIntel')
$osInfo   = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
$isServer = ($null -ne $osInfo) -and ($osInfo.ProductType -ne 1)

# ----- Detection report -----
$w = 64
$bar = '=' * $w
Write-Host ""
Write-Host $bar -ForegroundColor DarkCyan
Write-Host "  Hardware Speculation - Detection Report" -ForegroundColor Cyan
Write-Host $bar -ForegroundColor DarkCyan

if ($status.Success) {
    $rawHex = if ($null -ne $status.RawFlags) { '0x{0:X}' -f $status.RawFlags } else { 'N/A' }
    Write-ColorLog -Message ("CPU              : {0}" -f $status.CpuManufacturer) -Level "Verbose"
    Write-ColorLog -Message ("OS type          : {0}" -f $(if ($isServer) { 'Server' } else { 'Client' })) -Level "Verbose"
    Write-ColorLog -Message ("SpeculationFlags : {0}" -f $rawHex) -Level "Verbose"

    $mdsClr = if ($status.MdsVulnerable) { 'Warning' } else { 'Success' }
    $ssbClr = if ($status.SsbVulnerable) { 'Warning' } else { 'Success' }
    Write-ColorLog -Message ("MDS (ADV190013)  : {0}" -f $(if ($status.MdsVulnerable) { 'VULNERABLE' } else { 'Protected' })) -Level $mdsClr
    Write-ColorLog -Message ("SSB (ADV180012)  : {0}" -f $(if ($status.SsbVulnerable) { 'VULNERABLE (SSBD required, not active)' } else { 'Protected' })) -Level $ssbClr
    Write-ColorLog -Message ("  SSBD runtime    : SystemWide={0} Available(OS)={1} Supported(microcode)={2}" -f $status.SsbdSystemWide, $status.SsbdAvailable, $status.SsbdSupported) -Level "Verbose"
    if ($status.SsbVulnerable) {
        # Explain WHY it is still vulnerable so "ran remediation + rebooted, still flagged" is actionable.
        if (-not $status.SsbdSupported) {
            Write-ColorLog -Message "  SSB reason      : CPU microcode does NOT support SSBD - a firmware/BIOS update is required (ADV180002). Registry alone cannot activate it." -Level "Warning"
        } elseif (-not $status.SsbdAvailable) {
            Write-ColorLog -Message "  SSB reason      : OS support for SSBD absent - install the latest cumulative update, then reboot." -Level "Warning"
        } else {
            Write-ColorLog -Message "  SSB reason      : microcode + OS support present but SSBD not active - set FeatureSettingsOverride and REBOOT (KB4073119)." -Level "Warning"
        }
    }

}
else {
    Write-ColorLog -Message "Speculation query FAILED - state unknown." -Level "Error"
}

Write-Host ""
Write-Host "  [Registry] Session Manager\Memory Management" -ForegroundColor DarkCyan
foreach ($line in (Get-FeatureSettingsDecoding -Override $reg.Override -Mask $reg.Mask)) {
    Write-Host "    $line" -ForegroundColor DarkGray
}
$ssbReg = Test-SsbdRegCompliant -Reg $reg
$mdsReg = Test-MdsRegCompliant  -Reg $reg
Write-Host ("    SSB config compliant  (QID 91462): {0}" -f $ssbReg) -ForegroundColor $(if ($ssbReg) { 'Green' } else { 'Yellow' })
if ($isIntel) {
    Write-Host ("    MDS config compliant  (QID 91537): {0}" -f $mdsReg) -ForegroundColor $(if ($mdsReg) { 'Green' } else { 'Yellow' })
}
else {
    Write-Host  "    MDS config compliant  (QID 91537): N/A (non-Intel)" -ForegroundColor DarkGray
}
if ($hyperV.Present) {
    if ($isIntel) {
        Write-Host ("    Hyper-V mitigation key (QID 91537): {0} (value='{1}')" -f $hyperV.KeyOk, $(if ($hyperV.Value) { $hyperV.Value } else { 'NotSet' })) -ForegroundColor $(if ($hyperV.KeyOk) { 'Green' } else { 'Yellow' })
    }
    else {
        Write-Host  "    Hyper-V mitigation key (QID 91537): N/A (non-Intel)" -ForegroundColor DarkGray
    }
}

Write-Host ""
Write-Host $bar -ForegroundColor DarkCyan

# ----- Ivanti-style summary -----
# Scanner compliance is a REGISTRY-CONFIG check. SSB (QID 91462) applies to all CPUs;
# MDS (QID 91537) is Intel-specific (servers in particular). Hyper-V hosts also need
# the MinVmVersion key.
$mdsRelevant  = $isIntel
$ssbCompliant = $ssbReg
$mdsCompliant = (-not $mdsRelevant) -or $mdsReg
$hvRelevant   = $hyperV.Present -and $isIntel
$hvCompliant  = (-not $hvRelevant) -or $hyperV.KeyOk
$overallOk    = $ssbCompliant -and $mdsCompliant -and $hvCompliant

$expectedStr = "FeatureSettingsOverride in {72,8264}; Mask=3" + $(if ($hvRelevant) { "; MinVmVersion=1.0" } else { "" })
$ovHex = if ($null -ne $reg.Override) { '0x{0:X}' -f $reg.Override } else { 'missing' }
$mkHex = if ($null -ne $reg.Mask)     { '0x{0:X}' -f $reg.Mask }     else { 'missing' }
$foundStr = "Override=$ovHex; Mask=$mkHex; SSBok=$ssbReg; MDSok=$mdsReg; HVok=$hvCompliant"

Write-Host ""
Write-ColorLog -Message "--- DETECTION SUMMARY ---" -Level "Info"
if ($overallOk) {
    Write-ColorLog -Message "detected = false (Compliant)" -Level "Success"
}
else {
    Write-ColorLog -Message "detected = true (Non-Compliant)" -Level "Warning"
}
Write-ColorLog -Message "expected = $expectedStr" -Level "Verbose"
Write-ColorLog -Message "found    = $foundStr" -Level "Verbose"

# ----- Remediation -----
Invoke-Remediation -Status $status -Reg $reg -HyperV $hyperV -IsIntel $isIntel
