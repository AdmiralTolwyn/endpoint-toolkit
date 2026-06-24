<#
.SYNOPSIS
    Ivanti Custom Vulnerability Detection for MDS and SSB Hardware.
.DESCRIPTION
    Detects CPU hardware vulnerabilities based on user-defined switches.
    Uses ntdll.dll to query system speculation control information.

    This script is based on the official Microsoft SpeculationControl PowerShell module.
    
    DISCLAIMER:
    THIS SCRIPT IS PROVIDED "AS-IS" WITHOUT WARRANTY OF ANY KIND. 
    USE AT YOUR OWN RISK. THE AUTHOR AND CONTRIBUTORS ARE NOT RESPONSIBLE 
    FOR ANY DAMAGES OR DATA LOSS RESULTING FROM THE USE OF THIS SCRIPT.
    NOTE: ONLY LIMITED TESTING HAS BEEN PERFORMED ON THIS VERSION.
.NOTES
    Author:         Anton Romanyuk
    Version:        1.0
#>

function Get-HardwareSpeculationStatus {
    [CmdletBinding()]
    param (
        [switch]$MdsOnly,
        [switch]$SsbOnly,
        [switch]$Both
    )

    Write-Host " [DEBUG] --- Entering Get-HardwareSpeculationStatus ---" -ForegroundColor Cyan

    # Determine Scan Scope
    $CheckMds = $true
    $CheckSsb = $true
    
    if ($MdsOnly) { 
        $CheckSsb = $false
        Write-Host " [DEBUG] Scope: MdsOnly" -ForegroundColor Yellow 
    }
    elseif ($SsbOnly) { 
        $CheckMds = $false
        Write-Host " [DEBUG] Scope: SsbOnly" -ForegroundColor Yellow 
    }
    else { 
        Write-Host " [DEBUG] Scope: Both/Default" -ForegroundColor Yellow 
    }

    # Define flags from speculation control information.
    # NOTE (ADV180002 / KB4073119): SsbdRequired is a STATIC hardware-vulnerability
    # flag - it stays TRUE even after SSBD is fully mitigated. The flag that flips
    # once the mitigation is actually active (registry + reboot + CPU microcode) is
    # SsbdSystemWide. Detecting "still vulnerable" on SsbdRequired alone is wrong and
    # makes a remediated, rebooted device report vulnerable forever.
    $scfMdsHardwareProtected = 0x1000000
    $scfSsbdAvailable        = 0x100     # OS support for SSBD present
    $scfSsbdSupported        = 0x200     # CPU microcode supports SSBD
    $scfSsbdSystemWide       = 0x400     # SSBD ENABLED system-wide (flips after reg+reboot+microcode)
    $scfSsbdRequired         = 0x1000    # hardware is VULNERABLE / requires SSBD (STATIC - never clears)

    Write-Host " [DEBUG] Defining native ntdll methods..." -ForegroundColor Yellow
    
    # Using a Here-String for cleaner PS 5.1 parsing
    $NtQSIDefinition = @'
    [DllImport("ntdll.dll")] 
    public static extern int NtQuerySystemInformation(uint systemInformationClass, IntPtr systemInformation, uint systemInformationLength, IntPtr returnLength);
'@
    
    try {
        $ntdll = Add-Type -MemberDefinition $NtQSIDefinition -Name 'ntdll' -Namespace 'Win32' -PassThru
        Write-Host " [DEBUG] ntdll type added successfully." -ForegroundColor Green
    } catch {
        Write-Host " [DEBUG] FAILED to add ntdll type: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }

    $SYSTEM_SPECULATION_CONTROL_INFORMATION_LENGTH = 8
    [IntPtr]$ptr = [System.Runtime.InteropServices.Marshal]::AllocHGlobal($SYSTEM_SPECULATION_CONTROL_INFORMATION_LENGTH)
    [IntPtr]$retPtr = [System.Runtime.InteropServices.Marshal]::AllocHGlobal(4)

    # Initialize results hashtable
    $results = @{
        MdsVulnerable = $false
        SsbVulnerable = $false
        Success       = $false
        ScanScope     = @()
    }

    try {
        Write-Host " [DEBUG] Querying CPU Metadata via CIM..." -ForegroundColor Yellow
        $cpu = Get-CimInstance Win32_Processor
        Write-Host " [DEBUG] CPU Manufacturer: $($cpu.Manufacturer)" -ForegroundColor Green
        
        Write-Host " [DEBUG] Invoking NtQuerySystemInformation (Class 201)..." -ForegroundColor Yellow
        $retval = $ntdll::NtQuerySystemInformation(201, $ptr, $SYSTEM_SPECULATION_CONTROL_INFORMATION_LENGTH, $retPtr)
        
        # Check return value (0 = STATUS_SUCCESS)
        if ($retval -ne 0) {
            Write-Host " [DEBUG] Native query failed. RetVal: $retval" -ForegroundColor Red
            return $results
        }
        Write-Host " [DEBUG] NtQuerySystemInformation success." -ForegroundColor Green

        $flags = [uint32][System.Runtime.InteropServices.Marshal]::ReadInt32($ptr)
        Write-Host " [DEBUG] Raw Speculation Flags retrieved: 0x$($flags.ToString('X'))" -ForegroundColor Green

        # --- MDS Logic (ADV190013) ---
        if ($CheckMds) {
            $results.ScanScope += "MDS"
            $mdsProtected = (($flags -band $scfMdsHardwareProtected) -ne 0)
            Write-Host " [DEBUG] MDS Hardware Protected (Bitmask Check): $mdsProtected" -ForegroundColor Yellow
            
            # AMD and ARM architectures are considered protected
            if ($cpu.Manufacturer -eq "AuthenticAMD" -or $cpu.Architecture -eq 5 -or $cpu.Architecture -eq 12) {
                Write-Host " [DEBUG] MDS Protection inherited via CPU Vendor/Arch." -ForegroundColor Green
                $mdsProtected = $true
            }
            $results.MdsVulnerable = ($mdsProtected -eq $false)
        }

        # --- SSB Logic (ADV180012) ---
        if ($CheckSsb) {
            $results.ScanScope += "SSB"
            $ssbdRequired   = (($flags -band $scfSsbdRequired)   -ne 0)
            $ssbdSystemWide = (($flags -band $scfSsbdSystemWide) -ne 0)
            $ssbdAvailable  = (($flags -band $scfSsbdAvailable)  -ne 0)
            $ssbdSupported  = (($flags -band $scfSsbdSupported)  -ne 0)
            # Vulnerable ONLY if the hardware requires SSBD AND it is not enabled system-wide.
            # SsbdRequired alone is the static silicon flag and must NOT, by itself, mean vulnerable.
            $results.SsbVulnerable = ($ssbdRequired -and -not $ssbdSystemWide)
            $results.SsbdSystemWide = $ssbdSystemWide
            $results.SsbdAvailable  = $ssbdAvailable
            $results.SsbdSupported  = $ssbdSupported
            Write-Host " [DEBUG] SSB: Required=$ssbdRequired SystemWide=$ssbdSystemWide Available=$ssbdAvailable Supported=$ssbdSupported -> Vulnerable=$($results.SsbVulnerable)" -ForegroundColor Yellow
            if ($ssbdRequired -and -not $ssbdSystemWide) {
                if (-not $ssbdSupported) {
                    Write-Host " [DEBUG] SSB: SSBD not enabled - CPU microcode does not support SSBD. A firmware/BIOS update is required (ADV180002)." -ForegroundColor Red
                } elseif (-not $ssbdAvailable) {
                    Write-Host " [DEBUG] SSB: SSBD not enabled - OS support absent. Install the latest cumulative update." -ForegroundColor Red
                } else {
                    Write-Host " [DEBUG] SSB: microcode + OS support present but SSBD not active - set FeatureSettingsOverride and REBOOT (KB4073119)." -ForegroundColor Red
                }
            }
        }

        $results.Success = $true
    }
    catch {
        Write-Host " [DEBUG] EXCEPTION in try block: $($_.Exception.Message)" -ForegroundColor Red
    }
    finally {
        Write-Host " [DEBUG] Cleaning up memory pointers..." -ForegroundColor Yellow
        if ($ptr -ne [IntPtr]::Zero) { [System.Runtime.InteropServices.Marshal]::FreeHGlobal($ptr) }
        if ($retPtr -ne [IntPtr]::Zero) { [System.Runtime.InteropServices.Marshal]::FreeHGlobal($retPtr) }
    }

    Write-Host " [DEBUG] Success status: $($results.Success)" -ForegroundColor Green
    return $results
}

# --- Execution ---

# Default Execution - Change to -MdsOnly or -SsbOnly as needed
$Status = Get-HardwareSpeculationStatus -Both

$ExpectedString = "MdsProtected=True;SsbMitigated=True (SSBD not required OR enabled system-wide)"
$FoundString = "MdsVulnerable=$($Status.MdsVulnerable);SsbVulnerable=$($Status.SsbVulnerable);SsbdSystemWide=$($Status.SsbdSystemWide)"

if ($Status.Success -eq $false) {
    Write-Host "detected = false" -ForegroundColor Gray
    Write-Host "reason = Script failed to query ntdll system information." -ForegroundColor Red
}
elseif ($Status.MdsVulnerable -eq $true -or $Status.SsbVulnerable -eq $true) {
    # --- DETECTED (VULNERABLE) ---
    Write-Host "detected = true" -ForegroundColor Red
    
    if ($Status.MdsVulnerable -and $Status.SsbVulnerable) {
        Write-Host "reason = Hardware is vulnerable to both MDS (ADV190013) and SSB (ADV180012)." -ForegroundColor Yellow
    }
    elseif ($Status.MdsVulnerable) {
        Write-Host "reason = Hardware is vulnerable to MDS (ADV190013)." -ForegroundColor Yellow
    }
    else {
        Write-Host "reason = Hardware is vulnerable to SSB (ADV180012)." -ForegroundColor Yellow
    }
    
    Write-Host "expected = $ExpectedString" -ForegroundColor Cyan
    Write-Host "found = $FoundString" -ForegroundColor Magenta
}
else {
    # --- COMPLIANT ---
    Write-Host "detected = false" -ForegroundColor Green
    Write-Host "reason = Hardware is protected against targeted vulnerabilities ($($Status.ScanScope -join ', '))." -ForegroundColor Gray
    Write-Host "expected = $ExpectedString" -ForegroundColor Cyan
    Write-Host "found = $FoundString" -ForegroundColor Magenta
}
