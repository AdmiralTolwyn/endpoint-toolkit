<#
.SYNOPSIS
    Activates Hybrid Azure AD Joined AVD session hosts across all host pools.

.DESCRIPTION
    Scans the entire subscription for VMs tagged with 'HybridStatus = Pending'.
    For each pending VM:
      1. Resolves the parent host pool from the 'cm-resource-parent' tag
      2. Verifies the session host is registered and currently drained
      3. Triggers the Automatic-Device-Join scheduled task via Invoke-AzVMRunCommand
      4. Validates hybrid join status (DomainJoined + AzureAdJoined) via dsregcmd
      5. On success: undrains the session host and updates the tag to 'Active'

    Designed to run on a 15-minute schedule. No parameters required — discovery
    is fully tag-driven, so one pipeline covers all host pools.

.NOTES
    File:     scripts/Invoke-HybridActivator.ps1
    Pipeline: avd-activator.yml (scheduled every 15 min)
    Requires: Az.Compute, Az.DesktopVirtualization modules
    Tags:     HybridStatus (Pending → Active), cm-resource-parent (host pool link)

.LINK
    https://learn.microsoft.com/en-us/entra/identity/devices/hybrid-join-plan
#>

$ErrorActionPreference = "Stop"

Write-Host " [INIT] Scanning subscription for VMs with HybridStatus=Pending..." -ForegroundColor Cyan

# ─────────────────────────────────────────────────────────────────────────────
# 1. DISCOVER: Find all VMs tagged HybridStatus=Pending (subscription-wide)
# ─────────────────────────────────────────────────────────────────────────────
$PendingVms = Get-AzVM -Status | Where-Object { $_.Tags['HybridStatus'] -eq 'Pending' }

if ($PendingVms.Count -eq 0) {
    Write-Host " [OK] No pending VMs found. Nothing to do." -ForegroundColor Green
    return
}

Write-Host " [FOUND] $($PendingVms.Count) VM(s) with HybridStatus=Pending" -ForegroundColor Yellow

foreach ($Vm in $PendingVms) {
    $VmName = $Vm.Name
    $ComputeRg = $Vm.ResourceGroupName

    # ─────────────────────────────────────────────────────────────────────────
    # 2. RESOLVE HOST POOL from 'cm-resource-parent' tag
    #    The Bicep template stamps this tag on every session host VM at deploy
    #    time, pointing back to the host pool resource ID.
    # ─────────────────────────────────────────────────────────────────────────
    $HostPoolResourceId = $Vm.Tags['cm-resource-parent']
    if ([string]::IsNullOrEmpty($HostPoolResourceId)) {
        Write-Host " [SKIP] $VmName - missing 'cm-resource-parent' tag" -ForegroundColor DarkYellow
        continue
    }

    # Parse resource ID: /subscriptions/.../resourceGroups/<rg>/providers/Microsoft.DesktopVirtualization/hostpools/<name>
    $IdParts = $HostPoolResourceId.Split('/')
    $HostPoolRg   = $IdParts[4]   # Resource group at index 4
    $HostPoolName = $IdParts[-1]  # Host pool name is the last segment

    # ─────────────────────────────────────────────────────────────────────────
    # 3. FIND SESSION HOST ENTRY
    #    AVD session host names are FQDN-based (e.g. "VMNAME.domain.com").
    #    We match by splitting on '.' and comparing the short name.
    # ─────────────────────────────────────────────────────────────────────────
    $SessionHost = $null
    $SessionHostLeaf = $null
    $AllSH = Get-AzWvdSessionHost -ResourceGroupName $HostPoolRg -HostPoolName $HostPoolName -ErrorAction SilentlyContinue
    foreach ($SH in $AllSH) {
        # Session host Name format: "HostPoolName/vmname.domain.com"
        $Leaf = $SH.Name.Split('/')[-1]
        if ($Leaf.Split('.')[0] -eq $VmName) {
            $SessionHost = $SH
            $SessionHostLeaf = $Leaf
            break
        }
    }

    if ($null -eq $SessionHost) {
        Write-Host " [SKIP] $VmName - not registered in host pool '$HostPoolName'" -ForegroundColor DarkYellow
        continue
    }

    # Skip hosts that are already accepting sessions (not drained = already active)
    if ($SessionHost.AllowNewSession -eq $true) { continue }

    Write-Host "-------------------------------------------------------------"
    Write-Host " ACTIVATING: $VmName" -ForegroundColor Magenta
    Write-Host "   Host Pool: $HostPoolName ($HostPoolRg)"
    Write-Host "   Status: Drained & Pending Validation"

    # ─────────────────────────────────────────────────────────────────────────
    # 4. RUN COMMAND: TRIGGER HYBRID JOIN & CHECK STATUS
    #    Combines the join trigger + status check in a single remote call
    #    to minimize Invoke-AzVMRunCommand overhead.
    # ─────────────────────────────────────────────────────────────────────────
    $ScriptBlock = {
        # A. Trigger the built-in Workplace Join scheduled task
        Write-Output "Triggering Automatic-Device-Join task..."
        Get-ScheduledTask -TaskPath "\Microsoft\Windows\Workplace Join\" -TaskName "Automatic-Device-Join" | Start-ScheduledTask

        # B. Wait for the task to complete registration
        Start-Sleep -Seconds 15

        # C. Capture hybrid join status
        cmd /c "dsregcmd /status"
    }

    try {
        $Result = Invoke-AzVMRunCommand -ResourceGroupName $ComputeRg `
                                        -VMName $VmName `
                                        -CommandId 'RunPowerShellScript' `
                                        -ScriptBlock $ScriptBlock `
                                        -ErrorAction Stop

        $Output = $Result.Value[0].Message

        # ─────────────────────────────────────────────────────────────────────
        # 5. PARSE dsregcmd OUTPUT
        #    Both flags must be YES for a valid Hybrid Azure AD Join:
        #      - DomainJoined: YES  (on-prem AD)
        #      - AzureAdJoined: YES (Entra ID)
        # ─────────────────────────────────────────────────────────────────────
        $IsAzJoined = $Output -match "AzureAdJoined : YES"
        $IsDomainJoined = $Output -match "DomainJoined : YES"

        if ($IsAzJoined -and $IsDomainJoined) {
            Write-Host "   [PASS] Hybrid Join Confirmed." -ForegroundColor Green

            # A. UNBLOCK: Re-enable user sessions (undrain)
            Write-Host "   -> Enabling User Sessions..." -NoNewline
            Update-AzWvdSessionHost -ResourceGroupName $HostPoolRg `
                                    -HostPoolName $HostPoolName `
                                    -Name $SessionHostLeaf `
                                    -AllowNewSession:$true
            Write-Host " Done." -ForegroundColor Green

            # B. UPDATE TAG: Mark as Active so future scans skip this VM
            Write-Host "   -> Updating 'HybridStatus' tag..." -NoNewline
            $Tags = $Vm.Tags
            $Tags['HybridStatus'] = 'Active'
            Update-AzTag -ResourceId $Vm.Id -Tag $Tags -Operation Merge | Out-Null
            Write-Host " Done." -ForegroundColor Green

        } else {
            # Join not yet complete — will retry on the next scheduled cycle
            Write-Host "   [WAIT] Sync incomplete. (AD: $IsDomainJoined | Entra: $IsAzJoined)" -ForegroundColor Yellow
            Write-Host "          Retrying in next cycle."
        }
    } catch {
        # VM agent may not be ready yet, or the Run Command timed out
        Write-Host "   [ERR] VM Agent not reachable or Script failed." -ForegroundColor Red
    }
}