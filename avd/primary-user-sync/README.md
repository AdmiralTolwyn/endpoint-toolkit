# AVD Personal Desktop to Intune Primary User Sync
[Sync-AvdPersonalPrimaryUser.ps1](Sync-AvdPersonalPrimaryUser.ps1) populates the missing Microsoft Intune **Primary User** (`users/$ref`) on Entra ID-joined Azure Virtual Desktop (AVD) Personal session hosts based on each session host's `properties.assignedUser` in Azure Resource Manager (ARM).
---
## 1. Problem & Architecture
When AVD session hosts are deployed with `AADLoginForWindows` and automatic Intune enrollment (`mdmId = 0000000a-0000-0000-c000-000000000000`, as in [sessionHost-entraid.bicep](../bicep/modules/sessionHost-entraid.bicep) and [Microsoft Entra joined session hosts in Azure Virtual Desktop](https://learn.microsoft.com/en-us/azure/virtual-desktop/azure-ad-joined-session-hosts)), Windows enrolls into Intune under the **SYSTEM** account (`deviceenroller.exe /c /AutoEnrollMDM`). Because enrollment is device-driven rather than user-driven:
1. Intune sets **Primary User = None** (often represented in Graph inventory as an empty or all-zero `userId` `00000000-0000-0000-0000-000000000000`).
2. Per [Change a device's primary user in Microsoft Intune](https://learn.microsoft.com/en-us/intune/device-management/inventory-and-status/find-primary-user#company-portal-app-and-the-primary-user), when no Primary User is assigned, Company Portal detects the VM as a **Shared Device** (available apps can still be requested/installed in Company Portal 10.3.4651.0+, but uninstalling apps and self-service device actions such as reset, rename, and retire are unavailable, and user-to-device affinity in Intune troubleshooting/reporting is missing).
3. AVD does **not** natively synchronize `sessionHost.properties.assignedUser` (populated via [Personal desktop Direct or Automatic assignment](https://learn.microsoft.com/en-us/azure/virtual-desktop/configure-host-pool-personal-desktop-assignment-type)) to Intune.
### Execution Flow (Fail-Closed v1 Core)
```text
1. ARM GET   /subscriptions/{sub}/resourceGroups/{rg}/providers/Microsoft.DesktopVirtualization/hostPools/{pool}
             └──► Verifies hostPoolType == 'Personal' and lists sessionHosts (name + properties.assignedUser)
2. Graph GET /v1.0/deviceManagement/managedDevices?$filter=operatingSystem eq 'Windows' and (deviceName eq 'HOST1' ...)
             └──► Resolves Intune managedDevice record in batches of <=15 hostnames
             └──► Fails closed on duplicate Intune records (AmbiguousIntuneDevice) or stale records enrolled >24h before sessionHost creation (StaleIntuneDeviceRecord)
3. Graph GET /v1.0/users/{assignedUser}?$select=id,userPrincipalName
             └──► Resolves the assigned user's immutable Entra ObjectId (with mail fallback for B2B guests)
4. Graph GET /beta/deviceManagement/managedDevices/{intuneDeviceId}/users?$select=id,userPrincipalName
             └──► Inspects the authoritative Primary User relationship:
                  • Matching ObjectId  ──► AlreadyInSync (no write)
                  • Different ObjectId ──► PrimaryUserConflict (fail-closed; never overwrites)
                  • Empty              ──► Proceeds to step 5
5. Graph POST /beta/deviceManagement/managedDevices/{intuneDeviceId}/users/$ref
              Body: { "@odata.id": "https://graph.microsoft.com/beta/users/{userObjectId}" }
   (Optional) POST /v1.0/deviceManagement/managedDevices/{intuneDeviceId}/syncDevice (-SyncIntuneDevice)
```
---
## 2. Permissions & Setup
### Required Permissions
| Plane | Role / Permission | Type | Scope | Purpose |
| --- | --- | --- | --- | --- |
| **Azure RBAC** | [`Desktop Virtualization Reader`](https://learn.microsoft.com/en-us/azure/virtual-desktop/rbac#desktop-virtualization-reader) or [`Desktop Virtualization Host Pool Reader`](https://learn.microsoft.com/en-us/azure/virtual-desktop/rbac#desktop-virtualization-host-pool-reader) (or standard `Reader`) | RBAC Role | Host Pool Resource Group (or Host Pool resource) | Read Host Pool type (`hostpools/read`) and session host `assignedUser` (`hostpools/*/read`) |
| **Microsoft Graph** | [`DeviceManagementManagedDevices.ReadWrite.All`](https://learn.microsoft.com/en-us/graph/permissions-reference#devicemanagementmanageddevicesreadwriteall) | Application (or Delegated) | Tenant | Read Intune [`managedDevices`](https://learn.microsoft.com/en-us/graph/api/resources/intune-devices-manageddevice?view=graph-rest-1.0) and `/users`, set Primary User (`POST .../users/$ref`) |
| **Microsoft Graph** | [`User.Read.All`](https://learn.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0#permissions) | Application (or Delegated) | Tenant | Resolve `assignedUser` UPN/mail to Entra User `id` |
| **Microsoft Graph** *(Optional)* | [`DeviceManagementManagedDevices.PrivilegedOperations.All`](https://learn.microsoft.com/en-us/graph/api/intune-devices-manageddevice-syncdevice?view=graph-rest-1.0#permissions) | Application (or Delegated) | Tenant | **Optional** — only required when `-SyncIntuneDevice` is passed to trigger [`POST .../syncDevice`](https://learn.microsoft.com/en-us/graph/api/intune-devices-manageddevice-syncdevice?view=graph-rest-1.0) |
> **Prerequisites & Security Notes:**
> - **Target user licensing:** Per [Change a device's primary user in Microsoft Intune](https://learn.microsoft.com/en-us/intune/device-management/inventory-and-status/find-primary-user), the target Primary User must hold an active **Intune license**, and primary-user updates can take up to 10 minutes to reflect in the Intune portal after being set.
> - **Interactive / Delegated Intune RBAC:** When running interactively with delegated Graph permissions, the signed-in administrator must also hold an Intune RBAC role granting **Managed devices > Read** and **Managed devices > Set primary user** (and **Sync** if using `-SyncIntuneDevice`), such as **Intune Administrator**.
> - **Graph Application permissions are tenant-wide** and do not enforce Intune RBAC Scope Tags. An automation identity granted `DeviceManagementManagedDevices.ReadWrite.All` can modify Primary User on any Intune device in the tenant.
> - **Graph `/beta` Primary User endpoint:** Reading (`GET .../users`) and setting (`POST /beta/deviceManagement/managedDevices/{id}/users/$ref`, per [microsoft/Intune.Graph `Set-PrimaryUser.ps1`](https://github.com/microsoft/Intune.Graph/blob/main/Source/ManagedDevices/Set-PrimaryUser.ps1)) use the Microsoft Graph `/beta` endpoint.
> - Because `DeviceManagementManagedDevices.PrivilegedOperations.All` also permits remote device actions in Graph, `-SyncIntuneDevice` is **off by default** so you do not need to grant `PrivilegedOperations.All` unless desired. You can alternatively trigger local MDM check-in from inside the VM using [Invoke-IntuneClientSync.ps1](../../intune/mdm-sync-service/Invoke-IntuneClientSync.ps1).
> - The script is **fail-closed**: it only populates *missing* Primary Users and never overwrites an existing Primary User (`PrimaryUserConflict`) or clears unassigned hosts.
### Granting Graph App Roles to a Managed Identity (One-Time Setup)
Azure Portal does not provide a GUI to assign Microsoft Graph Application permissions to a Managed Identity. Run the following as a **Global Administrator** or **Privileged Role Administrator**:
```powershell
Connect-MgGraph -Scopes 'AppRoleAssignment.ReadWrite.All', 'Application.Read.All'
# Object (Principal) ID of the System-Assigned or User-Assigned Managed Identity
$miPrincipalId = '<Managed-Identity-Object-Id>'
$graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"
# Include 'DeviceManagementManagedDevices.PrivilegedOperations.All' only if using -SyncIntuneDevice
$requiredRoles = @(
    'DeviceManagementManagedDevices.ReadWrite.All',
    'User.Read.All'
)
foreach ($roleName in $requiredRoles) {
    $appRole = $graphSp.AppRoles | Where-Object {
        $_.Value -eq $roleName -and $_.AllowedMemberTypes -contains 'Application'
    }
    New-MgServicePrincipalAppRoleAssignment `
        -ServicePrincipalId $miPrincipalId `
        -PrincipalId $miPrincipalId `
        -ResourceId $graphSp.Id `
        -AppRoleId $appRole.Id
}
```
---
## 3. Running Manually (Engineer Workstation)
### Workstation Authentication Note (`Connect-AzAccount` vs. Microsoft Graph)
When you sign in interactively as a user with `Connect-AzAccount`, calling `Get-AzAccessToken -ResourceUrl 'https://graph.microsoft.com/'` mints a delegated token for the **Azure PowerShell** client application (`1950a258-227b-4e31-a9cf-717495945fc2`). In most tenants, that client application does **not** carry delegated `DeviceManagementManagedDevices.ReadWrite.All` scopes, so Graph returns `403 Forbidden`.
Use one of two approaches when running manually:
#### Option A: Interactive User (ARM via `Az.Accounts` + Graph Token via `Microsoft.Graph`)
```powershell
# 1. Sign in to Azure for ARM read access
Connect-AzAccount -Subscription '<Subscription-Id>'
# 2. Sign in to Microsoft Graph with Intune delegated scopes and extract the Bearer token
# Note: Extracting RequestMessage.Headers.Authorization.Parameter from Invoke-MgGraphRequest -OutputType HttpResponseMessage
# is an SDK workaround; the extracted token is valid for ~1 hour and is not auto-refreshed.
Connect-MgGraph -Scopes 'DeviceManagementManagedDevices.ReadWrite.All', 'User.Read.All'
$mgContextToken = (Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization?$select=id' -OutputType HttpResponseMessage).RequestMessage.Headers.Authorization.Parameter
# 3. Dry-run audit (-WhatIf) against a Personal Host Pool
.\avd\primary-user-sync\Sync-AvdPersonalPrimaryUser.ps1 `
    -ResourceGroupName 'rg-avd-sim-ams-core' `
    -HostPoolName 'vdpool-pers-01' `
    -GraphAccessToken $mgContextToken `
    -WhatIf | Format-Table ShortHostName, AssignedUser, PreviousPrimaryUser, NewPrimaryUser, Status, Detail -AutoSize
# 4. Apply to a single session host first
.\avd\primary-user-sync\Sync-AvdPersonalPrimaryUser.ps1 `
    -ResourceGroupName 'rg-avd-sim-ams-core' `
    -HostPoolName 'vdpool-pers-01' `
    -TargetSessionHostName 'avd-sh-001' `
    -GraphAccessToken $mgContextToken
# 5. Apply to all session hosts in the Personal Host Pool
.\avd\primary-user-sync\Sync-AvdPersonalPrimaryUser.ps1 `
    -ResourceGroupName 'rg-avd-sim-ams-core' `
    -HostPoolName 'vdpool-pers-01' `
    -GraphAccessToken $mgContextToken | Format-Table ShortHostName, AssignedUser, NewPrimaryUser, Status -AutoSize
```
#### Option B: Service Principal Session (`Connect-AzAccount -ServicePrincipal`)
When authenticated as a Service Principal that has been granted the Graph Application permissions above, `Get-AzAccessToken` returns an app-only token with `.default` roles—no `-GraphAccessToken` is needed:
```powershell
Connect-AzAccount -ServicePrincipal -Tenant '<Tenant-Id>' -ApplicationId '<App-Id>' -CertificateThumbprint '<Thumbprint>'
.\avd\primary-user-sync\Sync-AvdPersonalPrimaryUser.ps1 `
    -SubscriptionId '<Subscription-Id>' `
    -ResourceGroupName 'rg-avd-sim-ams-core' `
    -HostPoolName 'vdpool-pers-01' `
    -WhatIf
```
---
## 4. Running via Azure DevOps Pipeline or Azure Automation
### Option 1: Azure DevOps Pipeline Task (`AzurePowerShell@5`)
Grant the pipeline's Workload Identity Federation / Service Connection Service Principal `Desktop Virtualization Reader` (ARM) plus `DeviceManagementManagedDevices.ReadWrite.All` and `User.Read.All` (Graph Application permissions):
```yaml
- task: AzurePowerShell@5
  displayName: 'Sync AVD Personal AssignedUser to Intune Primary User'
  inputs:
    azureSubscription: $(serviceConnection)
    ScriptType: 'FilePath'
    ScriptPath: '$(System.DefaultWorkingDirectory)/avd/primary-user-sync/Sync-AvdPersonalPrimaryUser.ps1'
    ScriptArguments: >-
      -SubscriptionId "$(subscriptionId)"
      -ResourceGroupName "$(hostPoolRg)"
      -HostPoolName "$(hostPoolName)"
    azurePowerShellVersion: 'LatestVersion'
```
> **Timing note for new VM deployments:** Because `AADLoginForWindows` MDM auto-enrollment takes ~3–10 minutes after first boot to appear in Intune `managedDevices` (and Automatic user assignment happens later when the user first signs in), a post-deployment pipeline step should be paired with a recurring scheduled runbook.
### Option 2: Azure Automation Account (Managed Identity)
1. Enable **System-Assigned Managed Identity** (or attach a **User-Assigned Managed Identity**) on the Automation Account.
2. Assign `Desktop Virtualization Reader` on the Host Pool Resource Group(s) and grant the Graph Application permissions using the script in Section 2.
3. Import [Sync-AvdPersonalPrimaryUser.ps1](Sync-AvdPersonalPrimaryUser.ps1) as a **PowerShell 5.1** (or **7.4**, see [Azure Automation runbook types](https://learn.microsoft.com/en-us/azure/automation/automation-runbook-types)) Runbook. *(If your Automation Account rejects `#Requires` directives, remove `#Requires -Version 5.1` when importing.)*
4. Link to a schedule passing `HostPoolResourceId`:
   - `HostPoolResourceId`: `["/subscriptions/<subId>/resourceGroups/<rg>/providers/Microsoft.DesktopVirtualization/hostPools/<poolName>"]`
   - `AuthMode`: `ManagedIdentity`
   - *(Note: Per [Schedules in Azure Automation](https://learn.microsoft.com/en-us/azure/automation/shared-resources/schedules), recurring schedules have a 1-hour minimum interval. To run every 15 minutes, link 4 hourly schedules staggered at `:00`, `:15`, `:30`, and `:45`.)*
---
## 5. Result Status Reference
Every evaluated session host produces a structured `PSCustomObject` with one of the following `Status` values:
| Status | Mutates Intune? | Meaning |
| --- | --- | --- |
| `Updated` | **Yes** (`POST .../users/$ref`) | Intune Primary User was unset and has now been populated with `AssignedUser`. |
| `PrimaryUserUpdatedSyncFailed` | **Yes** (`POST .../users/$ref`) | Primary User was set, but the optional `-SyncIntuneDevice` call failed (e.g., missing `PrivilegedOperations.All`). |
| `AlreadyInSync` | No | Authoritative `/users` relationship ObjectId already matches the resolved Entra User ObjectId. |
| `UnassignedInAvd` | No | `sessionHost.properties.assignedUser` is empty; skipped. |
| `PendingIntuneEnrollment` | No | Session host has an `AssignedUser` in AVD, but no matching Windows `managedDevice` exists in Intune yet. |
| `PrimaryUserConflict` | No | Intune device already has a different Primary User assigned. Fails closed to prevent unintended reassignment. |
| `AmbiguousIntuneDevice` | No | Multiple Intune `managedDevice` records share the session host's short hostname. Remove stale Intune records first. |
| `AmbiguousAvdSessionHost` | No | Multiple evaluated AVD session hosts share the same short hostname across pools. |
| `StaleIntuneDeviceRecord` | No | The sole Intune device record has an empty/zero `azureADDeviceId` or was enrolled >24 hours before the current AVD `sessionHost` resource was created. |
| `EntraUserNotFound` | No | `AssignedUser` was not found in Entra ID via direct UPN lookup or B2B guest `mail` fallback (HTTP 404). |
| `AmbiguousEntraUser` | No | Multiple Entra users matched the B2B guest UPN/`mail` filter (HTTP 409). |
| `GraphUserLookupFailed` | No | Non-404 Graph failure while resolving the user in Entra ID (e.g., `403 Forbidden`). |
| `GraphLookupFailed` | No | Failed to read `/beta/deviceManagement/managedDevices/{id}/users`. |
| `GraphUpdateFailed` | No | `POST .../users/$ref` failed. |
| `InvalidSessionHostName` | No | Could not derive a valid short hostname from the AVD session host record. |
| `WhatIf` | No | `-WhatIf` (or declined confirmation) suppressed the `POST .../users/$ref` write. |
---
## 6. Offline Regression Tests
Run the AST-based regression test suite in Windows PowerShell 5.1 (no live Azure or Graph connectivity required):
```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\avd\primary-user-sync\Test-AvdPersonalPrimaryUser.ps1
```
---
## 7. Microsoft Learn & Official References
- [Using Azure Virtual Desktop single-session with Microsoft Intune](https://learn.microsoft.com/en-us/intune/solutions/azure-virtual-desktop)
- [Microsoft Entra joined session hosts in Azure Virtual Desktop](https://learn.microsoft.com/en-us/azure/virtual-desktop/azure-ad-joined-session-hosts)
- [Configure personal desktop assignment type in Azure Virtual Desktop](https://learn.microsoft.com/en-us/azure/virtual-desktop/configure-host-pool-personal-desktop-assignment-type)
- [Change a device's primary user in Microsoft Intune](https://learn.microsoft.com/en-us/intune/device-management/inventory-and-status/find-primary-user)
- [Built-in Azure RBAC roles for Azure Virtual Desktop (`Desktop Virtualization Reader`, `Desktop Virtualization Host Pool Reader`)](https://learn.microsoft.com/en-us/azure/virtual-desktop/rbac)
- [AVD REST API: Host Pools - Get (`2024-04-03`)](https://learn.microsoft.com/en-us/rest/api/desktopvirtualization/host-pools/get?view=rest-desktopvirtualization-2024-04-03)
- [AVD REST API: Session Hosts - List (`2024-04-03`)](https://learn.microsoft.com/en-us/rest/api/desktopvirtualization/session-hosts/list?view=rest-desktopvirtualization-2024-04-03)
- [Microsoft Graph: `managedDevice` resource type & `users` relationship](https://learn.microsoft.com/en-us/graph/api/resources/intune-devices-manageddevice?view=graph-rest-1.0)
- [Microsoft Intune PowerShell Sample: `Set-PrimaryUser.ps1` (`users/$ref`)](https://github.com/microsoft/Intune.Graph/blob/main/Source/ManagedDevices/Set-PrimaryUser.ps1)
- [Microsoft Graph: `syncDevice` action (`DeviceManagementManagedDevices.PrivilegedOperations.All`)](https://learn.microsoft.com/en-us/graph/api/intune-devices-manageddevice-syncdevice?view=graph-rest-1.0)
- [Microsoft Graph: Get user (`User.Read.All`)](https://learn.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0)
- [Microsoft Graph throttling guidance (`Retry-After`)](https://learn.microsoft.com/en-us/graph/throttling)
- [Azure Automation runbook types](https://learn.microsoft.com/en-us/azure/automation/automation-runbook-types) & [Schedules in Azure Automation](https://learn.microsoft.com/en-us/azure/automation/shared-resources/schedules)
