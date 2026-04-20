# Deploy WinGet REST Source (Patched)

Deployment scripts and patched Azure Functions package for deploying a private WinGet REST source with fixes for [#313](https://github.com/microsoft/winget-cli-restsource/issues/313) and [#314](https://github.com/microsoft/winget-cli-restsource/issues/314).

## What's Included

| File | Purpose |
|------|---------|
| `Deploy-WinGetSource.ps1` | Full deployment script with `-RestSourcePath` parameter |
| `Remove-WinGetSource.ps1` | Complete teardown of a deployed REST source |
| `WinGet.RestSource.Functions.zip` | Patched Azure Functions package (isolated worker model) |

## What the Patch Fixes

**Issue #313 — Isolated Worker Migration:**
The upstream `winget-cli-restsource` Azure Functions project uses the deprecated in-process hosting model (end of support November 2026). The patched zip migrates to the isolated worker model, which is required for new deployments on Azure Functions v4 with .NET 8.

**Issue #314 — APIM v2 Tier Support:**
The upstream module only supports `Developer`, `Basic`, and `Enhanced` APIM tiers. The patched PowerShell module adds `BasicV2` and `StandardV2` tiers — required for deployments that need VNet integration or private endpoints.

## Prerequisites

- PowerShell 7.4+
- Azure PowerShell modules: `Az.Accounts` (≥2.7.5), `Az.Network`
- `Microsoft.WinGet.RestSource` PowerShell module (from PSGallery)
- Azure subscription with permissions to create resources

## Quick Start

### Deploy with patched Functions (recommended)

```powershell
.\Deploy-WinGetSource.ps1 -Name "corpwinget" `
    -ResourceGroup "rg-winget-prod-001" `
    -Region "westeurope" `
    -Authentication "MicrosoftEntraId" `
    -RestSourcePath ".\WinGet.RestSource.Functions.zip"
```

### Deploy with APIM v2 (requires patched PowerShell module)

```powershell
.\Deploy-WinGetSource.ps1 -Name "corpwinget" `
    -ResourceGroup "rg-winget-prod-001" `
    -PerformanceTier "BasicV2" `
    -Authentication "MicrosoftEntraId" `
    -RestSourcePath ".\WinGet.RestSource.Functions.zip"
```

### Deploy without APIM (lower cost)

```powershell
.\Deploy-WinGetSource.ps1 -Name "corpwinget" `
    -ResourceGroup "rg-winget-prod-001" `
    -Authentication "MicrosoftEntraId" `
    -SkipAPIM `
    -RestSourcePath ".\WinGet.RestSource.Functions.zip"
```

### Remove a deployment

```powershell
.\Remove-WinGetSource.ps1 -Name "corpwinget" `
    -ResourceGroup "rg-winget-prod-001"
```

### Remove — purge soft-deleted resources only

```powershell
.\Remove-WinGetSource.ps1 -Name "corpwinget" `
    -ResourceGroup "rg-winget-prod-001" `
    -PurgeOnly
```

### Remove — skip Entra ID cleanup (keep app registrations)

```powershell
.\Remove-WinGetSource.ps1 -Name "corpwinget" `
    -ResourceGroup "rg-winget-prod-001" `
    -SkipEntraCleanup
```

## Key Parameters (Deploy)

| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `-Name` | Yes | — | Base name for Azure resources (3-24 alphanumeric) |
| `-ResourceGroup` | No | `rg-winget-prod-001` | Target resource group |
| `-Region` | No | `westeurope` | Azure region |
| `-PerformanceTier` | No | `Developer` | `Developer`, `Basic`, `Enhanced`, `BasicV2`, `StandardV2` |
| `-Authentication` | No | `None` | `None` or `MicrosoftEntraId` |
| `-RestSourcePath` | No | *(module default)* | Path to patched Functions zip |
| `-SkipAPIM` | No | `$false` | Deploy without API Management |
| `-EnablePrivateEndpoints` | No | `$false` | Enable private endpoints |
| `-ManifestPath` | No | — | Folder with manifests to publish post-deploy |
| `-RegisterSource` | No | `$false` | Register as local winget source after deploy |

## Key Parameters (Remove)

| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `-Name` | Yes | — | Base name used during deployment (e.g., "corpwinget") |
| `-ResourceGroup` | No | `rg-winget-prod-001` | Resource group to delete |
| `-SubscriptionName` | No | *(current context)* | Azure subscription name |
| `-SourceDisplayName` | No | `CorpWinGet` | Local winget source display name to unregister |
| `-SkipSourceRemoval` | No | `$false` | Skip removing the local winget source registration |
| `-SkipEntraCleanup` | No | `$false` | Skip deleting Entra ID app registrations |
| `-SkipSoftDeletePurge` | No | `$false` | Skip purging soft-deleted resources (APIM, Key Vault, etc.) |
| `-PurgeOnly` | No | `$false` | Only purge soft-deleted resources (don't delete the RG) |

The removal script supports `-WhatIf` for dry-run validation.

## How `-RestSourcePath` Works

When omitted, the script uses the Functions zip bundled with the `Microsoft.WinGet.RestSource` PowerShell module. When specified, it overrides the zip in all three deployment code paths:

1. **Default path** — passed to `New-WinGetSource` via splatting
2. **Split workflow** — used as `$restSourceZip` for `New-ARMObjects`
3. **ZipDeploy fallback** — used for manual zip upload if cold start fails

This means you can use the upstream PowerShell module for ARM template deployment while deploying the patched Functions code.

## Source

Patched fork: [AdmiralTolwyn/winget-cli-restsource](https://github.com/AdmiralTolwyn/winget-cli-restsource/tree/fix/apim-v2-tiers-and-isolated-worker)

See `CHANGES.md` in the fork for detailed change documentation.
