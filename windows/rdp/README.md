# RDP File Signer (Per-User, No Admin)

`Sign-RdpFile.ps1` digitally signs `.rdp` files for the current user, with **no administrator rights required**. It supports two distinct workflows:

1. ✅ **Batch-sign with an existing CA-issued code-signing certificate** *(fully supported, recommended)*
   Pass `-Thumbprint` (SHA1 or SHA256) of a code-signing cert already present in `Cert:\CurrentUser\My` or `Cert:\LocalMachine\My` — for example one issued by your Enterprise CA, a commercial CA (DigiCert, Sectigo, …), or imported from a PFX. The script wraps `rdpsign.exe` to sign one file or a whole tree, registers the thumbprint in the per-user RDP trusted-publishers policy so users don't see the publisher prompt, and **does not touch the trust stores** (your CA chain already handles trust). This is the intended way to use this script for real deployments.

2. ⚠️ **Generate and use a self-signed code-signing certificate** *(experimental, lab/testing only)*
   With no `-Thumbprint` supplied, the script creates a self-signed `CodeSigningCert` in the current user's profile and installs it into the per-user Trusted Root + Trusted Publishers stores so signed files open without warnings on **this user's** machine. Adding to per-user Root triggers the standard one-time CryptoAPI "Security Warning" dialog (Yes/No), and a self-signed cert in Root means *any* certificate it issues is trusted by that user. **Do not use this in production** — use an Enterprise or commercial CA distributed via Group Policy / Intune instead.

## Why

Microsoft's April 2026 cumulative updates (KB5083769 / KB5082200, addressing **CVE-2026-26151**) changed how Windows treats unsigned `.rdp` files:

- Every double-click shows a **"Caution: Unknown remote connection"** warning, with no "don't ask again" option.
- All clipboard / drive / printer redirection requested by the file is **blocked by default** and must be re-enabled on every connection.

Signing the `.rdp` file with a code signing certificate suppresses the unknown-publisher warning and restores redirection. The standard guidance is to install the signing certificate into `Cert:\LocalMachine\*` and trust it via Group Policy — both of which require administrator rights. This script performs the equivalent steps against the per-user stores instead, so any standard user can sign their own `.rdp` files without elevation.

## How it works

1. **Selects a signing certificate**:
   - With `-Thumbprint <SHA1|SHA256>`: looks up the cert in `Cert:\CurrentUser\My`, then falls back to `Cert:\LocalMachine\My` (a standard user can read — just not write — the machine store). Verifies the cert has the Code Signing EKU and an associated private key, and uses it as-is. **No new cert is generated and the trust stores are left untouched** — trust is expected to come from the existing CA chain.
   - Without `-Thumbprint`: reuses (matched by `-Subject`, default `CN=$env:USERNAME RDP Signing`) or creates a self-signed `CodeSigningCert` in `Cert:\CurrentUser\My` (SHA256 / RSA 2048 / 3-year validity, configurable via `-ValidYears`). The cert is tagged via `FriendlyName` prefix `EndpointToolkit:RDPSigning` so `-Cleanup` can find it later.
2. **Trusts the cert for the current user** *(self-signed mode only — skipped for CA-issued certs)* by importing it into:
   - `Cert:\CurrentUser\Root` (Trusted Root CAs — per-user). **First time only**, Windows shows a "Security Warning" dialog asking the user to confirm the Root install. 
   - `Cert:\CurrentUser\TrustedPublisher` (no prompt).
3. **Suppresses the "Verify the publisher of this remote connection" dialog** by writing the thumbprint to the per-user RDP trusted-publishers policy: `HKCU:\Software\Policies\Microsoft\Windows NT\Terminal Services\TrustedCertThumbprints` (REG_SZ, semicolon-separated). This is the same policy GPMC's "Specify SHA1 thumbprints of certificates representing trusted .rdp publishers" writes to, but in the user hive (writable without admin).
4. **Signs** each `.rdp` file via `rdpsign.exe /sha256 <thumbprint> <file>` (in-box on Windows 10/11).
5. **Optionally exports** the public `.cer` so other users can trust the cert with `-InstallCerOnly`.

End result: after the one-time Root install confirmation, double-clicking the signed `.rdp` opens the connection with no further dialogs and full redirection enabled.

## Usage

### Sign a single file

```powershell
.\Sign-RdpFile.ps1 -Path 'C:\Users\me\Desktop\MyServer.rdp'
```

### Sign every .rdp file in a folder

```powershell
.\Sign-RdpFile.ps1 -Path 'C:\RDP'
```

### Use a custom subject and export the public cert

```powershell
.\Sign-RdpFile.ps1 -Path 'C:\RDP' `
                   -Subject 'CN=Contoso RDP, O=Contoso, C=AU' `
                   -ExportCerPath 'C:\RDP\contoso-rdp.cer'
```

### Re-sign already-signed files

```powershell
.\Sign-RdpFile.ps1 -Path 'C:\RDP' -Force
```

### Trust someone else's signing cert (recipient side, no signing)

Distribute the `.cer` alongside the `.rdp`. Each recipient runs once:

```powershell
.\Sign-RdpFile.ps1 -InstallCerOnly -Path 'C:\Temp\contoso-rdp.cer'
```

This silently imports the public certificate into the recipient's per-user TrustedPublisher store and registers the thumbprint with the per-user RDP trusted-publishers policy. The Root install will prompt once with the standard CryptoAPI "Security Warning" — no admin required.

### Use an existing certificate (e.g. an imported PFX or CA-issued code-signing cert)

Pass either the standard SHA1 thumbprint (40 hex chars) or the SHA256 fingerprint (64 hex chars). Spaces / colons in the value are stripped automatically:

```powershell
# SHA1
.\Sign-RdpFile.ps1 -Path 'C:\RDP\MyServer.rdp' -Thumbprint A1B2C3D4E5F6...

# SHA256 (e.g. copied from a CA portal or `certutil -hashfile`)
.\Sign-RdpFile.ps1 -Path 'C:\RDP\MyServer.rdp' `
                   -Thumbprint '9F:86:D0:81:88:4C:7D:65:9A:2F:EA:A0:C5:5A:D0:15:A3:BF:4F:1B:2B:0B:82:2C:D1:5D:6C:15:B0:F0:0A:08'
```

When `-Thumbprint` resolves to an existing cert in `Cert:\CurrentUser\My`, the script does **not** create a new self-signed cert and does **not** modify the trust stores beyond ensuring the publisher policy is in place — trust for a real CA-issued code-signing cert is expected to be handled by the CA chain (or by GPO / Intune for self-signed).

### Clean up everything this script has installed

```powershell
.\Sign-RdpFile.ps1 -Cleanup
```

Removes every certificate tagged `EndpointToolkit:RDPSigning` from:

- `Cert:\CurrentUser\My`
- `Cert:\CurrentUser\Root`
- `Cert:\CurrentUser\TrustedPublisher`
- Leftover entries under `HKCU:\Software\Microsoft\SystemCertificates\Root\Certificates\*` (defensive sweep, in case an older build of this script wrote there directly)

…and removes their thumbprints from `HKCU:\...\Terminal Services\TrustedCertThumbprints`.

The signed `.rdp` files themselves are not touched (after cleanup they will revert to showing the unknown-publisher warning).

## Parameters

| Parameter | Description |
|-----------|-------------|
| `-Path` | One or more `.rdp` files or folders (Sign mode), or `.cer` files (InstallCer mode). Folders searched recursively. |
| `-Subject` | Subject DN used to look up or create the signing cert. Default: `CN=$env:USERNAME RDP Signing`. |
| `-Thumbprint` | Use a specific cert from `Cert:\CurrentUser\My` or `Cert:\LocalMachine\My` (CurrentUser preferred). Accepts SHA1 (40 hex chars) **or** SHA256 (64 hex chars) thumbprint; spaces / colons tolerated. Overrides `-Subject`. |
| `-ValidYears` | Lifetime of a newly created cert (1-10). Default: 3. |
| `-ExportCerPath` | Export the public certificate to this path for distribution. |
| `-Force` | Re-sign files that already contain a `signature:s:` line. |
| `-InstallCerOnly` | Trust a supplied `.cer` for the current user only and exit. |
| `-Cleanup` | Remove every cert + policy entry this script has installed for the current user, then exit. |

## Verifying the result

After signing, opening the `.rdp` file should connect directly with no warnings. To inspect the signature:

```powershell
Select-String -Path 'C:\RDP\MyServer.rdp' -Pattern '^(signscope|signature):s:'
```

To see what this script has installed:

```powershell
# Signing certs (private key)
Get-ChildItem Cert:\CurrentUser\My | Where-Object FriendlyName -like '*EndpointToolkit:RDPSigning*'

# Trusted publishers policy
Get-ItemProperty 'HKCU:\Software\Policies\Microsoft\Windows NT\Terminal Services' -Name TrustedCertThumbprints
```

## Caveats

- **Per-user trust only.** The cert is trusted only for the user that ran the script. Other users on the same machine — or on different machines — must run `-InstallCerOnly` against the exported `.cer`, or you must distribute trust via Group Policy / Intune / Enterprise CA.
- **Self-signed ≠ public trust.** This is appropriate for internal / lab / small-team scenarios. For files distributed to external users at scale, use an Enterprise CA or commercial code signing certificate instead.
- **Modifying a signed file invalidates the signature.** Re-run the script (with `-Force`) after any edit to the `.rdp` file.
- **Private key is non-exportable** by default (standard `New-SelfSignedCertificate` behaviour). To move signing to another machine, generate the cert there too, or add `-KeyExportPolicy Exportable` and export a PFX yourself.
- **Cleanup is tag-based.** Only certs whose `FriendlyName` contains `EndpointToolkit:RDPSigning` are removed. Hand-imported certs without the tag are left alone.

## Doing this properly (recommended for production)

This script is a workaround. The supported, scalable, audit-friendly approach is to use a real code-signing certificate, sign on a controlled build machine, and distribute trust centrally. Source: Microsoft Learn — [`rdpsign`](https://learn.microsoft.com/windows-server/administration/windows-commands/rdpsign) and the [Group Policy reference for Remote Desktop Connection Client](https://learn.microsoft.com/windows/client-management/mdm/policy-csp-remotedesktopservices).

### 1. Obtain a code-signing certificate

Pick one based on who needs to trust the file:

| Audience | Certificate type | Notes |
|----------|------------------|-------|
| Internal, AD-joined | **Enterprise CA** (AD CS) | Free if you already run AD CS. All domain-joined devices trust it automatically. |
| External users / mixed estate | **Commercial code-signing cert** (DigiCert, Sectigo, GlobalSign, Entrust, …) | Trusted by every Windows device out of the box. EV variants require a hardware token. |
| Lab / single user | Self-signed | Use this script. |

The certificate **must** have the `Code Signing` Enhanced Key Usage (OID `1.3.6.1.5.5.7.3.3`). A web-server / TLS cert will not work.

**Enterprise CA via PowerShell** (on a domain-joined machine, with a `CodeSigning` template published):

```powershell
Get-Certificate -Template 'CodeSigning' -CertStoreLocation Cert:\CurrentUser\My
```

**Commercial CA**: import the issued PFX on the signing machine:

```powershell
$pfxPwd = Read-Host -AsSecureString 'PFX password'
Import-PfxCertificate -FilePath 'C:\Certs\company-codesigning.pfx' `
                      -CertStoreLocation Cert:\CurrentUser\My `
                      -Password $pfxPwd
```

### 2. Find the thumbprint

```powershell
Get-ChildItem Cert:\CurrentUser\My |
    Where-Object { $_.EnhancedKeyUsageList.ObjectId -contains '1.3.6.1.5.5.7.3.3' } |
    Select-Object Subject, Thumbprint, NotAfter
```

Copy the `Thumbprint` value (no spaces).

### 3. Sign the .rdp file with rdpsign.exe

`rdpsign.exe` ships in-box at `C:\Windows\System32\rdpsign.exe`. Finalise the `.rdp` file first — any edit after signing invalidates the signature.

```powershell
rdpsign.exe /sha256 <THUMBPRINT> "C:\RDP\MyServer.rdp"
```

Bulk:

```powershell
$thumb = '<THUMBPRINT>'
Get-ChildItem 'C:\RDP\*.rdp' | ForEach-Object {
    & "$env:WINDIR\System32\rdpsign.exe" /sha256 $thumb $_.FullName
}
```

`rdpsign.exe` appends `signscope:s:` and `signature:s:` lines to the file. Verify:

```powershell
Select-String -Path 'C:\RDP\MyServer.rdp' -Pattern '^(signscope|signature):s:'
```

### 4. Distribute trust centrally (so users get no prompts)

To make signed files open with no warnings, the cert chain must be trusted **and** the signing thumbprint must be in the RDP trusted-publishers policy on each client.

**Group Policy** (`gpmc.msc` — domain admin):

1. **Trust the cert chain** (only needed for self-signed / internal Enterprise CA — commercial CAs are pre-trusted):
   `Computer Configuration → Windows Settings → Security Settings → Public Key Policies`
   - Import the public `.cer` into **Trusted Root Certification Authorities** (self-signed) or **Intermediate Certification Authorities** (Enterprise CA chain).
   - Import into **Trusted Publishers**.
2. **Suppress the "Verify the publisher" RDP dialog**:
   `Computer Configuration → Administrative Templates → Windows Components → Remote Desktop Services → Remote Desktop Connection Client → Specify SHA1 thumbprints of certificates representing trusted .rdp publishers` → Enabled, paste the signing-cert thumbprint(s) (semicolon-separated, no spaces).
3. (Optional, Microsoft recommended) `Allow .rdp files from unknown publishers` → **Disabled**, and `Allow .rdp files from valid publishers and user's default .rdp settings` → **Enabled**.

**Microsoft Intune** equivalents:

- Push the public `.cer` via **Devices → Configuration → Trusted certificate** profiles, scoped to Trusted Root and Trusted Publishers stores.
- Push the trusted-thumbprints policy via the **Settings Catalog**: search for *"Specify SHA1 thumbprints of certificates representing trusted .rdp publishers"* (under Administrative Templates → Remote Desktop Services → Remote Desktop Connection Client) and paste the thumbprint(s).

### 5. Operational hygiene

- **Sign on a single, controlled build machine** — ideally as part of a pipeline, not from each admin's workstation. Treat the signing key like any other secret.
- **Protect the private key.** Use a non-exportable cert, an HSM/TPM-backed key, or an EV token where available. Limit who can read the cert in the signing store.
- **Re-sign on every change.** Any edit to the `.rdp` after signing breaks the signature; bake signing into whatever produces the file.
- **Plan for rotation.** Keep the previous thumbprint in the trusted-publishers policy until all old `.rdp` files are re-signed with the new cert, then remove it.
- **Don't trust self-signed certs in production.** A self-signed code-signing cert deployed to `Trusted Root` means *any* certificate it issues is trusted by that user/machine. Use Enterprise CA or a commercial CA for anything beyond a lab.

## Exit codes

| Code | Meaning |
|------|---------|
| 0 | All files signed (or Cleanup completed, or InstallCerOnly succeeded). |
| 1 | No `.rdp` files found, or `-InstallCerOnly` failed. |
| 2 | Certificate setup failed. |
| 3 | One or more files failed to sign. |
