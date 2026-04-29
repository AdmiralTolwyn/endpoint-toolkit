<#
.SYNOPSIS
    Self-signs one or more .rdp files using a per-user code signing certificate.
    Requires NO administrator rights — certificate generation, trust, signing, and
    cleanup all run in the current user's profile.

    *** EXPERIMENTAL — NOT RECOMMENDED FOR PRODUCTION ***
    This script deliberately suppresses the RDP "Verify the publisher" dialog
    via the per-user RDP TrustedCertThumbprints policy. Adding the cert to the
    per-user Trusted Root store still triggers the standard one-time CryptoAPI
    confirmation dialog — the user must click Yes once. Use only for testing /
    lab / short-lived personal workflows. For real deployments, use an
    Enterprise CA or commercial code-signing certificate distributed via Group
    Policy or Intune.

.DESCRIPTION
    Microsoft's April 2026 cumulative updates (KB5083769 / KB5082200, CVE-2026-26151)
    block clipboard / drive redirection and force a "Caution: Unknown remote
    connection" warning every time an unsigned .rdp file is opened. The accepted
    fix is to digitally sign the .rdp file with a code signing certificate.

    This script automates the full self-signed flow without needing admin rights:

      1. Reuses (or creates) a CodeSigning certificate in Cert:\CurrentUser\My,
         tagged via FriendlyName so it can be located again for cleanup.
      2. Trusts that certificate for the current user by importing it into:
           - Cert:\CurrentUser\Root              (Trusted Root CAs - per-user)
           - Cert:\CurrentUser\TrustedPublisher  (Trusted Publishers - per-user)
         Adding to Root prompts the user with a one-time "Security Warning"
         dialog (CryptoAPI behaviour, no API to suppress for per-user Root).
      3. Adds the thumbprint to the per-user policy
           HKCU\Software\Policies\Microsoft\Windows NT\Terminal Services
           \TrustedCertThumbprints
         which suppresses the "Verify the publisher of this remote connection"
         dialog for files signed by this cert.
      4. Signs each .rdp file with rdpsign.exe /sha256.
      5. Optionally exports the public .cer for distribution to other users
         (who can run this script with -InstallCerOnly to trust it themselves).

    A -Cleanup mode removes every artefact tagged by this script (certs in My,
    Root, TrustedPublisher; the policy thumbprint entry).

    All of the above is what the article at
    https://pip.com.au/digitally-sign-rdp-files-a-complete-how-to/
    describes for the LocalMachine store (which needs admin) — re-targeted at
    the CurrentUser store and HKCU policy hive so any standard user can run it.

.PARAMETER Path
    One or more .rdp files (Sign mode), or .cer files (InstallCer mode), or
    folders containing .rdp files (folders are searched recursively).
    Not required in Cleanup mode.

.PARAMETER Subject
    Subject (CN=...) used when CREATING a new self-signed certificate, and
    when LOOKING UP an existing one in Cert:\CurrentUser\My.
    Default: "CN=$env:USERNAME RDP Signing".

.PARAMETER Thumbprint
    Use a specific certificate by thumbprint. Searched in Cert:\CurrentUser\My
    first, then Cert:\LocalMachine\My (a standard user can read but not write
    the machine store). Accepts either the standard SHA1 cert thumbprint
    (40 hex chars, what Windows shows as "Thumbprint") or the cert's SHA256
    fingerprint (64 hex chars). Spaces / colons are tolerated.
    Overrides -Subject lookup. Useful if you've already imported a PFX
    or have a code-signing cert from your CA in either store.

.PARAMETER ValidYears
    Lifetime of a newly created certificate. Default: 3 years. Ignored if a
    matching cert already exists or -Thumbprint is supplied.

.PARAMETER ExportCerPath
    Optional path to write the public certificate (.cer) for distribution.

.PARAMETER InstallCerOnly
    Install a previously exported .cer (passed via -Path) into the current
    user's Trusted Root + Trusted Publishers stores AND the per-user RDP
    trusted-publishers policy, then exit. Use this on recipient machines to
    trust someone else's signing cert without signing anything.

.PARAMETER Cleanup
    Remove every certificate this script has installed (matched by FriendlyName
    tag "EndpointToolkit:RDPSigning") from CurrentUser\My, the HKCU Root
    registry, and CurrentUser\TrustedPublisher, plus their thumbprints from
    the HKCU TrustedCertThumbprints policy. -Path is ignored in this mode.

.PARAMETER Force
    Re-sign files even if they already contain a signature line. By default
    already-signed files are skipped.

.EXAMPLE
    .\Sign-RdpFile.ps1 -Path 'C:\RDP\MyServer.rdp'

    Creates (or reuses) a self-signed cert, trusts it for the current user
    silently, adds it to the trusted-publishers RDP policy, and signs
    MyServer.rdp. End user double-clicks the file with no warnings.

.EXAMPLE
    .\Sign-RdpFile.ps1 -Path 'C:\RDP' -Subject 'CN=Contoso RDP, O=Contoso, C=AU' `
                       -ExportCerPath 'C:\RDP\contoso-rdp.cer'

    Signs every .rdp file under C:\RDP and exports the public cert so it can
    be distributed to other users.

.EXAMPLE
    .\Sign-RdpFile.ps1 -InstallCerOnly -Path 'C:\Temp\contoso-rdp.cer'

    Trusts the supplied public certificate for the current user only and
    bypasses both warning dialogs for files signed by it. No signing performed.

.EXAMPLE
    .\Sign-RdpFile.ps1 -Cleanup

    Removes every certificate / registry entry / policy thumbprint this script
    has ever created for the current user.

.NOTES
    Version : 1.2.2
    Author  : Anton Romanyuk
    Requires: Windows 10/11, PowerShell 5.1+, rdpsign.exe (in-box).

.DISCLAIMER
    This script is provided "AS IS" with no warranties and confers no rights.
    It is not supported under any Microsoft standard support program or service.
    Use of this script is entirely at your own risk. The customer is solely
    responsible for testing and validating this script in their environment
    before deploying to production. The author shall not be liable for any
    damage or data loss resulting from the use of this script.
#>

[CmdletBinding(DefaultParameterSetName = 'Sign')]
param(
    [Parameter(Mandatory, Position = 0, ParameterSetName = 'Sign')]
    [Parameter(Mandatory, Position = 0, ParameterSetName = 'InstallCer')]
    [string[]] $Path,

    [Parameter(ParameterSetName = 'Sign')]
    [string] $Subject = "CN=$env:USERNAME RDP Signing",

    [Parameter(ParameterSetName = 'Sign')]
    [string] $Thumbprint,

    [Parameter(ParameterSetName = 'Sign')]
    [ValidateRange(1, 10)]
    [int] $ValidYears = 3,

    [Parameter(ParameterSetName = 'Sign')]
    [string] $ExportCerPath,

    [Parameter(ParameterSetName = 'Sign')]
    [switch] $Force,

    [Parameter(Mandatory, ParameterSetName = 'InstallCer')]
    [switch] $InstallCerOnly,

    [Parameter(Mandatory, ParameterSetName = 'Cleanup')]
    [switch] $Cleanup
)

#region --- Constants ---

# Stable tag written into FriendlyName so we can locate our certs later for cleanup.
$Script:CertTag = 'EndpointToolkit:RDPSigning'

# Per-user RDP trusted-publishers policy. Same policy GPMC writes for "Specify SHA1
# thumbprints of certificates representing trusted .rdp publishers", but in HKCU
# (writable without admin). Adding a thumbprint here suppresses the
# "Verify the publisher of this remote connection" dialog for files signed by it.
$Script:TrustedPubPolicyKey   = 'HKCU:\Software\Policies\Microsoft\Windows NT\Terminal Services'
$Script:TrustedPubPolicyValue = 'TrustedCertThumbprints'

# HKCU root-cert registry path — used ONLY by cleanup, to find leftover entries.
$Script:HkcuRootKey = 'HKCU:\Software\Microsoft\SystemCertificates\Root\Certificates'

#endregion

#region --- Helpers ---

function Write-Log {
<#
.SYNOPSIS
    Writes a colour-coded, timestamped, level-tagged line to the console.
.DESCRIPTION
    Lightweight console logger used throughout the script. Format:
        [HH:mm:ss] [LEVEL] message
.PARAMETER Message
    Free-form text to print.
.PARAMETER Level
    INFO | WARN | ERROR | SUCCESS. Default INFO.
#>
    param(
        [string] $Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')]
        [string] $Level = 'INFO'
    )
    $ts = (Get-Date).ToString('HH:mm:ss')
    $color = switch ($Level) {
        'INFO'    { 'Gray' }
        'WARN'    { 'Yellow' }
        'ERROR'   { 'Red' }
        'SUCCESS' { 'Green' }
    }
    Write-Host ("[{0}] [{1}] {2}" -f $ts, $Level, $Message) -ForegroundColor $color
}

function Resolve-RdpFiles {
<#
.SYNOPSIS
    Expands -Path inputs into a flat list of .rdp FileInfo objects.
.DESCRIPTION
    Accepts any mix of file paths and folder paths:
      * Folders are searched recursively for *.rdp
      * Files with a .rdp extension are taken as-is
      * Anything else is logged at WARN and ignored
    Missing paths log a WARN and are skipped (never throw) so a single bad
    entry in a batch does not abort the whole signing pass.
.PARAMETER Inputs
    One or more file or folder paths.
.OUTPUTS
    System.Collections.Generic.List[System.IO.FileInfo]
#>
    param([string[]] $Inputs)
    $results = New-Object System.Collections.Generic.List[System.IO.FileInfo]
    foreach ($p in $Inputs) {
        if (-not (Test-Path -LiteralPath $p)) {
            Write-Log "Path not found: $p" -Level WARN
            continue
        }
        $item = Get-Item -LiteralPath $p
        if ($item.PSIsContainer) {
            Get-ChildItem -LiteralPath $item.FullName -Filter '*.rdp' -File -Recurse |
                ForEach-Object { [void]$results.Add($_) }
        }
        elseif ($item.Extension -eq '.rdp') {
            [void]$results.Add($item)
        }
        else {
            Write-Log "Skipping non-.rdp file: $($item.FullName)" -Level WARN
        }
    }
    return $results
}

function Get-CertSha256Thumbprint {
<#
.SYNOPSIS
    Computes a certificate's SHA256 fingerprint as a 64-char uppercase hex string.
.DESCRIPTION
    Windows exposes the SHA1 thumbprint via X509Certificate2.Thumbprint, but
    many CA portals and signing services hand out the SHA256 fingerprint
    instead. This helper lets -Thumbprint accept either form.
.PARAMETER Cert
    X509Certificate2 to fingerprint.
.OUTPUTS
    [string] uppercase hex SHA256 fingerprint, no separators.
#>
    param([Parameter(Mandatory)] [System.Security.Cryptography.X509Certificates.X509Certificate2] $Cert)
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        return ([System.BitConverter]::ToString($sha256.ComputeHash($Cert.RawData))).Replace('-', '').ToUpperInvariant()
    }
    finally { $sha256.Dispose() }
}

function Find-CertByThumbprint {
<#
.SYNOPSIS
    Locates a certificate by SHA1 (40 hex) or SHA256 (64 hex) thumbprint across one or more cert stores.
.DESCRIPTION
    Looks up a cert by either its SHA1 (40 hex chars, the value Windows shows
    as "Thumbprint") or its SHA256 (64 hex chars, what some portals / signing
    services hand out). Spaces and colons are tolerated.

    Searches the supplied -StorePaths in order and returns the first match.
    Default order is CurrentUser\My then LocalMachine\My, so a user-installed
    cert is preferred but a machine-wide one is still found (a standard user
    can READ LocalMachine\My without admin - they just can't write to it).
.PARAMETER Thumbprint
    SHA1 (40 hex) or SHA256 (64 hex) fingerprint. Spaces and colons accepted.
.PARAMETER StorePaths
    Cert: provider paths to scan, in priority order. Default:
        Cert:\CurrentUser\My, Cert:\LocalMachine\My
.OUTPUTS
    PSCustomObject with Cert and StorePath, or $null when no match.
#>
    param(
        [Parameter(Mandatory)] [string]   $Thumbprint,
        [string[]]                        $StorePaths = @('Cert:\CurrentUser\My', 'Cert:\LocalMachine\My')
    )

    $clean = ($Thumbprint -replace '[\s:]', '').ToUpperInvariant()
    if ($clean -notmatch '^[0-9A-F]+$') {
        throw "Thumbprint contains non-hex characters: '$Thumbprint'."
    }
    if ($clean.Length -ne 40 -and $clean.Length -ne 64) {
        throw "Thumbprint must be 40 hex chars (SHA1) or 64 hex chars (SHA256). Got $($clean.Length)."
    }

    foreach ($path in $StorePaths) {
        if (-not (Test-Path -LiteralPath $path)) { continue }
        $match = Get-ChildItem -LiteralPath $path -ErrorAction SilentlyContinue | Where-Object {
            if ($clean.Length -eq 40) { $_.Thumbprint -eq $clean }
            else                      { (Get-CertSha256Thumbprint -Cert $_) -eq $clean }
        } | Select-Object -First 1
        if ($match) {
            return [pscustomobject]@{ Cert = $match; StorePath = $path }
        }
    }
    return $null
}

function Get-OrCreateSigningCert {
<#
.SYNOPSIS
    Returns a usable code-signing certificate, creating a self-signed one if needed.
.DESCRIPTION
    Resolution order:
      1. If -Thumbprint is supplied, locate that exact cert in CurrentUser\My
         (or LocalMachine\My) and validate it has the Code Signing EKU and a
         private key. Throws an actionable error (with import/move snippet) if
         the cert exists in Root/Trust/CA/TrustedPublisher instead - rdpsign
         only signs from My.
      2. Otherwise, look in Cert:\CurrentUser\My for an existing, unexpired,
         private-key-bearing CodeSigning cert with the supplied -Subject.
         Reuse the latest match and (re-)tag its FriendlyName for cleanup.
      3. As a last resort, generate a new SHA256 / RSA-2048 self-signed cert
         valid for -ValidYears.
.PARAMETER Subject
    Distinguished name (CN=...). Used both for matching existing certs and as
    the Subject of any newly created cert.
.PARAMETER Thumbprint
    Optional. SHA1 or SHA256 fingerprint of an existing cert to reuse.
.PARAMETER ValidYears
    Lifetime of a newly created cert. Ignored when an existing cert is reused.
.OUTPUTS
    [System.Security.Cryptography.X509Certificates.X509Certificate2]
#>
    param(
        [string] $Subject,
        [string] $Thumbprint,
        [int]    $ValidYears
    )

    $store = 'Cert:\CurrentUser\My'

    if ($Thumbprint) {
        $hit = Find-CertByThumbprint -Thumbprint $Thumbprint
        if (-not $hit) {
            # Cert wasn't in My — check the other common stores so we can give a
            # specific, actionable error instead of a generic "not found".
            # rdpsign.exe ONLY signs from CurrentUser\My / LocalMachine\My, so a
            # cert living in Root/Trust/CA/TrustedPublisher needs to be moved or
            # imported into My first.
            $otherStores = @(
                'Cert:\CurrentUser\Root',  'Cert:\LocalMachine\Root',
                'Cert:\CurrentUser\Trust', 'Cert:\LocalMachine\Trust',
                'Cert:\CurrentUser\CA',    'Cert:\LocalMachine\CA',
                'Cert:\CurrentUser\TrustedPublisher', 'Cert:\LocalMachine\TrustedPublisher'
            )
            $misplaced = Find-CertByThumbprint -Thumbprint $Thumbprint -StorePaths $otherStores
            if ($misplaced) {
                $hint = @"
Certificate $($misplaced.Cert.Thumbprint) was found in $($misplaced.StorePath), not in CurrentUser\My / LocalMachine\My.
rdpsign.exe only signs using certificates from the 'Personal' (My) store.
Code-signing certs (and self-signed signing certs) must live in My — Root is for trust anchors only.

Fix: import/move the cert (with its private key) into Cert:\CurrentUser\My, e.g.:

  # If you have the original PFX:
  `$pwd = Read-Host -AsSecureString 'PFX password'
  Import-PfxCertificate -FilePath 'C:\path\to\codesigning.pfx' ``
                        -CertStoreLocation Cert:\CurrentUser\My ``
                        -Password `$pwd

  # If the cert in $($misplaced.StorePath) has an exportable private key:
  `$src = Get-Item '$($misplaced.StorePath)\$($misplaced.Cert.Thumbprint)'
  `$tmp = Join-Path `$env:TEMP 'codesign-move.pfx'
  `$pwd = ConvertTo-SecureString 'temp' -AsPlainText -Force
  Export-PfxCertificate -Cert `$src -FilePath `$tmp -Password `$pwd | Out-Null
  Import-PfxCertificate  -FilePath `$tmp -CertStoreLocation Cert:\CurrentUser\My -Password `$pwd | Out-Null
  Remove-Item `$tmp -Force
"@
                throw $hint
            }
            throw "No certificate matching thumbprint '$Thumbprint' found in CurrentUser\My or LocalMachine\My (tried SHA1 / SHA256)."
        }
        $cert     = $hit.Cert
        $foundIn  = $hit.StorePath
        if ($cert.EnhancedKeyUsageList.ObjectId -notcontains '1.3.6.1.5.5.7.3.3') {
            throw "Certificate $($cert.Thumbprint) does not have the Code Signing EKU (1.3.6.1.5.5.7.3.3)."
        }
        if (-not $cert.HasPrivateKey) {
            throw "Certificate $($cert.Thumbprint) found in $foundIn has no associated private key — cannot sign."
        }
        Write-Log "Using existing certificate by thumbprint: $($cert.Subject) [SHA1=$($cert.Thumbprint)] (from $foundIn)" -Level INFO
        return $cert
    }

    $existing = Get-ChildItem $store |
        Where-Object {
            $_.Subject -eq $Subject -and
            $_.NotAfter -gt (Get-Date) -and
            $_.HasPrivateKey -and
            $_.EnhancedKeyUsageList.ObjectId -contains '1.3.6.1.5.5.7.3.3'
        } |
        Sort-Object NotAfter -Descending |
        Select-Object -First 1

    if ($existing) {
        # Make sure it carries our cleanup tag even if it was created by an older
        # version of this script.
        if ($existing.FriendlyName -notmatch [regex]::Escape($Script:CertTag)) {
            try { $existing.FriendlyName = "$($Script:CertTag) | $Subject" } catch { }
        }
        Write-Log "Reusing existing signing certificate: $($existing.Subject) [$($existing.Thumbprint)] (expires $($existing.NotAfter.ToString('yyyy-MM-dd')))" -Level INFO
        return $existing
    }

    Write-Log "Creating new self-signed code signing certificate: $Subject" -Level INFO
    $cert = New-SelfSignedCertificate `
        -Type CodeSigningCert `
        -Subject $Subject `
        -KeyUsage DigitalSignature `
        -KeyAlgorithm RSA `
        -KeyLength 2048 `
        -HashAlgorithm SHA256 `
        -FriendlyName "$($Script:CertTag) | $Subject" `
        -CertStoreLocation $store `
        -NotAfter (Get-Date).AddYears($ValidYears)

    Write-Log "Created certificate: $($cert.Thumbprint) (expires $($cert.NotAfter.ToString('yyyy-MM-dd')))" -Level SUCCESS
    return $cert
}

function Add-CertToCurrentUserStore {
<#
.SYNOPSIS
    Imports the public portion of a certificate into a CurrentUser cert store.
.DESCRIPTION
    Used to populate Cert:\CurrentUser\Root (trust anchor) and
    Cert:\CurrentUser\TrustedPublisher (suppress publisher warning) without
    needing admin rights.

    Adding to Root triggers the standard one-time CryptoAPI security warning
    dialog the first time per user - that is intentional Windows behaviour for
    the per-user Root store and there is no documented API to suppress it.

    Only the public RawData is imported, and the FriendlyName tag is preserved
    so -Cleanup can find and remove it later.
.PARAMETER Cert
    X509Certificate2 to install (private key, if any, is NOT exported).
.PARAMETER StoreName
    'Root' or 'TrustedPublisher'.
#>
    param(
        [Parameter(Mandatory)] [System.Security.Cryptography.X509Certificates.X509Certificate2] $Cert,
        [Parameter(Mandatory)] [ValidateSet('Root', 'TrustedPublisher')] [string] $StoreName
    )

    $existing = Get-ChildItem "Cert:\CurrentUser\$StoreName" -ErrorAction SilentlyContinue |
        Where-Object { $_.Thumbprint -eq $Cert.Thumbprint }
    if ($existing) {
        Write-Log "Certificate already trusted in CurrentUser\$StoreName" -Level INFO
        return
    }

    # Tag the public-only copy too so cleanup can locate it.
    $publicOnly = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($Cert.RawData)
    if ($Cert.FriendlyName) { $publicOnly.FriendlyName = $Cert.FriendlyName }

    $store = [System.Security.Cryptography.X509Certificates.X509Store]::new(
        $StoreName,
        [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
    $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
    try {
        $store.Add($publicOnly)
        Write-Log "Trusted certificate in CurrentUser\$StoreName" -Level SUCCESS
    }
    finally {
        $store.Close()
    }
}

function Add-RdpTrustedPublisherThumbprint {
<#
.SYNOPSIS
    Adds a SHA1 thumbprint to the per-user RDP TrustedCertThumbprints policy.
.DESCRIPTION
    Adds the thumbprint to the per-user RDP trusted-publishers policy, which
    suppresses the "Verify the publisher of this remote connection" dialog
    for files signed with this cert.

    Stored as REG_SZ semicolon-separated, matching the format produced by the
    GPMC policy "Specify SHA1 thumbprints of certificates representing trusted
    .rdp publishers". Already-present thumbprints are not duplicated.
.PARAMETER Thumbprint
    SHA1 thumbprint of the cert (spaces tolerated, case-insensitive).
#>
    param([Parameter(Mandatory)] [string] $Thumbprint)

    $thumb = ($Thumbprint -replace '\s', '').ToUpperInvariant()

    if (-not (Test-Path -LiteralPath $Script:TrustedPubPolicyKey)) {
        New-Item -Path $Script:TrustedPubPolicyKey -Force | Out-Null
    }

    $current = (Get-ItemProperty -LiteralPath $Script:TrustedPubPolicyKey -Name $Script:TrustedPubPolicyValue -ErrorAction SilentlyContinue).$($Script:TrustedPubPolicyValue)
    $list = @()
    if ($current) {
        $list = $current.Split(';', [StringSplitOptions]::RemoveEmptyEntries) |
                ForEach-Object { $_.Trim().ToUpperInvariant() } |
                Where-Object { $_ }
    }

    if ($list -contains $thumb) {
        Write-Log "Thumbprint already present in RDP TrustedCertThumbprints policy" -Level INFO
        return
    }

    $list += $thumb
    Set-ItemProperty -LiteralPath $Script:TrustedPubPolicyKey `
                     -Name $Script:TrustedPubPolicyValue `
                     -Value ($list -join ';') `
                     -Type String -Force
    Write-Log "Added thumbprint to per-user RDP TrustedCertThumbprints policy (suppresses publisher prompt)" -Level SUCCESS
}

function Remove-RdpTrustedPublisherThumbprint {
<#
.SYNOPSIS
    Removes one or more thumbprints from the per-user RDP TrustedCertThumbprints policy.
.DESCRIPTION
    Reverses Add-RdpTrustedPublisherThumbprint. If the resulting list is empty
    the registry value is deleted entirely (rather than left as an empty
    REG_SZ) so the policy effectively reverts to default. Used by -Cleanup.
.PARAMETER Thumbprints
    SHA1 thumbprints to drop (spaces tolerated, case-insensitive).
#>
    param([Parameter(Mandatory)] [string[]] $Thumbprints)

    if (-not (Test-Path -LiteralPath $Script:TrustedPubPolicyKey)) { return }
    $current = (Get-ItemProperty -LiteralPath $Script:TrustedPubPolicyKey -Name $Script:TrustedPubPolicyValue -ErrorAction SilentlyContinue).$($Script:TrustedPubPolicyValue)
    if (-not $current) { return }

    $remove = $Thumbprints | ForEach-Object { ($_ -replace '\s', '').ToUpperInvariant() }
    $kept   = $current.Split(';', [StringSplitOptions]::RemoveEmptyEntries) |
              ForEach-Object { $_.Trim().ToUpperInvariant() } |
              Where-Object { $_ -and ($remove -notcontains $_) }

    if ($kept.Count -eq 0) {
        Remove-ItemProperty -LiteralPath $Script:TrustedPubPolicyKey -Name $Script:TrustedPubPolicyValue -Force -ErrorAction SilentlyContinue
        Write-Log "Cleared RDP TrustedCertThumbprints policy value" -Level SUCCESS
    }
    else {
        Set-ItemProperty -LiteralPath $Script:TrustedPubPolicyKey `
                         -Name $Script:TrustedPubPolicyValue `
                         -Value ($kept -join ';') `
                         -Type String -Force
        Write-Log "Removed $($Thumbprints.Count) thumbprint(s) from RDP TrustedCertThumbprints policy" -Level SUCCESS
    }
}

function Invoke-RdpSign {
<#
.SYNOPSIS
    Signs a single .rdp file with rdpsign.exe /sha256 using the supplied thumbprint.
.DESCRIPTION
    Wraps %WinDir%\System32\rdpsign.exe so the call sites stay tidy and every
    invocation produces a uniform result object.

    Behaviour:
      * If the file already contains a 'signature:s:' line and -Force is NOT
        set, the file is skipped (Status = 'Skipped').
      * Throws if rdpsign.exe is not present (unsupported Windows SKU).
      * Captures stdout+stderr and the exit code into the result Detail field
        so failures are diagnosable from the returned object alone.
.PARAMETER File
    FileInfo for the .rdp file to sign.
.PARAMETER Thumbprint
    SHA1 thumbprint of a cert in CurrentUser\My or LocalMachine\My.
.PARAMETER Force
    Re-sign even when an existing signature is present.
.OUTPUTS
    PSCustomObject (File, Status = Signed|Skipped|Failed, Detail).
#>
    param(
        [Parameter(Mandatory)] [System.IO.FileInfo] $File,
        [Parameter(Mandatory)] [string] $Thumbprint,
        [switch] $Force
    )

    if (-not $Force) {
        $hasSig = Select-String -LiteralPath $File.FullName -Pattern '^signature:s:' -Quiet -ErrorAction SilentlyContinue
        if ($hasSig) {
            Write-Log "Already signed (use -Force to re-sign): $($File.Name)" -Level WARN
            return [pscustomobject]@{ File = $File.FullName; Status = 'Skipped'; Detail = 'Already signed' }
        }
    }

    $rdpsign = Join-Path $env:WINDIR 'System32\rdpsign.exe'
    if (-not (Test-Path -LiteralPath $rdpsign)) {
        throw "rdpsign.exe not found at $rdpsign — unsupported Windows version."
    }

    $stdout = & $rdpsign /sha256 $Thumbprint $File.FullName 2>&1
    $exit   = $LASTEXITCODE

    if ($exit -eq 0) {
        Write-Log "Signed: $($File.FullName)" -Level SUCCESS
        return [pscustomobject]@{ File = $File.FullName; Status = 'Signed'; Detail = ($stdout -join '; ') }
    }
    else {
        Write-Log "rdpsign.exe failed (exit=$exit) for $($File.FullName): $($stdout -join '; ')" -Level ERROR
        return [pscustomobject]@{ File = $File.FullName; Status = 'Failed'; Detail = "exit=$exit; $($stdout -join '; ')" }
    }
}

function Invoke-Cleanup {
<#
.SYNOPSIS
    Removes every artefact this script has installed under the current user.
.DESCRIPTION
    Reverses everything Sign-RdpFile.ps1 does to the user profile:
      1. Deletes certificates tagged with $Script:CertTag from
         Cert:\CurrentUser\My (including the private key).
      2. Removes the matching public copies from Cert:\CurrentUser\Root
         (untrusts as a root CA for this user).
      3. Removes the matching public copies from Cert:\CurrentUser\TrustedPublisher.
      4. Strips the corresponding thumbprints from the per-user RDP
         TrustedCertThumbprints policy via Remove-RdpTrustedPublisherThumbprint.

    Identification is done by FriendlyName tag (set at create time) plus the
    set of thumbprints removed from My, so certs imported by this script - but
    not by the user's own hand - are the only ones touched.
#>
    Write-Log "Cleanup: removing certificates tagged '$($Script:CertTag)' from CurrentUser stores..." -Level INFO

    $tagPattern = "*$($Script:CertTag)*"
    $removedThumbs = New-Object System.Collections.Generic.HashSet[string]

    # 1) CurrentUser\My
    Get-ChildItem 'Cert:\CurrentUser\My' -ErrorAction SilentlyContinue |
        Where-Object { $_.FriendlyName -like $tagPattern } |
        ForEach-Object {
            try {
                Remove-Item -LiteralPath $_.PSPath -Force -DeleteKey
                [void]$removedThumbs.Add($_.Thumbprint.ToUpperInvariant())
                Write-Log "Removed cert from CurrentUser\My: $($_.Subject) [$($_.Thumbprint)]" -Level SUCCESS
            }
            catch {
                Write-Log "Failed to remove $($_.Thumbprint) from My: $_" -Level WARN
            }
        }

    # 2) CurrentUser\Root
    Get-ChildItem 'Cert:\CurrentUser\Root' -ErrorAction SilentlyContinue |
        Where-Object { $_.FriendlyName -like $tagPattern -or $removedThumbs.Contains($_.Thumbprint.ToUpperInvariant()) } |
        ForEach-Object {
            try {
                Remove-Item -LiteralPath $_.PSPath -Force
                [void]$removedThumbs.Add($_.Thumbprint.ToUpperInvariant())
                Write-Log "Removed cert from CurrentUser\Root: $($_.Thumbprint)" -Level SUCCESS
            }
            catch {
                Write-Log "Failed to remove $($_.Thumbprint) from Root: $_" -Level WARN
            }
        }

    # 3) CurrentUser\TrustedPublisher
    Get-ChildItem 'Cert:\CurrentUser\TrustedPublisher' -ErrorAction SilentlyContinue |
        Where-Object { $_.FriendlyName -like $tagPattern -or $removedThumbs.Contains($_.Thumbprint.ToUpperInvariant()) } |
        ForEach-Object {
            try {
                Remove-Item -LiteralPath $_.PSPath -Force
                [void]$removedThumbs.Add($_.Thumbprint.ToUpperInvariant())
                Write-Log "Removed cert from CurrentUser\TrustedPublisher: $($_.Thumbprint)" -Level SUCCESS
            }
            catch {
                Write-Log "Failed to remove $($_.Thumbprint) from TrustedPublisher: $_" -Level WARN
            }
        }

    # 4) HKCU Root registry — defensive sweep for leftover entries (older versions
    # of this script wrote here directly; keeps cleanup compatible with them).
    if (Test-Path -LiteralPath $Script:HkcuRootKey) {
        Get-ChildItem -LiteralPath $Script:HkcuRootKey -ErrorAction SilentlyContinue | ForEach-Object {
            $thumb = $_.PSChildName.ToUpperInvariant()
            if ($removedThumbs.Contains($thumb)) {
                try {
                    Remove-Item -LiteralPath $_.PSPath -Recurse -Force
                    Write-Log "Removed leftover Root registry entry: $thumb" -Level SUCCESS
                }
                catch {
                    Write-Log "Failed to remove leftover Root registry entry $thumb : $_" -Level WARN
                }
            }
        }
    }

    # 4) RDP trusted publisher policy
    if ($removedThumbs.Count -gt 0) {
        Remove-RdpTrustedPublisherThumbprint -Thumbprints @($removedThumbs)
    }
    else {
        Write-Log "No tagged certificates found — nothing to clean up." -Level INFO
    }

    Write-Log "Cleanup complete. Total certs removed: $($removedThumbs.Count)" -Level SUCCESS
}

#endregion

#region --- Main ---

# --- Cleanup mode ---
if ($PSCmdlet.ParameterSetName -eq 'Cleanup') {
    Invoke-Cleanup
    exit 0
}

# --- InstallCerOnly: trust an externally-supplied .cer for current user only ---
if ($PSCmdlet.ParameterSetName -eq 'InstallCer') {
    $exit = 0
    foreach ($p in $Path) {
        if (-not (Test-Path -LiteralPath $p)) {
            Write-Log "Cer file not found: $p" -Level ERROR
            $exit = 1
            continue
        }
        try {
            $cer = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new((Resolve-Path -LiteralPath $p).Path)
            # Ensure imported cer carries our cleanup tag.
            if (-not $cer.FriendlyName -or $cer.FriendlyName -notlike "*$($Script:CertTag)*") {
                $cer.FriendlyName = "$($Script:CertTag) | imported $($cer.Subject)"
            }
            Write-Log "Installing public certificate for current user: $($cer.Subject) [$($cer.Thumbprint)]" -Level INFO
            Add-CertToCurrentUserStore -Cert $cer -StoreName Root
            Add-CertToCurrentUserStore -Cert $cer -StoreName TrustedPublisher
            Add-RdpTrustedPublisherThumbprint -Thumbprint $cer.Thumbprint
        }
        catch {
            Write-Log "Failed to install $p : $_" -Level ERROR
            $exit = 1
        }
    }
    exit $exit
}

# --- Sign mode ---
$files = Resolve-RdpFiles -Inputs $Path
if ($files.Count -eq 0) {
    Write-Log 'No .rdp files found to sign.' -Level ERROR
    exit 1
}
Write-Log "Found $($files.Count) .rdp file(s) to process." -Level INFO

try {
    $cert = Get-OrCreateSigningCert -Subject $Subject -Thumbprint $Thumbprint -ValidYears $ValidYears
}
catch {
    Write-Log "Certificate setup failed: $_" -Level ERROR
    exit 2
}

# Detect whether this is a self-signed cert managed by this script (in which case
# we install it into the per-user Trusted Root + TrustedPublisher stores) vs a
# CA-issued code-signing cert the user already had. For the latter, the cert
# chain is expected to be trusted by an existing Root CA (commercial or
# Enterprise) — installing the leaf into Root would be wrong.
$isManagedCert = $cert.FriendlyName -like "*$($Script:CertTag)*"

if ($isManagedCert) {
    # NOTE: adding to CurrentUser\Root triggers a one-time CryptoAPI "Security
    # Warning" dialog — there is no supported API to suppress it for per-user Root.
    try {
        Add-CertToCurrentUserStore -Cert $cert -StoreName Root
        Add-CertToCurrentUserStore -Cert $cert -StoreName TrustedPublisher
    }
    catch {
        Write-Log "Failed to trust certificate in CurrentUser stores: $_" -Level WARN
        Write-Log 'Signing will continue, but the signed file may show as untrusted on this machine.' -Level WARN
    }
}
else {
    Write-Log "Cert appears to be CA-issued (no '$($Script:CertTag)' tag) — skipping per-user Root/TrustedPublisher install. Trust must come from the CA chain (or GPO/Intune)." -Level INFO
}

# Suppress the "Verify the publisher of this remote connection" dialog by adding
# the thumbprint to the per-user RDP trusted-publishers policy.
try {
    Add-RdpTrustedPublisherThumbprint -Thumbprint $cert.Thumbprint
}
catch {
    Write-Log "Failed to register thumbprint in RDP trusted-publishers policy: $_" -Level WARN
    Write-Log 'Publisher confirmation prompt may still appear when the file is opened.' -Level WARN
}

# Optional public-cert export for distribution
if ($ExportCerPath) {
    try {
        $dir = Split-Path -Parent $ExportCerPath
        if ($dir -and -not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }
        Export-Certificate -Cert $cert -FilePath $ExportCerPath -Force | Out-Null
        Write-Log "Exported public certificate to: $ExportCerPath" -Level SUCCESS
        Write-Log "Distribute this .cer with the .rdp file. Recipients run:" -Level INFO
        Write-Log "  .\Sign-RdpFile.ps1 -InstallCerOnly -Path '<path-to-cer>'" -Level INFO
    }
    catch {
        Write-Log "Failed to export .cer to '$ExportCerPath': $_" -Level WARN
    }
}

$results = New-Object System.Collections.Generic.List[object]
foreach ($f in $files) {
    try {
        $r = Invoke-RdpSign -File $f -Thumbprint $cert.Thumbprint -Force:$Force
        [void]$results.Add($r)
    }
    catch {
        Write-Log "Unexpected failure on $($f.FullName): $_" -Level ERROR
        [void]$results.Add([pscustomobject]@{ File = $f.FullName; Status = 'Failed'; Detail = $_.Exception.Message })
    }
}

$signed  = ($results | Where-Object Status -eq 'Signed').Count
$skipped = ($results | Where-Object Status -eq 'Skipped').Count
$failed  = ($results | Where-Object Status -eq 'Failed').Count

Write-Host ''
Write-Log "Done. Signed: $signed | Skipped: $skipped | Failed: $failed" -Level $(if ($failed) { 'ERROR' } else { 'SUCCESS' })

# Emit results for pipeline / caller consumption
$results

if ($failed -gt 0) { exit 3 } else { exit 0 }

#endregion
