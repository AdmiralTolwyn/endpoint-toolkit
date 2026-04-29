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
    Use a specific certificate from Cert:\CurrentUser\My by thumbprint.
    Accepts either the standard SHA1 cert thumbprint (40 hex chars, what
    Windows shows as "Thumbprint") or the cert's SHA256 fingerprint
    (64 hex chars). Spaces / colons are tolerated.
    Overrides -Subject lookup. Useful if you've already imported a PFX
    or have a code-signing cert from your CA in the user store.

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
    Version : 1.2.0
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
    param([Parameter(Mandatory)] [System.Security.Cryptography.X509Certificates.X509Certificate2] $Cert)
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        return ([System.BitConverter]::ToString($sha256.ComputeHash($Cert.RawData))).Replace('-', '').ToUpperInvariant()
    }
    finally { $sha256.Dispose() }
}

function Find-CertByThumbprint {
    <#
        Looks up a cert in Cert:\CurrentUser\My by either its SHA1 (40 hex chars,
        the value Windows shows as "Thumbprint") or its SHA256 (64 hex chars,
        what some portals / signing services hand out). Spaces are tolerated.
    #>
    param(
        [Parameter(Mandatory)] [string] $Thumbprint,
        [string] $StorePath = 'Cert:\CurrentUser\My'
    )

    $clean = ($Thumbprint -replace '[\s:]', '').ToUpperInvariant()
    if ($clean -notmatch '^[0-9A-F]+$') {
        throw "Thumbprint contains non-hex characters: '$Thumbprint'."
    }

    switch ($clean.Length) {
        40 {
            return Get-ChildItem $StorePath | Where-Object { $_.Thumbprint -eq $clean } | Select-Object -First 1
        }
        64 {
            return Get-ChildItem $StorePath | Where-Object {
                (Get-CertSha256Thumbprint -Cert $_) -eq $clean
            } | Select-Object -First 1
        }
        default {
            throw "Thumbprint must be 40 hex chars (SHA1) or 64 hex chars (SHA256). Got $($clean.Length)."
        }
    }
}

function Get-OrCreateSigningCert {
    param(
        [string] $Subject,
        [string] $Thumbprint,
        [int]    $ValidYears
    )

    $store = 'Cert:\CurrentUser\My'

    if ($Thumbprint) {
        $cert = Find-CertByThumbprint -Thumbprint $Thumbprint -StorePath $store
        if (-not $cert) { throw "No certificate matching thumbprint '$Thumbprint' found in $store (tried SHA1 / SHA256)." }
        if ($cert.EnhancedKeyUsageList.ObjectId -notcontains '1.3.6.1.5.5.7.3.3') {
            throw "Certificate $($cert.Thumbprint) does not have the Code Signing EKU (1.3.6.1.5.5.7.3.3)."
        }
        if (-not $cert.HasPrivateKey) {
            throw "Certificate $($cert.Thumbprint) has no associated private key in $store — cannot sign."
        }
        Write-Log "Using existing certificate by thumbprint: $($cert.Subject) [SHA1=$($cert.Thumbprint)]" -Level INFO
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
        Adds the thumbprint to the per-user RDP trusted-publishers policy, which
        suppresses the "Verify the publisher of this remote connection" dialog
        for files signed with this cert.

        Stored as REG_SZ semicolon-separated, matching the format produced by the
        GPMC policy "Specify SHA1 thumbprints of certificates representing trusted
        .rdp publishers".
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
