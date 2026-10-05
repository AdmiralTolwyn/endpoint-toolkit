<#
.SYNOPSIS
    Read-only connectivity check for Windows 11 Enterprise service endpoints.
.DESCRIPTION
    Tests DNS, direct TCP and HTTP(S) for Windows Update, Delivery Optimization, Microsoft Store,
    Defender, certificates, authentication, activation, settings, diagnostics, NCSI, WNS and
    Edge update endpoints listed at:
    https://learn.microsoft.com/en-us/windows/privacy/manage-windows-11-endpoints

    Wildcard entries are tested through representative hosts (Source = Representative).
    Any HTTP status proves the endpoint was reached, not that the service workflow succeeds.
    WNS is not plain HTTP, so it gets a TLS handshake only.
    ReachableNameMismatch = trusted chain but CDN certificate for another name; not interception.

    Proxy / interception evidence (ProxySignal column):
      - explicit proxy (WinINET manual, PAC, WPAD, or -Proxy)
      - TLS chain root not Microsoft/DigiCert/Baltimore (TLS inspection), checked through the same path
      - proxy headers (Via, Proxy-Agent, X-Squid-*, vendor names in Server/X-* headers) and 407
      - proxy-generated responses: vendor names or block/filter wording in the body, redirects to
        non-Microsoft hosts (e.g. proxy auth portals)
      - content integrity over HTTP: the Microsoft-signed root trust list (authrootstl.cab) must
        verify to a Microsoft root; NCSI must return its exact text
      - public Microsoft names resolving to private addresses
      - path through Zscaler per ip.zscaler.com (skip with -SkipZscalerCheck), and installed
        traffic-steering agents (Zscaler, Netskope, GlobalProtect, Umbrella)
    HTTPS with a Microsoft/DigiCert chain root cannot have been altered or answered by the proxy
    (SSL bypass); only HTTP content can be silently changed, hence the signed-file check.
    No evidence does not prove no proxy: a transparent proxy that passes TLS through untouched
    and strips headers is invisible from the client.

    Run once as the user and once as SYSTEM. Windows services generally use the WinHTTP proxy,
    while this script uses the current account's WinINET settings unless -Proxy is given.

    Result values:
      Reachable              HTTP response from the endpoint (any status), or TLS handshake OK (WNS)
      ReachableNameMismatch  trusted chain, but the CDN certificate is for another name
      UnexpectedContent      NCSI text or signed trust list does not match what Microsoft publishes
      ProxyResponse          the answer came from a proxy (block page or redirect to a non-Microsoft host)
      ProxyAuthRequired      the proxy demanded authentication (407)
      ProxyBlocked           the proxy refused the HTTPS tunnel (CONNECT not 200)
      TlsError               TLS failed for a reason other than a name mismatch
      DnsFailed              name not resolvable and no proxy in use
      Failed                 anything else (timeouts, refused connections)
    HTTP transport failures (reset, EOF, timeout) are retried once when the host is otherwise reachable;
    a success on retry is OK, with the first error shown as a note.

    Console rows are tagged OK, WARN (reachable, but proxy/interception evidence) or FAIL, followed by
    details for every non-OK endpoint and a summary. Use -CsvPath for the full data per endpoint.
.PARAMETER TimeoutSec
    Timeout per TCP connect and per HTTP request.
.PARAMETER Proxy
    Force all probes through this proxy, e.g. the WinHTTP proxy that services use.
.PARAMETER SkipZscalerCheck
    Do not contact ip.zscaler.com.
.PARAMETER CsvPath
    Optional path to export the results.
.OUTPUTS
    Exit code 0 = all reachable, no proxy evidence; 1 = at least one endpoint failed;
    2 = all reachable but proxy/interception evidence found.
.EXAMPLE
    .\Test-WindowsServiceEndpoints.ps1 -CsvPath .\endpoints.csv
.EXAMPLE
    .\Test-WindowsServiceEndpoints.ps1 -Proxy http://proxy.contoso.com:8080
#>
# Must follow the help block: a leading #Requires stops Get-Help from finding the script help.
#Requires -Version 5.1
[CmdletBinding()]
param(
    [ValidateRange(1, 120)]
    [int]$TimeoutSec = 10,

    [ValidateScript({ $_.IsAbsoluteUri })]
    [Uri]$Proxy,

    [switch]$SkipZscalerCheck,

    [string]$CsvPath
)

$ErrorActionPreference = 'Stop'
# SystemDefault (0) lets Windows choose TLS 1.2/1.3; legacy .NET Framework defaults (Ssl3|Tls) need TLS 1.2 added.
if ([int][Net.ServicePointManager]::SecurityProtocol -ne 0) {
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
}
Add-Type -AssemblyName System.Security

$endpoints = @'
Group,Host,Scheme,Path,Source,Check
Windows Update,sls.update.microsoft.com,https,/,Representative
Windows Update,fe2cr.update.microsoft.com,https,/,Representative
Windows Update,fe3cr.delivery.mp.microsoft.com,https,/,Representative
Windows Update,tlu.dl.delivery.mp.microsoft.com,http,/,Representative
Windows Update,download.windowsupdate.com,http,/,Representative
Windows Update,tsfe.trafficshaping.dsp.mp.microsoft.com,https,/,Documented
Windows Update,adl.windows.com,https,/,Documented
Delivery Optimization,geo.prod.do.dsp.mp.microsoft.com,https,/,Representative
Delivery Optimization,kv801.prod.do.dsp.mp.microsoft.com,https,/,Representative
Delivery Optimization,cp801.prod.do.dsp.mp.microsoft.com,https,/,Representative
Delivery Optimization,disc801.prod.do.dsp.mp.microsoft.com,https,/,Representative
Microsoft Store,displaycatalog.mp.microsoft.com,https,/,Representative
Microsoft Store,storeedgefd.dsx.mp.microsoft.com,http,/,Documented
Microsoft Store,livetileedge.dsx.mp.microsoft.com,https,/,Documented
Microsoft Store,storecatalogrevocation.storequality.microsoft.com,https,/,Documented
Defender,wdcp.microsoft.com,https,/,Documented
Defender,definitionupdates.microsoft.com,https,/,Documented
Defender,checkappexec.microsoft.com,https,/,Documented
Defender,ping-edge.smartscreen.microsoft.com,https,/,Documented
Defender,nav-edge.smartscreen.microsoft.com,https,/,Documented
Defender,data-edge.smartscreen.microsoft.com,http,/,Documented
Certificates,ctldl.windowsupdate.com,http,/msdownload/update/v3/static/trustedr/en/authrootstl.cab,Documented,ctl
Certificates,ocsp.digicert.com,http,/,Documented
Authentication,login.live.com,https,/,Documented
Activation,licensing.mp.microsoft.com,https,/,Documented
Settings,settings-win.data.microsoft.com,https,/,Documented
Settings,settings.data.microsoft.com,https,/,Documented
Diagnostics,v10.events.data.microsoft.com,https,/,Documented
Diagnostics,self.events.data.microsoft.com,https,/,Documented
Diagnostics,functional.events.data.microsoft.com,https,/,Documented
Diagnostics,watson.telemetry.microsoft.com,https,/,Representative
NCSI,www.msftconnecttest.com,http,/connecttest.txt,Documented,ncsi
Push Notifications,client.wns.windows.com,tls,/,Representative
Edge Update,msedge.api.cdp.microsoft.com,https,/,Documented
'@ | ConvertFrom-Csv

$trustedRootPattern = 'O=(Microsoft Corporation|DigiCert Inc|Baltimore)'
$proxyHeaderPattern = '^(Via|Proxy-Agent|Proxy-Connection|X-BlueCoat-Via|X-Squid-Error|X-Cache-Lookup|X-Zscaler.*|X-Forcepoint.*|X-Netskope.*|X-Iboss.*|X-Proxy.*)$'
$proxyVendorPattern = 'zscaler|blue ?coat|proxysg|squid|forcepoint|websense|web gateway|skyhigh|netskope|iboss|palo ?alto|fortigate|fortinet|check ?point|sophos|barracuda|menlo|ironport|umbrella|trend ?micro'
# Microsoft download CDNs add their own Via (observed: '1.1 varnish' on tlu.dl.delivery.mp.microsoft.com).
$cdnViaPattern = 'varnish|akamai|cloudfront|fastly|edgecast|azurefd|msedge'
$proxyPagePattern = '\b(zscaler|bluecoat|blue coat|proxysg|forcepoint|websense|skyhigh security|netskope|iboss|palo alto networks|fortiguard|fortigate|check point software|sophos|barracuda|menlo security|ironport|cisco umbrella|trend micro)\b|blocked by (your )?(organi[sz]ation|administrator|company|security policy)|access to this (site|page|resource|website) (is|has been) (denied|blocked)|acceptable use policy|captive portal'
$microsoftDomainPattern = '(^|\.)(microsoft\.com|windowsupdate\.com|windows\.com|windows\.net|live\.com|live\.net|msftconnecttest\.com|digicert\.com|office\.com|office\.net|microsoftonline\.com|msn\.com|bing\.com|msedge\.net|azureedge\.net|azurefd\.net|akamaized\.net|akamai\.net|msauth\.net|skype\.com|xboxlive\.com)$'
$steeringAgents = [ordered]@{
    ZSATunnel         = 'Zscaler Client Connector'
    stAgentSvc        = 'Netskope Client'
    PanGPS            = 'Palo Alto GlobalProtect'
    csc_umbrellaagent = 'Cisco Umbrella (Secure Client)'
    Umbrella_RC       = 'Cisco Umbrella Roaming Client'
}

function Get-ProxyConfiguration {
    <#
    .SYNOPSIS
        Reads the WinINET and WinHTTP proxy configuration of the current account.
    .DESCRIPTION
        Uses machine-wide WinINET settings when the ProxySettingsPerUser=0 policy is set, decodes the
        DefaultConnectionSettings flags (manual proxy, PAC URL, auto-detect) and the WinHttpSettings
        blob that 'netsh winhttp show proxy' displays. Configuration only: whether WPAD/PAC actually
        returns a proxy is determined per URL by Get-EffectiveProxy.
    .OUTPUTS
        PSCustomObject with Scope, Manual ('None' or proxy list), Pac ('None' or URL), AutoDetect (bool)
        and WinHttp ('Direct' or 'proxy (bypass: list)').
    #>
    $settingsKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings'
    $scope = 'Per user'
    $policy = Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\Internet Settings' -ErrorAction SilentlyContinue
    if ($policy -and $policy.ProxySettingsPerUser -eq 0) {
        $settingsKey = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings'
        $scope = 'Per machine (policy)'
    }
    $inet = Get-ItemProperty $settingsKey -ErrorAction SilentlyContinue
    $connections = Get-ItemProperty "$settingsKey\Connections" -ErrorAction SilentlyContinue

    # DefaultConnectionSettings byte 8: 0x02 manual proxy, 0x04 PAC URL, 0x08 auto-detect (WPAD).
    $flags = 0
    if ($connections -and $connections.DefaultConnectionSettings -and $connections.DefaultConnectionSettings.Length -gt 8) {
        $flags = $connections.DefaultConnectionSettings[8]
    }
    elseif ($inet -and $inet.ProxyEnable -eq 1) { $flags = 0x02 }

    $manual = 'None'
    if (($flags -band 0x02) -and $inet.ProxyServer) { $manual = $inet.ProxyServer }
    $pac = 'None'
    if (($flags -band 0x04) -and $inet.AutoConfigURL) { $pac = $inet.AutoConfigURL }

    # WinHttpSettings: DWORD size, DWORD counter, DWORD flags, DWORD proxy length, proxy, DWORD bypass length, bypass.
    $winHttp = 'Direct'
    $bytes = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Connections' -ErrorAction SilentlyContinue).WinHttpSettings
    if ($bytes -and $bytes.Length -ge 16) {
        $proxyLength = [BitConverter]::ToInt32($bytes, 12)
        if ($proxyLength -gt 0 -and $bytes.Length -ge (16 + $proxyLength)) {
            $winHttp = [Text.Encoding]::ASCII.GetString($bytes, 16, $proxyLength)
            $bypassOffset = 16 + $proxyLength
            if ($bytes.Length -ge ($bypassOffset + 4)) {
                $bypassLength = [BitConverter]::ToInt32($bytes, $bypassOffset)
                if ($bypassLength -gt 0 -and $bytes.Length -ge ($bypassOffset + 4 + $bypassLength)) {
                    $winHttp += ' (bypass: {0})' -f [Text.Encoding]::ASCII.GetString($bytes, $bypassOffset + 4, $bypassLength)
                }
            }
        }
    }

    [pscustomobject]@{
        Scope      = $scope
        Manual     = $manual
        Pac        = $pac
        AutoDetect = [bool]($flags -band 0x08)
        WinHttp    = $winHttp
    }
}

function Get-EffectiveProxy {
    <#
    .SYNOPSIS
        Returns the proxy that will be used for a URL, or $null for a direct connection.
    .DESCRIPTION
        Returns the script's -Proxy parameter when given. Otherwise asks the system proxy of the current
        account, which evaluates manual settings, bypass lists, PAC and WPAD. The resolver is created once
        and cached in $script:SystemProxy so WPAD discovery is not repeated for every endpoint.
    .PARAMETER Uri
        Target URL to resolve the proxy for.
    .OUTPUTS
        System.Uri of the proxy, or $null when the URL is reached directly.
    #>
    param([Uri]$Uri)
    if ($Proxy) { return $Proxy }
    if (-not $script:SystemProxy) { $script:SystemProxy = [Net.WebRequest]::GetSystemWebProxy() }
    $resolved = $script:SystemProxy.GetProxy($Uri)
    if (-not $resolved -or $resolved.AbsoluteUri -eq $Uri.AbsoluteUri) { return $null }
    return $resolved
}

function Get-ProxyHeaderSignal {
    <#
    .SYNOPSIS
        Returns the response headers that indicate a proxy in the path.
    .DESCRIPTION
        A header is returned when a Server, Via, Proxy-* or X-* header names a known proxy vendor, or when
        its name is proxy-specific (Via, Proxy-Agent, X-Squid-Error, X-Zscaler-*, ...). Via headers added
        by CDNs (varnish, Akamai, ...) are ignored because Microsoft download endpoints send them.
    .PARAMETER Headers
        Headers formatted as 'Name: value', from the HTTP response and/or the proxy CONNECT reply.
    .OUTPUTS
        System.String for each suspicious header, unchanged.
    #>
    param([string[]]$Headers)
    foreach ($header in $Headers) {
        if (-not $header) { continue }
        $name = ($header -split ':', 2)[0].Trim()
        if ($header -match $proxyVendorPattern -and $name -match '^(Server|Via|Proxy-.*|X-.*)$') { $header }
        elseif ($name -eq 'Via' -and $header -match $cdnViaPattern) { continue }
        elseif ($name -match $proxyHeaderPattern) { $header }
    }
}

function Test-PrivateAddress {
    <#
    .SYNOPSIS
        Tests whether an IP address is private, loopback, link-local or unspecified.
    .DESCRIPTION
        Public Microsoft names resolving to such addresses indicate DNS redirection to a proxy or sinkhole.
        Covers IPv4 10/8, 127/8, 0/8, 169.254/16, 172.16/12, 192.168/16 and IPv6 ::1, fc00::/7, fe80::/10.
    .PARAMETER Address
        IP address in string form.
    .OUTPUTS
        System.Boolean
    #>
    param([string]$Address)
    return $Address -match '^(10\.|127\.|0\.|169\.254\.|192\.168\.|172\.(1[6-9]|2\d|3[01])\.|::1$|f[cd][0-9a-f]{0,2}:|fe80:)'
}

function Test-Dns {
    <#
    .SYNOPSIS
        Resolves a host name with the Windows resolver.
    .DESCRIPTION
        Uses the same resolver as Windows services (hosts file, DNS client cache, NRPT). A failure is not
        fatal when an explicit proxy is used, because the proxy resolves the name instead.
    .PARAMETER HostName
        Name to resolve.
    .OUTPUTS
        Hashtable: Ok (bool), Addresses (string[]), Detail (addresses joined, or the error message).
    #>
    param([string]$HostName)
    try {
        $addresses = @([Net.Dns]::GetHostAddresses($HostName) | ForEach-Object { $_.IPAddressToString })
        return @{ Ok = $true; Addresses = $addresses; Detail = ($addresses -join ' ') }
    }
    catch {
        return @{ Ok = $false; Addresses = @(); Detail = $_.Exception.GetBaseException().Message }
    }
}

function Test-Tcp {
    <#
    .SYNOPSIS
        Tests a direct TCP connection to a host and port, bypassing any proxy.
    .DESCRIPTION
        Shows whether the firewall allows direct egress. A timeout here is expected on networks that only
        allow traffic through an explicit proxy.
    .PARAMETER HostName
        Target host.
    .PARAMETER Port
        Target TCP port.
    .PARAMETER TimeoutMs
        Connect timeout in milliseconds.
    .OUTPUTS
        System.String: 'Open', 'Timeout' or the socket error message.
    #>
    param([string]$HostName, [int]$Port, [int]$TimeoutMs)
    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $async = $client.BeginConnect($HostName, $Port, $null, $null)
        if (-not $async.AsyncWaitHandle.WaitOne($TimeoutMs)) { return 'Timeout' }
        $client.EndConnect($async)
        return 'Open'
    }
    catch {
        return $_.Exception.GetBaseException().Message
    }
    finally {
        $client.Close()
    }
}

function Test-Tls {
    <#
    .SYNOPSIS
        Performs a TLS handshake with a host on port 443 and captures the certificate chain.
    .DESCRIPTION
        When a proxy is given, opens an HTTP CONNECT tunnel first, so the certificate is the one the proxy
        path delivers. A chain root outside Microsoft/DigiCert then indicates TLS inspection. The validation
        callback always accepts, so the certificate is captured even when invalid; trust is reported via
        PolicyErrors. The callback writes to $script:TlsPolicyErrors and $script:TlsRoot because it runs as
        a delegate outside this function's scope. CONNECT is sent without proxy credentials, so a proxy that
        requires Windows authentication answers 407 here even when Test-Http succeeds. The handshake offers
        the same protocols as Test-Http (ServicePointManager.SecurityProtocol); the parameterless
        AuthenticateAsClient overload falls back to TLS 1.0 on .NET Framework without
        SystemDefaultTlsVersions, which Microsoft endpoints reject.
    .PARAMETER HostName
        Host to connect to; also used for SNI and certificate name validation.
    .PARAMETER ProxyUri
        Optional HTTP proxy to tunnel through.
    .PARAMETER TimeoutMs
        Connect, send and receive timeout in milliseconds.
    .OUTPUTS
        Hashtable: Ok (trusted and name matches), Subject, Issuer, Root (chain root subject), Protocol
        (negotiated TLS version), PolicyErrors, Error, ProxyStatus (CONNECT status code) and ProxyHeaders
        (CONNECT reply headers).
    #>
    param([string]$HostName, [Uri]$ProxyUri, [int]$TimeoutMs)
    $result = @{ Ok = $false; Subject = $null; Issuer = $null; Root = $null; Protocol = $null; PolicyErrors = $null; Error = $null; ProxyStatus = $null; ProxyHeaders = @() }
    $script:TlsPolicyErrors = $null
    $script:TlsRoot = $null
    $callback = [Net.Security.RemoteCertificateValidationCallback] {
        param($s, $c, $ch, $e)
        $script:TlsPolicyErrors = $e
        if ($ch -and $ch.ChainElements.Count -gt 0) { $script:TlsRoot = $ch.ChainElements[$ch.ChainElements.Count - 1].Certificate.Subject }
        $true
    }
    $connectHost = $HostName
    $connectPort = 443
    if ($ProxyUri) {
        $connectHost = $ProxyUri.Host
        $connectPort = $ProxyUri.Port
    }
    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $async = $client.BeginConnect($connectHost, $connectPort, $null, $null)
        if (-not $async.AsyncWaitHandle.WaitOne($TimeoutMs)) { $result.Error = 'TCP timeout'; return $result }
        $client.EndConnect($async)
        $client.ReceiveTimeout = $TimeoutMs
        $client.SendTimeout = $TimeoutMs
        $stream = $client.GetStream()

        if ($ProxyUri) {
            $connect = [Text.Encoding]::ASCII.GetBytes("CONNECT ${HostName}:443 HTTP/1.1`r`nHost: ${HostName}:443`r`n`r`n")
            $stream.Write($connect, 0, $connect.Length)
            $reply = New-Object System.Text.StringBuilder
            while ($reply.Length -lt 16384) {
                $byte = $stream.ReadByte()
                if ($byte -lt 0) { break }
                [void]$reply.Append([char]$byte)
                if ($reply.Length -ge 4 -and $reply.ToString($reply.Length - 4, 4) -eq "`r`n`r`n") { break }
            }
            $lines = @($reply.ToString() -split "`r`n" | Where-Object { $_ })
            if ($lines.Count -gt 0 -and $lines[0] -match '^HTTP/\d\.\d (\d{3})') { $result.ProxyStatus = [int]$Matches[1] }
            $result.ProxyHeaders = @($lines | Select-Object -Skip 1)
            if ($result.ProxyStatus -ne 200) {
                $result.Error = 'Proxy CONNECT returned: {0}' -f ($lines | Select-Object -First 1)
                return $result
            }
        }

        $ssl = New-Object System.Net.Security.SslStream($stream, $false, $callback)
        $ssl.AuthenticateAsClient($HostName, $null, [Security.Authentication.SslProtocols][int][Net.ServicePointManager]::SecurityProtocol, $false)
        $result.Protocol = [string]$ssl.SslProtocol
        $result.Subject = $ssl.RemoteCertificate.Subject
        $result.Issuer = $ssl.RemoteCertificate.Issuer
        $result.Root = $script:TlsRoot
        $result.PolicyErrors = [string]$script:TlsPolicyErrors
        $result.Ok = $result.PolicyErrors -eq 'None'
    }
    catch {
        $result.Error = $_.Exception.GetBaseException().Message
    }
    finally {
        $client.Close()
    }
    return $result
}

function Test-Http {
    <#
    .SYNOPSIS
        Sends one HTTP GET and captures status, headers, redirect target and the start of the body.
    .DESCRIPTION
        Redirects are not followed, so a proxy redirecting to its authentication portal remains visible.
        Error statuses (4xx/5xx) are returned as statuses, not errors. The proxy uses the current account's
        default credentials, matching how Windows authenticates to NTLM/Kerberos proxies.
    .PARAMETER Uri
        URL to request.
    .PARAMETER ProxyUri
        Proxy to use; $null forces a direct connection.
    .PARAMETER TimeoutMs
        Request and read timeout in milliseconds.
    .PARAMETER MaxBytes
        Maximum number of body bytes to read.
    .OUTPUTS
        Hashtable: Status (int or $null), Error, Body (UTF-8 text), BodyBytes, Location and Headers
        ('Name: value' strings).
    #>
    param([Uri]$Uri, [Uri]$ProxyUri, [int]$TimeoutMs, [int]$MaxBytes = 65536)
    $result = @{ Status = $null; Error = $null; Body = $null; BodyBytes = $null; Location = $null; Headers = @() }
    $request = [Net.HttpWebRequest]::Create($Uri)
    $request.Method = 'GET'
    $request.AllowAutoRedirect = $false
    $request.Timeout = $TimeoutMs
    $request.ReadWriteTimeout = $TimeoutMs
    $request.UserAgent = 'Windows-Endpoint-Check'
    if ($ProxyUri) {
        $webProxy = New-Object System.Net.WebProxy($ProxyUri)
        $webProxy.UseDefaultCredentials = $true
        $request.Proxy = $webProxy
    }
    else {
        $request.Proxy = $null
    }

    $response = $null
    try {
        $response = $request.GetResponse()
    }
    catch [Net.WebException] {
        $response = $_.Exception.Response
        if (-not $response) {
            $result.Error = '{0}: {1}' -f $_.Exception.Status, $_.Exception.GetBaseException().Message
        }
    }
    catch {
        $result.Error = $_.Exception.GetBaseException().Message
    }

    try {
        if ($response) {
            $result.Status = [int]$response.StatusCode
            $result.Headers = @($response.Headers.AllKeys | ForEach-Object { '{0}: {1}' -f $_, $response.Headers[$_] })
            $result.Location = $response.Headers['Location']
            $buffer = New-Object byte[] 8192
            $memory = New-Object System.IO.MemoryStream
            $body = $response.GetResponseStream()
            try {
                while ($memory.Length -lt $MaxBytes) {
                    $read = $body.Read($buffer, 0, [int][Math]::Min($buffer.Length, $MaxBytes - $memory.Length))
                    if ($read -le 0) { break }
                    $memory.Write($buffer, 0, $read)
                }
            }
            catch {
                $result.Error = 'Body read: {0}' -f $_.Exception.GetBaseException().Message
            }
            finally {
                $body.Close()
            }
            $result.BodyBytes = $memory.ToArray()
            $result.Body = [Text.Encoding]::UTF8.GetString($result.BodyBytes)
        }
    }
    finally {
        if ($response) { $response.Close() }
    }
    return $result
}

function Test-SignedCtl {
    <#
    .SYNOPSIS
        Verifies that downloaded authrootstl.cab content is the genuine Microsoft-signed root trust list.
    .DESCRIPTION
        The file is served over plain HTTP, so a proxy can replace it without detection by TLS. The CAB
        wrapper is unsigned; the authroot.stl inside is a PKCS#7 certificate trust list signed by
        'Microsoft Certificate Trust List Publisher'. Extracts it with expand.exe into a temporary folder
        (always removed), verifies the signature and requires a Microsoft signer and a Microsoft chain root.
        Any change to the signed content breaks the hash; a changed wrapper fails extraction.
    .PARAMETER Bytes
        Downloaded CAB content.
    .OUTPUTS
        $null when genuine; otherwise a System.String describing why verification failed.
    #>
    param([byte[]]$Bytes)
    # expand.exe writes to stderr on bad input; must not become a terminating error.
    $ErrorActionPreference = 'Continue'
    $work = Join-Path $env:TEMP ('ctlcheck-' + [guid]::NewGuid().ToString('N'))
    [void](New-Item -Path $work -ItemType Directory)
    try {
        $cab = Join-Path $work 'authrootstl.cab'
        $extract = Join-Path $work 'x'
        [void](New-Item -Path $extract -ItemType Directory)
        [IO.File]::WriteAllBytes($cab, $Bytes)
        $null = & expand.exe $cab '-F:*' $extract 2>&1
        $stl = Get-ChildItem -Path $extract -File | Select-Object -First 1
        if (-not $stl) { return 'CAB could not be extracted' }

        $cms = New-Object System.Security.Cryptography.Pkcs.SignedCms
        $cms.Decode([IO.File]::ReadAllBytes($stl.FullName))
        $cms.CheckSignature($true)
        $signer = $cms.SignerInfos[0].Certificate
        $chain = New-Object System.Security.Cryptography.X509Certificates.X509Chain
        $chain.ChainPolicy.RevocationMode = 'NoCheck'
        [void]$chain.ChainPolicy.ExtraStore.AddRange($cms.Certificates)
        if (-not $chain.Build($signer)) { return 'Signer chain not trusted: {0}' -f (($chain.ChainStatus | ForEach-Object { $_.Status }) -join ',') }
        $root = $chain.ChainElements[$chain.ChainElements.Count - 1].Certificate.Subject
        if ($signer.Subject -notmatch 'O=Microsoft Corporation' -or $root -notmatch 'O=Microsoft Corporation') {
            return 'Not Microsoft-signed: signer {0}, root {1}' -f $signer.Subject, $root
        }
        return $null
    }
    catch {
        return 'Signature check failed: {0}' -f $_.Exception.GetBaseException().Message.Trim()
    }
    finally {
        Remove-Item -Path $work -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Get-RowStatus {
    <#
    .SYNOPSIS
        Maps a result row to the console tag OK, WARN or FAIL.
    .DESCRIPTION
        FAIL for any result other than Reachable/ReachableNameMismatch; WARN when the endpoint was reached
        but proxy or interception evidence exists; otherwise OK. Drives row colours, the summary and the verdict.
    .PARAMETER Row
        Result object produced for one endpoint.
    .OUTPUTS
        System.String: 'OK', 'WARN' or 'FAIL'.
    #>
    param([psobject]$Row)
    if ($Row.Result -notin 'Reachable', 'ReachableNameMismatch') { return 'FAIL' }
    if ($Row.ProxySignal) { return 'WARN' }
    return 'OK'
}

function Get-CommonName {
    <#
    .SYNOPSIS
        Returns the CN part of a certificate subject for compact display.
    .PARAMETER Subject
        Distinguished name such as 'CN=DigiCert Global Root G2, OU=www.digicert.com, O=DigiCert Inc, C=US'.
    .OUTPUTS
        System.String: the CN value, or the full subject when it has no CN.
    #>
    param([string]$Subject)
    if ($Subject -match 'CN=([^,]+)') { return $Matches[1] }
    return $Subject
}

function Write-Section {
    <#
    .SYNOPSIS
        Writes a section heading with an underline to the console.
    .PARAMETER Title
        Heading text.
    .OUTPUTS
        None. Writes to the host only.
    #>
    param([string]$Title)
    Write-Host ''
    Write-Host $Title -ForegroundColor Cyan
    Write-Host ('-' * $Title.Length) -ForegroundColor DarkCyan
}

function Write-KeyValue {
    <#
    .SYNOPSIS
        Writes an aligned 'key   value' line to the console.
    .DESCRIPTION
        The key is written dim and padded to a fixed width; an empty key produces a continuation line.
        The value uses the console's default colour unless -Color is given.
    .PARAMETER Key
        Label text.
    .PARAMETER Value
        Value text.
    .PARAMETER Color
        Optional colour for the value.
    .OUTPUTS
        None. Writes to the host only.
    #>
    param([string]$Key, [string]$Value, [ConsoleColor]$Color)
    Write-Host ('  {0,-17}' -f $Key) -NoNewline -ForegroundColor DarkGray
    if ($PSBoundParameters.ContainsKey('Color')) { Write-Host $Value -ForegroundColor $Color }
    else { Write-Host $Value }
}

function Write-EndpointRow {
    <#
    .SYNOPSIS
        Writes one colour-coded result line for an endpoint.
    .DESCRIPTION
        Format: [ OK ] host  response  note. The response column shows the HTTP status, or 'TLS ok' for a
        successful handshake without HTTP. The note gives the failure or the most specific proxy signal and
        is truncated to the console width; the Details section repeats it in full.
    .PARAMETER Row
        Result object produced for one endpoint.
    .PARAMETER Signals
        Proxy/interception signals for the endpoint.
    .PARAMETER HostWidth
        Column width for the host name.
    .PARAMETER Width
        Console width in characters.
    .OUTPUTS
        None. Writes to the host only.
    #>
    param([psobject]$Row, [string[]]$Signals, [int]$HostWidth, [int]$Width)
    $status = Get-RowStatus -Row $Row
    $tags = @{ OK = ' OK '; WARN = 'WARN'; FAIL = 'FAIL' }
    $colors = @{ OK = 'Green'; WARN = 'Yellow'; FAIL = 'Red' }

    $response = '--'
    if ($Row.HttpStatus) { $response = 'HTTP {0}' -f $Row.HttpStatus }
    elseif ($Row.Result -match '^Reachable') { $response = 'TLS ok' }

    $specific = @($Signals | Where-Object { $_ -notmatch '^Explicit proxy' })
    $reason = $null
    if ($specific.Count -gt 0) { $reason = $specific[0] }
    elseif ($Row.Detail -and $status -eq 'FAIL') { $reason = $Row.Detail }
    elseif ($Signals.Count -gt 0) { $reason = $Signals[0] }

    $note = ''
    if ($status -eq 'FAIL') { $note = $Row.Result }
    if ($reason) { $note = (@($note, $reason) | Where-Object { $_ }) -join ': ' }
    if ($status -eq 'OK' -and $Row.Result -eq 'ReachableNameMismatch') { $note = 'name mismatch (CDN cert)' }
    elseif ($status -eq 'OK' -and $Row.Detail) { $note = $Row.Detail }
    if ($specific.Count -gt 1) { $note += ' (+{0} more)' -f ($specific.Count - 1) }

    # Narrow consoles wrap rather than cut notes to nothing; Details repeats them in full.
    $room = [Math]::Max(40, $Width - $HostWidth - 22)
    if ($note.Length -gt $room) { $note = $note.Substring(0, $room - 3) + '...' }
    $noteColor = 'DarkGray'
    if ($status -ne 'OK') { $noteColor = $colors[$status] }

    Write-Host '  [' -NoNewline -ForegroundColor DarkGray
    Write-Host $tags[$status] -NoNewline -ForegroundColor $colors[$status]
    Write-Host '] ' -NoNewline -ForegroundColor DarkGray
    Write-Host ($Row.Host.PadRight($HostWidth) + '  ') -NoNewline
    Write-Host $response.PadRight(9) -NoNewline -ForegroundColor DarkGray
    Write-Host (' ' + $note) -ForegroundColor $noteColor
}

$timeoutMs = $TimeoutSec * 1000
$identity = [Security.Principal.WindowsIdentity]::GetCurrent().Name
$config = Get-ProxyConfiguration
$started = Get-Date

$consoleWidth = 120
try { if ($Host.UI.RawUI.WindowSize.Width -ge 60) { $consoleWidth = $Host.UI.RawUI.WindowSize.Width } } catch { }
$hostWidth = ($endpoints | ForEach-Object { $_.Host.Length } | Measure-Object -Maximum).Maximum

$agents = @(foreach ($name in $steeringAgents.Keys) {
    $service = Get-Service -Name $name -ErrorAction SilentlyContinue
    if ($service) { '{0} ({1}: {2})' -f $steeringAgents[$name], $name, $service.Status }
})

$title = 'Windows service endpoint check'
Write-Host ''
Write-Host $title -ForegroundColor Cyan
Write-Host ('=' * $title.Length) -ForegroundColor DarkCyan
Write-KeyValue 'Account' ('{0}  (PowerShell {1})' -f $identity, $PSVersionTable.PSVersion)
Write-KeyValue 'Started' $started.ToString('yyyy-MM-dd HH:mm:ss')
Write-KeyValue 'WinINET' ('{0}; manual proxy {1}; PAC {2}; auto-detect {3}' -f $config.Scope, $config.Manual, $config.Pac, $config.AutoDetect)
Write-KeyValue 'WinHTTP' $config.WinHttp
if ($Proxy) { Write-KeyValue 'Forced proxy' $Proxy.AbsoluteUri -Color Yellow }
if ($agents.Count -gt 0) { Write-KeyValue 'Steering agents' ($agents -join '; ') -Color Yellow }
else { Write-KeyValue 'Steering agents' 'None found' }

$results = New-Object System.Collections.Generic.List[object]
$signalsByHost = @{}
$currentGroup = $null
$index = 0
foreach ($endpoint in $endpoints) {
    $index++
    Write-Progress -Activity 'Testing Windows service endpoints' -Status $endpoint.Host -PercentComplete ($index * 100 / $endpoints.Count)
    if ($endpoint.Group -ne $currentGroup) {
        if ($null -eq $currentGroup) { Write-Section 'Endpoints' }
        $currentGroup = $endpoint.Group
        Write-Host ''
        Write-Host ('  ' + $currentGroup) -ForegroundColor White
    }

    $port = 443
    if ($endpoint.Scheme -eq 'http') { $port = 80 }
    $isNcsi = $endpoint.Check -eq 'ncsi'
    $isCtl = $endpoint.Check -eq 'ctl'
    $isTlsOnly = $endpoint.Scheme -eq 'tls'
    if ($isTlsOnly) { $uri = [Uri]('https://{0}{1}' -f $endpoint.Host, $endpoint.Path) }
    else { $uri = [Uri]('{0}://{1}{2}' -f $endpoint.Scheme, $endpoint.Host, $endpoint.Path) }

    $proxyUri = Get-EffectiveProxy -Uri $uri
    $proxyLabel = 'Direct'
    if ($proxyUri) { $proxyLabel = $proxyUri.AbsoluteUri }

    $dns = Test-Dns -HostName $endpoint.Host
    $tcp = 'Skipped (DNS failed)'
    if ($dns.Ok) { $tcp = Test-Tcp -HostName $endpoint.Host -Port $port -TimeoutMs $timeoutMs }

    if ($isTlsOnly) {
        $http = @{ Status = $null; Error = $null; Body = $null; BodyBytes = $null; Location = $null; Headers = @() }
    }
    else {
        $maxBytes = 65536
        if ($isCtl) { $maxBytes = 4MB }
        $http = Test-Http -Uri $uri -ProxyUri $proxyUri -TimeoutMs $timeoutMs -MaxBytes $maxBytes
    }

    # .NET Framework reports TrustFailure; PowerShell 7 (.NET) reports a rejected validation callback.
    $tlsErrorPattern = 'TrustFailure|SecureChannelFailure|RemoteCertificateValidationCallback'
    $retryNote = $null
    if (-not $http.Status -and $http.Error -and $http.Error -notmatch $tlsErrorPattern -and ($tcp -eq 'Open' -or $proxyUri)) {
        $firstError = $http.Error
        $http = Test-Http -Uri $uri -ProxyUri $proxyUri -TimeoutMs $timeoutMs -MaxBytes $maxBytes
        if ($http.Status) { $retryNote = 'Succeeded on retry; first attempt: {0}' -f $firstError }
    }

    $tls = $null
    if ($uri.Scheme -eq 'https' -and ($dns.Ok -or $proxyUri)) {
        $tls = Test-Tls -HostName $endpoint.Host -ProxyUri $proxyUri -TimeoutMs $timeoutMs
    }

    $signals = New-Object System.Collections.Generic.List[string]
    if ($proxyUri) { $signals.Add("Explicit proxy $proxyLabel") }
    if ($http.Status -eq 407 -or ($tls -and $tls.ProxyStatus -eq 407)) { $signals.Add('Proxy authentication required (407)') }
    if ($tls -and $tls.ProxyStatus -and $tls.ProxyStatus -notin 200, 407) { $signals.Add("Proxy refused CONNECT ($($tls.ProxyStatus))") }
    $headers = @($http.Headers)
    if ($tls) { $headers += $tls.ProxyHeaders }
    foreach ($header in (Get-ProxyHeaderSignal -Headers $headers | Select-Object -Unique)) { $signals.Add("Header $header") }
    if ($tls -and $tls.Root -and $tls.Root -notmatch $trustedRootPattern) { $signals.Add("TLS inspection: chain root '$(Get-CommonName -Subject $tls.Root)'") }
    $privateAddresses = @($dns.Addresses | Where-Object { Test-PrivateAddress $_ })
    if ($privateAddresses.Count -gt 0) { $signals.Add("DNS resolves to private address $($privateAddresses -join ' ')") }

    $contentAltered = $false
    if ($isNcsi -and $http.Status -and $http.Body -ne 'Microsoft Connect Test') {
        $contentAltered = $true
        $signals.Add('NCSI content altered')
    }
    if ($isCtl -and $http.Status) {
        $ctlError = 'Expected HTTP 200 for signed file, got {0}' -f $http.Status
        if ($http.Status -eq 200) { $ctlError = Test-SignedCtl -Bytes $http.BodyBytes }
        if ($ctlError) {
            $contentAltered = $true
            $signals.Add("Signed content not genuine: $ctlError")
        }
    }
    $proxyAnswered = $false
    $target = $null
    if ($http.Location -and [Uri]::TryCreate($uri, $http.Location, [ref]$target) -and $target.Host -notmatch $microsoftDomainPattern) {
        $proxyAnswered = $true
        $signals.Add("Redirect to non-Microsoft host $($target.Host)")
    }
    if (-not $isCtl -and $http.Body -and $http.Body -match $proxyPagePattern) {
        $proxyAnswered = $true
        $signals.Add("Proxy-generated content ('$($Matches[0])' in body)")
    }

    $httpTlsError = $http.Error -match $tlsErrorPattern

    # Any HTTP status wins: explicit-proxy clients may legitimately fail local DNS and direct TCP.
    if ($http.Status -eq 407 -or ($isTlsOnly -and $tls -and $tls.ProxyStatus -eq 407)) { $outcome = 'ProxyAuthRequired' }
    # A refused tunnel means any HTTPS status came from the proxy, not from Microsoft.
    elseif ($tls -and $tls.ProxyStatus -and $tls.ProxyStatus -notin 200, 407) { $outcome = 'ProxyBlocked' }
    elseif ($http.Status) {
        $outcome = 'Reachable'
        if ($contentAltered) { $outcome = 'UnexpectedContent' }
        elseif ($proxyAnswered) { $outcome = 'ProxyResponse' }
    }
    elseif ($isTlsOnly -and $tls -and $tls.Ok) { $outcome = 'Reachable' }
    elseif ($tls -and $tls.PolicyErrors -eq 'RemoteCertificateNameMismatch') { $outcome = 'ReachableNameMismatch' }
    elseif (($tls -and $tls.Subject -and -not $tls.Ok) -or $httpTlsError) { $outcome = 'TlsError' }
    elseif (-not $dns.Ok -and -not $proxyUri) { $outcome = 'DnsFailed' }
    else { $outcome = 'Failed' }

    $detail = $http.Error
    if ($tls -and -not $tls.Ok) {
        $tlsReason = $tls.PolicyErrors
        if ($tls.Error) { $tlsReason = $tls.Error }
        $tlsDetail = 'TLS: {0}' -f $tlsReason
        if ($tls.Subject) { $tlsDetail += '; Subject: {0}' -f $tls.Subject }
        if ($detail -and -not $httpTlsError) { $detail = '{0} | {1}' -f $detail, $tlsDetail } else { $detail = $tlsDetail }
    }
    if (-not $detail -and -not $dns.Ok) { $detail = $dns.Detail }
    if (-not $detail) { $detail = $retryNote }

    $certIssuer = $null
    $certRoot = $null
    $tlsVersion = $null
    if ($tls) {
        $certIssuer = $tls.Issuer
        $certRoot = $tls.Root
        $tlsVersion = $tls.Protocol
    }

    Write-Verbose ('{0}: DNS [{1}] | proxy {2} | root [{3}] | headers [{4}]' -f $endpoint.Host, $dns.Detail, $proxyLabel, $certRoot, ($headers -join '; '))

    $row = [pscustomobject]@{
        Group       = $endpoint.Group
        Host        = $endpoint.Host
        Url         = $uri.AbsoluteUri
        Source      = $endpoint.Source
        Result      = $outcome
        HttpStatus  = $http.Status
        DirectTcp   = $tcp
        Proxy       = $proxyLabel
        CertIssuer  = $certIssuer
        CertRoot    = $certRoot
        TlsVersion  = $tlsVersion
        ProxySignal = ($signals -join '; ')
        Addresses   = $dns.Detail
        Detail      = $detail
    }
    $results.Add($row)
    $signalsByHost[$endpoint.Host] = $signals.ToArray()
    Write-EndpointRow -Row $row -Signals $signals.ToArray() -HostWidth $hostWidth -Width $consoleWidth
}
Write-Progress -Activity 'Testing Windows service endpoints' -Completed

$attention = @($results | Where-Object { (Get-RowStatus -Row $_) -ne 'OK' })
if ($attention.Count -gt 0) {
    Write-Section 'Details'
    foreach ($row in $attention) {
        $status = Get-RowStatus -Row $row
        $color = 'Yellow'
        if ($status -eq 'FAIL') { $color = 'Red' }
        Write-Host ''
        Write-Host ('  {0}  ' -f $row.Host) -NoNewline
        Write-Host $row.Result -ForegroundColor $color
        Write-Host ('      {0}' -f $row.Url) -ForegroundColor DarkGray
        foreach ($signal in $signalsByHost[$row.Host]) { Write-Host ('      - {0}' -f $signal) -ForegroundColor $color }
        if ($row.Detail) { Write-Host ('      - {0}' -f $row.Detail) -ForegroundColor DarkGray }
    }
}

$viaZscaler = $false
$zscalerText = 'Skipped (-SkipZscalerCheck)'
$zscalerColor = 'DarkGray'
if (-not $SkipZscalerCheck) {
    Write-Progress -Activity 'Testing Windows service endpoints' -Status 'ip.zscaler.com'
    $zscalerUri = [Uri]'https://ip.zscaler.com/'
    $zscalerProxy = Get-EffectiveProxy -Uri $zscalerUri
    $zscaler = Test-Http -Uri $zscalerUri -ProxyUri $zscalerProxy -TimeoutMs $timeoutMs -MaxBytes 262144
    # ip.zscaler.com intermittently returns 502/504 or times out; retry once.
    if ($zscaler.Status -ne 200) { $zscaler = Test-Http -Uri $zscalerUri -ProxyUri $zscalerProxy -TimeoutMs $timeoutMs -MaxBytes 262144 }
    Write-Progress -Activity 'Testing Windows service endpoints' -Completed
    # Only the negative wording is verified; any other Zscaler page is treated as being on the Zscaler path.
    if (-not $zscaler.Status) { $zscalerText = 'Unknown ({0})' -f $zscaler.Error }
    elseif ($zscaler.Status -ne 200) { $zscalerText = 'Unknown (check service returned HTTP {0}; retry later)' -f $zscaler.Status }
    elseif ($zscaler.Body -match 'not going through the Zscaler proxy') {
        $zscalerText = 'Not via Zscaler'
        $zscalerColor = 'Green'
    }
    elseif ($zscaler.Body -match 'Zscaler') {
        $viaZscaler = $true
        $zscalerText = 'Via Zscaler - Microsoft/DigiCert chain roots mean SSL bypass; other roots mean inspection'
        $zscalerColor = 'Yellow'
    }
    else { $zscalerText = 'Unknown (HTTP {0}, unrecognised page)' -f $zscaler.Status }
}

# Name mismatch with an otherwise trusted chain = CDN not serving that hostname over HTTPS, not interception.
$failed = @($results | Where-Object { $_.Result -notin 'Reachable', 'ReachableNameMismatch' })
$intercepted = @($results | Where-Object { $_.ProxySignal })
$okCount = @($results | Where-Object { (Get-RowStatus -Row $_) -eq 'OK' }).Count
$warnCount = $results.Count - $okCount - $failed.Count

Write-Section 'Summary'
Write-Host ('  {0,-17}' -f 'Endpoints') -NoNewline -ForegroundColor DarkGray
Write-Host ('{0} OK' -f $okCount) -NoNewline -ForegroundColor Green
Write-Host ', ' -NoNewline -ForegroundColor DarkGray
$warnColor = 'DarkGray'
if ($warnCount -gt 0) { $warnColor = 'Yellow' }
Write-Host ('{0} warning' -f $warnCount) -NoNewline -ForegroundColor $warnColor
Write-Host ', ' -NoNewline -ForegroundColor DarkGray
$failColor = 'DarkGray'
if ($failed.Count -gt 0) { $failColor = 'Red' }
Write-Host ('{0} failed' -f $failed.Count) -NoNewline -ForegroundColor $failColor
Write-Host ('   (of {0})' -f $results.Count) -ForegroundColor DarkGray

$roots = @($results | Where-Object { $_.CertRoot } | Group-Object CertRoot | Sort-Object Count -Descending)
$rootKey = 'TLS chain roots'
foreach ($root in $roots) {
    $rootColor = 'DarkGray'
    if ($root.Name -notmatch $trustedRootPattern) { $rootColor = 'Yellow' }
    Write-KeyValue $rootKey ('{0,3} x {1}' -f $root.Count, (Get-CommonName -Subject $root.Name)) -Color $rootColor
    $rootKey = ''
}

# Without a chain the TLS-inspection check could not run, so 'no evidence' would overstate coverage.
$httpsRows = @($results | Where-Object { $_.Url -like 'https://*' })
$unchecked = @($httpsRows | Where-Object { -not $_.CertRoot })
if ($unchecked.Count -gt 0) {
    Write-KeyValue 'TLS not checked' ('{0} of {1} HTTPS endpoints - handshake failed, interception not assessed (see Detail)' -f $unchecked.Count, $httpsRows.Count) -Color Yellow
}
$versions = @($results | Where-Object { $_.TlsVersion } | Group-Object TlsVersion | Sort-Object Name | ForEach-Object { '{0} x {1}' -f $_.Count, $_.Name })
if ($versions.Count -gt 0) { Write-KeyValue 'TLS versions' ($versions -join ', ') -Color DarkGray }

if ($intercepted.Count -gt 0) { Write-KeyValue 'Proxy evidence' ('{0} endpoint(s) - see Details' -f $intercepted.Count) -Color Yellow }
elseif ($unchecked.Count -gt 0) { Write-KeyValue 'Proxy evidence' 'None found, but TLS inspection was not assessed on every endpoint' -Color Yellow }
else { Write-KeyValue 'Proxy evidence' 'None found on the tested path' -Color Green }

# Compare what was actually resolved: WPAD auto-detect is on by default and usually finds nothing.
$userProxied = @($results | Where-Object { $_.Proxy -ne 'Direct' }).Count -gt 0
if ($Proxy) { Write-KeyValue 'Service path' 'Forced proxy used for every probe' -Color DarkGray }
elseif ($userProxied -ne ($config.WinHttp -ne 'Direct')) {
    Write-KeyValue 'Service path' 'Differs from WinHTTP - rerun with -Proxy <WinHTTP proxy> or as SYSTEM' -Color Yellow
}
else { Write-KeyValue 'Service path' ('Same as WinHTTP ({0})' -f $config.WinHttp) -Color DarkGray }

# All-direct results are ambiguous with a PAC: deliberate bypass, or PAC not loaded. A neutral URL tells them apart.
if (-not $Proxy -and ($config.Pac -ne 'None' -or $config.AutoDetect)) {
    $directCount = @($results | Where-Object { $_.Proxy -eq 'Direct' }).Count
    $generalProxy = Get-EffectiveProxy -Uri ([Uri]'http://www.example.com/')
    if ($generalProxy) {
        Write-KeyValue 'PAC / WPAD' ('Loaded - general traffic via {0}; {1} of {2} Microsoft endpoints direct' -f $generalProxy.Authority, $directCount, $results.Count) -Color DarkGray
    }
    else {
        Write-KeyValue 'PAC / WPAD' 'DIRECT even for www.example.com - PAC not loaded, unreachable, or returns DIRECT for everything' -Color Yellow
    }
}

Write-KeyValue 'Zscaler' $zscalerText -Color $zscalerColor
Write-KeyValue 'Duration' ('{0:N0} s' -f ((Get-Date) - $started).TotalSeconds) -Color DarkGray

if ($CsvPath) {
    $results | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
    Write-KeyValue 'CSV' $CsvPath -Color DarkGray
}

$exitCode = 0
$verdict = 'PASS  All endpoints reachable, no proxy or interception evidence'
$verdictColor = 'Green'
if ($failed.Count -gt 0) {
    $exitCode = 1
    $verdict = 'FAIL  {0} of {1} endpoints not reachable as expected' -f $failed.Count, $results.Count
    $verdictColor = 'Red'
}
elseif ($intercepted.Count -gt 0 -or $viaZscaler) {
    $exitCode = 2
    $verdict = 'WARN  All endpoints reachable, but proxy or interception evidence found'
    $verdictColor = 'Yellow'
}
Write-Host ''
Write-Host ('  {0}  (exit code {1})' -f $verdict, $exitCode) -ForegroundColor $verdictColor
Write-Host ''
exit $exitCode
