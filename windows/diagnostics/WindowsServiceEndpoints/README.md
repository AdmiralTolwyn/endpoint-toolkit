# Test-WindowsServiceEndpoints

**Version:** 1.0
**Author:** Anton Romanyuk

> **Disclaimer:** This script is provided "as-is" without warranty of any kind, express or implied. Use at your own risk. The author assumes no liability for any damage or data loss resulting from its use. Always test in a non-production environment before deployment.

Read-only connectivity check for the Microsoft endpoints that Windows 11 Enterprise services depend on — Windows Update, Delivery Optimization, Microsoft Store, Defender, certificate trust, authentication, activation, settings, diagnostics, NCSI, push notifications and Edge update. Beyond "can I reach it", it looks for **a proxy in the middle**: explicit proxies, TLS inspection, proxy-generated block pages, redirects to proxy portals and content altered over plain HTTP.

## Problem

"Windows Update / Store / Defender doesn't work on this network" usually ends in a firewall-or-proxy discussion where nobody has hard evidence:

- `Test-NetConnection host -Port 443` only proves a TCP handshake. It says nothing about TLS inspection, a proxy answering on behalf of Microsoft, or whether the SYSTEM account takes the same path as the user.
- A proxy can return **HTTP 200** with its own block or "caution" page, or a redirect to its login portal — a naive HTTP check calls that success.
- With **SSL inspection** the proxy can rewrite HTTPS responses; with an **SSL exception** it cannot. Telling the two apart requires looking at the certificate chain actually delivered.
- Plain-HTTP endpoints (Windows Update downloads, certificate trust lists, OCSP, NCSI) can be altered silently, without any header.
- The Microsoft endpoint list contains wildcards (`*.windowsupdate.com`) that cannot be tested directly.

## What it checks

Endpoints come from [Connection endpoints for Windows 11 Enterprise](https://learn.microsoft.com/en-us/windows/privacy/manage-windows-11-endpoints). Wildcard entries are tested through representative hosts (`Source = Representative` in the CSV).

| Group | Hosts |
|---|---|
| Windows Update | `sls.update.microsoft.com`, `fe2cr.update.microsoft.com`, `fe3cr.delivery.mp.microsoft.com`, `tlu.dl.delivery.mp.microsoft.com`, `download.windowsupdate.com`, `tsfe.trafficshaping.dsp.mp.microsoft.com`, `adl.windows.com` |
| Delivery Optimization | `geo`, `kv801`, `cp801`, `disc801` `.prod.do.dsp.mp.microsoft.com` |
| Microsoft Store | `displaycatalog.mp.microsoft.com`, `storeedgefd.dsx.mp.microsoft.com`, `livetileedge.dsx.mp.microsoft.com`, `storecatalogrevocation.storequality.microsoft.com` |
| Defender | `wdcp.microsoft.com`, `definitionupdates.microsoft.com`, `checkappexec.microsoft.com`, `ping-edge` / `nav-edge` / `data-edge` `.smartscreen.microsoft.com` |
| Certificates | `ctldl.windowsupdate.com` (signed trust list), `ocsp.digicert.com` |
| Authentication / Activation | `login.live.com`, `licensing.mp.microsoft.com` |
| Settings / Diagnostics | `settings-win.data.microsoft.com`, `settings.data.microsoft.com`, `v10` / `self` / `functional` `.events.data.microsoft.com`, `watson.telemetry.microsoft.com` |
| NCSI | `www.msftconnecttest.com/connecttest.txt` |
| Push Notifications | `client.wns.windows.com` (TLS handshake only — WNS is not plain HTTP) |
| Edge Update | `msedge.api.cdp.microsoft.com` |

Per endpoint the script runs:

1. **DNS** with the Windows resolver (hosts file, cache, NRPT).
2. **Direct TCP** to 80/443, bypassing any proxy — shows whether direct egress is allowed.
3. **HTTP GET** through the proxy Windows would use (manual, PAC, WPAD, or `-Proxy`). Redirects are not followed; any HTTP status proves the endpoint answered.
4. **TLS handshake** through the same path (HTTP `CONNECT` when proxied), capturing the certificate chain root.

Transport failures (connection reset, EOF, timeout) are retried once when the host is otherwise reachable; a success on retry is reported as OK with a note.

> Any HTTP status — including `403`, `404` or `5xx` on a bare `/` — means the endpoint was **reached**. It does not mean an update scan or Store download will succeed.

## Proxy and interception detection

| Signal | How it is detected |
|---|---|
| Explicit proxy | WinINET manual proxy, PAC URL, WPAD auto-detect (resolved per URL), or `-Proxy` |
| PAC / WPAD evaluation | When a PAC URL or auto-detect is configured, the proxy for neutral `www.example.com` is resolved (no request sent). A proxy there with Microsoft endpoints direct = deliberate bypass; DIRECT there too = PAC not loaded, unreachable, or direct for everything |
| TLS inspection | Chain root is not Microsoft / DigiCert / Baltimore — checked through the proxy path, so an inspecting proxy's re-signed certificate is caught. Also catches transparent firewall decryption (e.g. NGFW forward-trust CA) with no proxy configured |
| Proxy headers | `Via`, `Proxy-Agent`, `X-Squid-*`, `X-Zscaler-*`, vendor names in `Server` / `X-*` headers; CDN `Via` values (varnish, Akamai, …) are ignored |
| Proxy authentication | `407` on the request or on `CONNECT` |
| Refused tunnel | `CONNECT` answered with anything other than `200` → `ProxyBlocked` |
| Proxy-generated page | Vendor names (Zscaler, Netskope, Forcepoint, …) or block/filter wording in the body — catches block pages served as **HTTP 200** |
| Redirect to proxy portal | `Location` pointing to a non-Microsoft host (e.g. a proxy authentication portal) |
| Content integrity over HTTP | `authrootstl.cab` is downloaded and its inner certificate trust list must verify as Microsoft-signed to a Microsoft root; NCSI must return exactly `Microsoft Connect Test` |
| DNS redirection | Public Microsoft names resolving to private, loopback or link-local addresses |
| Zscaler path | `ip.zscaler.com` reports whether traffic traverses Zscaler (retried once; the service is intermittently unavailable) |
| Steering agents | Installed Zscaler Client Connector, Netskope, GlobalProtect or Cisco Umbrella services — tunnel-mode clients set no Windows proxy at all |

### SSL exception vs SSL inspection

| Proxy mode | Can the proxy alter the response or answer 200 itself? | What the script shows |
|---|---|---|
| **SSL exception / bypass** (HTTPS) | No — TLS is end-to-end; the proxy can only allow or refuse the tunnel | Chain root is Microsoft / DigiCert → the response is genuinely from Microsoft |
| **SSL inspection** (HTTPS) | Yes | Chain root is the proxy's CA (e.g. *Zscaler Root CA*) → `TLS inspection` signal |
| **Plain HTTP** | Yes, always, invisibly | Signed trust-list and NCSI content checks; body and redirect checks |

## Result values

| Result | Meaning |
|---|---|
| `Reachable` | HTTP response from the endpoint (any status), or TLS handshake OK (WNS) |
| `ReachableNameMismatch` | Trusted chain, but the CDN certificate is issued for another name (e.g. `adl.windows.com` → `a248.e.akamai.net`) — not interception |
| `UnexpectedContent` | NCSI text or the signed trust list does not match what Microsoft publishes |
| `ProxyResponse` | The answer came from a proxy (block page or redirect to a non-Microsoft host) |
| `ProxyAuthRequired` | The proxy demanded authentication (`407`) |
| `ProxyBlocked` | The proxy refused the HTTPS tunnel |
| `TlsError` | TLS failed for a reason other than a name mismatch |
| `DnsFailed` | Name not resolvable and no proxy in use |
| `Failed` | Anything else (timeouts, refused connections) |

Console rows are tagged **OK**, **WARN** (reachable, but proxy/interception evidence) or **FAIL**.

## Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `TimeoutSec` | `int` (1–120) | `10` | Timeout per TCP connect and per HTTP request. |
| `Proxy` | `Uri` | — | Force every probe through this proxy, e.g. the WinHTTP proxy that services use. |
| `SkipZscalerCheck` | `switch` | Off | Do not contact `ip.zscaler.com`. |
| `CsvPath` | `string` | — | Export one row per endpoint (URL, result, HTTP status, direct TCP, proxy, certificate issuer/root, proxy signals, DNS addresses, detail). |

`-Verbose` logs DNS addresses, proxy, chain root and response headers per endpoint.

## Exit codes

| Code | Meaning |
|---|---|
| `0` | All endpoints reachable, no proxy or interception evidence |
| `1` | At least one endpoint failed |
| `2` | All endpoints reachable, but proxy or interception evidence found |

## Usage

### Basic

```powershell
.\Test-WindowsServiceEndpoints.ps1
```

### Export results

```powershell
.\Test-WindowsServiceEndpoints.ps1 -CsvPath C:\Temp\endpoints.csv
```

### Test the path Windows services take

Services such as Windows Update and Delivery Optimization run as SYSTEM and generally use the **WinHTTP** proxy, while an interactive run uses the account's **WinINET** settings. The summary flags when the two differ. Either force the WinHTTP proxy:

```powershell
netsh winhttp show proxy
.\Test-WindowsServiceEndpoints.ps1 -Proxy http://proxy.contoso.com:8080
```

…or run as SYSTEM:

```powershell
psexec -s -i powershell.exe -ExecutionPolicy Bypass -File .\Test-WindowsServiceEndpoints.ps1
```

### Example output

```
Windows service endpoint check
==============================
  Account          CONTOSO\user  (PowerShell 5.1.26100.9444)
  WinINET          Per user; manual proxy None; PAC None; auto-detect True
  WinHTTP          Direct
  Steering agents  None found

Endpoints
---------
  Windows Update
  [ OK ] sls.update.microsoft.com                           HTTP 404
  [ OK ] adl.windows.com                                    TLS ok    name mismatch (CDN cert)
  ...
Summary
-------
  Endpoints        34 OK, 0 warning, 0 failed   (of 34)
  TLS chain roots   15 x DigiCert Global Root G2
                     7 x Microsoft Root Certificate Authority 2011
  Proxy evidence   None found on the tested path
  Service path     Same as WinHTTP (Direct)
  Zscaler          Not via Zscaler

  PASS  All endpoints reachable, no proxy or interception evidence  (exit code 0)
```

## Limitations

- **No evidence is not proof of no proxy.** A transparent proxy that passes TLS through untouched and strips its headers is invisible from the client.
- Over plain HTTP, a bare `403` from a proxy is indistinguishable from the origin's `403`; only the signed trust list and NCSI detect replaced content.
- `CONNECT` is sent without proxy credentials. A proxy that requires Windows authentication answers `407` to the TLS probe even when the HTTP probe authenticates successfully, and the TLS-inspection check is then unavailable.
- The trusted-root list (Microsoft, DigiCert, Baltimore), proxy-header and block-page patterns are heuristics; validate once behind your own proxy.
- Representative hosts for wildcard entries were chosen from observed Windows traffic, not from the Microsoft article; confirm them against your firewall logs.
- The `ip.zscaler.com` positive ("via Zscaler") wording and the GlobalProtect / Umbrella service names are not yet verified on a device using those products.

## Requirements

- Windows PowerShell **5.1** or PowerShell 7.
- No elevation required. Run as SYSTEM to test the service path.
- Read-only: the script makes **no changes** to the system. It writes only a temporary folder for the trust-list signature check (always removed) and the optional CSV.
