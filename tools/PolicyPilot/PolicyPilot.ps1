<#
.SYNOPSIS
    PolicyPilot - Endpoint Policy Audit & Configuration Workbench.
.DESCRIPTION
    WPF application that reads Group Policy Objects (GPO), Intune/MDM policies, and local
    RSoP data to generate comprehensive policy documentation, detect duplicate and conflicting
    settings across policy sources, and produce enriched HTML reports.

    Supports four scan modes:
    - Local:    Scans policies applied to this machine via gpresult /x (no modules required)
    - AD:       Scans all domain GPOs via RSAT GroupPolicy module
    - Intune:   Scans Intune device configuration via Microsoft Graph
    - Combined: Merges AD + Intune + local policy data

    Features include IME/GPO/MDM log viewers with live tail, conflict detection with GPO
    precedence resolution, snapshot diff, baseline comparison, HTML/CSV/REG export,
    ADMX/CSP database enrichment for locale-independent policy names, and headless
    report generation mode.
.PARAMETER Headless
    Run without UI - generate an HTML report directly and exit.
.PARAMETER ReportType
    Scan mode for headless operation: Local, AD, Intune, or Combined.
.PARAMETER OutputPath
    Output path for the headless HTML report. Default: auto-generated in reports/ folder.
.NOTES
    Author : Anton Romanyuk
    Version: 0.1.0
    Date   : 2026-04-07
    Requires: PowerShell 5.1+, WPF (PresentationFramework)
    Requires: admx_metadata.json + csp_metadata.json (built by Build-AdmxDatabase.ps1 / Build-CspDatabase.ps1)
    Optional: RSAT GroupPolicy module (AD mode), Microsoft.Graph.DeviceManagement (Intune mode)
.EXAMPLE
    .\PolicyPilot.ps1
    Launches the WPF UI with default settings.
.EXAMPLE
    .\PolicyPilot.ps1 -Headless -ReportType Local -OutputPath C:\Reports\policy.html
    Generates a local policy HTML report without showing the UI.
.EXAMPLE
    .\PolicyPilot.ps1 -Headless -ReportType Combined
    Generates a combined AD+Intune+local report in headless mode.
#>
param(
    [switch]$Headless,
    # ValidateSet removed — PS 5.1 .GetNewClosure() captures all scope variables
    # and re-validates; empty $ReportType (GUI mode) fails ValidateSet.
    # Validated manually in headless path instead.
    [string]$ReportType,
    [string]$OutputPath
)
# ╔═══════════════════════════════════════════════════════════════════════════════╗
# ║  PolicyPilot - Endpoint Policy Audit & Configuration Workbench  ║
# ║  Version 0.1.0  |  PowerShell 5.1+ WPF Application                         ║
# ╚═══════════════════════════════════════════════════════════════════════════════╝

$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Web

# ── DPI Awareness ──
try {
    if (-not ([System.Management.Automation.PSTypeName]'DpiHelper').Type) {
        Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class DpiHelper {
    [DllImport("shcore.dll")] public static extern int SetProcessDpiAwareness(int value);
}
"@
    }
    [void][DpiHelper]::SetProcessDpiAwareness(2)
} catch { try { Write-DebugLog "Unhandled: $_" -Level ERROR } catch {} }

# -- Hide console window (the black box behind the WPF UI) --
try {
    if (-not ([System.Management.Automation.PSTypeName]'ConsoleHelper').Type) {
        Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class ConsoleHelper {
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]   public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@
    }
    $consoleHwnd = [ConsoleHelper]::GetConsoleWindow()
    if ($consoleHwnd -ne [IntPtr]::Zero) {
        [void][ConsoleHelper]::ShowWindow($consoleHwnd, 0)   # 0 = SW_HIDE
    }
} catch {}

# ═══════════════════════════════════════════════════════════════════════════════
# Compiled C# - PolicyLogParser: fast parse + classify for IME / GPO / MDM logs
# ═══════════════════════════════════════════════════════════════════════════════
try {
if (-not ([System.Management.Automation.PSTypeName]'PolicyLogParser').Type) {
$_plpSrc = @"
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

public static class PolicyLogParser
{
    // ── Result arrays (set by each method) ──
    public static string[] FormattedLines;
    public static byte[]   LineTypes;
    public static string[] RawEntries;
    public static int      TotalLines;
    public static int      ErrorCount;
    public static int      WarningCount;

    // ════════════════════════════════════════════════════════════════
    // IME patterns
    // ════════════════════════════════════════════════════════════════
    static readonly Regex ImeRxCMTrace = new Regex(
        @"<!\[LOG\[(?<msg>.*?)\]LOG\]!><time=""(?<time>\d{1,2}:\d{2}:\d{2})\.\d+"" date=""(?<date>\d{1,2}-\d{1,2}-\d{4})"" component=""(?<comp>[^""]*)"" context=""[^""]*"" type=""(?<type>\d)"" thread=""(?<thread>\d+)"" file=""[^""]*"">",
        RegexOptions.Compiled | RegexOptions.Singleline);
    static readonly Regex ImeRxPlainTS = new Regex(
        @"^<!\[LOG\[(?<msg>.*?)\]LOG\]!>", RegexOptions.Compiled | RegexOptions.Singleline);

    static readonly Regex ImeRxError   = new Regex(@"(?i)(error|exception|fail(ed|ure)?|fatal|HRESULT\s*[:=]\s*0x8|StatusCode\s*[:=]\s*[45]\d{2}|could not|unable to|Access denied|not found.*critical|ExpectedPolicies)", RegexOptions.Compiled);
    static readonly Regex ImeRxWarning = new Regex(@"(?i)(warn(ing)?|timeout|retry|retrying|expired|fallback|skipp(ed|ing)|already exists|not applicable|timed out|throttl)", RegexOptions.Compiled);
    static readonly Regex ImeRxSuccess = new Regex(@"(?i)(successfully|completed|installed|applied|compliance.*(true|met)|detected|remediated|execution completed|exit code[:=]\s*0\b|Win32App.*successfully)", RegexOptions.Compiled);
    static readonly Regex ImeRxPolicy  = new Regex(@"(?i)(policy|assignment|targeting|scope|applicability|evaluation|SideCar|StatusService|check-in|PolicyId)", RegexOptions.Compiled);
    static readonly Regex ImeRxApp     = new Regex(@"(?i)(Win32App|IntuneApp|\.intunewin|ContentManager|download.*app|install.*app|detection rule|requirement rule|AppInstall|AppWorkload|Content.*download)", RegexOptions.Compiled);
    static readonly Regex ImeRxScript  = new Regex(@"(?i)(HealthScript|Proactive.*remediation|Sensor|AgentExecutor|PowerShell|script.*execution|ScriptHandler|remediationScript|detectionScript)", RegexOptions.Compiled);
    static readonly Regex ImeRxSync    = new Regex(@"(?i)(sync|check.?in|session|enrollment|OMA-?DM|SyncML|MDM.*session|DeviceEnroller|DMClient|polling|schedule)", RegexOptions.Compiled);

    static readonly string[] ImeBadgeMap = {"INFO","ERR","WARN","OK","APP","SCRP","SYNC","POL"};

    static byte ClassifyImeSeverity(string msg, string cmTraceType)
    {
        if (cmTraceType == "3") return 1;
        if (cmTraceType == "2") return 2;
        if (ImeRxError.IsMatch(msg))   return 1;
        if (ImeRxWarning.IsMatch(msg)) return 2;
        if (ImeRxSuccess.IsMatch(msg)) return 3;
        if (ImeRxApp.IsMatch(msg))     return 4;
        if (ImeRxScript.IsMatch(msg))  return 5;
        if (ImeRxSync.IsMatch(msg))    return 6;
        if (ImeRxPolicy.IsMatch(msg))  return 7;
        return 0;
    }

    // ════════════════════════════════════════════════════════════════
    // GPO patterns
    // ════════════════════════════════════════════════════════════════
    static readonly Regex GpoRxGpsvcLine = new Regex(
        @"^GPSVC\((?<pid>[0-9a-fA-F]+)\.(?<tid>[0-9a-fA-F]+)\)\s+(?<time>\d{2}:\d{2}:\d{2}:\d{3})\s+(?<func>\w+):\s+(?<msg>.*)", RegexOptions.Compiled);

    static readonly Regex GpoRxError   = new Regex(@"(?i)(error|fail(ed|ure)?|denied|not found|could not|unable to|HRESULT|exception|0x8\w{7}|Access is denied|no.*domain controller|unreachable)", RegexOptions.Compiled);
    static readonly Regex GpoRxWarning = new Regex(@"(?i)(warn(ing)?|timeout|retry|retrying|slow link|loopback|no changes|skipped|disabled|not applied|blocked|WMI filter.*false|security filter.*denied)", RegexOptions.Compiled);
    static readonly Regex GpoRxSuccess = new Regex(@"(?i)(success(fully)?|completed|applied|processed|linked|no errors|exit code[:=]\s*0\b)", RegexOptions.Compiled);
    static readonly Regex GpoRxCSE     = new Regex(@"(?i)(client.side extension|CSE|registry settings|security settings|folder redirection|software installation|scripts extension|administrative templates|drive maps|printers|preferences|internet explorer|firewall|wireless|wired|public key)", RegexOptions.Compiled);
    static readonly Regex GpoRxPolicy  = new Regex(@"(?i)(Group Policy|GPO|LGPO|policy (processing|application)|linked.*GPO.*list|filtering|WMI filter|security filter|scope of management)", RegexOptions.Compiled);
    static readonly Regex GpoRxNetwork = new Regex(@"(?i)(domain controller|DC[:=]|LDAP|site[:=]|network|bandwidth|slow link|NLA|DNS|connectivity)", RegexOptions.Compiled);
    static readonly Regex GpoRxSync    = new Regex(@"(?i)(gpupdate|background.*processing|foreground.*processing|manual.*refresh|user.*policy|computer.*policy|processing mode|loopback)", RegexOptions.Compiled);

    static readonly string[] GpoBadgeMap = {"INFO","ERR","WARN","OK","CSE","NET","SYNC","POL"};

    static byte ClassifyGpoSeverity(string msg, int eventLevel)
    {
        if (eventLevel == 1 || eventLevel == 2) return 1;
        if (eventLevel == 3) return 2;
        if (GpoRxError.IsMatch(msg))   return 1;
        if (GpoRxWarning.IsMatch(msg)) return 2;
        if (GpoRxSuccess.IsMatch(msg)) return 3;
        if (GpoRxCSE.IsMatch(msg))     return 4;
        if (GpoRxNetwork.IsMatch(msg)) return 5;
        if (GpoRxSync.IsMatch(msg))    return 6;
        if (GpoRxPolicy.IsMatch(msg))  return 7;
        return 0;
    }

    // ════════════════════════════════════════════════════════════════
    // MDM patterns
    // ════════════════════════════════════════════════════════════════
    static readonly Regex MdmRxError   = new Regex(@"(?i)(error|fail(ed|ure)?|denied|not found|could not|unable to|HRESULT|exception|0x8\w{7}|rejected|unauthorized|timeout)", RegexOptions.Compiled);
    static readonly Regex MdmRxWarning = new Regex(@"(?i)(warn(ing)?|retry|retrying|skipped|not applicable|conflict|pending|throttl)", RegexOptions.Compiled);
    static readonly Regex MdmRxSuccess = new Regex(@"(?i)(success(fully)?|completed|accepted|applied|result.*=.*0\b|status.*200|status.*OK)", RegexOptions.Compiled);
    static readonly Regex MdmRxSyncML  = new Regex(@"(?i)(SyncML|SyncBody|SyncHdr|OMA-?DM|<Add>|<Replace>|<Get>|<Delete>|<Exec>|<Alert>|<Status>|<Final>)", RegexOptions.Compiled);
    static readonly Regex MdmRxPolicy  = new Regex(@"(?i)(policy|configuration|compliance|enrollment|MDM session|check-?in|CSP|\./Device/|\./User/|\./Vendor/)", RegexOptions.Compiled);
    static readonly Regex MdmRxNetwork = new Regex(@"(?i)(http|https|push notification|WNS|MPNS|server|endpoint|DMClient|connection|channel)", RegexOptions.Compiled);

    static readonly string[] MdmBadgeMap = {"INFO","ERR","WARN","OK","SYNC","POL","NET"};

    static byte ClassifyMdmSeverity(string msg, int eventLevel)
    {
        if (eventLevel == 1 || eventLevel == 2) return 1;
        if (eventLevel == 3) return 2;
        if (MdmRxError.IsMatch(msg))   return 1;
        if (MdmRxWarning.IsMatch(msg)) return 2;
        if (MdmRxSuccess.IsMatch(msg)) return 3;
        if (MdmRxSyncML.IsMatch(msg))  return 4;
        if (MdmRxPolicy.IsMatch(msg))  return 5;
        if (MdmRxNetwork.IsMatch(msg)) return 6;
        return 0;
    }

    // ════════════════════════════════════════════════════════════════
    // Shared helpers
    // ════════════════════════════════════════════════════════════════
    static bool PassesContentFilter(string text, string filter)
    {
        if (string.IsNullOrEmpty(filter)) return true;
        if (filter.Length > 2 && filter[0] == '/' && filter[filter.Length - 1] == '/')
        {
            try { return Regex.IsMatch(text, filter.Substring(1, filter.Length - 2), RegexOptions.IgnoreCase); }
            catch { return text.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0; }
        }
        return text.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0;
    }

    static bool PassesCategoryFilter(string filterName, byte lt)
    {
        if (string.IsNullOrEmpty(filterName) || filterName == "All") return true;
        if (filterName == "Error")   return lt == 1;
        if (filterName == "Warning") return lt == 2;
        if (filterName == "Info")    return lt != 1 && lt != 2;
        return true;
    }

    // ════════════════════════════════════════════════════════════════
    // ParseAndClassifyImeFile - stream-read + parse CMTrace + classify
    // ════════════════════════════════════════════════════════════════
    public static void ParseAndClassifyImeFile(
        string filePath, int lineLimit, string categoryFilter, string contentFilter)
    {
        string allText; bool skipped = false;
        using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read,
                                       FileShare.ReadWrite | FileShare.Delete))
        {
            if (lineLimit > 0 && fs.Length > 0)
            {
                long seekBack = Math.Min(fs.Length, (long)lineLimit * 240);
                if (seekBack < fs.Length) { fs.Seek(-seekBack, SeekOrigin.End); skipped = true; }
            }
            using (var sr = new StreamReader(fs))
            {
                if (skipped) sr.ReadLine();
                allText = sr.ReadToEnd();
            }
        }
        var lines = allText.Split(new[]{'\r','\n'}, StringSplitOptions.RemoveEmptyEntries);
        int startIdx = (skipped && lineLimit > 0 && lines.Length > lineLimit)
                        ? lines.Length - lineLimit : 0;
        ClassifyImeCore(lines, startIdx, lines.Length, categoryFilter, contentFilter);
    }

    // ════════════════════════════════════════════════════════════════
    // ClassifyImeLines - batch classify pre-read raw lines (timer)
    // ════════════════════════════════════════════════════════════════
    public static void ClassifyImeLines(
        string[] rawLines, string categoryFilter, string contentFilter)
    {
        ClassifyImeCore(rawLines, 0, rawLines.Length, categoryFilter, contentFilter);
    }

    static void ClassifyImeCore(
        string[] lines, int startIdx, int endIdx,
        string categoryFilter, string contentFilter)
    {
        int cap = endIdx - startIdx;
        var fmt  = new List<string>(cap);
        var typs = new List<byte>(cap);
        var raws = new List<string>(cap);
        int errors = 0, warnings = 0;
        var sb = new StringBuilder(256);

        for (int i = startIdx; i < endIdx; i++)
        {
            var line = lines[i];
            if (string.IsNullOrWhiteSpace(line)) continue;

            string msg, time, date, comp, type, thread;
            var m = ImeRxCMTrace.Match(line);
            if (m.Success)
            {
                msg = m.Groups["msg"].Value;   time = m.Groups["time"].Value;
                date = m.Groups["date"].Value;  comp = m.Groups["comp"].Value;
                type = m.Groups["type"].Value;  thread = m.Groups["thread"].Value;
            }
            else
            {
                var m2 = ImeRxPlainTS.Match(line);
                msg = m2.Success ? m2.Groups["msg"].Value : line;
                time = ""; date = ""; comp = ""; type = "1"; thread = "";
            }

            byte lt = ClassifyImeSeverity(msg, type);
            if (!PassesCategoryFilter(categoryFilter, lt)) continue;
            if (!PassesContentFilter(msg + " " + comp + " " + line, contentFilter)) continue;

            if (lt == 1) errors++;
            if (lt == 2) warnings++;

            sb.Clear();
            sb.Append('[').Append(ImeBadgeMap[lt]).Append("] ");
            if (time.Length > 0) { if (date.Length > 0) sb.Append(date).Append(' '); sb.Append(time).Append(' '); }
            if (comp.Length > 0) sb.Append('[').Append(comp).Append("] ");
            sb.Append(msg);
            if (thread.Length > 0) sb.Append("  T:").Append(thread);

            fmt.Add(sb.ToString());
            typs.Add(lt);
            raws.Add(line);
        }

        FormattedLines = fmt.ToArray();
        LineTypes = new byte[typs.Count]; typs.CopyTo(LineTypes);
        RawEntries = raws.ToArray();
        TotalLines = fmt.Count; ErrorCount = errors; WarningCount = warnings;
    }

    // ════════════════════════════════════════════════════════════════
    // ClassifyGpoEntries - classify pre-parsed event log entries
    // ════════════════════════════════════════════════════════════════
    public static void ClassifyGpoEntries(
        string[] messages, int[] levels, string[] components,
        string[] eventIds, string[] times, string[] threads,
        string categoryFilter, string contentFilter)
    {
        int n = messages.Length;
        var fmt  = new List<string>(n);
        var typs = new List<byte>(n);
        var raws = new List<string>(n);
        int errors = 0, warnings = 0;
        var sb = new StringBuilder(256);

        for (int i = 0; i < n; i++)
        {
            string msg = messages[i]; int level = levels[i];
            byte lt = ClassifyGpoSeverity(msg, level);
            if (!PassesCategoryFilter(categoryFilter, lt)) continue;
            if (!PassesContentFilter(msg + " " + components[i] + " " + eventIds[i], contentFilter)) continue;

            if (lt == 1) errors++;
            if (lt == 2) warnings++;

            sb.Clear();
            sb.Append('[').Append(GpoBadgeMap[lt]).Append("] ");
            if (times[i].Length > 0) sb.Append(times[i]).Append(' ');
            if (components[i].Length > 0) sb.Append('[').Append(components[i]).Append("] ");
            if (eventIds[i].Length > 0) sb.Append("ID:").Append(eventIds[i]).Append(' ');
            sb.Append(msg);
            if (threads[i].Length > 0) sb.Append("  T:").Append(threads[i]);

            fmt.Add(sb.ToString());
            typs.Add(lt);
            raws.Add(msg);
        }

        FormattedLines = fmt.ToArray();
        LineTypes = new byte[typs.Count]; typs.CopyTo(LineTypes);
        RawEntries = raws.ToArray();
        TotalLines = fmt.Count; ErrorCount = errors; WarningCount = warnings;
    }

    // ════════════════════════════════════════════════════════════════
    // ParseAndClassifyGpsvcFile - parse gpsvc.log + classify
    // ════════════════════════════════════════════════════════════════
    public static void ParseAndClassifyGpsvcFile(
        string filePath, string categoryFilter, string contentFilter)
    {
        var lines = File.ReadAllLines(filePath);
        int n = lines.Length;
        var fmt  = new List<string>(n);
        var typs = new List<byte>(n);
        var raws = new List<string>(n);
        int errors = 0, warnings = 0;
        var sb = new StringBuilder(256);

        for (int i = 0; i < n; i++)
        {
            var line = lines[i];
            if (string.IsNullOrWhiteSpace(line)) continue;

            string msg, time, comp, thread;
            var m = GpoRxGpsvcLine.Match(line);
            if (m.Success)
            {
                msg = m.Groups["msg"].Value;  time = m.Groups["time"].Value;
                comp = m.Groups["func"].Value; thread = m.Groups["tid"].Value;
            }
            else { msg = line; time = ""; comp = ""; thread = ""; }

            byte lt = ClassifyGpoSeverity(msg, 4);
            if (!PassesCategoryFilter(categoryFilter, lt)) continue;
            if (!PassesContentFilter(msg + " " + comp, contentFilter)) continue;

            if (lt == 1) errors++;
            if (lt == 2) warnings++;

            sb.Clear();
            sb.Append('[').Append(GpoBadgeMap[lt]).Append("] ");
            if (time.Length > 0) sb.Append(time).Append(' ');
            if (comp.Length > 0) sb.Append('[').Append(comp).Append("] ");
            sb.Append(msg);
            if (thread.Length > 0) sb.Append("  T:").Append(thread);

            fmt.Add(sb.ToString());
            typs.Add(lt);
            raws.Add(msg);
        }

        FormattedLines = fmt.ToArray();
        LineTypes = new byte[typs.Count]; typs.CopyTo(LineTypes);
        RawEntries = raws.ToArray();
        TotalLines = fmt.Count; ErrorCount = errors; WarningCount = warnings;
    }

    // ════════════════════════════════════════════════════════════════
    // ClassifyMdmEntries - classify pre-parsed MDM event log entries
    // ════════════════════════════════════════════════════════════════
    public static void ClassifyMdmEntries(
        string[] messages, int[] levels, string[] components,
        string[] eventIds, string[] times, string[] threads,
        string categoryFilter, string contentFilter)
    {
        int n = messages.Length;
        var fmt  = new List<string>(n);
        var typs = new List<byte>(n);
        var raws = new List<string>(n);
        int errors = 0, warnings = 0;
        var sb = new StringBuilder(512);

        for (int i = 0; i < n; i++)
        {
            string msg = messages[i]; int level = levels[i];
            byte lt = ClassifyMdmSeverity(msg, level);
            if (!PassesCategoryFilter(categoryFilter, lt)) continue;
            if (!PassesContentFilter(msg + " " + components[i] + " " + eventIds[i], contentFilter)) continue;

            if (lt == 1) errors++;
            if (lt == 2) warnings++;

            sb.Clear();
            sb.Append('[').Append(MdmBadgeMap[lt]).Append("] ");
            if (times[i].Length > 0) sb.Append(times[i]).Append(' ');
            if (components[i].Length > 0) sb.Append('[').Append(components[i]).Append("] ");
            if (eventIds[i].Length > 0) sb.Append("ID:").Append(eventIds[i]).Append(' ');
            sb.Append(msg);
            if (threads[i].Length > 0) sb.Append("  T:").Append(threads[i]);

            fmt.Add(sb.ToString());
            typs.Add(lt);
            raws.Add(msg);
        }

        FormattedLines = fmt.ToArray();
        LineTypes = new byte[typs.Count]; typs.CopyTo(LineTypes);
        RawEntries = raws.ToArray();
        TotalLines = fmt.Count; ErrorCount = errors; WarningCount = warnings;
    }
}
"@
Add-Type -TypeDefinition $_plpSrc
}
$Script:HasCSharpParser = $true
} catch {
    Write-Host "[PolicyPilot] C# PolicyLogParser compile failed: $_"
    $Script:HasCSharpParser = $false
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 1: GLOBAL STATE
# ═══════════════════════════════════════════════════════════════════════════════

$Script:AppVersion    = '0.1.0'
$Script:AppDir        = Split-Path -Parent $MyInvocation.MyCommand.Definition
$Script:PrefsPath     = Join-Path $Script:AppDir 'user_prefs.json'
$Script:LogPath       = Join-Path $Script:AppDir 'debug.log'
$Script:SnapshotDir   = Join-Path $Script:AppDir 'snapshots'
$Script:ReportsDir    = Join-Path $Script:AppDir 'reports'
$Script:AchievementsFile = Join-Path $Script:AppDir 'achievements.json'

# Ensure directories exist
foreach ($d in @($Script:SnapshotDir, $Script:ReportsDir)) {
    if (-not (Test-Path $d)) { New-Item -Path $d -ItemType Directory -Force }
}

# ── ADMX + CSP Metadata Databases ──
# Used for locale-independent policy name resolution and HTML report enrichment.
# If the JSON files are missing, the tool falls back to locale-dependent names.
$Script:AdmxDbPath = Join-Path $Script:AppDir 'admx_metadata.json'
$Script:CspDbPath  = Join-Path $Script:AppDir 'csp_metadata.json'
$Script:AdmxDb     = $null   # keyed by "AdmxFile/PolicyName", each has RegistryKey, ValueName, Friendly, Desc, etc.
$Script:AdmxByReg  = @{}     # keyed by "HKLM_RegistryKey\ValueName" (normalized) for fast O(1) lookup
$Script:CspDb      = $null   # keyed by "Area/PolicyName", each has GPMapping with registry key
$Script:CspByReg   = @{}     # keyed by normalized registry path for CSP enrichment
$Script:AdmxDbAge  = $null   # days since DB was generated (show stale warning if >90)
$Script:CspDbAge   = $null

function Load-MetadataDatabases {
    # Helper: ConvertFrom-Json with -AsHashtable on PS7+, fallback on PS5.1
    $useHashtable = $PSVersionTable.PSVersion.Major -ge 6

    # Load ADMX database
    if (Test-Path $Script:AdmxDbPath) {
        try {
            $rawJson = Get-Content $Script:AdmxDbPath -Raw -Encoding UTF8
            # PS 5.1 cannot parse JSON with empty-string keys — strip all "": ... entries
            if (-not $useHashtable) {
                # Remove "": "value", "": number, "": null patterns (empty-key properties in EnumValues)
                $rawJson = $rawJson -replace '""\s*:\s*"[^"]*"\s*,?\s*', '' -replace '""\s*:\s*\d+\s*,?\s*', '' -replace '""\s*:\s*null\s*,?\s*', ''
                # Clean up any resulting syntax issues (trailing commas, double commas)
                $rawJson = $rawJson -replace ',\s*,', ',' -replace ',\s*\}', '}' -replace ',\s*\]', ']'
            }
            if ($useHashtable) {
                $raw = $rawJson | ConvertFrom-Json -AsHashtable
                $Script:AdmxDb = $raw.policies
                $policyKeys = $raw.policies.Keys
            } else {
                $raw = $rawJson | ConvertFrom-Json
                $Script:AdmxDb = $raw.policies
                $policyKeys = @($raw.policies.PSObject.Properties | ForEach-Object { $_.Name })
            }
            if ($raw._metadata -and $raw._metadata.generatedAt) {
                $Script:AdmxDbAge = [int]((Get-Date) - [datetime]$raw._metadata.generatedAt).TotalDays
            }
            # Build registry-key index for O(1) lookup
            $keyOnlyCount = 0
            foreach ($key in $policyKeys) {
                $entry = if ($useHashtable) { $Script:AdmxDb[$key] } else { $Script:AdmxDb.$key }
                if ($entry.RegistryKey) {
                    if ($entry.ValueName) {
                        # Full key+value index (exact match)
                        $normKey = ($entry.RegistryKey -replace '\\$','').ToLower() + '\' + $entry.ValueName.ToLower()
                        $Script:AdmxByReg[$normKey] = $entry
                    } else {
                        # Key-only index (fallback for policies that use key existence = enabled)
                        $normKey = ($entry.RegistryKey -replace '\\$','').ToLower()
                        if (-not $Script:AdmxByReg.ContainsKey($normKey)) {
                            $Script:AdmxByReg[$normKey] = $entry
                            $keyOnlyCount++
                        }
                    }
                }
            }
            Write-Host "[PolicyPilot] ADMX database loaded: $($Script:AdmxByReg.Count) indexed ($keyOnlyCount key-only fallbacks, age: $($Script:AdmxDbAge)d)" -ForegroundColor DarkGray
        } catch { Write-Host "[PolicyPilot] ADMX database load failed: $_" -ForegroundColor Yellow }
    } else {
        Write-Host "[PolicyPilot] ADMX database not found: $Script:AdmxDbPath (policy names will use OS locale)" -ForegroundColor Yellow
    }

    # Load CSP database
    if (Test-Path $Script:CspDbPath) {
        try {
            $cspJson = Get-Content $Script:CspDbPath -Raw -Encoding UTF8
            if ($useHashtable) {
                $cspRaw = $cspJson | ConvertFrom-Json -AsHashtable
                $cspKeys = $cspRaw.Keys
            } else {
                $cspRaw = $cspJson | ConvertFrom-Json
                $cspKeys = @($cspRaw.PSObject.Properties | ForEach-Object { $_.Name })
            }
            $Script:CspDb = $cspRaw
            if (($useHashtable -and $cspRaw._metadata -and $cspRaw._metadata.generatedAt) -or
                (-not $useHashtable -and $cspRaw._metadata -and $cspRaw._metadata.generatedAt)) {
                $Script:CspDbAge = [int]((Get-Date) - [datetime]$cspRaw._metadata.generatedAt).TotalDays
            } else {
                # No metadata block - use file modification date
                $Script:CspDbAge = [int]((Get-Date) - (Get-Item $Script:CspDbPath).LastWriteTime).TotalDays
            }
            # Build registry-key index from GPMapping + CSP path index
            $cspKeyOnly = 0
            $Script:CspByPath = @{}  # keyed by CSP OMA-URI path (e.g. "Defender/AllowRealtimeMonitoring")
            foreach ($key in $cspKeys) {
                if ($key -eq '_metadata') { continue }
                $entry = if ($useHashtable) { $cspRaw[$key] } else { $cspRaw.$key }
                # CSP path index (always available)
                $cspInfo = @{
                    CspPath       = $key
                    Friendly      = $entry.Friendly
                    Desc          = $entry.Desc
                    Scope         = $entry.Scope
                    Default       = $entry.Def
                    AllowedValues = $entry.AllowedValues
                }
                $Script:CspByPath[$key] = $cspInfo
                # Registry-key index from GPMapping
                if ($entry.GPMapping -and $entry.GPMapping.'Registry Key Name') {
                    $regKey = $entry.GPMapping.'Registry Key Name'
                    $valName = $entry.GPMapping.'Registry Value Name'
                    if ($regKey -and $valName) {
                        $normKey = ($regKey -replace '\\$','').ToLower() + '\' + $valName.ToLower()
                        $Script:CspByReg[$normKey] = $cspInfo
                    } elseif ($regKey) {
                        # Key-only fallback
                        $normKey = ($regKey -replace '\\$','').ToLower()
                        if (-not $Script:CspByReg.ContainsKey($normKey)) {
                            $Script:CspByReg[$normKey] = $cspInfo
                            $cspKeyOnly++
                        }
                    }
                }
            }
            $ageLabel = if ($Script:CspDbAge) { ", age: $($Script:CspDbAge)d" } else { '' }
            Write-Host "[PolicyPilot] CSP database loaded: $($Script:CspByReg.Count) reg-indexed ($cspKeyOnly key-only), $($Script:CspByPath.Count) by CSP path$ageLabel" -ForegroundColor DarkGray
        } catch { Write-Host "[PolicyPilot] CSP database load failed: $_" -ForegroundColor Yellow }
    }
}

<#
.SYNOPSIS
    Resolves a locale-independent English policy display name from a registry key + value.
.DESCRIPTION
    Looks up the registry path in the ADMX database index. If found, returns the English
    display name from the ADMX. If not found, returns the original (locale-dependent) name
    with an [ADMX?] marker so the report consumer knows the name is unresolved.
.PARAMETER RegistryKey
    The registry key path (e.g. 'Software\Policies\Microsoft\Windows\System').
.PARAMETER ValueName
    The registry value name.
.PARAMETER FallbackName
    The original locale-dependent name to use if no ADMX match is found.
.OUTPUTS
    [hashtable] with keys: Name (resolved display name), Desc (English description),
    Category (ADMX category path), Source ('ADMX'|'CSP'|'Local'), CspPath (if from CSP),
    AllowedValues (if available from CSP/ADMX).
#>
function Resolve-PolicyFromRegistry {
    param(
        [string]$RegistryKey,
        [string]$ValueName,
        [string]$FallbackName = ''
    )
    $result = @{
        Name          = $FallbackName
        Desc          = ''
        Category      = ''
        Source         = 'Local'
        CspPath       = ''
        AllowedValues = $null
        AdmxFile      = ''
    }

    if (-not $RegistryKey) { return $result }

    $normKey = ($RegistryKey -replace '\\$','').ToLower()
    if ($ValueName) { $normKey += '\' + $ValueName.ToLower() }

    # Try ADMX database first (authoritative for GPO policy names)
    if ($Script:AdmxByReg.Count -gt 0) {
        # Try full key+value match first, then key-only fallback
        $admx = $null
        if ($Script:AdmxByReg.ContainsKey($normKey)) {
            $admx = $Script:AdmxByReg[$normKey]
        } elseif ($ValueName) {
            # Try key-only (strip value name) for policies indexed without value
            $keyOnly = ($RegistryKey -replace '\\$','').ToLower()
            if ($Script:AdmxByReg.ContainsKey($keyOnly)) { $admx = $Script:AdmxByReg[$keyOnly] }
        } elseif ($normKey -match '^(.+)\\([^\\]+)$') {
            # RegistryKey may contain keyPath\valueName combined — try parent key as fallback
            $parentKey = $Matches[1]
            if ($Script:AdmxByReg.ContainsKey($parentKey)) { $admx = $Script:AdmxByReg[$parentKey] }
        }
        if ($admx) {
            $result.Name     = if ($admx.Friendly) { $admx.Friendly } else { $FallbackName }
            $result.Desc     = if ($admx.Desc) { $admx.Desc } else { '' }
            $result.Category = if ($admx.Category) { $admx.Category } else { '' }
            $result.Source   = 'ADMX'
            $result.AdmxFile = if ($admx.AdmxFile) { $admx.AdmxFile } else { '' }
            if ($admx.Elements) { $result.AllowedValues = $admx.Elements }
            return $result
        }
    }

    # Try CSP database (for Intune-delivered policies)
    if ($Script:CspByReg.Count -gt 0) {
        $csp = $null
        if ($Script:CspByReg.ContainsKey($normKey)) {
            $csp = $Script:CspByReg[$normKey]
        } elseif ($ValueName) {
            $keyOnly = ($RegistryKey -replace '\\$','').ToLower()
            if ($Script:CspByReg.ContainsKey($keyOnly)) { $csp = $Script:CspByReg[$keyOnly] }
        } elseif ($normKey -match '^(.+)\\([^\\]+)$') {
            $parentKey = $Matches[1]
            if ($Script:CspByReg.ContainsKey($parentKey)) { $csp = $Script:CspByReg[$parentKey] }
        }
        if ($csp) {
            $result.Name     = if ($csp.Friendly) { $csp.Friendly } else { $FallbackName }
            $result.Desc     = if ($csp.Desc) { $csp.Desc } else { '' }
            $result.Source   = 'CSP'
            $result.CspPath  = $csp.CspPath
            if ($csp.AllowedValues) { $result.AllowedValues = $csp.AllowedValues }
            return $result
        }
    }

    # No match - return fallback with marker
    if ($FallbackName -and $Script:AdmxByReg.Count -gt 0) {
        # DB is loaded but this specific key wasn't found
        $result.Source = 'Local'
    }
    return $result
}

Load-MetadataDatabases

# Application state
$Script:AllGPOs          = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
$Script:AllSettings      = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
$Script:AllConflicts     = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
$Script:AllIntuneApps    = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
$Script:TopConflicts     = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
$Script:ActiveTab        = 'Dashboard'
$Script:AnimationsDisabled = $false
$Script:ScanData         = $null   # raw scan result for snapshot/export
$Script:NotConfiguredSettings = [System.Collections.Generic.List[PSCustomObject]]::new()
$Script:CountUpTimers    = @{}
$Script:CountUpAnims     = @{}
$Script:PrereqsMet       = $false
$Script:CONFETTI_COUNT    = 50
$Script:Achievements      = @{}
$Script:SidebarCollapsed  = $false
$Global:ConsoleReady     = $false

# -- Thread Synchronization Bridge (same pattern as WinGetManifestManager) --
$Global:SyncHash = [Hashtable]::Synchronized(@{
    StatusQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new())
})
$Global:BgJobs          = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new())
$Global:TimerProcessing = $false
$Script:TIMER_INTERVAL_MS = 50

# Preferences
$Script:Prefs = @{
    IsLightMode    = $false
    DomainOverride = ''
    DCOverride     = ''
    OUScope        = ''
    ForceRefresh   = $false
    IncludeDisabled = $false
    IncludeUnlinked = $false
    ShowRegistryPaths = $true
    ScanMode = 'Intune'  # 'Local' | 'AD' | 'Intune' | 'Combined'
    AchievementsCollapsed = $false
}

# Logging
$Script:LogLineCount = 0
$Script:MaxLogLines  = 500
$Script:FullLogSB    = [System.Text.StringBuilder]::new(8192)
$Script:FullLogLines = 0

# Tabs definition
$Script:Tabs = @('Dashboard','GPOList','Settings','Conflicts','IntuneApps','Report','IMELogs','GPOLogs','MDMSync','Tools','AppSettings')

# ── Rotating File Log ──
try {
    if ((Test-Path $Script:LogPath) -and (Get-Item $Script:LogPath).Length -gt 2MB) {
        $prevLog = $Script:LogPath + '.prev'
        if (Test-Path $prevLog) { Remove-Item $prevLog -Force -ErrorAction SilentlyContinue }
        Rename-Item $Script:LogPath $prevLog -Force -ErrorAction SilentlyContinue
    }
    $header = "`n=== PolicyPilot v$($Script:AppVersion) - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="
    [System.IO.File]::AppendAllText($Script:LogPath, $header + "`r`n")
} catch { try { Write-DebugLog "Unhandled: $_" -Level ERROR } catch {} }

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 2: PREFERENCES LOAD / SAVE
# ═══════════════════════════════════════════════════════════════════════════════

function Load-Preferences {
    if (Test-Path $Script:PrefsPath) {
        try {
            $json = [System.IO.File]::ReadAllText($Script:PrefsPath) | ConvertFrom-Json
            foreach ($prop in $json.PSObject.Properties) {
                if ($Script:Prefs.ContainsKey($prop.Name)) { $Script:Prefs[$prop.Name] = $prop.Value }
            }
        } catch { try { Write-DebugLog "Unhandled: $_" -Level ERROR } catch {} }
    }
}
function Save-Preferences {
    try { $Script:Prefs | ConvertTo-Json -Depth 4 | Set-Content -Path $Script:PrefsPath -Encoding UTF8 -Force } catch { try { Write-DebugLog "Unhandled: $_" -Level ERROR } catch {} }
}
Load-Preferences

# ═══════════════════════════════════════════════════════════════════════════════
# ACHIEVEMENTS SYSTEM
# ═══════════════════════════════════════════════════════════════════════════════
$Script:AchievementDefs = @(
    @{ Id='first_scan';     Icon=[char]::ConvertFromUtf32(0x1F50D); Name='First Look';        Desc='Completed your first scan' }
    @{ Id='five_scans';     Icon=[char]::ConvertFromUtf32(0x1F4CA); Name='Analyst';           Desc='Completed 5 scans' }
    @{ Id='ten_scans';      Icon=[char]::ConvertFromUtf32(0x1F3C6); Name='Power Auditor';     Desc='Completed 10 scans' }
    @{ Id='first_export';   Icon=[char]::ConvertFromUtf32(0x1F4C4); Name='Documenter';        Desc='Exported your first HTML report' }
    @{ Id='csv_export';     Icon=[char]::ConvertFromUtf32(0x1F4CA); Name='Data Miner';        Desc='Exported settings to CSV' }
    @{ Id='first_snapshot'; Icon=[char]::ConvertFromUtf32(0x1F4F8); Name='Time Capsule';      Desc='Saved your first snapshot' }
    @{ Id='snapshot_compare';Icon=[char]::ConvertFromUtf32(0x1F50E); Name='Delta Detective';  Desc='Compared two snapshots' }
    @{ Id='hundred_settings';Icon=[char]::ConvertFromUtf32(0x1F4AF); Name='Century Club';     Desc='Scanned 100+ settings' }
    @{ Id='five_hundred';   Icon=[char]::ConvertFromUtf32(0x1F525); Name='Policy Guru';       Desc='Scanned 500+ settings' }
    @{ Id='conflict_found'; Icon=[char]::ConvertFromUtf32(0x26A1);  Name='Conflict Spotter';  Desc='Found your first policy conflict' }
    @{ Id='zero_conflicts'; Icon=[char]::ConvertFromUtf32(0x2728);  Name='Pristine Config';   Desc='Completed a scan with zero conflicts' }
    @{ Id='intune_mode';    Icon=[char]::ConvertFromUtf32(0x2601);  Name='Cloud Native';      Desc='Ran scan in Intune mode' }
    @{ Id='ad_mode';        Icon=[char]::ConvertFromUtf32(0x1F3E2); Name='Domain Veteran';    Desc='Ran scan in AD mode' }
    @{ Id='combined_mode';  Icon=[char]::ConvertFromUtf32(0x1F310); Name='Hybrid Hero';       Desc='Ran scan in Combined mode' }
    @{ Id='night_owl';      Icon=[char]::ConvertFromUtf32(0x1F989); Name='Night Owl';         Desc='Scanned between midnight and 5am' }
    @{ Id='early_bird';     Icon=[char]::ConvertFromUtf32(0x1F305); Name='Early Bird';        Desc='Scanned between 5am and 7am' }
    @{ Id='weekend_warrior';Icon=[char]::ConvertFromUtf32(0x1F6E1); Name='Weekend Warrior';   Desc='Scanned on a weekend' }
    @{ Id='theme_toggle';   Icon=[char]::ConvertFromUtf32(0x1F3A8); Name='Chameleon';         Desc='Toggled the theme' }
    @{ Id='speed_demon';    Icon=[char]::ConvertFromUtf32(0x26A1);  Name='Speed Demon';       Desc='Scan completed in under 5 seconds' }
    @{ Id='ten_gpos';       Icon=[char]::ConvertFromUtf32(0x1F4DA); Name='Policy Librarian';  Desc='Scanned 10+ GPOs in one scan' }
)

function Load-Achievements {
    if (Test-Path $Script:AchievementsFile) {
        try {
            $raw = Get-Content $Script:AchievementsFile -Raw | ConvertFrom-Json
            $Script:Achievements = @{}
            foreach ($p in $raw.PSObject.Properties) { $Script:Achievements[$p.Name] = $p.Value }
        } catch { $Script:Achievements = @{} }
    }
}

# ── Brush cache (perf: avoid re-creating SolidColorBrush per call) ──
$Script:GlobalBrushCache = @{}
function Get-CachedBrush {
    param([string]$ColorString)
    if ($Script:GlobalBrushCache.ContainsKey($ColorString)) {
        return $Script:GlobalBrushCache[$ColorString]
    }
    $b = [System.Windows.Media.SolidColorBrush]::new(
        [System.Windows.Media.ColorConverter]::ConvertFromString($ColorString))
    $b.Freeze()
    $Script:GlobalBrushCache[$ColorString] = $b
    return $b
}

function Save-Achievements {
    $Script:Achievements | ConvertTo-Json | Set-Content $Script:AchievementsFile -Force
}

function Unlock-Achievement {
    param([string]$Id)
    if ($Script:Achievements.ContainsKey($Id)) { return }
    $def = $Script:AchievementDefs | Where-Object { $_.Id -eq $Id }
    if (-not $def) { return }
    $Script:Achievements[$Id] = (Get-Date).ToString('o')
    Save-Achievements
    Show-Toast "Achievement: $($def.Name)" $def.Desc 'success'
    Start-ConfettiAnimation
    Render-Achievements
    Write-DebugLog "Achievement unlocked: $($def.Name)" -Level 'SUCCESS'
}

function Render-Achievements {
    if (-not $ui.pnlAchievements) { return }
    $ui.pnlAchievements.Children.Clear()
    $unlocked = 0
    foreach ($def in $Script:AchievementDefs) {
        $isUnlocked = $Script:Achievements.ContainsKey($def.Id)
        if ($isUnlocked) { $unlocked++ }
        $badge = [System.Windows.Controls.Border]::new()
        $badge.Width = 30; $badge.Height = 30
        $badge.CornerRadius = [System.Windows.CornerRadius]::new(6)
        $badge.Margin = [System.Windows.Thickness]::new(0,0,4,4)
        $lbl = [System.Windows.Controls.TextBlock]::new()
        $lbl.HorizontalAlignment = 'Center'; $lbl.VerticalAlignment = 'Center'
        if ($isUnlocked) {
            $badge.Background = $Window.FindResource('ThemeSelectedBg')
            $badge.BorderBrush = $Window.FindResource('ThemeWarning')
            $badge.BorderThickness = [System.Windows.Thickness]::new(1)
            $badge.ToolTip = "$($def.Name): $($def.Desc)"
            $lbl.Text = $def.Icon; $lbl.FontSize = 14
            $lbl.Foreground = $Window.FindResource('ThemeTextPrimary')
            $lbl.FontFamily = [System.Windows.Media.FontFamily]::new('Segoe UI Emoji')
        } else {
            $badge.Background = $Window.FindResource('ThemeDeepBg')
            $badge.BorderBrush = $Window.FindResource('ThemeBorder')
            $badge.BorderThickness = [System.Windows.Thickness]::new(1)
            $badge.ToolTip = '???'
            $lbl.Text = '?'; $lbl.FontSize = 12
            $lbl.Foreground = $Window.FindResource('ThemeTextDisabled')
        }
        $badge.Child = $lbl
        [void]$ui.pnlAchievements.Children.Add($badge)
    }
    if ($ui.lblAchievementCount) { $ui.lblAchievementCount.Text = "$unlocked/$($Script:AchievementDefs.Count)" }
}

function Start-ConfettiAnimation {
    if ($Script:AnimationsDisabled) { return }
    if (-not $ui.cnvConfetti) { return }
    if ($ui.cnvConfetti.Visibility -eq 'Visible') { return }
    $w = $Window.ActualWidth; $h = $Window.ActualHeight
    if ($w -le 0 -or $h -le 0) { return }
    $ui.cnvConfetti.Children.Clear()
    $ui.cnvConfetti.Visibility = 'Visible'
    $colors = @('#FF4444','#FFD700','#00C853','#60CDFF','#0078D4','#B388FF','#FF6D00','#E040FB')
    $r = [System.Random]::new()
    $bc = [System.Windows.Media.BrushConverter]::new()
    for ($i = 0; $i -lt $Script:CONFETTI_COUNT; $i++) {
        $sz = $r.Next(4,10)
        $rect = [System.Windows.Shapes.Rectangle]::new()
        $rect.Width = $sz; $rect.Height = $sz * ($r.NextDouble() * 1.5 + 0.5)
        $rect.Fill = $bc.ConvertFromString($colors[$r.Next($colors.Count)])
        $rect.RadiusX = if ($r.Next(3) -eq 0) { $sz/2 } else { 1 }; $rect.RadiusY = $rect.RadiusX
        $rect.Opacity = 0.9
        $rect.RenderTransform = [System.Windows.Media.RotateTransform]::new($r.Next(360))
        $x0 = $r.NextDouble() * $w
        [System.Windows.Controls.Canvas]::SetLeft($rect, $x0)
        [System.Windows.Controls.Canvas]::SetTop($rect, -20)
        [void]$ui.cnvConfetti.Children.Add($rect)
        $fall = [System.Windows.Media.Animation.DoubleAnimation]::new()
        $fall.From = $r.Next(-40,-10); $fall.To = $h + 20
        $fall.Duration = [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds($r.Next(2000,4500)))
        $ease = [System.Windows.Media.Animation.CubicEase]::new(); $ease.EasingMode = 'EaseIn'
        $fall.EasingFunction = $ease
        $drift = [System.Windows.Media.Animation.DoubleAnimation]::new()
        $drift.From = $x0; $drift.To = $x0 + $r.Next(-120,120); $drift.Duration = $fall.Duration
        $spin = [System.Windows.Media.Animation.DoubleAnimation]::new()
        $spin.From = 0; $spin.To = $r.Next(-720,720); $spin.Duration = $fall.Duration
        $rect.BeginAnimation([System.Windows.Controls.Canvas]::TopProperty, $fall)
        $rect.BeginAnimation([System.Windows.Controls.Canvas]::LeftProperty, $drift)
        $rect.RenderTransform.BeginAnimation([System.Windows.Media.RotateTransform]::AngleProperty, $spin)
    }
    $ct = [System.Windows.Threading.DispatcherTimer]::new()
    $ct.Interval = [TimeSpan]::FromMilliseconds(5000)
    $ct.Tag = $ui.cnvConfetti
    $ct.Add_Tick({ try { $this.Tag.Children.Clear(); $this.Tag.Visibility = 'Collapsed'; $this.Stop() } catch { try { Write-DebugLog "Unhandled: $_" -Level ERROR } catch {} } })
    $ct.Start()
}

function Start-StatusPulse {
    if (-not $ui.statusBarDot) { return }
    $ui.statusBarDot.Fill = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#FFB900')
    if ($Script:AnimationsDisabled) { return }
    $pulse = [System.Windows.Media.Animation.DoubleAnimation]::new()
    $pulse.From = 1.0; $pulse.To = 0.3
    $pulse.Duration = [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(600))
    $pulse.AutoReverse = $true
    $pulse.RepeatBehavior = [System.Windows.Media.Animation.RepeatBehavior]::Forever
    $ui.statusBarDot.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $pulse)
}
function Stop-StatusPulse {
    if (-not $ui.statusBarDot) { return }
    $ui.statusBarDot.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $null)
    $ui.statusBarDot.Opacity = 1.0
}

function Check-ScanAchievements {
    param([int]$SettingCount, [int]$GpoCount, [int]$ConflictCount, [string]$ScanMode, [double]$ElapsedSec)
    Unlock-Achievement 'first_scan'
    # Count total scans
    $scanCount = ($Script:Achievements.Keys | Where-Object { $_ -eq 'first_scan' }).Count
    $totalScans = 1
    if ($Script:Achievements.ContainsKey('_scan_count')) { $totalScans = [int]$Script:Achievements['_scan_count'] + 1 }
    $Script:Achievements['_scan_count'] = $totalScans
    if ($totalScans -ge 5)  { Unlock-Achievement 'five_scans' }
    if ($totalScans -ge 10) { Unlock-Achievement 'ten_scans' }
    # Setting milestones
    if ($SettingCount -ge 100) { Unlock-Achievement 'hundred_settings' }
    if ($SettingCount -ge 500) { Unlock-Achievement 'five_hundred' }
    if ($GpoCount -ge 10)     { Unlock-Achievement 'ten_gpos' }
    # Conflicts
    if ($ConflictCount -gt 0)  { Unlock-Achievement 'conflict_found' }
    if ($ConflictCount -eq 0 -and $SettingCount -gt 0) { Unlock-Achievement 'zero_conflicts' }
    # Scan mode
    if ($ScanMode -match 'Intune')   { Unlock-Achievement 'intune_mode' }
    if ($ScanMode -match '^AD')      { Unlock-Achievement 'ad_mode' }
    if ($ScanMode -match 'Combined') { Unlock-Achievement 'combined_mode' }
    # Time-based
    $hour = (Get-Date).Hour
    if ($hour -ge 0 -and $hour -lt 5) { Unlock-Achievement 'night_owl' }
    if ($hour -ge 5 -and $hour -lt 7) { Unlock-Achievement 'early_bird' }
    if ((Get-Date).DayOfWeek -in @('Saturday','Sunday')) { Unlock-Achievement 'weekend_warrior' }
    # Speed demon
    if ($ElapsedSec -gt 0 -and $ElapsedSec -lt 5) { Unlock-Achievement 'speed_demon' }
    Save-Achievements
}

Load-Achievements

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 3: WRITE-DEBUGLOG
# ═══════════════════════════════════════════════════════════════════════════════

function Write-DebugLog {
    param(
        [string]$Message,
        [ValidateSet('INFO','SUCCESS','WARN','ERROR','DEBUG','STEP','SYSTEM')]
        [string]$Level = 'INFO'
    )
    $timestamp = Get-Date -Format 'HH:mm:ss.fff'
    $line = "[$timestamp] [$Level] $Message"
    Write-Host $line -ForegroundColor DarkGray
    try { [System.IO.File]::AppendAllText($Script:LogPath, $line + "`r`n") } catch {}
    if ($Script:BgLogActive -and $Script:BgLogStream) { try { $Script:BgLogStream.WriteLine($line) } catch {} }
    $Script:FullLogLines++
    if ($Script:FullLogLines -gt ($Script:MaxLogLines * 2)) {
        $text = $Script:FullLogSB.ToString()
        $nl = $text.IndexOf("`n")
        if ($nl -ge 0) {
            [void]$Script:FullLogSB.Clear()
            [void]$Script:FullLogSB.Append($text.Substring($nl + 1))
            $Script:FullLogLines--
        }
    }
    [void]$Script:FullLogSB.AppendLine($line)

    # Write to console panel if available
    if ($Global:ConsoleReady -and $ui.paraLog) {
        try {
            $Color = if ($Script:Prefs.IsLightMode) {
                switch ($Level) {
                    'ERROR'   { '#CC0000' }
                    'WARN'    { '#B86E00' }
                    'SUCCESS' { '#008A2E' }
                    'STEP'    { '#0078D4' }
                    'SYSTEM'  { '#666666' }
                    'DEBUG'   { '#8888AA' }
                    default   { '#444444' }
                }
            } else {
                switch ($Level) {
                    'ERROR'   { '#FF4040' }
                    'WARN'    { '#FF9100' }
                    'SUCCESS' { '#16C60C' }
                    'STEP'    { '#60CDFF' }
                    'SYSTEM'  { '#7A7A84' }
                    'DEBUG'   { '#B8860B' }
                    default   { '#888888' }
                }
            }
            if ($ui.paraLog.Inlines.Count -gt 0) {
                [void]$ui.paraLog.Inlines.Add([System.Windows.Documents.LineBreak]::new())
            }
            $run = [System.Windows.Documents.Run]::new($line)
            $run.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($Color)
            [void]$ui.paraLog.Inlines.Add($run)
            $ui.logScroller.ScrollToEnd()
        } catch { Write-Host "paraLog error: $_" -ForegroundColor Red }
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 3a: UTILITY - NUMBERED-LIST FORMATTER
# ═══════════════════════════════════════════════════════════════════════════════
# Splits concatenated numbered-list registry values (e.g. "1url.com2url2.com3url3.com")
# into readable "1. url.com\n2. url2.com\n3. url3.com" format.
# Handles Chrome extension IDs, URLs, ext;url forcelist pairs, and file paths.
# Uses multi-candidate selection with average-length heuristic for disambiguation
# when values themselves contain the delimiter number followed by letters.
function Format-NumberedList([string]$raw) {
    if (-not $raw -or $raw.Length -lt 4) { return $raw }
    # Strip U+F000 PUA separator used by ADMX ingested list policies
    $raw = $raw -replace '\uF000',''
    # Must start with "1" followed by a non-digit
    if ($raw[0] -ne '1' -or ($raw.Length -gt 1 -and [char]::IsDigit($raw[1]))) { return $raw }

    $items     = [System.Collections.Generic.List[string]]::new()
    $remaining = $raw.Substring(1)  # consume the "1" prefix
    $expected  = 2
    $lengthSum = 0
    $pfxRe     = $null   # item-start pattern inferred from item 1

    while ($remaining.Length -gt 0) {
        $idxStr  = "$expected"
        # Match the expected sequential number followed by 2+ letters, backslash,
        # asterisk, semicolon, or an http(s) scheme start.
        $pattern = "$([regex]::Escape($idxStr))(?=[a-zA-Z]{2}|[\*\\;]|https?://)"
        $allM    = [regex]::Matches($remaining, $pattern)

        $best = $null
        if ($allM.Count -eq 1 -and $allM[0].Index -ge 1) {
            $best = $allM[0]
        } elseif ($allM.Count -gt 1) {
            # 1) Prefix consistency: prefer candidate whose remaining text
            #    starts like previous items (e.g. https://, domain, ext-ID).
            if ($pfxRe) {
                foreach ($cm in $allM) {
                    if ($cm.Index -lt 1) { continue }
                    $after = $remaining.Substring($cm.Index + $cm.Length)
                    if ($after -match $pfxRe) { $best = $cm; break }
                }
            }
            # 2) Average-length proximity (needs 2+ items).
            if (-not $best -and $items.Count -ge 2) {
                $avg = $lengthSum / $items.Count
                $bestDist = [int]::MaxValue
                foreach ($cm in $allM) {
                    if ($cm.Index -lt 1) { continue }
                    $dist = [Math]::Abs($cm.Index - $avg)
                    if ($dist -lt $bestDist) { $bestDist = $dist; $best = $cm }
                }
            }
            # 3) Fallback: first valid match.
            if (-not $best) {
                foreach ($cm in $allM) { if ($cm.Index -ge 1) { $best = $cm; break } }
            }
        }

        if ($best) {
            $item = $remaining.Substring(0, $best.Index)
            $items.Add($item)
            $lengthSum += $item.Length
            $remaining  = $remaining.Substring($best.Index + $best.Length)
            $expected++
            # After item 1, infer the common item-start pattern for disambiguation
            if ($items.Count -eq 1) {
                if     ($item -match '^https?://')            { $pfxRe = '^https?://' }
                elseif ($item -match '^[a-z]{20,}')           { $pfxRe = '^[a-z]{10,}' }
                elseif ($item -match '^[a-z][\w.-]*\.[a-z]') { $pfxRe = '^[a-z][\w.-]*\.' }
            }
        } else {
            $items.Add($remaining)
            break
        }
    }

    if ($items.Count -ge 2) {
        $result = for ($i = 0; $i -lt $items.Count; $i++) { "$($i + 1). $($items[$i])" }
        return ($result -join "`n")
    }
    # Single-item ADMX list: strip the leading index (e.g. "1extensions" → "extensions")
    if ($items.Count -eq 1) { return $items[0] }
    return $raw
}

# ===============================================================================
# SECTION 3b: BACKGROUND WORK ENGINE
# ===============================================================================

function Start-BackgroundWork {
    param(
        [ScriptBlock]$Work,
        [ScriptBlock]$OnComplete,
        [hashtable]$Variables = @{},
        [hashtable]$Context = @{}
    )

    $ISS = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $ISS.ExecutionPolicy = [Microsoft.PowerShell.ExecutionPolicy]::Bypass

    $RS = [RunspaceFactory]::CreateRunspace($ISS)
    $RS.ApartmentState = [System.Threading.ApartmentState]::STA
    $RS.ThreadOptions  = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
    $RS.Open()

    $PS = [PowerShell]::Create()
    $PS.Runspace = $RS

    foreach ($k in $Variables.Keys) {
        $PS.Runspace.SessionStateProxy.SetVariable($k, $Variables[$k])
    }

    $PS.AddScript($Work)

    Write-DebugLog "BgWork: launching runspace (vars=$($Variables.Keys -join ',' ))" -Level DEBUG
    $Async = $PS.BeginInvoke()

    [void]$Global:BgJobs.Add(@{
        PS          = $PS
        Runspace    = $RS
        AsyncResult = $Async
        OnComplete  = $OnComplete
        Context     = $Context
        StartedAt   = (Get-Date)
    })

    Write-DebugLog "BgWork: queued job #$($Global:BgJobs.Count)" -Level DEBUG
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 4: PREREQUISITE CHECKS
# ═══════════════════════════════════════════════════════════════════════════════

function Test-Prerequisites {
    Write-DebugLog 'Checking prerequisites...' -Level STEP
    $results = [System.Collections.Generic.List[string]]::new()
    $allOk = $true
    $mode = $Script:Prefs.ScanMode

    [void]$results.Add("Scan Mode: $mode")
    [void]$results.Add("")

    if ($mode -eq 'AD') {
        # AD mode requires RSAT GroupPolicy module
        $gpMod = Get-Module -ListAvailable -Name GroupPolicy -ErrorAction SilentlyContinue
        if ($gpMod) {
            [void]$results.Add("[OK]  GroupPolicy module v$($gpMod.Version)")
        } else {
            [void]$results.Add("[FAIL] GroupPolicy module NOT found")
            [void]$results.Add("       Install via Features on Demand:")
            [void]$results.Add("       Add-WindowsCapability -Online -Name Rsat.GroupPolicy.Management.Tools~~~~0.0.1.0")
            [void]$results.Add("       Or: Settings > System > Optional Features > RSAT: Group Policy Management Tools")
            [void]$results.Add("")
            [void]$results.Add("       [TIP] PolicyPilot can install this for you - click 'Install RSAT GP Tools' below")
            $allOk = $false
            $Script:RsatMissing = $true
        }

        try {
            $domain = if ($Script:Prefs.DomainOverride) { $Script:Prefs.DomainOverride } else { [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name }
            [void]$results.Add("[OK]  Domain: $domain")
        } catch {
            [void]$results.Add("[FAIL] Cannot reach Active Directory domain")
            [void]$results.Add("       Ensure this machine is domain-joined or specify a domain override")
            $allOk = $false
        }

        if ($gpMod) {
            try {
                Import-Module GroupPolicy -ErrorAction Stop
                $testGPO = Get-GPO -All -ErrorAction Stop | Select-Object -First 1
                if ($testGPO) {
                    [void]$results.Add("[OK]  GPO read access confirmed")
                } else {
                    [void]$results.Add("[WARN] Get-GPO returned 0 GPOs - domain may be empty")
                }
            } catch {
                [void]$results.Add("[FAIL] Cannot read GPOs: $($_.Exception.Message)")
                $allOk = $false
            }
        }
    } elseif ($mode -eq 'Intune') {
        $graphMod = Get-Module -ListAvailable Microsoft.Graph.DeviceManagement -ErrorAction SilentlyContinue
        $graphLegacy = Get-Module -ListAvailable Microsoft.Graph.Intune -ErrorAction SilentlyContinue
        if ($graphMod -or $graphLegacy) {
            [void]$results.Add("[OK]  Microsoft.Graph module found")
        } else {
            [void]$results.Add("[FAIL] Microsoft.Graph.DeviceManagement not found")
            [void]$results.Add("  Install: Install-Module Microsoft.Graph -Scope CurrentUser")
            $allOk = $false
        }
        [void]$results.Add("")
        [void]$results.Add("Intune mode uses Microsoft Graph to read")
        [void]$results.Add("device configuration, compliance policies,")
        [void]$results.Add("and settings catalog.")
    } else {
        # Local RSoP mode - just needs gpresult.exe (built into Windows)
        $gpresult = Get-Command gpresult.exe -ErrorAction SilentlyContinue
        if ($gpresult) {
            [void]$results.Add("[OK]  gpresult.exe found")
        } else {
            [void]$results.Add("[FAIL] gpresult.exe not found")
            $allOk = $false
        }

        try {
            $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
            if ($cs.PartOfDomain) {
                [void]$results.Add("[OK]  Domain-joined: $($cs.Domain)")
            } else {
                [void]$results.Add("[WARN] Not domain-joined - only local policies will be shown")
            }
        } catch {
            [void]$results.Add("[WARN] Cannot determine domain membership")
        }

        [void]$results.Add("")
        [void]$results.Add("Local mode scans policies applied to THIS machine")
        [void]$results.Add("using gpresult (no RSAT required).")
    }

    $Script:PrereqsMet = $allOk
    $statusText = ($results -join "`n")
    Write-DebugLog "Prerequisites ($mode): $( if ($allOk) {'PASSED'} else {'FAILED'} )" -Level $(if ($allOk) {'SUCCESS'} else {'ERROR'})
    return @{ Passed = $allOk; Details = $statusText }
}

function Install-RsatGpTools {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Show-Toast 'Admin Required' 'RSAT installation requires running as Administrator' 'error'
        return
    }
    Show-Toast 'Installing RSAT' 'Installing Group Policy Management Tools via Features on Demand...' 'info'
    Write-DebugLog 'Installing RSAT GroupPolicy tools via Add-WindowsCapability...' -Level STEP
    Start-BackgroundWork -Work {
        try {
            $result = Add-WindowsCapability -Online -Name 'Rsat.GroupPolicy.Management.Tools~~~~0.0.1.0' -ErrorAction Stop
            return @{ Success = $true; State = $result.State; RestartNeeded = $result.RestartNeeded }
        } catch {
            return @{ Success = $false; Error = $_.Exception.Message }
        }
    } -OnComplete {
        param($Results, $Errors)
        $r = $Results | Select-Object -First 1
        if ($r.Success) {
            $msg = "RSAT GP Tools installed (State: $($r.State))"
            if ($r.RestartNeeded) { $msg += ' - restart may be required' }
            Write-DebugLog $msg -Level SUCCESS
            Show-Toast 'RSAT Installed' $msg 'success'
            # Re-check prerequisites
            $result = Test-Prerequisites
            if ($ui.PrereqDetailStatus) { $ui.PrereqDetailStatus.Text = $result.Details }
            if ($ui.PrereqStatus)       { $ui.PrereqStatus.Text = $result.Details }
        } else {
            $errMsg = if ($r.Error) { $r.Error } elseif ($Errors.Count -gt 0) { $Errors[0].ToString() } else { 'Unknown error' }
            Write-DebugLog "RSAT install failed: $errMsg" -Level ERROR
            Show-Toast 'Install Failed' "RSAT installation failed: $errMsg" 'error'
        }
    }.GetNewClosure() -Variables @{} -Context @{ Name = 'InstallRSAT' }
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 5: GPO SCANNING ENGINE
# ═══════════════════════════════════════════════════════════════════════════════

function Invoke-GPOScan {
    Write-DebugLog 'Starting GPO scan...' -Level STEP

    # Prepare domain/DC params
    $gpoParams = @{}
    $DomainOvr = $Script:Prefs.DomainOverride
    $DcOvr     = $Script:Prefs.DCOverride
    if ($DomainOvr) { $gpoParams['Domain'] = $DomainOvr }
    if ($DcOvr)     { $gpoParams['Server'] = $DcOvr }

    # ── LDAP-based GPO enumeration (replaces slow Get-GPO -All) ──
    $rawGPOs = @()
    try {
        $domain = if ($DomainOvr) { $DomainOvr }
                  else { try { [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name } catch { '(unknown)' } }
        $ldapRoot = if ($DcOvr -and $DomainOvr) { "LDAP://$DcOvr/CN=Policies,CN=System,$(($DomainOvr.Split('.') | ForEach-Object { "DC=$_" }) -join ',')" }
                    elseif ($DomainOvr) { "LDAP://CN=Policies,CN=System,$(($DomainOvr.Split('.') | ForEach-Object { "DC=$_" }) -join ',')" }
                    elseif ($DcOvr) { $rootDse = [ADSI]"LDAP://$DcOvr/RootDSE"; "LDAP://$DcOvr/CN=Policies,CN=System,$($rootDse.defaultNamingContext)" }
                    else { $rootDse = [ADSI]'LDAP://RootDSE'; "LDAP://CN=Policies,CN=System,$($rootDse.defaultNamingContext)" }
        $searchRoot = [ADSI]$ldapRoot
        $searcher = [System.DirectoryServices.DirectorySearcher]::new($searchRoot)
        $searcher.Filter = '(objectClass=groupPolicyContainer)'
        $searcher.PageSize = 100
        @('displayName','name','flags','versionNumber','whenCreated','whenChanged','gPCWMIFilter') | ForEach-Object { [void]$searcher.PropertiesToLoad.Add($_) }
        $ldapResults = $searcher.FindAll()
        $rawList = [System.Collections.Generic.List[object]]::new()
        foreach ($entry in $ldapResults) {
            $props = $entry.Properties
            $dn = $props['displayname']
            $displayName = if ($dn -and $dn.Count -gt 0) { "$($dn[0])" } else { '' }
            $guidRaw = if ($props['name'] -and $props['name'].Count -gt 0) { "$($props['name'][0])" } else { '' }
            $guid = $guidRaw.Trim('{','}')
            $flags = if ($props['flags'] -and $props['flags'].Count -gt 0) { [int]$props['flags'][0] } else { 0 }
            $ver = if ($props['versionNumber'] -and $props['versionNumber'].Count -gt 0) { [int]$props['versionNumber'][0] } else { 0 }
            $whenCreated = if ($props['whencreated'] -and $props['whencreated'].Count -gt 0) { [datetime]$props['whencreated'][0] } else { [datetime]::MinValue }
            $whenChanged = if ($props['whenchanged'] -and $props['whenchanged'].Count -gt 0) { [datetime]$props['whenchanged'][0] } else { [datetime]::MinValue }
            $userVer = ($ver -shr 16) -band 0xFFFF
            $compVer = $ver -band 0xFFFF
            $gpoStatus = switch ($flags) { 1 { 'UserSettingsDisabled' } 2 { 'ComputerSettingsDisabled' } 3 { 'AllSettingsDisabled' } default { 'AllSettingsEnabled' } }
            $wmiFilterName = ''
            $wmiRef = if ($props['gpcwmifilter'] -and $props['gpcwmifilter'].Count -gt 0) { "$($props['gpcwmifilter'][0])" } else { $null }
            if ($wmiRef -and $wmiRef -match ';') { $wmiFilterName = ($wmiRef -split ';')[1] }
            if (-not $displayName -or -not $guid) { continue }
            [void]$rawList.Add([PSCustomObject]@{
                DisplayName      = $displayName
                Id               = [Guid]$guid
                GpoStatus        = $gpoStatus
                CreationTime     = $whenCreated
                ModificationTime = $whenChanged
                WmiFilter        = if ($wmiFilterName) { [PSCustomObject]@{ Name = $wmiFilterName } } else { $null }
                User             = [PSCustomObject]@{ DSVersion = $userVer }
                Computer         = [PSCustomObject]@{ DSVersion = $compVer }
            })
        }
        $ldapResults.Dispose()
        $rawGPOs = @($rawList)
    } catch {
        Write-DebugLog "LDAP GPO enumeration failed: $($_.Exception.Message)" -Level ERROR
        return $null
    }

    Write-DebugLog "Found $($rawGPOs.Count) GPOs via LDAP" -Level INFO

    # Auto-discover DC for cross-domain scans (reuse single connection)
    if ($DomainOvr -and -not $DcOvr) {
        try {
            $ctx = [System.DirectoryServices.ActiveDirectory.DirectoryContext]::new('Domain', $DomainOvr)
            $targetDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($ctx)
            $dc = $targetDomain.FindDomainController()
            $DcOvr = $dc.Name
            $gpoParams['Server'] = $DcOvr
            Write-DebugLog "Using DC: $DcOvr for cross-domain calls" -Level SUCCESS
        } catch {
            Write-DebugLog "DC auto-discovery failed: $($_.Exception.Message)" -Level WARN
        }
    }

    $allSettingsList = [System.Collections.Generic.List[PSCustomObject]]::new()
    $gpoRecords     = [System.Collections.Generic.List[PSCustomObject]]::new()
    $ns = @{ gp = 'http://www.microsoft.com/GroupPolicy/Settings'
             types = 'http://www.microsoft.com/GroupPolicy/Settings/Registry'
             q1 = 'http://www.microsoft.com/GroupPolicy/Settings/Registry'
             sec = 'http://www.microsoft.com/GroupPolicy/Settings/Security' }

    # ── Batch LDAP: discover GPO link locations ──
    $gpoLinkMap = @{}
    try {
        $baseDN = if ($DomainOvr) { ($DomainOvr.Split('.') | ForEach-Object { "DC=$_" }) -join ',' }
                  else { $rootDse2 = [ADSI]'LDAP://RootDSE'; "$($rootDse2.defaultNamingContext)" }
        $linkLdap = if ($DcOvr) { "LDAP://$DcOvr/$baseDN" } else { "LDAP://$baseDN" }
        $linkSearcher = [System.DirectoryServices.DirectorySearcher]::new([ADSI]$linkLdap)
        $linkSearcher.Filter = '(gPLink=*)'
        $linkSearcher.PageSize = 200
        @('distinguishedName','gPLink') | ForEach-Object { [void]$linkSearcher.PropertiesToLoad.Add($_) }
        $linkResults = $linkSearcher.FindAll()
        foreach ($lr in $linkResults) {
            $lrDN = "$($lr.Properties['distinguishedname'][0])"
            $gPLinkVal = "$($lr.Properties['gplink'][0])"
            $linkMatches = [regex]::Matches($gPLinkVal, '\[LDAP://[Cc][Nn]=\{([0-9a-fA-F\-]+)\}[^]]*;(\d+)\]')
            foreach ($m in $linkMatches) {
                $linkedGuid = $m.Groups[1].Value.ToUpper()
                $flagBits = [int]$m.Groups[2].Value
                $enforced = ($flagBits -band 2) -ne 0
                $disabled = ($flagBits -band 1) -ne 0
                if ($disabled) { continue }
                if (-not $gpoLinkMap.ContainsKey($linkedGuid)) { $gpoLinkMap[$linkedGuid] = [System.Collections.Generic.List[object]]::new() }
                [void]$gpoLinkMap[$linkedGuid].Add(@{ SOMPath = $lrDN; Enforced = $enforced })
            }
        }
        $linkResults.Dispose()
        Write-DebugLog "LDAP link discovery: $($gpoLinkMap.Count) GPOs have links" -Level INFO
    } catch {
        Write-DebugLog "LDAP link discovery failed: $($_.Exception.Message)" -Level WARN
    }

    # Skip GPOs with no settings (DSVersion=0 or AllSettingsDisabled)
    $linkedGPOs = @($rawGPOs | Where-Object {
        ($_.User.DSVersion -gt 0 -or $_.Computer.DSVersion -gt 0) -and
        ($_.GpoStatus -ne 'AllSettingsDisabled')
    })
    $skipped = $rawGPOs.Count - $linkedGPOs.Count
    if ($skipped -gt 0) { Write-DebugLog "Skipping $skipped empty/disabled GPOs" -Level INFO }
    $total = $linkedGPOs.Count
    $idx = 0

    foreach ($gpo in $linkedGPOs) {
        $idx++
        $pct = 15 + [int](($idx / $total) * 80)
        if ($idx % 25 -eq 0 -or $idx -eq 1 -or $idx -eq $total) { Write-DebugLog "GPO $idx/$total ($pct%): $($gpo.DisplayName)" -Level INFO }

        # ── SYSVOL direct read (fast SMB), fall back to Get-GPOReport (slow cmdlet) ──
        $xmlText = $null
        $sysvolBase = "\\\\$domain\\SYSVOL\\$domain\\Policies\\{$($gpo.Id.ToString())}".ToUpper()
        try {
            $polXml = [System.Text.StringBuilder]::new()
            [void]$polXml.AppendLine('<?xml version="1.0" encoding="utf-8"?>')
            [void]$polXml.AppendLine('<GPO xmlns="http://www.microsoft.com/GroupPolicy/Settings" xmlns:types="http://www.microsoft.com/GroupPolicy/Types">')
            [void]$polXml.AppendLine("<Identifier><Identifier>{$($gpo.Id.ToString())}</Identifier></Identifier>")
            [void]$polXml.AppendLine("<Name>$([System.Security.SecurityElement]::Escape($gpo.DisplayName))</Name>")
            foreach ($scope in @('Computer','User')) {
                $polPath = [IO.Path]::Combine($sysvolBase, $(if ($scope -eq 'Computer') {'Machine'} else {'User'}), 'registry.pol')
                if (-not [IO.File]::Exists($polPath)) { continue }
                [void]$polXml.AppendLine("<$scope>")
                [void]$polXml.AppendLine('<ExtensionData><Extension xmlns:q1="http://www.microsoft.com/GroupPolicy/Settings/Registry" xsi:type="q1:RegistrySettings" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">')
                try {
                    $polBytes = [IO.File]::ReadAllBytes($polPath)
                    if ($polBytes.Length -gt 8) {
                        $pos = 8
                        while ($pos -lt $polBytes.Length - 2) {
                            if ($polBytes[$pos] -ne 0x5B -or $polBytes[$pos+1] -ne 0x00) { $pos++; continue }
                            $pos += 2
                            $keyStart = $pos
                            while ($pos -lt $polBytes.Length - 3 -and -not ($polBytes[$pos] -eq 0 -and $polBytes[$pos+1] -eq 0 -and $polBytes[$pos+2] -eq 0x3B -and $polBytes[$pos+3] -eq 0)) { $pos += 2 }
                            $regKey = [System.Text.Encoding]::Unicode.GetString($polBytes, $keyStart, $pos - $keyStart)
                            $pos += 4
                            $vnStart = $pos
                            while ($pos -lt $polBytes.Length - 3 -and -not ($polBytes[$pos] -eq 0 -and $polBytes[$pos+1] -eq 0 -and $polBytes[$pos+2] -eq 0x3B -and $polBytes[$pos+3] -eq 0)) { $pos += 2 }
                            $valName = [System.Text.Encoding]::Unicode.GetString($polBytes, $vnStart, $pos - $vnStart)
                            $pos += 4
                            if ($pos + 4 -gt $polBytes.Length) { break }
                            $regType = [BitConverter]::ToUInt32($polBytes, $pos); $pos += 4
                            $pos += 2
                            if ($pos + 4 -gt $polBytes.Length) { break }
                            $dataSize = [BitConverter]::ToUInt32($polBytes, $pos); $pos += 4
                            $dataVal = ''
                            if ($dataSize -gt 0 -and $pos + $dataSize -le $polBytes.Length) {
                                switch ($regType) {
                                    1 { $dataVal = [System.Text.Encoding]::Unicode.GetString($polBytes, $pos, $dataSize).TrimEnd("`0") }
                                    4 { if ($dataSize -ge 4) { $dataVal = [BitConverter]::ToUInt32($polBytes, $pos).ToString() } }
                                    default { $dataVal = [BitConverter]::ToString($polBytes, $pos, [math]::Min($dataSize, 64)) }
                                }
                                $pos += $dataSize
                            }
                            if ($pos -lt $polBytes.Length - 1 -and $polBytes[$pos] -eq 0x5D -and $polBytes[$pos+1] -eq 0) { $pos += 2 }
                            $typeStr = switch ($regType) { 1 {'REG_SZ'} 2 {'REG_EXPAND_SZ'} 4 {'REG_DWORD'} 7 {'REG_MULTI_SZ'} 11 {'REG_QWORD'} 3 {'REG_BINARY'} default {"Type_$regType"} }
                            [void]$polXml.AppendLine('<q1:RegistrySetting>')
                            [void]$polXml.AppendLine("<q1:KeyPath>$([System.Security.SecurityElement]::Escape($regKey))</q1:KeyPath>")
                            [void]$polXml.AppendLine("<q1:ValueName>$([System.Security.SecurityElement]::Escape($valName))</q1:ValueName>")
                            [void]$polXml.AppendLine("<q1:Value>$([System.Security.SecurityElement]::Escape($dataVal))</q1:Value>")
                            [void]$polXml.AppendLine("<q1:Type>$typeStr</q1:Type>")
                            [void]$polXml.AppendLine('</q1:RegistrySetting>')
                        }
                    }
                } catch {
                    Write-DebugLog ".pol parse error for $scope in $($gpo.DisplayName): $($_.Exception.Message)" -Level WARN
                }
                [void]$polXml.AppendLine('</Extension></ExtensionData>')
                [void]$polXml.AppendLine("</$scope>")
            }
            [void]$polXml.AppendLine('</GPO>')
            $xmlText = $polXml.ToString()
        } catch {
            # SYSVOL read failed — fall back to Get-GPOReport
            try {
                $reportParams = @{ Guid = $gpo.Id; ReportType = 'Xml' }
                if ($DomainOvr) { $reportParams['Domain'] = $DomainOvr }
                if ($DcOvr)     { $reportParams['Server'] = $DcOvr }
                $xmlText = Get-GPOReport @reportParams -ErrorAction Stop
            } catch {
                Write-DebugLog "Both SYSVOL and GPOReport failed for '$($gpo.DisplayName)': $($_.Exception.Message)" -Level WARN
            }
        }

        # Parse link info & settings from XML
        $linkLocations = [System.Collections.Generic.List[string]]::new()
        $linkDetails   = [System.Collections.Generic.List[PSCustomObject]]::new()
        $settingCount = 0
        $isEnforced   = $false

        # Populate link data from pre-fetched LDAP gPLink map
        $gpoGuidUpper = $gpo.Id.ToString().Trim('{','}').ToUpper()
        if ($gpoLinkMap.ContainsKey($gpoGuidUpper)) {
            foreach ($lnk in $gpoLinkMap[$gpoGuidUpper]) {
                [void]$linkLocations.Add($lnk.SOMPath)
                if ($lnk.Enforced) { $isEnforced = $true }
                [void]$linkDetails.Add([PSCustomObject]@{
                    SOMPath   = $lnk.SOMPath
                    Enforced  = $lnk.Enforced
                    LinkOrder = 0
                    Enabled   = $true
                })
            }
        }

        if ($xmlText) {
            try {
                $xdoc = [xml]$xmlText
                $nsMgr = [System.Xml.XmlNamespaceManager]::new($xdoc.NameTable)
                $nsMgr.AddNamespace('gp', 'http://www.microsoft.com/GroupPolicy/Settings')

                # Parse settings from Computer and User sections
                foreach ($scope in @('Computer','User')) {
                    $scopeNode = $xdoc.SelectSingleNode("//gp:$scope", $nsMgr)
                    if (-not $scopeNode) { continue }
                    $enabledNode = $scopeNode.SelectSingleNode('gp:Enabled', $nsMgr)
                    if ($enabledNode -and $enabledNode.InnerText -eq 'false') { continue }

                    # Extension data
                    $extNodes = $scopeNode.SelectNodes('gp:ExtensionData/gp:Extension', $nsMgr)
                    foreach ($ext in $extNodes) {
                        # Administrative Template policies (q1:Policy nodes)
                        $policies = $ext.SelectNodes('q1:Policy', $nsMgr)
                        if (-not $policies -or $policies.Count -eq 0) {
                            # Try without namespace prefix (some GPOs use different schema)
                            $policies = $ext.ChildNodes | Where-Object { $_.LocalName -eq 'Policy' }
                        }
                        foreach ($pol in $policies) {
                            $settingCount++
                            $polName = $pol.SelectSingleNode('*[local-name()="Name"]')
                            $polState = $pol.SelectSingleNode('*[local-name()="State"]')
                            $polCategory = $pol.SelectSingleNode('*[local-name()="Category"]')
                            $polExplain = $pol.SelectSingleNode('*[local-name()="Explain"]')

                            # Collect sub-setting values
                            $subValues = [System.Collections.Generic.List[string]]::new()
                            foreach ($child in $pol.ChildNodes) {
                                $localName = $child.LocalName
                                if ($localName -in @('Name','State','Category','Supported','Explain')) { continue }
                                $subName  = $child.SelectSingleNode('*[local-name()="Name"]')
                                $subValue = $child.SelectSingleNode('*[local-name()="Value"]')
                                $subState = $child.SelectSingleNode('*[local-name()="State"]')
                                if ($subName -and ($subValue -or $subState)) {
                                    $val = if ($subValue) {
                                        $innerName = $subValue.SelectSingleNode('*[local-name()="Name"]')
                                        if ($innerName) { $innerName.InnerText } else { $subValue.InnerText }
                                    } elseif ($subState) { $subState.InnerText }
                                    else { '' }
                                    [void]$subValues.Add("$($subName.InnerText)=$val")
                                }
                            }

                            $settingKey = "$scope|$(if ($polCategory) {$polCategory.InnerText} else {'(unknown)'})\$(if ($polName) {$polName.InnerText} else {'(unnamed)'})"
                            $stateVal   = if ($polState) { $polState.InnerText } else { 'Unknown' }

                            # Resolve English policy name from ADMX/CSP database (locale-independent)
                            $localPolicyName = if ($polName) { $polName.InnerText } else { '(unnamed)' }
                            $localCategory   = if ($polCategory) { $polCategory.InnerText } else { '' }
                            $localExplain    = if ($polExplain) { $polExplain.InnerText } else { '' }
                            $resolved = Resolve-PolicyFromRegistry -RegistryKey '' -ValueName '' -FallbackName $localPolicyName
                            # For admin template settings, the ADMX match is by name - try category-based match too
                            # (admin template settings don't always have a registry key exposed in gpresult XML)

                            [void]$allSettingsList.Add([PSCustomObject]@{
                                SettingKey   = $settingKey
                                PolicyName   = $localPolicyName
                                State        = $stateVal
                                Scope        = $scope
                                Category     = $localCategory
                                GPOName      = $gpo.DisplayName
                                GPOGuid      = $gpo.Id.ToString()
                                RegistryKey  = ''
                                ValueData    = ($subValues -join '; ')
                                Explain      = $localExplain
                                Source       = 'Domain GPO'
                                IntuneGroup  = 'Group Policy'
                            })
                        }

                        # Registry-based settings (q1:RegistrySetting nodes - non-administrative template)
                        $regSettings = $ext.ChildNodes | Where-Object { $_.LocalName -eq 'RegistrySetting' }
                        foreach ($reg in $regSettings) {
                            $settingCount++
                            $regHive  = $reg.SelectSingleNode('*[local-name()="KeyPath"]')
                            $regName  = $reg.SelectSingleNode('*[local-name()="ValueName"]')
                            $regData  = $reg.SelectSingleNode('*[local-name()="Value"]')
                            $regType  = $reg.SelectSingleNode('*[local-name()="Type"]')

                            $keyPath = if ($regHive) { $regHive.InnerText } else { '' }
                            $valName = if ($regName) { $regName.InnerText } else { '(Default)' }
                            $settingKey = "$scope|$keyPath\$valName"

                            # Resolve English policy name from ADMX/CSP database
                            $resolved = Resolve-PolicyFromRegistry -RegistryKey $keyPath -ValueName $valName -FallbackName "$keyPath\$valName"
                            $resolvedName     = $resolved.Name
                            $resolvedCategory = if ($resolved.Category) { $resolved.Category } else { 'Registry Settings' }
                            $resolvedExplain  = if ($resolved.Desc) { $resolved.Desc }
                                                elseif ($regType) { "Type: $($regType.InnerText)" }
                                                else { '' }
                            $resolvedGroup    = if ($resolved.Source -eq 'ADMX') { 'Administrative Templates' }
                                                elseif ($resolved.Source -eq 'CSP') { 'CSP Policy' }
                                                else { 'Registry Settings' }

                            [void]$allSettingsList.Add([PSCustomObject]@{
                                SettingKey   = $settingKey
                                PolicyName   = $resolvedName
                                State        = 'Registry'
                                Scope        = $scope
                                Category     = $resolvedCategory
                                GPOName      = $gpo.DisplayName
                                GPOGuid      = $gpo.Id.ToString()
                                RegistryKey  = $keyPath
                                ValueData    = if ($regData) { $regData.InnerText } else { '' }
                                Explain      = $resolvedExplain
                                Source       = 'Domain GPO'
                                IntuneGroup  = $resolvedGroup
                                AdmxSource   = $resolved.Source
                                CspPath      = $resolved.CspPath
                                AdmxFile     = $resolved.AdmxFile
                            })
                        }
                    }
                }
            } catch {
                Write-DebugLog "XML parse error for '$($gpo.DisplayName)': $($_.Exception.Message)" -Level WARN
            }
        }

        # H3: Security group filtering via LDAP ACL (replaces slow Get-GPPermission)
        $secFilter = [System.Collections.Generic.List[string]]::new()
        try {
            $gpoDN = "CN={$($gpo.Id.ToString())},CN=Policies,CN=System,$(($domain.Split('.') | ForEach-Object { "DC=$_" }) -join ',')"
            $gpoLdap = if ($DcOvr) { "LDAP://$DcOvr/$gpoDN" } else { "LDAP://$gpoDN" }
            $gpoEntry = [ADSI]$gpoLdap
            $sd = $gpoEntry.ObjectSecurity
            $applyGpGuid = [Guid]'edacfd8f-ffb3-11d1-b41d-00a0c968f939'
            foreach ($ace in $sd.GetAccessRules($true, $false, [System.Security.Principal.NTAccount])) {
                if ($ace.AccessControlType -eq 'Allow' -and $ace.ObjectType -eq $applyGpGuid) {
                    [void]$secFilter.Add("$($ace.IdentityReference)")
                }
            }
        } catch { }
        $secFilterStr = if ($secFilter.Count -gt 0) { $secFilter -join '; ' } else { 'Authenticated Users' }

        # H6: GP Preferences items (Drive Maps, Printers, Scheduled Tasks, etc.)
        if ($xmlText) {
            try {
                $prefTags = @(
                    @{ Tag = 'DriveMapSettings'; SubTag = 'Drive'; NameAttr = 'name'; Cat = 'Drive Maps' }
                    @{ Tag = 'PrinterSettings'; SubTag = 'SharedPrinter'; NameAttr = 'name'; Cat = 'Printers' }
                    @{ Tag = 'ScheduledTasks';  SubTag = 'Task';  NameAttr = 'name'; Cat = 'Scheduled Tasks' }
                    @{ Tag = 'FolderOptions';   SubTag = 'OpenWith'; NameAttr = 'name'; Cat = 'Folder Options' }
                    @{ Tag = 'EnvironmentVariables'; SubTag = 'EnvironmentVariable'; NameAttr = 'name'; Cat = 'Environment Vars' }
                    @{ Tag = 'RegistrySettings'; SubTag = 'Registry'; NameAttr = 'name'; Cat = 'Registry Prefs' }
                    @{ Tag = 'Shortcuts';       SubTag = 'Shortcut'; NameAttr = 'name'; Cat = 'Shortcuts' }
                    @{ Tag = 'Files';           SubTag = 'File'; NameAttr = 'name'; Cat = 'Files' }
                    @{ Tag = 'IniFiles';        SubTag = 'Ini'; NameAttr = 'name'; Cat = 'INI Files' }
                    @{ Tag = 'Services';        SubTag = 'NTService'; NameAttr = 'name'; Cat = 'Services' }
                    @{ Tag = 'DataSources';     SubTag = 'DataSource'; NameAttr = 'name'; Cat = 'Data Sources' }
                )
                foreach ($pt in $prefTags) {
                    $prefNodes = $xdoc.GetElementsByTagName($pt.Tag)
                    foreach ($pn in $prefNodes) {
                        $items = $pn.GetElementsByTagName($pt.SubTag)
                        foreach ($item in $items) {
                            $settingCount++
                            $iName = $item.GetAttribute($pt.NameAttr)
                            if (-not $iName) { $iName = $item.GetAttribute('name') }
                            if (-not $iName) { $iName = "$($pt.Cat)-$settingCount" }
                            $action = $item.SelectSingleNode('Properties/@action')
                            $actionVal = if ($action) { $action.Value } else { 'Update' }
                            # Collect key properties
                            $props = $item.SelectSingleNode('Properties')
                            $propStr = ''
                            if ($props) {
                                $propParts = [System.Collections.Generic.List[string]]::new()
                                foreach ($attr in $props.Attributes) {
                                    if ($attr.Name -ne 'action') { [void]$propParts.Add("$($attr.Name)=$($attr.Value)") }
                                }
                                $propStr = $propParts -join '; '
                            }
                            [void]$allSettingsList.Add([PSCustomObject]@{
                                SettingKey   = "Preferences|$($pt.Cat)\$iName"
                                PolicyName   = $iName
                                State        = $actionVal
                                Scope        = 'Computer'
                                Category     = "GP Preferences: $($pt.Cat)"
                                GPOName      = $gpo.DisplayName
                                GPOGuid      = $gpo.Id.ToString()
                                RegistryKey  = ''
                                ValueData    = $propStr
                                Explain      = ''
                                Source       = 'Domain GPO'
                                IntuneGroup  = 'GP Preferences'
                            })
                        }
                    }
                }
            } catch { }
        }

        # Determine status text
        $statusText = switch ($gpo.GpoStatus.ToString()) {
            'AllSettingsEnabled'              { 'Enabled' }
            'AllSettingsDisabled'             { 'Disabled' }
            'UserSettingsDisabled'            { 'Computer Only' }
            'ComputerSettingsDisabled'        { 'User Only' }
            default                           { $gpo.GpoStatus.ToString() }
        }

        # Compute link order (lowest = highest precedence)
        $linkOrder = if ($linkDetails.Count -gt 0) {
            ($linkDetails | ForEach-Object { $_.LinkOrder } | Sort-Object | Select-Object -First 1)
        } else { 0 }

        [void]$gpoRecords.Add([PSCustomObject]@{
            DisplayName      = $gpo.DisplayName
            Id               = $gpo.Id.ToString()
            Status           = $statusText
            GpoStatus        = $gpo.GpoStatus.ToString()
            SettingCount     = $settingCount
            LinkCount        = $linkLocations.Count
            Links            = ($linkLocations -join '; ')
            CreationTime     = $gpo.CreationTime.ToString('yyyy-MM-dd HH:mm')
            ModificationTime = $gpo.ModificationTime.ToString('yyyy-MM-dd HH:mm')
            WmiFilter        = if ($gpo.WmiFilter) { $gpo.WmiFilter.Name } else { '' }
            IsLinked         = ($linkLocations.Count -gt 0)
            Enforced         = $isEnforced
            LinkOrder        = $linkOrder
            SecurityFiltering = $secFilterStr
            LinkDetails      = $linkDetails
        })
    }


    $domain = if ($Script:Prefs.DomainOverride) { $Script:Prefs.DomainOverride }
              else { try { [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name } catch { '(unknown)' } }

    $scanResult = @{
        Timestamp   = (Get-Date)
        Domain      = $domain
        GPOs        = $gpoRecords
        Settings    = $allSettingsList
    }

    Write-DebugLog "Scan complete: $($gpoRecords.Count) GPOs, $($allSettingsList.Count) settings parsed" -Level SUCCESS
    return $scanResult
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 5b: LOCAL RSoP SCAN ENGINE (gpresult)
# ═══════════════════════════════════════════════════════════════════════════════

function Invoke-LocalRSoPScan {
    Write-DebugLog 'Starting local RSoP scan via gpresult...' -Level STEP

    $gpoRecords     = [System.Collections.Generic.List[PSCustomObject]]::new()
    $allSettingsList = [System.Collections.Generic.List[PSCustomObject]]::new()
    $gpoIdCounter   = 0
    $settIdCounter  = 0
    $domain         = 'LocalMachine'

    try {
        $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
        if ($cs.PartOfDomain) { $domain = $cs.Domain }
    } catch { }

    # M4: Loopback processing mode detection
    $loopbackMode = 'Not Configured'
    try {
        $lbReg = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\System' -Name 'UserPolicyMode' -ErrorAction SilentlyContinue
        if ($lbReg -and $null -ne $lbReg.UserPolicyMode) {
            $loopbackMode = switch ([int]$lbReg.UserPolicyMode) {
                1 { 'Replace' }
                2 { 'Merge' }
                default { "Unknown ($($lbReg.UserPolicyMode))" }
            }
        }
    } catch { }

    # gpresult /scope computer captures computer policies (requires admin)
    $tmpFile = [IO.Path]::Combine([IO.Path]::GetTempPath(), 'PolicyPilot_RSoP.xml')
    if (Test-Path $tmpFile) { Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue }
    Write-DebugLog "Running gpresult /scope computer /x $tmpFile /f" -Level INFO

    try {
        $errFile = [IO.Path]::Combine([IO.Path]::GetTempPath(), 'gpresult_err.txt')
        $proc = Start-Process -FilePath 'gpresult.exe' `
            -ArgumentList '/scope','computer','/x',$tmpFile,'/f' `
            -NoNewWindow -Wait -PassThru -RedirectStandardError $errFile

        if (-not $proc -or ($null -ne $proc.ExitCode -and $proc.ExitCode -ne 0) -or -not (Test-Path $tmpFile)) {
            # Fallback: try user-only scope
            $exitInfo = if ($proc -and $null -ne $proc.ExitCode) { $proc.ExitCode } else { 'N/A' }
            Write-DebugLog "gpresult /scope computer failed (exit $exitInfo), trying /scope user" -Level WARN
            $proc = Start-Process -FilePath 'gpresult.exe' `
                -ArgumentList '/scope','user','/x',$tmpFile,'/f' `
                -NoNewWindow -Wait -PassThru -RedirectStandardError $errFile
        }

        $finalFailed = (-not $proc) -or (-not (Test-Path $tmpFile))
        if (-not $finalFailed -and $proc -and $null -ne $proc.ExitCode -and $proc.ExitCode -ne 0) { $finalFailed = $true }
        if ($finalFailed) {
            $exitCode = if ($proc -and $null -ne $proc.ExitCode) { $proc.ExitCode } else { 'N/A' }
            Write-DebugLog "gpresult failed completely (exit $exitCode)" -Level ERROR
            return $null
        }

        [xml]$rsop = Get-Content -Path $tmpFile -Raw -Encoding UTF8
        $rootNs = $rsop.DocumentElement.NamespaceURI
        $nsMgr  = New-Object System.Xml.XmlNamespaceManager($rsop.NameTable)
        $nsMgr.AddNamespace('r', $rootNs)

        # Process both Computer and User results sections
        foreach ($scopeTag in @('ComputerResults','UserResults')) {
            $scope = if ($scopeTag -eq 'ComputerResults') { 'Computer' } else { 'User' }
            $scopeNode = $rsop.SelectSingleNode("//r:$scopeTag", $nsMgr)
            if (-not $scopeNode) {
                Write-DebugLog "No $scopeTag section in RSoP XML" -Level INFO
                continue
            }

            # 1. Parse GPO list from this scope
            $gpoNodes = $scopeNode.SelectNodes('r:GPO', $nsMgr)
            $scopeGpoMap = @{} # Path -> GPO record
            foreach ($gNode in $gpoNodes) {
                $gpoIdCounter++
                $gName   = $gNode.SelectSingleNode('r:Name', $nsMgr)
                $gPath   = $gNode.SelectSingleNode('r:Path', $nsMgr)
                $gEnable = $gNode.SelectSingleNode('r:Enabled', $nsMgr)
                $gLink   = $gNode.SelectSingleNode('r:Link/r:SOMPath', $nsMgr)

                # M1: WMI Filter evaluation status from RSoP XML
                $gFilterAllowed = $gNode.SelectSingleNode('r:FilterAllowed', $nsMgr)
                $wmiFilterStatus = if ($gFilterAllowed) {
                    if ($gFilterAllowed.InnerText -eq 'true') { 'Passed' } else { 'Denied' }
                } else { '' }

                # H2: Enforced + link order from RSoP
                $gEnforced = $gNode.SelectSingleNode('r:Link/r:NoOverride', $nsMgr)
                $gLinkOrder = $gNode.SelectSingleNode('r:Link/r:SOMOrder', $nsMgr)
                $enforced = ($gEnforced -and $gEnforced.InnerText -eq 'true')

                $displayName = if ($gName) { $gName.InnerText } else { "GPO-$gpoIdCounter" }
                $pathId      = if ($gPath) { $gPath.InnerText } else { "GPO-$gpoIdCounter" }
                $isEnabled   = if ($gEnable) { $gEnable.InnerText -eq 'true' } else { $true }
                $linkPath    = if ($gLink) { $gLink.InnerText } else { '' }

                $gpoRec = [PSCustomObject]@{
                    Id              = $gpoIdCounter
                    DisplayName     = $displayName
                    GpoId           = $pathId
                    Status          = if ($isEnabled) { 'Enabled' } else { 'Disabled' }
                    CreatedTime     = ''
                    ModifiedTime    = ''
                    WmiFilter       = ''
                    WmiFilterStatus = $wmiFilterStatus
                    LinkPath        = $linkPath
                    IsLinked        = [bool]$linkPath
                    UserVersion     = '0'
                    ComputerVersion = '0'
                    Description     = "Applied via RSoP ($scope scope)"
                    Enforced        = $enforced
                    LinkOrder       = if ($gLinkOrder) { [int]$gLinkOrder.InnerText } else { 0 }
                    SecurityFiltering = ''
                }
                [void]$gpoRecords.Add($gpoRec)
                $scopeGpoMap[$pathId] = $gpoRec
            }

            # Build path -> friendly GPO name map for setting GPOName resolution
            $pathToGpoName = @{}
            foreach ($g in $gpoRecords) { if ($g.GpoId -and $g.DisplayName) { $pathToGpoName[$g.GpoId] = $g.DisplayName } }

            # 2. Parse settings from ExtensionData sections
            $extDataNodes = $scopeNode.SelectNodes('r:ExtensionData', $nsMgr)
            foreach ($extData in $extDataNodes) {
                $extNameNode = $extData.SelectSingleNode('*[local-name()="Name"]')
                $category    = if ($extNameNode) { $extNameNode.InnerText } else { 'General' }

                $extNode = $extData.SelectSingleNode('*[local-name()="Extension"]')
                if (-not $extNode) { continue }

                # Parse Policy elements
                $policies = $extNode.SelectNodes('*[local-name()="Policy"]')
                foreach ($pol in $policies) {
                    $settIdCounter++
                    $polName  = $pol.SelectSingleNode('*[local-name()="Name"]')
                    $polState = $pol.SelectSingleNode('*[local-name()="State"]')
                    $polGPO   = $pol.SelectSingleNode('*[local-name()="GPO"]')
                    $polCat   = $pol.SelectSingleNode('*[local-name()="Category"]')
                    $polValue = $pol.SelectSingleNode('*[local-name()="Value"]')

                    $sName   = if ($polName)  { $polName.InnerText }  else { "Policy-$settIdCounter" }
                    $sState  = if ($polState) { $polState.InnerText } else { 'Applied' }
                    $sGPORaw = if ($polGPO)   { $polGPO.InnerText }   else { '' }
                    $sGPO    = if ($pathToGpoName[$sGPORaw]) { $pathToGpoName[$sGPORaw] } else { $sGPORaw }
                    $sCat    = if ($polCat)   { $polCat.InnerText }   else { $category }
                    $sValue  = if ($polValue) { $polValue.InnerText } else { $sState }
                    $sValue  = Format-NumberedList $sValue

                    $settRec = [PSCustomObject]@{
                        Id          = $settIdCounter
                        GPOName     = $sGPO
                        GPOGuid     = $sGPORaw
                        Category    = $sCat
                        PolicyName  = $sName
                        SettingKey  = "$sCat\$sName"
                        State       = $sState
                        RegistryKey = ''
                        ValueData   = $sValue
                        Scope       = $scope
                        Source      = 'Local GPO'
                        IntuneGroup = 'Group Policy'
                    }
                    [void]$allSettingsList.Add($settRec)
                }

                # Parse RegistrySetting elements
                $regSettings = $extNode.SelectNodes('*[local-name()="RegistrySetting"]')
                foreach ($reg in $regSettings) {
                    $settIdCounter++
                    $regGPO     = $reg.SelectSingleNode('*[local-name()="GPO"]')
                    $regKey     = $reg.SelectSingleNode('*[local-name()="KeyPath"]')
                    $regVal     = $reg.SelectSingleNode('*[local-name()="ValueName"]')
                    $regAdm     = $reg.SelectSingleNode('*[local-name()="AdmSetting"]')

                    $sGPORaw = if ($regGPO) { $regGPO.InnerText }  else { '' }
                    $sGPO    = if ($pathToGpoName[$sGPORaw]) { $pathToGpoName[$sGPORaw] } else { $sGPORaw }
                    $keyPath = if ($regKey) { $regKey.InnerText }  else { '' }
                    $valName = if ($regVal) { $regVal.InnerText }  else { '' }
                    $admSet  = if ($regAdm) { $regAdm.InnerText }  else { '' }

                    $fullPath = if ($valName) { "$keyPath\$valName" } else { $keyPath }
                    $sName    = if ($valName) { $valName } else { $keyPath.Split('\')[-1] }

                    $settRec = [PSCustomObject]@{
                        Id          = $settIdCounter
                        GPOName     = $sGPO
                        GPOGuid     = $sGPORaw
                        Category    = "$category (Registry)"
                        PolicyName  = $sName
                        SettingKey  = $fullPath
                        State       = 'Applied'
                        RegistryKey = $fullPath
                        ValueData   = "AdmSetting=$admSet"
                        Scope       = $scope
                        Source      = 'Local GPO'
                        IntuneGroup = 'Registry Settings'
                    }
                    [void]$allSettingsList.Add($settRec)
                }

                # Parse Account / SecurityOptions / Audit elements
                foreach ($tagName in @('Account','SecurityOptions','Audit')) {
                    $items = $extNode.SelectNodes("*[local-name()='$tagName']")
                    foreach ($item in $items) {
                        $settIdCounter++
                        $iName  = $item.SelectSingleNode('*[local-name()="Name"]')
                        $iGPO   = $item.SelectSingleNode('*[local-name()="GPO"]')
                        $iValue = $item.SelectSingleNode('*[local-name()="SettingNumber"] | *[local-name()="SettingBoolean"] | *[local-name()="SettingString"] | *[local-name()="Value"]')

                        $sName  = if ($iName)  { $iName.InnerText }  else { "$tagName-$settIdCounter" }
                        $sGPORaw = if ($iGPO)   { $iGPO.InnerText }   else { '' }
                        $sGPO   = if ($pathToGpoName[$sGPORaw]) { $pathToGpoName[$sGPORaw] } else { $sGPORaw }
                        $sValue = if ($iValue) { $iValue.InnerText } else { '' }

                        $settRec = [PSCustomObject]@{
                            Id          = $settIdCounter
                            GPOName     = $sGPO
                            GPOGuid     = $sGPORaw
                            Category    = "$category ($tagName)"
                            PolicyName  = $sName
                            SettingKey  = "$category\$sName"
                            State       = 'Applied'
                            RegistryKey = ''
                            ValueData   = $sValue
                            Scope       = $scope
                            Source      = 'Local GPO'
                            IntuneGroup = 'Group Policy'
                        }
                        [void]$allSettingsList.Add($settRec)
                    }
                }

                # M5: Script elements - extract Type (Logon/Logoff/Startup/Shutdown), Command, Parameters
                $scriptItems = $extNode.SelectNodes("*[local-name()='Script']")
                foreach ($sItem in $scriptItems) {
                    $settIdCounter++
                    $sName    = $sItem.SelectSingleNode('*[local-name()="Name"]')
                    $sGPO     = $sItem.SelectSingleNode('*[local-name()="GPO"]')
                    $sType    = $sItem.SelectSingleNode('*[local-name()="Type"]')
                    $sCommand = $sItem.SelectSingleNode('*[local-name()="Command"]')
                    $sParams  = $sItem.SelectSingleNode('*[local-name()="Parameters"]')
                    $sOrder   = $sItem.SelectSingleNode('*[local-name()="Order"]')

                    $scriptType = if ($sType) { $sType.InnerText } else { 'Unknown' }
                    $scriptCmd  = if ($sCommand) { $sCommand.InnerText } else { '' }
                    $scriptPrm  = if ($sParams) { $sParams.InnerText } else { '' }
                    $scriptOrd  = if ($sOrder) { $sOrder.InnerText } else { '' }
                    $scriptName = if ($sName) { $sName.InnerText } else { "$scriptType Script $settIdCounter" }

                    $scriptValue = $scriptCmd
                    if ($scriptPrm) { $scriptValue += " $scriptPrm" }
                    if ($scriptOrd) { $scriptValue += " (Order: $scriptOrd)" }

                    [void]$allSettingsList.Add([PSCustomObject]@{
                        Id          = $settIdCounter
                        GPOName     = if ($sGPO) { $sGPORaw = $sGPO.InnerText; if ($pathToGpoName[$sGPORaw]) { $pathToGpoName[$sGPORaw] } else { $sGPORaw } } else { '' }
                        GPOGuid     = if ($sGPO) { $sGPO.InnerText } else { '' }
                        Category    = "$category (Script: $scriptType)"
                        PolicyName  = $scriptName
                        SettingKey  = "$category\$scriptType\$scriptName"
                        State       = 'Applied'
                        RegistryKey = ''
                        ValueData   = $scriptValue
                        Scope       = $scope
                        Source      = 'Local GPO'
                        IntuneGroup = 'Scripts'
                        ScriptType  = $scriptType
                        ScriptPath  = $scriptCmd
                        ScriptParams = $scriptPrm
                    })
                }
            }

            Write-DebugLog "RSoP ${scope}: $($gpoNodes.Count) GPOs, settings parsed" -Level SUCCESS
        }
    } catch {
        Write-DebugLog "RSoP scan error: $($_.Exception.Message)" -Level ERROR
    } finally {
        if (Test-Path $tmpFile) { Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue }
    }

    # ── WMI RSoP enrichment: replace locale-dependent Policy entries with registry-keyed entries ──
    try {
        $rsopWmi = @(Get-CimInstance -Namespace 'root\RSOP\Computer' -ClassName 'RSOP_RegistryPolicySetting' -ErrorAction Stop)
        if ($rsopWmi.Count -gt 0) {
            Write-DebugLog "WMI RSoP: $($rsopWmi.Count) registry policy settings found" -Level INFO
            $guidToName = @{}
            foreach ($g in $gpoRecords) {
                if ($g.GpoId -match '\{([0-9a-fA-F-]+)\}') { $guidToName[$Matches[1].ToUpper()] = $g.DisplayName }
            }
            $replaceableEntries = @($allSettingsList | Where-Object { $_.Category -notmatch 'Security|Account|Audit|Script' })
            foreach ($pe in $replaceableEntries) { [void]$allSettingsList.Remove($pe) }
            Write-DebugLog "Replaced $($replaceableEntries.Count) gpresult entries with WMI RSoP data" -Level DEBUG
            foreach ($wmi in $rsopWmi) {
                if ($wmi.deleted) { continue }
                $settIdCounter++
                $wmiKey = $wmi.registryKey; $wmiVal = $wmi.valueName; $wmiGpoId = $wmi.GPOID
                $wmiPrec = if ($wmi.precedence) { [int]$wmi.precedence } else { 1 }
                $wmiValueData = ''
                if ($null -ne $wmi.value -and $wmi.value.Count -gt 0) {
                    switch ($wmi.valueType) {
                        4 { if ($wmi.value.Count -ge 4) { $wmiValueData = [string][BitConverter]::ToUInt32($wmi.value, 0) } else { $wmiValueData = [string]$wmi.value } }
                        1 { $wmiValueData = [System.Text.Encoding]::Unicode.GetString($wmi.value).TrimEnd("`0") }
                        7 { $wmiValueData = [System.Text.Encoding]::Unicode.GetString($wmi.value).TrimEnd("`0") -replace "`0", '; ' }
                        3 { $wmiValueData = "[Binary $($wmi.value.Count)B]" }
                        default { if ($wmi.value.Count -gt 0) { $wmiValueData = "[Type$($wmi.valueType) $($wmi.value.Count)B]" } }
                    }
                }
                $gpoGuid = ''; $gpoName = ''
                if ($wmiGpoId -match '\{([0-9a-fA-F-]+)\}') {
                    $gpoGuid = $Matches[1].ToUpper()
                    $gpoName = if ($guidToName[$gpoGuid]) { $guidToName[$gpoGuid] } else { $wmiGpoId }
                }
                $fullPath = if ($wmiVal) { "$wmiKey\$wmiVal" } else { $wmiKey }
                $sName = if ($wmiVal) { $wmiVal } else { $wmiKey.Split('\')[-1] }
                $wmiState = if ($wmiPrec -eq 1) { 'Applied' } else { 'Superseded' }
                [void]$allSettingsList.Add([PSCustomObject]@{
                    Id=$settIdCounter; GPOName=$gpoName; GPOGuid=$wmiGpoId
                    Category='Administrative Templates (WMI)'; PolicyName=$sName; SettingKey=$fullPath
                    State=$wmiState; RegistryKey=$fullPath; ValueData=$wmiValueData
                    Scope='Computer'; Source='WMI RSoP'; IntuneGroup='Group Policy'
                    Precedence=$wmiPrec
                })
            }
            Write-DebugLog "WMI RSoP: added $($rsopWmi.Count) enrichable settings" -Level SUCCESS
        }
    } catch {
        Write-DebugLog "WMI RSoP unavailable: $($_.Exception.Message)" -Level DEBUG
    }

    $scanResult = @{
        Timestamp      = [datetime]::Now
        Domain         = $domain
        GPOs           = $gpoRecords
        Settings       = $allSettingsList
        LoopbackMode   = $loopbackMode
    }

    Write-DebugLog "Local RSoP scan complete: $($gpoRecords.Count) GPOs, $($allSettingsList.Count) settings" -Level SUCCESS
    return $scanResult
}

# ===============================================================================

# SECTION 5c: INTUNE POLICY SCAN ENGINE (Microsoft Graph)

# ===============================================================================



function Invoke-IntunePolicyScan {
    function Format-PolicyName([string]$name) {
        ($name -creplace '([a-z])([A-Z])', '$1 $2' -creplace '([A-Z]+)([A-Z][a-z])', '$1 $2').Trim()
    }

    # H4: Helper to fetch assignments for a given policy endpoint
    function Get-PolicyAssignments([string]$policyId, [string]$policyType) {
        $assignments = [System.Collections.Generic.List[string]]::new()
        try {
            $uri = switch ($policyType) {
                'DeviceConfig' { "https://graph.microsoft.com/v1.0/deviceManagement/deviceConfigurations/$policyId/assignments" }
                'Compliance'   { "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies/$policyId/assignments" }
                'SettingsCatalog' { "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/$policyId/assignments" }
            }
            $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
            foreach ($a in $resp.value) {
                $target = $a.target
                $targetType = $target.'@odata.type' -replace '#microsoft.graph.',''
                $groupId = $target.groupId
                $label = switch -Wildcard ($targetType) {
                    '*allLicensedUsers*' { 'All Users' }
                    '*allDevices*'       { 'All Devices' }
                    '*exclusionGroup*'   { "Exclude: $groupId" }
                    default              { $groupId }
                }
                [void]$assignments.Add($label)
            }
        } catch { }
        return ($assignments -join '; ')
    }

    $gpoRecords     = [System.Collections.Generic.List[PSCustomObject]]::new()

    $allSettingsList = [System.Collections.Generic.List[PSCustomObject]]::new()

    $gpoIdCounter   = 0

    $settIdCounter  = 0



    try {

        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

        Import-Module Microsoft.Graph.DeviceManagement -ErrorAction Stop

    } catch {

        try { Import-Module Microsoft.Graph.Intune -ErrorAction Stop } catch {

            return @{ Error = 'Microsoft.Graph.DeviceManagement not found. Install: Install-Module Microsoft.Graph -Scope CurrentUser' }

        }

    }

    try {

        $ctx = Get-MgContext -ErrorAction SilentlyContinue

        if (-not $ctx) {

            Connect-MgGraph -Scopes @('DeviceManagementConfiguration.Read.All','DeviceManagementManagedDevices.Read.All') -ErrorAction Stop

        }

    } catch { return @{ Error = "Graph auth failed: $($_.Exception.Message)" } }



    $domain = try { (Get-MgContext).TenantId } catch { 'Intune' }



    # Device Configuration Profiles

    try {

        $configs = Get-MgDeviceManagementDeviceConfiguration -All -ErrorAction Stop

        foreach ($cfg in $configs) {

            $gpoIdCounter++; $sc = 0

            # H7: Capture scope tags
            $scopeTags = $cfg.AdditionalProperties['roleScopeTagIds']
            $scopeTagStr = if ($scopeTags -is [System.Collections.IEnumerable]) { ($scopeTags | ForEach-Object { "$_" }) -join ', ' } elseif ($scopeTags) { "$scopeTags" } else { '0' }
            # H4: Fetch assignments
            $assignStr = Get-PolicyAssignments -policyId $cfg.Id -policyType 'DeviceConfig'

            [void]$gpoRecords.Add([PSCustomObject]@{

                Id=$gpoIdCounter; DisplayName=$cfg.DisplayName; GpoId=$cfg.Id; Status='Enabled'

                CreatedTime=if($cfg.CreatedDateTime){$cfg.CreatedDateTime.ToString('yyyy-MM-dd HH:mm')}else{''}

                ModifiedTime=if($cfg.LastModifiedDateTime){$cfg.LastModifiedDateTime.ToString('yyyy-MM-dd HH:mm')}else{''}

                WmiFilter=''; LinkPath='Intune > Device Configuration'; IsLinked=$true

                UserVersion='0'; ComputerVersion='0'; Description=if($cfg.Description){$cfg.Description}else{'Device Configuration'}

                SettingCount=0; LinkCount=1; Links='Intune > Device Configuration'
                ScopeTags=$scopeTagStr; Assignments=$assignStr

            })

            $odt = $cfg.AdditionalProperties['@odata.type'] -replace '#microsoft.graph.',''

            foreach ($kv in $cfg.AdditionalProperties.GetEnumerator()) {

                if ($kv.Key.StartsWith('@') -or $kv.Key -in @('id','displayName','description','version','createdDateTime','lastModifiedDateTime','roleScopeTagIds')) { continue }

                $pv = $kv.Value; if ($null -eq $pv) { continue }

                if ($pv -is [System.Collections.IEnumerable] -and $pv -isnot [string]) { $pv = ($pv | ForEach-Object { $_.ToString() }) -join ', ' }

                $settIdCounter++; $sc++

                [void]$allSettingsList.Add([PSCustomObject]@{

                    Id=$settIdCounter; GPOName=$cfg.DisplayName; GPOGuid=$cfg.Id

                    Category="DeviceConfig ($odt)"; PolicyName=(Format-PolicyName $kv.Key)

                    SettingKey="DeviceConfig|$odt\$($kv.Key)"; State='Applied'

                    RegistryKey=''; ValueData="$pv"; Scope='Device'; Source='Intune'; IntuneGroup='Configuration Profiles'

                })

            }

            $gpoRecords[$gpoRecords.Count - 1].SettingCount = $sc

        }

    } catch { try { Write-DebugLog "Unhandled: $_" -Level ERROR } catch {} }



    # Compliance Policies

    try {

        $compliance = Get-MgDeviceManagementDeviceCompliancePolicy -All -ErrorAction Stop

        foreach ($cp in $compliance) {

            $gpoIdCounter++; $sc = 0

            # H7: Capture scope tags
            $scopeTags = $cp.AdditionalProperties['roleScopeTagIds']
            $scopeTagStr = if ($scopeTags -is [System.Collections.IEnumerable]) { ($scopeTags | ForEach-Object { "$_" }) -join ', ' } elseif ($scopeTags) { "$scopeTags" } else { '0' }
            # H4: Fetch assignments
            $assignStr = Get-PolicyAssignments -policyId $cp.Id -policyType 'Compliance'

            [void]$gpoRecords.Add([PSCustomObject]@{

                Id=$gpoIdCounter; DisplayName=$cp.DisplayName; GpoId=$cp.Id; Status='Enabled'

                CreatedTime=if($cp.CreatedDateTime){$cp.CreatedDateTime.ToString('yyyy-MM-dd HH:mm')}else{''}

                ModifiedTime=if($cp.LastModifiedDateTime){$cp.LastModifiedDateTime.ToString('yyyy-MM-dd HH:mm')}else{''}

                WmiFilter=''; LinkPath='Intune > Compliance'; IsLinked=$true

                UserVersion='0'; ComputerVersion='0'; Description=if($cp.Description){$cp.Description}else{'Compliance Policy'}

                SettingCount=0; LinkCount=1; Links='Intune > Compliance'
                ScopeTags=$scopeTagStr; Assignments=$assignStr

            })

            $odt = $cp.AdditionalProperties['@odata.type'] -replace '#microsoft.graph.',''

            foreach ($kv in $cp.AdditionalProperties.GetEnumerator()) {

                if ($kv.Key.StartsWith('@') -or $kv.Key -in @('id','displayName','description','version','createdDateTime','lastModifiedDateTime','roleScopeTagIds')) { continue }

                $pv = $kv.Value; if ($null -eq $pv) { continue }

                if ($pv -is [System.Collections.IEnumerable] -and $pv -isnot [string]) { $pv = ($pv | ForEach-Object { $_.ToString() }) -join ', ' }

                $settIdCounter++; $sc++

                [void]$allSettingsList.Add([PSCustomObject]@{

                    Id=$settIdCounter; GPOName=$cp.DisplayName; GPOGuid=$cp.Id

                    Category="Compliance ($odt)"; PolicyName=(Format-PolicyName $kv.Key)

                    SettingKey="Compliance|$odt\$($kv.Key)"; State='Applied'

                    RegistryKey=''; ValueData="$pv"; Scope='Device'; Source='Intune'; IntuneGroup='Device Compliance'

                })

            }

            $gpoRecords[$gpoRecords.Count - 1].SettingCount = $sc

        }

    } catch { try { Write-DebugLog "Unhandled: $_" -Level ERROR } catch {} }



    # Settings Catalog

    try {

        $catPols = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies' -ErrorAction Stop

        foreach ($pol in $catPols.value) {

            $gpoIdCounter++; $sc = 0

            # H7: Capture scope tags
            $scopeTagStr = if ($pol.roleScopeTagIds) { ($pol.roleScopeTagIds | ForEach-Object { "$_" }) -join ', ' } else { '0' }
            # H4: Fetch assignments
            $assignStr = Get-PolicyAssignments -policyId $pol.id -policyType 'SettingsCatalog'

            [void]$gpoRecords.Add([PSCustomObject]@{

                Id=$gpoIdCounter; DisplayName=$pol.name; GpoId=$pol.id; Status='Enabled'

                CreatedTime=if($pol.createdDateTime){([datetime]$pol.createdDateTime).ToString('yyyy-MM-dd HH:mm')}else{''}

                ModifiedTime=if($pol.lastModifiedDateTime){([datetime]$pol.lastModifiedDateTime).ToString('yyyy-MM-dd HH:mm')}else{''}

                WmiFilter=''; LinkPath='Intune > Settings Catalog'; IsLinked=$true

                UserVersion='0'; ComputerVersion='0'; Description=if($pol.description){$pol.description}else{'Settings Catalog'}

                SettingCount=0; LinkCount=1; Links='Intune > Settings Catalog'
                ScopeTags=$scopeTagStr; Assignments=$assignStr

            })

            try {

                $pu = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/$($pol.id)/settings"

                $pss = Invoke-MgGraphRequest -Method GET -Uri $pu -ErrorAction Stop

                foreach ($s in $pss.value) {

                    $settIdCounter++; $sc++

                    $defId = $s.settingInstance.settingDefinitionId

                    $sName = if ($defId) { $defId.Split('_')[-1] } else { "Setting-$settIdCounter" }

                    $sVal = ''

                    if ($s.settingInstance.PSObject.Properties['simpleSettingValue']) { $sVal = "$($s.settingInstance.simpleSettingValue.value)" }

                    elseif ($s.settingInstance.PSObject.Properties['choiceSettingValue']) { $sVal = "$($s.settingInstance.choiceSettingValue.value)" }

                    else { $sVal = ($s.settingInstance | ConvertTo-Json -Compress -Depth 3 -ErrorAction SilentlyContinue) }

                    [void]$allSettingsList.Add([PSCustomObject]@{

                        Id=$settIdCounter; GPOName=$pol.name; GPOGuid=$pol.id

                        Category='Settings Catalog'; PolicyName=(Format-PolicyName $sName)

                        SettingKey="SettingsCatalog|$defId"; State='Applied'

                        RegistryKey=if($defId){$defId}else{''}; ValueData=$sVal; Scope='Device'; Source='Intune'; IntuneGroup='Configuration Profiles'

                    })

                }

            } catch { try { Write-DebugLog "Unhandled: $_" -Level ERROR } catch {} }

            $gpoRecords[$gpoRecords.Count - 1].SettingCount = $sc

        }

    } catch { try { Write-DebugLog "Unhandled: $_" -Level ERROR } catch {} }



    return @{ Timestamp=[datetime]::Now; Domain=$domain; GPOs=$gpoRecords; Settings=$allSettingsList }

}


# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 6: CONFLICT DETECTION ENGINE
# ═══════════════════════════════════════════════════════════════════════════════

function Find-Conflicts {
    param([System.Collections.Generic.List[PSCustomObject]]$SettingsList)

    Write-DebugLog 'Running conflict detection...' -Level STEP

    # Build GPO lookup for precedence info (Enforced, LinkOrder)
    $gpoLookup = @{}
    if ($Script:ScanData -and $Script:ScanData.GPOs) {
        foreach ($g in $Script:ScanData.GPOs) {
            $gpoLookup[$g.DisplayName] = $g
        }
    }

    # Group settings by SettingKey (case-insensitive to prevent registry path case duplication)
    $settingMap = @{}
    foreach ($s in $SettingsList) {
        $key = $s.SettingKey
        if ([string]::IsNullOrWhiteSpace($key)) { continue }
        $normKey = $key.ToLower()
        if (-not $settingMap.ContainsKey($normKey)) {
            $settingMap[$normKey] = [System.Collections.Generic.List[PSCustomObject]]::new()
        }
        [void]$settingMap[$normKey].Add($s)
    }

    $conflicts = [System.Collections.Generic.List[PSCustomObject]]::new()

    # Collect all group keys for parent-child dedup
    $allGroupKeys = @($settingMap.Keys)

    foreach ($kv in $settingMap.GetEnumerator()) {
        $group = $kv.Value
        if ($group.Count -lt 2) { continue }

        # Skip parent-key groups when child keys exist (e.g. skip ...\Personalization if ...\Personalization\NoLockScreenCamera exists)
        $thisKey = $kv.Key
        $hasChildKey = $false
        foreach ($oKey in $allGroupKeys) {
            if ($oKey -ne $thisKey -and $oKey.StartsWith("$thisKey\") -and $settingMap[$oKey].Count -ge 2) {
                $hasChildKey = $true; break
            }
        }
        if ($hasChildKey) { continue }

        # Deduplicate: keep one entry per GPO (prefer the winning/Applied entry)
        $byGpo = @{}
        foreach ($s in $group) {
            $gName = $s.GPOName
            if (-not $gName) { continue }
            if (-not $byGpo.ContainsKey($gName) -or $s.State -eq 'Applied') {
                $byGpo[$gName] = $s
            }
        }
        $dedupGroup = @($byGpo.Values)
        if ($dedupGroup.Count -lt 2) { continue }

        # Check if all values are the same
        $uniqueStates = @($dedupGroup | Select-Object -ExpandProperty State -Unique)
        $uniqueValues = @($dedupGroup | Where-Object { $_.ValueData } | Select-Object -ExpandProperty ValueData -Unique)

        $severity = if ($uniqueStates.Count -gt 1 -or $uniqueValues.Count -gt 1) {
            'Conflict'
        } else {
            'Redundant'
        }

        $gpoNames = ($dedupGroup | Select-Object -ExpandProperty GPOName -Unique) -join ', '
        # Concise conflict summary: one line per GPO with its value
        $values = ($dedupGroup | ForEach-Object {
            $val = if ($_.ValueData -and $_.ValueData -ne '0' -and $_.ValueData -ne '') {
                $v = "$($_.ValueData)"
                if ($v.Length -gt 60) { $v = $v.Substring(0, 57) + '...' }
                " = $v"
            } else { '' }
            "$($_.GPOName)$( if ($_.State -eq 'Superseded') { ' [LOSES]' } else { ' [WINS]' })$val"
        }) -join ' | '

        # H1: Determine winning GPO using proper precedence
        # Priority: 1) WMI RSoP Precedence field (lowest = winner)
        #           2) Enforced GPOs always win (lowest link order among enforced)
        #           3) Lowest link order (closer to OU = higher precedence)
        #           4) Last in processing order (RSoP: last entry = winning)
        $winnerGPO = ''
        $hasPrec = $dedupGroup[0].PSObject.Properties.Name -contains 'Precedence'
        if ($hasPrec) {
            $winner = $dedupGroup | Sort-Object { if ($_.Precedence) { [int]$_.Precedence } else { 999 } } | Select-Object -First 1
            $winnerGPO = $winner.GPOName
        } else {
            $uniqueGPOs = @($dedupGroup | Select-Object -ExpandProperty GPOName -Unique)
            $enforcedCandidates = @()
            $normalCandidates = @()
            foreach ($gName in $uniqueGPOs) {
                $gInfo = $gpoLookup[$gName]
                if ($gInfo -and $gInfo.Enforced) {
                    $enforcedCandidates += @{ Name = $gName; LinkOrder = if ($gInfo.LinkOrder) { $gInfo.LinkOrder } else { 999 } }
                } else {
                    $normalCandidates += @{ Name = $gName; LinkOrder = if ($gInfo -and $gInfo.LinkOrder) { $gInfo.LinkOrder } else { 999 } }
                }
            }
            if ($enforcedCandidates.Count -gt 0) {
                $winnerGPO = ($enforcedCandidates | Sort-Object { $_.LinkOrder } | Select-Object -First 1).Name
            } elseif ($normalCandidates.Count -gt 0) {
                $winnerGPO = ($normalCandidates | Sort-Object { $_.LinkOrder } | Select-Object -First 1).Name
            } else {
                $winnerGPO = $uniqueGPOs[-1]
            }
        }

        $settingRef = $dedupGroup[0]
        # Only flag as conflict if multiple DIFFERENT GPOs are involved
        $uniqueGPOs = @($dedupGroup | Select-Object -ExpandProperty GPOName -Unique | Where-Object { $_ })
        if ($uniqueGPOs.Count -lt 2 -and $severity -eq 'Conflict') {
            continue
        }

        # Resolve registry path to friendly ADMX/CSP policy name
        $friendlyName = $settingRef.SettingKey
        $admxSource = ''
        if ($settingRef.RegistryKey) {
            $resolved = Resolve-PolicyFromRegistry -RegistryKey $settingRef.RegistryKey -ValueName '' -FallbackName $settingRef.PolicyName
            if ($resolved.Source -in @('ADMX','CSP')) {
                $friendlyName = $resolved.Name
                $admxSource = $resolved.Source
                if ($resolved.Category) { $friendlyName = "$($resolved.Category) > $($resolved.Name)" }
            }
        }
        # Fall back to PolicyName if SettingKey is just a raw path
        if ($friendlyName -eq $settingRef.SettingKey -and $settingRef.PolicyName -and $settingRef.PolicyName -ne $settingRef.SettingKey) {
            $friendlyName = $settingRef.PolicyName
        }

        [void]$conflicts.Add([PSCustomObject]@{
            Severity    = $severity
            SettingKey  = $friendlyName
            RegistryPath = $settingRef.SettingKey
            Scope       = $settingRef.Scope
            Category    = $settingRef.Category
            GPONames    = $gpoNames
            GPOCount    = $uniqueGPOs.Count
            Values      = $values
            WinnerGPO   = $winnerGPO
            AdmxSource  = $admxSource
            Details     = $group
        })
    }

    # Sort: Conflicts first, then Redundant
    $sorted = $conflicts | Sort-Object @{Expression={if ($_.Severity -eq 'Conflict') {0} else {1}}}, SettingKey
    Write-DebugLog "Conflict detection: $(@($sorted | Where-Object Severity -eq 'Conflict').Count) conflicts, $(@($sorted | Where-Object Severity -eq 'Redundant').Count) redundant" -Level INFO
    return $sorted
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 7: XAML LOADING
# ═══════════════════════════════════════════════════════════════════════════════

Add-Type -AssemblyName System.Windows.Forms   # for DoEvents

$xamlPath = Join-Path $Script:AppDir 'PolicyPilot_UI.xaml'
if (-not (Test-Path $xamlPath)) {
    if ($Headless) {
        Write-Host "[PolicyPilot] CRITICAL: 'PolicyPilot_UI.xaml' not found in: $($Script:AppDir)" -ForegroundColor Red
    } else {
        [System.Windows.MessageBox]::Show(
            "CRITICAL: 'PolicyPilot_UI.xaml' not found in:`n$($Script:AppDir)",
            'PolicyPilot', 'OK', 'Error')
    }
    exit 1
}

$xamlContent = [System.IO.File]::ReadAllText($xamlPath)
$xml = [xml]$xamlContent
$reader = [System.Xml.XmlNodeReader]::new($xml)
$Window = [System.Windows.Markup.XamlReader]::Load($reader)
$reader.Close()

# ── Named element references ──
$ui = @{}
@(
    'WindowBorder','TitleBar','TitleText','VersionBadge','BtnMinimize','BtnMaximize','BtnClose','MaximizeIcon','BtnHelp'
    'DomainBadge','StatusDot','StatusText'
    # Nav
    'NavDashboard','NavGPOList','NavSettings','NavConflicts','NavIntuneApps','NavReport','NavIMELogs','NavGPOLogs','NavMDMSync','NavTools','NavAppSettings'
    # Sidebars
    'SidebarDashboard','SidebarGPOList','SidebarSettings','SidebarConflicts','SidebarIntuneApps','SidebarReport','SidebarIMELogs','SidebarGPOLogs','SidebarMDMSync','SidebarTools','SidebarAppSettings'
    # Dashboard sidebar
    'CmbScanMode','BtnScanText','BtnScanGPOs','ScanProgressPanel','ScanProgressBar','ScanProgressText'
    'DashLastScanTime','DashDomainName','SidebarActiveGPOs','SidebarTotalSettings','SidebarConflictsStat','SidebarRedundant'
    'BtnSaveSnapshot','BtnLoadSnapshot'
    # GPO List sidebar
    'TxtGPOSearch','FilterGPOAll','FilterGPOEnabled','FilterGPODisabled'
    'FilterScopeBoth','FilterScopeComputer','FilterScopeUser'
    'GPODetailCard','GPODetailName','GPODetailGuid','GPODetailStatus','GPODetailLinks','GPODetailModified'
    # Settings sidebar
    'TxtSettingSearch','FilterSettingAll','FilterSettingComputer','FilterSettingUser','CmbCategoryFilter'
    'FilterCfgAll','FilterCfgConfigured','FilterCfgNotConfigured'
    'FilterValAll','FilterValNonDefault','FilterValDefault'
    'FilterGroupAll','FilterGroupEndpointSec','FilterGroupAcctProt','FilterGroupCompliance'
    'FilterGroupWinUpdate','FilterGroupAppMgmt','FilterGroupRegistry','FilterGroupConfigProf'
    'FilterGroupADMXProf','FilterGroupGPO','FilterGroupPpkg'
    'StatConfigured','StatNotConfigured','StatNonDefault','StatCoverage'
    # Conflicts sidebar
    'FilterConflictAll','FilterConflictOnly','FilterRedundantOnly','ConflictSummaryText'
    # IntuneApps sidebar
    'IntuneAppsSummaryText','BtnResetAppInstall','CmbAppTypeFilter','CmbAppStatusFilter'
    'AppAdminBanner','AppAdminBannerText','BtnDismissAppBanner'
    # Report sidebar
    'BtnExportHtml','BtnExportCsv','BtnExportConflictsCsv'
    'ChkIncludeDisabled','ChkIncludeUnlinked','ChkShowRegistryPaths'
    # App Settings sidebar
    'BtnThemeDark','BtnThemeLight','TxtDomainOverride','TxtDCOverride'
    'TxtOUScope','BtnDetectOU','TxtDetectOUStatus','ChkForceRefresh'
    'BtnDetectDC','TxtDetectDCStatus'
    'PrereqStatus','PrereqDetailStatus','BtnRecheckPrereqs'
    'AboutText'
    # Panels
    'PanelDashboard','PanelGPOList','PanelSettings','PanelConflicts','PanelIntuneApps','PanelReport','PanelIMELogs','PanelGPOLogs','PanelMDMSync','PanelTools','PanelAppSettings'
    # Dashboard content
    'StatActiveGPOs','StatTotalSettings','StatConflicts','StatRedundant','StatUnlinked'
    'DashEmptyState','DashDataState','DashTopConflictsGrid','BtnViewAllConflicts','BtnGetStarted','BtnGetStartedText'
    # Data grids
    'GPOListGrid','GPOListSubtitle','GPOListEmptyState'
    'SettingsGrid','SettingsSubtitle','SettingsEmptyState'
    'ConflictsGrid','ConflictsSubtitle','ConflictsEmptyState'
    'IntuneAppsGrid','IntuneAppsSubtitle','IntuneAppsEmptyState'
    # Report
    'ReportScroller','ReportPreviewText'
    # Toast
    'ToastBorder','ToastAccentBar','ToastIcon','ToastTitle','ToastMessage'
    # Console panel
    'pnlBottomPanel','btnClearLog','btnHideBottom','rtbActivityLog','docActivityLog','paraLog','logScroller','splitterBottom'
    # Status bar
    'statusBarDot','lblStatusBar','lblStatusDetail','btnToggleConsole','lblVersionBar'
    # Shimmer
    'pnlGlobalProgress','brdGlobalShimmer','lblGlobalProgress'
    # Sidebar scan status indicators + IntuneApps scan button
    'ScanStatusGPOList','LblScanStatusGPOList','BtnGoToDashGPOList'
    'ScanStatusSettings','LblScanStatusSettings','BtnGoToDashSettings'
    'ScanStatusConflicts','LblScanStatusConflicts','BtnGoToDashConflicts'
    'BtnScan_IntuneApps'
    'ScanStatusReport','LblScanStatusReport','BtnGoToDashReport'
    # New feature UI elements
    'DetailPaneText','SettingsDetailPane','BtnCloseDetailPane','BtnCompareSnapshot','BtnExportScript','BtnPrintReport','BtnExportReg',
    'BtnImportBaseline','BtnExportBaseline','BtnSimulateImpact','BaselineStatus','ChartCanvas','ChartPanel','BtnOUTreeView','TxtDomainOverride'
    # Tools panel elements
    'BtnRegPolicyManager','BtnRegEnrollments','BtnRegProvisioning','BtnRegIME','BtnRegDeclaredConfig','BtnRegDesktopAppMgmt','BtnRegRebootURIs'
    'BtnRegPolicyManager2','BtnRegEnrollments2','BtnRegProvisioning2','BtnRegIME2','BtnRegDeclaredConfig2','BtnRegDesktopAppMgmt2','BtnRegRebootURIs2'
    'BtnMmpCSync','MmpCSyncStatus','TxtStatusCode','BtnLookupStatus','StatusCodeResult'
    'TxtBase64Input','BtnDecodeBase64','Base64Result'
    'TxtStatusCode2','BtnLookupStatus2','StatusCodeResult2','TxtBase64Input2','BtnDecodeBase642','Base64Result2'
    # ETW Trace elements
    'BtnEtwStart','BtnEtwStop','BtnEtwClear','EtwTraceStatus'
    'BtnEtwStart2','BtnEtwStop2','BtnEtwClear2','EtwTraceStatus2','EtwResultScroller','EtwResultBox'
    # Phase 2: Autopilot, WiFi/VPN, Node Cache, Background Log
    'TxtAutopilotHash','BtnDecodeAutopilot','AutopilotResult'
    'TxtAutopilotHash2','BtnDecodeAutopilot2','AutopilotResult2'
    'BtnLoadWifi','BtnLoadVpn','NetworkProfileStatus'
    'BtnLoadWifi2','BtnLoadVpn2','NetworkProfileResult'
    'BtnLoadNodeCache','NodeCacheStatus'
    'BtnLoadNodeCache2','NodeCacheStatus2','NodeCacheScroller','NodeCacheResultBox'
    'BtnBgLogToggle','BgLogStatus'
    'BtnBgLogToggle2','BgLogStatus2'
    # Sidebar collapse & Achievements
    'BtnHamburger','SidebarBorder','SidebarColumn','AchievementsPanel','btnToggleAchievements','txtAchievementChevron','pnlAchievements','lblAchievementCount','cnvConfetti','ChkHideInternal'
    # MDM Enrollment card
    'MdmEnrollmentCard','MdmProviderText','MdmUPNText','MdmEnrollTypeText','MdmStateText'
    # LAPS + Certificate cards
    'LapsStatusCard','LapsBackupText','LapsPasswordAgeText','LapsComplexityText','LapsPostAuthText','LapsLastRotationText'
    'CertInventoryCard','CertCountText','CertExpiringText','CertDetailText'
    # App Status + Compliance + Import cards
    'AppStatusCard','AppInstalledText','AppFailedText','AppPendingText','AppSuccessRateText'
    'ComplianceCard','ComplianceStatusBadge','CompliancePolicyCountText','ComplianceAppRateText','ComplianceIssuesText'
    'BtnImportMdmXml'
    # Script Policies + Config Profiles + Enrollment Issues cards
    'ScriptPoliciesCard','ScriptTotalText','ScriptFailedText','ScriptDetailText'
    'ConfigProfileCard','ProfileTotalText','ProfileErrorText','ProfileDetailText'
    'ProvisioningPkgCard','PpkgTotalText','PpkgFailText','PpkgDetailText'
    'EnrollmentIssuesCard','EnrollmentIssuesText'
    # IME Logs
    'CmbImeLogSource','CmbImeLineLimit','ChkImeLiveTail','ImeModeHint'
    'FilterImeAll','FilterImeError','FilterImeWarning','FilterImeInfo'
    'TxtImeContentFilter','PresetImeNone','PresetImeWin32','PresetImePowershell','PresetImePolicy','PresetImeNetwork','PresetImeCheckin','PresetImeCert','PresetImeTimeout'
    'CmbImeSavedFilters','BtnImeSaveFilter'
    'ImeStatLines','ImeStatErrors','ImeStatWarnings','ImeFollowIndicator'
    'BtnIntuneSync','ImeSyncStatus','ImeSyncAdminBanner','ImeSyncBtnLabel'
    'ImeLogSubtitle','cnvImeHeatmap'
    'TxtImeSearch','ImeSearchCount','BtnImeSearchPrev','BtnImeSearchNext','BtnImeWrapToggle','BtnImeClearLog'
    'lbImeLogs','cnvImeMinimap'
    'ImeDensityBar','ImeDensityError','ImeDensityWarn','ImeFindingsList'
    # GPO Logs
    'CmbGpoLogSource','ChkGpoLiveTail','GpoModeHint'
    'FilterGpoAll','FilterGpoError','FilterGpoWarning','FilterGpoInfo'
    'TxtGpoContentFilter','PresetGpoNone','PresetGpoSecurity','PresetGpoRegistry','PresetGpoScripts','PresetGpoPrefs','PresetGpoCSE','PresetGpoSlowLink','PresetGpoLoopback'
    'CmbGpoSavedFilters','BtnGpoSaveFilter'
    'GpoStatLines','GpoStatErrors','GpoStatWarnings','GpoFollowIndicator'
    'BtnGpoDebugToggle','GpoDebugToggleLabel','GpoDebugStatus','GppTraceStatus'
    'BtnGpoRefresh','GpoRefreshStatus'
    'GpoAdminBanner','GpoAdminBannerText'
    'GpoLogSubtitle','cnvGpoHeatmap'
    'TxtGpoSearch','GpoSearchCount','BtnGpoSearchPrev','BtnGpoSearchNext','BtnGpoWrapToggle','BtnGpoClearLog'
    'lbGpoLogs','cnvGpoMinimap'
    'GpoDensityBar','GpoDensityError','GpoDensityWarn','GpoFindingsList'
    # MDM Sync
    'CmbMdmLogSource','ChkMdmLiveTail','MdmModeHint'
    'FilterMdmAll','FilterMdmError','FilterMdmWarning','FilterMdmInfo'
    'TxtMdmContentFilter','PresetMdmNone','PresetMdmSync','PresetMdmCSP','PresetMdmCert','PresetMdmCompliance','PresetMdmEnroll','PresetMdmWipe','PresetMdmBitLocker'
    'CmbMdmSavedFilters','BtnMdmSaveFilter'
    'MdmStatLines','MdmStatErrors','MdmStatWarnings','MdmFollowIndicator'
    'MdmAdminBanner','MdmAdminBannerText'
    'MdmLogSubtitle','cnvMdmHeatmap'
    'TxtMdmSearch','MdmSearchCount','BtnMdmSearchPrev','BtnMdmSearchNext','BtnMdmWrapToggle','BtnMdmClearLog'
    'lbMdmLogs','cnvMdmMinimap'
) | ForEach-Object { $ui[$_] = $Window.FindName($_) }

# Re-bind data grids
if ($ui.GPOListGrid)    { $ui.GPOListGrid.ItemsSource    = $Script:AllGPOs }
if ($ui.SettingsGrid)   { $ui.SettingsGrid.ItemsSource   = $Script:AllSettings }
    $settingsView = [System.Windows.Data.CollectionViewSource]::GetDefaultView($Script:AllSettings)
    if ($settingsView) {
        $settingsView.GroupDescriptions.Clear()
        [void]$settingsView.GroupDescriptions.Add([System.Windows.Data.PropertyGroupDescription]::new('IntuneGroup'))
    }
if ($ui.ConflictsGrid)  { $ui.ConflictsGrid.ItemsSource  = $Script:AllConflicts }
if ($ui.DashTopConflictsGrid) { $ui.DashTopConflictsGrid.ItemsSource = $Script:TopConflicts }

$ui.VersionBadge.Text = "v$($Script:AppVersion)"
if ($ui.lblVersionBar) { $ui.lblVersionBar.Text = "v$($Script:AppVersion)" }


# ── Sidebar collapse / expand ──
$ui.BtnHamburger.Add_Click({
    $Script:SidebarCollapsed = -not $Script:SidebarCollapsed
    if ($Script:SidebarCollapsed) {
        $ui.SidebarColumn.Width = [System.Windows.GridLength]::new(0)
        $ui.SidebarBorder.Visibility = 'Collapsed'
        # hamburger icon stays the same (no toggle icon)
    } else {
        $ui.SidebarColumn.Width = [System.Windows.GridLength]::new(260)
        $ui.SidebarBorder.Visibility = 'Visible'
        # hamburger icon stays the same
    }
})

# ── Achievement panel collapse / expand ──
$ui.btnToggleAchievements.Add_Click({
    if ($ui.pnlAchievements.Visibility -eq 'Visible') {
        $ui.pnlAchievements.Visibility = 'Collapsed'
        $ui.txtAchievementChevron.RenderTransform = [System.Windows.Media.RotateTransform]::new(180)
        $Script:Prefs.AchievementsCollapsed = $true
    } else {
        $ui.pnlAchievements.Visibility = 'Visible'
        $ui.txtAchievementChevron.RenderTransform = [System.Windows.Media.RotateTransform]::new(0)
        $Script:Prefs.AchievementsCollapsed = $false
    }
    Save-Preferences
})

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 8: THEME ENGINE
# ═══════════════════════════════════════════════════════════════════════════════

$Script:ThemeDark = @{
    ThemeAppBg='#FF111113'; ThemePanelBg='#FF18181B'; ThemeCardBg='#F01E1E1E'
    ThemeInputBg='#FF141414'; ThemeDeepBg='#FF0D0D0D'
    ThemeAccent='#FF0078D4'; ThemeAccentHover='#FF1A8AD4'; ThemeAccentLight='#FF60CDFF'
    ThemeGreenAccent='#FF00C853'
    ThemeTextPrimary='#FFFFFFFF'; ThemeTextBody='#FFE0E0E0'; ThemeTextSecondary='#FFA1A1AA'
    ThemeTextMuted='#FFB0B0B8'; ThemeTextDim='#FFA0A0A8'
    ThemeHoverBg='#FF27272B'; ThemeSelectedBg='#FF2A2A2A'; ThemePressedBg='#FF1A1A1A'
    ThemeError='#FFFF5000'; ThemeWarning='#FFF59E0B'; ThemeSuccess='#FF00C853'
    ThemeBorder='#19FFFFFF'; ThemeBorderCard='#19FFFFFF'; ThemeBorderElevated='#FF333333'
    ThemeScrollTrack='#08FFFFFF'; ThemeScrollThumb='#40FFFFFF'
    ThemeOutputBg='#FF0F0F11'; ThemeBorderSubtle='#19FFFFFF'
    ThemeCardAltBg='#F01A1A1A'; ThemeSurfaceBg='#FF1F1F23'
    ThemeTextDisabled='#FF505058'; ThemeTextFaintest='#FF9A9AA2'
    ThemeAccentDim='#1A0078D4'; ThemeErrorDim='#20FF5000'
    ThemeBorderHover='#FF444444'
    ThemeSidebarBg='#FF111113'; ThemeSidebarBorder='#00000000'
    ThemeProgressEdge='#18FFFFFF'
    FilterGreen='#FF6BCB77'; FilterRed='#FFFF6B6B'; FilterYellow='#FFFFD93D'
    FilterCyan='#FF4FC3F7'; FilterPurple='#FFBA68C8'; FilterOrange='#FFFFB74D'
    FilterBlue='#FF90CAF9'; FilterLightGreen='#FFA5D6A7'; FilterLightPurple='#FFCE93D8'
    FilterBadgeBg='#FF2C2C30'; FilterBadgeText='#FFB0B0B8'
    AdminBannerBg='#30F59E0B'; CategoryBadgeBg='#409333EA'
}
$Script:ThemeLight = @{
    ThemeAppBg='#FFF5F5F5'; ThemePanelBg='#FFFFFFFF'; ThemeCardBg='#F0FFFFFF'
    ThemeInputBg='#FFF0F0F0'; ThemeDeepBg='#FFEBEBEB'
    ThemeAccent='#FF0067C0'; ThemeAccentHover='#FF1A7FD4'; ThemeAccentLight='#FF0078D4'
    ThemeGreenAccent='#FF00A844'
    ThemeTextPrimary='#FF1A1A1A'; ThemeTextBody='#FF333333'; ThemeTextSecondary='#FF666666'
    ThemeTextMuted='#FF999999'; ThemeTextDim='#FFAAAAAA'
    ThemeHoverBg='#FFEAEAEA'; ThemeSelectedBg='#FFE0E0E0'; ThemePressedBg='#FFD5D5D5'
    ThemeError='#FFDC3B00'; ThemeWarning='#FFD48A00'; ThemeSuccess='#FF00A844'
    ThemeBorder='#19000000'; ThemeBorderCard='#19000000'; ThemeBorderElevated='#FFCCCCCC'
    ThemeScrollTrack='#08000000'; ThemeScrollThumb='#40000000'
    ThemeOutputBg='#FFF8F8F8'; ThemeBorderSubtle='#19000000'
    ThemeCardAltBg='#F0F8F8F8'; ThemeSurfaceBg='#FFEAEAEA'
    ThemeTextDisabled='#FFCCCCCC'; ThemeTextFaintest='#FFB0B0B0'
    ThemeAccentDim='#1A0067C0'; ThemeErrorDim='#20DC3B00'
    ThemeBorderHover='#FFBBBBBB'
    ThemeSidebarBg='#FFF5F5F5'; ThemeSidebarBorder='#19000000'
    ThemeProgressEdge='#18000000'
    FilterGreen='#FF2E7D32'; FilterRed='#FFC62828'; FilterYellow='#FFE65100'
    FilterCyan='#FF0277BD'; FilterPurple='#FF7B1FA2'; FilterOrange='#FFBF360C'
    FilterBlue='#FF1565C0'; FilterLightGreen='#FF388E3C'; FilterLightPurple='#FF6A1B9A'
    FilterBadgeBg='#FFE0E0E0'; FilterBadgeText='#FF555555'
    AdminBannerBg='#30E8A000'; CategoryBadgeBg='#207B1FA2'
}

$Script:ThemeHighContrast = @{
    ThemeAppBg='#FF000000'; ThemePanelBg='#FF000000'; ThemeCardBg='#F0000000'
    ThemeInputBg='#FF000000'; ThemeDeepBg='#FF000000'
    ThemeAccent='#FF00FFFF'; ThemeAccentHover='#FF00CCCC'; ThemeAccentLight='#FF00FFFF'
    ThemeGreenAccent='#FF00FF00'
    ThemeTextPrimary='#FFFFFFFF'; ThemeTextBody='#FFFFFFFF'; ThemeTextSecondary='#FFFFFF00'
    ThemeTextMuted='#FFFFFF00'; ThemeTextDim='#FF00FFFF'
    ThemeHoverBg='#FF333333'; ThemeSelectedBg='#FF0000AA'; ThemePressedBg='#FF000066'
    ThemeError='#FFFF0000'; ThemeWarning='#FFFFFF00'; ThemeSuccess='#FF00FF00'
    ThemeBorder='#FFFFFFFF'; ThemeBorderCard='#FFFFFFFF'; ThemeBorderElevated='#FFFFFFFF'
    ThemeScrollTrack='#40FFFFFF'; ThemeScrollThumb='#FFFFFFFF'
    ThemeOutputBg='#FF000000'; ThemeBorderSubtle='#FFFFFFFF'
    ThemeCardAltBg='#F0111111'; ThemeSurfaceBg='#FF111111'
    ThemeTextDisabled='#FF808080'; ThemeTextFaintest='#FF808080'
    ThemeAccentDim='#4000FFFF'; ThemeErrorDim='#40FF0000'
    ThemeBorderHover='#FF00FFFF'
    ThemeSidebarBg='#FF000000'; ThemeSidebarBorder='#FFFFFFFF'
    ThemeProgressEdge='#40FFFFFF'
    FilterGreen='#FF00FF00'; FilterRed='#FFFF0000'; FilterYellow='#FFFFFF00'
    FilterCyan='#FF00FFFF'; FilterPurple='#FFFF00FF'; FilterOrange='#FFFFFF00'
    FilterBlue='#FF00FFFF'; FilterLightGreen='#FF00FF00'; FilterLightPurple='#FFFF00FF'
    FilterBadgeBg='#FFFFFFFF'; FilterBadgeText='#FF000000'
    AdminBannerBg='#40FFFF00'; CategoryBadgeBg='#40FFFFFF'
}

function Set-Theme([hashtable]$palette, [bool]$IsLight = $false) {

    $isLight = $IsLight



    # Swap dot grid color

    try {

        $dotBrush = $Window.Resources['DotGridBrush']

        if ($dotBrush -and $dotBrush.Drawing) {

            $dotColor = if ($isLight) { '#14000000' } else { '#1EFFFFFF' }

            $dotBrush.Drawing.Brush = (Get-CachedBrush $dotColor)

        }

    } catch { <# glow rebuild non-critical #> }



    # Rebuild radial glow

    try {

        $contentGrid = $ui.ContentArea

        if ($contentGrid) {

            foreach ($child in $contentGrid.Children) {

                if ($child -is [System.Windows.Controls.Border] -and $child.Background -is [System.Windows.Media.RadialGradientBrush]) {

                    $rg = [System.Windows.Media.RadialGradientBrush]::new()

                    $rg.Center = [System.Windows.Point]::new(0.5, 0); $rg.GradientOrigin = [System.Windows.Point]::new(0.5, 0)

                    $rg.RadiusX = 0.8; $rg.RadiusY = 0.5

                    if ($isLight) {

                        [void]$rg.GradientStops.Add([System.Windows.Media.GradientStop]::new([System.Windows.Media.ColorConverter]::ConvertFromString('#0A00C853'), 0.0))

                        [void]$rg.GradientStops.Add([System.Windows.Media.GradientStop]::new([System.Windows.Media.ColorConverter]::ConvertFromString('#00FFFFFF'), 1.0))

                    } else {

                        [void]$rg.GradientStops.Add([System.Windows.Media.GradientStop]::new([System.Windows.Media.ColorConverter]::ConvertFromString('#1800C853'), 0.0))

                        [void]$rg.GradientStops.Add([System.Windows.Media.GradientStop]::new([System.Windows.Media.ColorConverter]::ConvertFromString('#00000000'), 1.0))

                    }

                    $child.Background = $rg; break

                }

            }

        }

    } catch { <# glow rebuild non-critical #> }



    # Rebuild linear gradient glow

    try {

        $contentGrid = $ui.ContentArea

        if ($contentGrid) {

            foreach ($child in $contentGrid.Children) {

                if ($child -is [System.Windows.Controls.Border] -and $child.Background -is [System.Windows.Media.LinearGradientBrush]) {

                    $lg = [System.Windows.Media.LinearGradientBrush]::new()

                    $lg.StartPoint = [System.Windows.Point]::new(0.5, 0); $lg.EndPoint = [System.Windows.Point]::new(0.5, 1)

                    if ($isLight) {

                        [void]$lg.GradientStops.Add([System.Windows.Media.GradientStop]::new([System.Windows.Media.ColorConverter]::ConvertFromString('#180078D4'), 0.0))

                        [void]$lg.GradientStops.Add([System.Windows.Media.GradientStop]::new([System.Windows.Media.ColorConverter]::ConvertFromString('#0C0078D4'), 0.35))

                        [void]$lg.GradientStops.Add([System.Windows.Media.GradientStop]::new([System.Windows.Media.ColorConverter]::ConvertFromString('#060078D4'), 0.6))

                        [void]$lg.GradientStops.Add([System.Windows.Media.GradientStop]::new([System.Windows.Media.ColorConverter]::ConvertFromString('#00FFFFFF'), 1.0))

                    } else {

                        [void]$lg.GradientStops.Add([System.Windows.Media.GradientStop]::new([System.Windows.Media.ColorConverter]::ConvertFromString('#200078D4'), 0.0))

                        [void]$lg.GradientStops.Add([System.Windows.Media.GradientStop]::new([System.Windows.Media.ColorConverter]::ConvertFromString('#100078D4'), 0.35))

                        [void]$lg.GradientStops.Add([System.Windows.Media.GradientStop]::new([System.Windows.Media.ColorConverter]::ConvertFromString('#080078D4'), 0.6))

                        [void]$lg.GradientStops.Add([System.Windows.Media.GradientStop]::new([System.Windows.Media.ColorConverter]::ConvertFromString('#00000000'), 1.0))

                    }

                    $child.Background = $lg; break

                }

            }

        }

    } catch { <# glow rebuild non-critical #> }



    # Apply all brush keys

    foreach ($kv in $palette.GetEnumerator()) {

        $newBrush = (Get-CachedBrush $kv.Value)

        $Window.Resources[$kv.Key] = $newBrush

    }

    # Repaint achievements (they use FindResource snapshots, not DynamicResource)
    Render-Achievements
}

# Apply saved theme
if ($Script:Prefs.IsLightMode) { Set-Theme $Script:ThemeLight -IsLight $true } else { Set-Theme $Script:ThemeDark -IsLight $false }

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 9: TOAST NOTIFICATION
# ═══════════════════════════════════════════════════════════════════════════════

function Set-ToastType([string]$type) {
    if (-not $ui.ToastIcon -or -not $ui.ToastAccentBar) { return }
    switch ($type) {
        'success' {
            $ui.ToastIcon.Text = [char]0xE73E
            $ui.ToastIcon.Foreground = $Window.Resources['ThemeSuccess']
            $ui.ToastAccentBar.Background = $Window.Resources['ThemeSuccess']
        }
        'warning' {
            $ui.ToastIcon.Text = [char]0xE7BA
            $ui.ToastIcon.Foreground = $Window.Resources['ThemeWarning']
            $ui.ToastAccentBar.Background = $Window.Resources['ThemeWarning']
        }
        'error' {
            $ui.ToastIcon.Text = [char]0xEA39
            $ui.ToastIcon.Foreground = $Window.Resources['ThemeError']
            $ui.ToastAccentBar.Background = $Window.Resources['ThemeError']
        }
        default {
            $ui.ToastIcon.Text = [char]0xE946
            $ui.ToastIcon.Foreground = $Window.Resources['ThemeAccent']
            $ui.ToastAccentBar.Background = $Window.Resources['ThemeAccent']
        }
    }
}

function Show-Toast([string]$title, [string]$message, [string]$type = 'info') {
    Set-ToastType $type
    if (-not $ui.ToastBorder) { return }
    $ui.ToastTitle.Text   = $title
    $ui.ToastMessage.Text = $message
    $ui.ToastBorder.Visibility = 'Visible'
    $ui.ToastBorder.IsHitTestVisible = $true

    $tt = $ui.ToastBorder.RenderTransform
    if (-not $Script:AnimationsDisabled -and $tt -is [System.Windows.Media.TranslateTransform]) {
        $tt.X = 400
        $anim = [System.Windows.Media.Animation.DoubleAnimation]::new(400, 0,
            [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(350)))
        $anim.EasingFunction = [System.Windows.Media.Animation.CubicEase]::new()
        $anim.EasingFunction.EasingMode = 'EaseOut'
        $tt.BeginAnimation([System.Windows.Media.TranslateTransform]::XProperty, $anim)
    } elseif ($tt -is [System.Windows.Media.TranslateTransform]) { $tt.X = 0 }

    if ($Script:ToastTimer)     { $Script:ToastTimer.Stop() }
    if ($Script:ToastHideTimer) { $Script:ToastHideTimer.Stop() }

    $ToastRef = $ui.ToastBorder

    $Script:ToastHideTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $Script:ToastHideTimer.Interval = [TimeSpan]::FromMilliseconds(350)
    $Script:ToastHideTimer.Add_Tick({
        $this.Stop()
        $ToastRef.Visibility = 'Collapsed'
    }.GetNewClosure())

    $HideTimerRef = $Script:ToastHideTimer
    $Script:ToastTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $Script:ToastTimer.Interval = [TimeSpan]::FromSeconds(4)
    $Script:ToastTimer.Add_Tick({
        $this.Stop()
        $ttInner = $ToastRef.RenderTransform
        if (-not $Script:AnimationsDisabled -and $ttInner -is [System.Windows.Media.TranslateTransform]) {
            $anim = [System.Windows.Media.Animation.DoubleAnimation]::new(0, 400,
                [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(300)))
            $anim.EasingFunction = [System.Windows.Media.Animation.CubicEase]::new()
            $anim.EasingFunction.EasingMode = 'EaseIn'
            $ttInner.BeginAnimation([System.Windows.Media.TranslateTransform]::XProperty, $anim)
            $HideTimerRef.Start()
        } else { $ToastRef.Visibility = 'Collapsed' }
    }.GetNewClosure())
    $Script:ToastTimer.Start()

    if (-not $Script:ToastClickWired) {
        $Script:ToastClickWired = $true
        $ToastBorderRef = $ui.ToastBorder
        $ui.ToastBorder.Add_MouseLeftButtonDown({
            if ($Script:ToastTimer) { $Script:ToastTimer.Stop() }
            $ttInner = $ToastBorderRef.RenderTransform
            if (-not $Script:AnimationsDisabled -and $ttInner -is [System.Windows.Media.TranslateTransform]) {
                $anim = [System.Windows.Media.Animation.DoubleAnimation]::new(0, 400,
                    [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(200)))
                $anim.EasingFunction = [System.Windows.Media.Animation.CubicEase]::new()
                $anim.EasingFunction.EasingMode = 'EaseIn'
                $ttInner.BeginAnimation([System.Windows.Media.TranslateTransform]::XProperty, $anim)
                if ($Script:ToastHideTimer) { $Script:ToastHideTimer.Start() }
            } else { $ToastBorderRef.Visibility = 'Collapsed' }
        }.GetNewClosure())
    }
    Write-DebugLog "Toast: $title - $message" -Level DEBUG
}

# --- Themed MessageBox replacement (dark/light mode aware) ---
function Show-ThemedMessageBox {
    param([string]$Message, [string]$Title = 'PolicyPilot', [string]$Buttons = 'OK',
          [string]$Icon = 'Info', [string]$DefaultButton = 'OK')
    $Palette = if ($Script:Prefs.IsLightMode) { $Script:ThemeLight } else { $Script:ThemeDark }
    $Br = { param([string]$Key) (Get-CachedBrush $Palette[$Key]) }
    $IconChar = switch ($Icon) { 'Warning' { [char]0xE7BA } 'Error' { [char]0xEA39 } 'Question' { '?' } default { [char]0xE946 } }
    $IconColor = switch ($Icon) { 'Warning' { '#FFFFC107' } 'Error' { '#FFEF5350' } 'Question' { '#FF42A5F5' } default { '#FF42A5F5' } }

    $Dlg = New-Object System.Windows.Window
    $Dlg.Title = $Title; $Dlg.SizeToContent = 'WidthAndHeight'; $Dlg.MinWidth = 380; $Dlg.MaxWidth = 520
    $Dlg.ResizeMode = 'NoResize'; $Dlg.WindowStartupLocation = 'CenterOwner'
    $Dlg.Owner = $Window; $Dlg.Background = (& $Br 'ThemeCardBg')
    $Dlg.WindowStyle = 'None'; $Dlg.AllowsTransparency = $true

    $Border = New-Object System.Windows.Controls.Border
    $Border.Background = (& $Br 'ThemeCardBg'); $Border.BorderBrush = (& $Br 'ThemeBorderElevated')
    $Border.BorderThickness = [System.Windows.Thickness]::new(1); $Border.CornerRadius = [System.Windows.CornerRadius]::new(12)
    $Border.Padding = [System.Windows.Thickness]::new(24, 20, 24, 18)
    $Border.Effect = [System.Windows.Media.Effects.DropShadowEffect]@{ Color = [System.Windows.Media.Colors]::Black; Direction = 270; ShadowDepth = 6; BlurRadius = 24; Opacity = 0.4 }

    $Stack = New-Object System.Windows.Controls.StackPanel

    # Header with icon
    $HdrPanel = New-Object System.Windows.Controls.StackPanel; $HdrPanel.Orientation = 'Horizontal'
    $HdrPanel.Margin = [System.Windows.Thickness]::new(0, 0, 0, 14)
    $IconBadge = New-Object System.Windows.Controls.Border; $IconBadge.Width = 32; $IconBadge.Height = 32
    $IconBadge.CornerRadius = [System.Windows.CornerRadius]::new(8); $IconBadge.Background = (Get-CachedBrush ($IconColor -replace 'FF','33'))
    $IconTB = New-Object System.Windows.Controls.TextBlock; $IconTB.Text = $IconChar; $IconTB.FontSize = 16
    $IconTB.FontFamily = [System.Windows.Media.FontFamily]::new('Segoe MDL2 Assets'); $IconTB.Foreground = (Get-CachedBrush $IconColor)
    $IconTB.HorizontalAlignment = 'Center'; $IconTB.VerticalAlignment = 'Center'; $IconBadge.Child = $IconTB
    [void]$HdrPanel.Children.Add($IconBadge)
    $TitleTB = New-Object System.Windows.Controls.TextBlock; $TitleTB.Text = $Title; $TitleTB.FontSize = 14; $TitleTB.FontWeight = 'SemiBold'
    $TitleTB.Foreground = (& $Br 'ThemeTextPrimary'); $TitleTB.VerticalAlignment = 'Center'; $TitleTB.Margin = [System.Windows.Thickness]::new(12,0,0,0)
    [void]$HdrPanel.Children.Add($TitleTB)
    [void]$Stack.Children.Add($HdrPanel)

    # Separator
    $Sep = New-Object System.Windows.Controls.Border; $Sep.Height = 1; $Sep.Background = (& $Br 'ThemeBorder')
    $Sep.Margin = [System.Windows.Thickness]::new(0, 0, 0, 14); [void]$Stack.Children.Add($Sep)

    # Message body
    $MsgTB = New-Object System.Windows.Controls.TextBlock; $MsgTB.Text = $Message; $MsgTB.FontSize = 12.5
    $MsgTB.Foreground = (& $Br 'ThemeTextSecondary'); $MsgTB.TextWrapping = 'Wrap'; $MsgTB.Margin = [System.Windows.Thickness]::new(0, 0, 0, 18)
    [void]$Stack.Children.Add($MsgTB)

    # Buttons
    $BtnPanel = New-Object System.Windows.Controls.StackPanel; $BtnPanel.Orientation = 'Horizontal'; $BtnPanel.HorizontalAlignment = 'Right'
    $DlgRef = $Dlg
    $mkBtn = {
        param([string]$text, [bool]$isPrimary, [string]$tag)
        $b = New-Object System.Windows.Controls.Button; $b.Content = $text; $b.MinWidth = 80; $b.Tag = $tag
        $b.Padding = [System.Windows.Thickness]::new(16, 8, 16, 8); $b.FontSize = 12; $b.Cursor = [System.Windows.Input.Cursors]::Hand
        $b.BorderThickness = [System.Windows.Thickness]::new(1); $b.Margin = [System.Windows.Thickness]::new(6, 0, 0, 0)
        if ($isPrimary) {
            $b.Background = (& $Br 'ThemeAccent'); $b.Foreground = (Get-CachedBrush '#FFFFFFFF'); $b.BorderBrush = (& $Br 'ThemeAccent')
        } else {
            $b.Background = (& $Br 'ThemeCardAltBg'); $b.Foreground = (& $Br 'ThemeTextBody'); $b.BorderBrush = (& $Br 'ThemeBorder')
        }
        $b
    }
    $dlgState = @{ Result = if ($Buttons -eq 'YesNo') { 'No' } else { 'OK' } }
    if ($Buttons -eq 'YesNo') {
        $noBtn = & $mkBtn 'No' $false 'No'; $noBtn.Add_Click({ $dlgState.Result = 'No'; $DlgRef.Close() }.GetNewClosure())
        [void]$BtnPanel.Children.Add($noBtn)
        $yesBtn = & $mkBtn 'Yes' $true 'Yes'; $yesBtn.Add_Click({ $dlgState.Result = 'Yes'; $DlgRef.Close() }.GetNewClosure())
        [void]$BtnPanel.Children.Add($yesBtn)
    } else {
        $okBtn = & $mkBtn 'OK' $true 'OK'; $okBtn.Add_Click({ $dlgState.Result = 'OK'; $DlgRef.Close() }.GetNewClosure())
        [void]$BtnPanel.Children.Add($okBtn)
    }
    [void]$Stack.Children.Add($BtnPanel)

    $Border.Child = $Stack; $Dlg.Content = $Border
    $Dlg.ShowDialog() | Out-Null
    return $dlgState.Result
}

# --- Themed InputBox replacement (dark/light mode aware) ---
function Show-ThemedInputBox {
    param([string]$Prompt, [string]$Title = 'Input', [string]$DefaultValue = '')
    $Palette = if ($Script:Prefs.IsLightMode) { $Script:ThemeLight } else { $Script:ThemeDark }
    $Br = { param([string]$Key) (Get-CachedBrush $Palette[$Key]) }

    $Dlg = New-Object System.Windows.Window
    $Dlg.Title = $Title; $Dlg.Width = 400; $Dlg.SizeToContent = 'Height'
    $Dlg.ResizeMode = 'NoResize'; $Dlg.WindowStartupLocation = 'CenterOwner'
    $Dlg.Owner = $Window; $Dlg.Background = (& $Br 'ThemeCardBg')
    $Dlg.WindowStyle = 'None'; $Dlg.AllowsTransparency = $true

    $Border = New-Object System.Windows.Controls.Border
    $Border.Background = (& $Br 'ThemeCardBg'); $Border.BorderBrush = (& $Br 'ThemeBorderElevated')
    $Border.BorderThickness = [System.Windows.Thickness]::new(1); $Border.CornerRadius = [System.Windows.CornerRadius]::new(12)
    $Border.Padding = [System.Windows.Thickness]::new(24, 20, 24, 18)
    $Border.Effect = [System.Windows.Media.Effects.DropShadowEffect]@{ Color = [System.Windows.Media.Colors]::Black; Direction = 270; ShadowDepth = 6; BlurRadius = 24; Opacity = 0.4 }

    $Stack = New-Object System.Windows.Controls.StackPanel

    # Title
    $TitleTB = New-Object System.Windows.Controls.TextBlock; $TitleTB.Text = $Title; $TitleTB.FontSize = 14; $TitleTB.FontWeight = 'SemiBold'
    $TitleTB.Foreground = (& $Br 'ThemeTextPrimary'); $TitleTB.Margin = [System.Windows.Thickness]::new(0, 0, 0, 10)
    [void]$Stack.Children.Add($TitleTB)

    $Sep = New-Object System.Windows.Controls.Border; $Sep.Height = 1; $Sep.Background = (& $Br 'ThemeBorder')
    $Sep.Margin = [System.Windows.Thickness]::new(0, 0, 0, 14); [void]$Stack.Children.Add($Sep)

    # Prompt
    $PromptTB = New-Object System.Windows.Controls.TextBlock; $PromptTB.Text = $Prompt; $PromptTB.FontSize = 12
    $PromptTB.Foreground = (& $Br 'ThemeTextSecondary'); $PromptTB.TextWrapping = 'Wrap'; $PromptTB.Margin = [System.Windows.Thickness]::new(0, 0, 0, 10)
    [void]$Stack.Children.Add($PromptTB)

    # Input TextBox
    $InputTB = New-Object System.Windows.Controls.TextBox; $InputTB.Text = $DefaultValue; $InputTB.FontSize = 12.5
    $InputTB.Padding = [System.Windows.Thickness]::new(8, 6, 8, 6); $InputTB.Margin = [System.Windows.Thickness]::new(0, 0, 0, 18)
    $InputTB.Background = (& $Br 'ThemeCardAltBg'); $InputTB.Foreground = (& $Br 'ThemeTextPrimary')
    $InputTB.BorderBrush = (& $Br 'ThemeBorder'); $InputTB.BorderThickness = [System.Windows.Thickness]::new(1)
    $InputTB.CaretBrush = (& $Br 'ThemeTextPrimary')
    [void]$Stack.Children.Add($InputTB)

    # Buttons
    $BtnPanel = New-Object System.Windows.Controls.StackPanel; $BtnPanel.Orientation = 'Horizontal'; $BtnPanel.HorizontalAlignment = 'Right'
    $DlgRef = $Dlg; $InputRef = $InputTB
    $cancelBtn = New-Object System.Windows.Controls.Button; $cancelBtn.Content = 'Cancel'; $cancelBtn.MinWidth = 80
    $cancelBtn.Padding = [System.Windows.Thickness]::new(16, 8, 16, 8); $cancelBtn.FontSize = 12; $cancelBtn.Cursor = [System.Windows.Input.Cursors]::Hand
    $cancelBtn.Background = (& $Br 'ThemeCardAltBg'); $cancelBtn.Foreground = (& $Br 'ThemeTextBody')
    $cancelBtn.BorderBrush = (& $Br 'ThemeBorder'); $cancelBtn.BorderThickness = [System.Windows.Thickness]::new(1)
    $cancelBtn.Add_Click({ $Script:_inputResult = ''; $DlgRef.Close() }.GetNewClosure())
    [void]$BtnPanel.Children.Add($cancelBtn)
    $okBtn = New-Object System.Windows.Controls.Button; $okBtn.Content = 'OK'; $okBtn.MinWidth = 80
    $okBtn.Padding = [System.Windows.Thickness]::new(16, 8, 16, 8); $okBtn.FontSize = 12; $okBtn.Cursor = [System.Windows.Input.Cursors]::Hand
    $okBtn.Background = (& $Br 'ThemeAccent'); $okBtn.Foreground = (Get-CachedBrush '#FFFFFFFF')
    $okBtn.BorderBrush = (& $Br 'ThemeAccent'); $okBtn.BorderThickness = [System.Windows.Thickness]::new(1)
    $okBtn.Margin = [System.Windows.Thickness]::new(6, 0, 0, 0)
    $okBtn.Add_Click({ $Script:_inputResult = $InputRef.Text; $DlgRef.Close() }.GetNewClosure())
    [void]$BtnPanel.Children.Add($okBtn)
    [void]$Stack.Children.Add($BtnPanel)

    $Border.Child = $Stack; $Dlg.Content = $Border
    $Script:_inputResult = ''
    $Dlg.Add_ContentRendered({ $InputRef.Focus(); $InputRef.SelectAll() }.GetNewClosure())
    $Dlg.ShowDialog() | Out-Null
    return $Script:_inputResult
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 10: HELPER FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════════

function Set-Status([string]$text, [string]$color = '#FF8B8B93') {
    $ui.StatusText.Text = $text
    try {
        $brush = (Get-CachedBrush $color)
        $ui.StatusDot.Fill = $brush
    } catch { try { Write-DebugLog "Unhandled: $_" -Level ERROR } catch {} }
    # Mirror to status bar
    if ($ui.lblStatusBar) { $ui.lblStatusBar.Text = $text }
    if ($ui.statusBarDot) {
        try { $ui.statusBarDot.Fill = $brush } catch { try { Write-DebugLog "Unhandled: $_" -Level ERROR } catch {} }
    }
}

function Invoke-CountUp([System.Windows.Controls.TextBlock]$textBlock, [int]$target) {
    $key = $textBlock.Name
    if ($Script:AnimationsDisabled -or $target -le 0) { $textBlock.Text = "$target"; return }
    $steps = [Math]::Min(15, $target)
    $Script:CountUpAnims[$key] = @{ step = 0; target = $target; steps = $steps; tb = $textBlock }
    # Shared master timer - no .GetNewClosure(), so $Script: resolves to PolicyPilot scope
    if (-not $Script:CountUpMasterTimer) {
        $Script:CountUpMasterTimer = [System.Windows.Threading.DispatcherTimer]::new()
        $Script:CountUpMasterTimer.Interval = [TimeSpan]::FromMilliseconds(25)
        $Script:CountUpMasterTimer.Add_Tick({
            $anyActive = $false
            foreach ($k in @($Script:CountUpAnims.Keys)) {
                $s = $Script:CountUpAnims[$k]
                if ($s.step -ge $s.steps) { continue }
                $anyActive = $true
                $s.step++
                $progress = [Math]::Min(1.0, $s.step / $s.steps)
                $eased = 1 - (1 - $progress) * (1 - $progress)
                $val = [Math]::Round($s.target * $eased)
                $s.tb.Text = "$val"
                if ($s.step -ge $s.steps) { $s.tb.Text = "$($s.target)" }
            }
            if (-not $anyActive) { $Script:CountUpMasterTimer.Stop() }
        })
    }
    $Script:CountUpMasterTimer.Start()
}

function Invoke-TabFade($element) {
    if (-not $element) { return }
    $element.Visibility = 'Visible'
    if ($Script:AnimationsDisabled) { $element.Opacity = 1; return }
    $element.Opacity = 0
    $fadeIn = [System.Windows.Media.Animation.DoubleAnimation]::new(0, 1,
        [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(200)))
    $fadeIn.EasingFunction = [System.Windows.Media.Animation.CubicEase]::new()
    $fadeIn.EasingFunction.EasingMode = 'EaseOut'
    $element.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $fadeIn)
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 11: NAVIGATION
# ═══════════════════════════════════════════════════════════════════════════════

function Switch-Tab([string]$tab) {
    # Stop live tails when navigating away from their tab
    if ($Script:ActiveTab -eq 'MDMSync' -and $tab -ne 'MDMSync') {
        if ($Script:MdmTailing) { Stop-MdmTail }
    }
    if ($Script:ActiveTab -eq 'IMELogs' -and $tab -ne 'IMELogs') {
        if ($Script:ImeTailing) { Stop-ImeTail }
    }
    if ($Script:ActiveTab -eq 'GPOLogs' -and $tab -ne 'GPOLogs') {
        if ($Script:GpoTailing) { Stop-GpoTail }
    }
    $Script:ActiveTab = $tab
    foreach ($t in $Script:Tabs) {
        $panel   = $ui["Panel$t"]
        $sidebar = $ui["Sidebar$t"]
        $nav     = $ui["Nav$t"]
        if ($t -eq $tab) {
            if ($panel)   { Invoke-TabFade $panel }
            if ($sidebar) { Invoke-TabFade $sidebar }
        } else {
            if ($panel -and $panel.Visibility -eq 'Visible') {
                if (-not $Script:AnimationsDisabled) { $panel.Opacity = 0 }
                $panel.Visibility = 'Collapsed'
            }
            if ($sidebar -and $sidebar.Visibility -eq 'Visible') {
                if (-not $Script:AnimationsDisabled) { $sidebar.Opacity = 0 }
                $sidebar.Visibility = 'Collapsed'
            }
        }
        if ($nav) { $nav.Tag = if ($t -eq $tab) { 'Active' } else { $null } }
    }
    Set-Status $tab '#00C853'
    Write-DebugLog "Navigated to $tab" -Level STEP
}

# Wire nav buttons
$ui.NavDashboard.Add_Click({ Switch-Tab 'Dashboard' })
$ui.NavGPOList.Add_Click({ Switch-Tab 'GPOList' })
$ui.NavSettings.Add_Click({ Switch-Tab 'Settings' })
$ui.NavConflicts.Add_Click({ Switch-Tab 'Conflicts' })
$ui.NavIntuneApps.Add_Click({ Switch-Tab 'IntuneApps' })
$ui.NavReport.Add_Click({ Switch-Tab 'Report'; Update-ReportPreview })
$ui.NavIMELogs.Add_Click({
    Switch-Tab 'IMELogs'
    if ($ui.ChkImeLiveTail -and $ui.ChkImeLiveTail.IsChecked) {
        if (-not $Script:ImeTailing) { Write-Host '[IME] Auto-starting tail on tab switch'; Start-ImeTail }
    } else {
        if ($Script:ImeStats.Lines -eq 0) { Write-Host '[IME] Auto-loading log on tab switch'; Load-ImeLogFile }
    }
})
$ui.NavGPOLogs.Add_Click({
    Switch-Tab 'GPOLogs'
    if ($ui.ChkGpoLiveTail -and $ui.ChkGpoLiveTail.IsChecked) {
        if (-not $Script:GpoTailing) { Write-Host '[GPO] Auto-starting tail on tab switch'; Start-GpoTail }
    } else {
        if ($Script:GpoStats.Lines -eq 0) { Write-Host '[GPO] Auto-loading log on tab switch'; Load-GpoLogFile }
    }
})
$ui.NavTools.Add_Click({
    Write-DebugLog 'NavTools clicked - calling Switch-Tab Tools' -Level DEBUG
    Switch-Tab 'Tools'
})
$ui.NavAppSettings.Add_Click({ Switch-Tab 'AppSettings' })

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 12: UPDATE DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════════

function Update-Dashboard {
    if (-not $Script:ScanData) {
        $ui.DashEmptyState.Visibility = 'Visible'
        $ui.DashDataState.Visibility  = 'Collapsed'
        foreach ($tb in @($ui.StatActiveGPOs, $ui.StatTotalSettings, $ui.StatConflicts, $ui.StatRedundant, $ui.StatUnlinked)) {
            if ($tb) { $tb.Text = '-' }
        }
        $ui.SidebarActiveGPOs.Text    = '-'
        $ui.SidebarTotalSettings.Text = '-'
        # Sidebar conflict/redundant stats
        $ui.SidebarConflictsStat.Text = '-'
        $ui.DashLastScanTime.Text = 'No scan performed'
        $ui.DashDomainName.Text   = ''
        # Disable exports
        $ui.BtnExportHtml.IsEnabled        = $false
        $ui.BtnExportCsv.IsEnabled         = $false
        $ui.BtnExportConflictsCsv.IsEnabled = $false
        $ui.BtnSaveSnapshot.IsEnabled       = $false
        return
    }

    $gpos     = $Script:ScanData.GPOs
    $settings = $Script:ScanData.Settings
    $conflicts = $Script:AllConflicts

    $activeCount = @($gpos | Where-Object { $_.Status -ne 'Disabled' }).Count
    $totalSettings = $settings.Count
    $conflictCount   = @($conflicts | Where-Object Severity -eq 'Conflict').Count
    $redundantCount  = @($conflicts | Where-Object Severity -eq 'Redundant').Count
    $unlinkedCount   = @($gpos | Where-Object { -not $_.IsLinked }).Count

    # Stat cards
    Invoke-CountUp $ui.StatActiveGPOs     $activeCount
    Invoke-CountUp $ui.StatTotalSettings  $totalSettings
    Invoke-CountUp $ui.StatConflicts      $conflictCount
    Invoke-CountUp $ui.StatRedundant      $redundantCount
    Invoke-CountUp $ui.StatUnlinked       $unlinkedCount

    # Sidebar quick stats
    $ui.SidebarActiveGPOs.Text    = "$activeCount"
    $ui.SidebarTotalSettings.Text = "$totalSettings"

    # Scan info
    $ui.DashLastScanTime.Text = if ($Script:ScanData.Timestamp -is [datetime]) { $Script:ScanData.Timestamp.ToString('yyyy-MM-dd HH:mm:ss') } else { $Script:ScanData.Timestamp }
    $ui.DashDomainName.Text   = $Script:ScanData.Domain
    $ui.DomainBadge.Text      = $Script:ScanData.Domain

    # MDM Enrollment info
    if ($scanResult.MdmInfo -and $scanResult.MdmInfo.ProviderID -and $ui.MdmEnrollmentCard) {
        $mi = $scanResult.MdmInfo
        $ui.MdmEnrollmentCard.Visibility = 'Visible'
        $ui.MdmProviderText.Text    = $mi.ProviderID
        $ui.MdmUPNText.Text         = $mi.EnrollmentUPN
        $ui.MdmEnrollTypeText.Text  = $mi.EnrollmentType
        $ui.MdmStateText.Text       = $mi.EnrollmentState
        Write-DebugLog "OnComplete: MDM enrolled via $($mi.ProviderID), UPN=$($mi.EnrollmentUPN), Certs=$($mi.MdmDiag.Certificates.Count)" -Level INFO

        # LAPS card
        if ($mi.MdmDiag.LAPS -and $mi.MdmDiag.LAPS.BackupDirectory -and $mi.MdmDiag.LAPS.BackupDirectory -ne 'Not Configured') {
            if ($ui.LapsStatusCard) { $ui.LapsStatusCard.Visibility = 'Visible' }
            if ($ui.LapsBackupText)      { $ui.LapsBackupText.Text      = $mi.MdmDiag.LAPS.BackupDirectory }
            if ($ui.LapsPasswordAgeText) { $ui.LapsPasswordAgeText.Text = $mi.MdmDiag.LAPS.PasswordAgeDays }
            if ($ui.LapsComplexityText)  { $ui.LapsComplexityText.Text  = $mi.MdmDiag.LAPS.PasswordComplexity }
            if ($ui.LapsPostAuthText)    { $ui.LapsPostAuthText.Text    = $mi.MdmDiag.LAPS.PostAuthActions }
            $lapsRotation = if ($mi.MdmDiag.LAPS.LocalLastPasswordUpdate) { "Last rotation: $($mi.MdmDiag.LAPS.LocalLastPasswordUpdate)" } else { '' }
            if ($mi.MdmDiag.LAPS.LocalAzurePasswordExpiry) { $lapsRotation += if ($lapsRotation) { "`nExpiry: $($mi.MdmDiag.LAPS.LocalAzurePasswordExpiry)" } else { "Expiry: $($mi.MdmDiag.LAPS.LocalAzurePasswordExpiry)" } }
            if ($ui.LapsLastRotationText) { $ui.LapsLastRotationText.Text = $lapsRotation }
            Write-DebugLog "OnComplete: LAPS card populated - Backup=$($mi.MdmDiag.LAPS.BackupDirectory)" -Level DEBUG
        }

        # Certificate inventory card
        $certs = $mi.MdmDiag.Certificates
        if ($certs -and $certs.Count -gt 0) {
            if ($ui.CertInventoryCard) { $ui.CertInventoryCard.Visibility = 'Visible' }
            if ($ui.CertCountText)    { $ui.CertCountText.Text    = "$($certs.Count)" }
            $expiring = @($certs | Where-Object { $_.ExpireDays -ne '' -and [int]$_.ExpireDays -le 30 })
            if ($ui.CertExpiringText) { $ui.CertExpiringText.Text = "$($expiring.Count)" }
            $detailLines = @($certs | ForEach-Object {
                $exp = if ($_.ExpireDays -ne '') { "($($_.ExpireDays)d)" } else { '' }
                "$($_.Store): $($_.IssuedTo) $exp"
            })
            if ($ui.CertDetailText) { $ui.CertDetailText.Text = ($detailLines | Select-Object -First 5) -join "`n" }
            Write-DebugLog "OnComplete: Cert card - $($certs.Count) certs, $($expiring.Count) expiring within 30d" -Level DEBUG
        }

        # N1: App install summary card
        $appSum = $mi.MdmDiag.AppSummary
        if ($appSum -and $appSum.Total -gt 0) {
            if ($ui.AppStatusCard) { $ui.AppStatusCard.Visibility = 'Visible' }
            if ($ui.AppInstalledText)   { $ui.AppInstalledText.Text   = "$($appSum.Installed)/$($appSum.Total)" }
            if ($ui.AppFailedText)      { $ui.AppFailedText.Text      = "$($appSum.Failed)" }
            if ($ui.AppPendingText)     { $ui.AppPendingText.Text     = "$($appSum.Pending)" }
            if ($ui.AppSuccessRateText) { $ui.AppSuccessRateText.Text = "Success rate: $($appSum.SuccessRate)%" }
            Write-DebugLog "OnComplete: App card - $($appSum.Installed)/$($appSum.Total) installed, $($appSum.Failed) failed" -Level DEBUG
        }

        # N5: Compliance health card
        $comp = $mi.MdmDiag.Compliance
        if ($comp -and $comp.Status -ne 'N/A') {
            if ($ui.ComplianceCard) { $ui.ComplianceCard.Visibility = 'Visible' }
            if ($ui.ComplianceStatusBadge) {
                $ui.ComplianceStatusBadge.Text = $comp.Status
                $badgeBg = switch ($comp.Status) {
                    'Compliant'     { '#22C55E' }
                    'Non-compliant' { '#EF4444' }
                    'At Risk'       { '#F59E0B' }
                    default         { '#71717A' }
                }
                $ui.ComplianceStatusBadge.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString($badgeBg)
            }
            if ($ui.CompliancePolicyCountText) { $ui.CompliancePolicyCountText.Text = "$($comp.ConfiguredPolicies)" }
            if ($ui.ComplianceAppRateText -and $comp.AppSuccessRate) { $ui.ComplianceAppRateText.Text = "$($comp.AppSuccessRate)%" }
            if ($ui.ComplianceIssuesText -and $comp.Issues.Count -gt 0) {
                $ui.ComplianceIssuesText.Text = ($comp.Issues | ForEach-Object { "⚠ $_" }) -join "`n"
            }
            Write-DebugLog "OnComplete: Compliance card - $($comp.Status), $($comp.Issues.Count) issues" -Level DEBUG
        }

        # N4: Script policies card
        $scripts = $mi.MdmDiag.ScriptPolicies
        if ($scripts -and $scripts.Count -gt 0) {
            if ($ui.ScriptPoliciesCard) { $ui.ScriptPoliciesCard.Visibility = 'Visible' }
            if ($ui.ScriptTotalText) { $ui.ScriptTotalText.Text = "$($scripts.Count)" }
            $sFailed = @($scripts | Where-Object { $_.Result -match 'Fail' }).Count
            if ($ui.ScriptFailedText) { $ui.ScriptFailedText.Text = "$sFailed" }
            $sDetail = @($scripts | Where-Object { $_.ScriptName } | ForEach-Object {
                $icon = if ($_.Result -match 'Success') { [char]0x2705 } else { [char]0x274C }
                "$icon $($_.ScriptName) ($($_.ScriptType))"
            })
            if ($ui.ScriptDetailText -and $sDetail.Count -gt 0) { $ui.ScriptDetailText.Text = ($sDetail | Select-Object -First 5) -join "`n" }
            Write-DebugLog "OnComplete: Script policies - $($scripts.Count) total, $sFailed failed" -Level DEBUG
        }

        # N8: Config profiles card
        $profiles = $mi.MdmDiag.ConfigProfiles
        if ($profiles -and $profiles.Count -gt 0) {
            if ($ui.ConfigProfileCard) { $ui.ConfigProfileCard.Visibility = 'Visible' }
            if ($ui.ProfileTotalText) { $ui.ProfileTotalText.Text = "$($profiles.Count)" }
            $pErrors = @($profiles | Where-Object { $_.Status -match 'Error' }).Count
            if ($ui.ProfileErrorText) { $ui.ProfileErrorText.Text = "$pErrors" }
            $pDetail = @($profiles | Where-Object { $_.Status -match 'Error' } | ForEach-Object { "$([char]0x274C) $($_.Name): $($_.Status)" })
            if ($pDetail.Count -eq 0) { $pDetail = @($profiles | Select-Object -First 3 | ForEach-Object { "$([char]0x2705) $($_.Name)" }) }
            if ($ui.ProfileDetailText -and $pDetail.Count -gt 0) { $ui.ProfileDetailText.Text = ($pDetail | Select-Object -First 5) -join "`n" }
            Write-DebugLog "OnComplete: Config profiles - $($profiles.Count) total, $pErrors errors" -Level DEBUG
        }

        # N10: Provisioning packages card
        $ppkgs = $mi.MdmDiag.ProvisioningPackages
        if ($ppkgs -and $ppkgs.Count -gt 0) {
            if ($ui.ProvisioningPkgCard) { $ui.ProvisioningPkgCard.Visibility = 'Visible' }
            if ($ui.PpkgTotalText) { $ui.PpkgTotalText.Text = "$($ppkgs.Count)" }
            $ppkgFails = @($ppkgs | Where-Object { $_.TotalFailures -gt 0 }).Count
            if ($ui.PpkgFailText) { $ui.PpkgFailText.Text = "$ppkgFails" }
            $ppkgDetail = @($ppkgs | ForEach-Object {
                $icon = if ($_.TotalFailures -gt 0) { [char]0x274C } else { [char]0x2705 }
                $desc = if ($_.PackageName) { $_.PackageName } elseif ($_.FriendlyName) { $_.FriendlyName -replace ' \(.*', '' } else { $_.FileName }
                $owner = if ($_.Owner) { " [$($_.Owner)]" } else { '' }
                "$icon $desc$owner"
            })
            if ($ui.PpkgDetailText -and $ppkgDetail.Count -gt 0) { $ui.PpkgDetailText.Text = ($ppkgDetail | Select-Object -First 5) -join "`n" }
            Write-DebugLog "OnComplete: Provisioning packages - $($ppkgs.Count) total, $ppkgFails with failures" -Level DEBUG
        }

        # N9: Enrollment issues card
        $eIssues = $mi.MdmDiag.EnrollmentIssues
        if ($eIssues -and $eIssues.Count -gt 0) {
            if ($ui.EnrollmentIssuesCard) { $ui.EnrollmentIssuesCard.Visibility = 'Visible' }
            $issueLines = @($eIssues | ForEach-Object {
                $icon = if ($_.Severity -eq 'Error') { [char]0x274C } else { [char]0x26A0 }
                $code = if ($_.ErrorCode) { " [$($_.ErrorCode)]" } else { '' }
                "$icon $($_.Issue)$code"
            })
            if ($ui.EnrollmentIssuesText) { $ui.EnrollmentIssuesText.Text = $issueLines -join "`n" }
            Write-DebugLog "OnComplete: Enrollment issues - $($eIssues.Count) problems" -Level DEBUG
        }
    }

    # Show data state - hide Top Issues entirely when nothing to show
    $ui.DashEmptyState.Visibility = 'Collapsed'
    $Script:TopConflicts.Clear()
    if (($conflictCount + $redundantCount) -gt 0) {
        $ui.DashDataState.Visibility = 'Visible'
        $top = @($conflicts | Where-Object { $_.PSObject.Properties.Name -contains 'Severity' } | Select-Object -First 10)
        foreach ($c in $top) { [void]$Script:TopConflicts.Add($c) }
    } else {
        $ui.DashDataState.Visibility = 'Collapsed'
    }

    # Subtitles
    $ui.GPOListSubtitle.Text   = "$($gpos.Count) GPOs in $($Script:ScanData.Domain)"
    $ui.SettingsSubtitle.Text  = "$totalSettings settings across $($gpos.Count) GPOs $([char]0xB7) $($Script:ScanData.Domain)"
    $ui.ConflictsSubtitle.Text = "$conflictCount conflicts, $redundantCount redundant $([char]0xB7) $($Script:ScanData.Domain)"
    $ui.IntuneAppsSubtitle.Text = "$($Script:AllIntuneApps.Count) managed apps $([char]0xB7) $($Script:ScanData.Domain)"

    # Conflict summary sidebar
    $summaryParts = @()
    if ($conflictCount -gt 0)  { $summaryParts += "$conflictCount conflicting settings (different values in multiple GPOs)" }
    if ($redundantCount -gt 0) { $summaryParts += "$redundantCount redundant settings (same value duplicated)" }
    if ($unlinkedCount -gt 0)  { $summaryParts += "$unlinkedCount unlinked GPOs (cleanup candidates)" }
    if ($summaryParts.Count -eq 0) { $summaryParts += 'No conflicts or redundancies detected!' }
    $ui.ConflictSummaryText.Text = $summaryParts -join "`n`n"

    # Enable exports
    $ui.BtnExportHtml.IsEnabled         = $true
    $ui.BtnExportCsv.IsEnabled          = $true
    $ui.BtnExportConflictsCsv.IsEnabled = ($conflicts.Count -gt 0)
    $ui.BtnSaveSnapshot.IsEnabled       = $true
    if ($ui.BtnExportScript)     { $ui.BtnExportScript.IsEnabled = $true }
    if ($ui.BtnPrintReport)      { $ui.BtnPrintReport.IsEnabled = $true }
    if ($ui.BtnExportReg)        { $ui.BtnExportReg.IsEnabled = $true }
    if ($ui.BtnCompareSnapshot)  { $ui.BtnCompareSnapshot.IsEnabled = $true }
    Update-DashboardCharts
    Update-BaselineCompliance
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 13: SCAN BUTTON HANDLER
# ═══════════════════════════════════════════════════════════════════════════════

$ui.BtnScanGPOs.Add_Click({
    if (-not $Script:PrereqsMet) {
        # Lightweight inline check — avoid blocking on Get-GPO or domain queries
        $mode = $Script:Prefs.ScanMode
        $quickFail = $false
        if ($mode -eq 'AD' -and -not (Get-Module -ListAvailable -Name GroupPolicy -ErrorAction SilentlyContinue)) {
            Show-Toast 'RSAT Missing' 'GroupPolicy module not found. Check App Settings.' 'error'
            $quickFail = $true
        } elseif ($mode -eq 'Local' -and -not (Get-Command gpresult.exe -ErrorAction SilentlyContinue)) {
            Show-Toast 'gpresult Missing' 'gpresult.exe not found.' 'error'
            $quickFail = $true
        }
        if ($quickFail) { return }
    }
    # Multi-domain: if TxtDomainOverride has content, use it as the domain override
    if ($ui.TxtDomainOverride -and $ui.TxtDomainOverride.Text.Trim()) {
        $Script:Prefs.DomainOverride = $ui.TxtDomainOverride.Text.Trim().Split(",")[0].Trim()
        Write-DebugLog "Domain override from UI: $($Script:Prefs.DomainOverride)" -Level INFO
    }

    Set-Status 'Scanning...' '#F59E0B'
    $ui.BtnScanGPOs.IsEnabled = $false
    if ($ui.ScanProgressPanel) { $ui.ScanProgressPanel.Visibility = 'Visible'; if ($ui.ScanProgressBar) { $ui.ScanProgressBar.Value = 0 }; if ($ui.ScanProgressText) { $ui.ScanProgressText.Text = '' } }
    if ($ui.pnlGlobalProgress) { $ui.pnlGlobalProgress.Visibility = 'Visible'; if ($ui.lblGlobalProgress) { $ui.lblGlobalProgress.Text = 'Scanning...' } }

    $scanMode   = $Script:Prefs.ScanMode
    $domainOvr  = $Script:Prefs.DomainOverride
    $dcOvr      = $Script:Prefs.DCOverride
    $ouScope    = if ($ui.TxtOUScope -and $ui.TxtOUScope.Text.Trim()) { $ui.TxtOUScope.Text.Trim() } else { '' }
    $forceRefresh = if ($ui.ChkForceRefresh) { $ui.ChkForceRefresh.IsChecked -eq $true } else { $false }

    Write-DebugLog "Starting $scanMode scan..." -Level STEP
    Write-DebugLog "Parameters: domain=$domainOvr, dc=$dcOvr, ou=$ouScope, force=$forceRefresh" -Level DEBUG

    $Script:ScanStartTime = [DateTime]::Now
    # â”€â”€ Heavy work runs in background runspace â”€â”€
    Start-BackgroundWork -Variables @{
        ScanMode   = $scanMode
        DomainOvr  = $domainOvr
        DcOvr      = $dcOvr
        OUScope      = $ouScope
        ForceRefresh = $forceRefresh
        SyncH      = $Global:SyncHash
        ScriptRoot = $PSScriptRoot
    } -Context @{ Name = "$scanMode Scan" } -Work {
        try {
        # -- Background thread: no $ui, no $Script: access --
        $gpoRecords     = [System.Collections.Generic.List[PSCustomObject]]::new()
        $allSettingsList = [System.Collections.Generic.List[PSCustomObject]]::new()
        $appsList          = [System.Collections.Generic.List[PSCustomObject]]::new()
        $imeAdded       = 0
        $gpoIdCounter   = 0
        $settIdCounter  = 0
        $domain         = 'LocalMachine'
        $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[$ScanMode] Work scriptblock entered - mode=$ScanMode domain=$DomainOvr dc=$DcOvr ou=$OUScope force=$ForceRefresh";Level='DEBUG'})

        # ── Intune Category → Group mapping ──
        function Get-IntuneGroup ($cat, $source) {
            if ($cat -like 'Endpoint Security:*') { return 'Endpoint Security' }
            if ($cat -like 'Device Security:*')   { return 'Account Protection' }
            if ($cat -like 'Compliance:*')         { return 'Device Compliance' }
            if ($cat -eq 'Windows Update')         { return 'Windows Update' }
            if ($cat -eq 'App Management')         { return 'App Management' }
            if ($cat -eq 'Registry Settings' -or $cat -like '*(Registry)*') { return 'Registry Settings' }
            if ($cat -in @('Microsoft Edge','Power Management','Device Restrictions') -or
               $cat -like 'System:*' -or $cat -like 'Connectivity:*' -or $cat -like 'Microsoft:*') {
                return 'Configuration Profiles'
            }
            if ($source -eq 'Intune') { return 'ADMX Profiles' }
            return 'Group Policy'
        }

        # Numbered-list formatter (mirrors script-level copy - see SECTION 3a)
        function Format-NumberedList([string]$raw) {
            if (-not $raw -or $raw.Length -lt 4) { return $raw }
            $raw = $raw -replace '\uF000',''
            if ($raw[0] -ne '1' -or ($raw.Length -gt 1 -and [char]::IsDigit($raw[1]))) { return $raw }
            $items = [System.Collections.Generic.List[string]]::new()
            $remaining = $raw.Substring(1)
            $expected  = 2
            $lengthSum = 0
            $pfxRe     = $null
            while ($remaining.Length -gt 0) {
                $idxStr  = "$expected"
                $pattern = "$([regex]::Escape($idxStr))(?=[a-zA-Z]{2}|[\*\\;]|https?://)"
                $allM    = [regex]::Matches($remaining, $pattern)
                $best = $null
                if ($allM.Count -eq 1 -and $allM[0].Index -ge 1) {
                    $best = $allM[0]
                } elseif ($allM.Count -gt 1) {
                    if ($pfxRe) {
                        foreach ($cm in $allM) {
                            if ($cm.Index -lt 1) { continue }
                            $after = $remaining.Substring($cm.Index + $cm.Length)
                            if ($after -match $pfxRe) { $best = $cm; break }
                        }
                    }
                    if (-not $best -and $items.Count -ge 2) {
                        $avg = $lengthSum / $items.Count
                        $bestDist = [int]::MaxValue
                        foreach ($cm in $allM) {
                            if ($cm.Index -lt 1) { continue }
                            $dist = [Math]::Abs($cm.Index - $avg)
                            if ($dist -lt $bestDist) { $bestDist = $dist; $best = $cm }
                        }
                    }
                    if (-not $best) {
                        foreach ($cm in $allM) { if ($cm.Index -ge 1) { $best = $cm; break } }
                    }
                }
                if ($best) {
                    $item = $remaining.Substring(0, $best.Index)
                    $items.Add($item)
                    $lengthSum += $item.Length
                    $remaining  = $remaining.Substring($best.Index + $best.Length)
                    $expected++
                    if ($items.Count -eq 1) {
                        if     ($item -match '^https?://')            { $pfxRe = '^https?://' }
                        elseif ($item -match '^[a-z]{20,}')           { $pfxRe = '^[a-z]{10,}' }
                        elseif ($item -match '^[a-z][\w.-]*\.[a-z]') { $pfxRe = '^[a-z][\w.-]*\.' }
                    }
                } else {
                    $items.Add($remaining)
                    break
                }
            }
            if ($items.Count -ge 2) {
                $result = for ($i = 0; $i -lt $items.Count; $i++) { "$($i + 1). $($items[$i])" }
                return ($result -join "`n")
            }
            if ($items.Count -eq 1) { return $items[0] }
            return $raw
        }

        $runLocal  = $ScanMode -in @('Local','Combined')
        $runIntune = $ScanMode -in @('Intune','Combined')
        $runAD     = $ScanMode -eq 'AD'

        if ($runLocal) {
            # â”€â”€ LOCAL RSoP via gpresult â”€â”€
            $SyncH.StatusQueue.Enqueue(@{Type='Log';Text='[Local] Running gpresult RSoP scan...';Level='STEP'})
            try {
                $SyncH.StatusQueue.Enqueue(@{Type='Log';Text='[Local] Checking Win32_ComputerSystem for domain membership...';Level='DEBUG'})
                $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
                if ($cs.PartOfDomain) { $domain = $cs.Domain }
            } catch { }

            $tmpFile = [IO.Path]::Combine([IO.Path]::GetTempPath(), 'PolicyPilot_RSoP.xml')
            $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Local] RSoP XML path: $tmpFile";Level='DEBUG'})
            # Clean up any stale file from previous run
            if (Test-Path $tmpFile) { Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue }
            try {
                $errFile = [IO.Path]::Combine([IO.Path]::GetTempPath(), 'gpresult_err.txt')
                $outFile = [IO.Path]::Combine([IO.Path]::GetTempPath(), 'gpresult_out.txt')
                $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
                $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Local] Admin: $isAdmin, TEMP: $([IO.Path]::GetTempPath())";Level='DEBUG'})
                # Note: /scope computer requires admin; captures computer policy only (user RSoP may be empty for local admin accounts)
                $gpArgs  = @('/scope', 'computer', '/x', $tmpFile, '/f')
                $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Local] gpresult args: $($gpArgs -join ' ')";Level='DEBUG'})
                $SyncH.StatusQueue.Enqueue(@{Type='Log';Text='[Local] Launching gpresult.exe (RSoP generation may take 30-60s)...';Level='INFO'})
                $gpError = $null
                $proc = try {
                    Start-Process -FilePath 'gpresult.exe' -ArgumentList $gpArgs `
                        -WindowStyle Hidden -PassThru `
                        -RedirectStandardError $errFile -RedirectStandardOutput $outFile
                } catch { $gpError = $_.Exception.Message; $null }

                if (-not $proc) {
                    $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Local] Start-Process returned null. Error: $gpError";Level='WARN'})
                    $SyncH.StatusQueue.Enqueue(@{Type='Log';Text='[Local] Retrying without -RedirectStandardError...';Level='WARN'})
                    $proc = try {
                        Start-Process -FilePath 'gpresult.exe' -ArgumentList $gpArgs `
                            -WindowStyle Hidden -PassThru
                    } catch { $gpError = $_.Exception.Message; $null }
                    if (-not $proc) {
                        $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Local] Start-Process still null. Error: $gpError";Level='ERROR'})
                    }
                }

                # Poll with timeout (do NOT use -Wait — it blocks the runspace and prevents timeout)
                $timeout = 120; $elapsed = 0; $interval = 2
                if ($proc) {
                    while (-not $proc.HasExited -and $elapsed -lt $timeout) {
                        Start-Sleep -Seconds $interval; $elapsed += $interval
                        if ($elapsed % 10 -eq 0) {
                            $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Local] gpresult running... (${elapsed}s)";Level='DEBUG'})
                        }
                    }
                    if (-not $proc.HasExited) {
                        try { $proc.Kill() } catch { }
                        return @{ Error = 'gpresult timed out after 120s - is the domain controller reachable?' }
                    }
                }

                # Log process result details
                try {
                    $exitCode = if ($proc) { $proc.ExitCode } else { 'NULL-PROC' }
                    $fileExists = Test-Path $tmpFile
                    $fileSize = if ($fileExists) { (Get-Item $tmpFile).Length } else { 0 }
                    $stderrContent = if ($errFile -and (Test-Path $errFile)) { (Get-Content $errFile -Raw -ErrorAction SilentlyContinue) } else { '' }
                    $stdoutContent = if ($outFile -and (Test-Path $outFile)) { (Get-Content $outFile -Raw -ErrorAction SilentlyContinue) } else { '' }
                    $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Local] gpresult exit=$exitCode, file=$fileExists (${fileSize}b), stderr='$(if($stderrContent){$stderrContent.Trim()})', stdout='$(if($stdoutContent){$stdoutContent.Trim()})'";Level='DEBUG'})
                } catch {
                    $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Local] DIAGNOSTIC ERROR: $($_.Exception.Message) at line $($_.InvocationInfo.ScriptLineNumber)";Level='ERROR'})
                }

                # Check result
                $timeout = 120

                $gpFailed = $true
                if ($proc) {
                    $gpFailed = ($null -ne $proc.ExitCode -and $proc.ExitCode -ne 0) -or (-not (Test-Path $tmpFile))
                }
                if ($gpFailed) {
                    $exitInfo = if ($proc -and $null -ne $proc.ExitCode) { $proc.ExitCode } else { 'N/A' }
                    $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Local] Full RSoP failed (exit=$exitInfo), trying user-scope only...";Level='WARN'})
                    $gpArgs2 = @('/scope', 'user', '/x', $tmpFile, '/f')
                    $proc2 = try {
                        Start-Process -FilePath 'gpresult.exe' -ArgumentList $gpArgs2 `
                            -WindowStyle Hidden -PassThru
                    } catch { $gpError = $_.Exception.Message; $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Local] User-scope Start-Process error: $($_.Exception.Message)";Level='ERROR'}); $null }
                    if ($proc2) {
                        $elapsed2 = 0
                        while (-not $proc2.HasExited -and $elapsed2 -lt $timeout) {
                            Start-Sleep -Seconds 2; $elapsed2 += 2
                        }
                        if (-not $proc2.HasExited) { try { $proc2.Kill() } catch { } }
                    }
                    $proc = $proc2
                }

                # Final check after all attempts
                $finalFailed = (-not $proc) -or (-not (Test-Path $tmpFile))
                if (-not $finalFailed -and $proc -and $null -ne $proc.ExitCode -and $proc.ExitCode -ne 0) { $finalFailed = $true }
                if ($finalFailed) {
                    $exitCode = if ($proc -and $null -ne $proc.ExitCode) { $proc.ExitCode } else { 'N/A' }
                    $errText = ''
                    if (Test-Path $errFile) { $raw = Get-Content $errFile -Raw -ErrorAction SilentlyContinue; if ($raw) { $errText = $raw.Trim() } }
                    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
                    $hint = if (-not $isAdmin) { ' (not running as admin - computer policies require elevation)' } else { '' }
                    return @{ Error = "gpresult failed (exit $exitCode)$hint$(if ($errText) { ": $errText" })" }
                }

                $SyncH.StatusQueue.Enqueue(@{Type='Log';Text='[Local] gpresult completed, parsing RSoP XML...';Level='INFO'})
                [xml]$rsop = Get-Content -Path $tmpFile -Raw -Encoding UTF8
                $rootNs = $rsop.DocumentElement.NamespaceURI
                $nsMgr  = New-Object System.Xml.XmlNamespaceManager($rsop.NameTable)
                $nsMgr.AddNamespace('r', $rootNs)

                foreach ($scopeTag in @('ComputerResults','UserResults')) {
                    $scope = if ($scopeTag -eq 'ComputerResults') { 'Computer' } else { 'User' }
                    $scopeNode = $rsop.SelectSingleNode("//r:$scopeTag", $nsMgr)
                    if (-not $scopeNode) { continue }

                    $gpoNodes = $scopeNode.SelectNodes('r:GPO', $nsMgr)
                    foreach ($gNode in $gpoNodes) {
                        $gpoIdCounter++
                        $gName   = $gNode.SelectSingleNode('r:Name', $nsMgr)
                        $gPath   = $gNode.SelectSingleNode('r:Path', $nsMgr)
                        $gEnable = $gNode.SelectSingleNode('r:Enabled', $nsMgr)
                        $gLink   = $gNode.SelectSingleNode('r:Link/r:SOMPath', $nsMgr)

                        # M1: WMI filter status
                        $gFilterAllowed = $gNode.SelectSingleNode('r:FilterAllowed', $nsMgr)
                        $wmiFilterStatus = if ($gFilterAllowed) { if ($gFilterAllowed.InnerText -eq 'true') { 'Passed' } else { 'Denied' } } else { '' }
                        # H2: Enforced + link order
                        $gEnforced = $gNode.SelectSingleNode('r:Link/r:NoOverride', $nsMgr)
                        $gLinkOrder = $gNode.SelectSingleNode('r:Link/r:SOMOrder', $nsMgr)

                        $displayName = if ($gName) { $gName.InnerText } else { "GPO-$gpoIdCounter" }
                        $pathId      = if ($gPath) { $gPath.InnerText } else { "GPO-$gpoIdCounter" }
                        $isEnabled   = if ($gEnable) { $gEnable.InnerText -eq 'true' } else { $true }
                        $linkPath    = if ($gLink) { $gLink.InnerText } else { '' }

                        [void]$gpoRecords.Add([PSCustomObject]@{
                            Id              = $gpoIdCounter
                            DisplayName     = $displayName
                            GpoId           = $pathId
                            Status          = if ($isEnabled) { 'Enabled' } else { 'Disabled' }
                            CreatedTime     = ''
                            ModifiedTime    = ''
                            WmiFilter       = ''
                            WmiFilterStatus = $wmiFilterStatus
                            LinkPath        = $linkPath
                            IsLinked        = [bool]$linkPath
                            UserVersion     = '0'
                            ComputerVersion = '0'
                            Description     = "Applied via RSoP ($scope scope)"
                            Enforced        = ($gEnforced -and $gEnforced.InnerText -eq 'true')
                            LinkOrder       = if ($gLinkOrder) { [int]$gLinkOrder.InnerText } else { 0 }
                            SecurityFiltering = ''
                        })
                    }

                    # Build path -> friendly GPO name map for setting GPOName resolution
                    $pathToGpoName = @{}
                    foreach ($g in $gpoRecords) { if ($g.GpoId -and $g.DisplayName) { $pathToGpoName[$g.GpoId] = $g.DisplayName } }

                    $extDataNodes = $scopeNode.SelectNodes('r:ExtensionData', $nsMgr)
                    foreach ($extData in $extDataNodes) {
                        $extNameNode = $extData.SelectSingleNode('*[local-name()="Name"]')
                        $category    = if ($extNameNode) { $extNameNode.InnerText } else { 'General' }
                        $extNode = $extData.SelectSingleNode('*[local-name()="Extension"]')
                        if (-not $extNode) { continue }

                        $policies = $extNode.SelectNodes('*[local-name()="Policy"]')
                        foreach ($pol in $policies) {
                            $settIdCounter++
                            $polName  = $pol.SelectSingleNode('*[local-name()="Name"]')
                            $polState = $pol.SelectSingleNode('*[local-name()="State"]')
                            $polGPO   = $pol.SelectSingleNode('*[local-name()="GPO"]')
                            $polCat   = $pol.SelectSingleNode('*[local-name()="Category"]')
                            $polValue = $pol.SelectSingleNode('*[local-name()="Value"]')
                            $sName   = if ($polName)  { $polName.InnerText }  else { "Policy-$settIdCounter" }
                            $sState  = if ($polState) { $polState.InnerText } else { 'Applied' }
                            $sGPORaw = if ($polGPO)   { $polGPO.InnerText }   else { '' }
                            $sGPO    = if ($pathToGpoName[$sGPORaw]) { $pathToGpoName[$sGPORaw] } else { $sGPORaw }
                            $sCat    = if ($polCat)    { $polCat.InnerText }   else { $category }
                            $sValue  = if ($polValue)  { $polValue.InnerText } else { $sState }
                            $sValue  = Format-NumberedList $sValue
                            [void]$allSettingsList.Add([PSCustomObject]@{
                                Id=($settIdCounter); GPOName=$sGPO; GPOGuid=$sGPORaw; Category=$sCat
                                PolicyName=$sName; SettingKey="$sCat\$sName"; State=$sState
                                RegistryKey=''; ValueData=$sValue; Scope=$scope; Source='Local GPO'; IntuneGroup=(Get-IntuneGroup $sCat 'Local GPO')
                            })
                        }

                        $regSettings = $extNode.SelectNodes('*[local-name()="RegistrySetting"]')
                        foreach ($reg in $regSettings) {
                            $settIdCounter++
                            $regGPO  = $reg.SelectSingleNode('*[local-name()="GPO"]')
                            $regKey  = $reg.SelectSingleNode('*[local-name()="KeyPath"]')
                            $regVal  = $reg.SelectSingleNode('*[local-name()="ValueName"]')
                            $regAdm  = $reg.SelectSingleNode('*[local-name()="AdmSetting"]')
                            $sGPORaw = if ($regGPO) { $regGPO.InnerText } else { '' }
                            $sGPO    = if ($pathToGpoName[$sGPORaw]) { $pathToGpoName[$sGPORaw] } else { $sGPORaw }
                            $keyPath = if ($regKey) { $regKey.InnerText } else { '' }
                            $valName = if ($regVal) { $regVal.InnerText } else { '' }
                            $admSet  = if ($regAdm) { $regAdm.InnerText } else { '' }
                            $fullPath = if ($valName) { "$keyPath\$valName" } else { $keyPath }
                            $sName    = if ($valName) { $valName } else { $keyPath.Split('\')[-1] }
                            [void]$allSettingsList.Add([PSCustomObject]@{
                                Id=($settIdCounter); GPOName=$sGPO; GPOGuid=$sGPORaw; Category="$category (Registry)"
                                PolicyName=$sName; SettingKey=$fullPath; State='Applied'
                                RegistryKey=$fullPath; ValueData="AdmSetting=$admSet"; Scope=$scope; Source='Local GPO'; IntuneGroup=(Get-IntuneGroup "$category (Registry)" 'Local GPO')
                            })
                        }

                        foreach ($tagName in @('Account','SecurityOptions','Audit')) {
                            $items = $extNode.SelectNodes("*[local-name()='$tagName']")
                            foreach ($item in $items) {
                                $settIdCounter++
                                $iName  = $item.SelectSingleNode('*[local-name()="Name"]')
                                $iGPO   = $item.SelectSingleNode('*[local-name()="GPO"]')
                                $iValue = $item.SelectSingleNode('*[local-name()="SettingNumber"] | *[local-name()="SettingBoolean"] | *[local-name()="SettingString"] | *[local-name()="Value"]')
                                $sName  = if ($iName)  { $iName.InnerText }  else { "$tagName-$settIdCounter" }
                                $sGPORaw = if ($iGPO)   { $iGPO.InnerText }   else { '' }
                                $sGPO   = if ($pathToGpoName[$sGPORaw]) { $pathToGpoName[$sGPORaw] } else { $sGPORaw }
                                $sValue = if ($iValue)  { $iValue.InnerText } else { '' }
                                [void]$allSettingsList.Add([PSCustomObject]@{
                                    Id=($settIdCounter); GPOName=$sGPO; GPOGuid=$sGPORaw; Category="$category ($tagName)"
                                    PolicyName=$sName; SettingKey="$category\$sName"; State='Applied'
                                    RegistryKey=''; ValueData=$sValue; Scope=$scope; Source='Local GPO'; IntuneGroup=(Get-IntuneGroup "$category ($tagName)" 'Local GPO')
                                })
                            }
                        }

                        # M5: Script elements - extract Type, Command, Parameters
                        $scriptItems = $extNode.SelectNodes("*[local-name()='Script']")
                        foreach ($sItem in $scriptItems) {
                            $settIdCounter++
                            $sName    = $sItem.SelectSingleNode('*[local-name()="Name"]')
                            $sGPO     = $sItem.SelectSingleNode('*[local-name()="GPO"]')
                            $sType    = $sItem.SelectSingleNode('*[local-name()="Type"]')
                            $sCommand = $sItem.SelectSingleNode('*[local-name()="Command"]')
                            $sParams  = $sItem.SelectSingleNode('*[local-name()="Parameters"]')
                            $scriptType = if ($sType) { $sType.InnerText } else { 'Unknown' }
                            $scriptCmd  = if ($sCommand) { $sCommand.InnerText } else { '' }
                            $scriptPrm  = if ($sParams) { $sParams.InnerText } else { '' }
                            $scriptName = if ($sName) { $sName.InnerText } else { "$scriptType Script $settIdCounter" }
                            $scriptValue = $scriptCmd
                            if ($scriptPrm) { $scriptValue += " $scriptPrm" }
                            [void]$allSettingsList.Add([PSCustomObject]@{
                                Id=($settIdCounter); GPOName=(if ($sGPO) { $sGPORaw = $sGPO.InnerText; if ($pathToGpoName[$sGPORaw]) { $pathToGpoName[$sGPORaw] } else { $sGPORaw } } else { '' }); GPOGuid=(if ($sGPO) { $sGPO.InnerText } else { '' })
                                Category="$category (Script: $scriptType)"; PolicyName=$scriptName
                                SettingKey="$category\$scriptType\$scriptName"; State='Applied'
                                RegistryKey=''; ValueData=$scriptValue; Scope=$scope; Source='Local GPO'
                                IntuneGroup='Scripts'; ScriptType=$scriptType; ScriptPath=$scriptCmd; ScriptParams=$scriptPrm
                            })
                        }
                    }
                }

            # ── WMI RSoP enrichment: replace locale-dependent Policy entries with registry-keyed entries ──
            # RSOP_RegistryPolicySetting provides registryKey+valueName for ADMX policies that
            # gpresult <Policy> elements lack. This enables ADMX database lookup for English names.
            try {
                $rsopWmi = @(Get-CimInstance -Namespace 'root\RSOP\Computer' -ClassName 'RSOP_RegistryPolicySetting' -ErrorAction Stop)
                if ($rsopWmi.Count -gt 0) {
                    $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Local] WMI RSoP: $($rsopWmi.Count) registry policy settings found";Level='INFO'})
                    # Build GUID -> friendly GPO name map
                    $guidToName = @{}
                    foreach ($g in $gpoRecords) {
                        if ($g.GpoId -match '\{([0-9a-fA-F-]+)\}') { $guidToName[$Matches[1].ToUpper()] = $g.DisplayName }
                    }
                    # Remove gpresult entries that WMI RSoP replaces:
                    # 1) <Policy> entries (German, no reg key) — always replaced
                    # 2) <RegistrySetting> entries (have reg key) — WMI covers same data + precedence
                    # Keep: Security/Account/Audit/Script entries (not in WMI RSoP)
                    $replaceableEntries = @($allSettingsList | Where-Object { $_.Category -notmatch 'Security|Account|Audit|Script' })
                    foreach ($pe in $replaceableEntries) { [void]$allSettingsList.Remove($pe) }
                    $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Local] Replaced $($replaceableEntries.Count) gpresult entries with WMI RSoP data";Level='DEBUG'})
                    # Add WMI entries with proper registry keys (includes precedence for conflict detection)
                    foreach ($wmi in $rsopWmi) {
                        # Skip deleted entries (registry cleanup markers)
                        if ($wmi.deleted) { continue }
                        $settIdCounter++
                        $wmiKey = $wmi.registryKey
                        $wmiVal = $wmi.valueName
                        $wmiGpoId = $wmi.GPOID
                        $wmiPrec = if ($wmi.precedence) { [int]$wmi.precedence } else { 1 }
                        # Decode value based on type
                        $wmiValueData = ''
                        if ($null -ne $wmi.value -and $wmi.value.Count -gt 0) {
                            switch ($wmi.valueType) {
                                4 { # REG_DWORD
                                    if ($wmi.value.Count -ge 4) { $wmiValueData = [string][BitConverter]::ToUInt32($wmi.value, 0) }
                                    else { $wmiValueData = [string]$wmi.value }
                                }
                                1 { # REG_SZ
                                    $wmiValueData = [System.Text.Encoding]::Unicode.GetString($wmi.value).TrimEnd("`0")
                                }
                                7 { # REG_MULTI_SZ
                                    $wmiValueData = [System.Text.Encoding]::Unicode.GetString($wmi.value).TrimEnd("`0") -replace "`0", '; '
                                }
                                3 { # REG_BINARY
                                    $wmiValueData = "[Binary $($wmi.value.Count)B]"
                                }
                                default {
                                    if ($wmi.value.Count -gt 0) { $wmiValueData = "[Type$($wmi.valueType) $($wmi.value.Count)B]" }
                                }
                            }
                        }
                        # Resolve GPO GUID to friendly name
                        $gpoGuid = ''
                        $gpoName = ''
                        if ($wmiGpoId -match '\{([0-9a-fA-F-]+)\}') {
                            $gpoGuid = $Matches[1].ToUpper()
                            $gpoName = if ($guidToName[$gpoGuid]) { $guidToName[$gpoGuid] } else { $wmiGpoId }
                        }
                        $fullPath = if ($wmiVal) { "$wmiKey\$wmiVal" } else { $wmiKey }
                        $sName = if ($wmiVal) { $wmiVal } else { $wmiKey.Split('\')[-1] }
                        $wmiState = if ($wmiPrec -eq 1) { 'Applied' } else { 'Superseded' }
                        [void]$allSettingsList.Add([PSCustomObject]@{
                            Id=($settIdCounter); GPOName=$gpoName; GPOGuid=$wmiGpoId
                            Category='Administrative Templates (WMI)'; PolicyName=$sName; SettingKey=$fullPath
                            State=$wmiState; RegistryKey=$fullPath; ValueData=$wmiValueData
                            Scope='Computer'; Source='WMI RSoP'; IntuneGroup='Group Policy'
                            Precedence=$wmiPrec
                        })
                    }
                    $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Local] WMI RSoP: added $($rsopWmi.Count) enrichable settings";Level='SUCCESS'})
                }
            } catch {
                $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Local] WMI RSoP unavailable (non-admin or non-domain): $($_.Exception.Message)";Level='DEBUG'})
                # Fall through — keep gpresult Policy entries as-is (German names, no enrichment)
            }

            $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Local] Scan complete: $($gpoRecords.Count) GPOs, $($allSettingsList.Count) settings";Level='SUCCESS'})
            } catch {
                return @{ Error = $_.Exception.Message }
            } finally {
                if (Test-Path $tmpFile) { Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue }
            }

        }

        if ($runIntune) {
            # ── Intune scan via LOCAL MDM registry (enriched with CSP metadata) ──
            $SyncH.StatusQueue.Enqueue(@{Type='Log';Text='[Intune] Starting local MDM policy scan...';Level='STEP'})

            # ── CSP Metadata Dictionary ─────────────────────────────────────────────
            # Loaded from csp_metadata.json (generated by Build-CspDatabase.ps1)
            # Falls back to inline minimal set if JSON not found
            $cspJsonPath = if ($ScriptRoot) { Join-Path $ScriptRoot 'csp_metadata.json' } else { '' }
            $cspMeta = @{}
            if ($cspJsonPath -and (Test-Path $cspJsonPath)) {
                try {
                    $SyncH.StatusQueue.Enqueue(@{Type='Log';Text='[Intune] Loading CSP metadata from csp_metadata.json...';Level='INFO'})
                    $jsonRaw = [System.IO.File]::ReadAllText($cspJsonPath, [System.Text.Encoding]::UTF8)
                    $jsonObj = $jsonRaw | ConvertFrom-Json
                    foreach ($prop in $jsonObj.PSObject.Properties) {
                        $cspMeta[$prop.Name] = @{
                            Friendly = $prop.Value.Friendly
                            Desc     = $prop.Value.Desc
                            Cat      = $prop.Value.Cat
                            Def      = $prop.Value.Def
                            Scope    = $prop.Value.Scope
                            Editions = $prop.Value.Editions
                            MinVer   = $prop.Value.MinVersion
                            Format   = $prop.Value.Format
                            AV       = if ($prop.Value.AllowedValues) { $prop.Value.AllowedValues } else { $null }
                            GP       = if ($prop.Value.GPMapping) { $prop.Value.GPMapping } else { $null }
                        }
                    }
                    $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] Loaded $($cspMeta.Count) CSP metadata entries";Level='INFO'})
                    # Staleness check: warn if JSON older than 90 days
                    $jsonAge = ((Get-Date) - (Get-Item $cspJsonPath).LastWriteTime).Days
                    if ($jsonAge -gt 90) {
                        $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] CSP metadata is $jsonAge days old. Run Build-CspDatabase.ps1 to refresh.";Level='WARN'})
                    }
                } catch {
                    $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] CSP JSON load error: $($_.Exception.Message)";Level='WARN'})
                }
            }
            if ($cspMeta.Count -eq 0) {
                $SyncH.StatusQueue.Enqueue(@{Type='Log';Text='[Intune] Using built-in minimal CSP metadata';Level='INFO'})
                $cspMeta = @{
                    'Defender/AllowRealtimeMonitoring'=@{Friendly='Allow Real-time Monitoring';Desc='Enable Windows Defender real-time protection';Cat='Endpoint Security: Antivirus';Def='1'}
                    'Defender/CloudBlockLevel'=@{Friendly='Cloud Block Level';Desc='Aggressiveness of cloud blocking';Cat='Endpoint Security: Antivirus';Def='0'}
                    'Defender/EnableNetworkProtection'=@{Friendly='Network Protection';Desc='Block connections to malicious domains';Cat='Endpoint Security: Network Protection';Def='0'}
                    'DeviceLock/MaxInactivityTimeDeviceLock'=@{Friendly='Max Inactivity Lock';Desc='Minutes before device auto-locks';Cat='Device Security: Password';Def='0'}
                    'Update/AllowAutoUpdate'=@{Friendly='Allow Auto Update';Desc='Configure automatic update behavior';Cat='Windows Update';Def='1'}
                    'System/AllowTelemetry'=@{Friendly='Diagnostic Data Level';Desc='Windows telemetry level';Cat='System: Telemetry';Def='3'}
                }
            }

            # ── Friendly CSP area names ──────────────────────────────────────
            $cspAreaNames = @{
                'Defender'='Windows Defender'; 'Update'='Windows Update'; 'DeviceLock'='Device Lock'
                'Browser'='Microsoft Edge'; 'WiFi'='WiFi'; 'Bluetooth'='Bluetooth'; 'Camera'='Camera'
                'ApplicationManagement'='Application Management'; 'System'='System'; 'Privacy'='Privacy'
                'Experience'='User Experience'; 'Power'='Power Management'; 'Firewall'='Windows Firewall'
                'BitLocker'='BitLocker'; 'DeviceGuard'='Device Guard'; 'DeviceHealthMonitoring'='Health Monitoring'
                'WindowsDefenderApplicationGuard'='Application Guard'; 'Accounts'='Accounts'
                'RemoteDesktop'='Remote Desktop'; 'Storage'='Storage'; 'TextInput'='Text Input'
                'Search'='Windows Search'; 'Start'='Start Menu'; 'Notifications'='Notifications'
                'LockDown'='Lockdown'; 'Kerberos'='Kerberos'; 'Maps'='Maps'; 'Messaging'='Messaging'
                'CredentialProviders'='Credential Providers'; 'Authentication'='Authentication'
                'AboveLock'='Above Lock Screen'; 'Cellular'='Cellular'; 'Handwriting'='Handwriting'
                'DeliveryOptimization'='Delivery Optimization'; 'Connectivity'='Connectivity'
                'Cryptography'='Cryptography'; 'DataProtection'='Data Protection'
                'ErrorReporting'='Error Reporting'; 'ExploitGuard'='Exploit Guard'
                'Games'='Games'; 'InternetExplorer'='Internet Explorer'; 'Licensing'='Licensing'
                'LocalPoliciesSecurityOptions'='Local Security Options'; 'MixedReality'='Mixed Reality'
                'NetworkIsolation'='Network Isolation'; 'Printers'='Printers'
                'RemoteAssistance'='Remote Assistance'; 'RemoteManagement'='Remote Management'
                'RestrictedGroups'='Restricted Groups'; 'Security'='Security'; 'Settings'='Settings'
                'SmartScreen'='SmartScreen'; 'Speech'='Speech'; 'TaskScheduler'='Task Scheduler'
                'TimeLanguageSettings'='Time & Language'; 'Troubleshooting'='Troubleshooting'
                'UserRights'='User Rights'; 'VPNv2'='VPN'; 'WindowsInkWorkspace'='Windows Ink'
                'WindowsSandbox'='Windows Sandbox'; 'WirelessDisplay'='Wireless Display'
                'ADMX_MicrosoftEdge'='Edge (ADMX)'; 'ADMX_GroupPolicy'='Group Policy (ADMX)'
            }

            # Helper: strip MDM suffixes to get base setting name
            function Get-BaseSetting([string]$name) {
                $name -replace '_(ProviderSet|WinningProvider|LastWrite|CurrentChannel|ADMXInstanceData)$',''
            }
            function Format-PolicyName([string]$name) {
                # Strip ADMX metadata suffixes (PascalCase, no spaces):
                #   URLBlocklist_URLBlocklistDesc_ListSet → URLBlocklist
                #   PreventInstallationOfMatchingDeviceIDs_DeviceInstall_IDs_Deny_List_ListSet → PreventInstallationOfMatchingDeviceIDs
                #   LetAppsRunInBackground_ForceAllowTheseApps → LetAppsRunInBackground
                $cleaned = $name -replace '_.*Desc(_ListSet)?$','' -replace '_.*_ListSet$','' -replace '_ListSet$',''
                # PascalCase → spaces, then clean up any remaining underscores
                $spaced = ($cleaned -creplace '([a-z])([A-Z])', '$1 $2' -creplace '([A-Z]+)([A-Z][a-z])', '$1 $2').Trim()
                # Fix common abbreviation splits: "I Ds" → "IDs", "U Rl" → "URL", etc.
                $spaced = $spaced -creplace '\bI Ds\b','IDs' -creplace '\bU Rl\b','URL' -creplace '\bU Xs?\b','UX'
                $spaced -replace '_',' - '
            }
            # Helper: resolve ADMX tilde-delimited area names to friendly names
            # e.g. "chromeIntuneV1~Policy~googlechrome~Extensions" → "Google Chrome > Extensions"
            # e.g. "microsoft_edgev140~Policy~microsoft_edge~Network" → "Microsoft Edge > Network"
            function Format-AdmxAreaName([string]$raw) {
                if ($raw -notmatch '~') { return $null }  # not an ADMX area
                $parts = $raw -split '~'
                # ADMX format: <ingestionSource>~Policy~<area>[~<subArea>...]
                # Skip everything up to and including 'Policy'
                $policyIdx = [Array]::IndexOf($parts, 'Policy')
                $meaningful = if ($policyIdx -ge 0 -and $policyIdx -lt $parts.Count - 1) {
                    @($parts[($policyIdx + 1)..($parts.Count - 1)])
                } else {
                    @($parts[-1])
                }
                $friendly = @($meaningful | ForEach-Object {
                    $p = $_
                    switch -Regex ($p) {
                        '^googlechrome$'  { 'Google Chrome'; break }
                        '^microsoft_edge$' { 'Microsoft Edge'; break }
                        '^microsoft_'     { ($p -replace '^microsoft_','Microsoft ') -creplace '([a-z])([A-Z])','$1 $2'; break }
                        default {
                            $p = ($p -creplace '([a-z])([A-Z])', '$1 $2').Trim()
                            if ($p.Length -gt 0) { $p.Substring(0,1).ToUpper() + $p.Substring(1) } else { $p }
                        }
                    }
                })
                # Deduplicate consecutive identical segments
                $deduped = [System.Collections.Generic.List[string]]::new()
                foreach ($f in $friendly) {
                    if ($deduped.Count -eq 0 -or $deduped[$deduped.Count - 1] -ne $f) { [void]$deduped.Add($f) }
                }
                return ($deduped -join ' > ')
            }

            # Check MDM enrollment & collect rich metadata
            $enrolled = $false
            $mdmInfo = @{
                EnrollmentUPN = ''; ProviderID = ''; AADTenantID = ''; EnrollmentType = ''
                EnrollmentState = ''; EnrollmentGUID = ''; EntDMID = ''; CertRenewTime = ''
                MdmDiag = @{ DeviceInfo = @{}; ManagedPolicies = 0; Certificates = @(); LAPS = @{}; PolicyMeta = @{}; AppSummary = @{}; Compliance = @{}; ScriptPolicies = @(); ConfigProfiles = @(); EnrollmentIssues = @(); ProvisioningPackages = @() }
            }
            $enrollPath = 'HKLM:\SOFTWARE\Microsoft\Enrollments'
            if (Test-Path $enrollPath) {
                $enrollKeys = Get-ChildItem $enrollPath -ErrorAction SilentlyContinue |
                    Where-Object { (Get-ItemProperty $_.PSPath -Name 'ProviderID' -ErrorAction SilentlyContinue).ProviderID }
                if ($enrollKeys) {
                    $enrolled = $true
                    # Pick the primary MDM enrollment (MS DM Server = Intune)
                    $primaryKey = $enrollKeys | Where-Object {
                        (Get-ItemProperty $_.PSPath -Name 'ProviderID' -EA SilentlyContinue).ProviderID -eq 'MS DM Server'
                    } | Select-Object -First 1
                    if (-not $primaryKey) { $primaryKey = $enrollKeys[0] }
                    $ep = Get-ItemProperty $primaryKey.PSPath -EA SilentlyContinue
                    $mdmInfo.ProviderID      = "$($ep.ProviderID)"
                    $mdmInfo.EnrollmentUPN   = "$($ep.UPN)" -replace '@[a-f0-9-]+$',''
                    $mdmInfo.AADTenantID     = "$($ep.AADTenantID)"
                    $mdmInfo.EnrollmentGUID  = $primaryKey.PSChildName
                    $mdmInfo.EnrollmentState = switch ("$($ep.EnrollmentState)") { '1' { 'Enrolled' } '0' { 'Not Enrolled' } default { "State $($ep.EnrollmentState)" } }
                    $enrollTypeMap = @{ '6'='MDM'; '12'='WMI Bridge (SCCM)'; '26'='ConfigMgr Co-mgmt'; '28'='Local Authority'; '30'='Deploy Authority'; '31'='Cloud Authority' }
                    $mdmInfo.EnrollmentType  = if ($enrollTypeMap["$($ep.EnrollmentType)"]) { $enrollTypeMap["$($ep.EnrollmentType)"] } else { "Type $($ep.EnrollmentType)" }
                    # Read DMClient subkey for EntDMID and cert renewal
                    $dmPath = Join-Path $primaryKey.PSPath "DMClient\$($ep.ProviderID)"
                    $dm = Get-ItemProperty $dmPath -EA SilentlyContinue
                    if ($dm) {
                        $mdmInfo.EntDMID = "$($dm.EntDMID)"
                        $mdmInfo.CertRenewTime = "$($dm.CertRenewTimeStamp)"
                    }
                    $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] Enrolled via: $($mdmInfo.ProviderID)  UPN: $($mdmInfo.EnrollmentUPN)  Type: $($mdmInfo.EnrollmentType)";Level='INFO'})

                    # ── N9: Enrollment troubleshooting ──
                    $enrollIssues = [System.Collections.Generic.List[hashtable]]::new()
                    # Check all enrollment keys for problems
                    foreach ($ek in $enrollKeys) {
                        $ekProps = Get-ItemProperty $ek.PSPath -EA SilentlyContinue
                        if (-not $ekProps) { continue }
                        $ekGuid = $ek.PSChildName
                        $ekState = "$($ekProps.EnrollmentState)"
                        # Failed states: 0=Not Enrolled, 2=Pending, 3=Partial
                        if ($ekState -in @('0','2','3')) {
                            $stateText = switch ($ekState) { '0' { 'Not Enrolled' } '2' { 'Enrollment Pending' } '3' { 'Partial Enrollment' } default { "State $ekState" } }
                            $enrollIssues.Add(@{ Guid = $ekGuid; Provider = "$($ekProps.ProviderID)"; Issue = $stateText; Severity = 'Warning' })
                        }
                        # Check for error codes in the Status subkey
                        $statusPath = Join-Path $ek.PSPath 'Status'
                        if (Test-Path $statusPath) {
                            $statusProps = Get-ItemProperty $statusPath -EA SilentlyContinue
                            if ($statusProps.ErrorCode -and $statusProps.ErrorCode -ne 0) {
                                $errorHex = '0x{0:X}' -f $statusProps.ErrorCode
                                $enrollErrorMap = @{
                                    '0x80180001' = 'Device already enrolled'
                                    '0x80180002' = 'Device not found in directory'
                                    '0x80180003' = 'Enrollment profile not found'
                                    '0x80180005' = 'Server rejected the request'
                                    '0x80180006' = 'Enrollment blocked by server policy'
                                    '0x80180014' = 'Enrollment limit reached (too many devices)'
                                    '0x80180026' = 'Certificate request timeout'
                                    '0x801c0003' = 'Azure AD join failed - device limit exceeded'
                                    '0x801c001D' = 'Hybrid Azure AD join - connector issue'
                                    '0x80070774' = 'RPC server unavailable (DC connectivity)'
                                    '0x80072EE7' = 'DNS/network resolution failure'
                                    '0x80072F8F' = 'SSL/TLS certificate trust failure'
                                }
                                $friendlyError = if ($enrollErrorMap[$errorHex]) { $enrollErrorMap[$errorHex] } else { "Error code $errorHex" }
                                $enrollIssues.Add(@{ Guid = $ekGuid; Provider = "$($ekProps.ProviderID)"; Issue = $friendlyError; Severity = 'Error'; ErrorCode = $errorHex })
                            }
                        }
                        # Check DMClient last-sync errors
                        $dmcPath = Join-Path $ek.PSPath "DMClient\$($ekProps.ProviderID)"
                        if (Test-Path $dmcPath) {
                            $dmcProps = Get-ItemProperty $dmcPath -EA SilentlyContinue
                            if ($dmcProps.LastError -and $dmcProps.LastError -ne 0) {
                                $lastErrHex = '0x{0:X}' -f $dmcProps.LastError
                                $enrollIssues.Add(@{ Guid = $ekGuid; Provider = "$($ekProps.ProviderID)"; Issue = "Last sync error: $lastErrHex"; Severity = 'Warning'; ErrorCode = $lastErrHex })
                            }
                        }
                    }
                    if ($enrollIssues.Count -gt 0) {
                        $mdmInfo.MdmDiag.EnrollmentIssues = @($enrollIssues)
                        $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] Enrollment issues: $($enrollIssues.Count) problems detected";Level='WARN'})
                    }
                }
            }
            if (-not $enrolled) {
                $SyncH.StatusQueue.Enqueue(@{Type='Log';Text='[Intune] WARNING: No MDM enrollment detected on this device';Level='WARN'})
            }

            # Run mdmdiagnosticstool.exe to capture the diagnostic report (only if enrolled)
            $mdmDiagDir = $null
            if ($enrolled) { try {
                $mdmDiagDir = Join-Path ([IO.Path]::GetTempPath()) "PolicyPilot_MdmDiag_$(Get-Random)"
                [void][IO.Directory]::CreateDirectory($mdmDiagDir)
                $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] Running mdmdiagnosticstool.exe -> $mdmDiagDir";Level='INFO'})
                $mdmProc = Start-Process -FilePath 'mdmdiagnosticstool.exe' -ArgumentList '-out', "`"$mdmDiagDir`"" -NoNewWindow -Wait -PassThru -ErrorAction Stop
                $mdmHtml = Join-Path $mdmDiagDir 'MDMDiagReport.html'
                if ((Test-Path $mdmHtml) -and $mdmProc.ExitCode -eq 0) {
                    $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] MDM diagnostic report captured ($([math]::Round((Get-Item $mdmHtml).Length/1KB))KB)";Level='INFO'})
                    $mdmContent = [IO.File]::ReadAllText($mdmHtml)
                    # Parse Device Info section
                    $diIdx = $mdmContent.IndexOf('Device Info')
                    if ($diIdx -gt 0) {
                        $diEnd = $mdmContent.IndexOf('</section>', $diIdx)
                        $diSect = $mdmContent.Substring($diIdx, $diEnd - $diIdx)
                        $diRx = [regex]'(?s)LabelColumn[^>]*>([^<]+)</td>\s*<td[^>]*>([^<]*)</td>'
                        foreach ($m in $diRx.Matches($diSect)) {
                            $mdmInfo.MdmDiag.DeviceInfo[$m.Groups[1].Value.Trim()] = $m.Groups[2].Value.Trim()
                        }
                    }
                    # Parse Certificates section
                    $certIdx = $mdmContent.IndexOf('Certificates')
                    if ($certIdx -gt 0) {
                        $certEnd = $mdmContent.IndexOf('</section>', $certIdx)
                        $certSect = $mdmContent.Substring($certIdx, $certEnd - $certIdx)
                        $certRx = [regex]'(?s)<tr>\s*<td[^>]*>([^<]*)</td>\s*<td[^>]*>([^<]*)</td>\s*<td[^>]*>([^<]*)</td>\s*<td[^>]*>([^<]*)</td>'
                        foreach ($cm in $certRx.Matches($certSect)) {
                            $mdmInfo.MdmDiag.Certificates += @{ IssuedTo=$cm.Groups[1].Value.Trim(); IssuedBy=$cm.Groups[2].Value.Trim(); Expiration=$cm.Groups[3].Value.Trim(); Purpose=$cm.Groups[4].Value.Trim() }
                        }
                    }
                    # Count managed policies
                    $mpIdx = $mdmContent.IndexOf('Managed policies')
                    if ($mpIdx -gt 0) {
                        $mpEnd = $mdmContent.IndexOf('</section>', $mpIdx)
                        $mpSect = $mdmContent.Substring($mpIdx, $mpEnd - $mpIdx)
                        $mdmInfo.MdmDiag.ManagedPolicies = ([regex]'<tr>').Matches($mpSect).Count - 1
                    }
                    # Store raw path for optional reference
                    $mdmInfo.MdmDiag.ReportPath = $mdmHtml

                    # ── Known provisioning package friendly names (fallback when cmdlet unavailable) ──
                    $ppkgFriendlyNames = @{
                        'SecureStart.Settings.ppkg'              = 'Secured-core PC (UEFI security enforcement — enables VBS, HVCI, System Guard)'
                        'Microsoft.Windows.Cosa.Desktop.Client.ppkg' = 'COSA Desktop Client (cellular operator settings & APN database)'
                        'Microsoft.Windows.Cosa.Desktop.ppkg'   = 'COSA Desktop (cellular connectivity metadata)'
                        'ppkg_AzureADJoin.ppkg'                 = 'Azure AD Join (bulk enrollment provisioning)'
                        'Microsoft.Windows.SecureStartup.ppkg'  = 'Secure Startup (BitLocker & pre-boot authentication)'
                        'Microsoft.Windows.WindowsUpdate.ppkg'  = 'Windows Update (delivery optimization & policies)'
                        'Microsoft.Windows.Autopilot.ppkg'      = 'Windows Autopilot (OOBE enrollment profile)'
                    }

                    # ── Parse MDMDiagReport.xml for LAPS, policy metadata, and certificate resources ──
                    $mdmXmlPath = Join-Path $mdmDiagDir 'MDMDiagReport.xml'
                    if (Test-Path $mdmXmlPath) {
                        try {
                            [xml]$mdmDoc = [IO.File]::ReadAllText($mdmXmlPath)
                            $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] Parsing MDMDiagReport.xml ($([math]::Round((Get-Item $mdmXmlPath).Length/1KB))KB)";Level='INFO'})

                            # ── N3: LAPS policy section ──
                            $lapsNode = $mdmDoc.MDMEnterpriseDiagnosticsReport.LAPS
                            if ($lapsNode) {
                                $lapsCSP = $lapsNode.Laps_CSP_Policy
                                $lapsLocal = $lapsNode.Laps_Local_State
                                $BackupDirectoryMap = @{
                                    '0' = "Disabled (password won't be backed up)"
                                    '1' = 'Backup to Microsoft Entra ID only'
                                    '2' = 'Backup to Active Directory only'
                                }
                                $PasswordComplexityMap = @{
                                    '1' = 'Large letters'; '2' = 'Large + small letters'
                                    '3' = 'Letters + numbers'; '4' = 'Letters + numbers + special characters'
                                    '5' = 'Letters + numbers + special (improved readability)'
                                    '6' = 'Passphrase (long words)'; '7' = 'Passphrase (short words)'
                                    '8' = 'Passphrase (short words, unique prefixes)'
                                }
                                $PostAuthActionMap = @{
                                    '1' = 'Reset password'
                                    '3' = 'Reset password and logoff'
                                    '5' = 'Reset password and reboot'
                                    '11' = 'Reset password, logoff, and terminate processes'
                                }
                                $mdmInfo.MdmDiag.LAPS = @{
                                    BackupDirectory           = if ($lapsCSP.BackupDirectory) { $BackupDirectoryMap["$($lapsCSP.BackupDirectory)"] } else { 'Not Configured' }
                                    PasswordAgeDays           = if ($lapsCSP.PasswordAgeDays) { "$($lapsCSP.PasswordAgeDays) days" } else { '' }
                                    PasswordComplexity        = if ($lapsCSP.PasswordComplexity) { $PasswordComplexityMap["$($lapsCSP.PasswordComplexity)"] } else { '' }
                                    PasswordLength            = "$($lapsCSP.PasswordLength)"
                                    PostAuthActions           = if ($lapsCSP.PostAuthenticationActions) { $PostAuthActionMap["$($lapsCSP.PostAuthenticationActions)"] } else { '' }
                                    PostAuthResetDelay        = if ($lapsCSP.PostAuthenticationResetDelay) { "$($lapsCSP.PostAuthenticationResetDelay) hours" } else { '' }
                                    AutoManageEnabled         = if ("$($lapsCSP.AutomaticAccountManagementEnabled)" -eq '1') { 'Yes' } elseif ("$($lapsCSP.AutomaticAccountManagementEnabled)" -eq '0') { 'No' } else { '' }
                                    AutoManageTarget          = switch ("$($lapsCSP.AutomaticAccountManagementTarget)") { '0' { 'Built-in Administrator' } '1' { 'Custom account' } default { '' } }
                                    LocalLastPasswordUpdate   = if ($lapsLocal.LastPasswordUpdateTime) { try { [datetime]::FromFileTimeUtc([long]$lapsLocal.LastPasswordUpdateTime).ToString('yyyy-MM-dd HH:mm UTC') } catch { "$($lapsLocal.LastPasswordUpdateTime)" } } else { '' }
                                    LocalAzurePasswordExpiry  = if ($lapsLocal.AzurePasswordExpiryTime) { try { [datetime]::FromFileTimeUtc([long]$lapsLocal.AzurePasswordExpiryTime).ToString('yyyy-MM-dd HH:mm UTC') } catch { "$($lapsLocal.AzurePasswordExpiryTime)" } } else { '' }
                                    LocalManagedAccountName   = "$($lapsLocal.LastManagedAccountNameOrPrefix)"
                                }
                                $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] LAPS policy found: Backup=$($mdmInfo.MdmDiag.LAPS.BackupDirectory), PwdAge=$($mdmInfo.MdmDiag.LAPS.PasswordAgeDays)";Level='INFO'})
                            }

                            # ── N6: Policy metadata (AreaMetadata → default values, redirection paths) ──
                            $metaNode = $mdmDoc.MDMEnterpriseDiagnosticsReport.PolicyManagerMeta
                            if ($metaNode) {
                                $policyMetaHash = @{}
                                foreach ($area in $metaNode.AreaMetadata) {
                                    $areaName = $area.PolicyAreaName
                                    if (-not $areaName) { continue }
                                    foreach ($pm in $area.PolicyMetadata) {
                                        $pName = $pm.PolicyName
                                        if (-not $pName) { continue }
                                        $metaKey = "$areaName/$pName"
                                        $policyMetaHash[$metaKey] = @{
                                            DefaultValue = "$($pm.value)"
                                            RegPath      = if ($pm.RegKeyPathRedirect) { "$($pm.RegKeyPathRedirect)" } elseif ($pm.grouppolicyPath) { "GP: $($pm.grouppolicyPath)" } else { '' }
                                        }
                                    }
                                }
                                $mdmInfo.MdmDiag.PolicyMeta = $policyMetaHash
                                $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] PolicyManagerMeta: $($policyMetaHash.Count) metadata entries loaded";Level='INFO'})
                            }

                            # ── N2: Certificate resources (enrich with local cert store lookup) ──
                            $resNode = $mdmDoc.MDMEnterpriseDiagnosticsReport.Resources
                            if ($resNode) {
                                $certList = [System.Collections.Generic.List[hashtable]]::new()
                                foreach ($enrollment in $resNode.Enrollment) {
                                    foreach ($scope in $enrollment.Scope) {
                                        foreach ($resource in $scope.Resources.ChildNodes.'#Text') {
                                            if ($resource -match 'RootCATrustedCertificates') {
                                                $thumbprint = ($resource | Split-Path -Leaf -EA SilentlyContinue)
                                                if (-not $thumbprint -or $thumbprint.Length -lt 20) { continue }
                                                $certStoreName = if ($resource -match '/Root/') { 'Root CA' } elseif ($resource -match '/CA/') { 'Intermediate CA' } elseif ($resource -match '/TrustedPublisher/') { 'Trusted Publisher' } else { 'Other' }
                                                $pathType = if ($resource -match '^\.\/device\/') { 'LocalMachine' } else { 'CurrentUser' }
                                                $certStoreMap = @{ 'Root CA'='Root'; 'Intermediate CA'='CA'; 'Trusted Publisher'='TrustedPublisher'; 'Other'='My' }
                                                $certPath = "Cert:\$pathType\$($certStoreMap[$certStoreName])\$thumbprint"
                                                $certInfo = @{
                                                    Store = $certStoreName; Thumbprint = $thumbprint; Scope = $pathType
                                                    IssuedTo = ''; IssuedBy = ''; ValidFrom = ''; ValidTo = ''; ExpireDays = ''
                                                }
                                                if (Test-Path $certPath) {
                                                    $cert = Get-Item $certPath -EA SilentlyContinue
                                                    if ($cert) {
                                                        $certInfo.IssuedTo   = ($cert.Subject -replace '^.*CN=([^,]+).*$','$1' -replace '^CN=','')
                                                        $certInfo.IssuedBy   = ($cert.Issuer -replace '^.*CN=([^,]+).*$','$1' -replace '^CN=','')
                                                        $certInfo.ValidFrom  = $cert.NotBefore.ToString('yyyy-MM-dd')
                                                        $certInfo.ValidTo    = $cert.NotAfter.ToString('yyyy-MM-dd')
                                                        $certInfo.ExpireDays = [math]::Round(($cert.NotAfter - (Get-Date)).TotalDays, 0)
                                                    }
                                                }
                                                $certList.Add($certInfo)
                                            }
                                        }
                                    }
                                }
                                if ($certList.Count -gt 0) {
                                    # Replace basic HTML-parsed certs with richer XML-based cert inventory
                                    $mdmInfo.MdmDiag.Certificates = @($certList)
                                    $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] Certificate inventory: $($certList.Count) certs from XML Resources (local lookup enriched)";Level='INFO'})
                                }
                            }

                            # ── N8: Configuration profile delivery status ──
                            $cfgProfiles = [System.Collections.Generic.List[hashtable]]::new()
                            # Check ConfigurationSources
                            $cfgSrc = $mdmDoc.MDMEnterpriseDiagnosticsReport.ConfigurationSources
                            if ($cfgSrc) {
                                foreach ($src in $cfgSrc.ChildNodes) {
                                    if ($src.LocalName -eq '#comment') { continue }
                                    $profileName = $src.LocalName
                                    $profileState = if ($src.State) { $src.State } elseif ($src.InnerText) { $src.InnerText.Trim() } else { 'Unknown' }
                                    $cfgProfiles.Add(@{ Name = $profileName; Status = $profileState; Source = 'ConfigurationSources' })
                                }
                            }
                            # Check Policies node for profile-level status
                            $policiesNode = $mdmDoc.MDMEnterpriseDiagnosticsReport.Policies
                            if ($policiesNode) {
                                foreach ($policy in $policiesNode.ChildNodes) {
                                    if ($policy.LocalName -eq '#comment' -or -not $policy.LocalName) { continue }
                                    $scope = $policy.LocalName  # typically 'device' or 'user'
                                    foreach ($area in $policy.ChildNodes) {
                                        if ($area.LocalName -eq '#comment' -or -not $area.LocalName) { continue }
                                        $areaName = $area.LocalName
                                        $statusAttr = $area.GetAttribute('Status')
                                        $errAttr = $area.GetAttribute('Error')
                                        if ($statusAttr -or $errAttr) {
                                            $profStatus = if ($errAttr -and $errAttr -ne '0' -and $errAttr -ne '') { "Error: $errAttr" } elseif ($statusAttr) { $statusAttr } else { 'Applied' }
                                            $cfgProfiles.Add(@{ Name = "$scope/$areaName"; Status = $profStatus; Source = 'Policies' })
                                        }
                                    }
                                }
                            }
                            if ($cfgProfiles.Count -gt 0) {
                                $mdmInfo.MdmDiag.ConfigProfiles = @($cfgProfiles)
                                $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] Config profiles: $($cfgProfiles.Count) profile status entries";Level='INFO'})
                            }

                            # ── N9: Provisioning Packages (.ppkg) ──
                            # Parse <ProvisioningPackages><Result> elements for installed ppkg info.
                            $ppkgNode = $mdmDoc.MDMEnterpriseDiagnosticsReport.ProvisioningPackages
                            $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] ProvisioningPackages XML node present: $([bool]$ppkgNode)";Level='DEBUG'})
                            if ($ppkgNode) {
                                $ppkgList = [System.Collections.Generic.List[hashtable]]::new()
                                foreach ($result in $ppkgNode.Result) {
                                    if (-not $result.PackageId) { continue }
                                    $pkgFile = "$($result.PackageFileName)"
                                    $friendlyName = if ($ppkgFriendlyNames.ContainsKey($pkgFile)) { $ppkgFriendlyNames[$pkgFile] } else { '' }
                                    # Parse provisioning XML entries under Statistics
                                    $xmlEntries = [System.Collections.Generic.List[hashtable]]::new()
                                    $statsNode = $result.Statistics
                                    if ($statsNode) {
                                        foreach ($xmlNode in $statsNode.XML) {
                                            $xmlEntries.Add(@{
                                                XMLName          = "$($xmlNode.XMLName)"
                                                Area             = "$($xmlNode.Area)"
                                                Message          = "$($xmlNode.Message)"
                                                LastResult       = "$($xmlNode.LastResult)"
                                                NumberOfFailures = "$($xmlNode.NumberOfFailures)"
                                            })
                                        }
                                    }
                                    $totalFailures = ($xmlEntries | ForEach-Object { [int]$_.NumberOfFailures } | Measure-Object -Sum).Sum
                                    $ppkgList.Add(@{
                                        PackageId    = "$($result.PackageId)"
                                        FileName     = $pkgFile
                                        FriendlyName = $friendlyName
                                        XMLEntries   = @($xmlEntries)
                                        TotalFailures = $totalFailures
                                        Status       = if ($totalFailures -gt 0) { 'HasFailures' } else { 'OK' }
                                    })
                                }
                                if ($ppkgList.Count -gt 0) {
                                    $mdmInfo.MdmDiag.ProvisioningPackages = @($ppkgList)
                                    $failCount = @($ppkgList | Where-Object { $_.TotalFailures -gt 0 }).Count
                                    $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] Provisioning packages: $($ppkgList.Count) packages found ($failCount with failures)";Level='INFO'})
                                } else {
                                    $SyncH.StatusQueue.Enqueue(@{Type='Log';Text='[Intune] ProvisioningPackages XML node found but 0 <Result> entries';Level='DEBUG'})
                                }
                            } else {
                                $SyncH.StatusQueue.Enqueue(@{Type='Log';Text='[Intune] No <ProvisioningPackages> section in MDMDiagReport.xml — will try Get-ProvisioningPackage cmdlet';Level='DEBUG'})
                            }
                        } catch {
                            $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] MDMDiagReport.xml parse error: $($_.Exception.Message)";Level='WARN'})
                        }
                    }

                    # ── Enrich provisioning packages via Get-ProvisioningPackage cmdlet ──
                    # The cmdlet returns PackageName (human-readable), Owner, Version, PackagePath, IsApplied,
                    # which the MDMDiagReport.xml does not. Cross-reference by PackageId.
                    try {
                        if (Get-Command Get-ProvisioningPackage -ErrorAction SilentlyContinue) {
                            $SyncH.StatusQueue.Enqueue(@{Type='Log';Text='[Intune] Querying Get-ProvisioningPackage -AllInstalledPackages...';Level='INFO'})
                            $cmdletPkgs = @(Get-ProvisioningPackage -AllInstalledPackages -ErrorAction Stop)
                            $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] Get-ProvisioningPackage returned $($cmdletPkgs.Count) packages";Level='INFO'})

                            # Build lookup from existing XML-parsed packages by PackageId
                            $existingById = @{}
                            if ($mdmInfo.MdmDiag.ProvisioningPackages) {
                                foreach ($p in $mdmInfo.MdmDiag.ProvisioningPackages) { $existingById[$p.PackageId] = $p }
                            }

                            $enrichedList = [System.Collections.Generic.List[hashtable]]::new()
                            foreach ($cp in $cmdletPkgs) {
                                $cpId = "$($cp.PackageId)"
                                if ($existingById.ContainsKey($cpId)) {
                                    # Enrich existing XML-parsed entry with cmdlet metadata
                                    $existing = $existingById[$cpId]
                                    if ($cp.PackageName) { $existing.PackageName = "$($cp.PackageName)" }
                                    if (-not $existing.FriendlyName -and $cp.PackageName) {
                                        $existing.FriendlyName = "$($cp.PackageName)"
                                    }
                                    $existing.Owner    = "$($cp.Owner)"
                                    $existing.Version  = "$($cp.Version)"
                                    $existing.Rank     = "$($cp.Rank)"
                                    $existing.IsApplied = if ($cp.IsApplied) { $true } else { $false }
                                    if ($cp.PackagePath) { $existing.PackagePath = "$($cp.PackagePath)" }
                                    $enrichedList.Add($existing)
                                    $existingById.Remove($cpId)
                                } else {
                                    # Package found by cmdlet but not in MDMDiagReport.xml — add it
                                    $pkgFile = if ($cp.PackagePath) { [System.IO.Path]::GetFileName($cp.PackagePath) } else { '' }
                                    $friendlyName = "$($cp.PackageName)"
                                    if (-not $friendlyName -and $ppkgFriendlyNames.ContainsKey($pkgFile)) {
                                        $friendlyName = $ppkgFriendlyNames[$pkgFile]
                                    }
                                    $enrichedList.Add(@{
                                        PackageId    = $cpId
                                        FileName     = $pkgFile
                                        FriendlyName = $friendlyName
                                        PackageName  = "$($cp.PackageName)"
                                        Owner        = "$($cp.Owner)"
                                        Version      = "$($cp.Version)"
                                        Rank         = "$($cp.Rank)"
                                        IsApplied    = if ($cp.IsApplied) { $true } else { $false }
                                        PackagePath  = "$($cp.PackagePath)"
                                        XMLEntries   = @()
                                        TotalFailures = 0
                                        Status       = 'OK'
                                    })
                                }
                            }
                            # Add any XML-only packages that the cmdlet didn't return
                            foreach ($remaining in $existingById.Values) { $enrichedList.Add($remaining) }

                            $mdmInfo.MdmDiag.ProvisioningPackages = @($enrichedList)
                            $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] Provisioning packages enriched: $($enrichedList.Count) total (cmdlet=$($cmdletPkgs.Count), XML-only=$($existingById.Count))";Level='INFO'})
                        } else {
                            $SyncH.StatusQueue.Enqueue(@{Type='Log';Text='[Intune] Get-ProvisioningPackage cmdlet not available — using XML data only';Level='DEBUG'})
                        }
                    } catch {
                        $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] Get-ProvisioningPackage failed: $($_.Exception.Message) — using XML data only";Level='WARN'})
                    }
                } else {
                    $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] mdmdiagnosticstool exit code: $($mdmProc.ExitCode)";Level='WARN'})
                }
            } catch {
                $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] mdmdiagnosticstool failed: $($_.Exception.Message)";Level='WARN'})
            } }


            # ──────────────────────────────────────────────────────────────────
            # 1) PolicyManager current device policies (enriched)
            # ──────────────────────────────────────────────────────────────────
            $pmPath = 'HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device'
            $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] Checking: $pmPath  exists=$(Test-Path $pmPath)";Level='DEBUG'})
            if (Test-Path $pmPath) {
                $SyncH.StatusQueue.Enqueue(@{Type='Log';Text='[Intune] Reading PolicyManager device policies...';Level='INFO'})
                $areas = Get-ChildItem $pmPath -ErrorAction SilentlyContinue
                $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] Found $($areas.Count) CSP policy areas";Level='DEBUG'})
                $areaIdx = 0; $areaTotal = $areas.Count
                foreach ($area in $areas) {
                    $areaIdx++
                    if ($areaIdx % 10 -eq 0 -or $areaIdx -eq 1 -or $areaIdx -eq $areaTotal) {
                        $areaPct = [math]::Round($areaIdx / [math]::Max($areaTotal,1) * 100)
                        $SyncH.StatusQueue.Enqueue(@{Type='Progress';Value=$areaPct})
                        $SyncH.StatusQueue.Enqueue(@{Type='Status';Text="Reading policy area $areaIdx/$areaTotal ($areaPct%)";Color='#F59E0B'})
                    }
                    $areaName = $area.PSChildName
                    $friendlyArea = if ($cspAreaNames[$areaName]) { $cspAreaNames[$areaName] }
                                    elseif ($areaName -match '~') { $admxFriendly = Format-AdmxAreaName $areaName; if ($admxFriendly) { $admxFriendly } else { $areaName } }
                                    else { $areaName }
                    $props = Get-ItemProperty $area.PSPath -ErrorAction SilentlyContinue
                    if (-not $props) { continue }

                    # Collect and group properties: skip PS*, _ProviderSet, _WinningProvider
                    $policyProps = @{}; $providerInfo = @{}
                    foreach ($pn in $props.PSObject.Properties) {
                        if ($pn.Name -like 'PS*') { continue }
                        $baseName = Get-BaseSetting $pn.Name
                        if ($pn.Name -match '_WinningProvider$') {
                            $providerInfo[$baseName] = "$($pn.Value)"
                        } elseif ($pn.Name -match '_(ProviderSet|LastWrite|CurrentChannel|ADMXInstanceData)$') {
                            continue  # skip internal MDM tracking
                        } else {
                            $policyProps[$pn.Name] = $pn.Value
                        }
                    }

                    if ($policyProps.Count -eq 0) { continue }

                    $settingCount2 = 0
                    $gpoIdCounter++
                    [void]$gpoRecords.Add([PSCustomObject]@{
                        Id=$gpoIdCounter; DisplayName=$friendlyArea; GpoId=$areaName; Status='Applied'
                        CreatedTime=''; ModifiedTime=''
                        WmiFilter=''; LinkPath="Intune > $friendlyArea"; IsLinked=$true
                        UserVersion='0'; ComputerVersion='0'; Description="CSP: $areaName - Managed by Intune/MDM"
                        SettingCount=0; LinkCount=1; Links='Intune MDM'
                    })

                    foreach ($key in $policyProps.Keys) {
                        $settIdCounter++; $settingCount2++
                        $lookupKey = "$areaName/$key"
                        $meta = $cspMeta[$lookupKey]
                        $xmlMeta = $mdmInfo.MdmDiag.PolicyMeta[$lookupKey]
                        $policyName = if ($meta) { $meta.Friendly } else { Format-PolicyName $key }
                        $category   = if ($meta) { $meta.Cat } else { $friendlyArea }
                        $desc       = if ($meta) { $meta.Desc } else { '' }
                        $provider   = if ($providerInfo[$key]) { $providerInfo[$key] } else { '' }
                        $val = $policyProps[$key]
                        $valStr = if ($val -is [System.Array]) { ($val | ForEach-Object { "$_" }) -join "`n" } else { "$val" }
                        $valStr = Format-NumberedList $valStr
                        $displayVal = if ($desc -and $valStr -match '^\d+$') { "$valStr - $desc" } else { $valStr }
                        $defVal = if ($meta -and $meta.Def) { $meta.Def } elseif ($xmlMeta -and $xmlMeta.DefaultValue) { $xmlMeta.DefaultValue } else { '' }
                        $regKey = if ($xmlMeta -and $xmlMeta.RegPath) { $xmlMeta.RegPath } else { "HKLM\SOFTWARE\Microsoft\PolicyManager\current\device\$areaName\$key" }

                        [void]$allSettingsList.Add([PSCustomObject]@{
                            Id=$settIdCounter; GPOName=$friendlyArea; GPOGuid=$areaName
                            Category=$category; PolicyName=$policyName
                            SettingKey="$areaName/$key"; State='Applied'
                            RegistryKey=$regKey
                            ValueData=$displayVal; Scope='Device'; DefaultValue=$defVal; Source='Intune'; IntuneGroup=(Get-IntuneGroup $category 'Intune')
                        })
                    }
                    $gpoRecords[$gpoRecords.Count - 1].SettingCount = $settingCount2
                    if ($settingCount2 -gt 0) {
                        $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune]   $friendlyArea ($areaName): $settingCount2 policies";Level='DEBUG'})
                    }
                }
                $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] Device scan: $($gpoRecords.Count) areas, $($allSettingsList.Count) policies";Level='DEBUG'})
            } else {
                $SyncH.StatusQueue.Enqueue(@{Type='Log';Text='[Intune] PolicyManager device path not found';Level='WARN'})
            }

            # ──────────────────────────────────────────────────────────────────
            # 2) PolicyManager current user policies (enriched)
            # ──────────────────────────────────────────────────────────────────
            $pmUserPath = 'HKLM:\SOFTWARE\Microsoft\PolicyManager\current\user'
            $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] Checking: $pmUserPath  exists=$(Test-Path $pmUserPath)";Level='DEBUG'})
            if (Test-Path $pmUserPath) {
                $SyncH.StatusQueue.Enqueue(@{Type='Log';Text='[Intune] Reading PolicyManager user policies...';Level='INFO'})
                $userDirs = Get-ChildItem $pmUserPath -ErrorAction SilentlyContinue
                foreach ($userDir in $userDirs) {
                    $userAreas = Get-ChildItem $userDir.PSPath -ErrorAction SilentlyContinue
                    foreach ($area in $userAreas) {
                        $areaName = $area.PSChildName
                        $friendlyArea = if ($cspAreaNames[$areaName]) { $cspAreaNames[$areaName] }
                                        elseif ($areaName -match '~') { $admxFriendly = Format-AdmxAreaName $areaName; if ($admxFriendly) { $admxFriendly } else { $areaName } }
                                        else { $areaName }
                        $props = Get-ItemProperty $area.PSPath -ErrorAction SilentlyContinue
                        if (-not $props) { continue }

                        $policyProps = @{}; $providerInfo = @{}
                        foreach ($pn in $props.PSObject.Properties) {
                            if ($pn.Name -like 'PS*') { continue }
                            $baseName = Get-BaseSetting $pn.Name
                            if ($pn.Name -match '_WinningProvider$') {
                                $providerInfo[$baseName] = "$($pn.Value)"
                            } elseif ($pn.Name -match '_(ProviderSet|LastWrite|CurrentChannel|ADMXInstanceData)$') {
                                continue
                            } else {
                                $policyProps[$pn.Name] = $pn.Value
                            }
                        }

                        if ($policyProps.Count -eq 0) { continue }

                        $settingCount2 = 0
                        $gpoIdCounter++
                        [void]$gpoRecords.Add([PSCustomObject]@{
                            Id=$gpoIdCounter; DisplayName="$friendlyArea (User)"; GpoId="$($userDir.PSChildName)/$areaName"; Status='Applied'
                            CreatedTime=''; ModifiedTime=''
                            WmiFilter=''; LinkPath="Intune > $friendlyArea > User"; IsLinked=$true
                            UserVersion='0'; ComputerVersion='0'; Description="CSP: $areaName - User scope MDM policy"
                            SettingCount=0; LinkCount=1; Links='Intune MDM (User)'
                        })
                        foreach ($key in $policyProps.Keys) {
                            $settIdCounter++; $settingCount2++
                            $lookupKey = "$areaName/$key"
                            $meta = $cspMeta[$lookupKey]
                            $xmlMeta = $mdmInfo.MdmDiag.PolicyMeta[$lookupKey]
                            $policyName = if ($meta) { $meta.Friendly } else { Format-PolicyName $key }
                            $category   = if ($meta) { $meta.Cat } else { "$friendlyArea (User)" }
                            $val = $policyProps[$key]
                            $desc = if ($meta) { $meta.Desc } else { '' }
                            $valStr = if ($val -is [System.Array]) { ($val | ForEach-Object { "$_" }) -join "`n" } else { "$val" }
                            $valStr = Format-NumberedList $valStr
                            $displayVal = if ($desc -and $valStr -match '^\d+$') { "$valStr - $desc" } else { $valStr }
                            $defVal = if ($meta -and $meta.Def) { $meta.Def } elseif ($xmlMeta -and $xmlMeta.DefaultValue) { $xmlMeta.DefaultValue } else { '' }
                            $regKey = if ($xmlMeta -and $xmlMeta.RegPath) { $xmlMeta.RegPath } else { ($area.PSPath -replace 'Microsoft.PowerShell.Core\\Registry::','') }

                            [void]$allSettingsList.Add([PSCustomObject]@{
                                Id=$settIdCounter; GPOName="$friendlyArea (User)"; GPOGuid="$($userDir.PSChildName)/$areaName"
                                Category=$category; PolicyName=$policyName
                                SettingKey="$areaName/$key"; State='Applied'
                                RegistryKey=$regKey
                                ValueData=$displayVal; Scope='User'; DefaultValue=$defVal; Source='Intune'; IntuneGroup=(Get-IntuneGroup $category 'Intune')
                            })
                        }
                        $gpoRecords[$gpoRecords.Count - 1].SettingCount = $settingCount2
                    }
                }
                $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] User scan: $($gpoRecords.Count) areas, $($allSettingsList.Count) policies";Level='DEBUG'})
            }

            # ──────────────────────────────────────────────────────────────────
            # 3) MDM WMI Bridge policies (Win10+)
            # ──────────────────────────────────────────────────────────────────
            $SyncH.StatusQueue.Enqueue(@{Type='Log';Text='[Intune] Reading MDM WMI Bridge (dmmap)...';Level='INFO'})
            try {
                $mdmClasses = Get-CimClass -Namespace 'root/cimv2/mdm/dmmap' -ErrorAction Stop |
                    Where-Object { $_.CimClassName -like 'MDM_Policy_Config01_*' }
                foreach ($cls in $mdmClasses) {
                    $className = $cls.CimClassName
                    $areaName  = ($className -replace 'MDM_Policy_Config01_','') -replace '\d+$',''
                    $friendlyArea = if ($cspAreaNames[$areaName]) { "WMI: $($cspAreaNames[$areaName])" } else { "WMI: $areaName" }
                    try {
                        $instances = Get-CimInstance -Namespace 'root/cimv2/mdm/dmmap' -ClassName $className -ErrorAction Stop
                        foreach ($inst in $instances) {
                            $settingCount2 = 0
                            $gpoIdCounter++
                            [void]$gpoRecords.Add([PSCustomObject]@{
                                Id=$gpoIdCounter; DisplayName=$friendlyArea; GpoId=$className; Status='Applied'
                                CreatedTime=''; ModifiedTime=''
                                WmiFilter=''; LinkPath='Intune > MDM WMI Bridge'; IsLinked=$true
                                UserVersion='0'; ComputerVersion='0'; Description="WMI Bridge: $className"
                                SettingCount=0; LinkCount=1; Links='Intune MDM WMI'
                            })
                            foreach ($prop in $inst.CimInstanceProperties) {
                                if ($prop.Name -in @('InstanceID','ParentID') -or $null -eq $prop.Value) { continue }
                                $settIdCounter++; $settingCount2++
                                $lookupKey = "$areaName/$($prop.Name)"
                                $meta = $cspMeta[$lookupKey]
                                $policyName = if ($meta) { $meta.Friendly } else { Format-PolicyName $prop.Name }
                                $category   = if ($meta) { $meta.Cat } else { $friendlyArea }

                                [void]$allSettingsList.Add([PSCustomObject]@{
                                    Id=$settIdCounter; GPOName=$friendlyArea; GPOGuid=$className
                                    Category=$category; PolicyName=$policyName
                                    SettingKey="WMI|$className\$($prop.Name)"; State='Applied'
                                    RegistryKey="root/cimv2/mdm/dmmap:$className"; ValueData="$($prop.Value)"; Scope='Device'; DefaultValue=$(if ($meta -and $meta.Def) { $meta.Def } else { '' }); Source='Intune'; IntuneGroup=(Get-IntuneGroup $category 'Intune')
                                })
                            }
                            $gpoRecords[$gpoRecords.Count - 1].SettingCount = $settingCount2
                        }
                    } catch { $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] WMI $className`: $($_.Exception.Message)";Level='DEBUG'}) }
                }
                $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] WMI Bridge: $($mdmClasses.Count) config classes";Level='INFO'})
            } catch {
                $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] MDM WMI Bridge not available: $($_.Exception.Message)";Level='WARN'})
            }

            $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] Scan complete: $($gpoRecords.Count) policy areas, $($allSettingsList.Count) settings";Level='SUCCESS'})
        }

        # ──────────────────────────────────────────────────────────────────
        # Intune App Tracking (Win32 + LOB + Store apps from IME registry)
        # ──────────────────────────────────────────────────────────────────
        if ($runIntune) {
            $SyncH.StatusQueue.Enqueue(@{Type='Log';Text='[Intune] Scanning managed app installations...';Level='STEP'})
            $appId = 0
            # Helper: get registry key last-write timestamp via Win32 RegQueryInfoKey
            $regTsSig = @'
using System;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
public static class RegKeyTs {
    [DllImport("advapi32.dll")] public static extern int RegQueryInfoKey(
        SafeRegistryHandle hKey, IntPtr cls, IntPtr clsLen, IntPtr reserved,
        IntPtr subKeys, IntPtr maxSubKey, IntPtr maxClass, IntPtr values,
        IntPtr maxValName, IntPtr maxValData, IntPtr secDesc, out long lastWrite);
}
'@
            try { Add-Type -TypeDefinition $regTsSig -ErrorAction SilentlyContinue } catch { try { Write-DebugLog "Unhandled: $_" -Level ERROR } catch {} }
            function Get-RegistryKeyTimestamp ([string]$Path) {
                try {
                    $item = Get-Item $Path -ErrorAction Stop
                    $ft = [long]0
                    if ([RegKeyTs]::RegQueryInfoKey($item.Handle,[IntPtr]::Zero,[IntPtr]::Zero,
                        [IntPtr]::Zero,[IntPtr]::Zero,[IntPtr]::Zero,[IntPtr]::Zero,[IntPtr]::Zero,
                        [IntPtr]::Zero,[IntPtr]::Zero,[IntPtr]::Zero,[ref]$ft) -eq 0 -and $ft -gt 0) {
                        return [datetime]::FromFileTimeUtc($ft).ToLocalTime().ToString('yyyy-MM-dd HH:mm')
                    }
                } catch { try { Write-DebugLog "Unhandled: $_" -Level ERROR } catch {} }
                return ''
            }


            # Build Id->Name map + ReportingState map from IME AppWorkload logs
            $imeNameMap  = @{}
            $imeStateMap = @{}
            $imeLogDir = 'C:\ProgramData\Microsoft\IntuneManagementExtension\Logs'
            if (Test-Path $imeLogDir) {
                $logFiles = @(Get-ChildItem $imeLogDir -Filter 'AppWorkload*.log' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)
                foreach ($lf in $logFiles) {
                    try {
                        $logPath = Join-Path ([IO.Path]::GetTempPath()) "PolicyPilot_$($lf.Name)"
                        [IO.File]::Copy($lf.FullName, $logPath, $true)
                        $logText = [IO.File]::ReadAllText($logPath)
                        # Extract Id -> Name from policy JSON
                        foreach ($m in [regex]::Matches($logText, '"Id"\s*:\s*"([^"]+)"[^}]{0,500}?"Name"\s*:\s*"([^"]+)"')) {
                            $imeNameMap[$m.Groups[1].Value] = $m.Groups[2].Value
                        }
                        # Extract ReportingState JSON per app (last entry wins for each app)
                        foreach ($m in [regex]::Matches($logText, 'ReportingState: (\{[^}]+\})')) {
                            try {
                                $rs = $m.Groups[1].Value | ConvertFrom-Json
                                if ($rs.ApplicationId) { $imeStateMap[$rs.ApplicationId] = $rs }
                            } catch {}
                        }
                        [IO.File]::Delete($logPath)
                    } catch { try { Write-DebugLog "Unhandled: $_" -Level ERROR } catch {} }
                }
                if ($imeNameMap.Count -gt 0) {
                    $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] IME log name map: $($imeNameMap.Count) apps, $($imeStateMap.Count) with state";Level='DEBUG'})
                }

                # ── N4: Parse HealthScripts*.log for remediation/detection script results ──
                $scriptPoliciesList = [System.Collections.Generic.List[hashtable]]::new()
                $hsLogFiles = @(Get-ChildItem $imeLogDir -Filter 'HealthScripts*.log' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)
                if ($hsLogFiles.Count -gt 0) {
                    $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] Parsing $($hsLogFiles.Count) HealthScripts log file(s)...";Level='INFO'})
                    $seenScripts = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                    foreach ($hsLf in $hsLogFiles) {
                        try {
                            $hsLogPath = Join-Path ([IO.Path]::GetTempPath()) "PolicyPilot_$($hsLf.Name)"
                            [IO.File]::Copy($hsLf.FullName, $hsLogPath, $true)
                            $hsText = [IO.File]::ReadAllText($hsLogPath)
                            # Parse detection script results: "policyId":"<guid>","result":<code>,"resultType":"Detection"
                            foreach ($dm in [regex]::Matches($hsText, '"policyId"\s*:\s*"([^"]+)"[^}]{0,800}?"result"\s*:\s*(\d+)[^}]{0,200}?"resultType"\s*:\s*"([^"]+)"')) {
                                $policyId = $dm.Groups[1].Value
                                $result = $dm.Groups[2].Value
                                $resultType = $dm.Groups[3].Value
                                $scriptKey = "$policyId|$resultType"
                                if ($seenScripts.Contains($scriptKey)) { continue }
                                [void]$seenScripts.Add($scriptKey)
                                $resultText = switch ($result) { '0' { 'Success' } '1' { 'Failure' } default { "Exit code $result" } }
                                $scriptPoliciesList.Add(@{
                                    PolicyId = $policyId; ScriptType = $resultType
                                    Result = $resultText; ExitCode = $result
                                    ScriptName = ''
                                })
                            }
                            # Parse script names: "policyId":"<guid>","policyName":"<name>"
                            foreach ($nm in [regex]::Matches($hsText, '"policyId"\s*:\s*"([^"]+)"[^}]{0,400}?"policyName"\s*:\s*"([^"]+)"')) {
                                $pid2 = $nm.Groups[1].Value
                                $pName = $nm.Groups[2].Value
                                foreach ($sp in $scriptPoliciesList) {
                                    if ($sp.PolicyId -eq $pid2 -and -not $sp.ScriptName) { $sp.ScriptName = $pName }
                                }
                            }
                            # Parse remediation execution: "policyId":"<guid>","remediationStatus":<code>
                            foreach ($rm in [regex]::Matches($hsText, '"policyId"\s*:\s*"([^"]+)"[^}]{0,600}?"remediationStatus"\s*:\s*(\d+)')) {
                                $pid3 = $rm.Groups[1].Value
                                $remStatus = $rm.Groups[2].Value
                                $remText = switch ($remStatus) { '1' { 'Remediation Success' } '2' { 'Remediation Failed' } '3' { 'Remediation Skipped' } default { "Remediation status $remStatus" } }
                                $scriptKey3 = "$pid3|Remediation"
                                if (-not $seenScripts.Contains($scriptKey3)) {
                                    [void]$seenScripts.Add($scriptKey3)
                                    $existingName = ($scriptPoliciesList | Where-Object { $_.PolicyId -eq $pid3 } | Select-Object -First 1).ScriptName
                                    $scriptPoliciesList.Add(@{
                                        PolicyId = $pid3; ScriptType = 'Remediation'
                                        Result = $remText; ExitCode = $remStatus
                                        ScriptName = if ($existingName) { $existingName } else { '' }
                                    })
                                }
                            }
                            [IO.File]::Delete($hsLogPath)
                        } catch { try { Write-DebugLog "Unhandled: $_" -Level ERROR } catch {} }
                    }
                    if ($scriptPoliciesList.Count -gt 0) {
                        $mdmInfo.MdmDiag.ScriptPolicies = @($scriptPoliciesList)
                        $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] HealthScripts: $($scriptPoliciesList.Count) script policy results parsed";Level='INFO'})
                    }
                }
            }

            # Helper: resolve an MSI product code or app GUID to a friendly display name + version
            # Checks Uninstall registry (64-bit + 32-bit) for DisplayName / DisplayVersion
            function Resolve-AppDisplayName ([string]$AppGuid, $MsiProps) {
                # 1) Check Uninstall registry (64-bit + 32-bit)
                foreach ($root in @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
                                    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall')) {
                    $regPath = Join-Path $root $AppGuid
                    if (Test-Path $regPath) {
                        $p = Get-ItemProperty $regPath -ErrorAction SilentlyContinue
                        if ($p.DisplayName) {
                            return @{ Name = "$($p.DisplayName)"; Version = "$($p.DisplayVersion)" }
                        }
                    }
                }
                # 2) MSI packed-GUID lookup in Classes\Installer\Products
                if ($AppGuid -match '^\{') {
                    $g = $AppGuid -replace '[{}\-]',''
                    if ($g.Length -eq 32) {
                        $c = $g.ToCharArray()
                        [array]::Reverse($c, 0, 8); [array]::Reverse($c, 8, 4); [array]::Reverse($c, 12, 4)
                        $tail = $g.Substring(16); $sw = for ($i=0;$i -lt $tail.Length;$i+=2) { $tail[$i+1]; $tail[$i] }
                        $packed = (-join $c[0..15]) + (-join $sw)
                        $prodPath = "HKLM:\SOFTWARE\Classes\Installer\Products\$packed"
                        if (Test-Path $prodPath) {
                            $pp = Get-ItemProperty $prodPath -ErrorAction SilentlyContinue
                            if ($pp.ProductName) {
                                $ver = if ($MsiProps.ProductVersion) { $MsiProps.ProductVersion } else { '' }
                                return @{ Name = $pp.ProductName; Version = $ver }
                            }
                        }
                    }
                }
                # 3) Extract name from DownloadUrlList (Intune download URL contains MSI filename)
                if ($MsiProps.DownloadUrlList) {
                    try {
                        $seg = ([uri]($MsiProps.DownloadUrlList -split '[\s]')[0]).Segments[-1]
                        $fn  = [System.IO.Path]::GetFileNameWithoutExtension($seg)
                        if ($fn) {
                            $ver = if ($MsiProps.ProductVersion) { $MsiProps.ProductVersion } else { '' }
                            return @{ Name = $fn; Version = $ver }
                        }
                    } catch { try { Write-DebugLog "Unhandled: $_" -Level ERROR } catch {} }
                }
                # 4) Fallback: IME log name map
                if ($imeNameMap.ContainsKey($AppGuid)) {
                    return @{ Name = $imeNameMap[$AppGuid]; Version = '' }
                }
                return $null
            }

            # 1) IntuneManagementExtension - Win32 & LOB apps
            # Structure: Win32Apps\{userId}\{appGuid}_N (revision suffix) + GRS\ (retry schedule, skip)
            $imePath = 'HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps'
            $win32Seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            $win32RegMap = @{}  # appGuid -> registry path (for ALL entries, even skipped, so IME backfill can reset)
            $win32Skipped = 0
            if (Test-Path $imePath) {
                $userDirs = $null
                try { $userDirs = @(Get-ChildItem $imePath -ErrorAction Stop) } catch {
                    $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] Win32Apps registry: access denied (run elevated for full results)";Level='WARN'})
                }
                foreach ($userDir in $userDirs) {
                    $userId = $userDir.PSChildName
                    $appKeys = Get-ChildItem $userDir.PSPath -ErrorAction SilentlyContinue
                    foreach ($appKey in $appKeys) {
                        $rawName = $appKey.PSChildName
                        # Skip GRS (Global Retry Schedule) keys - not apps
                        if ($rawName -eq 'GRS') { continue }
                        # Strip revision suffix (_1, _2, etc.) to get the real app GUID
                        $appGuid = $rawName -replace '_\d+$', ''
                        # Skip non-GUID entries (metadata keys)
                        if ($appGuid -notmatch '^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$') { continue }
                        # Deduplicate: keep only latest revision per app across all users
                        if ($win32Seen.Contains($appGuid)) { continue }
                        [void]$win32Seen.Add($appGuid)
                        # Save registry path for ALL entries (even skipped) so IME backfill apps can be reset
                        $win32RegMap[$appGuid] = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps\$userId\$rawName"

                        $props = Get-ItemProperty $appKey.PSPath -ErrorAction SilentlyContinue
                        if (-not $props) { continue }

                        $compState = $props.ComplianceStateMessage
                        $resultMsg = $props.EnforcementStateMessage
                        # Skip stale/orphaned entries with no compliance or enforcement data
                        if (-not $compState -and -not $resultMsg) { $win32Skipped++; continue }
                        $stateObj  = $null; $resultObj = $null
                        try { if ($compState) { $stateObj  = $compState | ConvertFrom-Json } } catch { try { Write-DebugLog "Unhandled: $_" -Level ERROR } catch {} }
                        try { if ($resultMsg) { $resultObj = $resultMsg | ConvertFrom-Json } } catch { try { Write-DebugLog "Unhandled: $_" -Level ERROR } catch {} }

                        $installState = switch ($stateObj.ComplianceState) {
                            1 { 'Installed' } 2 { 'Not Installed' } 3 { 'Failed' }
                            4 { 'Not Applicable' } 5 { 'Pending' } default { "Unknown ($($stateObj.ComplianceState))" }
                        }
                        # Skip apps Intune deemed not applicable to this device
                        if ($installState -eq 'Not Applicable') { $win32Skipped++; continue }
                        $enfState = switch ($resultObj.EnforcementState) {
                            1000 { 'Success' } 2000 { 'In Progress' } 3000 { 'Requirements Not Met' }
                            4000 { 'Failed' } 5000 { 'Pending' } default { "$($resultObj.EnforcementState)" }
                        }
                        $errorCode = if ($resultObj.ErrorCode -and $resultObj.ErrorCode -ne 0) { "0x{0:X}" -f $resultObj.ErrorCode } else { '' }
                        $lastAttempt = ''
                        try { if ($stateObj.DesiredState) { $lastAttempt = $stateObj.DesiredState } } catch { try { Write-DebugLog "Unhandled: $_" -Level ERROR } catch {} }

                        $lastMod = Get-RegistryKeyTimestamp $appKey.PSPath

                        # Resolve Win32 app GUID to a friendly name via Uninstall registry
                        $resolved = Resolve-AppDisplayName $appGuid $props
                        $friendlyName = if ($resolved) { $resolved.Name } elseif ($resultObj.TargetingMethod) { "Win32 App ($($resultObj.TargetingMethod))" } else { 'Unknown Win32 App' }
                        $appVersion = if ($resolved) { $resolved.Version } else { '' }

                        $appId++
                        [void]$appsList.Add([PSCustomObject]@{
                            Id            = $appId
                            AppId         = $appGuid
                            AppName       = $friendlyName
                            AppVersion    = $appVersion
                            AppType       = 'Win32/LOB'
                            InstallState  = $installState
                            EnforcementState = $enfState
                            ErrorCode     = $errorCode
                            UserId        = $userId
                            RegistryKey   = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps\$userId\$rawName"
                            LastModified  = $lastMod
                            Source        = 'Registry'
                        })
                    }
                }
                $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] Win32/LOB apps from registry: $($appsList.Count) (scanned=$($win32Seen.Count), skipped $win32Skipped stale/N-A)";Level='DEBUG'})
            }

            # 2) Enrolled apps from EnterpriseDesktopAppManagement (MSI Line-of-Business)
            $msiBiz = 'HKLM:\SOFTWARE\Microsoft\EnterpriseDesktopAppManagement'
            $msiCountBefore = $appsList.Count
            if (Test-Path $msiBiz) {
                $msiKeys = Get-ChildItem $msiBiz -Recurse -ErrorAction SilentlyContinue |
                    Where-Object { $_.PSChildName -match '^\{' }
                foreach ($mk in $msiKeys) {
                    $msiProps = Get-ItemProperty $mk.PSPath -ErrorAction SilentlyContinue

                    $lastMod = ''
                    try { if ($msiProps.CreationTime -and $msiProps.CreationTime -gt 0) { $lastMod = [datetime]::FromFileTime($msiProps.CreationTime).ToString('yyyy-MM-dd HH:mm') } } catch { try { Write-DebugLog "Unhandled: $_" -Level ERROR } catch {} }

                    # Resolve MSI product code to friendly name + version
                    $msiGuid = $mk.PSChildName
                    $resolved = Resolve-AppDisplayName $msiGuid $msiProps
                    $friendlyName = if ($msiProps.Name) { $msiProps.Name }
                                    elseif ($resolved) { $resolved.Name }
                                    else { 'Unknown MSI App' }
                    $appVersion = if ($resolved) { $resolved.Version } elseif ($msiProps.ProductVersion) { $msiProps.ProductVersion } else { '' }
                    $msiInstallState = if ($msiProps.Status -eq 70) { 'Installed' } elseif ($msiProps.Status -eq 60) { 'Failed' } else { "Status: $($msiProps.Status)" }
                    $msiError = if ($msiProps.LastError -and $msiProps.LastError -ne 0) { "0x{0:X}" -f $msiProps.LastError } else { '' }

                    $appId++
                    [void]$appsList.Add([PSCustomObject]@{
                        Id            = $appId
                        AppId         = $msiGuid
                        AppName       = $friendlyName
                        AppVersion    = $appVersion
                        AppType       = 'MSI/LOB'
                        InstallState  = $msiInstallState
                        EnforcementState = $(if ($msiProps.Status -eq 70) { 'Success' } elseif ($msiProps.Status -eq 60) { 'Failed' } elseif ($msiProps.Status) { 'In Progress' } else { '' })
                        ErrorCode     = $msiError
                        UserId        = 'Device'
                        RegistryKey   = ($mk.PSPath -replace 'Microsoft.PowerShell.Core\\Registry::','')
                        LastModified  = $lastMod
                        Source        = 'Registry'
                    })
                }
            }
            $msiCount = $appsList.Count - $msiCountBefore
            if ($msiCount -gt 0) { $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] MSI/LOB apps from registry: $msiCount";Level='DEBUG'}) }

            # 3) Store apps from PackageManagement (UWP/MSIX deployed by Intune)
            $storePath = 'HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\SideCarPolicies\StatusServiceReports'
            $storeCountBefore = $appsList.Count
            if (Test-Path $storePath) {
                $storeKeys = Get-ChildItem $storePath -ErrorAction SilentlyContinue
                foreach ($sk in $storeKeys) {
                    $sProps = Get-ItemProperty $sk.PSPath -ErrorAction SilentlyContinue
                    if (-not $sProps) { continue }
                    $reportJson = $sProps.'(default)'
                    if (-not $reportJson) { $reportJson = ($sProps.PSObject.Properties | Where-Object { $_.Name -notlike 'PS*' } | Select-Object -First 1).Value }
                    $report = $null
                    try { if ($reportJson) { $report = $reportJson | ConvertFrom-Json } } catch { try { Write-DebugLog "Unhandled: $_" -Level ERROR } catch {} }
                    if ($report) {
                        $lastMod = Get-RegistryKeyTimestamp $sk.PSPath

                        $appId++
                        $storeState = if ($report.ComplianceState -eq 1 -or $report.InstallState -eq 1) { 'Installed' }
                                      elseif ($report.ComplianceState -eq 3) { 'Failed' }
                                      else { "State: $($report.ComplianceState)" }
                        $storeName  = if ($report.AppName) { $report.AppName } else { 'Unknown Store App' }
                        $storeVer   = if ($report.AppVersion) { $report.AppVersion } else { '' }
                        [void]$appsList.Add([PSCustomObject]@{
                            Id            = $appId
                            AppId         = $sk.PSChildName
                            AppName       = $storeName
                            AppVersion    = $storeVer
                            AppType       = 'Store/MSIX'
                            InstallState  = $storeState
                            EnforcementState = ''
                            ErrorCode     = ''
                            UserId        = 'Device'
                            RegistryKey   = ($sk.PSPath -replace 'Microsoft.PowerShell.Core\\Registry::','')
                            LastModified  = $lastMod
                            Source        = 'Registry'
                        })
                    }
                }
            }

            # 4) Fill in apps known to IME but missing from registry sources
            $storeCount = $appsList.Count - $storeCountBefore
            if ($storeCount -gt 0) { $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] Store/MSIX apps from registry: $storeCount";Level='DEBUG'}) }
            $regTotal = $appsList.Count
            $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] Registry-discovered apps total: $regTotal";Level='INFO'})
            $knownIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            foreach ($a in $appsList) { [void]$knownIds.Add($a.AppId) }
            $imeAdded = 0
            foreach ($kv in $imeNameMap.GetEnumerator()) {
                if (-not $knownIds.Contains($kv.Key)) {
                    $appId++
                    # Try Uninstall registry for version
                    $ver = ''
                    foreach ($root in @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
                                        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall')) {
                        $rp = Join-Path $root $kv.Key
                        if (Test-Path $rp) { $pp = Get-ItemProperty $rp -ErrorAction SilentlyContinue; if ($pp.DisplayVersion) { $ver = "$($pp.DisplayVersion)"; break } }
                    }
                    # Use ReportingState from log if available, otherwise fall back to defaults
                    $rs = $imeStateMap[$kv.Key]
                    $iState = if ($rs) {
                        switch ($rs.DetectionState) { 1 { 'Installed' } 2 { 'Not Installed' } default { 'Unknown' } }
                    } else { 'Unknown' }
                    $eState = if ($rs) {
                        switch ($rs.EnforcementState) {
                            1000 { 'Success' } 2000 { 'In Progress' } 3000 { 'Requirements Not Met' }
                            4000 { 'Failed' } 5000 { 'Pending' } default { if ($rs.EnforcementState) { "$($rs.EnforcementState)" } else { '' } }
                        }
                    } else { '' }
                    $eErr = if ($rs -and $rs.EnforcementErrorCode) { "0x{0:X}" -f [int]$rs.EnforcementErrorCode } else { '' }
                    $tgtUser = if ($rs -and $rs.TargetingType -eq 1) { 'User' } elseif ($rs -and $rs.TargetingType -eq 2) { 'Device' } else { 'Device' }
                    # Use Win32Apps registry path if we found one during scan (even if skipped as stale/N-A)
                    $w32RegPath = if ($win32RegMap.ContainsKey($kv.Key)) { $win32RegMap[$kv.Key] } else { '' }
                    [void]$appsList.Add([PSCustomObject]@{
                        Id            = $appId
                        AppId         = $kv.Key
                        AppName       = $kv.Value
                        AppVersion    = $ver
                        AppType       = 'Win32/LOB'
                        InstallState  = $iState
                        EnforcementState = $eState
                        ErrorCode     = $eErr
                        UserId        = $tgtUser
                        RegistryKey   = $w32RegPath
                        LastModified  = ''
                        Source        = 'IME Log'
                    })
                    $imeAdded++
                }
            }
            if ($imeAdded -gt 0) {
                $imeWithRegKey = @($appsList | Where-Object { $_.Source -eq 'IME Log' -and $_.RegistryKey }).Count
                $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] IME log backfill: $imeAdded additional apps ($imeWithRegKey with registry key for reset)";Level='INFO'})
            }

            $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] Total managed apps: $($appsList.Count) (registry=$regTotal, IME log=$imeAdded)";Level='INFO'})
        }

        if ($runAD) {
            # â”€â”€ AD Mode via Get-GPO â”€â”€
            $SyncH.StatusQueue.Enqueue(@{Type='Log';Text='[AD] Loading GroupPolicy module, fetching GPOs...';Level='STEP'})
            try {
                Import-Module GroupPolicy -ErrorAction Stop
                $gpoParams = @{}
                if ($DomainOvr) { $gpoParams['Domain'] = $DomainOvr }
                if ($DcOvr)     { $gpoParams['Server'] = $DcOvr }
                if ($OUScope) {
                    # OU-scoped scan: only GPOs that apply to the target OU (linked + inherited)
                    $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[AD] OU-scoped scan: $OUScope";Level='INFO'})
                    $inhParams = @{ Target = $OUScope }
                    if ($DomainOvr) { $inhParams['Domain'] = $DomainOvr }
                    if ($DcOvr)     { $inhParams['Server'] = $DcOvr }
                    $inh = Get-GPInheritance @inhParams -ErrorAction Stop
                    $ouGpoLinks = @($inh.InheritedGpoLinks)
                    $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[AD] Get-GPInheritance returned $($ouGpoLinks.Count) GPO links for OU";Level='INFO'})
                    # Batch-resolve via LDAP instead of N individual Get-GPO calls
                    $ouGuids = @($ouGpoLinks | ForEach-Object { $_.GpoId.ToString().Trim('{','}') } | Where-Object { $_ })
                    $SyncH.StatusQueue.Enqueue(@{Type='Status';Text="Resolving $($ouGuids.Count) GPOs via LDAP...";Color='#F59E0B'})
                    $domain = if ($DomainOvr) { $DomainOvr }
                              else { try { [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name } catch { '(unknown)' } }
                    $ldapRoot = if ($DcOvr -and $DomainOvr) { "LDAP://$DcOvr/CN=Policies,CN=System,$(($DomainOvr.Split('.') | ForEach-Object { "DC=$_" }) -join ',')" }
                                elseif ($DomainOvr) { "LDAP://CN=Policies,CN=System,$(($DomainOvr.Split('.') | ForEach-Object { "DC=$_" }) -join ',')" }
                                elseif ($DcOvr) { $rdse = [ADSI]"LDAP://$DcOvr/RootDSE"; "LDAP://$DcOvr/CN=Policies,CN=System,$($rdse.defaultNamingContext)" }
                                else { $rdse = [ADSI]'LDAP://RootDSE'; "LDAP://CN=Policies,CN=System,$($rdse.defaultNamingContext)" }
                    $ouSearchRoot = [ADSI]$ldapRoot
                    $ouSearcher = [System.DirectoryServices.DirectorySearcher]::new($ouSearchRoot)
                    # Build OR filter for all GUIDs: (|(name={guid1})(name={guid2})...)
                    $guidFilter = ($ouGuids | ForEach-Object { "(name={$_})" }) -join ''
                    $ouSearcher.Filter = "(&(objectClass=groupPolicyContainer)(|$guidFilter))"
                    $ouSearcher.PageSize = 100
                    @('displayName','name','flags','versionNumber','whenCreated','whenChanged','gPCWMIFilter') | ForEach-Object { [void]$ouSearcher.PropertiesToLoad.Add($_) }
                    $ouLdapResults = $ouSearcher.FindAll()
                    $rawGPOs = [System.Collections.Generic.List[object]]::new()
                    $resolvedGuids = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                    foreach ($entry in $ouLdapResults) {
                        $props = $entry.Properties
                        $displayName = if ($props['displayname'] -and $props['displayname'].Count -gt 0) { "$($props['displayname'][0])" } else { '' }
                        $guidRaw = if ($props['name'] -and $props['name'].Count -gt 0) { "$($props['name'][0])" } else { '' }
                        $guid = $guidRaw.Trim('{','}')
                        $flags = if ($props['flags'] -and $props['flags'].Count -gt 0) { [int]$props['flags'][0] } else { 0 }
                        $ver = if ($props['versionNumber'] -and $props['versionNumber'].Count -gt 0) { [int]$props['versionNumber'][0] } else { 0 }
                        $whenCreated = if ($props['whencreated'] -and $props['whencreated'].Count -gt 0) { [datetime]$props['whencreated'][0] } else { [datetime]::MinValue }
                        $whenChanged = if ($props['whenchanged'] -and $props['whenchanged'].Count -gt 0) { [datetime]$props['whenchanged'][0] } else { [datetime]::MinValue }
                        $userVer = ($ver -shr 16) -band 0xFFFF
                        $compVer = $ver -band 0xFFFF
                        $gpoStatus = switch ($flags) { 1 { 'UserSettingsDisabled' } 2 { 'ComputerSettingsDisabled' } 3 { 'AllSettingsDisabled' } default { 'AllSettingsEnabled' } }
                        $wmiFilterName = ''
                        $wmiRef = if ($props['gpcwmifilter'] -and $props['gpcwmifilter'].Count -gt 0) { "$($props['gpcwmifilter'][0])" } else { $null }
                        if ($wmiRef -and $wmiRef -match ';') { $wmiFilterName = ($wmiRef -split ';')[1] }
                        if ($displayName -and $guid) {
                            [void]$rawGPOs.Add([PSCustomObject]@{
                                DisplayName      = $displayName
                                Id               = [Guid]$guid
                                GpoStatus        = $gpoStatus
                                CreationTime     = $whenCreated
                                ModificationTime = $whenChanged
                                WmiFilter        = if ($wmiFilterName) { [PSCustomObject]@{ Name = $wmiFilterName } } else { $null }
                                User             = [PSCustomObject]@{ DSVersion = $userVer }
                                Computer         = [PSCustomObject]@{ DSVersion = $compVer }
                            })
                            [void]$resolvedGuids.Add($guid)
                        }
                    }
                    $ouLdapResults.Dispose()
                    # Log orphaned GPO links (link exists but GPO object gone)
                    $orphaned = @($ouGuids | Where-Object { -not $resolvedGuids.Contains($_) })
                    if ($orphaned.Count -gt 0) {
                        $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[AD] $($orphaned.Count) orphaned GPO link(s) skipped (GPO deleted but link remains): $($orphaned[0..([math]::Min(4,$orphaned.Count-1)) -join ', '])";Level='WARN'})
                    }
                    $rawGPOs = @($rawGPOs)
                    $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[AD] OU LDAP resolve: $($rawGPOs.Count) GPOs resolved from $($ouGuids.Count) links";Level='SUCCESS'})
                } else {
                    # LDAP-based enumeration with progress instead of blocking Get-GPO -All
                    $SyncH.StatusQueue.Enqueue(@{Type='Status';Text='Discovering GPOs via LDAP...';Color='#F59E0B'})
                    $domain = if ($DomainOvr) { $DomainOvr }
                              else { try { [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name } catch { '(unknown)' } }
                    $ldapRoot = if ($DcOvr -and $DomainOvr) { "LDAP://$DcOvr/CN=Policies,CN=System,$(($DomainOvr.Split('.') | ForEach-Object { "DC=$_" }) -join ',')" }
                                elseif ($DomainOvr) { "LDAP://CN=Policies,CN=System,$(($DomainOvr.Split('.') | ForEach-Object { "DC=$_" }) -join ',')" }
                                elseif ($DcOvr) { $rootDse = [ADSI]"LDAP://$DcOvr/RootDSE"; "LDAP://$DcOvr/CN=Policies,CN=System,$($rootDse.defaultNamingContext)" }
                                else { $rootDse = [ADSI]'LDAP://RootDSE'; "LDAP://CN=Policies,CN=System,$($rootDse.defaultNamingContext)" }
                    $searchRoot = [ADSI]$ldapRoot
                    $searcher = [System.DirectoryServices.DirectorySearcher]::new($searchRoot)
                    $searcher.Filter = '(objectClass=groupPolicyContainer)'
                    $searcher.PageSize = 100
                    @('displayName','name','flags','versionNumber','whenCreated','whenChanged','gPCWMIFilter') | ForEach-Object { [void]$searcher.PropertiesToLoad.Add($_) }
                    $ldapResults = $searcher.FindAll()
                    $rawGPOs = [System.Collections.Generic.List[object]]::new()
                    $ldapCount = 0
                    foreach ($entry in $ldapResults) {
                        $ldapCount++
                        $props = $entry.Properties
                        $dn = $props['displayname']
                        $displayName = if ($dn -and $dn.Count -gt 0) { "$($dn[0])" } else { '' }
                        $guidRaw = if ($props['name'] -and $props['name'].Count -gt 0) { "$($props['name'][0])" } else { '' }
                        $guid = $guidRaw.Trim('{','}')
                        $flags = if ($props['flags'] -and $props['flags'].Count -gt 0) { [int]$props['flags'][0] } else { 0 }
                        $ver = if ($props['versionNumber'] -and $props['versionNumber'].Count -gt 0) { [int]$props['versionNumber'][0] } else { 0 }
                        $whenCreated = if ($props['whencreated'] -and $props['whencreated'].Count -gt 0) { [datetime]$props['whencreated'][0] } else { [datetime]::MinValue }
                        $whenChanged = if ($props['whenchanged'] -and $props['whenchanged'].Count -gt 0) { [datetime]$props['whenchanged'][0] } else { [datetime]::MinValue }
                        # versionNumber: upper 16 bits = user, lower 16 bits = computer
                        $userVer = ($ver -shr 16) -band 0xFFFF
                        $compVer = $ver -band 0xFFFF
                        $gpoStatus = switch ($flags) {
                            1 { 'UserSettingsDisabled' }
                            2 { 'ComputerSettingsDisabled' }
                            3 { 'AllSettingsDisabled' }
                            default { 'AllSettingsEnabled' }
                        }
                        # WMI filter: extract display name from DN reference if present
                        $wmiFilterName = ''
                        $wmiRef = if ($props['gpcwmifilter'] -and $props['gpcwmifilter'].Count -gt 0) { "$($props['gpcwmifilter'][0])" } else { $null }
                        if ($wmiRef -and $wmiRef -match ';') { $wmiFilterName = ($wmiRef -split ';')[1] }
                        if (-not $displayName -or -not $guid) { continue }
                        [void]$rawGPOs.Add([PSCustomObject]@{
                            DisplayName      = $displayName
                            Id               = [Guid]$guid
                            GpoStatus        = $gpoStatus
                            CreationTime     = $whenCreated
                            ModificationTime = $whenChanged
                            WmiFilter        = if ($wmiFilterName) { [PSCustomObject]@{ Name = $wmiFilterName } } else { $null }
                            User             = [PSCustomObject]@{ DSVersion = $userVer }
                            Computer         = [PSCustomObject]@{ DSVersion = $compVer }
                        })
                        if ($ldapCount % 100 -eq 0) {
                            $SyncH.StatusQueue.Enqueue(@{Type='Status';Text="Discovered $ldapCount GPOs...";Color='#F59E0B'})
                            $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[AD] LDAP discovery: $ldapCount GPOs found so far...";Level='DEBUG'})
                        }
                    }
                    $ldapResults.Dispose()
                    $rawGPOs = @($rawGPOs)
                    $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[AD] LDAP discovery complete: $($rawGPOs.Count) GPOs in ${domain}";Level='SUCCESS'})
                }
            } catch {
                return @{ Error = "Get-GPO failed: $($_.Exception.Message)" }
            }

            $domain = if ($DomainOvr) { $DomainOvr }
                      else { try { [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name } catch { '(unknown)' } }

            # Auto-discover a single DC for cross-domain scans so every Get-GPOReport/Get-GPPermission
            # call reuses the same connection instead of doing DC location per call (huge perf win)
            if ($DomainOvr -and -not $DcOvr) {
                try {
                    $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[AD] Auto-discovering DC for $DomainOvr...";Level='INFO'})
                    $ctx = [System.DirectoryServices.ActiveDirectory.DirectoryContext]::new('Domain', $DomainOvr)
                    $targetDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($ctx)
                    $dc = $targetDomain.FindDomainController()
                    $DcOvr = $dc.Name
                    $gpoParams['Server'] = $DcOvr
                    $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[AD] Using DC: $DcOvr for all cross-domain calls";Level='SUCCESS'})
                } catch {
                    $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[AD] DC auto-discovery failed: $($_.Exception.Message) — will use default DC location";Level='WARN'})
                }
            }

            # Smart filter: skip GPOs with no settings (empty DSVersion=0 or AllSettingsDisabled)
            $linkedGPOs = @($rawGPOs | Where-Object {
                ($_.User.DSVersion -gt 0 -or $_.Computer.DSVersion -gt 0) -and
                ($_.GpoStatus -ne 'AllSettingsDisabled')
            })
            $skipped = $rawGPOs.Count - $linkedGPOs.Count

            # ── Batch LDAP: discover GPO link locations (which OUs/Sites/Domains link to each GPO) ──
            $gpoLinkMap = @{}
            try {
                $SyncH.StatusQueue.Enqueue(@{Type='Status';Text='Discovering GPO link locations via LDAP...';Color='#F59E0B'})
                $baseDN = if ($DomainOvr) { ($DomainOvr.Split('.') | ForEach-Object { "DC=$_" }) -join ',' }
                          else { $rootDse2 = [ADSI]'LDAP://RootDSE'; "$($rootDse2.defaultNamingContext)" }
                $linkLdap = if ($DcOvr) { "LDAP://$DcOvr/$baseDN" } else { "LDAP://$baseDN" }
                $linkSearcher = [System.DirectoryServices.DirectorySearcher]::new([ADSI]$linkLdap)
                $linkSearcher.Filter = '(gPLink=*)'
                $linkSearcher.PageSize = 200
                @('distinguishedName','gPLink','gPOptions') | ForEach-Object { [void]$linkSearcher.PropertiesToLoad.Add($_) }
                $linkResults = $linkSearcher.FindAll()
                foreach ($lr in $linkResults) {
                    $dn = "$($lr.Properties['distinguishedname'][0])"
                    $gpLink = "$($lr.Properties['gplink'][0])"
                    # gPLink format: [LDAP://cn={GUID},cn=policies,cn=system,DC=...;0][LDAP://...;2]
                    foreach ($match in [regex]::Matches($gpLink, '\[LDAP://[Cc][Nn]=\{([0-9a-fA-F-]+)\}[^;]*;(\d+)\]')) {
                        $linkGuid = $match.Groups[1].Value.ToUpper()
                        $linkFlags = [int]$match.Groups[2].Value
                        $enforced = ($linkFlags -band 2) -ne 0
                        $disabled = ($linkFlags -band 1) -ne 0
                        if (-not $disabled) {
                            if (-not $gpoLinkMap.ContainsKey($linkGuid)) { $gpoLinkMap[$linkGuid] = [System.Collections.Generic.List[PSCustomObject]]::new() }
                            [void]$gpoLinkMap[$linkGuid].Add([PSCustomObject]@{ SOMPath = $dn; Enforced = $enforced })
                        }
                    }
                }
                $linkResults.Dispose()
                $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[AD] Link discovery: $($gpoLinkMap.Count) GPOs have active links";Level='SUCCESS'})
            } catch {
                $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[AD] Link discovery failed: $($_.Exception.Message) — links will be empty";Level='WARN'})
            }
            if ($skipped -gt 0) {
                $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[AD] Skipping $skipped empty/disabled GPOs";Level='INFO'})
            }
            $total = $linkedGPOs.Count
            $idx = 0
            $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[AD] Found $($rawGPOs.Count) GPOs ($total with settings), fetching reports...";Level='INFO'})

            $gpoXmlDocs = @{}

            $idx = 0
            # ── Per-GPO XML disk cache ──
            # Cache fetched XML reports to $env:TEMP\PolicyPilot_GPOCache\<domain>\ so a
            # crash during a long scan doesn't lose all progress. Cache entries expire after 4 hours.
            $cacheDomain = if ($DomainOvr) { $DomainOvr -replace '[\\/:*?"<>|]','_' } else { 'local' }
            $cacheDir = [IO.Path]::Combine($env:TEMP, 'PolicyPilot_GPOCache', $cacheDomain)
            if (-not (Test-Path $cacheDir)) { [void][IO.Directory]::CreateDirectory($cacheDir) }
            $cacheMaxAge = [TimeSpan]::FromHours(4)
            $cacheHits = 0; $cacheMisses = 0
            if ($ForceRefresh) {
                $SyncH.StatusQueue.Enqueue(@{Type='Log';Text='[AD] Force refresh enabled — disk cache will be bypassed';Level='INFO'})
            }
            $loopStart = [DateTime]::Now

            foreach ($gpo in $linkedGPOs) {
                $idx++
                $pctDone = [math]::Round($idx / $total * 100)
                $SyncH.StatusQueue.Enqueue(@{Type='Status';Text="Processing GPO $idx/$total - $($gpo.DisplayName)";Color='#F59E0B'})
                $SyncH.StatusQueue.Enqueue(@{Type='Progress';Value=$pctDone})
                # Log progress every 25 GPOs or on the first one
                if ($idx -eq 1 -or $idx % 25 -eq 0 -or $idx -eq $total) {
                    $elapsed = ([DateTime]::Now - $loopStart).TotalSeconds
                    $eta = if ($idx -gt 1) { [math]::Round(($elapsed / ($idx - 1)) * ($total - $idx)) } else { '?' }
                    $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[AD] GPO $idx/$total ($pctDone%) — cache:$cacheHits fetch:$cacheMisses — ETA ${eta}s";Level='INFO'})
                }
                # Use pre-fetched XML from bulk report, fall back to disk cache, then per-GPO fetch
                $guid = $gpo.Id.ToString().Trim('{','}').ToUpper()
                $xmlText = $gpoXmlDocs[$guid]
                $cacheFile = [IO.Path]::Combine($cacheDir, "$guid.xml")
                if (-not $xmlText -and -not $ForceRefresh) {
                    # Check disk cache — skip if GPO was modified after the cached file
                    if ([IO.File]::Exists($cacheFile)) {
                        $cacheAge = [DateTime]::Now - [IO.File]::GetLastWriteTime($cacheFile)
                        $gpoModified = $gpo.ModificationTime
                        $cacheWritten = [IO.File]::GetLastWriteTime($cacheFile)
                        if ($cacheAge -lt $cacheMaxAge -and (!$gpoModified -or $gpoModified -lt $cacheWritten)) {
                            $xmlText = [IO.File]::ReadAllText($cacheFile, [System.Text.Encoding]::UTF8)
                            $cacheHits++
                        }
                    }
                }
                if (-not $xmlText) {
                    # Try SYSVOL direct read first (fast SMB), fall back to Get-GPOReport (slow cmdlet)
                    $sysvolBase = "\\\\$domain\\SYSVOL\\$domain\\Policies\\{$($gpo.Id.ToString())}".ToUpper()
                    $fetchSw = [System.Diagnostics.Stopwatch]::StartNew()
                    $fetchMethod = 'SYSVOL'
                    try {
                        # Build a minimal GPO XML from SYSVOL .pol files + gPCFileSysPath
                        $polXml = [System.Text.StringBuilder]::new()
                        [void]$polXml.AppendLine('<?xml version="1.0" encoding="utf-8"?>')
                        [void]$polXml.AppendLine('<GPO xmlns="http://www.microsoft.com/GroupPolicy/Settings" xmlns:types="http://www.microsoft.com/GroupPolicy/Types">')
                        [void]$polXml.AppendLine("<Identifier><Identifier>{$($gpo.Id.ToString())}</Identifier></Identifier>")
                        [void]$polXml.AppendLine("<Name>$([System.Security.SecurityElement]::Escape($gpo.DisplayName))</Name>")
                        # Parse Machine and User registry.pol
                        foreach ($scope in @('Computer','User')) {
                            $polPath = [IO.Path]::Combine($sysvolBase, $(if ($scope -eq 'Computer') {'Machine'} else {'User'}), 'registry.pol')
                            if (-not [IO.File]::Exists($polPath)) { continue }
                            [void]$polXml.AppendLine("<$scope>")
                            [void]$polXml.AppendLine('<ExtensionData><Extension xmlns:q1="http://www.microsoft.com/GroupPolicy/Settings/Registry" xsi:type="q1:RegistrySettings" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">')
                            try {
                                $polBytes = [IO.File]::ReadAllBytes($polPath)
                                # registry.pol format: header (8 bytes: PReg\x01\x00\x00\x00), then entries
                                # Each entry: [key;valueName;type(4LE);size(4LE);data(size)]  with UTF-16LE strings
                                if ($polBytes.Length -gt 8) {
                                    $pos = 8 # skip header
                                    while ($pos -lt $polBytes.Length - 2) {
                                        # Find opening bracket [
                                        if ($polBytes[$pos] -ne 0x5B -or $polBytes[$pos+1] -ne 0x00) { $pos++; continue }
                                        $pos += 2
                                        # Read key (null-terminated UTF-16LE)
                                        $keyStart = $pos
                                        while ($pos -lt $polBytes.Length - 3 -and -not ($polBytes[$pos] -eq 0 -and $polBytes[$pos+1] -eq 0 -and $polBytes[$pos+2] -eq 0x3B -and $polBytes[$pos+3] -eq 0)) { $pos += 2 }
                                        $regKey = [System.Text.Encoding]::Unicode.GetString($polBytes, $keyStart, $pos - $keyStart)
                                        $pos += 4 # skip null + semicolon
                                        # Read value name
                                        $vnStart = $pos
                                        while ($pos -lt $polBytes.Length - 3 -and -not ($polBytes[$pos] -eq 0 -and $polBytes[$pos+1] -eq 0 -and $polBytes[$pos+2] -eq 0x3B -and $polBytes[$pos+3] -eq 0)) { $pos += 2 }
                                        $valName = [System.Text.Encoding]::Unicode.GetString($polBytes, $vnStart, $pos - $vnStart)
                                        $pos += 4 # skip null + semicolon
                                        # Read type (4 bytes LE)
                                        if ($pos + 4 -gt $polBytes.Length) { break }
                                        $regType = [BitConverter]::ToUInt32($polBytes, $pos); $pos += 4
                                        $pos += 2 # semicolon
                                        # Read size (4 bytes LE)
                                        if ($pos + 4 -gt $polBytes.Length) { break }
                                        $dataSize = [BitConverter]::ToUInt32($polBytes, $pos); $pos += 4
                                        # Read data
                                        $dataVal = ''
                                        if ($dataSize -gt 0 -and $pos + $dataSize -le $polBytes.Length) {
                                            switch ($regType) {
                                                1 { $dataVal = [System.Text.Encoding]::Unicode.GetString($polBytes, $pos, $dataSize).TrimEnd("`0") } # REG_SZ
                                                4 { if ($dataSize -ge 4) { $dataVal = [BitConverter]::ToUInt32($polBytes, $pos).ToString() } }       # REG_DWORD
                                                default { $dataVal = [BitConverter]::ToString($polBytes, $pos, [math]::Min($dataSize, 64)) }
                                            }
                                            $pos += $dataSize
                                        }
                                        # Skip closing bracket ]
                                        if ($pos -lt $polBytes.Length - 1 -and $polBytes[$pos] -eq 0x5D -and $polBytes[$pos+1] -eq 0) { $pos += 2 }
                                        $typeStr = switch ($regType) { 1 {'REG_SZ'} 2 {'REG_EXPAND_SZ'} 4 {'REG_DWORD'} 7 {'REG_MULTI_SZ'} 11 {'REG_QWORD'} 3 {'REG_BINARY'} default {"Type_$regType"} }
                                        [void]$polXml.AppendLine('<q1:RegistrySetting>')
                                        [void]$polXml.AppendLine("<q1:KeyPath>$([System.Security.SecurityElement]::Escape($regKey))</q1:KeyPath>")
                                        [void]$polXml.AppendLine("<q1:ValueName>$([System.Security.SecurityElement]::Escape($valName))</q1:ValueName>")
                                        [void]$polXml.AppendLine("<q1:Value>$([System.Security.SecurityElement]::Escape($dataVal))</q1:Value>")
                                        [void]$polXml.AppendLine("<q1:Type>$typeStr</q1:Type>")
                                        [void]$polXml.AppendLine('</q1:RegistrySetting>')
                                    }
                                }
                            } catch {
                                $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[AD] .pol parse error for $scope in $($gpo.DisplayName): $($_.Exception.Message)";Level='WARN'})
                            }
                            [void]$polXml.AppendLine('</Extension></ExtensionData>')
                            [void]$polXml.AppendLine("</$scope>")
                        }
                        # Read link info from gPLink attribute on the GPO's LDAP entry
                        [void]$polXml.AppendLine('</GPO>')
                        $xmlText = $polXml.ToString()
                    } catch {
                        # SYSVOL read failed — fall back to Get-GPOReport
                        $fetchMethod = 'GPOReport'
                        try {
                            $reportParams = @{ Guid = $gpo.Id; ReportType = 'Xml' }
                            if ($DomainOvr) { $reportParams['Domain'] = $DomainOvr }
                            if ($DcOvr)     { $reportParams['Server'] = $DcOvr }
                            $xmlText = Get-GPOReport @reportParams -ErrorAction Stop
                        } catch {
                            $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[AD] Both SYSVOL and GPOReport failed for $($gpo.DisplayName): $($_.Exception.Message)";Level='WARN'})
                        }
                    }
                    $fetchSw.Stop()
                    if ($xmlText) {
                        # Save to disk cache
                        try { [IO.File]::WriteAllText($cacheFile, $xmlText, [System.Text.Encoding]::UTF8) } catch {}
                        $cacheMisses++
                        if ($fetchSw.Elapsed.TotalSeconds -gt 5) {
                            $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[AD] Slow fetch ($fetchMethod): $($gpo.DisplayName) took $([math]::Round($fetchSw.Elapsed.TotalSeconds,1))s";Level='WARN'})
                        }
                    }
                }
                $linkLocations = [System.Collections.Generic.List[string]]::new()
                $linkDetails   = [System.Collections.Generic.List[PSCustomObject]]::new()
                $settingCount = 0
                $isEnforced   = $false

                # Populate link data from pre-fetched LDAP gPLink map
                $gpoGuidUpper = $gpo.Id.ToString().Trim('{','}').ToUpper()
                if ($gpoLinkMap.ContainsKey($gpoGuidUpper)) {
                    foreach ($lnk in $gpoLinkMap[$gpoGuidUpper]) {
                        [void]$linkLocations.Add($lnk.SOMPath)
                        if ($lnk.Enforced) { $isEnforced = $true }
                        [void]$linkDetails.Add([PSCustomObject]@{
                            SOMPath   = $lnk.SOMPath
                            Enforced  = $lnk.Enforced
                            LinkOrder = 0
                        })
                    }
                }

                if ($xmlText) {
                    try {
                        $xdoc = [xml]$xmlText
                        $nsMgr2 = [System.Xml.XmlNamespaceManager]::new($xdoc.NameTable)
                        $nsMgr2.AddNamespace('gp', 'http://www.microsoft.com/GroupPolicy/Settings')
                        $nsMgr2.AddNamespace('q1', 'http://www.microsoft.com/GroupPolicy/Settings/Registry')

                        # Parse full LinksTo elements for Enforced + LinkOrder
                        $linksToNodes = $xdoc.SelectNodes('//gp:LinksTo', $nsMgr2)
                        foreach ($lt in $linksToNodes) {
                            $somPath    = $lt.SelectSingleNode('gp:SOMPath', $nsMgr2)
                            $noOverride = $lt.SelectSingleNode('gp:NoOverride', $nsMgr2)
                            $somOrder   = $lt.SelectSingleNode('gp:SOMOrder', $nsMgr2)
                            $path = if ($somPath) { $somPath.InnerText } else { '' }
                            if ($path) { [void]$linkLocations.Add($path) }
                            $enforced = ($noOverride -and $noOverride.InnerText -eq 'true')
                            if ($enforced) { $isEnforced = $true }
                            [void]$linkDetails.Add([PSCustomObject]@{
                                SOMPath   = $path
                                Enforced  = $enforced
                                LinkOrder = if ($somOrder) { [int]$somOrder.InnerText } else { 0 }
                            })
                        }

                        foreach ($scope in @('Computer','User')) {
                            $scopeNode = $xdoc.SelectSingleNode("//gp:$scope", $nsMgr2)
                            if (-not $scopeNode) { continue }
                            $enabledNode = $scopeNode.SelectSingleNode('gp:Enabled', $nsMgr2)
                            if ($enabledNode -and $enabledNode.InnerText -eq 'false') { continue }

                            $extNodes = $scopeNode.SelectNodes('gp:ExtensionData/gp:Extension', $nsMgr2)
                            foreach ($ext in $extNodes) {
                                $policies = $ext.SelectNodes('q1:Policy', $nsMgr2)
                                if (-not $policies -or $policies.Count -eq 0) {
                                    $policies = $ext.ChildNodes | Where-Object { $_.LocalName -eq 'Policy' }
                                }
                                foreach ($pol in $policies) {
                                    $settingCount++
                                    $polName = $pol.SelectSingleNode('*[local-name()="Name"]')
                                    $polState = $pol.SelectSingleNode('*[local-name()="State"]')
                                    $polCategory = $pol.SelectSingleNode('*[local-name()="Category"]')
                                    $polExplain = $pol.SelectSingleNode('*[local-name()="Explain"]')
                                    $subValues = [System.Collections.Generic.List[string]]::new()
                                    foreach ($child in $pol.ChildNodes) {
                                        if ($child.LocalName -in @('Name','State','Category','Supported','Explain')) { continue }
                                        $subName  = $child.SelectSingleNode('*[local-name()="Name"]')
                                        $subValue = $child.SelectSingleNode('*[local-name()="Value"]')
                                        $subState = $child.SelectSingleNode('*[local-name()="State"]')
                                        if ($subName -and ($subValue -or $subState)) {
                                            $val = if ($subValue) {
                                                $innerName = $subValue.SelectSingleNode('*[local-name()="Name"]')
                                                if ($innerName) { $innerName.InnerText } else { $subValue.InnerText }
                                            } elseif ($subState) { $subState.InnerText } else { '' }
                                            [void]$subValues.Add("$($subName.InnerText)=$val")
                                        }
                                    }
                                    $settingKey = "$scope|$(if ($polCategory) {$polCategory.InnerText} else {'(unknown)'})\$(if ($polName) {$polName.InnerText} else {'(unnamed)'})"
                                    $stateVal   = if ($polState) { $polState.InnerText } else { 'Unknown' }
                                    [void]$allSettingsList.Add([PSCustomObject]@{
                                        SettingKey   = $settingKey
                                        PolicyName   = if ($polName) { $polName.InnerText } else { '(unnamed)' }
                                        State        = $stateVal
                                        Scope        = $scope
                                        Category     = if ($polCategory) { $polCategory.InnerText } else { '' }
                                        GPOName      = $gpo.DisplayName
                                        GPOGuid      = $gpo.Id.ToString()
                                        RegistryKey  = ''
                                        ValueData    = ($subValues -join '; ')
                                        Explain      = if ($polExplain) { $polExplain.InnerText } else { '' }
                                        Source       = 'Domain GPO'
                                        IntuneGroup  = 'Group Policy'
                                    })
                                }
                                $regSettings = $ext.ChildNodes | Where-Object { $_.LocalName -eq 'RegistrySetting' }
                                foreach ($reg in $regSettings) {
                                    $settingCount++
                                    $regHive = $reg.SelectSingleNode('*[local-name()="KeyPath"]')
                                    $regName = $reg.SelectSingleNode('*[local-name()="ValueName"]')
                                    $regData = $reg.SelectSingleNode('*[local-name()="Value"]')
                                    $regType = $reg.SelectSingleNode('*[local-name()="Type"]')
                                    $keyPath2 = if ($regHive) { $regHive.InnerText } else { '' }
                                    $valName2 = if ($regName) { $regName.InnerText } else { '(Default)' }
                                    $settingKey2 = "$scope|$keyPath2\$valName2"
                                    [void]$allSettingsList.Add([PSCustomObject]@{
                                        SettingKey   = $settingKey2
                                        PolicyName   = "$keyPath2\$valName2"
                                        State        = 'Registry'
                                        Scope        = $scope
                                        Category     = 'Registry Settings'
                                        GPOName      = $gpo.DisplayName
                                        GPOGuid      = $gpo.Id.ToString()
                                        RegistryKey  = $keyPath2
                                        ValueData    = if ($regData) { $regData.InnerText } else { '' }
                                        Explain      = if ($regType) { "Type: $($regType.InnerText)" } else { '' }
                                        Source       = 'Domain GPO'
                                        IntuneGroup  = 'Registry Settings'
                                    })
                                }
                            }
                        }
                    } catch { $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[AD] GPO processing failed: $($_.Exception.Message)";Level='WARN'}) }
                }

                $statusText = switch ($gpo.GpoStatus.ToString()) {
                    'AllSettingsEnabled'       { 'Enabled' }
                    'AllSettingsDisabled'      { 'Disabled' }
                    'UserSettingsDisabled'     { 'Computer Only' }
                    'ComputerSettingsDisabled' { 'User Only' }
                    default                    { $gpo.GpoStatus.ToString() }
                }
                # H3: Security group filtering via LDAP ACL (replaces slow Get-GPPermission)
                $secFilter = [System.Collections.Generic.List[string]]::new()
                try {
                    $gpoDN = "CN={$($gpo.Id.ToString())},CN=Policies,CN=System,$(($domain.Split('.') | ForEach-Object { "DC=$_" }) -join ',')"
                    $gpoLdap = if ($DcOvr) { "LDAP://$DcOvr/$gpoDN" } else { "LDAP://$gpoDN" }
                    $gpoEntry = [ADSI]$gpoLdap
                    $sd = $gpoEntry.ObjectSecurity
                    # ApplyGroupPolicy extended right GUID
                    $applyGpGuid = [Guid]'edacfd8f-ffb3-11d1-b41d-00a0c968f939'
                    foreach ($ace in $sd.GetAccessRules($true, $false, [System.Security.Principal.NTAccount])) {
                        if ($ace.AccessControlType -eq 'Allow' -and $ace.ObjectType -eq $applyGpGuid) {
                            [void]$secFilter.Add("$($ace.IdentityReference)")
                        }
                    }
                } catch { }
                $secFilterStr = if ($secFilter.Count -gt 0) { $secFilter -join '; ' } else { 'Authenticated Users' }

                $linkOrder = if ($linkDetails.Count -gt 0) { ($linkDetails | ForEach-Object { $_.LinkOrder } | Sort-Object | Select-Object -First 1) } else { 0 }

                [void]$gpoRecords.Add([PSCustomObject]@{
                    DisplayName      = $gpo.DisplayName
                    Id               = $gpo.Id.ToString()
                    Status           = $statusText
                    GpoStatus        = $gpo.GpoStatus.ToString()
                    SettingCount     = $settingCount
                    LinkCount        = $linkLocations.Count
                    Links            = ($linkLocations -join '; ')
                    CreationTime     = $gpo.CreationTime.ToString('yyyy-MM-dd HH:mm')
                    ModificationTime = $gpo.ModificationTime.ToString('yyyy-MM-dd HH:mm')
                    WmiFilter        = if ($gpo.WmiFilter) { $gpo.WmiFilter.Name } else { '' }
                    IsLinked         = ($linkLocations.Count -gt 0)
                    Enforced         = $isEnforced
                    LinkOrder        = $linkOrder
                    SecurityFiltering = $secFilterStr
                    LinkDetails      = $linkDetails
                })
            }
        }

        # Log GPO XML cache stats
        if ($cacheHits -gt 0 -or $cacheMisses -gt 0) {
            $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[AD] GPO XML cache: $cacheHits hits, $cacheMisses fetched (cache: $cacheDir)";Level='INFO'})
        }

        # ── Set domain label based on scan mode ──
        if ($ScanMode -eq 'Intune') {
            $domain = 'Intune (Local MDM)'
        } elseif ($ScanMode -eq 'Combined') {
            $domain = if ($domain -ne 'LocalMachine') { "$domain (Co-managed)" } else { 'Co-managed (Local + Intune)' }
        }

        # ── Inject provisioning package settings into the settings list ──
        # Each ppkg XMLEntry becomes a setting row so it participates in filtering + conflict detection
        $ppkgSettings = 0
        if ($mdmInfo.MdmDiag.ProvisioningPackages) {
            foreach ($ppkg in $mdmInfo.MdmDiag.ProvisioningPackages) {
                $pkgLabel = if ($ppkg.PackageName) { $ppkg.PackageName } elseif ($ppkg.FriendlyName) { $ppkg.FriendlyName } else { $ppkg.FileName }
                if ($ppkg.XMLEntries -and $ppkg.XMLEntries.Count -gt 0) {
                    foreach ($xe in $ppkg.XMLEntries) {
                        $settIdCounter++
                        $xeState = if ([int]$xe.NumberOfFailures -gt 0) { 'Failed' } else { 'Applied' }
                        [void]$allSettingsList.Add([PSCustomObject]@{
                            Id=$settIdCounter; GPOName=$pkgLabel; GPOGuid=$ppkg.PackageId
                            Category="Provisioning Package"; PolicyName=$xe.XMLName
                            SettingKey="ppkg/$($ppkg.PackageId)/$($xe.XMLName)"; State=$xeState
                            RegistryKey=$xe.Area; ValueData=$xe.Message
                            Scope='Device'; Source='Provisioning Package'; IntuneGroup='Provisioning Packages'
                        })
                        $ppkgSettings++
                    }
                } else {
                    # Package with no XMLEntries — add a summary row so it's visible
                    $settIdCounter++
                    [void]$allSettingsList.Add([PSCustomObject]@{
                        Id=$settIdCounter; GPOName=$pkgLabel; GPOGuid=$ppkg.PackageId
                        Category="Provisioning Package"; PolicyName="$pkgLabel (package)"
                        SettingKey="ppkg/$($ppkg.PackageId)"; State=if ($ppkg.TotalFailures -gt 0) { 'HasFailures' } else { 'Applied' }
                        RegistryKey=''; ValueData="Owner=$($ppkg.Owner), Version=$($ppkg.Version)"
                        Scope='Device'; Source='Provisioning Package'; IntuneGroup='Provisioning Packages'
                    })
                    $ppkgSettings++
                }
            }
            if ($ppkgSettings -gt 0) {
                $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] Injected $ppkgSettings provisioning package settings into settings list";Level='INFO'})
            }
        }

        # ── N1+N5: Compute app install summary & compliance health ──
        if ($enrolled -and $appsList.Count -gt 0) {
            $instCnt  = @($appsList | Where-Object { $_.InstallState -eq 'Installed' }).Count
            $failCnt  = @($appsList | Where-Object { $_.InstallState -eq 'Failed' }).Count
            $pendCnt  = @($appsList | Where-Object { $_.InstallState -eq 'Pending' }).Count
            $naCnt    = @($appsList | Where-Object { $_.InstallState -eq 'Not Applicable' }).Count
            $unkCnt   = @($appsList | Where-Object { $_.InstallState -like 'Unknown*' -or $_.InstallState -like 'Status:*' }).Count
            # Success rate denominator excludes Unknown/unresolved apps (IME backfill, stale entries)
            $applicableCount = $appsList.Count - $unkCnt
            $mdmInfo.MdmDiag.AppSummary = @{
                Total = $appsList.Count; Installed = $instCnt; Failed = $failCnt
                Pending = $pendCnt; NotApplicable = $naCnt; Unknown = $unkCnt
                SuccessRate = if ($applicableCount -gt 0) { [math]::Round(($instCnt / $applicableCount) * 100, 0) } else { 100 }
            }
        }
        # Compliance: separate hard failures (scripts/profiles/enrollment) from soft warnings (app installs)
        $compIssues = [System.Collections.Generic.List[string]]::new()
        $compWarnings = [System.Collections.Generic.List[string]]::new()
        if ($mdmInfo.MdmDiag.AppSummary -and $mdmInfo.MdmDiag.AppSummary.Failed -gt 0) { $compWarnings.Add("$($mdmInfo.MdmDiag.AppSummary.Failed) app(s) failed to install") }
        if ($mdmInfo.MdmDiag.AppSummary -and $mdmInfo.MdmDiag.AppSummary.Pending -gt 0) { $compWarnings.Add("$($mdmInfo.MdmDiag.AppSummary.Pending) app(s) pending install") }
        # N4: script policy failures
        if ($mdmInfo.MdmDiag.ScriptPolicies) {
            $failedScripts = @($mdmInfo.MdmDiag.ScriptPolicies | Where-Object { $_.Result -match 'Fail' })
            if ($failedScripts.Count -gt 0) { $compIssues.Add("$($failedScripts.Count) remediation/detection script(s) failed") }
        }
        # N8: config profile errors
        if ($mdmInfo.MdmDiag.ConfigProfiles) {
            $profileErrors = @($mdmInfo.MdmDiag.ConfigProfiles | Where-Object { $_.Status -match 'Error' })
            if ($profileErrors.Count -gt 0) { $compIssues.Add("$($profileErrors.Count) config profile(s) have errors") }
        }
        # N9: enrollment issues
        if ($mdmInfo.MdmDiag.EnrollmentIssues -and $mdmInfo.MdmDiag.EnrollmentIssues.Count -gt 0) { $compIssues.Add("$($mdmInfo.MdmDiag.EnrollmentIssues.Count) enrollment issue(s) detected") }
        # Compliance: script/profile/enrollment failures → Non-compliant; app install issues → At Risk only
        $compStatus = if (-not $enrolled) { 'N/A' }
                      elseif ($compIssues.Count -eq 0 -and $compWarnings.Count -eq 0) { 'Compliant' }
                      elseif ($compIssues.Count -gt 0) { 'Non-compliant' }
                      else { 'At Risk' }
        $allIssues = [System.Collections.Generic.List[string]]::new()
        foreach ($i in $compIssues)  { [void]$allIssues.Add($i) }
        foreach ($w in $compWarnings) { [void]$allIssues.Add($w) }
        $mdmInfo.MdmDiag | Add-Member -NotePropertyName 'Compliance' -NotePropertyValue @{
            Status = $compStatus
            ConfiguredPolicies = $allSettingsList.Count
            Issues = @($allIssues)
            AppSuccessRate = if ($mdmInfo.MdmDiag.AppSummary) { $mdmInfo.MdmDiag.AppSummary.SuccessRate } else { $null }
        } -Force
        if ($enrolled) {
            $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[Intune] Compliance: $compStatus ($($compIssues.Count) issues, $($allSettingsList.Count) policies configured)";Level='INFO'})
        }

        $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[$ScanMode] Scan complete: $($gpoRecords.Count) GPOs, $($allSettingsList.Count) settings";Level='SUCCESS'})
        $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[$ScanMode] Returning result: GPOs=$($gpoRecords.Count) Settings=$($allSettingsList.Count) Domain=$domain";Level='DEBUG'})

        # Cleanup temp MDM diagnostics
        if ($mdmDiagDir -and (Test-Path $mdmDiagDir)) {
            try { Remove-Item $mdmDiagDir -Recurse -Force -EA SilentlyContinue } catch { try { Write-DebugLog "Unhandled: $_" -Level ERROR } catch {} }
        }

        return @{
            Timestamp = [datetime]::Now
            Domain    = $domain
            GPOs      = $gpoRecords
            Settings  = $allSettingsList
            Apps      = $appsList
            CspMeta   = $cspMeta
            CspDbAge  = $jsonAge          # days since csp_metadata.json was last updated (null if not loaded)
            CspDbCount = $cspMeta.Count   # number of CSP entries loaded
            MdmInfo   = $mdmInfo
            ImeBackfillCount = $imeAdded
        }

        } catch {
            $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[$ScanMode] FATAL ERROR: $($_.Exception.Message)";Level='ERROR'})
            $SyncH.StatusQueue.Enqueue(@{Type='Log';Text="[$ScanMode] Stack: $($_.ScriptStackTrace)";Level='ERROR'})
            return @{ Error = $_.Exception.Message }
        }
    } -OnComplete {
        param($Results, $Errors, $Ctx)
        Write-DebugLog "OnComplete: Results=$($Results.Count) Errors=$($Errors.Count) ResultType=$($Results.GetType().Name)" -Level DEBUG
        # â”€â”€ OnComplete runs on UI thread -- safe to touch $ui, $Script: â”€â”€
        $ui.BtnScanGPOs.IsEnabled = $true
        $ui.ScanProgressPanel.Visibility = 'Collapsed'
        if ($ui.pnlGlobalProgress) { $ui.pnlGlobalProgress.Visibility = 'Collapsed'; if ($ui.lblGlobalProgress) { $ui.lblGlobalProgress.Text = '' } }

        $scanResult = if ($Results -and $Results.Count -gt 0) { $Results[$Results.Count - 1] } else { $null }
        Write-DebugLog "OnComplete: scanResult type=$(if ($scanResult) { $scanResult.GetType().Name } else { 'NULL' }) keys=$(if ($scanResult -is [hashtable]) { $scanResult.Keys -join ',' } else { 'n/a' })" -Level DEBUG

        # Check for errors from background thread
        if ($Errors -and $Errors.Count -gt 0) {
            foreach ($e in $Errors) { Write-DebugLog "Scan BgError: $($e.ToString())" -Level ERROR }
        }
        if (-not $scanResult -or $scanResult.Error) {
            $errMsg = if ($scanResult.Error) { $scanResult.Error } else { 'Scan returned no data' }
            Show-Toast 'Scan Failed' $errMsg 'error'
            Set-Status 'Scan failed' '#FF5000'
            Write-DebugLog "Scan failed: $errMsg" -Level ERROR
            return
        }

        $Script:ScanData = $scanResult
        if ($scanResult.CspMeta) { $Script:CspMetaKeys = $scanResult.CspMeta }
        $Script:CspDbAge   = $scanResult.CspDbAge
        $Script:CspDbCount = $scanResult.CspDbCount

        Write-Host "[SCAN] Loading GPOs=$($scanResult.GPOs.Count) Settings=$($scanResult.Settings.Count) into collections"
        $Script:AllGPOs.Clear()
        foreach ($g in $scanResult.GPOs) { [void]$Script:AllGPOs.Add($g) }
        $Script:AllSettings.Clear()
        foreach ($s in $scanResult.Settings) { [void]$Script:AllSettings.Add($s) }

        # ── Save scan snapshot to disk for crash recovery / fast reload ──
        try {
            $snapshotDir = [IO.Path]::Combine($env:TEMP, 'PolicyPilot_GPOCache')
            if (-not (Test-Path $snapshotDir)) { [void][IO.Directory]::CreateDirectory($snapshotDir) }
            $snapshotPath = [IO.Path]::Combine($snapshotDir, 'last_scan.clixml')
            $snapshot = @{
                Timestamp = $scanResult.Timestamp
                Domain    = $scanResult.Domain
                ScanMode  = $scanResult.ScanMode
                GPOs      = @($scanResult.GPOs)
                Settings  = @($scanResult.Settings)
                Apps      = if ($scanResult.Apps) { @($scanResult.Apps) } else { @() }
                MdmInfo   = $scanResult.MdmInfo
            }
            $snapshot | Export-Clixml -Path $snapshotPath -Force
            Write-DebugLog "Scan snapshot saved: $snapshotPath ($([math]::Round((Get-Item $snapshotPath).Length / 1KB))KB)" -Level DEBUG
        } catch {
            Write-DebugLog "Failed to save scan snapshot: $($_.Exception.Message)" -Level WARN
        }

        # Post-scan ADMX enrichment: resolve raw registry paths to English policy names
        if ($Script:AdmxByReg.Count -gt 0 -or $Script:CspByReg.Count -gt 0) {
            $enriched = 0
            foreach ($s in $Script:AllSettings) {
                if ($s.RegistryKey) {
                    # RegistryKey already contains "keyPath\valueName" — pass as-is with empty ValueName
                    # to avoid double-appending the value name in Resolve-PolicyFromRegistry
                    $resolved = Resolve-PolicyFromRegistry -RegistryKey $s.RegistryKey -ValueName '' -FallbackName $s.PolicyName
                    if ($resolved.Source -in @('ADMX','CSP')) {
                        try { $s.PolicyName = $resolved.Name } catch { $s | Add-Member -NotePropertyName 'PolicyName' -NotePropertyValue $resolved.Name -Force }
                        if ($resolved.Category) { try { $s.Category = $resolved.Category } catch { $s | Add-Member -NotePropertyName 'Category' -NotePropertyValue $resolved.Category -Force } }
                        if ($resolved.Desc) {
                            $hasProp = $s.PSObject.Properties.Name -contains 'Explain'
                            if ($hasProp) { try { $s.Explain = $resolved.Desc } catch { } }
                            else { $s | Add-Member -NotePropertyName 'Explain' -NotePropertyValue $resolved.Desc -Force }
                        }
                        $enriched++
                    }
                }
            }
            if ($enriched -gt 0) { Write-DebugLog "OnComplete: ADMX/CSP enriched $enriched settings with English names" -Level INFO }
        } else {
            Write-DebugLog "OnComplete: ADMX/CSP databases not loaded - settings show raw names. Run Build-AdmxDatabase.ps1 to generate." -Level WARN
            Show-Toast 'ADMX Database Missing' 'Policy names show raw registry paths. Place admx_metadata.json next to PolicyPilot.ps1 for English names.' 'warning'
        }

        Write-Host "[SCAN] Collections loaded - AllGPOs=$($Script:AllGPOs.Count) AllSettings=$($Script:AllSettings.Count)"
        Write-DebugLog "OnComplete: Loaded $($Script:AllGPOs.Count) GPOs, $($Script:AllSettings.Count) settings into UI collections" -Level DEBUG

        # Toggle empty-state vs data-grid visibility
        if ($Script:AllGPOs.Count -gt 0) {
            if ($ui.GPOListEmptyState) { $ui.GPOListEmptyState.Visibility = 'Collapsed' }
            if ($ui.GPOListGrid)       { $ui.GPOListGrid.Visibility = 'Visible' }
        }
        if ($Script:AllSettings.Count -gt 0) {
            if ($ui.SettingsEmptyState) { $ui.SettingsEmptyState.Visibility = 'Collapsed' }
            if ($ui.SettingsGrid)       { $ui.SettingsGrid.Visibility = 'Visible' }
        }

        $conflictResults = Find-Conflicts $scanResult.Settings
        $Script:AllConflicts.Clear()
        foreach ($c in $conflictResults) { if ($c.PSObject.Properties.Name -contains 'Severity') { [void]$Script:AllConflicts.Add($c) } }
        Write-DebugLog "OnComplete: Found $($Script:AllConflicts.Count) conflicts" -Level DEBUG

        if ($Script:AllConflicts.Count -gt 0) {
            if ($ui.ConflictsEmptyState) { $ui.ConflictsEmptyState.Visibility = 'Collapsed' }
            if ($ui.ConflictsGrid)       { $ui.ConflictsGrid.Visibility = 'Visible' }
        } else {
            if ($ui.ConflictsEmptyState) { $ui.ConflictsEmptyState.Visibility = 'Visible' }
            if ($ui.ConflictsGrid)       { $ui.ConflictsGrid.Visibility = 'Collapsed' }
        }

        $Script:AllIntuneApps.Clear()
        if ($scanResult.Apps) { foreach ($a in $scanResult.Apps) { [void]$Script:AllIntuneApps.Add($a) } }
        Write-DebugLog "OnComplete: Loaded $($Script:AllIntuneApps.Count) Intune apps" -Level DEBUG
        if ($ui.IntuneAppsGrid) { Apply-IntuneAppsFilter }

        if ($Script:AllIntuneApps.Count -gt 0) {
            if ($ui.IntuneAppsEmptyState) { $ui.IntuneAppsEmptyState.Visibility = 'Collapsed' }
            if ($ui.IntuneAppsGrid)       { $ui.IntuneAppsGrid.Visibility = 'Visible' }
        } else {
            if ($ui.IntuneAppsEmptyState) { $ui.IntuneAppsEmptyState.Visibility = 'Visible' }
            if ($ui.IntuneAppsGrid)       { $ui.IntuneAppsGrid.Visibility = 'Collapsed' }
        }

        # Update IntuneApps sidebar summary & enable reset
        $appCount = $Script:AllIntuneApps.Count
        $failedCount = @($Script:AllIntuneApps | Where-Object { $_.InstallState -eq 'Failed' }).Count
        $installedCount = @($Script:AllIntuneApps | Where-Object { $_.InstallState -eq 'Installed' }).Count
        if ($ui.IntuneAppsSummaryText) { $ui.IntuneAppsSummaryText.Text = "$appCount apps tracked`n$installedCount installed, $failedCount failed" }
        # Show admin warning banner if not elevated
        $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        if ($ui.AppAdminBanner) {
            if (-not $isAdmin -and $appCount -gt 0) {
                $ui.AppAdminBanner.Visibility = 'Visible'
                $imeNote = if ($scanResult.ImeBackfillCount -and $scanResult.ImeBackfillCount -gt 0) { " ($($scanResult.ImeBackfillCount) from IME log)" } else { '' }
                if ($ui.AppAdminBannerText) { $ui.AppAdminBannerText.Text = "Not running as admin $([char]0x2014) app reset is disabled and Win32 registry data unavailable$imeNote. Run elevated for full app discovery and reset." }
            } else {
                $ui.AppAdminBanner.Visibility = 'Collapsed'
            }
        }
        if ($ui.BtnResetAppInstall) {
            $ui.BtnResetAppInstall.IsEnabled = ($appCount -gt 0 -and $isAdmin)
            $ui.BtnResetAppInstall.ToolTip = if ($isAdmin) { 'Delete the IME tracking registry key for the selected app so Intune re-evaluates it on next sync' } else { 'Requires elevation $([char]0x2014) restart PolicyPilot as Administrator to enable app reset' }
        }

        if ($ui.CmbCategoryFilter) {
            $ui.CmbCategoryFilter.Items.Clear()
            [void]$ui.CmbCategoryFilter.Items.Add('(All Categories)')
            $cats = @($scanResult.Settings | Select-Object -ExpandProperty Category -Unique | Where-Object { $_ } | Sort-Object)
            foreach ($cat in $cats) { [void]$ui.CmbCategoryFilter.Items.Add($cat) }
            $ui.CmbCategoryFilter.SelectedIndex = 0
        }

        # ── Intune Category → Group mapping (main thread) ──
        function Get-IntuneGroup ($cat, $source) {
            if ($cat -like 'Endpoint Security:*') { return 'Endpoint Security' }
            if ($cat -like 'Device Security:*')   { return 'Account Protection' }
            if ($cat -like 'Compliance:*')         { return 'Device Compliance' }
            if ($cat -eq 'Windows Update')         { return 'Windows Update' }
            if ($cat -eq 'App Management')         { return 'App Management' }
            if ($cat -eq 'Registry Settings' -or $cat -like '*(Registry)*') { return 'Registry Settings' }
            if ($cat -in @('Microsoft Edge','Power Management','Device Restrictions') -or
               $cat -like 'System:*' -or $cat -like 'Connectivity:*' -or $cat -like 'Microsoft:*') {
                return 'Configuration Profiles'
            }
            if ($source -eq 'Intune') { return 'ADMX Profiles' }
            return 'Group Policy'
        }

        # Generate "Not Configured" CSP gap analysis for Intune scans
        $Script:NotConfiguredSettings.Clear()
        if ($Script:Prefs.ScanMode -in @('Intune','Combined')) {
            $configuredKeys = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            foreach ($s in $scanResult.Settings) {
                if ($s.SettingKey) { $null = $configuredKeys.Add($s.SettingKey) }
            }
            $ncId = 0
            foreach ($key in $Script:CspMetaKeys.Keys) {
                if (-not $configuredKeys.Contains($key)) {
                    $meta = $Script:CspMetaKeys[$key]
                    $ncId++
                    $ncCat = if ($meta.Cat) { $meta.Cat } else { 'Uncategorized' }
                    $ncName = if ($meta.Friendly) { $meta.Friendly } else { $key }
                    $ncDef = if ($meta.Def) { $meta.Def } else { '' }
                    [void]$Script:NotConfiguredSettings.Add([PSCustomObject]@{
                        Id=$ncId; GPOName=''; GPOGuid=''
                        Category=$ncCat
                        PolicyName=$ncName
                        SettingKey=$key; State='Not Configured'
                        RegistryKey=''; ValueData=''; Scope='Device'
                        DefaultValue=$ncDef; Source='CSP Baseline'; IntuneGroup=(Get-IntuneGroup $ncCat 'Intune')
                    })
                }
            }
            Write-DebugLog "Gap analysis: $($configuredKeys.Count) configured, $($Script:NotConfiguredSettings.Count) not configured from $($Script:CspMetaKeys.Count) known CSPs" -Level INFO
        }

        # Update Settings sidebar summary
        $cfgCount = $scanResult.Settings.Count
        $ncCount  = $Script:NotConfiguredSettings.Count
        $ndCount  = @($scanResult.Settings | Where-Object { $_.DefaultValue -and $_.ValueData -and "$($_.ValueData)" -ne "$($_.DefaultValue)" -and "$($_.ValueData)" -notlike "$($_.DefaultValue) *" }).Count
        $totalKnown = $cfgCount + $ncCount
        $coverage = if ($totalKnown -gt 0) { [math]::Round(($cfgCount / $totalKnown) * 100, 0) } else { 0 }
        if ($ui.StatConfigured)    { $ui.StatConfigured.Text    = "$cfgCount" }
        if ($ui.StatNotConfigured) { $ui.StatNotConfigured.Text = "$ncCount" }
        if ($ui.StatNonDefault)    { $ui.StatNonDefault.Text    = "$ndCount" }
        if ($ui.StatCoverage)      { $ui.StatCoverage.Text      = "${coverage}%" }

        Write-Host "[SCAN] Running Update-Dashboard..."
        try { Update-Dashboard } catch { Write-Host "[SCAN] Update-Dashboard CRASHED: $($_.Exception.Message)"; Write-Host "[SCAN]   Inner: $(if($_.Exception.InnerException){$_.Exception.InnerException.Message}else{'none'})"; Write-DebugLog "Update-Dashboard CRASH: $($_.Exception.Message)" -Level ERROR }
        Write-Host "[SCAN] Running Update-ReportPreview..."
        try { Update-ReportPreview } catch { Write-Host "[SCAN] Update-ReportPreview CRASHED: $($_.Exception.Message)"; Write-DebugLog "Update-ReportPreview CRASH: $($_.Exception.Message)" -Level ERROR }

        $activeCount    = @($scanResult.GPOs | Where-Object { $_.Status -ne 'Disabled' }).Count
        $conflictCount  = @($conflictResults | Where-Object Severity -eq 'Conflict').Count
        $redundantCount = @($conflictResults | Where-Object Severity -eq 'Redundant').Count
        $msg = "$activeCount GPOs, $($scanResult.Settings.Count) settings"
        if ($conflictCount -gt 0)  { $msg += ", $conflictCount conflicts" }
        if ($redundantCount -gt 0) { $msg += ", $redundantCount redundant" }

        Show-Toast 'Scan Complete' $msg 'success'

        # ── Achievement checks ──
        $elapsedSec = if ($Script:ScanStartTime) { ([DateTime]::Now - $Script:ScanStartTime).TotalSeconds } else { 999 }
        Write-Host "[SCAN] Running Check-ScanAchievements..."
        try { Check-ScanAchievements -SettingCount $scanResult.Settings.Count -GpoCount $scanResult.GPOs.Count -ConflictCount $conflictCount -ScanMode $Script:Prefs.ScanMode -ElapsedSec $elapsedSec } catch { Write-Host "[SCAN] Check-ScanAchievements CRASHED: $($_.Exception.Message)"; Write-DebugLog "Achievements CRASH: $($_.Exception.Message)" -Level ERROR }
        if ($ui.lblStatusDetail) { $ui.lblStatusDetail.Text = $msg }
        Set-Status 'Ready' '#00C853'

        # ── Update sidebar scan status indicators ──
        $scanTime = (Get-Date).ToString('HH:mm')
        $modeLabel = switch ($Script:Prefs.ScanMode) { 'Intune' { 'Intune' } 'Local' { 'Local' } 'AD' { 'AD' } 'Combined' { 'Combined' } default { $Script:Prefs.ScanMode } }
        $statusText = "$modeLabel scan at $scanTime - $msg"
        foreach ($lbl in @('LblScanStatusGPOList','LblScanStatusSettings','LblScanStatusConflicts','LblScanStatusReport')) {
            if ($ui[$lbl]) { $ui[$lbl].Text = $statusText }
        }

        Write-DebugLog "Scan mode: $($Script:Prefs.ScanMode) | Domain: $($scanResult.Domain) | GPOs: $($scanResult.GPOs.Count) | Settings: $($scanResult.Settings.Count)" -Level INFO
        Write-DebugLog "Scan loaded into UI: $msg" -Level SUCCESS
    }
})

# GetStarted button → same as scan
$ui.BtnGetStarted.Add_Click({
    $ui.BtnScanGPOs.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
})


# Sidebar "Go to Dashboard" buttons → navigate to Dashboard tab
foreach ($btnName in @('BtnGoToDashGPOList','BtnGoToDashSettings','BtnGoToDashConflicts','BtnGoToDashReport')) {
    if ($ui[$btnName]) {
        $ui[$btnName].Add_Click({ Switch-Tab 'Dashboard' })
    }
}
# IntuneApps scan button → ensure scan mode includes Intune, then scan
if ($ui.BtnScan_IntuneApps) {
    $ui.BtnScan_IntuneApps.Add_Click({
        if ($Script:Prefs.ScanMode -notin @('Intune','Combined')) {
            $ui.CmbScanMode.SelectedIndex = 0  # Switch to Intune mode
        }
        $ui.BtnScanGPOs.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
    })
}
# View All Conflicts → switch to Conflicts tab
$ui.BtnViewAllConflicts.Add_Click({ Switch-Tab 'Conflicts' })

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 14: FILTER HANDLERS
# ═══════════════════════════════════════════════════════════════════════════════

function Apply-GPOFilter {
    $search = $ui.TxtGPOSearch.Text.Trim()
    $statusFilter = if ($ui.FilterGPOEnabled.Tag -eq 'Active') { 'Enabled' }
                    elseif ($ui.FilterGPODisabled.Tag -eq 'Active') { 'Disabled' }
                    else { $null }
    $scopeFilter = if ($ui.FilterScopeComputer.Tag -eq 'Active') { 'Computer' }
                   elseif ($ui.FilterScopeUser.Tag -eq 'Active') { 'User' }
                   else { $null }

    if ($ui.GPOListGrid) { $ui.GPOListGrid.ItemsSource = $null }
    $Script:AllGPOs.Clear()
    if (-not $Script:ScanData) {
        if ($ui.GPOListGrid) { $ui.GPOListGrid.ItemsSource = $Script:AllGPOs }
        return
    }

    foreach ($g in $Script:ScanData.GPOs) {
        if ($search -and $g.DisplayName -notlike "*$search*" -and "$($g.GpoId)" -notlike "*$search*") { continue }
        if ($statusFilter -eq 'Enabled'  -and $g.Status -eq 'Disabled') { continue }
        if ($statusFilter -eq 'Disabled' -and $g.Status -ne 'Disabled') { continue }
        if ($scopeFilter -and $g.Scope -and $g.Scope -ne $scopeFilter) { continue }
        [void]$Script:AllGPOs.Add($g)
    }
    if ($ui.GPOListGrid) { $ui.GPOListGrid.ItemsSource = $Script:AllGPOs }
    $ui.GPOListSubtitle.Text = "$($Script:AllGPOs.Count) of $($Script:ScanData.GPOs.Count) GPOs"
}

function Apply-SettingFilter {
    $search = $ui.TxtSettingSearch.Text.Trim()
    $scopeFilter = if ($ui.FilterSettingComputer.Tag -eq 'Active') { 'Computer' }
                   elseif ($ui.FilterSettingUser.Tag -eq 'Active') { 'User' }
                   else { $null }
    $catFilter = if ($ui.CmbCategoryFilter -and $ui.CmbCategoryFilter.SelectedIndex -gt 0) {
        $ui.CmbCategoryFilter.SelectedItem
    } else { $null }

    # Configuration filter (Configured / Not Set / All)
    $cfgFilter = if ($ui.FilterCfgConfigured -and $ui.FilterCfgConfigured.Tag -eq 'Active') { 'Configured' }
                 elseif ($ui.FilterCfgNotConfigured -and $ui.FilterCfgNotConfigured.Tag -eq 'Active') { 'NotConfigured' }
                 else { $null }

    # Value status filter (Non-Default / Default / All)
    $valFilter = if ($ui.FilterValNonDefault -and $ui.FilterValNonDefault.Tag -eq 'Active') { 'NonDefault' }
                 elseif ($ui.FilterValDefault -and $ui.FilterValDefault.Tag -eq 'Active') { 'Default' }
                 else { $null }

    # Group filter (IntuneGroup)
    $groupFilter = if ($ui.FilterGroupEndpointSec -and $ui.FilterGroupEndpointSec.Tag -eq 'Active') { 'Endpoint Security' }
                   elseif ($ui.FilterGroupAcctProt -and $ui.FilterGroupAcctProt.Tag -eq 'Active') { 'Account Protection' }
                   elseif ($ui.FilterGroupCompliance -and $ui.FilterGroupCompliance.Tag -eq 'Active') { 'Device Compliance' }
                   elseif ($ui.FilterGroupWinUpdate -and $ui.FilterGroupWinUpdate.Tag -eq 'Active') { 'Windows Update' }
                   elseif ($ui.FilterGroupAppMgmt -and $ui.FilterGroupAppMgmt.Tag -eq 'Active') { 'App Management' }
                   elseif ($ui.FilterGroupRegistry -and $ui.FilterGroupRegistry.Tag -eq 'Active') { 'Registry Settings' }
                   elseif ($ui.FilterGroupConfigProf -and $ui.FilterGroupConfigProf.Tag -eq 'Active') { 'Configuration Profiles' }
                   elseif ($ui.FilterGroupADMXProf -and $ui.FilterGroupADMXProf.Tag -eq 'Active') { 'ADMX Profiles' }
                   elseif ($ui.FilterGroupGPO -and $ui.FilterGroupGPO.Tag -eq 'Active') { 'Group Policy' }
                   elseif ($ui.FilterGroupPpkg -and $ui.FilterGroupPpkg.Tag -eq 'Active') { 'Provisioning Packages' }
                   else { $null }
    # Detach ItemsSource to prevent per-item UI updates during batch rebuild
    if ($ui.SettingsGrid) { $ui.SettingsGrid.ItemsSource = $null }
    $Script:AllSettings.Clear()
    if (-not $Script:ScanData) {
        if ($ui.SettingsGrid) { $ui.SettingsGrid.ItemsSource = $Script:AllSettings }
        return
    }

    # Build source list based on configuration filter
    $sourceList = [System.Collections.Generic.List[object]]::new()
    if ($cfgFilter -eq 'NotConfigured') {
        foreach ($nc in $Script:NotConfiguredSettings) { [void]$sourceList.Add($nc) }
    } elseif ($cfgFilter -eq 'Configured') {
        foreach ($s in $Script:ScanData.Settings) { [void]$sourceList.Add($s) }
    } else {
        foreach ($s in $Script:ScanData.Settings) { [void]$sourceList.Add($s) }
        foreach ($nc in $Script:NotConfiguredSettings) { [void]$sourceList.Add($nc) }
    }

    foreach ($s in $sourceList) {
        if ($search -and $s.PolicyName -notlike "*$search*" -and $s.Category -notlike "*$search*" -and $s.RegistryKey -notlike "*$search*" -and $s.GPOName -notlike "*$search*" -and "$($s.ValueData)" -notlike "*$search*") { continue }
        if ($scopeFilter -and $s.Scope -ne $scopeFilter) { continue }
        if ($catFilter -and $s.Category -ne $catFilter) { continue }
        if ($groupFilter -and $s.IntuneGroup -ne $groupFilter) { continue }
        if ($valFilter -eq 'NonDefault' -and $s.DefaultValue -and "$($s.ValueData)" -eq "$($s.DefaultValue)") { continue }
        if ($valFilter -eq 'Default' -and ($s.DefaultValue -eq $null -or "$($s.ValueData)" -ne "$($s.DefaultValue)")) { continue }
        [void]$Script:AllSettings.Add($s)
    }
    # Re-apply grouping after filter
    $v = [System.Windows.Data.CollectionViewSource]::GetDefaultView($Script:AllSettings)
    if ($v -and $v.GroupDescriptions.Count -eq 0) {
        [void]$v.GroupDescriptions.Add([System.Windows.Data.PropertyGroupDescription]::new('IntuneGroup'))
    }
    # Re-attach ItemsSource (triggers single UI refresh instead of per-item)
    if ($ui.SettingsGrid) { $ui.SettingsGrid.ItemsSource = $Script:AllSettings }
    $total = $Script:ScanData.Settings.Count + $Script:NotConfiguredSettings.Count
    $ui.SettingsSubtitle.Text = "$($Script:AllSettings.Count) of $total settings"
}

function Apply-ConflictFilter {
    $sevFilter = if ($ui.FilterConflictOnly.Tag -eq 'Active') { 'Conflict' }
                 elseif ($ui.FilterRedundantOnly.Tag -eq 'Active') { 'Redundant' }
                 else { $null }

    $Script:AllConflicts.Clear()
    if (-not $Script:ScanData) { return }

    $source = @(Find-Conflicts $Script:ScanData.Settings)
    foreach ($c in $source) {
        if ($c.PSObject.Properties.Name -notcontains 'Severity') { continue }
        if ($sevFilter -and $c.Severity -ne $sevFilter) { continue }
        [void]$Script:AllConflicts.Add($c)
    }
    $totalIssues = @($source | Where-Object { $_.PSObject.Properties.Name -contains 'Severity' }).Count
    $ui.ConflictsSubtitle.Text = "$($Script:AllConflicts.Count) of $totalIssues issues"
}

# GPO filter pill wiring
function Set-FilterPillGroup([System.Windows.Controls.Button[]]$pills, [System.Windows.Controls.Button]$active) {
    foreach ($p in $pills) { $p.Tag = $null }
    $active.Tag = 'Active'
}

# Wire pill groups - use $this (sender) instead of .GetNewClosure() to avoid function-scope issues
$gpoStatusPills = @($ui.FilterGPOAll, $ui.FilterGPOEnabled, $ui.FilterGPODisabled)
foreach ($pill in $gpoStatusPills) { $pill.Add_Click({ foreach ($b in $gpoStatusPills) { $b.Tag = $null }; $this.Tag = 'Active'; Apply-GPOFilter }) }

$scopePills = @($ui.FilterScopeBoth, $ui.FilterScopeComputer, $ui.FilterScopeUser)
foreach ($pill in $scopePills) { $pill.Add_Click({ foreach ($b in $scopePills) { $b.Tag = $null }; $this.Tag = 'Active'; Apply-GPOFilter }) }

# Setting filter pills
$settingScopePills = @($ui.FilterSettingAll, $ui.FilterSettingComputer, $ui.FilterSettingUser)
foreach ($pill in $settingScopePills) { $pill.Add_Click({ foreach ($b in $settingScopePills) { $b.Tag = $null }; $this.Tag = 'Active'; Apply-SettingFilter }) }

# Configuration filter pills (Configured / Not Set)
$cfgPills = @($ui.FilterCfgAll, $ui.FilterCfgConfigured, $ui.FilterCfgNotConfigured)
if ($ui.FilterCfgAll) { foreach ($pill in $cfgPills) { $pill.Add_Click({ foreach ($b in $cfgPills) { $b.Tag = $null }; $this.Tag = 'Active'; Apply-SettingFilter }) } }

# Value status filter pills (Non-Default / Default)
$valPills = @($ui.FilterValAll, $ui.FilterValNonDefault, $ui.FilterValDefault)
if ($ui.FilterValAll) { foreach ($pill in $valPills) { $pill.Add_Click({ foreach ($b in $valPills) { $b.Tag = $null }; $this.Tag = 'Active'; Apply-SettingFilter }) } }

# Group filter pills (IntuneGroup)
$groupPills = @($ui.FilterGroupAll, $ui.FilterGroupEndpointSec, $ui.FilterGroupAcctProt,
               $ui.FilterGroupCompliance, $ui.FilterGroupWinUpdate, $ui.FilterGroupAppMgmt,
               $ui.FilterGroupRegistry, $ui.FilterGroupConfigProf, $ui.FilterGroupADMXProf, $ui.FilterGroupGPO, $ui.FilterGroupPpkg) | Where-Object { $_ }
if ($ui.FilterGroupAll) { foreach ($pill in $groupPills) { $pill.Add_Click({ foreach ($b in $groupPills) { $b.Tag = $null }; $this.Tag = 'Active'; Apply-SettingFilter }) } }

# Conflict filter pills
$conflictPills = @($ui.FilterConflictAll, $ui.FilterConflictOnly, $ui.FilterRedundantOnly)
foreach ($pill in $conflictPills) { $pill.Add_Click({ foreach ($b in $conflictPills) { $b.Tag = $null }; $this.Tag = 'Active'; Apply-ConflictFilter }) }
# Search boxes
$ui.TxtGPOSearch.Add_TextChanged({ Apply-GPOFilter })
$ui.TxtSettingSearch.Add_TextChanged({ Apply-SettingFilter })
if ($ui.CmbCategoryFilter) { $ui.CmbCategoryFilter.Add_SelectionChanged({ Apply-SettingFilter }) }

# GPO list selection → detail card
$ui.GPOListGrid.Add_SelectionChanged({
    $sel = $ui.GPOListGrid.SelectedItem
    if (-not $sel) { return }
    $ui.GPODetailName.Text     = $sel.DisplayName
    $ui.GPODetailGuid.Text     = $sel.Id
    $ui.GPODetailStatus.Text   = "Status: $($sel.Status)"
    $ui.GPODetailLinks.Text    = if ($sel.Links) { "Links: $($sel.Links)" } else { 'Not linked to any OU' }
    $ui.GPODetailModified.Text = "Modified: $($sel.ModificationTime)"
})

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 15: EXPORT - HTML
# ═══════════════════════════════════════════════════════════════════════════════


function Update-ReportPreview {
    if (-not $Script:ScanData) {
        if ($ui.ReportPreviewText) { $ui.ReportPreviewText.Text = "Run a scan and click 'Export HTML Report' to generate documentation." }
        return
    }
    $sb = [System.Text.StringBuilder]::new()
    $gpos = $Script:ScanData.GPOs
    $settings = $Script:ScanData.Settings
    $conflicts = @($Script:AllConflicts | Where-Object { $_.PSObject.Properties.Name -contains 'Severity' })
    $apps = $Script:AllIntuneApps

    # Header
    $null = $sb.AppendLine("GPO DOCUMENTATION PREVIEW")
    $null = $sb.AppendLine("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm')")
    $null = $sb.AppendLine("Scan mode: $($Script:Prefs.ScanMode)")
    $null = $sb.AppendLine("$([char]0x2500)" * 60)
    $null = $sb.AppendLine()

    # Summary
    $null = $sb.AppendLine("SUMMARY")
    $null = $sb.AppendLine("  Active GPOs / Areas:    $($gpos.Count)")
    $enabledCount = @($gpos | Where-Object { $_.Status -eq 'Enabled' }).Count
    $null = $sb.AppendLine("  Enabled:                $enabledCount")
    $null = $sb.AppendLine("  Total Settings:         $($settings.Count)")
    $null = $sb.AppendLine("  Conflicts:              $(@($conflicts | Where-Object { $_.Severity -eq 'Conflict' }).Count)")
    $null = $sb.AppendLine("  Redundancies:           $(@($conflicts | Where-Object { $_.Severity -eq 'Redundant' }).Count)")
    if ($apps.Count -gt 0) {
        $null = $sb.AppendLine("  Intune Apps:            $($apps.Count)")
    }
    if ($Script:NotConfiguredSettings.Count -gt 0) {
        $null = $sb.AppendLine("  Not Configured (CSP):   $($Script:NotConfiguredSettings.Count)")
    }
    $null = $sb.AppendLine()

    # Settings by Category
    $null = $sb.AppendLine("SETTINGS BY CATEGORY")
    $groups = $settings | Group-Object Category | Sort-Object Count -Descending
    foreach ($g in $groups) {
        $null = $sb.AppendLine("  $($g.Name)  ($($g.Count))")
    }
    $null = $sb.AppendLine()

    # Settings by Intune Group
    $null = $sb.AppendLine("SETTINGS BY INTUNE GROUP")
    $iGroups = $settings | Group-Object IntuneGroup | Sort-Object Count -Descending
    foreach ($g in $iGroups) {
        $null = $sb.AppendLine("  $($g.Name)  ($($g.Count))")
    }
    $null = $sb.AppendLine()

    # GPO / Area Inventory
    $null = $sb.AppendLine("GPO / AREA INVENTORY")
    $null = $sb.AppendLine("$([char]0x2500)" * 60)
    foreach ($gpo in ($gpos | Sort-Object DisplayName)) {
        $status = if ($gpo.Status) { " [$($gpo.Status)]" } else { '' }
        $null = $sb.AppendLine("  $($gpo.DisplayName)$status  -  $($gpo.SettingCount) settings")
    }
    $null = $sb.AppendLine()

    # Non-default Values
    $nonDef = @($settings | Where-Object { $_.DefaultValue -and $_.ValueData -and "$($_.ValueData)" -ne "$($_.DefaultValue)" -and "$($_.ValueData)" -notlike "$($_.DefaultValue) *" })
    if ($nonDef.Count -gt 0) {
        $null = $sb.AppendLine("NON-DEFAULT VALUES ($($nonDef.Count))")
        $null = $sb.AppendLine("$([char]0x2500)" * 60)
        foreach ($s in ($nonDef | Sort-Object Category, PolicyName)) {
            $null = $sb.AppendLine("  $($s.PolicyName)")
            $null = $sb.AppendLine("    Value: $($s.ValueData)  |  Default: $($s.DefaultValue)")
        }
        $null = $sb.AppendLine()
    }

    # Conflicts
    if ($conflicts.Count -gt 0) {
        $null = $sb.AppendLine("CONFLICTS & REDUNDANCIES ($($conflicts.Count))")
        $null = $sb.AppendLine("$([char]0x2500)" * 60)
        foreach ($c in ($conflicts | Sort-Object Severity, SettingKey)) {
            $null = $sb.AppendLine("  [$($c.Severity)] $($c.SettingKey)")
            $null = $sb.AppendLine("    GPOs: $($c.GPONames)  |  Values: $($c.Values)")
        }
    }

    if ($ui.ReportPreviewText) { $ui.ReportPreviewText.Text = $sb.ToString() }
}

function Export-Html {
    if (-not $Script:ScanData) {
        Show-Toast 'No Data' 'Run a scan first before exporting.' 'warning'
        return
    }

    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Filter = 'HTML files (*.html)|*.html'
    $dlg.FileName = "GPO_Documentation_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    $dlg.InitialDirectory = $Script:ReportsDir
    if (-not $dlg.ShowDialog()) { return }

    Write-DebugLog "Exporting HTML to $($dlg.FileName)" -Level STEP
    $htmlContent = Build-HtmlReport
    try {
        [System.IO.File]::WriteAllText($dlg.FileName, $htmlContent, [System.Text.Encoding]::UTF8)
        Write-DebugLog "HTML report saved: $($dlg.FileName)" -Level SUCCESS
        Show-Toast 'Report Exported' "Saved to $($dlg.FileName)" 'success'
        Unlock-Achievement 'first_export'
    } catch {
        Write-DebugLog "HTML export failed: $($_.Exception.Message)" -Level ERROR
        Show-Toast 'Export Failed' $_.Exception.Message 'error'
    }
}

function Build-HtmlReport {
    $gpos      = $Script:ScanData.GPOs
    $settings  = $Script:ScanData.Settings
    $conflicts = [System.Collections.Generic.List[PSCustomObject]]@(Find-Conflicts $settings)
    $domain    = $Script:ScanData.Domain
    $genDate   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $apps      = $Script:AllIntuneApps
    $notCfg    = $Script:NotConfiguredSettings
    $mdmInfo   = $Script:ScanData.MdmInfo

    # Collect device info for report header
    $deviceName = $env:COMPUTERNAME
    $userName   = "$env:USERDOMAIN\$env:USERNAME"
    $osInfo     = try { $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop; "$($os.Caption) ($($os.Version))" } catch { 'N/A' }
    $scanMode   = $Script:Prefs.ScanMode
    $scanTime   = if ($Script:ScanData.Timestamp -is [datetime]) { $Script:ScanData.Timestamp.ToString('yyyy-MM-dd HH:mm:ss') } else { $genDate }

    $activeCount    = @($gpos | Where-Object { $_.Status -ne 'Disabled' }).Count
    $totalSettings  = $settings.Count
    $conflictCount  = @($conflicts | Where-Object Severity -eq 'Conflict').Count
    $redundantCount = @($conflicts | Where-Object Severity -eq 'Redundant').Count
    $unlinkedCount  = @($gpos | Where-Object { -not $_.IsLinked }).Count
    $appCount       = if ($apps) { $apps.Count } else { 0 }
    $notCfgCount    = if ($notCfg) { $notCfg.Count } else { 0 }

    $enc = { param($s) [System.Web.HttpUtility]::HtmlEncode("$s") }

    $html = [System.Text.StringBuilder]::new(65536)
    [void]$html.Append(@"
<!DOCTYPE html>
<html lang="en" data-theme="dark">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>Intune / GPO Policy Report - $(& $enc $domain)</title>
<style>
:root {
  --bg: #0a0a0c; --bg2: #111114; --card: #18181b; --card-hover: #1e1e22;
  --border: rgba(255,255,255,0.08); --border-strong: rgba(255,255,255,0.15);
  --accent: #60cdff; --accent-dim: rgba(96,205,255,0.12); --accent-text: #60cdff;
  --text: #e4e4e7; --text-bright: #fafafa; --muted: #71717a; --subtle: #52525b;
  --green: #22c55e; --green-dim: rgba(34,197,94,0.12);
  --red: #ef4444; --red-dim: rgba(239,68,68,0.12);
  --yellow: #eab308; --yellow-dim: rgba(234,179,8,0.12);
  --orange: #f97316; --purple: #a855f7;
  --radius: 10px; --radius-sm: 6px;
  --shadow: 0 1px 3px rgba(0,0,0,0.4), 0 4px 12px rgba(0,0,0,0.3);
  --font: 'Segoe UI Variable Display','Segoe UI',-apple-system,system-ui,sans-serif;
  --font-mono: 'Cascadia Mono','Consolas','Courier New',monospace;
  --mono: 'Cascadia Code','Cascadia Mono','Fira Code',Consolas,monospace;
}
[data-theme="light"] {
  --bg: #f8f9fa; --bg2: #ffffff; --card: #ffffff; --card-hover: #f4f4f5;
  --border: rgba(0,0,0,0.08); --border-strong: rgba(0,0,0,0.15);
  --accent: #0078d4; --accent-dim: rgba(0,120,212,0.08); --accent-text: #0066b8;
  --text: #18181b; --text-bright: #09090b; --muted: #71717a; --subtle: #a1a1aa;
  --green: #16a34a; --green-dim: rgba(22,163,74,0.08);
  --red: #dc2626; --red-dim: rgba(220,38,38,0.08);
  --yellow: #ca8a04; --yellow-dim: rgba(202,138,4,0.08);
  --orange: #ea580c; --purple: #9333ea;
  --shadow: 0 1px 3px rgba(0,0,0,0.06), 0 4px 12px rgba(0,0,0,0.04);
}
*, *::before, *::after { margin:0; padding:0; box-sizing:border-box; }
html { scroll-behavior: smooth; }
body { font-family:var(--font); background:var(--bg); color:var(--text); line-height:1.6; }

/* Sticky toolbar */
.toolbar { position:sticky; top:0; z-index:100; background:var(--bg2); border-bottom:1px solid var(--border);
  padding:12px 32px; display:flex; align-items:center; gap:16px; backdrop-filter:blur(12px); }
.toolbar h1 { font-size:16px; font-weight:600; color:var(--text-bright); white-space:nowrap; }
.toolbar .domain-badge { background:var(--accent-dim); color:var(--accent-text); padding:3px 10px;
  border-radius:20px; font-size:11px; font-weight:600; }
.toolbar .spacer { flex:1; }
.search-box { background:var(--card); border:1px solid var(--border); border-radius:var(--radius-sm);
  padding:6px 12px; color:var(--text); font-size:13px; width:260px; outline:none; font-family:var(--font); }
.search-box:focus { border-color:var(--accent); box-shadow:0 0 0 2px var(--accent-dim); }
.search-box::placeholder { color:var(--subtle); }
.btn { background:var(--card); border:1px solid var(--border); border-radius:var(--radius-sm);
  padding:6px 14px; color:var(--text); font-size:12px; cursor:pointer; font-family:var(--font);
  transition:all 0.15s; }
.btn:hover { background:var(--card-hover); border-color:var(--border-strong); }
.btn-icon { padding:6px 8px; font-size:16px; line-height:1; }

.container { max-width:1320px; margin:0 auto; padding:24px 32px 64px; }

/* Stats bar */
.stats { display:grid; grid-template-columns:repeat(auto-fit,minmax(140px,1fr)); gap:10px; margin-bottom:28px; }
.stat-card { background:var(--card); border:1px solid var(--border); border-radius:var(--radius);
  padding:16px 18px; transition:all 0.15s; }
.stat-card:hover { border-color:var(--border-strong); box-shadow:var(--shadow); }
.stat-num { font-size:26px; font-weight:700; color:var(--text-bright); line-height:1.1; }
.stat-label { font-size:11px; color:var(--muted); margin-top:3px; letter-spacing:0.3px; }
.stat-card.green .stat-num { color:var(--green); }
.stat-card.red .stat-num { color:var(--red); }
.stat-card.yellow .stat-num { color:var(--yellow); }
.stat-card.accent .stat-num { color:var(--accent-text); }

/* Meta info */
.meta-bar { display:flex; gap:24px; flex-wrap:wrap; margin-bottom:24px; color:var(--muted); font-size:12px; }
.meta-bar strong { color:var(--text); }

/* Device Info */
.device-info { display:grid; grid-template-columns:repeat(auto-fit,minmax(260px,1fr)); gap:8px 24px;
  background:var(--card); border:1px solid var(--border); border-radius:var(--radius);
  padding:16px 20px; margin-bottom:20px; font-size:12px; color:var(--muted); }
.device-info .di-item { display:flex; gap:8px; align-items:baseline; padding:2px 0; }
.device-info .di-label { min-width:110px; font-weight:600; color:var(--subtle); letter-spacing:0.3px; }
.device-info .di-value { color:var(--text); }

/* Sections */
.section { margin-bottom:20px; }
.section-header { background:var(--card); border:1px solid var(--border); border-radius:var(--radius);
  padding:14px 18px; cursor:pointer; display:flex; align-items:center; gap:12px;
  transition:all 0.15s; user-select:none; }
.section-header:hover { background:var(--card-hover); }
.section-header .icon { font-size:18px; width:24px; text-align:center; }
.section-header .title { font-size:14px; font-weight:600; color:var(--text-bright); flex:1; }
.section-header .count { background:var(--accent-dim); color:var(--accent-text); padding:2px 10px;
  border-radius:12px; font-size:11px; font-weight:600; }
.section-header .chevron { color:var(--subtle); transition:transform 0.2s; font-size:12px; }
details[open] .chevron { transform:rotate(90deg); }
.section-body { border:1px solid var(--border); border-top:none; border-radius:0 0 var(--radius) var(--radius);
  background:var(--bg2); overflow:hidden; }

/* Tables */
table { width:100%; border-collapse:collapse; font-size:12px; table-layout:fixed; }
th { background:var(--card); color:var(--muted); text-align:left; padding:10px 14px;
  font-weight:600; font-size:10.5px; text-transform:uppercase; letter-spacing:0.6px;
  border-bottom:1px solid var(--border); position:sticky; top:0; z-index:1; }
td { padding:9px 14px; border-bottom:1px solid var(--border); vertical-align:top;
  word-wrap:break-word; overflow-wrap:break-word; }
tr:hover td { background:var(--card-hover); }
tr.highlight td { background:var(--accent-dim); }
td.val { font-family:var(--mono); font-size:11px; word-break:break-all; overflow-wrap:break-word; }
details.val-list { cursor:pointer; }
details.val-list > summary { list-style:none; }
details.val-list > summary::-webkit-details-marker { display:none; }
details.val-list > summary em { color:var(--muted); font-size:10px; }
td.policy-name { font-weight:500; word-wrap:break-word; overflow-wrap:break-word; }

/* CSP Reference block (collapsed by default under each setting row) */
tr.csp-row td { padding:0; border-bottom:none; }
.csp-ref { margin:0; font-size:11.5px; }
.csp-ref > summary { padding:7px 14px; cursor:pointer; color:var(--muted); font-size:10.5px;
  list-style:none; user-select:none; display:flex; align-items:center; gap:6px;
  transition:color .15s, background .15s; }
.csp-ref > summary:hover { color:var(--accent-text); background:var(--accent-dim); }
.csp-ref > summary::-webkit-details-marker { display:none; }
.csp-ref > summary .csp-chevron { font-size:9px; transition:transform .2s ease; display:inline-block; color:var(--subtle); }
.csp-ref[open] > summary .csp-chevron { transform:rotate(90deg); }
.csp-ref > summary .csp-sum-icon { font-size:12px; opacity:0.5; }
.csp-ref > summary .csp-sum-label { font-weight:500; }
.csp-ref .csp-body { padding:14px 18px 16px 18px; background:var(--card); border-left:3px solid var(--accent);
  margin:0 12px 10px 12px; border-radius:var(--radius-sm); border-top:none; }
.csp-ref .csp-desc { color:var(--text); margin-bottom:10px; line-height:1.55; font-size:12px; }
.csp-ref .csp-grid { display:flex; flex-wrap:wrap; gap:6px; margin-bottom:10px; }
.csp-ref .csp-grid .csp-pill { display:inline-flex; align-items:center; gap:5px; padding:4px 10px;
  border-radius:20px; background:var(--bg); border:1px solid var(--border); font-size:10.5px; }
.csp-ref .csp-grid .csp-pill .csp-label { color:var(--muted); font-weight:400; }
.csp-ref .csp-grid .csp-pill .csp-val { color:var(--text-bright); font-weight:600; }
.csp-ref .csp-section-divider { height:1px; background:var(--border); margin:10px 0; }
.csp-ref .csp-av { margin-top:0; }
.csp-ref .csp-av-title { color:var(--muted); font-size:9.5px; text-transform:uppercase; letter-spacing:0.8px;
  font-weight:600; margin-bottom:6px; }
.csp-ref .csp-av-list { display:grid; grid-template-columns:repeat(auto-fill, minmax(240px,1fr)); gap:3px 16px; }
.csp-ref .csp-av-item { font-family:var(--mono); font-size:11px; color:var(--text); padding:3px 8px;
  border-radius:var(--radius-sm); background:var(--bg); display:flex; align-items:baseline; gap:6px; }
.csp-ref .csp-av-item .av-key { color:var(--accent-text); font-weight:700; min-width:16px; }
.csp-ref .csp-av-item .av-eq { color:var(--subtle); }
.csp-ref .csp-gp { margin-top:0; }
.csp-ref .csp-gp-title { color:var(--muted); font-size:9.5px; text-transform:uppercase; letter-spacing:0.8px;
  font-weight:600; margin-bottom:6px; }
.csp-ref .csp-gp-grid { display:grid; grid-template-columns:repeat(auto-fill, minmax(280px,1fr)); gap:4px 12px; }
.csp-ref .csp-gp-item { font-size:11px; color:var(--text); padding:3px 8px; border-radius:var(--radius-sm);
  background:var(--bg); display:flex; gap:6px; min-width:0; overflow:hidden; }
.csp-ref .csp-gp-item .csp-label { color:var(--muted); white-space:nowrap; min-width:100px; flex-shrink:0; font-weight:500; }
.csp-ref .csp-gp-item span:last-child { overflow-wrap:break-word; word-break:break-all; min-width:0; }
.csp-ref .csp-stale { color:var(--orange); font-size:10px; margin-top:10px; display:flex; align-items:center; gap:4px; }
.csp-ref .csp-none { color:var(--subtle); font-style:italic; padding:6px 14px; font-size:10.5px; }
.csp-db-warn { display:inline-flex; align-items:center; gap:4px; padding:4px 10px; border-radius:20px;
  background:var(--yellow-dim); color:var(--yellow); font-size:10px; font-weight:600; }

/* Badges */
.badge { display:inline-block; padding:2px 8px; border-radius:4px; font-size:10px; font-weight:600; }
.badge-enabled { background:var(--green-dim); color:var(--green); }
.badge-disabled { background:var(--red-dim); color:var(--red); }
.badge-conflict { background:var(--red-dim); color:var(--red); }
.badge-redundant { background:var(--yellow-dim); color:var(--yellow); }
.badge-notconfigured { background:rgba(113,113,122,0.15); color:var(--muted); }
.badge-installed { background:var(--green-dim); color:var(--green); }
.badge-failed { background:var(--red-dim); color:var(--red); }
.badge-pending { background:var(--yellow-dim); color:var(--yellow); }
.badge-scope { background:var(--accent-dim); color:var(--accent-text); }

/* Category chip */
.cat-chip { display:inline-block; padding:2px 8px; border-radius:4px; font-size:10px;
  background:rgba(168,85,247,0.18); color:var(--text); font-weight:500; }

/* Filter bar */
.filter-bar { padding:12px 14px; display:flex; gap:8px; flex-wrap:wrap; align-items:center;
  border-bottom:1px solid var(--border); }
.filter-bar select { background:var(--card); border:1px solid var(--border); border-radius:var(--radius-sm);
  padding:4px 8px; color:var(--text); font-size:11px; font-family:var(--font); outline:none; }
.filter-bar select:focus { border-color:var(--accent); }
.filter-bar .label { font-size:11px; color:var(--muted); }
.filter-bar .result-count { margin-left:auto; font-size:11px; color:var(--muted); }

/* Non-default highlight */
tr.non-default td:first-child { border-left:3px solid var(--orange); }

/* Sub-group heading */
.group-heading { background:var(--card); padding:10px 14px; font-size:12px; font-weight:600;
  color:var(--accent-text); border-bottom:1px solid var(--border); display:flex; align-items:center; gap:8px; }
.group-heading .cnt { color:var(--muted); font-weight:400; }

/* Empty state */
.empty { padding:32px; text-align:center; color:var(--muted); font-size:13px; }

/* Footer */
.footer { text-align:center; color:var(--subtle); font-size:11px; margin-top:48px;
  padding-top:16px; border-top:1px solid var(--border); }

/* Responsive */
@media (max-width:768px) {
  .toolbar { padding:10px 16px; flex-wrap:wrap; }
  .search-box { width:100%; }
  .container { padding:16px; }
  .stats { grid-template-columns:repeat(2,1fr); }
  .meta-bar { flex-direction:column; gap:4px; }
  td.val { max-width:180px; }
}
/* Print */
@media print {
  .toolbar { position:static; }
  .btn, .search-box, .filter-bar select { display:none; }
  body { background:#fff; color:#000; }
  .stat-card, .section-header, .section-body, th { background:#f5f5f5; border-color:#ddd; }
  .stat-num, .section-header .title { color:#000; }
  td { border-color:#ddd; }
  details { open:true; }
  details > .section-body { display:block !important; }
}
</style>
</head>
<body>
<div class="toolbar">
  <h1>&#x1F4CB; Policy Report</h1>
  <span class="domain-badge">$(& $enc $domain)</span>
  <span class="spacer"></span>
  <input type="text" class="search-box" id="globalSearch" placeholder="Search policies, values, categories..." />
  <button class="btn" onclick="expandAll()" title="Expand all sections">&#x25BC; Expand</button>
  <button class="btn" onclick="collapseAll()" title="Collapse all sections">&#x25B6; Collapse</button>
  <button class="btn btn-icon" id="themeToggle" onclick="toggleTheme()" title="Toggle dark/light mode">&#x263E;</button>
  <button class="btn" onclick="window.print()" title="Print report">&#x1F5A8;</button>
</div>

<div class="container">
<div class="meta-bar">
  <span>Generated: <strong>$genDate</strong></span>
  <span>Scan mode: <strong>$($Script:Prefs.ScanMode)</strong></span>
  <span>PolicyPilot <strong>v$($Script:AppVersion)</strong></span>
</div>

<div class="device-info">
  <div class="di-item"><span class="di-label">Computer Name</span><span class="di-value">$(& $enc $deviceName)</span></div>
  <div class="di-item"><span class="di-label">Domain</span><span class="di-value">$(& $enc $domain)</span></div>
  <div class="di-item"><span class="di-label">Logged-on User</span><span class="di-value">$(& $enc $userName)</span></div>
  <div class="di-item"><span class="di-label">Operating System</span><span class="di-value">$(& $enc $osInfo)</span></div>
  <div class="di-item"><span class="di-label">Scan Mode</span><span class="di-value">$(& $enc $scanMode)</span></div>
  <div class="di-item"><span class="di-label">Scan Time</span><span class="di-value">$scanTime</span></div>
  <div class="di-item"><span class="di-label">Loopback Mode</span><span class="di-value">$(if ($Script:ScanData.LoopbackMode) { & $enc $Script:ScanData.LoopbackMode } else { 'Not Configured' })</span></div>
"@)
    # Add MDM enrollment info if available
    if ($mdmInfo -and $mdmInfo.ProviderID) {
        [void]$html.Append(@"
  <div class="di-item"><span class="di-label">MDM Provider</span><span class="di-value">$(& $enc $mdmInfo.ProviderID)</span></div>
  <div class="di-item"><span class="di-label">Enrollment UPN</span><span class="di-value">$(& $enc $mdmInfo.EnrollmentUPN)</span></div>
  <div class="di-item"><span class="di-label">Enrollment Type</span><span class="di-value">$(& $enc $mdmInfo.EnrollmentType)</span></div>
  <div class="di-item"><span class="di-label">Enrollment State</span><span class="di-value">$(& $enc $mdmInfo.EnrollmentState)</span></div>
  <div class="di-item"><span class="di-label">AAD Tenant ID</span><span class="di-value" style="font-family:var(--font-mono);font-size:11px">$(& $enc $mdmInfo.AADTenantID)</span></div>
  <div class="di-item"><span class="di-label">Device MDM ID</span><span class="di-value" style="font-family:var(--font-mono);font-size:11px">$(& $enc $mdmInfo.EntDMID)</span></div>
"@)
    }
    [void]$html.Append(@"
</div>

<div class="stats">
  <div class="stat-card green"><div class="stat-num">$activeCount</div><div class="stat-label">Active GPOs / Areas</div></div>
  <div class="stat-card"><div class="stat-num">$totalSettings</div><div class="stat-label">Total Settings</div></div>
  <div class="stat-card red"><div class="stat-num">$conflictCount</div><div class="stat-label">Conflicts</div></div>
  <div class="stat-card yellow"><div class="stat-num">$redundantCount</div><div class="stat-label">Redundant</div></div>
  <div class="stat-card"><div class="stat-num">$unlinkedCount</div><div class="stat-label">Unlinked GPOs</div></div>
"@)

    if ($appCount -gt 0) {
        [void]$html.Append("<div class=`"stat-card accent`"><div class=`"stat-num`">$appCount</div><div class=`"stat-label`">Intune Apps</div></div>")
    }
    if ($notCfgCount -gt 0) {
        [void]$html.Append("<div class=`"stat-card`"><div class=`"stat-num`">$notCfgCount</div><div class=`"stat-label`">Not Configured (CSP)</div></div>")
    }

    [void]$html.Append('</div>') # close .stats

    # CSP database info for report enrichment
    $cspDb    = $Script:CspMetaKeys
    $cspDbAge = $Script:CspDbAge
    $cspStale = $cspDbAge -and $cspDbAge -gt 90
    if ($cspStale) {
        [void]$html.Append("<div style=`"margin:12px 0 0`"><span class=`"csp-db-warn`">&#x26A0; CSP reference database is $cspDbAge days old &mdash; run Build-CspDatabase.ps1 to refresh</span></div>")
    }

    # ── Section 1: All Settings ──
    $categories = @($settings | Group-Object Category | Sort-Object Count -Descending)
    $scopes = @($settings | ForEach-Object { $_.Scope } | Sort-Object -Unique)
    $states = @($settings | ForEach-Object { $_.State } | Sort-Object -Unique | Where-Object { $_ })

    [void]$html.Append(@"
<details class="section" open>
<summary class="section-header">
  <span class="icon">&#x2699;</span>
  <span class="title">All Policy Settings</span>
  <span class="count">$totalSettings</span>
  <span class="chevron">&#x25B6;</span>
</summary>
<div class="section-body">
<div class="filter-bar">
  <span class="label">Category:</span>
  <select id="fltCat" onchange="filterSettings()">
    <option value="">All Categories</option>
"@)
    foreach ($cat in $categories) {
        [void]$html.Append("<option value=`"$(& $enc $cat.Name)`">$(& $enc $cat.Name) ($($cat.Count))</option>")
    }
    [void]$html.Append('</select><span class="label">Scope:</span><select id="fltScope" onchange="filterSettings()"><option value="">All Scopes</option>')
    foreach ($sc in $scopes) {
        [void]$html.Append("<option value=`"$(& $enc $sc)`">$(& $enc $sc)</option>")
    }
    [void]$html.Append('</select><span class="label">State:</span><select id="fltState" onchange="filterSettings()"><option value="">All States</option>')
    foreach ($st in $states) {
        [void]$html.Append("<option value=`"$(& $enc $st)`">$(& $enc $st)</option>")
    }
    [void]$html.Append('</select><label style="display:flex;align-items:center;gap:4px;font-size:11px;color:var(--muted);cursor:pointer"><input type="checkbox" id="fltNonDefault" onchange="filterSettings()"/> Non-default only</label>')
    [void]$html.Append('<span class="result-count" id="settingsCount"></span></div>')

    # Grouped by GPO
    $grouped = $settings | Group-Object GPOName | Sort-Object Name
    [void]$html.Append('<div id="settingsContainer">')
    foreach ($grp in $grouped) {
        $gpoNameEnc = & $enc $grp.Name
        [void]$html.Append("<div class=`"group-heading`">$gpoNameEnc <span class=`"cnt`">($($grp.Count) settings)</span></div>")
        [void]$html.Append('<table class="settings-table"><thead><tr><th style="width:28%">Policy Name</th><th style="width:7%">State</th><th style="width:7%">Scope</th><th style="width:20%">Category</th><th style="width:25%">Value</th><th style="width:13%">Default</th></tr></thead><tbody>')
        foreach ($s in ($grp.Group | Sort-Object Category, PolicyName)) {
            $isNonDef = $s.DefaultValue -and $s.ValueData -and "$($s.ValueData)" -ne "$($s.DefaultValue)" -and "$($s.ValueData)" -notlike "$($s.DefaultValue) *"
            $rowClass = if ($isNonDef) { ' class="non-default"' } else { '' }
            $pnEnc  = & $enc $s.PolicyName
            $catEnc = & $enc $s.Category
            $v      = (& $enc $s.ValueData) -replace "`n", '<br>'
            # Wrap long numbered lists in a collapsible <details> element
            $lineCount = ($s.ValueData -split "`n").Count
            if ($lineCount -gt 5) {
                $firstLines = (($s.ValueData -split "`n") | Select-Object -First 3 | ForEach-Object { & $enc $_ }) -join '<br>'
                $v = "<details class=`"val-list`"><summary>$firstLines<br><em>… $lineCount items total (click to expand)</em></summary>$v</details>"
            }
            $dv     = & $enc $s.DefaultValue
            $stEnc  = & $enc $s.State
            $scEnc  = & $enc $s.Scope
            [void]$html.Append("<tr$rowClass data-cat=`"$(& $enc $s.Category)`" data-scope=`"$scEnc`" data-state=`"$stEnc`" data-nd=`"$(if ($isNonDef) {'1'} else {'0'})`"><td class=`"policy-name`">$pnEnc</td><td>$stEnc</td><td><span class=`"badge badge-scope`">$scEnc</span></td><td><span class=`"cat-chip`">$catEnc</span></td><td class=`"val`">$v</td><td class=`"val`" style=`"color:var(--subtle)`">$dv</td></tr>")

            # ── ADMX / CSP Reference row (collapsed enrichment panel) ──
            # Uses registry-key-based lookup from the ADMX + CSP databases loaded at startup.
            # Falls back to locale-dependent data with a "no metadata" label if not found.
            $regKey  = $s.RegistryKey
            # RegistryKey already contains "keyPath\valueName" — pass as-is
            $regVal  = ''
            $resolved = if ($regKey) { Resolve-PolicyFromRegistry -RegistryKey $regKey -ValueName '' -FallbackName $s.PolicyName } else { $null }

            if ($resolved -and $resolved.Source -in @('ADMX','CSP')) {
                $srcBadge = if ($resolved.Source -eq 'ADMX') { "&#x1F4D6; ADMX: $($resolved.AdmxFile)" } else { "&#x2601; CSP: $($resolved.CspPath)" }
                [void]$html.Append("<tr class=`"csp-row`"><td colspan=`"6`"><details class=`"csp-ref`"><summary><span class=`"csp-chevron`">&#x25B6;</span><span class=`"csp-sum-icon`">$(if ($resolved.Source -eq 'ADMX') {'&#x1F4D6;'} else {'&#x2601;'})</span> <span class=`"csp-sum-label`">$($resolved.Source) Reference:</span> $(& $enc $resolved.Name)</summary><div class=`"csp-body`">")
                if ($resolved.Desc) { [void]$html.Append("<div class=`"csp-desc`">$(& $enc $resolved.Desc)</div>") }
                [void]$html.Append('<div class="csp-grid">')
                [void]$html.Append("<span class=`"csp-pill`"><span class=`"csp-label`">Source</span><span class=`"csp-val`">$($resolved.Source)</span></span>")
                if ($resolved.AdmxFile) { [void]$html.Append("<span class=`"csp-pill`"><span class=`"csp-label`">ADMX</span><span class=`"csp-val`">$(& $enc $resolved.AdmxFile)</span></span>") }
                if ($resolved.CspPath)  { [void]$html.Append("<span class=`"csp-pill`"><span class=`"csp-label`">CSP</span><span class=`"csp-val`">$(& $enc $resolved.CspPath)</span></span>") }
                if ($resolved.Category) { [void]$html.Append("<span class=`"csp-pill`"><span class=`"csp-label`">Category</span><span class=`"csp-val`">$(& $enc $resolved.Category)</span></span>") }
                if ($regKey)            { [void]$html.Append("<span class=`"csp-pill`"><span class=`"csp-label`">Registry</span><span class=`"csp-val`" style=`"font-family:var(--mono);font-size:10px`">$(& $enc $regKey)\$(& $enc $regVal)</span></span>") }
                [void]$html.Append('</div>')
                # Allowed values from ADMX Elements or CSP AllowedValues
                if ($resolved.AllowedValues) {
                    [void]$html.Append('<div class="csp-section-divider"></div><div class="csp-av"><div class="csp-av-title">Allowed Values</div><div class="csp-av-list">')
                    $avObj = $resolved.AllowedValues
                    if ($avObj -is [array]) {
                        foreach ($elem in $avObj) {
                            if ($elem.EnumValues) {
                                $evProps = if ($elem.EnumValues -is [hashtable]) { $elem.EnumValues.GetEnumerator() } elseif ($elem.EnumValues.PSObject) { $elem.EnumValues.PSObject.Properties } else { @() }
                                foreach ($ev in $evProps) {
                                    $evK = if ($ev.Key) { $ev.Key } else { $ev.Name }
                                    [void]$html.Append("<div class=`"csp-av-item`"><span class=`"av-key`">$(& $enc $evK)</span><span class=`"av-eq`">=</span>$(& $enc $ev.Value)</div>")
                                }
                            } elseif ($elem.Type -and $elem.Id) {
                                $info = "[$($elem.Type)] $($elem.Id)"
                                if ($elem.MinValue) { $info += " (min: $($elem.MinValue))" }
                                if ($elem.MaxValue) { $info += " (max: $($elem.MaxValue))" }
                                [void]$html.Append("<div class=`"csp-av-item`">$(& $enc $info)</div>")
                            }
                        }
                    } elseif ($avObj.PSObject) {
                        foreach ($av in $avObj.PSObject.Properties) {
                            [void]$html.Append("<div class=`"csp-av-item`"><span class=`"av-key`">$(& $enc $av.Name)</span><span class=`"av-eq`">=</span>$(& $enc $av.Value)</div>")
                        }
                    }
                    [void]$html.Append('</div></div>')
                }
                if ($Script:AdmxDbAge -and $Script:AdmxDbAge -gt 90) { [void]$html.Append("<div class=`"csp-stale`">&#x26A0; ADMX database is $($Script:AdmxDbAge) days old - consider regenerating</div>") }
                [void]$html.Append('</div></details></td></tr>')
            } elseif ($regKey -and $Script:AdmxByReg.Count -eq 0) {
                # DB not loaded - show hint
                [void]$html.Append("<tr class=`"csp-row`"><td colspan=`"6`"><span class=`"csp-db-warn`">&#x26A0; ADMX metadata not available - run Build-AdmxDatabase.ps1 to enable policy name resolution</span></td></tr>")
            }
        }
        [void]$html.Append('</tbody></table>')
    }
    [void]$html.Append('</div></div></details>')

    # ── Section 2: GPO / Area Inventory ──
    [void]$html.Append(@"
<details class="section" open>
<summary class="section-header">
  <span class="icon">&#x1F4C1;</span>
  <span class="title">GPO / Area Inventory</span>
  <span class="count">$($gpos.Count)</span>
  <span class="chevron">&#x25B6;</span>
</summary>
<div class="section-body">
<table><thead><tr><th>GPO / Area Name</th><th>Status</th><th>Settings</th><th>Links</th><th>Enforced</th><th>Link Order</th><th>Security Filtering</th><th>WMI Filter</th><th>Scope Tags</th><th>Assignments</th><th>Created</th><th>Modified</th></tr></thead><tbody>
"@)
    foreach ($g in ($gpos | Sort-Object DisplayName)) {
        $nameEnc = & $enc $g.DisplayName
        $badgeClass = if ($g.Status -eq 'Disabled') { 'badge-disabled' } else { 'badge-enabled' }
        $linksEnc = & $enc $g.Links
        $enfBadge = if ($g.Enforced) { '<span class="badge badge-conflict">Enforced</span>' } else { '' }
        $loVal = if ($g.LinkOrder) { $g.LinkOrder } else { '' }
        $secFilt = & $enc $g.SecurityFiltering
        $wmiVal = & $enc $g.WmiFilter
        $wmiStatus = if ($g.WmiFilterStatus) { " ($($g.WmiFilterStatus))" } else { '' }
        $scopeVal = & $enc $g.ScopeTags
        $assignVal = & $enc $g.Assignments
        [void]$html.Append("<tr><td class=`"policy-name`">$nameEnc</td><td><span class=`"badge $badgeClass`">$($g.Status)</span></td><td>$($g.SettingCount)</td><td title=`"$linksEnc`">$($g.LinkCount)</td><td>$enfBadge</td><td>$loVal</td><td style=`"font-size:10px`">$secFilt</td><td>$wmiVal$wmiStatus</td><td>$scopeVal</td><td style=`"font-size:10px`">$assignVal</td><td style=`"color:var(--muted)`">$($g.CreationTime)</td><td style=`"color:var(--muted)`">$($g.ModificationTime)</td></tr>")
    }
    [void]$html.Append('</tbody></table></div></details>')

    # ── Section 3: Conflicts & Redundancies ──
    [void]$html.Append(@"
<details class="section"$(if ($conflicts.Count -gt 0) { ' open' } else { '' })>
<summary class="section-header">
  <span class="icon">&#x26A0;</span>
  <span class="title">Conflicts &amp; Redundancies</span>
  <span class="count">$($conflicts.Count)</span>
  <span class="chevron">&#x25B6;</span>
</summary>
<div class="section-body">
"@)
    if ($conflicts.Count -eq 0) {
        [void]$html.Append('<div class="empty">&#x2705; No conflicts or redundancies detected. Your configuration is clean.</div>')
    } else {
        [void]$html.Append('<table><thead><tr><th style="width:8%">Severity</th><th style="width:28%">Setting</th><th style="width:6%">Scope</th><th style="width:22%">Affected GPOs</th><th style="width:14%">Winner GPO</th><th style="width:22%">Values</th></tr></thead><tbody>')
        foreach ($c in $conflicts) {
            $sevBadge = if ($c.Severity -eq 'Conflict') { 'badge-conflict' } else { 'badge-redundant' }
            $settEnc = & $enc $c.SettingKey
            $regEnc  = if ($c.RegistryPath -and $c.RegistryPath -ne $c.SettingKey) { "<br><span style=`"font-size:10px;color:var(--muted);font-family:var(--mono)`">$(& $enc $c.RegistryPath)</span>" } else { '' }
            $srcBadge = if ($c.AdmxSource) { " <span class=`"badge badge-scope`" style=`"font-size:9px`">$($c.AdmxSource)</span>" } else { '' }
            $gpoEnc  = & $enc $c.GPONames
            $valEnc  = & $enc $c.Values
            $winEnc  = & $enc $c.WinnerGPO
            [void]$html.Append("<tr><td><span class=`"badge $sevBadge`">$($c.Severity)</span></td><td class=`"policy-name`">$settEnc$srcBadge$regEnc</td><td>$($c.Scope)</td><td>$gpoEnc</td><td style=`"font-weight:600;color:var(--green)`">$winEnc</td><td class=`"val`">$valEnc</td></tr>")
        }
        [void]$html.Append('</tbody></table>')
    }
    [void]$html.Append('</div></details>')

    # ── Section 4: Intune Apps (if any) ──
    if ($appCount -gt 0) {
        $installedCnt = @($apps | Where-Object { $_.InstallState -eq 'Installed' }).Count
        $failedCnt    = @($apps | Where-Object { $_.InstallState -eq 'Failed' }).Count
        [void]$html.Append(@"
<details class="section" open>
<summary class="section-header">
  <span class="icon">&#x1F4E6;</span>
  <span class="title">Intune Managed Apps</span>
  <span class="count">$appCount apps &middot; $installedCnt installed &middot; $failedCnt failed</span>
  <span class="chevron">&#x25B6;</span>
</summary>
<div class="section-body">
<table><thead><tr><th>Application</th><th>Type</th><th>Version</th><th>Enforcement</th><th>Install State</th></tr></thead><tbody>
"@)
        foreach ($a in ($apps | Sort-Object AppName)) {
            $appNameEnc = & $enc $a.AppName
            $typeEnc = & $enc $a.AppType
            $verEnc = & $enc $a.AppVersion
            $statusBadge = switch ($a.InstallState) {
                'Installed' { 'badge-installed' }
                'Failed'    { 'badge-failed' }
                default     { 'badge-pending' }
            }
            [void]$html.Append("<tr><td class=`"policy-name`">$appNameEnc</td><td>$typeEnc</td><td class=`"val`">$verEnc</td><td>$(& $enc $a.EnforcementState)</td><td><span class=`"badge $statusBadge`">$(& $enc $a.InstallState)</span></td></tr>")
        }
        [void]$html.Append('</tbody></table></div></details>')
    }

    # ── Section 5: MDM Enrollment & Certificates ──
    if ($mdmInfo -and $mdmInfo.ProviderID) {
        $certCount = if ($mdmInfo.MdmDiag.Certificates) { $mdmInfo.MdmDiag.Certificates.Count } else { 0 }
        $mgdPolCount = $mdmInfo.MdmDiag.ManagedPolicies
        [void]$html.Append(@"
<details class="section" open>
<summary class="section-header">
  <span class="icon">&#x1F511;</span>
  <span class="title">MDM Enrollment &amp; Certificates</span>
  <span class="count">$certCount certs &middot; $mgdPolCount managed policies</span>
  <span class="chevron">&#x25B6;</span>
</summary>
<div class="section-body">
<div class="device-info" style="border:none;margin:0;padding:12px 16px;">
  <div class="di-item"><span class="di-label">MDM Provider</span><span class="di-value">$(& $enc $mdmInfo.ProviderID)</span></div>
  <div class="di-item"><span class="di-label">Enrollment UPN</span><span class="di-value">$(& $enc $mdmInfo.EnrollmentUPN)</span></div>
  <div class="di-item"><span class="di-label">Enrollment Type</span><span class="di-value">$(& $enc $mdmInfo.EnrollmentType)</span></div>
  <div class="di-item"><span class="di-label">Enrollment State</span><span class="di-value" style="color:var(--green)">$(& $enc $mdmInfo.EnrollmentState)</span></div>
  <div class="di-item"><span class="di-label">AAD Tenant ID</span><span class="di-value" style="font-family:var(--font-mono);font-size:11px">$(& $enc $mdmInfo.AADTenantID)</span></div>
  <div class="di-item"><span class="di-label">Device MDM ID</span><span class="di-value" style="font-family:var(--font-mono);font-size:11px">$(& $enc $mdmInfo.EntDMID)</span></div>
  <div class="di-item"><span class="di-label">Enrollment GUID</span><span class="di-value" style="font-family:var(--font-mono);font-size:11px">$(& $enc $mdmInfo.EnrollmentGUID)</span></div>
  <div class="di-item"><span class="di-label">Cert Renewal</span><span class="di-value">$(& $enc $mdmInfo.CertRenewTime)</span></div>
</div>
"@)
        # LAPS subsection
        $laps = $mdmInfo.MdmDiag.LAPS
        if ($laps -and $laps.BackupDirectory -and $laps.BackupDirectory -ne 'Not Configured') {
            [void]$html.Append(@"
<div class="device-info" style="border-top:1px solid var(--border);margin:12px 0 0;padding:12px 16px 0;">
  <div style="font-weight:600;font-size:12px;margin-bottom:8px;">&#x1F512; LAPS (Local Admin Password Solution)</div>
  <div class="di-item"><span class="di-label">Backup Directory</span><span class="di-value">$(& `$enc `$laps.BackupDirectory)</span></div>
  <div class="di-item"><span class="di-label">Password Age</span><span class="di-value">$(& `$enc `$laps.PasswordAgeDays)</span></div>
  <div class="di-item"><span class="di-label">Complexity</span><span class="di-value">$(& `$enc `$laps.PasswordComplexity)</span></div>
  <div class="di-item"><span class="di-label">Password Length</span><span class="di-value">$(& `$enc `$laps.PasswordLength)</span></div>
  <div class="di-item"><span class="di-label">Post-auth Action</span><span class="di-value">$(& `$enc `$laps.PostAuthActions)</span></div>
  <div class="di-item"><span class="di-label">Post-auth Reset Delay</span><span class="di-value">$(& `$enc `$laps.PostAuthResetDelay)</span></div>
  <div class="di-item"><span class="di-label">Auto-manage Enabled</span><span class="di-value">$(& `$enc `$laps.AutoManageEnabled)</span></div>
  <div class="di-item"><span class="di-label">Auto-manage Target</span><span class="di-value">$(& `$enc `$laps.AutoManageTarget)</span></div>
  <div class="di-item"><span class="di-label">Last Password Rotation</span><span class="di-value">$(& `$enc `$laps.LocalLastPasswordUpdate)</span></div>
  <div class="di-item"><span class="di-label">Azure Password Expiry</span><span class="di-value">$(& `$enc `$laps.LocalAzurePasswordExpiry)</span></div>
  <div class="di-item"><span class="di-label">Managed Account</span><span class="di-value">$(& `$enc `$laps.LocalManagedAccountName)</span></div>
</div>
"@)
        }
        if ($certCount -gt 0) {
            $hasCertDetail = ($mdmInfo.MdmDiag.Certificates[0].PSObject -ne $null -and $mdmInfo.MdmDiag.Certificates[0].Keys -contains 'Store')
            if ($hasCertDetail) {
                [void]$html.Append('<table><thead><tr><th>Store</th><th>Issued To</th><th>Issued By</th><th>Valid From</th><th>Valid To</th><th>Expires (days)</th><th>Thumbprint</th></tr></thead><tbody>')
                foreach ($cert in $mdmInfo.MdmDiag.Certificates) {
                    $expStyle = if ($cert.ExpireDays -ne '' -and [int]$cert.ExpireDays -le 30) { ' style="color:var(--red);font-weight:600"' } elseif ($cert.ExpireDays -ne '' -and [int]$cert.ExpireDays -le 90) { ' style="color:var(--yellow)"' } else { '' }
                    [void]$html.Append("<tr><td>$(& $enc $cert.Store)</td><td class=`"policy-name`">$(& $enc $cert.IssuedTo)</td><td>$(& $enc $cert.IssuedBy)</td><td>$(& $enc $cert.ValidFrom)</td><td$expStyle>$(& $enc $cert.ValidTo)</td><td$expStyle>$($cert.ExpireDays)</td><td style=`"font-family:var(--mono);font-size:10px`">$(& $enc $cert.Thumbprint)</td></tr>")
                }
            } else {
                [void]$html.Append('<table><thead><tr><th>Issued To</th><th>Issued By</th><th>Expiration</th><th>Purpose</th></tr></thead><tbody>')
                foreach ($cert in $mdmInfo.MdmDiag.Certificates) {
                    [void]$html.Append("<tr><td class=`"policy-name`">$(& $enc $cert.IssuedTo)</td><td>$(& $enc $cert.IssuedBy)</td><td>$(& $enc $cert.Expiration)</td><td>$(& $enc $cert.Purpose)</td></tr>")
                }
            }
            [void]$html.Append('</tbody></table>')
        }
        [void]$html.Append('</div></details>')
    }

    # ── Section 5b: Compliance & App Summary ──
    if ($mdmInfo -and $mdmInfo.MdmDiag.Compliance -and $mdmInfo.MdmDiag.Compliance.Status -ne 'N/A') {
        $comp = $mdmInfo.MdmDiag.Compliance
        $appSum = $mdmInfo.MdmDiag.AppSummary
        $compColor = switch ($comp.Status) { 'Compliant' { 'var(--green)' } 'Non-compliant' { 'var(--red)' } 'At Risk' { 'var(--yellow)' } default { 'var(--muted)' } }
        [void]$html.Append(@"
<details class="section" open>
<summary class="section-header">
  <span class="icon" style="color:$compColor">&#x2705;</span>
  <span class="title">Compliance &amp; App Install Summary</span>
  <span class="count" style="background:$(if($comp.Status -eq 'Compliant'){'rgba(34,197,94,0.12)'}else{'rgba(245,158,11,0.12)'});color:$compColor">$($comp.Status)</span>
  <span class="chevron">&#x25B6;</span>
</summary>
<div class="section-body">
<div class="device-info" style="border:none;margin:0;padding:12px 16px;">
  <div class="di-item"><span class="di-label">Overall Status</span><span class="di-value" style="color:$compColor;font-weight:600">$(& $enc $comp.Status)</span></div>
  <div class="di-item"><span class="di-label">Configured Policies</span><span class="di-value">$($comp.ConfiguredPolicies)</span></div>
"@)
        if ($appSum -and $appSum.Total -gt 0) {
            [void]$html.Append(@"
  <div class="di-item"><span class="di-label">App Success Rate</span><span class="di-value">$($appSum.SuccessRate)%</span></div>
  <div class="di-item"><span class="di-label">Apps Installed</span><span class="di-value" style="color:var(--green)">$($appSum.Installed) / $($appSum.Total)</span></div>
  <div class="di-item"><span class="di-label">Apps Failed</span><span class="di-value"$(if($appSum.Failed -gt 0){' style="color:var(--red);font-weight:600"'}else{''})>$($appSum.Failed)</span></div>
  <div class="di-item"><span class="di-label">Apps Pending</span><span class="di-value"$(if($appSum.Pending -gt 0){' style="color:var(--yellow)"'}else{''})>$($appSum.Pending)</span></div>
"@)
        }
        if ($comp.Issues -and $comp.Issues.Count -gt 0) {
            [void]$html.Append('<div style="margin-top:8px;padding:8px 12px;background:rgba(245,158,11,0.08);border-radius:6px;border-left:3px solid var(--yellow)">')
            foreach ($issue in $comp.Issues) {
                [void]$html.Append("<div style=`"font-size:12px;color:var(--yellow);margin:2px 0`">&#x26A0; $(& $enc $issue)</div>")
            }
            [void]$html.Append('</div>')
        }
        [void]$html.Append('</div></div></details>')
    }

    # ── Section 5c: Script Policies (Proactive Remediations) ──
    if ($mdmInfo -and $mdmInfo.MdmDiag.ScriptPolicies -and $mdmInfo.MdmDiag.ScriptPolicies.Count -gt 0) {
        $spList = $mdmInfo.MdmDiag.ScriptPolicies
        $spFailed = @($spList | Where-Object { $_.Result -match 'Fail' }).Count
        [void]$html.Append(@"
<details class="section">
<summary class="section-header">
  <span class="icon">&#x1F4DD;</span>
  <span class="title">Script Policies (Proactive Remediations)</span>
  <span class="count">$($spList.Count) scripts &middot; $spFailed failed</span>
  <span class="chevron">&#x25B6;</span>
</summary>
<div class="section-body">
<table><thead><tr><th>Script Name</th><th>Type</th><th>Result</th><th>Exit Code</th></tr></thead><tbody>
"@)
        foreach ($sp in ($spList | Sort-Object ScriptName)) {
            $spNameEnc = if ($sp.ScriptName) { & $enc $sp.ScriptName } else { & $enc $sp.PolicyId }
            $resBadge = if ($sp.Result -match 'Success') { 'badge-installed' } elseif ($sp.Result -match 'Fail') { 'badge-failed' } else { 'badge-pending' }
            [void]$html.Append("<tr><td class=`"policy-name`">$spNameEnc</td><td>$(& $enc $sp.ScriptType)</td><td><span class=`"badge $resBadge`">$(& $enc $sp.Result)</span></td><td class=`"val`">$($sp.ExitCode)</td></tr>")
        }
        [void]$html.Append('</tbody></table></div></details>')
    }

    # ── Section 5d: Config Profile Status ──
    if ($mdmInfo -and $mdmInfo.MdmDiag.ConfigProfiles -and $mdmInfo.MdmDiag.ConfigProfiles.Count -gt 0) {
        $cpList = $mdmInfo.MdmDiag.ConfigProfiles
        $cpErrors = @($cpList | Where-Object { $_.Status -match 'Error' }).Count
        [void]$html.Append(@"
<details class="section">
<summary class="section-header">
  <span class="icon">&#x2699;</span>
  <span class="title">Configuration Profile Status</span>
  <span class="count">$($cpList.Count) profiles &middot; $cpErrors errors</span>
  <span class="chevron">&#x25B6;</span>
</summary>
<div class="section-body">
<table><thead><tr><th>Profile</th><th>Status</th><th>Source</th></tr></thead><tbody>
"@)
        foreach ($cp in ($cpList | Sort-Object Name)) {
            $cpNameEnc = & $enc $cp.Name
            $cpStatus = & $enc $cp.Status
            $cpBadge = if ($cp.Status -match 'Error') { 'badge-failed' } elseif ($cp.Status -match 'Applied|Success') { 'badge-installed' } else { 'badge-pending' }
            [void]$html.Append("<tr><td class=`"policy-name`">$cpNameEnc</td><td><span class=`"badge $cpBadge`">$cpStatus</span></td><td>$(& $enc $cp.Source)</td></tr>")
        }
        [void]$html.Append('</tbody></table></div></details>')
    }

    # ── Section 5e: Provisioning Packages (.ppkg) ──
    if ($mdmInfo -and $mdmInfo.MdmDiag.ProvisioningPackages -and $mdmInfo.MdmDiag.ProvisioningPackages.Count -gt 0) {
        $ppkgList = $mdmInfo.MdmDiag.ProvisioningPackages
        $ppkgFailCount = @($ppkgList | Where-Object { $_.TotalFailures -gt 0 }).Count
        $ppkgOpen = if ($ppkgFailCount -gt 0) { ' open' } else { '' }
        [void]$html.Append(@"
<details class="section"$ppkgOpen>
<summary class="section-header">
  <span class="icon">&#x1F4E6;</span>
  <span class="title">Provisioning Packages (.ppkg)</span>
  <span class="count">$($ppkgList.Count) packages$(if($ppkgFailCount -gt 0){" &middot; <span style=`"color:var(--red)`">$ppkgFailCount with failures</span>"})</span>
  <span class="chevron">&#x25B6;</span>
</summary>
<div class="section-body">
<p style="font-size:11px;color:var(--muted);margin:0 0 12px;padding:0 14px;">Provisioning packages (.ppkg) apply configuration at imaging/OOBE time. Settings enforced by .ppkg (e.g. Secured-core VBS/HVCI) can override or conflict with MDM/GPO policies &mdash; check for conflicts if policies appear to revert.</p>
<table><thead><tr><th style="width:22%">Package Name</th><th style="width:22%">Description</th><th style="width:12%">Owner</th><th style="width:6%">Ver</th><th style="width:22%">Package ID</th><th style="width:6%">Settings</th><th style="width:10%">Status</th></tr></thead><tbody>
"@)
        foreach ($ppkg in ($ppkgList | Sort-Object FileName)) {
            $pkgNameEnc = if ($ppkg.PackageName) { & $enc $ppkg.PackageName } else { & $enc $ppkg.FileName }
            $descEnc = if ($ppkg.FriendlyName -and $ppkg.FriendlyName -ne $ppkg.PackageName) { & $enc $ppkg.FriendlyName } elseif ($ppkg.FriendlyName) { '' } else { '<span style="color:var(--muted);font-style:italic">Unknown package</span>' }
            $ownerEnc = if ($ppkg.Owner) { & $enc $ppkg.Owner } else { '<span style="color:var(--muted)">—</span>' }
            $verEnc = if ($ppkg.Version) { & $enc $ppkg.Version } else { '<span style="color:var(--muted)">—</span>' }
            $idShort = & $enc $ppkg.PackageId
            $xmlCount = $ppkg.XMLEntries.Count
            $statusBadge = if ($ppkg.TotalFailures -gt 0) { 'badge-failed' } else { 'badge-installed' }
            $statusText = if ($ppkg.TotalFailures -gt 0) { "$($ppkg.TotalFailures) failures" } else { 'OK' }
            [void]$html.Append("<tr><td class=`"policy-name`">$pkgNameEnc</td><td style=`"font-size:11px`">$descEnc</td><td>$ownerEnc</td><td class=`"val`">$verEnc</td><td style=`"font-family:var(--mono);font-size:10px`">$idShort</td><td>$xmlCount</td><td><span class=`"badge $statusBadge`">$statusText</span></td></tr>")
            # Sub-rows for provisioned XML settings
            if ($xmlCount -gt 0) {
                [void]$html.Append('<tr class="csp-row"><td colspan="7"><details class="csp-ref"><summary><span class="csp-chevron">&#x25B6;</span> <span class="csp-sum-icon">&#x1F4CB;</span> <span class="csp-sum-label">Provisioned settings')
                [void]$html.Append(" ($xmlCount)</span></summary><div class=`"csp-body`">")
                [void]$html.Append('<table style="width:100%;margin:0"><thead><tr><th>Setting (XMLName)</th><th>Area</th><th>Message</th><th>Result</th><th>Failures</th></tr></thead><tbody>')
                foreach ($xe in $ppkg.XMLEntries) {
                    $xName = & $enc $xe.XMLName
                    $xArea = & $enc $xe.Area
                    $xMsg = & $enc $xe.Message
                    $xRes = $xe.LastResult
                    $xFail = $xe.NumberOfFailures
                    $xStyle = if ([int]$xFail -gt 0) { ' style="color:var(--red);font-weight:600"' } else { '' }
                    [void]$html.Append("<tr><td class=`"policy-name`">$xName</td><td>$xArea</td><td style=`"font-size:11px`">$xMsg</td><td class=`"val`">$xRes</td><td$xStyle>$xFail</td></tr>")
                }
                [void]$html.Append('</tbody></table></div></details></td></tr>')
            }
        }
        [void]$html.Append('</tbody></table></div></details>')
    }

    # ── Section 5f: Enrollment Issues ──
    if ($mdmInfo -and $mdmInfo.MdmDiag.EnrollmentIssues -and $mdmInfo.MdmDiag.EnrollmentIssues.Count -gt 0) {
        $eiList = $mdmInfo.MdmDiag.EnrollmentIssues
        [void]$html.Append(@"
<details class="section" open>
<summary class="section-header">
  <span class="icon" style="color:var(--red)">&#x26D4;</span>
  <span class="title">Enrollment Issues</span>
  <span class="count" style="background:rgba(239,68,68,0.12);color:var(--red)">$($eiList.Count)</span>
  <span class="chevron">&#x25B6;</span>
</summary>
<div class="section-body">
<table><thead><tr><th>Provider</th><th>Issue</th><th>Severity</th><th>Error Code</th></tr></thead><tbody>
"@)
        foreach ($ei in $eiList) {
            $sevBadge = if ($ei.Severity -eq 'Error') { 'badge-failed' } else { 'badge-pending' }
            $errCode = if ($ei.ErrorCode) { & $enc $ei.ErrorCode } else { '' }
            [void]$html.Append("<tr><td>$(& $enc $ei.Provider)</td><td class=`"policy-name`">$(& $enc $ei.Issue)</td><td><span class=`"badge $sevBadge`">$(& $enc $ei.Severity)</span></td><td class=`"val`">$errCode</td></tr>")
        }
        [void]$html.Append('</tbody></table></div></details>')
    }

    # ── Section 6: Unlinked GPOs ──
    $unlinked = @($gpos | Where-Object { -not $_.IsLinked })
    if ($unlinked.Count -gt 0) {
        [void]$html.Append(@"
<details class="section">
<summary class="section-header">
  <span class="icon" style="color:var(--yellow)">&#x26A0;</span>
  <span class="title">Unlinked GPOs</span>
  <span class="count" style="background:var(--yellow-dim);color:var(--yellow)">$($unlinked.Count)</span>
  <span class="chevron">&#x25B6;</span>
</summary>
<div class="section-body">
<table><thead><tr><th style="width:40%">GPO Name</th><th style="width:12%">Status</th><th style="width:10%">Settings</th><th style="width:38%">Modified</th></tr></thead><tbody>
"@)
        foreach ($u in ($unlinked | Sort-Object DisplayName)) {
            $nameEnc = & $enc $u.DisplayName
            [void]$html.Append("<tr><td class=`"policy-name`">$nameEnc</td><td>$($u.Status)</td><td>$($u.SettingCount)</td><td style=`"color:var(--muted)`">$($u.ModificationTime)</td></tr>")
        }
        [void]$html.Append('</tbody></table></div></details>')
    }

    # ── Section 7: Not Configured (CSP defaults) ──
    if ($notCfgCount -gt 0) {
        [void]$html.Append(@"
<details class="section">
<summary class="section-header">
  <span class="icon" style="color:var(--muted)">&#x2B55;</span>
  <span class="title">Not Configured (CSP Defaults)</span>
  <span class="count" style="background:rgba(113,113,122,0.12);color:var(--muted)">$notCfgCount</span>
  <span class="chevron">&#x25B6;</span>
</summary>
<div class="section-body">
<table><thead><tr><th>Policy Name</th><th>Category</th><th>Default Value</th><th>Scope</th></tr></thead><tbody>
"@)
        foreach ($nc in ($notCfg | Sort-Object Category, PolicyName)) {
            [void]$html.Append("<tr><td class=`"policy-name`">$(& $enc $nc.PolicyName)</td><td><span class=`"cat-chip`">$(& $enc $nc.Category)</span></td><td class=`"val`">$(& $enc $nc.DefaultValue)</td><td><span class=`"badge badge-scope`">$(& $enc $nc.Scope)</span></td></tr>")
        }
        [void]$html.Append('</tbody></table></div></details>')
    }

    # ── Settings by Category summary ──
    [void]$html.Append(@"
<details class="section">
<summary class="section-header">
  <span class="icon">&#x1F4CA;</span>
  <span class="title">Settings by Category Breakdown</span>
  <span class="count">$($categories.Count) categories</span>
  <span class="chevron">&#x25B6;</span>
</summary>
<div class="section-body">
<table><thead><tr><th>Category</th><th>Settings Count</th><th style="width:60%">Distribution</th></tr></thead><tbody>
"@)
    foreach ($cat in $categories) {
        $pct = [math]::Round(($cat.Count / [math]::Max($totalSettings,1)) * 100, 1)
        $catEnc = & $enc $cat.Name
        [void]$html.Append("<tr><td><span class=`"cat-chip`">$catEnc</span></td><td>$($cat.Count)</td><td><div style=`"background:var(--border);border-radius:4px;height:18px;overflow:hidden`"><div style=`"background:var(--accent);height:100%;width:${pct}%;border-radius:4px;min-width:2px`"></div></div><span style=`"font-size:10px;color:var(--muted)`">${pct}%</span></td></tr>")
    }
    [void]$html.Append('</tbody></table></div></details>')

    # ── Footer + JS ──
    [void]$html.Append(@"

<div class="footer">
  Generated by <strong>PolicyPilot v$($Script:AppVersion)</strong> on $genDate &middot;
  Scan mode: $($Script:Prefs.ScanMode) &middot; Domain: $(& $enc $domain)
</div>
</div>

<script>
function toggleTheme() {
  const html = document.documentElement;
  const isDark = html.getAttribute('data-theme') === 'dark';
  html.setAttribute('data-theme', isDark ? 'light' : 'dark');
  document.getElementById('themeToggle').textContent = isDark ? '\u2600' : '\u263E';
  try { localStorage.setItem('gpo-theme', isDark ? 'light' : 'dark'); } catch(e) {}
}
(function() {
  try {
    const saved = localStorage.getItem('gpo-theme');
    if (saved === 'light') { document.documentElement.setAttribute('data-theme','light'); document.getElementById('themeToggle').textContent = '\u2600'; }
  } catch(e) {}
})();

function expandAll() { document.querySelectorAll('details.section').forEach(d => d.open = true); }
function collapseAll() { document.querySelectorAll('details.section').forEach(d => d.open = false); }

// Global search
document.getElementById('globalSearch').addEventListener('input', function() {
  const q = this.value.toLowerCase().trim();
  document.querySelectorAll('table tbody tr').forEach(function(row) {
    if (!q) { row.style.display = ''; return; }
    row.style.display = row.textContent.toLowerCase().includes(q) ? '' : 'none';
  });
  // Auto-expand sections with matches
  if (q) {
    document.querySelectorAll('details.section').forEach(function(d) {
      const hasVisible = d.querySelector('tbody tr:not([style*="display: none"])');
      if (hasVisible) d.open = true;
    });
  }
  updateSettingsCount();
});

// Settings-specific filters
function filterSettings() {
  const cat = document.getElementById('fltCat').value;
  const scope = document.getElementById('fltScope').value;
  const state = document.getElementById('fltState').value;
  const ndOnly = document.getElementById('fltNonDefault').checked;

  document.querySelectorAll('#settingsContainer tr[data-cat]').forEach(function(row) {
    let show = true;
    if (cat && row.dataset.cat !== cat) show = false;
    if (scope && row.dataset.scope !== scope) show = false;
    if (state && row.dataset.state !== state) show = false;
    if (ndOnly && row.dataset.nd !== '1') show = false;
    row.style.display = show ? '' : 'none';
    // Toggle associated CSP reference row
    var next = row.nextElementSibling;
    if (next && next.classList.contains('csp-row')) next.style.display = show ? '' : 'none';
  });

  // Hide empty group headings
  document.querySelectorAll('#settingsContainer .group-heading').forEach(function(gh) {
    const table = gh.nextElementSibling;
    if (!table) return;
    const visibleRows = table.querySelectorAll('tbody tr:not([style*="display: none"])');
    gh.style.display = visibleRows.length > 0 ? '' : 'none';
    table.style.display = visibleRows.length > 0 ? '' : 'none';
  });
  updateSettingsCount();
}

function updateSettingsCount() {
  const visible = document.querySelectorAll('#settingsContainer tr[data-cat]:not([style*="display: none"])').length;
  const total = document.querySelectorAll('#settingsContainer tr[data-cat]').length;
  const el = document.getElementById('settingsCount');
  if (el) el.textContent = visible === total ? total + ' settings' : visible + ' of ' + total + ' settings';
}
updateSettingsCount();
</script>
</body>
</html>
"@)

    return $html.ToString()
}


# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 15.5: NEW FEATURES - Detail Pane, Context Menus, Diff, Script Gen, etc.
# ═══════════════════════════════════════════════════════════════════════════════

# --- #4: OU Tree (build from GPO Links) ---
function Build-OUTreeText {
    if (-not $Script:ScanData) { return "No scan data" }
    $gpos = $Script:ScanData.GPOs
    $tree = @{}
    foreach ($gpo in $gpos) {
        $links = "$($gpo.Links)" -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        foreach ($ouPath in $links) {
            if (-not $tree.ContainsKey($ouPath)) { $tree[$ouPath] = [System.Collections.Generic.List[string]]::new() }
            [void]$tree[$ouPath].Add($gpo.DisplayName)
        }
    }
    if ($tree.Count -eq 0) { return "No OU link data available (links parsed from GPO data)" }
    $sb = [System.Text.StringBuilder]::new()
    $null = $sb.AppendLine("OU LINK HIERARCHY")
    $null = $sb.AppendLine("$([char]0x2500)" * 50)
    foreach ($ou in ($tree.Keys | Sort-Object)) {
        $parts = $ou -split '/' | Where-Object { $_ }
        $indent = '  ' * [math]::Max(0, $parts.Count - 1)
        $null = $sb.AppendLine("${indent}$([char]0x2514) $ou")
        foreach ($gpoName in ($tree[$ou] | Sort-Object)) {
            $null = $sb.AppendLine("${indent}    $([char]0x251C) $gpoName")
        }
    }
    $sb.ToString()
}

# --- #15: Multi-Domain Support ---
$Script:DomainList = [System.Collections.Generic.List[string]]::new()

function Get-DomainList {
    $domains = [System.Collections.Generic.List[string]]::new()
    $override = $Script:Prefs.DomainOverride
    if ($override) {
        # Support comma/semicolon-separated domain list
        $override -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ } | ForEach-Object { [void]$domains.Add($_) }
    }
    if ($domains.Count -eq 0) {
        # Use current domain
        try { [void]$domains.Add([System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name) } catch { }
    }
    $Script:DomainList = $domains
    $domains
}

# --- #5: Snapshot Diff ---
function Compare-Snapshots {
    param([hashtable]$Baseline, [hashtable]$Current)
    $diffs = [System.Collections.Generic.List[PSCustomObject]]::new()

    # Build lookup by composite key
    $baseSettings = @{}
    foreach ($s in $Baseline.Settings) {
        $key = "$($s.GPOName)|$($s.SettingKey)|$($s.Scope)"
        $baseSettings[$key] = $s
    }
    $currSettings = @{}
    foreach ($s in $Current.Settings) {
        $key = "$($s.GPOName)|$($s.SettingKey)|$($s.Scope)"
        $currSettings[$key] = $s
    }

    # Find added / modified
    foreach ($key in $currSettings.Keys) {
        if (-not $baseSettings.ContainsKey($key)) {
            $c = $currSettings[$key]
            [void]$diffs.Add([PSCustomObject]@{ Change='Added'; PolicyName=$c.PolicyName; Category=$c.Category; Scope=$c.Scope; GPOName=$c.GPOName; OldValue=''; NewValue=$c.ValueData })
        } elseif ("$($currSettings[$key].ValueData)" -ne "$($baseSettings[$key].ValueData)") {
            $c = $currSettings[$key]; $b = $baseSettings[$key]
            [void]$diffs.Add([PSCustomObject]@{ Change='Modified'; PolicyName=$c.PolicyName; Category=$c.Category; Scope=$c.Scope; GPOName=$c.GPOName; OldValue=$b.ValueData; NewValue=$c.ValueData })
        }
    }

    # Find removed
    foreach ($key in $baseSettings.Keys) {
        if (-not $currSettings.ContainsKey($key)) {
            $b = $baseSettings[$key]
            [void]$diffs.Add([PSCustomObject]@{ Change='Removed'; PolicyName=$b.PolicyName; Category=$b.Category; Scope=$b.Scope; GPOName=$b.GPOName; OldValue=$b.ValueData; NewValue='' })
        }
    }

    # GPO-level diffs
    $baseGPOs = @{}; foreach ($g in $Baseline.GPOs) { $baseGPOs[$g.DisplayName] = $g }
    $currGPOs = @{}; foreach ($g in $Current.GPOs) { $currGPOs[$g.DisplayName] = $g }
    foreach ($name in $currGPOs.Keys) {
        if (-not $baseGPOs.ContainsKey($name)) {
            [void]$diffs.Add([PSCustomObject]@{ Change='GPO Added'; PolicyName=$name; Category='GPO'; Scope=''; GPOName=$name; OldValue=''; NewValue='New' })
        }
    }
    foreach ($name in $baseGPOs.Keys) {
        if (-not $currGPOs.ContainsKey($name)) {
            [void]$diffs.Add([PSCustomObject]@{ Change='GPO Removed'; PolicyName=$name; Category='GPO'; Scope=''; GPOName=$name; OldValue='Existed'; NewValue='' })
        }
    }

    $diffs
}

# --- #24: Script Generation ---
function Export-AsScript {
    if (-not $Script:ScanData) {
        Show-Toast 'No Data' 'Run a scan first.' 'warning'
        return
    }
    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Filter = 'PowerShell scripts (*.ps1)|*.ps1'
    $dlg.FileName = "GPO_Recreate_$(Get-Date -Format 'yyyyMMdd_HHmmss').ps1"
    $dlg.InitialDirectory = $Script:ReportsDir
    if (-not $dlg.ShowDialog()) { return }

    $sb = [System.Text.StringBuilder]::new()
    $null = $sb.AppendLine('#Requires -Modules GroupPolicy')
    $null = $sb.AppendLine('# GPO Recreation Script')
    $null = $sb.AppendLine("# Generated by PolicyPilot on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    $null = $sb.AppendLine("# Source domain: $($Script:ScanData.Domain)")
    $null = $sb.AppendLine("# WARNING: Review before running. This script modifies Group Policy.")
    $null = $sb.AppendLine('')
    $null = $sb.AppendLine('param([switch]$WhatIf)')
    $null = $sb.AppendLine('')

    $grouped = $Script:ScanData.Settings | Group-Object GPOName
    foreach ($group in $grouped) {
        $gpoName = $group.Name
        $null = $sb.AppendLine("# --- GPO: $gpoName ---")
        $null = $sb.AppendLine("Write-Host `"Processing GPO: $gpoName`"")
        $null = $sb.AppendLine("`$gpo = Get-GPO -Name '$($gpoName -replace "'","''")' -ErrorAction SilentlyContinue")
        $null = $sb.AppendLine("if (-not `$gpo) { `$gpo = New-GPO -Name '$($gpoName -replace "'","''")' -WhatIf:`$WhatIf }")
        $null = $sb.AppendLine('')

        foreach ($s in $group.Group) {
            if ($s.RegistryKey -and $s.ValueData) {
                $regPath = $s.RegistryKey -replace '\\[^\\]+$', ''
                $valName = $s.RegistryKey -replace '^.*\\', ''
                $null = $sb.AppendLine("# $($s.PolicyName)")
                $null = $sb.AppendLine("if (-not `$WhatIf) { Set-GPRegistryValue -Name '$($gpoName -replace "'","''")' -Key '$regPath' -ValueName '$valName' -Value '$($s.ValueData -replace "'","''")' -Type String }")
                $null = $sb.AppendLine('')
            }
        }
    }

    [System.IO.File]::WriteAllText($dlg.FileName, $sb.ToString(), [System.Text.Encoding]::UTF8)
    Write-DebugLog "Script exported: $($dlg.FileName)" -Level SUCCESS
    Show-Toast 'Script Exported' "GPO recreation script saved" 'success'
}

# --- #16: Print via HTML in browser ---
function Invoke-PrintReport {
    if (-not $Script:ScanData) {
        Show-Toast 'No Data' 'Run a scan first.' 'warning'
        return
    }
    $tmpFile = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "PolicyPilot_Print_$(Get-Date -Format 'yyyyMMddHHmmss').html")
    # Reuse Export-Html logic to temp file
    $script:_printPath = $tmpFile
    Export-HtmlToPath $tmpFile
    Start-Process $tmpFile
    Write-DebugLog "Opened report in browser for printing" -Level SUCCESS
}

function Export-HtmlToPath([string]$Path) {
    $htmlContent = Build-HtmlReport
    [System.IO.File]::WriteAllText($Path, $htmlContent, [System.Text.Encoding]::UTF8)
}

# --- #25: Registry .reg Export ---
function Export-RegistryFile {
    if (-not $Script:ScanData) {
        Show-Toast 'No Data' 'Run a scan first.' 'warning'
        return
    }
    $regSettings = @($Script:ScanData.Settings | Where-Object { $_.RegistryKey })
    if ($regSettings.Count -eq 0) {
        Show-Toast 'No Registry Settings' 'No registry-based settings found in scan data.' 'info'
        return
    }
    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Filter = 'Registry files (*.reg)|*.reg'
    $dlg.FileName = "GPO_Registry_$(Get-Date -Format 'yyyyMMdd_HHmmss').reg"
    $dlg.InitialDirectory = $Script:ReportsDir
    if (-not $dlg.ShowDialog()) { return }

    $sb = [System.Text.StringBuilder]::new()
    $null = $sb.AppendLine('Windows Registry Editor Version 5.00')
    $null = $sb.AppendLine('')
    $null = $sb.AppendLine('; GPO Registry Export')
    $null = $sb.AppendLine("; Generated by PolicyPilot on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    $null = $sb.AppendLine("; Domain: $($Script:ScanData.Domain)")
    $null = $sb.AppendLine('; WARNING: Review carefully before importing.')
    $null = $sb.AppendLine('')

    $grouped = $regSettings | Group-Object { ($_.RegistryKey -replace '\\[^\\]+$','') }
    foreach ($g in ($grouped | Sort-Object Name)) {
        $null = $sb.AppendLine("[$($g.Name)]")
        foreach ($s in $g.Group) {
            $valName = $s.RegistryKey -replace '^.*\\',''
            $val = "$($s.ValueData)" -replace '\\','\\' -replace '"','\"'
            $null = $sb.AppendLine("`"$valName`"=`"$val`"")
        }
        $null = $sb.AppendLine('')
    }

    [System.IO.File]::WriteAllText($dlg.FileName, $sb.ToString(), [System.Text.Encoding]::UTF8)
    Write-DebugLog "Registry export: $($dlg.FileName)" -Level SUCCESS
    Show-Toast 'Registry Exported' "$($regSettings.Count) registry settings exported to .reg file" 'success'
}

# --- #27: Impact Simulator (simple remove-and-re-diff) ---
function Invoke-ImpactSimulation {
    param([string]$GPONameToRemove)
    if (-not $Script:ScanData -or -not $GPONameToRemove) { return $null }

    $remainingSettings = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($s in $Script:ScanData.Settings) {
        if ($s.GPOName -ne $GPONameToRemove) { [void]$remainingSettings.Add($s) }
    }
    $removedSettings = @($Script:ScanData.Settings | Where-Object { $_.GPOName -eq $GPONameToRemove })
    $originalConflicts = Find-Conflicts $Script:ScanData.Settings
    $newConflicts = Find-Conflicts $remainingSettings

    $resolvedConflicts = @($originalConflicts | Where-Object { $_.GPONames -like "*$GPONameToRemove*" }) |
        Where-Object { $_.SettingKey -notin @($newConflicts | ForEach-Object { $_.SettingKey }) }

    [PSCustomObject]@{
        GPORemoved        = $GPONameToRemove
        SettingsLost      = $removedSettings.Count
        SettingsRemaining = $remainingSettings.Count
        ConflictsBefore   = $originalConflicts.Count
        ConflictsAfter    = $newConflicts.Count
        ConflictsResolved = $resolvedConflicts.Count
        OrphanedSettings  = @($removedSettings | Where-Object { $_.SettingKey -notin @($remainingSettings | ForEach-Object { $_.SettingKey }) })
        Details           = $removedSettings
    }
}

# --- #12: Baseline Comparison Framework ---
$Script:ActiveBaseline = $null

function Import-Baseline {
    $dlg = [Microsoft.Win32.OpenFileDialog]::new()
    $dlg.Filter = 'JSON Baseline (*.json)|*.json'
    $dlg.Title = 'Import Security Baseline'
    $dlg.InitialDirectory = $Script:SnapshotDir
    if (-not $dlg.ShowDialog()) { return }

    try {
        $json = [System.IO.File]::ReadAllText($dlg.FileName) | ConvertFrom-Json
        $Script:ActiveBaseline = @{
            Name     = if ($json.Name) { $json.Name } else { [System.IO.Path]::GetFileNameWithoutExtension($dlg.FileName) }
            Settings = @{}
        }
        foreach ($entry in $json.Settings) {
            $Script:ActiveBaseline.Settings[$entry.SettingKey] = @{
                ExpectedValue = $entry.ExpectedValue
                Severity      = if ($entry.Severity) { $entry.Severity } else { 'Medium' }
                Reference     = if ($entry.Reference) { $entry.Reference } else { '' }
            }
        }
        Write-DebugLog "Baseline loaded: $($Script:ActiveBaseline.Name) ($($Script:ActiveBaseline.Settings.Count) settings)" -Level SUCCESS
        Show-Toast 'Baseline Loaded' "$($Script:ActiveBaseline.Name): $($Script:ActiveBaseline.Settings.Count) settings" 'success'

        if ($Script:ScanData) { Update-BaselineCompliance }
    } catch {
        Write-DebugLog "Baseline import error: $($_.Exception.Message)" -Level ERROR
        Show-Toast 'Import Failed' $_.Exception.Message 'error'
    }
}

function Update-BaselineCompliance {
    if (-not $Script:ActiveBaseline -or -not $Script:ScanData) { return }
    $pass = 0; $fail = 0; $missing = 0
    foreach ($key in $Script:ActiveBaseline.Settings.Keys) {
        $expected = $Script:ActiveBaseline.Settings[$key]
        $actual = $Script:ScanData.Settings | Where-Object { $_.SettingKey -eq $key } | Select-Object -First 1
        if (-not $actual) { $missing++ }
        elseif ("$($actual.ValueData)" -eq "$($expected.ExpectedValue)") { $pass++ }
        else { $fail++ }
    }
    $total = $Script:ActiveBaseline.Settings.Count
    $pct = if ($total -gt 0) { [math]::Round(($pass / $total) * 100, 1) } else { 0 }
    $ui.BaselineStatus.Text = "$($Script:ActiveBaseline.Name): ${pct}% compliant ($pass pass, $fail fail, $missing missing)"
    Write-DebugLog "Baseline compliance: $pct% ($pass/$total pass)" -Level INFO
}


# ── N7: Import MDMDiagReport.xml from another device ──
function Import-MdmXml {
    $dlg = [Microsoft.Win32.OpenFileDialog]::new()
    $dlg.Filter = 'MDM Diagnostic Report (*.xml)|*.xml|All Files (*.*)|*.*'
    $dlg.Title = 'Import MDMDiagReport.xml'
    if (-not $dlg.ShowDialog()) { return }

    try {
        $xmlContent = [System.IO.File]::ReadAllText($dlg.FileName)
        [xml]$mdmDoc = $xmlContent

        if (-not $mdmDoc.MDMEnterpriseDiagnosticsReport) {
            Show-Toast 'Invalid File' 'Not a valid MDMDiagReport.xml (missing MDMEnterpriseDiagnosticsReport root).' 'error'
            return
        }

        $importName = [System.IO.Path]::GetFileName($dlg.FileName)
        Write-DebugLog "Import-MdmXml: Parsing $importName ($([math]::Round((Get-Item $dlg.FileName).Length/1KB))KB)" -Level INFO

        $importInfo = @{ LAPS = @{}; PolicyMeta = @{}; Certificates = @(); AppSummary = @{}; Compliance = @{} }

        # ── LAPS ──
        $lapsNode = $mdmDoc.MDMEnterpriseDiagnosticsReport.LAPS
        if ($lapsNode) {
            $lapsCSP = $lapsNode.Laps_CSP_Policy
            $lapsLocal = $lapsNode.Laps_Local_State
            $BackupDirectoryMap = @{ '0'="Disabled (password won't be backed up)"; '1'='Backup to Microsoft Entra ID only'; '2'='Backup to Active Directory only' }
            $PasswordComplexityMap = @{ '1'='Large letters'; '2'='Large + small letters'; '3'='Letters + numbers'; '4'='Letters + numbers + special characters'; '5'='Letters + numbers + special (improved readability)'; '6'='Passphrase (long words)'; '7'='Passphrase (short words)'; '8'='Passphrase (short words, unique prefixes)' }
            $PostAuthActionMap = @{ '1'='Reset password'; '3'='Reset password + logoff'; '5'='Reset password + reboot'; '11'='Reset password + terminate processes' }
            $importInfo.LAPS = @{
                BackupDirectory      = if ($lapsCSP.BackupDirectory) { if ($BackupDirectoryMap["$($lapsCSP.BackupDirectory)"]) { $BackupDirectoryMap["$($lapsCSP.BackupDirectory)"] } else { "$($lapsCSP.BackupDirectory)" } } else { 'Not Configured' }
                PasswordAgeDays      = if ($lapsCSP.PasswordAgeDays) { "$($lapsCSP.PasswordAgeDays) days" } else { [char]0x2014 }
                PasswordComplexity   = if ($lapsCSP.PasswordComplexity) { if ($PasswordComplexityMap["$($lapsCSP.PasswordComplexity)"]) { $PasswordComplexityMap["$($lapsCSP.PasswordComplexity)"] } else { "$($lapsCSP.PasswordComplexity)" } } else { [char]0x2014 }
                PasswordLength       = if ($lapsCSP.PasswordLength) { "$($lapsCSP.PasswordLength)" } else { [char]0x2014 }
                PostAuthActions      = if ($lapsCSP.PostAuthenticationActions) { if ($PostAuthActionMap["$($lapsCSP.PostAuthenticationActions)"]) { $PostAuthActionMap["$($lapsCSP.PostAuthenticationActions)"] } else { "$($lapsCSP.PostAuthenticationActions)" } } else { [char]0x2014 }
                PostAuthResetDelay   = if ($lapsCSP.PostAuthenticationResetDelay) { "$($lapsCSP.PostAuthenticationResetDelay) hours" } else { [char]0x2014 }
                AutoManageEnabled    = if ($lapsCSP.AutomaticAccountManagementEnabled) { if ($lapsCSP.AutomaticAccountManagementEnabled -eq '1') { 'Yes' } else { 'No' } } else { [char]0x2014 }
                AutoManageTarget     = if ($lapsCSP.AutomaticAccountManagementTarget) { "$($lapsCSP.AutomaticAccountManagementTarget)" } else { [char]0x2014 }
            }
            if ($lapsLocal) {
                if ($lapsLocal.LastPasswordUpdateTime) {
                    try { $ftVal = [long]$lapsLocal.LastPasswordUpdateTime; $importInfo.LAPS.LocalLastPasswordUpdate = [datetime]::FromFileTimeUtc($ftVal).ToString('yyyy-MM-dd HH:mm UTC') } catch { $importInfo.LAPS.LocalLastPasswordUpdate = "$($lapsLocal.LastPasswordUpdateTime)" }
                }
                if ($lapsLocal.AzurePasswordExpiryTime) {
                    try { $ftVal = [long]$lapsLocal.AzurePasswordExpiryTime; $importInfo.LAPS.LocalAzurePasswordExpiry = [datetime]::FromFileTimeUtc($ftVal).ToString('yyyy-MM-dd HH:mm UTC') } catch { $importInfo.LAPS.LocalAzurePasswordExpiry = "$($lapsLocal.AzurePasswordExpiryTime)" }
                }
                if ($lapsLocal.ManagedAccountName) { $importInfo.LAPS.LocalManagedAccountName = "$($lapsLocal.ManagedAccountName)" }
            }
            Write-DebugLog "Import-MdmXml: LAPS parsed $([char]0x2014) Backup=$($importInfo.LAPS.BackupDirectory)" -Level DEBUG
        }

        # ── Certificates (no local store lookup for remote device) ──
        $resNode = $mdmDoc.MDMEnterpriseDiagnosticsReport.Resources
        if ($resNode) {
            $certList = [System.Collections.Generic.List[hashtable]]::new()
            foreach ($enrollment in $resNode.Enrollment) {
                foreach ($scope in $enrollment.Scope) {
                    foreach ($resource in $scope.Resources.ChildNodes.'#Text') {
                        if ($resource -match 'RootCATrustedCertificates') {
                            $thumbprint = ($resource | Split-Path -Leaf -EA SilentlyContinue)
                            if (-not $thumbprint -or $thumbprint.Length -lt 20) { continue }
                            $certStoreName = if ($resource -match '/Root/') { 'Root CA' } elseif ($resource -match '/CA/') { 'Intermediate CA' } elseif ($resource -match '/TrustedPublisher/') { 'Trusted Publisher' } else { 'Other' }
                            $pathType = if ($resource -match '^\.\/device\/') { 'LocalMachine' } else { 'CurrentUser' }
                            $certList.Add(@{ Store = $certStoreName; Thumbprint = $thumbprint; Scope = $pathType; IssuedTo = ''; IssuedBy = ''; ValidFrom = ''; ValidTo = ''; ExpireDays = '' })
                        }
                    }
                }
            }
            if ($certList.Count -gt 0) { $importInfo.Certificates = @($certList) }
            Write-DebugLog "Import-MdmXml: $($certList.Count) certificate references found" -Level DEBUG
        }

        # ── PolicyMeta ──
        $metaNode = $mdmDoc.MDMEnterpriseDiagnosticsReport.PolicyManagerMeta
        if ($metaNode) {
            $policyMetaHash = @{}
            foreach ($areaGroup in $metaNode.AreaMetadata) {
                $areaName = $areaGroup.AreaName
                foreach ($pm in $areaGroup.Policies.ChildNodes) {
                    if ($pm.LocalName -and $pm.DefaultValue -ne $null) {
                        $policyMetaHash["$areaName/$($pm.LocalName)"] = @{
                            DefaultValue = "$($pm.DefaultValue)"
                            RegPath = if ($pm.RegKeyPathRedirect) { "$($pm.RegKeyPathRedirect)" } elseif ($pm.grouppolicyPath) { "GP: $($pm.grouppolicyPath)" } else { '' }
                        }
                    }
                }
            }
            $importInfo.PolicyMeta = $policyMetaHash
            Write-DebugLog "Import-MdmXml: PolicyMeta $([char]0x2014) $($policyMetaHash.Count) entries" -Level DEBUG
        }

        # Compliance for imported data
        $compIssues = [System.Collections.Generic.List[string]]::new()
        $importInfo.Compliance = @{
            Status = if ($compIssues.Count -eq 0) { 'Compliant' } else { 'At Risk' }
            ConfiguredPolicies = $importInfo.PolicyMeta.Count
            Issues = @($compIssues)
            AppSuccessRate = $null
        }

        # Update current scan data with imported MDM info
        if (-not $Script:ScanData) { $Script:ScanData = @{ MdmInfo = @{ ProviderID = '(Imported)'; EnrollmentUPN = ''; EnrollmentType = ''; EnrollmentState = ''; MdmDiag = @{} } } }
        if (-not $Script:ScanData.MdmInfo) { $Script:ScanData.MdmInfo = @{ ProviderID = '(Imported)'; EnrollmentUPN = ''; EnrollmentType = ''; EnrollmentState = ''; MdmDiag = @{} } }
        $Script:ScanData.MdmInfo.MdmDiag.LAPS = $importInfo.LAPS
        $Script:ScanData.MdmInfo.MdmDiag.Certificates = $importInfo.Certificates
        $Script:ScanData.MdmInfo.MdmDiag.PolicyMeta = $importInfo.PolicyMeta
        $Script:ScanData.MdmInfo.MdmDiag.Compliance = $importInfo.Compliance

        # Refresh dashboard cards
        $mi = $Script:ScanData.MdmInfo
        if ($ui.MdmEnrollmentCard) { $ui.MdmEnrollmentCard.Visibility = 'Visible' }
        if ($ui.MdmProviderText) { $ui.MdmProviderText.Text = if ($mi.ProviderID) { $mi.ProviderID } else { '(Imported)' } }
        if ($ui.MdmStateText)    { $ui.MdmStateText.Text = "Imported: $importName" }

        # LAPS card
        if ($importInfo.LAPS.BackupDirectory -and $importInfo.LAPS.BackupDirectory -ne 'Not Configured') {
            if ($ui.LapsStatusCard) { $ui.LapsStatusCard.Visibility = 'Visible' }
            if ($ui.LapsBackupText)      { $ui.LapsBackupText.Text      = $importInfo.LAPS.BackupDirectory }
            if ($ui.LapsPasswordAgeText) { $ui.LapsPasswordAgeText.Text = $importInfo.LAPS.PasswordAgeDays }
            if ($ui.LapsComplexityText)  { $ui.LapsComplexityText.Text  = $importInfo.LAPS.PasswordComplexity }
            if ($ui.LapsPostAuthText)    { $ui.LapsPostAuthText.Text    = $importInfo.LAPS.PostAuthActions }
            $lapsRotation = if ($importInfo.LAPS.LocalLastPasswordUpdate) { "Last rotation: $($importInfo.LAPS.LocalLastPasswordUpdate)" } else { '' }
            if ($importInfo.LAPS.LocalAzurePasswordExpiry) { $lapsRotation += if ($lapsRotation) { "`nExpiry: $($importInfo.LAPS.LocalAzurePasswordExpiry)" } else { "Expiry: $($importInfo.LAPS.LocalAzurePasswordExpiry)" } }
            if ($ui.LapsLastRotationText) { $ui.LapsLastRotationText.Text = $lapsRotation }
        }

        # Cert card
        if ($importInfo.Certificates.Count -gt 0) {
            if ($ui.CertInventoryCard) { $ui.CertInventoryCard.Visibility = 'Visible' }
            if ($ui.CertCountText) { $ui.CertCountText.Text = "$($importInfo.Certificates.Count)" }
            $expCnt = @($importInfo.Certificates | Where-Object { $_.ExpireDays -ne '' -and [int]$_.ExpireDays -le 30 }).Count
            if ($ui.CertExpiringText) { $ui.CertExpiringText.Text = "$expCnt" }
            $detLines = @($importInfo.Certificates | ForEach-Object { "$($_.Store): $($_.Thumbprint.Substring(0, [math]::Min(16, $_.Thumbprint.Length)))..." })
            if ($ui.CertDetailText) { $ui.CertDetailText.Text = ($detLines | Select-Object -First 5) -join "`n" }
        }

        # Compliance card
        $comp = $importInfo.Compliance
        if ($comp.Status -ne 'N/A') {
            if ($ui.ComplianceCard) { $ui.ComplianceCard.Visibility = 'Visible' }
            if ($ui.ComplianceStatusBadge) {
                $ui.ComplianceStatusBadge.Text = $comp.Status
                $badgeBg = switch ($comp.Status) { 'Compliant' { '#22C55E' } 'Non-compliant' { '#EF4444' } 'At Risk' { '#F59E0B' } default { '#71717A' } }
                $ui.ComplianceStatusBadge.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString($badgeBg)
            }
            if ($ui.CompliancePolicyCountText) { $ui.CompliancePolicyCountText.Text = "$($comp.ConfiguredPolicies)" }
            if ($ui.ComplianceIssuesText -and $comp.Issues.Count -gt 0) { $ui.ComplianceIssuesText.Text = ($comp.Issues | ForEach-Object { "$([char]0x26A0) $_" }) -join "`n" }
        }

        Show-Toast 'MDM XML Imported' "Parsed $importName`nLAPS: $($importInfo.LAPS.BackupDirectory)`nCerts: $($importInfo.Certificates.Count)`nMeta: $($importInfo.PolicyMeta.Count) entries" 'success'
        Write-DebugLog "Import-MdmXml: Complete" -Level SUCCESS
    } catch {
        Write-DebugLog "Import-MdmXml error: $($_.Exception.Message)" -Level ERROR
        Show-Toast 'Import Failed' $_.Exception.Message 'error'
    }
}

function Export-BaselineTemplate {
    if (-not $Script:ScanData) {
        Show-Toast 'No Data' 'Run a scan to create a baseline template from current settings.' 'warning'
        return
    }
    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Filter = 'JSON Baseline (*.json)|*.json'
    $dlg.FileName = "Baseline_$(Get-Date -Format 'yyyyMMdd').json"
    $dlg.InitialDirectory = $Script:SnapshotDir
    if (-not $dlg.ShowDialog()) { return }

    $entries = @($Script:ScanData.Settings | Where-Object { $_.ValueData } | ForEach-Object {
        @{ SettingKey = $_.SettingKey; ExpectedValue = "$($_.ValueData)"; Severity = 'Medium'; Reference = '' }
    })
    $baseline = @{ Name = "Custom Baseline $(Get-Date -Format 'yyyy-MM-dd')"; Settings = $entries }
    $baseline | ConvertTo-Json -Depth 4 -Compress | Set-Content -Path $dlg.FileName -Encoding UTF8
    Write-DebugLog "Baseline template exported: $($dlg.FileName)" -Level SUCCESS
    Show-Toast 'Baseline Created' "$($entries.Count) settings saved as baseline template" 'success'
}

# --- #9+#10: Copy/Clipboard helpers ---
function Copy-GridRowToClipboard([object]$Item) {
    if (-not $Item) { return }
    $props = $Item.PSObject.Properties | ForEach-Object { "$($_.Name): $($_.Value)" }
    [System.Windows.Clipboard]::SetText(($props -join "`r`n"))
    Show-Toast 'Copied' 'Row data copied to clipboard' 'info'
}

function Copy-GridCellToClipboard([object]$Item, [string]$Property) {
    if (-not $Item -or -not $Property) { return }
    $val = $Item.$Property
    if ($val) { [System.Windows.Clipboard]::SetText("$val") }
}

function Export-SelectedToCsv {
    param([System.Collections.IList]$Items, [string]$DefaultName)
    if (-not $Items -or $Items.Count -eq 0) {
        Show-Toast 'No Selection' 'Select rows first.' 'info'
        return
    }
    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Filter = 'CSV files (*.csv)|*.csv'
    $dlg.FileName = "${DefaultName}_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $dlg.InitialDirectory = $Script:ReportsDir
    if (-not $dlg.ShowDialog()) { return }

    $list = [System.Collections.Generic.List[object]]::new()
    foreach ($item in $Items) { [void]$list.Add($item) }
    $list | Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
    Show-Toast 'Exported' "$($list.Count) rows saved to CSV" 'success'
}

# --- #1: GPO Precedence (basic link-order) ---
function Resolve-GPOPrecedence {
    param([array]$Settings, [array]$GPOs)
    # Build link-order map from GPO Links property
    $linkOrder = @{}
    foreach ($gpo in $GPOs) {
        $links = "$($gpo.Links)" -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        if ($links.Count -gt 0) {
            # Lower link index = higher precedence (applied later = wins)
            $linkOrder[$gpo.DisplayName] = $links.Count
        } else {
            $linkOrder[$gpo.DisplayName] = 999
        }
    }
    $linkOrder
}

# --- #17: Dashboard Chart helpers ---  
function Update-DashboardCharts {
    if (-not $Script:ScanData -or -not $ui.ChartCanvas) { return }
    $canvas = $ui.ChartCanvas
    $canvas.Children.Clear()

    if ($ui.ChartPanel) { $ui.ChartPanel.Visibility = "Visible" }
    $gpos = $Script:ScanData.GPOs
    $settings = $Script:ScanData.Settings

    # Stacked bar: GPO status distribution
    $enabled = @($gpos | Where-Object { $_.Status -eq 'Enabled' -or $_.Status -eq 'Computer Only' -or $_.Status -eq 'User Only' }).Count
    $disabled = @($gpos | Where-Object { $_.Status -eq 'Disabled' }).Count
    $total = $enabled + $disabled
    if ($total -eq 0) { return }

    # Bar label
    $lbl = [System.Windows.Controls.TextBlock]::new()
    $lbl.Text = "GPO Status"
    $lbl.FontSize = 10
    $lbl.Foreground = (Get-CachedBrush '#FFA1A1AA')
    $lbl.FontWeight = 'SemiBold'
    [System.Windows.Controls.Canvas]::SetLeft($lbl, 0)
    [System.Windows.Controls.Canvas]::SetTop($lbl, 0)
    [void]$canvas.Children.Add($lbl)

    $barWidth = 220
    $barH = 20
    $barTop = 20

    # Enabled bar
    $ew = [math]::Max(2, [math]::Round(($enabled / $total) * $barWidth))
    $enabledRect = [System.Windows.Shapes.Rectangle]::new()
    $enabledRect.Width = $ew; $enabledRect.Height = $barH
    $enabledRect.RadiusX = 4; $enabledRect.RadiusY = 4
    $enabledRect.Fill = (Get-CachedBrush '#FF00C853')
    [System.Windows.Controls.Canvas]::SetLeft($enabledRect, 0)
    [System.Windows.Controls.Canvas]::SetTop($enabledRect, $barTop)
    [void]$canvas.Children.Add($enabledRect)

    # Disabled bar
    if ($disabled -gt 0) {
        $dw = $barWidth - $ew
        $disabledRect = [System.Windows.Shapes.Rectangle]::new()
        $disabledRect.Width = [math]::Max(2, $dw); $disabledRect.Height = $barH
        $disabledRect.RadiusX = 4; $disabledRect.RadiusY = 4
        $disabledRect.Fill = (Get-CachedBrush '#FFFF5000')
        [System.Windows.Controls.Canvas]::SetLeft($disabledRect, $ew)
        [System.Windows.Controls.Canvas]::SetTop($disabledRect, $barTop)
        [void]$canvas.Children.Add($disabledRect)
    }

    # Legend
    $leg1 = [System.Windows.Controls.TextBlock]::new()
    $leg1.Text = "Enabled: $enabled"
    $leg1.FontSize = 9; $leg1.Foreground = (Get-CachedBrush '#FF00C853')
    [System.Windows.Controls.Canvas]::SetLeft($leg1, 0)
    [System.Windows.Controls.Canvas]::SetTop($leg1, 44)
    [void]$canvas.Children.Add($leg1)

    $leg2 = [System.Windows.Controls.TextBlock]::new()
    $leg2.Text = "Disabled: $disabled"
    $leg2.FontSize = 9; $leg2.Foreground = (Get-CachedBrush '#FFFF5000')
    [System.Windows.Controls.Canvas]::SetLeft($leg2, 100)
    [System.Windows.Controls.Canvas]::SetTop($leg2, 44)
    [void]$canvas.Children.Add($leg2)

    # Settings scope bar
    $compCount = @($settings | Where-Object { $_.Scope -eq 'Computer' -or $_.Scope -eq 'Device' }).Count
    $userCount = @($settings | Where-Object { $_.Scope -eq 'User' }).Count
    $sTotal = $compCount + $userCount
    if ($sTotal -gt 0) {
        $lbl2 = [System.Windows.Controls.TextBlock]::new()
        $lbl2.Text = "Settings Scope"
        $lbl2.FontSize = 10
        $lbl2.Foreground = (Get-CachedBrush '#FFA1A1AA')
        $lbl2.FontWeight = 'SemiBold'
        [System.Windows.Controls.Canvas]::SetLeft($lbl2, 0)
        [System.Windows.Controls.Canvas]::SetTop($lbl2, 68)
        [void]$canvas.Children.Add($lbl2)

        $cw = [math]::Max(2, [math]::Round(($compCount / $sTotal) * $barWidth))
        $compRect = [System.Windows.Shapes.Rectangle]::new()
        $compRect.Width = $cw; $compRect.Height = $barH
        $compRect.RadiusX = 4; $compRect.RadiusY = 4
        $compRect.Fill = (Get-CachedBrush '#FF0078D4')
        [System.Windows.Controls.Canvas]::SetLeft($compRect, 0)
        [System.Windows.Controls.Canvas]::SetTop($compRect, 88)
        [void]$canvas.Children.Add($compRect)

        $uw = $barWidth - $cw
        if ($userCount -gt 0) {
            $userRect = [System.Windows.Shapes.Rectangle]::new()
            $userRect.Width = [math]::Max(2, $uw); $userRect.Height = $barH
            $userRect.RadiusX = 4; $userRect.RadiusY = 4
            $userRect.Fill = (Get-CachedBrush '#FF60CDFF')
            [System.Windows.Controls.Canvas]::SetLeft($userRect, $cw)
            [System.Windows.Controls.Canvas]::SetTop($userRect, 88)
            [void]$canvas.Children.Add($userRect)
        }

        $leg3 = [System.Windows.Controls.TextBlock]::new()
        $leg3.Text = "Computer: $compCount"
        $leg3.FontSize = 9; $leg3.Foreground = (Get-CachedBrush '#FF0078D4')
        [System.Windows.Controls.Canvas]::SetLeft($leg3, 0)
        [System.Windows.Controls.Canvas]::SetTop($leg3, 112)
        [void]$canvas.Children.Add($leg3)

        $leg4 = [System.Windows.Controls.TextBlock]::new()
        $leg4.Text = "User: $userCount"
        $leg4.FontSize = 9; $leg4.Foreground = (Get-CachedBrush '#FF60CDFF')
        [System.Windows.Controls.Canvas]::SetLeft($leg4, 100)
        [System.Windows.Controls.Canvas]::SetTop($leg4, 112)
        [void]$canvas.Children.Add($leg4)
    }

    # Top 5 categories mini-bar
    $catGroups = $settings | Group-Object Category | Sort-Object Count -Descending | Select-Object -First 5
    if ($catGroups.Count -gt 0) {
        $lbl3 = [System.Windows.Controls.TextBlock]::new()
        $lbl3.Text = "Top Categories"
        $lbl3.FontSize = 10
        $lbl3.Foreground = (Get-CachedBrush '#FFA1A1AA')
        $lbl3.FontWeight = 'SemiBold'
        [System.Windows.Controls.Canvas]::SetLeft($lbl3, 0)
        [System.Windows.Controls.Canvas]::SetTop($lbl3, 136)
        [void]$canvas.Children.Add($lbl3)

        $maxCat = ($catGroups | Measure-Object Count -Maximum).Maximum
        $yOff = 156
        foreach ($cg in $catGroups) {
            $catLbl = [System.Windows.Controls.TextBlock]::new()
            $catLbl.Text = "$($cg.Name) ($($cg.Count))"
            $catLbl.FontSize = 9
            $catLbl.Foreground = (Get-CachedBrush '#FFE0E0E0')
            $catLbl.MaxWidth = 220
            $catLbl.TextTrimming = 'CharacterEllipsis'
            [System.Windows.Controls.Canvas]::SetLeft($catLbl, 0)
            [System.Windows.Controls.Canvas]::SetTop($catLbl, $yOff)
            [void]$canvas.Children.Add($catLbl)
            $yOff += 14

            $bw = [math]::Max(4, [math]::Round(($cg.Count / $maxCat) * $barWidth))
            $catBar = [System.Windows.Shapes.Rectangle]::new()
            $catBar.Width = $bw; $catBar.Height = 8
            $catBar.RadiusX = 3; $catBar.RadiusY = 3
            $catBar.Fill = (Get-CachedBrush '#FF0078D4')
            $catBar.Opacity = 0.7
            [System.Windows.Controls.Canvas]::SetLeft($catBar, 0)
            [System.Windows.Controls.Canvas]::SetTop($catBar, $yOff)
            [void]$canvas.Children.Add($catBar)
            $yOff += 16
        }
    }
}


# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 15.5: NEW FEATURES - Detail Pane, Context Menus, Diff, Script Gen, etc.
# ═══════════════════════════════════════════════════════════════════════════════

# --- #2: Detail Pane helper ---
function Show-DetailPane {
    param([object]$Item, [string]$Mode)
    if (-not $Item) { return }
    $sb = [System.Text.StringBuilder]::new()
    switch ($Mode) {
        'Setting' {
            $null = $sb.AppendLine("SETTING DETAIL")
            $null = $sb.AppendLine("$([char]0x2500)" * 40)
            $null = $sb.AppendLine("Policy Name:   $($Item.PolicyName)")
            $null = $sb.AppendLine("Category:      $($Item.Category)")
            $null = $sb.AppendLine("Scope:         $($Item.Scope)")
            $null = $sb.AppendLine("State:         $($Item.State)")
            $null = $sb.AppendLine("Value:         $($Item.ValueData)")
            if ($Item.DefaultValue) { $null = $sb.AppendLine("Default:       $($Item.DefaultValue)") }
            if ($Item.RegistryKey)  { $null = $sb.AppendLine("Registry:      $($Item.RegistryKey)") }
            $null = $sb.AppendLine("GPO / Source:  $($Item.GPOName)")
            if ($Item.IntuneGroup)  { $null = $sb.AppendLine("Intune Group:  $($Item.IntuneGroup)") }
            # M5: Script details
            if ($Item.ScriptType)   { $null = $sb.AppendLine("Script Type:   $($Item.ScriptType)") }
            if ($Item.ScriptPath)   { $null = $sb.AppendLine("Script Path:   $($Item.ScriptPath)") }
            if ($Item.ScriptParams) { $null = $sb.AppendLine("Script Params: $($Item.ScriptParams)") }
            if ($Item.Explain)      { $null = $sb.AppendLine(""); $null = $sb.AppendLine("DESCRIPTION"); $null = $sb.AppendLine($Item.Explain) }
            # Enriched CSP metadata from csp_metadata.json
            $cspKey = $Item.SettingKey
            $cspInfo = if ($cspKey -and $Script:CspMetaKeys) { $Script:CspMetaKeys[$cspKey] } else { $null }
            if ($cspInfo) {
                if ($cspInfo.Desc -and -not $Item.Explain) {
                    $null = $sb.AppendLine("")
                    $null = $sb.AppendLine("DESCRIPTION")
                    $null = $sb.AppendLine($cspInfo.Desc)
                }
                $null = $sb.AppendLine("")
                $null = $sb.AppendLine("CSP METADATA")
                $null = $sb.AppendLine("$([char]0x2500)" * 40)
                if ($cspInfo.Scope)    { $null = $sb.AppendLine("CSP Scope:     $($cspInfo.Scope)") }
                if ($cspInfo.Editions) { $null = $sb.AppendLine("Editions:      $($cspInfo.Editions)") }
                if ($cspInfo.MinVer)   { $null = $sb.AppendLine("Min Version:   $($cspInfo.MinVer)") }
                if ($cspInfo.Format)   { $null = $sb.AppendLine("Format:        $($cspInfo.Format)") }
                if ($cspInfo.AV) {
                    $avObj = $cspInfo.AV
                    $avProps = if ($avObj -is [hashtable]) { $avObj.GetEnumerator() } elseif ($avObj.PSObject.Properties) { $avObj.PSObject.Properties } else { $null }
                    if ($avProps) {
                        $null = $sb.AppendLine("")
                        $null = $sb.AppendLine("ALLOWED VALUES")
                        foreach ($av in $avProps) {
                            $avKey = if ($av.Key) { $av.Key } else { $av.Name }
                            $avVal = $av.Value
                            if ($avKey -eq '_format') {
                                $null = $sb.AppendLine("  $avVal")
                            } else {
                                $null = $sb.AppendLine("  $avKey = $avVal")
                            }
                        }
                    }
                }
                if ($cspInfo.GP) {
                    $gpObj = $cspInfo.GP
                    $gpProps = if ($gpObj -is [hashtable]) { $gpObj.GetEnumerator() } elseif ($gpObj.PSObject.Properties) { $gpObj.PSObject.Properties } else { $null }
                    if ($gpProps) {
                        $null = $sb.AppendLine("")
                        $null = $sb.AppendLine("GROUP POLICY MAPPING")
                        foreach ($gp in $gpProps) {
                            $gpKey = if ($gp.Key) { $gp.Key } else { $gp.Name }
                            $gpVal = $gp.Value
                            $null = $sb.AppendLine("  ${gpKey}: $gpVal")
                        }
                    }
                }
                # Staleness warning
                if ($Script:CspDbAge -and $Script:CspDbAge -gt 90) {
                    $null = $sb.AppendLine("")
                    $null = $sb.AppendLine("[!] CSP database is $($Script:CspDbAge) days old. Run Build-CspDatabase.ps1 to refresh.")
                }
            } else {
                # No CSP metadata found for this setting
                if ($Item.Source -eq 'Intune' -or $Item.IntuneGroup -in @('Device Configuration','Registry Settings')) {
                    $null = $sb.AppendLine("")
                    $null = $sb.AppendLine("CSP METADATA")
                    $null = $sb.AppendLine("$([char]0x2500)" * 40)
                    if (-not $Script:CspMetaKeys -or $Script:CspMetaKeys.Count -eq 0) {
                        $null = $sb.AppendLine("(CSP database not loaded - run Build-CspDatabase.ps1)")
                    } else {
                        $null = $sb.AppendLine("(No CSP reference found for this setting key)")
                        $null = $sb.AppendLine("Setting Key: $cspKey")
                    }
                }
            }
        }
        'GPO' {
            $null = $sb.AppendLine("GPO / AREA DETAIL")
            $null = $sb.AppendLine("$([char]0x2500)" * 40)
            $null = $sb.AppendLine("Display Name:     $($Item.DisplayName)")
            $null = $sb.AppendLine("Status:           $($Item.Status)")
            $null = $sb.AppendLine("Setting Count:    $($Item.SettingCount)")
            $null = $sb.AppendLine("Link Count:       $($Item.LinkCount)")
            $null = $sb.AppendLine("Links:            $($Item.Links)")
            if ($Item.Enforced)          { $null = $sb.AppendLine("Enforced:         Yes") }
            if ($Item.LinkOrder)         { $null = $sb.AppendLine("Link Order:       $($Item.LinkOrder)") }
            if ($Item.SecurityFiltering) { $null = $sb.AppendLine("Security Filter:  $($Item.SecurityFiltering)") }
            if ($Item.WmiFilter)         { $null = $sb.AppendLine("WMI Filter:       $($Item.WmiFilter)") }
            if ($Item.WmiFilterStatus)   { $null = $sb.AppendLine("WMI Filter Status:$($Item.WmiFilterStatus)") }
            if ($Item.ScopeTags)         { $null = $sb.AppendLine("Scope Tags:       $($Item.ScopeTags)") }
            if ($Item.Assignments)       { $null = $sb.AppendLine("Assignments:      $($Item.Assignments)") }
            if ($Item.CreationTime)      { $null = $sb.AppendLine("Created:          $($Item.CreationTime)") }
            if ($Item.ModificationTime)  { $null = $sb.AppendLine("Modified:         $($Item.ModificationTime)") }
        }
        'Conflict' {
            $null = $sb.AppendLine("CONFLICT DETAIL")
            $null = $sb.AppendLine("$([char]0x2500)" * 40)
            $null = $sb.AppendLine("Severity:      $($Item.Severity)")
            $null = $sb.AppendLine("Setting:       $($Item.SettingKey)")
            $null = $sb.AppendLine("Scope:         $($Item.Scope)")
            $null = $sb.AppendLine("Category:      $($Item.Category)")
            $null = $sb.AppendLine("GPO Count:     $($Item.GPOCount)")
            $null = $sb.AppendLine("GPOs:          $($Item.GPONames)")
            $null = $sb.AppendLine("Values:        $($Item.Values)")
            $null = $sb.AppendLine("Winner GPO:    $($Item.WinnerGPO)")
        }
        'IntuneApp' {
            $null = $sb.AppendLine("INTUNE APP DETAIL")
            $null = $sb.AppendLine("$([char]0x2500)" * 40)
            $null = $sb.AppendLine("App Name:      $($Item.AppName)")
            $null = $sb.AppendLine("App Type:      $($Item.AppType)")
            $null = $sb.AppendLine("Install State: $($Item.InstallState)")
            $null = $sb.AppendLine("Enforcement:   $($Item.EnforcementState)")
            if ($Item.ErrorCode) { $null = $sb.AppendLine("Error Code:    $($Item.ErrorCode)") }
            $null = $sb.AppendLine("User / Device: $($Item.UserId)")
            if ($Item.RegistryKey) { $null = $sb.AppendLine("Registry:      $($Item.RegistryKey)") }
            $null = $sb.AppendLine("Last Modified: $($Item.LastModified)")
        }
    }
    $sb.ToString()
}

# --- #5: Snapshot Diff ---
function Compare-Snapshots {
    param([hashtable]$Baseline, [hashtable]$Current)
    $diffs = [System.Collections.Generic.List[PSCustomObject]]::new()

    # Build lookup by composite key
    $baseSettings = @{}
    foreach ($s in $Baseline.Settings) {
        $key = "$($s.GPOName)|$($s.SettingKey)|$($s.Scope)"
        $baseSettings[$key] = $s
    }
    $currSettings = @{}
    foreach ($s in $Current.Settings) {
        $key = "$($s.GPOName)|$($s.SettingKey)|$($s.Scope)"
        $currSettings[$key] = $s
    }

    # Find added / modified
    foreach ($key in $currSettings.Keys) {
        if (-not $baseSettings.ContainsKey($key)) {
            $c = $currSettings[$key]
            [void]$diffs.Add([PSCustomObject]@{ Change='Added'; PolicyName=$c.PolicyName; Category=$c.Category; Scope=$c.Scope; GPOName=$c.GPOName; OldValue=''; NewValue=$c.ValueData })
        } elseif ("$($currSettings[$key].ValueData)" -ne "$($baseSettings[$key].ValueData)") {
            $c = $currSettings[$key]; $b = $baseSettings[$key]
            [void]$diffs.Add([PSCustomObject]@{ Change='Modified'; PolicyName=$c.PolicyName; Category=$c.Category; Scope=$c.Scope; GPOName=$c.GPOName; OldValue=$b.ValueData; NewValue=$c.ValueData })
        }
    }

    # Find removed
    foreach ($key in $baseSettings.Keys) {
        if (-not $currSettings.ContainsKey($key)) {
            $b = $baseSettings[$key]
            [void]$diffs.Add([PSCustomObject]@{ Change='Removed'; PolicyName=$b.PolicyName; Category=$b.Category; Scope=$b.Scope; GPOName=$b.GPOName; OldValue=$b.ValueData; NewValue='' })
        }
    }

    # GPO-level diffs
    $baseGPOs = @{}; foreach ($g in $Baseline.GPOs) { $baseGPOs[$g.DisplayName] = $g }
    $currGPOs = @{}; foreach ($g in $Current.GPOs) { $currGPOs[$g.DisplayName] = $g }
    foreach ($name in $currGPOs.Keys) {
        if (-not $baseGPOs.ContainsKey($name)) {
            [void]$diffs.Add([PSCustomObject]@{ Change='GPO Added'; PolicyName=$name; Category='GPO'; Scope=''; GPOName=$name; OldValue=''; NewValue='New' })
        }
    }
    foreach ($name in $baseGPOs.Keys) {
        if (-not $currGPOs.ContainsKey($name)) {
            [void]$diffs.Add([PSCustomObject]@{ Change='GPO Removed'; PolicyName=$name; Category='GPO'; Scope=''; GPOName=$name; OldValue='Existed'; NewValue='' })
        }
    }

    $diffs
}

# --- #24: Script Generation ---
function Export-AsScript {
    if (-not $Script:ScanData) {
        Show-Toast 'No Data' 'Run a scan first.' 'warning'
        return
    }
    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Filter = 'PowerShell scripts (*.ps1)|*.ps1'
    $dlg.FileName = "GPO_Recreate_$(Get-Date -Format 'yyyyMMdd_HHmmss').ps1"
    $dlg.InitialDirectory = $Script:ReportsDir
    if (-not $dlg.ShowDialog()) { return }

    $sb = [System.Text.StringBuilder]::new()
    $null = $sb.AppendLine('#Requires -Modules GroupPolicy')
    $null = $sb.AppendLine('# GPO Recreation Script')
    $null = $sb.AppendLine("# Generated by PolicyPilot on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    $null = $sb.AppendLine("# Source domain: $($Script:ScanData.Domain)")
    $null = $sb.AppendLine("# WARNING: Review before running. This script modifies Group Policy.")
    $null = $sb.AppendLine('')
    $null = $sb.AppendLine('param([switch]$WhatIf)')
    $null = $sb.AppendLine('')

    $grouped = $Script:ScanData.Settings | Group-Object GPOName
    foreach ($group in $grouped) {
        $gpoName = $group.Name
        $null = $sb.AppendLine("# --- GPO: $gpoName ---")
        $null = $sb.AppendLine("Write-Host `"Processing GPO: $gpoName`"")
        $null = $sb.AppendLine("`$gpo = Get-GPO -Name '$($gpoName -replace "'","''")' -ErrorAction SilentlyContinue")
        $null = $sb.AppendLine("if (-not `$gpo) { `$gpo = New-GPO -Name '$($gpoName -replace "'","''")' -WhatIf:`$WhatIf }")
        $null = $sb.AppendLine('')

        foreach ($s in $group.Group) {
            if ($s.RegistryKey -and $s.ValueData) {
                $regPath = $s.RegistryKey -replace '\\[^\\]+$', ''
                $valName = $s.RegistryKey -replace '^.*\\', ''
                $null = $sb.AppendLine("# $($s.PolicyName)")
                $null = $sb.AppendLine("if (-not `$WhatIf) { Set-GPRegistryValue -Name '$($gpoName -replace "'","''")' -Key '$regPath' -ValueName '$valName' -Value '$($s.ValueData -replace "'","''")' -Type String }")
                $null = $sb.AppendLine('')
            }
        }
    }

    [System.IO.File]::WriteAllText($dlg.FileName, $sb.ToString(), [System.Text.Encoding]::UTF8)
    Write-DebugLog "Script exported: $($dlg.FileName)" -Level SUCCESS
    Show-Toast 'Script Exported' "GPO recreation script saved" 'success'
}

# --- #16: Print via HTML in browser ---
function Invoke-PrintReport {
    if (-not $Script:ScanData) {
        Show-Toast 'No Data' 'Run a scan first.' 'warning'
        return
    }
    $tmpFile = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "PolicyPilot_Print_$(Get-Date -Format 'yyyyMMddHHmmss').html")
    # Reuse Export-Html logic to temp file
    $script:_printPath = $tmpFile
    Export-HtmlToPath $tmpFile
    Start-Process $tmpFile
    Write-DebugLog "Opened report in browser for printing" -Level SUCCESS
}

function Export-HtmlToPath([string]$Path) {
    $htmlContent = Build-HtmlReport
    [System.IO.File]::WriteAllText($Path, $htmlContent, [System.Text.Encoding]::UTF8)
}

# --- #25: Registry .reg Export ---
function Export-RegistryFile {
    if (-not $Script:ScanData) {
        Show-Toast 'No Data' 'Run a scan first.' 'warning'
        return
    }
    $regSettings = @($Script:ScanData.Settings | Where-Object { $_.RegistryKey })
    if ($regSettings.Count -eq 0) {
        Show-Toast 'No Registry Settings' 'No registry-based settings found in scan data.' 'info'
        return
    }
    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Filter = 'Registry files (*.reg)|*.reg'
    $dlg.FileName = "GPO_Registry_$(Get-Date -Format 'yyyyMMdd_HHmmss').reg"
    $dlg.InitialDirectory = $Script:ReportsDir
    if (-not $dlg.ShowDialog()) { return }

    $sb = [System.Text.StringBuilder]::new()
    $null = $sb.AppendLine('Windows Registry Editor Version 5.00')
    $null = $sb.AppendLine('')
    $null = $sb.AppendLine('; GPO Registry Export')
    $null = $sb.AppendLine("; Generated by PolicyPilot on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    $null = $sb.AppendLine("; Domain: $($Script:ScanData.Domain)")
    $null = $sb.AppendLine('; WARNING: Review carefully before importing.')
    $null = $sb.AppendLine('')

    $grouped = $regSettings | Group-Object { ($_.RegistryKey -replace '\\[^\\]+$','') }
    foreach ($g in ($grouped | Sort-Object Name)) {
        $null = $sb.AppendLine("[$($g.Name)]")
        foreach ($s in $g.Group) {
            $valName = $s.RegistryKey -replace '^.*\\',''
            $val = "$($s.ValueData)" -replace '\\','\\' -replace '"','\"'
            $null = $sb.AppendLine("`"$valName`"=`"$val`"")
        }
        $null = $sb.AppendLine('')
    }

    [System.IO.File]::WriteAllText($dlg.FileName, $sb.ToString(), [System.Text.Encoding]::UTF8)
    Write-DebugLog "Registry export: $($dlg.FileName)" -Level SUCCESS
    Show-Toast 'Registry Exported' "$($regSettings.Count) registry settings exported to .reg file" 'success'
}

# --- #27: Impact Simulator (simple remove-and-re-diff) ---
function Invoke-ImpactSimulation {
    param([string]$GPONameToRemove)
    if (-not $Script:ScanData -or -not $GPONameToRemove) { return $null }

    $remainingSettings = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($s in $Script:ScanData.Settings) {
        if ($s.GPOName -ne $GPONameToRemove) { [void]$remainingSettings.Add($s) }
    }
    $removedSettings = @($Script:ScanData.Settings | Where-Object { $_.GPOName -eq $GPONameToRemove })
    $originalConflicts = Find-Conflicts $Script:ScanData.Settings
    $newConflicts = Find-Conflicts $remainingSettings

    $resolvedConflicts = @($originalConflicts | Where-Object { $_.GPONames -like "*$GPONameToRemove*" }) |
        Where-Object { $_.SettingKey -notin @($newConflicts | ForEach-Object { $_.SettingKey }) }

    [PSCustomObject]@{
        GPORemoved        = $GPONameToRemove
        SettingsLost      = $removedSettings.Count
        SettingsRemaining = $remainingSettings.Count
        ConflictsBefore   = $originalConflicts.Count
        ConflictsAfter    = $newConflicts.Count
        ConflictsResolved = $resolvedConflicts.Count
        OrphanedSettings  = @($removedSettings | Where-Object { $_.SettingKey -notin @($remainingSettings | ForEach-Object { $_.SettingKey }) })
        Details           = $removedSettings
    }
}

# --- #12: Baseline Comparison Framework ---
$Script:ActiveBaseline = $null

function Import-Baseline {
    $dlg = [Microsoft.Win32.OpenFileDialog]::new()
    $dlg.Filter = 'JSON Baseline (*.json)|*.json'
    $dlg.Title = 'Import Security Baseline'
    $dlg.InitialDirectory = $Script:SnapshotDir
    if (-not $dlg.ShowDialog()) { return }

    try {
        $json = [System.IO.File]::ReadAllText($dlg.FileName) | ConvertFrom-Json
        $Script:ActiveBaseline = @{
            Name     = if ($json.Name) { $json.Name } else { [System.IO.Path]::GetFileNameWithoutExtension($dlg.FileName) }
            Settings = @{}
        }
        foreach ($entry in $json.Settings) {
            $Script:ActiveBaseline.Settings[$entry.SettingKey] = @{
                ExpectedValue = $entry.ExpectedValue
                Severity      = if ($entry.Severity) { $entry.Severity } else { 'Medium' }
                Reference     = if ($entry.Reference) { $entry.Reference } else { '' }
            }
        }
        Write-DebugLog "Baseline loaded: $($Script:ActiveBaseline.Name) ($($Script:ActiveBaseline.Settings.Count) settings)" -Level SUCCESS
        Show-Toast 'Baseline Loaded' "$($Script:ActiveBaseline.Name): $($Script:ActiveBaseline.Settings.Count) settings" 'success'

        if ($Script:ScanData) { Update-BaselineCompliance }
    } catch {
        Write-DebugLog "Baseline import error: $($_.Exception.Message)" -Level ERROR
        Show-Toast 'Import Failed' $_.Exception.Message 'error'
    }
}

function Update-BaselineCompliance {
    if (-not $Script:ActiveBaseline -or -not $Script:ScanData) { return }
    $pass = 0; $fail = 0; $missing = 0
    foreach ($key in $Script:ActiveBaseline.Settings.Keys) {
        $expected = $Script:ActiveBaseline.Settings[$key]
        $actual = $Script:ScanData.Settings | Where-Object { $_.SettingKey -eq $key } | Select-Object -First 1
        if (-not $actual) { $missing++ }
        elseif ("$($actual.ValueData)" -eq "$($expected.ExpectedValue)") { $pass++ }
        else { $fail++ }
    }
    $total = $Script:ActiveBaseline.Settings.Count
    $pct = if ($total -gt 0) { [math]::Round(($pass / $total) * 100, 1) } else { 0 }
    $ui.BaselineStatus.Text = "$($Script:ActiveBaseline.Name): ${pct}% compliant ($pass pass, $fail fail, $missing missing)"
    Write-DebugLog "Baseline compliance: $pct% ($pass/$total pass)" -Level INFO
}

function Export-BaselineTemplate {
    if (-not $Script:ScanData) {
        Show-Toast 'No Data' 'Run a scan to create a baseline template from current settings.' 'warning'
        return
    }
    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Filter = 'JSON Baseline (*.json)|*.json'
    $dlg.FileName = "Baseline_$(Get-Date -Format 'yyyyMMdd').json"
    $dlg.InitialDirectory = $Script:SnapshotDir
    if (-not $dlg.ShowDialog()) { return }

    $entries = @($Script:ScanData.Settings | Where-Object { $_.ValueData } | ForEach-Object {
        @{ SettingKey = $_.SettingKey; ExpectedValue = "$($_.ValueData)"; Severity = 'Medium'; Reference = '' }
    })
    $baseline = @{ Name = "Custom Baseline $(Get-Date -Format 'yyyy-MM-dd')"; Settings = $entries }
    $baseline | ConvertTo-Json -Depth 4 -Compress | Set-Content -Path $dlg.FileName -Encoding UTF8
    Write-DebugLog "Baseline template exported: $($dlg.FileName)" -Level SUCCESS
    Show-Toast 'Baseline Created' "$($entries.Count) settings saved as baseline template" 'success'
}

# --- #9+#10: Copy/Clipboard helpers ---
function Copy-GridRowToClipboard([object]$Item) {
    if (-not $Item) { return }
    $props = $Item.PSObject.Properties | ForEach-Object { "$($_.Name): $($_.Value)" }
    [System.Windows.Clipboard]::SetText(($props -join "`r`n"))
    Show-Toast 'Copied' 'Row data copied to clipboard' 'info'
}

function Copy-GridCellToClipboard([object]$Item, [string]$Property) {
    if (-not $Item -or -not $Property) { return }
    $val = $Item.$Property
    if ($val) { [System.Windows.Clipboard]::SetText("$val") }
}

function Export-SelectedToCsv {
    param([System.Collections.IList]$Items, [string]$DefaultName)
    if (-not $Items -or $Items.Count -eq 0) {
        Show-Toast 'No Selection' 'Select rows first.' 'info'
        return
    }
    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Filter = 'CSV files (*.csv)|*.csv'
    $dlg.FileName = "${DefaultName}_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $dlg.InitialDirectory = $Script:ReportsDir
    if (-not $dlg.ShowDialog()) { return }

    $list = [System.Collections.Generic.List[object]]::new()
    foreach ($item in $Items) { [void]$list.Add($item) }
    $list | Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
    Show-Toast 'Exported' "$($list.Count) rows saved to CSV" 'success'
}

# --- #1: GPO Precedence (basic link-order) ---
function Resolve-GPOPrecedence {
    param([array]$Settings, [array]$GPOs)
    # Build link-order map from GPO Links property
    $linkOrder = @{}
    foreach ($gpo in $GPOs) {
        $links = "$($gpo.Links)" -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        if ($links.Count -gt 0) {
            # Lower link index = higher precedence (applied later = wins)
            $linkOrder[$gpo.DisplayName] = $links.Count
        } else {
            $linkOrder[$gpo.DisplayName] = 999
        }
    }
    $linkOrder
}

# --- #17: Dashboard Chart helpers ---  
function Update-DashboardCharts {
    if (-not $Script:ScanData -or -not $ui.ChartCanvas) { return }
    $canvas = $ui.ChartCanvas
    $canvas.Children.Clear()

    $gpos = $Script:ScanData.GPOs
    $settings = $Script:ScanData.Settings

    # Stacked bar: GPO status distribution
    $enabled = @($gpos | Where-Object { $_.Status -eq 'Enabled' -or $_.Status -eq 'Computer Only' -or $_.Status -eq 'User Only' }).Count
    $disabled = @($gpos | Where-Object { $_.Status -eq 'Disabled' }).Count
    $total = $enabled + $disabled
    if ($total -eq 0) { return }

    # Bar label
    $lbl = [System.Windows.Controls.TextBlock]::new()
    $lbl.Text = "GPO Status"
    $lbl.FontSize = 10
    $lbl.Foreground = (Get-CachedBrush '#FFA1A1AA')
    $lbl.FontWeight = 'SemiBold'
    [System.Windows.Controls.Canvas]::SetLeft($lbl, 0)
    [System.Windows.Controls.Canvas]::SetTop($lbl, 0)
    [void]$canvas.Children.Add($lbl)

    $barWidth = 220
    $barH = 20
    $barTop = 20

    # Enabled bar
    $ew = [math]::Max(2, [math]::Round(($enabled / $total) * $barWidth))
    $enabledRect = [System.Windows.Shapes.Rectangle]::new()
    $enabledRect.Width = $ew; $enabledRect.Height = $barH
    $enabledRect.RadiusX = 4; $enabledRect.RadiusY = 4
    $enabledRect.Fill = (Get-CachedBrush '#FF00C853')
    [System.Windows.Controls.Canvas]::SetLeft($enabledRect, 0)
    [System.Windows.Controls.Canvas]::SetTop($enabledRect, $barTop)
    [void]$canvas.Children.Add($enabledRect)

    # Disabled bar
    if ($disabled -gt 0) {
        $dw = $barWidth - $ew
        $disabledRect = [System.Windows.Shapes.Rectangle]::new()
        $disabledRect.Width = [math]::Max(2, $dw); $disabledRect.Height = $barH
        $disabledRect.RadiusX = 4; $disabledRect.RadiusY = 4
        $disabledRect.Fill = (Get-CachedBrush '#FFFF5000')
        [System.Windows.Controls.Canvas]::SetLeft($disabledRect, $ew)
        [System.Windows.Controls.Canvas]::SetTop($disabledRect, $barTop)
        [void]$canvas.Children.Add($disabledRect)
    }

    # Legend
    $leg1 = [System.Windows.Controls.TextBlock]::new()
    $leg1.Text = "Enabled: $enabled"
    $leg1.FontSize = 9; $leg1.Foreground = (Get-CachedBrush '#FF00C853')
    [System.Windows.Controls.Canvas]::SetLeft($leg1, 0)
    [System.Windows.Controls.Canvas]::SetTop($leg1, 44)
    [void]$canvas.Children.Add($leg1)

    $leg2 = [System.Windows.Controls.TextBlock]::new()
    $leg2.Text = "Disabled: $disabled"
    $leg2.FontSize = 9; $leg2.Foreground = (Get-CachedBrush '#FFFF5000')
    [System.Windows.Controls.Canvas]::SetLeft($leg2, 100)
    [System.Windows.Controls.Canvas]::SetTop($leg2, 44)
    [void]$canvas.Children.Add($leg2)

    # Settings scope bar
    $compCount = @($settings | Where-Object { $_.Scope -eq 'Computer' -or $_.Scope -eq 'Device' }).Count
    $userCount = @($settings | Where-Object { $_.Scope -eq 'User' }).Count
    $sTotal = $compCount + $userCount
    if ($sTotal -gt 0) {
        $lbl2 = [System.Windows.Controls.TextBlock]::new()
        $lbl2.Text = "Settings Scope"
        $lbl2.FontSize = 10
        $lbl2.Foreground = (Get-CachedBrush '#FFA1A1AA')
        $lbl2.FontWeight = 'SemiBold'
        [System.Windows.Controls.Canvas]::SetLeft($lbl2, 0)
        [System.Windows.Controls.Canvas]::SetTop($lbl2, 68)
        [void]$canvas.Children.Add($lbl2)

        $cw = [math]::Max(2, [math]::Round(($compCount / $sTotal) * $barWidth))
        $compRect = [System.Windows.Shapes.Rectangle]::new()
        $compRect.Width = $cw; $compRect.Height = $barH
        $compRect.RadiusX = 4; $compRect.RadiusY = 4
        $compRect.Fill = (Get-CachedBrush '#FF0078D4')
        [System.Windows.Controls.Canvas]::SetLeft($compRect, 0)
        [System.Windows.Controls.Canvas]::SetTop($compRect, 88)
        [void]$canvas.Children.Add($compRect)

        $uw = $barWidth - $cw
        if ($userCount -gt 0) {
            $userRect = [System.Windows.Shapes.Rectangle]::new()
            $userRect.Width = [math]::Max(2, $uw); $userRect.Height = $barH
            $userRect.RadiusX = 4; $userRect.RadiusY = 4
            $userRect.Fill = (Get-CachedBrush '#FF60CDFF')
            [System.Windows.Controls.Canvas]::SetLeft($userRect, $cw)
            [System.Windows.Controls.Canvas]::SetTop($userRect, 88)
            [void]$canvas.Children.Add($userRect)
        }

        $leg3 = [System.Windows.Controls.TextBlock]::new()
        $leg3.Text = "Computer: $compCount"
        $leg3.FontSize = 9; $leg3.Foreground = (Get-CachedBrush '#FF0078D4')
        [System.Windows.Controls.Canvas]::SetLeft($leg3, 0)
        [System.Windows.Controls.Canvas]::SetTop($leg3, 112)
        [void]$canvas.Children.Add($leg3)

        $leg4 = [System.Windows.Controls.TextBlock]::new()
        $leg4.Text = "User: $userCount"
        $leg4.FontSize = 9; $leg4.Foreground = (Get-CachedBrush '#FF60CDFF')
        [System.Windows.Controls.Canvas]::SetLeft($leg4, 100)
        [System.Windows.Controls.Canvas]::SetTop($leg4, 112)
        [void]$canvas.Children.Add($leg4)
    }

    # Top 5 categories mini-bar
    $catGroups = $settings | Group-Object Category | Sort-Object Count -Descending | Select-Object -First 5
    if ($catGroups.Count -gt 0) {
        $lbl3 = [System.Windows.Controls.TextBlock]::new()
        $lbl3.Text = "Top Categories"
        $lbl3.FontSize = 10
        $lbl3.Foreground = (Get-CachedBrush '#FFA1A1AA')
        $lbl3.FontWeight = 'SemiBold'
        [System.Windows.Controls.Canvas]::SetLeft($lbl3, 0)
        [System.Windows.Controls.Canvas]::SetTop($lbl3, 136)
        [void]$canvas.Children.Add($lbl3)

        $maxCat = ($catGroups | Measure-Object Count -Maximum).Maximum
        $yOff = 156
        foreach ($cg in $catGroups) {
            $catLbl = [System.Windows.Controls.TextBlock]::new()
            $catLbl.Text = "$($cg.Name) ($($cg.Count))"
            $catLbl.FontSize = 9
            $catLbl.Foreground = (Get-CachedBrush '#FFE0E0E0')
            $catLbl.MaxWidth = 220
            $catLbl.TextTrimming = 'CharacterEllipsis'
            [System.Windows.Controls.Canvas]::SetLeft($catLbl, 0)
            [System.Windows.Controls.Canvas]::SetTop($catLbl, $yOff)
            [void]$canvas.Children.Add($catLbl)
            $yOff += 14

            $bw = [math]::Max(4, [math]::Round(($cg.Count / $maxCat) * $barWidth))
            $catBar = [System.Windows.Shapes.Rectangle]::new()
            $catBar.Width = $bw; $catBar.Height = 8
            $catBar.RadiusX = 3; $catBar.RadiusY = 3
            $catBar.Fill = (Get-CachedBrush '#FF0078D4')
            $catBar.Opacity = 0.7
            [System.Windows.Controls.Canvas]::SetLeft($catBar, 0)
            [System.Windows.Controls.Canvas]::SetTop($catBar, $yOff)
            [void]$canvas.Children.Add($catBar)
            $yOff += 16
        }
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 15b: RESTORE LAST SCAN FROM DISK CACHE
# ═══════════════════════════════════════════════════════════════════════════════

function Restore-LastScan {
    $snapshotPath = [IO.Path]::Combine($env:TEMP, 'PolicyPilot_GPOCache', 'last_scan.clixml')
    if (-not (Test-Path $snapshotPath)) {
        Show-Toast 'No Cached Scan' 'No previous scan snapshot found.' 'warning'
        return $false
    }
    $age = [DateTime]::Now - (Get-Item $snapshotPath).LastWriteTime
    if ($age.TotalHours -gt 24) {
        Show-Toast 'Stale Cache' "Last scan is $([math]::Round($age.TotalHours,1))h old — consider rescanning." 'warning'
    }
    try {
        $snapshot = Import-Clixml -Path $snapshotPath
        $Script:ScanData = $snapshot
        $Script:AllGPOs.Clear()
        foreach ($g in $snapshot.GPOs) { [void]$Script:AllGPOs.Add($g) }
        $Script:AllSettings.Clear()
        foreach ($s in $snapshot.Settings) { [void]$Script:AllSettings.Add($s) }
        if ($snapshot.Apps) {
            $Script:AllIntuneApps.Clear()
            foreach ($a in $snapshot.Apps) { [void]$Script:AllIntuneApps.Add($a) }
        }

        Update-Dashboard $snapshot
        Find-Conflicts
        Apply-GPOFilter
        Apply-SettingFilter
        Set-Status 'Restored from cache' '#22C55E'
        $ts = if ($snapshot.Timestamp -is [datetime]) { $snapshot.Timestamp.ToString('HH:mm') } else { "$($snapshot.Timestamp)" }
        Show-Toast 'Scan Restored' "$($snapshot.GPOs.Count) GPOs, $($snapshot.Settings.Count) settings (from $ts)" 'success'
        Write-DebugLog "Restored scan snapshot: $($snapshot.GPOs.Count) GPOs, $($snapshot.Settings.Count) settings from $($snapshot.Domain) ($ts)" -Level SUCCESS
        return $true
    } catch {
        Write-DebugLog "Failed to restore scan snapshot: $($_.Exception.Message)" -Level ERROR
        Show-Toast 'Restore Failed' $_.Exception.Message 'error'
        return $false
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 16: EXPORT - CSV
# ═══════════════════════════════════════════════════════════════════════════════

function Export-SettingsCsv {
    if (-not $Script:ScanData) {
        Show-Toast 'No Data' 'Run a scan first.' 'warning'
        return
    }
    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Filter = 'CSV files (*.csv)|*.csv'
    $dlg.FileName = "GPO_AllSettings_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $dlg.InitialDirectory = $Script:ReportsDir
    if (-not $dlg.ShowDialog()) { return }

    $Script:ScanData.Settings |
        Select-Object GPOName, GPOGuid, Scope, Category, PolicyName, State, RegistryKey, ValueData |
        Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8

    Write-DebugLog "CSV exported: $($dlg.FileName)" -Level SUCCESS
    Show-Toast 'CSV Exported' "$($Script:ScanData.Settings.Count) rows saved" 'success'
    Unlock-Achievement 'csv_export'
}

function Export-ConflictsCsv {
    if (-not $Script:ScanData) {
        Show-Toast 'No Data' 'Run a scan first.' 'warning'
        return
    }
    $conflicts = Find-Conflicts $Script:ScanData.Settings
    if ($conflicts.Count -eq 0) {
        Show-Toast 'No Conflicts' 'No duplicates or conflicts found.' 'info'
        return
    }
    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Filter = 'CSV files (*.csv)|*.csv'
    $dlg.FileName = "GPO_Conflicts_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $dlg.InitialDirectory = $Script:ReportsDir
    if (-not $dlg.ShowDialog()) { return }

    $conflicts |
        Select-Object Severity, SettingKey, RegistryPath, Scope, Category, GPONames, GPOCount, Values, WinnerGPO |
        Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8

    Write-DebugLog "Conflicts CSV exported: $($dlg.FileName)" -Level SUCCESS
    Show-Toast 'Conflicts Exported' "$($conflicts.Count) rows saved" 'success'
}

# Wire export buttons
$ui.BtnExportHtml.Add_Click({ Export-Html })
$ui.BtnExportCsv.Add_Click({ Export-SettingsCsv })
$ui.BtnExportConflictsCsv.Add_Click({ Export-ConflictsCsv })

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 17: SNAPSHOT SAVE / LOAD
# ═══════════════════════════════════════════════════════════════════════════════

$ui.BtnSaveSnapshot.Add_Click({
    if (-not $Script:ScanData) {
        Show-Toast 'No Data' 'Run a scan first.' 'warning'
        return
    }
    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Filter = 'JSON files (*.json)|*.json'
    $dlg.FileName = "GPO_Snapshot_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    $dlg.InitialDirectory = $Script:SnapshotDir
    if (-not $dlg.ShowDialog()) { return }

    $snapshot = @{
        Version   = $Script:AppVersion
        Timestamp = if ($Script:ScanData.Timestamp -is [datetime]) { $Script:ScanData.Timestamp.ToString('o') } else { $Script:ScanData.Timestamp }
        Domain    = $Script:ScanData.Domain
        GPOs      = @($Script:ScanData.GPOs | ForEach-Object { $_ | Select-Object * })
        Settings  = @($Script:ScanData.Settings | ForEach-Object { $_ | Select-Object * })
    }
    $snapshot | ConvertTo-Json -Depth 6 -Compress | Set-Content -Path $dlg.FileName -Encoding UTF8 -Force
    Write-DebugLog "Snapshot saved: $($dlg.FileName)" -Level SUCCESS
    Show-Toast 'Snapshot Saved' "Saved $($snapshot.GPOs.Count) GPOs to snapshot" 'success'
    Unlock-Achievement 'first_snapshot'
})

$ui.BtnLoadSnapshot.Add_Click({
    $dlg = [Microsoft.Win32.OpenFileDialog]::new()
    $dlg.Filter = 'JSON files (*.json)|*.json'
    $dlg.InitialDirectory = $Script:SnapshotDir
    if (-not $dlg.ShowDialog()) { return }

    try {
        $json = [System.IO.File]::ReadAllText($dlg.FileName) | ConvertFrom-Json
        $gpoList = [System.Collections.Generic.List[PSCustomObject]]::new()
        foreach ($g in $json.GPOs) { [void]$gpoList.Add($g) }
        $settingsList = [System.Collections.Generic.List[PSCustomObject]]::new()
        foreach ($s in $json.Settings) { [void]$settingsList.Add($s) }

        $Script:ScanData = @{
            Timestamp  = [DateTime]::Parse($json.Timestamp)
            Domain     = $json.Domain
            GPOs       = $gpoList
            Settings   = $settingsList
        }

        $Script:AllGPOs.Clear()
        foreach ($g in $gpoList) { [void]$Script:AllGPOs.Add($g) }
        $Script:AllSettings.Clear()
        foreach ($s in $settingsList) { [void]$Script:AllSettings.Add($s) }

        $conflictResults = Find-Conflicts $settingsList
        $Script:AllConflicts.Clear()
        foreach ($c in $conflictResults) { [void]$Script:AllConflicts.Add($c) }

        Update-Dashboard
        Show-Toast 'Snapshot Loaded' "Loaded $($gpoList.Count) GPOs from $(Split-Path -Leaf $dlg.FileName)" 'success'
        Write-DebugLog "Snapshot loaded: $($dlg.FileName)" -Level SUCCESS
    } catch {
        Write-DebugLog "Snapshot load error: $($_.Exception.Message)" -Level ERROR
        Show-Toast 'Load Failed' $_.Exception.Message 'error'
    }
})


# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 17.5: NEW FEATURE HANDLERS
# ═══════════════════════════════════════════════════════════════════════════════

# --- #2: Detail Pane - Settings grid selection ---
$ui.SettingsGrid.Add_SelectionChanged({
    $sel = $ui.SettingsGrid.SelectedItem
    if (-not $sel -or -not $ui.DetailPaneText) { return }
    $ui.DetailPaneText.Text = Show-DetailPane -Item $sel -Mode 'Setting'
    if ($ui.SettingsDetailPane) { $ui.SettingsDetailPane.Visibility = 'Visible' }
})

# --- #2: Detail Pane close button ---
if ($ui.BtnCloseDetailPane) {
    $ui.BtnCloseDetailPane.Add_Click({
        if ($ui.SettingsDetailPane) { $ui.SettingsDetailPane.Visibility = 'Collapsed' }
        $ui.SettingsGrid.UnselectAll()
    })
}

# --- #2: Detail Pane - Conflicts grid selection ---
$ui.ConflictsGrid.Add_SelectionChanged({
    $sel = $ui.ConflictsGrid.SelectedItem
    if (-not $sel -or -not $ui.DetailPaneText) { return }
    $ui.DetailPaneText.Text = Show-DetailPane -Item $sel -Mode 'Conflict'
})

# --- #2: Detail Pane - IntuneApps grid selection ---
$ui.IntuneAppsGrid.Add_SelectionChanged({
    $sel = $ui.IntuneAppsGrid.SelectedItem
    if (-not $sel -or -not $ui.DetailPaneText) { return }
    $ui.DetailPaneText.Text = Show-DetailPane -Item $sel -Mode 'IntuneApp'
})

# --- #9: Context Menus (built in code, applied to DataGrids) ---
function Build-GridContextMenu([string]$GridName, [string]$ExportPrefix) {
    $cm = [System.Windows.Controls.ContextMenu]::new()
    $cm.Background = (Get-CachedBrush '#FF1E1E1E')
    $cm.Foreground = (Get-CachedBrush '#FFE0E0E0')
    $cm.BorderBrush = (Get-CachedBrush '#FF333333')
    $gridRef = $ui[$GridName]  # capture for closures ($ui and $this are null inside .GetNewClosure())

    $miCopyRow = [System.Windows.Controls.MenuItem]::new()
    $miCopyRow.Header = 'Copy Row'
    $miCopyRow.Add_Click({
        if ($gridRef -and $gridRef.SelectedItem) { Copy-GridRowToClipboard $gridRef.SelectedItem }
    }.GetNewClosure())
    [void]$cm.Items.Add($miCopyRow)

    $miCopyCell = [System.Windows.Controls.MenuItem]::new()
    $miCopyCell.Header = 'Copy Cell Value'
    $miCopyCell.Add_Click({
        if ($gridRef -and $gridRef.CurrentCell -and $gridRef.SelectedItem) {
            $colHeader = $gridRef.CurrentCell.Column.Header
            $propName = if ($gridRef.CurrentCell.Column.Binding) { $gridRef.CurrentCell.Column.Binding.Path.Path } else { $colHeader }
            Copy-GridCellToClipboard $gridRef.SelectedItem $propName
        }
    }.GetNewClosure())
    [void]$cm.Items.Add($miCopyCell)

    $sep = [System.Windows.Controls.Separator]::new()
    [void]$cm.Items.Add($sep)

    $miExportSel = [System.Windows.Controls.MenuItem]::new()
    $miExportSel.Header = 'Export Selected to CSV'
    $miExportSel.Add_Click({
        if ($gridRef) {
            $items = [System.Collections.Generic.List[object]]::new()
            foreach ($item in $gridRef.SelectedItems) { [void]$items.Add($item) }
            Export-SelectedToCsv -Items $items -DefaultName $GridName
        }
    }.GetNewClosure())
    [void]$cm.Items.Add($miExportSel)


    [void]$cm.Items.Add([System.Windows.Controls.Separator]::new())

    # Open in Notepad
    $miNotepad = [System.Windows.Controls.MenuItem]::new()
    $miNotepad.Header = 'Open in Notepad'
    $miNotepad.Add_Click({
        if ($gridRef -and $gridRef.SelectedItem) {
            $props = $gridRef.SelectedItem.PSObject.Properties | ForEach-Object { "$($_.Name): $($_.Value)" }
            $text = $props -join "`r`n"
            $tmpFile = [IO.Path]::Combine([IO.Path]::GetTempPath(), "PolicyPilot_$([DateTime]::Now.Ticks).txt")
            [IO.File]::WriteAllText($tmpFile, $text, [System.Text.Encoding]::UTF8)
            Start-Process notepad.exe $tmpFile
        }
    }.GetNewClosure())
    [void]$cm.Items.Add($miNotepad)

    # Search on Microsoft Learn
    $miLearn = [System.Windows.Controls.MenuItem]::new()
    $miLearn.Header = 'Search on Microsoft Learn'
    $miLearn.Add_Click({
        if ($gridRef -and $gridRef.SelectedItem) {
            $term = if ($gridRef.SelectedItem.PolicyName) { $gridRef.SelectedItem.PolicyName }
                    elseif ($gridRef.SelectedItem.SettingKey) { $gridRef.SelectedItem.SettingKey }
                    elseif ($gridRef.SelectedItem.AppName) { $gridRef.SelectedItem.AppName }
                    else { "$($gridRef.SelectedItem)" }
            $url = "https://learn.microsoft.com/en-us/search/?terms=$([Uri]::EscapeDataString($term))"
            Start-Process $url
        }
    }.GetNewClosure())
    [void]$cm.Items.Add($miLearn)

    # Decode Base64 (cell value)
    $miBase64 = [System.Windows.Controls.MenuItem]::new()
    $miBase64.Header = 'Decode Base64 Value'
    $miBase64.Add_Click({
        if ($gridRef -and $gridRef.CurrentCell -and $gridRef.SelectedItem) {
            $propName = if ($gridRef.CurrentCell.Column.Binding) { $gridRef.CurrentCell.Column.Binding.Path.Path } else { $gridRef.CurrentCell.Column.Header }
            $val = "$($gridRef.SelectedItem.$propName)"
            if ($val) {
                $decoded = Invoke-Base64Decode $val
                [System.Windows.Clipboard]::SetText($decoded)
                Show-Toast 'Base64 Decoded' "Result copied to clipboard" 'success'
            }
        }
    }.GetNewClosure())
    [void]$cm.Items.Add($miBase64)

    return $cm
}

# Apply context menus to all DataGrids
foreach ($gridName in @('GPOListGrid','SettingsGrid','ConflictsGrid','IntuneAppsGrid','DashTopConflictsGrid')) {
    if ($ui[$gridName]) {
        $ui[$gridName].ContextMenu = Build-GridContextMenu $gridName $gridName
    }
}


# --- #6: Search Highlighting via LoadingRow ---
$Script:SearchHighlightBrush = (Get-CachedBrush '#40F59E0B')
$Script:SearchHighlightTerm = ''

# Track search terms globally
$ui.TxtGPOSearch.Add_TextChanged({
    $Script:SearchHighlightTerm = $ui.TxtGPOSearch.Text.Trim()
    if ($ui.GPOListGrid) { $ui.GPOListGrid.Items.Refresh() }
})
$ui.TxtSettingSearch.Add_TextChanged({
    $Script:SearchHighlightTerm = $ui.TxtSettingSearch.Text.Trim()
    if ($ui.SettingsGrid) { $ui.SettingsGrid.Items.Refresh() }
})

# Highlight matching rows on GPO grid
$ui.GPOListGrid.Add_LoadingRow({
    param($s, $e)
    $item = $e.Row.DataContext
    if (-not $item -or -not $Script:SearchHighlightTerm) {
        $e.Row.Background = [System.Windows.Media.Brushes]::Transparent
        return
    }
    $term = $Script:SearchHighlightTerm
    if ($item.DisplayName -like "*$term*") {
        $e.Row.Background = $Script:SearchHighlightBrush
    } else {
        $e.Row.Background = [System.Windows.Media.Brushes]::Transparent
    }
})

# Highlight matching rows on Settings grid
$ui.SettingsGrid.Add_LoadingRow({
    param($s, $e)
    $item = $e.Row.DataContext
    if (-not $item -or -not $Script:SearchHighlightTerm) {
        $e.Row.Background = [System.Windows.Media.Brushes]::Transparent
        return
    }
    $term = $Script:SearchHighlightTerm
    $match = $item.PolicyName -like "*$term*" -or $item.Category -like "*$term*" -or $item.GPOName -like "*$term*" -or "$($item.ValueData)" -like "*$term*"
    if ($match) {
        $e.Row.Background = $Script:SearchHighlightBrush
    } else {
        $e.Row.Background = [System.Windows.Media.Brushes]::Transparent
    }
})
# --- #5: Snapshot Diff button ---
if ($ui.BtnCompareSnapshot) {
    $ui.BtnCompareSnapshot.Add_Click({
        if (-not $Script:ScanData) {
            Show-Toast 'No Data' 'Run a scan first, then compare against a saved snapshot.' 'warning'
            return
        }
        $dlg = [Microsoft.Win32.OpenFileDialog]::new()
        $dlg.Filter = 'JSON Snapshot (*.json)|*.json'
        $dlg.Title = 'Select Baseline Snapshot to Compare'
        $dlg.InitialDirectory = $Script:SnapshotDir
        if (-not $dlg.ShowDialog()) { return }

        try {
            $json = [System.IO.File]::ReadAllText($dlg.FileName) | ConvertFrom-Json
            $baseSettings = [System.Collections.Generic.List[PSCustomObject]]::new()
            foreach ($s in $json.Settings) { [void]$baseSettings.Add($s) }
            $baseGPOs = [System.Collections.Generic.List[PSCustomObject]]::new()
            foreach ($g in $json.GPOs) { [void]$baseGPOs.Add($g) }
            $baseline = @{ Settings = $baseSettings; GPOs = $baseGPOs; Domain = $json.Domain; Timestamp = $json.Timestamp }

            $diffs = Compare-Snapshots -Baseline $baseline -Current $Script:ScanData
            if ($diffs.Count -eq 0) {
                Show-Toast 'No Differences' 'Current scan matches the baseline snapshot.' 'success'
                return
            }

            # Show diff results in report preview
            $sb = [System.Text.StringBuilder]::new()
            $null = $sb.AppendLine("SNAPSHOT COMPARISON")
            $null = $sb.AppendLine("$([char]0x2500)" * 60)
            $null = $sb.AppendLine("Baseline: $(Split-Path -Leaf $dlg.FileName)")
            $null = $sb.AppendLine("Current:  $(Get-Date -Format 'yyyy-MM-dd HH:mm')")
            $null = $sb.AppendLine("Changes:  $($diffs.Count)")
            $null = $sb.AppendLine()

            $added    = @($diffs | Where-Object Change -like '*Added*')
            $removed  = @($diffs | Where-Object Change -like '*Removed*')
            $modified = @($diffs | Where-Object Change -eq 'Modified')

            if ($added.Count -gt 0) {
                $null = $sb.AppendLine("ADDED ($($added.Count))")
                foreach ($d in $added) { $null = $sb.AppendLine("  + $($d.PolicyName)  [$($d.GPOName)]  Value: $($d.NewValue)") }
                $null = $sb.AppendLine()
            }
            if ($removed.Count -gt 0) {
                $null = $sb.AppendLine("REMOVED ($($removed.Count))")
                foreach ($d in $removed) { $null = $sb.AppendLine("  - $($d.PolicyName)  [$($d.GPOName)]  Was: $($d.OldValue)") }
                $null = $sb.AppendLine()
            }
            if ($modified.Count -gt 0) {
                $null = $sb.AppendLine("MODIFIED ($($modified.Count))")
                foreach ($d in $modified) { $null = $sb.AppendLine("  ~ $($d.PolicyName)  [$($d.GPOName)]"); $null = $sb.AppendLine("    Old: $($d.OldValue)  ->  New: $($d.NewValue)") }
                $null = $sb.AppendLine()
            }

            if ($ui.ReportPreviewText) { $ui.ReportPreviewText.Text = $sb.ToString() }
            Switch-Tab 'Report'
            Show-Toast 'Comparison Complete' "$($diffs.Count) differences found" 'info'
    Unlock-Achievement 'snapshot_compare'
        } catch {
            Write-DebugLog "Snapshot compare error: $($_.Exception.Message)" -Level ERROR
            Show-Toast 'Compare Failed' $_.Exception.Message 'error'
        }
    })
}

# --- Wire new export buttons ---
if ($ui.BtnExportScript) { $ui.BtnExportScript.Add_Click({ Export-AsScript }) }
if ($ui.BtnPrintReport)  { $ui.BtnPrintReport.Add_Click({ Invoke-PrintReport }) }
if ($ui.BtnExportReg)    { $ui.BtnExportReg.Add_Click({ Export-RegistryFile }) }

# --- #12: Baseline buttons ---
if ($ui.BtnImportBaseline)  { $ui.BtnImportBaseline.Add_Click({ Import-Baseline }) }
if ($ui.BtnExportBaseline)  { $ui.BtnExportBaseline.Add_Click({ Export-BaselineTemplate }) }

# --- N7: Import MDM XML button ---
if ($ui.BtnImportMdmXml) { $ui.BtnImportMdmXml.Add_Click({ Import-MdmXml }) }

# --- #27: Impact Simulator button ---
if ($ui.BtnSimulateImpact) {
    $ui.BtnSimulateImpact.Add_Click({
        $sel = $ui.GPOListGrid.SelectedItem
        if (-not $sel) {
            Show-Toast 'No GPO Selected' 'Select a GPO in the GPO List tab first.' 'info'
            return
        }
        $result = Invoke-ImpactSimulation -GPONameToRemove $sel.DisplayName
        if (-not $result) { return }

        $sb = [System.Text.StringBuilder]::new()
        $null = $sb.AppendLine("IMPACT SIMULATION: Remove `"$($result.GPORemoved)`"")
        $null = $sb.AppendLine("$([char]0x2500)" * 60)
        $null = $sb.AppendLine("Settings lost:       $($result.SettingsLost)")
        $null = $sb.AppendLine("Settings remaining:  $($result.SettingsRemaining)")
        $null = $sb.AppendLine("Conflicts before:    $($result.ConflictsBefore)")
        $null = $sb.AppendLine("Conflicts after:     $($result.ConflictsAfter)")
        $null = $sb.AppendLine("Conflicts resolved:  $($result.ConflictsResolved)")
        $null = $sb.AppendLine()

        if ($result.OrphanedSettings.Count -gt 0) {
            $null = $sb.AppendLine("ORPHANED SETTINGS (no other GPO provides these):")
            foreach ($s in $result.OrphanedSettings) {
                $null = $sb.AppendLine("  ! $($s.PolicyName) [$($s.Category)]")
            }
        } else {
            $null = $sb.AppendLine("No orphaned settings - all settings have coverage from other GPOs.")
        }

        if ($ui.ReportPreviewText) { $ui.ReportPreviewText.Text = $sb.ToString() }
        Switch-Tab 'Report'
        Show-Toast 'Simulation Complete' "Removing $($result.GPORemoved): $($result.SettingsLost) settings affected" 'info'
    })
}

# ── OU Tree View Handler ──────────────────────────────────────────────────
if ($ui.BtnOUTreeView) {
    $ui.BtnOUTreeView.Add_Click({
        $sel = $ui.GPOListGrid.SelectedItem
        if (-not $sel) {
            $ui.ReportPreviewText.Text = "Select a GPO first to view its OU tree."
            return
        }
        $gpoName = $sel.Name
        $gpoData = $Script:ScanData.GPOs | Where-Object { $_.Name -eq $gpoName } | Select-Object -First 1
        if ($gpoData) {
            $tree = Build-OUTreeText -GPO $gpoData
            $ui.ReportPreviewText.Text = $tree
        } else {
            $ui.ReportPreviewText.Text = "No detailed data found for $gpoName."
        }
    }.GetNewClosure())
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 18: THEME & APP SETTINGS HANDLERS
# ═══════════════════════════════════════════════════════════════════════════════



# -- Help button (title bar) --

$ui.BtnHelp = $Window.FindName('BtnHelp')

if ($ui.BtnHelp) {

    $ui.BtnHelp.Add_Click({

        $HelpText = "KEYBOARD SHORTCUTS`n" +

            "-------------------------------------------`n" +

            "Ctrl+E          Export HTML report`n" +

            "Ctrl+S          Save snapshot`n" +

            "Ctrl+L          Toggle light/dark mode`n" +

            "F1              Show this help dialog`n" +

            "F5              Re-scan policies`n" +

            "Ctrl+1..7       Switch tabs`n" +

            "Ctrl+P          Print / Open report in browser`n" +

            "Ctrl+D          Compare snapshot (diff)`n" +

            "Ctrl+R          Re-scan policies`n`n" +

            "SCAN MODES`n" +

            "-------------------------------------------`n" +

            "Local Machine   Uses gpresult /x to read`n" +

            "                policies applied to this PC`n" +

            "                (no RSAT required)`n`n" +

            "Active Directory  Uses Get-GPO -All from`n" +

            "                  GroupPolicy RSAT module`n`n" +

            "Intune          Reads device configuration`n" +

            "                profiles, compliance policies`n" +

            "                and settings catalog via`n" +

            "                Microsoft Graph`n`n" +

            "TABS`n" +

            "-------------------------------------------`n" +

            "Dashboard       Overview cards, quick stats`n" +

            "GPO List        Full inventory with filters`n" +

            "Settings        All settings with search`n" +

            "Conflicts       Conflicting/redundant items`n" +

            "Report          HTML preview`n" +

            "App Settings    Theme, scan mode, exports`n`n" +

            "TIPS`n" +

            "-------------------------------------------`n" +

            "* Activity log in bottom panel`n" +

            "* Toggle via status bar button`n" +

            "* Export HTML, CSV, or JSON snapshot`n" +

            "* Load snapshots for offline comparison"



        $Palette = if ($Script:Prefs.IsLightMode) { $Script:ThemeLight } else { $Script:ThemeDark }

        $Br = { param([string]$Key) (Get-CachedBrush $Palette[$Key]) }

        $DlgHeight = [Math]::Max(460, $Window.ActualHeight - 120)



        $Dlg = New-Object System.Windows.Window

        $Dlg.Title = 'Help - PolicyPilot'

        $Dlg.Width = 520; $Dlg.Height = $DlgHeight

        $Dlg.ResizeMode = 'NoResize'

        $Dlg.WindowStartupLocation = 'CenterOwner'

        $Dlg.Owner = $Window

        $Dlg.Background = (& $Br 'ThemeCardBg')

        $Dlg.WindowStyle = 'None'

        $Dlg.AllowsTransparency = $true



        $OuterBorder = New-Object System.Windows.Controls.Border

        $OuterBorder.Background = (& $Br 'ThemeCardBg')

        $OuterBorder.BorderBrush = (& $Br 'ThemeBorderElevated')

        $OuterBorder.BorderThickness = [System.Windows.Thickness]::new(1)

        $OuterBorder.CornerRadius = [System.Windows.CornerRadius]::new(12)

        $OuterBorder.Padding = [System.Windows.Thickness]::new(28, 24, 28, 20)

        $OuterBorder.Effect = [System.Windows.Media.Effects.DropShadowEffect]@{

            Color = [System.Windows.Media.Colors]::Black; Direction = 270; ShadowDepth = 8; BlurRadius = 28; Opacity = 0.45

        }



        $Root = New-Object System.Windows.Controls.DockPanel

        $Root.LastChildFill = $true



        # Header

        $HdrGrid = New-Object System.Windows.Controls.Grid

        $HdrGrid.Margin = [System.Windows.Thickness]::new(0, 0, 0, 16)

        $Col1 = New-Object System.Windows.Controls.ColumnDefinition; $Col1.Width = [System.Windows.GridLength]::new(0, [System.Windows.GridUnitType]::Auto)

        $Col2 = New-Object System.Windows.Controls.ColumnDefinition; $Col2.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)

        [void]$HdrGrid.ColumnDefinitions.Add($Col1); $HdrGrid.ColumnDefinitions.Add($Col2)

        $Badge = New-Object System.Windows.Controls.Border

        $Badge.Width = 36; $Badge.Height = 36; $Badge.CornerRadius = [System.Windows.CornerRadius]::new(8)

        $Badge.Background = (& $Br 'ThemeAccentDim')

        $BadgeT = New-Object System.Windows.Controls.TextBlock; $BadgeT.Text = '?'; $BadgeT.FontSize = 17; $BadgeT.FontWeight = 'Bold'

        $BadgeT.Foreground = (& $Br 'ThemeAccentLight'); $BadgeT.HorizontalAlignment = 'Center'; $BadgeT.VerticalAlignment = 'Center'

        $Badge.Child = $BadgeT

        [void][System.Windows.Controls.Grid]::SetColumn($Badge, 0); $HdrGrid.Children.Add($Badge)

        $TitleT = New-Object System.Windows.Controls.TextBlock

        $TitleT.Text = "Help - PolicyPilot v$($Script:AppVersion)"; $TitleT.FontSize = 15; $TitleT.FontWeight = 'Bold'

        $TitleT.Foreground = (& $Br 'ThemeTextPrimary'); $TitleT.VerticalAlignment = 'Center'; $TitleT.Margin = [System.Windows.Thickness]::new(14,0,0,0)

        [void][System.Windows.Controls.Grid]::SetColumn($TitleT, 1); $HdrGrid.Children.Add($TitleT)

        [System.Windows.Controls.DockPanel]::SetDock($HdrGrid, [System.Windows.Controls.Dock]::Top)

        [void]$Root.Children.Add($HdrGrid)



        $Sep = New-Object System.Windows.Controls.Border; $Sep.Height = 1

        $Sep.Background = (& $Br 'ThemeBorder'); $Sep.Margin = [System.Windows.Thickness]::new(0, 0, 0, 14)

        [System.Windows.Controls.DockPanel]::SetDock($Sep, [System.Windows.Controls.Dock]::Top)

        [void]$Root.Children.Add($Sep)



        # Close button

        $BtnPanel = New-Object System.Windows.Controls.StackPanel

        $BtnPanel.Orientation = 'Horizontal'; $BtnPanel.HorizontalAlignment = 'Right'

        $BtnPanel.Margin = [System.Windows.Thickness]::new(0, 14, 0, 0)

        $CloseBtn = New-Object System.Windows.Controls.Button

        $CloseBtn.Content = 'Close'; $CloseBtn.MinWidth = 80; $CloseBtn.Padding = [System.Windows.Thickness]::new(16, 8, 16, 8)

        $CloseBtn.FontSize = 12; $CloseBtn.Cursor = [System.Windows.Input.Cursors]::Hand

        $CloseBtn.Background = (& $Br 'ThemeCardAltBg'); $CloseBtn.Foreground = (& $Br 'ThemeTextBody')

        $CloseBtn.BorderBrush = (& $Br 'ThemeBorder'); $CloseBtn.BorderThickness = [System.Windows.Thickness]::new(1)

        $DlgRef = $Dlg

        $CloseBtn.Add_Click({ $DlgRef.Close() }.GetNewClosure())

        [void]$BtnPanel.Children.Add($CloseBtn)

        [System.Windows.Controls.DockPanel]::SetDock($BtnPanel, [System.Windows.Controls.Dock]::Bottom)

        [void]$Root.Children.Add($BtnPanel)



        # Content

        $SV = New-Object System.Windows.Controls.ScrollViewer

        $SV.VerticalScrollBarVisibility = 'Auto'; $SV.HorizontalScrollBarVisibility = 'Disabled'

        $MsgTB = New-Object System.Windows.Controls.TextBlock

        $MsgTB.Text = $HelpText; $MsgTB.FontSize = 11.5

        $MsgTB.FontFamily = [System.Windows.Media.FontFamily]::new('Cascadia Code, Cascadia Mono, Consolas, Courier New')

        $MsgTB.Foreground = (& $Br 'ThemeTextSecondary'); $MsgTB.TextWrapping = 'NoWrap'; $MsgTB.LineHeight = 18

        $SV.Content = $MsgTB

        [void]$Root.Children.Add($SV)



        $OuterBorder.Child = $Root; $Dlg.Content = $OuterBorder

        $Dlg.ShowDialog()

    })

}



# -- Theme toggle button (title bar) --

$ui.BtnThemeToggle = $Window.FindName('BtnThemeToggle')

if ($ui.BtnThemeToggle) {

    $ui.BtnThemeToggle.Add_Click({

        $newLight = -not $Script:Prefs.IsLightMode

        if ($newLight) {

            Set-Theme $Script:ThemeLight -IsLight $true

            $Script:Prefs.IsLightMode = $true

            $ui.BtnThemeLight.Tag = 'Active'; $ui.BtnThemeDark.Tag = $null

            $ui.BtnThemeToggle.Content = [char]0x263E

        } else {

            Set-Theme $Script:ThemeDark -IsLight $false

            $Script:Prefs.IsLightMode = $false

            $ui.BtnThemeDark.Tag = 'Active'; $ui.BtnThemeLight.Tag = $null

            $ui.BtnThemeToggle.Content = [char]0x2600

        }

        Save-Preferences
        Unlock-Achievement 'theme_toggle'

    })

}
$ui.BtnThemeDark.Add_Click({
    Set-Theme $Script:ThemeDark -IsLight $false
    $Script:Prefs.IsLightMode = $false
    $ui.BtnThemeDark.Tag = 'Active'
    $ui.BtnThemeLight.Tag = $null
    if ($ui.BtnThemeHC) { $ui.BtnThemeHC.Tag = $null }
    Save-Preferences
})
$ui.BtnThemeLight.Add_Click({
    Set-Theme $Script:ThemeLight -IsLight $true
    $Script:Prefs.IsLightMode = $true
    $ui.BtnThemeLight.Tag = 'Active'
    $ui.BtnThemeDark.Tag = $null
    if ($ui.BtnThemeHC) { $ui.BtnThemeHC.Tag = $null }
    Save-Preferences
})
# Domain/DC override → save on text change
$ui.TxtDomainOverride.Add_LostFocus({ $Script:Prefs.DomainOverride = $ui.TxtDomainOverride.Text.Trim(); Save-Preferences })
if ($ui.BtnThemeHC) {
    $ui.BtnThemeHC.Add_Click({
        Set-Theme $Script:ThemeHighContrast -IsLight $false
        $Script:Prefs.IsLightMode = $false
        $ui.BtnThemeHC.Tag = 'Active'
        $ui.BtnThemeDark.Tag = $null
        $ui.BtnThemeLight.Tag = $null
        Save-Preferences
    })
}
$ui.TxtDCOverride.Add_LostFocus({ $Script:Prefs.DCOverride = $ui.TxtDCOverride.Text.Trim(); Save-Preferences })

# Detect My DC button — discovers a DC for the target domain
if ($ui.BtnDetectDC) {
    $ui.BtnDetectDC.Add_Click({
        if ($ui.TxtDetectDCStatus) { $ui.TxtDetectDCStatus.Text = 'Discovering...' }
        try {
            $targetDomain = if ($Script:Prefs.DomainOverride) { $Script:Prefs.DomainOverride } else { $null }
            if ($targetDomain) {
                $ctx = [System.DirectoryServices.ActiveDirectory.DirectoryContext]::new('Domain', $targetDomain)
                $domObj = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($ctx)
            } else {
                $domObj = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            }
            $dc = $domObj.FindDomainController()
            $dcName = $dc.Name
            if ($ui.TxtDCOverride) { $ui.TxtDCOverride.Text = $dcName }
            $Script:Prefs.DCOverride = $dcName
            Save-Preferences
            if ($ui.TxtDetectDCStatus) { $ui.TxtDetectDCStatus.Text = "Detected: $dcName" }
            Write-DebugLog "Detected DC: $dcName for domain $($domObj.Name)" -Level SUCCESS
        } catch {
            if ($ui.TxtDetectDCStatus) { $ui.TxtDetectDCStatus.Text = "Error: $($_.Exception.Message)" }
            Write-DebugLog "Detect DC failed: $($_.Exception.Message)" -Level ERROR
        }
    })
}

# OU Scope → save on text change
if ($ui.TxtOUScope) {
    $ui.TxtOUScope.Add_LostFocus({ $Script:Prefs.OUScope = $ui.TxtOUScope.Text.Trim(); Save-Preferences })
}

# Detect My OU button — discovers the OU of the current computer via LDAP
if ($ui.BtnDetectOU) {
    $ui.BtnDetectOU.Add_Click({
        if ($ui.TxtDetectOUStatus) { $ui.TxtDetectOUStatus.Text = 'Detecting...' }
        try {
            $compName = $env:COMPUTERNAME
            $domain = if ($Script:Prefs.DomainOverride) { $Script:Prefs.DomainOverride } else { $null }
            $root = if ($domain) { "LDAP://$domain" } else { 'LDAP://RootDSE' }
            $rootDSE = [ADSI]$root
            $baseDN = if ($domain) {
                ($domain.Split('.') | ForEach-Object { "DC=$_" }) -join ','
            } else {
                $rootDSE.defaultNamingContext[0]
            }
            $searchRoot = [ADSI]"LDAP://$baseDN"
            $searcher = [System.DirectoryServices.DirectorySearcher]::new($searchRoot)
            $searcher.Filter = "(&(objectCategory=computer)(name=$compName))"
            $searcher.PropertiesToLoad.Add('distinguishedName') | Out-Null
            $result = $searcher.FindOne()
            if ($result) {
                $dn = $result.Properties['distinguishedname'][0]
                # OU is everything after the first comma (strip the computer CN)
                $ou = ($dn -split ',', 2)[1]
                if ($ui.TxtOUScope) { $ui.TxtOUScope.Text = $ou }
                $Script:Prefs.OUScope = $ou
                Save-Preferences
                if ($ui.TxtDetectOUStatus) { $ui.TxtDetectOUStatus.Text = "Detected: $ou" }
                Write-DebugLog "Detected computer OU: $ou (from $dn)" -Level SUCCESS
            } else {
                if ($ui.TxtDetectOUStatus) { $ui.TxtDetectOUStatus.Text = "Computer '$compName' not found in AD" }
                Write-DebugLog "Detect OU: computer '$compName' not found in AD" -Level WARN
            }
        } catch {
            if ($ui.TxtDetectOUStatus) { $ui.TxtDetectOUStatus.Text = "Error: $($_.Exception.Message)" }
            Write-DebugLog "Detect OU failed: $($_.Exception.Message)" -Level ERROR
        }
    })
}

# Force Refresh checkbox
if ($ui.ChkForceRefresh) {
    $ui.ChkForceRefresh.Add_Checked({   $Script:Prefs.ForceRefresh = $true;  Save-Preferences })
    $ui.ChkForceRefresh.Add_Unchecked({ $Script:Prefs.ForceRefresh = $false; Save-Preferences })
}

# Export option checkboxes
$ui.ChkIncludeDisabled.Add_Checked({   $Script:Prefs.IncludeDisabled = $true;  Save-Preferences })
$ui.ChkIncludeDisabled.Add_Unchecked({ $Script:Prefs.IncludeDisabled = $false; Save-Preferences })
$ui.ChkIncludeUnlinked.Add_Checked({   $Script:Prefs.IncludeUnlinked = $true;  Save-Preferences })
$ui.ChkIncludeUnlinked.Add_Unchecked({ $Script:Prefs.IncludeUnlinked = $false; Save-Preferences })
$ui.ChkShowRegistryPaths.Add_Checked({   $Script:Prefs.ShowRegistryPaths = $true;  Save-Preferences })
$ui.ChkShowRegistryPaths.Add_Unchecked({ $Script:Prefs.ShowRegistryPaths = $false; Save-Preferences })

# Recheck prereqs
$ui.BtnRecheckPrereqs.Add_Click({
    if ($ui.PrereqDetailStatus) { $ui.PrereqDetailStatus.Text = 'Checking prerequisites...' }
    if ($ui.PrereqStatus)       { $ui.PrereqStatus.Text = 'Checking...' }
    Start-BackgroundWork -Work {
        param($SyncH)
        $results = [System.Collections.Generic.List[string]]::new()
        $allOk = $true
        $rsatMissing = $false
        $mode = $SyncH.ScanMode

        [void]$results.Add("Scan Mode: $mode")
        [void]$results.Add("")

        if ($mode -eq 'AD') {
            $gpMod = Get-Module -ListAvailable -Name GroupPolicy -ErrorAction SilentlyContinue
            if ($gpMod) { [void]$results.Add("[OK]  GroupPolicy module v$($gpMod.Version)") }
            else {
                [void]$results.Add("[FAIL] GroupPolicy module NOT found")
                [void]$results.Add("       Add-WindowsCapability -Online -Name Rsat.GroupPolicy.Management.Tools~~~~0.0.1.0")
                [void]$results.Add("       [TIP] Click 'Install RSAT GP Tools' below")
                $allOk = $false; $rsatMissing = $true
            }
            try {
                $dom = if ($SyncH.DomainOverride) { $SyncH.DomainOverride } else { [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name }
                [void]$results.Add("[OK]  Domain: $dom")
            } catch { [void]$results.Add("[FAIL] Cannot reach AD domain"); $allOk = $false }
            if ($gpMod) {
                try {
                    Import-Module GroupPolicy -ErrorAction Stop
                    $t = Get-GPO -All -ErrorAction Stop | Select-Object -First 1
                    if ($t) { [void]$results.Add("[OK]  GPO read access confirmed") }
                    else { [void]$results.Add("[WARN] Get-GPO returned 0 GPOs") }
                } catch { [void]$results.Add("[FAIL] Cannot read GPOs: $($_.Exception.Message)"); $allOk = $false }
            }
        } elseif ($mode -eq 'Intune') {
            $gm = Get-Module -ListAvailable Microsoft.Graph.DeviceManagement -ErrorAction SilentlyContinue
            $gl = Get-Module -ListAvailable Microsoft.Graph.Intune -ErrorAction SilentlyContinue
            if ($gm -or $gl) { [void]$results.Add("[OK]  Microsoft.Graph module found") }
            else { [void]$results.Add("[FAIL] Microsoft.Graph not found"); $allOk = $false }
        } else {
            $gpr = Get-Command gpresult.exe -ErrorAction SilentlyContinue
            if ($gpr) { [void]$results.Add("[OK]  gpresult.exe found") } else { [void]$results.Add("[FAIL] gpresult.exe not found"); $allOk = $false }
            try { $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop; if ($cs.PartOfDomain) { [void]$results.Add("[OK]  Domain-joined: $($cs.Domain)") } else { [void]$results.Add("[WARN] Not domain-joined") } } catch { [void]$results.Add("[WARN] Cannot determine domain") }
            [void]$results.Add(""); [void]$results.Add("Local mode: gpresult (no RSAT required).")
        }
        return @{ Passed = $allOk; Details = ($results -join "`n"); RsatMissing = $rsatMissing }
    } -OnComplete {
        param($Results, $Errors)
        $r = $Results | Select-Object -First 1
        if ($r) {
            $Script:PrereqsMet = $r.Passed
            $Script:RsatMissing = $r.RsatMissing
            if ($ui.PrereqDetailStatus) { $ui.PrereqDetailStatus.Text = $r.Details }
            if ($ui.PrereqStatus)       { $ui.PrereqStatus.Text = $r.Details }
            if ($r.Passed) {
                Show-Toast 'Prerequisites OK' 'All checks passed' 'success'
            } else {
                Show-Toast 'Prerequisites Failed' 'See App Settings for details' 'error'
                if ($r.RsatMissing) {
                    $answer = Show-ThemedMessageBox -Message "RSAT Group Policy Management Tools are not installed.`n`nInstall now via Features on Demand?`nRequires admin and may take a few minutes." `
                        -Title 'Install RSAT GP Tools' -Buttons 'YesNo' -Icon 'Question'
                    if ($answer -eq 'Yes') { Install-RsatGpTools }
                }
            }
        }
    }.GetNewClosure() -Variables @{ ScanMode = $Script:Prefs.ScanMode; DomainOverride = $Script:Prefs.DomainOverride } -Context @{ Name = 'PrereqRecheck' }
})

# ═══════════════════════════════════════════════════════════════════════════════
# Console panel handlers
if ($ui.btnToggleConsole) {
    $ui.btnToggleConsole.Add_Click({
        $mainGrid = $ui.WindowBorder.Child
        if ($ui.pnlBottomPanel.Visibility -eq 'Visible') {
            $ui.pnlBottomPanel.Visibility = 'Collapsed'
            if ($ui.splitterBottom) { $ui.splitterBottom.Visibility = 'Collapsed' }
            $mainGrid.RowDefinitions[3].Height = [System.Windows.GridLength]::new(0)
            $mainGrid.RowDefinitions[4].MinHeight = 0
            $mainGrid.RowDefinitions[4].Height = [System.Windows.GridLength]::new(0)
        } else {
            $mainGrid.RowDefinitions[3].Height = [System.Windows.GridLength]::new(4)
            $mainGrid.RowDefinitions[4].MinHeight = 80
            $mainGrid.RowDefinitions[4].Height = [System.Windows.GridLength]::new(200)
            $ui.pnlBottomPanel.Visibility = 'Visible'
            if ($ui.splitterBottom) { $ui.splitterBottom.Visibility = 'Visible' }
        }
    })
}
if ($ui.btnHideBottom) {
    $ui.btnHideBottom.Add_Click({
        $mainGrid = $ui.WindowBorder.Child
        $ui.pnlBottomPanel.Visibility = 'Collapsed'
        if ($ui.splitterBottom) { $ui.splitterBottom.Visibility = 'Collapsed' }
        $mainGrid.RowDefinitions[3].Height = [System.Windows.GridLength]::new(0)
        $mainGrid.RowDefinitions[4].MinHeight = 0
        $mainGrid.RowDefinitions[4].Height = [System.Windows.GridLength]::new(0)
    })
}
if ($ui.btnClearLog) {
    $ui.btnClearLog.Add_Click({
        if ($ui.paraLog) { $ui.paraLog.Inlines.Clear() }
    })
}


# ═══════════════════════════════════════════════════════════════════════════════
# Intune Apps - Dismiss admin banner
if ($ui.BtnDismissAppBanner) {
    $ui.BtnDismissAppBanner.Add_Click({
        if ($ui.AppAdminBanner) { $ui.AppAdminBanner.Visibility = 'Collapsed' }
    }.GetNewClosure())
}

# ═══════════════════════════════════════════════════════════════════════════════
# Intune Apps - Reset app install state handler
if ($ui.BtnResetAppInstall) {
    $ui.BtnResetAppInstall.Add_Click({
        $sel = $ui.IntuneAppsGrid.SelectedItem
        if (-not $sel) {
            Show-Toast 'No App Selected' 'Select an app from the grid to reset its install state.' 'error'
            return
        }
        $appId   = $sel.AppId
        $appName = $sel.AppName
        $regKey  = $sel.RegistryKey
        Write-DebugLog "Reset app requested: $appName ($appId) key=$(if($regKey){$regKey}else{'<none - log-backfill app>'})" -Level DEBUG

        if (-not $regKey) {
            Write-DebugLog "Reset app skipped: $appName has no registry key (discovered from IME logs only)" -Level WARN
            Show-Toast 'No Registry Key' "This app was discovered from IME logs and has no registry tracking key to reset.`n`nIt may have already been removed or hasn't synced yet." 'warning'
            return
        }

        $confirm = Show-ThemedMessageBox -Message "Reset install state for:`n`n$appName`n($appId)`n`nThis removes the IME tracking registry key, causing Intune to re-evaluate and re-install the app on next sync.`n`nContinue?" -Title 'Reset App Install State' -Buttons 'YesNo' -Icon 'Warning'
        if ($confirm -ne 'Yes') {
            Write-DebugLog "Reset app cancelled by user: $appName" -Level DEBUG
            return
        }

        try {
            $regPath = "Registry::$regKey"
            if (-not (Test-Path $regPath)) {
                Show-Toast 'Key Not Found' "The registry key no longer exists:`n$regKey" 'error'
                return
            }
            Remove-Item $regPath -Recurse -Force -ErrorAction Stop
            Write-DebugLog "Reset app: $appName ($appId) - removed $regKey" -Level INFO

            # Also remove GRS (Global Retry Schedule) entry for this app to force full re-evaluation
            if ($regKey -match 'Win32Apps\\([^\\]+)\\') {
                $grsUserId = $Matches[1]
                $grsBase = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps\$grsUserId\GRS"
                if (Test-Path $grsBase) {
                    $grsKeys = Get-ChildItem $grsBase -ErrorAction SilentlyContinue
                    foreach ($gk in $grsKeys) {
                        $gkProps = $gk | Get-ItemProperty -ErrorAction SilentlyContinue
                        $gkNames = $gkProps.PSObject.Properties.Name | Where-Object { $_ -like '*-*-*-*-*' }
                        if ($gkNames -match [regex]::Escape($appId)) {
                            $gkPath = $gk.Name -replace 'HKEY_LOCAL_MACHINE', 'Registry::HKEY_LOCAL_MACHINE'
                            Remove-Item $gkPath -Recurse -Force -ErrorAction SilentlyContinue
                            Write-DebugLog "Reset app: removed GRS entry for $appId" -Level DEBUG
                        }
                    }
                }
            }

            Show-Toast 'App Reset' "$appName will be re-evaluated on next Intune sync." 'success'
            # Remove from UI list
            $toRemove = $Script:AllIntuneApps | Where-Object { $_.AppId -eq $appId -and $_.RegistryKey -eq $regKey } | Select-Object -First 1
            if ($toRemove) { $Script:AllIntuneApps.Remove($toRemove) }
        } catch {
            Write-DebugLog "Reset app failed: $($_.Exception.Message)" -Level ERROR
            Show-Toast 'Reset Failed' $_.Exception.Message 'error'
        }
    }.GetNewClosure())
}

# Intune Apps - Unified filter function (type + status + hide internal)
function Apply-IntuneAppsFilter {
    $items = $Script:AllIntuneApps
    # Hide internal packages
    if ($ui.ChkHideInternal -and $ui.ChkHideInternal.IsChecked) {
        $items = @($items | Where-Object {
            $_.AppName -notmatch 'InventoryAdaptorPackaging|IntuneWindowsAgent|EPMAgent|Microsoft EPM Agent'
        })
    }
    # Type filter
    if ($ui.CmbAppTypeFilter -and $ui.CmbAppTypeFilter.SelectedItem) {
        $typeSel = $ui.CmbAppTypeFilter.SelectedItem.Content
        if ($typeSel -and $typeSel -ne 'All Types') {
            if ($typeSel -eq 'Resettable') {
                $items = @($items | Where-Object { $_.RegistryKey })
            } else {
                $items = @($items | Where-Object { $_.AppType -eq $typeSel })
            }
        }
    }
    # Status filter
    if ($ui.CmbAppStatusFilter -and $ui.CmbAppStatusFilter.SelectedItem) {
        $statusSel = $ui.CmbAppStatusFilter.SelectedItem.Content
        if ($statusSel -and $statusSel -ne 'All Statuses') {
            $items = @($items | Where-Object { $_.InstallState -eq $statusSel })
        }
    }
    $ui.IntuneAppsGrid.ItemsSource = $items

    # Toggle empty state vs grid based on filtered results
    if ($items -and @($items).Count -gt 0) {
        if ($ui.IntuneAppsEmptyState) { $ui.IntuneAppsEmptyState.Visibility = 'Collapsed' }
        if ($ui.IntuneAppsGrid)       { $ui.IntuneAppsGrid.Visibility = 'Visible' }
    } else {
        if ($ui.IntuneAppsEmptyState) { $ui.IntuneAppsEmptyState.Visibility = 'Visible' }
        if ($ui.IntuneAppsGrid)       { $ui.IntuneAppsGrid.Visibility = 'Collapsed' }
    }
}

# Intune Apps - Filter by type (delegates to unified filter)
if ($ui.CmbAppTypeFilter) {
    $ui.CmbAppTypeFilter.Add_SelectionChanged({ Apply-IntuneAppsFilter }.GetNewClosure())
}

# Intune Apps - Filter by status (delegates to unified filter)
if ($ui.CmbAppStatusFilter) {
    $ui.CmbAppStatusFilter.Add_SelectionChanged({ Apply-IntuneAppsFilter }.GetNewClosure())
}

# Intune Apps - Hide internal packages toggle
if ($ui.ChkHideInternal) {
    $ui.ChkHideInternal.Add_Checked({ Apply-IntuneAppsFilter }.GetNewClosure())
    $ui.ChkHideInternal.Add_Unchecked({ Apply-IntuneAppsFilter }.GetNewClosure())
}


# ═══════════════════════════════════════════════════════════════════════════════
# Scan mode ComboBox handler
$ui.CmbScanMode.Add_SelectionChanged({
    $selected = $ui.CmbScanMode.SelectedItem
    if ($selected) {
        $mode = $selected.Tag
        $Script:Prefs.ScanMode = $mode
        $Script:PrereqsMet = $false  # Force re-check with new mode
        Save-Preferences
        if ($mode -eq 'Local') {
            $ui.BtnScanText.Text = 'Scan Local Policies'
            $ui.BtnScanGPOs.ToolTip = 'Scan policies applied to this machine via gpresult'
            if ($ui.BtnGetStartedText) { $ui.BtnGetStartedText.Text = 'Scan Local Policies' }
        } elseif ($mode -eq 'Intune') {
            $ui.BtnScanText.Text = 'Scan Intune Policies'
            $ui.BtnScanGPOs.ToolTip = 'Scan Intune/MDM policies from local registry'
            if ($ui.BtnGetStartedText) { $ui.BtnGetStartedText.Text = 'Scan Intune Policies' }
        } elseif ($mode -eq 'Combined') {
            $ui.BtnScanText.Text = 'Scan Co-managed Policies'
            $ui.BtnScanGPOs.ToolTip = 'Scan both GP policies (gpresult) and Intune/MDM policies'
            if ($ui.BtnGetStartedText) { $ui.BtnGetStartedText.Text = 'Scan Co-managed Policies' }
        } else {
            $ui.BtnScanText.Text = 'Scan Domain GPOs'
            $ui.BtnScanGPOs.ToolTip = 'Scan all GPOs from Active Directory using Get-GPO -All'
            if ($ui.BtnGetStartedText) { $ui.BtnGetStartedText.Text = 'Scan Domain GPOs' }
        }
    }
})

# Initialize scan mode from saved prefs
$modeIdx = switch ($Script:Prefs.ScanMode) { 'Intune' { 0 } 'Local' { 1 } 'AD' { 2 } 'Combined' { 3 } default { 0 } }
$ui.CmbScanMode.SelectedIndex = $modeIdx
# Explicitly sync button text (SelectionChanged won't fire if index matches XAML default)
$scanLabel = switch ($Script:Prefs.ScanMode) {
    'Local'    { 'Scan Local Policies' }
    'AD'       { 'Scan Domain GPOs' }
    'Combined' { 'Scan Co-managed Policies' }
    default    { 'Scan Intune Policies' }
}
$ui.BtnScanText.Text = $scanLabel
if ($ui.BtnGetStartedText) { $ui.BtnGetStartedText.Text = $scanLabel }

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 19: IME LOG VIEWER - LIVE TAIL + RICH PARSING + FORCE SYNC
# ═══════════════════════════════════════════════════════════════════════════════

# ── IME Log Parsing Engine (CMTrace/JSON hybrid format) ──
# Pre-compiled regexes for IME log classification (mirrors AIB LogMonitor pattern)

$Script:ImeRxError    = [regex]::new('(?i)(error|exception|fail(ed|ure)?|fatal|HRESULT\s*[:=]\s*0x8|StatusCode\s*[:=]\s*[45]\d{2}|could not|unable to|Access denied|not found.*critical|ExpectedPolicies)', 'Compiled')
$Script:ImeRxWarning  = [regex]::new('(?i)(warn(ing)?|timeout|retry|retrying|expired|fallback|skipp(ed|ing)|already exists|not applicable|timed out|throttl)', 'Compiled')
$Script:ImeRxSuccess  = [regex]::new('(?i)(successfully|completed|installed|applied|compliance.*(true|met)|detected|remediated|execution completed|exit code[:=]\s*0\b|Win32App.*successfully)', 'Compiled')
$Script:ImeRxPolicy   = [regex]::new('(?i)(policy|assignment|targeting|scope|applicability|evaluation|SideCar|StatusService|check-in|PolicyId)', 'Compiled')
$Script:ImeRxApp      = [regex]::new('(?i)(Win32App|IntuneApp|\.intunewin|ContentManager|download.*app|install.*app|detection rule|requirement rule|AppInstall|AppWorkload|Content.*download)', 'Compiled')
$Script:ImeRxScript   = [regex]::new('(?i)(HealthScript|Proactive.*remediation|Sensor|AgentExecutor|PowerShell|script.*execution|ScriptHandler|remediationScript|detectionScript)', 'Compiled')
$Script:ImeRxSync     = [regex]::new('(?i)(sync|check.?in|session|enrollment|OMA-?DM|SyncML|MDM.*session|DeviceEnroller|DMClient|polling|schedule)', 'Compiled')
$Script:ImeRxJson     = [regex]::new('^\s*[\[{]', 'Compiled')

# CMTrace log format: <![LOG[message]LOG]!><time="HH:mm:ss.fff" date="M-D-YYYY" component="X" context="" type="T" thread="N" file="F:L">
$Script:ImeRxCMTrace  = [regex]::new('<!\[LOG\[(?<msg>.*?)\]LOG\]!><time="(?<time>\d{1,2}:\d{2}:\d{2})\.\d+" date="(?<date>\d{1,2}-\d{1,2}-\d{4})" component="(?<comp>[^"]*)" context="[^"]*" type="(?<type>\d)" thread="(?<thread>\d+)" file="(?<file>[^"]*)">', 'Compiled,Singleline')
# Simpler fallback for plain-text IME log lines with timestamp
$Script:ImeRxPlainTS  = [regex]::new('^<!\[LOG\[(?<msg>.*?)\]LOG\]!>', 'Compiled,Singleline')

function Get-ImeLineColor {
    param([string]$Line, [string]$CmTraceType)
    # CMTrace type field: 1=Info, 2=Warning, 3=Error
    if ($CmTraceType -eq '3') { return @{ Color = '#D13438'; Bold = $true;  Cat = 'Error'   } }
    if ($CmTraceType -eq '2') { return @{ Color = '#FFB900'; Bold = $false; Cat = 'Warning' } }
    # Content-based classification
    if ($Script:ImeRxError.IsMatch($Line))   { return @{ Color = '#D13438'; Bold = $true;  Cat = 'Error'   } }
    if ($Script:ImeRxWarning.IsMatch($Line)) { return @{ Color = '#FFB900'; Bold = $false; Cat = 'Warning' } }
    if ($Script:ImeRxSuccess.IsMatch($Line)) { return @{ Color = '#107C10'; Bold = $true;  Cat = 'Success' } }
    if ($Script:ImeRxApp.IsMatch($Line))     { return @{ Color = '#8764B8'; Bold = $false; Cat = 'App'     } }
    if ($Script:ImeRxScript.IsMatch($Line))  { return @{ Color = '#FF8C00'; Bold = $false; Cat = 'Script'  } }
    if ($Script:ImeRxSync.IsMatch($Line))    { return @{ Color = '#60CDFF'; Bold = $false; Cat = 'Sync'    } }
    if ($Script:ImeRxPolicy.IsMatch($Line))  { return @{ Color = '#0078D4'; Bold = $false; Cat = 'Policy'  } }
    return @{ Color = '#C0C0C0'; Bold = $false; Cat = 'Info' }
}

function Get-ImeSeverityBadge {
    param([string]$Cat)
    switch ($Cat) {
        'Error'   { return @{ Label = 'ERR';  Bg = '#D13438'; Fg = '#FFFFFF' } }
        'Warning' { return @{ Label = 'WARN'; Bg = '#FFB900'; Fg = '#1A1A1A' } }
        'Success' { return @{ Label = 'OK';   Bg = '#107C10'; Fg = '#FFFFFF' } }
        'App'     { return @{ Label = 'APP';  Bg = '#8764B8'; Fg = '#FFFFFF' } }
        'Script'  { return @{ Label = 'SCRP'; Bg = '#FF8C00'; Fg = '#FFFFFF' } }
        'Sync'    { return @{ Label = 'SYNC'; Bg = '#60CDFF'; Fg = '#1A1A1A' } }
        'Policy'  { return @{ Label = 'POL';  Bg = '#0078D4'; Fg = '#FFFFFF' } }
        default   { return @{ Label = 'INFO'; Bg = '#333333'; Fg = '#AAAAAA' } }
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# SHARED: Virtualized ListBox helpers for log viewers
# ═══════════════════════════════════════════════════════════════════════════════

function Get-ListBoxScrollViewer([System.Windows.Controls.ListBox]$lb) {
    if (-not $lb -or [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($lb) -eq 0) { return $null }
    $border = [System.Windows.Media.VisualTreeHelper]::GetChild($lb, 0)
    if ($border -is [System.Windows.Controls.Border]) {
        for ($i = 0; $i -lt [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($border); $i++) {
            $child = [System.Windows.Media.VisualTreeHelper]::GetChild($border, $i)
            if ($child -is [System.Windows.Controls.ScrollViewer]) { return $child }
        }
    }
    return $null
}

function Resolve-Brush([string]$Hex, [hashtable]$Cache) {
    if (-not $Cache.ContainsKey($Hex)) {
        $b = [System.Windows.Media.BrushConverter]::new().ConvertFromString($Hex); $b.Freeze(); $Cache[$Hex] = $b
    }
    $Cache[$Hex]
}

# Badge foreground: dark text for bright backgrounds, white for dark backgrounds
$Script:BadgeFgLookup = @{
    '#FFB900' = '#1B1B1B'; '#FF8C00' = '#1B1B1B'; '#60CDFF' = '#1B1B1B'; '#C0C0C0' = '#1B1B1B'
}
$Script:BadgeInfoBg = '#444444'; $Script:BadgeInfoFg = '#888888'

function Apply-LogListBoxColors {
    param(
        [System.Windows.Controls.ListBox]$ListBox,
        [System.Collections.Generic.List[byte]]$LineTypes,
        [array]$ColorMap,
        [hashtable]$BrushCache
    )
    if (-not $ListBox -or -not $LineTypes -or $ListBox.Items.Count -eq 0) { return }
    $gen = $ListBox.ItemContainerGenerator
    if (-not $gen) { return }
    $sv = Get-ListBoxScrollViewer $ListBox
    if (-not $sv) { return }
    $firstVisible = [int][Math]::Floor($sv.VerticalOffset)
    $lastVisible  = [int][Math]::Min($firstVisible + [int]$sv.ViewportHeight + 2, $ListBox.Items.Count - 1)
    if ($firstVisible -lt 0) { $firstVisible = 0 }
    $monoFont = [System.Windows.Media.FontFamily]::new('Cascadia Mono, Consolas, Courier New')
    for ($i = $firstVisible; $i -le $lastVisible; $i++) {
        $container = $gen.ContainerFromIndex($i)
        if ($container -is [System.Windows.Controls.ListBoxItem]) {
            $lt = if ($i -lt $LineTypes.Count) { $LineTypes[$i] } else { 0 }
            $cm = $ColorMap[$lt]
            $text = [string]$ListBox.Items[$i]

            $tb = [System.Windows.Controls.TextBlock]::new()
            $tb.TextWrapping = 'NoWrap'
            $tb.TextTrimming = 'CharacterEllipsis'
            $tb.FontFamily = $monoFont; $tb.FontSize = 12

            # Parse badge from "[XXX] rest..."
            $cb = $text.IndexOf('] ')
            if ($cb -gt 0 -and $text[0] -eq '[') {
                $badgeLabel = $text.Substring(1, $cb - 1)
                $restText   = $text.Substring($cb + 2)

                # Badge pill colors
                $bgHex = if ($lt -eq 0) { $Script:BadgeInfoBg } else { $cm.Fg }
                $fgHex = if ($lt -eq 0) { $Script:BadgeInfoFg }
                         elseif ($Script:BadgeFgLookup.ContainsKey($cm.Fg)) { $Script:BadgeFgLookup[$cm.Fg] }
                         else { '#FFFFFF' }

                $bRun = [System.Windows.Documents.Run]::new(" $badgeLabel ")
                $bRun.Background = Resolve-Brush $bgHex $BrushCache
                $bRun.Foreground = Resolve-Brush $fgHex $BrushCache
                $bRun.FontSize = 9; $bRun.FontWeight = 'Bold'; $bRun.FontFamily = $monoFont
                [void]$tb.Inlines.Add($bRun)

                [void]$tb.Inlines.Add([System.Windows.Documents.Run]::new(' '))

                $tRun = [System.Windows.Documents.Run]::new($restText)
                $tRun.Foreground = Resolve-Brush $cm.Fg $BrushCache
                if ($cm.Bold) { $tRun.FontWeight = 'Bold' }
                [void]$tb.Inlines.Add($tRun)
            } else {
                $tRun = [System.Windows.Documents.Run]::new($text)
                $tRun.Foreground = Resolve-Brush $cm.Fg $BrushCache
                if ($cm.Bold) { $tRun.FontWeight = 'Bold' }
                [void]$tb.Inlines.Add($tRun)
            }

            $container.Content = $tb
            $container.Background = [System.Windows.Media.Brushes]::Transparent
            $container.Padding = [System.Windows.Thickness]::new(0)
            $container.Margin  = [System.Windows.Thickness]::new(0)
        }
    }
}

function Rebuild-LogMinimap {
    param([System.Windows.Controls.Canvas]$Canvas, [System.Collections.Generic.List[byte]]$LineTypes, [array]$ColorMap, [hashtable]$BrushCache, [int]$MaxSamples = 2000)
    if (-not $Canvas) { return }
    $Canvas.Children.Clear()
    $total = $LineTypes.Count; if ($total -eq 0) { return }
    $ch = $Canvas.ActualHeight; if ($ch -lt 10) { return }
    $step = [Math]::Max(1, [int]($total / $MaxSamples))
    for ($i = 0; $i -lt $total; $i += $step) {
        $lt = $LineTypes[$i]; if ($lt -eq 0) { continue }
        $hex = $ColorMap[$lt].Fg
        if (-not $BrushCache.ContainsKey($hex)) { $b = [System.Windows.Media.BrushConverter]::new().ConvertFromString($hex); $b.Freeze(); $BrushCache[$hex] = $b }
        $y = ($i / $total) * $ch
        $dot = [System.Windows.Shapes.Ellipse]::new(); $dot.Width = 6; $dot.Height = 3; $dot.Fill = $BrushCache[$hex]
        [System.Windows.Controls.Canvas]::SetTop($dot, $y); [System.Windows.Controls.Canvas]::SetLeft($dot, 2)
        [void]$Canvas.Children.Add($dot)
    }
}

function Rebuild-LogHeatmap {
    param([System.Windows.Controls.Canvas]$Canvas, [System.Collections.Generic.List[byte]]$LineTypes, [array]$ColorMap, [hashtable]$BrushCache, [int]$MaxSamples = 2000)
    if (-not $Canvas) { return }
    $Canvas.Children.Clear()
    $total = $LineTypes.Count; if ($total -eq 0) { return }
    $cw = $Canvas.ActualWidth; if ($cw -lt 10) { return }
    $step = [Math]::Max(1, [int]($total / $MaxSamples))
    $barW = [Math]::Max(2, $cw / [Math]::Max($total / $step, 200))
    for ($i = 0; $i -lt $total; $i += $step) {
        $lt = $LineTypes[$i]; if ($lt -eq 0) { continue }
        $hex = $ColorMap[$lt].Fg
        if (-not $BrushCache.ContainsKey($hex)) { $b = [System.Windows.Media.BrushConverter]::new().ConvertFromString($hex); $b.Freeze(); $BrushCache[$hex] = $b }
        $x = ($i / $total) * $cw
        $rect = [System.Windows.Shapes.Rectangle]::new(); $rect.Width = $barW; $rect.Height = 14; $rect.Fill = $BrushCache[$hex]; $rect.Opacity = 0.7
        [System.Windows.Controls.Canvas]::SetLeft($rect, $x)
        [void]$Canvas.Children.Add($rect)
    }
}


# Viewport indicator on minimap (shows current scroll position)
function Update-MinimapViewportIndicator([System.Windows.Controls.Canvas]$Canvas, [System.Windows.Controls.ScrollViewer]$SV, [int]$TotalLines) {
    if (-not $Canvas -or -not $SV -or $TotalLines -le 0) { return }
    $H = $Canvas.ActualHeight; if ($H -le 0) { return }
    $topRatio    = $SV.VerticalOffset / $TotalLines
    $heightRatio = $SV.ViewportHeight / $TotalLines
    $vpTop = $topRatio * $H; $vpH = [Math]::Max($heightRatio * $H, 3)
    $existing = $null
    foreach ($c in $Canvas.Children) { if ($c.Tag -eq 'viewport') { $existing = $c; break } }
    if (-not $existing) {
        $existing = [System.Windows.Shapes.Rectangle]::new()
        $existing.Tag = 'viewport'; $existing.Width = 12; $existing.Opacity = 0.3
        $b = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Colors]::White); $b.Freeze()
        $existing.Fill = $b; $existing.IsHitTestVisible = $false
        [void]$Canvas.Children.Add($existing)
    }
    $existing.Height = $vpH
    [System.Windows.Controls.Canvas]::SetTop($existing, $vpTop)
    [System.Windows.Controls.Canvas]::SetLeft($existing, 0)
}

# Update sidebar density bar + findings list for a log viewer
function Update-LogSidebarExtras {
    param(
        [System.Collections.Generic.List[byte]]$LineTypes,
        [System.Windows.Controls.Border]$DensityError,
        [System.Windows.Controls.Border]$DensityWarn,
        [System.Windows.Controls.ListBox]$FindingsList,
        [System.Collections.Generic.List[string]]$FormattedLines,
        [System.Windows.Controls.ListBox]$LogListBox,
        [array]$ColorMap
    )
    if (-not $LineTypes -or $LineTypes.Count -eq 0) { return }
    $total = $LineTypes.Count
    $errCount = 0; $warnCount = 0
    for ($i = 0; $i -lt $total; $i++) {
        $lt = $LineTypes[$i]
        if ($lt -eq 1) { $errCount++ }
        elseif ($lt -eq 2) { $warnCount++ }
    }
    if ($DensityError) { $DensityError.Width = [Math]::Max(1, ($errCount / $total) * 200) }
    if ($DensityWarn) { $DensityWarn.Width = [Math]::Max(1, ($warnCount / $total) * 200) }
    if ($FindingsList -and $FormattedLines) {
        $FindingsList.Items.Clear()
        $brConv = [System.Windows.Media.BrushConverter]::new()
        $count = 0
        for ($i = 0; $i -lt $total -and $count -lt 25; $i++) {
            $lt = $LineTypes[$i]; if ($lt -ne 1 -and $lt -ne 2) { continue }
            $color = $ColorMap[$lt].Fg
            $sp = [System.Windows.Controls.StackPanel]::new(); $sp.Orientation = 'Horizontal'; $sp.Margin = [System.Windows.Thickness]::new(0,1,0,1)
            $dot = [System.Windows.Shapes.Ellipse]::new(); $dot.Width = 5; $dot.Height = 5; $dot.VerticalAlignment = 'Center'; $dot.Margin = [System.Windows.Thickness]::new(0,0,5,0)
            $db = $brConv.ConvertFromString($color); $db.Freeze(); $dot.Fill = $db
            [void]$sp.Children.Add($dot)
            $lnTb = [System.Windows.Controls.TextBlock]::new(); $lnTb.Text = "L$($i+1)"; $lnTb.FontSize = 9; $lnTb.Width = 42
            $lnFg = $brConv.ConvertFromString('#FF888892'); $lnFg.Freeze(); $lnTb.Foreground = $lnFg; $lnTb.VerticalAlignment = 'Center'
            [void]$sp.Children.Add($lnTb)
            $rawLine = if ($i -lt $FormattedLines.Count) { $FormattedLines[$i] } else { '' }
            if ($rawLine.Length -gt 8) { $rawLine = $rawLine.Substring(8) }
            if ($rawLine.Length -gt 50) { $rawLine = $rawLine.Substring(0,50) + '...' }
            $descTb = [System.Windows.Controls.TextBlock]::new(); $descTb.Text = $rawLine; $descTb.FontSize = 9; $descTb.TextTrimming = 'CharacterEllipsis'
            $dfg = $brConv.ConvertFromString($color); $dfg.Freeze(); $descTb.Foreground = $dfg; $descTb.VerticalAlignment = 'Center'
            [void]$sp.Children.Add($descTb)
            $sp.Tag = $i; $sp.Cursor = 'Hand'
            $sp.Add_MouseLeftButtonDown({
                param($sender,$e)
                $idx = $sender.Tag
                if ($idx -ge 0 -and $idx -lt $LogListBox.Items.Count) {
                    $LogListBox.ScrollIntoView($LogListBox.Items[$idx])
                    $LogListBox.SelectedIndex = $idx
                }
            }.GetNewClosure())
            [void]$FindingsList.Items.Add($sp)
            $count++
        }
    }
}
# â”€â”€ IME Log State â”€â”€
$Script:ImeLogQueue      = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new())
$Script:ImeTailing       = $false
$Script:ImeWatcher       = $null
$Script:ImeMaxPerTick     = 150
$Script:ImeLastFilePos    = 0
$Script:ImeActiveLogPath  = ''
$Script:ImeActiveFilter   = 'All'
$Script:ImeContentFilter   = ''
$Script:ImeSavedFilters    = @{}
$Script:ImeLastTimestamp   = $null
$Script:ImeStats = @{ Lines = 0; Errors = 0; Warnings = 0 }
$Script:ImeSearchMatches   = [System.Collections.Generic.List[int]]::new()
$Script:ImeSearchIndex     = -1
$Script:ImeBrushCache      = @{}
$Script:ImeLineLimit       = 5000
$Script:ImeLineLimitOptions = @(1000, 5000, 10000, 0)
$Script:ImeFormattedLines  = [System.Collections.Generic.List[string]]::new()
$Script:ImeLineTypes       = [System.Collections.Generic.List[byte]]::new()
$Script:ImeRawEntries      = [System.Collections.Generic.List[string]]::new()
$Script:ImeColorMap = @( @{ Fg = '#C0C0C0'; Bold = $false }, @{ Fg = '#D13438'; Bold = $true }, @{ Fg = '#FFB900'; Bold = $false }, @{ Fg = '#107C10'; Bold = $true }, @{ Fg = '#8764B8'; Bold = $false }, @{ Fg = '#FF8C00'; Bold = $false }, @{ Fg = '#60CDFF'; Bold = $false }, @{ Fg = '#0078D4'; Bold = $false } )
$Script:ImeBadgeMap = @('INFO','ERR','WARN','OK','APP','SCRP','SYNC','POL')
$Script:ImeLogDir = 'C:\ProgramData\Microsoft\IntuneManagementExtension\Logs'
function Update-ImeStatsDisplay {
    if ($ui.ImeStatLines)    { $ui.ImeStatLines.Text    = $Script:ImeStats.Lines }
    if ($ui.ImeStatErrors)   { $ui.ImeStatErrors.Text   = $Script:ImeStats.Errors }
    if ($ui.ImeStatWarnings) { $ui.ImeStatWarnings.Text = $Script:ImeStats.Warnings }
}
function Clear-ImeLogDisplay {
    $Script:ImeFormattedLines.Clear(); $Script:ImeLineTypes.Clear(); $Script:ImeRawEntries.Clear()
    $ui.lbImeLogs.ItemsSource = $null
    $Script:ImeStats.Lines = 0; $Script:ImeStats.Errors = 0; $Script:ImeStats.Warnings = 0
    Update-ImeStatsDisplay
    if ($ui.cnvImeHeatmap) { $ui.cnvImeHeatmap.Children.Clear() }; if ($ui.cnvImeMinimap) { $ui.cnvImeMinimap.Children.Clear() }
    $Script:ImeSearchMatches.Clear(); $Script:ImeSearchIndex = -1
    if ($ui.ImeSearchCount) { $ui.ImeSearchCount.Text = '' }
}
function Parse-ImeLogLine {
    param([string]$RawLine)
    # Try CMTrace format first
    $m = $Script:ImeRxCMTrace.Match($RawLine)
    if ($m.Success) {
        return @{
            Message   = $m.Groups['msg'].Value
            Time      = $m.Groups['time'].Value
            Date      = $m.Groups['date'].Value
            Component = $m.Groups['comp'].Value
            Type      = $m.Groups['type'].Value
            Thread    = $m.Groups['thread'].Value
            File      = $m.Groups['file'].Value
        }
    }
    # Fallback: simpler CMTrace
    $m2 = $Script:ImeRxPlainTS.Match($RawLine)
    if ($m2.Success) {
        return @{ Message = $m2.Groups['msg'].Value; Time = ''; Date = ''; Component = ''; Type = '1'; Thread = ''; File = '' }
    }
    # Pure text line
    return @{ Message = $RawLine; Time = ''; Date = ''; Component = ''; Type = '1'; Thread = ''; File = '' }
}

function Classify-ImeLogEntry {
    param([hashtable]$Parsed, [string]$RawLine)
    $msg  = $Parsed.Message
    $rules = Get-ImeLineColor -Line $msg -CmTraceType $Parsed.Type
    if ($Script:ImeActiveFilter -ne 'All') {
        $cat = $rules.Cat
        $pass = switch ($Script:ImeActiveFilter) { 'Error' { $cat -eq 'Error' }; 'Warning' { $cat -eq 'Warning' }; 'Info' { $cat -notin @('Error','Warning') }; default { $true } }
        if (-not $pass) { return $null }
    }
    if ($Script:ImeContentFilter -and $Script:ImeContentFilter.Length -gt 0) {
        $pattern = $Script:ImeContentFilter
        $textToSearch = "$msg $($Parsed.Component) $RawLine"
        if ($pattern.StartsWith('/') -and $pattern.EndsWith('/')) {
            try { if ($textToSearch -notmatch $pattern.Substring(1, $pattern.Length - 2)) { return $null } } catch { if ($textToSearch -notlike "*$pattern*") { return $null } }
        } else { if ($textToSearch -notlike "*$pattern*") { return $null } }
    }
    $Script:ImeStats.Lines++; if ($rules.Cat -eq 'Error') { $Script:ImeStats.Errors++ }; if ($rules.Cat -eq 'Warning') { $Script:ImeStats.Warnings++ }
    $lineType = switch ($rules.Cat) { 'Error' { [byte]1 }; 'Warning' { [byte]2 }; 'Success' { [byte]3 }; 'App' { [byte]4 }; 'Script' { [byte]5 }; 'Sync' { [byte]6 }; 'Policy' { [byte]7 }; default { [byte]0 } }
    $badge = $Script:ImeBadgeMap[$lineType]
    $sb = [System.Text.StringBuilder]::new(256)
    [void]$sb.Append("[$badge] ")
    if ($Parsed.Time) { if ($Parsed.Date) { [void]$sb.Append("$($Parsed.Date) ") }; [void]$sb.Append("$($Parsed.Time) ") }
    if ($Parsed.Component) { [void]$sb.Append("[$($Parsed.Component)] ") }
    [void]$sb.Append($msg)
    if ($Parsed.Thread) { [void]$sb.Append("  T:$($Parsed.Thread)") }
    return @{ Formatted = $sb.ToString(); LineType = $lineType; Raw = $RawLine }
}

# ── File Tailing Engine ──
function Read-ImeLogDelta {
    if (-not $Script:ImeActiveLogPath -or -not (Test-Path $Script:ImeActiveLogPath)) {
        Write-Host "[IME] Read-ImeLogDelta skipped - path='$($Script:ImeActiveLogPath)' exists=$((Test-Path $Script:ImeActiveLogPath -ErrorAction SilentlyContinue))"
        return
    }
    try {
        $fs = [System.IO.FileStream]::new($Script:ImeActiveLogPath, 'Open', 'Read', 'ReadWrite,Delete')
        $fs.Seek($Script:ImeLastFilePos, 'Begin')
        $sr = [System.IO.StreamReader]::new($fs)
        $linesRead = 0
        while ($null -ne ($line = $sr.ReadLine())) {
            if ($line.Trim()) {
                $Script:ImeLogQueue.Enqueue($line)
                $linesRead++
            }
        }
        $Script:ImeLastFilePos = $fs.Position
        $sr.Close()
        $fs.Close()
        if ($linesRead -gt 0) { Write-Verbose "[IME] Read-ImeLogDelta: $linesRead lines queued, pos=$($Script:ImeLastFilePos), queueSize=$($Script:ImeLogQueue.Count)" }
    } catch {
        Write-Host "[IME] Read-ImeLogDelta ERROR: $_"
        Write-DebugLog "IME tail read error: $_" -Level WARN
    }
}

function Start-ImeTail {
    $source = $ui.CmbImeLogSource.SelectedItem.Content
    Write-Host "[IME] Start-ImeTail called - source='$source', dir='$Script:ImeLogDir'"
    $logFile = Join-Path $Script:ImeLogDir "$source.log"
    if (-not (Test-Path $logFile)) {
        Write-Host "[IME] Log file NOT FOUND: $logFile"
        Show-Toast 'Log Not Found' "Cannot find $logFile" -Type Warning
        return
    }
    Write-Host "[IME] Log file exists: $logFile (size=$([System.IO.FileInfo]::new($logFile).Length))"
    $Script:ImeActiveLogPath = $logFile
    $Script:ImeLastFilePos = 0
    $Script:ImeTailing = $true
    $Script:ImeLastTimestamp = $null

    Clear-ImeLogDisplay
    $ui.ImeLogSubtitle.Text = "Tailing: $source.log"

    # Initial backfill - read existing file
    Read-ImeLogDelta

    # Set up FileSystemWatcher for live updates
    if ($Script:ImeWatcher) {
        $Script:ImeWatcher.EnableRaisingEvents = $false
        $Script:ImeWatcher.Dispose()
    }
    $Script:ImeWatcher = [System.IO.FileSystemWatcher]::new()
    $Script:ImeWatcher.Path = $Script:ImeLogDir
    $Script:ImeWatcher.Filter = "$source.log"
    $Script:ImeWatcher.NotifyFilter = [System.IO.NotifyFilters]::LastWrite -bor [System.IO.NotifyFilters]::Size
    $Script:ImeWatcher.EnableRaisingEvents = $true

    # Use SynchronizingObject to avoid cross-thread issues - instead use a flag
    $Script:ImeFswChanged = $false
    Register-ObjectEvent -InputObject $Script:ImeWatcher -EventName Changed -Action {
        $Script:ImeFswChanged = $true
    } -SourceIdentifier 'ImeLogWatcher'

    # UI state (buttons removed - tailing is automatic)

    Write-Host "[IME] Tail started - FSW watching '$source.log', timer running, ImeTailing=$Script:ImeTailing"
    Write-DebugLog "IME tail started: $logFile" -Level STEP
    Show-Toast 'IME Log Tail' "Now tailing $source.log" -Type Info
}

function Stop-ImeTail {
    $Script:ImeTailing = $false
    if ($Script:ImeWatcher) {
        $Script:ImeWatcher.EnableRaisingEvents = $false
        $Script:ImeWatcher.Dispose()
        $Script:ImeWatcher = $null
    }
    Unregister-Event -SourceIdentifier 'ImeLogWatcher' -ErrorAction SilentlyContinue
    $ui.ImeLogSubtitle.Text = 'Intune Management Extension - Tail stopped'
    if ($ui.ImeFollowIndicator) { $ui.ImeFollowIndicator.Text = [char]0x25A0 + ' Stopped' }
    Write-DebugLog 'IME tail stopped' -Level STEP
}

# ── Load log file for analysis (tail-read + batched render) ──
function Load-ImeLogFile {
    $source = $ui.CmbImeLogSource.SelectedItem.Content
    $logFile = Join-Path $Script:ImeLogDir "$source.log"
    Write-Host "[IME] Load-ImeLogFile: source='$source', path='$logFile'"
    if (-not (Test-Path $logFile)) { Write-Host "[IME] Log file NOT FOUND: $logFile"; Show-Toast 'Log Not Found' "Cannot find $logFile" -Type Warning; return }
    if ($Script:ImeTailing) { Stop-ImeTail }
    Clear-ImeLogDisplay
    $ui.ImeLogSubtitle.Text = "Loading: $source.log..."
    $lineLimit = $Script:ImeLineLimit
    $skipped = $false
    if ($Script:HasCSharpParser) {
        # ── C# fast path: file read + CMTrace parse + classify in one native call ──
        try {
            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            [PolicyLogParser]::ParseAndClassifyImeFile($logFile, $lineLimit, $Script:ImeActiveFilter, $Script:ImeContentFilter)
            $Script:ImeFormattedLines.AddRange([PolicyLogParser]::FormattedLines)
            $Script:ImeLineTypes.AddRange([PolicyLogParser]::LineTypes)
            $Script:ImeRawEntries.AddRange([PolicyLogParser]::RawEntries)
            $Script:ImeStats.Lines    = [PolicyLogParser]::TotalLines
            $Script:ImeStats.Errors   = [PolicyLogParser]::ErrorCount
            $Script:ImeStats.Warnings = [PolicyLogParser]::WarningCount
            $skipped = ($lineLimit -gt 0)
            $sw.Stop()
            Write-Host "[IME] C# ParseAndClassifyImeFile: $([PolicyLogParser]::TotalLines) lines in $($sw.ElapsedMilliseconds)ms"
        } catch {
            Write-Host "[IME] C# path failed, falling back to PS: $_"
            $Script:HasCSharpParser = $false
            Clear-ImeLogDisplay
        }
    }
    if (-not $Script:HasCSharpParser) {
        # ── PowerShell fallback ──
        try {
            $fs = [System.IO.FileStream]::new($logFile, 'Open', 'Read', 'ReadWrite,Delete')
            $sr = [System.IO.StreamReader]::new($fs)
            if ($lineLimit -gt 0 -and $fs.Length -gt 0) { $seekBack = [long][Math]::Min($fs.Length, [long]$lineLimit * 240); if ($seekBack -lt $fs.Length) { $fs.Seek(-$seekBack, 'End'); [void]$sr.ReadLine(); $skipped = $true } }
            $allText = $sr.ReadToEnd(); $sr.Close(); $fs.Close()
            $allLines = $allText -split "\r?\n"
            if ($skipped -and $lineLimit -gt 0 -and $allLines.Count -gt $lineLimit) { $allLines = $allLines[($allLines.Count - $lineLimit)..($allLines.Count - 1)] }
        } catch { Write-Host "[IME] Load-ImeLogFile read error: $_"; $ui.ImeLogSubtitle.Text = "Error reading $source.log"; return }
        $total = $allLines.Count
        Write-Host "[IME] Load-ImeLogFile: $total lines to classify (limit=$lineLimit, skipped=$skipped) [PS fallback]"
        for ($i = 0; $i -lt $total; $i++) {
            $line = $allLines[$i]; if (-not $line.Trim()) { continue }
            $parsed = Parse-ImeLogLine $line
            $result = Classify-ImeLogEntry $parsed $line
            if ($result) { $Script:ImeFormattedLines.Add($result.Formatted); $Script:ImeLineTypes.Add($result.LineType); $Script:ImeRawEntries.Add($result.Raw) }
        }
    }
    $ui.lbImeLogs.ItemsSource = $Script:ImeFormattedLines
    $ui.lbImeLogs.UpdateLayout()
    $imeSv = Get-ListBoxScrollViewer $ui.lbImeLogs
    if ($imeSv) { $imeSv.Add_ScrollChanged({ Apply-LogListBoxColors $ui.lbImeLogs $Script:ImeLineTypes $Script:ImeColorMap $Script:ImeBrushCache; $sv = Get-ListBoxScrollViewer $ui.lbImeLogs; Update-MinimapViewportIndicator $ui.cnvImeMinimap $sv $ui.lbImeLogs.Items.Count }) }
    Apply-LogListBoxColors $ui.lbImeLogs $Script:ImeLineTypes $Script:ImeColorMap $Script:ImeBrushCache
    if ($ui.cnvImeHeatmap) { $ui.cnvImeHeatmap.Visibility = 'Visible'; $ui.cnvImeHeatmap.UpdateLayout() }
    Rebuild-LogHeatmap $ui.cnvImeHeatmap $Script:ImeLineTypes $Script:ImeColorMap $Script:ImeBrushCache
    Rebuild-LogMinimap $ui.cnvImeMinimap $Script:ImeLineTypes $Script:ImeColorMap $Script:ImeBrushCache
    Update-LogSidebarExtras -LineTypes $Script:ImeLineTypes -DensityError $ui.ImeDensityError -DensityWarn $ui.ImeDensityWarn -FindingsList $ui.ImeFindingsList -FormattedLines $Script:ImeFormattedLines -LogListBox $ui.lbImeLogs -ColorMap $Script:ImeColorMap
    $rendered = $Script:ImeFormattedLines.Count
    Update-ImeStatsDisplay
    $limitTag = if ($skipped) { " (last $lineLimit)" } else { '' }
    $ui.ImeLogSubtitle.Text = "Loaded: $source.log ($rendered lines$limitTag)"
    if ($ui.ImeFollowIndicator) { $ui.ImeFollowIndicator.Text = [char]0x25A0 + " Analysis" }
    if ($rendered -gt 0) { $ui.lbImeLogs.ScrollIntoView($ui.lbImeLogs.Items[$rendered - 1]) }
    Write-Host "[IME] Load-ImeLogFile complete: $rendered lines classified"
    Show-Toast 'Log Loaded' "$rendered lines from $source.log$limitTag" -Type Info
}

# â”€â”€ IME DispatcherTimer â”€â”€
$Script:ImeTimer = [System.Windows.Threading.DispatcherTimer]::new()
$Script:ImeTimer.Interval = [TimeSpan]::FromMilliseconds(80)
$Script:ImeStatsThrottle = [DateTime]::MinValue
$Script:ImePollThrottle  = [DateTime]::MinValue
$Script:ImeTimer.Add_Tick({
    if (-not $Script:ImeTailing) { return }
    $needsRead = $Script:ImeFswChanged
    if (-not $needsRead) { $now = [DateTime]::Now; if (($now - $Script:ImePollThrottle).TotalMilliseconds -gt 500) { $Script:ImePollThrottle = $now; if ($Script:ImeActiveLogPath -and (Test-Path $Script:ImeActiveLogPath)) { $len = ([System.IO.FileInfo]::new($Script:ImeActiveLogPath)).Length; if ($len -gt $Script:ImeLastFilePos) { $needsRead = $true } } } }
    if ($needsRead) { $Script:ImeFswChanged = $false; Read-ImeLogDelta }
    $batchCount = 0; $contentAdded = $false
    # Drain queue into array for C# batch processing
    $batchLines = [System.Collections.Generic.List[string]]::new()
    while ($Script:ImeLogQueue.Count -gt 0 -and $batchCount -lt $Script:ImeMaxPerTick) {
        $batchLines.Add($Script:ImeLogQueue.Dequeue()); $batchCount++
    }
    if ($batchLines.Count -gt 0) {
        if ($Script:HasCSharpParser) {
            [PolicyLogParser]::ClassifyImeLines([string[]]$batchLines.ToArray(), $Script:ImeActiveFilter, $Script:ImeContentFilter)
            if ([PolicyLogParser]::TotalLines -gt 0) {
                $Script:ImeFormattedLines.AddRange([PolicyLogParser]::FormattedLines)
                $Script:ImeLineTypes.AddRange([PolicyLogParser]::LineTypes)
                $Script:ImeRawEntries.AddRange([PolicyLogParser]::RawEntries)
                $Script:ImeStats.Lines    += [PolicyLogParser]::TotalLines
                $Script:ImeStats.Errors   += [PolicyLogParser]::ErrorCount
                $Script:ImeStats.Warnings += [PolicyLogParser]::WarningCount
                $contentAdded = $true
            }
        } else {
            foreach ($raw in $batchLines) {
                $parsed = Parse-ImeLogLine $raw
                $result = Classify-ImeLogEntry $parsed $raw
                if ($result) { $Script:ImeFormattedLines.Add($result.Formatted); $Script:ImeLineTypes.Add($result.LineType); $Script:ImeRawEntries.Add($result.Raw); $contentAdded = $true }
            }
        }
    }
    if ($contentAdded) {
        $ui.lbImeLogs.ItemsSource = $null; $ui.lbImeLogs.ItemsSource = $Script:ImeFormattedLines
        $sv = Get-ListBoxScrollViewer $ui.lbImeLogs
        if ($sv) {
            $isLatched = ($sv.VerticalOffset + $sv.ViewportHeight) -ge ($sv.ExtentHeight - 20.0)
            if ($isLatched -and $Script:ImeFormattedLines.Count -gt 0) { $ui.lbImeLogs.ScrollIntoView($ui.lbImeLogs.Items[$Script:ImeFormattedLines.Count - 1]); if ($ui.ImeFollowIndicator) { $ui.ImeFollowIndicator.Text = [char]0x25BC + ' Following' } }
            else { if ($ui.ImeFollowIndicator) { $ui.ImeFollowIndicator.Text = [char]0x25A0 + ' Paused' } }
        }
        $ui.lbImeLogs.UpdateLayout(); Apply-LogListBoxColors $ui.lbImeLogs $Script:ImeLineTypes $Script:ImeColorMap $Script:ImeBrushCache
        $now2 = [DateTime]::Now; if (($now2 - $Script:ImeStatsThrottle).TotalMilliseconds -gt 500) { Update-ImeStatsDisplay; $Script:ImeStatsThrottle = $now2 }
    }
})
$Script:ImeTimer.Start()

# ── Intune Sync Trigger ──
function Invoke-IntuneSync {
    $ui.BtnIntuneSync.IsEnabled = $false
    $ui.ImeSyncStatus.Text = 'Triggering sync...'
    Write-Host "[IME] Invoke-IntuneSync called - ImeTailing=$Script:ImeTailing"
    Write-DebugLog 'Triggering Intune device sync' -Level STEP

    # Auto-start IME log tail if not already tailing
    if (-not $Script:ImeTailing) {
        Write-Host '[IME] Auto-starting tail for sync monitoring'
        Start-ImeTail
        Write-DebugLog 'Auto-started IME log tail for sync monitoring' -Level INFO
    }

    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if ($isAdmin) {
        Start-BackgroundWork -Work {
            $results = @()
            try {
                $tasks = Get-ScheduledTask -TaskPath '\Microsoft\Windows\EnterpriseMgmt\*' -ErrorAction Stop |
                    Where-Object { $_.TaskName -eq 'PushLaunch' }
                foreach ($task in $tasks) {
                    Start-ScheduledTask -InputObject $task -ErrorAction Stop
                    $enrollId = ($task.TaskPath -split '\\' | Where-Object { $_ -match '^\w{8}-' }) | Select-Object -First 1
                    $results += "Triggered PushLaunch for $enrollId"
                }
            } catch {
                $results += "Scheduled task trigger failed: $($_.Exception.Message)"
            }
            return $results
        } -OnComplete {
            param($results)
            $msg = ($results | ForEach-Object { $_ }) -join "`n"
            $ui.ImeSyncStatus.Text = $msg
            $ui.BtnIntuneSync.IsEnabled = $true
            Show-Toast 'Intune Sync' $msg -Type Info
            Write-DebugLog "Sync result: $msg" -Level STEP
        }.GetNewClosure() -Variables @{} -Context @{ Name = 'IntuneSync' }
    } else {
        # Non-admin: elevate to trigger PushLaunch scheduled task
        $ui.ImeSyncStatus.Text = 'Elevating to trigger sync...'
        try {
            $tmpOut = [System.IO.Path]::Combine($env:TEMP, 'intune_sync_result.txt')
            $script = @'
$results = @()
try {
    $tasks = Get-ScheduledTask -TaskPath '\Microsoft\Windows\EnterpriseMgmt\*' -ErrorAction Stop | Where-Object { $_.TaskName -eq 'PushLaunch' }
    foreach ($task in $tasks) {
        Start-ScheduledTask -InputObject $task -ErrorAction Stop
        $enrollId = ($task.TaskPath -split '\\' | Where-Object { $_ -match '^\w{8}-' }) | Select-Object -First 1
        $results += "Triggered PushLaunch for $enrollId"
    }
} catch {
    $results += "Failed: $($_.Exception.Message)"
}
$results -join "`n" | Set-Content -Path '__TMPOUT__' -Encoding UTF8 -Force
'@
            $script = $script.Replace('__TMPOUT__', $tmpOut)
            $encoded = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($script))
            Start-Process powershell.exe -ArgumentList "-NoProfile -EncodedCommand $encoded" -Verb RunAs -Wait -WindowStyle Hidden
            $msg = ''
            if (Test-Path $tmpOut) {
                $msg = (Get-Content $tmpOut -Raw -ErrorAction SilentlyContinue).Trim()
                Remove-Item $tmpOut -Force -ErrorAction SilentlyContinue
            }
            if (-not $msg) { $msg = 'Sync triggered (elevated)' }
            $ui.ImeSyncStatus.Text = $msg
            Show-Toast 'Intune Sync' $msg -Type Info
            Write-DebugLog "Sync result (elevated): $msg" -Level STEP
        } catch {
            $ui.ImeSyncStatus.Text = "Elevation cancelled or failed: $($_.Exception.Message)"
            Show-Toast 'Intune Sync' 'Elevation cancelled' -Type Warning
            Write-DebugLog "Intune sync elevation failed: $($_.Exception.Message)" -Level WARN
        }
        $ui.BtnIntuneSync.IsEnabled = $true
    }
}


# ═══════════════════════════════════════════════════════════════════════════════
# SECTION: TOOLS PANEL - Registry Links, Status Codes, Base64, MMP-C
# ═══════════════════════════════════════════════════════════════════════════════

# --- Registry Quick-Links ---
$Script:RegistryLinks = @{
    'BtnRegPolicyManager'   = 'HKLM\SOFTWARE\Microsoft\PolicyManager'
    'BtnRegEnrollments'     = 'HKLM\SOFTWARE\Microsoft\Enrollments'
    'BtnRegProvisioning'    = 'HKLM\SOFTWARE\Microsoft\Provisioning'
    'BtnRegIME'             = 'HKLM\SOFTWARE\Microsoft\IntuneManagementExtension'
    'BtnRegDeclaredConfig'  = 'HKLM\SOFTWARE\Microsoft\DeclaredConfiguration'
    'BtnRegDesktopAppMgmt'  = 'HKLM\SOFTWARE\Microsoft\EnterpriseDesktopAppManagement'
    'BtnRegRebootURIs'      = 'HKLM\SOFTWARE\Microsoft\Provisioning\SyncML\RebootRequiredURIs'
}

function Open-RegistryKey([string]$Path) {
    try {
        # Set regedit's last key so it opens at the right location
        Set-ItemProperty -Path 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit' -Name 'LastKey' -Value "Computer\$Path" -ErrorAction SilentlyContinue
        Start-Process regedit.exe
        Write-DebugLog "Opened registry: $Path" -Level INFO
        Show-Toast 'Registry' "Opened $($Path -replace '.*\\','')" 'info'
    } catch {
        Write-DebugLog "Failed to open registry: $_" -Level ERROR
        Show-Toast 'Registry' "Failed: $($_.Exception.Message)" 'error'
    }
}

# Wire sidebar registry buttons
foreach ($btnName in $Script:RegistryLinks.Keys) {
    $regPath = $Script:RegistryLinks[$btnName]
    if ($ui[$btnName]) {
        $ui[$btnName].Tag = $regPath
        $ui[$btnName].Add_Click({ Open-RegistryKey $this.Tag }.GetNewClosure())
    }
    # Also wire the panel duplicate buttons (suffixed with "2")
    $btn2 = $btnName + '2'
    if ($ui[$btn2]) {
        $ui[$btn2].Tag = $regPath
        $ui[$btn2].Add_Click({ Open-RegistryKey $this.Tag }.GetNewClosure())
    }
}

# --- MDM/SyncML Status Code Dictionary ---
$Script:StatusCodes = @{
    # SyncML/OMA-DM status codes
    200 = 'OK - Command completed successfully'
    201 = 'Item added - New data item created'
    202 = 'Accepted for processing - Request accepted but not yet completed'
    204 = 'No content - Command completed, no data returned'
    206 = 'Partial content - Only part of the command completed'
    207 = 'Conflict resolved with merge'
    208 = 'Conflict resolved with dominant version'
    209 = 'Conflict resolved with duplicate'
    210 = 'Delete without archive - Successfully deleted'
    211 = 'Item not deleted - Not found'
    212 = 'Authentication accepted'
    213 = 'Chunked item accepted - Waiting for more data'
    214 = 'Operation cancelled - Not executed'
    215 = 'Not executed - User interaction required'
    216 = 'Atomic rollback OK - Command undone'
    # Redirection
    300 = 'Multiple choices'
    301 = 'Moved permanently'
    302 = 'Found (moved temporarily)'
    303 = 'See other'
    304 = 'Not modified'
    305 = 'Use proxy'
    # Client errors
    400 = 'Bad request - Malformed command'
    401 = 'Unauthorized - Valid credentials required'
    403 = 'Forbidden - Command understood but refused'
    404 = 'Not found - Target URI not found'
    405 = 'Command not allowed on target'
    406 = 'Optional feature not supported'
    407 = 'Authentication required'
    408 = 'Request timeout'
    409 = 'Conflict - Update conflict'
    410 = 'Gone - Target no longer available'
    411 = 'Size required'
    412 = 'Incomplete command'
    413 = 'Request entity too large'
    414 = 'URI too long'
    415 = 'Unsupported media type or format'
    416 = 'Requested range not satisfiable'
    417 = 'Retry later'
    418 = 'Already exists - Add target already exists'
    420 = 'Device full - No more storage'
    421 = 'Unknown search grammar'
    422 = 'Bad CGI script'
    423 = 'Soft-delete conflict'
    424 = 'Object size mismatch'
    425 = 'Permission denied'
    # Server errors
    500 = 'Command failed - Internal server error'
    501 = 'Not implemented'
    502 = 'Bad gateway'
    503 = 'Service unavailable - Slow sync required'
    504 = 'Gateway timeout'
    505 = 'DTD version not supported'
    506 = 'Processing error'
    507 = 'Atomic failed - One or more operations in atomic block failed'
    508 = 'Refresh required - Full dataset needed instead of delta'
    509 = 'Reserved'
    510 = 'Data store failure'
    511 = 'Server failure'
    512 = 'Application synchronization failed'
    513 = 'Protocol version not supported'
    514 = 'Operation cancelled - Sending side cancelled'
    516 = 'Atomic rollback failed - Cannot undo successfully'
    517 = 'Atomic response too large to fit in a single message'
    # Common HRESULT codes (decimal)
    -2147024891 = 'E_ACCESSDENIED (0x80070005) - Access denied'
    -2147023728 = 'ERROR_NOT_FOUND (0x80070490) - Element not found'
    -2147024894 = 'ERROR_FILE_NOT_FOUND (0x80070002) - File not found'
    -2147024809 = 'E_INVALIDARG (0x80070057) - Invalid argument'
    -2147467259 = 'E_FAIL (0x80004005) - Unspecified failure'
    -2016345812 = 'MDM_E_SYNCML_STATUS_CODE_404 - CSP node not found'
    -2016345807 = 'MDM_E_SYNCML_STATUS_CODE_409 - CSP conflict'
    -2016345795 = 'MDM_E_SYNCML_STATUS_CODE_500 - CSP internal error'
}

function Invoke-StatusCodeLookup([string]$CodeStr) {
    $code = $null
    if (-not [int]::TryParse($CodeStr.Trim(), [ref]$code)) {
        # Try hex
        if ($CodeStr.Trim() -match '^0x([0-9A-Fa-f]+)$') {
            $code = [Convert]::ToInt32($Matches[1], 16)
        } else {
            return "Invalid input: enter a numeric code (decimal or 0x hex)"
        }
    }
    $result = $Script:StatusCodes[$code]
    if ($result) {
        return "$code = $result"
    }
    # Fallback: try Win32 error
    try {
        $win32 = [System.ComponentModel.Win32Exception]::new($code)
        if ($win32.Message -and $win32.Message -ne 'Unknown error') {
            return "$code = Win32: $($win32.Message)"
        }
    } catch {}
    # Fallback: try HRESULT conversion
    if ($code -gt 0) {
        $hr = [int]($code -bor 0x80070000)
        try {
            $win32h = [System.ComponentModel.Win32Exception]::new($code)
            if ($win32h.Message -ne 'Unknown error') { return "$code = HRESULT 0x$($hr.ToString('X8')): $($win32h.Message)" }
        } catch {}
    }
    return "$code = Unknown code. Not found in SyncML/MDM or Win32 tables."
}

# Wire status code lookup (sidebar)
if ($ui.BtnLookupStatus) {
    $ui.BtnLookupStatus.Add_Click({
        $code = $ui.TxtStatusCode.Text
        if ($code) { $ui.StatusCodeResult.Text = Invoke-StatusCodeLookup $code }
    }.GetNewClosure())
}
# Wire status code lookup (panel)
if ($ui.BtnLookupStatus2) {
    $ui.BtnLookupStatus2.Add_Click({
        $code = $ui.TxtStatusCode2.Text
        if ($code) { $ui.StatusCodeResult2.Text = Invoke-StatusCodeLookup $code }
    }.GetNewClosure())
}

# --- Base64 Decoder ---
function Invoke-Base64Decode([string]$Input) {
    if (-not $Input -or $Input.Trim().Length -eq 0) { return 'No input provided' }
    try {
        $bytes = [Convert]::FromBase64String($Input.Trim())
        $decoded = [System.Text.Encoding]::UTF8.GetString($bytes)
        # Try JSON pretty-print
        try {
            $json = $decoded | ConvertFrom-Json -ErrorAction Stop
            $decoded = $json | ConvertTo-Json -Depth 10
        } catch {}
        return $decoded
    } catch {
        # Try URL-safe Base64 variant
        try {
            $padded = $Input.Trim().Replace('-', '+').Replace('_', '/')
            switch ($padded.Length % 4) { 2 { $padded += '==' }; 3 { $padded += '=' } }
            $bytes = [Convert]::FromBase64String($padded)
            return [System.Text.Encoding]::UTF8.GetString($bytes)
        } catch {
            return "Failed to decode: not valid Base64 ($($_.Exception.Message))"
        }
    }
}

# Wire Base64 decoder (sidebar)
if ($ui.BtnDecodeBase64) {
    $ui.BtnDecodeBase64.Add_Click({
        $input = $ui.TxtBase64Input.Text
        $ui.Base64Result.Text = Invoke-Base64Decode $input
    }.GetNewClosure())
}
# Wire Base64 decoder (panel)
if ($ui.BtnDecodeBase642) {
    $ui.BtnDecodeBase642.Add_Click({
        $input = $ui.TxtBase64Input2.Text
        $result = Invoke-Base64Decode $input
        $ui.Base64Result2.Text = $result
        $ui.Base64Result2.Visibility = 'Visible'
    }.GetNewClosure())
}

# --- MMP-C Sync Trigger ---
function Invoke-MmpCSync {
    if ($ui.BtnMmpCSync) { $ui.BtnMmpCSync.IsEnabled = $false }
    if ($ui.MmpCSyncStatus) { $ui.MmpCSyncStatus.Text = 'Triggering MMP-C sync...' }
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    Write-DebugLog "MMP-C sync triggered (admin=$isAdmin)" -Level STEP

    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if ($isAdmin) {
        Start-BackgroundWork -Work {
            $results = @()
            try {
                # MMP-C uses "Schedule #3 created by enrollment client" tasks
                $tasks = Get-ScheduledTask -TaskPath '\Microsoft\Windows\EnterpriseMgmt\*' -ErrorAction Stop |
                    Where-Object { $_.TaskName -like 'Schedule #3*' }
                if ($tasks) {
                    foreach ($task in $tasks) {
                        $enrollId = ($task.TaskPath -split '\\' | Where-Object { $_ -match '^\w{8}-' }) | Select-Object -First 1
                        $results += "Starting MMP-C Schedule #3 for $enrollId..."
                        Start-ScheduledTask -InputObject $task -ErrorAction Stop
                        $results += "Triggered MMP-C Schedule #3 for $enrollId"
                    }
                } else {
                    $results += "No MMP-C scheduled tasks found (Schedule #3). Device may not be co-managed."
                }
            } catch {
                $results += "MMP-C trigger failed: $($_.Exception.Message)"
            }
            return $results
        } -OnComplete {
            param($results)
            $msg = ($results | ForEach-Object { $_ }) -join "`n"
            if ($ui.MmpCSyncStatus) { $ui.MmpCSyncStatus.Text = $msg }
            if ($ui.BtnMmpCSync) { $ui.BtnMmpCSync.IsEnabled = $true }
            Show-Toast 'MMP-C Sync' $msg 'info'
            Write-DebugLog "MMP-C result: $msg" -Level STEP
        }.GetNewClosure() -Variables @{} -Context @{ Name = 'MmpCSync' }
    } else {
        if ($ui.MmpCSyncStatus) { $ui.MmpCSyncStatus.Text = 'MMP-C sync requires admin privileges.' }
        if ($ui.BtnMmpCSync) { $ui.BtnMmpCSync.IsEnabled = $true }
        Show-Toast 'MMP-C Sync' 'Admin privileges required to trigger scheduled tasks.' 'warning'
    }
}

if ($ui.BtnMmpCSync) {
    $ui.BtnMmpCSync.Add_Click({ Invoke-MmpCSync }.GetNewClosure())
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION: SYNCML ETW TRACE - Option A (logman capture + Get-WinEvent parse)
# ═══════════════════════════════════════════════════════════════════════════════

$Script:EtwTraceActive  = $false
$Script:EtwSessionName  = 'PolicyPilotTrace'
$Script:EtwTracePath    = Join-Path $env:TEMP 'PolicyPilot_ETW.etl'
$Script:EtwProviderGuid = '{0EC685CD-64E4-4375-92AD-4086B6AF5F1D}'  # OmaDmClient

function Start-EtwTrace {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        $msg = 'ETW tracing requires admin privileges. Restart PolicyPilot as Administrator.'
        if ($ui.EtwTraceStatus)  { $ui.EtwTraceStatus.Text  = $msg }
        if ($ui.EtwTraceStatus2) { $ui.EtwTraceStatus2.Text = $msg }
        Show-Toast 'ETW Trace' $msg 'warning'
        Write-DebugLog "ETW trace: admin required" -Level WARN
        return
    }

    # Clean up any stale session
    Write-DebugLog "ETW trace: cleaning up stale session '$($Script:EtwSessionName)'" -Level DEBUG
    try { & logman stop $Script:EtwSessionName -ets 2>$null | Out-Null } catch {}
    if (Test-Path $Script:EtwTracePath) {
        Remove-Item $Script:EtwTracePath -Force -ErrorAction SilentlyContinue
    }

    # Start trace session
    $logmanArgs = @('create', 'trace', $Script:EtwSessionName,
              '-p', $Script:EtwProviderGuid, '0xFFFFFFFFFFFFFFFF', '0xFF',
              '-o', $Script:EtwTracePath,
              '-ets',
              '-f', 'bincirc',
              '-max', '32')
    $proc = Start-Process -FilePath 'logman.exe' -ArgumentList $logmanArgs -NoNewWindow -Wait -PassThru -RedirectStandardOutput "$env:TEMP\pp_logman_out.txt" -RedirectStandardError "$env:TEMP\pp_logman_err.txt"

    if ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq -1) {
        $Script:EtwTraceActive = $true
        $ts = Get-Date -Format 'HH:mm:ss'
        $msg = "Trace active since $ts. Trigger an MDM sync, then click Stop & Analyze."
        if ($ui.EtwTraceStatus)  { $ui.EtwTraceStatus.Text  = $msg }
        if ($ui.EtwTraceStatus2) { $ui.EtwTraceStatus2.Text = $msg }
        if ($ui.BtnEtwStart)  { $ui.BtnEtwStart.IsEnabled  = $false }
        if ($ui.BtnEtwStop)   { $ui.BtnEtwStop.IsEnabled   = $true }
        if ($ui.BtnEtwStart2) { $ui.BtnEtwStart2.IsEnabled = $false }
        if ($ui.BtnEtwStop2)  { $ui.BtnEtwStop2.IsEnabled  = $true }
        Show-Toast 'ETW Trace' 'Trace session started - capturing OMA-DM events' 'info'
        Write-DebugLog "ETW trace started: session=$($Script:EtwSessionName) path=$($Script:EtwTracePath)" -Level STEP
    } else {
        $errText = ''
        if (Test-Path "$env:TEMP\pp_logman_err.txt") { $errText = [System.IO.File]::ReadAllText("$env:TEMP\pp_logman_err.txt").Trim() }
        if (-not $errText -and (Test-Path "$env:TEMP\pp_logman_out.txt")) { $errText = [System.IO.File]::ReadAllText("$env:TEMP\pp_logman_out.txt").Trim() }
        $msg = "logman failed (exit $($proc.ExitCode)): $errText"
        if ($ui.EtwTraceStatus)  { $ui.EtwTraceStatus.Text  = $msg }
        if ($ui.EtwTraceStatus2) { $ui.EtwTraceStatus2.Text = $msg }
        Show-Toast 'ETW Trace' "Failed to start: $errText" 'error'
        Write-DebugLog "ETW trace start failed: $msg" -Level ERROR
    }
}

function Stop-EtwTrace {
    if (-not $Script:EtwTraceActive) {
        $msg = 'No active trace session.'
        if ($ui.EtwTraceStatus)  { $ui.EtwTraceStatus.Text  = $msg }
        if ($ui.EtwTraceStatus2) { $ui.EtwTraceStatus2.Text = $msg }
        return
    }

    $msg = 'Stopping trace and parsing events...'
    if ($ui.EtwTraceStatus)  { $ui.EtwTraceStatus.Text  = $msg }
    if ($ui.EtwTraceStatus2) { $ui.EtwTraceStatus2.Text = $msg }

    Write-DebugLog "ETW trace stopping: session=$($Script:EtwSessionName)" -Level STEP
    try { & logman stop $Script:EtwSessionName -ets 2>$null | Out-Null } catch {}
    $Script:EtwTraceActive = $false
    Write-DebugLog "ETW trace stopped, parsing ETL: $($Script:EtwTracePath)" -Level DEBUG

    if ($ui.BtnEtwStart)  { $ui.BtnEtwStart.IsEnabled  = $true }
    if ($ui.BtnEtwStop)   { $ui.BtnEtwStop.IsEnabled   = $false }
    if ($ui.BtnEtwStart2) { $ui.BtnEtwStart2.IsEnabled = $true }
    if ($ui.BtnEtwStop2)  { $ui.BtnEtwStop2.IsEnabled  = $false }

    $etlPath = $Script:EtwTracePath
    Start-BackgroundWork -Work {
        param($etlPath)
        $result = [System.Text.StringBuilder]::new(16384)
        $eventCount = 0
        $syncmlCount = 0

        if (-not (Test-Path $etlPath)) {
            [void]$result.AppendLine("No ETW data captured (file not found: $etlPath)")
            return @{ Text = $result.ToString(); Events = 0; SyncML = 0 }
        }

        try {
            $events = Get-WinEvent -Path $etlPath -Oldest -ErrorAction Stop
            $eventCount = $events.Count
            [void]$result.AppendLine("=== SyncML ETW Trace Results ===")
            [void]$result.AppendLine("Captured $eventCount events from OmaDmClient provider")
            [void]$result.AppendLine("ETL file: $etlPath")
            [void]$result.AppendLine("=" * 50)
            [void]$result.AppendLine()

            foreach ($evt in $events) {
                $ts = $evt.TimeCreated.ToString('HH:mm:ss.fff')
                $id = $evt.Id
                $evtMsg = if ($evt.Message) { $evt.Message } else { '' }
                $xmlBody = $null

                # Check message for SyncML XML
                if ($evtMsg -match '<SyncML[\s>]') {
                    if ($evtMsg -match '(?s)(<SyncML.*?</SyncML>)') {
                        $xmlBody = $Matches[1]
                    }
                }
                # Check event properties for SyncML XML
                if (-not $xmlBody -and $evt.Properties) {
                    foreach ($prop in $evt.Properties) {
                        $val = "$($prop.Value)"
                        if ($val -match '<SyncML[\s>]') {
                            if ($val -match '(?s)(<SyncML.*?</SyncML>)') {
                                $xmlBody = $Matches[1]
                            } else {
                                $xmlBody = $val
                            }
                            break
                        }
                    }
                }

                if ($xmlBody) {
                    $syncmlCount++
                    [void]$result.AppendLine("[$ts] Event $id - SyncML Message #$syncmlCount")
                    [void]$result.AppendLine("-" * 40)
                    try {
                        $xdoc = [xml]$xmlBody
                        $sw = [System.IO.StringWriter]::new()
                        $xw = [System.Xml.XmlTextWriter]::new($sw)
                        $xw.Formatting = 'Indented'
                        $xw.Indentation = 2
                        $xdoc.WriteTo($xw)
                        $xw.Flush()
                        [void]$result.AppendLine($sw.ToString())
                        $xw.Dispose(); $sw.Dispose()
                    } catch {
                        [void]$result.AppendLine($xmlBody)
                    }
                    [void]$result.AppendLine()
                } else {
                    # Non-SyncML: show Event ID + first 200 chars
                    $shortMsg = if ($evtMsg.Length -gt 200) { $evtMsg.Substring(0, 200) + '...' } else { $evtMsg }
                    if ($shortMsg) {
                        [void]$result.AppendLine("[$ts] Event ${id}: $shortMsg")
                    } else {
                        # No message - show property values
                        $propText = ($evt.Properties | ForEach-Object { "$($_.Value)" }) -join ' | '
                        if ($propText.Length -gt 200) { $propText = $propText.Substring(0, 200) + '...' }
                        [void]$result.AppendLine("[$ts] Event ${id}: $propText")
                    }
                }
            }

            [void]$result.AppendLine()
            [void]$result.AppendLine("=== Summary: $eventCount events, $syncmlCount SyncML messages ===")
        } catch {
            [void]$result.AppendLine("Failed to parse ETL: $($_.Exception.Message)")
            [void]$result.AppendLine("The .etl file may have no events or an unsupported format.")
            [void]$result.AppendLine("ETL path: $etlPath")
        }

        return @{ Text = $result.ToString(); Events = $eventCount; SyncML = $syncmlCount }
    } -OnComplete {
        param($result)
        $text = $result.Text
        $evts = $result.Events
        $syncml = $result.SyncML

        if ($ui.EtwResultBox) {
            $ui.EtwResultBox.Text = $text
            if ($ui.EtwResultScroller) { $ui.EtwResultScroller.ScrollToTop() }
        }
        $statusMsg = "Trace complete: $evts events captured, $syncml SyncML messages found."
        if ($ui.EtwTraceStatus)  { $ui.EtwTraceStatus.Text  = $statusMsg }
        if ($ui.EtwTraceStatus2) { $ui.EtwTraceStatus2.Text = $statusMsg }
        Show-Toast 'ETW Trace' $statusMsg 'info'
        Write-DebugLog "ETW trace complete: $evts events, $syncml SyncML messages" -Level STEP
    }.GetNewClosure() -Variables @{ etlPath = $etlPath } -Context @{ Name = 'EtwTrace' }
}

function Clear-EtwResults {
    if ($ui.EtwResultBox)    { $ui.EtwResultBox.Text = '' }
    if ($ui.EtwTraceStatus)  { $ui.EtwTraceStatus.Text = '' }
    if ($ui.EtwTraceStatus2) { $ui.EtwTraceStatus2.Text = '' }
}

# Wire ETW trace buttons (sidebar)
if ($ui.BtnEtwStart) { $ui.BtnEtwStart.Add_Click({ Start-EtwTrace }.GetNewClosure()) }
if ($ui.BtnEtwStop)  { $ui.BtnEtwStop.Add_Click({ Stop-EtwTrace }.GetNewClosure()) }
if ($ui.BtnEtwClear) { $ui.BtnEtwClear.Add_Click({ Clear-EtwResults }.GetNewClosure()) }
# Wire ETW trace buttons (panel)
if ($ui.BtnEtwStart2) { $ui.BtnEtwStart2.Add_Click({ Start-EtwTrace }.GetNewClosure()) }
if ($ui.BtnEtwStop2)  { $ui.BtnEtwStop2.Add_Click({ Stop-EtwTrace }.GetNewClosure()) }
if ($ui.BtnEtwClear2) { $ui.BtnEtwClear2.Add_Click({ Clear-EtwResults }.GetNewClosure()) }

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION: AUTOPILOT HARDWARE HASH DECODER
# ═══════════════════════════════════════════════════════════════════════════════

function Decode-AutopilotHash([string]$Hash) {
    if (-not $Hash -or $Hash.Trim().Length -lt 100) {
        return 'Invalid hash - paste the full Base64-encoded hardware hash (typically 2-4 KB).'
    }
    try {
        $bytes = [Convert]::FromBase64String($Hash.Trim())
    } catch {
        return "Not valid Base64: $($_.Exception.Message)"
    }

    $sb = [System.Text.StringBuilder]::new(2048)
    [void]$sb.AppendLine("=== Autopilot Hardware Hash ===")
    [void]$sb.AppendLine("Total size: $($bytes.Length) bytes")
    [void]$sb.AppendLine()

    # The hash starts with a 2-byte version, then TLV (Type-Length-Value) entries
    $offset = 0
    if ($bytes.Length -lt 4) { return $sb.ToString() + "Hash too short to parse." }

    $version = [BitConverter]::ToUInt16($bytes, 0)
    [void]$sb.AppendLine("Version: $version")
    $offset = 2

    $typeNames = @{
        1 = 'SMBIOS Data'
        2 = 'Disk Serial'
        3 = 'MAC Address(es)'
        4 = 'SMBIOS Hash'
        5 = 'Disk Serial Hash'
        6 = 'MAC Address Hash'
        7 = 'TPM EK Public Key Hash'
        8 = 'SMBIOS UUID'
    }

    while ($offset -lt ($bytes.Length - 4)) {
        try {
            $type = [BitConverter]::ToUInt16($bytes, $offset)
            $length = [BitConverter]::ToUInt16($bytes, $offset + 2)
            $offset += 4

            if ($length -eq 0 -or ($offset + $length) -gt $bytes.Length) { break }

            $data = $bytes[$offset..($offset + $length - 1)]
            $typeName = if ($typeNames[$type]) { $typeNames[$type] } else { "Unknown (Type $type)" }
            [void]$sb.AppendLine("--- $typeName (Type=$type, Length=$length) ---")

            switch ($type) {
                1 {
                    # SMBIOS - try to extract ASCII strings
                    $text = [System.Text.Encoding]::ASCII.GetString($data) -replace '[^\x20-\x7E]', '.'
                    # Try to find manufacturer, model, serial in the blob
                    $strings = [System.Text.Encoding]::ASCII.GetString($data) -split '\x00+' | Where-Object { $_.Length -gt 2 -and $_ -match '[A-Za-z]' }
                    if ($strings.Count -ge 3) {
                        [void]$sb.AppendLine("  Manufacturer: $($strings[0])")
                        [void]$sb.AppendLine("  Model:        $($strings[1])")
                        [void]$sb.AppendLine("  Serial:       $($strings[2])")
                        if ($strings.Count -gt 3) {
                            [void]$sb.AppendLine("  Other:        $($strings[3..($strings.Count-1)] -join ', ')")
                        }
                    } else {
                        [void]$sb.AppendLine("  Raw strings: $($strings -join ' | ')")
                    }
                }
                2 {
                    # Disk serial
                    $diskSerial = [System.Text.Encoding]::ASCII.GetString($data) -replace '[^\x20-\x7E]', ''
                    [void]$sb.AppendLine("  Serial: $($diskSerial.Trim())")
                }
                3 {
                    # MAC addresses - 6 bytes per MAC
                    $macCount = [Math]::Floor($length / 6)
                    for ($m = 0; $m -lt $macCount; $m++) {
                        $mac = ($data[($m*6)..($m*6+5)] | ForEach-Object { $_.ToString('X2') }) -join ':'
                        [void]$sb.AppendLine("  MAC[$m]: $mac")
                    }
                }
                {$_ -in 4,5,6,7} {
                    # Hash values - display as hex
                    $hex = ($data | ForEach-Object { $_.ToString('X2') }) -join ''
                    [void]$sb.AppendLine("  Hash: $($hex.Substring(0, [Math]::Min(64, $hex.Length)))$(if($hex.Length -gt 64){'...'})")
                }
                8 {
                    # SMBIOS UUID - 16 bytes
                    if ($data.Length -ge 16) {
                        try {
                            $guid = [Guid]::new($data[0..15], 0)
                            [void]$sb.AppendLine("  UUID: $guid")
                        } catch {
                            $hex = ($data[0..15] | ForEach-Object { $_.ToString('X2') }) -join ''
                            [void]$sb.AppendLine("  UUID (raw): $hex")
                        }
                    }
                }
                default {
                    # Generic hex dump
                    $hex = ($data[0..([Math]::Min(63, $data.Length-1))] | ForEach-Object { $_.ToString('X2') }) -join ' '
                    [void]$sb.AppendLine("  Data: $hex$(if($data.Length -gt 64){'...'})")
                }
            }
            [void]$sb.AppendLine()
            $offset += $length
        } catch {
            [void]$sb.AppendLine("Parse error at offset $offset`: $($_.Exception.Message)")
            break
        }
    }
    return $sb.ToString()
}

# Wire Autopilot decode (sidebar)
if ($ui.BtnDecodeAutopilot) {
    $ui.BtnDecodeAutopilot.Add_Click({
        $hash = $ui.TxtAutopilotHash.Text
        $ui.AutopilotResult.Text = Decode-AutopilotHash $hash
    }.GetNewClosure())
}
# Wire Autopilot decode (panel)
if ($ui.BtnDecodeAutopilot2) {
    $ui.BtnDecodeAutopilot2.Add_Click({
        $hash = $ui.TxtAutopilotHash2.Text
        $result = Decode-AutopilotHash $hash
        $ui.AutopilotResult2.Text = $result
        $ui.AutopilotResult2.Visibility = 'Visible'
    }.GetNewClosure())
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION: WIFI / VPN PROFILE VIEWER
# ═══════════════════════════════════════════════════════════════════════════════

function Get-WifiProfiles {
    $sb = [System.Text.StringBuilder]::new(4096)
    try {
        $output = & netsh wlan show profiles 2>&1
        if ($LASTEXITCODE -ne 0) {
            return "WLAN service not available: $output"
        }
        $profiles = $output | Select-String 'All User Profile\s*:\s*(.+)' | ForEach-Object { $_.Matches[0].Groups[1].Value.Trim() }
        if (-not $profiles -or $profiles.Count -eq 0) {
            return "No saved WiFi profiles found."
        }
        [void]$sb.AppendLine("=== WiFi Profiles ($($profiles.Count) found) ===")
        [void]$sb.AppendLine()
        foreach ($profile in $profiles) {
            [void]$sb.AppendLine("  $profile")
            # Get details for each profile
            $detail = & netsh wlan show profile name="$profile" 2>&1
            $auth = ($detail | Select-String 'Authentication\s*:\s*(.+)' | Select-Object -First 1)
            $cipher = ($detail | Select-String 'Cipher\s*:\s*(.+)' | Select-Object -First 1)
            $connMode = ($detail | Select-String 'Connection mode\s*:\s*(.+)' | Select-Object -First 1)
            if ($auth) { [void]$sb.AppendLine("    Auth: $($auth.Matches[0].Groups[1].Value.Trim())") }
            if ($cipher) { [void]$sb.AppendLine("    Cipher: $($cipher.Matches[0].Groups[1].Value.Trim())") }
            if ($connMode) { [void]$sb.AppendLine("    Mode: $($connMode.Matches[0].Groups[1].Value.Trim())") }
            [void]$sb.AppendLine()
        }
    } catch {
        [void]$sb.AppendLine("Error enumerating WiFi profiles: $($_.Exception.Message)")
    }
    return $sb.ToString()
}

function Get-VpnProfiles {
    $sb = [System.Text.StringBuilder]::new(4096)
    try {
        $vpns = Get-VpnConnection -ErrorAction Stop
        if (-not $vpns -or $vpns.Count -eq 0) {
            return "No VPN connections configured."
        }
        [void]$sb.AppendLine("=== VPN Connections ($($vpns.Count) found) ===")
        [void]$sb.AppendLine()
        foreach ($vpn in $vpns) {
            [void]$sb.AppendLine("  $($vpn.Name)")
            [void]$sb.AppendLine("    Server:  $($vpn.ServerAddress)")
            [void]$sb.AppendLine("    Tunnel:  $($vpn.TunnelType)")
            [void]$sb.AppendLine("    Auth:    $($vpn.AuthenticationMethod -join ', ')")
            [void]$sb.AppendLine("    Status:  $(if($vpn.ConnectionStatus -eq 'Connected'){'Connected'}else{'Disconnected'})")
            [void]$sb.AppendLine()
        }
    } catch {
        [void]$sb.AppendLine("VPN not available: $($_.Exception.Message)")
    }
    return $sb.ToString()
}

# Wire WiFi (sidebar)
if ($ui.BtnLoadWifi) {
    $ui.BtnLoadWifi.Add_Click({
        Write-DebugLog 'Loading WiFi profiles (sidebar)' -Level DEBUG
        $ui.NetworkProfileStatus.Text = 'Loading WiFi profiles...'
        Start-BackgroundWork -Work {
            function Get-WifiProfiles {
                $sb = [System.Text.StringBuilder]::new(4096)
                try {
                    $output = & netsh wlan show profiles 2>&1
                    if ($LASTEXITCODE -ne 0) {
                        return "WLAN service not available: $output"
                    }
                    $profiles = $output | Select-String 'All User Profile\s*:\s*(.+)' | ForEach-Object { $_.Matches[0].Groups[1].Value.Trim() }
                    if (-not $profiles -or $profiles.Count -eq 0) {
                        return "No saved WiFi profiles found."
                    }
                    [void]$sb.AppendLine("=== WiFi Profiles ($($profiles.Count) found) ===")
                    [void]$sb.AppendLine()
                    foreach ($profile in $profiles) {
                        [void]$sb.AppendLine("  $profile")
                        $detail = & netsh wlan show profile name="$profile" 2>&1
                        $auth = ($detail | Select-String 'Authentication\s*:\s*(.+)' | Select-Object -First 1)
                        $cipher = ($detail | Select-String 'Cipher\s*:\s*(.+)' | Select-Object -First 1)
                        $connMode = ($detail | Select-String 'Connection mode\s*:\s*(.+)' | Select-Object -First 1)
                        if ($auth) { [void]$sb.AppendLine("    Auth: $($auth.Matches[0].Groups[1].Value.Trim())") }
                        if ($cipher) { [void]$sb.AppendLine("    Cipher: $($cipher.Matches[0].Groups[1].Value.Trim())") }
                        if ($connMode) { [void]$sb.AppendLine("    Mode: $($connMode.Matches[0].Groups[1].Value.Trim())") }
                        [void]$sb.AppendLine()
                    }
                } catch {
                    [void]$sb.AppendLine("Error enumerating WiFi profiles: $($_.Exception.Message)")
                }
                return $sb.ToString()
            }
            Get-WifiProfiles
        } -OnComplete {
            param($result)
            $text = $result | ForEach-Object { $_ }
            if ($ui.NetworkProfileStatus) { $ui.NetworkProfileStatus.Text = 'Done' }
            Show-Toast 'WiFi Profiles' 'WiFi profiles loaded' 'info'
            Write-DebugLog "WiFi profiles loaded" -Level INFO
        }.GetNewClosure() -Variables @{} -Context @{ Name = 'WiFiProfiles' }
    }.GetNewClosure())
}
# Wire WiFi (panel)
if ($ui.BtnLoadWifi2) {
    $ui.BtnLoadWifi2.Add_Click({
        Write-DebugLog 'Loading WiFi profiles (panel)' -Level DEBUG
        Start-BackgroundWork -Work {
            function Get-WifiProfiles {
                $sb = [System.Text.StringBuilder]::new(4096)
                try {
                    $output = & netsh wlan show profiles 2>&1
                    if ($LASTEXITCODE -ne 0) {
                        return "WLAN service not available: $output"
                    }
                    $profiles = $output | Select-String 'All User Profile\s*:\s*(.+)' | ForEach-Object { $_.Matches[0].Groups[1].Value.Trim() }
                    if (-not $profiles -or $profiles.Count -eq 0) {
                        return "No saved WiFi profiles found."
                    }
                    [void]$sb.AppendLine("=== WiFi Profiles ($($profiles.Count) found) ===")
                    [void]$sb.AppendLine()
                    foreach ($profile in $profiles) {
                        [void]$sb.AppendLine("  $profile")
                        $detail = & netsh wlan show profile name="$profile" 2>&1
                        $auth = ($detail | Select-String 'Authentication\s*:\s*(.+)' | Select-Object -First 1)
                        $cipher = ($detail | Select-String 'Cipher\s*:\s*(.+)' | Select-Object -First 1)
                        $connMode = ($detail | Select-String 'Connection mode\s*:\s*(.+)' | Select-Object -First 1)
                        if ($auth) { [void]$sb.AppendLine("    Auth: $($auth.Matches[0].Groups[1].Value.Trim())") }
                        if ($cipher) { [void]$sb.AppendLine("    Cipher: $($cipher.Matches[0].Groups[1].Value.Trim())") }
                        if ($connMode) { [void]$sb.AppendLine("    Mode: $($connMode.Matches[0].Groups[1].Value.Trim())") }
                        [void]$sb.AppendLine()
                    }
                } catch {
                    [void]$sb.AppendLine("Error enumerating WiFi profiles: $($_.Exception.Message)")
                }
                return $sb.ToString()
            }
            Get-WifiProfiles
        } -OnComplete {
            param($result)
            $text = $result | ForEach-Object { $_ }
            if ($ui.NetworkProfileResult) {
                $ui.NetworkProfileResult.Text = $text -join "`n"
                $ui.NetworkProfileResult.Visibility = 'Visible'
            }
        }.GetNewClosure() -Variables @{} -Context @{ Name = 'WiFiProfiles2' }
    }.GetNewClosure())
}
# Wire VPN (sidebar)
if ($ui.BtnLoadVpn) {
    $ui.BtnLoadVpn.Add_Click({
        Write-DebugLog 'Loading VPN connections (sidebar)' -Level DEBUG
        $ui.NetworkProfileStatus.Text = 'Loading VPN connections...'
        Start-BackgroundWork -Work {
            function Get-VpnProfiles {
                $sb = [System.Text.StringBuilder]::new(4096)
                try {
                    $vpns = Get-VpnConnection -ErrorAction Stop
                    if (-not $vpns -or $vpns.Count -eq 0) {
                        return "No VPN connections configured."
                    }
                    [void]$sb.AppendLine("=== VPN Connections ($($vpns.Count) found) ===")
                    [void]$sb.AppendLine()
                    foreach ($vpn in $vpns) {
                        [void]$sb.AppendLine("  $($vpn.Name)")
                        [void]$sb.AppendLine("    Server:  $($vpn.ServerAddress)")
                        [void]$sb.AppendLine("    Tunnel:  $($vpn.TunnelType)")
                        [void]$sb.AppendLine("    Auth:    $($vpn.AuthenticationMethod -join ', ')")
                        [void]$sb.AppendLine("    Status:  $(if($vpn.ConnectionStatus -eq 'Connected'){'Connected'}else{'Disconnected'})")
                        [void]$sb.AppendLine()
                    }
                } catch {
                    [void]$sb.AppendLine("VPN not available: $($_.Exception.Message)")
                }
                return $sb.ToString()
            }
            Get-VpnProfiles
        } -OnComplete {
            param($result)
            $text = $result | ForEach-Object { $_ }
            if ($ui.NetworkProfileStatus) { $ui.NetworkProfileStatus.Text = 'Done' }
            Show-Toast 'VPN' 'VPN connections loaded' 'info'
            Write-DebugLog 'VPN profiles loaded' -Level INFO
        }.GetNewClosure() -Variables @{} -Context @{ Name = 'VpnProfiles' }
    }.GetNewClosure())
}
# Wire VPN (panel)
if ($ui.BtnLoadVpn2) {
    $ui.BtnLoadVpn2.Add_Click({
        Write-DebugLog 'Loading VPN connections (panel)' -Level DEBUG
        Start-BackgroundWork -Work {
            function Get-VpnProfiles {
                $sb = [System.Text.StringBuilder]::new(4096)
                try {
                    $vpns = Get-VpnConnection -ErrorAction Stop
                    if (-not $vpns -or $vpns.Count -eq 0) {
                        return "No VPN connections configured."
                    }
                    [void]$sb.AppendLine("=== VPN Connections ($($vpns.Count) found) ===")
                    [void]$sb.AppendLine()
                    foreach ($vpn in $vpns) {
                        [void]$sb.AppendLine("  $($vpn.Name)")
                        [void]$sb.AppendLine("    Server:  $($vpn.ServerAddress)")
                        [void]$sb.AppendLine("    Tunnel:  $($vpn.TunnelType)")
                        [void]$sb.AppendLine("    Auth:    $($vpn.AuthenticationMethod -join ', ')")
                        [void]$sb.AppendLine("    Status:  $(if($vpn.ConnectionStatus -eq 'Connected'){'Connected'}else{'Disconnected'})")
                        [void]$sb.AppendLine()
                    }
                } catch {
                    [void]$sb.AppendLine("VPN not available: $($_.Exception.Message)")
                }
                return $sb.ToString()
            }
            Get-VpnProfiles
        } -OnComplete {
            param($result)
            $text = $result | ForEach-Object { $_ }
            if ($ui.NetworkProfileResult) {
                $ui.NetworkProfileResult.Text = $text -join "`n"
                $ui.NetworkProfileResult.Visibility = 'Visible'
            }
            Write-DebugLog 'VPN profiles loaded (panel)' -Level INFO
        }.GetNewClosure() -Variables @{} -Context @{ Name = 'VpnProfiles2' }
    }.GetNewClosure())
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION: MDM NODE CACHE EXPLORER
# ═══════════════════════════════════════════════════════════════════════════════

function Get-NodeCacheData {
    $sb = [System.Text.StringBuilder]::new(8192)
    $totalNodes = 0
    $basePaths = @(
        'HKLM:\SOFTWARE\Microsoft\Provisioning\NodeCache'
    )

    foreach ($basePath in $basePaths) {
        if (-not (Test-Path $basePath)) {
            [void]$sb.AppendLine("Registry path not found: $basePath")
            continue
        }

        [void]$sb.AppendLine("=== MDM Node Cache ===")
        [void]$sb.AppendLine("Path: $basePath")
        [void]$sb.AppendLine()

        try {
            $cspFolders = Get-ChildItem -Path $basePath -ErrorAction Stop
            foreach ($csp in $cspFolders) {
                $cspName = $csp.PSChildName
                $devicePath = Join-Path $csp.PSPath 'Device'
                $userPath = Join-Path $csp.PSPath 'User'

                foreach ($scope in @(@{Name='Device';Path=$devicePath}, @{Name='User';Path=$userPath})) {
                    if (-not (Test-Path $scope.Path)) { continue }
                    $serverPaths = Get-ChildItem -Path $scope.Path -ErrorAction SilentlyContinue
                    foreach ($server in $serverPaths) {
                        $nodesPath = Join-Path $server.PSPath 'Nodes'
                        if (-not (Test-Path $nodesPath)) { continue }

                        [void]$sb.AppendLine("--- CSP: $cspName | Scope: $($scope.Name) | Server: $($server.PSChildName) ---")

                        $nodes = Get-ChildItem -Path $nodesPath -ErrorAction SilentlyContinue
                        foreach ($node in $nodes) {
                            $totalNodes++
                            $nodeUri = $node.GetValue('NodeUri')
                            $expectedVal = $node.GetValue('ExpectedValue')

                            if ($nodeUri) {
                                $shortVal = if ($expectedVal -and $expectedVal.Length -gt 80) { $expectedVal.Substring(0, 80) + '...' } else { $expectedVal }
                                [void]$sb.AppendLine("  $nodeUri = $shortVal")
                            }
                        }
                        [void]$sb.AppendLine()
                    }
                }
            }
        } catch {
            [void]$sb.AppendLine("Error reading node cache: $($_.Exception.Message)")
        }
    }

    [void]$sb.AppendLine("=== Total: $totalNodes cached nodes ===")
    return @{ Text = $sb.ToString(); Count = $totalNodes }
}

# Wire Node Cache (sidebar)
if ($ui.BtnLoadNodeCache) {
    $ui.BtnLoadNodeCache.Add_Click({
        Write-DebugLog 'Loading MDM node cache (sidebar)' -Level DEBUG
        $ui.NodeCacheStatus.Text = 'Loading node cache...'
        Start-BackgroundWork -Work {
            function Get-NodeCacheData {
                $sb = [System.Text.StringBuilder]::new(8192)
                $totalNodes = 0
                $basePaths = @('HKLM:\SOFTWARE\Microsoft\Provisioning\NodeCache')
                foreach ($basePath in $basePaths) {
                    if (-not (Test-Path $basePath)) {
                        [void]$sb.AppendLine("Registry path not found: $basePath")
                        continue
                    }
                    [void]$sb.AppendLine("=== MDM Node Cache ===")
                    [void]$sb.AppendLine("Path: $basePath")
                    [void]$sb.AppendLine()
                    try {
                        $cspFolders = Get-ChildItem -Path $basePath -ErrorAction Stop
                        foreach ($csp in $cspFolders) {
                            $cspName = $csp.PSChildName
                            $devicePath = Join-Path $csp.PSPath 'Device'
                            $userPath   = Join-Path $csp.PSPath 'User'
                            foreach ($scope in @(@{Name='Device';Path=$devicePath}, @{Name='User';Path=$userPath})) {
                                if (-not (Test-Path $scope.Path)) { continue }
                                $serverPaths = Get-ChildItem -Path $scope.Path -ErrorAction SilentlyContinue
                                foreach ($server in $serverPaths) {
                                    $nodesPath = Join-Path $server.PSPath 'Nodes'
                                    if (-not (Test-Path $nodesPath)) { continue }
                                    [void]$sb.AppendLine("--- CSP: $cspName | Scope: $($scope.Name) | Server: $($server.PSChildName) ---")
                                    $nodes = Get-ChildItem -Path $nodesPath -ErrorAction SilentlyContinue
                                    foreach ($node in $nodes) {
                                        $totalNodes++
                                        $nodeUri     = $node.GetValue('NodeUri')
                                        $expectedVal = $node.GetValue('ExpectedValue')
                                        if ($nodeUri) {
                                            $shortVal = if ($expectedVal -and $expectedVal.Length -gt 80) { $expectedVal.Substring(0,80) + '...' } else { $expectedVal }
                                            [void]$sb.AppendLine("  $nodeUri = $shortVal")
                                        }
                                    }
                                    [void]$sb.AppendLine()
                                }
                            }
                        }
                    } catch {
                        [void]$sb.AppendLine("Error reading node cache: $($_.Exception.Message)")
                    }
                }
                [void]$sb.AppendLine("=== Total: $totalNodes cached nodes ===")
                return @{ Text = $sb.ToString(); Count = $totalNodes }
            }
            Get-NodeCacheData
        } -OnComplete {
            param($result)
            $count = $result.Count
            if ($ui.NodeCacheStatus) { $ui.NodeCacheStatus.Text = "$count cached nodes found" }
            Show-Toast 'Node Cache' "$count cached nodes loaded" 'info'
            Write-DebugLog "Node cache loaded: $count nodes" -Level INFO
        }.GetNewClosure() -Variables @{} -Context @{ Name = 'NodeCache' }
    }.GetNewClosure())
}
# Wire Node Cache (panel)
if ($ui.BtnLoadNodeCache2) {
    $ui.BtnLoadNodeCache2.Add_Click({
        Write-DebugLog 'Loading MDM node cache (panel)' -Level DEBUG
        $ui.NodeCacheStatus2.Text = 'Loading node cache...'
        Start-BackgroundWork -Work {
            function Get-NodeCacheData {
                $sb = [System.Text.StringBuilder]::new(8192)
                $totalNodes = 0
                $basePaths = @('HKLM:\SOFTWARE\Microsoft\Provisioning\NodeCache')
                foreach ($basePath in $basePaths) {
                    if (-not (Test-Path $basePath)) {
                        [void]$sb.AppendLine("Registry path not found: $basePath")
                        continue
                    }
                    [void]$sb.AppendLine("=== MDM Node Cache ===")
                    [void]$sb.AppendLine("Path: $basePath")
                    [void]$sb.AppendLine()
                    try {
                        $cspFolders = Get-ChildItem -Path $basePath -ErrorAction Stop
                        foreach ($csp in $cspFolders) {
                            $cspName = $csp.PSChildName
                            $devicePath = Join-Path $csp.PSPath 'Device'
                            $userPath   = Join-Path $csp.PSPath 'User'
                            foreach ($scope in @(@{Name='Device';Path=$devicePath}, @{Name='User';Path=$userPath})) {
                                if (-not (Test-Path $scope.Path)) { continue }
                                $serverPaths = Get-ChildItem -Path $scope.Path -ErrorAction SilentlyContinue
                                foreach ($server in $serverPaths) {
                                    $nodesPath = Join-Path $server.PSPath 'Nodes'
                                    if (-not (Test-Path $nodesPath)) { continue }
                                    [void]$sb.AppendLine("--- CSP: $cspName | Scope: $($scope.Name) | Server: $($server.PSChildName) ---")
                                    $nodes = Get-ChildItem -Path $nodesPath -ErrorAction SilentlyContinue
                                    foreach ($node in $nodes) {
                                        $totalNodes++
                                        $nodeUri     = $node.GetValue('NodeUri')
                                        $expectedVal = $node.GetValue('ExpectedValue')
                                        if ($nodeUri) {
                                            $shortVal = if ($expectedVal -and $expectedVal.Length -gt 80) { $expectedVal.Substring(0,80) + '...' } else { $expectedVal }
                                            [void]$sb.AppendLine("  $nodeUri = $shortVal")
                                        }
                                    }
                                    [void]$sb.AppendLine()
                                }
                            }
                        }
                    } catch {
                        [void]$sb.AppendLine("Error reading node cache: $($_.Exception.Message)")
                    }
                }
                [void]$sb.AppendLine("=== Total: $totalNodes cached nodes ===")
                return @{ Text = $sb.ToString(); Count = $totalNodes }
            }
            Get-NodeCacheData
        } -OnComplete {
            param($result)
            $text = $result.Text
            $count = $result.Count
            if ($ui.NodeCacheResultBox) { $ui.NodeCacheResultBox.Text = $text }
            if ($ui.NodeCacheScroller) { $ui.NodeCacheScroller.ScrollToTop() }
            if ($ui.NodeCacheStatus2) { $ui.NodeCacheStatus2.Text = "$count cached nodes found" }
            Show-Toast 'Node Cache' "$count cached nodes loaded" 'info'
            Write-DebugLog "Node cache loaded: $count nodes" -Level INFO
        }.GetNewClosure() -Variables @{} -Context @{ Name = 'NodeCache2' }
    }.GetNewClosure())
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION: BACKGROUND LOG EXPORT
# ═══════════════════════════════════════════════════════════════════════════════

$Script:BgLogActive = $false
$Script:BgLogPath   = $null
$Script:BgLogStream = $null

function Toggle-BackgroundLog {
    if ($Script:BgLogActive) {
        # Stop logging
        if ($Script:BgLogStream) {
            try { $Script:BgLogStream.Flush(); $Script:BgLogStream.Close(); $Script:BgLogStream.Dispose() } catch {}
            $Script:BgLogStream = $null
        }
        $Script:BgLogActive = $false
        $msg = "Background logging stopped. File: $($Script:BgLogPath)"
        if ($ui.BgLogStatus)  { $ui.BgLogStatus.Text  = $msg }
        if ($ui.BgLogStatus2) { $ui.BgLogStatus2.Text = $msg }
        if ($ui.BtnBgLogToggle)  { $ui.BtnBgLogToggle.Content.Children[0].Text = [char]0xE896; $ui.BtnBgLogToggle.Content.Children[1].Text = 'Enable Background Logging' }
        if ($ui.BtnBgLogToggle2) { $ui.BtnBgLogToggle2.Content.Children[0].Text = [char]0xE896; $ui.BtnBgLogToggle2.Content.Children[1].Text = 'Enable Background Logging' }
        Show-Toast 'Background Log' "Logging stopped. Saved to $($Script:BgLogPath)" 'info'
        Write-DebugLog "Background logging stopped: $($Script:BgLogPath)" -Level INFO
    } else {
        # Start logging
        $desktop = [Environment]::GetFolderPath('Desktop')
        $timestamp = Get-Date -Format 'MM-dd-yy_H-mm-ss'
        $Script:BgLogPath = Join-Path $desktop "PolicyPilot-Log-$env:COMPUTERNAME-$timestamp.log"
        try {
            $Script:BgLogStream = [System.IO.StreamWriter]::new($Script:BgLogPath, $true, [System.Text.Encoding]::UTF8)
            $Script:BgLogStream.AutoFlush = $true
            $Script:BgLogStream.WriteLine("=== PolicyPilot Background Log === $(Get-Date) === $env:COMPUTERNAME ===")
            $Script:BgLogActive = $true
            $msg = "Logging to: $($Script:BgLogPath)"
            if ($ui.BgLogStatus)  { $ui.BgLogStatus.Text  = $msg }
            if ($ui.BgLogStatus2) { $ui.BgLogStatus2.Text = $msg }
            if ($ui.BtnBgLogToggle)  { $ui.BtnBgLogToggle.Content.Children[0].Text = [char]0xE71A; $ui.BtnBgLogToggle.Content.Children[1].Text = 'Stop Background Logging' }
            if ($ui.BtnBgLogToggle2) { $ui.BtnBgLogToggle2.Content.Children[0].Text = [char]0xE71A; $ui.BtnBgLogToggle2.Content.Children[1].Text = 'Stop Background Logging' }
            Show-Toast 'Background Log' "Logging started: $($Script:BgLogPath)" 'info'
            Write-DebugLog "Background logging started: $($Script:BgLogPath)" -Level INFO
        } catch {
            $msg = "Failed to create log file: $($_.Exception.Message)"
            if ($ui.BgLogStatus)  { $ui.BgLogStatus.Text  = $msg }
            if ($ui.BgLogStatus2) { $ui.BgLogStatus2.Text = $msg }
            Show-Toast 'Background Log' $msg 'error'
        }
    }
}

# Wire Background Log toggle (sidebar + panel)
if ($ui.BtnBgLogToggle)  { $ui.BtnBgLogToggle.Add_Click({ Toggle-BackgroundLog }.GetNewClosure()) }
if ($ui.BtnBgLogToggle2) { $ui.BtnBgLogToggle2.Add_Click({ Toggle-BackgroundLog }.GetNewClosure()) }



# ── IME Search ──
function Search-ImeHighlight {
    param([string]$Term)
    $Script:ImeSearchMatches.Clear(); $Script:ImeSearchIndex = -1
    if ([string]::IsNullOrWhiteSpace($Term)) {
        if ($ui.ImeSearchCount) { $ui.ImeSearchCount.Text = '' }
        Apply-LogListBoxColors $ui.lbImeLogs $Script:ImeLineTypes $Script:ImeColorMap $Script:ImeBrushCache
        return
    }
    $rx = [regex]::new([regex]::Escape($Term), 'IgnoreCase')
    for ($i = 0; $i -lt $Script:ImeRawEntries.Count; $i++) {
        if ($rx.IsMatch($Script:ImeRawEntries[$i]) -or $rx.IsMatch($Script:ImeFormattedLines[$i])) { [void]$Script:ImeSearchMatches.Add($i) }
    }
    if ($ui.ImeSearchCount) { $ui.ImeSearchCount.Text = "$($Script:ImeSearchMatches.Count) matches" }
    if ($Script:ImeSearchMatches.Count -gt 0) { $Script:ImeSearchIndex = 0; Navigate-ImeSearchMatch 0 }
}

function Navigate-ImeSearchMatch([int]$Index) {
    if ($Index -lt 0 -or $Index -ge $Script:ImeSearchMatches.Count) { return }
    $Script:ImeSearchIndex = $Index
    $lineIdx = $Script:ImeSearchMatches[$Index]
    $ui.lbImeLogs.ScrollIntoView($ui.lbImeLogs.Items[$lineIdx]); $ui.lbImeLogs.SelectedIndex = $lineIdx
    if ($ui.ImeSearchCount) { $ui.ImeSearchCount.Text = "$($Index+1)/$($Script:ImeSearchMatches.Count)" }
}

# ── Wire up IME Log UI events ──
if ($ui.BtnImeWrapToggle) {
    $ui.BtnImeWrapToggle.Add_Click({
        $scroll = $ui.lbImeLogs.GetValue([System.Windows.Controls.ScrollViewer]::HorizontalScrollBarVisibilityProperty)
        if ($scroll -eq 'Disabled') { $ui.lbImeLogs.SetValue([System.Windows.Controls.ScrollViewer]::HorizontalScrollBarVisibilityProperty, [System.Windows.Controls.ScrollBarVisibility]::Auto) }
        else { $ui.lbImeLogs.SetValue([System.Windows.Controls.ScrollViewer]::HorizontalScrollBarVisibilityProperty, [System.Windows.Controls.ScrollBarVisibility]::Disabled) }
    })
}
if ($ui.BtnImeClearLog) {
    $ui.BtnImeClearLog.Add_Click({ Clear-ImeLogDisplay })
}
if ($ui.BtnIntuneSync) {
    $ui.BtnIntuneSync.Add_Click({ Invoke-IntuneSync })
}
# Mode toggle: Live tail vs Analysis
if ($ui.ChkImeLiveTail) {
    $ui.ChkImeLiveTail.Add_Checked({
        if ($ui.ImeModeHint) { $ui.ImeModeHint.Text = 'Streaming new log entries in real time' }
        if ($Script:ImeTailing) { return }  # already tailing
        Start-ImeTail
    })
    $ui.ChkImeLiveTail.Add_Unchecked({
        if ($ui.ImeModeHint) { $ui.ImeModeHint.Text = 'Full log loaded for analysis and search' }
        if ($Script:ImeTailing) { Stop-ImeTail }
        Load-ImeLogFile
    })
}
# Log source change: restart current mode
if ($ui.CmbImeLogSource) {
    $ui.CmbImeLogSource.Add_SelectionChanged({
        if ($ui.ChkImeLiveTail -and $ui.ChkImeLiveTail.IsChecked) {
            if ($Script:ImeTailing) { Stop-ImeTail }
            Start-ImeTail
        } else {
            Load-ImeLogFile
        }
    })
}
# Line limit change: update variable and reload if in analysis mode
if ($ui.CmbImeLineLimit) {
    $ui.CmbImeLineLimit.Add_SelectionChanged({
        $idx = $ui.CmbImeLineLimit.SelectedIndex
        $Script:ImeLineLimit = $Script:ImeLineLimitOptions[$idx]
        Write-Host "[IME] Line limit changed to: $($Script:ImeLineLimit) (idx=$idx)"
        # Reload if currently in analysis mode (not tailing)
        if (-not ($ui.ChkImeLiveTail -and $ui.ChkImeLiveTail.IsChecked)) {
            Load-ImeLogFile
        }
    })
}
if ($ui.TxtImeSearch) {
    $Script:ImeSearchDebounce = $null
    $ui.TxtImeSearch.Add_TextChanged({
        if ($Script:ImeSearchDebounce) { $Script:ImeSearchDebounce.Stop() }
        $Script:ImeSearchDebounce = [System.Windows.Threading.DispatcherTimer]::new()
        $Script:ImeSearchDebounce.Interval = [TimeSpan]::FromMilliseconds(300)
        $Script:ImeSearchDebounce.Add_Tick({
            $Script:ImeSearchDebounce.Stop()
            Search-ImeHighlight $ui.TxtImeSearch.Text
        })
        $Script:ImeSearchDebounce.Start()
    })
}
if ($ui.BtnImeSearchNext) {
    $ui.BtnImeSearchNext.Add_Click({
        if ($Script:ImeSearchMatches.Count -gt 0) {
            $next = ($Script:ImeSearchIndex + 1) % $Script:ImeSearchMatches.Count
            Navigate-ImeSearchMatch $next
        }
    })
}
if ($ui.BtnImeSearchPrev) {
    $ui.BtnImeSearchPrev.Add_Click({
        if ($Script:ImeSearchMatches.Count -gt 0) {
            $prev = $Script:ImeSearchIndex - 1
            if ($prev -lt 0) { $prev = $Script:ImeSearchMatches.Count - 1 }
            Navigate-ImeSearchMatch $prev
        }
    })
}

# Severity filter pills
foreach ($filterName in @('FilterImeAll','FilterImeError','FilterImeWarning','FilterImeInfo')) {
    if ($ui[$filterName]) {
        $ui[$filterName].Add_Click({
            param($sender)
            $label = $sender.Content
            $Script:ImeActiveFilter = switch ($label) {
                'All'      { 'All' }
                'Errors'   { 'Error' }
                'Warnings' { 'Warning' }
                'Info'     { 'Info' }
                default    { 'All' }
            }
            # Update pill active states
            foreach ($fn in @('FilterImeAll','FilterImeError','FilterImeWarning','FilterImeInfo')) {
                if ($ui[$fn]) { $ui[$fn].Tag = if ($ui[$fn] -eq $sender) { 'Active' } else { $null } }
            }
            # Re-apply with new filter
            if ($ui.ChkImeLiveTail -and $ui.ChkImeLiveTail.IsChecked) {
                # Live mode: re-read from start with filter
                if ($Script:ImeTailing) {
                    $prevPath = $Script:ImeActiveLogPath
                    Stop-ImeTail
                    $Script:ImeActiveLogPath = $prevPath
                    $Script:ImeLastFilePos = 0
                    $Script:ImeTailing = $true
                    Clear-ImeLogDisplay
                    Read-ImeLogDelta
                    # Show heatmap if we have data
                    if ($ui.cnvImeHeatmap -and $Script:ImeStats.Lines -gt 0) { $ui.cnvImeHeatmap.Visibility = 'Visible' }

                    # Re-create watcher
                    $source = $ui.CmbImeLogSource.SelectedItem.Content
                    $Script:ImeWatcher = [System.IO.FileSystemWatcher]::new()
                    $Script:ImeWatcher.Path = $Script:ImeLogDir
                    $Script:ImeWatcher.Filter = "$source.log"
                    $Script:ImeWatcher.NotifyFilter = [System.IO.NotifyFilters]::LastWrite -bor [System.IO.NotifyFilters]::Size
                    $Script:ImeWatcher.EnableRaisingEvents = $true
                    Register-ObjectEvent -InputObject $Script:ImeWatcher -EventName Changed -Action {
                        $Script:ImeFswChanged = $true
                    } -SourceIdentifier 'ImeLogWatcher'
                }
            } else {
                # Analysis mode: reload with filter
                Load-ImeLogFile
            }
        })
    }
}

# ═══════════════════════════════════════════════════════════════════════════════

# ── IME Content Filter Handlers ──
# Preset keyword map
$Script:ImePresets = @{
    'Win32App'   = 'Win32App'
    'PowerShell' = '/powershell|remediation|proactive|script/i'
    'Policy'     = '/policy|compliance|configuration|assignment/i'
    'Network'    = '/network|connectivity|endpoint|service point|http/i'
    'Check-in'   = '/check.in|sync|heartbeat|enrollment/i'
    'Certs'      = '/certificate|SCEP|cert enroll|PKCS/i'
    'Timeout'    = '/timeout|retry|retrying|timed out/i'
}

function Apply-ImeContentFilter([string]$filterText) {
    $Script:ImeContentFilter = $filterText
    if ($ui.TxtImeContentFilter) { $ui.TxtImeContentFilter.Text = $filterText }
    # Reload log with new filter
    if ($Script:ImeTailing) {
        $prevPath = $Script:ImeActiveLogPath
        Stop-ImeTail
        $Script:ImeActiveLogPath = $prevPath
        $Script:ImeLastFilePos = 0
        $Script:ImeTailing = $true
        Clear-ImeLogDisplay
        Read-ImeLogDelta
        $source = $ui.CmbImeLogSource.SelectedItem.Content
        $Script:ImeWatcher = [System.IO.FileSystemWatcher]::new()
        $Script:ImeWatcher.Path = $Script:ImeLogDir
        $Script:ImeWatcher.Filter = "$source.log"
        $Script:ImeWatcher.NotifyFilter = [System.IO.NotifyFilters]::LastWrite -bor [System.IO.NotifyFilters]::Size
        $Script:ImeWatcher.EnableRaisingEvents = $true
        Register-ObjectEvent -InputObject $Script:ImeWatcher -EventName Changed -Action { $Script:ImeFswChanged = $true } -SourceIdentifier 'ImeLogWatcher'
    } else {
        Load-ImeLogFile
    }
}

# Wire preset buttons
foreach ($presetName in @('PresetImeNone','PresetImeWin32','PresetImePowershell','PresetImePolicy','PresetImeNetwork','PresetImeCheckin','PresetImeCert','PresetImeTimeout')) {
    if ($ui[$presetName]) {
        $ui[$presetName].Add_Click({
            param($sender)
            $label = $sender.Content
            # Update preset pill states
            foreach ($pn in @('PresetImeNone','PresetImeWin32','PresetImePowershell','PresetImePolicy','PresetImeNetwork','PresetImeCheckin','PresetImeCert','PresetImeTimeout')) {
                if ($ui[$pn]) { $ui[$pn].Tag = if ($ui[$pn] -eq $sender) { 'Active' } else { $null } }
            }
            if ($label -eq 'None') { Apply-ImeContentFilter '' } else { Apply-ImeContentFilter ($Script:ImePresets[$label]) }
        })
    }
}

# Wire TextBox
if ($ui.TxtImeContentFilter) {
    $Script:ImeContentFilterTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $Script:ImeContentFilterTimer.Interval = [TimeSpan]::FromMilliseconds(500)
    $Script:ImeContentFilterTimer.Add_Tick({
        $Script:ImeContentFilterTimer.Stop()
        Apply-ImeContentFilter $ui.TxtImeContentFilter.Text
    })
    $ui.TxtImeContentFilter.Add_TextChanged({
        $Script:ImeContentFilterTimer.Stop()
        $Script:ImeContentFilterTimer.Start()
        # Clear preset pill active states when user types
        foreach ($pn in @('PresetImeNone','PresetImeWin32','PresetImePowershell','PresetImePolicy','PresetImeNetwork','PresetImeCheckin','PresetImeCert','PresetImeTimeout')) {
            if ($ui[$pn]) { $ui[$pn].Tag = $null }
        }
        if (-not $ui.TxtImeContentFilter.Text) {
            if ($ui.PresetImeNone) { $ui.PresetImeNone.Tag = 'Active' }
        }
    })
}

# Wire save filter button
if ($ui.BtnImeSaveFilter) {
    $ui.BtnImeSaveFilter.Add_Click({
        $filterText = $ui.TxtImeContentFilter.Text
        if (-not $filterText) { return }
        $name = Show-ThemedInputBox -Prompt 'Name for this filter:' -Title 'Save Filter' -DefaultValue $filterText
        if ($name) {
            $Script:ImeSavedFilters[$name] = $filterText
            [void]$ui.CmbImeSavedFilters.Items.Add($name)
            $ui.CmbImeSavedFilters.SelectedItem = $name
        }
    })
}

if ($ui.CmbImeSavedFilters) {
    $ui.CmbImeSavedFilters.Add_SelectionChanged({
        $sel = $ui.CmbImeSavedFilters.SelectedItem
        if ($sel -and $Script:ImeSavedFilters.ContainsKey($sel)) {
            Apply-ImeContentFilter $Script:ImeSavedFilters[$sel]
        }
    })
}
# SECTION 19.5: GPO LOG VIEWER - LIVE TAIL + DEBUG LOGGING + GPUPDATE
# ═══════════════════════════════════════════════════════════════════════════════

# ── GPO Log Parsing Engine ──
$Script:GpoRxError    = [regex]::new('(?i)(error|fail(ed|ure)?|denied|not found|could not|unable to|HRESULT|exception|0x8\w{7}|Access is denied|no.*domain controller|unreachable)', 'Compiled')
$Script:GpoRxWarning  = [regex]::new('(?i)(warn(ing)?|timeout|retry|retrying|slow link|loopback|no changes|skipped|disabled|not applied|blocked|WMI filter.*false|security filter.*denied)', 'Compiled')
$Script:GpoRxSuccess  = [regex]::new('(?i)(success(fully)?|completed|applied|processed|linked|no errors|exit code[:=]\s*0\b)', 'Compiled')
$Script:GpoRxCSE      = [regex]::new('(?i)(client.side extension|CSE|registry settings|security settings|folder redirection|software installation|scripts extension|administrative templates|drive maps|printers|preferences|internet explorer|firewall|wireless|wired|public key)', 'Compiled')
$Script:GpoRxPolicy   = [regex]::new('(?i)(Group Policy|GPO|LGPO|policy (processing|application)|linked.*GPO.*list|filtering|WMI filter|security filter|scope of management)', 'Compiled')
$Script:GpoRxNetwork  = [regex]::new('(?i)(domain controller|DC[:=]|LDAP|site[:=]|network|bandwidth|slow link|NLA|DNS|connectivity)', 'Compiled')
$Script:GpoRxSync     = [regex]::new('(?i)(gpupdate|background.*processing|foreground.*processing|manual.*refresh|user.*policy|computer.*policy|processing mode|loopback)', 'Compiled')

# gpsvc.log line regex: "GPSVC(PID.TID) HH:MM:SS:mmm FunctionName: message"
$Script:GpoRxGpsvcLine = [regex]::new('^GPSVC\((?<pid>[0-9a-fA-F]+)\.(?<tid>[0-9a-fA-F]+)\)\s+(?<time>\d{2}:\d{2}:\d{2}:\d{3})\s+(?<func>\w+):\s+(?<msg>.*)', 'Compiled')

function Get-GpoLineColor {
    param([string]$Line, [int]$EventLevel)
    # Event level: 1=Critical, 2=Error, 3=Warning, 4=Information, 5=Verbose
    if ($EventLevel -eq 1 -or $EventLevel -eq 2) { return @{ Color = '#D13438'; Bold = $true;  Cat = 'Error'   } }
    if ($EventLevel -eq 3) { return @{ Color = '#FFB900'; Bold = $false; Cat = 'Warning' } }
    # Content-based classification
    if ($Script:GpoRxError.IsMatch($Line))   { return @{ Color = '#D13438'; Bold = $true;  Cat = 'Error'   } }
    if ($Script:GpoRxWarning.IsMatch($Line)) { return @{ Color = '#FFB900'; Bold = $false; Cat = 'Warning' } }
    if ($Script:GpoRxSuccess.IsMatch($Line)) { return @{ Color = '#107C10'; Bold = $true;  Cat = 'Success' } }
    if ($Script:GpoRxCSE.IsMatch($Line))     { return @{ Color = '#8764B8'; Bold = $false; Cat = 'CSE'     } }
    if ($Script:GpoRxNetwork.IsMatch($Line)) { return @{ Color = '#FF8C00'; Bold = $false; Cat = 'Network' } }
    if ($Script:GpoRxSync.IsMatch($Line))    { return @{ Color = '#60CDFF'; Bold = $false; Cat = 'Sync'    } }
    if ($Script:GpoRxPolicy.IsMatch($Line))  { return @{ Color = '#0078D4'; Bold = $false; Cat = 'Policy'  } }
    return @{ Color = '#C0C0C0'; Bold = $false; Cat = 'Info' }
}

function Get-GpoSeverityBadge {
    param([string]$Cat)
    switch ($Cat) {
        'Error'   { return @{ Label = 'ERR';  Bg = '#D13438'; Fg = '#FFFFFF' } }
        'Warning' { return @{ Label = 'WARN'; Bg = '#FFB900'; Fg = '#1A1A1A' } }
        'Success' { return @{ Label = 'OK';   Bg = '#107C10'; Fg = '#FFFFFF' } }
        'CSE'     { return @{ Label = 'CSE';  Bg = '#8764B8'; Fg = '#FFFFFF' } }
        'Network' { return @{ Label = 'NET';  Bg = '#FF8C00'; Fg = '#FFFFFF' } }
        'Sync'    { return @{ Label = 'SYNC'; Bg = '#60CDFF'; Fg = '#1A1A1A' } }
        'Policy'  { return @{ Label = 'POL';  Bg = '#0078D4'; Fg = '#FFFFFF' } }
        default   { return @{ Label = 'INFO'; Bg = '#333333'; Fg = '#AAAAAA' } }
    }
}

# â”€â”€ GPO Log State â”€â”€
$Script:GpoLogQueue       = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new())
$Script:GpoTailing        = $false
$Script:GpoWatcher        = $null
$Script:GpoMaxPerTick     = 150
$Script:GpoLastFilePos    = 0
$Script:GpoActiveLogPath  = ''
$Script:GpoActiveFilter   = 'All'
$Script:GpoContentFilter   = ''
$Script:GpoSavedFilters    = @{}
$Script:GpoActiveSource   = 'EventLog'
$Script:GpoLastEventTime  = $null
$Script:GpoStats = @{ Lines = 0; Errors = 0; Warnings = 0 }
$Script:GpoSearchMatches  = [System.Collections.Generic.List[int]]::new()
$Script:GpoSearchIndex    = -1
$Script:GpoBrushCache     = @{}
$Script:GpoDebugLogPath   = "$env:WINDIR\debug\usermode\gpsvc.log"
$Script:GpoDebugRegPath   = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Diagnostics'
$Script:GpoFormattedLines  = [System.Collections.Generic.List[string]]::new()
$Script:GpoLineTypes       = [System.Collections.Generic.List[byte]]::new()
$Script:GpoRawEntries      = [System.Collections.Generic.List[string]]::new()
$Script:GpoColorMap = @( @{ Fg = '#C0C0C0'; Bold = $false }, @{ Fg = '#D13438'; Bold = $true }, @{ Fg = '#FFB900'; Bold = $false }, @{ Fg = '#107C10'; Bold = $true }, @{ Fg = '#8764B8'; Bold = $false }, @{ Fg = '#FF8C00'; Bold = $false }, @{ Fg = '#60CDFF'; Bold = $false }, @{ Fg = '#0078D4'; Bold = $false } )
$Script:GpoBadgeMap = @('INFO','ERR','WARN','OK','CSE','NET','SYNC','POL')
function Update-GpoStatsDisplay {
    if ($ui.GpoStatLines)    { $ui.GpoStatLines.Text    = $Script:GpoStats.Lines }
    if ($ui.GpoStatErrors)   { $ui.GpoStatErrors.Text   = $Script:GpoStats.Errors }
    if ($ui.GpoStatWarnings) { $ui.GpoStatWarnings.Text = $Script:GpoStats.Warnings }
}
function Clear-GpoLogDisplay {
    $Script:GpoFormattedLines.Clear(); $Script:GpoLineTypes.Clear(); $Script:GpoRawEntries.Clear()
    $ui.lbGpoLogs.ItemsSource = $null
    $Script:GpoStats.Lines = 0; $Script:GpoStats.Errors = 0; $Script:GpoStats.Warnings = 0
    Update-GpoStatsDisplay
    if ($ui.cnvGpoHeatmap) { $ui.cnvGpoHeatmap.Children.Clear() }; if ($ui.cnvGpoMinimap) { $ui.cnvGpoMinimap.Children.Clear() }
    $Script:GpoSearchMatches.Clear(); $Script:GpoSearchIndex = -1
    if ($ui.GpoSearchCount) { $ui.GpoSearchCount.Text = '' }
}
function Classify-GpoLogEntry {
    param([hashtable]$Parsed)
    $msg   = $Parsed.Message
    $rules = Get-GpoLineColor -Line $msg -EventLevel $Parsed.Level
    if ($Script:GpoActiveFilter -ne 'All') {
        $cat = $rules.Cat
        $pass = switch ($Script:GpoActiveFilter) { 'Error' { $cat -eq 'Error' }; 'Warning' { $cat -eq 'Warning' }; 'Info' { $cat -notin @('Error','Warning') }; default { $true } }
        if (-not $pass) { return $null }
    }
    if ($Script:GpoContentFilter -and $Script:GpoContentFilter.Length -gt 0) {
        $pattern = $Script:GpoContentFilter
        $textToSearch = "$($Parsed.Message) $($Parsed.Component) $($Parsed.EventId)"
        if ($pattern.StartsWith('/') -and $pattern.EndsWith('/')) {
            try { if ($textToSearch -notmatch $pattern.Substring(1, $pattern.Length - 2)) { return $null } } catch { if ($textToSearch -notlike "*$pattern*") { return $null } }
        } else { if ($textToSearch -notlike "*$pattern*") { return $null } }
    }
    $Script:GpoStats.Lines++; if ($rules.Cat -eq 'Error') { $Script:GpoStats.Errors++ }; if ($rules.Cat -eq 'Warning') { $Script:GpoStats.Warnings++ }
    $lineType = switch ($rules.Cat) { 'Error' { [byte]1 }; 'Warning' { [byte]2 }; 'Success' { [byte]3 }; 'CSE' { [byte]4 }; 'Network' { [byte]5 }; 'Sync' { [byte]6 }; 'Policy' { [byte]7 }; default { [byte]0 } }
    $badge = $Script:GpoBadgeMap[$lineType]
    $sb = [System.Text.StringBuilder]::new(256)
    [void]$sb.Append("[$badge] ")
    if ($Parsed.Time)      { [void]$sb.Append("$($Parsed.Time) ") }
    if ($Parsed.Component) { [void]$sb.Append("[$($Parsed.Component)] ") }
    if ($Parsed.EventId)   { [void]$sb.Append("ID:$($Parsed.EventId) ") }
    [void]$sb.Append($msg)
    if ($Parsed.Thread) { [void]$sb.Append("  T:$($Parsed.Thread)") }
    return @{ Formatted = $sb.ToString(); LineType = $lineType; Raw = $msg }
}

# ── Event Log Reader ──
function Read-GpoEventLogDelta {
    Write-Verbose "[GPO] Read-GpoEventLogDelta called - lastEventTime=$($Script:GpoLastEventTime)"
    try {
        $filter = @{ LogName = 'Microsoft-Windows-GroupPolicy/Operational' }
        if ($Script:GpoLastEventTime) {
            $filter['StartTime'] = $Script:GpoLastEventTime.AddMilliseconds(1)
        }
        $events = Get-WinEvent -FilterHashtable $filter -MaxEvents 500 -ErrorAction Stop |
            Sort-Object TimeCreated
        Write-Verbose "[GPO] EventLog: $($events.Count) events found"
        foreach ($ev in $events) {
            $parsed = @{
                Message   = ($ev.Message -replace '\r?\n', ' ').Trim()
                Time      = $ev.TimeCreated.ToString('M-d-yyyy HH:mm:ss')
                Component = $ev.ProviderName -replace 'Microsoft-Windows-',''
                EventId   = "$($ev.Id)"
                Level     = [int]$ev.Level
                Thread    = "$($ev.ProcessId)"
            }
            $Script:GpoLogQueue.Enqueue($parsed)
            $Script:GpoLastEventTime = $ev.TimeCreated
        }
    } catch {
        if ($_.Exception.Message -match 'unauthorized|access') {
            Write-Host "[GPO] EventLog: ACCESS DENIED - needs admin"
            $Script:GpoLogQueue.Enqueue(@{
                Message = "Access denied - reading Group Policy event log requires administrator privileges. Run PolicyPilot as admin."
                Time = (Get-Date).ToString('M-d-yyyy HH:mm:ss'); Component = 'System'; EventId = ''; Level = 2; Thread = ''
            })
            $Script:GpoTailing = $false
        } elseif ($_.Exception.Message -match 'No events were found') {
            Write-Verbose "[GPO] EventLog: No events found (this is normal if no GP activity yet)"
        } else {
            Write-Host "[GPO] EventLog ERROR: $($_.Exception.Message)"
            Write-DebugLog "GPO event log read error: $_" -Level WARN
        }
    }
}

# ── gpsvc.log File Reader ──
function Read-GpoDebugLogDelta {
    if (-not (Test-Path $Script:GpoDebugLogPath)) { return }
    try {
        $fs = [System.IO.FileStream]::new($Script:GpoDebugLogPath, 'Open', 'Read', 'ReadWrite,Delete')
        $fs.Seek($Script:GpoLastFilePos, 'Begin')
        $sr = [System.IO.StreamReader]::new($fs)
        while ($null -ne ($line = $sr.ReadLine())) {
            if ($line.Trim()) {
                $m = $Script:GpoRxGpsvcLine.Match($line)
                if ($m.Success) {
                    $parsed = @{
                        Message   = $m.Groups['msg'].Value
                        Time      = $m.Groups['time'].Value
                        Component = $m.Groups['func'].Value
                        EventId   = ''
                        Level     = 4
                        Thread    = $m.Groups['tid'].Value
                    }
                } else {
                    $parsed = @{
                        Message = $line; Time = ''; Component = ''; EventId = ''; Level = 4; Thread = ''
                    }
                }
                $Script:GpoLogQueue.Enqueue($parsed)
            }
        }
        $Script:GpoLastFilePos = $fs.Position
        $sr.Close()
        $fs.Close()
    } catch {
        Write-DebugLog "GPO debug log read error: $_" -Level WARN
    }
}

# ── System Event Log Reader (no admin required) ──
$Script:GpoSystemLastTime = $null
function Read-GpoSystemLogDelta {
    Write-Verbose "[GPO] Read-GpoSystemLogDelta called - lastSystemTime=$($Script:GpoSystemLastTime)"
    try {
        $filter = @{ LogName = 'System'; ProviderName = @('Microsoft-Windows-GroupPolicy','GroupPolicy') }
        if ($Script:GpoSystemLastTime) {
            $filter['StartTime'] = $Script:GpoSystemLastTime.AddMilliseconds(1)
        }
        $events = Get-WinEvent -FilterHashtable $filter -MaxEvents 500 -ErrorAction Stop |
            Sort-Object TimeCreated
        Write-Verbose "[GPO] SystemLog: $($events.Count) events found"
        foreach ($ev in $events) {
            $parsed = @{
                Message   = ($ev.Message -replace '\r?\n', ' ').Trim()
                Time      = $ev.TimeCreated.ToString('M-d-yyyy HH:mm:ss')
                Component = 'GroupPolicy'
                EventId   = "$($ev.Id)"
                Level     = [int]$ev.Level
                Thread    = "$($ev.ProcessId)"
            }
            $Script:GpoLogQueue.Enqueue($parsed)
            $Script:GpoSystemLastTime = $ev.TimeCreated
        }
    } catch {
        if ($_.Exception.Message -match 'No events were found') {
            Write-Verbose "[GPO] SystemLog: No events found (this is normal if no GP activity yet)"
        } else {
            Write-Host "[GPO] SystemLog ERROR: $($_.Exception.Message)"
            Write-DebugLog "GPO system log read error: $_" -Level WARN
        }
    }
}

# ── GP Preferences Trace Log Reader ──
$Script:GppTraceDir   = "$env:ProgramData\GroupPolicy\Preference\Trace"
$Script:GppLastFilePos = 0
$Script:GppActiveLogPath = ''
$Script:GppWatcher     = $null
$Script:GppFswChanged  = $false

function Read-GppTraceLogDelta {
    if (-not $Script:GppActiveLogPath -or -not (Test-Path $Script:GppActiveLogPath)) { return }
    try {
        $fs = [System.IO.FileStream]::new($Script:GppActiveLogPath, 'Open', 'Read', 'ReadWrite,Delete')
        $fs.Seek($Script:GppLastFilePos, 'Begin')
        $sr = [System.IO.StreamReader]::new($fs)
        while ($null -ne ($line = $sr.ReadLine())) {
            if ($line.Trim()) {
                $ts = ''; $comp = 'GPPrefs'
                # GPP trace lines often start with timestamp like "2026-02-27 10:30:45.123"
                if ($line -match '^(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})') {
                    $ts = ($Matches[1] -split '\s+')[-1]
                }
                $lvl = 4  # default: informational
                if ($line -match '(?i)\berror\b|fail(ed|ure)?|exception|HRESULT|0x8') { $lvl = 2 }
                elseif ($line -match '(?i)\bwarn(ing)?\b') { $lvl = 3 }
                $Script:GpoLogQueue.Enqueue(@{
                    Message = $line; Time = $ts; Component = $comp; EventId = ''; Level = $lvl; Thread = ''
                })
            }
        }
        $Script:GppLastFilePos = $fs.Position
        $sr.Close()
        $fs.Close()
    } catch {
        Write-DebugLog "GPP trace log read error: $_" -Level WARN
    }
}

function Get-GppTraceFiles {
    $files = @()
    if (Test-Path $Script:GppTraceDir) {
        $files = Get-ChildItem -Path $Script:GppTraceDir -Filter '*.log' -File -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending
    }
    return $files
}

function Start-GpoTail {
    $sourceItem = $ui.CmbGpoLogSource.SelectedItem.Content
    Write-Host "[GPO] Start-GpoTail called - source='$sourceItem'"

    Clear-GpoLogDisplay
    $Script:GpoTailing = $true
    $Script:GpoLastEventTime = $null
    $Script:GpoSystemLastTime = $null
    $Script:GpoLastFilePos = 0
    $Script:GppLastFilePos = 0

    # Cleanup any previous watchers
    if ($Script:GpoWatcher) { $Script:GpoWatcher.EnableRaisingEvents = $false; $Script:GpoWatcher.Dispose(); $Script:GpoWatcher = $null }
    if ($Script:GppWatcher) { $Script:GppWatcher.EnableRaisingEvents = $false; $Script:GppWatcher.Dispose(); $Script:GppWatcher = $null }
    Unregister-Event -SourceIdentifier 'GpoLogWatcher' -ErrorAction SilentlyContinue
    Unregister-Event -SourceIdentifier 'GppLogWatcher' -ErrorAction SilentlyContinue

    if ($sourceItem -match 'Operational') {
        $Script:GpoActiveSource = 'EventLog'
        $ui.GpoLogSubtitle.Text = "Tailing: Microsoft-Windows-GroupPolicy/Operational"
        Read-GpoEventLogDelta
    } elseif ($sourceItem -match 'System Log') {
        $Script:GpoActiveSource = 'SystemLog'
        $ui.GpoLogSubtitle.Text = "Tailing: System Log (GroupPolicy events)"
        Read-GpoSystemLogDelta
    } elseif ($sourceItem -match 'Preferences') {
        $Script:GpoActiveSource = 'PrefsTrace'
        $traceFiles = Get-GppTraceFiles
        if ($traceFiles.Count -eq 0) {
            Show-Toast 'GP Prefs Trace' "No trace files found in $Script:GppTraceDir`nEnable via gpedit.msc > Computer Configuration > Admin Templates > System > Group Policy > Logging and Tracing" -Type Warning
            $Script:GpoTailing = $false
            return
        }
        # Use the most recently written trace file
        $Script:GppActiveLogPath = $traceFiles[0].FullName
        $ui.GpoLogSubtitle.Text = "Tailing: $($traceFiles[0].Name) (GP Preferences)"
        Read-GppTraceLogDelta

        # FSW for the trace directory
        $Script:GppFswChanged = $false
        $Script:GppWatcher = [System.IO.FileSystemWatcher]::new()
        $Script:GppWatcher.Path = $Script:GppTraceDir
        $Script:GppWatcher.Filter = '*.log'
        $Script:GppWatcher.NotifyFilter = [System.IO.NotifyFilters]::LastWrite -bor [System.IO.NotifyFilters]::Size
        $Script:GppWatcher.EnableRaisingEvents = $true
        Register-ObjectEvent -InputObject $Script:GppWatcher -EventName Changed -Action {
            $Script:GppFswChanged = $true
        } -SourceIdentifier 'GppLogWatcher'
    } elseif ($sourceItem -match 'Combined|All Sources') {
        $Script:GpoActiveSource = 'Combined'
        $ui.GpoLogSubtitle.Text = "Tailing: All Sources (Combined)"
        # Read all available sources
        Read-GpoSystemLogDelta  # always available
        try { Read-GpoEventLogDelta } catch { try { Write-DebugLog "Unhandled: $_" -Level ERROR } catch {} }  # may fail without admin
        if (Test-Path $Script:GpoDebugLogPath) { Read-GpoDebugLogDelta }
        $traceFiles = Get-GppTraceFiles
        if ($traceFiles.Count -gt 0) {
            $Script:GppActiveLogPath = $traceFiles[0].FullName
            Read-GppTraceLogDelta
        }
        # FSW for debug log
        if (Test-Path $Script:GpoDebugLogPath) {
            $Script:GpoFswChanged = $false
            $Script:GpoWatcher = [System.IO.FileSystemWatcher]::new()
            $Script:GpoWatcher.Path = [System.IO.Path]::GetDirectoryName($Script:GpoDebugLogPath)
            $Script:GpoWatcher.Filter = 'gpsvc.log'
            $Script:GpoWatcher.NotifyFilter = [System.IO.NotifyFilters]::LastWrite -bor [System.IO.NotifyFilters]::Size
            $Script:GpoWatcher.EnableRaisingEvents = $true
            Register-ObjectEvent -InputObject $Script:GpoWatcher -EventName Changed -Action {
                $Script:GpoFswChanged = $true
            } -SourceIdentifier 'GpoLogWatcher'
        }
        # FSW for GP Prefs trace
        if (Test-Path $Script:GppTraceDir) {
            $Script:GppFswChanged = $false
            $Script:GppWatcher = [System.IO.FileSystemWatcher]::new()
            $Script:GppWatcher.Path = $Script:GppTraceDir
            $Script:GppWatcher.Filter = '*.log'
            $Script:GppWatcher.NotifyFilter = [System.IO.NotifyFilters]::LastWrite -bor [System.IO.NotifyFilters]::Size
            $Script:GppWatcher.EnableRaisingEvents = $true
            Register-ObjectEvent -InputObject $Script:GppWatcher -EventName Changed -Action {
                $Script:GppFswChanged = $true
            } -SourceIdentifier 'GppLogWatcher'
        }
    } else {
        # gpsvc.log (Debug)
        $Script:GpoActiveSource = 'DebugLog'
        if (-not (Test-Path $Script:GpoDebugLogPath)) {
            Show-Toast 'Debug Log Not Found' "No gpsvc.log found at $Script:GpoDebugLogPath.`nEnable debug logging first." -Type Warning
            $Script:GpoTailing = $false
            return
        }
        $ui.GpoLogSubtitle.Text = "Tailing: gpsvc.log (Debug)"
        $Script:GpoActiveLogPath = $Script:GpoDebugLogPath
        Read-GpoDebugLogDelta

        # FileSystemWatcher for gpsvc.log
        $Script:GpoFswChanged = $false
        $Script:GpoWatcher = [System.IO.FileSystemWatcher]::new()
        $Script:GpoWatcher.Path = [System.IO.Path]::GetDirectoryName($Script:GpoDebugLogPath)
        $Script:GpoWatcher.Filter = 'gpsvc.log'
        $Script:GpoWatcher.NotifyFilter = [System.IO.NotifyFilters]::LastWrite -bor [System.IO.NotifyFilters]::Size
        $Script:GpoWatcher.EnableRaisingEvents = $true
        Register-ObjectEvent -InputObject $Script:GpoWatcher -EventName Changed -Action {
            $Script:GpoFswChanged = $true
        } -SourceIdentifier 'GpoLogWatcher'
    }

    Write-Host "[GPO] Tail started - source=$($Script:GpoActiveSource), tailing=$Script:GpoTailing, queueSize=$($Script:GpoLogQueue.Count)"
    Write-DebugLog "GPO tail started: $sourceItem" -Level STEP
    Show-Toast 'GPO Log Tail' "Now tailing $sourceItem" -Type Info
}

function Stop-GpoTail {
    $Script:GpoTailing = $false
    if ($Script:GpoWatcher) {
        $Script:GpoWatcher.EnableRaisingEvents = $false
        $Script:GpoWatcher.Dispose()
        $Script:GpoWatcher = $null
    }
    if ($Script:GppWatcher) {
        $Script:GppWatcher.EnableRaisingEvents = $false
        $Script:GppWatcher.Dispose()
        $Script:GppWatcher = $null
    }
    Unregister-Event -SourceIdentifier 'GpoLogWatcher' -ErrorAction SilentlyContinue
    Unregister-Event -SourceIdentifier 'GppLogWatcher' -ErrorAction SilentlyContinue
    $ui.GpoLogSubtitle.Text = 'Group Policy - Tail stopped'
    if ($ui.GpoFollowIndicator) { $ui.GpoFollowIndicator.Text = [char]0x25A0 + ' Stopped' }
    Write-DebugLog 'GPO tail stopped' -Level STEP
}

function Load-GpoLogFile {
    $sourceItem = $ui.CmbGpoLogSource.SelectedItem.Content
    Write-Host "[GPO] Load-GpoLogFile called -- source='$sourceItem'"
    if ($Script:GpoTailing) { Stop-GpoTail }
    Clear-GpoLogDisplay
    $ui.GpoLogSubtitle.Text = "Loading: $sourceItem..."

    # Helper: classify WinEvent objects via C# or PS fallback
    $classifyGpoEvents = {
        param([array]$Events, [string]$DefaultComp)
        if ($Events.Count -eq 0) { return }
        if ($Script:HasCSharpParser) {
            $n = $Events.Count
            $msgs = [string[]]::new($n); $lvls = [int[]]::new($n); $comps = [string[]]::new($n)
            $eids = [string[]]::new($n); $times = [string[]]::new($n); $threads = [string[]]::new($n)
            for ($i = 0; $i -lt $n; $i++) {
                $ev = $Events[$i]
                $msgs[$i] = ($ev.Message -replace '\r?\n', ' ').Trim()
                $lvls[$i] = [int]$ev.Level
                $comps[$i] = if ($DefaultComp) { $DefaultComp } else { $ev.ProviderName -replace 'Microsoft-Windows-','' }
                $eids[$i] = "$($ev.Id)"; $times[$i] = $ev.TimeCreated.ToString('M-d-yyyy HH:mm:ss'); $threads[$i] = "$($ev.ProcessId)"
            }
            [PolicyLogParser]::ClassifyGpoEntries($msgs, $lvls, $comps, $eids, $times, $threads, $Script:GpoActiveFilter, $Script:GpoContentFilter)
            if ([PolicyLogParser]::TotalLines -gt 0) {
                $Script:GpoFormattedLines.AddRange([PolicyLogParser]::FormattedLines)
                $Script:GpoLineTypes.AddRange([PolicyLogParser]::LineTypes)
                $Script:GpoRawEntries.AddRange([PolicyLogParser]::RawEntries)
                $Script:GpoStats.Lines += [PolicyLogParser]::TotalLines; $Script:GpoStats.Errors += [PolicyLogParser]::ErrorCount; $Script:GpoStats.Warnings += [PolicyLogParser]::WarningCount
            }
        } else {
            foreach ($ev in $Events) {
                $parsed = @{ Message = ($ev.Message -replace '\r?\n', ' ').Trim(); Time = $ev.TimeCreated.ToString('M-d-yyyy HH:mm:ss'); Component = $(if ($DefaultComp) { $DefaultComp } else { $ev.ProviderName -replace 'Microsoft-Windows-','' }); EventId = "$($ev.Id)"; Level = [int]$ev.Level; Thread = "$($ev.ProcessId)" }
                $result = Classify-GpoLogEntry $parsed; if ($result) { $Script:GpoFormattedLines.Add($result.Formatted); $Script:GpoLineTypes.Add($result.LineType); $Script:GpoRawEntries.Add($result.Raw) }
            }
        }
    }

    # Helper: classify gpsvc.log file via C# or PS fallback
    $classifyGpsvcFile = {
        param([string]$FilePath)
        if ($Script:HasCSharpParser) {
            [PolicyLogParser]::ParseAndClassifyGpsvcFile($FilePath, $Script:GpoActiveFilter, $Script:GpoContentFilter)
            if ([PolicyLogParser]::TotalLines -gt 0) {
                $Script:GpoFormattedLines.AddRange([PolicyLogParser]::FormattedLines)
                $Script:GpoLineTypes.AddRange([PolicyLogParser]::LineTypes)
                $Script:GpoRawEntries.AddRange([PolicyLogParser]::RawEntries)
                $Script:GpoStats.Lines += [PolicyLogParser]::TotalLines; $Script:GpoStats.Errors += [PolicyLogParser]::ErrorCount; $Script:GpoStats.Warnings += [PolicyLogParser]::WarningCount
            }
        } else {
            $gpsvcLines = [System.IO.File]::ReadAllLines($FilePath)
            foreach ($line in $gpsvcLines) {
                if (-not $line.Trim()) { continue }
                $m = $Script:GpoRxGpsvcLine.Match($line)
                if ($m.Success) { $parsed = @{ Message = $m.Groups['msg'].Value; Time = $m.Groups['time'].Value; Component = $m.Groups['func'].Value; EventId = ''; Level = 4; Thread = $m.Groups['tid'].Value } }
                else { $parsed = @{ Message = $line; Time = ''; Component = ''; EventId = ''; Level = 4; Thread = '' } }
                $result = Classify-GpoLogEntry $parsed
                if ($result) { $Script:GpoFormattedLines.Add($result.Formatted); $Script:GpoLineTypes.Add($result.LineType); $Script:GpoRawEntries.Add($result.Raw) }
            }
        }
    }

  try {
    if ($sourceItem -match 'Operational') {
        try {
            $events = Get-WinEvent -FilterHashtable @{ LogName = 'Microsoft-Windows-GroupPolicy/Operational' } -MaxEvents 5000 -ErrorAction Stop | Sort-Object TimeCreated
            & $classifyGpoEvents $events $null
        } catch {
            if ($_.Exception.Message -match 'unauthorized|access') { Show-Toast 'Access Denied' 'Reading GP Operational log requires administrator privileges.' -Type Warning }
            elseif ($_.Exception.Message -notmatch 'No events were found') { Write-Host "[GPO] Load error: $_" }
        }
    } elseif ($sourceItem -match 'System Log') {
        try {
            $events = Get-WinEvent -FilterHashtable @{ LogName = 'System'; ProviderName = @('Microsoft-Windows-GroupPolicy','GroupPolicy') } -MaxEvents 5000 -ErrorAction Stop | Sort-Object TimeCreated
            & $classifyGpoEvents $events 'GroupPolicy'
        } catch { if ($_.Exception.Message -notmatch 'No events were found') { Write-Host "[GPO] Load error: $_" } }
    } elseif ($sourceItem -match 'Preferences') {
        $traceFiles = Get-GppTraceFiles
        if ($traceFiles.Count -gt 0) {
            $traceFile = $traceFiles[0].FullName; $allText = [System.IO.File]::ReadAllText($traceFile); $gppLines = $allText -split "\r?\n"
            foreach ($line in $gppLines) { if (-not $line.Trim()) { continue }; $ts = ''; $comp = 'GPPrefs'; $lvl = 4
                if ($line -match '^\d{4}-\d{2}-\d{2}\s+(\d{2}:\d{2}:\d{2})') { $ts = $Matches[1] }
                if ($line -match '(?i)\berror\b|fail(ed|ure)?|exception|HRESULT|0x8') { $lvl = 2 } elseif ($line -match '(?i)\bwarn(ing)?\b') { $lvl = 3 }
                $parsed = @{ Message = $line; Time = $ts; Component = $comp; EventId = ''; Level = $lvl; Thread = '' }
                $result = Classify-GpoLogEntry $parsed; if ($result) { $Script:GpoFormattedLines.Add($result.Formatted); $Script:GpoLineTypes.Add($result.LineType); $Script:GpoRawEntries.Add($result.Raw) }
            }
        } else { Show-Toast 'GP Prefs Trace' "No trace files found in $Script:GppTraceDir" -Type Warning }
    } elseif ($sourceItem -match 'Combined|All Sources') {
        # System log (no admin)
        try {
            $events = Get-WinEvent -FilterHashtable @{ LogName = 'System'; ProviderName = @('Microsoft-Windows-GroupPolicy','GroupPolicy') } -MaxEvents 5000 -ErrorAction Stop | Sort-Object TimeCreated
            & $classifyGpoEvents $events 'GroupPolicy'
        } catch {
            if ($_.Exception.Message -notmatch 'No events were found') { Write-Host "[GPO] SystemLog: $_" }
        }
        # Operational (may need admin)
        try {
            $events = Get-WinEvent -FilterHashtable @{ LogName = 'Microsoft-Windows-GroupPolicy/Operational' } -MaxEvents 5000 -ErrorAction Stop | Sort-Object TimeCreated
            & $classifyGpoEvents $events $null
        } catch {
            try { Write-DebugLog "Unhandled: $_" -Level ERROR } catch {}
        }
        # gpsvc.log
        if (Test-Path $Script:GpoDebugLogPath) {
            try { & $classifyGpsvcFile $Script:GpoDebugLogPath } catch { Write-Host "[GPO] gpsvc.log read: $_" }
        }
    } else {
        # gpsvc.log (Debug)
        if (Test-Path $Script:GpoDebugLogPath) {
            try { & $classifyGpsvcFile $Script:GpoDebugLogPath } catch { Write-Host "[GPO] gpsvc.log: $_" }
        } else {
            Show-Toast 'Debug Log Not Found' "No gpsvc.log at $Script:GpoDebugLogPath" -Type Warning
        }
    }
  } finally { }

    $ui.lbGpoLogs.ItemsSource = $Script:GpoFormattedLines
    $ui.lbGpoLogs.UpdateLayout()
    $gpoSv = Get-ListBoxScrollViewer $ui.lbGpoLogs
    if ($gpoSv) { $gpoSv.Add_ScrollChanged({ Apply-LogListBoxColors $ui.lbGpoLogs $Script:GpoLineTypes $Script:GpoColorMap $Script:GpoBrushCache; $sv = Get-ListBoxScrollViewer $ui.lbGpoLogs; Update-MinimapViewportIndicator $ui.cnvGpoMinimap $sv $ui.lbGpoLogs.Items.Count }) }
    Apply-LogListBoxColors $ui.lbGpoLogs $Script:GpoLineTypes $Script:GpoColorMap $Script:GpoBrushCache
    if ($ui.cnvGpoHeatmap) { $ui.cnvGpoHeatmap.Visibility = 'Visible'; $ui.cnvGpoHeatmap.UpdateLayout() }
    Rebuild-LogHeatmap $ui.cnvGpoHeatmap $Script:GpoLineTypes $Script:GpoColorMap $Script:GpoBrushCache
    Rebuild-LogMinimap $ui.cnvGpoMinimap $Script:GpoLineTypes $Script:GpoColorMap $Script:GpoBrushCache
    Update-LogSidebarExtras -LineTypes $Script:GpoLineTypes -DensityError $ui.GpoDensityError -DensityWarn $ui.GpoDensityWarn -FindingsList $ui.GpoFindingsList -FormattedLines $Script:GpoFormattedLines -LogListBox $ui.lbGpoLogs -ColorMap $Script:GpoColorMap
    $rendered = $Script:GpoFormattedLines.Count
    Update-GpoStatsDisplay
    $ui.GpoLogSubtitle.Text = "Loaded: $sourceItem ($rendered entries)"
    if ($ui.GpoFollowIndicator) { $ui.GpoFollowIndicator.Text = [char]0x25A0 + ' Analysis' }
    if ($rendered -gt 0) { $ui.lbGpoLogs.ScrollIntoView($ui.lbGpoLogs.Items[$rendered - 1]) }
    Write-Host "[GPO] Load-GpoLogFile complete: $rendered entries classified"
    Show-Toast 'Log Loaded' "$rendered entries from $sourceItem" -Type Info
}

# â”€â”€ GPO DispatcherTimer â”€â”€
$Script:GpoTimer = [System.Windows.Threading.DispatcherTimer]::new()
$Script:GpoTimer.Interval = [TimeSpan]::FromMilliseconds(80)
$Script:GpoStatsThrottle = [DateTime]::MinValue
$Script:GpoEventPollLast = [DateTime]::MinValue
$Script:GpoPollThrottle  = [DateTime]::MinValue
$Script:GpoTimer.Add_Tick({
    if (-not $Script:GpoTailing) { return }
    # Poll based on active source
    if ($Script:GpoActiveSource -eq 'EventLog') {
        $now = [DateTime]::Now; if (($now - $Script:GpoEventPollLast).TotalSeconds -ge 2) { $Script:GpoEventPollLast = $now; Read-GpoEventLogDelta }
    } elseif ($Script:GpoActiveSource -eq 'SystemLog') {
        $now = [DateTime]::Now; if (($now - $Script:GpoEventPollLast).TotalSeconds -ge 2) { $Script:GpoEventPollLast = $now; Read-GpoSystemLogDelta }
    } elseif ($Script:GpoActiveSource -eq 'PrefsTrace') {
        $needsRead = $Script:GppFswChanged
        if (-not $needsRead) { $now = [DateTime]::Now; if (($now - $Script:GpoPollThrottle).TotalMilliseconds -gt 500) { $Script:GpoPollThrottle = $now; if ($Script:GppActiveLogPath -and (Test-Path $Script:GppActiveLogPath)) { $len = ([System.IO.FileInfo]::new($Script:GppActiveLogPath)).Length; if ($len -gt $Script:GppLastFilePos) { $needsRead = $true } } } }
        if ($needsRead) { $Script:GppFswChanged = $false; Read-GppTraceLogDelta }
    } elseif ($Script:GpoActiveSource -eq 'Combined') {
        $now = [DateTime]::Now; if (($now - $Script:GpoEventPollLast).TotalSeconds -ge 2) { $Script:GpoEventPollLast = $now; Read-GpoSystemLogDelta; try { Read-GpoEventLogDelta } catch { try { Write-DebugLog "Unhandled: $_" -Level ERROR } catch {} } }
        $needsGpo = $Script:GpoFswChanged; $needsGpp = $Script:GppFswChanged
        if (-not $needsGpo -or -not $needsGpp) {
            $nowP = [DateTime]::Now; if (($nowP - $Script:GpoPollThrottle).TotalMilliseconds -gt 500) { $Script:GpoPollThrottle = $nowP
                if (-not $needsGpo -and (Test-Path $Script:GpoDebugLogPath)) { $len = ([System.IO.FileInfo]::new($Script:GpoDebugLogPath)).Length; if ($len -gt $Script:GpoLastFilePos) { $needsGpo = $true } }
                if (-not $needsGpp -and $Script:GppActiveLogPath -and (Test-Path $Script:GppActiveLogPath)) { $len = ([System.IO.FileInfo]::new($Script:GppActiveLogPath)).Length; if ($len -gt $Script:GppLastFilePos) { $needsGpp = $true } }
            }
        }
        if ($needsGpo) { $Script:GpoFswChanged = $false; Read-GpoDebugLogDelta }; if ($needsGpp) { $Script:GppFswChanged = $false; Read-GppTraceLogDelta }
    } else {
        $needsRead = $Script:GpoFswChanged
        if (-not $needsRead) { $now = [DateTime]::Now; if (($now - $Script:GpoPollThrottle).TotalMilliseconds -gt 500) { $Script:GpoPollThrottle = $now; if ($Script:GpoActiveLogPath -and (Test-Path $Script:GpoActiveLogPath)) { $len = ([System.IO.FileInfo]::new($Script:GpoActiveLogPath)).Length; if ($len -gt $Script:GpoLastFilePos) { $needsRead = $true } } } }
        if ($needsRead) { $Script:GpoFswChanged = $false; Read-GpoDebugLogDelta }
    }
    # Drain queue
    $batchCount = 0; $contentAdded = $false
    $batchEntries = [System.Collections.Generic.List[hashtable]]::new()
    while ($Script:GpoLogQueue.Count -gt 0 -and $batchCount -lt $Script:GpoMaxPerTick) {
        $batchEntries.Add($Script:GpoLogQueue.Dequeue()); $batchCount++
    }
    if ($batchEntries.Count -gt 0) {
        if ($Script:HasCSharpParser) {
            $n = $batchEntries.Count
            $msgs = [string[]]::new($n); $lvls = [int[]]::new($n); $comps = [string[]]::new($n)
            $eids = [string[]]::new($n); $times = [string[]]::new($n); $threads = [string[]]::new($n)
            for ($i = 0; $i -lt $n; $i++) {
                $e = $batchEntries[$i]; $msgs[$i] = $e.Message; $lvls[$i] = [int]$e.Level
                $comps[$i] = if ($e.Component) { $e.Component } else { '' }
                $eids[$i] = if ($e.EventId) { $e.EventId } else { '' }
                $times[$i] = if ($e.Time) { $e.Time } else { '' }
                $threads[$i] = if ($e.Thread) { $e.Thread } else { '' }
            }
            [PolicyLogParser]::ClassifyGpoEntries($msgs, $lvls, $comps, $eids, $times, $threads, $Script:GpoActiveFilter, $Script:GpoContentFilter)
            if ([PolicyLogParser]::TotalLines -gt 0) {
                $Script:GpoFormattedLines.AddRange([PolicyLogParser]::FormattedLines); $Script:GpoLineTypes.AddRange([PolicyLogParser]::LineTypes); $Script:GpoRawEntries.AddRange([PolicyLogParser]::RawEntries)
                $Script:GpoStats.Lines += [PolicyLogParser]::TotalLines; $Script:GpoStats.Errors += [PolicyLogParser]::ErrorCount; $Script:GpoStats.Warnings += [PolicyLogParser]::WarningCount
                $contentAdded = $true
            }
        } else {
            foreach ($entry in $batchEntries) {
                $result = Classify-GpoLogEntry $entry
                if ($result) { $Script:GpoFormattedLines.Add($result.Formatted); $Script:GpoLineTypes.Add($result.LineType); $Script:GpoRawEntries.Add($result.Raw); $contentAdded = $true }
            }
        }
    }
    if ($contentAdded) {
        $ui.lbGpoLogs.ItemsSource = $null; $ui.lbGpoLogs.ItemsSource = $Script:GpoFormattedLines
        $sv = Get-ListBoxScrollViewer $ui.lbGpoLogs
        if ($sv) {
            $isLatched = ($sv.VerticalOffset + $sv.ViewportHeight) -ge ($sv.ExtentHeight - 20.0)
            if ($isLatched -and $Script:GpoFormattedLines.Count -gt 0) { $ui.lbGpoLogs.ScrollIntoView($ui.lbGpoLogs.Items[$Script:GpoFormattedLines.Count - 1]); if ($ui.GpoFollowIndicator) { $ui.GpoFollowIndicator.Text = [char]0x25BC + ' Following' } }
            else { if ($ui.GpoFollowIndicator) { $ui.GpoFollowIndicator.Text = [char]0x25A0 + ' Paused' } }
        }
        $ui.lbGpoLogs.UpdateLayout(); Apply-LogListBoxColors $ui.lbGpoLogs $Script:GpoLineTypes $Script:GpoColorMap $Script:GpoBrushCache
        $now2 = [DateTime]::Now; if (($now2 - $Script:GpoStatsThrottle).TotalMilliseconds -gt 500) { Update-GpoStatsDisplay; $Script:GpoStatsThrottle = $now2 }
    }
})
$Script:GpoTimer.Start()

# ── GPO Debug Logging Toggle ──
function Update-GpoDebugToggleUI {
    $debugEnabled = $false
    try {
        if (Test-Path $Script:GpoDebugRegPath) {
            $val = Get-ItemProperty -Path $Script:GpoDebugRegPath -Name 'GPSvcDebugLevel' -ErrorAction SilentlyContinue
            if ($val -and $val.GPSvcDebugLevel -gt 0) { $debugEnabled = $true }
        }
    } catch { try { Write-DebugLog "Unhandled: $_" -Level ERROR } catch {} }
    if ($debugEnabled) {
        if ($ui.GpoDebugToggleLabel) { $ui.GpoDebugToggleLabel.Text = 'Disable Debug Logging' }
        if ($ui.GpoDebugStatus) { $ui.GpoDebugStatus.Text = "Debug logging enabled`ngpsvc.log: $Script:GpoDebugLogPath" }
    } else {
        if ($ui.GpoDebugToggleLabel) { $ui.GpoDebugToggleLabel.Text = 'Enable Debug Logging' }
        if ($ui.GpoDebugStatus) { $ui.GpoDebugStatus.Text = '' }
    }
}

function Toggle-GpoDebugLogging {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    $debugEnabled = $false
    try {
        if (Test-Path $Script:GpoDebugRegPath) {
            $val = Get-ItemProperty -Path $Script:GpoDebugRegPath -Name 'GPSvcDebugLevel' -ErrorAction SilentlyContinue
            if ($val -and $val.GPSvcDebugLevel -gt 0) { $debugEnabled = $true }
        }
    } catch { try { Write-DebugLog "Unhandled: $_" -Level ERROR } catch {} }

    if ($debugEnabled) {
        # Disable debug logging
        if ($isAdmin) {
            try {
                Remove-ItemProperty -Path $Script:GpoDebugRegPath -Name 'GPSvcDebugLevel' -Force -ErrorAction Stop
                Show-Toast 'Debug Logging' 'GPO debug logging disabled. Restart gpsvc service for changes to take effect.' -Type Info
            } catch {
                Show-Toast 'Error' "Failed to disable debug logging: $($_.Exception.Message)" -Type Error
            }
        } else {
            try {
                Start-Process powershell.exe -ArgumentList "-NoProfile -Command `"Remove-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Diagnostics' -Name 'GPSvcDebugLevel' -Force`"" -Verb RunAs -Wait -WindowStyle Hidden
                Show-Toast 'Debug Logging' 'GPO debug logging disabled (elevated). Restart gpsvc service for changes to take effect.' -Type Info
            } catch {
                Show-Toast 'Error' "Elevation cancelled or failed: $($_.Exception.Message)" -Type Warning
            }
        }
    } else {
        # Enable debug logging
        $regCmd = "if(-not(Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Diagnostics')){New-Item -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Diagnostics' -Force}; Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Diagnostics' -Name 'GPSvcDebugLevel' -Value 0x00030002 -Type DWord -Force"
        if ($isAdmin) {
            try {
                Invoke-Expression $regCmd
                Show-Toast 'Debug Logging' "GPO debug logging enabled.`nLog: $Script:GpoDebugLogPath`nRestart gpsvc service for changes to take effect." -Type Info
            } catch {
                Show-Toast 'Error' "Failed to enable debug logging: $($_.Exception.Message)" -Type Error
            }
        } else {
            try {
                Start-Process powershell.exe -ArgumentList "-NoProfile -Command `"$regCmd`"" -Verb RunAs -Wait -WindowStyle Hidden
                Show-Toast 'Debug Logging' "GPO debug logging enabled (elevated).`nLog: $Script:GpoDebugLogPath`nRestart gpsvc service for changes to take effect." -Type Info
            } catch {
                Show-Toast 'Error' "Elevation cancelled or failed: $($_.Exception.Message)" -Type Warning
            }
        }
    }
    Update-GpoDebugToggleUI
}

# ── Force GPUpdate ──
function Invoke-GpoRefresh {
    $ui.BtnGpoRefresh.IsEnabled = $false
    Write-DebugLog 'Triggering gpupdate /force' -Level STEP

    # Auto-start GPO log tail if not already tailing
    if (-not $Script:GpoTailing) {
        Write-Host '[GPO] Auto-starting tail for gpupdate monitoring'
        Start-GpoTail
    }

    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if ($isAdmin) {
        $ui.GpoRefreshStatus.Text = 'Running gpupdate /force (elevated)...'
        Start-BackgroundWork -Work {
            $output = & gpupdate.exe /force 2>&1 | Out-String
            return $output
        } -OnComplete {
            param($result)
            $brief = ($result -split '\r?\n' | Where-Object { $_ -match '\S' } | Select-Object -Last 3) -join '; '
            $ui.GpoRefreshStatus.Text = $brief
            $ui.BtnGpoRefresh.IsEnabled = $true
            Show-Toast 'GPUpdate' $brief -Type Info
            Write-DebugLog "GPUpdate result: $brief" -Level STEP
        }.GetNewClosure() -Variables @{} -Context @{ Name = 'GPUpdate' }
    } else {
        $ui.GpoRefreshStatus.Text = 'Elevating for gpupdate /force (machine policy needs admin)...'
        try {
            $tmpOut = [System.IO.Path]::Combine($env:TEMP, 'gpupdate_result.txt')
            $cmd = "gpupdate.exe /force 2>&1 | Out-String | Set-Content -Path '$tmpOut' -Encoding UTF8 -Force"
            Start-Process powershell.exe -ArgumentList "-NoProfile -Command `"$cmd`"" -Verb RunAs -Wait -WindowStyle Hidden
            $brief = ''
            if (Test-Path $tmpOut) {
                $brief = (Get-Content $tmpOut -Raw -ErrorAction SilentlyContinue) -replace '\r?\n', '; '
                $brief = ($brief -split ';' | Where-Object { $_ -match '\S' } | Select-Object -Last 3) -join '; '
                Remove-Item $tmpOut -Force -ErrorAction SilentlyContinue
            }
            if (-not $brief) { $brief = 'gpupdate completed (elevated)' }
            $ui.GpoRefreshStatus.Text = $brief
            Show-Toast 'GPUpdate' $brief -Type Info
            Write-DebugLog "GPUpdate result (elevated): $brief" -Level STEP
        } catch {
            $ui.GpoRefreshStatus.Text = "Elevation cancelled or failed: $($_.Exception.Message)"
            Show-Toast 'GPUpdate' 'Elevation cancelled or failed' -Type Warning
            Write-DebugLog "GPUpdate elevation failed: $($_.Exception.Message)" -Level WARN
        }
        $ui.BtnGpoRefresh.IsEnabled = $true
    }
}

# ── GPO Search ──
function Search-GpoHighlight {
    param([string]$Term)
    $Script:GpoSearchMatches.Clear(); $Script:GpoSearchIndex = -1
    if ([string]::IsNullOrWhiteSpace($Term)) {
        if ($ui.GpoSearchCount) { $ui.GpoSearchCount.Text = '' }
        Apply-LogListBoxColors $ui.lbGpoLogs $Script:GpoLineTypes $Script:GpoColorMap $Script:GpoBrushCache
        return
    }
    $rx = [regex]::new([regex]::Escape($Term), 'IgnoreCase')
    for ($i = 0; $i -lt $Script:GpoRawEntries.Count; $i++) {
        if ($rx.IsMatch($Script:GpoRawEntries[$i]) -or $rx.IsMatch($Script:GpoFormattedLines[$i])) { [void]$Script:GpoSearchMatches.Add($i) }
    }
    if ($ui.GpoSearchCount) { $ui.GpoSearchCount.Text = "$($Script:GpoSearchMatches.Count) matches" }
    if ($Script:GpoSearchMatches.Count -gt 0) { $Script:GpoSearchIndex = 0; Navigate-GpoSearchMatch 0 }
}

function Navigate-GpoSearchMatch([int]$Index) {
    if ($Index -lt 0 -or $Index -ge $Script:GpoSearchMatches.Count) { return }
    $Script:GpoSearchIndex = $Index
    $lineIdx = $Script:GpoSearchMatches[$Index]
    $ui.lbGpoLogs.ScrollIntoView($ui.lbGpoLogs.Items[$lineIdx]); $ui.lbGpoLogs.SelectedIndex = $lineIdx
    if ($ui.GpoSearchCount) { $ui.GpoSearchCount.Text = "$($Index+1)/$($Script:GpoSearchMatches.Count)" }
}

# Initialize debug toggle state
Update-GpoDebugToggleUI

# ── Admin Banner Init ──
$Script:IsAdminSession = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $Script:IsAdminSession) {
    if ($ui.ImeSyncAdminBanner) { $ui.ImeSyncAdminBanner.Visibility = 'Visible' }
    if ($ui.ImeSyncBtnLabel)    { $ui.ImeSyncBtnLabel.Text = 'Force Intune Sync (Elevate)' }
    if ($ui.GpoAdminBanner)     { $ui.GpoAdminBanner.Visibility = 'Visible' }
    if ($ui.GpoAdminBannerText) { $ui.GpoAdminBannerText.Text = "Not running as admin - GP Operational log, debug logging, and gpupdate require elevation" }
}

# ── GP Preferences Trace Status ──
$gppFiles = Get-GppTraceFiles
if ($ui.GppTraceStatus) {
    if ($gppFiles.Count -gt 0) {
        $ui.GppTraceStatus.Text = "$($gppFiles.Count) trace file(s) found"
    } else {
        $ui.GppTraceStatus.Text = "No trace files - enable logging in gpedit.msc"
    }
}

# ── Wire up GPO Log UI events ──
if ($ui.BtnGpoWrapToggle) {
    $ui.BtnGpoWrapToggle.Add_Click({
        $scroll = $ui.lbGpoLogs.GetValue([System.Windows.Controls.ScrollViewer]::HorizontalScrollBarVisibilityProperty)
        if ($scroll -eq 'Disabled') { $ui.lbGpoLogs.SetValue([System.Windows.Controls.ScrollViewer]::HorizontalScrollBarVisibilityProperty, [System.Windows.Controls.ScrollBarVisibility]::Auto) }
        else { $ui.lbGpoLogs.SetValue([System.Windows.Controls.ScrollViewer]::HorizontalScrollBarVisibilityProperty, [System.Windows.Controls.ScrollBarVisibility]::Disabled) }
    })
}
if ($ui.BtnGpoClearLog)    { $ui.BtnGpoClearLog.Add_Click({ Clear-GpoLogDisplay }) }
if ($ui.BtnGpoDebugToggle) { $ui.BtnGpoDebugToggle.Add_Click({ Toggle-GpoDebugLogging }) }
if ($ui.BtnGpoRefresh)     { $ui.BtnGpoRefresh.Add_Click({ Invoke-GpoRefresh }) }

# Mode toggle: Live tail vs Analysis
if ($ui.ChkGpoLiveTail) {
    $ui.ChkGpoLiveTail.Add_Checked({
        if ($ui.GpoModeHint) { $ui.GpoModeHint.Text = 'Streaming new log entries in real time' }
        if ($Script:GpoTailing) { return }  # already tailing
        Start-GpoTail
    })
    $ui.ChkGpoLiveTail.Add_Unchecked({
        if ($ui.GpoModeHint) { $ui.GpoModeHint.Text = 'Full log loaded for analysis and search' }
        if ($Script:GpoTailing) { Stop-GpoTail }
        Load-GpoLogFile
    })
}
# Log source change: restart current mode
if ($ui.CmbGpoLogSource) {
    $ui.CmbGpoLogSource.Add_SelectionChanged({
        if ($ui.ChkGpoLiveTail -and $ui.ChkGpoLiveTail.IsChecked) {
            if ($Script:GpoTailing) { Stop-GpoTail }
            Start-GpoTail
        } else {
            Load-GpoLogFile
        }
    })
}

if ($ui.TxtGpoSearch) {
    $Script:GpoSearchDebounce = $null
    $ui.TxtGpoSearch.Add_TextChanged({
        if ($Script:GpoSearchDebounce) { $Script:GpoSearchDebounce.Stop() }
        $Script:GpoSearchDebounce = [System.Windows.Threading.DispatcherTimer]::new()
        $Script:GpoSearchDebounce.Interval = [TimeSpan]::FromMilliseconds(300)
        $Script:GpoSearchDebounce.Add_Tick({
            $Script:GpoSearchDebounce.Stop()
            Search-GpoHighlight $ui.TxtGpoSearch.Text
        })
        $Script:GpoSearchDebounce.Start()
    })
}
if ($ui.BtnGpoSearchNext) {
    $ui.BtnGpoSearchNext.Add_Click({
        if ($Script:GpoSearchMatches.Count -gt 0) {
            $next = ($Script:GpoSearchIndex + 1) % $Script:GpoSearchMatches.Count
            Navigate-GpoSearchMatch $next
        }
    })
}
if ($ui.BtnGpoSearchPrev) {
    $ui.BtnGpoSearchPrev.Add_Click({
        if ($Script:GpoSearchMatches.Count -gt 0) {
            $prev = $Script:GpoSearchIndex - 1
            if ($prev -lt 0) { $prev = $Script:GpoSearchMatches.Count - 1 }
            Navigate-GpoSearchMatch $prev
        }
    })
}

# GPO Severity filter pills
foreach ($filterName in @('FilterGpoAll','FilterGpoError','FilterGpoWarning','FilterGpoInfo')) {
    if ($ui[$filterName]) {
        $ui[$filterName].Add_Click({
            param($sender)
            $label = $sender.Content
            $Script:GpoActiveFilter = switch ($label) {
                'All'      { 'All' }
                'Errors'   { 'Error' }
                'Warnings' { 'Warning' }
                'Info'     { 'Info' }
                default    { 'All' }
            }
            foreach ($fn in @('FilterGpoAll','FilterGpoError','FilterGpoWarning','FilterGpoInfo')) {
                if ($ui[$fn]) { $ui[$fn].Tag = if ($ui[$fn] -eq $sender) { 'Active' } else { $null } }
            }
            # Re-apply with new filter - respect mode toggle
            if ($ui.ChkGpoLiveTail -and $ui.ChkGpoLiveTail.IsChecked) {
                # Live mode: re-read from start with filter
                if ($Script:GpoTailing) {
                    $prevSource = $Script:GpoActiveSource
                    Stop-GpoTail
                    Clear-GpoLogDisplay
                    $Script:GpoActiveSource = $prevSource
                    $Script:GpoLastEventTime = $null
                    $Script:GpoSystemLastTime = $null
                    $Script:GpoLastFilePos = 0
                    $Script:GppLastFilePos = 0
                    $Script:GpoTailing = $true
                    switch ($prevSource) {
                        'EventLog'   { Read-GpoEventLogDelta }
                        'SystemLog'  { Read-GpoSystemLogDelta }
                        'PrefsTrace' { Read-GppTraceLogDelta }
                        'Combined'   {
                            Read-GpoSystemLogDelta
                            try { Read-GpoEventLogDelta } catch { try { Write-DebugLog "Unhandled: $_" -Level ERROR } catch {} }
                            if (Test-Path $Script:GpoDebugLogPath) { Read-GpoDebugLogDelta }
                            if ($Script:GppActiveLogPath -and (Test-Path $Script:GppActiveLogPath)) { Read-GppTraceLogDelta }
                        }
                        'DebugLog'   { Read-GpoDebugLogDelta }
                    }
                }
            } else {
                # Analysis mode: reload with filter
                Load-GpoLogFile
            }
        })
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 19.7: MDM SYNC VIEWER - OMA-DM / SyncML DIAGNOSTIC EVENT LOGS
# ═══════════════════════════════════════════════════════════════════════════════

# ── MDM Regex Patterns ──
$Script:MdmRxError    = [regex]::new('(?i)(error|fail(ed|ure)?|denied|not found|could not|unable to|HRESULT|exception|0x8\w{7}|rejected|unauthorized|timeout)', 'Compiled')
$Script:MdmRxWarning  = [regex]::new('(?i)(warn(ing)?|retry|retrying|skipped|not applicable|conflict|pending|throttl)', 'Compiled')
$Script:MdmRxSuccess  = [regex]::new('(?i)(success(fully)?|completed|accepted|applied|result.*=.*0\b|status.*200|status.*OK)', 'Compiled')
$Script:MdmRxSyncML   = [regex]::new('(?i)(SyncML|SyncBody|SyncHdr|OMA-?DM|<Add>|<Replace>|<Get>|<Delete>|<Exec>|<Alert>|<Status>|<Final>)', 'Compiled')
$Script:MdmRxPolicy   = [regex]::new('(?i)(policy|configuration|compliance|enrollment|MDM session|check-?in|CSP|./Device/|./User/|./Vendor/)', 'Compiled')
$Script:MdmRxNetwork  = [regex]::new('(?i)(http|https|push notification|WNS|MPNS|server|endpoint|DMClient|connection|channel)', 'Compiled')

function Get-MdmLineColor {
    param([string]$Line, [int]$EventLevel)
    if ($EventLevel -eq 1 -or $EventLevel -eq 2) { return @{ Color = '#D13438'; Bold = $true;  Cat = 'Error'   } }
    if ($EventLevel -eq 3) { return @{ Color = '#FFB900'; Bold = $false; Cat = 'Warning' } }
    if ($Script:MdmRxError.IsMatch($Line))   { return @{ Color = '#D13438'; Bold = $true;  Cat = 'Error'   } }
    if ($Script:MdmRxWarning.IsMatch($Line)) { return @{ Color = '#FFB900'; Bold = $false; Cat = 'Warning' } }
    if ($Script:MdmRxSuccess.IsMatch($Line)) { return @{ Color = '#107C10'; Bold = $true;  Cat = 'Success' } }
    if ($Script:MdmRxSyncML.IsMatch($Line))  { return @{ Color = '#60CDFF'; Bold = $false; Cat = 'SyncML'  } }
    if ($Script:MdmRxPolicy.IsMatch($Line))  { return @{ Color = '#0078D4'; Bold = $false; Cat = 'Policy'  } }
    if ($Script:MdmRxNetwork.IsMatch($Line)) { return @{ Color = '#FF8C00'; Bold = $false; Cat = 'Network' } }
    return @{ Color = '#C0C0C0'; Bold = $false; Cat = 'Info' }
}

function Get-MdmSeverityBadge {
    param([string]$Cat)
    switch ($Cat) {
        'Error'   { return @{ Label = 'ERR';  Bg = '#D13438'; Fg = '#FFFFFF' } }
        'Warning' { return @{ Label = 'WARN'; Bg = '#FFB900'; Fg = '#1A1A1A' } }
        'Success' { return @{ Label = 'OK';   Bg = '#107C10'; Fg = '#FFFFFF' } }
        'SyncML'  { return @{ Label = 'SYNC'; Bg = '#60CDFF'; Fg = '#1A1A1A' } }
        'Policy'  { return @{ Label = 'POL';  Bg = '#0078D4'; Fg = '#FFFFFF' } }
        'Network' { return @{ Label = 'NET';  Bg = '#FF8C00'; Fg = '#FFFFFF' } }
        default   { return @{ Label = 'INFO'; Bg = '#333333'; Fg = '#AAAAAA' } }
    }
}

# â”€â”€ MDM Log State â”€â”€
$Script:MdmLogQueue       = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new())
$Script:MdmTailing        = $false
$Script:MdmMaxPerTick     = 150
$Script:MdmActiveFilter   = 'All'
$Script:MdmContentFilter   = ''
$Script:MdmSavedFilters    = @{}
$Script:MdmActiveSource   = 'AdminDiag'
$Script:MdmLastEventTime  = $null
$Script:MdmStats = @{ Lines = 0; Errors = 0; Warnings = 0 }
$Script:MdmSearchMatches  = [System.Collections.Generic.List[int]]::new()
$Script:MdmSearchIndex    = -1
$Script:MdmBrushCache     = @{}
$Script:MdmFormattedLines  = [System.Collections.Generic.List[string]]::new()
$Script:MdmLineTypes       = [System.Collections.Generic.List[byte]]::new()
$Script:MdmRawEntries      = [System.Collections.Generic.List[string]]::new()
$Script:MdmColorMap = @( @{ Fg = '#C0C0C0'; Bold = $false }, @{ Fg = '#D13438'; Bold = $true }, @{ Fg = '#FFB900'; Bold = $false }, @{ Fg = '#107C10'; Bold = $true }, @{ Fg = '#60CDFF'; Bold = $false }, @{ Fg = '#0078D4'; Bold = $false }, @{ Fg = '#FF8C00'; Bold = $false } )
$Script:MdmBadgeMap = @('INFO','ERR','WARN','OK','SYNC','POL','NET')
function Update-MdmStatsDisplay {
    if ($ui.MdmStatLines)    { $ui.MdmStatLines.Text    = $Script:MdmStats.Lines }
    if ($ui.MdmStatErrors)   { $ui.MdmStatErrors.Text   = $Script:MdmStats.Errors }
    if ($ui.MdmStatWarnings) { $ui.MdmStatWarnings.Text = $Script:MdmStats.Warnings }
}
function Clear-MdmLogDisplay {
    $Script:MdmFormattedLines.Clear(); $Script:MdmLineTypes.Clear(); $Script:MdmRawEntries.Clear()
    $ui.lbMdmLogs.ItemsSource = $null
    $Script:MdmStats.Lines = 0; $Script:MdmStats.Errors = 0; $Script:MdmStats.Warnings = 0
    Update-MdmStatsDisplay
    if ($ui.cnvMdmHeatmap) { $ui.cnvMdmHeatmap.Children.Clear() }; if ($ui.cnvMdmMinimap) { $ui.cnvMdmMinimap.Children.Clear() }
    $Script:MdmSearchMatches.Clear(); $Script:MdmSearchIndex = -1
    if ($ui.MdmSearchCount) { $ui.MdmSearchCount.Text = '' }
}
function Classify-MdmLogEntry {
    param([hashtable]$Parsed)
    $msg   = $Parsed.Message
    $rules = Get-MdmLineColor -Line $msg -EventLevel $Parsed.Level
    if ($Script:MdmActiveFilter -ne 'All') {
        $cat = $rules.Cat
        $pass = switch ($Script:MdmActiveFilter) { 'Error' { $cat -eq 'Error' }; 'Warning' { $cat -eq 'Warning' }; 'Info' { $cat -notin @('Error','Warning') }; default { $true } }
        if (-not $pass) { return $null }
    }
    if ($Script:MdmContentFilter -and $Script:MdmContentFilter.Length -gt 0) {
        $pattern = $Script:MdmContentFilter
        $textToSearch = "$($Parsed.Message) $($Parsed.Component) $($Parsed.EventId)"
        if ($pattern.StartsWith('/') -and $pattern.EndsWith('/')) {
            try { if ($textToSearch -notmatch $pattern.Substring(1, $pattern.Length - 2)) { return $null } } catch { if ($textToSearch -notlike "*$pattern*") { return $null } }
        } else { if ($textToSearch -notlike "*$pattern*") { return $null } }
    }
    $Script:MdmStats.Lines++; if ($rules.Cat -eq 'Error') { $Script:MdmStats.Errors++ }; if ($rules.Cat -eq 'Warning') { $Script:MdmStats.Warnings++ }
    $lineType = switch ($rules.Cat) { 'Error' { [byte]1 }; 'Warning' { [byte]2 }; 'Success' { [byte]3 }; 'SyncML' { [byte]4 }; 'Policy' { [byte]5 }; 'Network' { [byte]6 }; default { [byte]0 } }
    $badge = $Script:MdmBadgeMap[$lineType]
    $sb = [System.Text.StringBuilder]::new(512)
    [void]$sb.Append("[$badge] ")
    if ($Parsed.Time) { [void]$sb.Append("$($Parsed.Time) ") }
    if ($Parsed.Component) { [void]$sb.Append("[$($Parsed.Component)] ") }
    if ($Parsed.EventId) { [void]$sb.Append("ID:$($Parsed.EventId) ") }
    if ($Parsed.Thread) { [void]$sb.Append("T:$($Parsed.Thread) ") }
    $displayMsg = $msg; if ($displayMsg.Length -gt 300) { $displayMsg = $displayMsg.Substring(0, 300) + ' [...]' }
    [void]$sb.Append($displayMsg)
    return @{ Formatted = $sb.ToString(); LineType = $lineType; Raw = $msg }
}

# ── MDM Event Log Readers ──
function Get-MdmEventLogName {
    $sourceItem = if ($ui.CmbMdmLogSource -and $ui.CmbMdmLogSource.SelectedItem) { $ui.CmbMdmLogSource.SelectedItem.Content } else { 'Admin Diagnostics' }
    if ($sourceItem -match 'Operational') {
        return 'Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Operational'
    }
    return 'Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Admin'
}

function Read-MdmEventLogDelta {
    $logName = Get-MdmEventLogName
    Write-Verbose "[MDM] Read-MdmEventLogDelta - log='$logName' lastEventTime=$($Script:MdmLastEventTime)"
    try {
        $filter = @{ LogName = $logName }
        if ($Script:MdmLastEventTime) {
            $filter['StartTime'] = $Script:MdmLastEventTime.AddMilliseconds(1)
        }
        $events = Get-WinEvent -FilterHashtable $filter -MaxEvents 500 -ErrorAction Stop |
            Sort-Object TimeCreated
        Write-Verbose "[MDM] EventLog: $($events.Count) events found"
        foreach ($ev in $events) {
            $parsed = @{
                Message   = ($ev.Message -replace '\r?\n', ' ').Trim()
                Time      = $ev.TimeCreated.ToString('M-d-yyyy HH:mm:ss')
                Component = $ev.ProviderName -replace 'Microsoft-Windows-',''
                EventId   = "$($ev.Id)"
                Level     = [int]$ev.Level
                Thread    = "$($ev.ProcessId)"
            }
            $Script:MdmLogQueue.Enqueue($parsed)
            $Script:MdmLastEventTime = $ev.TimeCreated
        }
    } catch {
        if ($_.Exception.Message -match 'unauthorized|access') {
            Write-Host "[MDM] EventLog: ACCESS DENIED - needs admin"
            $Script:MdmLogQueue.Enqueue(@{
                Message = "Access denied - reading MDM diagnostic event log requires administrator privileges. Run PolicyPilot as admin."
                Time = (Get-Date).ToString('M-d-yyyy HH:mm:ss'); Component = 'System'; EventId = ''; Level = 2; Thread = ''
            })
            $Script:MdmTailing = $false
        } elseif ($_.Exception.Message -match 'No events were found') {
            Write-Verbose "[MDM] EventLog: No events found"
        } else {
            Write-Host "[MDM] EventLog ERROR: $($_.Exception.Message)"
        }
    }
}

# ── Start / Stop Tail ──
function Start-MdmTail {
    $logName = Get-MdmEventLogName
    Write-Host "[MDM] Start-MdmTail called - source='$logName'"

    Clear-MdmLogDisplay
    $Script:MdmTailing = $true
    $Script:MdmLastEventTime = $null

    $sourceLabel = if ($ui.CmbMdmLogSource -and $ui.CmbMdmLogSource.SelectedItem) { $ui.CmbMdmLogSource.SelectedItem.Content } else { 'Admin Diagnostics' }
    $ui.MdmLogSubtitle.Text = "Tailing: $logName"

    Read-MdmEventLogDelta
    Show-Toast 'MDM Sync Tail' "Now tailing $sourceLabel" -Type Info
}

function Stop-MdmTail {
    $Script:MdmTailing = $false
    $ui.MdmLogSubtitle.Text = 'MDM Sync - Tail stopped'
    if ($ui.MdmFollowIndicator) { $ui.MdmFollowIndicator.Text = [char]0x25A0 + ' Stopped' }
    Write-Host '[MDM] Tail stopped'
}

# ── Analysis Mode: Load bulk ──
function Load-MdmLogFile {
    $logName = Get-MdmEventLogName
    $sourceLabel = if ($ui.CmbMdmLogSource -and $ui.CmbMdmLogSource.SelectedItem) { $ui.CmbMdmLogSource.SelectedItem.Content } else { 'Admin Diagnostics' }
    if ($Script:MdmTailing) { Stop-MdmTail }
    Clear-MdmLogDisplay
    $ui.MdmLogSubtitle.Text = "Loading: $sourceLabel..."
    try {
        $filter = @{ LogName = $logName }
        $events = Get-WinEvent -FilterHashtable $filter -MaxEvents 5000 -ErrorAction Stop | Sort-Object TimeCreated
        Write-Host "[MDM] Load: $($events.Count) events from $logName"
        if ($Script:HasCSharpParser -and $events.Count -gt 0) {
            $n = $events.Count
            $msgs = [string[]]::new($n); $lvls = [int[]]::new($n); $comps = [string[]]::new($n)
            $eids = [string[]]::new($n); $times = [string[]]::new($n); $threads = [string[]]::new($n)
            for ($i = 0; $i -lt $n; $i++) {
                $ev = $events[$i]
                $msgs[$i] = ($ev.Message -replace '\r?\n', ' ').Trim()
                $lvls[$i] = [int]$ev.Level; $comps[$i] = $ev.ProviderName -replace 'Microsoft-Windows-',''
                $eids[$i] = "$($ev.Id)"; $times[$i] = $ev.TimeCreated.ToString('M-d-yyyy HH:mm:ss'); $threads[$i] = "$($ev.ProcessId)"
            }
            [PolicyLogParser]::ClassifyMdmEntries($msgs, $lvls, $comps, $eids, $times, $threads, $Script:MdmActiveFilter, $Script:MdmContentFilter)
            if ([PolicyLogParser]::TotalLines -gt 0) {
                $Script:MdmFormattedLines.AddRange([PolicyLogParser]::FormattedLines); $Script:MdmLineTypes.AddRange([PolicyLogParser]::LineTypes); $Script:MdmRawEntries.AddRange([PolicyLogParser]::RawEntries)
                $Script:MdmStats.Lines += [PolicyLogParser]::TotalLines; $Script:MdmStats.Errors += [PolicyLogParser]::ErrorCount; $Script:MdmStats.Warnings += [PolicyLogParser]::WarningCount
            }
        } else {
            foreach ($ev in $events) {
                $parsed = @{ Message = ($ev.Message -replace '\r?\n', ' ').Trim(); Time = $ev.TimeCreated.ToString('M-d-yyyy HH:mm:ss'); Component = $ev.ProviderName -replace 'Microsoft-Windows-',''; EventId = "$($ev.Id)"; Level = [int]$ev.Level; Thread = "$($ev.ProcessId)" }
                $result = Classify-MdmLogEntry $parsed
                if ($result) { $Script:MdmFormattedLines.Add($result.Formatted); $Script:MdmLineTypes.Add($result.LineType); $Script:MdmRawEntries.Add($result.Raw) }
            }
        }
    } catch {
        if ($_.Exception.Message -match 'unauthorized|access') {
            $errP = @{ Message = "Access denied - reading MDM diagnostic event log requires administrator privileges."; Time = (Get-Date).ToString('M-d-yyyy HH:mm:ss'); Component = 'System'; EventId = ''; Level = 2; Thread = '' }
            $r = Classify-MdmLogEntry $errP; if ($r) { $Script:MdmFormattedLines.Add($r.Formatted); $Script:MdmLineTypes.Add($r.LineType); $Script:MdmRawEntries.Add($r.Raw) }
        } elseif ($_.Exception.Message -match 'No events were found') { Write-Host "[MDM] Load: No events found in $logName"
        } else { Write-Host "[MDM] Load ERROR: $($_.Exception.Message)" }
    }
    $ui.lbMdmLogs.ItemsSource = $Script:MdmFormattedLines
    $ui.lbMdmLogs.UpdateLayout()
    $mdmSv = Get-ListBoxScrollViewer $ui.lbMdmLogs
    if ($mdmSv) { $mdmSv.Add_ScrollChanged({ Apply-LogListBoxColors $ui.lbMdmLogs $Script:MdmLineTypes $Script:MdmColorMap $Script:MdmBrushCache }) }
    Apply-LogListBoxColors $ui.lbMdmLogs $Script:MdmLineTypes $Script:MdmColorMap $Script:MdmBrushCache
    if ($ui.cnvMdmHeatmap) { $ui.cnvMdmHeatmap.Visibility = 'Visible'; $ui.cnvMdmHeatmap.UpdateLayout() }
    Rebuild-LogHeatmap $ui.cnvMdmHeatmap $Script:MdmLineTypes $Script:MdmColorMap $Script:MdmBrushCache
    Rebuild-LogMinimap $ui.cnvMdmMinimap $Script:MdmLineTypes $Script:MdmColorMap $Script:MdmBrushCache
    $rendered = $Script:MdmFormattedLines.Count
    Update-MdmStatsDisplay
    $ui.MdmLogSubtitle.Text = "Loaded: $sourceLabel ($rendered entries)"
    if ($rendered -gt 0) { $ui.lbMdmLogs.ScrollIntoView($ui.lbMdmLogs.Items[$rendered - 1]) }
}

# â”€â”€ MDM DispatcherTimer â”€â”€
$Script:MdmTimer = [System.Windows.Threading.DispatcherTimer]::new()
$Script:MdmTimer.Interval = [TimeSpan]::FromMilliseconds(80)
$Script:MdmStatsThrottle = [DateTime]::MinValue
$Script:MdmEventPollLast = [DateTime]::MinValue
$Script:MdmTimer.Add_Tick({
    if (-not $Script:MdmTailing) { return }
    $now = [DateTime]::Now
    if (($now - $Script:MdmEventPollLast).TotalSeconds -ge 2) {
        $Script:MdmEventPollLast = $now
        Read-MdmEventLogDelta
    }
    $batchCount = 0; $contentAdded = $false
    $batchEntries = [System.Collections.Generic.List[hashtable]]::new()
    while ($Script:MdmLogQueue.Count -gt 0 -and $batchCount -lt $Script:MdmMaxPerTick) {
        $batchEntries.Add($Script:MdmLogQueue.Dequeue()); $batchCount++
    }
    if ($batchEntries.Count -gt 0) {
        if ($Script:HasCSharpParser) {
            $n = $batchEntries.Count
            $msgs = [string[]]::new($n); $lvls = [int[]]::new($n); $comps = [string[]]::new($n)
            $eids = [string[]]::new($n); $times = [string[]]::new($n); $threads = [string[]]::new($n)
            for ($i = 0; $i -lt $n; $i++) {
                $e = $batchEntries[$i]; $msgs[$i] = $e.Message; $lvls[$i] = [int]$e.Level
                $comps[$i] = if ($e.Component) { $e.Component } else { '' }
                $eids[$i] = if ($e.EventId) { $e.EventId } else { '' }
                $times[$i] = if ($e.Time) { $e.Time } else { '' }
                $threads[$i] = if ($e.Thread) { $e.Thread } else { '' }
            }
            [PolicyLogParser]::ClassifyMdmEntries($msgs, $lvls, $comps, $eids, $times, $threads, $Script:MdmActiveFilter, $Script:MdmContentFilter)
            if ([PolicyLogParser]::TotalLines -gt 0) {
                $Script:MdmFormattedLines.AddRange([PolicyLogParser]::FormattedLines); $Script:MdmLineTypes.AddRange([PolicyLogParser]::LineTypes); $Script:MdmRawEntries.AddRange([PolicyLogParser]::RawEntries)
                $Script:MdmStats.Lines += [PolicyLogParser]::TotalLines; $Script:MdmStats.Errors += [PolicyLogParser]::ErrorCount; $Script:MdmStats.Warnings += [PolicyLogParser]::WarningCount
                $contentAdded = $true
            }
        } else {
            foreach ($entry in $batchEntries) {
                $result = Classify-MdmLogEntry $entry
                if ($result) { $Script:MdmFormattedLines.Add($result.Formatted); $Script:MdmLineTypes.Add($result.LineType); $Script:MdmRawEntries.Add($result.Raw); $contentAdded = $true }
            }
        }
    }
    if ($contentAdded) {
        $ui.lbMdmLogs.ItemsSource = $null
        $ui.lbMdmLogs.ItemsSource = $Script:MdmFormattedLines
        $sv = Get-ListBoxScrollViewer $ui.lbMdmLogs
        if ($sv) {
            $isLatched = ($sv.VerticalOffset + $sv.ViewportHeight) -ge ($sv.ExtentHeight - 20.0)
            if ($isLatched -and $Script:MdmFormattedLines.Count -gt 0) {
                $ui.lbMdmLogs.ScrollIntoView($ui.lbMdmLogs.Items[$Script:MdmFormattedLines.Count - 1])
                if ($ui.MdmFollowIndicator) { $ui.MdmFollowIndicator.Text = [char]0x25BC + ' Following' }
            } else {
                if ($ui.MdmFollowIndicator) { $ui.MdmFollowIndicator.Text = [char]0x25A0 + ' Paused' }
            }
        }
        $ui.lbMdmLogs.UpdateLayout()
        Apply-LogListBoxColors $ui.lbMdmLogs $Script:MdmLineTypes $Script:MdmColorMap $Script:MdmBrushCache
        $now2 = [DateTime]::Now
        if (($now2 - $Script:MdmStatsThrottle).TotalMilliseconds -gt 500) { Update-MdmStatsDisplay; $Script:MdmStatsThrottle = $now2 }
    }
})
$Script:MdmTimer.Start()

# ── MDM Search ──
function Search-MdmHighlight {
    param([string]$Term)
    $Script:MdmSearchMatches.Clear()
    $Script:MdmSearchIndex = -1
    if ([string]::IsNullOrWhiteSpace($Term)) {
        if ($ui.MdmSearchCount) { $ui.MdmSearchCount.Text = '' }
        Apply-LogListBoxColors $ui.lbMdmLogs $Script:MdmLineTypes $Script:MdmColorMap $Script:MdmBrushCache
        return
    }
    $rx = [regex]::new([regex]::Escape($Term), 'IgnoreCase')
    for ($i = 0; $i -lt $Script:MdmRawEntries.Count; $i++) {
        if ($rx.IsMatch($Script:MdmRawEntries[$i]) -or $rx.IsMatch($Script:MdmFormattedLines[$i])) {
            [void]$Script:MdmSearchMatches.Add($i)
        }
    }
    if ($ui.MdmSearchCount) { $ui.MdmSearchCount.Text = "$($Script:MdmSearchMatches.Count) matches" }
    if ($Script:MdmSearchMatches.Count -gt 0) { $Script:MdmSearchIndex = 0; Navigate-MdmSearchMatch 0 }
}

function Navigate-MdmSearchMatch([int]$Index) {
    if ($Index -lt 0 -or $Index -ge $Script:MdmSearchMatches.Count) { return }
    $Script:MdmSearchIndex = $Index
    $lineIdx = $Script:MdmSearchMatches[$Index]
    $ui.lbMdmLogs.ScrollIntoView($ui.lbMdmLogs.Items[$lineIdx])
    $ui.lbMdmLogs.SelectedIndex = $lineIdx
    if ($ui.MdmSearchCount) { $ui.MdmSearchCount.Text = "$($Index+1)/$($Script:MdmSearchMatches.Count)" }
}

# ── MDM Admin Banner Init ──
if (-not $Script:IsAdminSession) {
    if ($ui.MdmAdminBanner) { $ui.MdmAdminBanner.Visibility = 'Visible' }
    if ($ui.MdmAdminBannerText) { $ui.MdmAdminBannerText.Text = "Not running as admin - MDM diagnostic event logs require elevation" }
}

# ── Wire up MDM Sync Nav ──
$ui.NavMDMSync.Add_Click({
    Switch-Tab 'MDMSync'
    if ($ui.ChkMdmLiveTail -and $ui.ChkMdmLiveTail.IsChecked) {
        if (-not $Script:MdmTailing) { Write-Host '[MDM] Auto-starting tail on tab switch'; Start-MdmTail }
    } else {
        if ($Script:MdmStats.Lines -eq 0) { Write-Host '[MDM] Auto-loading log on tab switch'; Load-MdmLogFile }
    }
})

# ── Wire up MDM Log UI events ──
if ($ui.BtnMdmWrapToggle) {
    $ui.BtnMdmWrapToggle.Add_Click({
        $scroll = $ui.lbMdmLogs.GetValue([System.Windows.Controls.ScrollViewer]::HorizontalScrollBarVisibilityProperty)
        if ($scroll -eq 'Disabled') { $ui.lbMdmLogs.SetValue([System.Windows.Controls.ScrollViewer]::HorizontalScrollBarVisibilityProperty, [System.Windows.Controls.ScrollBarVisibility]::Auto) }
        else { $ui.lbMdmLogs.SetValue([System.Windows.Controls.ScrollViewer]::HorizontalScrollBarVisibilityProperty, [System.Windows.Controls.ScrollBarVisibility]::Disabled) }
    })
}
if ($ui.BtnMdmClearLog) { $ui.BtnMdmClearLog.Add_Click({ Clear-MdmLogDisplay }) }

# Mode toggle: Live tail vs Analysis
if ($ui.ChkMdmLiveTail) {
    $ui.ChkMdmLiveTail.Add_Checked({
        if ($ui.MdmModeHint) { $ui.MdmModeHint.Text = 'Streaming new MDM diagnostic entries in real time' }
        if ($Script:MdmTailing) { return }
        Start-MdmTail
    })
    $ui.ChkMdmLiveTail.Add_Unchecked({
        if ($ui.MdmModeHint) { $ui.MdmModeHint.Text = 'Full log loaded for analysis and search' }
        if ($Script:MdmTailing) { Stop-MdmTail }
        Load-MdmLogFile
    })
}

# Log source change: restart current mode
if ($ui.CmbMdmLogSource) {
    $ui.CmbMdmLogSource.Add_SelectionChanged({
        if ($ui.ChkMdmLiveTail -and $ui.ChkMdmLiveTail.IsChecked) {
            if ($Script:MdmTailing) { Stop-MdmTail }
            Start-MdmTail
        } else {
            Load-MdmLogFile
        }
    })
}

# Search with 300ms debounce
if ($ui.TxtMdmSearch) {
    $Script:MdmSearchDebounce = $null
    $ui.TxtMdmSearch.Add_TextChanged({
        if ($Script:MdmSearchDebounce) { $Script:MdmSearchDebounce.Stop() }
        $Script:MdmSearchDebounce = [System.Windows.Threading.DispatcherTimer]::new()
        $Script:MdmSearchDebounce.Interval = [TimeSpan]::FromMilliseconds(300)
        $Script:MdmSearchDebounce.Add_Tick({
            $Script:MdmSearchDebounce.Stop()
            Search-MdmHighlight $ui.TxtMdmSearch.Text
        })
        $Script:MdmSearchDebounce.Start()
    })
}
if ($ui.BtnMdmSearchNext) {
    $ui.BtnMdmSearchNext.Add_Click({
        if ($Script:MdmSearchMatches.Count -gt 0) {
            $next = ($Script:MdmSearchIndex + 1) % $Script:MdmSearchMatches.Count
            Navigate-MdmSearchMatch $next
        }
    })
}
if ($ui.BtnMdmSearchPrev) {
    $ui.BtnMdmSearchPrev.Add_Click({
        if ($Script:MdmSearchMatches.Count -gt 0) {
            $prev = $Script:MdmSearchIndex - 1
            if ($prev -lt 0) { $prev = $Script:MdmSearchMatches.Count - 1 }
            Navigate-MdmSearchMatch $prev
        }
    })
}

# MDM Severity filter pills
foreach ($filterName in @('FilterMdmAll','FilterMdmError','FilterMdmWarning','FilterMdmInfo')) {
    if ($ui[$filterName]) {
        $ui[$filterName].Add_Click({
            param($sender)
            $label = $sender.Content
            $Script:MdmActiveFilter = switch ($label) {
                'All' { 'All' }; 'Errors' { 'Error' }; 'Warnings' { 'Warning' }; 'Info' { 'Info' }; default { 'All' }
            }
            foreach ($fn in @('FilterMdmAll','FilterMdmError','FilterMdmWarning','FilterMdmInfo')) {
                if ($ui[$fn]) { $ui[$fn].Tag = if ($ui[$fn] -eq $sender) { 'Active' } else { $null } }
            }
            if ($ui.ChkMdmLiveTail -and $ui.ChkMdmLiveTail.IsChecked) {
                if ($Script:MdmTailing) {
                    Stop-MdmTail
                    Clear-MdmLogDisplay
                    $Script:MdmLastEventTime = $null
                    $Script:MdmTailing = $true
                    Read-MdmEventLogDelta
                }
            } else {
                Load-MdmLogFile
            }
        })
    }
}

# Minimap click → scroll to position
if ($ui.cnvMdmMinimap) {
    $ui.cnvMdmMinimap.Add_MouseLeftButtonDown({
        param($sender, $e)
        $pos = $e.GetPosition($sender)
        $ratio = $pos.Y / [Math]::Max($sender.ActualHeight, 1)
        $sv = Get-ListBoxScrollViewer $ui.lbMdmLogs
        if ($sv) { $sv.ScrollToVerticalOffset($ratio * $sv.ExtentHeight) }
    })
}

# Cleanup on window close
$Window.Add_Closed({
    if ($Script:ImeTailing) { Stop-ImeTail }
    if ($Script:ImeTimer) { $Script:ImeTimer.Stop() }
    if ($Script:GpoTailing) { Stop-GpoTail }
    if ($Script:GpoTimer) { $Script:GpoTimer.Stop() }
    if ($Script:MdmTailing) { Stop-MdmTail }
    if ($Script:MdmTimer) { $Script:MdmTimer.Stop() }
    # Stop ETW trace if active
    if ($Script:EtwTraceActive) {
        try { & logman stop $Script:EtwSessionName -ets 2>$null | Out-Null } catch {}
    }
    # Close background log if active
    if ($Script:BgLogActive -and $Script:BgLogStream) {
        try { $Script:BgLogStream.Flush(); $Script:BgLogStream.Close(); $Script:BgLogStream.Dispose() } catch {}
        $Script:BgLogActive = $false
    }
})

$Window.Add_PreviewKeyDown({
    if ($_.Key -eq 'F1') { if ($ui.BtnHelp) { $ui.BtnHelp.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) }; $_.Handled = $true }
    if ($_.Key -eq 'F5') { $ui.BtnScanGPOs.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)); $_.Handled = $true }
    if ($_.Key -eq 'L' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'Control') { if ($ui.BtnThemeToggle) { $ui.BtnThemeToggle.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) }; $_.Handled = $true }
    if ($_.Key -eq 'D1' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'Control') { Switch-Tab 'Dashboard'; $_.Handled = $true }
    if ($_.Key -eq 'D2' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'Control') { Switch-Tab 'GPOList'; $_.Handled = $true }
    if ($_.Key -eq 'D3' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'Control') { Switch-Tab 'Settings'; $_.Handled = $true }
    if ($_.Key -eq 'D4' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'Control') { Switch-Tab 'Conflicts'; $_.Handled = $true }
    if ($_.Key -eq 'D5' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'Control') { Switch-Tab 'IntuneApps'; $_.Handled = $true }
    if ($_.Key -eq 'D6' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'Control') { Switch-Tab 'Report'; $_.Handled = $true }
    if ($_.Key -eq 'R' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'Control') {
        $ui.BtnScanGPOs.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
        $_.Handled = $true
    }
    if ($_.Key -eq 'E' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'Control') { Export-Html; $_.Handled = $true }
    if ($_.Key -eq 'S' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'Control') {
        $ui.BtnSaveSnapshot.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
        $_.Handled = $true
    }
    if ($_.Key -eq 'P' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'Control') { Invoke-PrintReport; $_.Handled = $true }
    if ($_.Key -eq 'D' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'Control') { if ($ui.BtnCompareSnapshot) { $ui.BtnCompareSnapshot.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) }; $_.Handled = $true }
    if ($_.Key -eq 'D7' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'Control') { Switch-Tab 'IMELogs'; $_.Handled = $true }
    if ($_.Key -eq 'D8' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'Control') { Switch-Tab 'GPOLogs'; $_.Handled = $true }
    if ($_.Key -eq 'D9' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'Control') { Switch-Tab 'MDMSync'; $_.Handled = $true }
    if ($_.Key -eq 'D0' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'Control') { Switch-Tab 'Tools'; $_.Handled = $true }
    if ($_.Key -eq 'OemComma' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'Control') { Switch-Tab 'AppSettings'; $_.Handled = $true }
    if ($_.Key -eq 'B' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'Control') { $ui.BtnHamburger.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)); $_.Handled = $true }
})

# ═══════════════════════════════════════════════════════════════════════════════

# ── MDM Content Filter Handlers ──
$Script:MdmPresets = @{
    'Sync'       = '/sync|session|check.in|heartbeat/i'
    'CSP'        = '/CSP|configuration service|./Device/|./User/|OMA-URI/i'
    'Certs'      = '/certificate|SCEP|cert enroll|PKCS|WiFi cert/i'
    'Compliance'  = '/compliance|compliant|non.compliant|policy eval/i'
    'Enrollment'  = '/enroll|unenroll|join|register|MDM.*registration/i'
    'Wipe/Reset'  = '/wipe|reset|retire|factory|selective wipe/i'
    'BitLocker'   = '/bitlocker|encrypt|recovery key|TPM/i'
}

function Apply-MdmContentFilter([string]$filterText) {
    $Script:MdmContentFilter = $filterText
    if ($ui.TxtMdmContentFilter) { $ui.TxtMdmContentFilter.Text = $filterText }
    if ($Script:MdmTailing) {
        $Script:MdmLastEventTime = $null
        Clear-MdmLogDisplay
        Read-MdmEventLogDelta
    } else {
        Load-MdmLogFile
    }
}

foreach ($presetName in @('PresetMdmNone','PresetMdmSync','PresetMdmCSP','PresetMdmCert','PresetMdmCompliance','PresetMdmEnroll','PresetMdmWipe','PresetMdmBitLocker')) {
    if ($ui[$presetName]) {
        $ui[$presetName].Add_Click({
            param($sender)
            $label = $sender.Content
            foreach ($pn in @('PresetMdmNone','PresetMdmSync','PresetMdmCSP','PresetMdmCert','PresetMdmCompliance','PresetMdmEnroll','PresetMdmWipe','PresetMdmBitLocker')) {
                if ($ui[$pn]) { $ui[$pn].Tag = if ($ui[$pn] -eq $sender) { 'Active' } else { $null } }
            }
            if ($label -eq 'None') { Apply-MdmContentFilter '' } else { Apply-MdmContentFilter ($Script:MdmPresets[$label]) }
        })
    }
}

if ($ui.TxtMdmContentFilter) {
    $Script:MdmContentFilterTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $Script:MdmContentFilterTimer.Interval = [TimeSpan]::FromMilliseconds(500)
    $Script:MdmContentFilterTimer.Add_Tick({
        $Script:MdmContentFilterTimer.Stop()
        Apply-MdmContentFilter $ui.TxtMdmContentFilter.Text
    })
    $ui.TxtMdmContentFilter.Add_TextChanged({
        $Script:MdmContentFilterTimer.Stop()
        $Script:MdmContentFilterTimer.Start()
        foreach ($pn in @('PresetMdmNone','PresetMdmSync','PresetMdmCSP','PresetMdmCert','PresetMdmCompliance','PresetMdmEnroll','PresetMdmWipe','PresetMdmBitLocker')) {
            if ($ui[$pn]) { $ui[$pn].Tag = $null }
        }
        if (-not $ui.TxtMdmContentFilter.Text) { if ($ui.PresetMdmNone) { $ui.PresetMdmNone.Tag = 'Active' } }
    })
}

if ($ui.BtnMdmSaveFilter) {
    $ui.BtnMdmSaveFilter.Add_Click({
        $filterText = $ui.TxtMdmContentFilter.Text
        if (-not $filterText) { return }
        $name = Show-ThemedInputBox -Prompt 'Name for this filter:' -Title 'Save Filter' -DefaultValue $filterText
        if ($name) {
            $Script:MdmSavedFilters[$name] = $filterText
            [void]$ui.CmbMdmSavedFilters.Items.Add($name)
            $ui.CmbMdmSavedFilters.SelectedItem = $name
        }
    })
}

if ($ui.CmbMdmSavedFilters) {
    $ui.CmbMdmSavedFilters.Add_SelectionChanged({
        $sel = $ui.CmbMdmSavedFilters.SelectedItem
        if ($sel -and $Script:MdmSavedFilters.ContainsKey($sel)) { Apply-MdmContentFilter $Script:MdmSavedFilters[$sel] }
    })
}

# ── MDM Content Filter Handlers ──
$Script:MdmPresets = @{
    'Sync'       = '/sync|session|check.in|heartbeat/i'
    'CSP'        = '/CSP|configuration service|./Device/|./User/|OMA-URI/i'
    'Certs'      = '/certificate|SCEP|cert enroll|PKCS|WiFi cert/i'
    'Compliance'  = '/compliance|compliant|non.compliant|policy eval/i'
    'Enrollment'  = '/enroll|unenroll|join|register|MDM.*registration/i'
    'Wipe/Reset'  = '/wipe|reset|retire|factory|selective wipe/i'
    'BitLocker'   = '/bitlocker|encrypt|recovery key|TPM/i'
}

function Apply-MdmContentFilter([string]$filterText) {
    $Script:MdmContentFilter = $filterText
    if ($ui.TxtMdmContentFilter) { $ui.TxtMdmContentFilter.Text = $filterText }
    if ($Script:MdmTailing) {
        $Script:MdmLastEventTime = $null
        Clear-MdmLogDisplay
        Read-MdmEventLogDelta
    } else {
        Load-MdmLogFile
    }
}

foreach ($presetName in @('PresetMdmNone','PresetMdmSync','PresetMdmCSP','PresetMdmCert','PresetMdmCompliance','PresetMdmEnroll','PresetMdmWipe','PresetMdmBitLocker')) {
    if ($ui[$presetName]) {
        $ui[$presetName].Add_Click({
            param($sender)
            $label = $sender.Content
            foreach ($pn in @('PresetMdmNone','PresetMdmSync','PresetMdmCSP','PresetMdmCert','PresetMdmCompliance','PresetMdmEnroll','PresetMdmWipe','PresetMdmBitLocker')) {
                if ($ui[$pn]) { $ui[$pn].Tag = if ($ui[$pn] -eq $sender) { 'Active' } else { $null } }
            }
            if ($label -eq 'None') { Apply-MdmContentFilter '' } else { Apply-MdmContentFilter ($Script:MdmPresets[$label]) }
        })
    }
}

if ($ui.TxtMdmContentFilter) {
    $Script:MdmContentFilterTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $Script:MdmContentFilterTimer.Interval = [TimeSpan]::FromMilliseconds(500)
    $Script:MdmContentFilterTimer.Add_Tick({
        $Script:MdmContentFilterTimer.Stop()
        Apply-MdmContentFilter $ui.TxtMdmContentFilter.Text
    })
    $ui.TxtMdmContentFilter.Add_TextChanged({
        $Script:MdmContentFilterTimer.Stop()
        $Script:MdmContentFilterTimer.Start()
        foreach ($pn in @('PresetMdmNone','PresetMdmSync','PresetMdmCSP','PresetMdmCert','PresetMdmCompliance','PresetMdmEnroll','PresetMdmWipe','PresetMdmBitLocker')) {
            if ($ui[$pn]) { $ui[$pn].Tag = $null }
        }
        if (-not $ui.TxtMdmContentFilter.Text) { if ($ui.PresetMdmNone) { $ui.PresetMdmNone.Tag = 'Active' } }
    })
}

if ($ui.BtnMdmSaveFilter) {
    $ui.BtnMdmSaveFilter.Add_Click({
        $filterText = $ui.TxtMdmContentFilter.Text
        if (-not $filterText) { return }
        $name = Show-ThemedInputBox -Prompt 'Name for this filter:' -Title 'Save Filter' -DefaultValue $filterText
        if ($name) {
            $Script:MdmSavedFilters[$name] = $filterText
            [void]$ui.CmbMdmSavedFilters.Items.Add($name)
            $ui.CmbMdmSavedFilters.SelectedItem = $name
        }
    })
}

if ($ui.CmbMdmSavedFilters) {
    $ui.CmbMdmSavedFilters.Add_SelectionChanged({
        $sel = $ui.CmbMdmSavedFilters.SelectedItem
        if ($sel -and $Script:MdmSavedFilters.ContainsKey($sel)) { Apply-MdmContentFilter $Script:MdmSavedFilters[$sel] }
    })
}
# SECTION 20: WINDOW CHROME & CLEANUP
# ═══════════════════════════════════════════════════════════════════════════════

$ui.TitleBar.Add_MouseLeftButtonDown({
    if ($_.ClickCount -eq 2) {
        if ($Window.WindowState -eq 'Maximized') { $Window.WindowState = 'Normal' }
        else { $Window.WindowState = 'Maximized' }
    } else { $Window.DragMove() }
})

$ui.BtnMinimize.Add_Click({ $Window.WindowState = 'Minimized' })
$ui.BtnMaximize.Add_Click({
    if ($Window.WindowState -eq 'Maximized') { $Window.WindowState = 'Normal' }
    else { $Window.WindowState = 'Maximized' }
})
$ui.BtnClose.Add_Click({ $Window.Close() })

$Window.Add_StateChanged({
    if ($Window.WindowState -eq 'Maximized') {
        $ui.MaximizeIcon.Text = [char]0xE923
        $ui.WindowBorder.CornerRadius = [System.Windows.CornerRadius]::new(0)
        $ui.WindowBorder.BorderThickness = [System.Windows.Thickness]::new(0)
        # Constrain to work area so window doesn't extend behind taskbar
        $wa = [System.Windows.SystemParameters]::WorkArea
        $Window.MaxHeight = $wa.Height + 14
        $Window.MaxWidth  = $wa.Width + 14
        $ui.WindowBorder.Margin = [System.Windows.Thickness]::new(7)
    } else {
        $ui.MaximizeIcon.Text = [char]0xE922
        $ui.WindowBorder.CornerRadius = [System.Windows.CornerRadius]::new(8)
        $ui.WindowBorder.BorderThickness = [System.Windows.Thickness]::new(1)
        $Window.MaxHeight = [double]::PositiveInfinity
        $Window.MaxWidth  = [double]::PositiveInfinity
        $ui.WindowBorder.Margin = [System.Windows.Thickness]::new(0)
    }
})

# Entry animation
$Window.Add_ContentRendered({
    if ($Script:AnimationsDisabled) {
        $ui.WindowBorder.Opacity = 1
    } else {
        $ui.WindowBorder.RenderTransform = [System.Windows.Media.TranslateTransform]::new(0, 10)
        $fadeIn = [System.Windows.Media.Animation.DoubleAnimation]::new(0, 1,
            [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(250)))
        $fadeIn.EasingFunction = [System.Windows.Media.Animation.CubicEase]::new()
        $fadeIn.EasingFunction.EasingMode = 'EaseOut'
        $ui.WindowBorder.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $fadeIn)
        $slideUp = [System.Windows.Media.Animation.DoubleAnimation]::new(10, 0,
            [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(250)))
        $slideUp.EasingFunction = [System.Windows.Media.Animation.CubicEase]::new()
        $slideUp.EasingFunction.EasingMode = 'EaseOut'
        $ui.WindowBorder.RenderTransform.BeginAnimation(
            [System.Windows.Media.TranslateTransform]::YProperty, $slideUp)
    }

    $Global:ConsoleReady = $true

    # Open console panel on launch so user sees debug output
    if ($ui.pnlBottomPanel) { $ui.pnlBottomPanel.Visibility = 'Visible' }
    Write-DebugLog "PolicyPilot ready" -Level SYSTEM
    Write-DebugLog "Scan mode: $($Script:Prefs.ScanMode) | Theme: $(if ($Script:Prefs.IsLightMode) {'Light'} else {'Dark'})" -Level INFO

    # Check for a cached scan snapshot and offer to restore
    $snapshotPath = [IO.Path]::Combine($env:TEMP, 'PolicyPilot_GPOCache', 'last_scan.clixml')
    if (Test-Path $snapshotPath) {
        $snapAge = [DateTime]::Now - (Get-Item $snapshotPath).LastWriteTime
        if ($snapAge.TotalHours -le 24) {
            $ageLabel = if ($snapAge.TotalMinutes -lt 60) { "$([math]::Round($snapAge.TotalMinutes))m ago" } else { "$([math]::Round($snapAge.TotalHours,1))h ago" }
            Write-DebugLog "Found scan snapshot ($ageLabel) — auto-restoring" -Level INFO
            Restore-LastScan | Out-Null
        }
    }

    Render-Achievements
    if ($Script:Prefs.AchievementsCollapsed) {
        $ui.pnlAchievements.Visibility = 'Collapsed'
        $ui.txtAchievementChevron.RenderTransform = [System.Windows.Media.RotateTransform]::new(180)
    }
    # Auto-check prerequisites on startup (background to avoid blocking UI)
    if ($ui.PrereqDetailStatus) { $ui.PrereqDetailStatus.Text = 'Checking prerequisites...' }
    if ($ui.PrereqStatus)       { $ui.PrereqStatus.Text = 'Checking...' }
    Start-BackgroundWork -Work {
        param($SyncH)
        $results = [System.Collections.Generic.List[string]]::new()
        $allOk = $true
        $rsatMissing = $false
        $mode = $SyncH.ScanMode

        [void]$results.Add("Scan Mode: $mode")
        [void]$results.Add("")

        if ($mode -eq 'AD') {
            $gpMod = Get-Module -ListAvailable -Name GroupPolicy -ErrorAction SilentlyContinue
            if ($gpMod) {
                [void]$results.Add("[OK]  GroupPolicy module v$($gpMod.Version)")
            } else {
                [void]$results.Add("[FAIL] GroupPolicy module NOT found")
                [void]$results.Add("       Install via Features on Demand:")
                [void]$results.Add("       Add-WindowsCapability -Online -Name Rsat.GroupPolicy.Management.Tools~~~~0.0.1.0")
                [void]$results.Add("       Or: Settings > System > Optional Features > RSAT: Group Policy Management Tools")
                [void]$results.Add("")
                [void]$results.Add("       [TIP] PolicyPilot can install this for you - click 'Install RSAT GP Tools' below")
                $allOk = $false
                $rsatMissing = $true
            }
            try {
                $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name
                [void]$results.Add("[OK]  Domain: $domain")
            } catch {
                [void]$results.Add("[FAIL] Cannot reach Active Directory domain")
                [void]$results.Add("       Ensure this machine is domain-joined or specify a domain override")
                $allOk = $false
            }
            if ($gpMod) {
                try {
                    Import-Module GroupPolicy -ErrorAction Stop
                    $testGPO = Get-GPO -All -ErrorAction Stop | Select-Object -First 1
                    if ($testGPO) { [void]$results.Add("[OK]  GPO read access confirmed") }
                    else { [void]$results.Add("[WARN] Get-GPO returned 0 GPOs - domain may be empty") }
                } catch {
                    [void]$results.Add("[FAIL] Cannot read GPOs: $($_.Exception.Message)")
                    $allOk = $false
                }
            }
        } elseif ($mode -eq 'Intune') {
            $graphMod = Get-Module -ListAvailable Microsoft.Graph.DeviceManagement -ErrorAction SilentlyContinue
            $graphLegacy = Get-Module -ListAvailable Microsoft.Graph.Intune -ErrorAction SilentlyContinue
            if ($graphMod -or $graphLegacy) { [void]$results.Add("[OK]  Microsoft.Graph module found") }
            else { [void]$results.Add("[FAIL] Microsoft.Graph.DeviceManagement not found"); [void]$results.Add("  Install: Install-Module Microsoft.Graph -Scope CurrentUser"); $allOk = $false }
            [void]$results.Add(""); [void]$results.Add("Intune mode uses Microsoft Graph to read"); [void]$results.Add("device configuration, compliance policies,"); [void]$results.Add("and settings catalog.")
        } else {
            $gpresult = Get-Command gpresult.exe -ErrorAction SilentlyContinue
            if ($gpresult) { [void]$results.Add("[OK]  gpresult.exe found") }
            else { [void]$results.Add("[FAIL] gpresult.exe not found"); $allOk = $false }
            try {
                $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
                if ($cs.PartOfDomain) { [void]$results.Add("[OK]  Domain-joined: $($cs.Domain)") }
                else { [void]$results.Add("[WARN] Not domain-joined - only local policies will be shown") }
            } catch { [void]$results.Add("[WARN] Cannot determine domain membership") }
            [void]$results.Add(""); [void]$results.Add("Local mode scans policies applied to THIS machine"); [void]$results.Add("using gpresult (no RSAT required).")
        }
        return @{ Passed = $allOk; Details = ($results -join "`n"); RsatMissing = $rsatMissing }
    } -OnComplete {
        param($Results, $Errors)
        $r = $Results | Select-Object -First 1
        if ($r) {
            $Script:PrereqsMet = $r.Passed
            $Script:RsatMissing = $r.RsatMissing
            if ($ui.PrereqDetailStatus) { $ui.PrereqDetailStatus.Text = $r.Details }
            if ($ui.PrereqStatus)       { $ui.PrereqStatus.Text = $r.Details }
            Write-DebugLog "Prereqs ($($Script:Prefs.ScanMode)): $(if ($r.Passed) {'PASSED'} else {'FAILED'})" -Level $(if ($r.Passed) {'SUCCESS'} else {'ERROR'})
        } else {
            $errMsg = if ($Errors.Count -gt 0) { $Errors[0].ToString() } else { 'Unknown error' }
            if ($ui.PrereqDetailStatus) { $ui.PrereqDetailStatus.Text = "Prerequisite check failed: $errMsg" }
            Write-DebugLog "Prereq check error: $errMsg" -Level ERROR
        }
    }.GetNewClosure() -Variables @{ ScanMode = $Script:Prefs.ScanMode } -Context @{ Name = 'PrereqCheck' }
}.GetNewClosure())

# Save on close
$Window.Add_Closing({
    Write-DebugLog 'Application closing - saving preferences' -Level SYSTEM
    if ($Script:BgTimer) { $Script:BgTimer.Stop() }
    foreach ($j in @($Global:BgJobs)) { try { $j.PS.Stop(); $j.PS.Dispose(); $j.Runspace.Dispose() } catch { try { Write-DebugLog "Unhandled: $_" -Level ERROR } catch {} } }
    $Global:BgJobs.Clear()
    Save-Preferences
})

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 21: INITIALIZE & SHOW
# ═══════════════════════════════════════════════════════════════════════════════

# Apply saved domain/DC/OU overrides to UI
if ($Script:Prefs.DomainOverride) { $ui.TxtDomainOverride.Text = $Script:Prefs.DomainOverride }
if ($Script:Prefs.DCOverride)     { $ui.TxtDCOverride.Text = $Script:Prefs.DCOverride }
if ($Script:Prefs.OUScope -and $ui.TxtOUScope) { $ui.TxtOUScope.Text = $Script:Prefs.OUScope }
if ($Script:Prefs.ForceRefresh -and $ui.ChkForceRefresh) { $ui.ChkForceRefresh.IsChecked = $true }

# Initialize export option checkboxes from prefs
$ui.ChkIncludeDisabled.IsChecked   = $Script:Prefs.IncludeDisabled
$ui.ChkIncludeUnlinked.IsChecked   = $Script:Prefs.IncludeUnlinked
$ui.ChkShowRegistryPaths.IsChecked = $Script:Prefs.ShowRegistryPaths

# Set initial theme pill state
if ($Script:Prefs.IsLightMode) {
    $ui.BtnThemeLight.Tag = 'Active'; $ui.BtnThemeDark.Tag = $null
} else {
    $ui.BtnThemeDark.Tag = 'Active'; $ui.BtnThemeLight.Tag = $null
}

# Set theme toggle button icon
if ($ui.BtnThemeToggle) { $ui.BtnThemeToggle.Content = if ($Script:Prefs.IsLightMode) { [char]0x263E } else { [char]0x2600 } }


# Set initial scan mode from prefs
switch ($Script:Prefs.ScanMode) {
    'Intune' { $ui.CmbScanMode.SelectedIndex = 0 }
    'Local'  { $ui.CmbScanMode.SelectedIndex = 1 }
    'AD'     { $ui.CmbScanMode.SelectedIndex = 2 }
    'Combined' { $ui.CmbScanMode.SelectedIndex = 3 }
    default  { $ui.CmbScanMode.SelectedIndex = 0 }
}
# Set initial tab
Switch-Tab 'Dashboard'

Write-DebugLog "PolicyPilot v$($Script:AppVersion) started" -Level SYSTEM


# -- DispatcherTimer: polls background jobs every 50ms --
$Script:BgTimer = New-Object System.Windows.Threading.DispatcherTimer
$Script:BgTimer.Interval = [TimeSpan]::FromMilliseconds($Script:TIMER_INTERVAL_MS)
$Script:BgTimer.Add_Tick({
    if ($Global:TimerProcessing) { return }
    $Global:TimerProcessing = $true
    try {
        for ($bi = $Global:BgJobs.Count - 1; $bi -ge 0; $bi--) {
            $Job = $Global:BgJobs[$bi]
            if ($Job.AsyncResult.IsCompleted) {
                $elapsed = [math]::Round(((Get-Date) - $Job.StartedAt).TotalSeconds, 1)
                Write-DebugLog "BgJob[$bi]: COMPLETED after ${elapsed}s" -Level DEBUG
                try {
                    $BgResult = $Job.PS.EndInvoke($Job.AsyncResult)
                    Write-DebugLog "BgJob[$bi]: EndInvoke returned $($BgResult.Count) result(s), Errors=$($Job.PS.Streams.Error.Count) Warnings=$($Job.PS.Streams.Warning.Count) Verbose=$($Job.PS.Streams.Verbose.Count)" -Level DEBUG
                    $BgErrors = @($Job.PS.Streams.Error)
                    if ($BgErrors.Count -gt 0) {
                        foreach ($e in $BgErrors) {
                            Write-DebugLog "BgJob[$bi] ERROR: $($e.ToString())" -Level ERROR
                        }
                    }
                    if ($Job.PS.Streams.Warning.Count -gt 0) { foreach ($w in $Job.PS.Streams.Warning) { Write-DebugLog "BgJob[$bi] WARN: $($w.ToString())" -Level WARN } }
                    & $Job.OnComplete $BgResult $BgErrors $Job.Context
                } catch {
                    Write-Host "[BGJOB] Callback EXCEPTION: $($_.Exception.Message)"
                    if ($_.Exception.InnerException) {
                        Write-Host "[BGJOB]   Inner: $($_.Exception.InnerException.Message)"
                        if ($_.Exception.InnerException.InnerException) {
                            Write-Host "[BGJOB]   Inner2: $($_.Exception.InnerException.InnerException.Message)"
                        }
                    }
                    Write-DebugLog "BgJob[$bi] callback EXCEPTION: $($_.Exception.Message)" -Level ERROR
                }
                try { $Job.PS.Dispose() } catch { }
                try { $Job.Runspace.Dispose() } catch { }
                $Global:BgJobs.RemoveAt($bi)
            } else {
                $elapsed = [math]::Round(((Get-Date) - $Job.StartedAt).TotalSeconds, 1)
                $sec = [math]::Floor($elapsed / 5)
                if ($sec -ge 1 -and $sec -ne $Job.LastBucket) {
                    $Job.LastBucket = $sec
                    Write-DebugLog "BgJob[$bi] '$($Job.Context.Name)': still running (${elapsed}s)..." -Level DEBUG
                }
                # Update global progress label if available
                if ($ui.lblGlobalProgress -and $Job.Context.Name) {
                    $ui.lblGlobalProgress.Text = "$($Job.Context.Name) running (${elapsed}s)..."
                }
            }
        }
        # Drain status queue
        while ($Global:SyncHash.StatusQueue.Count -gt 0) {
            $msg = $Global:SyncHash.StatusQueue.Dequeue()
            if ($msg.Type -eq 'Status') {
                Set-Status $msg.Text $msg.Color
            }
            if ($msg.Type -eq 'Progress') {
                if ($ui.ScanProgressBar) { $ui.ScanProgressBar.Value = $msg.Value }
                if ($ui.ScanProgressText) { $ui.ScanProgressText.Text = "$($msg.Value)%" }
            }
            if ($msg.Type -eq 'Log') {
                Write-DebugLog $msg.Text -Level $msg.Level
            }
        }
    } finally {
        $Global:TimerProcessing = $false
    }
})
$Script:BgTimer.Start()
Write-DebugLog "DispatcherTimer started (${Script:TIMER_INTERVAL_MS}ms)" -Level SYSTEM

# Set initial MaxHeight for Maximized start (WindowStyle=None needs explicit constraint)
$wa = [System.Windows.SystemParameters]::WorkArea
$Window.MaxHeight = $wa.Height + 14
$Window.MaxWidth  = $wa.Width + 14

# Start on top, then release so other windows can overlay
$Window.Topmost = $true
$WindowRef = $Window
$Window.Add_ContentRendered({ $WindowRef.Topmost = $false }.GetNewClosure())

# Catch WPF rendering/dispatcher exceptions that escape try/catch blocks
$Window.Dispatcher.Add_UnhandledException({
    param($sender, $e)
    $ex = $e.Exception
    Write-Host "`n[WPF-CRASH] Dispatcher UnhandledException!" -ForegroundColor Red
    Write-Host "[WPF-CRASH] Type: $($ex.GetType().FullName)" -ForegroundColor Red
    Write-Host "[WPF-CRASH] Message: $($ex.Message)" -ForegroundColor Red
    Write-Host "[WPF-CRASH] Source: $($ex.Source)" -ForegroundColor Red
    $inner = $ex.InnerException
    $depth = 0
    while ($inner -and $depth -lt 5) {
        $depth++
        Write-Host "[WPF-CRASH]   Inner[$depth]: $($inner.GetType().FullName) - $($inner.Message)" -ForegroundColor Yellow
        $inner = $inner.InnerException
    }
    Write-Host "[WPF-CRASH] StackTrace:`n$($ex.StackTrace)" -ForegroundColor DarkGray
    Write-DebugLog "WPF DISPATCHER CRASH: $($ex.Message)" -Level ERROR
    if ($ex.InnerException) { Write-DebugLog "  Inner: $($ex.InnerException.Message)" -Level ERROR }
    $e.Handled = $true
})

# ===============================================================================
# HEADLESS MODE - generate report without showing UI
# ===============================================================================
if ($Headless) {
    if (-not $ReportType) {
        Write-Host "[PolicyPilot] ERROR: -ReportType is required in headless mode (Local, AD, Intune, Combined)" -ForegroundColor Red
        Write-Host "Usage: .\PolicyPilot.ps1 -Headless -ReportType <Local|AD|Intune|Combined> [-OutputPath <path>]" -ForegroundColor Yellow
        exit 1
    }
    if ($ReportType -notin @('Local','AD','Intune','Combined')) {
        Write-Host "[PolicyPilot] ERROR: Invalid -ReportType '$ReportType'. Must be: Local, AD, Intune, or Combined" -ForegroundColor Red
        exit 1
    }

    $Script:Prefs.ScanMode = $ReportType
    Write-Host "[PolicyPilot] Headless mode - generating $ReportType policy report..." -ForegroundColor Cyan

    # Determine output path
    if (-not $OutputPath) {
        $OutputPath = Join-Path $Script:ReportsDir "PolicyPilot_${ReportType}_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    }
    if (-not [System.IO.Path]::IsPathRooted($OutputPath)) {
        $OutputPath = Join-Path (Get-Location).Path $OutputPath
    }
    $outDir = Split-Path -Parent $OutputPath
    if ($outDir -and -not (Test-Path $outDir)) { New-Item -Path $outDir -ItemType Directory -Force | Out-Null }

    # Prerequisites check
    $prereq = Test-Prerequisites
    if (-not $prereq.Passed) {
        Write-Host "[PolicyPilot] Prerequisite check FAILED:" -ForegroundColor Red
        Write-Host $prereq.Details
        exit 1
    }
    Write-Host "[PolicyPilot] Prerequisites OK" -ForegroundColor Green

    # Run scan
    Write-Host "[PolicyPilot] Running $ReportType scan..." -ForegroundColor Cyan
    $scanResult = $null
    switch ($ReportType) {
        'Local' {
            $scanResult = Invoke-LocalRSoPScan
        }
        'AD' {
            $scanResult = Invoke-GPOScan
        }
        'Intune' {
            $scanResult = Invoke-IntunePolicyScan
        }
        'Combined' {
            Write-Host "[PolicyPilot]   Running Local RSoP scan..." -ForegroundColor Gray
            $localResult = Invoke-LocalRSoPScan
            Write-Host "[PolicyPilot]   Running Intune scan..." -ForegroundColor Gray
            $intuneResult = Invoke-IntunePolicyScan
            if ($localResult -and -not $localResult.Error -and $intuneResult -and -not $intuneResult.Error) {
                $mergedGPOs = [System.Collections.Generic.List[PSCustomObject]]::new()
                $mergedSettings = [System.Collections.Generic.List[PSCustomObject]]::new()
                foreach ($g in $localResult.GPOs) { [void]$mergedGPOs.Add($g) }
                foreach ($g in $intuneResult.GPOs) { [void]$mergedGPOs.Add($g) }
                foreach ($s in $localResult.Settings) { [void]$mergedSettings.Add($s) }
                foreach ($s in $intuneResult.Settings) { [void]$mergedSettings.Add($s) }
                $domain = if ($localResult.Domain -ne 'LocalMachine') { "$($localResult.Domain) (Co-managed)" } else { 'Co-managed (Local + Intune)' }
                $scanResult = @{ Timestamp=[datetime]::Now; Domain=$domain; GPOs=$mergedGPOs; Settings=$mergedSettings }
            } elseif ($localResult -and -not $localResult.Error) {
                $scanResult = $localResult
            } elseif ($intuneResult -and -not $intuneResult.Error) {
                $scanResult = $intuneResult
            } else {
                $scanResult = @{ Error = "Both scans failed" }
            }
        }
    }

    if (-not $scanResult -or $scanResult.Error) {
        $errMsg = if ($scanResult.Error) { $scanResult.Error } else { 'Scan returned no data' }
        Write-Host "[PolicyPilot] Scan FAILED: $errMsg" -ForegroundColor Red
        exit 1
    }

    $gpoCount = $scanResult.GPOs.Count
    $settCount = $scanResult.Settings.Count
    Write-Host "[PolicyPilot] Scan complete: $gpoCount GPOs/areas, $settCount settings" -ForegroundColor Green

    # Populate script-level state for Build-HtmlReport
    $Script:ScanData = $scanResult
    $Script:AllIntuneApps.Clear()

    # Run conflict detection
    $conflicts = Find-Conflicts $scanResult.Settings
    $conflictCount  = @($conflicts | Where-Object Severity -eq 'Conflict').Count
    $redundantCount = @($conflicts | Where-Object Severity -eq 'Redundant').Count
    if ($conflictCount -gt 0 -or $redundantCount -gt 0) {
        Write-Host "[PolicyPilot] Conflicts: $conflictCount, Redundancies: $redundantCount" -ForegroundColor Yellow
    }

    # Generate and write report
    Write-Host "[PolicyPilot] Generating HTML report..." -ForegroundColor Cyan
    $html = Build-HtmlReport
    [System.IO.File]::WriteAllText($OutputPath, $html, [System.Text.Encoding]::UTF8)
    $fileSize = [math]::Round((Get-Item $OutputPath).Length / 1KB, 1)
    Write-Host "[PolicyPilot] Report saved: $OutputPath ($($fileSize) KB)" -ForegroundColor Green
    Write-Host "[PolicyPilot] Done." -ForegroundColor Cyan
    exit 0
}

# Show window
$Window.ShowDialog()