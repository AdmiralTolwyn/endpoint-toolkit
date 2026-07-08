#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Creates and activates a custom power plan for Modern Standby devices.
.DESCRIPTION
    Duplicates the Balanced plan and applies:
      - Power button  : Shut down (DC + AC)
      - Lid close     : Sleep     (DC + AC)
      - Screen timeout: 5 min     (DC + AC)
      - Sleep after   : 15 min (DC) / 60 min (AC)
.NOTES
    Target: Win11 25H2 Modern Standby (S0 Low Power Idle). PS 5.1 compatible.
#>

$PlanName = 'Modern Standby - Corporate'

# ── Subgroup & setting GUIDs ───────────────────────────────────────────
# https://learn.microsoft.com/en-us/windows-hardware/customize/power-settings/power-button-and-lid-settings
$SUB_BUTTONS  = '4f971e89-eebd-4455-a8de-9e59040e7347'
#https://learn.microsoft.com/en-us/windows-hardware/customize/power-settings/power-button-and-lid-settings-power-button-action
$PBUTTONACTION = '7648efa3-dd9c-4e3e-b566-50f929386280'
# https://learn.microsoft.com/en-us/windows-hardware/customize/power-settings/power-button-and-lid-settings-lid-switch-close-action
$LIDACTION     = '5ca83367-6e45-459f-a27b-476b1d01c936'

# https://learn.microsoft.com/en-us/windows-hardware/customize/power-settings/display-settings 
$SUB_VIDEO    = '7516b95f-f776-4464-8c53-06167f40cc99'
# https://learn.microsoft.com/en-us/windows-hardware/customize/power-settings/display-settings-display-idle-timeout
$VIDEOIDLE    = '3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e'

# https://learn.microsoft.com/en-us/windows-hardware/customize/power-settings/sleep-settings
$SUB_SLEEP    = '238c9fa8-0aad-41ed-83f4-97be242c8f20'
# https://learn.microsoft.com/en-us/windows-hardware/customize/power-settings/sleep-settings-sleep-idle-timeout
$STANDBYIDLE  = '29f6c1db-86da-48c5-9fdb-f2b67b1f44da'

# ── Action values for power button / lid close ────────────────────────
# https://learn.microsoft.com/en-us/windows-hardware/customize/power-settings/power-button-and-lid-settings-power-button-action
# 0 = Do nothing, 1 = Sleep, 2 = Hibernate, 3 = Shut down, 4 = Turn off display
$DoNothing = 0
$Sleep     = 1
$Hibernate = 2
$ShutDown  = 3

# ── Check for existing plan with the same name ────────────────────────
$existing = powercfg /list | Where-Object { $_ -match [regex]::Escape($PlanName) }
if ($existing) {
    # Extract GUID from the line
    if ($existing -match '([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})') {
        $PlanGuid = $Matches[1]
        Write-Host "Plan '$PlanName' already exists ($PlanGuid). Updating settings."
    }
}
else {
    # Duplicate Balanced (381b4222-f694-41f0-9685-ff5bb260df2e)
    $output = powercfg /duplicatescheme 381b4222-f694-41f0-9685-ff5bb260df2e 2>&1
    if ($output -match '([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})') {
        $PlanGuid = $Matches[1]
    }
    else {
        Write-Error "Failed to duplicate Balanced plan: $output"
        exit 1
    }
    powercfg /changename $PlanGuid $PlanName "Corporate power plan for Modern Standby devices"
    Write-Host "Created plan '$PlanName' ($PlanGuid)."
}

# ── Unhide button/lid settings (hidden by default on Modern Standby) ──
powercfg -attributes $SUB_BUTTONS $PBUTTONACTION -ATTRIB_HIDE
powercfg -attributes $SUB_BUTTONS $LIDACTION -ATTRIB_HIDE

# ── Apply settings ─────────────────────────────────────────────────────

# Power button → Shut down (DC + AC)
powercfg /setdcvalueindex $PlanGuid $SUB_BUTTONS $PBUTTONACTION $ShutDown
powercfg /setacvalueindex $PlanGuid $SUB_BUTTONS $PBUTTONACTION $ShutDown

# Lid close → Sleep (DC + AC)
powercfg /setdcvalueindex $PlanGuid $SUB_BUTTONS $LIDACTION $Sleep
powercfg /setacvalueindex $PlanGuid $SUB_BUTTONS $LIDACTION $Sleep

# Screen off after 5 min (DC + AC) — value in seconds
powercfg /setdcvalueindex $PlanGuid $SUB_VIDEO $VIDEOIDLE 300
powercfg /setacvalueindex $PlanGuid $SUB_VIDEO $VIDEOIDLE 300

# Sleep after 15 min (DC) / 60 min (AC) — value in seconds
powercfg /setdcvalueindex $PlanGuid $SUB_SLEEP $STANDBYIDLE 900
powercfg /setacvalueindex $PlanGuid $SUB_SLEEP $STANDBYIDLE 3600

# ── Activate ───────────────────────────────────────────────────────────
powercfg /setactive $PlanGuid

# ── Verify ─────────────────────────────────────────────────────────────
Write-Host "`n=== Active plan ==="
powercfg /getactivescheme

Write-Host "`n=== Settings applied ==="
$settings = @(
    @{ Label = 'Power button (DC)'; Sub = $SUB_BUTTONS; Set = $PBUTTONACTION; Expect = $ShutDown }
    @{ Label = 'Power button (AC)'; Sub = $SUB_BUTTONS; Set = $PBUTTONACTION; Expect = $ShutDown }
    @{ Label = 'Lid close (DC)';    Sub = $SUB_BUTTONS; Set = $LIDACTION;     Expect = $Sleep }
    @{ Label = 'Lid close (AC)';    Sub = $SUB_BUTTONS; Set = $LIDACTION;     Expect = $Sleep }
    @{ Label = 'Screen off (DC)';   Sub = $SUB_VIDEO;   Set = $VIDEOIDLE;     Expect = 300 }
    @{ Label = 'Screen off (AC)';   Sub = $SUB_VIDEO;   Set = $VIDEOIDLE;     Expect = 300 }
    @{ Label = 'Sleep after (DC)';  Sub = $SUB_SLEEP;   Set = $STANDBYIDLE;   Expect = 900 }
    @{ Label = 'Sleep after (AC)';  Sub = $SUB_SLEEP;   Set = $STANDBYIDLE;   Expect = 3600 }
)

foreach ($s in $settings) {
    $query = powercfg /query $PlanGuid $s.Sub $s.Set
    $dcVal = ($query | Select-String 'Current DC Power Setting Index:') -replace '.*:\s*', ''
    $acVal = ($query | Select-String 'Current AC Power Setting Index:') -replace '.*:\s*', ''
    $label = $s.Label
    if ($label -match 'DC') {
        $hex = $dcVal; $dec = [Convert]::ToInt64($hex, 16)
    }
    else {
        $hex = $acVal; $dec = [Convert]::ToInt64($hex, 16)
    }
    $status = if ($dec -eq $s.Expect) { 'OK' } else { "MISMATCH (got $dec, expected $($s.Expect))" }
    Write-Host ("  {0,-20} {1,6}  [{2}]" -f $label, $dec, $status)
}
