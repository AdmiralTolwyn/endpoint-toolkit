# AudioArtifactHunter

**Author:** Anton Romanyuk

> **Disclaimer:** These scripts are provided "as-is" without warranty of any kind, express or implied. Use at your own risk. The author assumes no liability for any damage or data loss resulting from their use. Always test in a non-production environment before deployment.

A set of evidence-collection tools for investigating **intermittent audio artifacts** on Windows endpoints — an unexpected loud transient ("pop", "bang", "screech") in a headset, a volume level that changes on its own, a mute that does not stick, or a notification sound that arrives at the wrong level.

The tools are deliberately **evidence-first**. They do not attempt to "fix" the audio stack. They capture what the endpoint actually did, with timestamps precise enough to correlate a user-reported moment against the audio samples, the endpoint state, and the event log.

## Problem

Intermittent audio artifacts are one of the hardest classes of endpoint ticket, because by the time anybody looks, the evidence is gone:

- The user cannot reproduce it on demand, so a live screen-share never catches it.
- Nothing in the Windows UI records **what the level was** at the moment of the artifact, so "it went loud" cannot be distinguished from "the volume was reset to 100%".
- Windows event logs roll over, and the relevant records are spread across audio, device, remoting and power providers that nobody queries together.
- A recording made after the fact proves nothing about the original signal path.

Worst of all, the two candidate explanations require completely different fixes and look identical to the user:

| Hypothesis | What actually happened | Where the fix lives |
|---|---|---|
| **Signal-level** | The rendered audio stream itself contained a loud transient (a notification sound, a glitch on stream transition, a buffer discontinuity). | Windows / application audio path |
| **Gain-level** | The stream was normal, but the endpoint volume or an amplifier changed underneath it. | Endpoint volume, remoting volume sync, headset/dock firmware |

These tools exist to separate those two, because you cannot fix an artifact you have not attributed. A loopback capture shows the digital signal; the endpoint-state monitor shows the gain applied to it. Running both simultaneously is the whole point.

## Contents

| File | Ver | Purpose | Records audio? |
|---|---|---|---|
| [`AudioLoopbackCapture.cs`](AudioLoopbackCapture.cs) | 1.3.0 | WASAPI interop core: loopback recorder, endpoint-state reader, WAV writer. Compiled at runtime by the scripts below; not run directly. | — |
| [`Start-AudioEndpointStateMonitor.ps1`](Start-AudioEndpointStateMonitor.ps1) | 1.3.0 | Samples default render endpoint identity, device state, master volume and mute, plus per-application sessions. | **No** |
| [`Start-AudioLoopbackRecorder.ps1`](Start-AudioLoopbackRecorder.ps1) | 1.0.0 | Rolling WASAPI loopback capture with level logging, automatic peak triggers and operator-marked preservation. | **Yes** |
| [`Invoke-AudioStimulus.ps1`](Invoke-AudioStimulus.ps1) | 1.0.0 | Fires controlled, logged audio stimuli so an intermittent artifact can be provoked rather than waited for. | Plays audio |
| [`Get-AudioStackInventory.ps1`](Get-AudioStackInventory.ps1) | 1.0.0 | Full audio-stack configuration snapshot, and a diff between two snapshots. | **No** |
| [`Invoke-AudioEventCorrelation.ps1`](Invoke-AudioEventCorrelation.ps1) | 1.1.0 | Cross-provider event timeline around an incident; profiles retained date ranges in an archived log collection. | **No** |
| [`Set-NotificationSoundState.ps1`](Set-NotificationSoundState.ps1) | 1.0.0 | Reversibly silences Windows notification/ringtone/alarm sounds while leaving the visual toast intact. | **No** (makes changes) |

Only `Set-NotificationSoundState.ps1` modifies the system. Everything else is read-only apart from writing into its own output folder.

## Before you run anything: consent and safety

Two of these tools need an explicit acknowledgement, and both gates are deliberate.

**`Start-AudioLoopbackRecorder.ps1` records audio.** Loopback capture records everything the machine renders — which in practice includes meeting audio and the voice of anyone on the far end of a call. `-AcknowledgeAudioCapture` is a **mandatory** parameter; the script will not run without it. Before deploying it:

- Confirm you have authorization to record on that endpoint, from whoever owns that decision in your organisation (privacy/works council/legal, as applicable).
- Tell the user what is being captured and for how long.
- Treat the output folder as sensitive. The script ACLs the directory and writes SHA-256 hashes of preserved evidence, but retention and disposal are your responsibility.

**`Invoke-AudioStimulus.ps1` deliberately plays sounds into a headset**, during an investigation whose whole subject is unexpectedly loud audio. `-AcknowledgeHearingSafety` is required. Use `-ListPlan` first — it prints the planned sequence and writes nothing. Do not run an unattended stimulus sequence against a headset somebody is wearing without briefing them first and having them set a safe volume.

## The scripts

### Start-AudioEndpointStateMonitor.ps1

The low-risk starting point, and usually the highest-value one. It captures **no audio at all**, so it is normally the easiest tool to get approved and the one to leave running longest.

It polls the default render endpoint for the Console, Multimedia and Communications roles and records identity, device state, master volume scalar and mute, plus per-application audio session state. Because it logs the **volume scalar over time**, it answers the gain-level question directly: if the artifact coincides with the scalar jumping, the signal was never the problem.

| Parameter | Default | Description |
|---|---|---|
| `OutputPath` | — | **Required.** CSV path; a `-sessions.csv` sibling is written alongside. |
| `SampleIntervalMilliseconds` | | Polling interval. |
| `HeartbeatSeconds` | | Console heartbeat cadence. |
| `DurationHours` | | Stop automatically after this long. |
| `ExcludeSessions` | Off | Skip per-application session sampling. |
| `StopFileName` | `STOP-MONITOR.txt` | Create this file to stop cleanly. |
| `Quiet` | Off | Suppress console output. |

```powershell
.\Start-AudioEndpointStateMonitor.ps1 -OutputPath C:\Temp\Audio\endpoint-state.csv -DurationHours 8
```

### Start-AudioLoopbackRecorder.ps1

Continuous WASAPI loopback capture with **rolling retention** — it keeps only the last N segments, so it can run for hours without filling the disk, and preserves a segment only when something interesting happens.

Preservation is triggered two ways: automatically when the peak crosses `-TriggerThresholdDbfs` (with a debounce so one event produces one row, not hundreds), and manually when the user drops a `MARK-INCIDENT.txt` file into the output folder the moment they hear something. The manual marker matters — it captures the user's *perception*, which is the only ground truth for "that was the noise I'm complaining about".

The automatic trigger is **onset-gated**, and this matters more than it sounds. A bare peak threshold does not work in practice: in a pilot on a laptop in a Teams call, a plain −6 dBFS threshold fired six times in two minutes — normal speech peaks — which under rolling retention would evict the real event before anyone looked at it. So the trigger only opens if the preceding `-OnsetQuietSeconds` were at or below `-OnsetQuietDbfs`, i.e. it looks for a loud sound emerging *out of quiet*, which is what "it suddenly banged" actually means. Use `-AbsoluteTrigger` to disable the gate when you genuinely want every peak.

It measures true per-sample RMS alongside peak, and logs buffer discontinuities and silent/timestamp-error flags rather than papering over them, so a gap in the capture is visible as a gap instead of being silently padded.

| Parameter | Default | Description |
|---|---|---|
| `AcknowledgeAudioCapture` | — | **Mandatory.** Consent gate; see above. |
| `OutputDirectory` | | Where segments and logs are written. |
| `ListDevices` | Off | Enumerate render endpoints and exit. No capture, no consent gate. |
| `DeviceId` | Default endpoint | Capture a specific endpoint. |
| `SegmentSeconds` / `RetainSegments` | | Rolling window size and depth. |
| `TriggerThresholdDbfs` | `-3` | Peak level at or above which a segment is auto-preserved. Pick this from measured `levels.csv` data rather than guessing. |
| `OnsetQuietDbfs` | `-40` | The trigger opens only if the preceding quiet window peaked at or below this. |
| `OnsetQuietSeconds` | `2` | Length of the quiet window the onset gate inspects. |
| `AbsoluteTrigger` | Off | Disable the onset gate; any peak at or above the threshold preserves, whatever preceded it. |
| `MaxPreservedMegabytes` | | Hard cap on preserved evidence. |
| `RestartOnEndpointChange` / `MaxEndpointRestarts` | | Follow the default endpoint across device changes (common on remoted sessions and USB headsets). |
| `DurationHours`, `StopFileName`, `LogPath`, `KeepRollingOnStop` | | Run length and shutdown behaviour. |

```powershell
# Enumerate endpoints first - read-only, no consent gate
.\Start-AudioLoopbackRecorder.ps1 -ListDevices

# Then capture
.\Start-AudioLoopbackRecorder.ps1 -OutputDirectory C:\Temp\Audio -AcknowledgeAudioCapture -RestartOnEndpointChange
```

### Invoke-AudioStimulus.ps1

Waiting for an intermittent fault is expensive. This fires **controlled, timestamped stimuli** instead, so each one can be correlated against the recorder and the state monitor. Stimulus types: `ToastDefaultSound`, `ToastSilent`, `DirectWav`, `SystemSound`, `ManualCue`.

`ToastSilent` is the control case: it raises the same notification with no sound bound. If the artifact still occurs on `ToastSilent`, the notification WAV is not the source — which immediately invalidates the most common assumption.

Use `-ListPlan` to print the sequence and estimated duration without playing anything. Point `-RecorderOutputDirectory` at a running recorder and it drops incident markers automatically.

```powershell
.\Invoke-AudioStimulus.ps1 -ListPlan
.\Invoke-AudioStimulus.ps1 -OutputDirectory C:\Temp\Stim -RecorderOutputDirectory C:\Temp\Audio -AcknowledgeHearingSafety
```

### Get-AudioStackInventory.ps1

A single read-only snapshot of everything that can plausibly change how audio is rendered: OS and build, audio services, endpoints and their properties, audio processing objects (APOs), drivers, reliability history, remoting/USB configuration and sound-scheme bindings.

Its real value is the **diff**. Take a baseline on a healthy machine (or on the same machine before a change), then compare:

```powershell
.\Get-AudioStackInventory.ps1 -OutputPath .\good.json
.\Get-AudioStackInventory.ps1 -OutputPath .\bad.json
.\Get-AudioStackInventory.ps1 -CompareWith .\bad.json -Baseline .\good.json
```

Run elevated for the broadest registry and file coverage. `-MeasureSoundAssets` additionally measures the peak level of the bound WAV assets, which is how you check whether a notification sound is itself unusually hot.

### Invoke-AudioEventCorrelation.ps1

Builds a cross-provider timeline around a known incident time, and reports event bursts. It reads live logs, `.evtx` files, or an archived `.zip` collection.

`-Scan` profiles what date range each log **actually retains** before you invest any time in analysis — the answer to "we collected logs, why is the incident not in them?" is very often that the relevant provider rolled over hours earlier. Run `-Scan` first.

`-IncidentTime` requires an ISO 8601 value **with a UTC offset**, enforced by pattern. This is intentional: correlating a headset artifact across providers is exactly where a silent local-vs-UTC mix-up produces a confident, wrong answer.

```powershell
.\Invoke-AudioEventCorrelation.ps1 -Path .\logs.zip -Scan
.\Invoke-AudioEventCorrelation.ps1 -Path .\logs.zip -IncidentTime '2026-09-04T14:32:10+02:00' -WindowMinutes 10
```

### Set-NotificationSoundState.ps1

The one tool here that changes the system — a **reversible mitigation**, not a diagnostic.

It blanks the WAV binding of system sound events so the notification still appears visually but produces no audio. This is the native Windows mechanism (several stock events ship with an empty binding), not a hack. Silencing rather than disabling is the point: turning notifications off removes the visual alert the user needs, when the acoustic exposure comes from the audio alone.

Scopes: `Notifications`, `Ringtones`, `Alarms`, `Devices`, `All`. The `Devices` scope is worth knowing about — it silences only device arrival/removal sounds, which is the fix when a headset or dock repeatedly re-enumerates on a remoted session and announces itself audibly every time.

Every original value is written to an identity-bound backup before any change, and restore uses app/event-relative paths rather than arbitrary paths from the JSON.

```powershell
.\Set-NotificationSoundState.ps1 -WhatIf
.\Set-NotificationSoundState.ps1 -Scope Notifications -BackupPath C:\Temp\sounds.json
.\Set-NotificationSoundState.ps1 -Restore -BackupPath C:\Temp\sounds.json
```

These are per-user values. On a non-persistent session host, apply via Group Policy Preferences, a logon script, or `-DefaultUser` against the golden image.

## Suggested workflow

1. **Inventory** — `Get-AudioStackInventory.ps1` on the affected machine and on a known-good one. Diff them.
2. **Scan logs** — `Invoke-AudioEventCorrelation.ps1 -Scan` to confirm the incident window is even retained.
3. **Monitor** — leave `Start-AudioEndpointStateMonitor.ps1` running. No audio is recorded, so this is usually approvable immediately.
4. **Capture** — once authorized, add `Start-AudioLoopbackRecorder.ps1` and brief the user to create `MARK-INCIDENT.txt` when they hear it.
5. **Provoke** — if waiting is not working, run `Invoke-AudioStimulus.ps1` with the recorder and monitor already running.
6. **Correlate** — feed the incident timestamp back into `Invoke-AudioEventCorrelation.ps1`.
7. **Mitigate** — only once attributed. `Set-NotificationSoundState.ps1` if the source is a notification sound.

## Output artifacts

| File | Written by | Contents |
|---|---|---|
| `levels.csv` | Recorder | Per-segment peak and true RMS in dBFS. |
| `triggers.csv` | Recorder | One row per debounced trigger, with peak and duration. |
| `capture-events.csv` | Recorder | Buffer discontinuities, silent/timestamp-error flags, HRESULTs, QPC timestamps. |
| `endpoint-volume.csv` / `session-volume.csv` | Recorder | Master scalar and mute sampled during capture. |
| `session.json` | Recorder / Stimulus | Run parameters and environment. |
| `endpoint-generations.csv` | Recorder | Each endpoint change and restart. |
| `evidence-hashes.csv` | Recorder | SHA-256 of every preserved artifact. |
| `MARK-INCIDENT.txt` | **You / the user** | Drop-file that preserves the current segment. |
| `endpoint-state.csv` + `-sessions.csv` | Monitor | Endpoint identity, state, volume, mute; per-app sessions. |
| `stimulus-log.csv` | Stimulus | One row per stimulus with pre/post endpoint state. `RunId` distinguishes processes; `Cycle` is only meaningful within one `RunId`. |
| `stimulus-log-schema-<stamp>.csv` | Stimulus | An existing log whose column layout no longer matches is archived here rather than being appended to or overwritten. |
| `STOP-RECORDER.txt` / `STOP-MONITOR.txt` / `STOP-STIMULUS.txt` | **You** | Clean-shutdown drop-files. |

## What this can and cannot prove

Being explicit about the boundaries, because loopback capture is routinely over-read:

**It can show** that the rendered digital stream did or did not contain a loud transient at a given moment; that the endpoint volume scalar did or did not change; that the capture had a buffer discontinuity; and that a given event coincided in time.

**It cannot show** anything that happens *after* the loopback tap. Loopback captures the mix as Windows renders it, so it will **not** contain distortion introduced by the headset, dock, DAC, Bluetooth/DECT link, or a hardware amplifier — nor by a remoting client's own audio path on a different machine. An artifact the user clearly heard that is **absent** from the loopback capture is a genuine and useful result: it relocates the fault downstream of Windows. Do not read it as "nothing happened".

It also cannot establish causation from correlation alone. Coincidence in a timeline is a lead, not a root cause — which is what `ToastSilent` and the inventory diff are for.

## Requirements

- **PowerShell 5.1.** These target Windows PowerShell 5.1 and use no PowerShell 7+ syntax.
- `AudioLoopbackCapture.cs` must sit **alongside** the recorder, monitor and stimulus scripts — they compile it at runtime with `Add-Type`. No NuGet package, no third-party audio library; it uses the in-box .NET Framework compiler and the Windows Core Audio COM APIs directly.
- The scripts check the core's `CoreVersion` and require **1.3.0 or later**. .NET Framework cannot unload an assembly, so if an older build was already compiled into the current session you will get an explicit error rather than silently mismatched behaviour. The fix is a **new PowerShell process** (`powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\<script>.ps1 ...`), not re-running in the same window.
- Elevation is not required for capture or monitoring. `Get-AudioStackInventory.ps1` gives broader coverage elevated, and `Set-NotificationSoundState.ps1 -DefaultUser` requires it.
- A render endpoint must be present and active. Verify with `-ListDevices` first.

## References

- [WASAPI](https://learn.microsoft.com/windows/win32/coreaudio/wasapi)
- [Loopback recording](https://learn.microsoft.com/windows/win32/coreaudio/loopback-recording)
- [IMMDeviceEnumerator](https://learn.microsoft.com/windows/win32/api/mmdeviceapi/nn-mmdeviceapi-immdeviceenumerator)
- [IAudioClient](https://learn.microsoft.com/windows/win32/api/audioclient/nn-audioclient-iaudioclient)
- [IAudioEndpointVolume](https://learn.microsoft.com/windows/win32/api/endpointvolume/nn-endpointvolume-iaudioendpointvolume)
- [Audio processing objects (APOs)](https://learn.microsoft.com/windows-hardware/drivers/audio/audio-processing-object-architecture)
