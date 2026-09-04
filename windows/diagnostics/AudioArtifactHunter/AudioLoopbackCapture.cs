// AudioArtifactHunter - WASAPI loopback capture core.
//
// Records the WASAPI loopback stream exposed by a render endpoint into rolling
// WAV segments, while logging a per-second peak and RMS level history.
// When the peak exceeds a threshold the surrounding segments are preserved.
//
// Purpose: bracket an artefact relative to the endpoint's loopback tap. The tap
// location must be confirmed with the endpoint vendor before assigning component
// ownership, particularly for virtual endpoints such as Citrix HDX Audio.
//
// The per-second level log is written whether or not a trigger ever fires, so a
// measured level history exists for the whole monitoring period.
//
// Peak levels are computed from the raw float samples BEFORE clipping to the
// 16-bit output format, so an overload above 0 dBFS is still measured correctly
// rather than being reported as exactly 0.
//
// Interface IIDs below are the documented WASAPI/MMDevice values. If any is
// wrong the first Activate call fails with a clear HRESULT rather than
// misbehaving silently, so validate on one machine before fleet deployment.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace AudioArtifactHunter
{
    [ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
    internal class MMDeviceEnumeratorComObject
    {
    }

    [ComImport, Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMMDeviceEnumerator
    {
        [PreserveSig] int EnumAudioEndpoints(int dataFlow, int stateMask, out IntPtr devices);
        [PreserveSig] int GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice device);
        [PreserveSig] int GetDevice([MarshalAs(UnmanagedType.LPWStr)] string id, out IMMDevice device);
        [PreserveSig] int RegisterEndpointNotificationCallback(IntPtr client);
        [PreserveSig] int UnregisterEndpointNotificationCallback(IntPtr client);
    }

    [ComImport, Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMMDevice
    {
        [PreserveSig] int Activate(ref Guid iid, int clsCtx, IntPtr activationParams, [MarshalAs(UnmanagedType.IUnknown)] out object iface);
        [PreserveSig] int OpenPropertyStore(int access, out IntPtr propertyStore);
        [PreserveSig] int GetId([MarshalAs(UnmanagedType.LPWStr)] out string id);
        [PreserveSig] int GetState(out int state);
    }

    [ComImport, Guid("1CB9AD4C-DBFA-4C32-B178-C2F568A703B2"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IAudioClient
    {
        [PreserveSig] int Initialize(int shareMode, int streamFlags, long bufferDuration, long periodicity, IntPtr format, IntPtr sessionGuid);
        [PreserveSig] int GetBufferSize(out uint bufferFrames);
        [PreserveSig] int GetStreamLatency(out long latency);
        [PreserveSig] int GetCurrentPadding(out uint padding);
        [PreserveSig] int IsFormatSupported(int shareMode, IntPtr format, out IntPtr closestMatch);
        [PreserveSig] int GetMixFormat(out IntPtr format);
        [PreserveSig] int GetDevicePeriod(out long defaultPeriod, out long minimumPeriod);
        [PreserveSig] int Start();
        [PreserveSig] int Stop();
        [PreserveSig] int Reset();
        [PreserveSig] int SetEventHandle(IntPtr handle);
        [PreserveSig] int GetService(ref Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object service);
    }

    [ComImport, Guid("C8ADBD64-E71E-48A0-A4DE-185C395CD317"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IAudioCaptureClient
    {
        [PreserveSig] int GetBuffer(out IntPtr data, out uint frames, out uint flags, out ulong devicePosition, out ulong qpcPosition);
        [PreserveSig] int ReleaseBuffer(uint frames);
        [PreserveSig] int GetNextPacketSize(out uint frames);
    }

    [ComImport, Guid("5CDF2C82-841E-4546-9722-0CF74078229A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IAudioEndpointVolume
    {
        [PreserveSig] int RegisterControlChangeNotify(IntPtr notify);
        [PreserveSig] int UnregisterControlChangeNotify(IntPtr notify);
        [PreserveSig] int GetChannelCount(out uint channelCount);
        [PreserveSig] int SetMasterVolumeLevel(float levelDb, IntPtr eventContext);
        [PreserveSig] int SetMasterVolumeLevelScalar(float level, IntPtr eventContext);
        [PreserveSig] int GetMasterVolumeLevel(out float levelDb);
        [PreserveSig] int GetMasterVolumeLevelScalar(out float level);
        [PreserveSig] int SetChannelVolumeLevel(uint channel, float levelDb, IntPtr eventContext);
        [PreserveSig] int SetChannelVolumeLevelScalar(uint channel, float level, IntPtr eventContext);
        [PreserveSig] int GetChannelVolumeLevel(uint channel, out float levelDb);
        [PreserveSig] int GetChannelVolumeLevelScalar(uint channel, out float level);
        [PreserveSig] int SetMute([MarshalAs(UnmanagedType.Bool)] bool mute, IntPtr eventContext);
        [PreserveSig] int GetMute([MarshalAs(UnmanagedType.Bool)] out bool mute);
        [PreserveSig] int GetVolumeStepInfo(out uint step, out uint stepCount);
        [PreserveSig] int VolumeStepUp(IntPtr eventContext);
        [PreserveSig] int VolumeStepDown(IntPtr eventContext);
        [PreserveSig] int QueryHardwareSupport(out uint hardwareSupportMask);
        [PreserveSig] int GetVolumeRange(out float minimumDb, out float maximumDb, out float incrementDb);
    }

    // IAudioSessionControl (audiopolicy.h). Only GetState, GetDisplayName and
    // GetGroupingParam are ever called; the Set* and notification slots are
    // declared with IntPtr parameters purely to keep the vtable order correct
    // for QueryInterface to IAudioSessionControl2 and are never invoked.
    [ComImport, Guid("F4B1A599-7266-4319-A8CA-E70ACB11E8CD"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IAudioSessionControl
    {
        [PreserveSig] int GetState(out int state);
        [PreserveSig] int GetDisplayName([MarshalAs(UnmanagedType.LPWStr)] out string displayName);
        [PreserveSig] int SetDisplayName([MarshalAs(UnmanagedType.LPWStr)] string value, IntPtr eventContext);
        [PreserveSig] int GetIconPath([MarshalAs(UnmanagedType.LPWStr)] out string iconPath);
        [PreserveSig] int SetIconPath([MarshalAs(UnmanagedType.LPWStr)] string value, IntPtr eventContext);
        [PreserveSig] int GetGroupingParam(out Guid groupingParam);
        [PreserveSig] int SetGroupingParam(IntPtr overrideValue, IntPtr eventContext);
        [PreserveSig] int RegisterAudioSessionNotification(IntPtr newNotifications);
        [PreserveSig] int UnregisterAudioSessionNotification(IntPtr newNotifications);
    }

    // IAudioSessionControl2 (audiopolicy.h). A ComImport-derived interface must
    // repeat every base-interface slot first, in order, before its own new
    // methods; the base 9 slots here are byte-for-byte the same as
    // IAudioSessionControl above. Obtained from an IAudioSessionControl by a
    // C# cast, which performs QueryInterface via the RCW.
    [ComImport, Guid("bfb7ff88-7239-4fc9-8fa2-07c950be9c6d"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IAudioSessionControl2
    {
        [PreserveSig] int GetState(out int state);
        [PreserveSig] int GetDisplayName([MarshalAs(UnmanagedType.LPWStr)] out string displayName);
        [PreserveSig] int SetDisplayName([MarshalAs(UnmanagedType.LPWStr)] string value, IntPtr eventContext);
        [PreserveSig] int GetIconPath([MarshalAs(UnmanagedType.LPWStr)] out string iconPath);
        [PreserveSig] int SetIconPath([MarshalAs(UnmanagedType.LPWStr)] string value, IntPtr eventContext);
        [PreserveSig] int GetGroupingParam(out Guid groupingParam);
        [PreserveSig] int SetGroupingParam(IntPtr overrideValue, IntPtr eventContext);
        [PreserveSig] int RegisterAudioSessionNotification(IntPtr newNotifications);
        [PreserveSig] int UnregisterAudioSessionNotification(IntPtr newNotifications);
        [PreserveSig] int GetSessionIdentifier([MarshalAs(UnmanagedType.LPWStr)] out string sessionIdentifier);
        [PreserveSig] int GetSessionInstanceIdentifier([MarshalAs(UnmanagedType.LPWStr)] out string sessionInstanceIdentifier);
        [PreserveSig] int GetProcessId(out uint processId);
        // Returns S_OK (0) when this is the system sounds session, S_FALSE (1)
        // otherwise; both are success codes, so callers must not treat a
        // nonzero return as failure.
        [PreserveSig] int IsSystemSoundsSession();
        [PreserveSig] int SetDuckingPreference([MarshalAs(UnmanagedType.Bool)] bool optOut);
    }

    // IAudioSessionEnumerator (audiopolicy.h).
    [ComImport, Guid("E2F5BB11-0570-40CA-ACDD-3AA01277DEE8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IAudioSessionEnumerator
    {
        [PreserveSig] int GetCount(out int sessionCount);
        [PreserveSig] int GetSession(int index, out IAudioSessionControl session);
    }

    // IAudioSessionManager2 (audiopolicy.h). It derives from IAudioSessionManager,
    // so every ComImport-derived interface in C# must repeat the base slots
    // first, in order: GetAudioSessionControl and GetSimpleAudioVolume are the
    // two IAudioSessionManager slots, declared here but never called - only
    // GetSessionEnumerator is used.
    [ComImport, Guid("77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IAudioSessionManager2
    {
        [PreserveSig] int GetAudioSessionControl(IntPtr audioSessionGuid, int streamFlags, out IAudioSessionControl sessionControl);
        [PreserveSig] int GetSimpleAudioVolume(IntPtr audioSessionGuid, int streamFlags, out ISimpleAudioVolume audioVolume);
        [PreserveSig] int GetSessionEnumerator(out IAudioSessionEnumerator sessionEnumerator);
        [PreserveSig] int RegisterSessionNotification(IntPtr sessionNotification);
        [PreserveSig] int UnregisterSessionNotification(IntPtr sessionNotification);
        [PreserveSig] int RegisterDuckNotification([MarshalAs(UnmanagedType.LPWStr)] string sessionId, IntPtr duckNotification);
        [PreserveSig] int UnregisterDuckNotification(IntPtr duckNotification);
    }

    // ISimpleAudioVolume (audioclient.h). Obtained by QueryInterface on a
    // session control (a C# cast on the RCW), per the documented pattern for
    // reading per-session volume/mute.
    [ComImport, Guid("87CE5498-68D6-44E5-9215-6DA47EF883D8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface ISimpleAudioVolume
    {
        [PreserveSig] int SetMasterVolume(float level, IntPtr eventContext);
        [PreserveSig] int GetMasterVolume(out float level);
        [PreserveSig] int SetMute([MarshalAs(UnmanagedType.Bool)] bool mute, IntPtr eventContext);
        [PreserveSig] int GetMute([MarshalAs(UnmanagedType.Bool)] out bool mute);
    }

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    internal struct WaveFormatEx
    {
        public ushort FormatTag;
        public ushort Channels;
        public uint SamplesPerSec;
        public uint AvgBytesPerSec;
        public ushort BlockAlign;
        public ushort BitsPerSample;
        public ushort Size;
    }

    /// <summary>
    /// Writes 16-bit PCM WAV files and patches the RIFF sizes on close.
    /// </summary>
    internal sealed class WavWriter : IDisposable
    {
        private readonly FileStream _stream;
        private readonly BinaryWriter _writer;
        private readonly object _sync = new object();
        private int _dataBytes;

        public string Path { get; private set; }

        public WavWriter(string path, int sampleRate, int channels)
        {
            Path = path;
            _stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.Read);
            _writer = new BinaryWriter(_stream);

            _writer.Write(Encoding.ASCII.GetBytes("RIFF"));
            _writer.Write(0);
            _writer.Write(Encoding.ASCII.GetBytes("WAVE"));
            _writer.Write(Encoding.ASCII.GetBytes("fmt "));
            _writer.Write(16);
            _writer.Write((ushort)1);
            _writer.Write((ushort)channels);
            _writer.Write(sampleRate);
            _writer.Write(sampleRate * channels * 2);
            _writer.Write((ushort)(channels * 2));
            _writer.Write((ushort)16);
            _writer.Write(Encoding.ASCII.GetBytes("data"));
            _writer.Write(0);
        }

        public void WriteSamples(short[] samples, int count)
        {
            lock (_sync)
            {
                for (int i = 0; i < count; i++)
                {
                    _writer.Write(samples[i]);
                }

                _dataBytes += count * 2;
            }
        }

        public void Snapshot(string targetPath)
        {
            lock (_sync)
            {
                PatchHeader();
                File.Copy(Path, targetPath, true);
            }
        }

        public void Dispose()
        {
            lock (_sync)
            {
                try
                {
                    PatchHeader();
                }
                finally
                {
                    _writer.Close();
                    _stream.Dispose();
                }
            }
        }

        private void PatchHeader()
        {
            long endPosition = _stream.Position;
            _writer.Flush();
            _stream.Seek(4, SeekOrigin.Begin);
            _writer.Write(36 + _dataBytes);
            _stream.Seek(40, SeekOrigin.Begin);
            _writer.Write(_dataBytes);
            _writer.Flush();
            _stream.Seek(endPosition, SeekOrigin.Begin);
        }
    }

    public sealed class EndpointStateSnapshot
    {
        public int Role { get; internal set; }
        public string EndpointId { get; internal set; }
        public int DeviceState { get; internal set; }
        public float MasterVolumeScalar { get; internal set; }
        public bool Muted { get; internal set; }
        public int HResult { get; internal set; }
        public string Error { get; internal set; }
    }

    /// <summary>
    /// One Volume Mixer row: the per-application audio session state exposed
    /// through IAudioSessionControl2 and ISimpleAudioVolume. Distinct from
    /// EndpointStateSnapshot, which is the single device-wide master volume.
    /// </summary>
    public sealed class AudioSessionSnapshot
    {
        public int Role { get; internal set; }
        public string EndpointId { get; internal set; }
        public string SessionIdentifier { get; internal set; }
        public string SessionInstanceIdentifier { get; internal set; }
        public uint ProcessId { get; internal set; }
        public string ProcessName { get; internal set; }
        public string DisplayName { get; internal set; }
        public bool IsSystemSounds { get; internal set; }
        public int State { get; internal set; }
        public string StateName { get; internal set; }
        public float Volume { get; internal set; }
        public bool Muted { get; internal set; }
        public Guid GroupingParam { get; internal set; }
        public int HResult { get; internal set; }
        public string Error { get; internal set; }
    }

    public static class EndpointStateReader
    {
        /// <summary>
        /// Version of this capture core. The PowerShell scripts compare it
        /// with the version they require, because .NET Framework cannot unload
        /// an assembly: a session that compiled an older copy of this file
        /// keeps that copy until the process exits.
        /// </summary>
        public static string CoreVersion { get { return "1.3.0"; } }

        private const int DataFlowRender = 0;
        private const int ClsCtxAll = 23;
        private static readonly Guid IidAudioEndpointVolume = new Guid("5CDF2C82-841E-4546-9722-0CF74078229A");
        private static readonly Guid IidAudioSessionManager2 = new Guid("77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F");

        public static EndpointStateSnapshot ReadDefaultRenderEndpoint(int role)
        {
            EndpointStateSnapshot snapshot = new EndpointStateSnapshot();
            snapshot.Role = role;
            IMMDeviceEnumerator enumerator = null;
            IMMDevice device = null;
            IAudioEndpointVolume endpointVolume = null;

            try
            {
                enumerator = (IMMDeviceEnumerator)new MMDeviceEnumeratorComObject();
                int hr = enumerator.GetDefaultAudioEndpoint(DataFlowRender, role, out device);
                if (hr != 0)
                {
                    snapshot.HResult = hr;
                    snapshot.Error = "GetDefaultAudioEndpoint";
                    return snapshot;
                }

                string endpointId;
                hr = device.GetId(out endpointId);
                if (hr != 0)
                {
                    snapshot.HResult = hr;
                    snapshot.Error = "GetId";
                    return snapshot;
                }

                snapshot.EndpointId = endpointId;
                int deviceState;
                hr = device.GetState(out deviceState);
                if (hr != 0)
                {
                    snapshot.HResult = hr;
                    snapshot.Error = "GetState";
                    return snapshot;
                }

                snapshot.DeviceState = deviceState;
                Guid endpointVolumeIid = IidAudioEndpointVolume;
                object endpointVolumeObject;
                hr = device.Activate(ref endpointVolumeIid, ClsCtxAll, IntPtr.Zero, out endpointVolumeObject);
                if (hr != 0)
                {
                    snapshot.HResult = hr;
                    snapshot.Error = "Activate(IAudioEndpointVolume)";
                    return snapshot;
                }

                endpointVolume = (IAudioEndpointVolume)endpointVolumeObject;
                float scalar;
                hr = endpointVolume.GetMasterVolumeLevelScalar(out scalar);
                if (hr != 0)
                {
                    snapshot.HResult = hr;
                    snapshot.Error = "GetMasterVolumeLevelScalar";
                    return snapshot;
                }

                bool muted;
                hr = endpointVolume.GetMute(out muted);
                if (hr != 0)
                {
                    snapshot.HResult = hr;
                    snapshot.Error = "GetMute";
                    return snapshot;
                }

                snapshot.MasterVolumeScalar = scalar;
                snapshot.Muted = muted;
                snapshot.HResult = 0;
                return snapshot;
            }
            catch (Exception ex)
            {
                snapshot.HResult = ex.HResult;
                snapshot.Error = ex.Message;
                return snapshot;
            }
            finally
            {
                Release(endpointVolume);
                Release(device);
                Release(enumerator);
            }
        }

        /// <summary>
        /// Reads every audio session (Volume Mixer row) on the default render
        /// endpoint for a role. Read-only: never calls a Set* method. Includes
        /// expired sessions (state 2); the caller decides whether to filter
        /// them out.
        /// </summary>
        public static AudioSessionSnapshot[] ReadDefaultRenderSessions(int role)
        {
            IMMDeviceEnumerator enumerator = null;
            IMMDevice device = null;
            IAudioSessionManager2 sessionManager = null;

            try
            {
                enumerator = (IMMDeviceEnumerator)new MMDeviceEnumeratorComObject();
                int hr = enumerator.GetDefaultAudioEndpoint(DataFlowRender, role, out device);
                if (hr != 0)
                {
                    return new AudioSessionSnapshot[] { FailureSnapshot(role, null, hr, "GetDefaultAudioEndpoint") };
                }

                string endpointId;
                hr = device.GetId(out endpointId);
                if (hr != 0)
                {
                    return new AudioSessionSnapshot[] { FailureSnapshot(role, null, hr, "GetId") };
                }

                Guid sessionManagerIid = IidAudioSessionManager2;
                object sessionManagerObject;
                hr = device.Activate(ref sessionManagerIid, ClsCtxAll, IntPtr.Zero, out sessionManagerObject);
                if (hr != 0)
                {
                    return new AudioSessionSnapshot[] { FailureSnapshot(role, endpointId, hr, "Activate(IAudioSessionManager2)") };
                }

                sessionManager = (IAudioSessionManager2)sessionManagerObject;
                return EnumerateSessions(sessionManager, role, endpointId);
            }
            catch (Exception ex)
            {
                return new AudioSessionSnapshot[] { FailureSnapshot(role, null, ex.HResult, ex.Message) };
            }
            finally
            {
                Release(sessionManager);
                Release(device);
                Release(enumerator);
            }
        }

        /// <summary>
        /// Enumerates every session on an already-activated IAudioSessionManager2.
        /// Shared by ReadDefaultRenderSessions (which activates and releases the
        /// session manager per call) and LoopbackRecorder (which keeps one
        /// session manager activated for the life of the capture).
        /// </summary>
        internal static AudioSessionSnapshot[] EnumerateSessions(IAudioSessionManager2 sessionManager, int role, string endpointId)
        {
            IAudioSessionEnumerator sessionEnumerator = null;

            try
            {
                int hr = sessionManager.GetSessionEnumerator(out sessionEnumerator);
                if (hr != 0)
                {
                    return new AudioSessionSnapshot[] { FailureSnapshot(role, endpointId, hr, "GetSessionEnumerator") };
                }

                int count;
                hr = sessionEnumerator.GetCount(out count);
                if (hr != 0)
                {
                    return new AudioSessionSnapshot[] { FailureSnapshot(role, endpointId, hr, "GetCount") };
                }

                AudioSessionSnapshot[] results = new AudioSessionSnapshot[count];
                for (int i = 0; i < count; i++)
                {
                    results[i] = ReadOneSession(sessionEnumerator, i, role, endpointId);
                }

                return results;
            }
            catch (Exception ex)
            {
                return new AudioSessionSnapshot[] { FailureSnapshot(role, endpointId, ex.HResult, ex.Message) };
            }
            finally
            {
                Release(sessionEnumerator);
            }
        }

        private static AudioSessionSnapshot ReadOneSession(IAudioSessionEnumerator sessionEnumerator, int index, int role, string endpointId)
        {
            AudioSessionSnapshot snapshot = new AudioSessionSnapshot();
            snapshot.Role = role;
            snapshot.EndpointId = endpointId ?? string.Empty;
            snapshot.SessionIdentifier = string.Empty;
            snapshot.SessionInstanceIdentifier = string.Empty;
            snapshot.ProcessName = string.Empty;
            snapshot.DisplayName = string.Empty;
            snapshot.StateName = "Unknown";

            IAudioSessionControl control = null;
            IAudioSessionControl2 control2 = null;
            ISimpleAudioVolume simpleVolume = null;

            try
            {
                int hr = sessionEnumerator.GetSession(index, out control);
                if (hr != 0)
                {
                    snapshot.HResult = hr;
                    snapshot.Error = "GetSession";
                    return snapshot;
                }

                int state;
                hr = control.GetState(out state);
                if (hr != 0)
                {
                    snapshot.HResult = hr;
                    snapshot.Error = "GetState";
                    return snapshot;
                }

                snapshot.State = state;
                snapshot.StateName = SessionStateName(state);

                string displayName;
                if (control.GetDisplayName(out displayName) == 0)
                {
                    snapshot.DisplayName = displayName ?? string.Empty;
                }

                Guid groupingParam;
                if (control.GetGroupingParam(out groupingParam) == 0)
                {
                    snapshot.GroupingParam = groupingParam;
                }

                // Cast to IAudioSessionControl2 (QueryInterface via the RCW) for
                // the identifiers and process ID that only that interface exposes.
                control2 = (IAudioSessionControl2)control;

                string sessionIdentifier;
                if (control2.GetSessionIdentifier(out sessionIdentifier) == 0)
                {
                    snapshot.SessionIdentifier = sessionIdentifier ?? string.Empty;
                }

                string sessionInstanceIdentifier;
                if (control2.GetSessionInstanceIdentifier(out sessionInstanceIdentifier) == 0)
                {
                    snapshot.SessionInstanceIdentifier = sessionInstanceIdentifier ?? string.Empty;
                }

                uint processId;
                if (control2.GetProcessId(out processId) == 0)
                {
                    snapshot.ProcessId = processId;
                    snapshot.ProcessName = ResolveProcessName(processId);
                }

                // S_OK (0) means yes, S_FALSE (1) means no; both are success codes.
                snapshot.IsSystemSounds = (control2.IsSystemSoundsSession() == 0);

                // Cast to ISimpleAudioVolume (QueryInterface via the RCW) for the
                // per-session volume/mute state - the Volume Mixer row itself.
                simpleVolume = (ISimpleAudioVolume)control2;

                float volume;
                hr = simpleVolume.GetMasterVolume(out volume);
                if (hr != 0)
                {
                    snapshot.HResult = hr;
                    snapshot.Error = "GetMasterVolume";
                    return snapshot;
                }

                bool muted;
                hr = simpleVolume.GetMute(out muted);
                if (hr != 0)
                {
                    snapshot.HResult = hr;
                    snapshot.Error = "GetMute";
                    return snapshot;
                }

                snapshot.Volume = volume;
                snapshot.Muted = muted;
                snapshot.HResult = 0;
                return snapshot;
            }
            catch (Exception ex)
            {
                snapshot.HResult = ex.HResult;
                snapshot.Error = ex.Message;
                return snapshot;
            }
            finally
            {
                Release(simpleVolume);
                Release(control2);
                Release(control);
            }
        }

        private static AudioSessionSnapshot FailureSnapshot(int role, string endpointId, int hr, string error)
        {
            AudioSessionSnapshot snapshot = new AudioSessionSnapshot();
            snapshot.Role = role;
            snapshot.EndpointId = endpointId ?? string.Empty;
            snapshot.SessionIdentifier = string.Empty;
            snapshot.SessionInstanceIdentifier = string.Empty;
            snapshot.ProcessName = string.Empty;
            snapshot.DisplayName = string.Empty;
            snapshot.StateName = "Unknown";
            snapshot.HResult = hr;
            snapshot.Error = error;
            return snapshot;
        }

        internal static string SessionStateName(int state)
        {
            switch (state)
            {
                case 0: return "Inactive";
                case 1: return "Active";
                case 2: return "Expired";
                default: return "Unknown";
            }
        }

        private static string ResolveProcessName(uint processId)
        {
            if (processId == 0)
            {
                return string.Empty;
            }

            try
            {
                using (Process process = Process.GetProcessById((int)processId))
                {
                    return process.ProcessName;
                }
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                try { Marshal.FinalReleaseComObject(value); } catch (Exception) { }
            }
        }
    }

    /// <summary>
    /// Continuous WASAPI loopback recorder with rolling retention and
    /// peak-triggered preservation.
    /// </summary>
    public sealed class LoopbackRecorder : IDisposable
    {
        private const int DataFlowRender = 0;
        private const int RoleConsole = 0;
        private const int ShareModeShared = 0;
        private const int StreamFlagsLoopback = 0x00020000;
        private const int ClsCtxAll = 23;
        private const uint BufferFlagsDataDiscontinuity = 0x1;
        private const uint BufferFlagsSilent = 0x2;
        private const uint BufferFlagsTimestampError = 0x4;
        private const ushort FormatExtensible = 0xFFFE;
        private const ushort FormatIeeeFloat = 3;
        private const ushort FormatPcm = 1;

        // A single loud event spans many packets. Crossings within this many
        // seconds of each other are collapsed into one reported event so the
        // trigger log holds one row per artefact rather than one row per packet.
        private const double TriggerCooldownSeconds = 3.0;

        // The onset gate compares a loud packet against the peak of the
        // preceding quiet window, tracked in 100ms buckets. The two most
        // recent buckets are excluded because a transient's rise can straddle
        // a bucket boundary and would otherwise count as its own background.
        private const int OnsetBucketMilliseconds = 100;
        private const int OnsetGuardBuckets = 2;

        private static readonly Guid IidAudioClient = new Guid("1CB9AD4C-DBFA-4C32-B178-C2F568A703B2");
        private static readonly Guid IidAudioCaptureClient = new Guid("C8ADBD64-E71E-48A0-A4DE-185C395CD317");
        private static readonly Guid IidAudioEndpointVolume = new Guid("5CDF2C82-841E-4546-9722-0CF74078229A");
        private static readonly Guid IidAudioSessionManager2 = new Guid("77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F");
        private static readonly Guid SubtypeIeeeFloat = new Guid("00000003-0000-0010-8000-00AA00389B71");
        private static readonly Guid SubtypePcm = new Guid("00000001-0000-0010-8000-00AA00389B71");

        // Session-volume sampling cadence, independent of the 250ms endpoint
        // volume cadence and the 1s level-log cadence.
        private const double SessionSampleIntervalMilliseconds = 1000.0;
        private const double SessionHeartbeatSeconds = 30.0;

        private readonly string _rollingDirectory;
        private readonly string _preservedDirectory;
        private readonly string _levelLogPath;
        private readonly string _triggerLogPath;
        private readonly string _captureEventLogPath;
        private readonly string _endpointVolumeLogPath;
        private readonly string _sessionVolumeLogPath;
        private readonly string _requestedDeviceId;
        private readonly int _segmentSeconds;
        private readonly int _retainSegments;
        private readonly double _triggerThreshold;
        private readonly double _onsetQuietDbfs;
        private readonly double[] _onsetBuckets;
        private long _onsetCurrentBucket = -1;
        private long _onsetReadyBucket = -1;

        private Thread _worker;
        private volatile bool _stopRequested;
        private volatile bool _running;

        private WavWriter _segment;
        private readonly object _segmentSync = new object();
        private string _previousSegmentPath;
        private bool _preserveCurrentSegment;
        private DateTime _segmentStartUtc;

        private int _sampleRate;
        private int _channels;

        private double _secondPeak;
        private double _secondSumSquares;
        private long _secondFrames;
        private long _secondSamples;
        private DateTime _secondStartUtc;

        private bool _eventActive;
        private DateTime _eventStartUtc;
        private DateTime _eventLastCrossUtc;
        private double _eventPeak;
        private double _eventPreOnsetPeakDbfs;
        private string _eventPreservedPath;
        private string _eventOnsetSegmentPath;
        private ulong _eventDevicePosition;
        private ulong _eventQpcPosition;
        private DateTime _lastVolumeSampleUtc;
        private DateTime _lastVolumeRowUtc;
        private float _lastVolumeScalar = -1.0f;
        private bool _lastMute;
        private bool _hasVolumeSample;
        private bool _firstPacketSeen;

        private DateTime _lastSessionSampleUtc;
        private readonly Dictionary<string, SessionVolumeState> _sessionVolumeStates = new Dictionary<string, SessionVolumeState>(StringComparer.Ordinal);

        /// <summary>
        /// Last known (state, volume, muted) tuple for one audio session,
        /// keyed by SessionInstanceIdentifier, so a row is written only on
        /// change or heartbeat and a SessionGone row can report the process
        /// and display name that disappeared.
        /// </summary>
        private sealed class SessionVolumeState
        {
            public string StateName;
            public float Volume;
            public bool Muted;
            public DateTime LastRowUtc;
            public string SessionIdentifier;
            public uint ProcessId;
            public string ProcessName;
            public string DisplayName;
            public bool IsSystemSounds;
        }

        public string DeviceId { get; private set; }
        public int SampleRate { get { return _sampleRate; } }
        public int Channels { get { return _channels; } }
        public long TriggerCount { get; private set; }
        public long GatedCrossingCount { get; private set; }
        public bool OnsetGateEnabled { get { return _onsetBuckets != null; } }
        public double LastPeakDbfs { get; private set; }
        public string LastError { get; private set; }
        public bool IsRunning { get { return _running; } }
        public float LastEndpointVolumeScalar { get; private set; }
        public bool LastEndpointMuted { get; private set; }

        /// <summary>
        /// Endpoint generation number, set by the caller before Start when
        /// supervising restarts across endpoint replacement. Recorded on every
        /// CSV row so rows can be attributed to the generation that wrote them.
        /// </summary>
        public int Generation { get; set; }

        /// <summary>
        /// Creates a recorder. Nothing is captured until Start is called.
        /// </summary>
        /// <param name="outputDirectory">Root directory for all output.</param>
        /// <param name="deviceId">Endpoint ID to capture, or null for the default render endpoint.</param>
        /// <param name="segmentSeconds">Length of each rolling WAV segment.</param>
        /// <param name="retainSegments">Number of rolling segments kept on disk.</param>
        /// <param name="triggerThresholdDbfs">Peak level at or above which segments are preserved.</param>
        public LoopbackRecorder(string outputDirectory, string deviceId, int segmentSeconds, int retainSegments, double triggerThresholdDbfs)
            : this(outputDirectory, deviceId, segmentSeconds, retainSegments, triggerThresholdDbfs, 0.0, 0.0)
        {
        }

        /// <summary>
        /// Creates a recorder whose automatic trigger only opens from a quiet
        /// background: a packet at or above the threshold starts an event only
        /// when the peak of the preceding onsetQuietSeconds stayed at or below
        /// onsetQuietDbfs. Call audio, which sits near full scale for its whole
        /// duration, therefore does not preserve segments, while a transient
        /// out of an idle sink does. onsetQuietDbfs of 0 or onsetQuietSeconds of
        /// 0 disables the gate and restores the absolute threshold.
        /// </summary>
        public LoopbackRecorder(string outputDirectory, string deviceId, int segmentSeconds, int retainSegments, double triggerThresholdDbfs, double onsetQuietDbfs, double onsetQuietSeconds)
        {
            if (string.IsNullOrEmpty(outputDirectory))
            {
                throw new ArgumentNullException("outputDirectory");
            }

            _rollingDirectory = Path.Combine(outputDirectory, "rolling");
            _preservedDirectory = Path.Combine(outputDirectory, "preserved");
            _levelLogPath = Path.Combine(outputDirectory, "levels.csv");
            _triggerLogPath = Path.Combine(outputDirectory, "triggers.csv");
            _captureEventLogPath = Path.Combine(outputDirectory, "capture-events.csv");
            _endpointVolumeLogPath = Path.Combine(outputDirectory, "endpoint-volume.csv");
            _sessionVolumeLogPath = Path.Combine(outputDirectory, "session-volume.csv");
            _requestedDeviceId = deviceId;
            _segmentSeconds = segmentSeconds;
            _retainSegments = retainSegments;
            _triggerThreshold = triggerThresholdDbfs;
            _onsetQuietDbfs = onsetQuietDbfs;
            if (onsetQuietDbfs < 0.0 && onsetQuietSeconds > 0.0)
            {
                int bucketCount = (int)Math.Ceiling(onsetQuietSeconds * 1000.0 / OnsetBucketMilliseconds) + OnsetGuardBuckets;
                _onsetBuckets = new double[bucketCount];
            }

            LastPeakDbfs = -144.0;
            Generation = 1;

            Directory.CreateDirectory(_rollingDirectory);
            Directory.CreateDirectory(_preservedDirectory);

            EnsureCsvHeader(_levelLogPath, "TimestampUtc,PeakDbfs,RmsDbfs,Frames,EndpointId,Generation");
            EnsureCsvHeader(_triggerLogPath, "StartUtc,PeakDbfs,DurationSeconds,DevicePosition,Qpc100ns,PreservedSegment,OnsetSegment,EndpointId,Generation,PreOnsetPeakDbfs,ClosedBy");
            EnsureCsvHeader(_captureEventLogPath, "TimestampUtc,EndpointId,Event,HResult,Flags,Frames,DevicePosition,Qpc100ns,Details,Generation");
            EnsureCsvHeader(_endpointVolumeLogPath, "TimestampUtc,EndpointId,MasterScalar,MasterPercent,Muted,Changed,Generation");
            EnsureCsvHeader(_sessionVolumeLogPath, "TimestampUtc,EndpointId,Generation,SessionIdentifier,ProcessId,ProcessName,DisplayName,IsSystemSounds,StateName,Volume,Muted,Changed");
        }

        /// <summary>
        /// Creates the CSV with its header, or moves aside an existing file
        /// whose header differs so rows from two schemas never share a file.
        /// </summary>
        private static void EnsureCsvHeader(string path, string header)
        {
            if (File.Exists(path))
            {
                string existing = null;
                try
                {
                    using (StreamReader reader = new StreamReader(path, Encoding.UTF8, true))
                    {
                        existing = reader.ReadLine();
                    }
                }
                catch (IOException)
                {
                }

                if (existing == header)
                {
                    return;
                }

                string archived = Path.Combine(
                    Path.GetDirectoryName(path),
                    Path.GetFileNameWithoutExtension(path) + "-schema-" + DateTime.UtcNow.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture) + Path.GetExtension(path));
                try
                {
                    File.Move(path, archived);
                }
                catch (IOException)
                {
                    return;
                }
            }

            File.AppendAllText(path, header + "\r\n", Encoding.UTF8);
        }

        /// <summary>
        /// Records the packet peak in the current 100ms bucket and returns the
        /// peak of the quiet window that preceded it, in dBFS. Returns positive
        /// infinity while the gate is disabled or the window is not yet full,
        /// which callers treat as "not quiet".
        /// </summary>
        private double UpdateOnsetWindow(double packetPeak)
        {
            if (_onsetBuckets == null)
            {
                return double.PositiveInfinity;
            }

            int length = _onsetBuckets.Length;
            long bucket = DateTime.UtcNow.Ticks / (OnsetBucketMilliseconds * TimeSpan.TicksPerMillisecond);

            if (_onsetCurrentBucket < 0)
            {
                Array.Clear(_onsetBuckets, 0, length);
                _onsetCurrentBucket = bucket;
                _onsetReadyBucket = bucket + length;
            }
            else if (bucket != _onsetCurrentBucket)
            {
                long skipped = bucket - _onsetCurrentBucket;
                if (skipped > length)
                {
                    skipped = length;
                }

                for (long i = 1; i <= skipped; i++)
                {
                    _onsetBuckets[(int)((bucket - skipped + i) % length)] = 0.0;
                }

                _onsetCurrentBucket = bucket;
            }

            double trailing = 0.0;
            for (long b = bucket - length + 1; b <= bucket - OnsetGuardBuckets; b++)
            {
                double value = _onsetBuckets[(int)(b % length)];
                if (value > trailing)
                {
                    trailing = value;
                }
            }

            int current = (int)(bucket % length);
            if (packetPeak > _onsetBuckets[current])
            {
                _onsetBuckets[current] = packetPeak;
            }

            if (bucket < _onsetReadyBucket)
            {
                return double.PositiveInfinity;
            }

            return ToDbfs(trailing);
        }

        /// <summary>
        /// Starts capture on a dedicated MTA thread and returns immediately.
        /// All COM work happens on that thread to avoid apartment marshalling.
        /// </summary>
        public void Start()
        {
            if (_running)
            {
                return;
            }

            _stopRequested = false;
            _worker = new Thread(CaptureLoop);
            _worker.IsBackground = true;
            _worker.SetApartmentState(ApartmentState.MTA);
            _worker.Start();

            // Surface an immediate initialisation failure to the caller rather
            // than leaving a silently dead recorder behind.
            for (int i = 0; i < 100 && !_running && LastError == null; i++)
            {
                Thread.Sleep(50);
            }

            if (LastError != null)
            {
                throw new InvalidOperationException(LastError);
            }
        }

        /// <summary>
        /// Signals the capture loop to stop and waits for it to finish.
        /// </summary>
        public void Stop()
        {
            _stopRequested = true;

            if (_worker != null && _worker.IsAlive)
            {
                _worker.Join(5000);
            }
        }

        public void Dispose()
        {
            Stop();
        }

        public int PreserveRollingWindow(string targetDirectory)
        {
            Directory.CreateDirectory(targetDirectory);
            int copied = 0;

            lock (_segmentSync)
            {
                string activePath = _segment == null ? null : _segment.Path;

                foreach (string source in Directory.GetFiles(_rollingDirectory, "segment-*.wav"))
                {
                    string target = Path.Combine(targetDirectory, Path.GetFileName(source));
                    if (string.Equals(source, activePath, StringComparison.OrdinalIgnoreCase))
                    {
                        _segment.Snapshot(target);
                    }
                    else
                    {
                        File.Copy(source, target, true);
                    }

                    copied++;
                }
            }

            return copied;
        }

        private void CaptureLoop()
        {
            IntPtr formatPointer = IntPtr.Zero;
            IMMDeviceEnumerator enumerator = null;
            IMMDevice device = null;
            IAudioClient client = null;
            IAudioCaptureClient capture = null;
            IAudioEndpointVolume endpointVolume = null;
            IAudioSessionManager2 sessionManager = null;

            try
            {
                enumerator = (IMMDeviceEnumerator)new MMDeviceEnumeratorComObject();

                int hr;

                if (string.IsNullOrEmpty(_requestedDeviceId))
                {
                    hr = enumerator.GetDefaultAudioEndpoint(DataFlowRender, RoleConsole, out device);
                    Check(hr, "GetDefaultAudioEndpoint");
                }
                else
                {
                    hr = enumerator.GetDevice(_requestedDeviceId, out device);
                    Check(hr, "GetDevice");
                }

                string id;
                Check(device.GetId(out id), "GetId");
                DeviceId = id;

                Guid endpointVolumeIid = IidAudioEndpointVolume;
                object endpointVolumeObject;
                hr = device.Activate(ref endpointVolumeIid, ClsCtxAll, IntPtr.Zero, out endpointVolumeObject);
                if (hr == 0)
                {
                    endpointVolume = (IAudioEndpointVolume)endpointVolumeObject;
                    SampleEndpointVolume(endpointVolume, true);
                }
                else
                {
                    LogCaptureEvent("EndpointVolumeUnavailable", hr, 0, 0, 0, 0, "Activate(IAudioEndpointVolume)");
                }

                Guid sessionManagerIid = IidAudioSessionManager2;
                object sessionManagerObject;
                hr = device.Activate(ref sessionManagerIid, ClsCtxAll, IntPtr.Zero, out sessionManagerObject);
                if (hr == 0)
                {
                    sessionManager = (IAudioSessionManager2)sessionManagerObject;
                    SampleSessionVolumes(sessionManager, true);
                }
                else
                {
                    LogCaptureEvent("SessionManagerUnavailable", hr, 0, 0, 0, 0, "Activate(IAudioSessionManager2)");
                }

                Guid audioClientIid = IidAudioClient;
                object clientObject;
                Check(device.Activate(ref audioClientIid, ClsCtxAll, IntPtr.Zero, out clientObject), "Activate(IAudioClient)");
                client = (IAudioClient)clientObject;

                Check(client.GetMixFormat(out formatPointer), "GetMixFormat");
                WaveFormatEx format = (WaveFormatEx)Marshal.PtrToStructure(formatPointer, typeof(WaveFormatEx));

                bool isFloat = ResolveSampleType(format, formatPointer);

                _sampleRate = (int)format.SamplesPerSec;
                _channels = format.Channels;

                // A one second client buffer is far larger than needed and makes
                // the loop tolerant of scheduling delays on a busy host.
                Check(client.Initialize(ShareModeShared, StreamFlagsLoopback, 10000000L, 0, formatPointer, IntPtr.Zero), "Initialize");

                Guid captureIid = IidAudioCaptureClient;
                object captureObject;
                Check(client.GetService(ref captureIid, out captureObject), "GetService(IAudioCaptureClient)");
                capture = (IAudioCaptureClient)captureObject;

                Check(client.Start(), "Start");

                _secondStartUtc = DateTime.UtcNow;
                _segmentStartUtc = _secondStartUtc;
                _running = true;

                RotateSegment();

                int blockAlign = format.BlockAlign;
                int bytesPerSample = format.BitsPerSample / 8;
                short[] outputBuffer = new short[_sampleRate * _channels];

                while (!_stopRequested)
                {
                    uint packetFrames;
                    hr = capture.GetNextPacketSize(out packetFrames);
                    CheckCapture(hr, "GetNextPacketSize", 0, 0, 0, 0);

                    if (packetFrames == 0)
                    {
                        Thread.Sleep(10);
                        SampleEndpointVolume(endpointVolume, false);
                        SampleSessionVolumes(sessionManager, false);
                        FlushSecondIfDue();
                        RotateSegmentIfDue();
                        continue;
                    }

                    while (packetFrames != 0 && !_stopRequested)
                    {
                        IntPtr data;
                        uint frames;
                        uint flags;
                        ulong devicePosition;
                        ulong qpcPosition;

                        hr = capture.GetBuffer(out data, out frames, out flags, out devicePosition, out qpcPosition);
                        CheckCapture(hr, "GetBuffer", flags, frames, devicePosition, qpcPosition);

                        if (frames > 0)
                        {
                            bool silent = (flags & BufferFlagsSilent) != 0;
                            int sampleCount = (int)frames * _channels;

                            if (outputBuffer.Length < sampleCount)
                            {
                                outputBuffer = new short[sampleCount];
                            }

                            double packetSumSquares;
                            double packetPeak = ConvertSamples(data, (int)frames, blockAlign, bytesPerSample, isFloat, silent, outputBuffer, out packetSumSquares);

                            lock (_segmentSync)
                            {
                                _segment.WriteSamples(outputBuffer, sampleCount);
                            }

                            AccumulateLevel(packetPeak, packetSumSquares, frames, sampleCount);

                            if ((flags & BufferFlagsDataDiscontinuity) != 0)
                            {
                                // WASAPI reliably marks the very first packet after
                                // Start as discontinuous; that is expected stream
                                // startup behaviour, not a glitch worth flagging.
                                if (!_firstPacketSeen)
                                {
                                    LogCaptureEvent("DataDiscontinuity", 0, flags, frames, devicePosition, qpcPosition, "Expected discontinuity on first packet after Start; not a glitch.");
                                }
                                else
                                {
                                    LogCaptureEvent("DataDiscontinuity", 0, flags, frames, devicePosition, qpcPosition, "WASAPI reported a gap or stream-state transition.");
                                }
                            }

                            _firstPacketSeen = true;

                            if ((flags & BufferFlagsTimestampError) != 0)
                            {
                                LogCaptureEvent("TimestampError", 0, flags, frames, devicePosition, qpcPosition, "WASAPI marked the packet timestamp as uncertain.");
                            }

                            double peakDbfs = ToDbfs(packetPeak);
                            LastPeakDbfs = peakDbfs;
                            double preOnsetPeakDbfs = UpdateOnsetWindow(packetPeak);

                            if (peakDbfs >= _triggerThreshold)
                            {
                                if (_eventActive || _onsetBuckets == null || preOnsetPeakDbfs <= _onsetQuietDbfs)
                                {
                                    NoteThresholdCrossing(packetPeak, devicePosition, qpcPosition, preOnsetPeakDbfs);
                                }
                                else
                                {
                                    GatedCrossingCount++;
                                }
                            }
                        }

                        hr = capture.ReleaseBuffer(frames);
                        CheckCapture(hr, "ReleaseBuffer", flags, frames, devicePosition, qpcPosition);

                        FlushSecondIfDue();
                        SampleEndpointVolume(endpointVolume, false);
                        SampleSessionVolumes(sessionManager, false);
                        RotateSegmentIfDue();

                        hr = capture.GetNextPacketSize(out packetFrames);
                        CheckCapture(hr, "GetNextPacketSize", 0, 0, 0, 0);
                    }
                }
            }
            catch (Exception ex)
            {
                LastError = ex.Message;
            }
            finally
            {
                _running = false;

                CloseThresholdEvent(true);

                if (formatPointer != IntPtr.Zero)
                {
                    Marshal.FreeCoTaskMem(formatPointer);
                }

                CloseSegment();

                if (client != null)
                {
                    try { client.Stop(); } catch (Exception) { }
                }

                ReleaseComObject(capture);
                ReleaseComObject(client);
                ReleaseComObject(endpointVolume);
                ReleaseComObject(sessionManager);
                ReleaseComObject(device);
                ReleaseComObject(enumerator);
            }
        }

        /// <summary>
        /// Determines whether the mix format carries float or integer samples.
        /// </summary>
        private static bool ResolveSampleType(WaveFormatEx format, IntPtr formatPointer)
        {
            if (format.FormatTag == FormatIeeeFloat)
            {
                return true;
            }

            if (format.FormatTag == FormatPcm)
            {
                if (format.BitsPerSample != 16)
                {
                    throw new NotSupportedException("Unsupported PCM bit depth: " + format.BitsPerSample);
                }

                return false;
            }

            if (format.FormatTag == FormatExtensible)
            {
                if (format.Size < 22)
                {
                    throw new NotSupportedException("Invalid WAVEFORMATEXTENSIBLE cbSize: " + format.Size);
                }

                // WAVEFORMATEXTENSIBLE: 18 byte WAVEFORMATEX, 2 byte samples
                // union, 4 byte channel mask, then the 16 byte subformat GUID.
                byte[] raw = new byte[16];
                Marshal.Copy(new IntPtr(formatPointer.ToInt64() + 24), raw, 0, 16);
                Guid subFormat = new Guid(raw);

                if (subFormat == SubtypeIeeeFloat)
                {
                    return true;
                }

                if (subFormat == SubtypePcm)
                {
                    if (format.BitsPerSample != 16)
                    {
                        throw new NotSupportedException("Unsupported extensible PCM bit depth: " + format.BitsPerSample);
                    }

                    return false;
                }

                throw new NotSupportedException("Unsupported subformat: " + subFormat);
            }

            throw new NotSupportedException("Unsupported format tag: " + format.FormatTag);
        }

        /// <summary>
        /// Converts one captured packet to 16-bit PCM and returns the raw peak
        /// magnitude before clipping, so overloads above full scale are visible.
        /// </summary>
        private double ConvertSamples(IntPtr data, int frames, int blockAlign, int bytesPerSample, bool isFloat, bool silent, short[] output, out double sumSquares)
        {
            int sampleCount = frames * _channels;
            sumSquares = 0.0;

            if (silent)
            {
                Array.Clear(output, 0, sampleCount);
                return 0.0;
            }

            double peak = 0.0;

            if (isFloat)
            {
                float[] scratch = new float[sampleCount];
                Marshal.Copy(data, scratch, 0, sampleCount);

                for (int i = 0; i < sampleCount; i++)
                {
                    double magnitude = Math.Abs((double)scratch[i]);
                    sumSquares += (double)scratch[i] * scratch[i];
                    if (magnitude > peak)
                    {
                        peak = magnitude;
                    }

                    double clamped = scratch[i];
                    if (clamped > 1.0)
                    {
                        clamped = 1.0;
                    }
                    else if (clamped < -1.0)
                    {
                        clamped = -1.0;
                    }

                    output[i] = (short)(clamped * 32767.0);
                }
            }
            else
            {
                short[] scratch = new short[sampleCount];
                Marshal.Copy(data, scratch, 0, sampleCount);

                for (int i = 0; i < sampleCount; i++)
                {
                    double magnitude = Math.Abs((double)scratch[i]) / 32768.0;
                    sumSquares += magnitude * magnitude;
                    if (magnitude > peak)
                    {
                        peak = magnitude;
                    }

                    output[i] = scratch[i];
                }
            }

            return peak;
        }

        private void AccumulateLevel(double packetPeak, double packetSumSquares, uint frames, int samples)
        {
            if (packetPeak > _secondPeak)
            {
                _secondPeak = packetPeak;
            }

            _secondSumSquares += packetSumSquares;
            _secondFrames += frames;
            _secondSamples += samples;
        }

        private void FlushSecondIfDue()
        {
            DateTime now = DateTime.UtcNow;
            if ((now - _secondStartUtc).TotalSeconds < 1.0)
            {
                return;
            }

            double rms = 0.0;
            if (_secondSamples > 0)
            {
                rms = Math.Sqrt(_secondSumSquares / _secondSamples);
            }

            string line = string.Format(
                CultureInfo.InvariantCulture,
                "{0:o},{1:F2},{2:F2},{3},\"{4}\",{5}\r\n",
                _secondStartUtc,
                ToDbfs(_secondPeak),
                ToDbfs(rms),
                _secondFrames,
                DeviceId == null ? string.Empty : DeviceId.Replace("\"", "'"),
                Generation);

            try
            {
                File.AppendAllText(_levelLogPath, line, Encoding.UTF8);
            }
            catch (IOException)
            {
                // Losing one level row must never stop the capture.
            }

            _secondStartUtc = now;
            _secondPeak = 0.0;
            _secondSumSquares = 0.0;
            _secondFrames = 0;
            _secondSamples = 0;

            CloseThresholdEvent(false);
        }

        private void RotateSegmentIfDue()
        {
            if ((DateTime.UtcNow - _segmentStartUtc).TotalSeconds < _segmentSeconds)
            {
                return;
            }

            RotateSegment();
        }

        private void RotateSegment()
        {
            lock (_segmentSync)
            {
                CloseSegment();

                _segmentStartUtc = DateTime.UtcNow;
                string name = string.Format(
                    CultureInfo.InvariantCulture,
                    "segment-{0:yyyyMMdd-HHmmss-fff}.wav",
                    _segmentStartUtc);

                _segment = new WavWriter(Path.Combine(_rollingDirectory, name), _sampleRate, _channels);
                TrimRollingSegments();
            }
        }

        private void CloseSegment()
        {
            lock (_segmentSync)
            {
                if (_segment == null)
                {
                    return;
                }

                string path = _segment.Path;
                _segment.Dispose();
                _segment = null;

                if (_preserveCurrentSegment)
                {
                    PreserveFile(path);
                    _preserveCurrentSegment = false;
                }

                _previousSegmentPath = path;
            }
        }

        private void TrimRollingSegments()
        {
            try
            {
                string[] files = Directory.GetFiles(_rollingDirectory, "segment-*.wav");
                if (files.Length <= _retainSegments)
                {
                    return;
                }

                Array.Sort(files, StringComparer.Ordinal);

                for (int i = 0; i < files.Length - _retainSegments; i++)
                {
                    try
                    {
                        File.Delete(files[i]);
                    }
                    catch (IOException)
                    {
                    }
                }
            }
            catch (IOException)
            {
            }
        }

        /// <summary>
        /// Records that a packet crossed the trigger threshold, opening a new
        /// event or extending the one in progress.
        /// </summary>
        private void NoteThresholdCrossing(double packetPeak, ulong devicePosition, ulong qpcPosition, double preOnsetPeakDbfs)
        {
            DateTime now = DateTime.UtcNow;
            _eventLastCrossUtc = now;

            if (_eventActive)
            {
                if (packetPeak > _eventPeak)
                {
                    _eventPeak = packetPeak;
                }

                // The event is still open, so the segment now active must be
                // preserved too even though it was not the one open when the
                // event started.
                _preserveCurrentSegment = true;
                return;
            }

            _eventActive = true;
            _eventStartUtc = now;
            _eventPeak = packetPeak;
            _eventPreOnsetPeakDbfs = preOnsetPeakDbfs;
            _eventDevicePosition = devicePosition;
            _eventQpcPosition = qpcPosition;

            lock (_segmentSync)
            {
                _eventOnsetSegmentPath = _segment == null ? string.Empty : _segment.Path;
            }

            // The transient may straddle the segment boundary, so the segment
            // that preceded it is preserved alongside the current one.
            _eventPreservedPath = PreserveFile(_previousSegmentPath);
            _preserveCurrentSegment = true;
        }

        /// <summary>
        /// Closes an open threshold event once no further crossings have
        /// occurred within the cooldown, or immediately when forced at
        /// shutdown, and writes a single summary row. The duration is always
        /// measured from the first to the last real crossing.
        /// </summary>
        private void CloseThresholdEvent(bool force)
        {
            if (!_eventActive)
            {
                return;
            }

            if (!force && (DateTime.UtcNow - _eventLastCrossUtc).TotalSeconds < TriggerCooldownSeconds)
            {
                return;
            }

            _eventActive = false;
            TriggerCount++;

            string line = string.Format(
                CultureInfo.InvariantCulture,
                "{0:o},{1:F2},{2:F2},{3},{4},{5},{6},\"{7}\",{8},{9},{10}\r\n",
                _eventStartUtc,
                ToDbfs(_eventPeak),
                (_eventLastCrossUtc - _eventStartUtc).TotalSeconds,
                _eventDevicePosition,
                _eventQpcPosition,
                _eventPreservedPath,
                _eventOnsetSegmentPath,
                DeviceId == null ? string.Empty : DeviceId.Replace("\"", "'"),
                Generation,
                double.IsInfinity(_eventPreOnsetPeakDbfs) ? string.Empty : _eventPreOnsetPeakDbfs.ToString("F2", CultureInfo.InvariantCulture),
                force ? "Shutdown" : "Cooldown");

            try
            {
                File.AppendAllText(_triggerLogPath, line, Encoding.UTF8);
            }
            catch (IOException)
            {
            }
        }

        private string PreserveFile(string path)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
            {
                return string.Empty;
            }

            try
            {
                string target = Path.Combine(_preservedDirectory, Path.GetFileName(path));
                if (!File.Exists(target))
                {
                    File.Copy(path, target, false);
                }

                return target;
            }
            catch (IOException)
            {
                LogCaptureEvent("PreserveFailed", 0, 0, 0, 0, 0, "Could not copy " + path);
                return string.Empty;
            }
        }

        // WASAPI reports some non-error outcomes as positive HRESULTs (for
        // example AUDCLNT_S_BUFFER_EMPTY, 0x08890001). Only a negative HRESULT
        // is a genuine failure; a positive one is logged and otherwise ignored.
        private void CheckCapture(int hr, string operation, uint flags, uint frames, ulong devicePosition, ulong qpcPosition)
        {
            if (hr == 0)
            {
                return;
            }

            if (hr > 0)
            {
                LogCaptureEvent("SuccessCode", hr, flags, frames, devicePosition, qpcPosition, operation);
                return;
            }

            LogCaptureEvent("CaptureError", hr, flags, frames, devicePosition, qpcPosition, operation);
            Check(hr, operation);
        }

        private void SampleEndpointVolume(IAudioEndpointVolume endpointVolume, bool force)
        {
            if (endpointVolume == null)
            {
                return;
            }

            DateTime now = DateTime.UtcNow;
            if (!force && (now - _lastVolumeSampleUtc).TotalMilliseconds < 250.0)
            {
                return;
            }

            _lastVolumeSampleUtc = now;
            float scalar;
            bool muted;
            int volumeHr = endpointVolume.GetMasterVolumeLevelScalar(out scalar);
            int muteHr = endpointVolume.GetMute(out muted);

            if (volumeHr != 0 || muteHr != 0)
            {
                int error = volumeHr != 0 ? volumeHr : muteHr;
                LogCaptureEvent("EndpointVolumeReadError", error, 0, 0, 0, 0, volumeHr != 0 ? "GetMasterVolumeLevelScalar" : "GetMute");
                return;
            }

            bool changed = !_hasVolumeSample || Math.Abs(scalar - _lastVolumeScalar) > 0.0001f || muted != _lastMute;
            _hasVolumeSample = true;
            _lastVolumeScalar = scalar;
            _lastMute = muted;
            LastEndpointVolumeScalar = scalar;
            LastEndpointMuted = muted;

            // A row is written on change, plus a heartbeat every 30 seconds so
            // the file still proves the endpoint was being watched even when
            // volume and mute state never moved.
            bool heartbeatDue = (now - _lastVolumeRowUtc).TotalSeconds >= 30.0;
            if (!changed && !heartbeatDue)
            {
                return;
            }

            _lastVolumeRowUtc = now;

            string line = string.Format(
                CultureInfo.InvariantCulture,
                "{0:o},\"{1}\",{2:F6},{3:F2},{4},{5},{6}\r\n",
                now,
                DeviceId == null ? string.Empty : DeviceId.Replace("\"", "'"),
                scalar,
                scalar * 100.0f,
                muted,
                changed,
                Generation);

            try
            {
                File.AppendAllText(_endpointVolumeLogPath, line, Encoding.UTF8);
            }
            catch (IOException)
            {
            }
        }

        /// <summary>
        /// Samples every audio session (Volume Mixer row) on the endpoint
        /// already held by this recorder, at most once per
        /// SessionSampleIntervalMilliseconds unless force is set. A row is
        /// written per session only when its (state, volume, muted) tuple
        /// changed since the last sample, or every SessionHeartbeatSeconds.
        /// Read-only: never calls a Set* method.
        /// </summary>
        private void SampleSessionVolumes(IAudioSessionManager2 sessionManager, bool force)
        {
            if (sessionManager == null)
            {
                return;
            }

            DateTime now = DateTime.UtcNow;
            if (!force && (now - _lastSessionSampleUtc).TotalMilliseconds < SessionSampleIntervalMilliseconds)
            {
                return;
            }

            _lastSessionSampleUtc = now;

            AudioSessionSnapshot[] sessions;
            try
            {
                sessions = EndpointStateReader.EnumerateSessions(sessionManager, RoleConsole, DeviceId);
            }
            catch (Exception ex)
            {
                LogCaptureEvent("SessionSampleError", ex.HResult, 0, 0, 0, 0, ex.Message);
                return;
            }

            Dictionary<string, bool> seenThisPass = new Dictionary<string, bool>(StringComparer.Ordinal);

            foreach (AudioSessionSnapshot session in sessions)
            {
                if (session.HResult != 0)
                {
                    continue;
                }

                string key = session.SessionInstanceIdentifier;
                if (string.IsNullOrEmpty(key))
                {
                    continue;
                }

                seenThisPass[key] = true;

                SessionVolumeState previous;
                bool hasPrevious = _sessionVolumeStates.TryGetValue(key, out previous);
                bool changed = !hasPrevious
                    || previous.StateName != session.StateName
                    || Math.Abs(previous.Volume - session.Volume) > 0.0001f
                    || previous.Muted != session.Muted;
                bool heartbeatDue = !hasPrevious || (now - previous.LastRowUtc).TotalSeconds >= SessionHeartbeatSeconds;

                if (!changed && !heartbeatDue)
                {
                    continue;
                }

                WriteSessionVolumeRow(now, session, changed);

                SessionVolumeState state = hasPrevious ? previous : new SessionVolumeState();
                state.StateName = session.StateName;
                state.Volume = session.Volume;
                state.Muted = session.Muted;
                state.LastRowUtc = now;
                state.SessionIdentifier = session.SessionIdentifier;
                state.ProcessId = session.ProcessId;
                state.ProcessName = session.ProcessName;
                state.DisplayName = session.DisplayName;
                state.IsSystemSounds = session.IsSystemSounds;
                _sessionVolumeStates[key] = state;
            }

            // A previously tracked instance not seen this pass has gone away
            // (process exited, stream torn down) rather than merely changed.
            List<string> goneKeys = null;
            foreach (KeyValuePair<string, SessionVolumeState> entry in _sessionVolumeStates)
            {
                if (seenThisPass.ContainsKey(entry.Key))
                {
                    continue;
                }

                if (goneKeys == null)
                {
                    goneKeys = new List<string>();
                }

                goneKeys.Add(entry.Key);
            }

            if (goneKeys != null)
            {
                foreach (string key in goneKeys)
                {
                    WriteSessionGoneRow(now, key);
                    _sessionVolumeStates.Remove(key);
                }
            }
        }

        private void WriteSessionVolumeRow(DateTime nowUtc, AudioSessionSnapshot session, bool changed)
        {
            string line = string.Format(
                CultureInfo.InvariantCulture,
                "{0:o},\"{1}\",{2},\"{3}\",{4},\"{5}\",\"{6}\",{7},{8},{9:F6},{10},{11}\r\n",
                nowUtc,
                DeviceId == null ? string.Empty : DeviceId.Replace("\"", "'"),
                Generation,
                (session.SessionIdentifier ?? string.Empty).Replace("\"", "'"),
                session.ProcessId,
                (session.ProcessName ?? string.Empty).Replace("\"", "'"),
                (session.DisplayName ?? string.Empty).Replace("\"", "'"),
                session.IsSystemSounds,
                session.StateName,
                session.Volume,
                session.Muted,
                changed);

            try
            {
                File.AppendAllText(_sessionVolumeLogPath, line, Encoding.UTF8);
            }
            catch (IOException)
            {
            }
        }

        private void WriteSessionGoneRow(DateTime nowUtc, string sessionInstanceIdentifier)
        {
            SessionVolumeState state;
            if (!_sessionVolumeStates.TryGetValue(sessionInstanceIdentifier, out state))
            {
                return;
            }

            string line = string.Format(
                CultureInfo.InvariantCulture,
                "{0:o},\"{1}\",{2},\"{3}\",{4},\"{5}\",\"{6}\",{7},Gone,{8:F6},{9},{10}\r\n",
                nowUtc,
                DeviceId == null ? string.Empty : DeviceId.Replace("\"", "'"),
                Generation,
                (state.SessionIdentifier ?? string.Empty).Replace("\"", "'"),
                state.ProcessId,
                (state.ProcessName ?? string.Empty).Replace("\"", "'"),
                (state.DisplayName ?? string.Empty).Replace("\"", "'"),
                state.IsSystemSounds,
                state.Volume,
                state.Muted,
                true);

            try
            {
                File.AppendAllText(_sessionVolumeLogPath, line, Encoding.UTF8);
            }
            catch (IOException)
            {
            }
        }

        private void LogCaptureEvent(string eventName, int hr, uint flags, uint frames, ulong devicePosition, ulong qpcPosition, string details)
        {
            string safeDetails = (details ?? string.Empty).Replace('"', '\'');
            string line = string.Format(
                CultureInfo.InvariantCulture,
                "{0:o},\"{1}\",{2},0x{3:X8},0x{4:X8},{5},{6},{7},\"{8}\",{9}\r\n",
                DateTime.UtcNow,
                DeviceId == null ? string.Empty : DeviceId.Replace("\"", "'"),
                eventName,
                hr,
                flags,
                frames,
                devicePosition,
                qpcPosition,
                safeDetails,
                Generation);

            try
            {
                File.AppendAllText(_captureEventLogPath, line, Encoding.UTF8);
            }
            catch (IOException)
            {
            }
        }

        private static void ReleaseComObject(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                try { Marshal.FinalReleaseComObject(value); } catch (Exception) { }
            }
        }

        private static double ToDbfs(double magnitude)
        {
            if (magnitude <= 0.0000001)
            {
                return -144.0;
            }

            return 20.0 * Math.Log10(magnitude);
        }

        // A positive HRESULT is a success code (e.g. AUDCLNT_S_BUFFER_EMPTY),
        // not a failure, so only a negative HRESULT throws here.
        private static void Check(int hr, string operation)
        {
            if (hr < 0)
            {
                throw new InvalidOperationException(string.Format(
                    CultureInfo.InvariantCulture,
                    "{0} failed with HRESULT 0x{1:X8}.",
                    operation,
                    hr));
            }
        }
    }
}
