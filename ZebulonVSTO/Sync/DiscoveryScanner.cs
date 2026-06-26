using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;

namespace ZebulonVSTO.Sync {
    /// <summary>
    /// One peer discovered by a <see cref="DiscoveryScanner"/> scan. The IP/port
    /// are the peer's SYNC coordinates (taken from the ANNOUNCE message fields),
    /// i.e. exactly what a SENDER copies into its RemoteIP/RemotePort.
    /// </summary>
    public sealed class DiscoveredPeer {
        public string Host { get; }
        public string IP { get; }
        public int SyncPort { get; }
        public DiscoveryRole Role { get; }
        public string Version { get; }

        public DiscoveredPeer(string host, string ip, int syncPort, DiscoveryRole role, string version) {
            Host = host ?? string.Empty;
            IP = ip ?? string.Empty;
            SyncPort = syncPort;
            Role = role;
            Version = version ?? string.Empty;
        }
    }

    /// <summary>
    /// Broadcasts DISCOVER pings and collects ANNOUNCE replies for the setup
    /// wizard's auto path. Runs the send/receive on a background thread and
    /// reports peers via callbacks as they arrive; the caller (wizard) is
    /// responsible for marshalling those callbacks onto the UI thread.
    ///
    /// Pure networking — no PowerPoint Interop and no WPF dependency.
    /// </summary>
    public sealed class DiscoveryScanner {
        // UDP is lossy on a LAN, so ping a few times across the window; replies
        // that arrive while we're still sending are buffered by the socket.
        private const int ScanWindowMs = 1800;
        private const int ReceiveTimeoutMs = 150;
        private static readonly long[] SendMarksMs = { 0, 300, 600 };

        private static int _nextScanId;

        private readonly SyncManager _sync;
        private readonly ISyncLogger _logger;

        private Thread _thread;
        private UdpClient _client;
        private volatile bool _running;

        public DiscoveryScanner(SyncManager sync, ISyncLogger logger) {
            _sync = sync;
            _logger = logger;
        }

        public bool IsRunning() {
            return _running;
        }

        /// <summary>
        /// Start a scan. <paramref name="onPeerFound"/> fires once per unique peer
        /// (on the background thread); <paramref name="onCompleted"/> fires when the
        /// collection window closes naturally (NOT when <see cref="Stop"/> is called).
        /// <paramref name="want"/> = <see cref="DiscoveryRole.Unknown"/> lets every
        /// role answer (the wizard filters/greys client-side).
        /// </summary>
        public bool Start(DiscoveryRole want, Action<DiscoveredPeer> onPeerFound, Action onCompleted) {
            if (_running) {
                return false;
            }
            try {
                _client = new UdpClient { EnableBroadcast = true };
                _client.Client.ReceiveTimeout = ReceiveTimeoutMs;
            } catch (Exception e) {
                _logger?.LogError("Failed to open a discovery scan socket.", e);
                _client = null;
                return false;
            }

            int scanId = Interlocked.Increment(ref _nextScanId);
            _running = true;
            _thread = new Thread(() => ScanLoop(scanId, want, onPeerFound, onCompleted)) { IsBackground = true };
            _thread.Start();
            return true;
        }

        public void Stop() {
            if (!_running) {
                return;
            }
            _running = false;
            UdpClient client = _client;
            _client = null;
            client?.Close();
            _thread = null;
        }

        private void ScanLoop(int scanId, DiscoveryRole want,
                              Action<DiscoveredPeer> onPeerFound, Action onCompleted) {
            UdpClient client = _client;
            if (client == null) {
                return;
            }
            IPEndPoint broadcast = new IPEndPoint(IPAddress.Broadcast, SyncDefaults.DiscoveryPort);
            IPEndPoint source = new IPEndPoint(IPAddress.Any, 0);
            HashSet<string> seen = new HashSet<string>();
            byte[] discover = BuildDiscover(scanId, want);

            bool naturalEnd = false;
            try {
                Stopwatch clock = Stopwatch.StartNew();
                int sent = 0;
                while (_running) {
                    if (clock.ElapsedMilliseconds >= ScanWindowMs) {
                        naturalEnd = true;
                        break;
                    }
                    // Emit the scheduled pings as their marks come due.
                    while (sent < SendMarksMs.Length && clock.ElapsedMilliseconds >= SendMarksMs[sent]) {
                        try {
                            client.Send(discover, discover.Length, broadcast);
                        } catch (Exception e) {
                            _logger?.LogError("Failed to send a DISCOVER ping.", e);
                        }
                        sent++;
                    }

                    byte[] bytes;
                    try {
                        bytes = client.Receive(ref source);
                    } catch (SocketException) {
                        continue; // ReceiveTimeout elapsed — recheck clock / send marks
                    } catch (ObjectDisposedException) {
                        break;    // Stop() closed the socket
                    }
                    HandleAnnounce(bytes, scanId, seen, onPeerFound);
                }
            } catch (Exception e) {
                _logger?.LogError("Discovery scan failed.", e);
            } finally {
                UdpClient toClose = _client;
                _client = null;
                _running = false;
                toClose?.Close();
            }

            if (naturalEnd) {
                onCompleted?.Invoke();
            }
        }

        private void HandleAnnounce(byte[] bytes, int scanId, HashSet<string> seen,
                                    Action<DiscoveredPeer> onPeerFound) {
            SyncMessage message = SyncMessage.Parse(Encoding.UTF8.GetString(bytes));
            if (message == null || message.Type != (int)MessageType.ANNOUNCE) {
                return;
            }
            if (message.ID != scanId) {
                return; // stale reply from a previous scan
            }
            if (message.SenderIP == _sync.LocalIP) {
                return; // don't list ourselves
            }
            DiscoveryPayload payload = DiscoveryPayload.Parse(message.Data);
            if (!payload.Valid) {
                return;
            }
            string key = message.SenderIP + ":" + message.SenderPort;
            if (!seen.Add(key)) {
                return; // already reported in this scan
            }
            onPeerFound?.Invoke(new DiscoveredPeer(
                payload.Host, message.SenderIP, message.SenderPort, payload.Role, payload.Version));
        }

        private byte[] BuildDiscover(int scanId, DiscoveryRole want) {
            // SenderIP lets responders self-skip our own broadcast; SenderPort is
            // informational only (responders reply to the actual UDP source).
            SyncMessage message = new SyncMessage(scanId, _sync.LocalIP, SyncDefaults.DiscoveryPort,
                MessageType.DISCOVER, DiscoveryPayload.BuildDiscover(want));
            return Encoding.UTF8.GetBytes(message.ToJsonString());
        }
    }
}
