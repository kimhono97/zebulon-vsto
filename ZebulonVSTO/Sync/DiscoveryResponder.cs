using System;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Text;
using System.Threading;

namespace ZebulonVSTO.Sync {
    /// <summary>
    /// Always-on UDP responder for peer auto-discovery. Listens on
    /// <see cref="SyncDefaults.DiscoveryPort"/> from add-in startup to shutdown —
    /// independent of the start/stop-locked sync socket — so a peer can be found
    /// even before it has started syncing.
    ///
    /// On a valid DISCOVER it replies with an ANNOUNCE (its sync IP/port + role)
    /// unicast to the datagram's actual source endpoint. It is pure networking
    /// (no PowerPoint Interop), so — unlike <see cref="SyncManager"/>'s receive
    /// loop — it needs no UI-thread marshalling.
    /// </summary>
    public sealed class DiscoveryResponder {
        private readonly SyncManager _sync;
        private readonly ISyncLogger _logger;
        private readonly string _version;

        private Thread _thread;
        private UdpClient _client;
        private volatile bool _running;

        public DiscoveryResponder(SyncManager sync, ISyncLogger logger) {
            _sync = sync;
            _logger = logger;
            _version = ReadVersion();
        }

        public bool IsRunning() {
            return _running;
        }

        public bool Start() {
            if (_running) {
                return false;
            }
            try {
                // Parameterless ctor so ReuseAddress can be set BEFORE Bind:
                // lets multiple add-in instances on one host share the discovery
                // port (single-machine testing). UdpClient(port) would bind first.
                UdpClient client = new UdpClient { ExclusiveAddressUse = false };
                client.Client.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReuseAddress, true);
                client.Client.Bind(new IPEndPoint(IPAddress.Any, SyncDefaults.DiscoveryPort));
                _client = client;
            } catch (Exception e) {
                // Discovery is best-effort: if the port can't be opened the add-in
                // still loads and manual setup keeps working.
                _logger?.LogError("Failed to open discovery port " + SyncDefaults.DiscoveryPort + ".", e);
                _client = null;
                return false;
            }

            _running = true;
            _thread = new Thread(ReceiveLoop) { IsBackground = true };
            _thread.Start();
            return true;
        }

        public void Stop() {
            if (!_running) {
                return;
            }
            // Cooperative shutdown, same shape as SyncManager.StopSync: flip the
            // flag and close the socket so the blocking Receive throws
            // ObjectDisposedException and the background thread unwinds. No Join
            // needed (IsBackground), and this loop never touches the UI thread.
            _running = false;
            UdpClient client = _client;
            _client = null;
            client?.Close();
            _thread = null;
        }

        private void ReceiveLoop() {
            UdpClient client = _client;
            if (client == null) {
                return;
            }
            IPEndPoint source = new IPEndPoint(IPAddress.Any, SyncDefaults.DiscoveryPort);
            try {
                while (_running) {
                    byte[] bytes;
                    try {
                        bytes = client.Receive(ref source);
                    } catch (ObjectDisposedException) {
                        break; // socket closed by Stop()
                    } catch (SocketException e) {
                        if (!_running) {
                            break;
                        }
                        _logger?.LogError("Discovery socket error while receiving.", e);
                        continue;
                    }
                    HandleDatagram(client, bytes, source);
                }
            } catch (Exception e) {
                _logger?.LogError("Discovery responder failed.", e);
            }
        }

        private void HandleDatagram(UdpClient client, byte[] bytes, IPEndPoint source) {
            SyncMessage message = SyncMessage.Parse(Encoding.UTF8.GetString(bytes));
            if (message == null || message.Type != (int)MessageType.DISCOVER) {
                return; // not a discovery ping — ignore (ANNOUNCEs, stray traffic)
            }
            DiscoveryPayload payload = DiscoveryPayload.Parse(message.Data);
            if (!payload.Valid || !payload.IsQuery) {
                return;
            }
            // Don't answer our own broadcast. Compare the per-process InstanceId,
            // NOT LocalIP — another process on the same host shares our IP but has
            // a different InstanceId (so same-machine peers/stand-ins still match).
            if (payload.InstanceId == _sync.InstanceId) {
                return;
            }
            DiscoveryRole role = CurrentRole();
            // Honor an optional role filter: a scan for RECEIVERs shouldn't be
            // answered by a SENDER/IDLE. (v1 scanner sends no filter → all reply.)
            if (payload.Want != DiscoveryRole.Unknown && payload.Want != role) {
                return;
            }
            SendAnnounce(client, source, message.ID, role);
        }

        private void SendAnnounce(UdpClient client, IPEndPoint target, int echoId, DiscoveryRole role) {
            try {
                string data = DiscoveryPayload.BuildAnnounce(role, Environment.MachineName, _version, _sync.InstanceId);
                // SenderIP/SenderPort carry the SYNC coordinates (where to reach
                // me for sync), not the discovery port. ID echoes the DISCOVER's
                // id so the scanner can drop stale replies from a prior scan.
                SyncMessage announce = new SyncMessage(echoId, _sync.LocalIP, _sync.LocalPort,
                    MessageType.ANNOUNCE, data);
                byte[] bytes = Encoding.UTF8.GetBytes(announce.ToJsonString());
                client.Send(bytes, bytes.Length, target);
            } catch (Exception e) {
                _logger?.LogError("Failed to send an ANNOUNCE.", e);
            }
        }

        private DiscoveryRole CurrentRole() {
            if (!_sync.IsRunning()) {
                return DiscoveryRole.Idle;
            }
            switch (_sync.Mode) {
                case SyncManager.SyncMode.RECEIVER: return DiscoveryRole.Receiver;
                case SyncManager.SyncMode.SENDER: return DiscoveryRole.Sender;
                default: return DiscoveryRole.Idle;
            }
        }

        private static string ReadVersion() {
            try {
                Version version = Assembly.GetExecutingAssembly().GetName().Version;
                return version != null ? version.ToString() : string.Empty;
            } catch (Exception) {
                return string.Empty;
            }
        }
    }
}
