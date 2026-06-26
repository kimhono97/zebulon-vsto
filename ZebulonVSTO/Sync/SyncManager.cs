using System;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;

namespace ZebulonVSTO.Sync {
    public class SyncManager {
        public enum SyncMode {
            NONE, SENDER, RECEIVER
        }

        private static SyncManager _instance;
        private static int _nextMessageId;

        // Per-process identity for discovery self-recognition. Distinguishes "my
        // own broadcast" from another process on the SAME host (LocalIP can't —
        // it's identical for every process on a machine). Stable for the session.
        private readonly string _instanceId = Guid.NewGuid().ToString("N");

        private Thread _receiveThread;
        private UdpClient _client;
        private volatile bool _running;

        private readonly string _localIP;
        private string _remoteIP;
        private int _localPort;
        private int _remotePort;
        private SyncMode _mode;

        // Display-only label for the current remote (e.g. a discovered peer's
        // host name). NOT part of the wire protocol — purely for the ribbon's
        // status line. Set by the setup wizard before StartSync.
        private string _remoteLabel = "";

        // Host-provided collaborators, wired via Attach() at add-in startup so
        // this class carries no compile-time dependency on the VSTO Globals.
        private ISyncLogger _logger;
        private ISlideController _controller;
        private IStatusObserver _statusObserver;
        private bool _allowCustomCommands;

        // Most recent sender seen while in RECEIVER mode (display-only, for the
        // ribbon's status line). Written on the receive thread, read on the UI
        // thread; a torn read is harmless since it only drives a label.
        private string _lastPeerIP = "";
        private int _lastPeerPort;

        private SyncManager() {
            _receiveThread = null;
            _client = null;
            _running = false;

            _localIP = FindLocalIPAddress();
            _localPort = SyncDefaults.Port;
            _remoteIP = SyncDefaults.Broadcast;
            _remotePort = SyncDefaults.Port;
            _mode = SyncMode.NONE;
        }

        /// <summary>
        /// Connect the add-in host so the manager can log traffic and execute
        /// received commands. <paramref name="allowCustomCommands"/> mirrors the
        /// host's debug flag (gates console-issued custom commands).
        /// </summary>
        public void Attach(ISyncLogger logger, ISlideController controller,
                           IStatusObserver statusObserver, bool allowCustomCommands) {
            _logger = logger;
            _controller = controller;
            _statusObserver = statusObserver;
            _allowCustomCommands = allowCustomCommands;
        }

        public string LocalIP {
            get { return _localIP; }
        }
        public string InstanceId {
            get { return _instanceId; }
        }
        public int LocalPort {
            get { return _localPort; }
            set { if (!_running) { _localPort = value; } }
        }
        public string RemoteIP {
            get { return _remoteIP; }
            set { if (!_running) { _remoteIP = value; } }
        }
        public string RemoteLabel {
            get { return _remoteLabel; }
            set { if (!_running) { _remoteLabel = value ?? ""; } }
        }
        public int RemotePort {
            get { return _remotePort; }
            set { if (!_running) { _remotePort = value; } }
        }
        public SyncMode Mode {
            get { return _mode; }
            set { if (!_running) { _mode = value; } }
        }
        public string LastPeerIP {
            get { return _lastPeerIP; }
        }
        public int LastPeerPort {
            get { return _lastPeerPort; }
        }
        public bool IsRunning() {
            return _running;
        }

        public bool StartSync() {
            if (_running) {
                return false;
            }
            if (_mode != SyncMode.SENDER && _mode != SyncMode.RECEIVER) {
                _logger?.LogError("Cannot start sync: invalid mode.", new InvalidOperationException("Mode=" + _mode));
                return false;
            }

            try {
                _client = new UdpClient(_localPort);
            } catch (Exception e) {
                _logger?.LogError("Failed to open UDP port " + _localPort + ".", e);
                _client = null;
                return false;
            }

            _lastPeerIP = "";
            _lastPeerPort = 0;
            _running = true;
            _receiveThread = new Thread(ReceiveLoop) { IsBackground = true };
            _receiveThread.Start();
            _logger?.Log("---> Start Sync (Mode: " + _mode + ")");
            return true;
        }

        public void StopSync() {
            if (!_running) {
                return;
            }

            // Cooperative shutdown. StopSync runs on the UI thread, and the
            // receive thread marshals commands back to that same thread via
            // Dispatcher.Invoke — so we must NOT block the UI thread on a Join
            // here (the two would block each other until the OS timed out).
            // Instead we flip the flag and close the socket: the blocking
            // Receive throws ObjectDisposedException and the background
            // (IsBackground) receive thread unwinds on its own.
            _running = false;

            UdpClient client = _client;
            _client = null;
            client?.Close();
            _receiveThread = null;
            _logger?.Log("<--- Stop Sync");
        }

        private void ReceiveLoop() {
            // Capture the socket once so a concurrent StopSync (_client = null)
            // can't turn our reads into a NullReferenceException; Close() on the
            // captured instance still unblocks Receive as ObjectDisposedException.
            UdpClient client = _client;
            if (client == null) {
                return;
            }
            IPEndPoint endpoint = new IPEndPoint(IPAddress.Any, _localPort);
            try {
                while (_running) {
                    byte[] bytes;
                    try {
                        bytes = client.Receive(ref endpoint);
                    } catch (ObjectDisposedException) {
                        break; // socket closed by StopSync
                    } catch (SocketException e) {
                        if (!_running) {
                            break;
                        }
                        _logger?.LogError("Socket error while receiving.", e);
                        continue;
                    }

                    SyncMessage message = SyncMessage.Parse(Encoding.UTF8.GetString(bytes));
                    if (message == null) {
                        continue;
                    }
                    if (message.IsResponse()) {
                        LogMessage("▶RES", message);
                    } else if (_mode == SyncMode.RECEIVER) {
                        UpdateLastPeer(message.SenderIP, message.SenderPort);
                        LogMessage("▶REQ", message);
                        bool handled = ProcessRequest(message);
                        SendResponse(message, handled);
                    }
                }
            } catch (Exception e) {
                _logger?.LogError("Failed to receive sync messages.", e);
            }
        }

        private bool SendMessage(SyncMessage message) {
            UdpClient client = _client;
            if (client == null) {
                return false;
            }
            try {
                byte[] bytes = Encoding.UTF8.GetBytes(message.ToJsonString());
                client.Send(bytes, bytes.Length, _remoteIP, _remotePort);
                LogMessage("◀REQ", message);
            } catch (Exception e) {
                _logger?.LogError("Failed to send a sync message.", e);
                return false;
            }
            return true;
        }

        private bool SendResponse(SyncMessage received, bool result) {
            UdpClient client = _client;
            if (client == null) {
                return false;
            }
            try {
                SyncMessage message = new SyncMessage(received.ID, _localIP, _localPort,
                    MessageType.RESPONSE, result ? "Success" : "Failed");
                byte[] bytes = Encoding.UTF8.GetBytes(message.ToJsonString());
                client.Send(bytes, bytes.Length, received.SenderIP, received.SenderPort);
                LogMessage("◀RES", message);
            } catch (Exception e) {
                _logger?.LogError("Failed to send a sync response message.", e);
                return false;
            }
            return true;
        }

        private SyncMessage CreateSenderMessage(MessageType type, string data) {
            int id = Interlocked.Increment(ref _nextMessageId);
            return new SyncMessage(id, _localIP, _localPort, type, data);
        }

        public bool SendRequestMessage(string data, bool isCustom = false) {
            if (_mode != SyncMode.SENDER && (!_allowCustomCommands || !isCustom)) {
                return false;
            }
            SyncMessage message = CreateSenderMessage(isCustom ? MessageType.CUSTOM : MessageType.REQUEST, data);
            return SendMessage(message);
        }

        // Track the most recent sender and notify the host only when it changes,
        // so the ribbon refreshes once per new peer rather than per datagram.
        private void UpdateLastPeer(string ip, int port) {
            if (ip == _lastPeerIP && port == _lastPeerPort) {
                return;
            }
            _lastPeerIP = ip;
            _lastPeerPort = port;
            _statusObserver?.OnPeerChanged();
        }

        private bool ProcessRequest(SyncMessage received) {
            ParsedCommand command = CommandParser.Parse(received.Data);
            switch (command.Kind) {
                case CommandKind.Alert:
                    string sender = received.SenderIP + ":" + received.SenderPort;
                    return _controller != null && _controller.Alert(sender, command.Text);
                case CommandKind.Select:
                    return _controller != null && _controller.SelectSlide(command.SlideIndex);
                case CommandKind.ShowSlide:
                    return _controller != null && _controller.ShowSlide(command.SlideIndex);
                case CommandKind.HideSlide:
                    return _controller != null && _controller.HideSlide();
                default:
                    return false;
            }
        }

        public static SyncManager GetInstance() {
            if (_instance == null) {
                _instance = new SyncManager();
            }
            return _instance;
        }

        public static string FindLocalIPAddress() {
            try {
                IPHostEntry host = Dns.GetHostEntry(Dns.GetHostName());
                foreach (IPAddress ip in host.AddressList) {
                    if (ip.AddressFamily == AddressFamily.InterNetwork) {
                        return ip.ToString();
                    }
                }
            } catch (Exception) {
                // fall through to the loopback default
            }
            return SyncDefaults.Localhost;
        }

        private void LogMessage(string prefix, SyncMessage message) {
            if (_logger == null) {
                return;
            }
            string text = prefix + " ";
            if (message != null) {
                text += string.Format("[{0}:{1}][{2}][{3}] ", message.SenderIP, message.SenderPort,
                    message.ID, ((MessageType)message.Type).ToString());
                text += message.Data;
            }
            _logger.Log(text);
        }
    }
}
