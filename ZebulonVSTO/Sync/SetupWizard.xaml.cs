using System;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;

namespace ZebulonVSTO.Sync {
    /// <summary>
    /// Interaction logic for SetupWizard.xaml — a modal 2-screen, asymmetric
    /// setup flow. Mode select → (receiver: pick a port) or (sender: auto-discover
    /// a peer / manual entry). On finish it writes the connection settings onto
    /// <see cref="SyncManager"/> and starts sync; on a start failure it stays open.
    /// </summary>
    public partial class SetupWizard {
        private enum Screen { ModeSelect, ReceiverSetup, SenderSetup }
        private enum SenderMethod { Auto, Manual }

        /// <summary>A discovered-peer row for the auto results list. Public so
        /// WPF can bind <see cref="IsSelectable"/> in the item container style.</summary>
        public sealed class PeerRow {
            public DiscoveredPeer Peer { get; }
            public string Display { get; }
            public bool IsSelectable { get; }

            public PeerRow(DiscoveredPeer peer) {
                Peer = peer;
                IsSelectable = peer.Role == DiscoveryRole.Receiver;
                string badge = peer.Role == DiscoveryRole.Receiver ? "🟢 수신"
                             : peer.Role == DiscoveryRole.Sender ? "🟠 송신" : "⚪ 대기";
                Display = peer.Host + "    " + peer.IP + " : " + peer.SyncPort + "    " + badge;
            }

            public override string ToString() {
                return Display;
            }
        }

        private readonly SyncManager _sync;
        private readonly ObservableCollection<PeerRow> _peers = new ObservableCollection<PeerRow>();
        private DiscoveryScanner _scanner;

        private Screen _screen;
        private SenderMethod _method = SenderMethod.Auto;
        private readonly bool _loaded;

        public SetupWizard() {
            InitializeComponent();
            _sync = Globals.ThisAddIn.SyncMng;
            PeerListBox.ItemsSource = _peers;

            // Prefill from the manager's current values (retained for the session).
            LocalIpText.Text = _sync.LocalIP;
            ReceiverPortBox.Text = _sync.LocalPort.ToString();
            BroadcastPortBox.Text = SyncDefaults.Port.ToString();
            SenderLocalPortBox.Text = _sync.LocalPort.ToString();
            ManualRemoteIpBox.Text = _sync.RemoteIP == SyncDefaults.Broadcast ? "" : _sync.RemoteIP;
            ManualRemotePortBox.Text = _sync.RemotePort.ToString();
            ManualLocalPortBox.Text = _sync.LocalPort.ToString();

            _loaded = true; // guard: ignore control events that fire during init
            ShowScreen(Screen.ModeSelect);
        }

        #region Navigation

        private void ShowScreen(Screen screen) {
            _screen = screen;
            ClearErrors();

            ModePanel.Visibility = Vis(screen == Screen.ModeSelect);
            ReceiverPanel.Visibility = Vis(screen == Screen.ReceiverSetup);
            SenderPanel.Visibility = Vis(screen == Screen.SenderSetup);

            bool isSetup = screen != Screen.ModeSelect;
            BackButton.Visibility = Vis(isSetup);
            StartButton.Visibility = Vis(isSetup);

            if (screen == Screen.SenderSetup) {
                UpdateMethodPanels();
                if (_method == SenderMethod.Auto) {
                    StartScan();
                }
            } else {
                StopScan();
            }
        }

        private void SenderCard_Click(object sender, RoutedEventArgs e) {
            ShowScreen(Screen.SenderSetup);
        }
        private void ReceiverCard_Click(object sender, RoutedEventArgs e) {
            ShowScreen(Screen.ReceiverSetup);
        }
        private void BackButton_Click(object sender, RoutedEventArgs e) {
            ShowScreen(Screen.ModeSelect);
        }
        private void CancelButton_Click(object sender, RoutedEventArgs e) {
            StopScan();
            Close();
        }
        private void Window_Closed(object sender, EventArgs e) {
            StopScan();
        }

        #endregion

        #region Sender method toggle

        private void AutoRadio_Checked(object sender, RoutedEventArgs e) {
            if (!_loaded) {
                return;
            }
            _method = SenderMethod.Auto;
            UpdateMethodPanels();
            StartScan();
        }
        private void ManualRadio_Checked(object sender, RoutedEventArgs e) {
            if (!_loaded) {
                return;
            }
            _method = SenderMethod.Manual;
            UpdateMethodPanels();
            StopScan();
        }
        private void UpdateMethodPanels() {
            AutoSubPanel.Visibility = Vis(_method == SenderMethod.Auto);
            ManualSubPanel.Visibility = Vis(_method == SenderMethod.Manual);
        }

        #endregion

        #region Discovery scan

        private void StartScan() {
            StopScan();
            _peers.Clear();
            BroadcastOption.IsChecked = false;
            ScanStatusText.Text = "검색 중…";

            _scanner = new DiscoveryScanner(_sync, Globals.ThisAddIn);
            _scanner.Start(DiscoveryRole.Unknown,
                peer => Dispatcher.Invoke(new Action(() => OnPeerFound(peer))),
                () => Dispatcher.Invoke(new Action(OnScanCompleted)));
        }
        private void StopScan() {
            if (_scanner != null) {
                _scanner.Stop();
                _scanner = null;
            }
        }
        private void OnPeerFound(DiscoveredPeer peer) {
            _peers.Add(new PeerRow(peer));
            ScanStatusText.Text = "";
        }
        private void OnScanCompleted() {
            ScanStatusText.Text = _peers.Count == 0 ? "수신기를 찾지 못했습니다." : "";
        }
        private void RescanButton_Click(object sender, RoutedEventArgs e) {
            StartScan();
        }

        // Broadcast and a specific peer are mutually exclusive targets.
        private void BroadcastOption_Checked(object sender, RoutedEventArgs e) {
            if (!_loaded) {
                return;
            }
            PeerListBox.SelectedIndex = -1;
        }
        private void PeerListBox_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            if (!_loaded) {
                return;
            }
            if (PeerListBox.SelectedItem != null) {
                BroadcastOption.IsChecked = false;
            }
        }

        #endregion

        #region Apply / start

        private void StartButton_Click(object sender, RoutedEventArgs e) {
            ClearErrors();
            switch (_screen) {
                case Screen.ReceiverSetup:
                    ApplyReceiver();
                    break;
                case Screen.SenderSetup:
                    if (_method == SenderMethod.Auto) {
                        ApplySenderAuto();
                    } else {
                        ApplySenderManual();
                    }
                    break;
            }
        }

        private void ApplyReceiver() {
            int port;
            if (!InputValidation.TryParsePort(ReceiverPortBox.Text, out port)) {
                ShowError(ReceiverError, "수신 포트는 1–65535 범위여야 합니다.");
                return;
            }
            _sync.Mode = SyncManager.SyncMode.RECEIVER;
            _sync.LocalPort = port;
            _sync.RemoteLabel = "";
            StartAndClose(ReceiverError);
        }

        private void ApplySenderAuto() {
            if (BroadcastOption.IsChecked == true) {
                int bport;
                if (!InputValidation.TryParsePort(BroadcastPortBox.Text, out bport)) {
                    ShowError(SenderError, "브로드캐스트 포트는 1–65535 범위여야 합니다.");
                    return;
                }
                _sync.RemoteIP = SyncDefaults.Broadcast;
                _sync.RemotePort = bport;
                _sync.RemoteLabel = "";
            } else {
                PeerRow row = PeerListBox.SelectedItem as PeerRow;
                if (row == null) {
                    ShowError(SenderError, "수신기를 선택하거나 ‘모든 수신기로 전송’을 선택하세요.");
                    return;
                }
                _sync.RemoteIP = row.Peer.IP;
                _sync.RemotePort = row.Peer.SyncPort;
                _sync.RemoteLabel = row.Peer.Host;
            }

            int localPort;
            if (!InputValidation.TryParsePort(SenderLocalPortBox.Text, out localPort)) {
                ShowError(SenderError, "로컬 포트는 1–65535 범위여야 합니다.");
                return;
            }
            _sync.Mode = SyncManager.SyncMode.SENDER;
            _sync.LocalPort = localPort;
            StartAndClose(SenderError);
        }

        private void ApplySenderManual() {
            if (!InputValidation.IsValidIPv4(ManualRemoteIpBox.Text)) {
                ShowError(SenderError, "원격 IP가 올바르지 않습니다 (IPv4만).");
                return;
            }
            int remotePort;
            if (!InputValidation.TryParsePort(ManualRemotePortBox.Text, out remotePort)) {
                ShowError(SenderError, "원격 포트는 1–65535 범위여야 합니다.");
                return;
            }
            int localPort;
            if (!InputValidation.TryParsePort(ManualLocalPortBox.Text, out localPort)) {
                ShowError(SenderError, "로컬 포트는 1–65535 범위여야 합니다.");
                return;
            }
            _sync.Mode = SyncManager.SyncMode.SENDER;
            _sync.RemoteIP = ManualRemoteIpBox.Text;
            _sync.RemotePort = remotePort;
            _sync.RemoteLabel = "";
            _sync.LocalPort = localPort;
            StartAndClose(SenderError);
        }

        private void StartAndClose(TextBlock errorBlock) {
            StopScan();
            if (_sync.StartSync()) {
                DialogResult = true;
                Close();
            } else {
                ShowError(errorBlock, "시작 실패 — 포트가 사용 중일 수 있습니다.");
            }
        }

        #endregion

        #region Helpers

        private static Visibility Vis(bool visible) {
            return visible ? Visibility.Visible : Visibility.Collapsed;
        }
        private static void ShowError(TextBlock block, string message) {
            block.Text = message;
            block.Visibility = Visibility.Visible;
        }
        private void ClearErrors() {
            ReceiverError.Visibility = Visibility.Collapsed;
            SenderError.Visibility = Visibility.Collapsed;
        }

        #endregion
    }
}
