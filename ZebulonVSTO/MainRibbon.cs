using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

using ZebulonVSTO.Sync;

namespace ZebulonVSTO {
    [ComVisible(true)]
    public class MainRibbon : Office.IRibbonExtensibility {
        private Office.IRibbonUI _ribbon;
        private readonly Dictionary<string, bool> _enableMap;

        // Transient status/error shown in the ribbon's status label; cleared on
        // the next Start/Stop or settings change.
        private string _statusMessage;

        private SyncManager SyncMng {
            get { return Globals.ThisAddIn.SyncMng; }
        }

        public MainRibbon() {
            _enableMap = new Dictionary<string, bool>();
        }

        #region IRibbonExtensibility members

        public string GetCustomUI(string ribbonID) {
            return GetResourceText("ZebulonVSTO.MainRibbon.xml");
        }

        #endregion

        #region Ribbon callbacks

        public bool GetEnabled(IRibbonControl c) {
            string controlId = c.Id;
            bool enabled;
            if (!_enableMap.TryGetValue(controlId, out enabled)) {
                switch (controlId) {
                    case "DdMode":
                    case "EbLocalPort":
                        enabled = true;
                        break;
                    case "BtnSync":
                    case "BtnConsole":
                    case "EbRemoteIP":
                    case "EbRemotePort":
                        enabled = false;
                        break;
                }
                _enableMap.Add(controlId, enabled);
            }
            return enabled;
        }
        public string GetImage(IRibbonControl c) {
            switch (c.Id) {
                case "BtnSync":
                    return SyncMng.IsRunning() ? "RecordingStop" : "Synchronize";
            }
            return null; // 'no image' rather than an invalid imageMso lookup
        }
        public string GetLabel(IRibbonControl c) {
            switch (c.Id) {
                case "BtnSync":
                    return SyncMng.IsRunning() ? "동기화 중지" : "동기화 시작";
                case "LblStatus":
                    return BuildStatusText();
            }
            return "";
        }
        public string GetText(IRibbonControl c) {
            switch (c.Id) {
                case "EbLocalIP":
                    return SyncMng.LocalIP;
                case "EbLocalPort":
                    return SyncMng.LocalPort.ToString();
                case "EbRemoteIP":
                    return SyncMng.RemoteIP;
                case "EbRemotePort":
                    return SyncMng.RemotePort.ToString();
            }
            return "";
        }
        public void OnTextChange(IRibbonControl c, string text) {
            _statusMessage = null;
            switch (c.Id) {
                // EbLocalIP is read-only (auto-detected); no setter case.
                case "EbLocalPort": {
                    int port;
                    if (TryParsePort(text, out port)) {
                        SyncMng.LocalPort = port;
                    } else {
                        _statusMessage = "⚠ 잘못된 로컬 포트 (1–65535)";
                    }
                    break;
                }
                case "EbRemoteIP": {
                    IPAddress address;
                    if (IPAddress.TryParse(text, out address) && address.AddressFamily == AddressFamily.InterNetwork) {
                        SyncMng.RemoteIP = text;
                    } else {
                        _statusMessage = "⚠ 잘못된 원격 IP (IPv4만)";
                    }
                    break;
                }
                case "EbRemotePort": {
                    int port;
                    if (TryParsePort(text, out port)) {
                        SyncMng.RemotePort = port;
                    } else {
                        _statusMessage = "⚠ 잘못된 원격 포트 (1–65535)";
                    }
                    break;
                }
            }
            _ribbon?.Invalidate();
        }
        public int GetSelectedItemIndex(IRibbonControl c) {
            switch (c.Id) {
                case "DdMode":
                    switch (SyncMng.Mode) {
                        case SyncManager.SyncMode.SENDER: return 0;
                        case SyncManager.SyncMode.RECEIVER: return 1;
                    }
                    break;
            }
            return -1;
        }
        public void OnBtnAction(IRibbonControl c) {
            switch (c.Id) {
                case "BtnSync":
                    _statusMessage = null;
                    if (SyncMng.IsRunning()) {
                        Globals.ThisAddIn.HideSyncConsole();
                        SyncMng.StopSync();
                    } else if (!SyncMng.StartSync()) {
                        _statusMessage = "⚠ 시작 실패 — 포트 사용 중일 수 있음";
                    }
                    UpdateSyncSettingsUI();
                    break;
                case "BtnAbout":
                    Globals.ThisAddIn.ShowInfoDlg();
                    break;
                case "BtnConsole":
                    Globals.ThisAddIn.ShowSyncConsole();
                    break;
            }
        }
        public void OnDdAction(IRibbonControl c, string selectedId, int selectedIndex) {
            switch (c.Id) {
                case "DdMode":
                    _statusMessage = null;
                    switch (selectedId) {
                        case "ModeSender":
                            SyncMng.Mode = SyncManager.SyncMode.SENDER;
                            break;
                        case "ModeReceiver":
                            SyncMng.Mode = SyncManager.SyncMode.RECEIVER;
                            break;
                        default:
                            SyncMng.Mode = SyncManager.SyncMode.NONE;
                            break;
                    }
                    UpdateSyncSettingsUI();
                    break;
            }
        }

        private void UpdateSyncSettingsUI() {
            bool isRunning = SyncMng.IsRunning();
            bool isSenderMode = SyncMng.Mode == SyncManager.SyncMode.SENDER;
            bool isReceiverMode = SyncMng.Mode == SyncManager.SyncMode.RECEIVER;
            UpdateEnableMap("BtnSync", isSenderMode || isReceiverMode);
            UpdateEnableMap("BtnConsole", isRunning);
            UpdateEnableMap("DdMode", !isRunning);
            UpdateEnableMap("EbLocalPort", !isRunning);
            UpdateEnableMap("EbRemoteIP", !isRunning && isSenderMode);
            UpdateEnableMap("EbRemotePort", !isRunning && isSenderMode);
            _ribbon.Invalidate();
        }
        private void UpdateEnableMap(string key, bool value) {
            if (_enableMap.ContainsKey(key)) {
                _enableMap[key] = value;
            } else {
                _enableMap.Add(key, value);
            }
        }
        private string BuildStatusText() {
            if (!string.IsNullOrEmpty(_statusMessage)) {
                return _statusMessage;
            }
            if (!SyncMng.IsRunning()) {
                return "○ 중지됨";
            }
            if (SyncMng.Mode == SyncManager.SyncMode.RECEIVER) {
                return "● 수신 · " + SyncMng.LocalPort;
            }
            return "● 송신 · →" + SyncMng.RemoteIP + ":" + SyncMng.RemotePort;
        }
        private static bool TryParsePort(string text, out int port) {
            return int.TryParse(text, out port) && port >= 1 && port <= 65535;
        }
        public void Ribbon_Load(Office.IRibbonUI ribbonUI) {
            _ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName) {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i) {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0) {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i]))) {
                        if (resourceReader != null) {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
