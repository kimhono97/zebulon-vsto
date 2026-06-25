using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

using ZebulonVSTO.Sync;

namespace ZebulonVSTO {
    [ComVisible(true)]
    public class MainRibbon : Office.IRibbonExtensibility {
        private Office.IRibbonUI _ribbon;
        private readonly Dictionary<string, bool> _enableMap;

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
            return "";
        }
        public string GetLabel(IRibbonControl c) {
            switch (c.Id) {
                case "BtnSync":
                    return SyncMng.IsRunning() ? "Stop Sync" : "Start Sync";
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
            string controlId = c.Id;
            try {
                switch (controlId) {
                    // EbLocalIP is read-only (auto-detected); no setter case.
                    case "EbLocalPort":
                        SyncMng.LocalPort = int.Parse(text);
                        break;
                    case "EbRemoteIP":
                        SyncMng.RemoteIP = text;
                        break;
                    case "EbRemotePort":
                        SyncMng.RemotePort = int.Parse(text);
                        break;
                }
            } catch (Exception e) {
                Console.WriteLine(e.ToString());
            }
            _ribbon.InvalidateControl(controlId);
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
                    if (SyncMng.IsRunning()) {
                        Globals.ThisAddIn.HideSyncConsole();
                        SyncMng.StopSync();
                        UpdateSyncSettingsUI();
                    } else {
                        if (SyncMng.StartSync()) {
                            UpdateSyncSettingsUI();
                        }
                    }
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
