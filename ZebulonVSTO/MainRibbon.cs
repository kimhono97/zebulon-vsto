using Microsoft.Office.Core;
using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

using ZebulonVSTO.Sync;

namespace ZebulonVSTO {
    [ComVisible(true)]
    public class MainRibbon : Office.IRibbonExtensibility {
        private Office.IRibbonUI _ribbon;

        private SyncManager SyncMng {
            get { return Globals.ThisAddIn.SyncMng; }
        }

        #region IRibbonExtensibility members

        public string GetCustomUI(string ribbonID) {
            return GetResourceText("ZebulonVSTO.MainRibbon.xml");
        }

        #endregion

        #region Ribbon callbacks

        public bool GetEnabled(IRibbonControl c) {
            switch (c.Id) {
                case "BtnConsole":
                    return SyncMng.IsRunning();
            }
            return true;
        }
        public bool GetVisible(IRibbonControl c) {
            switch (c.Id) {
                // Self/peer lines appear only while running (minimal when stopped).
                case "LblLocal":
                case "LblRemote":
                    return SyncMng.IsRunning();
            }
            return true;
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
                case "LblState":
                    return BuildStateLine();
                case "LblLocal":
                    return BuildLocalLine();
                case "LblRemote":
                    return BuildRemoteLine();
            }
            return "";
        }
        public void OnBtnAction(IRibbonControl c) {
            switch (c.Id) {
                case "BtnSync":
                    if (SyncMng.IsRunning()) {
                        Globals.ThisAddIn.HideSyncConsole();
                        SyncMng.StopSync();
                    } else {
                        // Modal wizard collects settings and starts sync on finish.
                        Globals.ThisAddIn.ShowSetupWizard();
                    }
                    _ribbon?.Invalidate();
                    break;
                case "BtnAbout":
                    Globals.ThisAddIn.ShowInfoDlg();
                    break;
                case "BtnConsole":
                    Globals.ThisAddIn.ShowSyncConsole();
                    break;
            }
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI) {
            _ribbon = ribbonUI;
        }

        /// <summary>Refresh the peer status line after a peer change (invoked by
        /// the host on the UI thread). Cheap — invalidates a single label.</summary>
        public void RefreshPeerStatus() {
            _ribbon?.InvalidateControl("LblRemote");
        }

        #endregion

        #region Status text

        private string BuildStateLine() {
            if (!SyncMng.IsRunning()) {
                return "○ 중지됨";
            }
            return SyncMng.Mode == SyncManager.SyncMode.RECEIVER ? "● 수신 중" : "● 송신 중";
        }
        private string BuildLocalLine() {
            return "로컬  " + SyncMng.LocalIP + " : " + SyncMng.LocalPort;
        }
        private string BuildRemoteLine() {
            if (SyncMng.Mode == SyncManager.SyncMode.RECEIVER) {
                return string.IsNullOrEmpty(SyncMng.LastPeerIP)
                    ? "송신자  대기 중…"
                    : "송신자  " + SyncMng.LastPeerIP;
            }
            // SENDER: show the configured target.
            if (SyncMng.RemoteIP == SyncDefaults.Broadcast) {
                return "원격 → 전체(브로드캐스트) : " + SyncMng.RemotePort;
            }
            string label = SyncMng.RemoteLabel;
            if (!string.IsNullOrEmpty(label)) {
                return "원격 → " + label + "(" + SyncMng.RemoteIP + ") : " + SyncMng.RemotePort;
            }
            return "원격 → " + SyncMng.RemoteIP + " : " + SyncMng.RemotePort;
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
