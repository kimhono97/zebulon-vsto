using Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

// TODO:  리본(XML) 항목을 설정하려면 다음 단계를 수행하십시오:

// 1. 다음 코드 블록을 ThisAddin, ThisWorkbook 또는 ThisDocument 클래스에 복사합니다.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new MainRibbon();
//  }

// 2. 단추 클릭 등의 사용자 작업을 처리하려면 이 클래스의 "리본 콜백" 영역에서 콜백
//    메서드를 만듭니다. 참고: 리본 디자이너에서 이 리본을 내보낸 경우 이벤트 처리기의 코드를
//    콜백 메서드로 이동하고 리본 확장성(RibbonX) 프로그래밍 모델에서 사용할 수 있도록
//    코드를 수정해야 합니다.

// 3. 리본 XML 파일의 컨트롤 태그에 특성을 할당하여 사용자 코드의 적절한 콜백 메서드를 식별합니다.  

// 자세한 내용은 Visual Studio Tools for Office 도움말에서 리본 XML 설명서를 참조하십시오.


namespace ZebulonVSTO {
    [ComVisible(true)]
    public class MainRibbon: Office.IRibbonExtensibility {
        private Office.IRibbonUI ribbon;

        private Dictionary<string, bool> pEnableMap;
        
        private SyncManager SyncMng {
            get { return Globals.ThisAddIn.SyncMng; }
        }

        public MainRibbon() {
            this.pEnableMap = new Dictionary<string, bool>();
        }

        #region IRibbonExtensibility 멤버

        public string GetCustomUI(string ribbonID) {
            return GetResourceText("ZebulonVSTO.MainRibbon.xml");
        }

        public bool getEnabled(IRibbonControl c) {
            string strCompID = c.Id;
            bool bRet = false;
            if (!this.pEnableMap.TryGetValue(strCompID, out bRet)) {
                switch (strCompID) {
                    case "DdMode":
                        bRet = true;
                        break;
                    case "BtnSync":
                    case "EbLocalPort":
                    case "EbRemoteIP":
                    case "EbRemotePort":
                        bRet = false;
                        break;
                }
                this.pEnableMap.Add(strCompID, bRet);
            }
            return bRet;
        }
        public string getImage(IRibbonControl c) {
            string strCompID = c.Id;
            switch (strCompID) {
                case "BtnSync":
                    return SyncMng.IsRunning() ? "RecordingStop" : "Synchronize";
            }
            return "";
        }
        public string getLabel(IRibbonControl c) {
            string strCompID = c.Id;
            switch (strCompID) {
                case "BtnSync":
                    return SyncMng.IsRunning() ? "Stop Sync" : "Start Sync";
            }
            return "";
        }
        public string getText(IRibbonControl c) {
            string strCompID = c.Id;
            switch (strCompID) {
                case "EbLocalIP":
                    return SyncMng.LocalIP;
            }
            return "";
        }
        public int getSelectedItemIndex(IRibbonControl c) {
            string strCompID = c.Id;
            switch (strCompID) {
                case "DdMode":
                    switch (SyncMng.Mode) {
                        case SyncManager.SyncMode.SENDER:   return 0;
                        case SyncManager.SyncMode.RECEIVER: return 1;
                    }
                    break;
            }
            return -1;
        }

        public void OnBtnAction(IRibbonControl c) {
            string strCompID = c.Id;
            switch (strCompID) {
                case "BtnSync":
                    if (SyncMng.IsRunning()) {
                        SyncMng.StopSync();
                        updateSyncSettingUI();
                    } else {
                        if (SyncMng.StartSync()) {
                            updateSyncSettingUI();
                        }
                    }
                    break;
                case "BtnAbout":
                    Globals.ThisAddIn.ShowInfoDlg();
                    break;
            }
        }
        public void OnDdAction(IRibbonControl c, string strSelectedID, int nSelectedIndex) {
            string strCompID = c.Id;
            switch (strCompID) {
                case "DdMode":
                    //MessageBox.Show("ID=" + strSelectedID + ", Index=" + nSelectedIndex);
                    switch (strSelectedID) {
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
                    updateSyncSettingUI();
                    break;
            }
        }

        private void updateSyncSettingUI() {
            bool bRunning = SyncMng.IsRunning();
            bool bSndMode = (SyncMng.Mode == SyncManager.SyncMode.SENDER);
            bool bRcvMode = (SyncMng.Mode == SyncManager.SyncMode.RECEIVER);
            updateEnableMap("BtnSync", bSndMode || bRcvMode);
            updateEnableMap("DdMode", !bRunning);
            updateEnableMap("EbLocalPort", !bRunning && bRcvMode);
            updateEnableMap("EbRemoteIP", !bRunning && bSndMode);
            updateEnableMap("EbRemotePort", !bRunning && bSndMode);
            this.ribbon.Invalidate();
        }
        private void updateEnableMap(string strKey, bool bValue) {
            if (this.pEnableMap.ContainsKey(strKey)) {
                this.pEnableMap[strKey] = bValue;
            } else {
                this.pEnableMap.Add(strKey, bValue);
            }
        }

        #endregion

        #region 리본 콜백
        //여기서 콜백 메서드를 만듭니다. 콜백 메서드를 추가하는 방법에 대한 자세한 내용은 https://go.microsoft.com/fwlink/?LinkID=271226을 참조하세요.

        public void Ribbon_Load(Office.IRibbonUI ribbonUI) {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region 도우미

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
