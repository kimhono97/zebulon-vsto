using System;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;
using System.Diagnostics;

using ZebulonVSTO.Sync;
using System.Windows.Threading;

namespace ZebulonVSTO {
    public partial class ThisAddIn {
        private SyncManager pSMng = null;
        private SyncConsole pSCsl = null;
        private Dispatcher pDispatcher = null;
        public SyncManager SyncMng {
            get { return this.pSMng; }
        }

        public void ShowInfoDlg() {
            System.Reflection.Assembly pAssembly = System.Reflection.Assembly.GetExecutingAssembly();
            FileVersionInfo pInfo = FileVersionInfo.GetVersionInfo(pAssembly.Location);
            
            string strInfo = (" * Name\t\t: " + pInfo.ProductName + "");
            strInfo += ("\n * Version\t: " + pInfo.ProductVersion);
            strInfo += ("\n * Copyright\t: " + pInfo.LegalCopyright);

            MessageBox.Show(strInfo, "About");
        }
        public void ShowSyncConsole() {
            this.pSCsl = new SyncConsole(this.pSCsl);
            this.pSCsl.Show();
        }
        public void HideSyncConsole() {
            this.pSCsl.Close();
        }
        public void LogDebug(string strLine) {
            this.pDispatcher.Invoke(() => {
                this.pSCsl.LogText += (strLine + "\n");
            });
        }
        public void LogError(string strMsg, Exception e) {
            this.pDispatcher.Invoke(() => {
                this.pSCsl.LogText += ("<!> ERROR : " + strMsg + "\n" + e.ToString() + "\n");
            });
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e) {
            this.pSMng = SyncManager.GetInstance();
            this.pSCsl = new SyncConsole();
            this.pDispatcher = Dispatcher.CurrentDispatcher;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e) {
        }
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject() {
            return new MainRibbon();
        }

        #region VSTO에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InternalStartup() {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
