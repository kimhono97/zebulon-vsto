using System;
using System.Windows.Forms;
using System.Diagnostics;
using System.Windows.Threading;
using Microsoft.Office.Interop.PowerPoint;

using ZebulonVSTO.Sync;

namespace ZebulonVSTO {
    public partial class ThisAddIn {
        public bool DEBUG_MODE = true;

        private SyncManager pSMng = null;
        private SyncConsole pSCsl = null;
        private Dispatcher pDispatcher = null;

        private int nSlideShowWndIndex = -1;
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
            this.pSCsl.OpenWindow();
        }
        public void HideSyncConsole() {
            this.pSCsl.CloseWindow();
        }
        public void LogDebug(string strLine) {
            this.pSCsl.AppendLogLine(strLine);
        }
        public void LogError(string strMsg, Exception e) {
            this.pSCsl.AppendLogLine("<!> ERROR : " + strMsg + "\n" + e.ToString());
        }

        private SlideShowWindow GetCurrentSlideShowWnd() {
            if (this.nSlideShowWndIndex < 1 || this.nSlideShowWndIndex > this.Application.SlideShowWindows.Count) {
                this.nSlideShowWndIndex = -1;
                return null;
            }
            return this.Application.SlideShowWindows[this.nSlideShowWndIndex];
        }
        private void SetCurrentSlideShowWnd(SlideShowWindow pWnd) {
            for (int i = 0; i < this.Application.SlideShowWindows.Count; i++) {
                if (pWnd.Equals(this.Application.SlideShowWindows[i + 1])) {
                    this.nSlideShowWndIndex = i + 1;
                    return;
                }
            }
        }
        private void OnSelectSlide(SlideRange pRange) {
            int nIndex = pRange.SlideIndex;
            LogDebug("[Event] SelectSlide " + nIndex.ToString());
            this.pSMng.SendRequestMessage("select " + nIndex.ToString());
        }
        private void OnSlideShow(SlideShowWindow pWnd) {
            int nIndex = pWnd.View.Slide.SlideIndex;
            LogDebug("[Event] SlideShow " + nIndex.ToString());
            this.pSMng.SendRequestMessage("showslide " + nIndex.ToString());

            if (pWnd != GetCurrentSlideShowWnd()) {
                SetCurrentSlideShowWnd(pWnd);
            }
        }
        private void OnSlideShowEnd(Presentation pPres) {
            LogDebug("[Event] SlideShowEnd");
            this.pSMng.SendRequestMessage("hideslide");
            this.nSlideShowWndIndex = -1;
        }
        public bool DoSelectSlide(int nSlideIndex) {
            bool bRet = false;
            try {
                this.pDispatcher.Invoke(() => {
                    Presentation pPres = this.Application.ActivePresentation;
                    if (pPres != null) {
                        nSlideIndex = Math.Min(Math.Max(nSlideIndex, 1), pPres.Slides.Count);
                        pPres.Slides[nSlideIndex].Select();
                        bRet = true;
                    }
                });
            } catch (Exception e) {
                LogError(String.Format("Failed to select slide {0}.", nSlideIndex), e);
            }
            return bRet;
        }
        public bool DoSlideShow(int nSlideIndex) {
            bool bRet = false;
            try {
                this.pDispatcher.Invoke(() => {
                    SlideShowWindow pWnd = GetCurrentSlideShowWnd();
                    if (pWnd == null && this.Application.ActivePresentation != null) {
                        pWnd = this.Application.ActivePresentation.SlideShowSettings.Run();
                        SetCurrentSlideShowWnd(pWnd);
                    }
                    if (pWnd != null) {
                        nSlideIndex = Math.Min(Math.Max(nSlideIndex, 1), pWnd.Presentation.Slides.Count);
                        pWnd.View.GotoSlide(nSlideIndex);
                        bRet = true;
                    }
                });
            } catch (Exception e) {
                LogError(String.Format("Failed to start slide show {0}.", nSlideIndex), e);
            }
            return bRet;
        }
        public bool DoSlideShowEnd() {
            bool bRet = false;
            try {
                this.pDispatcher.Invoke(() => {
                    SlideShowWindow pWnd = GetCurrentSlideShowWnd();
                    if (pWnd != null) {
                        pWnd.View.Exit();
                        bRet = true;
                    }
                    this.nSlideShowWndIndex = -1;
                });
            } catch (Exception e) {
                LogError(String.Format("Failed to finish slide show."), e);
            }
            return bRet;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e) {
            this.pSMng = SyncManager.GetInstance();
            this.pSCsl = new SyncConsole();
            this.pDispatcher = Dispatcher.CurrentDispatcher;

            this.nSlideShowWndIndex = -1;

            this.Application.SlideShowEnd += this.OnSlideShowEnd;
            this.Application.SlideShowNextSlide += this.OnSlideShow;
            this.Application.SlideSelectionChanged += this.OnSelectSlide;
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e) {
            if (this.pSMng.IsRunning()) {
                this.pSMng.StopSync();
            }
            this.pSCsl = null;
            this.pDispatcher = null;
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
