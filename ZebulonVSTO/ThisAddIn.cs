using System;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Forms;
using System.Windows.Threading;
using Microsoft.Office.Interop.PowerPoint;

using ZebulonVSTO.Sync;

namespace ZebulonVSTO {
    public partial class ThisAddIn : ISyncLogger, ISlideController {
        // Gates whether the console may send custom (non-SENDER) commands.
        // TODO (deployment): drive from configuration and default to false.
        public bool DebugMode = true;

        private SyncManager _syncManager;
        private SyncConsole _syncConsole;
        private Dispatcher _dispatcher;

        private int _slideShowWindowIndex = -1;

        public SyncManager SyncMng {
            get { return _syncManager; }
        }

        public void ShowInfoDlg() {
            Assembly assembly = Assembly.GetExecutingAssembly();
            FileVersionInfo versionInfo = FileVersionInfo.GetVersionInfo(assembly.Location);

            string info = " * Name\t\t: " + versionInfo.ProductName;
            info += "\n * Version\t: " + versionInfo.ProductVersion;
            info += "\n * Copyright\t: " + versionInfo.LegalCopyright;

            MessageBox.Show(info, "About");
        }
        public void ShowSyncConsole() {
            _syncConsole = new SyncConsole(_syncConsole);
            _syncConsole.OpenWindow();
        }
        public void HideSyncConsole() {
            _syncConsole?.CloseWindow();
        }

        #region ISyncLogger

        public void Log(string line) {
            _syncConsole?.AppendLogLine(line);
        }
        public void LogError(string message, Exception error) {
            _syncConsole?.AppendLogLine("<!> ERROR : " + message + "\n" + error);
        }

        #endregion

        private SlideShowWindow GetCurrentSlideShowWindow() {
            if (_slideShowWindowIndex < 1 || _slideShowWindowIndex > Application.SlideShowWindows.Count) {
                _slideShowWindowIndex = -1;
                return null;
            }
            return Application.SlideShowWindows[_slideShowWindowIndex];
        }
        private void SetCurrentSlideShowWindow(SlideShowWindow window) {
            for (int i = 0; i < Application.SlideShowWindows.Count; i++) {
                if (window.Equals(Application.SlideShowWindows[i + 1])) {
                    _slideShowWindowIndex = i + 1;
                    return;
                }
            }
        }

        private void OnSelectSlide(SlideRange range) {
            int index = range.SlideIndex;
            Log("[Event] SelectSlide " + index);
            _syncManager.SendRequestMessage("select " + index);
        }
        private void OnSlideShow(SlideShowWindow window) {
            int index = window.View.Slide.SlideIndex;
            Log("[Event] SlideShow " + index);
            _syncManager.SendRequestMessage("showslide " + index);

            if (window != GetCurrentSlideShowWindow()) {
                SetCurrentSlideShowWindow(window);
            }
        }
        private void OnSlideShowEnd(Presentation presentation) {
            Log("[Event] SlideShowEnd");
            _syncManager.SendRequestMessage("hideslide");
            _slideShowWindowIndex = -1;
        }

        #region ISlideController

        public bool SelectSlide(int slideIndex) {
            if (_dispatcher == null) {
                return false;
            }
            bool result = false;
            try {
                _dispatcher.Invoke(() => {
                    Presentation presentation = Application.ActivePresentation;
                    if (presentation != null) {
                        slideIndex = Math.Min(Math.Max(slideIndex, 1), presentation.Slides.Count);
                        presentation.Slides[slideIndex].Select();
                        result = true;
                    }
                });
            } catch (Exception e) {
                LogError(string.Format("Failed to select slide {0}.", slideIndex), e);
            }
            return result;
        }
        public bool ShowSlide(int slideIndex) {
            if (_dispatcher == null) {
                return false;
            }
            bool result = false;
            try {
                _dispatcher.Invoke(() => {
                    SlideShowWindow window = GetCurrentSlideShowWindow();
                    if (window == null && Application.ActivePresentation != null) {
                        window = Application.ActivePresentation.SlideShowSettings.Run();
                        SetCurrentSlideShowWindow(window);
                    }
                    if (window != null) {
                        slideIndex = Math.Min(Math.Max(slideIndex, 1), window.Presentation.Slides.Count);
                        window.View.GotoSlide(slideIndex);
                        result = true;
                    }
                });
            } catch (Exception e) {
                LogError(string.Format("Failed to start slide show {0}.", slideIndex), e);
            }
            return result;
        }
        public bool HideSlide() {
            if (_dispatcher == null) {
                return false;
            }
            bool result = false;
            try {
                _dispatcher.Invoke(() => {
                    SlideShowWindow window = GetCurrentSlideShowWindow();
                    if (window != null) {
                        window.View.Exit();
                        result = true;
                    }
                    _slideShowWindowIndex = -1;
                });
            } catch (Exception e) {
                LogError("Failed to finish slide show.", e);
            }
            return result;
        }
        public bool Alert(string sender, string text) {
            if (_dispatcher == null) {
                return false;
            }
            bool result = false;
            try {
                _dispatcher.Invoke(() => {
                    result = MessageBox.Show(text, sender) == DialogResult.OK;
                });
            } catch (Exception e) {
                LogError("Failed to show alert.", e);
            }
            return result;
        }

        #endregion

        private void ThisAddIn_Startup(object sender, EventArgs e) {
            _syncManager = SyncManager.GetInstance();
            _syncManager.Attach(this, this, DebugMode);
            _syncConsole = new SyncConsole();
            _dispatcher = Dispatcher.CurrentDispatcher;

            _slideShowWindowIndex = -1;

            Application.SlideShowEnd += OnSlideShowEnd;
            Application.SlideShowNextSlide += OnSlideShow;
            Application.SlideSelectionChanged += OnSelectSlide;
        }
        private void ThisAddIn_Shutdown(object sender, EventArgs e) {
            if (_syncManager != null && _syncManager.IsRunning()) {
                _syncManager.StopSync();
            }
            _syncConsole = null;
            _dispatcher = null;
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject() {
            return new MainRibbon();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support — do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup() {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
