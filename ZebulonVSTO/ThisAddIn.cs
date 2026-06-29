using System;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Forms;
using System.Windows.Threading;
using Microsoft.Office.Interop.PowerPoint;

using ZebulonVSTO.Sync;

namespace ZebulonVSTO {
    public partial class ThisAddIn : ISyncLogger, ISlideController, IStatusObserver {
        // Gates whether the console may send custom (non-SENDER) commands.
        // Tied to the build configuration: on in Debug, OFF in Release (the
        // shipped/production build) since the DEBUG constant is only defined
        // for Debug builds.
#if DEBUG
        public bool DebugMode = true;
#else
        public bool DebugMode = false;
#endif

        private SyncManager _syncManager;
        private DiscoveryResponder _discoveryResponder;
        private SyncConsole _syncConsole;
        private MainRibbon _mainRibbon;
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
        public void ShowSetupWizard() {
            // Modal, owned by the PowerPoint main window so it stays on top and
            // blocks ribbon re-entry while the user configures a session.
            SetupWizard wizard = new SetupWizard();
            new System.Windows.Interop.WindowInteropHelper(wizard).Owner = (IntPtr)Application.HWND;
            wizard.ShowDialog();
        }

        #region ISyncLogger

        public void Log(string line) {
            _syncConsole?.AppendLogLine(line);
        }
        public void LogError(string message, Exception error) {
            _syncConsole?.AppendLogLine("<!> ERROR : " + message + "\n" + error);
        }

        #endregion

        #region IStatusObserver

        public void OnPeerChanged() {
            // Raised on the sync receive thread; marshal the ribbon refresh onto
            // the UI thread (the ribbon API is UI-thread affine).
            _dispatcher?.BeginInvoke(new Action(() => _mainRibbon?.RefreshPeerStatus()));
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
            _syncManager.Attach(this, this, this, DebugMode);
            _syncConsole = new SyncConsole();
            _dispatcher = Dispatcher.CurrentDispatcher;

            // Always-on peer-discovery responder (best-effort; failure to bind
            // the discovery port does not block the add-in or manual setup).
            _discoveryResponder = new DiscoveryResponder(_syncManager, this);
            _discoveryResponder.Start();

            _slideShowWindowIndex = -1;

            Application.SlideShowEnd += OnSlideShowEnd;
            Application.SlideShowNextSlide += OnSlideShow;
            Application.SlideSelectionChanged += OnSelectSlide;
        }
        private void ThisAddIn_Shutdown(object sender, EventArgs e) {
            _discoveryResponder?.Stop();
            _discoveryResponder = null;
            if (_syncManager != null && _syncManager.IsRunning()) {
                _syncManager.StopSync();
            }
            _syncConsole = null;
            _dispatcher = null;
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject() {
            _mainRibbon = new MainRibbon();
            return _mainRibbon;
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
