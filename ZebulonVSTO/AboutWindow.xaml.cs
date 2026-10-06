using System;
using System.Diagnostics;
using System.Reflection;
using System.Windows;
using System.Windows.Navigation;

using ZebulonVSTO.Slides;

namespace ZebulonVSTO {
    /// <summary>
    /// Interaction logic for AboutWindow.xaml — the ribbon's "Zebulon 정보"
    /// dialog: logo, product/version/copyright (from the assembly's
    /// FileVersionInfo), links to the repo and Zebulon Web, and a button that
    /// copies the version line for support requests. Pure WPF; no Interop.
    /// </summary>
    public partial class AboutWindow {
        private const string RepoUrl = "https://github.com/kimhono97/zebulon-vsto";

#if DEBUG
        private const string BuildConfig = "Debug";
#else
        private const string BuildConfig = "Release";
#endif

        private readonly string _versionLine;

        public AboutWindow() {
            InitializeComponent();

            FileVersionInfo info = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location);
            ProductText.Text = info.ProductName;
            VersionText.Text = "버전 " + info.ProductVersion + (BuildConfig == "Debug" ? " (Debug)" : "");
            CopyrightText.Text = info.LegalCopyright;
            _versionLine = info.ProductName + " " + info.ProductVersion + " (" + BuildConfig + ")";

            RepoLink.NavigateUri = new Uri(RepoUrl);
            RepoLink.ToolTip = RepoUrl;
            // WebBaseUrl's root is not Zebulon Web itself; the app lives under /zebulon.
            string webUrl = SlideGenDefaults.WebBaseUrl.TrimEnd('/') + "/zebulon";
            WebLink.NavigateUri = new Uri(webUrl);
            WebLink.ToolTip = webUrl;
        }

        private void Link_RequestNavigate(object sender, RequestNavigateEventArgs e) {
            try {
                Process.Start(e.Uri.AbsoluteUri);
            } catch (Exception ex) {
                MessageBox.Show(this, "브라우저를 열 수 없습니다.\n" + ex.Message, Title);
            }
            e.Handled = true;
        }

        private void CopyButton_Click(object sender, RoutedEventArgs e) {
            try {
                Clipboard.SetText(_versionLine);
                CopyButton.Content = "복사됨 ✓";
            } catch (Exception ex) {
                // The clipboard can be transiently locked by another process.
                MessageBox.Show(this, "클립보드에 복사할 수 없습니다.\n" + ex.Message, Title);
            }
        }
    }
}
