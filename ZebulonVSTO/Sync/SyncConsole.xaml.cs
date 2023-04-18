using System;
using System.Windows;
using System.Windows.Input;

namespace ZebulonVSTO.Sync {
    /// <summary>
    /// SyncConsole.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class SyncConsole {
        private bool bOpened = false;
        public string LogText {
            get { return this.TB_Log.Text; }
            set { this.TB_Log.Text = value; }
        }
        public string CommandText {
            get { return this.TB_Cmd.Text; }
            set { this.TB_Cmd.Text = value; }
        }

        public SyncConsole() {
            InitializeComponent();
            this.LogText = "";
            this.CommandText = "";
            this.bOpened = false;
        }
        public SyncConsole(SyncConsole pSC) {
            InitializeComponent();
            this.LogText = pSC.LogText;
            this.CommandText = "";
            this.SV_Log.ScrollToBottom();
            this.bOpened = false;
            pSC.CloseWindow();
        }

        public bool OpenWindow() {
            if (!this.bOpened) {
                Show();
                this.bOpened = true;
                return true;
            }
            return false;
        }
        public bool CloseWindow() {
            if (this.bOpened) {
                Close();
                this.bOpened = false;
                return true;
            }
            return false;
        }
        public void AppendLogLine(string strLine) {
            this.Dispatcher.Invoke(() => {
                string strTime = DateTime.Now.ToString("HH:mm:ss.ffff");
                LogText += (strTime + " " + strLine + "\n");
                this.SV_Log.ScrollToBottom();
            });
        }
        private void ProcessCommand() {
            if (CommandText.Length == 0) {
                return;
            }
            AppendLogLine(CommandText);
            Globals.ThisAddIn.SyncMng.SendRequestMessage(CommandText, true);
            CommandText = "";
        }

        private void Button_Click(object sender, RoutedEventArgs e) {
            ProcessCommand();
        }
        private void TB_Cmd_KeyDown(object sender, KeyEventArgs e) {
            if (e.Key == Key.Enter) {
                ProcessCommand();
            }
        }
    }
}
