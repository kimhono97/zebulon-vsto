using System;
using System.Windows;
using System.Windows.Input;

namespace ZebulonVSTO.Sync {
    /// <summary>
    /// SyncConsole.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class SyncConsole {
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
            LogText = "";
            CommandText = "";
        }
        public SyncConsole(SyncConsole pSC) {
            InitializeComponent();
            LogText = pSC.LogText;
            CommandText = "";
        }

        private void ProcessCommand() {
            if (CommandText.Length == 0) {
                return;
            }
            LogText += (CommandText + "\n");
            Globals.ThisAddIn.SyncMng.SendCustomMessage(CommandText);
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
