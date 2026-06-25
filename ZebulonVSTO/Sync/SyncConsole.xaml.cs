using System;
using System.Windows;
using System.Windows.Input;

namespace ZebulonVSTO.Sync {
    /// <summary>
    /// Interaction logic for SyncConsole.xaml — a live traffic log plus a
    /// manual command input box.
    /// </summary>
    public partial class SyncConsole {
        private bool _opened;

        public string LogText {
            get { return LogTextBox.Text; }
            set { LogTextBox.Text = value; }
        }
        public string CommandText {
            get { return CommandTextBox.Text; }
            set { CommandTextBox.Text = value; }
        }

        public SyncConsole() {
            InitializeComponent();
            LogText = "";
            CommandText = "";
            _opened = false;
        }
        public SyncConsole(SyncConsole previous) {
            InitializeComponent();
            LogText = previous.LogText;
            CommandText = "";
            LogScrollViewer.ScrollToBottom();
            _opened = false;
            previous.CloseWindow();
        }

        public bool OpenWindow() {
            if (!_opened) {
                Show();
                _opened = true;
                return true;
            }
            return false;
        }
        public bool CloseWindow() {
            if (_opened) {
                Close();
                _opened = false;
                return true;
            }
            return false;
        }
        public void AppendLogLine(string line) {
            Dispatcher.Invoke(() => {
                string time = DateTime.Now.ToString("HH:mm:ss.ffff");
                LogText += time + " " + line + "\n";
                LogScrollViewer.ScrollToBottom();
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

        private void EnterButton_Click(object sender, RoutedEventArgs e) {
            ProcessCommand();
        }
        private void CommandTextBox_KeyDown(object sender, KeyEventArgs e) {
            if (e.Key == Key.Enter) {
                ProcessCommand();
            }
        }
    }
}
