using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;

namespace ZebulonVSTO.Sync {
    /// <summary>
    /// Interaction logic for SyncConsole.xaml — a color-coded live traffic log
    /// (RichTextBox) plus a manual command input box.
    /// </summary>
    public partial class SyncConsole {
        private enum LogLevel { Default, Request, Response, Event, Error, Control }

        private sealed class LogEntry {
            public LogLevel Level;
            public string Text;
        }

        private bool _opened;
        private readonly List<LogEntry> _history = new List<LogEntry>();
        private Paragraph _paragraph;

        public string CommandText {
            get { return CommandTextBox.Text; }
            set { CommandTextBox.Text = value; }
        }

        public SyncConsole() {
            InitializeComponent();
            InitializeDocument();
            CommandText = "";
            _opened = false;
        }
        public SyncConsole(SyncConsole previous) {
            InitializeComponent();
            InitializeDocument();
            CommandText = "";
            _opened = false;
            foreach (LogEntry entry in previous._history) {
                AppendEntry(entry.Level, entry.Text);
            }
            LogRichTextBox.ScrollToEnd();
            previous.CloseWindow();
        }

        private void InitializeDocument() {
            _paragraph = new Paragraph { Margin = new Thickness(0) };
            LogRichTextBox.Document = new FlowDocument(_paragraph) {
                PagePadding = new Thickness(2, 0, 0, 0),
                Background = Brushes.Black,
                FontFamily = LogRichTextBox.FontFamily
            };
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
                AppendEntry(LevelFor(line), time + " " + line);
                LogRichTextBox.ScrollToEnd();
            });
        }

        private void AppendEntry(LogLevel level, string text) {
            _history.Add(new LogEntry { Level = level, Text = text });
            if (_paragraph.Inlines.Count > 0) {
                _paragraph.Inlines.Add(new LineBreak());
            }
            _paragraph.Inlines.Add(new Run(text) { Foreground = BrushFor(level) });
        }

        private static LogLevel LevelFor(string line) {
            if (string.IsNullOrEmpty(line)) {
                return LogLevel.Default;
            }
            if (line.IndexOf("ERROR", StringComparison.Ordinal) >= 0 || line.StartsWith("<!>")) {
                return LogLevel.Error;
            }
            if (line.StartsWith("--->") || line.StartsWith("<---")) {
                return LogLevel.Control;
            }
            if (line.StartsWith("[Event]")) {
                return LogLevel.Event;
            }
            if (line.IndexOf("REQ", StringComparison.Ordinal) >= 0) {
                return LogLevel.Request;
            }
            if (line.IndexOf("RES", StringComparison.Ordinal) >= 0) {
                return LogLevel.Response;
            }
            return LogLevel.Default;
        }

        private static Brush BrushFor(LogLevel level) {
            switch (level) {
                case LogLevel.Request: return Brushes.DeepSkyBlue;
                case LogLevel.Response: return Brushes.LightGreen;
                case LogLevel.Event: return Brushes.Silver;
                case LogLevel.Error: return Brushes.Salmon;
                case LogLevel.Control: return Brushes.Khaki;
                default: return Brushes.White;
            }
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
        private void ClearButton_Click(object sender, RoutedEventArgs e) {
            _history.Clear();
            _paragraph.Inlines.Clear();
        }
        private void CopyButton_Click(object sender, RoutedEventArgs e) {
            StringBuilder sb = new StringBuilder();
            foreach (LogEntry entry in _history) {
                sb.AppendLine(entry.Text);
            }
            try {
                Clipboard.SetText(sb.ToString());
            } catch (Exception) {
                // clipboard may be locked by another process; ignore
            }
        }
    }
}
