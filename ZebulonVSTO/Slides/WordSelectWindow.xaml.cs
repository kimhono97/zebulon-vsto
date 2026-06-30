using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace ZebulonVSTO.Slides {
    /// <summary>
    /// Interaction logic for WordSelectWindow.xaml — a modal that builds one Word
    /// passage (<see cref="WordItem"/>): up to 3 language/version slots (with ruby
    /// mode for ruby-capable versions), a book, a chapter, and one or more verse
    /// ranges. The verse ranges are picked by tapping verses in a preview list
    /// (tap a start verse, then a later verse to extend a range, then "범위 추가" to
    /// commit it) — mirroring the web SectionPicker / SlideEditorWord. On confirm it
    /// fetches each version's chapter from the Bible API and assembles the item.
    /// Provider I/O async.
    /// </summary>
    public partial class WordSelectWindow {
        private sealed class BookRow {
            public int Number { get; private set; }
            public string Display { get; private set; }
            public BookRow(int number, string display) { Number = number; Display = display; }
            public override string ToString() { return Display; }
        }

        /// <summary>A tappable verse row: number + preview text + range role.</summary>
        private sealed class VerseRow : INotifyPropertyChanged {
            public int Number { get; private set; }
            public string Text { get; private set; }

            private string _role = "";   // "", "single", "start", "end", "middle"
            private string _marker = "";  // "시작" / "끝" / ""

            public VerseRow(int number, string text) {
                Number = number;
                Text = text;
            }

            public string Role {
                get { return _role; }
                set { if (_role != value) { _role = value; Raise("Role"); } }
            }

            public string Marker {
                get { return _marker; }
                set { if (_marker != value) { _marker = value; Raise("Marker"); } }
            }

            public event PropertyChangedEventHandler PropertyChanged;
            private void Raise(string name) {
                PropertyChangedEventHandler h = PropertyChanged;
                if (h != null) {
                    h(this, new PropertyChangedEventArgs(name));
                }
            }
        }

        /// <summary>A committed-range chip ("S-E"), with its index in the selection.</summary>
        private sealed class SegmentChip {
            public int Index { get; private set; }
            public string Label { get; private set; }
            public SegmentChip(int index, string label) { Index = index; Label = label; }
        }

        private readonly BibleClient _bible = new BibleClient();
        private ComboBox[] _langCombos;
        private ComboBox[] _verCombos;
        private ComboBox[] _rubyCombos;
        private StackPanel[] _slots;

        private readonly ObservableCollection<BookRow> _bookView = new ObservableCollection<BookRow>();
        private List<BookRow> _allBooks = new List<BookRow>();

        private readonly ObservableCollection<VerseRow> _verses = new ObservableCollection<VerseRow>();
        private readonly ObservableCollection<SegmentChip> _chips = new ObservableCollection<SegmentChip>();
        private readonly VerseSelection _selection = new VerseSelection();

        // Cache fetched chapters by "code|book|chapter" so re-rendering (ruby/version
        // toggles) and the final all-versions assembly never refetch needlessly.
        private readonly Dictionary<string, BibleData> _chapterCache = new Dictionary<string, BibleData>();
        // Guards against a stale async chapter load overwriting a newer one.
        private int _chapterLoadGen;

        private int _bookNumber = -1;
        private bool _loaded;
        private bool _busy;
        private bool _closed;

        /// <summary>The assembled passage; non-null only when DialogResult == true.</summary>
        public WordItem Result { get; private set; }

        public WordSelectWindow() {
            InitializeComponent();
            _langCombos = new ComboBox[] { Lang0, Lang1, Lang2 };
            _verCombos = new ComboBox[] { Ver0, Ver1, Ver2 };
            _rubyCombos = new ComboBox[] { Ruby0, Ruby1, Ruby2 };
            _slots = new StackPanel[] { Slot0, Slot1, Slot2 };
            BookList.ItemsSource = _bookView;
            VerseList.ItemsSource = _verses;
            SegmentChips.ItemsSource = _chips;
            Loaded += delegate { Initialize(); };
        }

        private void Initialize() {
            for (int i = 1; i <= 3; i++) {
                LangCountCombo.Items.Add(i);
            }
            LangCountCombo.SelectedIndex = 2; // default language count = 3

            string[] defaults = { "ko-KR", "en-US", "zh-CN" };
            for (int slot = 0; slot < 3; slot++) {
                ComboBox lc = _langCombos[slot];
                foreach (BibleLanguage lang in LanguageCatalog.All) {
                    lc.Items.Add(lang);
                }
                int idx = IndexOfLanguage(defaults[slot]);
                lc.SelectedIndex = idx >= 0 ? idx : 0;

                ComboBox rb = _rubyCombos[slot];
                rb.Items.Add("한자만");
                rb.Items.Add("후리가나");
                rb.Items.Add("한자+후리가나");
                rb.SelectedIndex = 0;
            }

            _loaded = true;
            for (int slot = 0; slot < 3; slot++) {
                RefreshVersionCombo(slot);
            }
            RebuildBookList();
            UpdateSlotVisibility();
            RefreshSelection();
        }

        private static int IndexOfLanguage(string code) {
            for (int i = 0; i < LanguageCatalog.All.Count; i++) {
                if (string.Equals(LanguageCatalog.All[i].Code, code, StringComparison.OrdinalIgnoreCase)) {
                    return i;
                }
            }
            return -1;
        }

        private int LangCount {
            get { return LangCountCombo.SelectedIndex + 1; } // index 0->1, 1->2, 2->3
        }

        private void UpdateSlotVisibility() {
            for (int slot = 0; slot < 3; slot++) {
                _slots[slot].Visibility = slot < LangCount ? Visibility.Visible : Visibility.Collapsed;
            }
        }

        private BibleLanguage SlotLanguage(int slot) {
            return _langCombos[slot].SelectedItem as BibleLanguage;
        }

        private BibleVersion SlotVersion(int slot) {
            return _verCombos[slot].SelectedItem as BibleVersion;
        }

        private void RefreshVersionCombo(int slot) {
            ComboBox vc = _verCombos[slot];
            BibleLanguage lang = SlotLanguage(slot);
            vc.Items.Clear();
            if (lang != null) {
                foreach (BibleVersion v in BibleCatalog.VersionsForLanguage(lang.Code)) {
                    vc.Items.Add(v);
                }
            }
            if (vc.Items.Count > 0) {
                vc.SelectedIndex = 0;
            }
            UpdateRubyEnabled(slot);
        }

        private void UpdateRubyEnabled(int slot) {
            BibleVersion v = SlotVersion(slot);
            _rubyCombos[slot].IsEnabled = v != null && v.HasRuby;
        }

        private int SlotOf(object combo) {
            for (int i = 0; i < 3; i++) {
                if (ReferenceEquals(_langCombos[i], combo) || ReferenceEquals(_verCombos[i], combo)) {
                    return i;
                }
            }
            return -1;
        }

        private async void Lang_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            if (!_loaded) {
                return;
            }
            int slot = SlotOf(sender);
            if (slot < 0) {
                return;
            }
            RefreshVersionCombo(slot);
            if (slot == 0) {
                RebuildBookList();   // book list follows the primary language
                await LoadChapterAsync(); // primary version changed → re-render preview
            }
        }

        private async void Ver_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            if (!_loaded) {
                return;
            }
            int slot = SlotOf(sender);
            if (slot < 0) {
                return;
            }
            UpdateRubyEnabled(slot);
            if (slot == 0) {
                await LoadChapterAsync(); // primary version drives the preview text
            }
        }

        private async void Ruby_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            if (!_loaded) {
                return;
            }
            // Only the primary slot's ruby mode affects the preview rendering.
            if (ReferenceEquals(sender, Ruby0)) {
                await LoadChapterAsync();
            }
        }

        private void LangCountCombo_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            if (_loaded) {
                UpdateSlotVisibility();
            }
        }

        private void RebuildBookList() {
            BibleLanguage lang = SlotLanguage(0);
            string code = lang != null ? lang.Code : "en-US";
            _allBooks = new List<BookRow>();
            IReadOnlyList<string> names = BibleCatalog.AllBookNames(code);
            for (int i = 0; i < names.Count; i++) {
                _allBooks.Add(new BookRow(i + 1, (i + 1) + ". " + names[i]));
            }
            ApplyBookFilter();
        }

        private void ApplyBookFilter() {
            string f = (BookFilter.Text ?? "").Trim();
            _bookView.Clear();
            foreach (BookRow b in _allBooks) {
                if (f.Length == 0 || b.Display.IndexOf(f, StringComparison.OrdinalIgnoreCase) >= 0) {
                    _bookView.Add(b);
                }
            }
        }

        private void BookFilter_TextChanged(object sender, TextChangedEventArgs e) {
            if (_loaded) {
                ApplyBookFilter();
            }
        }

        private void BookList_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            BookRow b = BookList.SelectedItem as BookRow;
            if (b == null) {
                _bookNumber = -1;
                ChapterCombo.Items.Clear(); // → fires ChapterCombo_SelectionChanged → clears the verse list
                return;
            }
            _bookNumber = b.Number;
            int chapters = BibleCatalog.ChapterCount(_bookNumber);
            ChapterCombo.Items.Clear();
            for (int c = 1; c <= chapters; c++) {
                ChapterCombo.Items.Add(c);
            }
            if (ChapterCombo.Items.Count > 0) {
                ChapterCombo.SelectedIndex = 0; // → fires ChapterCombo_SelectionChanged → loads the chapter
            }
        }

        private async void ChapterCombo_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            if (!_loaded) {
                return;
            }
            // A book/chapter change resets the verse selection (mirrors the web).
            _selection.Reset();
            RefreshSelection();
            await LoadChapterAsync();
        }

        // (Re)load the primary version's chapter into the verse tap-list. Reads the
        // current book / primary version / chapter; clears the list when any is
        // unset. The verse selection is preserved (book/chapter handlers reset it).
        private async Task LoadChapterAsync() {
            if (!_loaded) {
                return;
            }
            int gen = ++_chapterLoadGen;
            BibleVersion v = SlotVersion(0);
            int chapter = ChapterCombo.SelectedItem is int c ? c : -1;
            if (_bookNumber < 1 || v == null || chapter < 1) {
                _verses.Clear();
                VerseCountText.Text = "";
                WordSelectStatus.Text = "";
                RefreshSelection();
                return;
            }
            WordSelectStatus.Text = "본문 불러오는 중…";
            try {
                BibleData data = await GetChapterCachedAsync(v.Code, _bookNumber, chapter);
                if (_closed || gen != _chapterLoadGen) {
                    return; // a newer load started, or the dialog closed
                }
                RenderVerses(data);
                WordSelectStatus.Text = _verses.Count == 0 ? "본문을 찾지 못했습니다." : "";
            } catch (Exception ex) {
                if (_closed || gen != _chapterLoadGen) {
                    return;
                }
                _verses.Clear();
                VerseCountText.Text = "";
                WordSelectStatus.Text = "오류: " + ex.Message;
                RefreshSelection();
            }
        }

        private void RenderVerses(BibleData data) {
            RubyMode mode = (RubyMode)(_rubyCombos[0].SelectedIndex < 0 ? 0 : _rubyCombos[0].SelectedIndex);
            List<string> verses = (data != null && data.Data != null) ? data.Data : new List<string>();
            _verses.Clear();
            for (int i = 0; i < verses.Count; i++) {
                string raw = verses[i] ?? "";
                // Plain-text the preview (a WPF placeholder can't render HTML); apply
                // the primary slot's ruby mode where the source carries ruby markup.
                string text = raw.IndexOf("<ruby>", StringComparison.OrdinalIgnoreCase) >= 0
                    ? RubyText.ApplyRubyMode(raw, mode)
                    : RubyText.ToPlainText(raw);
                _verses.Add(new VerseRow(i + 1, text));
            }
            VerseCountText.Text = "선택 가능한 절: " + verses.Count + "개";
            RefreshSelection();
        }

        // Recompute verse roles (in place — preserves scroll), rebuild the committed
        // chips, and refresh the pending/hint UI from the current selection.
        private void RefreshSelection() {
            foreach (VerseRow row in _verses) {
                string role = _selection.RoleOf(row.Number);
                row.Role = role;
                row.Marker = role == "start" ? "시작" : role == "end" ? "끝" : "";
            }

            _chips.Clear();
            IReadOnlyList<WordSection> segs = _selection.Segments;
            for (int i = 0; i < segs.Count; i++) {
                _chips.Add(new SegmentChip(i, RangeLabel(segs[i])));
            }

            if (_selection.HasPending) {
                PendingPanel.Visibility = Visibility.Visible;
                PendingLabel.Text = RangeLabel(_selection.Pending);
            } else {
                PendingPanel.Visibility = Visibility.Collapsed;
            }

            bool empty = segs.Count == 0 && !_selection.HasPending;
            SelectHint.Visibility = empty ? Visibility.Visible : Visibility.Collapsed;
        }

        private static string RangeLabel(WordSection s) {
            return s.End != s.Start ? s.Start + "-" + s.End : s.Start.ToString();
        }

        private void Verse_Click(object sender, RoutedEventArgs e) {
            VerseRow row = (sender as FrameworkElement)?.DataContext as VerseRow;
            if (row == null) {
                return;
            }
            _selection.Click(row.Number);
            RefreshSelection();
        }

        private void CommitPending_Click(object sender, RoutedEventArgs e) {
            _selection.CommitPending();
            RefreshSelection();
        }

        private void RemoveChip_Click(object sender, RoutedEventArgs e) {
            SegmentChip chip = (sender as FrameworkElement)?.DataContext as SegmentChip;
            if (chip == null) {
                return;
            }
            _selection.RemoveSegment(chip.Index);
            RefreshSelection();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e) {
            Close();
        }

        private async void AddButton_Click(object sender, RoutedEventArgs e) {
            if (_busy) {
                return;
            }
            if (_bookNumber < 1) {
                WordSelectStatus.Text = "책을 선택하세요.";
                return;
            }
            int chapter = ChapterCombo.SelectedItem is int c ? c : -1;
            if (chapter < 1) {
                WordSelectStatus.Text = "장을 선택하세요.";
                return;
            }
            List<WordSection> sections = _selection.DisplayRanges();
            if (sections.Count == 0) {
                WordSelectStatus.Text = "절 범위를 1개 이상 선택하세요.";
                return;
            }
            int count = LangCount;

            List<BibleVersion> versions = new List<BibleVersion>();
            List<RubyMode> rubyModes = new List<RubyMode>();
            for (int slot = 0; slot < count; slot++) {
                BibleVersion v = SlotVersion(slot);
                if (v == null) {
                    WordSelectStatus.Text = "버전을 선택하세요.";
                    return;
                }
                versions.Add(v);
                int rubyIdx = _rubyCombos[slot].SelectedIndex;
                rubyModes.Add((RubyMode)(rubyIdx < 0 ? 0 : rubyIdx)); // 0=Base,1=Furigana,2=Both
            }

            _busy = true;
            AddButton.IsEnabled = false;
            WordSelectStatus.Text = "성경 불러오는 중…";
            try {
                List<BibleData> chunk = new List<BibleData>();
                foreach (BibleVersion v in versions) {
                    chunk.Add(await GetChapterCachedAsync(v.Code, _bookNumber, chapter));
                }
                if (_closed) {
                    return; // the dialog was cancelled/closed during the fetch
                }
                WordItem item = WordAssembler.Build(chunk, chapter, sections, rubyModes);
                if (!HasAnyVerse(item)) {
                    WordSelectStatus.Text = "선택한 범위에 본문이 없습니다. 절 범위를 확인하세요.";
                    return;
                }
                Result = item;
                DialogResult = true;
                Close();
            } catch (Exception ex) {
                WordSelectStatus.Text = "오류: " + ex.Message;
            } finally {
                _busy = false;
                AddButton.IsEnabled = true;
            }
        }

        private async Task<BibleData> GetChapterCachedAsync(string code, int book, int chapter) {
            string key = code + "|" + book + "|" + chapter;
            BibleData cached;
            if (_chapterCache.TryGetValue(key, out cached)) {
                return cached;
            }
            BibleData data = await _bible.GetChapterAsync(code, book, chapter);
            if (data != null) {
                _chapterCache[key] = data;
            }
            return data;
        }

        protected override void OnClosed(EventArgs e) {
            _closed = true;
            base.OnClosed(e);
        }

        private static bool HasAnyVerse(WordItem item) {
            return HasText(item.First) || HasText(item.Second) || HasText(item.Third);
        }

        // Named HasText (not HasContent) to avoid shadowing ContentControl.HasContent.
        private static bool HasText(List<List<string>> sections) {
            if (sections == null) {
                return false;
            }
            foreach (List<string> sec in sections) {
                if (sec == null) {
                    continue;
                }
                foreach (string line in sec) {
                    if (!string.IsNullOrEmpty(line)) {
                        return true;
                    }
                }
            }
            return false;
        }
    }
}
