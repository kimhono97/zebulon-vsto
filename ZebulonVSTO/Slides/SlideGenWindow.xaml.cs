using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace ZebulonVSTO.Slides {
    /// <summary>
    /// Interaction logic for SlideGenWindow.xaml — a modal wizard that inserts
    /// Zebulon slides directly into the active deck. Steps: pick type → pick a
    /// bind layout → pick lyrics (from the Provider) → pick insert position →
    /// generate. The type step can also download a template from the Provider,
    /// save it (user-chosen folder/name), open it, and continue. Lyric rows show
    /// a derived title + path; a preview dialog shows the selected lyric's text.
    /// Mirrors the SetupWizard pattern (Visibility-toggled panels, modal,
    /// PowerPoint-owned). Provider I/O is async; Interop runs through
    /// <see cref="ISlideBuilder"/> (marshalled to the UI thread by the host).
    /// </summary>
    public partial class SlideGenWindow {
        private enum Step { Type, Layout, Data, Position }
        private enum TemplateKind { Praise, Word }

        private sealed class LayoutItem {
            public LayoutMatch Match { get; private set; }
            public LayoutItem(LayoutMatch match) { Match = match; }
            public override string ToString() {
                string name = string.IsNullOrEmpty(Match.Name) ? "(이름 없음)" : Match.Name;
                return "#" + Match.LayoutIndex + "   " + name;
            }
        }

        private readonly ISlideBuilder _builder;
        private readonly ProviderClient _provider = new ProviderClient();

        private List<LayoutDescriptor> _layouts = new List<LayoutDescriptor>();
        private LayoutMatchResult _praiseMatch = new LayoutMatchResult();
        private LayoutMatchResult _wordMatch = new LayoutMatchResult();

        private TemplateKind _kind = TemplateKind.Praise;
        private Step _step = Step.Type;
        private LayoutMatch _selectedBind;
        private int _emptyLayoutIndex = -1;
        private string _targetDeck; // null/"" = active presentation; set when a template is downloaded & opened

        private List<LyricEntry> _allEntries = new List<LyricEntry>();
        private readonly ObservableCollection<LyricEntry> _available = new ObservableCollection<LyricEntry>();
        private readonly ObservableCollection<LyricEntry> _selected = new ObservableCollection<LyricEntry>();
        private readonly HashSet<string> _selectedPaths = new HashSet<string>();
        private readonly Dictionary<string, Lyric> _lyricCache = new Dictionary<string, Lyric>();
        private readonly ObservableCollection<WordItem> _wordItems = new ObservableCollection<WordItem>();

        private bool _lyricsLoaded;
        private bool _templatesLoaded;
        private bool _busy;

        public SlideGenWindow() {
            InitializeComponent();
            _builder = Globals.ThisAddIn;
            AvailableListBox.ItemsSource = _available;
            SelectedListBox.ItemsSource = _selected;
            WordItemList.ItemsSource = _wordItems;
            Loaded += OnLoaded;
        }

        private async void OnLoaded(object sender, RoutedEventArgs e) {
            Initialize();
            await LoadTemplatesAsync();
        }

        // Scans the active deck and refreshes the type-step state. Idempotent — also
        // called again after a template is downloaded and opened.
        private void Initialize() {
            _layouts = _builder.ReadLayouts(_targetDeck);
            _praiseMatch = LayoutMatching.Match(_layouts, SlideGenDefaults.PraiseBoxCount);
            if (_praiseMatch.IsAvailable) {
                PraiseStatus.Text = "사용 가능 — bind 레이아웃 " + _praiseMatch.BindCandidates.Count + "개";
                PraiseCard.IsEnabled = true;
            } else if (string.IsNullOrEmpty(_targetDeck) && !_builder.HasActivePresentation()) {
                PraiseStatus.Text = "열린 프레젠테이션이 없습니다.";
                PraiseCard.IsEnabled = false;
            } else if (_praiseMatch.BindCandidates.Count > 0 && !_praiseMatch.HasEmpty) {
                PraiseStatus.Text = "구분용 빈(empty) 레이아웃이 없어 사용 불가.";
                PraiseCard.IsEnabled = false;
            } else {
                PraiseStatus.Text = "이 deck에 적합한 Praise 레이아웃이 없습니다.";
                PraiseCard.IsEnabled = false;
            }

            _wordMatch = LayoutMatching.Match(_layouts, SlideGenDefaults.WordBoxCount);
            if (_wordMatch.IsAvailable) {
                WordStatus.Text = "사용 가능 — bind 레이아웃 " + _wordMatch.BindCandidates.Count + "개";
                WordCard.IsEnabled = true;
            } else if (string.IsNullOrEmpty(_targetDeck) && !_builder.HasActivePresentation()) {
                WordStatus.Text = "열린 프레젠테이션이 없습니다.";
                WordCard.IsEnabled = false;
            } else if (_wordMatch.BindCandidates.Count > 0 && !_wordMatch.HasEmpty) {
                WordStatus.Text = "구분용 빈(empty) 레이아웃이 없어 사용 불가.";
                WordCard.IsEnabled = false;
            } else {
                WordStatus.Text = "이 deck에 적합한 Word 레이아웃이 없습니다.";
                WordCard.IsEnabled = false;
            }
            ShowStep(Step.Type);
        }

        #region Navigation

        private void ShowStep(Step step) {
            _step = step;
            ClearErrors();
            TypePanel.Visibility = Vis(step == Step.Type);
            LayoutPanel.Visibility = Vis(step == Step.Layout);
            DataPanelPraise.Visibility = Vis(step == Step.Data && _kind == TemplateKind.Praise);
            DataPanelWord.Visibility = Vis(step == Step.Data && _kind == TemplateKind.Word);
            PositionPanel.Visibility = Vis(step == Step.Position);

            BackButton.Visibility = Vis(step != Step.Type);
            NextButton.Visibility = Vis(step != Step.Type); // cards advance on the type step
            NextButton.Content = step == Step.Position ? "생성" : "다음 >";
        }

        private void PraiseCard_Click(object sender, RoutedEventArgs e) {
            if (_busy) {
                return; // ignore navigation while a download/open is in flight
            }
            _kind = TemplateKind.Praise;
            EnterLayoutStep();
        }

        private void WordCard_Click(object sender, RoutedEventArgs e) {
            if (_busy) {
                return;
            }
            _kind = TemplateKind.Word;
            EnterLayoutStep();
        }

        private void EnterLayoutStep() {
            LayoutMatchResult match = _kind == TemplateKind.Word ? _wordMatch : _praiseMatch;
            LayoutHint.Text = _kind == TemplateKind.Word
                ? "조건(마커 $1·$2·$3·$4)을 충족하는 bind 레이아웃:"
                : "조건(마커 $1·$2·$3)을 충족하는 bind 레이아웃:";
            LayoutListBox.Items.Clear();
            foreach (LayoutMatch m in match.BindCandidates) {
                LayoutListBox.Items.Add(new LayoutItem(m));
            }
            if (LayoutListBox.Items.Count > 0) {
                LayoutListBox.SelectedIndex = 0;
            }
            _emptyLayoutIndex = match.EmptyLayoutIndex;
            EmptyLayoutText.Text = "구분(empty) 레이아웃: #" + _emptyLayoutIndex;
            ShowStep(Step.Layout);
        }

        private void LayoutListBox_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            LayoutItem item = LayoutListBox.SelectedItem as LayoutItem;
            _selectedBind = item != null ? item.Match : null;
        }

        private void BackButton_Click(object sender, RoutedEventArgs e) {
            if (_busy) {
                return;
            }
            switch (_step) {
                case Step.Layout: ShowStep(Step.Type); break;
                case Step.Data: ShowStep(Step.Layout); break;
                case Step.Position: ShowStep(Step.Data); break;
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e) {
            Close();
        }

        private async void NextButton_Click(object sender, RoutedEventArgs e) {
            if (_busy) {
                return;
            }
            switch (_step) {
                case Step.Layout:
                    if (_selectedBind == null) {
                        ShowError(LayoutError, "레이아웃을 선택하세요.");
                        return;
                    }
                    await EnterDataStepAsync();
                    break;
                case Step.Data:
                    if (_kind == TemplateKind.Word) {
                        if (_wordItems.Count == 0) {
                            ShowError(WordDataError, "구절을 1개 이상 추가하세요.");
                            return;
                        }
                    } else if (_selected.Count == 0) {
                        ShowError(DataError, "가사를 1개 이상 선택하세요.");
                        return;
                    }
                    ShowStep(Step.Position);
                    break;
                case Step.Position:
                    await GenerateAsync();
                    break;
            }
        }

        #endregion

        #region Template download (type step)

        private async Task LoadTemplatesAsync() {
            if (_templatesLoaded) {
                return;
            }
            TemplateStatus.Text = "템플릿 목록 불러오는 중…";
            try {
                List<string> paths = await _provider.ListTemplatesAsync();
                TemplateCombo.Items.Clear();
                foreach (string p in paths ?? new List<string>()) {
                    TemplateCombo.Items.Add(TemplateEntry.From(p));
                }
                if (TemplateCombo.Items.Count > 0) {
                    TemplateCombo.SelectedIndex = 0;
                }
                _templatesLoaded = true;
                TemplateStatus.Text = "";
            } catch (Exception ex) {
                TemplateStatus.Text = "템플릿 목록을 불러오지 못했습니다: " + ex.Message;
            }
        }

        private async void DownloadButton_Click(object sender, RoutedEventArgs e) {
            if (_busy) {
                return;
            }
            TemplateEntry tpl = TemplateCombo.SelectedItem as TemplateEntry;
            if (tpl == null) {
                TemplateStatus.Text = "템플릿을 선택하세요.";
                return;
            }
            SetBusy(true);
            TemplateStatus.Text = "다운로드 중…";
            try {
                byte[] bytes = await _provider.DownloadTemplateAsync(tpl.Path);

                // Let the user choose the destination folder AND filename.
                Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog {
                    Title = "다운로드한 템플릿 저장",
                    Filter = "PowerPoint 프레젠테이션 (*.pptx)|*.pptx",
                    DefaultExt = ".pptx",
                    AddExtension = true,
                    OverwritePrompt = true,
                    FileName = DisplayNames.SafePptxFileName(tpl.Path)
                };
                string dir = _builder.GetActivePresentationFolder();
                if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir)) {
                    dir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                }
                dlg.InitialDirectory = dir;
                if (dlg.ShowDialog(this) != true) {
                    TemplateStatus.Text = "저장이 취소되었습니다.";
                    return;
                }

                string path = dlg.FileName;
                // Write off the UI thread, via a temp file, so a failed/partial
                // write never leaves a truncated .pptx in place.
                await Task.Run(() => {
                    string tmp = path + ".download.tmp";
                    try {
                        File.WriteAllBytes(tmp, bytes);
                        if (File.Exists(path)) {
                            File.Delete(path); // may throw if the target is open in PowerPoint
                        }
                        File.Move(tmp, path);
                    } catch {
                        // don't leave the temp file orphaned when the move fails
                        try {
                            if (File.Exists(tmp)) {
                                File.Delete(tmp);
                            }
                        } catch {
                            // best-effort cleanup
                        }
                        throw;
                    }
                });
                TemplateStatus.Text = "저장됨: " + path;

                MessageBoxResult ans = MessageBox.Show(this,
                    "저장한 템플릿을 열고 이어서 슬라이드 생성을 진행할까요?",
                    "슬라이드 생성", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (ans == MessageBoxResult.Yes) {
                    string opened = _builder.OpenPresentation(path);
                    if (!string.IsNullOrEmpty(opened)) {
                        _targetDeck = opened; // bind the wizard to this exact deck
                        Initialize();         // re-scan; re-enables the Praise/Word cards
                        TemplateStatus.Text = (_praiseMatch.IsAvailable || _wordMatch.IsAvailable)
                            ? "열었습니다. 위에서 타입을 선택하세요."
                            : "열었지만 적합한 레이아웃이 없습니다.";
                    } else {
                        TemplateStatus.Text = "열기에 실패했습니다.";
                    }
                }
            } catch (Exception ex) {
                TemplateStatus.Text = "오류: " + ex.Message;
            } finally {
                SetBusy(false);
            }
        }

        #endregion

        #region Data step

        private async Task EnterDataStepAsync() {
            ShowStep(Step.Data);
            if (_kind == TemplateKind.Word) {
                return; // Word passages are added via the WordSelectWindow dialog
            }
            if (_lyricsLoaded) {
                return;
            }
            SetBusy(true);
            DataStatus.Text = "목록 불러오는 중…";
            try {
                List<string> paths = await _provider.ListLyricsAsync();
                _allEntries = (paths ?? new List<string>()).Select(LyricEntry.From).ToList();
                _lyricsLoaded = true;
                ApplyFilter(FilterBox.Text);
                DataStatus.Text = _allEntries.Count + "개";
            } catch (Exception ex) {
                DataStatus.Text = "";
                ShowError(DataError, "목록을 불러오지 못했습니다: " + ex.Message);
            } finally {
                SetBusy(false);
            }
        }

        private void FilterBox_TextChanged(object sender, TextChangedEventArgs e) {
            if (_lyricsLoaded) {
                ApplyFilter(FilterBox.Text);
            }
        }

        private void ApplyFilter(string filter) {
            filter = (filter ?? "").Trim();
            _available.Clear();
            foreach (LyricEntry entry in _allEntries) {
                if (_selectedPaths.Contains(entry.Path)) {
                    continue;
                }
                if (filter.Length == 0 || Matches(entry, filter)) {
                    _available.Add(entry);
                }
            }
        }

        private static bool Matches(LyricEntry entry, string filter) {
            return (entry.DisplayName != null && entry.DisplayName.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0)
                || (entry.Path != null && entry.Path.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private void AvailableListBox_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e) {
            AddSelected();
        }

        private void SelectedListBox_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e) {
            RemoveSelected();
        }

        private void AddButton_Click(object sender, RoutedEventArgs e) {
            AddSelected();
        }

        private void AddSelected() {
            List<LyricEntry> picks = AvailableListBox.SelectedItems.Cast<LyricEntry>().ToList();
            foreach (LyricEntry entry in picks) {
                if (_selectedPaths.Add(entry.Path)) {
                    _selected.Add(entry);
                }
            }
            ApplyFilter(FilterBox.Text);
        }

        private void RemoveButton_Click(object sender, RoutedEventArgs e) {
            RemoveSelected();
        }

        private void RemoveSelected() {
            List<LyricEntry> picks = SelectedListBox.SelectedItems.Cast<LyricEntry>().ToList();
            foreach (LyricEntry entry in picks) {
                _selected.Remove(entry);
                _selectedPaths.Remove(entry.Path);
            }
            ApplyFilter(FilterBox.Text);
        }

        private void UpButton_Click(object sender, RoutedEventArgs e) {
            MoveSelected(-1);
        }

        private void DownButton_Click(object sender, RoutedEventArgs e) {
            MoveSelected(1);
        }

        private void MoveSelected(int direction) {
            int i = SelectedListBox.SelectedIndex;
            if (i < 0) {
                return;
            }
            int j = i + direction;
            if (j < 0 || j >= _selected.Count) {
                return;
            }
            _selected.Move(i, j);
            SelectedListBox.SelectedIndex = j;
        }

        private async void PreviewButton_Click(object sender, RoutedEventArgs e) {
            if (_busy) {
                return;
            }
            LyricEntry entry = (AvailableListBox.SelectedItem as LyricEntry)
                            ?? (SelectedListBox.SelectedItem as LyricEntry);
            if (entry == null) {
                ShowError(DataError, "미리볼 항목을 선택하세요.");
                return;
            }
            ClearErrors();
            SetBusy(true);
            DataStatus.Text = "미리보기 불러오는 중…";
            try {
                Lyric lyric = await GetLyricCachedAsync(entry.Path);
                DataStatus.Text = "";
                if (lyric == null) {
                    ShowError(DataError, "가사를 불러오지 못했습니다.");
                    return;
                }
                LyricPreviewWindow win = new LyricPreviewWindow(lyric) { Owner = this };
                win.ShowDialog();
            } catch (Exception ex) {
                DataStatus.Text = "";
                ShowError(DataError, "미리보기 오류: " + ex.Message);
            } finally {
                SetBusy(false);
            }
        }

        private void AddPassage_Click(object sender, RoutedEventArgs e) {
            if (_busy) {
                return;
            }
            ClearErrors();
            WordSelectWindow win = new WordSelectWindow { Owner = this };
            bool? ok = win.ShowDialog();
            if (ok == true && win.Result != null) {
                _wordItems.Add(win.Result);
            }
        }

        private void RemovePassage_Click(object sender, RoutedEventArgs e) {
            WordItem item = WordItemList.SelectedItem as WordItem;
            if (item != null) {
                _wordItems.Remove(item);
            }
        }

        private void PassageUp_Click(object sender, RoutedEventArgs e) {
            MovePassage(-1);
        }

        private void PassageDown_Click(object sender, RoutedEventArgs e) {
            MovePassage(1);
        }

        private void MovePassage(int direction) {
            int i = WordItemList.SelectedIndex;
            if (i < 0) {
                return;
            }
            int j = i + direction;
            if (j < 0 || j >= _wordItems.Count) {
                return;
            }
            _wordItems.Move(i, j);
            WordItemList.SelectedIndex = j;
        }

        #endregion

        #region Generate

        private async Task GenerateAsync() {
            if (_selectedBind == null) {
                ShowStep(Step.Layout);
                return;
            }
            InsertPosition pos = PosFront.IsChecked == true ? InsertPosition.Front
                               : PosEnd.IsChecked == true ? InsertPosition.End
                               : InsertPosition.AfterCurrent;
            SetBusy(true);
            GenStatus.Text = "준비 중…";
            try {
                List<SlidePlanItem> plan;
                if (_kind == TemplateKind.Word) {
                    plan = WordPlanner.BuildPlan(_wordItems.ToList());
                } else {
                    GenStatus.Text = "가사 불러오는 중…";
                    List<Lyric> lyrics = new List<Lyric>();
                    foreach (LyricEntry entry in _selected) {
                        Lyric lyric = await GetLyricCachedAsync(entry.Path);
                        if (lyric != null) {
                            lyrics.Add(lyric);
                        }
                    }
                    if (lyrics.Count == 0) {
                        GenStatus.Text = "가사를 불러오지 못했습니다.";
                        return;
                    }
                    plan = PraisePlanner.BuildPlan(lyrics);
                }
                if (plan.Count == 0) {
                    GenStatus.Text = "생성할 슬라이드가 없습니다.";
                    return;
                }
                LayoutSelection sel = new LayoutSelection {
                    BindLayoutIndex = _selectedBind.LayoutIndex,
                    BoxSignatures = _selectedBind.BoxSignatures,
                    EmptyLayoutIndex = _emptyLayoutIndex,
                    // The CN-box centering nudge is Praise-only (faithful to the
                    // Exporter); Word must not move its 2nd-language box.
                    CenterCnBox = _kind == TemplateKind.Praise
                };
                GenStatus.Text = "슬라이드 삽입 중…";
                int inserted = _builder.ExecutePlan(_targetDeck, sel, plan, pos);
                if (inserted > 0) {
                    DialogResult = true;
                    Close();
                } else {
                    GenStatus.Text = "삽입된 슬라이드가 없습니다.";
                }
            } catch (Exception ex) {
                GenStatus.Text = "오류: " + ex.Message;
            } finally {
                SetBusy(false);
            }
        }

        private async Task<Lyric> GetLyricCachedAsync(string path) {
            Lyric cached;
            if (_lyricCache.TryGetValue(path, out cached)) {
                return cached;
            }
            Lyric lyric = await _provider.GetLyricAsync(path);
            if (lyric != null) {
                _lyricCache[path] = lyric;
            }
            return lyric;
        }

        #endregion

        #region Helpers

        private void SetBusy(bool busy) {
            _busy = busy;
            NextButton.IsEnabled = !busy;
            BackButton.IsEnabled = !busy;
            DownloadButton.IsEnabled = !busy;
        }

        private static Visibility Vis(bool visible) {
            return visible ? Visibility.Visible : Visibility.Collapsed;
        }

        private static void ShowError(TextBlock block, string message) {
            block.Text = message;
            block.Visibility = Visibility.Visible;
        }

        private void ClearErrors() {
            TypeError.Visibility = Visibility.Collapsed;
            LayoutError.Visibility = Visibility.Collapsed;
            DataError.Visibility = Visibility.Collapsed;
            WordDataError.Visibility = Visibility.Collapsed;
        }

        #endregion
    }
}
