using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace ZebulonVSTO.Slides {
    /// <summary>
    /// Interaction logic for SlideGenWindow.xaml — a modal wizard that inserts
    /// Zebulon slides directly into the active deck. Steps: pick type → pick a
    /// bind layout → pick lyrics (from the Provider) → pick insert position →
    /// generate. Mirrors the SetupWizard pattern (Visibility-toggled panels,
    /// modal, PowerPoint-owned). Provider I/O is async; Interop runs through
    /// <see cref="ISlideBuilder"/> (marshalled to the UI thread by the host).
    /// </summary>
    public partial class SlideGenWindow {
        private enum Step { Type, Layout, Data, Position }

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

        private Step _step = Step.Type;
        private LayoutMatch _selectedBind;
        private int _emptyLayoutIndex = -1;

        private List<string> _allPaths = new List<string>();
        private readonly ObservableCollection<string> _available = new ObservableCollection<string>();
        private readonly ObservableCollection<string> _selected = new ObservableCollection<string>();
        private bool _lyricsLoaded;
        private bool _busy;

        public SlideGenWindow() {
            InitializeComponent();
            _builder = Globals.ThisAddIn;
            AvailableListBox.ItemsSource = _available;
            SelectedListBox.ItemsSource = _selected;
            Loaded += delegate { Initialize(); };
        }

        private void Initialize() {
            if (!_builder.HasActivePresentation()) {
                PraiseCard.IsEnabled = false;
                PraiseStatus.Text = "열린 프레젠테이션이 없습니다.";
                ShowStep(Step.Type);
                return;
            }
            _layouts = _builder.ReadActiveLayouts();
            _praiseMatch = LayoutMatching.Match(_layouts, SlideGenDefaults.PraiseBoxCount);
            if (_praiseMatch.IsAvailable) {
                PraiseStatus.Text = "사용 가능 — bind 레이아웃 " + _praiseMatch.BindCandidates.Count + "개";
                PraiseCard.IsEnabled = true;
            } else if (_praiseMatch.BindCandidates.Count > 0 && !_praiseMatch.HasEmpty) {
                PraiseStatus.Text = "구분용 빈(empty) 레이아웃이 없어 사용 불가.";
                PraiseCard.IsEnabled = false;
            } else {
                PraiseStatus.Text = "이 deck에 적합한 Praise 레이아웃이 없습니다.";
                PraiseCard.IsEnabled = false;
            }
            ShowStep(Step.Type);
        }

        #region Navigation

        private void ShowStep(Step step) {
            _step = step;
            ClearErrors();
            TypePanel.Visibility = Vis(step == Step.Type);
            LayoutPanel.Visibility = Vis(step == Step.Layout);
            DataPanel.Visibility = Vis(step == Step.Data);
            PositionPanel.Visibility = Vis(step == Step.Position);

            BackButton.Visibility = Vis(step != Step.Type);
            NextButton.Visibility = Vis(step != Step.Type); // cards advance on the type step
            NextButton.Content = step == Step.Position ? "생성" : "다음 >";
        }

        private void PraiseCard_Click(object sender, RoutedEventArgs e) {
            EnterLayoutStep();
        }

        private void EnterLayoutStep() {
            LayoutListBox.Items.Clear();
            foreach (LayoutMatch m in _praiseMatch.BindCandidates) {
                LayoutListBox.Items.Add(new LayoutItem(m));
            }
            if (LayoutListBox.Items.Count > 0) {
                LayoutListBox.SelectedIndex = 0;
            }
            _emptyLayoutIndex = _praiseMatch.EmptyLayoutIndex;
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
                    if (_selected.Count == 0) {
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

        #region Data step

        private async Task EnterDataStepAsync() {
            ShowStep(Step.Data);
            if (_lyricsLoaded) {
                return;
            }
            SetBusy(true);
            DataStatus.Text = "목록 불러오는 중…";
            try {
                List<string> paths = await _provider.ListLyricsAsync();
                _allPaths = paths ?? new List<string>();
                _lyricsLoaded = true;
                ApplyFilter(FilterBox.Text);
                DataStatus.Text = _allPaths.Count + "개";
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
            foreach (string p in _allPaths) {
                if (_selected.Contains(p)) {
                    continue;
                }
                if (filter.Length == 0 || p.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0) {
                    _available.Add(p);
                }
            }
        }

        private void AvailableListBox_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e) {
            AddSelected();
        }

        private void AddButton_Click(object sender, RoutedEventArgs e) {
            AddSelected();
        }

        private void AddSelected() {
            List<string> picks = AvailableListBox.SelectedItems.Cast<string>().ToList();
            foreach (string p in picks) {
                if (!_selected.Contains(p)) {
                    _selected.Add(p);
                }
            }
            ApplyFilter(FilterBox.Text);
        }

        private void RemoveButton_Click(object sender, RoutedEventArgs e) {
            List<string> picks = SelectedListBox.SelectedItems.Cast<string>().ToList();
            foreach (string p in picks) {
                _selected.Remove(p);
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
            GenStatus.Text = "가사 불러오는 중…";
            try {
                List<Lyric> lyrics = new List<Lyric>();
                foreach (string path in _selected) {
                    Lyric lyric = await _provider.GetLyricAsync(path);
                    if (lyric != null) {
                        lyrics.Add(lyric);
                    }
                }
                if (lyrics.Count == 0) {
                    GenStatus.Text = "가사를 불러오지 못했습니다.";
                    return;
                }
                List<SlidePlanItem> plan = PraisePlanner.BuildPlan(lyrics);
                LayoutSelection sel = new LayoutSelection {
                    BindLayoutIndex = _selectedBind.LayoutIndex,
                    BoxSignatures = _selectedBind.BoxSignatures,
                    EmptyLayoutIndex = _emptyLayoutIndex
                };
                GenStatus.Text = "슬라이드 삽입 중…";
                int inserted = _builder.ExecutePlan(sel, plan, pos);
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

        #endregion

        #region Helpers

        private void SetBusy(bool busy) {
            _busy = busy;
            NextButton.IsEnabled = !busy;
            BackButton.IsEnabled = !busy;
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
        }

        #endregion
    }
}
