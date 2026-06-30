using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace ZebulonVSTO.Slides {
    /// <summary>
    /// Interaction logic for LyricPreviewWindow.xaml — a read-only modal that
    /// renders a fetched <see cref="Lyric"/> (title → groups → pages, each page's
    /// KR/EN/CN lines) so the user can verify the right song/version before
    /// inserting. Pure WPF; no Interop. The caller fetches the Lyric (cached).
    /// </summary>
    public partial class LyricPreviewWindow {
        public LyricPreviewWindow(Lyric lyric) {
            InitializeComponent();
            Render(lyric);
        }

        private void Render(Lyric lyric) {
            if (lyric == null) {
                ContentPanel.Children.Add(new TextBlock { Text = "내용을 불러올 수 없습니다." });
                return;
            }

            string title = lyric.Title ?? "";
            if (!string.IsNullOrEmpty(title)) {
                Title = title;
            }
            ContentPanel.Children.Add(new TextBlock {
                Text = title,
                FontSize = 18,
                FontWeight = FontWeights.Bold,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 12)
            });

            if (lyric.Groups == null) {
                return;
            }
            foreach (LyricGroup g in lyric.Groups) {
                if (g == null) {
                    continue;
                }
                ContentPanel.Children.Add(new TextBlock {
                    Text = g.Name ?? "",
                    FontSize = 14,
                    FontWeight = FontWeights.Bold,
                    Foreground = Brushes.SteelBlue,
                    TextWrapping = TextWrapping.Wrap,
                    Margin = new Thickness(0, 10, 0, 4)
                });
                if (g.Pages == null) {
                    continue;
                }
                int pageNo = 0;
                foreach (List<string> page in g.Pages) {
                    pageNo++;
                    if (page == null) {
                        continue;
                    }
                    StackPanel sp = new StackPanel();
                    sp.Children.Add(new TextBlock {
                        Text = "p." + pageNo,
                        FontSize = 10,
                        Foreground = Brushes.Gray,
                        Margin = new Thickness(0, 0, 0, 2)
                    });
                    foreach (string line in page) {
                        if (!string.IsNullOrEmpty(line)) {
                            sp.Children.Add(new TextBlock {
                                Text = line,
                                TextWrapping = TextWrapping.Wrap,
                                Margin = new Thickness(0, 1, 0, 1)
                            });
                        }
                    }
                    ContentPanel.Children.Add(new Border {
                        BorderBrush = Brushes.LightGray,
                        BorderThickness = new Thickness(1),
                        Padding = new Thickness(8),
                        Margin = new Thickness(0, 0, 0, 6),
                        Child = sp
                    });
                }
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e) {
            Close();
        }
    }
}
