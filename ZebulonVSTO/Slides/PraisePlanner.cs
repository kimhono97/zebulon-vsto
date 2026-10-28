using System;
using System.Collections.Generic;

namespace ZebulonVSTO.Slides {
    /// <summary>
    /// Builds the ordered slide plan for Praise items, ported 1:1 from the
    /// Exporter (zebulon-exporter/py/pptx_exporter.py, PPTXFile_Praise.addItem,
    /// and its VENUE successor venue/items/praise.py). Per valid item: an empty
    /// separator, a title slide (box0 = raw title, box1/2 = " "), one bind slide
    /// per page (KR/EN/CN via TextTransforms), then an empty separator. The VENUE
    /// display switches drop the title slide (<c>titleSlide: false</c>) or both
    /// separators (<c>separators: false</c>); the "&lt;n&gt;. &lt;title&gt;" note
    /// then moves to the song's first slide. The box→placeholder mapping and the
    /// CN-box geometry tweak are resolved by the Interop layer, not here.
    /// COM-free; unit-tested.
    /// </summary>
    public static class PraisePlanner {
        public static List<SlidePlanItem> BuildPlan(IReadOnlyList<Lyric> items) {
            return BuildPlan(items, true, true);
        }

        public static List<SlidePlanItem> BuildPlan(IReadOnlyList<Lyric> items, bool titleSlide, bool separators) {
            List<SlidePlanItem> plan = new List<SlidePlanItem>();
            if (items == null) {
                return plan;
            }
            int itemNo = 0;
            foreach (Lyric item in items) {
                if (IsValid(item)) {
                    AddSong(plan, item, ++itemNo, titleSlide, separators);
                }
            }
            return plan;
        }

        /// <summary>False for items the Exporter rejects (no title or no groups); they are skipped and not numbered.</summary>
        public static bool IsValid(Lyric item) {
            return item != null && !string.IsNullOrEmpty(item.Title) && item.Groups != null && item.Groups.Count > 0;
        }

        /// <summary>Append one valid song as song number <paramref name="number"/> (venue/items/praise.py build).</summary>
        public static void AddSong(List<SlidePlanItem> plan, Lyric item, int number, bool titleSlide, bool separators) {
            const int boxCount = SlideGenDefaults.PraiseBoxCount;
            string title = item.Title;
            // "<n>. <title>" marks where a song starts in the presenter notes; it goes on the song's first slide.
            string pending = number + ". " + title;
            Func<string, string> notes = text => {
                string first = pending;
                pending = null;
                if (first == null) {
                    return text;
                }
                return text == null ? first : first + "\n" + text;
            };

            if (separators) {
                plan.Add(new SlidePlanItem { Kind = LayoutKind.Empty, Note = notes(null) });
            }
            if (titleSlide) {
                plan.Add(new SlidePlanItem {
                    Kind = LayoutKind.Bind,
                    BoxText = new Dictionary<int, string> { { 0, title }, { 1, " " }, { 2, " " } },
                    Note = notes(title + " - Title")
                });
            }
            foreach (LyricGroup g in item.Groups) {
                if (g == null || string.IsNullOrEmpty(g.Name) ||
                    g.Pages == null || g.Pages.Count == 0) {
                    continue;
                }
                int n = g.Pages.Count;
                for (int i = 0; i < n; i++) {
                    List<string> page = g.Pages[i];
                    if (page == null) {
                        continue;
                    }
                    int lim = Math.Min(page.Count, boxCount);
                    Dictionary<int, string> box = new Dictionary<int, string>();
                    for (int j = 0; j < lim; j++) {
                        box[j] = TextTransforms.TransformPraiseBox(j, page[j]);
                    }
                    plan.Add(new SlidePlanItem {
                        Kind = LayoutKind.Bind,
                        BoxText = box,
                        Note = notes(title + " - " + g.Name + " (" + (i + 1) + "/" + n + ")")
                    });
                }
            }
            if (separators) {
                plan.Add(new SlidePlanItem { Kind = LayoutKind.Empty, Note = title + " - End" });
            }
        }
    }
}
