using System;
using System.Collections.Generic;

namespace ZebulonVSTO.Slides {
    /// <summary>
    /// Builds the ordered slide plan for Praise items, ported 1:1 from the
    /// Exporter (zebulon-exporter/py/pptx_exporter.py, PPTXFile_Praise.addItem).
    /// Per valid item: an empty separator, a title slide (box0 = raw title,
    /// box1/2 = " "), one bind slide per page (KR/EN/CN via TextTransforms), then
    /// an empty separator. The box→placeholder mapping and the CN-box geometry
    /// tweak are resolved by the Interop layer, not here. COM-free; unit-tested.
    /// </summary>
    public static class PraisePlanner {
        public static List<SlidePlanItem> BuildPlan(IReadOnlyList<Lyric> items) {
            const int boxCount = SlideGenDefaults.PraiseBoxCount;
            List<SlidePlanItem> plan = new List<SlidePlanItem>();
            if (items == null) {
                return plan;
            }
            int itemNo = 0;
            foreach (Lyric item in items) {
                if (item == null || string.IsNullOrEmpty(item.Title) ||
                    item.Groups == null || item.Groups.Count == 0) {
                    continue; // matches the Exporter returning False for invalid items
                }
                string title = item.Title;
                itemNo++;
                plan.Add(new SlidePlanItem {
                    Kind = LayoutKind.Empty,
                    Note = itemNo + ". " + title
                });
                plan.Add(new SlidePlanItem {
                    Kind = LayoutKind.Bind,
                    BoxText = new Dictionary<int, string> { { 0, title }, { 1, " " }, { 2, " " } },
                    Note = title + " - Title"
                });
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
                            Note = title + " - " + g.Name + " (" + (i + 1) + "/" + n + ")"
                        });
                    }
                }
                plan.Add(new SlidePlanItem {
                    Kind = LayoutKind.Empty,
                    Note = title + " - End"
                });
            }
            return plan;
        }
    }
}
