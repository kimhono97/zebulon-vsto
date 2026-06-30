using System.Collections.Generic;

namespace ZebulonVSTO.Slides {
    /// <summary>
    /// Builds the slide plan for Word (scripture) items, ported 1:1 from the
    /// Exporter (zebulon-exporter/py/pptx_exporter.py, PPTXFile_Word.addItem).
    /// Per valid item: one empty separator (no note), then one bind slide per
    /// emitted verse line across all sections — box 0 = headText, boxes 1/2/3 =
    /// "{verse}. " + line (a box is omitted when its line is empty). No text
    /// transforms, no title slide, no trailing separator, no notes (unlike Praise).
    /// box→placeholder mapping is resolved by the Interop layer. COM-free; unit-tested.
    /// </summary>
    public static class WordPlanner {
        private const int LangCount = 3; // boxes 1,2,3 (box 0 = headText)

        public static List<SlidePlanItem> BuildPlan(IReadOnlyList<WordItem> items) {
            List<SlidePlanItem> plan = new List<SlidePlanItem>();
            if (items == null) {
                return plan;
            }
            foreach (WordItem item in items) {
                if (item == null || string.IsNullOrEmpty(item.HeadText) ||
                    item.Sections == null || item.Sections.Count == 0) {
                    continue; // matches the Exporter returning False for invalid items
                }
                List<List<string>> first = item.First ?? new List<List<string>>();
                List<List<string>> second = item.Second ?? new List<List<string>>();
                List<List<string>> third = item.Third ?? new List<List<string>>();
                int secN = item.Sections.Count;
                // Every language has fewer section-chunks than there are sections → no usable data.
                if (first.Count < secN && second.Count < secN && third.Count < secN) {
                    continue;
                }
                List<List<string>>[] langData = new List<List<string>>[] { first, second, third };
                string headText = item.HeadText;

                // One empty separator at the item start (no note).
                plan.Add(new SlidePlanItem { Kind = LayoutKind.Empty, Note = null });

                for (int secIndex = 0; secIndex < item.Sections.Count; secIndex++) {
                    WordSection sec = item.Sections[secIndex];
                    int start = sec.Start, end = sec.End;
                    if (start > end) {
                        int t = start; start = end; end = t;
                    }

                    List<string>[] secChunks = new List<string>[LangCount];
                    for (int li = 0; li < LangCount; li++) {
                        List<List<string>> ld = langData[li];
                        secChunks[li] = (secIndex < ld.Count && ld[secIndex] != null)
                            ? ld[secIndex]
                            : new List<string>();
                    }

                    int lineCount = end - start + 1; // inclusive
                    for (int ti = 0; ti < lineCount; ti++) {
                        string[] lineTexts = new string[LangCount];
                        int emptyCount = 0;
                        for (int li = 0; li < LangCount; li++) {
                            List<string> chunk = secChunks[li];
                            string s = (ti < chunk.Count) ? (chunk[ti] ?? "") : "";
                            lineTexts[li] = s;
                            if (s.Length == 0) {
                                emptyCount++;
                            }
                        }
                        // All languages empty for this line → stop this section.
                        if (emptyCount == LangCount) {
                            break;
                        }

                        string prefix = (start + ti) + ". "; // verse-number prefix
                        Dictionary<int, string> box = new Dictionary<int, string> { { 0, headText } };
                        for (int li = 0; li < LangCount; li++) {
                            if (lineTexts[li].Length > 0) {
                                box[li + 1] = prefix + lineTexts[li];
                            }
                        }
                        plan.Add(new SlidePlanItem { Kind = LayoutKind.Bind, BoxText = box, Note = null });
                    }
                }
            }
            return plan;
        }
    }
}
