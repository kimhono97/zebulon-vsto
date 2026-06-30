using System.Collections.Generic;
using System.Text;

namespace ZebulonVSTO.Slides {
    /// <summary>
    /// Verse-range helpers ported from jym-workbox state/projectIO.ts
    /// (coalesceSections / formatSections). COM-free; unit-tested.
    /// </summary>
    public static class SectionUtils {
        /// <summary>
        /// Sort and merge overlapping/adjacent ranges (merge when
        /// next.Start &lt;= last.End + 1). Null and inverted ranges (End &lt; Start) are
        /// dropped — faithful to the web coalesceSections, which filters with
        /// <c>r =&gt; r &amp;&amp; r.end &gt;= r.start</c> (it does not swap).
        /// </summary>
        public static List<WordSection> Coalesce(IEnumerable<WordSection> ranges) {
            List<WordSection> normalized = new List<WordSection>();
            if (ranges != null) {
                foreach (WordSection r in ranges) {
                    if (r == null || r.End < r.Start) {
                        continue; // matches the web filter(r => r && r.end >= r.start)
                    }
                    normalized.Add(new WordSection(r.Start, r.End));
                }
            }
            normalized.Sort((a, b) => a.Start != b.Start ? a.Start.CompareTo(b.Start) : a.End.CompareTo(b.End));

            List<WordSection> merged = new List<WordSection>();
            foreach (WordSection r in normalized) {
                if (merged.Count > 0) {
                    WordSection last = merged[merged.Count - 1];
                    if (r.Start <= last.End + 1) {
                        if (r.End > last.End) {
                            last.End = r.End;
                        }
                        continue;
                    }
                }
                merged.Add(new WordSection(r.Start, r.End));
            }
            return merged;
        }

        /// <summary>Render ranges as "16, 18-20" (single verse → "N", range → "S-E").</summary>
        public static string FormatSections(IReadOnlyList<WordSection> sections) {
            if (sections == null || sections.Count == 0) {
                return "";
            }
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < sections.Count; i++) {
                if (i > 0) {
                    sb.Append(", ");
                }
                WordSection s = sections[i];
                sb.Append(s.Start);
                if (s.End != s.Start) {
                    sb.Append("-").Append(s.End);
                }
            }
            return sb.ToString();
        }
    }
}
