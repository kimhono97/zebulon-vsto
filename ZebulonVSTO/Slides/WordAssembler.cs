using System;
using System.Collections.Generic;
using System.Text;

namespace ZebulonVSTO.Slides {
    /// <summary>
    /// Builds a <see cref="WordItem"/> from per-version chapter data + selected
    /// verse ranges, ported from jym-workbox word/helpers.ts buildWordItem.
    /// Each language slot's chapter verses are sliced by section (1-based inclusive)
    /// and resolved to plain text (ruby applied where present). headText =
    /// "{bookNames} {chapter}:{ranges}". COM-free; unit-tested.
    /// </summary>
    public static class WordAssembler {
        private const int SlotCount = 3;

        /// <param name="chapterChunk">Per-slot BibleData in version/slot order (≤3).</param>
        /// <param name="chapter">1-based chapter number.</param>
        /// <param name="sections">Coalesced verse ranges.</param>
        /// <param name="rubyModes">Per-slot ruby mode (Base when absent).</param>
        public static WordItem Build(IReadOnlyList<BibleData> chapterChunk, int chapter,
                                     IReadOnlyList<WordSection> sections, IReadOnlyList<RubyMode> rubyModes) {
            WordItem item = new WordItem {
                Sections = new List<WordSection>(),
                First = new List<List<string>>(),
                Second = new List<List<string>>(),
                Third = new List<List<string>>()
            };
            if (sections != null) {
                foreach (WordSection s in sections) {
                    item.Sections.Add(new WordSection(s.Start, s.End));
                }
            }
            int slots = chapterChunk != null ? chapterChunk.Count : 0;

            // headText: dedupe booknames across slots in order (from API response).
            List<string> bookNames = new List<string>();
            for (int i = 0; i < slots; i++) {
                BibleData d = chapterChunk[i];
                string bn = (d != null && d.BookName != null) ? d.BookName : "";
                if (bn.Length > 0 && !bookNames.Contains(bn)) {
                    bookNames.Add(bn);
                }
            }
            string bookNamesText;
            if (bookNames.Count == 0) {
                bookNamesText = "";
            } else if (bookNames.Count == 1) {
                bookNamesText = bookNames[0];
            } else {
                StringBuilder sb = new StringBuilder(bookNames[0]);
                sb.Append(" (");
                for (int i = 1; i < bookNames.Count; i++) {
                    if (i > 1) {
                        sb.Append(", ");
                    }
                    sb.Append(bookNames[i]);
                }
                sb.Append(")");
                bookNamesText = sb.ToString();
            }
            string ranges = SectionUtils.FormatSections(item.Sections);
            item.HeadText = (bookNamesText.Length > 0 ? bookNamesText + " " : "") + chapter + ":" + ranges;

            // Per-slot section slicing + ruby/HTML → plain text.
            List<List<string>>[] perSlot = new List<List<string>>[SlotCount];
            for (int slot = 0; slot < SlotCount; slot++) {
                perSlot[slot] = new List<List<string>>();
                if (slot >= slots) {
                    continue;
                }
                BibleData d = chapterChunk[slot];
                List<string> verses = (d != null && d.Data != null) ? d.Data : new List<string>();
                RubyMode mode = (rubyModes != null && slot < rubyModes.Count) ? rubyModes[slot] : RubyMode.Base;

                foreach (WordSection seg in item.Sections) {
                    bool segHasRuby = false;
                    List<string> raw = new List<string>();
                    for (int v = seg.Start - 1; v < seg.End; v++) {
                        string line = (v >= 0 && v < verses.Count) ? (verses[v] ?? "") : "";
                        if (line.IndexOf("<ruby>", StringComparison.OrdinalIgnoreCase) >= 0) {
                            segHasRuby = true;
                        }
                        raw.Add(line);
                    }
                    List<string> outLines = new List<string>();
                    foreach (string line in raw) {
                        // Ruby section → apply the slot's ruby mode; otherwise strip
                        // markup to plain text. NOTE: the web passes non-ruby verses
                        // verbatim, but a PowerPoint placeholder renders text literally
                        // (no HTML), so we plain-text every line to keep stray tags /
                        // entities off the slide — an intentional, output-improving
                        // divergence from the web/Exporter.
                        outLines.Add(segHasRuby ? RubyText.ApplyRubyMode(line, mode) : RubyText.ToPlainText(line));
                    }
                    perSlot[slot].Add(outLines);
                }
            }
            item.First = perSlot[0];
            item.Second = perSlot[1];
            item.Third = perSlot[2];
            return item;
        }
    }
}
