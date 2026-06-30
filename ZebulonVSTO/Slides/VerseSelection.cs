using System.Collections.Generic;

namespace ZebulonVSTO.Slides {
    /// <summary>
    /// Pure, COM-free verse-range selection state machine, ported 1:1 from the web
    /// SlideEditorWord tap-to-select flow (jym-workbox
    /// src/zebulon/components/SlideEditorWord.tsx: onVerseClick / commitPending /
    /// displayRanges / verseRole). Drives the <see cref="WordSelectWindow"/> verse
    /// picker:
    ///   - <see cref="Segments"/>: committed disjoint ranges (the chips).
    ///   - <see cref="Pending"/>: the range currently being picked (tap a start
    ///     verse, then a later verse to extend; any other tap restarts).
    ///   - <see cref="DisplayRanges"/>: committed + pending, coalesced — what
    ///     Confirm produces and what drives per-verse highlighting.
    /// Multi-range is preserved via <see cref="SectionUtils.Coalesce"/>. Unit-tested.
    /// </summary>
    public sealed class VerseSelection {
        private readonly List<WordSection> _segments = new List<WordSection>();
        private WordSection _pending; // null = nothing being picked

        /// <summary>Committed, coalesced ranges (in order) — one chip each.</summary>
        public IReadOnlyList<WordSection> Segments {
            get { return _segments; }
        }

        /// <summary>The in-progress range, or null.</summary>
        public WordSection Pending {
            get { return _pending; }
        }

        public bool HasPending {
            get { return _pending != null; }
        }

        /// <summary>
        /// Tap a verse. Mirrors the web onVerseClick: with no pending, start a
        /// single-verse pending; with a single pending and a later verse, extend it
        /// into a range; otherwise restart at the tapped verse. Backward taps (an
        /// earlier verse) restart rather than extend.
        /// </summary>
        public void Click(int verseNo) {
            if (verseNo < 1) {
                return;
            }
            if (_pending == null) {
                _pending = new WordSection(verseNo, verseNo);
                return;
            }
            if (_pending.Start == _pending.End && verseNo > _pending.Start) {
                _pending = new WordSection(_pending.Start, verseNo);
                return;
            }
            _pending = new WordSection(verseNo, verseNo);
        }

        /// <summary>Commit the pending range into the coalesced committed set.</summary>
        public void CommitPending() {
            if (_pending == null) {
                return;
            }
            List<WordSection> next = new List<WordSection>(_segments) { _pending };
            List<WordSection> merged = SectionUtils.Coalesce(next);
            _segments.Clear();
            _segments.AddRange(merged);
            _pending = null;
        }

        /// <summary>Remove the i-th committed segment (chip).</summary>
        public void RemoveSegment(int index) {
            if (index >= 0 && index < _segments.Count) {
                _segments.RemoveAt(index);
            }
        }

        /// <summary>Clear everything (called on a book/chapter change).</summary>
        public void Reset() {
            _segments.Clear();
            _pending = null;
        }

        /// <summary>
        /// Committed + pending, coalesced and sorted — the full selection that
        /// Confirm produces and that <see cref="RoleOf"/> reads.
        /// </summary>
        public List<WordSection> DisplayRanges() {
            List<WordSection> all = new List<WordSection>(_segments);
            if (_pending != null) {
                all.Add(_pending);
            }
            return SectionUtils.Coalesce(all);
        }

        /// <summary>
        /// Per-verse role for range-style rendering: "" (unselected), "single",
        /// "start", "end", or "middle". Mirrors the web verseRole.
        /// </summary>
        public string RoleOf(int verseNo) {
            foreach (WordSection r in DisplayRanges()) {
                if (verseNo < r.Start || verseNo > r.End) {
                    continue;
                }
                if (r.Start == r.End) {
                    return "single";
                }
                if (verseNo == r.Start) {
                    return "start";
                }
                if (verseNo == r.End) {
                    return "end";
                }
                return "middle";
            }
            return "";
        }
    }
}
