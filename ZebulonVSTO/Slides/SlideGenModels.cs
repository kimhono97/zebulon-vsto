using System.Collections.Generic;

namespace ZebulonVSTO.Slides {
    /// <summary>How a planned slide maps to a template layout.</summary>
    public enum LayoutKind { Empty, Bind }

    /// <summary>Where generated slides are inserted into the active deck.</summary>
    public enum InsertPosition { Front, AfterCurrent, End }

    /// <summary>
    /// COM-free snapshot of one placeholder on a layout: its text plus geometry
    /// (points). Produced by the Interop layer reading PowerPoint; consumed by
    /// the pure <see cref="LayoutMatching"/> and <see cref="VenueRules"/>.
    /// </summary>
    public class PlaceholderInfo {
        // Prompt text as PowerPoint reports it: paragraphs separated by '\r',
        // line breaks inside a paragraph as '\v'.
        public string Text { get; set; }
        public float Height { get; set; }
        public float Top { get; set; }
        public float Left { get; set; }
        public float Width { get; set; }

        /// <summary>A date, footer or slide-number placeholder: never a data box,
        /// and not counted as content (an empty layout may still carry one).</summary>
        public bool IsLatent { get; set; }

        /// <summary>False for placeholders that cannot hold text (e.g. picture).</summary>
        public bool HasTextFrame { get; set; } = true;
    }

    /// <summary>COM-free snapshot of one slide layout.</summary>
    public class LayoutDescriptor {
        // 1-based index into Presentation.Designs (one slide master per design).
        public int MasterIndex { get; set; }
        // 1-based index into that master's CustomLayouts collection.
        public int LayoutIndex { get; set; }
        public string Name { get; set; }
        public IReadOnlyList<PlaceholderInfo> Placeholders { get; set; }
    }

    /// <summary>
    /// Identifies the placeholder that backs one box by its geometry. The
    /// Exporter matches height + top only; left + width are added because
    /// some templates put two language boxes side by side at the same height
    /// (VENUE_Cheonbosan LyricPage $lang2 / $lang3), which height + top alone
    /// would resolve to the same placeholder.
    /// </summary>
    public class BoxSignature {
        public float Height { get; set; }
        public float Top { get; set; }
        public float Left { get; set; }
        public float Width { get; set; }

        public static BoxSignature Of(PlaceholderInfo ph) {
            return new BoxSignature { Height = ph.Height, Top = ph.Top, Left = ph.Left, Width = ph.Width };
        }
    }

    /// <summary>
    /// A layout satisfying the bind condition: all "$N " markers present (legacy
    /// PRAISE_/WORD_ templates), or a valid VENUE praise/word layout. One
    /// signature per box; null for an optional box the layout leaves out (VENUE).
    /// </summary>
    public class LayoutMatch {
        public int MasterIndex { get; set; }
        public int LayoutIndex { get; set; }
        public string Name { get; set; }
        public List<BoxSignature> BoxSignatures { get; set; }

        /// <summary>Found by the VENUE rules (<see cref="VenueRules"/>) rather
        /// than the legacy "$N " markers.</summary>
        public bool IsVenue { get; set; }
    }

    /// <summary>Result of scanning a deck's layouts for one template type.</summary>
    public class LayoutMatchResult {
        public List<LayoutMatch> BindCandidates { get; set; }
        public int EmptyMasterIndex { get; set; }
        public int EmptyLayoutIndex { get; set; }
        public string EmptyLayoutName { get; set; }

        public LayoutMatchResult() {
            BindCandidates = new List<LayoutMatch>();
            EmptyLayoutIndex = -1;
        }

        public bool HasEmpty { get { return EmptyLayoutIndex >= 1; } }
        public bool IsAvailable { get { return HasEmpty && BindCandidates.Count > 0; } }
    }

    /// <summary>The layout choice fed to the Interop slide builder.</summary>
    public class LayoutSelection {
        public int BindMasterIndex { get; set; }
        public int BindLayoutIndex { get; set; }
        public List<BoxSignature> BoxSignatures { get; set; }
        public int EmptyMasterIndex { get; set; }
        public int EmptyLayoutIndex { get; set; }

        /// <summary>
        /// Write a single space into every box a bind slide leaves unset, so the
        /// layout's "$lang2 …" prompt does not show in edit view. VENUE behaviour
        /// (exporter venue/deck.py EMPTY_TEXT); the legacy Exporter leaves them
        /// empty, so this is set only for VENUE layouts.
        /// </summary>
        public bool FillUnusedBoxes { get; set; }

        /// <summary>
        /// Apply the Praise-only CN-box vertical-centering nudge in the Interop
        /// layer (<c>ThisAddIn.ApplyBoxText</c>). The Exporter performs this tweak
        /// solely in <c>PPTXFile_Praise.addItem</c> (pptx_exporter.py); the Word
        /// path (<c>PPTXFile_Word.addItem</c>) does NOT — box 2 there is the 2nd
        /// language placeholder and must stay put. The wizard sets this true only
        /// for the Praise template.
        /// </summary>
        public bool CenterCnBox { get; set; }
    }

    /// <summary>
    /// One planned slide: which layout kind, the per-box text to set (Bind only),
    /// and the speaker-notes line. Produced by the pure planners, executed by the
    /// Interop slide builder.
    /// </summary>
    public class SlidePlanItem {
        public LayoutKind Kind { get; set; }
        public Dictionary<int, string> BoxText { get; set; }
        public string Note { get; set; }
    }
}
