using System.Collections.Generic;

namespace ZebulonVSTO.Slides {
    /// <summary>How a planned slide maps to a template layout.</summary>
    public enum LayoutKind { Empty, Bind }

    /// <summary>Where generated slides are inserted into the active deck.</summary>
    public enum InsertPosition { Front, AfterCurrent, End }

    /// <summary>
    /// COM-free snapshot of one placeholder on a layout: its text plus geometry
    /// (points). Produced by the Interop layer reading PowerPoint; consumed by
    /// the pure <see cref="LayoutMatching"/>.
    /// </summary>
    public class PlaceholderInfo {
        public string Text { get; set; }
        public float Height { get; set; }
        public float Top { get; set; }
    }

    /// <summary>COM-free snapshot of one slide layout (default master).</summary>
    public class LayoutDescriptor {
        // 1-based index into the default master's CustomLayouts collection.
        public int LayoutIndex { get; set; }
        public string Name { get; set; }
        public IReadOnlyList<PlaceholderInfo> Placeholders { get; set; }
    }

    /// <summary>
    /// Identifies the placeholder that backs one box by its geometry
    /// (height + top), mirroring the Exporter's placeholder matching.
    /// </summary>
    public class BoxSignature {
        public float Height { get; set; }
        public float Top { get; set; }
    }

    /// <summary>A layout satisfying the bind condition (all box markers present).</summary>
    public class LayoutMatch {
        public int LayoutIndex { get; set; }
        public string Name { get; set; }
        public List<BoxSignature> BoxSignatures { get; set; }
    }

    /// <summary>Result of scanning a deck's layouts for one template type.</summary>
    public class LayoutMatchResult {
        public List<LayoutMatch> BindCandidates { get; set; }
        public int EmptyLayoutIndex { get; set; }

        public LayoutMatchResult() {
            BindCandidates = new List<LayoutMatch>();
            EmptyLayoutIndex = -1;
        }

        public bool HasEmpty { get { return EmptyLayoutIndex >= 1; } }
        public bool IsAvailable { get { return HasEmpty && BindCandidates.Count > 0; } }
    }

    /// <summary>The layout choice fed to the Interop slide builder.</summary>
    public class LayoutSelection {
        public int BindLayoutIndex { get; set; }
        public List<BoxSignature> BoxSignatures { get; set; }
        public int EmptyLayoutIndex { get; set; }
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
