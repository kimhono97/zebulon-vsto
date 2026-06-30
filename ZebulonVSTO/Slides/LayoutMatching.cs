using System.Collections.Generic;

namespace ZebulonVSTO.Slides {
    /// <summary>
    /// Pure layout-matching logic, ported from the Exporter's template scan
    /// (zebulon-exporter/py/pptx_exporter.py, PPTXFile.__init__). Given COM-free
    /// descriptors of a deck's layouts, finds:
    ///   - bind-layout candidates: a layout whose placeholders carry the markers
    ///     "$1 ".."$N " (N = boxCount); each box's placeholder is captured as a
    ///     <see cref="BoxSignature"/> (placeholder Idx + geometry).
    ///   - the empty/separator layout: a layout with no placeholders.
    /// The Exporter silently used the *last* matching layout; here we return all
    /// bind candidates so the UI can let the user choose when several match.
    /// COM-free; unit-tested.
    /// </summary>
    public static class LayoutMatching {
        public static LayoutMatchResult Match(IReadOnlyList<LayoutDescriptor> layouts, int boxCount) {
            LayoutMatchResult result = new LayoutMatchResult();
            if (layouts == null) {
                return result;
            }
            foreach (LayoutDescriptor layout in layouts) {
                IReadOnlyList<PlaceholderInfo> phs = layout.Placeholders;
                if (phs == null || phs.Count == 0) {
                    result.EmptyLayoutIndex = layout.LayoutIndex; // last empty wins (faithful to Exporter)
                    continue;
                }
                BoxSignature[] sig = new BoxSignature[boxCount];
                foreach (PlaceholderInfo ph in phs) {
                    string text = ph.Text ?? "";
                    for (int k = 0; k < boxCount; k++) {
                        if (text.StartsWith(SlideGenDefaults.BoxMarker(k))) {
                            // Last matching placeholder wins, mirroring the Exporter.
                            sig[k] = new BoxSignature { Height = ph.Height, Top = ph.Top };
                        }
                    }
                }
                bool allFound = true;
                for (int k = 0; k < boxCount; k++) {
                    if (sig[k] == null) {
                        allFound = false;
                        break;
                    }
                }
                if (allFound) {
                    result.BindCandidates.Add(new LayoutMatch {
                        LayoutIndex = layout.LayoutIndex,
                        Name = layout.Name,
                        BoxSignatures = new List<BoxSignature>(sig)
                    });
                }
            }
            return result;
        }
    }
}
