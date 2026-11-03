using System.Collections.Generic;

namespace ZebulonVSTO.Slides {
    /// <summary>
    /// Pure layout-matching logic. Two template generations are recognised:
    ///   - legacy PRAISE_/WORD_ templates, ported from the Exporter's template scan
    ///     (zebulon-exporter/py/pptx_exporter.py, PPTXFile.__init__): a bind layout
    ///     carries the markers "$1 ".."$N " (N = boxCount), the empty/separator
    ///     layout has no placeholders (<see cref="Match"/>);
    ///   - VENUE_ templates, by the /venue rules in <see cref="VenueRules"/>: roles
    ///     by layout-name suffix, named or numbered tokens, optional languages
    ///     (<see cref="MatchVenue"/>).
    /// Each box's placeholder is captured as a <see cref="BoxSignature"/>
    /// (geometry). The Exporter silently used one layout; here we return all bind
    /// candidates so the UI can let the user choose when several match.
    /// COM-free; unit-tested.
    /// </summary>
    public static class LayoutMatching {
        /// <summary>
        /// Both generations for one role: VENUE candidates first, then legacy ones
        /// not already listed; the VENUE blank layout when there is one, else the
        /// legacy empty layout.
        /// </summary>
        public static LayoutMatchResult MatchAll(IReadOnlyList<LayoutDescriptor> layouts, string role, int boxCount) {
            LayoutMatchResult venue = MatchVenue(VenueRules.Scan(layouts), role, boxCount);
            LayoutMatchResult legacy = Match(layouts, boxCount);
            foreach (LayoutMatch m in legacy.BindCandidates) {
                if (!venue.BindCandidates.Exists(v => v.MasterIndex == m.MasterIndex && v.LayoutIndex == m.LayoutIndex)) {
                    venue.BindCandidates.Add(m);
                }
            }
            if (!venue.HasEmpty) {
                venue.EmptyMasterIndex = legacy.EmptyMasterIndex;
                venue.EmptyLayoutIndex = legacy.EmptyLayoutIndex;
                venue.EmptyLayoutName = legacy.EmptyLayoutName;
            }
            return venue;
        }

        /// <summary>
        /// VENUE bind candidates for <paramref name="role"/> (praise/word): every
        /// valid layout of that role, box k = token k+1, a null signature for an
        /// optional language the layout leaves out. Empty = the resolved blank layout.
        /// </summary>
        public static LayoutMatchResult MatchVenue(VenueScan scan, string role, int boxCount) {
            LayoutMatchResult result = new LayoutMatchResult();
            if (scan == null) {
                return result;
            }
            foreach (VenueLayout layout in scan.Candidates(role)) {
                List<BoxSignature> sigs = new List<BoxSignature>();
                for (int k = 0; k < boxCount; k++) {
                    PlaceholderInfo ph;
                    sigs.Add(layout.Boxes.TryGetValue(k + 1, out ph) ? BoxSignature.Of(ph) : null);
                }
                result.BindCandidates.Add(new LayoutMatch {
                    MasterIndex = layout.MasterIndex,
                    LayoutIndex = layout.LayoutIndex,
                    Name = layout.Name,
                    BoxSignatures = sigs,
                    IsVenue = true
                });
            }
            VenueLayout blank = scan.Resolve(VenueRules.Blank);
            if (blank != null) {
                result.EmptyMasterIndex = blank.MasterIndex;
                result.EmptyLayoutIndex = blank.LayoutIndex;
                result.EmptyLayoutName = blank.Name;
            }
            return result;
        }

        /// <summary>Legacy "$1 ".."$N " marker matching (PRAISE_/WORD_ templates).</summary>
        public static LayoutMatchResult Match(IReadOnlyList<LayoutDescriptor> layouts, int boxCount) {
            LayoutMatchResult result = new LayoutMatchResult();
            if (layouts == null) {
                return result;
            }
            foreach (LayoutDescriptor layout in layouts) {
                IReadOnlyList<PlaceholderInfo> phs = layout.Placeholders;
                if (phs == null || phs.Count == 0) {
                    // last empty wins (faithful to Exporter)
                    result.EmptyMasterIndex = layout.MasterIndex;
                    result.EmptyLayoutIndex = layout.LayoutIndex;
                    result.EmptyLayoutName = layout.Name;
                    continue;
                }
                BoxSignature[] sig = new BoxSignature[boxCount];
                foreach (PlaceholderInfo ph in phs) {
                    string text = ph.Text ?? "";
                    for (int k = 0; k < boxCount; k++) {
                        if (text.StartsWith(SlideGenDefaults.BoxMarker(k))) {
                            // Last matching placeholder wins, mirroring the Exporter.
                            sig[k] = BoxSignature.Of(ph);
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
                        MasterIndex = layout.MasterIndex,
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
