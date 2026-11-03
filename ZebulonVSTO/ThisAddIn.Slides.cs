using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using ZebulonVSTO.Slides;
using Core = Microsoft.Office.Core;

namespace ZebulonVSTO {
    /// <summary>
    /// <see cref="Slides.ISlideBuilder"/> implementation — the Interop half of the
    /// slide-generation feature. Reads the active deck's layouts and inserts
    /// planned slides, always marshalled onto the add-in UI thread via
    /// <c>_dispatcher</c> (same pattern as ISlideController). COM behaviour is
    /// verified at runtime via F5, not unit tests.
    /// </summary>
    public partial class ThisAddIn : Slides.ISlideBuilder {

        /// <summary>Open the slide-generation wizard, modal and owned by the
        /// PowerPoint main window (mirrors <see cref="ShowSetupWizard"/>).</summary>
        public void ShowSlideGenWindow() {
            Slides.SlideGenWindow window = new Slides.SlideGenWindow();
            new System.Windows.Interop.WindowInteropHelper(window).Owner = (IntPtr)Application.HWND;
            window.ShowDialog();
        }

        public bool HasActivePresentation() {
            if (_dispatcher == null) {
                return false;
            }
            bool has = false;
            try {
                _dispatcher.Invoke(() => { has = SafeActivePresentation() != null; });
            } catch {
                // ignore — treat as no presentation
            }
            return has;
        }

        public List<LayoutDescriptor> ReadLayouts(string presentationFullName) {
            List<LayoutDescriptor> result = new List<LayoutDescriptor>();
            if (_dispatcher == null) {
                return result;
            }
            try {
                _dispatcher.Invoke(() => {
                    Presentation pres = ResolvePresentation(presentationFullName);
                    if (pres == null) {
                        return;
                    }
                    // Every slide master (one per design), like the Exporter's
                    // /venue scan — not just the default pres.SlideMaster.
                    Designs designs = pres.Designs;
                    for (int d = 1; d <= designs.Count; d++) {
                        CustomLayouts layouts = designs[d].SlideMaster.CustomLayouts;
                        for (int i = 1; i <= layouts.Count; i++) {
                            CustomLayout layout = layouts[i];
                            List<PlaceholderInfo> phInfos = new List<PlaceholderInfo>();
                            Placeholders phs = layout.Shapes.Placeholders;
                            for (int p = 1; p <= phs.Count; p++) {
                                Shape sh = phs[p];
                                phInfos.Add(new PlaceholderInfo {
                                    Text = SafeText(sh),
                                    Height = sh.Height,
                                    Top = sh.Top,
                                    Left = sh.Left,
                                    Width = sh.Width,
                                    IsLatent = IsLatentPlaceholder(sh),
                                    HasTextFrame = sh.HasTextFrame == Core.MsoTriState.msoTrue
                                });
                            }
                            result.Add(new LayoutDescriptor {
                                MasterIndex = d,
                                LayoutIndex = i,
                                Name = SafeLayoutName(layout),
                                Placeholders = phInfos
                            });
                        }
                    }
                });
            } catch (Exception e) {
                LogError("Failed to read active layouts.", e);
            }
            return result;
        }

        public int ExecutePlan(string presentationFullName, LayoutSelection selection, List<SlidePlanItem> plan, InsertPosition position) {
            if (_dispatcher == null || selection == null || plan == null) {
                return 0;
            }
            int inserted = 0;
            try {
                _dispatcher.Invoke(() => {
                    // Target the deck the wizard scanned (by FullName), NOT whatever
                    // happens to be active — a download-opened deck may not be the
                    // active presentation under the modal dialog.
                    Presentation pres = ResolvePresentation(presentationFullName);
                    if (pres == null) {
                        return;
                    }
                    CustomLayout bindLayout = LayoutAt(pres, selection.BindMasterIndex, selection.BindLayoutIndex);
                    CustomLayout emptyLayout = LayoutAt(pres, selection.EmptyMasterIndex, selection.EmptyLayoutIndex);
                    if (bindLayout == null || emptyLayout == null) {
                        return;
                    }
                    int insertAt = ResolveStartIndex(pres, position);
                    foreach (SlidePlanItem item in plan) {
                        CustomLayout layout = item.Kind == LayoutKind.Bind ? bindLayout : emptyLayout;
                        Slide slide = pres.Slides.AddSlide(insertAt, layout);
                        if (item.Kind == LayoutKind.Bind && item.BoxText != null) {
                            ApplyBoxText(slide, selection, item.BoxText);
                        }
                        SetNote(slide, item.Note);
                        insertAt++;
                        inserted++;
                    }
                    // Bring the edited deck to the front so the user sees the result
                    // (it may not be the active window when opened via download).
                    if (inserted > 0) {
                        try {
                            if (pres.Windows.Count > 0) {
                                pres.Windows[1].Activate();
                            }
                        } catch {
                            // best-effort
                        }
                    }
                });
            } catch (Exception e) {
                LogError("Failed to insert generated slides.", e);
            }
            return inserted;
        }

        public string OpenPresentation(string path) {
            if (_dispatcher == null || string.IsNullOrEmpty(path)) {
                return "";
            }
            string fullName = "";
            try {
                _dispatcher.Invoke(() => {
                    // WithWindow=msoTrue so the deck gets a window; ReadOnly/Untitled=
                    // msoFalse keep it bound to the file the user just saved. Return
                    // its FullName so the wizard can target this exact deck regardless
                    // of which presentation is "active" under the modal dialog.
                    Presentation pres = Application.Presentations.Open(
                        path, Core.MsoTriState.msoFalse, Core.MsoTriState.msoFalse, Core.MsoTriState.msoTrue);
                    if (pres != null) {
                        try {
                            fullName = pres.FullName;
                        } catch {
                            fullName = path;
                        }
                    }
                });
            } catch (Exception e) {
                LogError("Failed to open downloaded presentation.", e);
            }
            return fullName;
        }

        public string GetActivePresentationFolder() {
            if (_dispatcher == null) {
                return "";
            }
            string folder = "";
            try {
                _dispatcher.Invoke(() => {
                    Presentation pres = SafeActivePresentation();
                    if (pres != null) {
                        try {
                            folder = pres.Path ?? ""; // empty for a never-saved deck
                        } catch {
                            folder = "";
                        }
                    }
                });
            } catch {
                // ignore — fall back to empty
            }
            return folder;
        }

        #region Interop helpers (UI thread only)

        private Presentation SafeActivePresentation() {
            try {
                if (Application.Presentations.Count == 0) {
                    return null;
                }
                return Application.ActivePresentation;
            } catch {
                return null;
            }
        }

        // Resolve the wizard's target deck: a specific open presentation by
        // FullName, or the active presentation when no target is given (manual
        // path). Falls back to the active deck if the named one isn't found.
        private Presentation ResolvePresentation(string fullName) {
            try {
                if (!string.IsNullOrEmpty(fullName)) {
                    Presentations all = Application.Presentations;
                    for (int i = 1; i <= all.Count; i++) {
                        Presentation p = all[i];
                        string fn = "";
                        try {
                            fn = p.FullName;
                        } catch {
                            // some presentations may not expose FullName; skip
                        }
                        if (string.Equals(fn, fullName, StringComparison.OrdinalIgnoreCase)) {
                            return p;
                        }
                    }
                }
            } catch {
                // fall through to the active presentation
            }
            return SafeActivePresentation();
        }

        // masterIndex is 1-based into pres.Designs; 0 (unset) = the default master.
        private static CustomLayout LayoutAt(Presentation pres, int masterIndex, int index) {
            Designs designs = pres.Designs;
            CustomLayouts layouts = masterIndex >= 1 && masterIndex <= designs.Count
                ? designs[masterIndex].SlideMaster.CustomLayouts
                : pres.SlideMaster.CustomLayouts;
            if (index < 1 || index > layouts.Count) {
                return null;
            }
            return layouts[index];
        }

        // Date / footer / slide number: never a data box (exporter LATENT_TYPES).
        private static bool IsLatentPlaceholder(Shape sh) {
            try {
                PpPlaceholderType t = sh.PlaceholderFormat.Type;
                return t == PpPlaceholderType.ppPlaceholderDate ||
                       t == PpPlaceholderType.ppPlaceholderFooter ||
                       t == PpPlaceholderType.ppPlaceholderSlideNumber;
            } catch {
                return false;
            }
        }

        private int ResolveStartIndex(Presentation pres, InsertPosition position) {
            int count = pres.Slides.Count;
            switch (position) {
                case InsertPosition.Front:
                    return 1;
                case InsertPosition.AfterCurrent:
                    try {
                        // Use the target deck's own window, not Application.ActiveWindow
                        // (which may belong to a different deck).
                        Slide cur = pres.Windows.Count > 0 ? pres.Windows[1].View.Slide as Slide : null;
                        return cur != null ? Math.Min(cur.SlideIndex + 1, count + 1) : count + 1;
                    } catch {
                        return count + 1;
                    }
                case InsertPosition.End:
                default:
                    return count + 1;
            }
        }

        private static void ApplyBoxText(Slide slide, LayoutSelection sel, Dictionary<int, string> boxText) {
            List<BoxSignature> sigs = sel.BoxSignatures;
            if (sigs == null) {
                return;
            }
            Placeholders phs = slide.Shapes.Placeholders;

            // Resolve every box to its placeholder ONCE, by geometry, before any
            // move — the CN tweak below changes box 2's top/height, which would
            // otherwise break a later geometry lookup.
            Shape[] boxShapes = new Shape[sigs.Count];
            for (int b = 0; b < sigs.Count; b++) {
                boxShapes[b] = FindPlaceholder(phs, sigs[b]);
            }

            // CN-box vertical-centering tweak (faithful to the Exporter): when the
            // EN box (1) holds a single line and the KR box (0) sits above EN,
            // nudge the CN box (2) to the midpoint and reset its height. This is
            // Praise-only — the Exporter does it solely in PPTXFile_Praise.addItem;
            // Word's box 2 is its 2nd-language placeholder and must NOT be moved
            // (gated by sel.CenterCnBox, which the wizard sets only for Praise).
            // VENUE layouts may leave optional boxes out (null signatures).
            string en;
            if (sel.CenterCnBox && sigs.Count >= 3 && boxShapes[2] != null &&
                sigs[0] != null && sigs[1] != null && sigs[2] != null &&
                boxText.TryGetValue(1, out en) && en != null && en.IndexOf('\n') < 0 &&
                sigs[0].Top < sigs[1].Top) {
                try {
                    float chW = boxShapes[2].Width;
                    boxShapes[2].Top = (sigs[2].Top + sigs[1].Top) / 2f;
                    boxShapes[2].Width = chW;
                    boxShapes[2].Height = sigs[2].Height;
                } catch {
                    // geometry tweak is best-effort
                }
            }

            for (int b = 0; b < boxShapes.Length; b++) {
                string text;
                if (!boxText.TryGetValue(b, out text)) {
                    if (!sel.FillUnusedBoxes) {
                        continue;
                    }
                    text = " "; // hide the "$lang2 …" prompt (VENUE EMPTY_TEXT)
                }
                Shape sh = boxShapes[b];
                if (sh == null) {
                    continue;
                }
                try {
                    if (sh.HasTextFrame == Core.MsoTriState.msoTrue) {
                        sh.TextFrame.TextRange.Text = text ?? "";
                    }
                } catch {
                    // skip a placeholder that won't accept text
                }
            }
        }

        // Map a box to its slide placeholder by geometry (height + top as in the
        // Exporter, plus left + width — see BoxSignature). Tolerant compare
        // because Interop reports points as floats.
        private static Shape FindPlaceholder(Placeholders phs, BoxSignature sig) {
            if (sig == null) {
                return null;
            }
            for (int p = 1; p <= phs.Count; p++) {
                Shape sh = phs[p];
                if (Math.Abs(sh.Height - sig.Height) < 0.5f && Math.Abs(sh.Top - sig.Top) < 0.5f &&
                    Math.Abs(sh.Left - sig.Left) < 0.5f && Math.Abs(sh.Width - sig.Width) < 0.5f) {
                    return sh;
                }
            }
            return null;
        }

        private static void SetNote(Slide slide, string note) {
            if (string.IsNullOrEmpty(note)) {
                return;
            }
            try {
                Shapes shapes = slide.NotesPage.Shapes;
                for (int i = 1; i <= shapes.Count; i++) {
                    Shape sh = shapes[i];
                    try {
                        if (sh.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderBody &&
                            sh.HasTextFrame == Core.MsoTriState.msoTrue) {
                            sh.TextFrame.TextRange.Text = note;
                            return;
                        }
                    } catch {
                        // not the notes body placeholder; keep looking
                    }
                }
            } catch {
                // notes are best-effort metadata; never fail the insert over them
            }
        }

        private static string SafeText(Shape sh) {
            try {
                if (sh.HasTextFrame == Core.MsoTriState.msoTrue && sh.TextFrame.HasText == Core.MsoTriState.msoTrue) {
                    return sh.TextFrame.TextRange.Text;
                }
            } catch {
                // ignore
            }
            return "";
        }

        private static string SafeLayoutName(CustomLayout layout) {
            try {
                return layout.Name;
            } catch {
                return "";
            }
        }

        #endregion
    }
}
