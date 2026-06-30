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

        public List<LayoutDescriptor> ReadActiveLayouts() {
            List<LayoutDescriptor> result = new List<LayoutDescriptor>();
            if (_dispatcher == null) {
                return result;
            }
            try {
                _dispatcher.Invoke(() => {
                    Presentation pres = SafeActivePresentation();
                    if (pres == null) {
                        return;
                    }
                    CustomLayouts layouts = pres.SlideMaster.CustomLayouts;
                    for (int i = 1; i <= layouts.Count; i++) {
                        CustomLayout layout = layouts[i];
                        List<PlaceholderInfo> phInfos = new List<PlaceholderInfo>();
                        Placeholders phs = layout.Shapes.Placeholders;
                        for (int p = 1; p <= phs.Count; p++) {
                            Shape sh = phs[p];
                            phInfos.Add(new PlaceholderInfo {
                                Text = SafeText(sh),
                                Height = sh.Height,
                                Top = sh.Top
                            });
                        }
                        result.Add(new LayoutDescriptor {
                            LayoutIndex = i,
                            Name = SafeLayoutName(layout),
                            Placeholders = phInfos
                        });
                    }
                });
            } catch (Exception e) {
                LogError("Failed to read active layouts.", e);
            }
            return result;
        }

        public int ExecutePlan(LayoutSelection selection, List<SlidePlanItem> plan, InsertPosition position) {
            if (_dispatcher == null || selection == null || plan == null) {
                return 0;
            }
            int inserted = 0;
            try {
                _dispatcher.Invoke(() => {
                    Presentation pres = SafeActivePresentation();
                    if (pres == null) {
                        return;
                    }
                    CustomLayouts layouts = pres.SlideMaster.CustomLayouts;
                    CustomLayout bindLayout = LayoutAt(layouts, selection.BindLayoutIndex);
                    CustomLayout emptyLayout = LayoutAt(layouts, selection.EmptyLayoutIndex);
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
                });
            } catch (Exception e) {
                LogError("Failed to insert generated slides.", e);
            }
            return inserted;
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

        private static CustomLayout LayoutAt(CustomLayouts layouts, int index) {
            if (index < 1 || index > layouts.Count) {
                return null;
            }
            return layouts[index];
        }

        private int ResolveStartIndex(Presentation pres, InsertPosition position) {
            int count = pres.Slides.Count;
            switch (position) {
                case InsertPosition.Front:
                    return 1;
                case InsertPosition.AfterCurrent:
                    try {
                        Slide cur = Application.ActiveWindow.View.Slide as Slide;
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
            // nudge the CN box (2) to the midpoint and reset its height.
            string en;
            if (sigs.Count >= 3 && boxShapes[2] != null &&
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

            foreach (KeyValuePair<int, string> kv in boxText) {
                if (kv.Key < 0 || kv.Key >= boxShapes.Length) {
                    continue;
                }
                Shape sh = boxShapes[kv.Key];
                if (sh == null) {
                    continue;
                }
                try {
                    if (sh.HasTextFrame == Core.MsoTriState.msoTrue) {
                        sh.TextFrame.TextRange.Text = kv.Value ?? "";
                    }
                } catch {
                    // skip a placeholder that won't accept text
                }
            }
        }

        // Map a box to its slide placeholder by geometry (height + top), mirroring
        // the Exporter. Tolerant compare because Interop reports points as floats.
        private static Shape FindPlaceholder(Placeholders phs, BoxSignature sig) {
            if (sig == null) {
                return null;
            }
            for (int p = 1; p <= phs.Count; p++) {
                Shape sh = phs[p];
                if (Math.Abs(sh.Height - sig.Height) < 0.5f && Math.Abs(sh.Top - sig.Top) < 0.5f) {
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
