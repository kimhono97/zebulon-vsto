namespace ZebulonVSTO.Slides {
    /// <summary>
    /// Pure per-box text rules for Praise slides, ported 1:1 from the Zebulon
    /// Exporter (zebulon-exporter/py/pptx_exporter.py, PPTXFile_Praise.addItem)
    /// so VSTO-inserted slides match the Exporter's output:
    ///   - box 0 (KR) and box 1 (EN): upper-cased
    ///   - box 0 (KR): a single-line value gets a leading newline (the Exporter's
    ///     vertical-centering hack)
    ///   - box 2 (CN): left as-is
    /// COM-free; unit-tested.
    /// </summary>
    public static class TextTransforms {
        public static string TransformPraiseBox(int boxIndex, string text) {
            if (text == null) {
                text = "";
            }
            if (boxIndex < 2) {
                text = text.ToUpperInvariant();
            }
            if (boxIndex == 0 && text.IndexOf('\n') < 0) {
                text = "\n" + text;
            }
            return text;
        }
    }
}
