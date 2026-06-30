namespace ZebulonVSTO.Slides {
    /// <summary>
    /// Shared defaults for the slide-generation feature. Kept free of any
    /// PowerPoint/COM dependency so the pure data/logic types can be unit-tested
    /// in isolation (mirrors <see cref="ZebulonVSTO.Sync.SyncDefaults"/>).
    /// </summary>
    public static class SlideGenDefaults {
        // Zebulon Web (jym-workbox) base URL. The Provider URL is normally
        // resolved at runtime via {WebBaseUrl}/api/proj?n=zebulon (urls.alias);
        // this also hosts the Bible API used in Phase B. Leave empty to skip
        // /api/proj resolution and use ProviderBaseUrl directly.
        // TODO(confirm): set the canonical Zebulon Web deployment URL.
        public const string WebBaseUrl = "https://jym-workbox.vercel.app";

        // Fallback Provider base URL (used when /api/proj resolution is skipped
        // or fails). TODO(confirm): verify this is the current deployment.
        public const string ProviderBaseUrl = "https://zebulon-provider-2.vercel.app";

        // Bind-layout box counts (must match the Exporter's boxCount).
        public const int PraiseBoxCount = 3; // [KR, EN, CN]
        public const int WordBoxCount = 4;   // [book, KR, EN, CN]

        /// <summary>
        /// Marker prefix a bind-layout placeholder's text must start with to map
        /// to box <paramref name="boxIndex"/> (0-based): "$1 ", "$2 ", … — the
        /// same convention the Exporter scans for (pptx_exporter.py).
        /// </summary>
        public static string BoxMarker(int boxIndex) {
            return "$" + (boxIndex + 1) + " ";
        }
    }
}
