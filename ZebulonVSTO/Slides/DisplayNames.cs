using System;
using System.IO;

namespace ZebulonVSTO.Slides {
    /// <summary>
    /// Pure helpers for turning server paths into human-readable display names
    /// and safe filenames. Mirrors the derivations in Zebulon Web
    /// (SlideEditorPraise.tsx, TemplateLoader.tsx) closely enough for parity.
    /// COM-free; unit-tested.
    /// </summary>
    public static class DisplayNames {
        /// <summary>"lyrics/A-Z/Alive KEC.md" → "Alive KEC" (basename, extension stripped).</summary>
        public static string LyricName(string path) {
            string name = StripExtension(Basename(path));
            return name.Length > 0 ? name : (path ?? "");
        }

        /// <summary>"templates/PRAISE_CornerStone_Origin.pptx" → "CornerStone_Origin"
        /// (basename, extension + leading TYPE_ prefix stripped).</summary>
        public static string TemplateName(string path) {
            string name = StripExtension(Basename(path));
            int underscore = name.IndexOf('_');
            if (underscore >= 0 && underscore + 1 < name.Length) {
                name = name.Substring(underscore + 1);
            }
            return name.Length > 0 ? name : StripExtension(Basename(path));
        }

        /// <summary>"templates/PRAISE_*.pptx" → "Praise", "templates/WORD_*.pptx" → "Word", else "".</summary>
        public static string TemplateKind(string path) {
            string baseName = Basename(path);
            if (baseName.StartsWith("PRAISE_", StringComparison.OrdinalIgnoreCase)) {
                return "Praise";
            }
            if (baseName.StartsWith("WORD_", StringComparison.OrdinalIgnoreCase)) {
                return "Word";
            }
            return "";
        }

        /// <summary>A safe Windows .pptx filename derived from a server path basename.</summary>
        public static string SafePptxFileName(string path) {
            string name = Basename(path);
            if (string.IsNullOrEmpty(name)) {
                return "template.pptx";
            }
            foreach (char c in Path.GetInvalidFileNameChars()) {
                name = name.Replace(c, '_');
            }
            if (!name.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase)) {
                name += ".pptx";
            }
            return name;
        }

        private static string Basename(string path) {
            if (string.IsNullOrEmpty(path)) {
                return "";
            }
            int slash = path.LastIndexOfAny(new[] { '/', '\\' });
            return slash >= 0 ? path.Substring(slash + 1) : path;
        }

        private static string StripExtension(string name) {
            if (string.IsNullOrEmpty(name)) {
                return "";
            }
            int dot = name.LastIndexOf('.');
            return dot > 0 ? name.Substring(0, dot) : name;
        }
    }

    /// <summary>A lyric list row: server path + derived display name (for the picker).</summary>
    public class LyricEntry {
        public string Path { get; set; }
        public string DisplayName { get; set; }

        public static LyricEntry From(string path) {
            return new LyricEntry { Path = path, DisplayName = DisplayNames.LyricName(path) };
        }
    }

    /// <summary>A template combo row: server path + kind + display label.</summary>
    public class TemplateEntry {
        public string Path { get; set; }
        public string Kind { get; set; }
        public string DisplayName { get; set; }

        public static TemplateEntry From(string path) {
            string kind = DisplayNames.TemplateKind(path);
            string name = DisplayNames.TemplateName(path);
            return new TemplateEntry {
                Path = path,
                Kind = kind,
                DisplayName = (kind.Length > 0 ? "[" + kind + "] " : "") + name
            };
        }
    }
}
