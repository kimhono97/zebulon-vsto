using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

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

        /// <summary>Provider template type → label: "VENUE" → "Venue", "PRAISE" → "Praise",
        /// "WORD" → "Word"; an unknown type is shown as given.</summary>
        public static string TemplateKind(string type) {
            switch ((type ?? "").ToUpperInvariant()) {
                case "VENUE": return "Venue";
                case "PRAISE": return "Praise";
                case "WORD": return "Word";
                default: return type ?? "";
            }
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

        /// <summary>From a GET /templates row; type and name come from the Provider.</summary>
        public static TemplateEntry From(ProviderTemplate t) {
            string kind = DisplayNames.TemplateKind(t.Type);
            string name = string.IsNullOrEmpty(t.Name) ? t.Path ?? "" : t.Name;
            return new TemplateEntry {
                Path = t.Path,
                Kind = kind,
                DisplayName = (kind.Length > 0 ? "[" + kind + "] " : "") + name
            };
        }
    }

    /// <summary>
    /// One row of the Provider's GET /templates: { path, type: "PRAISE" | "WORD" |
    /// "VENUE", name } (name = basename after the prefix, without .pptx).
    /// </summary>
    [DataContract]
    public class ProviderTemplate {
        [DataMember(Name = "path")] public string Path { get; set; }
        [DataMember(Name = "type")] public string Type { get; set; }
        [DataMember(Name = "name")] public string Name { get; set; }

        /// <summary>Parse the listing; rows without a path are dropped, malformed input yields [].</summary>
        public static List<ProviderTemplate> ParseList(string json) {
            List<ProviderTemplate> result = new List<ProviderTemplate>();
            if (string.IsNullOrEmpty(json)) {
                return result;
            }
            try {
                using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(json))) {
                    DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(List<ProviderTemplate>));
                    List<ProviderTemplate> list = serializer.ReadObject(ms) as List<ProviderTemplate>;
                    if (list != null) {
                        result.AddRange(list.Where(t => t != null && !string.IsNullOrEmpty(t.Path)));
                    }
                }
            } catch {
                // malformed listing → empty
            }
            return result;
        }
    }
}
