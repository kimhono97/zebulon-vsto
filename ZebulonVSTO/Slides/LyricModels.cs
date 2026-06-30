using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

namespace ZebulonVSTO.Slides {
    /// <summary>
    /// Lyric document as served by the Zebulon Provider (POST /lyric). Shape
    /// mirrors zebulon-provider/src/app/lyric/Lyric.ts:
    ///   { title: string, groups: [ { name: string, pages: string[][] } ] }
    /// where each page is an array of language strings [KR, EN, CN].
    /// COM-free; deserialized with the framework DataContractJsonSerializer
    /// (no System.Text.Json — same constraint as the rest of the add-in).
    /// </summary>
    [DataContract]
    public class Lyric {
        [DataMember(Name = "title")] public string Title { get; set; }
        [DataMember(Name = "groups")] public List<LyricGroup> Groups { get; set; }
    }

    [DataContract]
    public class LyricGroup {
        [DataMember(Name = "name")] public string Name { get; set; }
        [DataMember(Name = "pages")] public List<List<string>> Pages { get; set; }
    }

    /// <summary>JSON (de)serialization helpers for Provider responses.</summary>
    public static class LyricJson {
        public static Lyric ParseLyric(string json) {
            return Parse<Lyric>(json);
        }

        /// <summary>Parse a JSON string array (e.g. GET /lyric, GET /template).</summary>
        public static List<string> ParseStringArray(string json) {
            List<string> list = Parse<List<string>>(json);
            return list ?? new List<string>();
        }

        private static T Parse<T>(string json) where T : class {
            if (string.IsNullOrEmpty(json)) {
                return null;
            }
            try {
                using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(json))) {
                    DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(T));
                    return serializer.ReadObject(ms) as T;
                }
            } catch {
                return null;
            }
        }
    }
}
