using System;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

using ZebulonVSTO.Slides;

namespace ZebulonVSTO {
    /// <summary>
    /// Per-user preferences persisted as JSON at <see cref="PreferencesStore.DefaultPath"/>.
    /// One member per feature section so later features (e.g. sync ports) can add
    /// their own without disturbing existing ones. Unknown sections written by a
    /// newer build survive a round-trip through an older one (IExtensibleDataObject).
    /// COM-free; serialization uses the framework DataContractJsonSerializer.
    /// </summary>
    [DataContract]
    public sealed class Preferences : IExtensibleDataObject {
        /// <summary>WordSelectWindow's remembered languages/versions; null = never saved.</summary>
        [DataMember(Name = "wordSelect", EmitDefaultValue = false)]
        public WordSelectPrefs WordSelect { get; set; }

        public ExtensionDataObject ExtensionData { get; set; }

        public string ToJson() {
            DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(Preferences));
            using (MemoryStream ms = new MemoryStream()) {
                ser.WriteObject(ms, this);
                return Encoding.UTF8.GetString(ms.ToArray());
            }
        }

        /// <summary>Parses <paramref name="json"/>; blank or malformed input yields empty preferences.</summary>
        public static Preferences FromJson(string json) {
            if (string.IsNullOrWhiteSpace(json)) {
                return new Preferences();
            }
            try {
                DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(Preferences));
                using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(json))) {
                    return (ser.ReadObject(ms) as Preferences) ?? new Preferences();
                }
            } catch (SerializationException) {
                return new Preferences();
            }
        }
    }

    /// <summary>
    /// Best-effort load/save of <see cref="Preferences"/>. Never throws: a missing
    /// or unreadable file loads as empty preferences, and a failed save is reported
    /// via the return value only — remembering settings must never block the
    /// action that triggered it. Uninstall.ps1 removes the folder.
    /// </summary>
    public static class PreferencesStore {
        /// <summary>%APPDATA%\ZebulonVSTO\preferences.json — kept apart from the
        /// %LOCALAPPDATA% install folder so reinstalling/updating keeps it.</summary>
        public static string DefaultPath {
            get {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "ZebulonVSTO", "preferences.json");
            }
        }

        public static Preferences Load(string path = null) {
            try {
                string p = path ?? DefaultPath;
                return File.Exists(p) ? Preferences.FromJson(File.ReadAllText(p, Encoding.UTF8)) : new Preferences();
            } catch (Exception) {
                return new Preferences();
            }
        }

        public static bool Save(Preferences prefs, string path = null) {
            try {
                string p = path ?? DefaultPath;
                Directory.CreateDirectory(Path.GetDirectoryName(p));
                // Write a sibling temp file, then swap it in, so a crash mid-write
                // can't leave a truncated preferences.json behind.
                string tmp = p + ".tmp";
                File.WriteAllText(tmp, prefs.ToJson(), new UTF8Encoding(false));
                if (File.Exists(p)) {
                    File.Replace(tmp, p, null);
                } else {
                    File.Move(tmp, p);
                }
                return true;
            } catch (Exception) {
                return false;
            }
        }
    }
}
