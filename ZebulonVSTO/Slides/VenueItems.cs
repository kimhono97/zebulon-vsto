using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

namespace ZebulonVSTO.Slides {
    /// <summary>
    /// The Exporter's /venue/export payload ({options, items}), mirrored as the
    /// add-in's slide-plan input: the wizard builds <see cref="VenueItem"/>s and
    /// <see cref="VenuePlanner"/> turns them into slides, so a payload built by
    /// Zebulon Web could later be planned the same way. Mirrored: item types
    /// praise / word / blank (planned) and fixed / layout (carried, not planned),
    /// options fixedSlides / fontNames / footerText (carried, not applied).
    /// Not mirrored: a layout item's <c>data.fields</c> (a string or paragraph
    /// list per field, which DataContractJsonSerializer cannot type).
    /// COM-free; unit-tested.
    /// </summary>
    [DataContract]
    public class VenuePayload {
        [DataMember(Name = "options", EmitDefaultValue = false, Order = 0)] public VenueOptions Options { get; set; }
        [DataMember(Name = "items", Order = 1)] public List<VenueItem> Items { get; set; }

        public string ToJson() {
            DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(VenuePayload));
            using (MemoryStream ms = new MemoryStream()) {
                ser.WriteObject(ms, this);
                return Encoding.UTF8.GetString(ms.ToArray());
            }
        }

        /// <summary>Parses <paramref name="json"/>; null when it is blank or malformed.</summary>
        public static VenuePayload FromJson(string json) {
            if (string.IsNullOrWhiteSpace(json)) {
                return null;
            }
            try {
                DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(VenuePayload));
                using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(json))) {
                    return ser.ReadObject(ms) as VenuePayload;
                }
            } catch (SerializationException) {
                return null;
            }
        }
    }

    [DataContract]
    public class VenueOptions {
        /// <summary>"template" (default) or "items".</summary>
        [DataMember(Name = "fixedSlides", EmitDefaultValue = false)] public string FixedSlides { get; set; }
        [DataMember(Name = "fontNames", EmitDefaultValue = false)] public List<string> FontNames { get; set; }
        [DataMember(Name = "footerText", EmitDefaultValue = false)] public string FooterText { get; set; }
    }

    /// <summary>One payload item: <c>{type, layout?, data?, fontNames?}</c>.</summary>
    [DataContract]
    public class VenueItem {
        public const string PraiseType = "praise";
        public const string WordType = "word";
        public const string BlankType = "blank";
        public const string FixedType = "fixed";
        public const string LayoutType = "layout";

        [DataMember(Name = "type", Order = 0)] public string Type { get; set; }
        [DataMember(Name = "layout", EmitDefaultValue = false, Order = 1)] public string Layout { get; set; }
        [DataMember(Name = "data", EmitDefaultValue = false, Order = 2)] public VenueItemData Data { get; set; }
        [DataMember(Name = "fontNames", EmitDefaultValue = false, Order = 3)] public List<string> FontNames { get; set; }

        /// <summary>A praise item; the display switches are written only when they differ from the default (true).</summary>
        public static VenueItem Praise(Lyric lyric, bool titleSlide = true, bool separators = true) {
            return new VenueItem {
                Type = PraiseType,
                Data = new VenueItemData {
                    Title = lyric.Title,
                    Groups = lyric.Groups,
                    TitleSlide = titleSlide ? (bool?)null : false,
                    Separators = separators ? (bool?)null : false
                }
            };
        }

        /// <summary>A word item; the structured header when known, else the plain headText.</summary>
        public static VenueItem Word(WordItem item) {
            return new VenueItem {
                Type = WordType,
                Data = new VenueItemData {
                    Header = item.Header,
                    HeadText = item.Header == null ? item.HeadText : null,
                    Sections = item.Sections,
                    First = item.First,
                    Second = item.Second,
                    Third = item.Third
                }
            };
        }

        public static VenueItem Blank(string notes = null) {
            return new VenueItem { Type = BlankType, Data = notes == null ? null : new VenueItemData { Notes = notes } };
        }
    }

    /// <summary>
    /// An item's <c>data</c>: the union of the praise, word, blank and fixed
    /// schemas (members of other types stay unset and are not written).
    /// </summary>
    [DataContract]
    public class VenueItemData {
        // praise
        [DataMember(Name = "title", EmitDefaultValue = false)] public string Title { get; set; }
        [DataMember(Name = "titleSlide", EmitDefaultValue = false)] public bool? TitleSlide { get; set; }
        [DataMember(Name = "separators", EmitDefaultValue = false)] public bool? Separators { get; set; }
        [DataMember(Name = "groups", EmitDefaultValue = false)] public List<LyricGroup> Groups { get; set; }
        // word
        [DataMember(Name = "header", EmitDefaultValue = false)] public VenueHeader Header { get; set; }
        [DataMember(Name = "headText", EmitDefaultValue = false)] public string HeadText { get; set; }
        [DataMember(Name = "sections", EmitDefaultValue = false)] public List<WordSection> Sections { get; set; }
        [DataMember(Name = "first", EmitDefaultValue = false)] public List<List<string>> First { get; set; }
        [DataMember(Name = "second", EmitDefaultValue = false)] public List<List<string>> Second { get; set; }
        [DataMember(Name = "third", EmitDefaultValue = false)] public List<List<string>> Third { get; set; }
        // blank, layout
        [DataMember(Name = "notes", EmitDefaultValue = false)] public string Notes { get; set; }
        // fixed
        [DataMember(Name = "key", EmitDefaultValue = false)] public string Key { get; set; }
        [DataMember(Name = "index", EmitDefaultValue = false)] public int? Index { get; set; }

        public Lyric ToLyric() {
            return new Lyric { Title = Title, Groups = Groups };
        }

        public WordItem ToWordItem() {
            return new WordItem {
                Header = Header,
                HeadText = Header != null ? Header.Format() : HeadText,
                Sections = Sections,
                First = First,
                Second = Second,
                Third = Third
            };
        }
    }

    /// <summary>A word verse header <c>{books: [first, second?, third?], ref?}</c>.</summary>
    [DataContract]
    public class VenueHeader {
        [DataMember(Name = "books")] public List<string> Books { get; set; }
        [DataMember(Name = "ref", EmitDefaultValue = false)] public string Ref { get; set; }

        /// <summary>
        /// "first (second, third) ref", dropping empty names: "first (second) ref",
        /// "first ref" (venue/items/word.py setHeader, as plain text — the
        /// exporter's per-language fonts for the 2nd/3rd names are not applied).
        /// </summary>
        public string Format() {
            List<string> books = Books ?? new List<string>();
            StringBuilder sb = new StringBuilder(books.Count > 0 ? books[0] ?? "" : "");
            List<string> others = books.Skip(1).Take(2).Where(b => !string.IsNullOrEmpty(b)).ToList();
            if (others.Count > 0) {
                sb.Append(" (").Append(string.Join(", ", others)).Append(")");
            }
            if (!string.IsNullOrEmpty(Ref)) {
                sb.Append(" ").Append(Ref);
            }
            return sb.ToString();
        }
    }

    /// <summary>Result of <see cref="VenuePlanner.BuildPlan"/>.</summary>
    public class VenuePlan {
        public VenuePlan() {
            Slides = new List<SlidePlanItem>();
            Skipped = new List<int>();
        }

        public List<SlidePlanItem> Slides { get; private set; }

        /// <summary>Indexes of items that produced no slides: an unsupported type
        /// (fixed, layout, unknown) or invalid data.</summary>
        public List<int> Skipped { get; private set; }
    }

    /// <summary>
    /// Plans payload items in order, like the Exporter's DeckBuilder: praise via
    /// <see cref="PraisePlanner"/> (songs numbered across the payload), word via
    /// <see cref="WordPlanner"/>, blank as one empty slide with its notes.
    /// fixed and layout items are skipped: the template's fixed slides are already
    /// in the deck the user opened, and layout slides are not planned (yet).
    /// COM-free; unit-tested.
    /// </summary>
    public static class VenuePlanner {
        public static VenuePlan BuildPlan(IReadOnlyList<VenueItem> items) {
            VenuePlan result = new VenuePlan();
            if (items == null) {
                return result;
            }
            int songNo = 0;
            for (int i = 0; i < items.Count; i++) {
                VenueItem item = items[i];
                int before = result.Slides.Count;
                VenueItemData data = item != null ? item.Data : null;
                switch (item != null ? item.Type : null) {
                    case VenueItem.PraiseType:
                        Lyric lyric = data != null ? data.ToLyric() : null;
                        if (PraisePlanner.IsValid(lyric)) {
                            PraisePlanner.AddSong(result.Slides, lyric, ++songNo,
                                                  data.TitleSlide ?? true, data.Separators ?? true);
                        }
                        break;
                    case VenueItem.WordType:
                        if (data != null) {
                            result.Slides.AddRange(WordPlanner.BuildPlan(new[] { data.ToWordItem() }));
                        }
                        break;
                    case VenueItem.BlankType:
                        result.Slides.Add(new SlidePlanItem { Kind = LayoutKind.Empty, Note = data != null ? data.Notes : null });
                        break;
                }
                if (result.Slides.Count == before) {
                    result.Skipped.Add(i);
                }
            }
            return result;
        }
    }
}
