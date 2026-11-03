using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ZebulonVSTO.Slides {
    /// <summary>
    /// VENUE template rules, mirrored from the Exporter's /venue contract (spec
    /// 1.6; zebulon-exporter/venue/layout.py, fields.py, template.py and the
    /// praise/word/blank item rules). Given COM-free layout descriptors, finds
    /// each layout's role by name suffix (…LyricPage / …Word / …EmptyPage), its
    /// numbered ($N) or named ($lang1, $title.lang1, $chapter.lang1.ruby, …)
    /// tokens, and whether it is valid — the same answers as /venue/inspect's
    /// <c>roles</c> and <c>layouts[]</c>, so the two can be checked against each
    /// other (tests/ZebulonVSTO.Tests/Fixtures/venue).
    /// Not mirrored: placeholder idx (not exposed by PowerPoint COM), paragraph
    /// style samples (@list …), geometry/colour/background/media and fixed-slide
    /// read-back — they serve the web preview, which PowerPoint does itself.
    /// COM-free; unit-tested.
    /// </summary>
    public static class VenueRules {
        public const string Praise = "praise";
        public const string Word = "word";
        public const string Blank = "blank";
        public const int MaxLangs = 3;

        // "$1 ...", "$1" or "$1\r..." marks token 1 (layout.py TOKEN_PATTERN).
        private static readonly Regex NumberedToken = new Regex(@"^\s*\$([0-9]+)(?=\s|$)");
        // fields.py NAMED_TOKEN / NAMED_START / LANG_NAME.
        private static readonly Regex NamedToken = new Regex(
            @"^\s*\$([A-Za-z][A-Za-z0-9]*)(?:\.(lang[0-9]+))?(?:\.(ruby))?(?=\s|$)", RegexOptions.IgnoreCase);
        private static readonly Regex NamedStart = new Regex(@"^\s*\$[A-Za-z]");
        private static readonly Regex LangName = new Regex(@"^lang([0-9]+)$");

        // Item layout rules (venue/items/praise.py, word.py, blank.py).
        private static readonly RoleRule[] Rules = {
            new RoleRule(Praise, "LyricPage", new[] { 1 }, new[] { 2, 3 }),
            new RoleRule(Word, "Word", new[] { 1, 2 }, new[] { 3, 4 }),
            new RoleRule(Blank, "EmptyPage", new int[0], new int[0]),
        };

        // Name suffix -> kind hint (layout.py KINDS); longest suffix first.
        private static readonly string[][] Kinds = {
            new[] { "Chapter+Part", "chapterPart" }, new[] { "Chapter", "chapter" }, new[] { "Title", "title" },
            new[] { "Prayer", "prayer" }, new[] { "Order", "order" },
        };

        /// <summary>Classify every layout and resolve the praise/word/blank roles.</summary>
        public static VenueScan Scan(IReadOnlyList<LayoutDescriptor> layouts) {
            List<VenueLayout> result = new List<VenueLayout>();
            if (layouts != null) {
                foreach (LayoutDescriptor d in layouts) {
                    VenueLayout layout = ReadLayout(d);
                    RoleRule rule = RuleForName(layout.Name);
                    if (rule != null) {
                        ApplyRule(layout, rule);
                    }
                    result.Add(layout);
                }
            }
            return new VenueScan(result);
        }

        /// <summary>The kind hint for a layout name (title, chapter, chapterPart, prayer, order) or null.</summary>
        public static string KindOf(string layoutName) {
            string name = (layoutName ?? "").Trim().ToLowerInvariant();
            foreach (string[] k in Kinds) {
                if (name.EndsWith(k[0].ToLowerInvariant(), StringComparison.Ordinal)) {
                    return k[1];
                }
            }
            return null;
        }

        /// <summary>
        /// Parse a prompt's first paragraph (fields.py parseNamedToken). Returns
        /// null when it has no named token, a result with <see cref="NamedTokenParse.Error"/>
        /// set for a malformed one, else the token ("title.lang1", "lang1.ruby", "header", …).
        /// </summary>
        public static NamedTokenParse ParseNamedToken(string text) {
            text = text ?? "";
            Match m = NamedToken.Match(text);
            if (!m.Success) {
                return NamedStart.IsMatch(text)
                    ? NamedTokenParse.Fail("malformed token '" + FirstWord(text) + "'")
                    : null;
            }
            string name = m.Groups[1].Value.ToLowerInvariant();
            string lang = m.Groups[2].Value.ToLowerInvariant();
            bool ruby = m.Groups[3].Success;
            string whole = m.Value.Trim();
            if (LangName.IsMatch(name)) {
                if (lang.Length > 0) {
                    return NamedTokenParse.Fail("malformed token '" + whole + "'");
                }
                lang = name;
                name = "text";
            }
            int? n = null;
            if (lang.Length > 0) {
                int parsed;
                if (!int.TryParse(lang.Substring(4), out parsed) || parsed < 1 || parsed > MaxLangs) {
                    return NamedTokenParse.Fail("token '" + whole + "': languages are lang1..lang" + MaxLangs);
                }
                n = parsed;
            }
            if (ruby && n == null) {
                return NamedTokenParse.Fail("token '" + whole + "': .ruby needs a .langN before it");
            }
            string token = (name == "text" ? "lang" + n : name + (n != null ? ".lang" + n : "")) + (ruby ? ".ruby" : "");
            return new NamedTokenParse {
                Token = token,
                Field = name,
                Lang = n,
                Ruby = ruby,
                Label = text.Substring(m.Index + m.Length).Trim().Replace('\v', '\n')
            };
        }

        // layout.py readLayout / readField.
        private static VenueLayout ReadLayout(LayoutDescriptor d) {
            VenueLayout layout = new VenueLayout(d);
            if (d.Placeholders == null) {
                return layout;
            }
            foreach (PlaceholderInfo ph in d.Placeholders) {
                if (ph == null) {
                    continue;
                }
                if (!ph.IsLatent) {
                    layout.ContentPlaceholders++;
                }
                if (ph.IsLatent || !ph.HasTextFrame) {
                    continue;
                }
                string text = ph.Text ?? "";
                Match m = NumberedToken.Match(text);
                if (!m.Success) {
                    ReadField(layout, ph, FirstParagraph(text));
                    continue;
                }
                int token;
                if (!int.TryParse(m.Groups[1].Value, out token)) {
                    continue;
                }
                if (layout.Boxes.ContainsKey(token)) {
                    layout.Issues.Add("duplicate token $" + token);
                    continue;
                }
                layout.Boxes[token] = ph;
            }
            if (layout.Boxes.Count > 0 && layout.Fields.Count > 0) {
                layout.Issues.Add("mixes numbered ($N) and named tokens");
            }
            return layout;
        }

        private static void ReadField(VenueLayout layout, PlaceholderInfo ph, string firstLine) {
            NamedTokenParse parsed = ParseNamedToken(firstLine);
            if (parsed == null) {
                return;
            }
            if (parsed.Error != null) {
                layout.Issues.Add(parsed.Error);
                return;
            }
            if (layout.Fields.Any(f => f.Token == parsed.Token)) {
                layout.Issues.Add("duplicate token $" + parsed.Token);
                return;
            }
            layout.Fields.Add(new VenueField {
                Token = parsed.Token,
                Field = parsed.Field,
                Lang = parsed.Lang,
                Ruby = parsed.Ruby,
                Label = parsed.Label,
                Placeholder = ph
            });
        }

        // template.py __ruleForName: the longest matching suffix wins.
        private static RoleRule RuleForName(string name) {
            string lower = (name ?? "").Trim().ToLowerInvariant();
            RoleRule best = null;
            foreach (RoleRule r in Rules) {
                if (lower.EndsWith(r.Suffix.ToLowerInvariant(), StringComparison.Ordinal) &&
                    (best == null || r.Suffix.Length > best.Suffix.Length)) {
                    best = r;
                }
            }
            return best;
        }

        // layout.py applyRule + mapAliases.
        private static void ApplyRule(VenueLayout layout, RoleRule rule) {
            layout.Role = rule.Role;
            MapAliases(layout, rule.Role);
            List<int> missing = rule.Required.Where(t => !layout.Boxes.ContainsKey(t)).OrderBy(t => t).ToList();
            List<int> extra = layout.Boxes.Keys
                .Where(t => !rule.Required.Contains(t) && !rule.Optional.Contains(t)).OrderBy(t => t).ToList();
            if (missing.Count > 0) {
                layout.Issues.Add("missing token(s) " + string.Join(", ", missing.Select(t => "$" + t)));
            }
            if (extra.Count > 0) {
                layout.Issues.Add("unexpected token(s) " + string.Join(", ", extra.Select(t => "$" + t)));
            }
            if (rule.TokenCount == 0 && layout.ContentPlaceholders > 0 && extra.Count == 0) {
                layout.Issues.Add("blank layout should have no content placeholders");
            }
        }

        /// <summary>
        /// Praise and word layouts may use named tokens instead of $N: praise takes
        /// the languages in order as $1..3; word $header = $1 and the languages in
        /// order = $2..4. In both, $lang1.ruby (a reading line) takes the place of
        /// the second language when present. Other named tokens are an issue.
        /// </summary>
        private static void MapAliases(VenueLayout layout, string role) {
            if (layout.Fields.Count == 0 || layout.Boxes.Count > 0) {
                return;
            }
            string[] langs = layout.Fields.Any(f => f.Ruby)
                ? new[] { "lang1", "lang1.ruby", "lang2" }
                : new[] { "lang1", "lang2", "lang3" };
            Dictionary<string, int> mapping = new Dictionary<string, int>();
            if (role == Praise) {
                for (int i = 0; i < langs.Length; i++) {
                    mapping[langs[i]] = i + 1;
                }
            } else if (role == Word) {
                mapping["header"] = 1;
                for (int i = 0; i < langs.Length; i++) {
                    mapping[langs[i]] = i + 2;
                }
            }
            foreach (VenueField f in layout.Fields) {
                int n;
                if (!mapping.TryGetValue(f.Token, out n)) {
                    layout.Issues.Add("unexpected token $" + f.Token);
                    continue;
                }
                layout.Boxes[n] = f.Placeholder;
            }
        }

        // PowerPoint separates paragraphs with '\r' (python-pptx: paragraphs[0]).
        private static string FirstParagraph(string text) {
            int cut = text.IndexOfAny(new[] { '\r', '\n' });
            return cut >= 0 ? text.Substring(0, cut) : text;
        }

        private static string FirstWord(string text) {
            string[] parts = text.Trim().Split((char[])null, 2, StringSplitOptions.RemoveEmptyEntries);
            return parts.Length > 0 ? parts[0] : "";
        }

        private sealed class RoleRule {
            public readonly string Role;
            public readonly string Suffix;
            public readonly int[] Required;
            public readonly int[] Optional;

            public RoleRule(string role, string suffix, int[] required, int[] optional) {
                Role = role;
                Suffix = suffix;
                Required = required;
                Optional = optional;
            }

            public int TokenCount {
                get { return Required.Concat(Optional).DefaultIfEmpty(0).Max(); }
            }
        }
    }

    /// <summary>Result of <see cref="VenueRules.ParseNamedToken"/>.</summary>
    public class NamedTokenParse {
        public string Token { get; set; }
        public string Field { get; set; }
        public int? Lang { get; set; }
        public bool Ruby { get; set; }
        public string Label { get; set; }
        public string Error { get; set; }

        public static NamedTokenParse Fail(string error) {
            return new NamedTokenParse { Error = error };
        }
    }

    /// <summary>One named-token placeholder (inspect <c>layouts[].fields[]</c>, without idx/styles).</summary>
    public class VenueField {
        public string Token { get; set; }
        public string Field { get; set; }
        public int? Lang { get; set; }
        public bool Ruby { get; set; }
        public string Label { get; set; }
        public PlaceholderInfo Placeholder { get; set; }
    }

    /// <summary>One classified layout (inspect <c>layouts[]</c> subset).</summary>
    public class VenueLayout {
        public VenueLayout(LayoutDescriptor descriptor) {
            Descriptor = descriptor;
            Boxes = new SortedDictionary<int, PlaceholderInfo>();
            Fields = new List<VenueField>();
            Issues = new List<string>();
        }

        public LayoutDescriptor Descriptor { get; private set; }
        public string Name { get { return Descriptor.Name ?? ""; } }
        public int MasterIndex { get { return Descriptor.MasterIndex; } }
        public int LayoutIndex { get { return Descriptor.LayoutIndex; } }

        /// <summary>praise / word / blank, or null for any other layout.</summary>
        public string Role { get; set; }
        public string Kind { get { return VenueRules.KindOf(Name); } }

        /// <summary>Placeholders other than date/footer/slide number.</summary>
        public int ContentPlaceholders { get; set; }

        /// <summary>Token number ($N, or a named token's alias) -> placeholder.</summary>
        public SortedDictionary<int, PlaceholderInfo> Boxes { get; private set; }
        public List<VenueField> Fields { get; private set; }
        public List<string> Issues { get; private set; }

        public bool IsValid { get { return Role != null && Issues.Count == 0; } }

        /// <summary>Usable by a <c>layout</c> item: named tokens and no issues.</summary>
        public bool Fillable { get { return Fields.Count > 0 && Issues.Count == 0; } }
    }

    /// <summary>All layouts of a deck, classified (inspect <c>roles</c> + <c>layouts</c>).</summary>
    public class VenueScan {
        public VenueScan(List<VenueLayout> layouts) {
            Layouts = layouts;
        }

        public List<VenueLayout> Layouts { get; private set; }

        /// <summary>Valid layouts of a role, in deck order.</summary>
        public List<VenueLayout> Candidates(string role) {
            return Layouts.Where(l => l.Role == role && l.IsValid).ToList();
        }

        /// <summary>
        /// The layout the Exporter uses for a role when an item names none: the
        /// first valid one; for blank, else the first role-less layout without
        /// content placeholders (template.py resolve / __blankFallback). Null if none.
        /// </summary>
        public VenueLayout Resolve(string role) {
            VenueLayout first = Layouts.FirstOrDefault(l => l.Role == role && l.IsValid);
            if (first != null || role != VenueRules.Blank) {
                return first;
            }
            return Layouts.FirstOrDefault(l => l.Role == null && l.ContentPlaceholders == 0);
        }
    }
}
