using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using Xunit;
using ZebulonVSTO.Slides;

namespace ZebulonVSTO.Tests {
    public class VenueTokenTests {
        [Theory]
        [InlineData("$title.lang1 설교 제목", "title.lang1", "title", 1, false, "설교 제목")]
        [InlineData("$Title.LANG2", "title.lang2", "title", 2, false, "")]
        [InlineData("$chapter.lang1.ruby だいみだし", "chapter.lang1.ruby", "chapter", 1, true, "だいみだし")]
        [InlineData("$lang1 72p, 페이지", "lang1", "text", 1, false, "72p, 페이지")]
        [InlineData("$lang1.ruby", "lang1.ruby", "text", 1, true, "")]
        [InlineData("  $header 데살로니가전서", "header", "header", null, false, "데살로니가전서")]
        [InlineData("$name\v홍길동", "name", "name", null, false, "홍길동")]
        public void ParsesNamedTokens(string text, string token, string field, int? lang, bool ruby, string label) {
            NamedTokenParse p = VenueRules.ParseNamedToken(text);
            Assert.Null(p.Error);
            Assert.Equal(token, p.Token);
            Assert.Equal(field, p.Field);
            Assert.Equal(lang, p.Lang);
            Assert.Equal(ruby, p.Ruby);
            Assert.Equal(label, p.Label);
        }

        [Theory]
        [InlineData("$title.lang4 x")]
        [InlineData("$title.ruby x")]
        [InlineData("$lang1.lang2")]
        [InlineData("$title.lang1x")]
        public void ReportsMalformedTokens(string text) {
            NamedTokenParse p = VenueRules.ParseNamedToken(text);
            Assert.NotNull(p);
            Assert.NotNull(p.Error);
        }

        [Theory]
        [InlineData("설교 제목")]
        [InlineData("$1 한국어")]
        [InlineData("")]
        [InlineData(null)]
        public void NoNamedToken(string text) {
            Assert.Null(VenueRules.ParseNamedToken(text));
        }

        [Theory]
        [InlineData("CS Chapter+Part", "chapterPart")]
        [InlineData("BgArt Chapter", "chapter")]
        [InlineData("YoMori Title", "title")]
        [InlineData("DSM Prayer ", "prayer")]
        [InlineData("CS Order", "order")]
        [InlineData("UG Prepare", null)]
        [InlineData("CS LyricPage", null)]
        public void KindFromNameSuffix(string name, string kind) {
            Assert.Equal(kind, VenueRules.KindOf(name));
        }
    }

    public class VenueScanTests {
        private static PlaceholderInfo Ph(string text, float top = 0, float left = 0) {
            return new PlaceholderInfo { Text = text, Top = top, Left = left, Height = 50, Width = 100 };
        }

        private static PlaceholderInfo Footer() {
            return new PlaceholderInfo { Text = "account", IsLatent = true };
        }

        private static LayoutDescriptor Layout(int index, string name, params PlaceholderInfo[] phs) {
            return new LayoutDescriptor { MasterIndex = 1, LayoutIndex = index, Name = name, Placeholders = phs };
        }

        [Fact]
        public void PraiseWithReadingLine_TakesTheSecondBox() {
            VenueLayout l = VenueRules.Scan(new[] {
                Layout(1, "YoMori LyricPage", Ph("$lang2 한국어", 300), Ph("$lang1 日本語", 100), Ph("$lang1.ruby にほんご", 200))
            }).Layouts[0];

            Assert.Equal(VenueRules.Praise, l.Role);
            Assert.True(l.IsValid);
            Assert.Equal(new[] { 1, 2, 3 }, l.Boxes.Keys.ToArray());
            Assert.Equal(100f, l.Boxes[1].Top); // lang1
            Assert.Equal(200f, l.Boxes[2].Top); // lang1.ruby
            Assert.Equal(300f, l.Boxes[3].Top); // lang2
        }

        [Fact]
        public void Word_OptionalLanguagesMayBeMissing() {
            VenueLayout l = VenueRules.Scan(new[] {
                Layout(1, "UG Word", Ph("$header 창세기 1:1"), Ph("$lang1 본문"))
            }).Layouts[0];

            Assert.True(l.IsValid);
            Assert.Equal(new[] { 1, 2 }, l.Boxes.Keys.ToArray());
        }

        [Fact]
        public void Word_WithoutHeaderIsInvalid() {
            VenueLayout l = VenueRules.Scan(new[] { Layout(1, "X Word", Ph("$lang1 본문")) }).Layouts[0];

            Assert.False(l.IsValid);
            Assert.Contains("missing token(s) $1", l.Issues);
        }

        [Fact]
        public void NumberedTokensStillWork_OnRoleLayouts() {
            VenueLayout l = VenueRules.Scan(new[] {
                Layout(1, "Old LyricPage", Ph("$1 한국어\r두번째 문단"), Ph("$2 english"), Ph("$3 中文"))
            }).Layouts[0];

            Assert.True(l.IsValid);
            Assert.Empty(l.Fields);
            Assert.Equal(new[] { 1, 2, 3 }, l.Boxes.Keys.ToArray());
        }

        [Fact]
        public void MixedOrDuplicateOrForeignTokens_AreIssues() {
            VenueScan scan = VenueRules.Scan(new[] {
                Layout(1, "A LyricPage", Ph("$1 x"), Ph("$lang2 y")),
                Layout(2, "B LyricPage", Ph("$lang1 x"), Ph("$lang1 y")),
                Layout(3, "C LyricPage", Ph("$lang1 x"), Ph("$title.lang1 y")),
            });

            Assert.Contains("mixes numbered ($N) and named tokens", scan.Layouts[0].Issues);
            Assert.Contains("duplicate token $lang1", scan.Layouts[1].Issues);
            Assert.Contains("unexpected token $title.lang1", scan.Layouts[2].Issues);
            Assert.Empty(scan.Candidates(VenueRules.Praise));
        }

        [Fact]
        public void BlankLayout_IgnoresFooterButNotContent() {
            VenueScan scan = VenueRules.Scan(new[] {
                Layout(1, "A EmptyPage", Footer()),
                Layout(2, "B EmptyPage", Ph("본문")),
            });

            Assert.True(scan.Layouts[0].IsValid);
            Assert.Contains("blank layout should have no content placeholders", scan.Layouts[1].Issues);
            Assert.Equal("A EmptyPage", scan.Resolve(VenueRules.Blank).Name);
        }

        [Fact]
        public void BlankFallback_IsFirstRoleLessLayoutWithoutContent() {
            VenueScan scan = VenueRules.Scan(new[] {
                Layout(1, "Title Slide", Ph("제목")),
                Layout(2, "Blank", Footer()),
            });

            Assert.Equal("Blank", scan.Resolve(VenueRules.Blank).Name);
            Assert.Null(scan.Resolve(VenueRules.Praise));
        }

        [Fact]
        public void FillableLayouts_HaveFieldsAndKind() {
            VenueLayout l = VenueRules.Scan(new[] {
                Layout(1, "CS Order", Ph("$title.lang1 차례명"), Ph("$affiliation 소속"), Ph("$name 이름"))
            }).Layouts[0];

            Assert.Null(l.Role);
            Assert.False(l.IsValid);
            Assert.True(l.Fillable);
            Assert.Equal("order", l.Kind);
            Assert.Equal(new[] { "title.lang1", "affiliation", "name" }, l.Fields.Select(f => f.Token).ToArray());
            Assert.Equal("차례명", l.Fields[0].Label);
        }
    }

    public class VenueMatchTests {
        private static PlaceholderInfo Ph(string text, float top, float left = 0) {
            return new PlaceholderInfo { Text = text, Top = top, Left = left, Height = 50, Width = 100 };
        }

        private static LayoutDescriptor Layout(int master, int index, string name, params PlaceholderInfo[] phs) {
            return new LayoutDescriptor { MasterIndex = master, LayoutIndex = index, Name = name, Placeholders = phs };
        }

        [Fact]
        public void VenuePraise_LeavesMissingLanguagesNull() {
            List<LayoutDescriptor> layouts = new List<LayoutDescriptor> {
                Layout(1, 1, "UG EmptyPage", new PlaceholderInfo { IsLatent = true }),
                Layout(1, 2, "UG LyricPage", Ph("$lang1 가사", 10)),
            };

            LayoutMatchResult r = LayoutMatching.MatchAll(layouts, VenueRules.Praise, SlideGenDefaults.PraiseBoxCount);

            Assert.True(r.IsAvailable);
            LayoutMatch m = Assert.Single(r.BindCandidates);
            Assert.True(m.IsVenue);
            Assert.Equal(10f, m.BoxSignatures[0].Top);
            Assert.Null(m.BoxSignatures[1]);
            Assert.Null(m.BoxSignatures[2]);
            // The VENUE blank layout is found even though it carries a footer placeholder.
            Assert.Equal(1, r.EmptyLayoutIndex);
            Assert.Equal("UG EmptyPage", r.EmptyLayoutName);
        }

        [Fact]
        public void LayoutsOfEveryMasterAreScanned() {
            List<LayoutDescriptor> layouts = new List<LayoutDescriptor> {
                Layout(1, 1, "Office Title", Ph("제목", 0)),
                Layout(2, 1, "V EmptyPage"),
                Layout(2, 2, "V Word", Ph("$header h", 0), Ph("$lang1 a", 100)),
            };

            LayoutMatchResult r = LayoutMatching.MatchAll(layouts, VenueRules.Word, SlideGenDefaults.WordBoxCount);

            LayoutMatch m = Assert.Single(r.BindCandidates);
            Assert.Equal(2, m.MasterIndex);
            Assert.Equal(2, m.LayoutIndex);
            Assert.Equal(2, r.EmptyMasterIndex);
            Assert.Equal(1, r.EmptyLayoutIndex);
        }

        [Fact]
        public void LegacyMarkers_StillMatch_AndAreNotListedTwice() {
            List<LayoutDescriptor> layouts = new List<LayoutDescriptor> {
                Layout(1, 1, "Empty"),
                Layout(1, 2, "Old LyricPage", Ph("$1 a", 0), Ph("$2 b", 100), Ph("$3 c", 200)),
                Layout(1, 3, "Bind", Ph("$1 a", 0), Ph("$2 b", 100), Ph("$3 c", 200)),
            };

            LayoutMatchResult r = LayoutMatching.MatchAll(layouts, VenueRules.Praise, SlideGenDefaults.PraiseBoxCount);

            Assert.Equal(2, r.BindCandidates.Count);
            Assert.True(r.BindCandidates[0].IsVenue);   // "…LyricPage" by the VENUE rules
            Assert.False(r.BindCandidates[1].IsVenue);  // "Bind" by the legacy markers only
            Assert.Equal(1, r.EmptyLayoutIndex);        // legacy empty layout (no VENUE blank)
        }

        [Fact]
        public void SideBySideBoxes_GetDistinctSignatures() {
            // VENUE_Cheonbosan LyricPage: $lang2 and $lang3 share top and height.
            List<LayoutDescriptor> layouts = new List<LayoutDescriptor> {
                Layout(1, 1, "CbMt LyricPage", Ph("$lang1 a", 0), Ph("$lang2 b", 400, 10), Ph("$lang3 c", 400, 490)),
            };

            List<BoxSignature> sigs = LayoutMatching.MatchAll(layouts, VenueRules.Praise, 3).BindCandidates[0].BoxSignatures;

            Assert.Equal(10f, sigs[1].Left);
            Assert.Equal(490f, sigs[2].Left);
        }
    }

    /// <summary>
    /// VenueRules against /venue/inspect: each fixture holds a template's layout
    /// placeholders (as COM reports them) and the exporter's answers for them.
    /// Regenerate with Fixtures/venue/gen_fixtures.py when the templates or the
    /// exporter spec change.
    /// </summary>
    public class VenueFixtureTests {
        private static string FixtureDir {
            get { return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Fixtures", "venue"); }
        }

        public static IEnumerable<object[]> Fixtures() {
            return Directory.GetFiles(FixtureDir, "*.json").OrderBy(f => f).Select(f => new object[] { Path.GetFileName(f) });
        }

        [Fact]
        public void FixturesArePresent() {
            Assert.True(Fixtures().Count() >= 9);
        }

        [Theory]
        [MemberData(nameof(Fixtures))]
        public void MatchesExporterInspect(string file) {
            Fixture fx = Load(file);
            List<LayoutDescriptor> descriptors = Descriptors(fx);

            VenueScan scan = VenueRules.Scan(descriptors);

            for (int i = 0; i < fx.Layouts.Count; i++) {
                Expected want = fx.Layouts[i].Expected;
                VenueLayout got = scan.Layouts[i];
                string at = fx.Template + " / " + got.Name + ": ";
                Assert.True(want.Role == got.Role, at + "role " + got.Role + " != " + want.Role);
                Assert.True(want.Kind == got.Kind, at + "kind " + got.Kind + " != " + want.Kind);
                Assert.True(want.Valid == got.IsValid, at + "valid; issues: " + string.Join("; ", got.Issues));
                Assert.True(want.Fillable == got.Fillable, at + "fillable");
                Assert.True(want.IssueCount == got.Issues.Count, at + "issues: " + string.Join("; ", got.Issues));
                Assert.Equal(want.Tokens, got.Boxes.Keys.ToList());
                Assert.Equal(want.Fields, got.Fields.Select(f => f.Token).ToList());
            }
            AssertRole(fx.Template, fx.Roles.Praise, scan.Resolve(VenueRules.Praise));
            AssertRole(fx.Template, fx.Roles.Word, scan.Resolve(VenueRules.Word));
            AssertRole(fx.Template, fx.Roles.Blank, scan.Resolve(VenueRules.Blank));
        }

        [Theory]
        [MemberData(nameof(Fixtures))]
        public void PraiseAndWordBoxesAreDistinguishable(string file) {
            Fixture fx = Load(file);
            List<LayoutDescriptor> descriptors = Descriptors(fx);

            foreach (string role in new[] { VenueRules.Praise, VenueRules.Word }) {
                LayoutMatchResult r = LayoutMatching.MatchAll(descriptors, role, role == VenueRules.Praise ? 3 : 4);
                Assert.True(r.IsAvailable, fx.Template + ": no " + role + " layout");
                foreach (LayoutMatch m in r.BindCandidates) {
                    List<BoxSignature> sigs = m.BoxSignatures.Where(s => s != null).ToList();
                    int distinct = sigs.Select(s => s.Top + "|" + s.Height + "|" + s.Left + "|" + s.Width).Distinct().Count();
                    Assert.True(distinct == sigs.Count, fx.Template + " / " + m.Name + ": two boxes share a geometry");
                }
            }
        }

        // The fixture's placeholders as the Interop layer would report them (1-based indexes).
        private static List<LayoutDescriptor> Descriptors(Fixture fx) {
            return fx.Layouts.Select(l => new LayoutDescriptor {
                MasterIndex = l.Master + 1,
                LayoutIndex = l.Index + 1,
                Name = l.Name,
                Placeholders = l.Placeholders.Select(p => new PlaceholderInfo {
                    Text = p.Text, Top = p.Top, Height = p.Height, Left = p.Left, Width = p.Width,
                    IsLatent = p.Latent, HasTextFrame = p.HasTextFrame
                }).ToList()
            }).ToList();
        }

        private static void AssertRole(string template, string expected, VenueLayout got) {
            Assert.True(expected == (got == null ? null : got.Name),
                template + ": role layout " + (got == null ? "null" : got.Name) + " != " + expected);
        }

        private static Fixture Load(string file) {
            byte[] bytes = File.ReadAllBytes(Path.Combine(FixtureDir, file));
            DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(Fixture));
            using (MemoryStream ms = new MemoryStream(bytes)) {
                return (Fixture)ser.ReadObject(ms);
            }
        }

        [DataContract]
        public class Fixture {
            [DataMember(Name = "template")] public string Template { get; set; }
            [DataMember(Name = "roles")] public Roles Roles { get; set; }
            [DataMember(Name = "layouts")] public List<FixtureLayout> Layouts { get; set; }
        }

        [DataContract]
        public class Roles {
            [DataMember(Name = "praise")] public string Praise { get; set; }
            [DataMember(Name = "word")] public string Word { get; set; }
            [DataMember(Name = "blank")] public string Blank { get; set; }
        }

        [DataContract]
        public class FixtureLayout {
            [DataMember(Name = "name")] public string Name { get; set; }
            [DataMember(Name = "master")] public int Master { get; set; }
            [DataMember(Name = "index")] public int Index { get; set; }
            [DataMember(Name = "placeholders")] public List<FixturePlaceholder> Placeholders { get; set; }
            [DataMember(Name = "expected")] public Expected Expected { get; set; }
        }

        [DataContract]
        public class FixturePlaceholder {
            [DataMember(Name = "text")] public string Text { get; set; }
            [DataMember(Name = "top")] public float Top { get; set; }
            [DataMember(Name = "height")] public float Height { get; set; }
            [DataMember(Name = "left")] public float Left { get; set; }
            [DataMember(Name = "width")] public float Width { get; set; }
            [DataMember(Name = "latent")] public bool Latent { get; set; }
            [DataMember(Name = "hasTextFrame")] public bool HasTextFrame { get; set; }
        }

        [DataContract]
        public class Expected {
            [DataMember(Name = "role")] public string Role { get; set; }
            [DataMember(Name = "kind")] public string Kind { get; set; }
            [DataMember(Name = "valid")] public bool Valid { get; set; }
            [DataMember(Name = "fillable")] public bool Fillable { get; set; }
            [DataMember(Name = "tokens")] public List<int> Tokens { get; set; }
            [DataMember(Name = "fields")] public List<string> Fields { get; set; }
            [DataMember(Name = "issueCount")] public int IssueCount { get; set; }
        }
    }
}
