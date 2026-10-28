using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using Xunit;
using ZebulonVSTO.Slides;

namespace ZebulonVSTO.Tests {
    public class VenueHeaderTests {
        [Theory]
        [InlineData(new[] { "사도행전", "Acts", "使徒行传" }, "19:30-41", "사도행전 (Acts, 使徒行传) 19:30-41")]
        [InlineData(new[] { "사도행전", "", "使徒行传" }, "19:30", "사도행전 (使徒行传) 19:30")]
        [InlineData(new[] { "사도행전" }, "19:30", "사도행전 19:30")]
        [InlineData(new[] { "사도행전", "Acts" }, null, "사도행전 (Acts)")]
        public void FormatsLikeTheExporter(string[] books, string reference, string expected) {
            Assert.Equal(expected, new VenueHeader { Books = books.ToList(), Ref = reference }.Format());
        }

        [Fact]
        public void WordAssembler_HeaderFormatEqualsHeadText() {
            List<BibleData> chunk = new List<BibleData> {
                new BibleData { BookName = "창세기", Data = new List<string> { "태초에", "땅이" } },
                new BibleData { BookName = "Genesis", Data = new List<string> { "In the beginning", "Now" } },
                new BibleData { BookName = "창세기", Data = new List<string> { "a", "b" } },
            };
            WordItem item = WordAssembler.Build(chunk, 1, new List<WordSection> { new WordSection(1, 2) }, null);

            Assert.Equal(new[] { "창세기", "Genesis" }, item.Header.Books.ToArray());
            Assert.Equal("1:1-2", item.Header.Ref);
            Assert.Equal(item.HeadText, item.Header.Format());
        }
    }

    public class VenuePayloadTests {
        [Fact]
        public void Praise_WritesOnlyNonDefaultSwitches() {
            Lyric lyric = new Lyric {
                Title = "Alive",
                Groups = new List<LyricGroup> { new LyricGroup { Name = "V", Pages = new List<List<string>> { new List<string> { "a" } } } }
            };
            string json = new VenuePayload { Items = new List<VenueItem> {
                VenueItem.Praise(lyric), VenueItem.Praise(lyric, titleSlide: false) } }.ToJson();

            Assert.Contains("{\"type\":\"praise\",\"data\":{", json);
            Assert.Contains("\"title\":\"Alive\"", json);
            Assert.Equal(1, json.Split(new[] { "\"titleSlide\":false" }, StringSplitOptions.None).Length - 1);
            Assert.DoesNotContain("separators", json);
            Assert.DoesNotContain("options", json);
        }

        [Fact]
        public void Word_PrefersTheStructuredHeader() {
            WordItem w = new WordItem {
                HeadText = "창세기 1:1",
                Header = new VenueHeader { Books = new List<string> { "창세기" }, Ref = "1:1" },
                Sections = new List<WordSection> { new WordSection(1, 1) },
                First = new List<List<string>> { new List<string> { "태초에" } }
            };
            string json = new VenuePayload { Items = new List<VenueItem> { VenueItem.Word(w) } }.ToJson();

            Assert.Contains("\"header\":{\"books\":[\"창세기\"],\"ref\":\"1:1\"}", json);
            Assert.DoesNotContain("headText", json);
            Assert.Contains("\"sections\":[{\"end\":1,\"start\":1}]", json);
        }

        [Fact]
        public void ParsesEveryItemType() {
            VenuePayload p = VenuePayload.FromJson(
                "{\"options\":{\"fixedSlides\":\"items\",\"fontNames\":[\"A\",null]},\"items\":[" +
                "{\"type\":\"fixed\",\"data\":{\"key\":\"testimony\"}}," +
                "{\"type\":\"layout\",\"layout\":\"CS Prayer\",\"data\":{\"fields\":{\"subject.lang1\":\"1. x\"},\"notes\":\"n\"}}," +
                "{\"type\":\"blank\"}]}");

            Assert.Equal("items", p.Options.FixedSlides);
            Assert.Equal(new[] { "A", null }, p.Options.FontNames.ToArray());
            Assert.Equal("testimony", p.Items[0].Data.Key);
            Assert.Equal("CS Prayer", p.Items[1].Layout);
            Assert.Equal("n", p.Items[1].Data.Notes);
            Assert.Null(p.Items[2].Data);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("not json")]
        public void FromJson_BadInputIsNull(string json) {
            Assert.Null(VenuePayload.FromJson(json));
        }
    }

    public class VenuePlannerTests {
        private static VenueItem Song(string title, bool titleSlide = true, bool separators = true) {
            return VenueItem.Praise(new Lyric {
                Title = title,
                Groups = new List<LyricGroup> { new LyricGroup { Name = "V", Pages = new List<List<string>> { new List<string> { "a" } } } }
            }, titleSlide, separators);
        }

        [Fact]
        public void DefaultSwitches_MatchTheLegacyPraisePlan() {
            Lyric lyric = Song("Alive").Data.ToLyric();
            List<SlidePlanItem> legacy = PraisePlanner.BuildPlan(new[] { lyric });
            List<SlidePlanItem> venue = VenuePlanner.BuildPlan(new[] { Song("Alive") }).Slides;

            Assert.Equal(legacy.Select(s => s.Kind + "|" + s.Note), venue.Select(s => s.Kind + "|" + s.Note));
        }

        [Fact]
        public void SongsAreNumberedAcrossTheWholePayload_SkippingInvalidOnes() {
            VenuePlan plan = VenuePlanner.BuildPlan(new[] {
                Song("A"), VenueItem.Blank("x"), new VenueItem { Type = "praise", Data = new VenueItemData { Title = "no groups" } },
                Song("B", separators: false)
            });

            Assert.Equal("1. A", plan.Slides[0].Note);
            Assert.Equal("2. B\nB - Title", plan.Slides.Single(s => s.Note != null && s.Note.StartsWith("2. ")).Note);
            Assert.Equal(new[] { 2 }, plan.Skipped.ToArray());
        }

        [Fact]
        public void FixedLayoutAndUnknownItems_AreSkipped() {
            VenuePlan plan = VenuePlanner.BuildPlan(new[] {
                new VenueItem { Type = "fixed", Data = new VenueItemData { Key = "k" } },
                new VenueItem { Type = "layout", Layout = "CS Order" },
                new VenueItem { Type = "poster" },
                null,
                VenueItem.Blank()
            });

            Assert.Single(plan.Slides);
            Assert.Equal(LayoutKind.Empty, plan.Slides[0].Kind);
            Assert.Null(plan.Slides[0].Note);
            Assert.Equal(new[] { 0, 1, 2, 3 }, plan.Skipped.ToArray());
        }
    }

    /// <summary>
    /// VenuePlanner against /venue/export: Fixtures/venue/items/payload.json was
    /// exported on several templates, and each generated slide's role, notes and
    /// box text recorded. The plan must produce the same slides; a box the plan
    /// leaves unset is " " (LayoutSelection.FillUnusedBoxes), and boxes the
    /// template lacks are not compared (the Interop layer drops them).
    /// </summary>
    public class VenuePlannerFixtureTests {
        private static string ItemsDir {
            get { return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Fixtures", "venue", "items"); }
        }

        public static IEnumerable<object[]> Exports() {
            return Directory.GetFiles(ItemsDir, "*.slides.json").OrderBy(f => f).Select(f => new object[] { Path.GetFileName(f) });
        }

        [Fact]
        public void ExportsArePresent() {
            Assert.True(Exports().Count() >= 4);
        }

        [Theory]
        [MemberData(nameof(Exports))]
        public void PlansTheSameSlidesAsTheExporter(string file) {
            VenuePayload payload = VenuePayload.FromJson(File.ReadAllText(Path.Combine(ItemsDir, "payload.json")));
            Export export = Load(file);

            VenuePlan plan = VenuePlanner.BuildPlan(payload.Items);

            Assert.Empty(plan.Skipped);
            Assert.Equal(export.Slides.Count, plan.Slides.Count);
            for (int i = 0; i < plan.Slides.Count; i++) {
                ExportSlide want = export.Slides[i];
                SlidePlanItem got = plan.Slides[i];
                string at = export.Template + " slide " + (i + 1) + ": ";
                Assert.True((want.Role == "blank") == (got.Kind == LayoutKind.Empty), at + "kind " + got.Kind + " vs role " + want.Role);
                Assert.True(want.Notes == (got.Note ?? ""), at + "notes '" + got.Note + "' != '" + want.Notes + "'");
                foreach (KeyValuePair<string, string> box in want.Texts ?? new Dictionary<string, string>()) {
                    string text;
                    if (got.BoxText == null || !got.BoxText.TryGetValue(int.Parse(box.Key) - 1, out text)) {
                        text = " ";
                    }
                    Assert.True(box.Value == text, at + "$" + box.Key + " '" + text + "' != '" + box.Value + "'");
                }
            }
        }

        private static Export Load(string file) {
            DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(Export),
                new DataContractJsonSerializerSettings { UseSimpleDictionaryFormat = true });
            using (FileStream fs = File.OpenRead(Path.Combine(ItemsDir, file))) {
                return (Export)ser.ReadObject(fs);
            }
        }

        [DataContract]
        public class Export {
            [DataMember(Name = "template")] public string Template { get; set; }
            [DataMember(Name = "slides")] public List<ExportSlide> Slides { get; set; }
        }

        [DataContract]
        public class ExportSlide {
            [DataMember(Name = "role")] public string Role { get; set; }
            [DataMember(Name = "notes")] public string Notes { get; set; }
            [DataMember(Name = "texts")] public Dictionary<string, string> Texts { get; set; }
        }
    }
}
