using System.Collections.Generic;
using Xunit;
using ZebulonVSTO.Slides;

namespace ZebulonVSTO.Tests {
    public class TextTransformsTests {
        [Fact]
        public void Box0_SingleLine_UpperedAndNewlinePrepended() {
            Assert.Equal("\nABC", TextTransforms.TransformPraiseBox(0, "abc"));
        }

        [Fact]
        public void Box0_MultiLine_NoExtraNewline() {
            Assert.Equal("AB\nCD", TextTransforms.TransformPraiseBox(0, "ab\ncd"));
        }

        [Fact]
        public void Box0_Korean_GetsNewlinePrefix() {
            // No case change for Hangul, but the leading-newline rule still applies.
            Assert.Equal("\n주 예수", TextTransforms.TransformPraiseBox(0, "주 예수"));
        }

        [Fact]
        public void Box1_Uppered_NoNewline() {
            Assert.Equal("HELLO", TextTransforms.TransformPraiseBox(1, "hello"));
        }

        [Fact]
        public void Box2_Unchanged() {
            Assert.Equal("ni hao", TextTransforms.TransformPraiseBox(2, "ni hao"));
        }

        [Fact]
        public void Null_BecomesEmpty() {
            Assert.Equal("", TextTransforms.TransformPraiseBox(2, null));
        }
    }

    public class LayoutMatchingTests {
        private static PlaceholderInfo Ph(string text, float h = 10, float t = 20) {
            return new PlaceholderInfo { Text = text, Height = h, Top = t };
        }

        private static LayoutDescriptor Layout(int index, string name, params PlaceholderInfo[] phs) {
            return new LayoutDescriptor { LayoutIndex = index, Name = name, Placeholders = phs };
        }

        [Fact]
        public void FindsBindAndEmpty_ForPraise() {
            List<LayoutDescriptor> layouts = new List<LayoutDescriptor> {
                Layout(1, "Empty"), // no placeholders
                Layout(2, "Bind",
                    Ph("$1 한국어", 100, 10),
                    Ph("$2 english", 100, 120),
                    Ph("$3 中文", 100, 230)),
            };

            LayoutMatchResult r = LayoutMatching.Match(layouts, 3);

            Assert.True(r.IsAvailable);
            Assert.Equal(1, r.EmptyLayoutIndex);
            Assert.Single(r.BindCandidates);
            LayoutMatch m = r.BindCandidates[0];
            Assert.Equal(2, m.LayoutIndex);
            Assert.Equal(10f, m.BoxSignatures[0].Top);
            Assert.Equal(120f, m.BoxSignatures[1].Top);
            Assert.Equal(230f, m.BoxSignatures[2].Top);
        }

        [Fact]
        public void NoBind_WhenAMarkerMissing() {
            List<LayoutDescriptor> layouts = new List<LayoutDescriptor> {
                Layout(1, "Empty"),
                Layout(2, "Partial", Ph("$1 a"), Ph("$2 b")), // missing $3
            };

            LayoutMatchResult r = LayoutMatching.Match(layouts, 3);

            Assert.False(r.IsAvailable);
            Assert.Empty(r.BindCandidates);
        }

        [Fact]
        public void NotAvailable_WhenNoEmptyLayout() {
            List<LayoutDescriptor> layouts = new List<LayoutDescriptor> {
                Layout(1, "Bind", Ph("$1 a"), Ph("$2 b"), Ph("$3 c")),
            };

            LayoutMatchResult r = LayoutMatching.Match(layouts, 3);

            Assert.Single(r.BindCandidates);
            Assert.False(r.HasEmpty);
            Assert.False(r.IsAvailable);
        }

        [Fact]
        public void WordLayout_AlsoCountsAsPraise_ButPraiseLayoutDoesNotCountAsWord() {
            LayoutDescriptor threeBox = Layout(2, "Praise",
                Ph("$1 a"), Ph("$2 b"), Ph("$3 c"));
            LayoutDescriptor fourBox = Layout(3, "Word",
                Ph("$1 a"), Ph("$2 b"), Ph("$3 c"), Ph("$4 d"));
            List<LayoutDescriptor> layouts = new List<LayoutDescriptor> { Layout(1, "Empty"), threeBox, fourBox };

            // boxCount 3: both qualify (the 4-box layout still carries $1..$3).
            Assert.Equal(2, LayoutMatching.Match(layouts, 3).BindCandidates.Count);
            // boxCount 4: only the 4-box layout qualifies.
            List<LayoutMatch> word = LayoutMatching.Match(layouts, 4).BindCandidates;
            Assert.Single(word);
            Assert.Equal(3, word[0].LayoutIndex);
        }

        [Fact]
        public void LastEmptyLayoutWins() {
            List<LayoutDescriptor> layouts = new List<LayoutDescriptor> {
                Layout(1, "Empty A"),
                Layout(4, "Empty B"),
            };

            Assert.Equal(4, LayoutMatching.Match(layouts, 3).EmptyLayoutIndex);
        }
    }

    public class PraisePlannerTests {
        [Fact]
        public void BuildsExpectedSequence() {
            Lyric lyric = new Lyric {
                Title = "Amazing",
                Groups = new List<LyricGroup> {
                    new LyricGroup {
                        Name = "Verse 1",
                        Pages = new List<List<string>> {
                            new List<string> { "은혜", "grace", "恩典" },
                        }
                    }
                }
            };

            List<SlidePlanItem> plan = PraisePlanner.BuildPlan(new List<Lyric> { lyric });

            // empty separator, title, one page, empty separator
            Assert.Equal(4, plan.Count);

            Assert.Equal(LayoutKind.Empty, plan[0].Kind);
            Assert.Equal("1. Amazing", plan[0].Note);

            Assert.Equal(LayoutKind.Bind, plan[1].Kind);
            Assert.Equal("Amazing", plan[1].BoxText[0]); // raw title, not transformed
            Assert.Equal(" ", plan[1].BoxText[1]);
            Assert.Equal(" ", plan[1].BoxText[2]);
            Assert.Equal("Amazing - Title", plan[1].Note);

            Assert.Equal(LayoutKind.Bind, plan[2].Kind);
            Assert.Equal("\n은혜", plan[2].BoxText[0]); // box0 single-line → leading newline
            Assert.Equal("GRACE", plan[2].BoxText[1]);  // box1 uppercased
            Assert.Equal("恩典", plan[2].BoxText[2]);    // box2 unchanged
            Assert.Equal("Amazing - Verse 1 (1/1)", plan[2].Note);

            Assert.Equal(LayoutKind.Empty, plan[3].Kind);
            Assert.Equal("Amazing - End", plan[3].Note);
        }

        [Fact]
        public void NumbersMultipleItemsAndPages() {
            Lyric a = new Lyric {
                Title = "A",
                Groups = new List<LyricGroup> {
                    new LyricGroup { Name = "V", Pages = new List<List<string>> {
                        new List<string> { "x", "y", "z" },
                        new List<string> { "p", "q", "r" },
                    } }
                }
            };
            Lyric b = new Lyric {
                Title = "B",
                Groups = new List<LyricGroup> {
                    new LyricGroup { Name = "C", Pages = new List<List<string>> {
                        new List<string> { "1", "2", "3" },
                    } }
                }
            };

            List<SlidePlanItem> plan = PraisePlanner.BuildPlan(new List<Lyric> { a, b });

            // A: empty + title + 2 pages + empty = 5; B: empty + title + 1 page + empty = 4
            Assert.Equal(9, plan.Count);
            Assert.Equal("1. A", plan[0].Note);
            Assert.Equal("A - V (2/2)", plan[3].Note);
            Assert.Equal("2. B", plan[5].Note);
        }

        [Fact]
        public void SkipsInvalidItem() {
            Lyric bad = new Lyric { Title = "", Groups = null };
            Assert.Empty(PraisePlanner.BuildPlan(new List<Lyric> { bad }));
        }
    }

    public class LyricJsonTests {
        [Fact]
        public void ParsesNestedPages() {
            string json = "{\"title\":\"Alive\",\"groups\":[{\"name\":\"Verse 1\",\"pages\":[[\"가사\",\"lyric\",\"歌词\"],[\"둘째\",\"second\",\"第二\"]]}]}";

            Lyric lyric = LyricJson.ParseLyric(json);

            Assert.NotNull(lyric);
            Assert.Equal("Alive", lyric.Title);
            Assert.Single(lyric.Groups);
            Assert.Equal("Verse 1", lyric.Groups[0].Name);
            Assert.Equal(2, lyric.Groups[0].Pages.Count);
            Assert.Equal("가사", lyric.Groups[0].Pages[0][0]);
            Assert.Equal("第二", lyric.Groups[0].Pages[1][2]);
        }

        [Fact]
        public void ParsesStringArray() {
            List<string> list = LyricJson.ParseStringArray("[\"lyrics/A/x.md\",\"lyrics/B/y.md\"]");

            Assert.Equal(2, list.Count);
            Assert.Equal("lyrics/A/x.md", list[0]);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("not json")]
        [InlineData("{ broken")]
        public void ReturnsNullOnBadInput(string input) {
            Assert.Null(LyricJson.ParseLyric(input));
        }
    }
}
