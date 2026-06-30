using System.Collections.Generic;
using Xunit;
using ZebulonVSTO.Slides;

namespace ZebulonVSTO.Tests {
    public class SectionUtilsTests {
        private static WordSection S(int start, int end) {
            return new WordSection(start, end);
        }

        [Fact]
        public void Coalesce_MergesOverlappingAndAdjacent() {
            List<WordSection> r = SectionUtils.Coalesce(new List<WordSection> { S(18, 20), S(19, 22), S(16, 16) });
            Assert.Equal(2, r.Count);
            Assert.Equal(16, r[0].Start);
            Assert.Equal(16, r[0].End);
            Assert.Equal(18, r[1].Start);
            Assert.Equal(22, r[1].End);
        }

        [Fact]
        public void Coalesce_MergesAdjacent() {
            List<WordSection> r = SectionUtils.Coalesce(new List<WordSection> { S(1, 3), S(4, 5) });
            Assert.Single(r);
            Assert.Equal(1, r[0].Start);
            Assert.Equal(5, r[0].End);
        }

        [Fact]
        public void Coalesce_DropsNullAndInvertedRanges() {
            // Faithful to the web coalesceSections: filter(r => r && r.end >= r.start) — no swap.
            List<WordSection> r = SectionUtils.Coalesce(new List<WordSection> { null, S(8, 5), S(2, 4) });
            Assert.Single(r);
            Assert.Equal(2, r[0].Start);
            Assert.Equal(4, r[0].End);
        }

        [Fact]
        public void FormatSections_RendersSinglesAndRanges() {
            Assert.Equal("16, 18-20", SectionUtils.FormatSections(new List<WordSection> { S(16, 16), S(18, 20) }));
        }
    }

    public class VerseSelectionTests {
        [Fact]
        public void FirstTap_StartsSinglePending() {
            VerseSelection sel = new VerseSelection();
            sel.Click(5);
            Assert.True(sel.HasPending);
            Assert.Equal(5, sel.Pending.Start);
            Assert.Equal(5, sel.Pending.End);
            Assert.Empty(sel.Segments);
        }

        [Fact]
        public void TapLaterVerse_ExtendsPendingIntoRange() {
            VerseSelection sel = new VerseSelection();
            sel.Click(5);
            sel.Click(8);
            Assert.Equal(5, sel.Pending.Start);
            Assert.Equal(8, sel.Pending.End);
        }

        [Fact]
        public void TapEarlierVerse_RestartsPending() {
            VerseSelection sel = new VerseSelection();
            sel.Click(5);
            sel.Click(3); // not greater than start → restart
            Assert.Equal(3, sel.Pending.Start);
            Assert.Equal(3, sel.Pending.End);
        }

        [Fact]
        public void TapAfterRange_RestartsPending() {
            VerseSelection sel = new VerseSelection();
            sel.Click(5);
            sel.Click(8);  // pending is a range now
            sel.Click(10); // any tap on a range restarts as a single
            Assert.Equal(10, sel.Pending.Start);
            Assert.Equal(10, sel.Pending.End);
        }

        [Fact]
        public void Commit_MovesPendingIntoSegments() {
            VerseSelection sel = new VerseSelection();
            sel.Click(5);
            sel.Click(8);
            sel.CommitPending();
            Assert.False(sel.HasPending);
            Assert.Single(sel.Segments);
            Assert.Equal(5, sel.Segments[0].Start);
            Assert.Equal(8, sel.Segments[0].End);
        }

        [Fact]
        public void Commit_CoalescesOverlappingAndAdjacentRanges() {
            VerseSelection sel = new VerseSelection();
            sel.Click(5); sel.Click(8); sel.CommitPending(); // 5-8
            sel.Click(16); sel.CommitPending();               // 16
            sel.Click(9); sel.CommitPending();                // 9 adjacent to 8 → merge into 5-9
            Assert.Equal(2, sel.Segments.Count);
            Assert.Equal(5, sel.Segments[0].Start);
            Assert.Equal(9, sel.Segments[0].End);
            Assert.Equal(16, sel.Segments[1].Start);
            Assert.Equal(16, sel.Segments[1].End);
        }

        [Fact]
        public void RemoveSegment_DropsByIndex() {
            VerseSelection sel = new VerseSelection();
            sel.Click(1); sel.CommitPending();
            sel.Click(10); sel.CommitPending();
            sel.RemoveSegment(0);
            Assert.Single(sel.Segments);
            Assert.Equal(10, sel.Segments[0].Start);
        }

        [Fact]
        public void DisplayRanges_IncludePendingAndCoalesce() {
            VerseSelection sel = new VerseSelection();
            sel.Click(5); sel.Click(8); sel.CommitPending(); // segment 5-8
            sel.Click(9);                                     // pending 9 (adjacent)
            List<WordSection> ranges = sel.DisplayRanges();
            Assert.Single(ranges);
            Assert.Equal(5, ranges[0].Start);
            Assert.Equal(9, ranges[0].End);
        }

        [Fact]
        public void RoleOf_ClassifiesEndpointsAndMiddle() {
            VerseSelection sel = new VerseSelection();
            sel.Click(5); sel.Click(8); sel.CommitPending(); // 5-8
            sel.Click(12); sel.CommitPending();              // 12 (single)
            Assert.Equal("", sel.RoleOf(4));
            Assert.Equal("start", sel.RoleOf(5));
            Assert.Equal("middle", sel.RoleOf(6));
            Assert.Equal("end", sel.RoleOf(8));
            Assert.Equal("single", sel.RoleOf(12));
        }

        [Fact]
        public void Reset_ClearsEverything() {
            VerseSelection sel = new VerseSelection();
            sel.Click(5); sel.CommitPending();
            sel.Click(9);
            sel.Reset();
            Assert.Empty(sel.Segments);
            Assert.False(sel.HasPending);
            Assert.Empty(sel.DisplayRanges());
        }
    }

    public class WordPlannerTests {
        private static WordItem Item(string head, List<WordSection> sections,
            List<List<string>> first, List<List<string>> second, List<List<string>> third) {
            return new WordItem { HeadText = head, Sections = sections, First = first, Second = second, Third = third };
        }

        [Fact]
        public void BuildsEmptySeparatorThenOnePerVerse() {
            WordItem item = Item("Acts 19:30-31",
                new List<WordSection> { new WordSection(30, 31) },
                new List<List<string>> { new List<string> { "k30", "k31" } },
                new List<List<string>> { new List<string> { "e30", "e31" } },
                new List<List<string>> { new List<string> { "c30", "c31" } });

            List<SlidePlanItem> plan = WordPlanner.BuildPlan(new List<WordItem> { item });

            Assert.Equal(3, plan.Count);
            Assert.Equal(LayoutKind.Empty, plan[0].Kind);
            Assert.Null(plan[0].Note);

            Assert.Equal(LayoutKind.Bind, plan[1].Kind);
            Assert.Equal("Acts 19:30-31", plan[1].BoxText[0]);
            Assert.Equal("30. k30", plan[1].BoxText[1]);
            Assert.Equal("30. e30", plan[1].BoxText[2]);
            Assert.Equal("30. c30", plan[1].BoxText[3]);

            Assert.Equal("31. k31", plan[2].BoxText[1]);
            Assert.Equal("31. c31", plan[2].BoxText[3]);
        }

        [Fact]
        public void VerseNumbersFollowEachSectionStart() {
            WordItem item = Item("H",
                new List<WordSection> { new WordSection(1, 1), new WordSection(3, 3) },
                new List<List<string>> { new List<string> { "a" }, new List<string> { "c" } },
                new List<List<string>>(), new List<List<string>>());

            List<SlidePlanItem> plan = WordPlanner.BuildPlan(new List<WordItem> { item });

            Assert.Equal(3, plan.Count); // empty + 2 binds
            Assert.Equal("1. a", plan[1].BoxText[1]);
            Assert.Equal("3. c", plan[2].BoxText[1]);
            // empty languages omit their box entirely
            Assert.False(plan[1].BoxText.ContainsKey(2));
            Assert.False(plan[1].BoxText.ContainsKey(3));
        }

        [Fact]
        public void AllEmptyLineBreaksTheSection() {
            // section 1-3 but the first slot has an empty middle verse; other slots empty.
            WordItem item = Item("H",
                new List<WordSection> { new WordSection(1, 3) },
                new List<List<string>> { new List<string> { "a", "", "c" } },
                new List<List<string>>(), new List<List<string>>());

            List<SlidePlanItem> plan = WordPlanner.BuildPlan(new List<WordItem> { item });

            // verse 1 emits; verse 2 is empty in all langs → break (verse 3 never reached)
            Assert.Equal(2, plan.Count);
            Assert.Equal("1. a", plan[1].BoxText[1]);
        }

        [Fact]
        public void SkipsInvalidItem() {
            WordItem bad = Item("", new List<WordSection> { new WordSection(1, 1) },
                new List<List<string>>(), new List<List<string>>(), new List<List<string>>());
            Assert.Empty(WordPlanner.BuildPlan(new List<WordItem> { bad }));
        }
    }

    public class WordAssemblerTests {
        private static BibleData Chapter(string bookName, int verseCount) {
            List<string> data = new List<string>();
            for (int i = 1; i <= verseCount; i++) {
                data.Add("v" + i);
            }
            return new BibleData { BookName = bookName, Chapter = 19, Data = data };
        }

        [Fact]
        public void SlicesVersesAndComposesHeadText() {
            List<BibleData> chunk = new List<BibleData> {
                Chapter("사도행전", 31),
                Chapter("Acts", 31),
            };
            WordItem item = WordAssembler.Build(chunk, 19,
                new List<WordSection> { new WordSection(30, 31) },
                new List<RubyMode> { RubyMode.Base, RubyMode.Base });

            Assert.Equal("사도행전 (Acts) 19:30-31", item.HeadText);
            Assert.Single(item.First);
            Assert.Equal(new List<string> { "v30", "v31" }, item.First[0]);
            Assert.Equal(new List<string> { "v30", "v31" }, item.Second[0]);
            Assert.Empty(item.Third); // slot 2 absent
        }

        [Fact]
        public void SingleVersionHeadTextHasNoParens() {
            List<BibleData> chunk = new List<BibleData> { Chapter("사도행전", 31) };
            WordItem item = WordAssembler.Build(chunk, 19,
                new List<WordSection> { new WordSection(30, 30) },
                new List<RubyMode> { RubyMode.Base });
            Assert.Equal("사도행전 19:30", item.HeadText);
        }

        [Fact]
        public void RubySectionResolvedToBaseText() {
            List<BibleData> chunk = new List<BibleData> {
                new BibleData { BookName = "b", Chapter = 1, Data = new List<string> { "<ruby>神<rt>かみ</rt></ruby>を愛す" } }
            };
            WordItem item = WordAssembler.Build(chunk, 1,
                new List<WordSection> { new WordSection(1, 1) },
                new List<RubyMode> { RubyMode.Base });
            Assert.Equal("神を愛す", item.First[0][0]);
        }
    }

    public class BibleDataTests {
        [Fact]
        public void Parse_ReadsBooknameChapterData() {
            BibleData d = BibleData.Parse("{\"bookname\":\"창세기\",\"chapter\":1,\"data\":[\"태초에\",\"땅이\"]}");
            Assert.NotNull(d);
            Assert.Equal("창세기", d.BookName);
            Assert.Equal(1, d.Chapter);
            Assert.Equal(2, d.Data.Count);
            Assert.Equal("태초에", d.Data[0]);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("not json")]
        public void Parse_ReturnsNullOnBadInput(string input) {
            Assert.Null(BibleData.Parse(input));
        }
    }

    public class RubyTextTests {
        [Fact]
        public void Base_KeepsKanjiDropsReading() {
            Assert.Equal("神を", RubyText.ApplyRubyMode("<ruby>神<rt>かみ</rt></ruby>を", RubyMode.Base));
        }

        [Fact]
        public void Both_InsertsReadingInParens() {
            Assert.Equal("神(かみ)を", RubyText.ApplyRubyMode("<ruby>神<rt>かみ</rt></ruby>を", RubyMode.Both));
        }

        [Fact]
        public void Furigana_TakesReadingAndKatakanaToHiragana() {
            Assert.Equal("かみ", RubyText.ApplyRubyMode("<ruby>神<rt>カミ</rt></ruby>", RubyMode.Furigana));
        }

        [Fact]
        public void ToPlainText_StripsTagsAndDecodesEntities() {
            Assert.Equal("hi&x", RubyText.ToPlainText("<b>hi</b>&amp;x"));
        }
    }

    public class BibleCatalogTests {
        [Fact]
        public void HasAllVersionsAndBooks() {
            Assert.Equal(18, BibleCatalog.AllVersions.Count);
            Assert.Equal(66, BibleCatalog.BookCount);
        }

        [Fact]
        public void DefaultVersionPerLanguage() {
            Assert.Equal("KRTRV", BibleCatalog.VersionsForLanguage("ko-KR")[0].Code);
            Assert.Equal("ESV", BibleCatalog.VersionsForLanguage("en-US")[0].Code); // ALL orders EN as ESV-first
        }

        [Fact]
        public void ParseVersionIsCaseInsensitive_AndJdbHasRuby() {
            BibleVersion jdb = BibleCatalog.ParseVersion("jdb");
            Assert.NotNull(jdb);
            Assert.Equal("JDB", jdb.Code);
            Assert.True(jdb.HasRuby);
        }

        [Fact]
        public void ChapterCount_Genesis() {
            Assert.Equal(50, BibleCatalog.ChapterCount(1));
        }
    }
}
