using System;
using System.Collections.Generic;
using System.IO;
using Xunit;
using ZebulonVSTO.Slides;

namespace ZebulonVSTO.Tests {
    public class WordSelectPrefsTests {
        private static WordSlotPref Slot(string lang, string version, int ruby = 0) {
            return new WordSlotPref { Language = lang, Version = version, Ruby = ruby };
        }

        [Fact]
        public void Defaults_AreThreeLanguagesKoEnZh() {
            WordSelectPrefs p = WordSelectPrefs.Defaults();
            Assert.Equal(3, p.LangCount);
            Assert.Equal(new[] { "ko-KR", "en-US", "zh-CN" }, p.Slots.ConvertAll(s => s.Language));
            Assert.All(p.Slots, s => { Assert.Null(s.Version); Assert.Equal(0, s.Ruby); });
        }

        [Theory]
        [InlineData(0, 1)]
        [InlineData(2, 2)]
        [InlineData(7, 3)]
        public void Normalized_ClampsLangCount(int stored, int expected) {
            WordSelectPrefs p = new WordSelectPrefs { LangCount = stored }.Normalized();
            Assert.Equal(expected, p.LangCount);
        }

        [Fact]
        public void Normalized_PadsMissingSlotsWithDefaults() {
            WordSelectPrefs p = new WordSelectPrefs {
                LangCount = 1,
                Slots = new List<WordSlotPref> { Slot("ja-JP", "JDB", 1) }
            }.Normalized();
            Assert.Equal(3, p.Slots.Count);
            Assert.Equal("ja-JP", p.Slots[0].Language);
            Assert.Equal("JDB", p.Slots[0].Version);
            Assert.Equal(1, p.Slots[0].Ruby);
            Assert.Equal("en-US", p.Slots[1].Language);
            Assert.Equal("zh-CN", p.Slots[2].Language);
        }

        [Fact]
        public void Normalized_UnknownLanguageResetsThatSlot() {
            WordSelectPrefs p = new WordSelectPrefs {
                LangCount = 3,
                Slots = new List<WordSlotPref> { Slot("ko-KR", "KNT"), Slot("xx-XX", "ESV", 2), Slot(null, null) }
            }.Normalized();
            Assert.Equal("KNT", p.Slots[0].Version);
            Assert.Equal("en-US", p.Slots[1].Language);
            Assert.Null(p.Slots[1].Version);
            Assert.Equal(0, p.Slots[1].Ruby);
            Assert.Equal("zh-CN", p.Slots[2].Language);
        }

        [Fact]
        public void Normalized_DropsVersionOfAnotherLanguage_AndCanonicalizesCase() {
            WordSelectPrefs p = new WordSelectPrefs {
                LangCount = 2,
                Slots = new List<WordSlotPref> { Slot("ko-KR", "NIV"), Slot("en-US", "kjv") }
            }.Normalized();
            Assert.Null(p.Slots[0].Version);   // NIV is English → first Korean version
            Assert.Equal("KJV", p.Slots[1].Version);
        }

        [Theory]
        [InlineData(-1, 0)]
        [InlineData(2, 2)]
        [InlineData(3, 0)]
        public void Normalized_ResetsOutOfRangeRuby(int stored, int expected) {
            WordSelectPrefs p = new WordSelectPrefs {
                LangCount = 1,
                Slots = new List<WordSlotPref> { Slot("ja-JP", "JDB", stored) }
            }.Normalized();
            Assert.Equal(expected, p.Slots[0].Ruby);
        }
    }

    public class PreferencesTests {
        [Fact]
        public void Json_RoundTripsWordSelect() {
            Preferences src = new Preferences {
                WordSelect = new WordSelectPrefs {
                    LangCount = 2,
                    Slots = new List<WordSlotPref> {
                        new WordSlotPref { Language = "ja-JP", Version = "JDB", Ruby = 2 },
                        new WordSlotPref { Language = "ko-KR", Version = null, Ruby = 0 }
                    }
                }
            };
            Preferences back = Preferences.FromJson(src.ToJson());
            Assert.Equal(2, back.WordSelect.LangCount);
            Assert.Equal("ja-JP", back.WordSelect.Slots[0].Language);
            Assert.Equal("JDB", back.WordSelect.Slots[0].Version);
            Assert.Equal(2, back.WordSelect.Slots[0].Ruby);
            Assert.Null(back.WordSelect.Slots[1].Version);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("not json")]
        [InlineData("{\"wordSelect\":{\"langCount\":\"three\"}}")]
        public void FromJson_BlankOrMalformedYieldsEmpty(string json) {
            Assert.Null(Preferences.FromJson(json).WordSelect);
        }

        [Fact]
        public void Json_KeepsUnknownSectionsFromNewerBuilds() {
            string json = "{\"futureSection\":{\"port\":8291},\"wordSelect\":{\"langCount\":1,\"slots\":[]}}";
            Preferences p = Preferences.FromJson(json);
            p.WordSelect.LangCount = 2;
            string again = p.ToJson();
            Assert.Contains("\"futureSection\"", again);
            Assert.Contains("8291", again);
            Assert.Equal(2, Preferences.FromJson(again).WordSelect.LangCount);
        }

        [Fact]
        public void Store_SavesLoadsAndOverwrites() {
            string dir = Path.Combine(Path.GetTempPath(), "ZebulonVSTO.Tests." + Guid.NewGuid().ToString("N"));
            string path = Path.Combine(dir, "sub", "preferences.json");
            try {
                Assert.Null(PreferencesStore.Load(path).WordSelect); // missing file

                Preferences p = new Preferences { WordSelect = WordSelectPrefs.Defaults() };
                Assert.True(PreferencesStore.Save(p, path));       // creates the folder
                Assert.Equal(3, PreferencesStore.Load(path).WordSelect.LangCount);

                p.WordSelect.LangCount = 1;
                Assert.True(PreferencesStore.Save(p, path));       // replaces the existing file
                Assert.Equal(1, PreferencesStore.Load(path).WordSelect.LangCount);
                Assert.False(File.Exists(path + ".tmp"));
            } finally {
                if (Directory.Exists(dir)) {
                    Directory.Delete(dir, true);
                }
            }
        }
    }
}
