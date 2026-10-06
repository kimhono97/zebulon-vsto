using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace ZebulonVSTO.Slides {
    /// <summary>One remembered language/version/ruby slot of the Word dialog.</summary>
    [DataContract]
    public sealed class WordSlotPref {
        /// <summary>Language code (e.g. "ko-KR"); see <see cref="LanguageCatalog"/>.</summary>
        [DataMember(Name = "lang")]
        public string Language { get; set; }

        /// <summary>Version code (e.g. "KRTRV"); null = the language's first version.</summary>
        [DataMember(Name = "version")]
        public string Version { get; set; }

        /// <summary>Ruby mode as the <see cref="RubyMode"/> ordinal (0=Base, 1=Furigana, 2=Both).</summary>
        [DataMember(Name = "ruby")]
        public int Ruby { get; set; }
    }

    /// <summary>
    /// The language/version/ruby configuration <see cref="WordSelectWindow"/>
    /// remembers between runs (a congregation normally reuses one setup). Pure
    /// and COM-free: <see cref="Normalized"/> repairs whatever was read from
    /// disk against the current catalogs, so a stale or hand-edited file can
    /// never put the dialog into an invalid state.
    /// </summary>
    [DataContract]
    public sealed class WordSelectPrefs {
        public const int SlotCount = 3;

        /// <summary>Default language per slot (also the reset target).</summary>
        public static readonly string[] DefaultLanguages = { "ko-KR", "en-US", "zh-CN" };

        /// <summary>Visible slot count, 1..<see cref="SlotCount"/>.</summary>
        [DataMember(Name = "langCount")]
        public int LangCount { get; set; }

        /// <summary>All slots, hidden ones included, so raising the count restores them.</summary>
        [DataMember(Name = "slots")]
        public List<WordSlotPref> Slots { get; set; }

        /// <summary>The out-of-the-box configuration: 3 languages, ko/en/zh, first versions, base ruby.</summary>
        public static WordSelectPrefs Defaults() {
            WordSelectPrefs p = new WordSelectPrefs { LangCount = SlotCount, Slots = new List<WordSlotPref>() };
            for (int i = 0; i < SlotCount; i++) {
                p.Slots.Add(DefaultSlot(i));
            }
            return p;
        }

        private static WordSlotPref DefaultSlot(int slot) {
            return new WordSlotPref { Language = DefaultLanguages[slot], Version = null, Ruby = 0 };
        }

        /// <summary>
        /// A repaired copy: count clamped to 1..3, exactly 3 slots; an unknown
        /// language resets that slot to its default; a version that doesn't belong
        /// to the slot's language becomes null (first version); an out-of-range
        /// ruby mode becomes Base. Version codes are canonicalized to the catalog's
        /// casing (language codes must match exactly, as LanguageCatalog.ByCode does).
        /// </summary>
        public WordSelectPrefs Normalized() {
            WordSelectPrefs p = new WordSelectPrefs { Slots = new List<WordSlotPref>() };
            p.LangCount = LangCount < 1 ? 1 : LangCount > SlotCount ? SlotCount : LangCount;
            for (int i = 0; i < SlotCount; i++) {
                WordSlotPref src = (Slots != null && i < Slots.Count) ? Slots[i] : null;
                p.Slots.Add(NormalizeSlot(src, i));
            }
            return p;
        }

        private static WordSlotPref NormalizeSlot(WordSlotPref src, int slot) {
            BibleLanguage lang = src != null ? LanguageCatalog.ByCode(src.Language) : null;
            if (lang == null) {
                return DefaultSlot(slot);
            }
            string version = null;
            BibleVersion v = BibleCatalog.ParseVersion(src.Version);
            if (v != null && string.Equals(v.Language, lang.Code, StringComparison.OrdinalIgnoreCase)) {
                version = v.Code;
            }
            int ruby = src.Ruby >= (int)RubyMode.Base && src.Ruby <= (int)RubyMode.Both ? src.Ruby : (int)RubyMode.Base;
            return new WordSlotPref { Language = lang.Code, Version = version, Ruby = ruby };
        }
    }
}
