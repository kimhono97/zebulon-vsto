using System.Collections.Generic;

namespace ZebulonVSTO.Slides
{
    /// <summary>
    /// Describes a Bible language. The <see cref="Code"/> value is the BCP-47-style
    /// locale identifier ("lang-COUNTRY", e.g. "ko-KR") and matches BibleVersion.Language.
    /// Ported from jym-workbox/src/consts/LanguageConst.ts.
    /// </summary>
    public sealed class BibleLanguage
    {
        private readonly string _code;
        private readonly string _name;
        private readonly bool _rtl;

        public BibleLanguage(string code, string name, bool rtl)
        {
            _code = code;
            _name = name;
            _rtl = rtl;
        }

        /// <summary>Locale code, e.g. "ko-KR" (lowercase language, uppercase country).</summary>
        public string Code { get { return _code; } }

        /// <summary>Display (locale) name, e.g. "한국어".</summary>
        public string Name { get { return _name; } }

        /// <summary>True for right-to-left scripts (Hebrew).</summary>
        public bool Rtl { get { return _rtl; } }
    }

    /// <summary>
    /// Catalog of Bible languages, ported faithfully from LanguageConst.ts.
    /// The <see cref="All"/> ordering mirrors LanguageConst.ALL exactly.
    /// </summary>
    public static class LanguageCatalog
    {
        private static readonly BibleLanguage KoKr = new BibleLanguage("ko-KR", "한국어", false);
        private static readonly BibleLanguage EnUs = new BibleLanguage("en-US", "English", false);
        private static readonly BibleLanguage ZhCn = new BibleLanguage("zh-CN", "中文", false);
        private static readonly BibleLanguage ZhTw = new BibleLanguage("zh-TW", "漢文", false);
        private static readonly BibleLanguage JaJp = new BibleLanguage("ja-JP", "日本語", false);
        private static readonly BibleLanguage RuRu = new BibleLanguage("ru-RU", "Русский", false);
        private static readonly BibleLanguage MnMn = new BibleLanguage("mn-MN", "Монгол", false);
        private static readonly BibleLanguage HeIl = new BibleLanguage("he-IL", "עִבְרִית", true);
        private static readonly BibleLanguage ElGr = new BibleLanguage("el-GR", "ελληνικά", false);

        // Order matches LanguageConst.ALL:
        // KO_KR, EN_US, ZH_CN, MN_MN, JA_JP, HE_IL, EL_GR, RU_RU, ZH_TW
        private static readonly IReadOnlyList<BibleLanguage> _all = new List<BibleLanguage>
        {
            KoKr,
            EnUs,
            ZhCn,
            MnMn,
            JaJp,
            HeIl,
            ElGr,
            RuRu,
            ZhTw,
        }.AsReadOnly();

        /// <summary>All languages, in the same order as LanguageConst.ALL.</summary>
        public static IReadOnlyList<BibleLanguage> All
        {
            get { return _all; }
        }

        /// <summary>
        /// Finds a language by its <see cref="BibleLanguage.Code"/> (exact, case-sensitive
        /// match, e.g. "ko-KR"). Returns null if the code is unknown.
        /// </summary>
        public static BibleLanguage ByCode(string code)
        {
            if (code == null)
            {
                return null;
            }

            foreach (BibleLanguage lang in _all)
            {
                if (lang.Code == code)
                {
                    return lang;
                }
            }

            return null;
        }
    }
}
