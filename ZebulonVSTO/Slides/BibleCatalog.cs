using System;
using System.Collections.Generic;

namespace ZebulonVSTO.Slides {
    /// <summary>
    /// A faithful C# port of the Bible catalog from the Zebulon Web project
    /// (src/consts/BibleConst.ts + src/consts/LanguageConst.ts).
    ///
    /// Holds the 18 Bible versions (in BibleVersionConst.ALL order), the 66-book
    /// name arrays for every language present (ko-KR / en-US / zh-CN / zh-TW /
    /// ja-JP / ru-RU / mn-MN / he-IL / el-GR), the per-book chapter counts
    /// (NUM_CH), and accessors mirroring s_GetBookName / s_GetAllBookName /
    /// s_GetBookLength.
    ///
    /// Targets C# 7.3 (.NET Framework 4.7.2). COM-free (no Office / Interop, no
    /// WPF) so it can be link-compiled into the xUnit test project.
    /// </summary>
    public sealed class BibleVersion {
        private readonly string m_code;
        private readonly string m_name;
        private readonly string m_language;
        private readonly bool m_hasRuby;

        public BibleVersion(string code, string name, string language, bool hasRuby) {
            m_code = code;
            m_name = name;
            m_language = language;
            m_hasRuby = hasRuby;
        }

        /// <summary>The version key, e.g. "KRTRV", "ESV", "JDB". Always upper-case.</summary>
        public string Code { get { return m_code; } }

        /// <summary>The version's display name, e.g. "개역개정", "ESV", "新改訳".</summary>
        public string Name { get { return m_name; } }

        /// <summary>The language code this version maps to, e.g. "ko-KR", "ja-JP".</summary>
        public string Language { get { return m_language; } }

        /// <summary>
        /// Whether this version's verse data contains &lt;ruby&gt; (furigana)
        /// markup. Currently only JDB (新改訳).
        /// </summary>
        public bool HasRuby { get { return m_hasRuby; } }
    }

    public static class BibleCatalog {
        public const int BookCount = 66;

        // --- Language codes (LanguageConst.toString() => "{lang}-{COUNTRY}") --------
        private const string KO_KR = "ko-KR";
        private const string EN_US = "en-US";
        private const string ZH_CN = "zh-CN";
        private const string ZH_TW = "zh-TW";
        private const string JA_JP = "ja-JP";
        private const string RU_RU = "ru-RU";
        private const string MN_MN = "mn-MN";
        private const string HE_IL = "he-IL";
        private const string EL_GR = "el-GR";

        private const int NUM_OT = 39;
        private const int NUM_NT = 27;

        // --- Versions, in BibleVersionConst.ALL order (18 total) -------------------
        private static readonly BibleVersion[] s_all = new BibleVersion[] {
            new BibleVersion("KRTRV", "개역개정", KO_KR, false),
            new BibleVersion("KRTKO", "개역한글", KO_KR, false),
            new BibleVersion("KCOT", "공동번역", KO_KR, false),
            new BibleVersion("KNT", "새번역", KO_KR, false),
            new BibleVersion("KMMB", "현대인의 성경", KO_KR, false),
            new BibleVersion("ESV", "ESV", EN_US, false),
            new BibleVersion("NIV", "NIV", EN_US, false),
            new BibleVersion("KJV", "KJV", EN_US, false),
            new BibleVersion("NASB", "NASB", EN_US, false),
            new BibleVersion("CUV", "圣经", ZH_CN, false),
            new BibleVersion("CUVT", "聖經", ZH_TW, false),
            new BibleVersion("JDB", "新改訳", JA_JP, true),
            new BibleVersion("JNCOT", "新共同訳", JA_JP, false),
            new BibleVersion("JCLQT", "口語訳", JA_JP, false),
            new BibleVersion("RU", "БИБЛИЯ", RU_RU, false),
            new BibleVersion("MN", "АРИУН БИБЛИ", MN_MN, false),
            new BibleVersion("HE", "הקודש במקרא", HE_IL, false),
            new BibleVersion("EL", "ΒΙΒΛΟΣ", EL_GR, false),
        };

        private static readonly IReadOnlyList<BibleVersion> s_allReadOnly =
            Array.AsReadOnly(s_all);

        public static IReadOnlyList<BibleVersion> AllVersions {
            get { return s_allReadOnly; }
        }

        /// <summary>All versions whose Language equals the given code (case-insensitive), in ALL order.</summary>
        public static IReadOnlyList<BibleVersion> VersionsForLanguage(string languageCode) {
            List<BibleVersion> ret = new List<BibleVersion>();
            if (languageCode == null) {
                return ret.AsReadOnly();
            }
            for (int i = 0; i < s_all.Length; i++) {
                if (string.Equals(s_all[i].Language, languageCode, StringComparison.OrdinalIgnoreCase)) {
                    ret.Add(s_all[i]);
                }
            }
            return ret.AsReadOnly();
        }

        /// <summary>
        /// Find a version by code (case-insensitive, trimmed); null if unknown.
        /// Mirrors BibleVersionConst.parse.
        /// </summary>
        public static BibleVersion ParseVersion(string code) {
            if (code == null) {
                return null;
            }
            string key = code.Trim().ToUpperInvariant();
            for (int i = 0; i < s_all.Length; i++) {
                if (s_all[i].Code == key) {
                    return s_all[i];
                }
            }
            return null;
        }

        // --- Chapter counts (NUM_CH): 39 OT + 27 NT = 66 entries -------------------
        private static readonly int[] s_numCh = new int[] {
            // Old Testament (39)
            50, 40, 27, 36, 34, 24, 21, 4, 31, 24, 22, 25, 29, 36, 10, 13, 10, 42, 150, 31,
            12, 8, 66, 52, 5, 48, 12, 14, 3, 9, 1, 4, 7, 3, 3, 3, 2, 14, 4,
            // New Testament (27)
            28, 16, 24, 21, 28, 16, 16, 13, 6, 6, 4, 4, 5, 3, 6, 4, 3, 1, 13, 5,
            5, 3, 5, 1, 1, 1, 22
        };

        /// <summary>Chapter count for a book (book is 1..66). 0 if out of range. Mirrors s_GetBookLength.</summary>
        public static int ChapterCount(int book) {
            int idx = book - 1;
            if (idx < 0 || idx >= s_numCh.Length) {
                return 0;
            }
            return s_numCh[idx];
        }

        /// <summary>True if the book (1..66) is in the Old Testament (book &lt;= 39).</summary>
        public static bool IsOldTestament(int book) {
            return book <= NUM_OT;
        }

        // ===========================================================================
        // Book name arrays. OT arrays have 39 entries, NT arrays have 27 entries.
        // ===========================================================================

        // --- Korean (ko-KR) --------------------------------------------------------
        private static readonly string[] KR_OT = new string[] {
            "창세기", "출애굽기", "레위기", "민수기", "신명기", "여호수아", "사사기", "룻기", "사무엘상", "사무엘하",
            "열왕기상", "열왕기하", "역대상", "역대하", "에스라", "느헤미야", "에스더", "욥기", "시편", "잠언",
            "전도서", "아가", "이사야", "예레미야", "예레미야 애가", "에스겔", "다니엘", "호세아", "요엘", "아모스",
            "오바댜", "요나", "미가", "나훔", "하박국", "스바냐", "학개", "스가랴", "말라기"
        };
        private static readonly string[] KR_NT = new string[] {
            "마태복음", "마가복음", "누가복음", "요한복음", "사도행전", "로마서", "고린도전서", "고린도후서", "갈라디아서", "에베소서",
            "빌립보서", "골로새서", "데살로니가전서", "데살로니가후서", "디모데전서", "디모데후서", "디도서", "빌레몬서", "히브리서", "야고보서",
            "베드로전서", "베드로후서", "요한일서", "요한이서", "요한삼서", "유다서", "요한계시록"
        };
        private static readonly string[] KR_OTA = new string[] {
            "창", "출", "레", "민", "신", "수", "삿", "룻", "삼상", "삼하",
            "왕상", "왕하", "대상", "대하", "스", "느", "에", "욥", "시", "잠",
            "전", "아", "사", "렘", "애", "겔", "단", "호", "욜", "암",
            "옵", "욘", "미", "나", "합", "습", "학", "슥", "말"
        };
        private static readonly string[] KR_NTA = new string[] {
            "마", "막", "눅", "요", "행", "롬", "고전", "고후", "갈", "엡",
            "빌", "골", "살전", "살후", "딤전", "딤후", "딛", "몬", "히", "약",
            "벧전", "벧후", "요일", "요이", "요삼", "유", "계"
        };

        // --- English (en-US) -------------------------------------------------------
        private static readonly string[] EN_OT = new string[] {
            "Genesis", "Exodus", "Leviticus", "Numbers", "Deuteronomy", "Joshua", "Judges", "Ruth", "1 Samuel", "2 Samuel",
            "1 Kings", "2 Kings", "1 Chronicles", "2 Chronicles", "Ezra", "Nehemiah", "Esther", "Job", "Psalms", "Proverbs",
            "Ecclesiastes", "Song of Songs", "Isaiah", "Jeremiah", "Lamentations", "Ezekiel", "Daniel", "Hosea", "Joel", "Amos",
            "Obadiah", "Jonah", "Micah", "Nahum", "Habakkuk", "Zephaniah", "Haggai", "Zechariah", "Malachi"
        };
        private static readonly string[] EN_NT = new string[] {
            "Matthew", "Mark", "Luke", "John", "Acts", "Romans", "1 Corinthians", "2 Corinthians", "Galatians", "Ephesians",
            "Philippians", "Colossians", "1 Thessalonians", "2 Thessalonians", "1 Timothy", "2 Timothy", "Titus", "Philemon", "Hebrews", "James",
            "1 Peter", "2 Peter", "1 John", "2 John", "3 John", "Jude", "Revelation"
        };
        private static readonly string[] EN_OTA = new string[] {
            "GEN", "EXO", "LEV", "NUM", "DEU", "JOS", "JDG", "RUT", "1SA", "2SA",
            "1KI", "2KI", "1CH", "2CH", "EZR", "NEH", "EST", "JOB", "PSA", "PRO",
            "ECC", "SNG", "ISA", "JER", "LAM", "EZK", "DAN", "HOS", "JOL", "AMO",
            "OBA", "JON", "MIC", "NAM", "HAB", "ZEP", "HAG", "ZEC", "MAL"
        };
        private static readonly string[] EN_NTA = new string[] {
            "MAT", "MRK", "LUK", "JHN", "ACT", "ROM", "1CO", "2CO", "GAL", "EPH",
            "PHP", "COL", "1TH", "2TH", "1TI", "2TI", "TIT", "PHM", "HEB", "JAS",
            "1PE", "2PE", "1JN", "2JN", "3JN", "JUD", "REV"
        };

        // --- Chinese Simplified (zh-CN) --------------------------------------------
        private static readonly string[] CN_OT = new string[] {
            "创世记", "出埃及记", "利未记", "民数记", "申命记", "约书亚记", "士师记", "路得记", "撒母耳记上", "撒母耳记下",
            "列王记上", "列王记下", "历代志上", "历代志下", "以斯拉记", "尼希米记", "以斯帖记", "约伯记", "诗篇", "箴言",
            "传道书", "雅歌", "以赛亚书", "耶利米书", "耶利米哀歌", "以西结书", "但以理书", "何西阿书", "约珥书", "阿摩司书",
            "俄巴底亚书", "约拿书", "弥迦书", "那鸿书", "哈巴谷书", "西番雅书", "哈该书", "撒迦利亚书", "玛拉基书"
        };
        private static readonly string[] CN_NT = new string[] {
            "马太福音", "马可福音", "路加福音", "约翰福音", "使徒行传", "罗马书", "哥林多前书", "哥林多後书", "加拉太书", "以弗所书",
            "腓立比书", "歌罗西书", "帖撒罗尼迦前书", "帖撒罗尼迦後书", "提摩太前书", "提摩太後书", "提多书", "腓利门书", "希伯来书", "雅各书",
            "彼得前书", "彼得後书", "约翰一书", "约翰二书", "约翰三书", "犹大书", "启示录"
        };

        // --- Chinese Traditional (zh-TW) -------------------------------------------
        private static readonly string[] TW_OT = new string[] {
            "創世記", "出埃及記", "利未記", "民數記", "申命記", "約書亞記", "士師記", "路得記", "撒母耳記上", "撒母耳記下",
            "列王紀上", "列王紀下", "歷代志上", "歷代志下", "以斯拉記", "尼希米記", "以斯帖記", "約伯記", "詩篇", "箴言",
            "傳道書", "雅歌", "以賽亞書", "耶利米書", "耶利米哀歌", "以西結書", "但以理書", "何西阿書", "約珥書", "阿摩司書",
            "俄巴底亞書", "約拿書", "彌迦書", "那鴻書", "哈巴谷書", "西番雅書", "哈該書", "撒迦利亞書", "瑪拉基書"
        };
        private static readonly string[] TW_NT = new string[] {
            "馬太福音", "馬可福音", "路加福音", "約翰福音", "使徒行傳", "羅馬書", "哥林多前書", "哥林多後書", "加拉太書", "以弗所書",
            "腓立比書", "歌羅西書", "帖撒羅尼迦前書", "帖撒羅尼迦後書", "提摩太前書", "提摩太後書", "提多書", "腓利門書", "希伯來書", "雅各書",
            "彼得前書", "彼得後書", "約翰壹書", "約翰貳書", "約翰參書", "猶大書", "啟示錄"
        };

        // --- Japanese (ja-JP) ------------------------------------------------------
        private static readonly string[] JP_OT = new string[] {
            "創世記", "出エジプト記", "レビ記", "民数記", "申命記", "ヨシュア記", "士師記", "ルツ記", "サムエル記上", "サムエル記下",
            "列王記上", "列王記下", "歴代誌上", "歴代誌下", "エズラ記", "ネヘミヤ記", "エステル記", "ヨブ記", "詩篇", "箴言",
            "伝道の書", "雅歌", "イザヤ書", "エレミヤ書", "哀歌", "エゼキエル書", "ダニエル書", "ホセア書", "ヨエル書", "アモス書",
            "オバデヤ書", "ヨナ書", "ミカ書", "ナホム書", "ハバクク書", "ゼパニヤ書", "ハガイ書", "ゼカリヤ書", "マラキ書"
        };
        private static readonly string[] JP_NT = new string[] {
            "マタイによる福音書", "マルコによる福音書", "ルカによる福音書", "ヨハネによる福音書", "使徒行伝", "ローマ人への手紙", "コリント人への手紙第一", "コリント人への手紙第二", "ガラテヤ人への手紙", "エペソ人への手紙",
            "ピリピ人への手紙", "コロサイ人への手紙", "テサロニケ人への手紙第一", "テサロニケ人への手紙第二", "テモテへの手紙第一", "テモテへの手紙第二", "テトスへの手紙", "ピレモンへの手紙", "ヘブル人への手紙", "ヤコブの手紙",
            "ペテロの手紙第一", "ペテロの手紙第二", "ヨハネの手紙第一", "ヨハネの手紙第二", "ヨハネの手紙第三", "ユダの手紙", "ヨハネの黙示録"
        };

        // --- Russian (ru-RU) -------------------------------------------------------
        private static readonly string[] RU_OT = new string[] {
            "БЫТИЕ", "ИСХОД", "ЛЕВИТ", "ЧИСЛА", "ВТОРОЗАКОНИЕ", "ИИСУС НАВИН", "КНИГА СУДЕЙ", "РУФЬ", "1-Я ЦАРСТВ", "2-Я ЦАРСТВ",
            "3-Я ЦАРСТВ", "4-Я ЦАРСТВ", "1-Я ПАРАЛИПОМЕНОН", "2-Я ПАРАЛИПОМЕНОН", "ЕЗДРА", "НЕЕМИЯ", "ЕСФИРЬ", "ИОВ", "ПСАЛТИРЬ", "ПРИТЧИ",
            "ЕККЛЕСИАСТ", "ПЕСНИ ПЕСНЕЙ", "ИСАИЯ", "ИЕРЕМИЯ", "ПЛАЧ ИЕРЕМИИ", "ИЕЗЕКИИЛЬ", "ДАНИИЛ", "ОСИЯ", "ИОИЛЬ", "АМОС",
            "АВДИЯ", "ИОНА", "МИХЕЙ", "НАУМ", "АВВАКУМ", "СОФОНИЯ", "АГГЕЙ", "ЗАХАРИЯ", "МАЛАХИЯ"
        };
        private static readonly string[] RU_NT = new string[] {
            "МАТФЕЯ", "МАРКА", "ЛУКИ", "ИОАННА", "ДЕЯНИЯ", "РИМЛЯНАМ", "1-Е КОРИНФЯНАМ", "2-Е КОРИНФЯНАМ", "ГАЛАТАМ", "ЕФЕСЯНАМ",
            "ФИЛИППИЙЦАМ", "КОЛОССЯНАМ", "1-Е ФЕССАЛОНИКИЙЦАМ", "2-Е ФЕССАЛОНИКИЙЦАМ", "1-Е ТИМОФЕЮ", "2-Е ТИМОФЕЮ", "ТИТУ", "ФИЛИМОНУ", "ЕВРЕЯМ", "ИАКОВА",
            "1-E ПЕТРА", "2-E ПЕТРА", "1-E ИОАННА", "2-E ИОАННА", "3-E ИОАННА", "ИУДА", "ОТКРОВЕНИЕ"
        };

        // --- Mongolian (mn-MN) -----------------------------------------------------
        private static readonly string[] MN_OT = new string[] {
            "ЭХЛЭЛ", "ЕГИПЕТЭЭС ГАРСАН НЬ", "ЛЕВИТ", "ТООЛЛОГО", "ДЭД ХУУЛЬ", "ИОШУА", "ШҮҮГЧИД", "РУТ", "1 САМУЕЛ", "2 САМУЕЛ",
            "ХААДЫН ДЭЭД", "ХААДЫН ДЭД", "ШАСТИРЫН ДЭЭД", "ШАСТИРЫН ДЭД", "ЕЗРА", "НЕХЕМИА", "ЕСТЕР", "ИОВ", "ДУУЛАЛ", "СУРГААЛТ ҮГС။",
            "НОМЛОГЧИЙН ҮГС", "СОЛОМОНЫ ДУУН", "ИСАИА", "ИЕРЕМИА", "ГАШУУДАЛ", "ЕЗЕКИЕЛ", "ДАНИЕЛ", "ХОСЕА", "ИОЕЛ", "АМОС",
            "ОБАДИА", "ИОНА", "МИКА", "НАХУМ", "ХАБАККУК", "ЗЕФАНИА", "ХАГГАИ", "ЗЕХАРИА", "МАЛАХИ"
        };
        private static readonly string[] MN_NT = new string[] {
            "МАТАЙ", "МАРK", "ЛУК", "ИОХАН", "ҮЙЛС", "РОМ", "1 КОРИНТ", "2 КОРИНТ", "ГАЛАТ", "ЕФЕС",
            "ФИЛИППОЙ", "КОЛОССАЙ", "1 ТЕСАЛОНИК", "2 ТЕСАЛОНИК", "1 ТИМОТ", "2 ТИМОТ", "ТИТ", "ФИЛЕМОН", "ЕВРЕЙ", "ИАКОВ",
            "1 ПЕТР", "2 ПЕТР", "1 ИОХАН", "2 ИОХАН", "3 ИОХАН", "ИУДА", "ИЛЧЛЭЛТ"
        };

        // --- Hebrew (he-IL) --------------------------------------------------------
        private static readonly string[] HE_OT = new string[] {
            "בראשית", "שמות", "ויקרא", "במדבר", "דברים", "יהושע", "שפטים", "רות", "שמואל א", "שמואל ב",
            "מלכים א", "מלכים ב", "דברי הימים א", "דברי הימים ב", "עזרא", "נחמיה", "אסתר", "איוב", "תהלים", "משלי",
            "קהלת", "שיר השירים", "ישעיה", "ירמיה", "איכה", "יחזקאל", "דניאל", "הושע", "יואל", "עמוס",
            "עבדיה", "יונה", "מיכה", "נחום", "חבקוק", "צפניה", "חגי", "זכריה", "מלאכי"
        };
        private static readonly string[] HE_NT = new string[] {
            "מַתָּי", "מרקוםי", "לוקם", "יוחנן", "מעשי השליחים", "אל־הרומיים", "הראשונה אל־הקורינתים", "השנייה אל־הקורינתים", "אל־הגלטים", "אל־האפסיים",
            "אל־הפיליפיים", "אל־הקולומים", "הראשונה אל־התסלוניקים", "השנייה אל־התסלוניקים", "הראשונה אל־טימותיום", "השנייה אל־טימותיום", "אל־טיטום", "אל־פילימון", "אל־העברים", "יעקב",
            "פטרום הראשונה", "פטרום השנייה", "יוחנן הראשונה", "יוחנן השנייה", "יוחנן השלישית", "יהודה", "התגלות"
        };

        // --- Greek (el-GR) ---------------------------------------------------------
        private static readonly string[] EL_OT = new string[] {
            "ΓΕΝΕΣΗ", "ΕΞΟΔΟΣ", "ΛΕΥΙΤΙΚΟ", "ΑΡΙΘΜΟΙ", "ΔΕΥΤΕΡΟΝΟΜΙΟ", "ΙΗΣΟΥΣ ΤΟΥ ΝΑΥΗ", "ΚΡΙΤΕΣ", "ΡΟΥΘ", "1 ΣΑΜΟΥΗΛ", "2 ΣΑΜΟΥΗΛ",
            "1 ΒΑΣΙΛΕΩΝ", "2 ΒΑΣΙΛΕΩΝ", "1 ΧΡΟΝΙΚΩΝ", "2 ΧΡΟΝΙΚΩΝ", "ΕΣΔΡΑΣ", "ΝΕΕΜΙΑΣ", "ΕΣΘΗΡ", "ΙΩΒ", "ΨΑΛΜΟΙ", "ΠΑΡΟΙΜΙΕΣ",
            "ΕΚΚΛΗΣΙΑΣΤΗΣ", "ΑΣΜΑ ΣΟΛΟΜΩΝΤΟΣ (ΑΣΜΑ ΑΣΜΑΤΩΝ)", "ΗΣΑΪΑΣ", "ΙΕΡΕΜΙΑΣ", "ΘΡΗΝΟΙ", "ΙΕΖΕΚΙΗΛ", "ΔΑΝΙΗΛ", "ΩΣΗΕ", "ΙΩΗΛ", "ΑΜΩΣ",
            "ΑΒΔΙΟΥ", "ΙΩΝΑΣ", "ΜΙΧΑΙΑΣ", "ΝΑΟΥΜ", "ΑΒΒΑΚΟΥΜ", "ΣΟΦΟΝΙΑΣ", "ΑΓΓΑΙΟΣ", "ΖΑΧΑΡΙΑΣ", "ΜΑΛΑΧΙΑΣ"
        };
        private static readonly string[] EL_NT = new string[] {
            "ΜΑΘΘΑΙΟΝ", "ΜΑΡΚΟΝ", "ΛΟΥΚΑΝ", "ΙΩΑΝΝΗΝ", "ΠΡΑΞΕΙΣ ΑΠΟΣΤΟΛΩΝ", "ΠΡΟΣ ΡΩΜΑΙΟΥΣΣ", "ΚΟΡΙΝΘΙΟΥΣ Α’", "ΚΟΡΙΝΘΙΟΥΣ Β’", "ΓΑΛΑΤΑΣ", "ΕΦΕΣΙΟΥΣ",
            "ΦΙΛΙΠΠΗΣΙΟΥΣ", "ΚΟΛΟΣΣΑΕΙΣ", "ΘΕΣΣΑΛΟΝΙΚΕΙΣ Α’", "ΘΕΣΣΑΛΟΝΙΚΕΙΣ Β’", "ΤΙΜΟΘΕΟΝ Α’", "ΤΙΜΟΘΕΟΝ Β’", "ΤΙΤΟΝ", "ΦΙΛΗΜΟΝΑ", "ΕΒΡΑΙΟΥΣ", "ΙΑΚΩΒΟΥ",
            "ΕΠΙΣΤΟΛΗ Α’", "ΕΠΙΣΤΟΛΗ Β’", "ΙΩΑΝΝΟΥ ΕΠΙΣΤΟΛΗ Α’", "ΙΩΑΝΝΟΥ ΕΠΙΣΤΟΛΗ Β’", "ΙΩΑΝΝΟΥ ΕΠΙΣΤΟΛΗ Γ’", "ΙΟΥΔΑ ΕΠΙΣΤΟΛΗ", "ΑΠΟΚΑΛΥΨΙΣ"
        };

        /// <summary>
        /// Book name for a book (1..66) in a given language. Mirrors s_GetBookName:
        /// abbreviated forms exist only for ko-KR and en-US; every other language
        /// ignores <paramref name="abbreviated"/> and returns the full name. Unknown
        /// language codes fall back to en-US (the TS default). Returns "" if the book
        /// number is out of the 1..66 range.
        /// </summary>
        public static string BookName(int book, string languageCode, bool abbreviated = false) {
            int idx = book - 1;
            if (idx < 0 || idx >= NUM_OT + NUM_NT) {
                return "";
            }
            bool isOt = idx < NUM_OT;
            int localIdx = isOt ? idx : idx - NUM_OT;
            string lang = NormalizeLang(languageCode);

            if (isOt) {
                if (lang == KO_KR) { return (abbreviated ? KR_OTA : KR_OT)[localIdx]; }
                if (lang == ZH_CN) { return CN_OT[localIdx]; }
                if (lang == ZH_TW) { return TW_OT[localIdx]; }
                if (lang == JA_JP) { return JP_OT[localIdx]; }
                if (lang == RU_RU) { return RU_OT[localIdx]; }
                if (lang == MN_MN) { return MN_OT[localIdx]; }
                if (lang == HE_IL) { return HE_OT[localIdx]; }
                if (lang == EL_GR) { return EL_OT[localIdx]; }
                return (abbreviated ? EN_OTA : EN_OT)[localIdx];
            }

            if (lang == KO_KR) { return (abbreviated ? KR_NTA : KR_NT)[localIdx]; }
            if (lang == ZH_CN) { return CN_NT[localIdx]; }
            if (lang == ZH_TW) { return TW_NT[localIdx]; }
            if (lang == JA_JP) { return JP_NT[localIdx]; }
            if (lang == RU_RU) { return RU_NT[localIdx]; }
            if (lang == MN_MN) { return MN_NT[localIdx]; }
            if (lang == HE_IL) { return HE_NT[localIdx]; }
            if (lang == EL_GR) { return EL_NT[localIdx]; }
            return (abbreviated ? EN_NTA : EN_NT)[localIdx];
        }

        /// <summary>
        /// All 66 book names for a language (OT then NT). Mirrors s_GetAllBookName:
        /// abbreviated forms exist only for ko-KR and en-US; other languages ignore
        /// the flag. Unknown codes fall back to en-US.
        /// </summary>
        public static IReadOnlyList<string> AllBookNames(string languageCode, bool abbreviated = false) {
            string lang = NormalizeLang(languageCode);
            string[] ot;
            string[] nt;

            if (lang == KO_KR) {
                ot = abbreviated ? KR_OTA : KR_OT;
                nt = abbreviated ? KR_NTA : KR_NT;
            } else if (lang == ZH_CN) {
                ot = CN_OT; nt = CN_NT;
            } else if (lang == ZH_TW) {
                ot = TW_OT; nt = TW_NT;
            } else if (lang == JA_JP) {
                ot = JP_OT; nt = JP_NT;
            } else if (lang == RU_RU) {
                ot = RU_OT; nt = RU_NT;
            } else if (lang == MN_MN) {
                ot = MN_OT; nt = MN_NT;
            } else if (lang == HE_IL) {
                ot = HE_OT; nt = HE_NT;
            } else if (lang == EL_GR) {
                ot = EL_OT; nt = EL_NT;
            } else {
                ot = abbreviated ? EN_OTA : EN_OT;
                nt = abbreviated ? EN_NTA : EN_NT;
            }

            List<string> all = new List<string>(ot.Length + nt.Length);
            all.AddRange(ot);
            all.AddRange(nt);
            return all.AsReadOnly();
        }

        /// <summary>
        /// Map an arbitrary (possibly mixed-case) language code to one of the known
        /// canonical codes, or return the input unchanged (so unknowns fall through
        /// to the en-US default in the callers).
        /// </summary>
        private static string NormalizeLang(string languageCode) {
            if (languageCode == null) {
                return EN_US;
            }
            if (string.Equals(languageCode, KO_KR, StringComparison.OrdinalIgnoreCase)) { return KO_KR; }
            if (string.Equals(languageCode, EN_US, StringComparison.OrdinalIgnoreCase)) { return EN_US; }
            if (string.Equals(languageCode, ZH_CN, StringComparison.OrdinalIgnoreCase)) { return ZH_CN; }
            if (string.Equals(languageCode, ZH_TW, StringComparison.OrdinalIgnoreCase)) { return ZH_TW; }
            if (string.Equals(languageCode, JA_JP, StringComparison.OrdinalIgnoreCase)) { return JA_JP; }
            if (string.Equals(languageCode, RU_RU, StringComparison.OrdinalIgnoreCase)) { return RU_RU; }
            if (string.Equals(languageCode, MN_MN, StringComparison.OrdinalIgnoreCase)) { return MN_MN; }
            if (string.Equals(languageCode, HE_IL, StringComparison.OrdinalIgnoreCase)) { return HE_IL; }
            if (string.Equals(languageCode, EL_GR, StringComparison.OrdinalIgnoreCase)) { return EL_GR; }
            return languageCode;
        }
    }
}
