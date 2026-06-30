using System;
using System.Collections.Generic;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace ZebulonVSTO.Slides {
    /// <summary>
    /// Ruby (furigana) HTML text handling, ported 1:1 from Zebulon Web:
    ///   - jym-workbox/src/utils/ClientUtils.ts: trimRubyHtmlText + katakanaToHiragana
    ///   - jym-workbox/src/zebulon/components/word/helpers.ts: applyRubyMode
    ///
    /// The TypeScript original runs in the browser and relies on the DOM
    /// (DOMParser + Element.textContent). There is no DOM here, so the equivalent
    /// behavior is reproduced with regex/string operations plus
    /// System.Net.WebUtility.HtmlDecode (available on .NET Framework 4.7.2) for
    /// HTML-entity decoding. The net effect is identical:
    ///   - &lt;ruby&gt; markup is rewritten per mode (base / furigana / both),
    ///   - &lt;mark&gt; is always unwrapped,
    ///   - all remaining tags are stripped (mirrors reading doc.body.textContent),
    ///   - entities are decoded and the result is trimmed.
    ///
    /// RubyMode maps to trimRubyHtmlText's pattern strings as:
    ///   Base     -> "ruby"      (unwrap &lt;ruby&gt;, keep base text, drop &lt;rt&gt;/&lt;rp&gt;)
    ///   Furigana -> "rt"        (take &lt;rt&gt; reading text, then katakana -&gt; hiragana)
    ///   Both     -> "ruby(rt)"  ("base(reading)")
    ///
    /// COM-free (no Office/Interop, no WPF) so it can be link-compiled into the
    /// xUnit test project. Targets C# 7.3 / .NET Framework 4.7.2 — classic C# only.
    /// </summary>
    public enum RubyMode {
        Base,
        Furigana,
        Both
    }

    /// <summary>
    /// Pure, COM-free port of the Zebulon Web ruby-text helpers. See the file-level
    /// summary for the source mapping.
    /// </summary>
    public static class RubyText {
        // <ruby> ... </ruby>  (case-insensitive, dot matches newline, non-greedy inner).
        private static readonly Regex RubyBlockRegex = new Regex(
            @"<ruby\b[^>]*>(?<inner>.*?)</ruby\s*>",
            RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.CultureInvariant);

        // <rt ...> ... </rt>
        private static readonly Regex RtElementRegex = new Regex(
            @"<rt\b[^>]*>(?<inner>.*?)</rt\s*>",
            RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.CultureInvariant);

        // <rp ...> ... </rp>
        private static readonly Regex RpElementRegex = new Regex(
            @"<rp\b[^>]*>.*?</rp\s*>",
            RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.CultureInvariant);

        // <mark ...> ... </mark>  (unwrapped: tags removed, children kept).
        private static readonly Regex MarkOpenRegex = new Regex(
            @"<mark\b[^>]*>",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        private static readonly Regex MarkCloseRegex = new Regex(
            @"</mark\s*>",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        // Any remaining HTML tag (mirrors textContent dropping all element markup).
        private static readonly Regex AnyTagRegex = new Regex(
            @"<[^>]*>",
            RegexOptions.Singleline | RegexOptions.CultureInvariant);

        /// <summary>
        /// Port of applyRubyMode(line, mode):
        ///   furigana -&gt; katakanaToHiragana(trimRubyHtmlText(line, "rt"))
        ///   both     -&gt; trimRubyHtmlText(line, "ruby(rt)")
        ///   base     -&gt; trimRubyHtmlText(line)  (pattern "ruby")
        /// </summary>
        public static string ApplyRubyMode(string htmlLine, RubyMode mode) {
            if (mode == RubyMode.Furigana) {
                return KatakanaToHiragana(TrimRubyHtmlText(htmlLine, "rt"));
            }
            if (mode == RubyMode.Both) {
                return TrimRubyHtmlText(htmlLine, "ruby(rt)");
            }
            // RubyMode.Base and any default.
            return TrimRubyHtmlText(htmlLine, "ruby");
        }

        /// <summary>
        /// Strip all tags and decode entities, trimmed. Equivalent to Base mode for
        /// non-ruby lines (trimRubyHtmlText with the default "ruby" pattern on text
        /// that contains no &lt;ruby&gt; markup is just textContent).
        /// </summary>
        public static string ToPlainText(string htmlLine) {
            return TrimRubyHtmlText(htmlLine, "ruby");
        }

        /// <summary>
        /// Parse "base"/"furigana"/"both" (case-insensitive) into a RubyMode.
        /// Defaults to Base for null/unknown input (mirrors the "base" default).
        /// </summary>
        public static RubyMode ParseMode(string s) {
            if (s == null) {
                return RubyMode.Base;
            }
            string trimmed = s.Trim();
            if (string.Equals(trimmed, "furigana", StringComparison.OrdinalIgnoreCase)) {
                return RubyMode.Furigana;
            }
            if (string.Equals(trimmed, "both", StringComparison.OrdinalIgnoreCase)) {
                return RubyMode.Both;
            }
            // "base" and everything else.
            return RubyMode.Base;
        }

        /// <summary>
        /// Port of trimRubyHtmlText(htmlText, pattern). pattern is one of
        /// "ruby" (default), "rt", "ruby(rt)". See the TS original in ClientUtils.ts.
        /// </summary>
        private static string TrimRubyHtmlText(string htmlText, string pattern) {
            if (string.IsNullOrEmpty(htmlText)) {
                return "";
            }

            // 1) Rewrite each <ruby>...</ruby> block according to the pattern. This
            //    mirrors the per-rubyElement DOM mutation in the TS source.
            string working = RubyBlockRegex.Replace(htmlText, match => {
                string inner = match.Groups["inner"].Value;
                return RewriteRubyBlock(inner, pattern);
            });

            // 2) Always unwrap <mark> (drop the tags, keep the children).
            working = MarkOpenRegex.Replace(working, "");
            working = MarkCloseRegex.Replace(working, "");

            // 3) Strip every remaining tag — this reproduces doc.body.textContent,
            //    which yields only the concatenated text of all descendant nodes.
            working = AnyTagRegex.Replace(working, "");

            // 4) Decode HTML entities (textContent returns decoded text), then trim.
            return WebUtility.HtmlDecode(working).Trim();
        }

        /// <summary>
        /// Rewrites the inner HTML of a single &lt;ruby&gt; element per pattern,
        /// returning the replacement for the whole &lt;ruby&gt;...&lt;/ruby&gt; block.
        ///   "rt"        -&gt; concatenation of each &lt;rt&gt; textContent (trimmed).
        ///   "ruby(rt)"  -&gt; base text with "(reading)" inserted before each &lt;rt&gt;
        ///                    when the reading is non-empty; &lt;rt&gt;/&lt;rp&gt; removed.
        ///   "ruby"      -&gt; base text only; &lt;rt&gt;/&lt;rp&gt; removed.
        /// Note that the surrounding tags are left for the final tag-strip pass; only
        /// the ruby-specific structure is handled here, exactly as the DOM code does.
        /// </summary>
        private static string RewriteRubyBlock(string inner, string pattern) {
            if (pattern == "rt") {
                // Collect each <rt> textContent, trimmed, joined with "". The TS reads
                // rt.textContent, so strip any inner tags within the <rt> too.
                StringBuilder sb = new StringBuilder();
                foreach (Match rt in RtElementRegex.Matches(inner)) {
                    string rtText = StripTags(rt.Groups["inner"].Value).Trim();
                    sb.Append(rtText);
                }
                return sb.ToString();
            }

            if (pattern == "ruby(rt)") {
                // Insert "(reading)" before each non-empty <rt>, then remove the <rt>.
                string replaced = RtElementRegex.Replace(inner, match => {
                    string rtText = StripTags(match.Groups["inner"].Value).Trim();
                    return rtText.Length > 0 ? "(" + rtText + ")" : "";
                });
                // Remove any <rp> elements entirely.
                replaced = RpElementRegex.Replace(replaced, "");
                return replaced;
            }

            // pattern == "ruby" (default): drop <rt> and <rp>, keep base text.
            string baseOnly = RtElementRegex.Replace(inner, "");
            baseOnly = RpElementRegex.Replace(baseOnly, "");
            return baseOnly;
        }

        private static string StripTags(string html) {
            if (string.IsNullOrEmpty(html)) {
                return "";
            }
            return AnyTagRegex.Replace(html, "");
        }

        /// <summary>
        /// Port of katakanaToHiragana: every code point in the katakana block
        /// U+30A1..U+30F6 is shifted down by 0x60 into the corresponding hiragana
        /// (U+3041..U+3096). All other characters pass through unchanged.
        ///
        /// The TS implementation does this arithmetically:
        ///   str.replace(/[ァ-ヶ]/g, m =&gt; String.fromCharCode(m.charCodeAt(0) - 0x60))
        /// The FULL mapping table (every code point in the range) is materialized below
        /// for an explicit, auditable 1:1 port; the arithmetic shift produces the same
        /// result for each entry.
        /// </summary>
        public static string KatakanaToHiragana(string str) {
            if (string.IsNullOrEmpty(str)) {
                return str;
            }
            StringBuilder sb = new StringBuilder(str.Length);
            foreach (char c in str) {
                if (c >= 'ァ' && c <= 'ヶ') {
                    char mapped;
                    if (KatakanaToHiraganaMap.TryGetValue(c, out mapped)) {
                        sb.Append(mapped);
                    } else {
                        // Defensive: every code point in the range is in the table,
                        // but fall back to the same arithmetic the TS uses.
                        sb.Append((char)(c - 0x60));
                    }
                } else {
                    sb.Append(c);
                }
            }
            return sb.ToString();
        }

        /// <summary>
        /// Full katakana -&gt; hiragana mapping for U+30A1..U+30F6 (86 entries),
        /// each katakana code point paired with its hiragana counterpart 0x60 below.
        /// Transcribed in full; no entries elided.
        /// </summary>
        private static readonly Dictionary<char, char> KatakanaToHiraganaMap = new Dictionary<char, char> {
            { 'ァ', 'ぁ' }, // ァ -> ぁ
            { 'ア', 'あ' }, // ア -> あ
            { 'ィ', 'ぃ' }, // ィ -> ぃ
            { 'イ', 'い' }, // イ -> い
            { 'ゥ', 'ぅ' }, // ゥ -> ぅ
            { 'ウ', 'う' }, // ウ -> う
            { 'ェ', 'ぇ' }, // ェ -> ぇ
            { 'エ', 'え' }, // エ -> え
            { 'ォ', 'ぉ' }, // ォ -> ぉ
            { 'オ', 'お' }, // オ -> お
            { 'カ', 'か' }, // カ -> か
            { 'ガ', 'が' }, // ガ -> が
            { 'キ', 'き' }, // キ -> き
            { 'ギ', 'ぎ' }, // ギ -> ぎ
            { 'ク', 'く' }, // ク -> く
            { 'グ', 'ぐ' }, // グ -> ぐ
            { 'ケ', 'け' }, // ケ -> け
            { 'ゲ', 'げ' }, // ゲ -> げ
            { 'コ', 'こ' }, // コ -> こ
            { 'ゴ', 'ご' }, // ゴ -> ご
            { 'サ', 'さ' }, // サ -> さ
            { 'ザ', 'ざ' }, // ザ -> ざ
            { 'シ', 'し' }, // シ -> し
            { 'ジ', 'じ' }, // ジ -> じ
            { 'ス', 'す' }, // ス -> す
            { 'ズ', 'ず' }, // ズ -> ず
            { 'セ', 'せ' }, // セ -> せ
            { 'ゼ', 'ぜ' }, // ゼ -> ぜ
            { 'ソ', 'そ' }, // ソ -> そ
            { 'ゾ', 'ぞ' }, // ゾ -> ぞ
            { 'タ', 'た' }, // タ -> た
            { 'ダ', 'だ' }, // ダ -> だ
            { 'チ', 'ち' }, // チ -> ち
            { 'ヂ', 'ぢ' }, // ヂ -> ぢ
            { 'ッ', 'っ' }, // ッ -> っ
            { 'ツ', 'つ' }, // ツ -> つ
            { 'ヅ', 'づ' }, // ヅ -> づ
            { 'テ', 'て' }, // テ -> て
            { 'デ', 'で' }, // デ -> で
            { 'ト', 'と' }, // ト -> と
            { 'ド', 'ど' }, // ド -> ど
            { 'ナ', 'な' }, // ナ -> な
            { 'ニ', 'に' }, // ニ -> に
            { 'ヌ', 'ぬ' }, // ヌ -> ぬ
            { 'ネ', 'ね' }, // ネ -> ね
            { 'ノ', 'の' }, // ノ -> の
            { 'ハ', 'は' }, // ハ -> は
            { 'バ', 'ば' }, // バ -> ば
            { 'パ', 'ぱ' }, // パ -> ぱ
            { 'ヒ', 'ひ' }, // ヒ -> ひ
            { 'ビ', 'び' }, // ビ -> び
            { 'ピ', 'ぴ' }, // ピ -> ぴ
            { 'フ', 'ふ' }, // フ -> ふ
            { 'ブ', 'ぶ' }, // ブ -> ぶ
            { 'プ', 'ぷ' }, // プ -> ぷ
            { 'ヘ', 'へ' }, // ヘ -> へ
            { 'ベ', 'べ' }, // ベ -> べ
            { 'ペ', 'ぺ' }, // ペ -> ぺ
            { 'ホ', 'ほ' }, // ホ -> ほ
            { 'ボ', 'ぼ' }, // ボ -> ぼ
            { 'ポ', 'ぽ' }, // ポ -> ぽ
            { 'マ', 'ま' }, // マ -> ま
            { 'ミ', 'み' }, // ミ -> み
            { 'ム', 'む' }, // ム -> む
            { 'メ', 'め' }, // メ -> め
            { 'モ', 'も' }, // モ -> も
            { 'ャ', 'ゃ' }, // ャ -> ゃ
            { 'ヤ', 'や' }, // ヤ -> や
            { 'ュ', 'ゅ' }, // ュ -> ゅ
            { 'ユ', 'ゆ' }, // ユ -> ゆ
            { 'ョ', 'ょ' }, // ョ -> ょ
            { 'ヨ', 'よ' }, // ヨ -> よ
            { 'ラ', 'ら' }, // ラ -> ら
            { 'リ', 'り' }, // リ -> り
            { 'ル', 'る' }, // ル -> る
            { 'レ', 'れ' }, // レ -> れ
            { 'ロ', 'ろ' }, // ロ -> ろ
            { 'ヮ', 'ゎ' }, // ヮ -> ゎ
            { 'ワ', 'わ' }, // ワ -> わ
            { 'ヰ', 'ゐ' }, // ヰ -> ゐ
            { 'ヱ', 'ゑ' }, // ヱ -> ゑ
            { 'ヲ', 'を' }, // ヲ -> を
            { 'ン', 'ん' }, // ン -> ん
            { 'ヴ', 'ゔ' }, // ヴ -> ゔ
            { 'ヵ', 'ゕ' }, // ヵ -> ゕ
            { 'ヶ', 'ゖ' }  // ヶ -> ゖ
        };
    }
}
