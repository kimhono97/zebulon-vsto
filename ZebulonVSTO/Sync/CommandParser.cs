using System.Text;

namespace ZebulonVSTO.Sync {
    /// <summary>
    /// The kind of action a received command requests. <see cref="Unknown"/>
    /// represents an unrecognized or malformed command.
    /// </summary>
    public enum CommandKind {
        Unknown,
        Alert,
        Select,
        ShowSlide,
        HideSlide
    }

    /// <summary>
    /// The result of parsing a command string. Pure data — no COM/UI here.
    /// </summary>
    public sealed class ParsedCommand {
        public CommandKind Kind { get; }

        /// <summary>Message text for <see cref="CommandKind.Alert"/>; empty otherwise.</summary>
        public string Text { get; }

        /// <summary>Slide index for Select/ShowSlide; 0 otherwise.</summary>
        public int SlideIndex { get; }

        /// <summary>True when the command was recognized and its argument is well-formed.</summary>
        public bool IsValid { get { return Kind != CommandKind.Unknown; } }

        private ParsedCommand(CommandKind kind, string text, int slideIndex) {
            Kind = kind;
            Text = text;
            SlideIndex = slideIndex;
        }

        public static readonly ParsedCommand Invalid =
            new ParsedCommand(CommandKind.Unknown, string.Empty, 0);

        public static ParsedCommand Alert(string text) {
            return new ParsedCommand(CommandKind.Alert, text ?? string.Empty, 0);
        }
        public static ParsedCommand Indexed(CommandKind kind, int slideIndex) {
            return new ParsedCommand(kind, string.Empty, slideIndex);
        }
        public static ParsedCommand HideSlide() {
            return new ParsedCommand(CommandKind.HideSlide, string.Empty, 0);
        }
    }

    /// <summary>
    /// Translates a raw command string (the <c>Data</c> field of a request
    /// message) into a structured <see cref="ParsedCommand"/>. Stateless and
    /// side-effect free so it can be unit-tested without PowerPoint.
    ///
    /// Grammar: <c>&lt;verb&gt; [argument]</c>, verb matched case-insensitively.
    /// Recognized verbs: <c>alert &lt;text&gt;</c>, <c>select &lt;n&gt;</c>,
    /// <c>showslide &lt;n&gt;</c>, <c>hideslide</c>.
    /// </summary>
    public static class CommandParser {
        public static ParsedCommand Parse(string data) {
            if (string.IsNullOrEmpty(data)) {
                return ParsedCommand.Invalid;
            }

            // Fold full-width / compatibility characters to their ASCII forms.
            // An IME left in full-width mode turns "select 2" into the wider
            // "ｓｅｌｅｃｔ　２" (incl. an ideographic space U+3000); NFKC maps those
            // back so the command parses regardless of how it was typed/sent.
            data = data.Normalize(NormalizationForm.FormKC);

            string verb;
            string argument;
            int separator = data.IndexOf(' ');
            if (separator > 0) {
                verb = data.Substring(0, separator).ToLowerInvariant();
                argument = data.Length > separator + 1 ? data.Substring(separator + 1) : string.Empty;
            } else {
                verb = data.ToLowerInvariant();
                argument = string.Empty;
            }

            switch (verb) {
                case "alert":
                    return ParsedCommand.Alert(argument);
                case "select":
                    return ParseIndexed(CommandKind.Select, argument);
                case "showslide":
                    return ParseIndexed(CommandKind.ShowSlide, argument);
                case "hideslide":
                    return ParsedCommand.HideSlide();
                default:
                    return ParsedCommand.Invalid;
            }
        }

        private static ParsedCommand ParseIndexed(CommandKind kind, string argument) {
            int index;
            if (int.TryParse(argument, out index)) {
                return ParsedCommand.Indexed(kind, index);
            }
            return ParsedCommand.Invalid;
        }
    }
}
