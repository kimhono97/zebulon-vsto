using Xunit;
using ZebulonVSTO.Sync;

namespace ZebulonVSTO.Tests {
    public class CommandParserTests {
        [Fact]
        public void Alert_WithText_CapturesMessage() {
            ParsedCommand command = CommandParser.Parse("alert hello world");
            Assert.Equal(CommandKind.Alert, command.Kind);
            Assert.True(command.IsValid);
            Assert.Equal("hello world", command.Text);
        }

        [Fact]
        public void Alert_WithoutText_IsStillValidWithEmptyText() {
            ParsedCommand command = CommandParser.Parse("alert");
            Assert.Equal(CommandKind.Alert, command.Kind);
            Assert.Equal(string.Empty, command.Text);
        }

        [Theory]
        [InlineData("select 3", CommandKind.Select, 3)]
        [InlineData("showslide 5", CommandKind.ShowSlide, 5)]
        [InlineData("select -2", CommandKind.Select, -2)]
        public void IndexedCommands_ParseSlideIndex(string input, CommandKind expectedKind, int expectedIndex) {
            ParsedCommand command = CommandParser.Parse(input);
            Assert.Equal(expectedKind, command.Kind);
            Assert.True(command.IsValid);
            Assert.Equal(expectedIndex, command.SlideIndex);
        }

        [Fact]
        public void HideSlide_Parses() {
            ParsedCommand command = CommandParser.Parse("hideslide");
            Assert.Equal(CommandKind.HideSlide, command.Kind);
            Assert.True(command.IsValid);
        }

        [Theory]
        [InlineData("SELECT 2", CommandKind.Select)]
        [InlineData("HideSlide", CommandKind.HideSlide)]
        [InlineData("HIDESLIDE", CommandKind.HideSlide)] // guards against culture-sensitive ToLower (Turkish 'I')
        [InlineData("ShowSlide 1", CommandKind.ShowSlide)]
        public void Verbs_AreCaseInsensitive(string input, CommandKind expectedKind) {
            ParsedCommand command = CommandParser.Parse(input);
            Assert.Equal(expectedKind, command.Kind);
        }

        [Theory]
        // Full-width letters/digits and an ideographic space (U+3000) — what an
        // IME in full-width mode produces — must still parse via NFKC folding.
        [InlineData("ｓｅｌｅｃｔ　２", CommandKind.Select, 2)]      // "ｓｅｌｅｃｔ　２"
        [InlineData("ｓｈｏｗｓｌｉｄｅ　12", CommandKind.ShowSlide, 12)]
        public void FullWidth_InputIsFolded(string input, CommandKind expectedKind, int expectedIndex) {
            ParsedCommand command = CommandParser.Parse(input);
            Assert.Equal(expectedKind, command.Kind);
            Assert.True(command.IsValid);
            Assert.Equal(expectedIndex, command.SlideIndex);
        }

        [Fact]
        public void FullWidth_HideSlide_IsFolded() {
            // "ＨＩＤＥＳＬＩＤＥ" (full-width, upper-case) -> hideslide
            ParsedCommand command = CommandParser.Parse("ＨＩＤＥＳＬＩＤＥ");
            Assert.Equal(CommandKind.HideSlide, command.Kind);
        }

        [Theory]
        [InlineData("select abc")]   // non-integer argument
        [InlineData("select")]       // missing argument
        [InlineData("showslide x")]
        [InlineData("bogus 1")]      // unknown verb
        [InlineData("")]             // empty
        [InlineData(null)]           // null
        [InlineData(" select 2")]    // leading space => verb is unrecognized (preserves original behavior)
        public void MalformedOrUnknown_IsInvalid(string input) {
            ParsedCommand command = CommandParser.Parse(input);
            Assert.Equal(CommandKind.Unknown, command.Kind);
            Assert.False(command.IsValid);
        }
    }
}
