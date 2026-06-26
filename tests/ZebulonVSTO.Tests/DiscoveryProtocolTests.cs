using Xunit;
using ZebulonVSTO.Sync;

namespace ZebulonVSTO.Tests {
    public class DiscoveryProtocolTests {
        [Fact]
        public void Announce_RoundTrips_RoleHostVersion() {
            string wire = DiscoveryPayload.BuildAnnounce(DiscoveryRole.Receiver, "PC-A", "1.2.0");

            DiscoveryPayload parsed = DiscoveryPayload.Parse(wire);

            Assert.True(parsed.Valid);
            Assert.Equal(DiscoveryRole.Receiver, parsed.Role);
            Assert.Equal("PC-A", parsed.Host);
            Assert.Equal("1.2.0", parsed.Version);
        }

        [Fact]
        public void Discover_WithWant_RoundTrips() {
            string wire = DiscoveryPayload.BuildDiscover(DiscoveryRole.Receiver);

            DiscoveryPayload parsed = DiscoveryPayload.Parse(wire);

            Assert.True(parsed.Valid);
            Assert.True(parsed.IsQuery);
            Assert.Equal(DiscoveryRole.Receiver, parsed.Want);
        }

        [Fact]
        public void Discover_WithoutWant_OmitsFilter() {
            string wire = DiscoveryPayload.BuildDiscover(DiscoveryRole.Unknown);

            DiscoveryPayload parsed = DiscoveryPayload.Parse(wire);

            Assert.True(parsed.IsQuery);
            Assert.Equal(DiscoveryRole.Unknown, parsed.Want); // "any role may answer"
        }

        /// <summary>
        /// The exact wire strings are the cross-instance discovery contract.
        /// Pinning them here catches accidental format drift (a peer on the old
        /// format would silently never match), mirroring the SyncMessage
        /// frozen-field guard. The magic tag <c>ZSYNC1</c> namespaces our traffic.
        /// </summary>
        [Fact]
        public void WireFormat_IsFrozen() {
            Assert.Equal("ZSYNC1", DiscoveryPayload.Magic);
            Assert.Equal("ZSYNC1;q=1", DiscoveryPayload.BuildDiscover(DiscoveryRole.Unknown));
            Assert.Equal("ZSYNC1;q=1;want=RECEIVER", DiscoveryPayload.BuildDiscover(DiscoveryRole.Receiver));
            Assert.Equal("ZSYNC1;role=RECEIVER;host=PC-A;ver=1.2.0",
                DiscoveryPayload.BuildAnnounce(DiscoveryRole.Receiver, "PC-A", "1.2.0"));
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("NOPE;role=RECEIVER")]   // wrong magic
        [InlineData("zsync1;role=RECEIVER")] // magic is case-sensitive
        [InlineData("select 2")]             // a sync command, not discovery
        public void Parse_RejectsWithoutMagic(string input) {
            DiscoveryPayload parsed = DiscoveryPayload.Parse(input);
            Assert.False(parsed.Valid);
            Assert.Equal(DiscoveryRole.Unknown, parsed.Role);
        }

        [Fact]
        public void Parse_UnknownKeys_AreIgnored() {
            // Forward-compatibility: a newer peer may add keys we don't know.
            DiscoveryPayload parsed = DiscoveryPayload.Parse("ZSYNC1;role=RECEIVER;future=x;host=PC-A");

            Assert.True(parsed.Valid);
            Assert.Equal(DiscoveryRole.Receiver, parsed.Role);
            Assert.Equal("PC-A", parsed.Host);
        }

        [Fact]
        public void Parse_MagicOnly_YieldsValidWithDefaults() {
            DiscoveryPayload parsed = DiscoveryPayload.Parse("ZSYNC1");

            Assert.True(parsed.Valid);
            Assert.Equal(DiscoveryRole.Unknown, parsed.Role);
            Assert.Equal(string.Empty, parsed.Host);
            Assert.Equal(string.Empty, parsed.Version);
            Assert.False(parsed.IsQuery);
            Assert.Equal(DiscoveryRole.Unknown, parsed.Want);
        }

        [Theory]
        [InlineData("RECEIVER", DiscoveryRole.Receiver)]
        [InlineData("receiver", DiscoveryRole.Receiver)] // case-insensitive (Turkish-I safe)
        [InlineData("SENDER", DiscoveryRole.Sender)]
        [InlineData("IDLE", DiscoveryRole.Idle)]
        [InlineData("BOGUS", DiscoveryRole.Unknown)]
        public void Parse_Role_IsCaseInsensitiveAndTolerant(string roleValue, DiscoveryRole expected) {
            DiscoveryPayload parsed = DiscoveryPayload.Parse("ZSYNC1;role=" + roleValue);
            Assert.Equal(expected, parsed.Role);
        }

        [Theory]
        [InlineData("1", true)]
        [InlineData("0", false)]
        [InlineData("", false)]
        public void Parse_QueryFlag(string qValue, bool expected) {
            DiscoveryPayload parsed = DiscoveryPayload.Parse("ZSYNC1;q=" + qValue);
            Assert.Equal(expected, parsed.IsQuery);
        }

        [Fact]
        public void Parse_SkipsMalformedTokens_ButStaysValid() {
            // empty token, key with no value position, value-less token — all skipped.
            DiscoveryPayload parsed = DiscoveryPayload.Parse("ZSYNC1;;role=SENDER;=oops;noequals");

            Assert.True(parsed.Valid);
            Assert.Equal(DiscoveryRole.Sender, parsed.Role);
        }

        [Theory]
        [InlineData("PC;A")] // delimiter in host must not corrupt the grammar
        [InlineData("PC=A")]
        public void Announce_SanitizesDelimitersInHost(string rawHost) {
            string wire = DiscoveryPayload.BuildAnnounce(DiscoveryRole.Receiver, rawHost, "1.0");

            DiscoveryPayload parsed = DiscoveryPayload.Parse(wire);

            // Sanitized to spaces, so host is preserved as one field and role/ver
            // are unaffected (no injected tokens).
            Assert.True(parsed.Valid);
            Assert.Equal("PC A", parsed.Host);
            Assert.Equal(DiscoveryRole.Receiver, parsed.Role);
            Assert.Equal("1.0", parsed.Version);
        }
    }
}
