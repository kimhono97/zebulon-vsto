using Xunit;
using ZebulonVSTO.Sync;

namespace ZebulonVSTO.Tests {
    public class SyncMessageTests {
        [Fact]
        public void RoundTrip_PreservesAllFields() {
            SyncMessage original = new SyncMessage(7, "192.168.0.10", 8291, MessageType.REQUEST, "select 2");

            SyncMessage parsed = SyncMessage.Parse(original.ToJsonString());

            Assert.NotNull(parsed);
            Assert.Equal(7, parsed.ID);
            Assert.Equal("192.168.0.10", parsed.SenderIP);
            Assert.Equal(8291, parsed.SenderPort);
            Assert.Equal((int)MessageType.REQUEST, parsed.Type);
            Assert.Equal("select 2", parsed.Data);
        }

        /// <summary>
        /// The serialized member names are a frozen on-the-wire contract shared
        /// with every other peer instance. If this test fails, sync between
        /// instances is broken — do not "fix" it by changing the assertion.
        /// </summary>
        [Theory]
        [InlineData("\"SenderIP\"")]
        [InlineData("\"SenderPort\"")]
        [InlineData("\"ID\"")]
        [InlineData("\"Type\"")]
        [InlineData("\"Data\"")]
        public void Serialized_ContainsFrozenWireFieldName(string expectedFieldName) {
            string json = new SyncMessage(1, "127.0.0.1", 8291, MessageType.RESPONSE, "Success").ToJsonString();
            Assert.Contains(expectedFieldName, json);
        }

        [Fact]
        public void Parse_AcceptsExternallyShapedDatagram() {
            // Field order differs from our serializer output; must still parse by name.
            string external = "{\"SenderIP\":\"127.0.0.1\",\"SenderPort\":8291,\"ID\":1,\"Type\":1,\"Data\":\"select 2\"}";

            SyncMessage parsed = SyncMessage.Parse(external);

            Assert.NotNull(parsed);
            Assert.Equal(1, parsed.ID);
            Assert.Equal(1, parsed.Type);
            Assert.Equal("select 2", parsed.Data);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("not json")]
        [InlineData("{ broken")]
        public void Parse_ReturnsNullOnBadInput_WithoutThrowing(string input) {
            Assert.Null(SyncMessage.Parse(input));
        }

        [Theory]
        [InlineData(MessageType.RESPONSE, true)]
        [InlineData(MessageType.REQUEST, false)]
        [InlineData(MessageType.CUSTOM, false)]
        public void IsResponse_ReflectsType(MessageType type, bool expected) {
            SyncMessage message = new SyncMessage(1, "127.0.0.1", 8291, type, "");
            Assert.Equal(expected, message.IsResponse());
        }

        /// <summary>
        /// The integer Type encoding (CUSTOM=0/REQUEST=1/RESPONSE=2) is part of
        /// the frozen wire contract; reordering the enum would silently break
        /// interop. This locks it alongside the field-name guard.
        /// </summary>
        [Theory]
        [InlineData(MessageType.CUSTOM, "\"Type\":0")]
        [InlineData(MessageType.REQUEST, "\"Type\":1")]
        [InlineData(MessageType.RESPONSE, "\"Type\":2")]
        [InlineData(MessageType.DISCOVER, "\"Type\":3")]  // appended for discovery; must stay 3
        [InlineData(MessageType.ANNOUNCE, "\"Type\":4")]  // appended for discovery; must stay 4
        public void Serialized_PinsNumericTypeEncoding(MessageType type, string expectedFragment) {
            string json = new SyncMessage(1, "127.0.0.1", 8291, type, "x").ToJsonString();
            Assert.Contains(expectedFragment, json);
        }

        [Fact]
        public void Parse_PartialDatagram_FallsBackToConstructorDefaults() {
            // Missing SenderIP/SenderPort. DataContractJsonSerializer skips the
            // constructor, so the [OnDeserializing] hook must restore the
            // loopback/port defaults rather than leaving null/0.
            SyncMessage parsed = SyncMessage.Parse("{\"ID\":7,\"Type\":1,\"Data\":\"select 2\"}");

            Assert.NotNull(parsed);
            Assert.Equal(7, parsed.ID);
            Assert.Equal("select 2", parsed.Data);
            Assert.Equal("127.0.0.1", parsed.SenderIP);
            Assert.Equal(8291, parsed.SenderPort);
        }

        [Fact]
        public void Parse_DatagramMissingData_YieldsEmptyStringNotNull() {
            SyncMessage parsed = SyncMessage.Parse("{\"SenderIP\":\"10.0.0.1\",\"SenderPort\":9000,\"ID\":3,\"Type\":2}");

            Assert.NotNull(parsed);
            Assert.Equal(string.Empty, parsed.Data);
        }
    }
}
