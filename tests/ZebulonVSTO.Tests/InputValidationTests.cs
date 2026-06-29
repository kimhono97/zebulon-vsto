using Xunit;
using ZebulonVSTO.Sync;

namespace ZebulonVSTO.Tests {
    public class InputValidationTests {
        [Theory]
        [InlineData("8291", true, 8291)]
        [InlineData("1", true, 1)]
        [InlineData("65535", true, 65535)]
        [InlineData("0", false, 0)]
        [InlineData("65536", false, 0)]
        [InlineData("-1", false, 0)]
        [InlineData("abc", false, 0)]
        [InlineData("", false, 0)]
        [InlineData(null, false, 0)]
        public void TryParsePort_EnforcesRange(string input, bool expectedValid, int expectedPort) {
            int port;
            bool valid = InputValidation.TryParsePort(input, out port);
            Assert.Equal(expectedValid, valid);
            if (expectedValid) {
                Assert.Equal(expectedPort, port);
            }
        }

        [Theory]
        [InlineData("192.168.0.1")]
        [InlineData("10.0.0.255")]
        [InlineData("255.255.255.255")]
        [InlineData("127.0.0.1")]
        public void IsValidIPv4_AcceptsDottedQuads(string input) {
            Assert.True(InputValidation.IsValidIPv4(input));
        }

        [Theory]
        [InlineData("")]
        [InlineData(null)]
        [InlineData("abc")]
        [InlineData("example.com")]
        [InlineData("1.2.3.4.5")]  // too many octets
        [InlineData("::1")]         // IPv6 rejected (transport is IPv4-only)
        [InlineData("2001:db8::1")]
        public void IsValidIPv4_RejectsNonIPv4(string input) {
            Assert.False(InputValidation.IsValidIPv4(input));
        }
    }
}
