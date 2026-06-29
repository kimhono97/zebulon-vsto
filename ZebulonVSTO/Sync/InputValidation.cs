using System.Net;
using System.Net.Sockets;

namespace ZebulonVSTO.Sync {
    /// <summary>
    /// Pure input validators for connection settings, shared by the setup wizard
    /// (and previously inlined in the ribbon's edit-box handlers). COM/UI-free so
    /// they can be unit-tested in isolation.
    /// </summary>
    public static class InputValidation {
        /// <summary>True if <paramref name="text"/> is an integer in 1–65535;
        /// the parsed value is returned via <paramref name="port"/>.</summary>
        public static bool TryParsePort(string text, out int port) {
            return int.TryParse(text, out port) && port >= 1 && port <= 65535;
        }

        public static bool IsValidPort(string text) {
            int port;
            return TryParsePort(text, out port);
        }

        /// <summary>True if <paramref name="text"/> is a well-formed IPv4 literal.
        /// IPv6 is rejected (the transport is IPv4-only).</summary>
        public static bool IsValidIPv4(string text) {
            IPAddress address;
            return IPAddress.TryParse(text, out address)
                && address.AddressFamily == AddressFamily.InterNetwork;
        }
    }
}
