using System.Text;

namespace ZebulonVSTO.Sync {
    /// <summary>
    /// The role a peer advertises in a discovery ANNOUNCE (or that a DISCOVER
    /// asks for via <c>want</c>). <see cref="Unknown"/> means absent/unparseable;
    /// as a <c>want</c> it means "any role may answer".
    /// </summary>
    public enum DiscoveryRole {
        Unknown,
        Receiver,  // running as RECEIVER — a valid sync target
        Sender,    // running as SENDER
        Idle       // present but not running (no sync socket bound yet)
    }

    /// <summary>
    /// The parsed discovery payload carried in the <c>Data</c> field of a
    /// DISCOVER/ANNOUNCE <see cref="SyncMessage"/>. Pure data — no COM/UI/network
    /// here, so it can be unit-tested in isolation (like <see cref="CommandParser"/>).
    ///
    /// Wire form: <c>ZSYNC1;key=value;key=value</c>. The first token is a frozen
    /// magic+version tag (<see cref="Magic"/>); payloads not starting with it are
    /// rejected (<see cref="Valid"/> = false) so foreign/old-protocol datagrams
    /// are ignored. Unknown keys are skipped (forward-compatible); absent keys
    /// fall back to defaults.
    ///
    /// Recognized keys: <c>role</c>/<c>host</c>/<c>ver</c> (ANNOUNCE),
    /// <c>q</c>/<c>want</c> (DISCOVER). Address/port are NOT carried here — they
    /// live in the enclosing message's <c>SenderIP</c>/<c>SenderPort</c> so there
    /// is a single source of truth for "where to reach me for sync".
    /// </summary>
    public sealed class DiscoveryPayload {
        public const string Magic = "ZSYNC1";

        private const string RoleReceiver = "RECEIVER";
        private const string RoleSender = "SENDER";
        private const string RoleIdle = "IDLE";

        /// <summary>True when the payload carried the expected magic tag.</summary>
        public bool Valid { get; }

        // ANNOUNCE fields
        public DiscoveryRole Role { get; }
        public string Host { get; }
        public string Version { get; }

        // DISCOVER fields
        public bool IsQuery { get; }
        /// <summary>Requested role filter; <see cref="DiscoveryRole.Unknown"/> = any.</summary>
        public DiscoveryRole Want { get; }

        private DiscoveryPayload(bool valid, DiscoveryRole role, string host, string version,
                                 bool isQuery, DiscoveryRole want) {
            Valid = valid;
            Role = role;
            Host = host ?? string.Empty;
            Version = version ?? string.Empty;
            IsQuery = isQuery;
            Want = want;
        }

        public static readonly DiscoveryPayload Invalid =
            new DiscoveryPayload(false, DiscoveryRole.Unknown, string.Empty, string.Empty,
                                 false, DiscoveryRole.Unknown);

        public static DiscoveryPayload Parse(string data) {
            if (string.IsNullOrEmpty(data)) {
                return Invalid;
            }
            string[] tokens = data.Split(';');
            if (tokens[0] != Magic) {
                return Invalid;
            }

            DiscoveryRole role = DiscoveryRole.Unknown;
            string host = string.Empty;
            string version = string.Empty;
            bool isQuery = false;
            DiscoveryRole want = DiscoveryRole.Unknown;

            for (int i = 1; i < tokens.Length; i++) {
                string token = tokens[i];
                int eq = token.IndexOf('=');
                if (eq <= 0) {
                    continue; // empty or malformed token — skip
                }
                string key = token.Substring(0, eq);
                string value = token.Substring(eq + 1);
                switch (key) {
                    case "role": role = RoleFromWire(value); break;
                    case "host": host = value; break;
                    case "ver": version = value; break;
                    case "want": want = RoleFromWire(value); break;
                    case "q": isQuery = value == "1"; break;
                    // unknown keys ignored (forward-compatible)
                }
            }
            return new DiscoveryPayload(true, role, host, version, isQuery, want);
        }

        /// <summary>Build a DISCOVER ping payload. <paramref name="want"/> =
        /// <see cref="DiscoveryRole.Unknown"/> omits the filter (any role answers).</summary>
        public static string BuildDiscover(DiscoveryRole want) {
            StringBuilder sb = new StringBuilder(Magic);
            sb.Append(";q=1");
            if (want != DiscoveryRole.Unknown) {
                sb.Append(";want=").Append(RoleToWire(want));
            }
            return sb.ToString();
        }

        /// <summary>Build an ANNOUNCE reply payload.</summary>
        public static string BuildAnnounce(DiscoveryRole role, string host, string version) {
            StringBuilder sb = new StringBuilder(Magic);
            sb.Append(";role=").Append(RoleToWire(role));
            sb.Append(";host=").Append(Sanitize(host));
            sb.Append(";ver=").Append(Sanitize(version));
            return sb.ToString();
        }

        private static DiscoveryRole RoleFromWire(string value) {
            // Case-insensitive via ToUpperInvariant (guards against culture-
            // sensitive casing, e.g. Turkish 'I'); the wire form is upper-case.
            switch ((value ?? string.Empty).ToUpperInvariant()) {
                case RoleReceiver: return DiscoveryRole.Receiver;
                case RoleSender: return DiscoveryRole.Sender;
                case RoleIdle: return DiscoveryRole.Idle;
                default: return DiscoveryRole.Unknown;
            }
        }

        private static string RoleToWire(DiscoveryRole role) {
            switch (role) {
                case DiscoveryRole.Receiver: return RoleReceiver;
                case DiscoveryRole.Sender: return RoleSender;
                case DiscoveryRole.Idle: return RoleIdle;
                default: return string.Empty;
            }
        }

        // host/version are machine-supplied; strip our delimiters defensively so
        // a stray ';' or '=' can't corrupt the payload grammar on parse.
        private static string Sanitize(string value) {
            if (string.IsNullOrEmpty(value)) {
                return string.Empty;
            }
            return value.Replace(';', ' ').Replace('=', ' ');
        }
    }
}
