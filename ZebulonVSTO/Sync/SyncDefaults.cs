namespace ZebulonVSTO.Sync {
    /// <summary>
    /// Shared default values for the sync transport. Kept free of any
    /// PowerPoint/COM dependency so the pure data types (SyncMessage,
    /// CommandParser) can be unit-tested in isolation.
    /// </summary>
    public static class SyncDefaults {
        public const string Broadcast = "255.255.255.255";
        public const string Localhost = "127.0.0.1";
        public const int Port = 8291;

        // Dedicated discovery port for the always-on responder (peer auto-detect).
        // Kept separate from the sync Port so the responder can listen from
        // add-in startup, independent of the start/stop-locked sync socket.
        public const int DiscoveryPort = 8290;
    }
}
