namespace ZebulonVSTO.Sync {
    /// <summary>
    /// Sink for sync traffic/diagnostic logging. Implemented by the add-in host
    /// (ThisAddIn) so <see cref="SyncManager"/> need not reach into Globals.
    /// </summary>
    public interface ISyncLogger {
        void Log(string line);
        void LogError(string message, System.Exception error);
    }

    /// <summary>
    /// Notified when the sync status the ribbon shows may have changed (e.g. the
    /// RECEIVER saw traffic from a new sender). Implemented by the add-in host,
    /// which marshals a ribbon refresh onto the UI thread. Called from the
    /// receive thread, so implementations must not touch COM/UI directly.
    /// </summary>
    public interface IStatusObserver {
        void OnPeerChanged();
    }

    /// <summary>
    /// Executes received slide-navigation commands against PowerPoint. All
    /// implementations must marshal Interop calls onto the UI thread.
    /// Implemented by the add-in host (ThisAddIn).
    /// </summary>
    public interface ISlideController {
        bool SelectSlide(int slideIndex);
        bool ShowSlide(int slideIndex);
        bool HideSlide();

        /// <summary>Show an alert dialog originating from <paramref name="sender"/>.</summary>
        bool Alert(string sender, string text);
    }
}
