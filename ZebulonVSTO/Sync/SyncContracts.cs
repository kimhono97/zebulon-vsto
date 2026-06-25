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
