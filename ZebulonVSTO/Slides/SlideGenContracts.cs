using System.Collections.Generic;

namespace ZebulonVSTO.Slides {
    /// <summary>
    /// Host-implemented seam (by <c>ThisAddIn</c>) that performs the PowerPoint
    /// Interop half of slide generation, keeping the dialog and pure logic
    /// COM-free. All members marshal onto the add-in UI thread (mirrors
    /// <see cref="ZebulonVSTO.Sync.ISlideController"/>).
    /// </summary>
    public interface ISlideBuilder {
        /// <summary>True if a presentation is open to insert into.</summary>
        bool HasActivePresentation();

        /// <summary>Snapshot a deck's (default master) layouts. Pass a presentation
        /// FullName to target a specific deck, or null/"" for the active one.</summary>
        List<LayoutDescriptor> ReadLayouts(string presentationFullName);

        /// <summary>Insert the planned slides into the deck identified by
        /// <paramref name="presentationFullName"/> (null/"" = active); returns the
        /// number inserted.</summary>
        int ExecutePlan(string presentationFullName, LayoutSelection selection, List<SlidePlanItem> plan, InsertPosition position);

        /// <summary>Open a .pptx from disk (with a window). Returns the opened
        /// deck's FullName (to target it later), or "" on failure.</summary>
        string OpenPresentation(string path);

        /// <summary>The active deck's folder, or "" if none / unsaved. Lets the
        /// COM-free dialog default its save directory.</summary>
        string GetActivePresentationFolder();
    }
}
