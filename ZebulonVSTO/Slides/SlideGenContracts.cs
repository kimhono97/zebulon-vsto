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

        /// <summary>Snapshot the active deck's (default master) layouts.</summary>
        List<LayoutDescriptor> ReadActiveLayouts();

        /// <summary>Insert the planned slides; returns the number inserted.</summary>
        int ExecutePlan(LayoutSelection selection, List<SlidePlanItem> plan, InsertPosition position);
    }
}
