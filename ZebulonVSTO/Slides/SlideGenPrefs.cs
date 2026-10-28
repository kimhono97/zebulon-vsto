using System.Runtime.Serialization;

namespace ZebulonVSTO.Slides {
    /// <summary>
    /// SlideGenWindow's remembered options (the <c>slideGen</c> section of
    /// <see cref="ZebulonVSTO.Preferences"/>). COM-free; unit-tested.
    /// </summary>
    [DataContract]
    public sealed class SlideGenPrefs : IExtensibleDataObject {
        /// <summary>Also list the legacy PRAISE_/WORD_ templates (default: VENUE only).</summary>
        [DataMember(Name = "showLegacyTemplates", EmitDefaultValue = false)]
        public bool ShowLegacyTemplates { get; set; }

        public ExtensionDataObject ExtensionData { get; set; }
    }
}
