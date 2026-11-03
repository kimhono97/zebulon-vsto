using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

namespace ZebulonVSTO.Slides {
    /// <summary>A 1-based inclusive verse range (mirrors web WordSection and the
    /// exporter word item's <c>sections[]</c>).</summary>
    [DataContract]
    public class WordSection {
        [DataMember(Name = "start")] public int Start { get; set; }
        [DataMember(Name = "end")] public int End { get; set; }

        public WordSection() { }
        public WordSection(int start, int end) {
            Start = start;
            End = end;
        }
    }

    /// <summary>
    /// A Word (scripture) item to insert, mirroring the fields the Exporter
    /// consumes from the web's WordListItem. First/Second/Third are
    /// [sectionIndex][verseIndex] plain-text verse lines. COM-free.
    /// </summary>
    public class WordItem {
        public string HeadText { get; set; }

        /// <summary>The structured verse header the exporter's word item prefers
        /// ({books, ref}); null when no book name is known. Its
        /// <see cref="VenueHeader.Format"/> equals <see cref="HeadText"/>.</summary>
        public VenueHeader Header { get; set; }

        /// <summary>Book name per language slot ("" when unknown), so the header
        /// can be rebuilt for only the slots a layout shows
        /// (<see cref="VenueItem.Word(WordItem, IList{bool})"/>).</summary>
        public List<string> SlotBooks { get; set; }
        public List<WordSection> Sections { get; set; }
        public List<List<string>> First { get; set; }
        public List<List<string>> Second { get; set; }
        public List<List<string>> Third { get; set; }
    }

    /// <summary>
    /// Response of GET /api/bible (jym-workbox): a chapter's verses. data[i] is
    /// verse (i+1)'s text (may contain HTML/ruby markup). COM-free; parsed with the
    /// framework DataContractJsonSerializer.
    /// </summary>
    [DataContract]
    public class BibleData {
        [DataMember(Name = "bookname")] public string BookName { get; set; }
        [DataMember(Name = "chapter")] public int Chapter { get; set; }
        [DataMember(Name = "data")] public List<string> Data { get; set; }
        [DataMember(Name = "error")] public string Error { get; set; }

        public static BibleData Parse(string json) {
            if (string.IsNullOrEmpty(json)) {
                return null;
            }
            try {
                using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(json))) {
                    DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(BibleData));
                    return serializer.ReadObject(ms) as BibleData;
                }
            } catch {
                return null;
            }
        }
    }
}
