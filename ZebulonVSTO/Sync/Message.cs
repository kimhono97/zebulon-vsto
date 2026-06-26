using System;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

namespace ZebulonVSTO.Sync {
    public enum MessageType {
        // Existing sync types — order is part of the frozen wire contract.
        CUSTOM, REQUEST, RESPONSE,
        // Discovery types (appended; never reorder the values above).
        DISCOVER, ANNOUNCE
    }

    /// <summary>
    /// Wire DTO exchanged between sync peers as UTF-8 JSON.
    ///
    /// IMPORTANT: the JSON member names (SenderIP, SenderPort, ID, Type, Data)
    /// are a frozen on-the-wire contract shared with every other instance.
    /// They are pinned explicitly via <see cref="DataMemberAttribute.Name"/>;
    /// do NOT rename them. Serialization uses the framework
    /// <see cref="DataContractJsonSerializer"/> (no third-party dependency).
    /// </summary>
    [DataContract]
    public class SyncMessage {
        [DataMember(Name = "SenderIP")]
        public string SenderIP { get; set; }

        [DataMember(Name = "SenderPort")]
        public int SenderPort { get; set; }

        [DataMember(Name = "ID")]
        public int ID { get; set; }

        [DataMember(Name = "Type")]
        public int Type { get; set; }

        [DataMember(Name = "Data")]
        public string Data { get; set; }

        // A new serializer per call: DataContractJsonSerializer instance members
        // are not guaranteed thread-safe, and ToJsonString may run on both the
        // UI thread (sending a request) and the receive thread (sending a
        // response) at once. Contract metadata is cached by the framework per
        // type, so construction is cheap.
        private static DataContractJsonSerializer CreateSerializer() {
            return new DataContractJsonSerializer(typeof(SyncMessage));
        }

        public SyncMessage() {
            ID = -1;
            Type = -1;
            SenderIP = SyncDefaults.Localhost;
            SenderPort = SyncDefaults.Port;
            Data = "";
        }
        public SyncMessage(int id, string senderIP, int senderPort, MessageType type, string data) {
            ID = id;
            Type = (int)type;
            SenderIP = senderIP;
            SenderPort = senderPort;
            Data = data;
        }

        // DataContractJsonSerializer constructs instances WITHOUT running a
        // constructor, so members absent from a (malformed/partial) datagram
        // would otherwise land on CLR defaults (null SenderIP, 0 port). Seed
        // the same fallbacks the constructor uses before members are populated,
        // so present members override and absent ones keep sane values — this
        // matches the previous serializer's behavior and keeps Data non-null.
        [OnDeserializing]
        private void OnDeserializing(StreamingContext context) {
            ID = -1;
            Type = -1;
            SenderIP = SyncDefaults.Localhost;
            SenderPort = SyncDefaults.Port;
            Data = "";
        }

        public string ToJsonString() {
            try {
                using (var stream = new MemoryStream()) {
                    CreateSerializer().WriteObject(stream, this);
                    return Encoding.UTF8.GetString(stream.ToArray());
                }
            } catch (Exception e) {
                Console.WriteLine(e.ToString());
                return null;
            }
        }

        public bool IsResponse() {
            return Type == (int)MessageType.RESPONSE;
        }

        public static SyncMessage Parse(string json) {
            if (string.IsNullOrEmpty(json)) {
                return null;
            }
            try {
                using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(json))) {
                    return (SyncMessage)CreateSerializer().ReadObject(stream);
                }
            } catch (Exception e) {
                Console.WriteLine(e.ToString());
                return null;
            }
        }
    }
}
