using System;
using System.Text.Json;

namespace ZebulonVSTO.Sync {
    public enum MessageType {
        CUSTOM, REQUEST, RESPONSE
    }

    public class SyncMessage {
        public string SenderIP { get; set; }
        public int SenderPort { get; set; }
        public int ID { get; set; }
        public int Type { get; set; }
        public string Data { get; set; }

        public SyncMessage() {
            ID = -1;
            Type = -1;
            SenderIP = SyncManager.DEFAULT_LOCAL;
            SenderPort = SyncManager.DEFAULT_PORT;
            Data = "";
        }
        public SyncMessage(int nID, string strSenderIP, int nSenderPort, MessageType nType, string strData) {
            ID = nID;
            Type = (int)nType;
            SenderIP = strSenderIP;
            SenderPort = nSenderPort;
            Data = strData;
        }

        public string ToJSONString() {
            try {
                return JsonSerializer.Serialize(this);
            } catch (Exception e) {
                Console.WriteLine(e.ToString());
                return null;
            }
        }
        public bool IsResonse() {
            return Type == (uint)MessageType.RESPONSE;
        }

        public static SyncMessage Parse(string strData) {
            try {
                return JsonSerializer.Deserialize<SyncMessage>(strData);
            } catch (Exception e) {
                Console.WriteLine(e.ToString());
                return null;
            }
        } 
    }
}
