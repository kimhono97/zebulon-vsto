using System;
using System.Net.Sockets;
using System.Net;
using System.Threading;
using System.Windows.Forms;
using System.Text;

namespace ZebulonVSTO.Sync {
    public class SyncManager {
        public static string DEFAULT_TARGET = "255.255.255.255";
        public static string DEFAULT_LOCAL = "127.0.0.1";
        public static int DEFAULT_PORT = 8291;

        public enum SyncMode {
            NONE, SENDER, RECEIVER
        }

        private static SyncManager pInstance = null;
        private static int NextMessageID = 0;

        private Thread pThread;
        private UdpClient pClient;

        private string strLocalIP;
        private string strRemoteIP;
        private int nLocalPort;
        private int nRemotePort;
        private SyncMode nSyncMode; 
        private bool bRunning;
        
        private SyncManager() {
            this.pThread = null;
            this.pClient = null;

            this.strLocalIP = FindLocalIPAddress();
            this.nLocalPort = DEFAULT_PORT;
            this.strRemoteIP = DEFAULT_TARGET;
            this.nRemotePort = DEFAULT_PORT;
            this.nSyncMode = SyncMode.NONE;
            this.bRunning = false;
        }
        public string LocalIP {
            get { return this.strLocalIP; }
            set { }
        }
        public int LocalPort {
            get { return this.nLocalPort; }
            set {
                if (!this.bRunning) {
                    this.nLocalPort = value;
                }
            }
        }
        public string RemoteIP {
            get { return this.strRemoteIP;  }
            set {
                if (!this.bRunning) {
                    this.strRemoteIP = value;
                }
            }
        }
        public int RemotePort {
            get { return this.nRemotePort; }
            set {
                if (!this.bRunning) {
                    this.nRemotePort = value;
                }
            }
        }
        public SyncMode Mode {
            get { return this.nSyncMode; }
            set {
                if (!this.bRunning) {
                    this.nSyncMode = value;
                }
            }
        }
        public bool IsRunning () {
            return this.bRunning;
        }

        public bool StartSync() {
            if (this.bRunning) {
                return false;
            }

            switch (this.nSyncMode) {
                case SyncMode.SENDER:
                case SyncMode.RECEIVER:
                    this.pClient = new UdpClient(this.nLocalPort);
                    this.pThread = new Thread(new ThreadStart(ReceiveMessage));
                    break;
                default:
                    MessageBox.Show("Invalid sync mode", "Warning");
                    return false;
            }
            this.pThread.Start();
            MessageBox.Show("Start Sync", "ZebulonVSTO");
            Globals.ThisAddIn.LogDebug("---> Start Sync (Mode: " + this.nSyncMode.ToString() + ")");
            this.bRunning = true;
            return true;
        }

        public void StopSync() {
            if (!this.bRunning) {
                return;
            }

            if (this.pThread != null) {
                this.pThread.Abort();
                this.pThread = null;
            }
            if (this.pClient != null) {
                this.pClient.Close();
                this.pClient = null;
            }

            MessageBox.Show("Stop Sync", "ZebulonVSTO");
            Globals.ThisAddIn.LogDebug("<--- Stop Sync");
            this.bRunning = false;
        }

        private void ReceiveMessage() {
            IPEndPoint pEndPoint = new IPEndPoint(IPAddress.Parse(this.strLocalIP), this.nLocalPort);

            try {
                while (this.pClient != null) {
                    byte[] pBytes = this.pClient.Receive(ref pEndPoint);
                    string strJson = Encoding.UTF8.GetString(pBytes);
                    SyncMessage pMsg = SyncMessage.Parse(strJson);
                    if (pMsg != null) {
                        if (pMsg.IsResonse()) {
                            LogMessage("▶RES", pMsg);
                        } else if (this.nSyncMode == SyncMode.RECEIVER) {
                            LogMessage("▶REQ", pMsg);
                            bool bRet = ProcessRequest(pMsg);
                            SendResponse(pMsg, bRet);
                        }
                    }
                }
            } catch (ThreadAbortException e) {
                if (this.pClient != null) {
                    this.pClient.Close();
                    this.pClient = null;
                }
                Console.WriteLine(e.ToString());
            } catch (Exception e) {
                Globals.ThisAddIn.LogError("Failed to Receive Sync Messages.", e);
            }
        }
        private bool SendMessage(SyncMessage pMsg) {
            if (this.pClient == null) {
                return false;
            }

            try {
                string strJson = pMsg.ToJSONString();
                byte[] pBytes = Encoding.UTF8.GetBytes(strJson);
                this.pClient.Send(pBytes, pBytes.Length, this.strRemoteIP, this.nRemotePort);
                LogMessage("◀REQ", pMsg);
            } catch (Exception e) {
                Globals.ThisAddIn.LogError("Failed to Send a Sync Message.", e);
                return false;
            }
            return true;
        }
        private bool SendResponse(SyncMessage pReceived, bool bResult) {
            if (this.pClient == null) {
                return false;
            }

            try {
                SyncMessage pMsg = new SyncMessage(pReceived.ID, this.strLocalIP, this.nLocalPort, MessageType.RESPONSE, bResult ? "Success" : "Failed");
                string strJson = pMsg.ToJSONString();
                byte[] pBytes = Encoding.UTF8.GetBytes(strJson);
                this.pClient.Send(pBytes, pBytes.Length, pReceived.SenderIP, pReceived.SenderPort);
                LogMessage("◀RES", pMsg);
            } catch (Exception e) {
                Globals.ThisAddIn.LogError("Failed to Send a Sync Response Message.", e);
                return false;
            }

            return true;
        }
        private SyncMessage CreateSenderMessage(MessageType nType, string strData) {
            return new SyncMessage(NextMessageID++, this.strLocalIP, this.nLocalPort, nType, strData);
        }
        public bool SendRequestMessage(string strData, bool bIsCustom=false) {
            if (this.nSyncMode != SyncMode.SENDER && (!Globals.ThisAddIn.DEBUG_MODE || !bIsCustom)) {
                return false;
            }
            SyncMessage pMsg = CreateSenderMessage(bIsCustom ? MessageType.CUSTOM : MessageType.REQUEST, strData);
            return SendMessage(pMsg);
        }

        private bool ProcessRequest(SyncMessage pReceived) {
            string strData = pReceived.Data;
            if (strData != null && strData.Length > 0) {
                string strCommand = "";
                string strDetail = "";
                int nSep = strData.IndexOf(" ");
                if (nSep > 0) {
                    strCommand = strData.Substring(0, nSep).ToLower();
                    if (strData.Length > nSep + 1) {
                        strDetail = strData.Substring(nSep + 1);
                    }
                } else {
                    strCommand = strData.ToLower();
                }

                int nSlideIndex;
                switch (strCommand) {
                    case "alert":
                        string strSender = pReceived.SenderIP + ":" + pReceived.SenderPort.ToString();
                        DialogResult pRet = MessageBox.Show(strDetail, strSender);
                        return pRet == DialogResult.OK;
                    case "select":
                        if (int.TryParse(strDetail, out nSlideIndex)) {
                            return Globals.ThisAddIn.DoSelectSlide(nSlideIndex);
                        }
                        break;
                    case "showslide":
                        if (int.TryParse(strDetail, out nSlideIndex)) {
                            return Globals.ThisAddIn.DoSlideShow(nSlideIndex);
                        }
                        break;
                    case "hideslide":
                        return Globals.ThisAddIn.DoSlideShowEnd();
                }
            }
            return false;
        }

        public static SyncManager GetInstance() {
            if (pInstance == null) {
                pInstance = new SyncManager();
            }
            return pInstance;
        }
        public static string FindLocalIPAddress() {
            IPHostEntry host = Dns.GetHostEntry(Dns.GetHostName());
            foreach (IPAddress ip in host.AddressList) {
                if (ip.AddressFamily == AddressFamily.InterNetwork) {
                    return ip.ToString();
                }
            }
            return DEFAULT_LOCAL;
        }
        private static void LogMessage(string strPrefix, SyncMessage pMsg) {
            string strText = strPrefix + " ";
            if (pMsg != null) {
                strText += String.Format("[{0}:{1}][{2}][{3}] ", pMsg.SenderIP, pMsg.SenderPort, pMsg.ID, ((MessageType)pMsg.Type).ToString());
                strText += pMsg.Data;
            }
            Globals.ThisAddIn.LogDebug(strText);
        }
    }
}
