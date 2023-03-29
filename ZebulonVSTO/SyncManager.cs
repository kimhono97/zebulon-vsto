using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Sockets;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ZebulonVSTO {
    public class SyncManager {
        private static SyncManager pInstance = null;

        public enum SyncMode {
            NONE, SENDER, RECEIVER
        }

        private string strLocalIP;
        private string strRemoteIP;
        private int nLocalPort;
        private int nRemotePort;
        private SyncMode nSyncMode; 
        private bool bRunning;
        
        private SyncManager() {
            this.strLocalIP = FindLocalIPAddress();
            this.nLocalPort = 0;
            this.strRemoteIP = "";
            this.nRemotePort = 0;
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
            MessageBox.Show("StartSync");
            this.bRunning = true;
            return true;
        }

        public void StopSync() {
            if (!this.bRunning) {
                return;
            }
            MessageBox.Show("StopSync");
            this.bRunning = false;
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
            return null;
        }
    }
}
