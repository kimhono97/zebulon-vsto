using System;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;

namespace ZebulonVSTO.Slides {
    /// <summary>
    /// HTTP client for the Zebulon Web Bible API:
    ///   GET {WebBaseUrl}/api/bible?v={code}&amp;b={book 1-66}&amp;c={chapter}
    /// → <see cref="BibleData"/> (the chapter's verses). COM-free. Requires
    /// <see cref="SlideGenDefaults.WebBaseUrl"/> to be configured.
    /// </summary>
    public class BibleClient {
        // Shared transport (recommended HttpClient pattern; see ProviderClient).
        private static readonly HttpClient SharedHttp = new HttpClient();
        private readonly HttpClient _http;

        public BibleClient(HttpClient http = null) {
            try {
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            } catch {
                // ignore — environment may pin the protocol
            }
            _http = http ?? SharedHttp;
        }

        public async Task<BibleData> GetChapterAsync(string versionCode, int book, int chapter) {
            string web = (SlideGenDefaults.WebBaseUrl ?? "").TrimEnd('/');
            if (string.IsNullOrEmpty(web)) {
                throw new InvalidOperationException(
                    "Zebulon Web URL이 설정되지 않았습니다 (SlideGenDefaults.WebBaseUrl).");
            }
            string url = web + "/api/bible?v=" + Uri.EscapeDataString(versionCode ?? "")
                + "&b=" + book + "&c=" + chapter;
            string json = await _http.GetStringAsync(url).ConfigureAwait(false);
            BibleData data = BibleData.Parse(json);
            if (data == null) {
                throw new InvalidOperationException(
                    "성경 데이터 파싱 실패 (" + versionCode + " " + book + ":" + chapter + ").");
            }
            return data;
        }
    }
}
