using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;

namespace ZebulonVSTO.Slides {
    /// <summary>
    /// HTTP client for the Zebulon Provider (lyric data). COM-free. The Provider
    /// base URL is resolved at runtime from {WebBaseUrl}/api/proj?n=zebulon
    /// (urls.alias) — the same source Zebulon Web uses — and falls back to
    /// <see cref="SlideGenDefaults.ProviderBaseUrl"/> when WebBaseUrl is unset or
    /// resolution fails.
    /// </summary>
    public class ProviderClient {
        // One long-lived transport for the whole add-in: the recommended HttpClient
        // pattern, avoiding leaked handlers / socket churn across repeated dialog
        // opens. Tests may still inject their own HttpClient via the constructor.
        private static readonly HttpClient SharedHttp = new HttpClient();

        private readonly HttpClient _http;
        private string _providerUrl;
        private bool _resolved;

        public ProviderClient(HttpClient http = null) {
            // .NET Framework 4.7.2: ensure TLS 1.2 is enabled for the HTTPS
            // Vercel/Cloud Run endpoints (older process defaults may omit it).
            try {
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            } catch {
                // ignore — environment may pin the protocol
            }
            _http = http ?? SharedHttp;
            _providerUrl = StripTrailingSlash(SlideGenDefaults.ProviderBaseUrl);
        }

        public string ProviderUrl {
            get { return _providerUrl; }
        }

        /// <summary>Resolve the Provider URL via Zebulon Web's /api/proj (best effort, once).</summary>
        public async Task EnsureResolvedAsync() {
            if (_resolved) {
                return;
            }
            _resolved = true;
            string web = StripTrailingSlash(SlideGenDefaults.WebBaseUrl);
            if (string.IsNullOrEmpty(web)) {
                return; // no web base configured → keep the default provider URL
            }
            try {
                string json = await _http.GetStringAsync(web + "/api/proj?n=zebulon").ConfigureAwait(false);
                string alias = ParseProjAlias(json);
                if (!string.IsNullOrEmpty(alias)) {
                    _providerUrl = StripTrailingSlash(alias);
                }
            } catch {
                // keep the default provider URL
            }
        }

        /// <summary>GET /lyric → array of lyric file paths.</summary>
        public async Task<List<string>> ListLyricsAsync() {
            await EnsureResolvedAsync().ConfigureAwait(false);
            string json = await _http.GetStringAsync(_providerUrl + "/lyric").ConfigureAwait(false);
            return LyricJson.ParseStringArray(json);
        }

        /// <summary>POST /lyric { path } → parsed Lyric.</summary>
        public async Task<Lyric> GetLyricAsync(string path) {
            await EnsureResolvedAsync().ConfigureAwait(false);
            string body = "{\"path\":" + JsonQuote(path) + "}";
            using (StringContent content = new StringContent(body, Encoding.UTF8, "application/json")) {
                HttpResponseMessage resp = await _http.PostAsync(_providerUrl + "/lyric", content).ConfigureAwait(false);
                resp.EnsureSuccessStatusCode();
                string json = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
                return LyricJson.ParseLyric(json);
            }
        }

        /// <summary>GET /template → array of template paths (e.g. "templates/PRAISE_DSM.pptx").</summary>
        public async Task<List<string>> ListTemplatesAsync() {
            await EnsureResolvedAsync().ConfigureAwait(false);
            string json = await _http.GetStringAsync(_providerUrl + "/template").ConfigureAwait(false);
            return LyricJson.ParseStringArray(json);
        }

        /// <summary>
        /// POST /template { path } → raw .pptx bytes. The Provider can return HTTP
        /// 200 with a JSON {err} on bad input, so success is judged by the response
        /// Content-Type (must be the openxml presentation type), not the status code.
        /// </summary>
        public async Task<byte[]> DownloadTemplateAsync(string path) {
            await EnsureResolvedAsync().ConfigureAwait(false);
            string body = "{\"path\":" + JsonQuote(path) + "}";
            using (StringContent content = new StringContent(body, Encoding.UTF8, "application/json")) {
                HttpResponseMessage resp = await _http.PostAsync(_providerUrl + "/template", content).ConfigureAwait(false);
                string ctype = resp.Content.Headers.ContentType != null ? resp.Content.Headers.ContentType.MediaType : "";
                if (!resp.IsSuccessStatusCode || ctype.IndexOf("openxmlformats", StringComparison.OrdinalIgnoreCase) < 0) {
                    string err = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
                    if (err != null && err.Length > 200) {
                        err = err.Substring(0, 200);
                    }
                    throw new InvalidOperationException("템플릿 다운로드 실패: " + err);
                }
                return await resp.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
            }
        }

        private static string StripTrailingSlash(string url) {
            if (string.IsNullOrEmpty(url)) {
                return "";
            }
            return url.TrimEnd('/');
        }

        // Minimal JSON string escaping for the request body (paths may contain
        // spaces and non-ASCII characters, e.g. Korean filenames).
        private static string JsonQuote(string s) {
            if (s == null) {
                return "\"\"";
            }
            StringBuilder sb = new StringBuilder(s.Length + 2);
            sb.Append('"');
            foreach (char c in s) {
                switch (c) {
                    case '"': sb.Append("\\\""); break;
                    case '\\': sb.Append("\\\\"); break;
                    case '\b': sb.Append("\\b"); break;
                    case '\f': sb.Append("\\f"); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    default:
                        if (c < ' ') {
                            sb.Append("\\u").Append(((int)c).ToString("x4"));
                        } else {
                            sb.Append(c);
                        }
                        break;
                }
            }
            sb.Append('"');
            return sb.ToString();
        }

        [DataContract]
        private class ProjResponse {
            [DataMember(Name = "urls")] public ProjUrls Urls { get; set; }
        }

        [DataContract]
        private class ProjUrls {
            [DataMember(Name = "alias")] public string Alias { get; set; }
        }

        private static string ParseProjAlias(string json) {
            if (string.IsNullOrEmpty(json)) {
                return null;
            }
            try {
                using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(json))) {
                    DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(ProjResponse));
                    ProjResponse r = serializer.ReadObject(ms) as ProjResponse;
                    return r != null && r.Urls != null ? r.Urls.Alias : null;
                }
            } catch {
                return null;
            }
        }
    }
}
