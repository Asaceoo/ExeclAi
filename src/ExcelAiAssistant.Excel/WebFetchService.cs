using System;
using ExcelAiAssistant.Core.Diagnostics;
using System.Collections.Generic;
using System.Globalization;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using ExcelAiAssistant.Core.Models;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// 只读网页抓取工具（fetch_url）。仅支持 http/https，返回抓取到的纯文本（可脱敏），
    /// 不写工作簿、不执行任何脚本、不访问本地文件。协议白名单为硬性底线。
    /// </summary>
    internal static class WebFetchService
    {
        private const int DefaultMaxBytes = 256 * 1024;
        private const int MaxOutputChars = 12000;

        public static ToolResult Fetch(Dictionary<string, object> args)
        {
            var url = GetValue(args, "url");
            if (string.IsNullOrWhiteSpace(url))
            {
                return Error("INVALID_VALUE", "缺少 url 参数。");
            }

            url = url.Trim();
            // 硬性协议白名单：非 http/https 一律拒绝，保证本工具绝不会触及本地文件/内网协议等非互联网地址。
            if (!url.StartsWith("http://", StringComparison.OrdinalIgnoreCase) &&
                !url.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
            {
                return Error("INVALID_URL", "仅支持 http/https 协议，已拒绝该地址。");
            }

            // SSRF 底线：IP 字面量命中的内网/回环/链路本地/保留地址一律拒绝（含云元数据 169.254.169.254），
            // 避免 ReadOnly 自动放行下模型读到内网/云元数据并回传。主机名不解析，不触发注入式 DNS。
            if (!IsPublicHost(url))
            {
                return Error("BLOCKED_PRIVATE_ADDR", "已拒绝访问内网/回环/保留地址，仅允许公网地址。");
            }

            var timeoutMs = Clamp(GetInt(args, "timeout_ms", 30000), 1000, 60000);
            var maxBytes = Clamp(GetInt(args, "max_bytes", DefaultMaxBytes), 1024, 1024 * 1024);
            var sanitize = GetBool(args, "sanitize", true);
            var extract = GetValue(args, "extract");
            if (!string.IsNullOrWhiteSpace(extract) &&
                !extract.Equals("html", StringComparison.OrdinalIgnoreCase) &&
                !extract.Equals("raw", StringComparison.OrdinalIgnoreCase))
            {
                return Error("INVALID_EXTRACT", "extract 仅支持 html 或 raw。");
            }

            var raw = extract != null && extract.Equals("raw", StringComparison.OrdinalIgnoreCase);

            try
            {
                // .NET Framework 下即使 IE 代理被禁用，HttpClient 默认仍可能把请求导向残留代理地址（如 127.0.0.1:7897）导致超时。
                // 因此跟随 IE 代理启用开关：禁用则直连，启用则走系统代理。
                var proxyEnabled = SystemProxyEnabled();
                using (var handler = new HttpClientHandler())
                using (var client = new HttpClient(handler))
                {
                    handler.UseProxy = proxyEnabled;
                    if (proxyEnabled)
                    {
                        handler.Proxy = WebRequest.GetSystemWebProxy();
                    }

                    client.Timeout = TimeSpan.FromMilliseconds(timeoutMs);
                    client.DefaultRequestHeaders.UserAgent.TryParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64) JYYJAssistant/1.0");

                    using (var response = client.GetAsync(url).GetAwaiter().GetResult())
                    {
                        if (!response.IsSuccessStatusCode)
                        {
                            return Error("HTTP_" + (int)response.StatusCode,
                                "服务返回 HTTP " + (int)response.StatusCode + (string.IsNullOrEmpty(response.ReasonPhrase) ? "" : " " + response.ReasonPhrase));
                        }

                        var mediaType = response.Content.Headers.ContentType?.MediaType ?? "";
                        var encoding = ResolveEncoding(response.Content.Headers.ContentType?.CharSet ?? "");
                        var body = ReadLimited(response, encoding, maxBytes, out var truncated);

                        var compiled = raw ? body : ToPlainText(body, mediaType);
                        if (truncated)
                        {
                            compiled += "\n[内容超过 " + maxBytes + " 字节，已截断]";
                        }

                        if (sanitize)
                        {
                            compiled = Sanitize(compiled);
                        }

                        if (compiled.Length > MaxOutputChars)
                        {
                            compiled = compiled.Substring(0, MaxOutputChars) + "\n[结果过长，已截断]";
                        }

                        return new ToolResult { Ok = true, Data = compiled };
                    }
                }
            }
            catch (OperationCanceledException)
            {
                return Error("TIMEOUT", "请求超时（" + timeoutMs + "ms）。");
            }
            catch (HttpRequestException ex)
            {
                return Error("NETWORK_ERROR", "网络请求失败：" + ex.Message);
            }
            catch (Exception ex) when (!(ex is OutOfMemoryException))
            {
                return Error("FETCH_ERROR", "抓取失败：" + ex.Message);
            }
        }

        private static string ReadLimited(HttpResponseMessage response, Encoding encoding, int maxBytes, out bool truncated)
        {
            truncated = false;
            var buffer = new byte[4096];
            long total = 0;
            using (var stream = response.Content.ReadAsStreamAsync().GetAwaiter().GetResult())
            using (var bufferStream = new System.IO.MemoryStream())
            {
                int read;
                while ((read = stream.Read(buffer, 0, buffer.Length)) > 0)
                {
                    total += read;
                    if (total > maxBytes)
                    {
                        truncated = true;
                        break;
                    }

                    bufferStream.Write(buffer, 0, read);
                }

                // 整块缓冲后一次性解码：逐块 GetString 不会保留解码器状态，多字节字符一旦落在块边界
                // （UTF-8 中文 3 字节等在 4096 边界被切断）必然产出  乱码。内存上界 = maxBytes(≤1MB)。
                return encoding.GetString(bufferStream.GetBuffer(), 0, (int)bufferStream.Length);
            }
        }

        private static string ToPlainText(string text, string mediaType)
        {
            if (string.IsNullOrEmpty(text))
            {
                return text;
            }

            var head = text.TrimStart();
            var looksHtml = mediaType.IndexOf("html", StringComparison.OrdinalIgnoreCase) >= 0 ||
                head.StartsWith("<!DOCTYPE", StringComparison.OrdinalIgnoreCase) ||
                head.StartsWith("<html", StringComparison.OrdinalIgnoreCase) ||
                text.IndexOf("<body", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("</body", StringComparison.OrdinalIgnoreCase) >= 0;
            if (!looksHtml)
            {
                return text;
            }

            var s = Regex.Replace(text, "(?is)<(script|style|noscript|head)[^>]*>.*?</\\1>", " ");
            s = Regex.Replace(s, "(?is)<[^>]+>", " ");
            s = Regex.Replace(s, "&nbsp;?|&#160;", " ");
            s = Regex.Replace(s, "&amp;", "&");
            s = Regex.Replace(s, "&lt;", "<");
            s = Regex.Replace(s, "&gt;", ">");
            s = Regex.Replace(s, "&quot;", "\"");
            s = Regex.Replace(s, "&#\\d+;|&#x[0-9a-fA-F]+;", " ");
            s = Regex.Replace(s, "[\\r\\n\\t]+", "\n");
            s = Regex.Replace(s, "\\n{3,}", "\n\n");
            return s.Trim();
        }

        private static string Sanitize(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return text;
            }

            // 固定掩码，防止在与外部无关的本地处理返回前把敏感信息暴露给模型。
            text = Regex.Replace(text, @"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", "[邮箱已隐藏]");
            text = Regex.Replace(text, @"(?<!\d)1[3-9]\d{9}(?!\d)", "[手机已隐藏]");
            text = Regex.Replace(text, @"(?<!\d)\d{17}[\dXx](?!\d)", "[身份证已隐藏]");
            return text;
        }

        /// <summary>IP 字面量命中内网/回环/链路本地/保留地址则拒绝；主机名不解析（避免注入式 DNS 副作用），放行。</summary>
        private static bool IsPublicHost(string url)
        {
            if (!Uri.TryCreate(url, UriKind.Absolute, out var uri))
            {
                return false;
            }

            var hostType = Uri.CheckHostName(uri.Host);
            if (hostType == UriHostNameType.IPv4 || hostType == UriHostNameType.IPv6)
            {
                return !IPAddress.TryParse(uri.Host, out var ip) || !IsPrivateOrReserved(ip);
            }

            // M16 全量审查（SSRF 修复）：主机名不再无条件放行——DNS 重绑定/xip.io 类服务可把
            // 域名解析到回环/内网/云元数据。解析主机名的全部地址并逐一校验；解析失败按 fail-closed 拒绝。
            if (string.Equals(uri.Host, "localhost", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            try
            {
                var addresses = System.Net.Dns.GetHostAddresses(uri.Host);
                if (addresses == null || addresses.Length == 0)
                {
                    return false;
                }

                foreach (var address in addresses)
                {
                    if (IsPrivateOrReserved(address))
                    {
                        return false;
                    }
                }

                return true;
            }
            catch
            {
                return false; // 解析失败一律拒绝（fail-closed）
            }
        }

        private static bool IsPrivateOrReserved(IPAddress addr)
        {
            if (IPAddress.IsLoopback(addr)) return true;
            var bytes = addr.GetAddressBytes();
            if (bytes.Length == 16)
            {
                // IPv4-mapped IPv6（::ffff:127.0.0.1、::ffff:169.254.169.254 等）：映射回 IPv4 后
                // 按同一套规则复核，堵住"16 字节分支漏判 IPv4 私网"的绕过（M16 全量审查）。
                var isMapped = true;
                for (var i = 0; i < 10; i++)
                {
                    if (bytes[i] != 0) { isMapped = false; break; }
                }

                if (isMapped && bytes[10] == 0xff && bytes[11] == 0xff)
                {
                    return IsPrivateOrReserved(new IPAddress(new[] { bytes[12], bytes[13], bytes[14], bytes[15] }));
                }
            }

            if (bytes.Length == 4)
            {
                byte b0 = bytes[0];
                if (b0 == 0 || b0 == 10 || b0 == 127 || b0 >= 240) return true;                        // 0/8, 10/8, 127/8, 240/4
                if (b0 == 172 && bytes[1] >= 16 && bytes[1] <= 31) return true;                       // 172.16/12
                if (b0 == 192 && bytes[1] == 168) return true;                                        // 192.168/16
                if (b0 == 169 && bytes[1] == 254) return true;                                        // 169.254/16 链路本地(含云元数据)
                if (b0 == 100 && (bytes[1] & 0xC0) == 64) return true;                                // 100.64/10 CGNAT
                if (b0 == 192 && (bytes[1] & 0xF0) == 0 && (bytes[2] == 0 || bytes[2] == 2 || bytes[2] == 88)) return true; // 192.0.0.0/24, 192.0.2.0/24, 192.88.99.0/24
                if (b0 == 198 && (bytes[1] == 18 || bytes[1] == 19)) return true;                    // 198.18/15
                if (b0 == 203 && bytes[1] == 0 && bytes[2] == 113) return true;                      // 203.0.113.0/24 TEST-NET-3
                return false;
            }

            if (addr.IsIPv6LinkLocal || addr.IsIPv6SiteLocal) return true;
            if (addr.Equals(IPAddress.IPv6Loopback) || addr.Equals(IPAddress.IPv6Any)) return true;
            return false;
        }

        private static Encoding ResolveEncoding(string charset)
        {
            if (!string.IsNullOrWhiteSpace(charset))
            {
                try { return Encoding.GetEncoding(charset); }
                catch { }
            }

            return Encoding.UTF8;
        }

        private static int GetInt(Dictionary<string, object> args, string key, int fallback)
        {
            var value = GetValue(args, key);
            if (string.IsNullOrWhiteSpace(value))
            {
                return fallback;
            }

            if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed))
            {
                return parsed;
            }

            if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsedDouble))
            {
                return (int)parsedDouble;
            }

            return fallback;
        }

        private static bool GetBool(Dictionary<string, object> args, string key, bool fallback)
        {
            var value = GetValue(args, key);
            if (string.IsNullOrWhiteSpace(value))
            {
                return fallback;
            }

            switch (value.ToLowerInvariant())
            {
                case "true": case "1": case "yes": case "y": case "on": case "是": case "对":
                    return true;
                case "false": case "0": case "no": case "n": case "off": case "否": case "错":
                    return false;
                default:
                    return fallback;
            }
        }

        private static string GetValue(Dictionary<string, object> args, string key)
        {
            if (args != null && args.TryGetValue(key, out var value))
            {
                return value?.ToString();
            }

            return null;
        }

        private static int Clamp(int value, int min, int max)
        {
            return value < min ? min : value > max ? max : value;
        }

        private static bool SystemProxyEnabled()
        {
            try
            {
                var value = Microsoft.Win32.Registry.GetValue(
                    @"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings",
                    "ProxyEnable", 0);
                return value is int proxyEnabled && proxyEnabled == 1;
            }
            catch
            {
                return true;
            }
        }

        private static ToolResult Error(string code, string message)
        {
            return new ToolResult { Ok = false, ErrorCode = code, Data = message };
        }
    }
}