using System;
using System.Collections.Generic;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Linq;
using ExcelAiAssistant.Core.Diagnostics;
using ExcelAiAssistant.Core.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelAiAssistant.Core.Search
{
    /// <summary>
    /// cn.bing.com 网页抓取搜索兜底（联网搜索批次，主 A 兜 B 的 B）：
    /// 无服务端搜索能力的提供方（DeepSeek 官方等）或抓取兜底分支使用。
    /// 结果统一标注「网页抓取」来源类型；隐私红线：query 只传关键词，原始数据不出本机。
    /// fail-closed：网络失败/解析零结果返回 Ok=false + 人话 ErrorCode，不静默成功。
    /// 解析规则经真实网络实测（v3.8.1.83 批次：cn.bing.com b_algo 结构，9 块可解析）。
    /// </summary>
    public sealed class WebSearchService
    {
        private const string BingSearchUrl = "https://cn.bing.com/search?q=";
        private const string BrowserUa =
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36";

        private const int MaxResults = 8;

        /// <summary>从 ToolResult.Data 提取结果条数（供内置函数回执摘要）。</summary>
        public static int CountOf(ToolResult result)
        {
            try
            {
                if (Newtonsoft.Json.Linq.JToken.FromObject(result.Data) is Newtonsoft.Json.Linq.JObject data &&
                    data["results"] is Newtonsoft.Json.Linq.JArray results)
                {
                    return results.Count;
                }
            }
            catch (Newtonsoft.Json.JsonException)
            {
                AppLog.Warn("WebSearchService", "[LogCover] WebSearchService.cs:40  silent-catch covered (no exception variable)");
            }

            return 0;
        }

        /// <summary>解析校验夹具（单测用）：注入预置 HTML 走解析管线，不打网络。</summary>
        public static IReadOnlyList<SearchHit> ParseResults(string html, int maxResults = MaxResults)
        {
            var hits = new List<SearchHit>();
            if (string.IsNullOrEmpty(html))
            {
                return hits;
            }

            var blocks = Regex.Split(html, "<li class=\"b_algo\"", RegexOptions.IgnoreCase);
            for (var i = 1; i < blocks.Length && hits.Count < Math.Max(maxResults, 0); i++)
            {
                var block = blocks[i];
                var end = block.IndexOf("</li>", StringComparison.OrdinalIgnoreCase);
                if (end >= 0)
                {
                    block = block.Substring(0, end);
                }

                var link = Regex.Match(
                    block,
                    "<h2[^>]*><a[^>]*href=\"([^\"]+)\"[^>]*>(.*?)</a></h2>",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);
                if (!link.Success)
                {
                    continue;
                }

                var url = WebUtility.HtmlDecode(link.Groups[1].Value).Trim();
                var title = StripTags(link.Groups[2].Value);
                if (string.IsNullOrEmpty(url) || string.IsNullOrEmpty(title))
                {
                    continue;
                }

                var snippet = string.Empty;
                var para = Regex.Match(block, "<p[^>]*>(.*?)</p>", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                if (para.Success)
                {
                    snippet = StripTags(para.Groups[1].Value);
                }

                hits.Add(new SearchHit(title, url, snippet));
            }

            return hits;
        }

        public static ToolResult Search(Dictionary<string, object> args)
        {
            var query = GetString(args, "query");
            if (string.IsNullOrWhiteSpace(query))
            {
                return new ToolResult { Ok = false, ErrorCode = "INVALID_ARGS", Data = "缺少搜索关键词 query。" };
            }

            var trimmed = query.Trim();
            if (trimmed.Length > 120)
            {
                trimmed = trimmed.Substring(0, 120);
            }

            var limit = 5;
            if (args != null && args.TryGetValue("limit", out var rawLimit))
            {
                var parsed = Convert.ToInt32(rawLimit);
                if (parsed >= 1 && parsed <= MaxResults)
                {
                    limit = parsed;
                }
            }

            string html;
            try
            {
                EnsureTls12();
                using (var client = new WebClient { Encoding = Encoding.UTF8 })
                {
                    client.Proxy = WebRequest.DefaultWebProxy;
                    client.Headers[HttpRequestHeader.UserAgent] = BrowserUa;
                    client.Headers[HttpRequestHeader.AcceptLanguage] = "zh-CN,zh;q=0.9";
                    html = client.DownloadString(BingSearchUrl + Uri.EscapeDataString(trimmed));
                }
            }
            catch (Exception ex)
            {
                AppLog.Warn("search", "bing 抓取失败：" + ex.Message);
                return new ToolResult
                {
                    Ok = false,
                    ErrorCode = "SEARCH_NETWORK_FAILED",
                    Data = "搜索请求失败（cn.bing.com 不可达或被网络拦截）：" + ex.Message + "。请检查网络后重试，或改由模型提供方自带搜索承担。"
                };
            }

            var hits = ParseResults(html, limit);
            if (hits.Count == 0)
            {
                return new ToolResult
                {
                    Ok = false,
                    ErrorCode = "SEARCH_EMPTY",
                    Data = "搜索完成但未解析到结果（页面结构可能变化或触发了风控）。可换关键词重试。"
                };
            }

            AppLog.Info("search", $"bing 搜索 [{trimmed}] 返回 {hits.Count} 条");
            return new ToolResult
            {
                Ok = true,
                Data = new
                {
                    source = "bing_scrape",
                    source_label = "网页抓取（cn.bing.com）",
                    query = trimmed,
                    results = hits.Select(hit => new
                    {
                        title = hit.Title,
                        url = hit.Url,
                        snippet = hit.Snippet
                    }).ToList(),
                    privacy_note = "仅搜索关键词出网，工作簿数据未外发。引用外部结果时须注明为公开网页信息，不得当作权威统计口径。"
                }
            };
        }

        private static void EnsureTls12()
        {
            try
            {
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            }
            catch (Exception)
            {
                AppLog.Warn("WebSearchService", "[LogCover] WebSearchService.cs:179  silent-catch covered (no exception variable)");
                // 协议位已含或平台不支持时忽略（只升不降）。
        }
        }

        private static string StripTags(string html)
        {
            var text = Regex.Replace(html ?? string.Empty, "<[^>]+>", string.Empty);
            text = WebUtility.HtmlDecode(text);
            text = Regex.Replace(text, "\\s+", " ").Trim();
            return text.Length > 200 ? text.Substring(0, 200) + "…" : text;
        }

        private static string GetString(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.TryGetValue(key, out var value) || value == null)
            {
                return null;
            }

            if (value is Newtonsoft.Json.Linq.JValue jv)
            {
                return jv.Value?.ToString();
            }

            return value.ToString();
        }
    }

    /// <summary>单条搜索结果。</summary>
    public sealed class SearchHit
    {
        public SearchHit(string title, string url, string snippet)
        {
            Title = title;
            Url = url;
            Snippet = snippet;
        }

        public string Title { get; }

        public string Url { get; }

        public string Snippet { get; }
    }
}
