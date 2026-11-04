using System;
using System.Collections.Generic;

namespace ExcelAiAssistant.Core.Search
{
    /// <summary>
    /// 搜索提供方自动探测（联网搜索批次，用户拍板主 A 兜 B）：
    /// 按模型方案 BaseUrl 前缀探测提供方与其服务端搜索能力——
    /// 命中「服务端搜索」提供方时客户端注入对应搜索参数（智谱 web_search 工具 /
    /// Kimi 内置 $web_search / 通义 enable_search / OpenRouter :online 后缀 / 千帆 AI 搜索）；
    /// 未命中（DeepSeek 官方等无搜索能力提供方）降级产品 web_search 工具走 cn.bing.com 抓取兜底。
    /// 纯逻辑可单测；不发送任何网络请求。
    /// </summary>
    public static class SearchProviderDetector
    {
        public enum SearchCapability
        {
            /// <summary>提供方服务端搜索：客户端按 Provider 注入请求参数，模型返回搜索引用。</summary>
            ProviderNative,

            /// <summary>提供方无搜索能力：走产品 web_search 工具（bing 抓取兜底）。</summary>
            FallbackBing,

            /// <summary>用户在设置页显式关闭联网搜索。</summary>
            Disabled
        }

        public sealed class Detection
        {
            public Detection(string provider, SearchCapability capability, string note)
            {
                Provider = provider;
                Capability = capability;
                Note = note;
            }

            /// <summary>提供方标识（zhipu/moonshot/qwen/openrouter/qianfan/deepseek/unknown）。</summary>
            public string Provider { get; }

            public SearchCapability Capability { get; }

            /// <summary>给用户/AI 看的一句话判定依据（如「按 baseUrl 前缀 api.zhipuai.cn 命中智谱服务端搜索」）。</summary>
            public string Note { get; }
        }

        /// <summary>设置页总开关缺失/损坏时的默认值：联网搜索默认开启。</summary>
        public const bool DefaultEnabled = true;

        public static Detection Detect(string baseUrl, bool webSearchEnabled = DefaultEnabled)
        {
            if (!webSearchEnabled)
            {
                return new Detection("disabled", SearchCapability.Disabled, "联网搜索已在设置中关闭。");
            }

            var host = ExtractHost(baseUrl);
            if (string.IsNullOrEmpty(host))
            {
                return new Detection("unknown", SearchCapability.FallbackBing, "模型地址无法识别，搜索走 cn.bing.com 抓取兜底。");
            }

            if (host.IndexOf("zhipuai", StringComparison.OrdinalIgnoreCase) >= 0 ||
                host.IndexOf("bigmodel", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return new Detection("zhipu", SearchCapability.ProviderNative,
                    "命中智谱开放平台（服务端 web_search 工具），请求将注入搜索参数。");
            }

            if (host.IndexOf("moonshot", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return new Detection("moonshot", SearchCapability.ProviderNative,
                    "命中 Kimi/Moonshot（内置 $web_search），请求将注入搜索参数。");
            }

            if (host.IndexOf("dashscope", StringComparison.OrdinalIgnoreCase) >= 0 ||
                host.IndexOf("aliyuncs", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return new Detection("qwen", SearchCapability.ProviderNative,
                    "命中通义千问 DashScope（enable_search 参数），请求将注入搜索开关。");
            }

            if (host.IndexOf("openrouter", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return new Detection("openrouter", SearchCapability.ProviderNative,
                    "命中 OpenRouter（:online 后缀），请求将追加联网插件。");
            }

            if (host.IndexOf("bcebos", StringComparison.OrdinalIgnoreCase) >= 0 ||
                host.IndexOf("baidubce", StringComparison.OrdinalIgnoreCase) >= 0 ||
                host.IndexOf("qianfan", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return new Detection("qianfan", SearchCapability.ProviderNative,
                    "命中百度千帆（AI 搜索组件），请求将注入搜索参数。");
            }

            if (host.IndexOf("deepseek", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return new Detection("deepseek", SearchCapability.FallbackBing,
                    "DeepSeek 官方 API 无服务端搜索，搜索走 cn.bing.com 抓取兜底。");
            }

            return new Detection("unknown", SearchCapability.FallbackBing,
                "模型地址未匹配已知服务端搜索提供方，搜索走 cn.bing.com 抓取兜底。");
        }

        /// <summary>OpenRouter 服务端搜索通过模型名 :online 后缀启用。</summary>
        public static string ApplyModelSuffix(string provider, string modelName)
        {
            if (provider != "openrouter" || string.IsNullOrWhiteSpace(modelName))
            {
                return modelName;
            }

            return modelName.EndsWith(":online", StringComparison.OrdinalIgnoreCase)
                ? modelName
                : modelName + ":online";
        }

        private static string ExtractHost(string baseUrl)
        {
            if (string.IsNullOrWhiteSpace(baseUrl))
            {
                return string.Empty;
            }

            var text = baseUrl.Trim();
            if (!text.StartsWith("http://", StringComparison.OrdinalIgnoreCase) &&
                !text.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
            {
                text = "https://" + text;
            }

            return Uri.TryCreate(text, UriKind.Absolute, out var uri)
                ? uri.Host
                : string.Empty;
        }
    }
}
