using System.Collections.Generic;
using System.Linq;
using ExcelAiAssistant.Core.Agent;
using ExcelAiAssistant.Core.Search;
using Newtonsoft.Json.Linq;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>联网搜索纯逻辑单测（v3.8.1.83）：提供方探测/模型后缀/请求注入/bing 解析，全部离线。</summary>
    public class WebSearchLogicTests
    {
        [Theory]
        [InlineData("https://open.bigmodel.cn/api/paas/v4/chat/completions", "zhipu", SearchProviderDetector.SearchCapability.ProviderNative)]
        [InlineData("https://api.zhipuai.cn/v1/chat/completions", "zhipu", SearchProviderDetector.SearchCapability.ProviderNative)]
        [InlineData("https://api.moonshot.cn/v1", "moonshot", SearchProviderDetector.SearchCapability.ProviderNative)]
        [InlineData("https://dashscope.aliyuncs.com/compatible-mode/v1", "qwen", SearchProviderDetector.SearchCapability.ProviderNative)]
        [InlineData("https://openrouter.ai/api/v1", "openrouter", SearchProviderDetector.SearchCapability.ProviderNative)]
        [InlineData("https://qianfan.baidubce.com/v2", "qianfan", SearchProviderDetector.SearchCapability.ProviderNative)]
        [InlineData("https://api.deepseek.com/v1", "deepseek", SearchProviderDetector.SearchCapability.FallbackBing)]
        [InlineData("https://api.openai.com/v1", "unknown", SearchProviderDetector.SearchCapability.FallbackBing)]
        [InlineData("", "unknown", SearchProviderDetector.SearchCapability.FallbackBing)]
        public void Detect_MapsProvidersByBaseUrl(string baseUrl, string expectedProvider, SearchProviderDetector.SearchCapability expectedCapability)
        {
            var detection = SearchProviderDetector.Detect(baseUrl);
            Assert.Equal(expectedProvider, detection.Provider);
            Assert.Equal(expectedCapability, detection.Capability);
            Assert.False(string.IsNullOrEmpty(detection.Note));
        }

        [Fact]
        public void Detect_Disabled_ReturnsDisabled()
        {
            var detection = SearchProviderDetector.Detect("https://open.bigmodel.cn/api/paas/v4", webSearchEnabled: false);
            Assert.Equal("disabled", detection.Provider);
            Assert.Equal(SearchProviderDetector.SearchCapability.Disabled, detection.Capability);
        }

        [Fact]
        public void ApplyModelSuffix_OpenRouter_AppendsOnline()
        {
            Assert.Equal("openai/gpt-4o:online", SearchProviderDetector.ApplyModelSuffix("openrouter", "openai/gpt-4o"));
            Assert.Equal("openai/gpt-4o:online", SearchProviderDetector.ApplyModelSuffix("openrouter", "openai/gpt-4o:online"));
            Assert.Equal("deepseek-chat", SearchProviderDetector.ApplyModelSuffix("deepseek", "deepseek-chat"));
            Assert.Equal(string.Empty, SearchProviderDetector.ApplyModelSuffix("openrouter", string.Empty));
        }

        [Fact]
        public void Inject_Zhipu_AddsWebSearchToolEntry()
        {
            var request = new JObject
            {
                ["model"] = "glm-4-plus",
                ["tools"] = new JArray { new JObject { ["type"] = "function" } }
            };
            SearchRequestInjector.Apply(request, "zhipu");
            var tools = (JArray)request["tools"];
            Assert.Equal(2, tools.Count);
            Assert.Equal("web_search", tools[1].Value<string>("type"));
            Assert.True(tools[1]["web_search"].Value<bool>("enable"));
        }

        [Fact]
        public void Inject_Moonshot_AddsBuiltinFunction()
        {
            var request = new JObject { ["model"] = "moonshot-v1-8k" };
            SearchRequestInjector.Apply(request, "moonshot");
            var tools = (JArray)request["tools"];
            Assert.Single(tools);
            Assert.Equal("builtin_function", tools[0].Value<string>("type"));
            Assert.Equal("$web_search", tools[0]["function"].Value<string>("name"));
        }

        [Fact]
        public void Inject_Qwen_SetsEnableSearch()
        {
            var request = new JObject { ["model"] = "qwen-plus" };
            SearchRequestInjector.Apply(request, "qwen");
            Assert.True(request.Value<bool>("enable_search"));
            Assert.Null(request["tools"]);
        }

        [Fact]
        public void Inject_OpenRouter_AppendsOnlineSuffix()
        {
            var request = new JObject { ["model"] = "openai/gpt-4o" };
            SearchRequestInjector.Apply(request, "openrouter");
            Assert.Equal("openai/gpt-4o:online", request.Value<string>("model"));
        }

        [Fact]
        public void Inject_Qianfan_SetsWebSearch()
        {
            var request = new JObject { ["model"] = "ernie-4.0" };
            SearchRequestInjector.Apply(request, "qianfan");
            Assert.True(request["web_search"].Value<bool>("enable"));
        }

        [Fact]
        public void Inject_UnknownOrDisabled_NoMutation()
        {
            var request = new JObject { ["model"] = "deepseek-chat" };
            SearchRequestInjector.Apply(request, "deepseek");
            SearchRequestInjector.Apply(request, "unknown");
            Assert.Equal("deepseek-chat", request.Value<string>("model"));
            Assert.Null(request["enable_search"]);
            Assert.Null(request["tools"]);
            Assert.Null(request["plugins"]);
            Assert.Null(request["web_search"]);
        }

        [Fact]
        public void ParseBingHtml_ExtractsTitleUrlSnippet()
        {
            var html = "<html><body><ol id=\"b_results\">" +
                       "<li class=\"b_algo\"><h2><a href=\"https://example.com/a\">制造业<b>良率</b>基准报告</a></h2>" +
                       "<p>行业平均良率约为 <strong>85%</strong>，头部厂商可达 95%。</p></li>" +
                       "<li class=\"b_algo\"><h2><a href=\"https://example.com/b\">2026 制造业趋势</a></h2><p>趋势概述。</p></li>" +
                       "</ol></body></html>";
            var hits = WebSearchService.ParseResults(html);
            Assert.Equal(2, hits.Count);
            Assert.Equal("https://example.com/a", hits[0].Url);
            Assert.Equal("制造业良率基准报告", hits[0].Title);
            Assert.Contains("85%", hits[0].Snippet);
            Assert.Equal("2026 制造业趋势", hits[1].Title);
        }

        [Fact]
        public void ParseBingHtml_EmptyOrNoResults_ReturnsEmpty()
        {
            Assert.Empty(WebSearchService.ParseResults(string.Empty));
            Assert.Empty(WebSearchService.ParseResults("<html><body>风控页</body></html>"));
        }

        [Fact]
        public void ParseBingHtml_RespectsMaxResults()
        {
            var blocks = string.Concat(System.Linq.Enumerable.Repeat(
                "<li class=\"b_algo\"><h2><a href=\"https://example.com/x\">标题</a></h2><p>摘要</p></li>", 5));
            Assert.Equal(2, WebSearchService.ParseResults(blocks, maxResults: 2).Count);
        }
    }
}
