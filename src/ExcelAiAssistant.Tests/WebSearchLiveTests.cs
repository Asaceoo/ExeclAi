using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAiAssistant.Core.Search;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// 联网搜索真机闭环（v3.8.1.83）：环境变量守卫模式（同技能市场安装链先例）——
    /// gate 默认跳过不联网；真机验证时设置 XAI_LIVE_WEBSEARCH=1 手动运行：
    /// set XAI_LIVE_WEBSEARCH=1 &amp;&amp; dotnet test --filter WebSearchLiveTests
    /// </summary>
    public sealed class WebSearchLiveTests
    {
        private static bool LiveEnabled =>
            string.Equals(
                Environment.GetEnvironmentVariable("XAI_LIVE_WEBSEARCH"),
                "1",
                StringComparison.OrdinalIgnoreCase);

        [Fact]
        public void BingSearch_realNetwork_returnsAnnotatedResults()
        {
            if (!LiveEnabled)
            {
                return; // gate 默认跳过：真机联网验证需 XAI_LIVE_WEBSEARCH=1
            }

            var result = WebSearchService.Search(new Dictionary<string, object>
            {
                ["query"] = "manufacturing industry average yield rate",
                ["limit"] = 3
            });

            Assert.True(result.Ok, "真机 bing 搜索应成功：" + result.Warning);
            var data = (Newtonsoft.Json.Linq.JObject)Newtonsoft.Json.Linq.JToken.FromObject(result.Data);
            Assert.Equal("bing_scrape", data.Value<string>("source"));
            var results = (Newtonsoft.Json.Linq.JArray)data["results"];
            Assert.True(results.Count >= 1, "至少返回 1 条结果");
            var first = (Newtonsoft.Json.Linq.JObject)results[0];
            Console.WriteLine("DIAG " + results.ToString(Newtonsoft.Json.Formatting.None).Substring(0, Math.Min(600, results.ToString(Newtonsoft.Json.Formatting.None).Length)));
            Assert.False(string.IsNullOrWhiteSpace(first.Value<string>("title")));
            Assert.StartsWith("http", first.Value<string>("url"), StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void ProviderDetection_realMachine_assertsKnownMappings()
        {
            if (!LiveEnabled)
            {
                return;
            }

            Assert.Equal("zhipu", SearchProviderDetector.Detect("https://open.bigmodel.cn/api/paas/v4").Provider);
            Assert.Equal("moonshot", SearchProviderDetector.Detect("https://api.moonshot.cn/v1").Provider);
            Assert.Equal("qwen", SearchProviderDetector.Detect("https://dashscope.aliyuncs.com/compatible-mode/v1").Provider);
            Assert.Equal(SearchProviderDetector.SearchCapability.FallbackBing,
                SearchProviderDetector.Detect("https://api.deepseek.com/v1").Capability);
        }
    }
}
