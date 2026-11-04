using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using ExcelAiAssistant.Core.Agent;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    public sealed class AgentExperienceTests
    {
        [Fact]
        public void TokenEstimator_counts_cjk_as_one_and_latin_as_four_chars()
        {
            Assert.Equal(6, TokenEstimator.Estimate("中国中文测试"));
            Assert.Equal(3, TokenEstimator.Estimate("abcdefghijkl")); // 12 / 4 = 3
            Assert.Equal(0, TokenEstimator.Estimate(string.Empty));
            Assert.Equal(0, TokenEstimator.Estimate((string)null));
        }

        [Fact]
        public void TokenEstimator_sums_message_contents()
        {
            var messages = new List<ChatMessage>
            {
                ChatMessage.System("你是助手"),
                ChatMessage.User("请分析数据")
            };
            var estimate = TokenEstimator.Estimate(messages);
            var sumOfContents =
                TokenEstimator.Estimate("你是助手") + TokenEstimator.Estimate("请分析数据");
            Assert.True(estimate > 0);
            // 除正文外还需计入每条消息的角色与格式开销，故整体应大于正文之和。
            Assert.True(estimate > sumOfContents);
        }

        [Fact]
        public void SensitivityScanner_detects_email_mobile_and_id()
        {
            var text = "联系 13800138000 和 test@example.com，身份证 110105199003078512，随便写点。" +
                       "1000 2000"; // "1000"不应命中银行卡规则（太短）
            var detections = SensitivityScanner.Scan(text);

            Assert.Contains(detections, d => d.Kind == "email" && d.Count >= 1);
            Assert.Contains(detections, d => d.Kind == "cn_mobile" && d.Sample == "13800138000");
            Assert.Contains(detections, d => d.Kind == "cn_id");
            Assert.DoesNotContain(detections, d => d.Kind == "bank_card");
        }

        [Fact]
        public void SensitivityScanner_returns_empty_for_plain_text()
        {
            Assert.Empty(SensitivityScanner.Scan("今天天气不错，汇总第一季度销售数据。"));
            Assert.Empty(SensitivityScanner.Scan(string.Empty));
            Assert.Empty(SensitivityScanner.Scan(null));
        }

        [Fact]
        public void FriendlyError_maps_known_error_codes()
        {
            Assert.Contains("取消", FriendlyError.For("USER_DENIED", null));
            Assert.Contains("工作簿", FriendlyError.For("UNDO_WORKBOOK_MISMATCH", null));
            Assert.Equal("回退文本", FriendlyError.For("SOME_UNKNOWN", "回退文本"));
            Assert.Equal("操作失败，请重试。", FriendlyError.For(null, null));
        }

        [Fact]
        public async Task Auto_mode_degrades_to_fallback_when_native_tools_rejected()
        {
            var handler = new FallbackFakeHandler();
            var profile = new ModelProfile
            {
                Name = "测试",
                BaseUrl = "https://example.com/v1",
                ApiKey = "secret",
                ModelName = "gpt-test",
                ProtocolMode = AgentProtocolMode.Auto
            };
            using var client = new OpenAiCompatibleClient(profile, new HttpClient(handler));
            var tools = new List<ToolDescriptor>
            {
                new ToolDescriptor("read_excel", "读取", RiskLevel.ReadOnly, _ => new ToolResult { Ok = true })
            };

            var result = await client.CompleteAsync(
                new List<ChatMessage> { ChatMessage.User("请读数据") },
                tools,
                CancellationToken.None);

            Assert.True(result.FallbackUsed);
            Assert.Equal(2, handler.RequestCount); // 原生失败 → 降级重试一次
            Assert.True(handler.FirstRequestHadTools);
            Assert.False(handler.SecondRequestHadTools);
            Assert.Contains("action", result.Content);
        }

        [Fact]
        public void ToolEventView_builds_success_and_failure_views()
        {
            var success = ToolEventViewBuilder.Build(new ToolEventRecord
            {
                Tool = "preview_excel_range",
                Confirmed = true,
                Result = new ToolResult { Ok = true, Data = new { rows = 5, columns = 3 } }
            });
            var successToken = Newtonsoft.Json.Linq.JObject.FromObject(success);
            Assert.True((bool)successToken["ok"]);
            Assert.True((bool)successToken["confirmed"]);
            Assert.True(string.IsNullOrEmpty((string)successToken["error_code"]));
            Assert.Contains("rows=5", (string)successToken["summary"]);
            Assert.DoesNotContain("{", (string)successToken["summary"]);
            Assert.Contains("rows", (string)successToken["detail"]);

            var failure = ToolEventViewBuilder.Build(new ToolEventRecord
            {
                Tool = "write_excel_range",
                Confirmed = false,
                Result = new ToolResult { Ok = false, ErrorCode = "SHAPE_MISMATCH", Data = "行列数不匹配" }
            });
            var failureToken = Newtonsoft.Json.Linq.JObject.FromObject(failure);
            Assert.False((bool)failureToken["ok"]);
            Assert.False((bool)failureToken["confirmed"]);
            Assert.Equal("SHAPE_MISMATCH", (string)failureToken["error_code"]);
            Assert.Contains("行列数不匹配", (string)failureToken["error_message"]);
        }

        [Fact]
        public void ToolEventView_truncates_long_summary_and_handles_null()
        {
            var longText = new string('a', 500);
            var view = ToolEventViewBuilder.Build(new ToolEventRecord
            {
                Tool = "t",
                Confirmed = true,
                Result = new ToolResult { Ok = true, Data = longText }
            });
            var token = Newtonsoft.Json.Linq.JObject.FromObject(view);
            Assert.True(((string)token["summary"]).Length <= 141);
            Assert.EndsWith("…", (string)token["summary"]);
            Assert.Null(ToolEventViewBuilder.Build(null));
        }

        private sealed class FallbackFakeHandler : HttpMessageHandler
        {
            public int RequestCount { get; private set; }

            public bool FirstRequestHadTools { get; private set; } = true;

            public bool SecondRequestHadTools { get; private set; } = true;

            protected override async Task<HttpResponseMessage> SendAsync(
                HttpRequestMessage request,
                CancellationToken cancellationToken)
            {
                RequestCount++;
                var body = await request.Content.ReadAsStringAsync();
                var hadTools = body.Contains("\"tools\"");
                if (RequestCount == 1)
                {
                    FirstRequestHadTools = hadTools;
                    return new HttpResponseMessage(HttpStatusCode.BadRequest)
                    {
                        Content = new StringContent("{\"error\":\"tools not supported\"}")
                    };
                }

                SecondRequestHadTools = hadTools;
                return new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent(
                        "{\"choices\":[{\"message\":{\"content\":\"{\\\"action\\\":\\\"final\\\",\\\"content\\\":\\\"ok\\\"}\"}}]," +
                        "\"usage\":{\"prompt_tokens\":10,\"completion_tokens\":5,\"total_tokens\":15}}")
                };
            }
        }
    }
}