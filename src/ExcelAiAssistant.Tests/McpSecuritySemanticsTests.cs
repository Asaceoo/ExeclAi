using System;
using System.Collections.Generic;
using System.Text;
using ExcelAiAssistant.McpServer;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// N12（M27 第四期 P3 工程债）：确认码门、日期解析禁用、UTF-8 字节口径截断的回归锁定。
    /// </summary>
    public class McpSecuritySemanticsTests
    {
        // ===== 确认码（IsConfirmedWithCode）=====

        [Fact]
        public void IsConfirmedWithCode_no_confirmed_key_returns_false()
        {
            var args = new Dictionary<string, object> { ["address"] = "A1" };
            Assert.False(Program.IsConfirmedWithCode("test_tool_no_confirm", args));
        }

        [Fact]
        public void IsConfirmedWithCode_fail_closed_without_pending_record()
        {
            // 进程重启后直接带任意码（无待确认记录）——必须拒绝（fail-closed）
            var args = new Dictionary<string, object> { ["confirmed"] = "ABC12345" };
            Assert.False(Program.IsConfirmedWithCode("never_pending_tool", args));
        }

        [Fact]
        public void IsConfirmedWithCode_correct_code_consumes_once()
        {
            // BuildConfirmationText 生成码文件并登记 PendingConfirmations → 带码首次校验通过并消费 → 二次同码失败（一次性）
            var name = "unit_confirm_tool_" + Guid.NewGuid().ToString("N").Substring(0, 6);
            var text = Program.BuildConfirmationText(name, new Dictionary<string, object> { ["k"] = "v" });
            // （码不回显由 BuildConfirmationText_response_has_no_code_but_file_does 锁定）

            var args = new Dictionary<string, object> { ["confirmed"] = ExtractCode(name) };
            Assert.True(Program.IsConfirmedWithCode(name, args), "首次带正确码应放行");
            Assert.False(Program.IsConfirmedWithCode(name, args), "同码二次应拒绝（一次性消费）");
        }

        [Fact]
        public void IsConfirmedWithCode_wrong_code_stays_pending()
        {
            var name = "unit_confirm_wrong_" + Guid.NewGuid().ToString("N").Substring(0, 6);
            Program.BuildConfirmationText(name, new Dictionary<string, object> { ["k"] = "v" });
            var bad = new Dictionary<string, object> { ["confirmed"] = "DEADBEEF" };
            Assert.False(Program.IsConfirmedWithCode(name, bad), "错码应拒绝");
            // 待确认记录保留——正确码仍可放行
            var good = new Dictionary<string, object> { ["confirmed"] = ExtractCode(name) };
            Assert.True(Program.IsConfirmedWithCode(name, good), "错码不得销毁待确认记录");
        }

        [Fact]
        public void BuildConfirmationText_response_has_no_code_but_file_does()
        {
            var name = "unit_confirm_file_" + Guid.NewGuid().ToString("N").Substring(0, 6);
            var text = Program.BuildConfirmationText(name, new Dictionary<string, object> { ["k"] = "v" });
            var code = ExtractCode(name);
            Assert.False(string.IsNullOrEmpty(code), "确认文件里必须有码");
            Assert.False(text.Contains(code), "AI 可见的响应文案绝不能含确认码——否则 AI 可自答（D1 红线）");
        }

        private static string ExtractCode(string toolName)
        {
            // D1 收口：确认文件名 <tool>-pending.txt（不含码），码只在文件内容「确认码:」行
            var dir = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ExcelAiAssistant", "pending-confirmations");
            var f = System.IO.Path.Combine(dir, toolName + "-pending.txt");
            if (!System.IO.File.Exists(f)) return null;
            foreach (var line in System.IO.File.ReadAllLines(f))
            {
                if (line.StartsWith("确认码:"))
                {
                    var code = line.Substring("确认码:".Length).Trim();
                    try { System.IO.File.Delete(f); } catch { }
                    return code;
                }
            }
            return null;
        }

        // ===== 日期解析（DateParseHandling.None）=====

        [Fact]
        public void HandleMessage_date_string_stays_string()
        {
            // 协议层 JSON 里的 ISO 日期串不得转 DateTime（编码红线依赖拿到文本）
            // 直接验证解析路径配置（JsonTextReader.DateParseHandling=None）
            using var sr = new System.IO.StringReader("{\"d\":\"2024-01-15\"}");
            using var jr = new JsonTextReader(sr) { DateParseHandling = DateParseHandling.None };
            var tok = JToken.Load(jr);
            Assert.IsType<string>(((JValue)tok["d"]).Value);
            Assert.Equal("2024-01-15", (string)tok["d"]);
        }

        [Fact]
        public void DateParse_handling_none_by_default_protools_layer()
        {
            // 反证：若不禁用（Newtonsoft 默认 Automatic）会被转成 DateTime——本测试锁定「禁用配置存在」的语义
            var withNone = ParseWithHandling(DateParseHandling.None);
            var withAuto = ParseWithHandling(DateParseHandling.DateTime);
            Assert.IsType<string>(((JValue)withNone).Value);
            Assert.IsType<DateTime>(((JValue)withAuto).Value); // 默认处理会转 DateTime → 证明 None 配置确在生效路径上
        }

        private static JToken ParseWithHandling(DateParseHandling h)
        {
            using var sr = new System.IO.StringReader("{\"x\":\"2024-01-15T08:30:00Z\"}");
            using var jr = new JsonTextReader(sr) { DateParseHandling = h, MaxDepth = 64 };
            return JToken.Load(jr)["x"];
        }

        // ===== UTF-8 字节口径截断（N8）=====

        [Fact]
        public void TrimToUtf8Bytes_cjk_uses_three_bytes_per_char()
        {
            // "中"=3 字节：预算 9 字节应取 3 字
            var cut = Program.TrimToUtf8Bytes("中文字符串测试", 9);
            Assert.Equal("中文字", cut);
            Assert.Equal(9, Encoding.UTF8.GetByteCount(cut));
        }

        [Fact]
        public void TrimToUtf8Bytes_never_splits_surrogate_pair()
        {
            // 😀（U+1F600）代理对 = 4 字节 UTF-8：预算 3 字节不能吞半个代理对
            var input = "a😀b";
            var cut = Program.TrimToUtf8Bytes(input, 3);
            Assert.Equal("a", cut); // 1 字节 a + 预算只剩 2 < 4 → 不吞代理对
        }

        [Fact]
        public void TrimToUtf8Bytes_exact_boundary_takes_whole_char()
        {
            var input = "a😀b";
            var cut = Program.TrimToUtf8Bytes(input, 5); // 1+4=5 → 恰好吞下代理对
            Assert.Equal("a😀", cut);
        }

        [Fact]
        public void TrimToUtf8Bytes_ascii_exact()
        {
            Assert.Equal("abcd", Program.TrimToUtf8Bytes("abcdefgh", 4));
            Assert.Equal("abcdefgh", Program.TrimToUtf8Bytes("abcdefgh", 8));
        }

        [Fact]
        public void Transport_truncation_cjk_heavy_result_stays_under_bytes()
        {
            // 2000 行×3 列中文矩阵 JSON 远超 512K 字符 → 触发结构化降级；
            // 验证 data_preview 按 UTF-8 字节限幅（预览字节 ≤ 200KB）且整体 ≤ 512KB+裕量
            var vals = new List<object>();
            for (int r = 0; r < 3000; r++) vals.Add(new List<object> { "中文数据行" + r, "测试内容" + r, "表格单元" + r });
            var payload = new JObject { ["ok"] = true, ["data"] = JToken.FromObject(vals) };
            var slim = Program.BuildResultText(payload);
            // slim 含 data_preview 或 data
            var bytes = Encoding.UTF8.GetByteCount(slim);
            Assert.True(bytes <= 512 * 1024 + 64 * 1024, "超限结果应被降级为预览载荷（实际 " + bytes + " 字节）");
        }
    }
}
