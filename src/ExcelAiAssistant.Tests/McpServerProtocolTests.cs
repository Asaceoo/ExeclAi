using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.McpServer;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// MCP server 协议层纯逻辑单测（2026-11-03 对抗性审查收敛批次）：
    /// 三视角发现 F1/F2/F3/F7 的回归锁定 + schema 转换契约。
    /// </summary>
    public class McpServerProtocolTests
    {
        [Fact]
        public void CleanSchemaNode_strips_internal_name_keys()
        {
            // 与 M3ToolSchemas 实际产物同构：纯 Dictionary + object[]（非 JToken 混合体）
            var schema = new Dictionary<string, object>
            {
                ["name"] = "ignored",
                ["type"] = "object",
                ["properties"] = new Dictionary<string, object>
                {
                    ["address"] = new Dictionary<string, object>
                    {
                        ["name"] = "address",
                        ["type"] = "string",
                        ["description"] = "区域"
                    }
                }
            };
            var cleaned = (JObject)JToken.FromObject(Program.CleanSchemaNode(schema));
            var obj = (JObject)cleaned["properties"]!["address"]!;
            Assert.False(obj.ContainsKey("name"));
            Assert.Equal("string", (string)obj["type"]);
            Assert.Equal("区域", (string)obj["description"]);
        }

        [Fact]
        public void CleanSchemaNode_keeps_required_and_enum_arrays()
        {
            // required/enum 数组按 M3ToolSchemas 帮助方法的实际形状：object[] / string[]
            var schema = new Dictionary<string, object>
            {
                ["type"] = "object",
                ["required"] = new object[] { "address" },
                ["properties"] = new Dictionary<string, object>
                {
                    ["operator"] = new Dictionary<string, object>
                    {
                        ["name"] = "operator",
                        ["type"] = "string",
                        ["enum"] = new[] { "equals", "contains" }
                    }
                }
            };
            var cleaned = (JObject)JToken.FromObject(Program.CleanSchemaNode(schema));
            Assert.Equal(new[] { "address" }, ((JArray)cleaned["required"]!).Select(t2 => (string)t2!).ToArray());
            var op = (JObject)cleaned["properties"]!["operator"]!;
            Assert.False(op.ContainsKey("name"));
            Assert.Equal(2, ((JArray)op["enum"]!).Count);
        }

        [Fact]
        public void RiskHint_prefixes_by_risk_level()
        {
            var ro = new ToolDescriptor("t1", "只读工具", RiskLevel.ReadOnly, _ => new ToolResult { Ok = true });
            var w = new ToolDescriptor("t2", "写入工具", RiskLevel.NormalWrite, _ => new ToolResult { Ok = true });
            var s = new ToolDescriptor("t3", "结构工具", RiskLevel.StructuralChange, _ => new ToolResult { Ok = true });
            var d = new ToolDescriptor("t4", "危险工具", RiskLevel.Dangerous, _ => new ToolResult { Ok = true });
            Assert.Equal("", Program.RiskHint(ro));
            Assert.StartsWith("[写入]", Program.RiskHint(w));
            Assert.StartsWith("[结构变更]", Program.RiskHint(s));
            Assert.StartsWith("[危险级·需确认]", Program.RiskHint(d));
        }

        [Fact]
        public void BuildResultText_carries_full_result_shape()
        {
            var result = new ToolResult
            {
                Ok = true,
                Data = new { mode = "markdown", word_count = 12 },
                Warning = "样本警告",
                AffectedRanges = new List<string> { "Sheet1!A1:B4" }
            };
            var text = Program.BuildResultText(result);
            var payload = JObject.Parse(text);
            Assert.True((bool)payload["ok"]!);
            Assert.Equal(12, (int)payload["data"]!["word_count"]!);
            Assert.Equal("样本警告", (string)payload["warning"]!);
            Assert.Equal("Sheet1!A1:B4", (string)((JArray)payload["affected_ranges"]!)[0]!);
        }

        [Fact]
        public void BuildResultText_parses_json_string_data_as_object()
        {
            // Data 为 JSON 字符串时解析为对象（服务层常见：Data=JsonConvert.SerializeObject(...)）
            var result = new ToolResult { Ok = true, Data = "{\"a\": 1}" };
            var payload = JObject.Parse(Program.BuildResultText(result));
            Assert.Equal(1, (int)payload["data"]!["a"]!);
        }

        [Fact]
        public void BuildResultText_falls_back_to_string_for_non_json_data()
        {
            // Data 是非 JSON 纯文本时不抛异常，降级为字符串（F3 审查发现的隐式契约）
            var result = new ToolResult { Ok = false, ErrorCode = "X", Data = "这是纯文本数据不是JSON" };
            var payload = JObject.Parse(Program.BuildResultText(result));
            Assert.False((bool)payload["ok"]!);
            Assert.Equal("这是纯文本数据不是JSON", (string)payload["data"]!);
            Assert.Equal("X", (string)payload["error_code"]!);
        }

        [Fact]
        public void BuildResultText_marks_transport_truncation_over_512kb()
        {
            // R4 回归锁定（三视角审查升级契约）：超大 Data 必须结构化降级——
            // ① 输出保持合法 JSON（不可 Substring 拦腰砍）；② 带 transport_truncated 标记；
            // ③ hint 指明 max_cells/缩小 address 的分页路径；④ data_preview 存在且有限。
            var big = new string('x', 700 * 1024);
            var result = new ToolResult { Ok = true, Data = big };
            var text = Program.BuildResultText(result);

            var payload = JObject.Parse(text); // ① 必须可解析——非法 JSON 直接抛
            Assert.True((bool)payload["transport_truncated"]!);                       // ②
            // ③ hint 动态化（Y2-2）：无 tool 上下文时给通用缩小范围指引；max_cells 分支由
            //    Truncation_hint_adapts_to_tool_schema 专项锁定
            Assert.Contains("缩小", (string)payload["hint"]!);
            Assert.NotNull(payload["data_preview"]);                                  // ④
            Assert.True(text.Length < 512 * 1024, "瘦身后必须显著小于 512KB 上限");
        }

        [Fact]
        public void Unknown_tool_in_readonly_mode_lists_readonly_tools()
        {
            // F2 回归锁定：只读模式下未知（已被移除的写）工具必须给模式说明+可用只读工具清单
            var registry = new ToolRegistry();
            registry.Register(new ToolDescriptor("read_only_alpha", "只读A", RiskLevel.ReadOnly, _ => new ToolResult { Ok = true }));
            registry.Register(new ToolDescriptor("read_only_beta", "只读B", RiskLevel.ReadOnly, _ => new ToolResult { Ok = true }));
            Program._registry = registry;
            Program._readOnlyMode = true;

            var text = Program.BuildUnknownToolText("write_excel_cell");
            Assert.Contains("只读模式", text);
            Assert.Contains("--allow-write", text);
            Assert.Contains("read_only_alpha", text);
            Assert.Contains("read_only_beta", text);
            Assert.Contains("请勿重试写操作", text);

            // 读写模式下不附加模式说明
            Program._readOnlyMode = false;
            var text2 = Program.BuildUnknownToolText("nope");
            Assert.DoesNotContain("只读模式", text2);
        }

        [Fact]
        public void CleanSchemaNode_preserves_business_name_parameter()
        {
            // R3 回归锁定：rename/add_excel_worksheet 的业务参数 name 绝不能被路由键剥除逻辑干掉
            var schema = new Dictionary<string, object>
            {
                ["type"] = "object",
                ["required"] = new object[] { "name" },
                ["properties"] = new Dictionary<string, object>
                {
                    ["name"] = new Dictionary<string, object>
                    {
                        ["name"] = "name",           // 内部路由键（要剥）
                        ["type"] = "string",
                        ["description"] = "新名称。"
                    },
                    ["sheet"] = new Dictionary<string, object>
                    {
                        ["name"] = "sheet",
                        ["type"] = "string"
                    }
                }
            };
            var cleaned = (JObject)JToken.FromObject(Program.CleanSchemaNode(schema));
            var nameProp = (JObject)cleaned["properties"]!["name"]!;
            // 业务参数 name 键保留在 properties 里（键名层面）
            Assert.NotNull(cleaned["properties"]!["name"]);
            Assert.Equal("string", (string)nameProp["type"]);
            Assert.Equal("新名称。", (string)nameProp["description"]);
            // 属性定义对象内部的 name 路由键剥掉
            Assert.False(nameProp.ContainsKey("name"));
            // required 数组与 properties 键一致（schema 不再自相矛盾）
            Assert.Equal(new[] { "name" }, ((JArray)cleaned["required"]!).Select(t2 => (string)t2!).ToArray());
        }

        [Fact]
        public void Slim_truncates_oversized_warning()
        {
            // Y2-1 回归锁定：600KB warning 原样透传会再次超 512KB，必须同规格 preview 截断
            var big = new string('w', 600 * 1024);
            var result = new ToolResult { Ok = true, Warning = big, Data = "ok" };
            var text = Program.BuildResultText(result);
            Assert.True(text.Length < 512 * 1024, "含超大 warning 时瘦身后仍须低于上限");
            var payload = JObject.Parse(text);
            Assert.Contains("已截断", (string)payload["warning"]!);
        }

        [Fact]
        public void Truncation_hint_adapts_to_tool_schema()
        {
            // Y2-2 回归锁定：hint 按工具 schema 动态生成——有 max_cells 的工具提该参数，否则给缩小范围指引
            var withMaxCells = new ToolDescriptor("reader", "读", RiskLevel.ReadOnly, _ => new ToolResult { Ok = true },
                new Dictionary<string, object> { ["properties"] = new Dictionary<string, object> { ["max_cells"] = new Dictionary<string, object> { ["type"] = "integer" } } });
            var without = new ToolDescriptor("renderer", "渲染", RiskLevel.ReadOnly, _ => new ToolResult { Ok = true });

            var r1 = new ToolResult { Ok = true, Data = new string('x', 700 * 1024) };
            var text1 = Program.BuildResultText(r1, withMaxCells);
            Assert.Contains("max_cells", (string)JObject.Parse(text1)["hint"]!);

            var text2 = Program.BuildResultText(r1, without);
            Assert.DoesNotContain("max_cells", (string)JObject.Parse(text2)["hint"]!);
            Assert.Contains("缩小", (string)JObject.Parse(text2)["hint"]!);
        }
    }
}
