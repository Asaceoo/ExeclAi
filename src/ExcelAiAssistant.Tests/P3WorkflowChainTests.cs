using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Core.Workflow;
using Newtonsoft.Json.Linq;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>P3-D3 多步串联验证：run_workflow 把上一步 Python 的结构化结果（xai_result 产物）
    /// 经 {stepN.xxx} 占位符注入下一步 VBA 的 args 数据通道（「Python 算→VBA 写」）。
    /// 纯逻辑测试：用仿 run_python_code / run_vba_snippet 的工具描述符验证引擎链路，不依赖真机。</summary>
    public sealed class P3WorkflowChainTests
    {
        [Fact]
        public async Task Python_structured_result_flows_into_next_vba_step_args()
        {
            var registry = new ToolRegistry();
            // step0：仿 run_python_code 返回匿名对象，result 为 xai_result 的结构化 JToken。
            registry.Register(new ToolDescriptor(
                "run_python_code", "compute", RiskLevel.Dangerous,
                _ => new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        exit_success = true,
                        exit_code = 0,
                        output = "42",
                        result = JToken.Parse("{\"total\": 42, \"avg\": 21.0, \"items\": [1, 2, 3]}"),
                        result_truncated = false,
                        state_reset = false
                    }
                }));
            Dictionary<string, object> capturedVbaArgs = null;
            // step1：仿 run_vba_snippet，把收到的参数原样捕获，便于断言 {step0.result} 是否被真实替换。
            registry.Register(new ToolDescriptor(
                "run_vba_snippet", "write", RiskLevel.Dangerous,
                args => { capturedVbaArgs = args; return new ToolResult { Ok = true, Data = new { written = 1 } }; }));

            var executor = new WorkflowExecutor(registry, (_, __) => Task.FromResult(true));
            var template = new WorkflowTemplate
            {
                Name = "Python算→VBA写",
                Steps = new List<WorkflowStep>
                {
                    new WorkflowStep { Tool = "run_python_code", Description = "Python 计算" },
                    new WorkflowStep
                    {
                        Tool = "run_vba_snippet",
                        Description = "VBA 写回",
                        Arguments = new Dictionary<string, object>
                        {
                            { "snippet", "Sub F()\n Range(\"A1\").Value = XaiArg(\"total\")\nEnd Sub" },
                            // 整串占位符 → ResolveValue 返回 step0 的结构化 result（JToken 原样，而非字符串化）。
                            { "args", "{step0.result}" }
                        }
                    }
                }
            };

            var result = await executor.RunAsync(template, null, null, 0, CancellationToken.None);
            Assert.True(result.Completed, "Python 算→VBA 写两步都应成功");
            Assert.Equal(2, result.Steps.Count);

            // 占位符必须被替换：args 是 step0 的 JObject，而非字面 "{step0.result}"。
            Assert.NotNull(capturedVbaArgs);
            Assert.True(capturedVbaArgs.TryGetValue("args", out var vbaArgs));
            var token = Assert.IsAssignableFrom<JToken>(vbaArgs);
            Assert.Equal(42, (int)token["total"]);
            Assert.Equal(21.0, (double)token["avg"]);
            Assert.Equal(3, ((JArray)token["items"]).Count);
        }

        [Fact]
        public async Task Embedded_placeholder_injects_json_text_into_nested_template()
        {
            // 模板作者把上一步结果嵌入字符串模板（如拼装 JSON），应替换为 JSON 文本而非原样保留。
            var registry = new ToolRegistry();
            registry.Register(new ToolDescriptor(
                "run_python_code", "compute", RiskLevel.Dangerous,
                _ => new ToolResult
                {
                    Ok = true,
                    Data = new { result = JToken.Parse("{\"total\": 7}") }
                }));
            Dictionary<string, object> captured = null;
            registry.Register(new ToolDescriptor(
                "consume", "read", RiskLevel.ReadOnly,
                args => { captured = args; return new ToolResult { Ok = true }; }));

            var executor = new WorkflowExecutor(registry, (_, __) => Task.FromResult(true));
            var template = new WorkflowTemplate
            {
                Name = "嵌入",
                Steps = new List<WorkflowStep>
                {
                    new WorkflowStep { Tool = "run_python_code" },
                    new WorkflowStep
                    {
                        Tool = "consume",
                        Arguments = new Dictionary<string, object>
                        {
                            { "message", "结果 {step0.result} 已就绪" }
                        }
                    }
                }
            };

            var result = await executor.RunAsync(template, null, null, 0, CancellationToken.None);
            Assert.True(result.Completed);
            Assert.Equal("结果 {\"total\":7} 已就绪", Convert.ToString(captured["message"]));
        }
    }
}
