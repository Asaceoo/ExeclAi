using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using ExcelAiAssistant.Core.Agent;
using ExcelAiAssistant.Core.Configuration;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

internal static class Program
{
    private static int Main(string[] args)
    {
        var newKey = args.Length > 0 ? args[0] : string.Empty;
        var cfgPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ExcelAiAssistant", "models.json");

        var store = new ConfigurationStore(
            cfgPath,
            apiKey => ProtectedData.Protect(
                Encoding.UTF8.GetBytes(apiKey),
                Encoding.UTF8.GetBytes("ExcelAiAssistant:v1"),
                DataProtectionScope.CurrentUser),
            cipher => Encoding.UTF8.GetString(
                ProtectedData.Unprotect(
                    cipher,
                    Encoding.UTF8.GetBytes("ExcelAiAssistant:v1"),
                    DataProtectionScope.CurrentUser)));

        var profiles = new List<ModelProfile>(store.Load());
        var target = profiles.FirstOrDefault(p =>
            p.BaseUrl != null &&
            p.BaseUrl.IndexOf("askdiandian", StringComparison.OrdinalIgnoreCase) >= 0);
        if (target == null)
        {
            Console.WriteLine("ERROR 未找到小红书方案");
            return 1;
        }

        Console.WriteLine("方案: " + target.Name + " | " + target.ModelName);
        Console.WriteLine("存储 key 是否与用户提供 key 相同: " +
            (string.Equals(target.ApiKey, newKey, StringComparison.Ordinal) ? "是" : "否"));

        // 复现真实聊天流：带工具 + 流式
        try
        {
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(90));
            using var client = new OpenAiCompatibleClient(target);
            var tools = new List<ToolDescriptor>
            {
                new ToolDescriptor("get_current_excel_context", "获取当前 Excel 上下文（工作簿/工作表/选区）。",
                    RiskLevel.ReadOnly, _ => new ToolResult { Ok = true },
                    new Dictionary<string, object> { ["type"] = "object", ["properties"] = new JObject() }),
                new ToolDescriptor("write_excel_cell", "向指定单元格写入值。",
                    RiskLevel.NormalWrite, _ => new ToolResult { Ok = true },
                    new Dictionary<string, object>
                    {
                        ["type"] = "object",
                        ["properties"] = new Dictionary<string, object> { ["cell"] = new JObject { ["type"] = "string" } }
                    })
            };
            var messages = new List<ChatMessage>
            {
                ChatMessage.System("你是 Excel 助手。"),
                ChatMessage.User("帮我看看当前选区的值")
            };
            var result = client.CompleteAsync(messages, tools, cts.Token).GetAwaiter().GetResult();
            Console.WriteLine("带工具聊天结果: Content=" + Trunc(result.Content));
            Console.WriteLine("  Reasoning 长度=" + (result.Reasoning ?? "").Length);
            Console.WriteLine("  ToolCalls 数量=" + result.ToolCalls.Count);
            if (result.ToolCalls.Count > 0)
            {
                Console.WriteLine("  首个工具调用=" + result.ToolCalls[0].Name + " args=" + Trunc(JsonConvert.SerializeObject(result.ToolCalls[0].Arguments)));
            }
            Console.WriteLine("  FallbackUsed=" + result.FallbackUsed);
            Console.WriteLine("  结果: " + (result != null ? "OK" : "NULL"));
        }
        catch (Exception ex)
        {
            Console.WriteLine("带工具聊天异常: " + ex.GetType().Name + ": " + Trunc(ex.Message));
            var inner = ex.InnerException;
            var depth = 0;
            while (inner != null && depth < 5)
            {
                Console.WriteLine("  [内层" + depth + "] " + inner.GetType().Name + ": " + Trunc(inner.Message));
                inner = inner.InnerException;
                depth++;
            }
            return 1;
        }
        return 0;
    }

    private static string Trunc(string s)
    {
        if (string.IsNullOrEmpty(s)) return "(空)";
        return s.Length <= 200 ? s : s.Substring(0, 200) + "...";
    }
}
