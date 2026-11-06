using System.Collections.Generic;
using ExcelAiAssistant.Core.Models;
using Newtonsoft.Json;

namespace ExcelAiAssistant.Core.Agent
{
    /// <summary>
    /// 工具执行记录的可序列化展示视图，成功时给截断摘要、失败时给友好错误。
    /// 字段名以蛇形命名，与前端侧边栏 / 会话持久化保持一致。
    /// </summary>
    public sealed class ToolEventView
    {
        [JsonProperty("tool")]
        public string Tool { get; set; }

        [JsonProperty("ok")]
        public bool Ok { get; set; }

        [JsonProperty("confirmed")]
        public bool? Confirmed { get; set; }

        [JsonProperty("retryable")]
        public bool Retryable { get; set; }

        [JsonProperty("error_code")]
        public string ErrorCode { get; set; }

        [JsonProperty("error_message")]
        public string ErrorMessage { get; set; }

        [JsonProperty("summary")]
        public string Summary { get; set; }

        /// <summary>完整结果 JSON（供前端点击「详情」展开查看原始数据，不参与摘要行展示）。</summary>
        [JsonProperty("detail")]
        public string Detail { get; set; }

        public static ToolEventView FromRecord(ToolEventRecord record)
        {
            if (record == null)
            {
                return null;
            }

            var result = record.Result;
            var succeeded = result?.Ok == true;
            return new ToolEventView
            {
                Tool = record.Tool,
                Ok = succeeded,
                Confirmed = record.Confirmed,
                Retryable = !succeeded && record.Args != null,
                ErrorCode = succeeded ? null : result?.ErrorCode,
                ErrorMessage = succeeded ? null : Truncate(result?.Data?.ToString() ?? result?.ErrorCode ?? "操作失败。"),
                Summary = succeeded ? Truncate(ToolSummaryBuilder.Summarize(record.Tool, result.Data)) : null,
                Detail = succeeded ? Truncate(SerializeData(result.Data), 4000) : null
            };
        }
        private static string SerializeData(object data)
        {
            if (data == null)
            {
                return null;
            }

            if (data is string text)
            {
                return text;
            }

            return JsonConvert.SerializeObject(data, Formatting.None);
        }

        private static string Truncate(string value, int max = 140)
        {
            if (string.IsNullOrEmpty(value))
            {
                return value ?? string.Empty;
            }

            value = value.Trim();
            return value.Length <= max ? value : value.Substring(0, max) + "…";
        }
    }

    /// <summary>
    /// 把工具执行记录转成面向界面展示的视图，避免把完整的执行结果整份发给前端。
    /// </summary>
    public static class ToolEventViewBuilder
    {
        public static object Build(ToolEventRecord record)
        {
            return ToolEventView.FromRecord(record);
        }
    }

    /// <summary>
    /// 一轮助手消息里按真实时间顺序组织的「步骤」：type=reasoning 为模型思考片段，
    /// type=tool 为一次工具调用及其结果。前端据此交替渲染「思考 → 调用工具 → 拿结果 → …」。
    /// </summary>
    public sealed class AgentStepView
    {
        /// <summary>reasoning（思考片段）| tool（工具调用）。</summary>
        [JsonProperty("type")]
        public string Type { get; set; }

        /// <summary>type=reasoning 时的推理文字。</summary>
        [JsonProperty("text")]
        public string Text { get; set; }

        /// <summary>type=tool 时的工具事件视图（含结果）。</summary>
        [JsonProperty("tool_event")]
        public ToolEventView ToolEvent { get; set; }
    }
}