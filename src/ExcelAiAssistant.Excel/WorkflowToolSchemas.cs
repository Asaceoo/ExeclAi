using System.Collections.Generic;

namespace ExcelAiAssistant.Excel
{
    /// <summary>run_workflow 工具的 JSON Schema 参数描述（与 M3ToolSchemas 结构一致）。</summary>
    public static class WorkflowToolSchemas
    {
        public static Dictionary<string, object> RunWorkflow()
        {
            var map = new Dictionary<string, object>();
            map["workflow"] = Property("workflow", "string", "流程名称，如「数据体检并生成报告」。");
            map["description"] = Property("description", "string", "可选，本次执行说明。");

            var schema = new Dictionary<string, object>
            {
                ["type"] = "object",
                ["properties"] = map,
                ["required"] = new[] { "workflow" }
            };
            return schema;
        }

        private static Dictionary<string, object> Property(string name, string type, string description = null)
        {
            var result = new Dictionary<string, object> { ["name"] = name, ["type"] = type };
            if (description != null)
            {
                result["description"] = description;
            }

            return result;
        }
    }
}