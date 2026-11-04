using System;
using System.Collections.Generic;

namespace ExcelAiAssistant.Core.Workflow
{
    /// <summary>
    /// 流程定义模型：名称 + 描述 + 有序步骤列表。
    /// 每步 = 工具名 + 参数模板 + 前置说明；命中工具后风险级别取自工具注册表，保证与工具库不漂移。
    /// Arguments 中可用 ``{占位符}`` 形式的运行时参数，执行时由调用方覆盖。
    /// </summary>
    public sealed class WorkflowTemplate
    {
        public string Name { get; set; }

        public string Description { get; set; }

        public List<WorkflowStep> Steps { get; set; } = new List<WorkflowStep>();

        public bool IsBuiltIn { get; set; }

        public DateTime CreatedAtUtc { get; set; }
    }

    public sealed class WorkflowStep
    {
        public string Tool { get; set; }

        public string Description { get; set; }

        public Dictionary<string, object> Arguments { get; set; } = new Dictionary<string, object>();

        /// <summary>该步失败后是否允许单独重试（默认 true）。</summary>
        public bool AllowRetry { get; set; } = true;

        /// <summary>可选的步骤后自检注入：主步骤成功后自动执行对应 verify 工具并校验 verdict，未达标即停在当前步。</summary>
        public WorkflowVerify Verify { get; set; }
    }

    /// <summary>
    /// 流程步骤的自检规格：主步骤执行成功后自动注入并校验。
    /// Arguments 支持 ``{stepN.属性}`` / ``{selection}`` 占位符，执行时与主步骤一致地解析。
    /// </summary>
    public sealed class WorkflowVerify
    {
        public string Tool { get; set; }

        public string Description { get; set; }

        public Dictionary<string, object> Arguments { get; set; } = new Dictionary<string, object>();
    }
}