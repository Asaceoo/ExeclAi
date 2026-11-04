using System;
using System.Collections.Generic;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;

namespace ExcelAiAssistant.Core.Safety
{
    public sealed class SafetyDecision
    {
        public bool Allowed { get; set; }

        public bool RequiresConfirmation { get; set; }

        public string Reason { get; set; }
    }

    public sealed class SafetyOptions
    {
        public bool ConfirmNormalWrites { get; set; } = true;

        public bool EnableHighRiskTools { get; set; }
    }

    public sealed class SafetyPipeline
    {
        private readonly SafetyOptions options;

        public SafetyPipeline(SafetyOptions options)
        {
            this.options = options ?? new SafetyOptions();
        }

        public SafetyDecision Evaluate(ToolDescriptor tool)
        {
            if (tool == null)
            {
                throw new ArgumentNullException(nameof(tool));
            }

            if (tool.RiskLevel == RiskLevel.Dangerous && !options.EnableHighRiskTools)
            {
                return new SafetyDecision
                {
                    Allowed = false,
                    RequiresConfirmation = false,
                    Reason = "高危工具默认关闭，需要在设置中显式启用。"
                };
            }

            return new SafetyDecision
            {
                Allowed = true,
                RequiresConfirmation =
                    tool.RiskLevel >= RiskLevel.StructuralChange ||
                    (tool.RiskLevel == RiskLevel.NormalWrite && options.ConfirmNormalWrites),
                Reason = $"风险等级 L{(int)tool.RiskLevel}"
            };
        }
    }
}
