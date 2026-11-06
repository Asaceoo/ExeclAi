using System.Collections.Generic;

namespace ExcelAiAssistant.Core.Skills
{
    public sealed class Skill
    {
        public string Name { get; set; }

        public string Description { get; set; }

        public string Version { get; set; }

        public string RiskPolicy { get; set; } = "normal";

        public List<string> Tools { get; set; } = new List<string>();

        /// <summary>技能所属分类，用于目录分组展示；缺省归入「通用」。</summary>
        public string Category { get; set; } = "通用";

        /// <summary>技能类型：quick（快捷型，单步少步，可进前端 chip 行）或 orchestration（编排型，多步流程，需 AI 多轮对话）。</summary>
        public string Type { get; set; } = "quick";

        /// <summary>触发场景词（便于 AI 命中该技能）。</summary>
        public List<string> Triggers { get; set; } = new List<string>();

        /// <summary>典型用法示例（一句话，便于引导调用）。</summary>
        public List<string> Examples { get; set; } = new List<string>();

        public string Directory { get; set; }

        public string Body { get; set; }

        public bool Valid { get; set; } = true;

        public string InvalidReason { get; set; }

        public bool IsBuiltIn { get; set; }

        /// <summary>技能来源：builtin（产品内置）/ market（URL 或粘贴 SKILL.md 安装）/ local（手写或旧用户技能）。</summary>
        public string Source { get; set; }
    }
}