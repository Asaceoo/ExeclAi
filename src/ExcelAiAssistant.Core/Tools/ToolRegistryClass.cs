using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAiAssistant.Core.Models;

namespace ExcelAiAssistant.Core.Tools
{
    public sealed class ToolRegistry
    {
        private readonly Dictionary<string, ToolDescriptor> tools =
            new Dictionary<string, ToolDescriptor>(StringComparer.OrdinalIgnoreCase);

        public void Register(ToolDescriptor tool)
        {
            if (tool == null)
            {
                throw new ArgumentNullException(nameof(tool));
            }

            if (tools.ContainsKey(tool.Name))
            {
                throw new InvalidOperationException($"重复注册工具：{tool.Name}");
            }

            tools.Add(tool.Name, tool);
        }

        public IReadOnlyCollection<ToolDescriptor> Tools => tools.Values.ToList().AsReadOnly();

        /// <summary>
        /// 按名称移除工具（宿主感知过滤用：WPS 无 VBA 环境时剔除 VBA 执行类工具）。
        /// 返回是否发生了移除；名称不存在时返回 false。
        /// </summary>
        public bool Remove(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return false;
            }

            return tools.Remove(name.Trim());
        }

        public bool TryGet(string name, out ToolDescriptor tool)
        {
            return tools.TryGetValue(name, out tool);
        }

        /// <summary>
        /// 原位替换已注册工具（M17 守卫包装用：保留注册名与风险级别，仅换 Execute）。
        /// 名称不存在时抛 KeyNotFoundException——替换目标必须真实存在，防拼写漂移静默失效。
        /// </summary>
        public void Replace(ToolDescriptor tool)
        {
            if (tool == null)
            {
                throw new ArgumentNullException(nameof(tool));
            }

            if (!tools.TryGetValue(tool.Name, out var existing))
            {
                throw new KeyNotFoundException($"替换目标未注册：{tool.Name}");
            }

            // T1（审查者第二轮）：与 XML 注释契约对齐——Replace 只换 Execute 不变级别。
            // 等值校验同时挡住静默降级与静默升级（如需换级显式 Remove+Register）。
            if (tool.RiskLevel != existing.RiskLevel)
            {
                throw new InvalidOperationException($"Replace 不允许变更风险级别：{tool.Name} {existing.RiskLevel} → {tool.RiskLevel}（如需换级请显式 Remove+Register）");
            }

            tools[tool.Name] = tool;
        }

        public RiskLevel GetRiskLevel(string name)
        {
            return TryGet(name, out var tool)
                ? tool.RiskLevel
                : throw new KeyNotFoundException($"未知工具：{name}");
        }

        public void RestrictTo(IEnumerable<string> allowed)
        {
            if (allowed == null)
            {
                throw new ArgumentNullException(nameof(allowed));
            }

            var allowedSet = new HashSet<string>(allowed, StringComparer.OrdinalIgnoreCase);
            var removals = tools.Keys.Where(name => !allowedSet.Contains(name)).ToList();
            foreach (var name in removals)
            {
                tools.Remove(name);
            }
        }
    }
}
