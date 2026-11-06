using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ExcelAiAssistant.Core.Diagnostics;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Core.Workflow;
using Newtonsoft.Json;

namespace ExcelAiAssistant.Addin
{
    /// <summary>
    /// 用户自定义工具/模板管理协作者（阶段 B 自 SidebarBridge 抽出）。
    /// 方法体自 SidebarBridge 原样搬移（ListCustomTools/SaveUserTemplate/DeleteCustomTool/
    /// DeleteVbaTemplate/DeScopeVbaCode/UpdateCustomToolDescription + 2 个静态辅助）；
    /// 异常→JSON 兜底与 busy 发布经 executeAsync 委托保持单层结构（Bridge 侧透传）。
    /// </summary>
    internal sealed class CustomToolManager
    {
        private readonly Func<Func<Task<string>>, Task<string>> executeAsync;

        public CustomToolManager(Func<Func<Task<string>>, Task<string>> executeAsync)
        {
            this.executeAsync = executeAsync ?? throw new ArgumentNullException(nameof(executeAsync));
        }

        public Task<string> ListCustomTools(string category = null)
        {
            return executeAsync(() =>
            {
                var store = new CustomToolStore(CustomToolStore.DefaultRoot());
                store.SeedPresets(); // 三源之一：内置预置模板首次导入
                var tools = store.LoadByCategory(category)
                    .Select(t => new
                    {
                        name = t.Name,
                        description = t.Description,
                        code = t.Code,
                        created_at = t.CreatedAt,
                        updated_at = t.UpdatedAt,
                        category = t.EffectiveCategory(),
                        tags = t.Tags ?? new List<string>(),
                        version = string.IsNullOrWhiteSpace(t.Version) ? "1.0" : t.Version
                    });
                return Task.FromResult(JsonConvert.SerializeObject(new { ok = true, tools }));
            });
        }

        /// <summary>把用户手写的 VBA 代码保存为「用户模板」（三源之一：user），供模板库动态加载复用。</summary>
        public Task<string> SaveUserTemplate(string name, string description, string code, string tagsText)
        {
            return executeAsync(() =>
            {
                var store = new CustomToolStore(CustomToolStore.DefaultRoot());
                var baseName = string.IsNullOrWhiteSpace(name) ? null : name.Trim();
                var nameText = store.FindAvailableName(string.IsNullOrWhiteSpace(baseName) ? "my_template" : baseName);
                var tags = (tagsText ?? string.Empty)
                    .Split(new[] { ',', '，', ';', '；', ' ' }, StringSplitOptions.RemoveEmptyEntries)
                    .Take(8).ToList();
                var result = store.Save(
                    nameText,
                    string.IsNullOrWhiteSpace(description) ? "用户手写模板：" + nameText : description.Trim(),
                    code,
                    ToolSourceCategory.User,
                    tags,
                    null);
                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = result.Ok,
                    name = result.Ok ? nameText : null,
                    renamed = result.Ok && !string.Equals(baseName, nameText, StringComparison.OrdinalIgnoreCase)
                        ? new { from = baseName, to = nameText }
                        : null,
                    error = result.Ok ? null : (result.Data?.ToString() ?? result.ErrorCode),
                    message = result.Ok ? result.Data : null
                }));
            });
        }

        /// <summary>删除一个已沉淀的自定义工具。</summary>
        public Task<string> DeleteCustomTool(string name)
        {
            return executeAsync(() =>
            {
                var store = new CustomToolStore(CustomToolStore.DefaultRoot());
                var result = store.Delete(name ?? string.Empty);
                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = result.Ok,
                    error = result.Ok ? null : (result.Data?.ToString() ?? result.ErrorCode),
                    message = result.Ok ? result.Data : null
                }));
            });
        }

        /// <summary>VBA 面板「删除模板」：删除已沉淀的用户/AI 模板（预置模板保护，禁止删除）。</summary>
        public Task<string> DeleteVbaTemplate(string name)
        {
            return executeAsync(() =>
            {
                var store = new CustomToolStore(CustomToolStore.DefaultRoot());
                var tool = store.LoadAll().FirstOrDefault(t =>
                    string.Equals(t.Name, name ?? string.Empty, StringComparison.OrdinalIgnoreCase));
                if (tool == null)
                {
                    return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = $"模板“{name}”不存在。" }));
                }

                if (string.Equals(tool.EffectiveCategory(), ToolSourceCategory.Preset, StringComparison.Ordinal))
                {
                    return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = "内置预置模板不可删除。" }));
                }

                var result = store.Delete(tool.Name);
                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = result.Ok,
                    error = result.Ok ? null : (result.Data?.ToString() ?? result.ErrorCode),
                    message = result.Ok ? result.Data : null
                }));
            });
        }

        /// <summary>VBA 代码「去作用区域」：把固化区域（Range("C5:F22") 等）改写为 {selection} 占位符，返回改写后的代码。
        /// 供 VBA 面板「去作用区域后保存」使用，针对手写与 AI沉淀代码。</summary>
        public Task<string> DeScopeVbaCode(string code)
        {
            return executeAsync(() =>
            {
                var original = code ?? string.Empty;
                var rewritten = FlowScopeNormalizer.DescopeVbaCode(original);
                var changed = !string.Equals(rewritten, original, StringComparison.Ordinal);
                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = true,
                    code = rewritten,
                    changed,
                    replaced = changed ? FlowScopeNormalizerReferences(original) : 0
                }));
            });
        }

        private static int FlowScopeNormalizerReferences(string code)
        {
            int n = 0;
            if (code == null) return 0;
            n += CountRangeLiterals(code);
            return n;
        }

        private static int CountRangeLiterals(string code)
        {
            var pattern = @"(?i)(\.?Range\s*\(\s*)""(?:[^""]*!)?\$?[A-Z]{1,3}\$?[0-9]+(?:\s*:\s*\$?[A-Z]{1,3}\$?[0-9]+)?""\s*\)";
            return System.Text.RegularExpressions.Regex.Matches(code, pattern).Count;
        }

        /// <summary>更新已沉淀工具的用途说明（仅改描述，不覆盖代码）。</summary>
        public Task<string> UpdateCustomToolDescription(string name, string description)
        {
            return executeAsync(() =>
            {
                var store = new CustomToolStore(CustomToolStore.DefaultRoot());
                var tool = store.LoadAll().FirstOrDefault(t =>
                    string.Equals(t.Name, name ?? string.Empty, StringComparison.OrdinalIgnoreCase));
                if (tool == null)
                {
                    return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = $"自定义工具“{name}”不存在。" }));
                }

                if (string.IsNullOrWhiteSpace(description))
                {
                    return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = "描述不能为空。" }));
                }

                tool.Description = description.Trim();
                tool.UpdatedAt = DateTime.Now.ToString("s");
                try
                {
                    System.IO.File.WriteAllText(
                        store.ToFile(tool.Name),
                        JsonConvert.SerializeObject(tool, Formatting.Indented));
                    return Task.FromResult(JsonConvert.SerializeObject(new { ok = true, name = tool.Name, message = "描述已更新。" }));
                }
                catch (Exception ex)
                {
DiagnosticsLog.Write("[LogCover] CustomToolManager.cs:187  {ex.GetType().Name}: {ex.Message}");
                    return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = "保存失败：" + ex.Message }));
                }
            });
        }
    }
}
