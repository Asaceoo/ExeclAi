using System;
using ExcelAiAssistant.Core.Diagnostics;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelAiAssistant.Core.Models;
using Newtonsoft.Json;

namespace ExcelAiAssistant.Core.Tools
{
    /// <summary>自定义 VBA 工具/模板信息（存为一个 JSON 文件），三源一套统一存储：
    /// 预置(preset) / AI沉淀(ai) / 用户手写(user)。</summary>
    public sealed class CustomToolInfo
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public string Code { get; set; }
        public string CreatedAt { get; set; }
        public string UpdatedAt { get; set; }
        /// <summary>来源/分类：preset=预置、ai=AI沉淀、user=用户手写。旧数据缺省视为 ai。</summary>
        public string Category { get; set; }
        /// <summary>标签，用于模板搜索/筛选。</summary>
        public List<string> Tags { get; set; }
        /// <summary>模板版本（如 1.0；预置片段随产品升级迭代）。</summary>
        public string Version { get; set; }

        /// <summary>返回规范化分类；null/空白一律回落为 ai（向后兼容旧沉淀）。</summary>
        public string EffectiveCategory()
        {
            return string.IsNullOrWhiteSpace(Category) ? "ai" : Category.Trim();
        }
    }

    /// <summary>三源分类常量。</summary>
    public static class ToolSourceCategory
    {
        public const string Preset = "preset";
        public const string Ai = "ai";
        public const string User = "user";

        public static bool IsKnown(string category)
        {
            return string.Equals(category, Preset, StringComparison.Ordinal)
                || string.Equals(category, Ai, StringComparison.Ordinal)
                || string.Equals(category, User, StringComparison.Ordinal);
        }
    }

    /// <summary>
    /// 把 AI 自写代码沉淀为可复用工具的本地库（%LocalAppData%\XaiAssistant\custom-tools\）。
    /// 仅负责元数据（名称/描述/代码）的持久化与注册；真实的 VBA 执行由注册处的 runner 注入，
    /// 与 run_vba_snippet 走同一条受控安全链。工具名经严格白名单校验，防路径穿越与重复污染。
    /// </summary>
    public sealed class CustomToolStore
    {
        private static readonly Regex NamePattern = new Regex("^[A-Za-z][A-Za-z0-9_]*$");
        private readonly string _root;

        public CustomToolStore(string rootDir)
        {
            _root = rootDir;
            Directory.CreateDirectory(_root);
        }

        public static string DefaultRoot()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "XaiAssistant", "custom-tools");
        }

        public static bool IsValidName(string name)
        {
            return !string.IsNullOrWhiteSpace(name)
                && name.Length <= 40
                && NamePattern.IsMatch(name);
        }

        public string ToFile(string name)
        {
            if (!IsValidName(name))
            {
                throw new InvalidOperationException(
                    "工具名不合法：仅允许以字母开头，可由字母/数字/下划线组成（≤40 字符）。");
            }

            return Path.Combine(_root, name + ".json");
        }

        /// <summary>判断某工具名是否已被占用（用于沉淀时自动去重）。</summary>
        public bool Exists(string name)
        {
            return !string.IsNullOrWhiteSpace(name)
                && IsValidName(name)
                && File.Exists(ToFile(name));
        }

        /// <summary>在 baseName 基础上找到一个未被占用的工具名：base、base_2、base_3…。超出长度则截短前缀。</summary>
        public string FindAvailableName(string baseName)
        {
            if (string.IsNullOrWhiteSpace(baseName))
            {
                baseName = "vba_tool";
            }

            var candidate = SanitizeName(baseName);
            var original = candidate;
            var suffix = 2;
            while (Exists(candidate))
            {
                var tail = "_" + suffix;
                var maxLen = 40 - tail.Length;
                var prefix = original.Length > maxLen ? original.Substring(0, maxLen) : original;
                var next = prefix.TrimEnd('_') + tail;
                if (string.Equals(next, candidate, StringComparison.Ordinal) || !IsValidName(next))
                {
                    break;
                }

                candidate = next;
                suffix++;
                if (suffix > 1000)
                {
                    break;
                }
            }

            return candidate;
        }

        /// <summary>把任意来源的工具名收敛为合法 ASCII 名：剥离中文等非 [A-Za-z0-9_] 字符、确保以字母开头（覆盖「Sub 汇总」这类中文宏名的沉淀）。</summary>
        private static string SanitizeName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return "vba_tool";
            }

            var sb = new System.Text.StringBuilder();
            foreach (var ch in name.Trim())
            {
                var ascii = (ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z') || ch == '_';
                var digit = ch >= '0' && ch <= '9';
                if ((ascii && sb.Length > 0) || (digit && sb.Length > 0))
                {
                    sb.Append(ch);
                }
                else if (ascii && sb.Length == 0 && !digit)
                {
                    // 首字符允许字母或下划线，后续再统一兜底
                    sb.Append(ch);
                }
            }

            var s = sb.ToString();
            if (s.Length == 0 || !(char.IsLetter(s[0])))
            {
                if (s.Length > 0 && s[0] == '_')
                {
                    s = string.Empty;
                }
                else
                {
                    s = Regex.Replace(s, "^[^A-Za-z]+", string.Empty);
                }
            }

            if (string.IsNullOrEmpty(s) || !IsAsciiLetter(s[0]))
            {
                s = "vba_tool";
            }

            return s.Length > 40 ? s.Substring(0, 40) : s;
        }

        private static bool IsAsciiLetter(char ch)
        {
            return (ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z');
        }

        public IReadOnlyList<CustomToolInfo> LoadAll()
        {
            var list = new List<CustomToolInfo>();
            if (!Directory.Exists(_root))
            {
                return list;
            }

            foreach (var file in Directory.GetFiles(_root, "*.json"))
            {
                try
                {
                    var info = JsonConvert.DeserializeObject<CustomToolInfo>(File.ReadAllText(file));
                    if (info == null || string.IsNullOrWhiteSpace(info.Name) || string.IsNullOrWhiteSpace(info.Code))
                    {
                        continue;
                    }

                    list.Add(info);
                }
                catch
                {
                    AppLog.Warn("CustomToolStore", "[LogCover] CustomToolStore.cs:201  silent-catch covered (no exception variable)");
                    // 损坏条目跳过，不影响其它工具
                }
            }

            return list.OrderBy(t => t.Name, StringComparer.OrdinalIgnoreCase).ToList();
        }

        /// <summary>按来源分类过滤列出；category 空/未知时返回全部。</summary>
        public IReadOnlyList<CustomToolInfo> LoadByCategory(string category)
        {
            var all = LoadAll();
            if (string.IsNullOrWhiteSpace(category))
            {
                return all;
            }

            return all.Where(t => string.Equals(t.EffectiveCategory(), category.Trim(), StringComparison.Ordinal)).ToList();
        }

        /// <summary>
        /// 把内置预置模板一次性导入统一存储（三源之一：preset）。幂等：仅当同名模板不存在时才写入，
        /// 绝不覆盖已有的 ai/user 沉淀或用户改动；某条失败不影响其它。
        /// </summary>
        public void SeedPresets()
        {
            foreach (var preset in VbaPresetCatalog.All)
            {
                if (Exists(preset.Name))
                {
                    continue;
                }

                try
                {
                    Save(
                        preset.Name,
                        preset.Description,
                        preset.Code,
                        ToolSourceCategory.Preset,
                        preset.Tags,
                        preset.Version);
                }
                catch
                {
                    AppLog.Warn("CustomToolStore", "[LogCover] CustomToolStore.cs:245  silent-catch covered (no exception variable)");
                    // 单个预置导入失败不阻断后续
                }
            }
        }

        public ToolResult Save(string name, string description, string code)
        {
            return Save(name, description, code, ToolSourceCategory.Ai, null, null);
        }

        public ToolResult Save(string name, string description, string code, string category, IReadOnlyList<string> tags, string version)
        {
            if (!IsValidName(name))
            {
                return Fail("工具名不合法：仅允许以字母开头，可由字母/数字/下划线组成（≤40 字符）。");
            }

            if (string.IsNullOrWhiteSpace(description))
            {
                return Fail("请先为工具写一段用途说明（description），AI 据此判断何时调用。");
            }

            if (string.IsNullOrWhiteSpace(code))
            {
                return Fail("VBA 代码不能为空。");
            }

            var categoryText = string.IsNullOrWhiteSpace(category) ? ToolSourceCategory.Ai : category.Trim();
            if (!ToolSourceCategory.IsKnown(categoryText))
            {
                categoryText = ToolSourceCategory.Ai;
            }

            var now = DateTime.Now.ToString("s");
            var info = new CustomToolInfo
            {
                Name = name.Trim(),
                Description = description.Trim(),
                Code = code,
                CreatedAt = now,
                UpdatedAt = now,
                Category = categoryText,
                Tags = tags == null
                    ? new List<string>()
                    : tags.Where(t => !string.IsNullOrWhiteSpace(t))
                        .Select(t => t.Trim()).Distinct().Take(8).ToList(),
                Version = string.IsNullOrWhiteSpace(version) ? "1.0" : version.Trim()
            };

            try
            {
                File.WriteAllText(ToFile(info.Name), JsonConvert.SerializeObject(info, Formatting.Indented));
                return Ok($"已将代码沉淀为自定义工具“{info.Name}”。下次对话可直接调用；执行前仍会请你确认。描述：{info.Description}");
            }
            catch (Exception ex)
            {
                AppLog.Warn("CustomToolStore", $"[LogCover] CustomToolStore.cs:301  {ex.GetType().Name}: {ex.Message}");
                return Fail("保存自定义工具失败：" + ex.Message);
            }
        }

        public ToolResult Delete(string name)
        {
            if (!IsValidName(name))
            {
                return Fail("工具名不合法。");
            }

            try
            {
                var file = ToFile(name);
                if (File.Exists(file))
                {
                    File.Delete(file);
                    return Ok($"已删除自定义工具“{name}”。");
                }

                return Fail($"自定义工具“{name}”不存在。");
            }
            catch (Exception ex)
            {
                AppLog.Warn("CustomToolStore", $"[LogCover] CustomToolStore.cs:325  {ex.GetType().Name}: {ex.Message}");
                return Fail("删除自定义工具失败：" + ex.Message);
            }
        }

        /// <summary>把库内工具注册进 registry（预置模板除外——它们是模板库条目，非 AI 可调用工具），
        /// 执行时经 runner(code, args) 运行存好的 VBA。</summary>
        public void RegisterTo(ToolRegistry registry, Func<string, Dictionary<string, object>, ToolResult> runner)
        {
            foreach (var tool in LoadAll())
            {
                if (string.Equals(tool.EffectiveCategory(), ToolSourceCategory.Preset, StringComparison.Ordinal))
                {
                    continue; // 预置模板不自动注册为 AI 工具，仅作为模板库供用户插入使用
                }

                var registeredName = "ct_" + tool.Name;
                if (registry.TryGet(registeredName, out _))
                {
                    continue; // 避免与既有工具/重复文件撞名
                }

                var code = tool.Code;
                var description = "自定义工具："
                    + (string.IsNullOrWhiteSpace(tool.Description) ? tool.Name : tool.Description.Trim());
                registry.Register(new ToolDescriptor(
                    registeredName,
                    description,
                    RiskLevel.Dangerous,
                    args => runner(code, args),
                    CustomSchemas.NoPackedArgs()));
            }
        }

        private static ToolResult Ok(string message)
        {
            return new ToolResult { Ok = true, Data = message };
        }

        private static ToolResult Fail(string message)
        {
            return new ToolResult { Ok = false, Data = message, ErrorCode = "CUSTOM_TOOL_FAILED" };
        }
    }

    internal static class CustomSchemas
    {
        public static Dictionary<string, object> NoPackedArgs()
        {
            // 固定代码工具：逻辑内含在沉淀代码里，无需额外入参。
            return new Dictionary<string, object>
            {
                ["type"] = "object",
                ["properties"] = new Dictionary<string, object>(),
                ["required"] = new string[0]
            };
        }
    }
}