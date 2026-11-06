using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    public sealed class CustomToolStoreTests : IDisposable
    {
        private readonly string root;
        private readonly CustomToolStore store;

        public CustomToolStoreTests()
        {
            root = Path.Combine(Path.GetTempPath(), $"xai-custom-tools-{Guid.NewGuid():N}");
            store = new CustomToolStore(root);
        }

        [Fact]
        public void Preset_catalog_names_unique_and_fields_nonempty()
        {
            var all = VbaPresetCatalog.All;
            Assert.NotEmpty(all);
            // 名称合法且唯一（seed 时按名去重，重名会导致静默丢模板）
            Assert.Equal(all.Count, all.Select(t => t.Name).Distinct().Count());
            foreach (var template in all)
            {
                Assert.True(CustomToolStore.IsValidName(template.Name), "预置模板名不合法: " + template.Name);
                Assert.False(string.IsNullOrWhiteSpace(template.Description), "描述为空: " + template.Name);
                Assert.False(string.IsNullOrWhiteSpace(template.Code), "代码为空: " + template.Name);
                Assert.NotNull(template.Tags);
                Assert.NotEmpty(template.Tags);
            }
        }

        [Fact]
        public void Python_preset_templates_carry_python_tag()
        {
            // Python 面板模板库按「python」标签过滤（三源一套），预置源必须带该标签才能被展示；
            // 名称前缀 py_ 也须与 VBA 预置（sum_column 等）隔离，避免互相污染。
            var pythonPresets = VbaPresetCatalog.All
                .Where(t => t.Name.StartsWith("py_", StringComparison.Ordinal))
                .ToList();
            Assert.NotEmpty(pythonPresets);
            foreach (var template in pythonPresets)
            {
                Assert.Contains("python", template.Tags, StringComparer.OrdinalIgnoreCase);
            }
        }

        [Fact]
        public void Save_returns_false_for_invalid_or_blank_name()
        {
            Assert.False(store.Save("1bad", "描述", "Sub x(): End Sub").Ok);
            Assert.False(store.Save("", "描述", "code").Ok);
            Assert.False(store.Save("  ", "描述", "code").Ok);
            Assert.False(store.Save("ok_name", "", "code").Ok, "空描述应被拒绝");
        }

        [Fact]
        public void Save_and_reload_round_trips()
        {
            var ok = store.Save("my_tool", "我的工具", "Sub x(): MsgBox 1: End Sub");
            Assert.True(ok.Ok);

            var tools = store.LoadAll();
            Assert.Single(tools);
            Assert.Equal("my_tool", tools[0].Name);
            Assert.True(store.Exists("my_tool"));
        }

        [Fact]
        public void Find_available_name_dedupes_existing_names_with_suffix()
        {
            Assert.True(store.Save("mytool", "描述A", "Sub a(): End Sub").Ok);

            var first = store.FindAvailableName("mytool");
            Assert.NotEqual("mytool", first); // 已占用 → 应改名
            Assert.False(store.Exists(first));

            // 先把候选名也占用，验证继续递增到 _3
            Assert.True(store.Save(first, "描述B", "Sub b(): End Sub").Ok);
            var third = store.FindAvailableName("mytool");
            Assert.Equal("mytool_3", third);
            Assert.False(store.Exists(third));
        }

        [Fact]
        public void Save_with_an_available_name_does_not_fail()
        {
            Assert.True(store.Save("first", "描述", "Sub a(): End Sub").Ok);
            Assert.True(store.Save("second", "描述", "Sub b(): End Sub").Ok);
            Assert.Equal(2, store.LoadAll().Count);
        }

        [Fact]
        public void Find_available_name_sanitizes_chinese_macro_names()
        {
            // Sub 汇总() → AutoToolName 产出 proc_汇总：必须落成合法 ASCII 名，否则沉淀失败
            var mixed = store.FindAvailableName("proc_汇总");
            Assert.True(CustomToolStore.IsValidName(mixed), $"混合名应被合法化：{mixed}");

            var pureChinese = store.FindAvailableName("汇总工具");
            Assert.True(CustomToolStore.IsValidName(pureChinese), $"纯中文名应被合法化：{pureChinese}");
            Assert.True(store.Save(pureChinese, "中文宏", "Sub 汇总(): End Sub").Ok,
                "中文宏名经合法化后应能成功沉淀");
            Assert.Single(store.LoadAll());
        }

        [Fact]
        public void Legacy_save_defaults_to_ai_category()
        {
            Assert.True(store.Save("legacy_tool", "描述", "Sub x(): End Sub").Ok);
            var tool = Assert.Single(store.LoadAll());
            Assert.Equal(ToolSourceCategory.Ai, tool.EffectiveCategory());
        }

        [Fact]
        public void Save_with_category_tags_version_round_trips()
        {
            Assert.True(store.Save(
                "user_tpl", "用户模板", "Sub 汇总(): End Sub",
                ToolSourceCategory.User,
                new[] { "汇总", "求和" },
                "2.1").Ok);

            var tool = Assert.Single(store.LoadAll());
            Assert.Equal(ToolSourceCategory.User, tool.Category);
            Assert.Equal(new[] { "汇总", "求和" }, tool.Tags);
            Assert.Equal("2.1", tool.Version);
        }

        [Fact]
        public void Unknown_category_falls_back_to_ai()
        {
            Assert.True(store.Save(
                "odd", "描述", "Sub x(): End Sub",
                "banana", new[] { "a" }, null).Ok);
            Assert.Equal(ToolSourceCategory.Ai, Assert.Single(store.LoadAll()).EffectiveCategory());
        }

        [Fact]
        public void LoadByCategory_filters_by_source()
        {
            store.Save("ai_tool", "AI沉淀", "Sub a(): End Sub");                       //默认 ai
            store.Save("pr_tool", "预置", "Sub p(): End Sub", ToolSourceCategory.Preset, null, "1.0");
            store.Save("us_tool", "手写", "Sub u(): End Sub", ToolSourceCategory.User, null, null);

            Assert.Equal(new[] { "ai_tool" },
                store.LoadByCategory(ToolSourceCategory.Ai).Select(t => t.Name));
            Assert.Equal(new[] { "pr_tool" },
                store.LoadByCategory(ToolSourceCategory.Preset).Select(t => t.Name));
            Assert.Equal(new[] { "us_tool" },
                store.LoadByCategory(ToolSourceCategory.User).Select(t => t.Name));
            Assert.Equal(3, store.LoadByCategory(null).Count);
        }

        [Fact]
        public void SeedPresets_seeds_all_and_is_idempotent_and_never_overwrites()
        {
            // 先手工放一个与预置同名的 ai 沉淀，seed 绝不能覆盖它
            Assert.True(store.Save("sum_column", "我自己的汇总", "Sub mine(): End Sub").Ok);

            store.SeedPresets();
            store.SeedPresets(); // 幂等：第二次不重复、不叠加

            var all = store.LoadAll();
            Assert.Equal(VbaPresetCatalog.All.Count, all.Count); // 预置同名(sum_column)被已有 ai 沉淀占用而跳过

            var preset = all.Single(t => t.Name == "sum_column");
            Assert.Equal(ToolSourceCategory.Ai, preset.EffectiveCategory()); // 仍是 AI 沉淀
            Assert.Equal("Sub mine(): End Sub", preset.Code);                 // 代码未被覆盖

            var other = all.Single(t => t.Name == "auto_width");
            Assert.Equal(ToolSourceCategory.Preset, other.EffectiveCategory()); // 自动补入的预置
            Assert.False(string.IsNullOrWhiteSpace(other.Code));
        }

        [Fact]
        public void RegisterTo_skips_preset_templates()
        {
            store.Save("sum_column", "预置汇总", "Sub s(): End Sub", ToolSourceCategory.Preset, null, "1.0");
            store.Save("ai_tool", "AI沉淀", "Sub a(): End Sub");   //默认 ai
            store.Save("user_tpl", "手写", "Sub u(): End Sub", ToolSourceCategory.User, null, null);

            var registry = new ToolRegistry();
            store.RegisterTo(registry, (code, args) => new ToolResult { Ok = true, Data = "运行" });

            var names = registry.Tools.Select(t => t.Name).ToList();
            Assert.Contains("ct_ai_tool", names);
            Assert.Contains("ct_user_tpl", names);
            Assert.DoesNotContain("ct_sum_column", names); // 预置模板不自动注册为 AI 工具
        }

        public void Dispose()
        {
            if (Directory.Exists(root))
            {
                Directory.Delete(root, true);
            }
        }
    }
}