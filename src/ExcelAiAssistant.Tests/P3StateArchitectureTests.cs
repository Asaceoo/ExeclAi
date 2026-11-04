using System;
using ExcelAiAssistant.Excel;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>P3-D2/D4 状态架构契约的纯逻辑测试（不依赖 Excel 真机）。
    /// 覆盖「每轮新建运行模块」的命名契约：创建前缀与清理谓词必须共用同一常量（防漂移→模块泄漏），
    /// 以及状态护栏：清理谓词绝不命中状态模块、状态模块源码只含最小访问器且可显式清空。</summary>
    public sealed class P3StateArchitectureTests
    {
        [Fact]
        public void Run_module_prefix_shared_between_create_and_cleanup()
        {
            // 回归锁定：模块创建（ScriptRunnerService.RunVbaSnippet）用 RunModulePrefix 拼名，
            // 清理（CleanupRunModules）用 IsRunModuleName 回收。若两者漂移，每轮运行模块会泄漏进工作簿。
            var fresh = ScriptRunnerService.RunModulePrefix + "ab12cd34";
            Assert.True(ScriptRunnerService.IsRunModuleName(fresh),
                "按 RunModulePrefix 新建的模块必须被清理谓词命中");
            Assert.StartsWith(ScriptRunnerService.RunModulePrefix, fresh, StringComparison.Ordinal);
        }

        [Fact]
        public void Legacy_scratch_module_is_still_cleaned_up()
        {
            // 旧版本常驻草稿 JYYJHelper：升级后首次调用须被清理（整模块移除，不触发工程重编译）。
            Assert.True(ScriptRunnerService.IsRunModuleName("JYYJHelper"));
        }

        [Fact]
        public void State_module_is_never_cleaned_by_run_cleanup()
        {
            // 状态护栏：JYYJXaiState 承载持久状态，清理运行模块绝不能把它当运行模块移除，
            // 否则状态（gXaiState）随模块一起被删。模块名在运行时由 StateModuleName 赋值，
            // 此处锁定「该名字既不命中清理谓词、也不落在运行前缀之下」的契约。
            Assert.False(ScriptRunnerService.IsRunModuleName("JYYJXaiState"));
            Assert.False("JYYJXaiState".StartsWith(ScriptRunnerService.RunModulePrefix, StringComparison.Ordinal));
        }

        [Fact]
        public void Run_prefix_does_not_collide_with_state_module_prefix()
        {
            // 若未来运行模块前缀与状态模块名前缀重叠，清理会误删状态模块。此处锁定互不包含。
            Assert.False(ScriptRunnerService.RunModulePrefix.StartsWith("JYYJXaiState", StringComparison.Ordinal));
            Assert.False("JYYJXaiState".StartsWith(ScriptRunnerService.RunModulePrefix, StringComparison.Ordinal));
        }

        [Fact]
        public void Arbitrary_user_module_names_are_not_cleaned()
        {
            // 用户自建模块（非 XAI 命名空间）绝不能被清理逻辑误删。
            Assert.False(ScriptRunnerService.IsRunModuleName("Module1"));
            Assert.False(ScriptRunnerService.IsRunModuleName("ThisWorkbook"));
            Assert.False(ScriptRunnerService.IsRunModuleName(""));
            Assert.False(ScriptRunnerService.IsRunModuleName(null));
        }

        [Fact]
        public void State_module_source_has_explicit_clear_accessor()
        {
            // P3-D4 显式清除：state_clear=true 时 ClearVbaState 调用 XaiStateClearRaw，模块源码必须提供该过程。
            Assert.Contains("Public Sub XaiStateClearRaw()", ScriptRunnerService.StateModuleSource);
            Assert.Contains("gXaiState = \"\"", ScriptRunnerService.StateModuleSource);
        }

        [Fact]
        public void State_module_source_has_no_result_or_args_logic()
        {
            // P3-D2 分层：状态模块只承载模块级变量 + 最小 Raw 访问器，序列化/入参逻辑留在运行模块辅助块。
            // 若状态模块被塞入 XaiReturn/XaiArg，则重建时需改写它 → 触发重编译 → 状态丢失。
            Assert.DoesNotContain("XaiReturn", ScriptRunnerService.StateModuleSource);
            Assert.DoesNotContain("XaiArg", ScriptRunnerService.StateModuleSource);
            Assert.DoesNotContain("XaiJson", ScriptRunnerService.StateModuleSource);
        }

        [Fact]
        public void Entry_resolution_picks_user_procedure_not_helper()
        {
            // 回归：P3-D2 每轮运行模块 = 声明区 + 辅助块（含 XaiProbe9F2E 等）+ 用户片段。
            // 缺省入口探测必须命中用户片段的过程（最后一个声明），绝不能误中辅助块过程——
            // 若入口被解析成 XaiProbe9F2E（需参数），探针里 Application.Run p 会抛 449「参数不可选」。
            var module = ScriptRunnerService.VbaHelperDeclarations + "\n"
                + ScriptRunnerService.BuildVbaHelperBlock("{}", 1024 * 1024) + "\n"
                + "Sub T()\n Range(\"A1\").Value = 1\nEnd Sub";
            Assert.Equal("T", ScriptRunnerService.FindEntryProcedure(module));

            // 裸体被包成 RunSnippet 时入口固定为 RunSnippet，辅助块过程也不得被选中。
            var wrapped = ScriptRunnerService.VbaHelperDeclarations + "\n"
                + ScriptRunnerService.BuildVbaHelperBlock("{}", 1024 * 1024) + "\n"
                + "Public Sub RunSnippet()\n Range(\"A1\").Value = 1\nEnd Sub";
            Assert.Equal("RunSnippet", ScriptRunnerService.FindEntryProcedure(wrapped));

            // 保护断言：入口绝不是辅助块里带必填参数的过程（449 根因）。
            Assert.NotEqual("XaiProbe9F2E", ScriptRunnerService.FindEntryProcedure(module));
            Assert.NotEqual("XaiArg", ScriptRunnerService.FindEntryProcedure(module));
        }
    }
}
