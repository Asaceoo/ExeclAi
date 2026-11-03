using System.Collections.Generic;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>Sidebar 内联 JS 静态审计（踩坑手册 §4.1 的构建期拦截）：每次 dotnet test 都会执行，
    /// 内联脚本括号失衡/字符串未闭合会直接打红测试，防止「一个语法错误废掉整个侧边栏」上线。</summary>
    public sealed class SidebarInlineScriptTests
    {
        private static string LocateHtmlOrSkip()
        {
            return SidebarInlineScriptAuditor.LocateSidebarHtml();
        }

        [Fact]
        public void Sidebar_inline_scripts_pass_bracket_and_quote_scan()
        {
            var htmlPath = LocateHtmlOrSkip();
            Assert.True(htmlPath != null,
                "未找到 src/ExcelAiAssistant.Addin/Sidebar/index.html——仓库结构变化时请同步更新审计定位逻辑。");

            var problems = SidebarInlineScriptAuditor.Audit(System.IO.File.ReadAllText(htmlPath));
            Assert.True(problems.Count == 0,
                "内联 JS 存在括号/字符串失衡（踩坑手册 §4.1 高危）：" +
                SidebarInlineScriptAuditor.DescribeProblems(problems));
        }

        [Fact]
        public void Auditor_flags_unbalanced_brackets()
        {
            var problems = SidebarInlineScriptAuditor.Audit(
                "<script>function f() { return [1, 2;</script>");
            var problem = Assert.Single(problems);
            Assert.False(problem.Balanced);
        }

        [Fact]
        public void Auditor_flags_unterminated_string()
        {
            var problems = SidebarInlineScriptAuditor.Audit(
                "<script>var s = \"未闭合;</script>");
            Assert.Single(problems);
        }

        [Fact]
        public void Auditor_ignores_brackets_inside_strings_and_comments()
        {
            var ok = SidebarInlineScriptAuditor.Audit(
                "<script>// ( 伪括号\nvar s = \")]}\"; /* { */ var t = { a: 1 };</script>");
            Assert.Empty(ok);
        }

        [Fact]
        public void Auditor_ignores_brackets_and_quotes_inside_regex_literals()
        {
            // 真实踩坑样式：正则里含括号/字符类/引号，不得破坏状态机。
            var ok = SidebarInlineScriptAuditor.Audit(
                "<script>var re = /<w:t(?:\\s[^>]*)?>([\\s\\S]*?)<\\/w:t>/g; var q = /['\")]/g; var n = x / 2; var m = { a: /[{]/.test(s) };</script>");
            Assert.Empty(ok);
        }

        [Fact]
        public void Auditor_skips_external_script_tags()
        {
            Assert.Empty(SidebarInlineScriptAuditor.Audit(
                "<script src=\"app.js\"></script><script>var ok = 1;</script>"));
        }
    }
}
