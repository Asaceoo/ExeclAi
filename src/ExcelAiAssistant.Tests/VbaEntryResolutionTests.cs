using System;
using ExcelAiAssistant.Excel;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    public sealed class VbaEntryResolutionTests
    {
        [Fact]
        public void Finds_last_declared_procedure_when_no_explicit_name()
        {
            var code = "Sub 汇总()\n    Range(\"A1\") = 1\nEnd Sub";
            Assert.Equal("汇总", ScriptRunnerService.FindEntryProcedure(code));
        }

        [Fact]
        public void Plain_statements_without_procedure_return_null()
        {
            Assert.Null(ScriptRunnerService.FindEntryProcedure("Range(\"A1\") = 1"));
        }

        [Fact]
        public void Multiple_procedures_picks_the_last_one()
        {
            var code = "Sub First()\nEnd Sub\nPrivate Function Second() As Long\nEnd Function";
            Assert.Equal("Second", ScriptRunnerService.FindEntryProcedure(code));
        }

        [Fact]
        public void Wrapped_snippet_resolves_to_runsnippet()
        {
            var code = "Public Sub RunSnippet()\nRange(\"A1\") = 2\nEnd Sub";
            Assert.Equal("RunSnippet", ScriptRunnerService.FindEntryProcedure(code));
        }

        // 缺陷2回归：中文过程名在中文版 VBA 完全合法，DetectVbaCompileIssue 不得再把它误报为“中文标识符无法编译”。
        // 合法片段（含隐式变量、未声明变量，无任何编译问题）必须返回 null（不拦截、不误导）。
        [Fact]
        public void Compile_issue_detector_does_not_blame_valid_chinese_procedure_names()
        {
            var code = "Sub 题目1_多条件汇总()\n"
                       + "    s = 0\n"
                       + "    For i = 1 To 100\n"
                       + "        s = s + i\n"
                       + "    Next\n"
                       + "End Sub";
            Assert.Null(ScriptRunnerService.DetectVbaCompileIssue(code));
        }

        [Fact]
        public void Compile_issue_detector_does_not_blame_chinese_variable_references()
        {
            // 中文变量名同样合法；直接引用未声明的中文变量（Option Explicit 之下才报错）不得被判为编译问题。
            var code = "Sub 汇总()\n    Dim 总量 As Long\n    总量 = 总量 + 1\nEnd Sub";
            Assert.Null(ScriptRunnerService.DetectVbaCompileIssue(code));
        }

        // 上两条在“无 Option Explicit”的宽松编译语义下成立：中文名 + 隐式变量都不该触发误报。
        [Fact]
        public void Compile_issue_detector_still_catches_unclosed_quotes()
        {
            var code = "Sub 汇总()\n    Range(\"A1\") = \"未闭合\nEnd Sub";
            Assert.Contains("未闭合的双引号", ScriptRunnerService.DetectVbaCompileIssue(code));
        }

        [Fact]
        public void Compile_issue_detector_still_catches_missing_end_sub()
        {
            var code = "Sub 汇总()\n    Range(\"A1\") = 1";
            Assert.Contains("过程声明比过程结尾多", ScriptRunnerService.DetectVbaCompileIssue(code));
        }

        [Fact]
        public void Compile_issue_detector_null_for_blank()
        {
            Assert.Null(ScriptRunnerService.DetectVbaCompileIssue(null));
            Assert.Null(ScriptRunnerService.DetectVbaCompileIssue("   "));
        }

        [Fact]
        public void Find_entry_procedure_resolves_chinese_name_with_digits_and_underscore()
        {
            var code = "Sub 题目1_多条件汇总()\n    Range(\"A1\") = 1\nEnd Sub";
            Assert.Equal("题目1_多条件汇总", ScriptRunnerService.FindEntryProcedure(code));
        }

        [Fact]
        public void English_macro_disabled_error_classifies_specifically()
        {
            var ex = new Exception("Cannot run the macro 'JYYJHelper.AutoFitRows'. The macro may not be available in this workbook or all macros may be disabled.");
            Assert.Equal("VBA_MACROS_DISABLED", ScriptRunnerService.ClassifyVbaRunError(ex));
        }

        [Fact]
        public void Chinese_macro_disabled_error_classifies_specifically()
        {
            var ex = new Exception("宏“JYYJHelper.AutoFitRowHeight”在此工作簿中不可用或所有宏被禁用");
            Assert.Equal("VBA_MACROS_DISABLED", ScriptRunnerService.ClassifyVbaRunError(ex));
        }

        [Fact]
        public void Generic_runtime_error_is_not_macro_disabled()
        {
            Assert.Equal("VBA_RUN_ERROR", ScriptRunnerService.ClassifyVbaRunError(new Exception("VBA 运行时错误 438")));
            Assert.Equal("VBA_RUN_ERROR", ScriptRunnerService.ClassifyVbaRunError(new Exception("The remote procedure call failed")));
        }

        [Fact]
        public void Null_exception_classifies_as_generic()
        {
            Assert.Equal("VBA_RUN_ERROR", ScriptRunnerService.ClassifyVbaRunError(null));
        }

        [Fact]
        public void Probe_ok_is_success()
        {
            var (ok, num, desc, resultJson, truncated) = ScriptRunnerService.ParseVbaProbeResult("OK");
            Assert.True(ok);
            Assert.Equal(0, num);
            Assert.Equal("", desc);
            Assert.Null(resultJson);
            Assert.False(truncated);
        }

        [Fact]
        public void Probe_null_is_success()
        {
            var (ok, _, _, _, _) = ScriptRunnerService.ParseVbaProbeResult(null);
            Assert.True(ok);
        }

        [Fact]
        public void Probe_ok_with_json_returns_structured_result()
        {
            var (ok, _, _, resultJson, truncated) = ScriptRunnerService.ParseVbaProbeResult("OK|[1,2,3]");
            Assert.True(ok);
            Assert.Equal("[1,2,3]", resultJson);
            Assert.False(truncated);
        }

        [Fact]
        public void Probe_ok_with_truncated_json_marks_truncated()
        {
            var (ok, _, _, resultJson, truncated) = ScriptRunnerService.ParseVbaProbeResult("OKT|{\"a\":1");
            Assert.True(ok);
            Assert.Equal("{\"a\":1", resultJson);
            Assert.True(truncated);
        }

        [Fact]
        public void Probe_err_parses_number_and_description()
        {
            var (ok, num, desc, _, _) = ScriptRunnerService.ParseVbaProbeResult("ERR|1004|Method 'AutoFit' of object 'Range' failed");
            Assert.False(ok);
            Assert.Equal(1004, num);
            Assert.Equal("Method 'AutoFit' of object 'Range' failed", desc);
        }

        [Fact]
        public void Probe_unexpected_value_is_unknown_error()
        {
            var (ok, _, desc, _, _) = ScriptRunnerService.ParseVbaProbeResult("weird");
            Assert.False(ok);
            Assert.Equal("weird", desc);
        }

        [Theory]
        [InlineData(-2146788248, true)] // 0x800A9C68：第一轮“宏尚未就绪”瞬态错误 → 触发重试自愈
        [InlineData(-2146827284, false)] // 0x800A03EC：1004 普通运行时错误，不应触发重试
        public void Probe_likely_unavailable_heuristic(int hresult, bool expected)
        {
            Assert.Equal(expected, ScriptRunnerService.ProbeLikelyUnavailable(new System.Runtime.InteropServices.COMException("oops", hresult)));
        }

        [Theory]
        [InlineData("Cannot run the macro 'X'. The macro may not be available in this workbook or all macros may be disabled.", true)]
        [InlineData("VBA 运行时错误 1004：Method 'AutoFit' of object 'Range' failed", false)]
        public void Probe_likely_unavailable_message(string message, bool expected)
        {
            Assert.Equal(expected, ScriptRunnerService.ProbeLikelyUnavailable(new Exception(message)));
        }

        [Theory]
        [InlineData("Exception calling \"Run\" with \"2\" argument(s): \"无法运行“JYYJHelper.XaiProbe9F2E”宏。可能是因为该宏在此工作簿中不可用，或者所有的宏都被禁用。\"", true)]
        [InlineData("宏“X”在此工作簿中不可用", true)]
        [InlineData("无法运行“X”宏", true)]
        [InlineData("VBA 运行时错误 438：对象不支持该属性或方法", false)]
        public void Probe_likely_unavailable_chinese_transient(string message, bool expected)
        {
            // 回归：中文版 Excel 的瞬态「宏尚未就绪」消息（模块刚写入未编译时第一轮 Run 必现）
            // 必须触发重试自愈，否则误判 VBA_MACROS_DISABLED。
            Assert.Equal(expected, ScriptRunnerService.ProbeLikelyUnavailable(new Exception(message)));
        }

        [Theory]
        [InlineData("ws.Range(\"G12:K18\").AutoFit", true)]
        [InlineData("ws.Range(\"A1:B10\").AutoFit", true)]
        [InlineData("ws.Range(\"A1\").AutoFit", true)] // 带接收者的单格也会被匹配（保守提示，无副作用）
        [InlineData("ws.Range(\"G12\").Offset(2,3).AutoFit", true)]
        [InlineData("ws.Rows(\"12:18\").AutoFit", false)]
        [InlineData("ws.Columns(\"G:K\").AutoFit", false)]
        [InlineData("ws.Rows(\"12:18\").RowHeight = 0", false)]
        [InlineData("ws.Range(\"G12:K18\").EntireRow.AutoFit", true)]
        [InlineData("rows.AutoFit", false)]
        [InlineData("columns.AutoFit", false)]
        public void AutoFit2d_detection_catches_invalid_patterns(string code, bool expected)
        {
            var matches = ScriptRunnerService.RegexVbaAutoFit2D.IsMatch(code);
            Assert.Equal(expected, matches);
        }
    }
}