using System;
using System.Collections.Generic;
using System.Text;
using ExcelAiAssistant.Excel;
using Newtonsoft.Json.Linq;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>P1 最小闭环：VBA/Python 结构化回传与数据通道的纯逻辑测试（不依赖 Excel/Python 真机）。
    /// 覆盖 args 归一化、input_data 护栏、VBA 字符串字面量转义、辅助块生成、探针解析。</summary>
    public sealed class P1DataChannelTests
    {
        [Fact]
        public void Normalize_args_null_returns_empty_object()
        {
            Assert.Equal("{}", ScriptRunnerService.NormalizeArgsJson(null));
        }

        [Fact]
        public void Normalize_args_object_minifies_json()
        {
            var raw = new Dictionary<string, object> { ["column"] = "A", ["threshold"] = 100 };
            Assert.Equal("{\"column\":\"A\",\"threshold\":100}", ScriptRunnerService.NormalizeArgsJson(raw));
        }

        [Fact]
        public void Normalize_args_json_string_passes_through_minified()
        {
            Assert.Equal("{\"a\":1}", ScriptRunnerService.NormalizeArgsJson(" { \"a\" : 1 } "));
        }

        [Fact]
        public void Normalize_args_invalid_string_falls_back_to_string_value()
        {
            Assert.Equal("\"not json\"", ScriptRunnerService.NormalizeArgsJson("not json"));
        }

        [Fact]
        public void Normalize_args_jtoken_input_returns_compact_json()
        {
            // 真实 AI 调用时 args 可能是 JObject/JArray（已反序列化的 JToken），须原样紧凑化而非再包一层。
            var token = JToken.Parse("{\"column\":\"A\",\"nested\":{\"x\":[1,2]}}");
            Assert.Equal("{\"column\":\"A\",\"nested\":{\"x\":[1,2]}}", ScriptRunnerService.NormalizeArgsJson(token));

            var arr = JToken.Parse("[1,2,3]");
            Assert.Equal("[1,2,3]", ScriptRunnerService.NormalizeArgsJson(arr));
        }

        [Fact]
        public void Normalize_args_nested_dictionary_supported()
        {
            var raw = new Dictionary<string, object>
            {
                ["k"] = "v",
                ["nested"] = new Dictionary<string, object> { ["a"] = new object[] { 1, 2 } }
            };
            Assert.Equal("{\"k\":\"v\",\"nested\":{\"a\":[1,2]}}", ScriptRunnerService.NormalizeArgsJson(raw));
        }

        [Fact]
        public void Count_cells_counts_scalars_recursively()
        {
            var token = JToken.Parse("[[1,2],[3,4]]");
            Assert.Equal(4, ScriptRunnerService.CountCells(token));

            var obj = JToken.Parse("{\"a\":1,\"b\":[1,2,3]}");
            Assert.Equal(4, ScriptRunnerService.CountCells(obj));
        }

        [Fact]
        public void Normalize_input_data_absent_returns_null()
        {
            var (token, error) = ScriptRunnerService.NormalizeInputData(new Dictionary<string, object>());
            Assert.Null(token);
            Assert.Null(error);
        }

        [Fact]
        public void Normalize_input_data_valid_returns_token()
        {
            var args = new Dictionary<string, object> { ["input_data"] = "{\"x\":1}" };
            var (token, error) = ScriptRunnerService.NormalizeInputData(args);
            Assert.Null(error);
            Assert.NotNull(token);
            Assert.Equal(1, (int)token["x"]);
        }

        [Fact]
        public void Normalize_input_data_too_many_cells_rejected()
        {
            // 400x300 = 12 万 cell > 10 万上限 → INPUT_TOO_LARGE。
            var rows = new List<object>();
            for (var i = 0; i < 400; i++)
            {
                var row = new List<object>();
                for (var j = 0; j < 300; j++)
                {
                    row.Add(j);
                }

                rows.Add(row);
            }

            var args = new Dictionary<string, object> { ["input_data"] = rows };
            var (_, error) = ScriptRunnerService.NormalizeInputData(args);
            Assert.NotNull(error);
            Assert.Contains("过大", error);
        }

        [Fact]
        public void Vba_string_literal_escapes_quotes()
        {
            Assert.Equal("\"a\"\"b\"", ScriptRunnerService.VbaStringLiteral("a\"b"));
        }

        [Fact]
        public void Vba_string_literal_chunks_long_text()
        {
            var text = new string('x', 500);
            var literal = ScriptRunnerService.VbaStringLiteral(text);
            Assert.Contains("& _", literal);
            // 去掉续行符与引号后应还原原文。
            var restored = literal.Replace("\"", "").Replace(" & _", "").Replace("\n", "").Replace("    ", "");
            Assert.Equal(text, restored);
        }

        [Fact]
        public void Vba_string_literal_never_splits_escape_pair_at_chunk_boundary()
        {
            // 回归：旧实现按转义后文本 100 字符硬切，边界落在 "" 转义对中间时块尾单引号与包裹引号合并成转义对，
            // 字面量未闭合 → 模块编译失败 → 工具挂死。新实现按原始字符分块、逐块独立转义，杜绝切断转义对。
            var text = new string('x', 99) + "\"" + new string('y', 60);
            var literal = ScriptRunnerService.VbaStringLiteral(text);
            Assert.Equal(text, DecodeVbaLiteral(literal));
        }

        [Fact]
        public void Helper_json_parser_decodes_escaped_quote_without_backslash()
        {
            // 回归：XaiParseStr 的 Case """" 曾追加 "\"""（反斜杠+引号）而非单个引号，
            // 使含引号入参往返后多出反斜杠（自检 CheckVbaDataChannelLargeString 端到端覆盖，此处守住源码契约）。
            // 注意 helper 里有两条 Case """"：XaiEsc 编码侧（输出 \"，正确保留）与 XaiParseStr 解码侧（还原为单个引号 """"）。
            var helper = ScriptRunnerService.BuildVbaHelperBlock("{}", 1024);
            // 解码侧：JSON \" 应还原为单个引号（VBA 字面量 """"），而非 "\"""（反斜杠+引号）。
            Assert.Contains("Case \"\"\"\": sb = sb & \"\"\"\"", helper);
            // 编码侧：输出 JSON 时引号仍须转义为 \"，防止误伤编码路径。
            Assert.Contains("Case \"\"\"\": sb = sb & \"\\\"\"\"", helper);
        }

        private static string DecodeVbaLiteral(string literal)
        {
            var sep = "\" & _\n    \"";
            var sb = new StringBuilder();
            var pos = 0;
            var first = true;
            while (true)
            {
                var idx = literal.IndexOf(sep, pos, StringComparison.Ordinal);
                var end = idx < 0 ? literal.Length - 1 : idx;
                var start = first ? pos + 1 : pos;
                sb.Append(literal.Substring(start, end - start).Replace("\"\"", "\""));
                if (idx < 0)
                {
                    break;
                }

                pos = idx + sep.Length;
                first = false;
            }

            return sb.ToString();
        }

        [Fact]
        public void Vba_helper_block_contains_probe_and_args()
        {
            var block = ScriptRunnerService.BuildVbaHelperBlock("{\"k\":\"v\"}", 1024);
            Assert.Contains("XaiProbe9F2E", block);
            Assert.Contains("XaiReturn", block);
            Assert.Contains("XaiArg", block);
            Assert.Contains("XaiInitArgs", block);
            Assert.Contains("gXaiArgs = \"{\"\"k\"\":\"\"v\"\"}\"", block);
            Assert.Contains("XAI_RESULT_LIMIT = 1024", block);
        }

        [Fact]
        public void Vba_helper_block_escapes_quotes_in_args()
        {
            var block = ScriptRunnerService.BuildVbaHelperBlock("{\"a\":\"x\\\"y\"}", 1024);
            Assert.Contains("gXaiArgs = \"{\"\"a\"\":\"\"x\\\"\"y\"\"}\"", block);
        }

        [Fact]
        public void Vba_helper_block_parse_num_handles_large_ints_without_overflow()
        {
            // F1 修复：大整数（超出 Long 范围）回退 Double，避免 CLng 溢出使整个数据通道失败。
            var block = ScriptRunnerService.BuildVbaHelperBlock("{}", 1024);
            Assert.Contains("d = Val(txt)", block);
            Assert.Contains("If d > 2147483647# Or d < -2147483648# Then", block);
            Assert.Contains("XaiParseNum = d", block);
        }

        [Fact]
        public void Vba_helper_block_arg_iteration_is_bounds_safe()
        {
            // 回归：XaiArg 曾按 UBound 满幅 Step 2 遍历，奇数长/纯数组 args 会访问 obj(i+1) 越界抛「下标越界」。
            var block = ScriptRunnerService.BuildVbaHelperBlock("{}", 1024);
            Assert.Contains("To UBound(obj, 1) - 1 Step 2", block);
        }

        [Fact]
        public void Truncated_marker_is_consistent_between_worker_and_check()
        {
            // 回归：C# 判定曾用单字符省略号 U+2026，而 worker/Truncate 用 ASCII 三点，
            // 两者不等 → run_python_code 的 truncated 字段恒为 false，AI 感知不到输出被截断。
            Assert.Equal("...（已截断）", ScriptRunnerService.OutputTruncatedMarker);
            Assert.Contains(ScriptRunnerService.OutputTruncatedMarker, PythonWorkerHost.WorkerMain);
        }

        [Fact]
        public void Vba_helper_block_num_is_locale_safe()
        {
            // F2 修复：Str() 的小数点随区域设置，统一 Replace 为 '.' 保证输出合法 JSON。
            var block = ScriptRunnerService.BuildVbaHelperBlock("{}", 1024);
            Assert.Contains("Replace(Trim(Str(v)), \",\", \".\")", block);
        }

        [Fact]
        public void Vba_state_module_keeps_module_level_state_only()
        {
            // P3-D2：状态模块只承载模块级变量 + 三个最小访问器，编解码留在草稿模块辅助块。
            Assert.Contains("Private gXaiState As String", ScriptRunnerService.StateModuleSource);
            Assert.Contains("Public Function XaiStateRaw", ScriptRunnerService.StateModuleSource);
            Assert.Contains("Public Sub XaiStateRawSet", ScriptRunnerService.StateModuleSource);
            Assert.Contains("Public Sub XaiStateClearRaw", ScriptRunnerService.StateModuleSource);
        }

        [Fact]
        public void Vba_helper_declarations_are_static_and_detected()
        {
            // P3-D2：辅助块声明区须恒为 5 行静态文本（无每轮占位符），且被 IsStaticDeclarationSection 精确识别；
            // 旧版本 const 形式必须判定为「非静态」→ 触发整清重建，防止把旧 const 当静态保留导致「给常量赋值」编译错误。
            Assert.DoesNotContain("{XAI_RESULT_LIMIT}", ScriptRunnerService.VbaHelperDeclarations);
            Assert.Contains("Private XAI_RESULT_LIMIT As Long", ScriptRunnerService.VbaHelperDeclarations);
            Assert.True(ScriptRunnerService.IsStaticDeclarationSection(ScriptRunnerService.VbaHelperDeclarations));
            Assert.True(ScriptRunnerService.IsStaticDeclarationSection(
                "\r\n" + ScriptRunnerService.VbaHelperDeclarations + "\r\n")); // 首尾空白/空行容错
            Assert.False(ScriptRunnerService.IsStaticDeclarationSection(
                "Private Const XAI_RESULT_LIMIT As Long = 102400")); // 旧版本 const 形式
            Assert.False(ScriptRunnerService.IsStaticDeclarationSection(
                ScriptRunnerService.VbaHelperDeclarations + "\nPrivate gExtra As Long")); // 片段残留模块级声明
        }

        [Fact]
        public void Vba_helper_block_contains_xai_state_helpers()
        {
            // P3-D2：草稿模块辅助块提供 XaiStateGet/Set/Clear，经状态模块的 Raw 访问器读写。
            var block = ScriptRunnerService.BuildVbaHelperBlock("{}", 1024);
            Assert.Contains("Public Function XaiStateGet(key As String)", block);
            Assert.Contains("Public Sub XaiStateSet(key As String, v As Variant)", block);
            Assert.Contains("Public Sub XaiStateClear()", block);
            Assert.Contains("XaiStateRaw", block);
            Assert.Contains("XaiStateRawSet", block);
            Assert.Contains("XaiStateClearRaw", block);
        }

        [Fact]
        public void Vba_state_set_uses_dictionary_roundtrip()
        {
            // P3-D2 契约：XaiStateSet 用 Scripting.Dictionary 重组键值并交给 XaiJson 序列化，XaiStateGet 用 XaiParse 还原。
            var block = ScriptRunnerService.BuildVbaHelperBlock("{}", 1024);
            Assert.Contains("CreateObject(\"Scripting.Dictionary\")", block);
            Assert.Contains("obj = XaiParse(XaiStateRaw)", block);
            Assert.Contains("XaiStateRawSet XaiJson(d, 0)", block);
        }
    }
}
