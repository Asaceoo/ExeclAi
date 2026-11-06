using System;
using System.Collections.Generic;
using ExcelAiAssistant.Excel;
using Newtonsoft.Json.Linq;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>run_python_excel 的纯逻辑单测（不依赖 Excel/Python 真机）：
    /// 覆盖 write_ranges 解析、版本/位数判定、连接错误分类、前置连接代码、参数引用转义。</summary>
    public sealed class PythonExcelServiceTests
    {
        // ---------- SplitRangeAddress：地址拆分 ----------

        [Fact]
        public void Split_address_with_sheet_prefix()
        {
            var list = PythonExcelService.SplitRangeAddressList(new[] { "Sheet1!A1:B5" });
            var item = (dynamic)list[0];
            Assert.Equal("Sheet1", (string)item.sheet);
            Assert.Equal("A1:B5", (string)item.address);
        }

        [Fact]
        public void Split_address_without_sheet_defaults_null()
        {
            var list = PythonExcelService.SplitRangeAddressList(new[] { "C5:D9" });
            var item = (dynamic)list[0];
            Assert.Null((string)item.sheet);
            Assert.Equal("C5:D9", (string)item.address);
        }

        [Fact]
        public void Split_address_quoted_sheet_name_strips_quotes()
        {
            var list = PythonExcelService.SplitRangeAddressList(new[] { "'My Sheet'!A1" });
            var item = (dynamic)list[0];
            Assert.Equal("My Sheet", (string)item.sheet);
            Assert.Equal("A1", (string)item.address);
        }

        [Fact]
        public void Split_address_invalid_or_empty_yields_null()
        {
            var list = PythonExcelService.SplitRangeAddressList(new[] { "", "   ", "!A1", "Sheet1!" });
            foreach (var raw in list)
            {
                var item = (dynamic)raw;
                Assert.Null((string)item.sheet);
                Assert.Null((string)item.address);
            }
        }

        // ---------- ParseWriteRanges：write_ranges 全量解析 ----------

        [Fact]
        public void Parse_write_ranges_absent_ok_empty()
        {
            var (ok, error, ranges) = PythonExcelService.ParseWriteRangesForTest(new Dictionary<string, object>());
            Assert.True(ok);
            Assert.Null(error);
            Assert.Empty(ranges);
        }

        [Fact]
        public void Parse_write_ranges_string_entries()
        {
            var args = new Dictionary<string, object> { ["write_ranges"] = new object[] { "Sheet1!A1:B5", "C1" } };
            var (ok, _, ranges) = PythonExcelService.ParseWriteRangesForTest(args);
            Assert.True(ok);
            Assert.Equal(2, ranges.Count);
            var r0 = (dynamic)ranges[0];
            Assert.Equal("Sheet1", (string)r0.sheet);
            Assert.Equal("A1:B5", (string)r0.address);
            var r1 = (dynamic)ranges[1];
            Assert.Null((string)r1.sheet);
            Assert.Equal("C1", (string)r1.address);
        }

        [Fact]
        public void Parse_write_ranges_jobject_entries()
        {
            var args = new Dictionary<string, object>
            {
                ["write_ranges"] = new object[]
                {
                    JObject.FromObject(new { sheet = "data", address = "A1:C3" }),
                    JObject.Parse("{\"address\":\"F1\"}")
                }
            };
            var (ok, _, ranges) = PythonExcelService.ParseWriteRangesForTest(args);
            Assert.True(ok);
            Assert.Equal(2, ranges.Count);
            var r0 = (dynamic)ranges[0];
            Assert.Equal("data", (string)r0.sheet);
            Assert.Equal("A1:C3", (string)r0.address);
            var r1 = (dynamic)ranges[1];
            Assert.Null((string)r1.sheet);
            Assert.Equal("F1", (string)r1.address);
        }

        [Fact]
        public void Parse_write_ranges_null_items_skipped()
        {
            var args = new Dictionary<string, object> { ["write_ranges"] = new object[] { null, "A1", null } };
            var (ok, _, ranges) = PythonExcelService.ParseWriteRangesForTest(args);
            Assert.True(ok);
            Assert.Single(ranges);
        }

        [Fact]
        public void Parse_write_ranges_object_missing_address_rejected()
        {
            var args = new Dictionary<string, object> { ["write_ranges"] = new object[] { JObject.Parse("{\"sheet\":\"s\"}") } };
            var (ok, error, _) = PythonExcelService.ParseWriteRangesForTest(args);
            Assert.False(ok);
            Assert.Contains("address", error);
        }

        // ---------- ParseWriteRanges：JValue 字符串归一化（真机实证 Sheet1!A52:C60 曾被误报格式无效） ----------

        [Fact]
        public void Parse_write_ranges_jvalue_string_entries_accepted()
        {
            var args = new Dictionary<string, object>
            {
                ["write_ranges"] = new object[] { new JValue("Sheet1!A52:C60"), new JValue("A30:C30") }
            };
            var (ok, error, ranges) = PythonExcelService.ParseWriteRangesForTest(args);
            Assert.True(ok);
            Assert.Null(error);
            Assert.Equal(2, ranges.Count);
            var r0 = (dynamic)ranges[0];
            Assert.Equal("Sheet1", (string)r0.sheet);
            Assert.Equal("A52:C60", (string)r0.address);
            var r1 = (dynamic)ranges[1];
            Assert.Null((string)r1.sheet);
            Assert.Equal("A30:C30", (string)r1.address);
        }

        [Fact]
        public void Parse_write_ranges_jarray_of_jvalues_accepted()
        {
            var arr = JArray.Parse("[\"Sheet1!A52:C60\", \"A30:C30\"]");
            var args = new Dictionary<string, object> { ["write_ranges"] = arr };
            var (ok, error, ranges) = PythonExcelService.ParseWriteRangesForTest(args);
            Assert.True(ok);
            Assert.Null(error);
            Assert.Equal(2, ranges.Count);
        }

        [Fact]
        public void Parse_write_ranges_jvalue_string_has_no_bang_still_ok()
        {
            var args = new Dictionary<string, object> { ["write_ranges"] = new object[] { new JValue("$A$52:$C$60") } };
            var (ok, error, ranges) = PythonExcelService.ParseWriteRangesForTest(args);
            Assert.True(ok);
            Assert.Null(error);
            var r0 = (dynamic)ranges[0];
            Assert.Null((string)r0.sheet);
            Assert.Equal("$A$52:$C$60", (string)r0.address);
        }

        // ---------- VersionAtLeast39：版本判定 ----------

        [Theory]
        [InlineData("3.9", true)]
        [InlineData("3.9.7", true)]
        [InlineData("3.10", true)]
        [InlineData("3.13.1", true)]
        [InlineData("4.0", true)]
        [InlineData("3.8.10", false)]
        [InlineData("2.7", false)]
        [InlineData("", false)]
        [InlineData("abc", false)]
        [InlineData("3", false)]
        public void Version_at_least_39(string version, bool expected)
        {
            Assert.Equal(expected, PythonExcelService.VersionAtLeast39(version));
        }

        // ---------- ClassifyConnectionError：连接错误分类 ----------

        [Theory]
        [InlineData("Traceback ... Excel App not found", "无法连接当前 Excel 实例")]
        [InlineData("Could not connect to Excel", "无法连接当前 Excel 实例")]
        [InlineData("xw.apps.active failed: Excel is not running", "无法连接当前 Excel 实例")]
        [InlineData("No active books", "未能定位活动工作簿")]
        [InlineData("AttributeError: 'Books' object has no attribute 'active'", "未能定位活动工作簿")]
        [InlineData("xlwings.XlwingsError: Make sure to have \"pywin32\", a dependency of xlwings, installed.", "xlwings 的 Excel 引擎未加载")]
        public void Classify_connection_error_gives_hint(string output, string hint)
        {
            var message = PythonExcelService.ClassifyConnectionError(output);
            Assert.Contains(hint, message);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("NameError: name 'x' is not defined")]
        public void Classify_connection_error_unrelated_returns_empty(string output)
        {
            Assert.Equal("", PythonExcelService.ClassifyConnectionError(output));
        }

        // ---------- PrependConnectPreamble：前置连接代码 ----------

        [Fact]
        public void Prepend_preamble_adds_import_on_first_line()
        {
            var full = PythonExcelService.PrependConnectPreamble("print('hi')");
            Assert.Equal("import xlwings as xw\nprint('hi')", full);
        }

        [Fact]
        public void Prepend_preamble_handles_null_code()
        {
            Assert.Equal("import xlwings as xw\n", PythonExcelService.PrependConnectPreamble(null));
        }

        [Fact]
        public void Prepend_preamble_preserves_multiline_body()
        {
            var full = PythonExcelService.PrependConnectPreamble("ws = xw.books.active.sheets[0]\nws.range(\"A1\").value = 42");
            Assert.Equal("import xlwings as xw\nws = xw.books.active.sheets[0]\nws.range(\"A1\").value = 42", full);
        }

        // ---------- QuoteArg：进程参数转义 ----------

        [Theory]
        [InlineData("abc", "\"abc\"")]
        [InlineData("a b", "\"a b\"")]
        [InlineData("a\"b", "\"a\"\"b\"")]
        public void Quote_arg_wraps_and_escapes(string input, string expected)
        {
            Assert.Equal(expected, PythonExcelService.QuoteArg(input));
        }
    }
}
