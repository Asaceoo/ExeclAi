using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Excel;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// 锁定 ExcelImportService 批量落地的修复：
    /// (1) detected_encoding 在 encoding=auto 时按 BOM 嗅探回真实编码（不再拿 UTF-8 fallback 冒充）；
    /// (2) MergeDelimitedFiles 单文件不可读时隔离跳过，source_count/failed_files 等 extra 字段真正到达最终 Data。
    /// 纯文件操作，不依赖 COM，用 fake writer 捕获矩阵与回调。
    /// </summary>
    public sealed class ExcelImportServiceTests
    {
        private static string TempCsv(string content, bool bom)
        {
            var path = Path.Combine(Path.GetTempPath(), "imp_" + Guid.NewGuid().ToString("N") + ".csv");
            File.WriteAllText(path, content, bom ? new UTF8Encoding(true) : new UTF8Encoding(false));
            return path;
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (path != null && File.Exists(path)) File.Delete(path);
            }
            catch
            {
            }
        }

        private static ExcelImportService MakeService(
            Func<Dictionary<string, object>, object[,], string, string, ToolResult> writer = null)
        {
            return new ExcelImportService(new object(), writer);
        }

        [Fact]
        public void SniffBomEncoding_reports_real_bom()
        {
            var plain = TempCsv("a,b\n1,2", bom: false);
            var utf8Bom = TempCsv("a,b\n1,2", bom: true);
            var utf16 = Path.Combine(Path.GetTempPath(), "imp_" + Guid.NewGuid().ToString("N") + ".csv");
            File.WriteAllBytes(utf16, new byte[] { 0xFF, 0xFE, 0x61, 0x00 });

            try
            {
                Assert.Equal("utf-8", ExcelImportService.SniffBomEncoding(plain));
                Assert.Equal("utf-8-bom", ExcelImportService.SniffBomEncoding(utf8Bom));
                Assert.Equal("utf-16-le", ExcelImportService.SniffBomEncoding(utf16));
            }
            finally
            {
                TryDelete(plain);
                TryDelete(utf8Bom);
                TryDelete(utf16);
            }
        }

        [Theory]
        [InlineData(false, "utf-8")]
        [InlineData(true, "utf-8-bom")]
        public void Preview_detected_encoding_is_bom_aware(bool bom, string expected)
        {
            var path = TempCsv("a,b\n1,2\n3,4", bom);
            try
            {
                var result = MakeService().PreviewExternalFile(
                    new Dictionary<string, object> { ["file_path"] = path, ["encoding"] = "auto" });
                Assert.True(result.Ok, result.Data as string);
                dynamic data = result.Data;
                Assert.Equal(expected, (string)data.detected_encoding);
            }
            finally
            {
                TryDelete(path);
            }
        }

        [Fact]
        public void Merge_isolates_skipped_file_and_extra_reaches_data()
        {
            var csv1 = TempCsv("a,b\n1,2\n3,4", bom: false);
            var csv2 = TempCsv("a,b\n5,6", bom: false);
            var missing = Path.Combine(Path.GetTempPath(), "imp_none_" + Guid.NewGuid().ToString("N") + ".csv");
            object[,] captured = null;
            var service = MakeService(
                (_wargs, matrix, _mode, _tool) =>
                {
                    captured = matrix;
                    return new ToolResult { Ok = true };
                });

            try
            {
                var result = service.MergeDelimitedFiles(new Dictionary<string, object>
                {
                    ["file_paths"] = new[] { csv1, csv2, missing },
                    ["output_address"] = "A1",
                    ["has_header"] = true
                });

                Assert.True(result.Ok, result.Data as string);
                Assert.NotNull(captured);
                Assert.Equal(4, captured.GetLength(0)); // 表头 + 3 行数据
                Assert.Equal(2, captured.GetLength(1));

                // extra（source_count/failed_files/combined_*）必须真实到达最终 Data，而非仅参数传递层。
                var data = Assert.IsAssignableFrom<IDictionary<string, object>>(result.Data);
                Assert.Equal(2, Convert.ToInt32(data["source_count"])); // 实际合并 2 个，不存在的文件未计入
                Assert.Equal(4, Convert.ToInt32(data["combined_rows"]));
                Assert.Equal(2, Convert.ToInt32(data["combined_columns"]));
                Assert.True(data.ContainsKey("failed_files"));
                // 工具分类审查第九批改进：不存在的路径从「静默跳过」改为「计入 failed_files 显式回显」——
                // source_count 少 1 必须有解释（AI 使用方视角）。
                var failed = (data["failed_files"] as System.Collections.IEnumerable)?.Cast<object>().ToArray() ?? new object[0];
                Assert.Single(failed);
                Assert.Contains("文件不存在", Convert.ToString(failed[0]));
            }
            finally
            {
                TryDelete(csv1);
                TryDelete(csv2);
            }
        }

        [Fact]
        public void Merge_skips_unreadable_file_and_keeps_rest()
        {
            var normal = TempCsv("a,b\n1,2", bom: false);
            var locked = TempCsv("x,y\n7,8", bom: false);
            object[,] captured = null;
            var service = MakeService(
                (_wargs, matrix, _mode, _tool) =>
                {
                    captured = matrix;
                    return new ToolResult { Ok = true };
                });

            try
            {
                // FileShare.None 独占句柄使第二个打开在 Merge 的 ReadLines 抛 IOException，走 failed_files 隔离分支。
                using (var hold = new FileStream(locked, FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    var result = service.MergeDelimitedFiles(new Dictionary<string, object>
                    {
                        ["file_paths"] = new[] { normal, locked },
                        ["output_address"] = "A1",
                        ["has_header"] = true
                    });

                    Assert.True(result.Ok, result.Data as string);
                    Assert.NotNull(captured);
                    Assert.Equal(2, captured.GetLength(0)); // 表头 + 1 行（只有 normal 并入）
                    Assert.Equal(2, captured.GetLength(1));

                    var data = Assert.IsAssignableFrom<IDictionary<string, object>>(result.Data);
                    Assert.Equal(1, Convert.ToInt32(data["source_count"]));
                    var failed = ((System.Collections.IEnumerable)data["failed_files"])
                        .Cast<object>().Select(x => Convert.ToString(x)).ToArray();
                    Assert.Single(failed);
                    Assert.Contains(locked, failed);
                }

                TryDelete(locked);
            }
            finally
            {
                TryDelete(normal);
            }
        }

        [Fact]
        public void Merge_all_empty_files_returns_empty_result()
        {
            var empty = TempCsv("", bom: false);
            var missing = Path.Combine(Path.GetTempPath(), "imp_none_" + Guid.NewGuid().ToString("N") + ".csv");
            var service = MakeService(
                (_wargs, _matrix, _mode, _tool) => new ToolResult { Ok = true });
            try
            {
                var result = service.MergeDelimitedFiles(new Dictionary<string, object>
                {
                    ["file_paths"] = new[] { empty, missing },
                    ["output_address"] = "A1",
                    ["has_header"] = true
                });
                Assert.False(result.Ok);
                Assert.Equal("EMPTY_RESULT", result.ErrorCode);
            }
            finally
            {
                TryDelete(empty);
            }
        }
    }
}