using System.Collections.Generic;
using ExcelAiAssistant.Core.Diagnostics;
using ExcelAiAssistant.Excel;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>一键自检日志降噪（LogIssueExtractor）+ run_python_code 缺模块自愈解析（TryResolveMissingModule）。</summary>
    public class DiagnosticNoiseReductionTests
    {
        // ---------- LogIssueExtractor：心跳行与配置回显不误报（真机实证「发现 17 处异常」中 4 条为回合开始 Info 行） ----------

        [Fact]
        public void Heartbeat_turn_start_lines_are_skipped()
        {
            var lines = new List<string>
            {
                "16:30:09.370 [Info] [agent] 回合开始 工具上限=1000 空闲超时=300s",
                "16:25:28.879 [Info] [agent] 回合开始 工具上限=1000 空闲超时=300s"
            };
            Assert.Empty(LogIssueExtractor.Extract(lines));
        }

        [Fact]
        public void Idle_timeout_config_echo_alone_is_not_an_issue()
        {
            // 旧版泛化「超时」签名会把「空闲超时=300s」配置回显误报为异常；收紧后不应命中。
            var lines = new List<string>
            {
                "16:19:01.676 [Info] [agent] 回合开始 工具上限=1000 空闲超时=300s"
            };
            Assert.Empty(LogIssueExtractor.Extract(lines));
        }

        [Fact]
        public void Real_timeout_failure_is_still_reported()
        {
            var lines = new List<string>
            {
                "16:20:00.000 [Error] [python] 执行超时(limit=30000ms)，worker 已终止重建"
            };
            var issues = LogIssueExtractor.Extract(lines);
            Assert.Single(issues);
            Assert.Contains("执行超时", issues[0]);
        }

        [Fact]
        public void Info_level_tool_log_with_error_code_is_reported()
        {
            var lines = new List<string>
            {
                "16:21:21.616 [Info] [tool] 工具[create_excel_chart] 70ms EXCEL_COM_ERROR EXCEL_COM_ERROR: 缺少必填参数：source_address"
            };
            var issues = LogIssueExtractor.Extract(lines);
            Assert.Single(issues);
            Assert.Contains("EXCEL_COM_ERROR", issues[0]);
        }

        [Fact]
        public void Ok_false_in_info_is_reported_and_duplicates_collapsed()
        {
            var lines = new List<string>
            {
                "16:21:31.691 [Error] tool:run_python_excel ok=False WRITE_RANGE_INVALID write_ranges 项格式无效：Sheet1!A52:C60",
                "16:21:31.691 [Error] tool:run_python_excel ok=False WRITE_RANGE_INVALID write_ranges 项格式无效：Sheet1!A52:C60"
            };
            Assert.Single(LogIssueExtractor.Extract(lines));
        }

        [Fact]
        public void Same_failure_across_layers_and_rotated_copies_collapses_to_one()
        {
            // v3.7.1.58 实证：同一次 Python 失败以三种行文落日志（worker/工具事件/回合摘要），
            // 且主日志与轮转副本同文——旧整行精确去重挡不住，一键自检把 10 个事件虚报 31 处。
            var lines = new List<string>
            {
                "15:15:20.169 [Info] 工具[run_python_excel] PYTHON_RUN_ERROR PYTHON_RUN_ERROR: Traceback (most recent call last): File \"C:\\Temp\\xai_python_worker.py\", line 124",
                "15:15:20.167 [Info] [tool] 工具[run_python_excel] 1122ms PYTHON_RUN_ERROR PYTHON_RUN_ERROR: Traceback (most recent call last): File \"C:\\Temp\\xai_python_worker.py\", line 124",
                "15:15:20.167 [Error] [python] 执行失败 exit=1: Traceback (most recent call last): File \"C:\\Temp\\xai_python_worker.py\", line 124",
                "15:15:20.168 [Error] tool:run_python_excel ok=False PYTHON_RUN_ERROR PYTHON_RUN_ERROR: Traceback",
                "15:15:20.169 [Error] tool:run_python_excel ok=False PYTHON_RUN_ERROR PYTHON_RUN_ERROR: Traceback（轮转副本同文）"
            };
            Assert.Single(LogIssueExtractor.Extract(lines));
        }

        [Fact]
        public void Same_code_events_in_different_time_buckets_are_kept_separate()
        {
            // 不同时间点的独立重试不是重复：15:14:53 与 15:15:20 两次失败都要保留。
            var lines = new List<string>
            {
                "15:14:53.243 [Error] [python] 执行失败 exit=1: Traceback (most recent call last): line 193",
                "15:15:20.167 [Error] [python] 执行失败 exit=1: Traceback (most recent call last): line 153"
            };
            Assert.Equal(2, LogIssueExtractor.Extract(lines).Count);
        }

        [Fact]
        public void Successful_chart_probe_line_is_not_an_issue()
        {
            var lines = new List<string>
            {
                "16:33:00.000 [Info] 图表实况核对：ChartType=51（簇状柱形图） 系列 1 数据点 8 有标题"
            };
            Assert.Empty(LogIssueExtractor.Extract(lines));
        }

        [Fact]
        public void Extract_caps_at_30_entries()
        {
            var lines = new List<string>();
            for (var i = 0; i < 40; i++)
            {
                lines.Add(i.ToString("00") + ":00:00.000 [Error] tool:demo ok=False FAIL_" + i);
            }

            Assert.Equal(30, LogIssueExtractor.Extract(lines).Count);
        }

        // ---------- ScriptRunnerService.TryResolveMissingModule：缺模块白名单解析 ----------

        [Theory]
        [InlineData("Traceback ... ModuleNotFoundError: No module named 'PIL'", "PIL", "pillow")]
        [InlineData("ModuleNotFoundError: No module named 'cv2'", "cv2", "opencv-python")]
        [InlineData("ModuleNotFoundError: No module named 'sklearn'", "sklearn", "scikit-learn")]
        [InlineData("ModuleNotFoundError: No module named 'yaml'", "yaml", "pyyaml")]
        public void Missing_module_resolves_to_whitelisted_pip_package(string error, string module, string pip)
        {
            var ok = ScriptRunnerService.TryResolveMissingModule(error, out var m, out var p);
            Assert.True(ok);
            Assert.Equal(module, m);
            Assert.Equal(pip, p);
        }

        [Theory]
        [InlineData("ModuleNotFoundError: No module named 'some_obscure_pkg'")]
        [InlineData("NameError: name 'x' is not defined")]
        [InlineData("")]
        [InlineData(null)]
        public void Missing_module_outside_whitelist_is_not_auto_installed(string error)
        {
            Assert.False(ScriptRunnerService.TryResolveMissingModule(error, out _, out _));
        }

        [Fact]
        public void Missing_module_map_has_no_duplicate_pip_targets_collision()
        {
            // 白名单兜底：所有映射键非空、包名非空（防拼进空的 pip 参数）。
            foreach (var kv in ScriptRunnerService.MissingModulePipMap)
            {
                Assert.False(string.IsNullOrWhiteSpace(kv.Key));
                Assert.False(string.IsNullOrWhiteSpace(kv.Value));
            }
        }
    }
}
