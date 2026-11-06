using System;
using System.Collections.Generic;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;

namespace ExcelAiAssistant.Excel
{
    public static class ExcelSecondPhaseRegistration
    {
        private static readonly Func<Dictionary<string, object>, object[,], string, string, ToolResult> NoWriter =
            (_1, _2, _3, _4) => throw new InvalidOperationException("缺少写入回调。");

        public static void RegisterSecondPhaseTools(
            ToolRegistry registry,
            dynamic application,
            Func<Dictionary<string, object>, object[,], string, string, ToolResult> writer,
            ScriptOptions scriptOptions)
        {
            var markdown = new ExcelMarkdownService(application);
            var edit = new ExcelAdvancedEditService(application, writer ?? NoWriter);
            var workbook = new ExcelWorkbookService(application);
            var script = new ScriptRunnerService(application, scriptOptions ?? new ScriptOptions());
            var pyExcel = new PythonExcelService(application, scriptOptions ?? new ScriptOptions());
            var pyEnv = new PythonEnvService(application, scriptOptions ?? new ScriptOptions());

            Register(registry, "export_selection_to_markdown", "把选区导出为 Markdown 表格文本。", RiskLevel.ReadOnly,
                () => markdown.ExportToMarkdown, SecondPhaseSchemas.ExportMarkdown());
            Register(registry, "import_markdown_to_excel", "把 Markdown 表格文本导入为 Excel 区域。", RiskLevel.NormalWrite,
                () => args => markdown.ImportToExcel(args, writer ?? NoWriter), SecondPhaseSchemas.ImportMarkdown());

            Register(registry, "merge_or_split_cells", "合并/取消合并区域，或按相同内容合并、拆分并填充。", RiskLevel.StructuralChange,
                () => edit.ManageMergeSplit, SecondPhaseSchemas.MergeSplit());
            Register(registry, "reset_cell_value_type", "把区域单元格重置为文本、数字或日期类型。", RiskLevel.NormalWrite,
                () => edit.ResetValueType, SecondPhaseSchemas.ResetValueType());
            Register(registry, "transform_excel_formulas", "对区域公式批量应用 IFERROR/IFNA/ROUND 等转换。", RiskLevel.NormalWrite,
                () => edit.TransformFormulas, SecondPhaseSchemas.TransformFormulas());
            Register(registry, "convert_formula_references", "切换公式中的绝对/相对引用。", RiskLevel.NormalWrite,
                () => edit.ConvertFormulaReferences, SecondPhaseSchemas.ConvertRefs());

            Register(registry, "trace_formula_references", "定位公式直接引用（前导项）的单元格与值。", RiskLevel.ReadOnly,
                () => edit.TraceFormulaReferences, SecondPhaseSchemas.TraceRefs());
            Register(registry, "drill_excel_range", "下钻单元格：读取其显示值、公式与引用来源预览。", RiskLevel.ReadOnly,
                () => edit.DrillExcelRange, SecondPhaseSchemas.DrillRange());
            Register(registry, "copy_visible_rows_to_sheet", "把筛选后的可见行复制到目标区域。", RiskLevel.StructuralChange,
                () => edit.CopyVisibleRows, SecondPhaseSchemas.CopyVisibleRows());

            Register(registry, "apply_page_setup", "设置工作表的打印页面布局。", RiskLevel.NormalWrite,
                () => workbook.ApplyPageSetup, SecondPhaseSchemas.PageSetup());
            Register(registry, "export_sheets_to_pdf", "把整本或指定工作表导出为 PDF（目标已存在默认拒绝，overwrite=true 才覆盖并先备份；加载项/AgentRunner 入口导出后自动带出 verify_file，MCP/CLI 直调时请显式调用 verify_file）。", RiskLevel.NormalWrite,
                () => workbook.ExportSheetsToPdf, SecondPhaseSchemas.ExportPdf());
            Register(registry, "merge_workbooks", "把指定工作簿的各工作表合并进当前工作簿。", RiskLevel.StructuralChange,
                () => workbook.MergeWorkbooks, SecondPhaseSchemas.MergeWorkbooks());
            Register(registry, "split_workbook", "把当前工作簿按工作表拆分为独立文件。目标 xlsx 已存在时默认跳过（skipped_sheets/skipped_count），真失败才进 failed_sheets；overwrite=true 才覆盖。ok=true 仅代表至少一张成功。", RiskLevel.StructuralChange,
                () => workbook.SplitWorkbook, SecondPhaseSchemas.SplitWorkbook());
            Register(registry, "save_workbook", "保存当前工作簿（可选另存到指定路径 file_path，支持 .xlsx/.xlsm/.xls）。MCP/CLI 等外部入口完成写入后必须调用本工具落盘，否则宿主进程退出时全部改动丢失；拆分/导出后也可用本工具固化主簿状态。另存到已存在的其他文件需显式 overwrite=true（默认拒绝，防误覆盖）。", RiskLevel.NormalWrite,
                () => workbook.SaveWorkbook, SecondPhaseSchemas.SaveWorkbook());
            Register(registry, "list_external_links", "列出当前工作簿的外部链接与查询连接。", RiskLevel.ReadOnly,
                () => workbook.ListExternalLinks, SecondPhaseSchemas.ListExternalLinks());
            Register(registry, "replace_external_links", "替换工作簿中的外部链接。", RiskLevel.Dangerous,
                () => workbook.ReplaceExternalLinks, SecondPhaseSchemas.ReplaceExternalLinks());

            Register(registry, "validate_vba_snippet", "校验 VBA 代码片段语法/结构。", RiskLevel.NormalWrite,
                () => script.ValidateVbaSnippet, SecondPhaseSchemas.VbaValidate());
            Register(registry, "run_vba_snippet", "当现有工具无法完成用户需求时，编写 VBA 代码片段用本工具执行以达成目的（万能兜底工具）。代码必须直接操作当前 ActiveWorkbook/ActiveSheet/Selection，逻辑自包含、可独立运行，且必须能在数秒内自行终止（本工具无外部超时中断，死循环会一直运行直到手动中断），并在执行前先向用户展示代码确认。参数 snippet 为完整 VBA 代码（注意：VBE 代码模块源码在中文系统按 ANSI/GBK 存储，snippet 与 args 内请使用 ASCII/中文等系统代码页可表示字符，emoji 等非 BMP 字符注入后会被替换为 ? 而损坏）。数据通道：可传 args（JSON 对象/数组/标量，支持嵌套）注入模块级 gXaiArgs，代码内用 XaiArg(\"key\") 读取（返回字符串/数字/布尔/数组，缺键返回 Null）；需要把计算结果结构化交回时，调用 XaiReturn 值（支持标量/数组/字典），执行器以 result 字段返回 JSON（示例：XaiReturn Array(v, 42) → result=[\"A\",42]；结果超 result_limit_kb 时 result 为截断字符串预览且 result_truncated=true）。持久状态：用 XaiStateSet/XaiStateGet 在多次调用间保留键值状态，state_clear=true 显式清空；返回的 state_reset=true 表示状态模块刚重建、旧状态已丢失（工作簿首次调用或升级后），勿假定旧状态仍存在。状态仅存内存、Excel 关闭即清、不落盘。铁律（违者会把整段代码判定非法驳回）：① 大批量读写一律用数组/Range.Value2 整体赋值（先 arr=Range.Value2，处理数组后再一次写回），严禁在 For/While 循环里用 Cells()/Range().Value 逐格读写；② 严禁对多行多列矩形 Range 直接调用 .AutoFit（如 ws.Range(\"G12:K18\").AutoFit 运行时会抛 1004），改按行/列调用 .Rows(n:m).AutoFit/.Columns(x:y).AutoFit 或设 RowHeight/ColumnWidth=0；③ 开头关闭 Application.ScreenUpdating 与 Calculation=Manual，结束前恢复；④ 成功执行后必须再调用 save_vba_tool 沉淀。", RiskLevel.Dangerous,
                () => script.RunVbaSnippet, SecondPhaseSchemas.VbaRun());
            Register(registry, "run_python_code", "在受控 Python（常驻 worker）下执行纯计算/数据处理代码，结果以 stdout 文本返回。数据通道：可传 input_data（JSON 对象/数组/标量）注入命名空间变量 input_data 直接使用；需要结构化结果时调用 xai_result(obj)（支持 dict/list/标量，numpy/pandas 自动转换），执行器以 result 字段返回 JSON（示例：xai_result({'got': input_data['x']}) → result={\"got\":...}；结果超 result_limit_kb 时 result 为截断字符串预览且 result_truncated=true）。持久状态：persist=true 保留命名空间跨调用（变量/自定义函数可复用；注意 run_python_code 与 run_python_excel 共享同一 worker 的 persist 命名空间，两者变量互通）；xai_state 为常驻 dict（无论 persist 与否都保留）；state_clear=true 一并清空命名空间与 xai_state；返回的 state_reset=true 表示 worker 刚重建/超时重启/被清空，旧状态已丢失，勿假定仍存在。状态仅存内存、进程退出即清、不落盘。约束：① 默认 30 秒超时（可传 timeout_ms 放宽，上限 300 秒），严禁死循环、阻塞等待 input()、访问网络或超大迭代（如 range(1e9)），否则必然超时并重置 worker；② 只做纯 Python 计算，超大输出会被截断；③ 若要把结果写回 Excel 单元格，请改用 run_vba_snippet 用数组一次性写回，不要在 python 里拼大字符串再回填；④ 序列/字典下标访问前先做 len()/in 边界检查（IndexError/KeyError 高频源，行数据列数不足时先截断再处理）。", RiskLevel.Dangerous,
                () => script.RunPythonCode, SecondPhaseSchemas.RunPython());
            Register(registry, "run_python_excel", "在受控 Python（常驻 worker + xlwings）中直接读写当前 Excel 活动工作簿。已自动前置 import xlwings as xw；用 xw.apps.active 连当前实例、xw.books.active 定位活动簿（示例：xw.books.active.sheets[0].range(\"A1\").value = 42）。数据通道：可传 input_data（JSON 对象/数组/标量）注入 input_data 变量；需要结构化结果时调用 xai_result(obj)，执行器以 result 字段返回 JSON（结果超 result_limit_kb 时 result 为截断字符串预览且 result_truncated=true）。持久状态：persist=true 保留命名空间跨调用（与 run_python_code 共享同一 worker 的 persist 命名空间，变量互通）；xai_state 为常驻 dict（无论 persist 与否都保留）；state_clear=true 一并清空命名空间与 xai_state；返回的 state_reset=true 表示 worker 刚重建/超时重启/被清空，旧状态已丢失，勿假定仍存在。状态仅存内存、进程退出即清、不落盘。安全契约：本工具可绕过操作记录直写 Excel，风险与 run_vba_snippet 同级（Dangerous + 确认门 + 受控执行总开关）；首次使用自动探测 Python 版本/位数（要求 ≥3.9 且与 Excel 位数一致，不符给出安装指引）并自动 pip 安装缺失依赖 xlwings/pywin32（镜像兜底 → --user 兜底 → 装后 import 自检，失败分级提示）；AI 声明 write_ranges 时执行前快照该区域、执行失败自动回滚；若只需纯计算请用 run_python_code（低风险）。约束：默认 30 秒超时（可传 timeout_ms 放宽，上限 300 秒），严禁死循环/阻塞/访问网络；代码必须能在数秒内自行终止。另：openpyxl/xlw 磁盘引擎只处理未打开的文件，要改当前已打开的工作簿必须走 xw.books.active 直写（对已打开文件 to_excel/save 会报 PermissionError——Excel 进程持锁）。", RiskLevel.Dangerous,
                () => pyExcel.Run, SecondPhaseSchemas.RunPythonExcel());
            Register(registry, "run_cli_command", "在受控白名单下运行 CLI 命令。", RiskLevel.Dangerous,
                () => script.RunCliCommand, SecondPhaseSchemas.RunCli());
            Register(registry, "list_python_packages", "列出当前 Python 环境已安装的库（名称/版本/安装位置）与 Python 解释器路径、版本、位数。用户问装了哪些库/库在哪时调用。", RiskLevel.ReadOnly,
                () => pyEnv.ListPackages, SecondPhaseSchemas.ListPythonPackages());
            Register(registry, "install_python_packages", "常用 Python 库一键安装（pandas/matplotlib/openpyxl/xlwings 等，白名单+国内镜像兜底）。必须先把 python_setup_guide 返回的库目录与用途说明展示给用户、等用户明确选择要装哪些后再调用。", RiskLevel.Dangerous,
                () => pyEnv.InstallPackages, SecondPhaseSchemas.InstallPythonPackages());
            Register(registry, "python_setup_guide", "Python 环境检测与小白级安装教程：当前 Python 是否安装/位数是否匹配 Excel、逐步安装引导、常用库目录（每个库做什么的详细说明与安装命令）。用户环境缺 Python/缺库或询问如何装库时必调。", RiskLevel.ReadOnly,
                () => pyEnv.SetupGuide, SecondPhaseSchemas.PythonSetupGuide());
            Register(registry, "run_exe_command", "在受控白名单下运行可执行文件。", RiskLevel.Dangerous,
                () => script.RunExeCommand, SecondPhaseSchemas.RunExe());
        }

        private static void Register(
            ToolRegistry registry,
            string name,
            string description,
            RiskLevel risk,
            Func<Func<Dictionary<string, object>, ToolResult>> target,
            Dictionary<string, object> parameters)
        {
            var execute = target();
            registry.Register(new ToolDescriptor(name, description, risk, args => Execute(() => execute(args)), parameters));
        }

        private static ToolResult Execute(Func<ToolResult> action)
        {
            return ExcelToolHost.Execute(action);
        }
    }
}