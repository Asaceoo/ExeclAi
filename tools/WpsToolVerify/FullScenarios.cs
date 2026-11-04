using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ExcelAiAssistant.Core.Tools;
using Newtonsoft.Json;

namespace WpsToolHarness
{
    /// <summary>
    /// full 模式：130 工具全量真机审查（WPS 兼容审查专用）。
    /// 与 battery（品类代表）互补：按 8 个分组场景构造真实数据，逐工具执行并输出 RESULT 行。
    /// 分档口径：OK=正常；FAIL+预期错误码（UNDO_UNAVAILABLE/VBA_UNAVAILABLE/受控未开启等）=预期降级；
    /// 其余 FAIL/EXCEPTION=候选真故障或用例参数偏差（结果行附 SCHEMA_REQUIRED=… 供 Round 2 修正用例）。
    /// 每场景独立工作表，跑完即删；全程临时簿内进行，不触碰任何真实数据。
    /// </summary>
    internal static class FullScenarios
    {
        private static dynamic _app;
        private static ToolRegistry _registry;

        private static Dictionary<string, object> Args(params object[] kv)
        {
            var dict = new Dictionary<string, object>();
            for (var i = 0; i + 1 < kv.Length; i += 2)
            {
                dict[(string)kv[i]] = kv[i + 1];
            }
            return dict;
        }

        private static void Seed(string sheet, string address, object values)
        {
            RunOne("write_excel_range", Args("sheet", sheet, "address", address, "values", values));
        }

        private static void RunOne(string name, Dictionary<string, object> args)
        {
            if (!_registry.TryGet(name, out var tool))
            {
                Console.WriteLine("RESULT\t" + name + "\tNOT_REGISTERED\t-\t-");
                return;
            }

            try
            {
                var result = tool.Execute(args);
                var payload = result?.Data == null ? "-" : JsonConvert.SerializeObject(result.Data);
                if (payload.Length > 2000)
                {
                    payload = payload.Substring(0, 2000) + "…";
                }

                Console.WriteLine("RESULT\t" + name + "\t" + (result?.Ok == true ? "OK" : "FAIL") + "\t" + (result?.ErrorCode ?? "-") + "\t" + payload);
            }
            catch (Exception ex)
            {
                var message = ex.Message;
                if (message.Length > 200)
                {
                    message = message.Substring(0, 200) + "…";
                }

                Console.WriteLine("RESULT\t" + name + "\tEXCEPTION\t" + ex.GetType().Name + "\t" + message);
            }

            // FAIL 时自动附 schema 必填参数清单，供用例修正（Round 2）快速定位
            if (_registry.TryGet(name, out var desc) && desc.Parameters != null
                && desc.Parameters.TryGetValue("required", out var req) && req is object[] arr && arr.Length > 0)
            {
                Console.WriteLine("SCHEMA_REQUIRED\t" + name + "\t" + string.Join(",", arr));
            }
        }

        private static void NewSheet(string name)
        {
            try
            {
                foreach (dynamic s in _app.Worksheets)
                {
                    if (string.Equals(Convert.ToString(s.Name), name, StringComparison.OrdinalIgnoreCase))
                    {
                        s.Delete();
                        break;
                    }
                }
            }
            catch { }

            dynamic ws = _app.Worksheets.Add();
            ws.Name = name;
            ws.Activate();
        }

        private static void DropSheet(string name)
        {
            try
            {
                foreach (dynamic s in _app.Worksheets)
                {
                    if (string.Equals(Convert.ToString(s.Name), name, StringComparison.OrdinalIgnoreCase))
                    {
                        s.Delete();
                        break;
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// m7 模式（M6 复审专项）：schema 对象项契约 + 静默成功防线真机验证。
        /// ① import_excel_file_sheet 按 {name} 对象项命中导入（修复前解析出类名字符串 → imported_sheets=[] 静默）；
        /// ② 指定不存在表名 → NO_MATCHED_SHEET 显式报错（修复前静默 OK）；
        /// ③ merge_workbooks 全部路径无效 → NO_VALID_SOURCE 显式报错（修复前静默 OK source_count=0）。
        /// </summary>
        public static int RunM7(ToolRegistry registry, dynamic app)
        {
            _registry = registry;
            _app = app;
            var tempDir = Path.GetTempPath();
            var xlsxPath = Path.Combine(tempDir, "wps_m7_ext.xlsx");

            var before = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (dynamic s in _app.Worksheets)
            {
                before.Add(Convert.ToString(s.Name));
            }

            try
            {
                string extSheetName;
                dynamic wb2 = _app.Workbooks.Add();
                try
                {
                    extSheetName = Convert.ToString(wb2.Worksheets(1).Name);
                    wb2.Worksheets(1).Range("A1").Value2 = "名";
                    wb2.Worksheets(1).Range("B1").Value2 = 7;
                    wb2.SaveAs(xlsxPath);
                }
                finally
                {
                    wb2.Close(false);
                }

                Console.WriteLine("M7_EXT_SHEET_NAME=" + extSheetName);
                RunOne("import_excel_file_sheet", Args("file_path", xlsxPath, "copy_as_new_sheet", true,
                    "sheets", new object[] { new Dictionary<string, object> { { "name", extSheetName } } }));
                RunOne("import_excel_file_sheet", Args("file_path", xlsxPath, "copy_as_new_sheet", true,
                    "sheets", new object[] { new Dictionary<string, object> { { "name", "NoSuchSheet_M7" } } }));
                RunOne("merge_workbooks", Args("source_workbooks", new object[] {
                    new Dictionary<string, object> { { "path", Path.Combine(tempDir, "no_such_m7.xlsx") } } }));
            }
            finally
            {
                var toDelete = new List<string>();
                foreach (dynamic s in _app.Worksheets)
                {
                    var n7 = Convert.ToString(s.Name);
                    if (!before.Contains(n7))
                    {
                        toDelete.Add(n7);
                    }
                }

                foreach (var n7 in toDelete)
                {
                    DropSheet(n7);
                }

                try { if (File.Exists(xlsxPath)) File.Delete(xlsxPath); } catch { }
            }

            return 0;
        }

        /// <summary>
        /// m17 模式（M17 保护态守卫专项，双宿主闭环）：
        /// ① 保护表写入（write_excel_range / flag_cells）必须返回 SHEET_PROTECTED 结构化错误
        ///    （修复前为裸 COM 错误码）；② 豁免工具 freeze_panes 在保护表上原生可用；
        /// ③ 解保后同一写入恢复 OK（守卫不误伤正常路径）。返回失败断言数（0 = 全过）。
        /// </summary>
        public static int RunM17(ToolRegistry registry, dynamic app)
        {
            _registry = registry;
            _app = app;
            var failed = 0;

            void Check(bool condition, string name, string detail)
            {
                Console.WriteLine((condition ? "[PASS] " : "[FAIL] ") + name + "　→ " + detail);
                if (!condition)
                {
                    failed++;
                }
            }

            string Describe(ExcelAiAssistant.Core.Models.ToolResult result)
            {
                if (result == null)
                {
                    return "null";
                }

                return result.Ok ? "ok" : (result.ErrorCode ?? "-") + " " + JsonConvert.SerializeObject(result.Data);
            }

            ExcelAiAssistant.Core.Models.ToolResult Run(string name, Dictionary<string, object> args)
            {
                if (!_registry.TryGet(name, out var tool))
                {
                    Console.WriteLine("RESULT\t" + name + "\tNOT_REGISTERED\t-\t-");
                    return null;
                }

                try
                {
                    return tool.Execute(args);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("RESULT\t" + name + "\tEXCEPTION\t" + ex.GetType().Name + "\t" + ex.Message);
                    return null;
                }
            }

            dynamic FindItem(string name)
            {
                foreach (dynamic s in _app.Worksheets)
                {
                    if (string.Equals(Convert.ToString(s.Name), name, StringComparison.OrdinalIgnoreCase))
                    {
                        return s;
                    }
                }

                return null;
            }

            try
            {
                var add = Run("add_excel_worksheet", Args("name", "M17P"));
                Check(add != null && add.Ok, "M17 准备：新增工作表", Describe(add));

                var ws = FindItem("M17P");
                Check(ws != null, "M17 准备：定位工作表", ws == null ? "未找到 M17P" : "found");
                if (ws == null)
                {
                    return failed;
                }

                ws.Protect(); // harness 直接设置保护态（工具层无保护开关，属真机环境构造）
                try
                {
                    var w1 = Run("write_excel_range", Args("sheet", "M17P", "address", "A1",
                        "values", new object[] { new object[] { "x" } }));
                    Check(w1 != null && !w1.Ok && w1.ErrorCode == "SHEET_PROTECTED",
                        "M17 负向①：保护表写入返回 SHEET_PROTECTED", Describe(w1));

                    var w2 = Run("flag_cells", Args("sheet", "M17P", "address", "A1:B2"));
                    Check(w2 != null && !w2.Ok && w2.ErrorCode == "SHEET_PROTECTED",
                        "M17 负向②：保护表标记返回 SHEET_PROTECTED", Describe(w2));

                    var w3 = Run("freeze_panes", Args("sheet", "M17P", "cell", "B2"));
                    Check(w3 != null && w3.Ok,
                        "M17 豁免：保护表冻结窗格不受守卫拦截", Describe(w3));
                }
                finally
                {
                    // 无状态 COM 约定：不跨工具调用持有工作表 RCW（豁免工具 freeze_panes
                    // 内部可能释放 M17P 的 RCW，WPS 下 RCW 身份共享会连带分离本引用）。
                    // 解保前按名重新解析，与 Excel harness 用例16 的清理模式一致。
                    var wsForUnprotect = FindItem("M17P");
                    try { wsForUnprotect?.Unprotect(); }
                    catch (Exception unEx) { Console.WriteLine("M17_UNPROTECT_EX=" + unEx.Message); }
                }

                var w4 = Run("write_excel_range", Args("sheet", "M17P", "address", "A1",
                    "values", new object[] { new object[] { "x" } }));
                Check(w4 != null && w4.Ok, "M17 收尾：解保后写入恢复（守卫不误伤）", Describe(w4));
            }
            finally
            {
                DropSheet("M17P");
            }

            return failed;
        }

        public static int Run(ToolRegistry registry, dynamic app)
        {
            _registry = registry;
            _app = app;
            var tempDir = Path.GetTempPath();

            // ---------- F1 读写基础 ----------
            NewSheet("HF1");
            Seed("HF1", "A1:B3", new object[] { new object[] { "张三 ", 10 }, new object[] { "", 20 }, new object[] { "李四", 30 } });
            RunOne("write_excel_cell", Args("sheet", "HF1", "address", "K1", "value", "单元写入"));
            RunOne("insert_text_into_cell", Args("sheet", "HF1", "address", "L1", "value", "插入文本"));
            RunOne("clear_excel_range", Args("sheet", "HF1", "address", "L1"));
            RunOne("verify_data_clean", Args("sheet", "HF1", "address", "L1", "checks", new object[] { new Dictionary<string, object> { ["type"] = "all_blank" } }));
            RunOne("data_validation", Args("sheet", "HF1", "address", "E1:E5", "type", "list", "source", "甲,乙,丙"));
            RunOne("verify_data_validation", Args("sheet", "HF1", "address", "E1:E5", "expected_type", "list", "expected_source", "甲,乙,丙"));
            RunOne("create_excel_sparkline", Args("sheet", "HF1", "data_range", "A2:B2", "dest_range", "M2", "type", "column"));
            RunOne("verify_sparkline", Args("sheet", "HF1", "dest_range", "M2", "expected_type", "column", "expected_data_source", "A2:B2"));
            RunOne("fill_number_sequence", Args("sheet", "HF1", "address", "E1:E10", "start", 1, "step", 1));
            RunOne("fill_random_numbers", Args("sheet", "HF1", "address", "F1:F5", "min", 1, "max", 100));
            RunOne("fill_blank_cells_down", Args("sheet", "HF1", "address", "A1:B3"));
            RunOne("fill_formula", Args("sheet", "HF1", "address", "G1", "formula", "=SUM(A1:B1)"));
            RunOne("number_pad", Args("sheet", "HF1", "address", "H1:H3", "column", 1, "width", 6, "pad_char", "0"));
            RunOne("bankers_round_values", Args("sheet", "HF1", "address", "I1:I3", "digits", 1));
            RunOne("reset_cell_value_type", Args("sheet", "HF1", "address", "A1:B3", "target_type", "text"));
            RunOne("wrap_formulas_with_iferror", Args("sheet", "HF1", "address", "G1:G3"));
            RunOne("format_excel_range", Args("sheet", "HF1", "address", "A1:B1", "bold", true));
            RunOne("verify_cell_format", Args("sheet", "HF1", "address", "A1:B1"));
            RunOne("verify_write_result", Args("sheet", "HF1", "address", "A1:B3"));
            RunOne("get_excel_range_values_and_formulas", Args("sheet", "HF1", "address", "A1:B3"));
            RunOne("preview_excel_sheet", Args("sheet", "HF1"));
            RunOne("select_excel_range", Args("sheet", "HF1", "address", "A1:B3"));
            RunOne("get_current_selection_values_and_formulas", new Dictionary<string, object>());
            RunOne("find_excel_cells", Args("sheet", "HF1", "what", 10));
            RunOne("cell_comment", Args("sheet", "HF1", "address", "A1", "text", "审查批注"));
            DropSheet("HF1");

            // ---------- F2 文本处理 ----------
            NewSheet("HF2");
            Seed("HF2", "A1:B4", new object[] { new object[] { "张三 ", "2024-01-05" }, new object[] { "李四", "2024/02/06" }, new object[] { "王五", "01.03" }, new object[] { 123.45, "hello world" } });
            Seed("HF2", "D1:D4", new object[] { new object[] { "甲" }, new object[] { "乙" }, new object[] { "丙" }, new object[] { "丁" } });
            RunOne("text_clean", Args("sheet", "HF2", "address", "A1:A4"));
            RunOne("text_merge", Args("sheet", "HF2", "address", "A1:B4", "columns", new object[] { 1, 2 }, "delimiter", "-", "target_column", 3));
            RunOne("text_split", Args("sheet", "HF2", "address", "B1:B4", "column", 1, "delimiter", " "));
            RunOne("delete_text_matches", Args("sheet", "HF2", "address", "A1:A4", "pattern", "\\s+"));
            RunOne("verify_data_clean", Args("sheet", "HF2", "address", "A1:A4", "checks", new object[] { new Dictionary<string, object> { ["type"] = "no_text_matches", ["pattern"] = "\\s+" } }));
            RunOne("convert_text_to_pinyin", Args("sheet", "HF2", "address", "A1:A3"));
            RunOne("convert_chinese_variant", Args("sheet", "HF2", "address", "A1:A3"));
            RunOne("extract_text_characters", Args("sheet", "HF2", "address", "B1:B4"));
            RunOne("date_normalize", Args("sheet", "HF2", "address", "B1:B4", "column", 1, "output_format", "yyyy-MM-dd"));
            RunOne("amount_to_uppercase", Args("sheet", "HF2", "address", "A4"));
            RunOne("convert_1d_to_2d", Args("sheet", "HF2", "address", "D1:D4", "columns", 2));
            RunOne("convert_2d_to_1d", Args("sheet", "HF2", "address", "A1:B4"));
            RunOne("make_payslip_rows", Args("sheet", "HF2", "address", "A1:B4", "header_address", "A1:B1", "data_address", "A2:B4"));
            DropSheet("HF2");

            // ---------- F3 公式与分析 ----------
            NewSheet("HF3");
            Seed("HF3", "A1:B6", new object[] { new object[] { 10, 100 }, new object[] { 20, 200 }, new object[] { 30, 300 }, new object[] { 40, 400 }, new object[] { 50, 500 }, new object[] { 60, 600 } });
            Seed("HF3", "C1", new object[] { new object[] { "=SUM(A1:A6)" } });
            RunOne("convert_formula_references", Args("sheet", "HF3", "address", "A1:C6"));
            RunOne("transform_excel_formulas", Args("sheet", "HF3", "address", "C1", "actions", new object[] { new Dictionary<string, object> { ["type"] = "iferror" } }));
            RunOne("explain_excel_formula", Args("sheet", "HF3", "address", "C1"));
            RunOne("trace_formula_references", Args("sheet", "HF3", "address", "C1"));
            RunOne("verify_formula", Args("sheet", "HF3", "address", "C1"));
            RunOne("analyze_range_statistics", Args("sheet", "HF3", "address", "A1:B6"));
            RunOne("analyze_column_correlation", Args("sheet", "HF3", "address", "A1:B6"));
            RunOne("analyze_value_distribution", Args("sheet", "HF3", "address", "A1:A6"));
            // M14：列质量画像（数值列 → numbers 剖析 + 无质量提示；负路径 INVALID_COLUMN）
            RunOne("analyze_column_profile", Args("sheet", "HF3", "address", "A1:B6", "column", 1));
            RunOne("analyze_column_profile", Args("sheet", "HF3", "address", "A1:B6", "column", 9));
            RunOne("rate_change_analysis", Args("sheet", "HF3", "address", "A1:B6", "value_field", 2));
            RunOne("pick_numbers_to_sum", Args("sheet", "HF3", "candidates_address", "A1:A6", "address", "C3"));
            RunOne("fuzzy_match_columns", Args("sheet", "HF3", "first_address", "A1:A6", "second_address", "B1:B6"));
            RunOne("cross_reference_columns", Args("sheet", "HF3", "lookup_address", "A1:A6", "address", "B1:B6"));
            RunOne("drill_excel_range", Args("sheet", "HF3", "address", "A1:B6"));
            DropSheet("HF3");

            // ---------- F4 清洗与筛选 ----------
            NewSheet("HF4");
            Seed("HF4", "A1:B8", new object[] { new object[] { "张三", 1 }, new object[] { "", "" }, new object[] { "李四", 2 }, new object[] { "张三", 1 }, new object[] { "王五", 3 }, new object[] { "", "" }, new object[] { "赵六", 4 }, new object[] { "钱七", 5 } });
            RunOne("clean_excel_data", Args("sheet", "HF4", "address", "A1:B8", "operations", new object[] { new Dictionary<string, object> { ["tool"] = "remove_blank_rows" } }));
            RunOne("delete_blank_rows_range", Args("sheet", "HF4", "address", "A1:B8"));
            RunOne("filter_excel_data", Args("sheet", "HF4", "address", "A1:B8", "column", 1, "operator", "contains", "value", "张", "target_address", "D1"));
            RunOne("manage_excel_filter", Args("sheet", "HF4", "action", "clear"));
            RunOne("insert_blank_rows_every_n", Args("sheet", "HF4", "address", "A1:B6", "n", 2));
            RunOne("merge_duplicate_cells", Args("sheet", "HF4", "address", "A1:A8"));
            RunOne("merge_or_split_cells", Args("sheet", "HF4", "address", "A1:B2", "action", "merge"));
            RunOne("verify_merge", Args("sheet", "HF4", "address", "A1:B2", "expected_merged", "true"));
            RunOne("merge_or_split_cells", Args("sheet", "HF4", "address", "A1:B2", "action", "unmerge"));
            RunOne("verify_merge", Args("sheet", "HF4", "address", "A1:B2", "expected_merged", "false"));
            RunOne("verify_data_clean", Args("sheet", "HF4", "address", "A1:B8"));
            RunOne("verify_sort_filter", Args("sheet", "HF4", "address", "A1:B8"));
            DropSheet("HF4");

            // ---------- F4t 智能表与保护探测 ----------
            NewSheet("HF4T");
            Seed("HF4T", "A1:C4", new object[] { new object[] { "姓名", "金额", "城市" }, new object[] { "张三", 10, "北京" }, new object[] { "李四", 20, "上海" }, new object[] { "王五", 30, "广州" } });
            RunOne("create_excel_table", Args("sheet", "HF4T", "address", "A1:C5", "name", "tbl_HF4T", "has_headers", "true"));
            // WPS ET 实测：ListObjects.Add 会把区域向下扩展一行（A1:C5 → A1:C6），故不断言行数，只断名称/表头/落点
            RunOne("verify_table", Args("sheet", "HF4T", "name", "tbl_HF4T", "address", "A1:C5", "expected_headers", new object[] { "列1", "列2", "列3" }));
            RunOne("verify_table", Args("sheet", "HF4T", "address", "B2"));
            RunOne("verify_protection", Args("sheet", "HF4T", "address", "A2", "expected_writable", "true"));
            DropSheet("HF4T");

            // ---------- F4i 图片与冻结窗格（M13：insert_image / verify_image / freeze_panes / verify_freeze_panes） ----------
            NewSheet("HF4I");
            var harnessPng = Path.Combine(tempDir, "wps_harness_insert.png");
            File.WriteAllBytes(harnessPng, Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAwCAIAAAAuKetIAAAAYklEQVR4nO3PQQkAMAzAwAqrfyZrIvY4BoEIuMzZ/brhgga0oAEtaEALGtCCBrSgAS1oQAsa0IIGtKABLWhACxrQgga0oAEtaEALGtCCBrSgAS1oQAsa0IIGtKABLWhACx67nJIA4m9ID9wAAAAASUVORK5CYII="));
            // 显式尺寸插入 + 全期望核对（锚点/数量/宽/高）
            RunOne("insert_image", Args("sheet", "HF4I", "image_path", harnessPng, "anchor", "E2", "width", 40, "height", 30));
            RunOne("verify_image", Args("sheet", "HF4I", "anchor", "E2", "expected_count", 1, "expected_width", 40, "expected_height", 30));
            // 缺省尺寸路径：AddPicture(-1,-1) 按图片原生尺寸（64x48px @96dpi → 48x36 磅，WPS/Excel 实测一致）
            RunOne("insert_image", Args("sheet", "HF4I", "image_path", harnessPng, "anchor", "H5"));
            RunOne("verify_image", Args("sheet", "HF4I", "anchor", "H5", "expected_count", 2, "expected_width", 48, "expected_height", 36));
            // 负路径：文件不存在应显式报 IMAGE_NOT_FOUND（防静默成功）
            RunOne("insert_image", Args("sheet", "HF4I", "image_path", Path.Combine(tempDir, "no_such_img_m13.png"), "anchor", "A1"));
            // 冻结 → 核对状态+位置 → 解冻 → 核对已解除（WPS 实测必须 SplitRow/SplitColumn + FreezePanes=true）
            RunOne("freeze_panes", Args("sheet", "HF4I", "cell", "B2"));
            RunOne("verify_freeze_panes", Args("sheet", "HF4I", "expected_frozen", "true", "expected_cell", "B2"));
            RunOne("freeze_panes", Args("sheet", "HF4I", "unfreeze", "true"));
            RunOne("verify_freeze_panes", Args("sheet", "HF4I", "expected_frozen", "false"));
            // 单侧冻结：A2 仅冻结第 1 行（SplitColumn=0 边界行为真机取证）
            RunOne("freeze_panes", Args("sheet", "HF4I", "cell", "A2"));
            RunOne("verify_freeze_panes", Args("sheet", "HF4I", "expected_frozen", "true", "expected_cell", "A2"));
            RunOne("freeze_panes", Args("sheet", "HF4I", "unfreeze", "true"));
            // 负路径：A1 不能作为冻结锚点（显式 INVALID_ANCHOR）
            RunOne("freeze_panes", Args("sheet", "HF4I", "cell", "A1"));
            // 负路径（M13-R 审查闭环）：width=0 → INVALID_ARGS；损坏图片文件 → IMAGE_INVALID（而非裸 COM 异常）
            RunOne("insert_image", Args("sheet", "HF4I", "image_path", harnessPng, "anchor", "A9", "width", 0));
            var corruptPng = Path.Combine(tempDir, "wps_corrupt_m13.png");
            File.WriteAllBytes(corruptPng, new byte[] { 0x00, 0x01, 0x02, 0x03, 0xFF });
            RunOne("insert_image", Args("sheet", "HF4I", "image_path", corruptPng, "anchor", "A9"));
            DropSheet("HF4I");

            // ---------- F4h 超链接（M15：manage_hyperlinks / verify_hyperlink） ----------
            NewSheet("HF4H");
            Seed("HF4H", "A1:A3", new object[] { new object[] { "官网" }, new object[] { "跳转" }, new object[] { "空位" } });
            RunOne("manage_hyperlinks", Args("sheet", "HF4H", "action", "add", "address", "A1", "link_url", "https://example.com/wps", "display_text", "官网"));
            RunOne("verify_hyperlink", Args("sheet", "HF4H", "address", "A1", "expected_url", "https://example.com/wps", "expected_text", "官网"));
            RunOne("manage_hyperlinks", Args("sheet", "HF4H", "action", "add", "address", "A2", "link_subaddress", "HF4H!A1"));
            RunOne("verify_hyperlink", Args("sheet", "HF4H", "address", "A2", "expected_subaddress", "HF4H!A1"));
            RunOne("manage_hyperlinks", Args("sheet", "HF4H", "action", "list"));
            // 负路径：重复添加 → HYPERLINK_EXISTS；update/remove 无链接单元格 → HYPERLINK_NOT_FOUND；双参数互斥 → INVALID_ARGS
            RunOne("manage_hyperlinks", Args("sheet", "HF4H", "action", "add", "address", "A1", "link_url", "https://example.com/2"));
            RunOne("manage_hyperlinks", Args("sheet", "HF4H", "action", "update", "address", "A3", "link_url", "https://example.com/x"));
            RunOne("manage_hyperlinks", Args("sheet", "HF4H", "action", "remove", "address", "A3"));
            RunOne("manage_hyperlinks", Args("sheet", "HF4H", "action", "add", "address", "A3", "link_url", "https://a.com", "link_subaddress", "HF4H!A1"));
            RunOne("manage_hyperlinks", Args("sheet", "HF4H", "action", "update", "address", "A1", "link_url", "https://example.com/updated"));
            RunOne("verify_hyperlink", Args("sheet", "HF4H", "address", "A1", "expected_url", "https://example.com/updated"));
            RunOne("manage_hyperlinks", Args("sheet", "HF4H", "action", "remove", "address", "A1:A2"));
            RunOne("verify_hyperlink", Args("sheet", "HF4H", "address", "A1:A2", "expected_present", "false"));
            RunOne("manage_hyperlinks", Args("sheet", "HF4H", "action", "remove", "address", "A1"));
            RunOne("manage_hyperlinks", Args("sheet", "HF4H", "action", "remove_all"));
            RunOne("verify_hyperlink", Args("sheet", "HF4H", "expected_count", 0));
            RunOne("manage_hyperlinks", Args("sheet", "HF4H", "action", "remove_all"));
            DropSheet("HF4H");

            // ---------- F5 结构与页面 ----------
            NewSheet("HF5");
            NewSheet("HF5b");
            NewSheet("HF5c");
            Seed("HF5", "A1:B6", new object[] { new object[] { "组1", 1 }, new object[] { "组2", 2 }, new object[] { "组1", 3 }, new object[] { "组2", 4 }, new object[] { "组3", 5 }, new object[] { "组3", 6 } });
            RunOne("copy_paste_excel_range", Args("sheet", "HF5", "source_address", "A1:B3", "target_address", "D1"));
            RunOne("copy_visible_range_only", Args("sheet", "HF5", "address", "A1:B3", "target_address", "F1"));
            RunOne("copy_visible_rows_to_sheet", Args("sheet", "HF5", "source_address", "A1:B3", "target_sheet", "HF5b", "target_address", "A1"));
            RunOne("copy_range_as_image", Args("sheet", "HF5", "address", "A1:B3", "file_path", Path.Combine(tempDir, "wps_harness_img.png")));
            RunOne("create_sheet_by_value", Args("sheet", "HF5", "address", "A1:B6", "column", 1, "name_cell", "A1"));
            RunOne("create_toc_sheet", new Dictionary<string, object>());
            RunOne("insert_excel_dimension", Args("sheet", "HF5", "address", "A5", "dimension", "row"));
            RunOne("delete_excel_dimension", Args("sheet", "HF5", "address", "A5", "dimension", "row"));
            RunOne("split_sheet_by_field", Args("sheet", "HF5", "address", "A1:B6", "column", 1));
            // 区分实验定案：WPS 空白表/单格 UsedRange 的 Value2 走 0 基数组路径抛 BADINDEX
            // （ReadMatrix 已修复 LowerBound 适配）。HF5d=有数据源表，HF5b/HF5c=空源表，
            // 三表混合合并验证 merge_sheets 对空源表的容错。
            NewSheet("HF5d");
            Seed("HF5d", "A1:B2", new object[] { new object[] { "x", 1 }, new object[] { "y", 2 } });
            RunOne("merge_sheets", Args("sheet", "HF5", "source_sheets", new object[] { "HF5d", "HF5b", "HF5c" }, "target_sheet", "HF5"));
            RunOne("protect_sheet_watermark", Args("sheet", "HF5", "address", "A1:B6", "text", "审查水印"));
            RunOne("apply_page_setup", Args("sheet", "HF5", "orientation", "landscape"));
            RunOne("manage_excel_named_range", Args("sheet", "HF5", "action", "create", "name", "HF_NAMED", "reference", "HF5!$A$1:$B$3"));
            RunOne("verify_structure", Args("expected_sheet_exists", "HF5b", "expected_sheet_absent", "HF_NO_SUCH"));
            DropSheet("HF5");
            DropSheet("HF5b");
            DropSheet("HF5c");

            // ---------- F6 导入导出 ----------
            NewSheet("HF6");
            var csvPath = Path.Combine(tempDir, "wps_harness_import.csv");
            File.WriteAllText(csvPath, "姓名,数量\n甲,1\n乙,2\n丙,3\n", new UTF8Encoding(true));
            RunOne("import_delimited_file", Args("file_path", csvPath, "delimiter", ",", "target_address", "A1"));
            var mdPath = Path.Combine(tempDir, "wps_harness_import.md");
            File.WriteAllText(mdPath, "| 名称 | 数量 |\n| --- | --- |\n| 甲 | 1 |\n| 乙 | 2 |\n", new UTF8Encoding(true));
            RunOne("import_markdown_to_excel", Args("markdown", "| 名称 | 数量 |\n| --- | --- |\n| 甲 | 1 |\n| 乙 | 2 |", "address", "A1"));
            RunOne("export_selection_to_markdown", Args("sheet", "HF6", "address", "A1:B3"));
            RunOne("create_excel_chart", Args("sheet", "HF6", "source_address", "A1:B3", "chart_type", "ColumnClustered"));
            RunOne("export_excel_chart_image", Args("sheet", "HF6", "file_path", Path.Combine(tempDir, "wps_harness_chart.png"), "overwrite", true)); // M28：导出覆盖门需显式 overwrite（巡检幂等）
            RunOne("smart_paste_delimiter", Args("sheet", "HF6", "address", "E1", "text", "甲,乙,丙", "delimiter", ","));
            DropSheet("HF6");

            // ---------- F7 报告与撤销 ----------
            NewSheet("HF7");
            Seed("HF7", "A1:B6", new object[] { new object[] { 1, 10 }, new object[] { 2, 20 }, new object[] { 3, 30 }, new object[] { 4, 40 }, new object[] { 5, 50 }, new object[] { 6, 60 } });
            RunOne("generate_analysis_report", Args("sheet", "HF7", "address", "A1:B6"));
            RunOne("undo_excel_operation", new Dictionary<string, object>());
            DropSheet("HF7");

            // ---------- F8 危险/受控类（预期降级档：VBA/Python/受控执行环境） ----------
            RunOne("run_vba_snippet", Args("snippet", "Beep"));
            RunOne("validate_vba_snippet", Args("snippet", "Sub T()\nEnd Sub"));
            RunOne("save_vba_tool", Args("name", "hf_review_tool", "description", "审查用", "code", "Sub T()\nEnd Sub"));
            RunOne("run_python_code", Args("code", "print(1)"));
            RunOne("run_python_excel", Args("code", "print(1)"));
            RunOne("run_cli_command", Args("command", "cmd /c echo ok"));
            RunOne("run_exe_command", Args("path", "cmd", "arguments", "/c echo ok"));
            RunOne("preview_external_file", Args("file_path", csvPath));
            RunOne("replace_external_links", Args("old_link", "D:\\no\\such.xlsx", "new_link", "D:\\no\\such2.xlsx"));
            RunOne("ct_format_answer_sheet", Args("address", "A1:B2"));
            RunOne("fetch_url", Args("url", "https://localhost/no/such"));

            // ---------- F9 补测：全量审查缺口工具（19 个静态工具真机补跑） ----------
            RunOne("get_current_excel_context", new Dictionary<string, object>());
            NewSheet("HF9");
            // 记录 F9 前的表名集合：import/merge 会引入外部表，收尾按差集清理
            var hf9Before = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (dynamic s in _app.Worksheets) hf9Before.Add(Convert.ToString(s.Name));
            Seed("HF9", "A1:C4", new object[] {
                new object[] { "姓名", "数量", "地区" },
                new object[] { "甲", 1, "华东" },
                new object[] { "乙", null, "华北" },
                new object[] { "丙", 3, null } });
            RunOne("preview_excel_range", Args("sheet", "HF9", "address", "A1:C4"));
            RunOne("recommend_excel_chart", Args("sheet", "HF9", "address", "A1:C4"));
            RunOne("ai_fill_blank_cells", Args("sheet", "HF9", "address", "A1:C4"));
            RunOne("apply_fill_blank_cells", Args("sheet", "HF9", "address", "A1:C4", "fills",
                new Dictionary<string, object> { { "B3", 2 }, { "C4", "华南" } }));
            RunOne("append_calculation_scratchpad", Args("input", "2+3", "formula", "=2+3", "expected_output", "5"));
            RunOne("clear_calculation_scratchpad", new Dictionary<string, object>());
            RunOne("generate_permutation_list", Args("sheet", "HF9", "n", 4, "output_address", "F1"));
            RunOne("create_pivot_table", Args("sheet", "HF9", "address", "A1:C4", "row_field", 1, "value_field", 2, "aggregation", "sum", "output_address", "H1"));
            RunOne("verify_pivot", Args("sheet", "HF9", "target_address", "H1"));
            NewSheet("HF9b");
            Seed("HF9b", "A1:C4", new object[] {
                new object[] { "姓名", "数量", "地区" },
                new object[] { "甲", 1, "华东" },
                new object[] { "乙", 2, "华北" },
                new object[] { "丙", 4, "华南" } });
            RunOne("compare_two_sheets", Args("sheet_a", "HF9", "address_a", "A1:C4", "sheet_b", "HF9b", "address_b", "A1:C4", "output_address", "L1"));
            RunOne("flag_cells", Args("sheet", "HF9", "address", "A1:B2"));
            RunOne("verify_marker", Args("sheet", "HF9", "address", "A1:B2"));
            RunOne("verify_conditional_format", Args("sheet", "HF9", "address", "A1:C4"));
            RunOne("insert_recommended_chart", Args("sheet", "HF9", "address", "A1:C4", "title", "补测图"));
            RunOne("verify_chart", Args("sheet", "HF9"));

            // 外部文件类：临时 xlsx（import/merge 用）+ 两个 csv（merge_delimited 用）
            var xlsxPath = Path.Combine(tempDir, "wps_harness_ext.xlsx");
            var f9csv1 = Path.Combine(tempDir, "wps_harness_m1.csv");
            var f9csv2 = Path.Combine(tempDir, "wps_harness_m2.csv");
            var splitDir = Path.Combine(tempDir, "wps_harness_split");
            try
            {
                dynamic wb2 = _app.Workbooks.Add();
                wb2.Worksheets(1).Range("A1").Value2 = "名";
                wb2.Worksheets(1).Range("B1").Value2 = "量";
                wb2.Worksheets(1).Range("A2").Value2 = "a";
                wb2.Worksheets(1).Range("B2").Value2 = 7;
                wb2.SaveAs(xlsxPath);
                wb2.Close(false);
            }
            catch (Exception ex9) { Console.WriteLine("RESULT\tsetup_external_xlsx\tEXCEPTION\t" + ex9.GetType().Name + "\t" + ex9.Message); }
            File.WriteAllText(f9csv1, "名,量\na,1\n", Encoding.UTF8);
            File.WriteAllText(f9csv2, "名,量\nb,2\n", Encoding.UTF8);
            RunOne("import_excel_file_sheet", Args("file_path", xlsxPath, "copy_as_new_sheet", true));
            RunOne("merge_delimited_files", Args("file_paths", new object[] {
                new Dictionary<string, object> { { "path", f9csv1 } },
                new Dictionary<string, object> { { "path", f9csv2 } } }, "output_sheet", "HF9", "output_address", "P1"));
            RunOne("merge_workbooks", Args("source_workbooks", new object[] {
                new Dictionary<string, object> { { "path", xlsxPath } } }));
            RunOne("split_workbook", Args("output_folder", splitDir, "sheets", new object[] {
                new Dictionary<string, object> { { "name", "HF9b" } } }));

            // 清理 F9 引入的表与临时文件
            DropSheet("HF9");
            DropSheet("HF9b");
            try
            {
                var f9ToDelete = new List<string>();
                foreach (dynamic s in _app.Worksheets)
                {
                    var n9 = Convert.ToString(s.Name);
                    if (!hf9Before.Contains(n9)) f9ToDelete.Add(n9);
                }
                foreach (var n9 in f9ToDelete)
                {
                    foreach (dynamic s in _app.Worksheets)
                    {
                        if (string.Equals(Convert.ToString(s.Name), n9, StringComparison.OrdinalIgnoreCase)) { s.Delete(); break; }
                    }
                }
            }
            catch { }
            try { if (File.Exists(xlsxPath)) File.Delete(xlsxPath); } catch { }
            try { if (File.Exists(f9csv1)) File.Delete(f9csv1); } catch { }
            try { if (File.Exists(f9csv2)) File.Delete(f9csv2); } catch { }
            try { if (Directory.Exists(splitDir)) Directory.Delete(splitDir, true); } catch { }

            // 清理临时文件
            try { if (File.Exists(csvPath)) File.Delete(csvPath); } catch { }
            try { if (File.Exists(mdPath)) File.Delete(mdPath); } catch { }
            return 0;
        }
    }
}
