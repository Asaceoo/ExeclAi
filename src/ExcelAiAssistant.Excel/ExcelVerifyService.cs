using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelAiAssistant.Core.Models;
using Newtonsoft.Json.Linq;

using ExcelAiAssistant.Core.Diagnostics;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// 四类操作结果自检工具（只读，不修改数据）：
    /// verify_cell_format（数字/对齐/字体/边框/填充/保护）、verify_conditional_format、
    /// verify_sort_filter（含排序依据）、verify_formula（语法/预期/计算值）。
    /// 统一「预期值 vs 实际值」判定模式：AI 操作时传入预期参数，自检读取实际属性后逐项比对，
    /// 产出 verdict + issues + selfcheck 三要素（与 verify_chart 一致）。
    /// </summary>
    public sealed class ExcelVerifyService
    {
        private readonly dynamic application;

        public ExcelVerifyService(object application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
        }

        // ===================== 单元格格式自检 =====================

        public ToolResult VerifyCellFormat(Dictionary<string, object> args)
        {
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            dynamic font = null;
            dynamic interior = null;
            dynamic borders = null;
            dynamic protection = null;
            try
            {
                var sheetName = Optional(args, "sheet");
                var address = Required(args, "address");
                workbook = application.ActiveWorkbook ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                sheet = string.IsNullOrWhiteSpace(sheetName) ? application.ActiveSheet : workbook.Worksheets.Item[sheetName];
                range = sheet.Range[address];

                var actualNumberFormat = SafeStr(() => Convert.ToString(range.NumberFormat));
                // M16 全量审查：Font/Interior/Borders 显式捕获到变量、统一在 finally 释放。
                // 原实现声明了变量却从不赋值（死释放），实际每次 range.Font/... 内联访问都在泄漏子 RCW。
                font = range.Font;
                interior = range.Interior;
                borders = range.Borders;
                var actualBold = SafeBool(() => Convert.ToBoolean(font.Bold));
                var actualItalic = SafeBool(() => Convert.ToBoolean(font.Italic));
                var actualFontName = SafeStr(() => Convert.ToString(font.Name));
                var actualFontSize = SafeInt(() => Convert.ToInt32(font.Size));
                var actualFontColor = SafeInt(() => Convert.ToInt32(font.Color));
                var actualBackColor = SafeInt(() => Convert.ToInt32(interior.Color));
                var actualHAlign = SafeInt(() => Convert.ToInt32(range.HorizontalAlignment));
                var actualVAlign = SafeInt(() => Convert.ToInt32(range.VerticalAlignment));
                var actualBorderStyle = SafeInt(() => Convert.ToInt32(borders.LineStyle));
                var actualLocked = SafeBool(() => Convert.ToBoolean(range.Locked));

                var (verdict, issues, selfcheck) = ExcelVerifyJudgments.ComputeCellFormatVerdict(
                    actualNumberFormat, actualBold, actualItalic, actualFontName, actualFontSize, actualFontColor, actualBackColor,
                    actualHAlign, actualVAlign, actualBorderStyle, actualLocked,
                    Optional(args, "expected_number_format"),
                    TryOptionalBool(args, "expected_bold"),
                    TryOptionalBool(args, "expected_italic"),
                    Optional(args, "expected_font_name"),
                    TryOptionalInt(args, "expected_font_size"),
                    Optional(args, "expected_font_color"),
                    Optional(args, "expected_back_color"),
                    Optional(args, "expected_h_align"),
                    Optional(args, "expected_v_align"),
                    Optional(args, "expected_border_style"),
                    TryOptionalBool(args, "expected_locked"));

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        address,
                        actual_number_format = actualNumberFormat,
                        actual_bold = actualBold,
                        actual_italic = actualItalic,
                        actual_font_name = actualFontName,
                        actual_font_size = actualFontSize,
                        actual_font_color = actualFontColor,
                        actual_back_color = actualBackColor,
                        actual_h_align = actualHAlign,
                        actual_v_align = actualVAlign,
                        actual_border_style = actualBorderStyle,
                        actual_locked = actualLocked,
                        verdict,
                        issues,
                        selfcheck,
                        note = "单元格格式自检结果：请结合 issues / selfcheck 判断是否达标。"
                    }
                };
            }
            catch (Exception ex)
            {
                return Invalid("VERIFY_CELL_FORMAT_FAILED", "单元格格式自检失败：" + ex.Message);
            }
            finally
            {
                Release((object)protection);
                Release((object)borders);
                Release((object)interior);
                Release((object)font);
                Release((object)range);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        // ===================== 条件格式自检 =====================

        public ToolResult VerifyConditionalFormat(Dictionary<string, object> args)
        {
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            dynamic conditions = null;
            dynamic fc = null;
            try
            {
                var sheetName = Optional(args, "sheet");
                var address = Required(args, "address");
                var ruleIndex = TryOptionalInt(args, "rule_index") ?? 1;
                workbook = application.ActiveWorkbook ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                sheet = string.IsNullOrWhiteSpace(sheetName) ? application.ActiveSheet : workbook.Worksheets.Item[sheetName];
                range = sheet.Range[address];
                conditions = range.FormatConditions;
                var count = SafeInt(() => Convert.ToInt32(conditions.Count)) ?? 0;

                if (count <= 0)
                {
                    var r0 = ExcelVerifyJudgments.ComputeConditionalFormatVerdict(
                        null, null, null, null, null, null,
                        Optional(args, "expected_rule"), Optional(args, "expected_formula1"), Optional(args, "expected_formula2"),
                        TryOptionalInt(args, "expected_operator"), Optional(args, "expected_fill_color"), Optional(args, "expected_font_color"));
                    return new ToolResult
                    {
                        Ok = true,
                        Data = new
                        {
                            rule_found = false,
                            rule_count = 0,
                            verdict = r0.verdict,
                            issues = r0.issues,
                            selfcheck = r0.selfcheck,
                            note = "目标区域没有任何条件格式规则，说明条件格式并未真正生成。"
                        }
                    };
                }

                if (ruleIndex < 1 || ruleIndex > count)
                {
                    return new ToolResult
                    {
                        Ok = true,
                        Data = new
                        {
                            rule_found = false,
                            rule_count = count,
                            rule_index = ruleIndex,
                            verdict = "rule",
                            issues = new List<string> { $"规则序号非法：第 {ruleIndex} 条超出目标区域规则总数（{count}），无法校验。" },
                            selfcheck = "目标区域仅含 " + count + " 条规则，请用正确的 rule_index（1.." + count + "）重试。",
                            note = "rule_index 越界，未读取任何规则。"
                        }
                    };
                }

                fc = conditions.Item[ruleIndex];
                var typeCode = SafeInt(() => Convert.ToInt32(fc.Type));
                var opCode = SafeInt(() => Convert.ToInt32(fc.Operator));
                var actualType = MapConditionalType(typeCode, opCode);
                var actualFormula1 = SafeStr(() => Convert.ToString(fc.Formula1));
                var actualFormula2 = SafeStr(() => Convert.ToString(fc.Formula2));
                var actualOperator = opCode;
                // M16 全量审查：fc.Interior/Font 子 RCW 显式捕获并释放（原实现每次泄漏 2 个）。
                // v3.8.1.76 审查加固：WPS 宿主下 FormatCondition 的 COM dynamic 分派可能不暴露
                // Interior/Font（RuntimeBinderException「未包含 Interior 的定义」，踩坑 §2.1/2.2
                // dynamic 丢运行时成员同族）。读取降级为可空：颜色读不到时同时跳过颜色断言并在
                // issues 明示人工核对项，规则类型/公式/运算符校验照常执行，不再让整个自检一票否决。
                dynamic fcInterior = null;
                dynamic fcFont = null;
                var colorCheckSkipped = false;
                int? actualFillColor;
                int? actualFontColor;
                try
                {
                    try { fcInterior = fc.Interior; }
                    catch (Exception xaiInterior) { AppLog.Warn("verify", "ExcelVerifyService.cs:VerifyConditionalFormat Interior 读取降级: " + xaiInterior.GetType().Name + " " + xaiInterior.Message); }
                    try { fcFont = fc.Font; }
                    catch (Exception xaiFont) { AppLog.Warn("verify", "ExcelVerifyService.cs:VerifyConditionalFormat Font 读取降级: " + xaiFont.GetType().Name + " " + xaiFont.Message); }
                    actualFillColor = fcInterior == null ? (int?)null : SafeInt(() => Convert.ToInt32(fcInterior.Color));
                    actualFontColor = fcFont == null ? (int?)null : SafeInt(() => Convert.ToInt32(fcFont.Color));
                    colorCheckSkipped = fcInterior == null || fcFont == null
                        || actualFillColor == null || actualFontColor == null;
                }
                finally
                {
                    Release((object)fcInterior);
                    Release((object)fcFont);
                }

                var expectedFillColorArg = Optional(args, "expected_fill_color");
                var expectedFontColorArg = Optional(args, "expected_font_color");
                if (colorCheckSkipped)
                {
                    // 颜色读不到时不做颜色断言（expected 也置空跳过比对），避免把"读不到"误判为"颜色不符"。
                    expectedFillColorArg = null;
                    expectedFontColorArg = null;
                }

                var (verdict, issues, selfcheck) = ExcelVerifyJudgments.ComputeConditionalFormatVerdict(
                    actualType, actualFormula1, actualFormula2, actualOperator, actualFillColor, actualFontColor,
                    Optional(args, "expected_rule"), Optional(args, "expected_formula1"), Optional(args, "expected_formula2"),
                    TryOptionalInt(args, "expected_operator"), expectedFillColorArg, expectedFontColorArg);

                if (colorCheckSkipped)
                {
                    issues.Add("宿主 COM 未暴露 FormatCondition.Interior/Font（dynamic 分派限制，WPS 宿主已知形态），填充色/字体色读取与断言已跳过——请人工目视核对颜色。");
                    selfcheck = selfcheck + "；颜色断言因宿主 COM 分派限制被跳过。";
                }

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        rule_found = true,
                        rule_count = count,
                        rule_index = ruleIndex,
                        actual_rule_type = actualType,
                        actual_formula1 = actualFormula1,
                        actual_formula2 = actualFormula2,
                        actual_operator = actualOperator,
                        actual_fill_color = actualFillColor,
                        actual_font_color = actualFontColor,
                        color_check_skipped = colorCheckSkipped,
                        verdict,
                        issues,
                        selfcheck,
                        note = "条件格式自检结果：请结合 issues / selfcheck 判断是否达标。"
                    }
                };
            }
            catch (Exception ex)
            {
                return Invalid("VERIFY_CONDITIONAL_FORMAT_FAILED", "条件格式自检失败：" + ex.Message);
            }
            finally
            {
                Release((object)fc);
                Release((object)conditions);
                Release((object)range);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        // ===================== 排序与筛选自检 =====================

        public ToolResult VerifySortFilter(Dictionary<string, object> args)
        {
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            try
            {
                var sheetName = Optional(args, "sheet");
                var address = Required(args, "address");
                var headerRows = TryOptionalInt(args, "header_rows") ?? 1;

                // 工具级 fail-closed（§4A）：排序、筛选两侧都没有任何期望 = 无判定依据。
                // 单侧缺期望只是"不检查那一半"，不构成失败（sort_range 只传排序参数是合法用法）。
                var hasSortExpectation = TryOptionalInt(args, "expected_sort_column").HasValue
                    && !string.IsNullOrWhiteSpace(Optional(args, "expected_sort_direction"));
                var hasFilterExpectation = TryOptionalInt(args, "expected_filter_field").HasValue
                    && !string.IsNullOrWhiteSpace(Optional(args, "expected_filter_criteria1"));
                if (!hasSortExpectation && !hasFilterExpectation)
                {
                    return new ToolResult
                    {
                        Ok = true,
                        Data = new
                        {
                            address,
                            verdict = "no_expected",
                            issues = new List<string>
                            {
                                "未提供排序或筛选任一期望值（expected_sort_column+expected_sort_direction，或 expected_filter_field+expected_filter_criteria1），未做断言。"
                            },
                            selfcheck = "期望值缺失：排序筛选断言未执行（fail-closed）。"
                        }
                    };
                }

                workbook = application.ActiveWorkbook ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                sheet = string.IsNullOrWhiteSpace(sheetName) ? application.ActiveSheet : workbook.Worksheets.Item[sheetName];
                range = sheet.Range[address];
                object[,] matrix = (object[,])ReadMatrix(range); // 显式还原静态类型，截断 dynamic 分派传染链

                // ---- 排序数据级判定：检查排序列身体行是否按预期方向单调（支持主/次排序列复合字典序） ----
                var sortVerdict = ExcelVerifyJudgments.ComputeSortVerdict(
                    matrix, headerRows,
                    TryOptionalInt(args, "expected_sort_column"),
                    Optional(args, "expected_sort_direction"),
                    TryOptionalInt(args, "expected_sort_column2"),
                    Optional(args, "expected_sort_direction2"));

                // ---- 筛选数据级判定：检查每个数据行是否都满足 (列, 运算符, 条件) 谓词 ----
                var filterVerdict = ExcelVerifyJudgments.ComputeFilterVerdict(
                    matrix, headerRows,
                    TryOptionalInt(args, "expected_filter_field"),
                    Optional(args, "expected_filter_criteria1"),
                    Optional(args, "expected_filter_operator"));

                var issues = new List<string>();
                issues.AddRange(sortVerdict.issues);
                issues.AddRange(filterVerdict.issues);
                var verdict = "ok";
                foreach (var part in (sortVerdict.verdict + "+" + filterVerdict.verdict).Split('+'))
                {
                    if (!string.Equals(part, "ok", StringComparison.Ordinal) && !string.IsNullOrWhiteSpace(part))
                    {
                        ExcelVerifyJudgments.AppendVerdictFlag(ref verdict, part);
                    }
                }

                var selfcheck = new System.Text.StringBuilder();
                selfcheck.Append(verdict == "ok" ? "自检通过" : "自检未通过")
                         .Append("：排序筛选（基于区域数据矩阵实际值验证）；")
                         .Append(sortVerdict.selfcheck).Append("；").Append(filterVerdict.selfcheck);
                if (issues.Count > 0)
                {
                    selfcheck.Append("；问题：").Append(string.Join("；", issues));
                }

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        address,
                        header_rows = headerRows,
                        verdict,
                        issues,
                        selfcheck,
                        note = "排序筛选自检结果：基于区域数据矩阵实际值验证，与内存 sort_range / filter_excel_data 工具语义一致。"
                    }
                };
            }
            catch (Exception ex)
            {
                return Invalid("VERIFY_SORT_FILTER_FAILED", "排序筛选自检失败：" + ex.Message);
            }
            finally
            {
                Release((object)range);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        // ===================== 公式自检 =====================

        public ToolResult VerifyFormula(Dictionary<string, object> args)
        {
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            try
            {
                var sheetName = Optional(args, "sheet");
                var address = Required(args, "address");
                workbook = application.ActiveWorkbook ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                sheet = string.IsNullOrWhiteSpace(sheetName) ? application.ActiveSheet : workbook.Worksheets.Item[sheetName];
                range = sheet.Range[address];

                var hasFormula = SafeBool(() => Convert.ToBoolean(range.HasFormula)) ?? false;
                var actualFormula = SafeStr(() => Convert.ToString(range.Formula));
                // 错误格 Value2 令互操作返回带 0x800A 高位的 Int32 错误码（如 -2146826281=0x800A07D7 → #DIV/0!），
                // 统一映射为 "#..." 文本后再做错误判定。
                string actualValue = ToCellText(SafeObj(() => range.Value2), range);

                var (verdict, issues, selfcheck) = ExcelVerifyJudgments.ComputeFormulaVerdict(
                    hasFormula, actualFormula, actualValue,
                    Optional(args, "expected_formula"), Optional(args, "expected_value"),
                    TryOptionalDouble(args, "tolerance"));

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        address,
                        has_formula = hasFormula,
                        actual_formula = actualFormula,
                        actual_value = actualValue,
                        verdict,
                        issues,
                        selfcheck,
                        note = "公式自检结果：请结合 issues / selfcheck 判断是否达标。"
                    }
                };
            }
            catch (Exception ex)
            {
                return Invalid("VERIFY_FORMULA_FAILED", "公式自检失败：" + ex.Message);
            }
            finally
            {
                Release((object)range);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        // ===================== 内部工具 =====================

        /// <summary>把 xlFormatConditionType 常量（结合 Operator）映射为规范规则名，供判据比对。</summary>
        private static string MapConditionalType(int? typeCode, int? opCode)
        {
            if (!typeCode.HasValue)
            {
                return null;
            }

            return typeCode.Value switch
            {
                4 => "data_bar",   // xlDatabar
                8 => "duplicates", // xlUniqueValues
                1 => opCode switch // xlCellValue：具体规则由运算符决定
                {
                    5 => "greater", // xlGreater
                    6 => "less",    // xlLess
                    1 => "between", // xlBetween
                    _ => "cell_value"
                },
                _ => typeCode.Value.ToString()
            };
        }

        private static string Optional(Dictionary<string, object> args, string key)
        {
            return args.TryGetValue(key, out var value) && value != null
                ? Convert.ToString(value)
                : null;
        }

        // ===================== 文件核验（M28） =====================

        /// <summary>verify_file：导出/生成文件核验——存在性 / 大小下限 / 魔数与扩展名匹配（纯文件层，不依赖 COM）。</summary>
        public ToolResult VerifyFile(Dictionary<string, object> args)
        {
            try
            {
                var rawPath = Required(args, "file_path").Trim();
                // M28 审查（审查者）：verify_file 不能成为任意路径探测面——File.Exists 对 UNC 会触发
                // SMB/NTLM 外连；设备路径/含 NUL 同理。只接受本地固定磁盘路径（先校验、后 IO）。
                var isUnc = rawPath.Length >= 2 && ((rawPath[0] == '\\' && rawPath[1] == '\\') || (rawPath[0] == '/' && rawPath[1] == '/'));
                var isDevice = rawPath.Length >= 4 && rawPath[0] == '\\' && rawPath[1] == '\\' && rawPath[2] == '.' && rawPath[3] == '\\';
                if (isUnc || isDevice || rawPath.IndexOf('\0') >= 0)
                {
                    return Invalid("PATH_NOT_ALLOWED",
                        "verify_file 仅支持本地固定磁盘路径，拒绝 UNC/网络或设备路径（防凭据外连/信息探测）：" + rawPath);
                }

                // M28 审查（第三轮）：路径口径与写侧 PathGuard 统一——相对路径按 server CWD 解析、
                // 拒绝 ..、UNC/设备、网络/非固定盘（映射盘 Z: 也按卷类型拒绝，防 SMB/NTLM 外连）。
                string filePath;
                try
                {
                    filePath = ExcelAiAssistant.Core.Security.PathGuard.GuardOutputPath(rawPath);
                }
                catch (Exception pathException)
                {
                    return Invalid("PATH_NOT_ALLOWED", "verify_file 路径不允许：" + pathException.Message);
                }

                // 最终目标校验：允许本地 junction/OneDrive 占位（最终仍在本机固定盘），
                // 拒绝链接指向 UNC/设备的绕过；探测失败显式报 PATH_PROBE_FAILED（不与“符号链接”混淆）。
                var finalKind = ResolveFinalPathKind(filePath);
                if (finalKind == "network")
                {
                    return Invalid("PATH_NOT_ALLOWED", "verify_file 目标最终落在网络/设备路径（拒绝 SMB/NTLM 外连）：" + filePath);
                }

                if (finalKind == "unknown")
                {
                    return Invalid("PATH_PROBE_FAILED", "verify_file 无法确认目标为本地固定盘（权限/IO 探测失败）：" + filePath);
                }

                var issues = new List<object>();
                var expectedExists = TryOptionalBool(args, "expected_exists") ?? true;
                if (System.IO.Directory.Exists(filePath))
                {
                    // M28 审查：目录不是文件——显式 not_a_file，不能按 missing/ok 混淆
                    issues.Add("路径是目录而不是文件：" + filePath);
                    return FileVerdict(filePath, false, -1, "directory", FileSignature.ExpectedFromExtension(filePath), issues, "exists=true;kind=directory", "not_a_file");
                }

                var exists = System.IO.File.Exists(filePath);
                var selfcheck = new System.Text.StringBuilder("exists=" + exists);
                if (!exists)
                {
                    if (expectedExists)
                    {
                        issues.Add("文件不存在：" + filePath);
                    }

                    return FileVerdict(filePath, false, -1, null, FileSignature.ExpectedFromExtension(filePath), issues, selfcheck.ToString());
                }

                if (!expectedExists)
                {
                    issues.Add("文件不应存在但实际存在：" + filePath);
                }

                var size = new System.IO.FileInfo(filePath).Length;
                selfcheck.Append(";size=" + size);
                var minBytes = TryOptionalInt(args, "min_bytes");
                if (minBytes.HasValue && size < minBytes.Value)
                {
                    issues.Add("文件大小不足：期望至少 " + minBytes.Value + " 字节，实际 " + size + " 字节");
                }

                var expectedMagicArg = Optional(args, "expected_magic");
                var expectedMagic = !string.IsNullOrWhiteSpace(expectedMagicArg) && !string.Equals(expectedMagicArg, "auto", StringComparison.OrdinalIgnoreCase)
                    ? expectedMagicArg
                    : FileSignature.ExpectedFromExtension(filePath);
                var detected = FileSignature.Detect(filePath);
                selfcheck.Append(";magic=" + detected);
                if (!FileSignature.Matches(expectedMagic, detected))
                {
                    issues.Add("魔数不符：期望 " + expectedMagic + "，实际 " + detected + "（扩展名与内容可能不一致）");
                }

                return FileVerdict(filePath, true, size, detected, expectedMagic, issues, selfcheck.ToString());
            }
            catch (Exception ex)
            {
                return Invalid("VERIFY_FILE_FAILED", "文件核验失败：" + ex.Message);
            }
        }

        [System.Runtime.InteropServices.DllImport("kernel32", CharSet = System.Runtime.InteropServices.CharSet.Unicode, SetLastError = true)]
        private static extern int GetFinalPathNameByHandle(IntPtr hFile, System.Text.StringBuilder lpszFilePath, int cchFilePath, int dwFlags);

        /// <summary>
        /// M28 审查（第三轮）：解析文件最终目标判定 local / network / unknown。
        /// 允许本地 junction 与 OneDrive 占位（最终仍在本机固定盘），拒绝链接指向 UNC/设备；
        /// 文件不存在时按所在卷 DriveType 判定（PathGuard 已先校验固定盘）。
        /// </summary>
        private static string ResolveFinalPathKind(string filePath)
        {
            try
            {
                if (!System.IO.File.Exists(filePath))
                {
                    return "local";
                }

                using (var stream = new System.IO.FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read,
                           System.IO.FileShare.ReadWrite | System.IO.FileShare.Delete))
                {
                    var buffer = new System.Text.StringBuilder(1024);
                    var length = GetFinalPathNameByHandle(stream.SafeFileHandle.DangerousGetHandle(), buffer, buffer.Capacity, 0);
                    if (length <= 0)
                    {
                        return "unknown";
                    }

                    var final = buffer.ToString();
                    if (final.StartsWith("\\\\?\\UNC\\", StringComparison.OrdinalIgnoreCase))
                    {
                        return "network";
                    }

                    if (final.StartsWith("\\\\?\\", StringComparison.Ordinal))
                    {
                        final = final.Substring(4);
                    }

                    if (final.StartsWith("\\\\", StringComparison.Ordinal))
                    {
                        return "network";
                    }

                    var root = System.IO.Path.GetPathRoot(final);
                    return !string.IsNullOrWhiteSpace(root) && new System.IO.DriveInfo(root).DriveType == System.IO.DriveType.Fixed
                        ? "local"
                        : "network";
                }
            }
            catch
            {
                return "unknown";
            }
        }

        private static ToolResult FileVerdict(string filePath, bool exists, long sizeBytes, string magic, string expectedMagic, List<object> issues, string selfcheck, string verdictOverride = null)
        {
            var verdict = !string.IsNullOrEmpty(verdictOverride)
                ? verdictOverride
                : (issues.Count == 0 ? "ok" : (exists ? "mismatch" : "missing"));
            return new ToolResult
            {
                // M28 审查（AI 使用方）：断言型自检的 Ok 必须表达「判定通过」——否则自动自检只看到
                // Ok=true 会把 missing/mismatch 记成成功（导出后自检实际永不失败）。判定详情看 verdict。
                Ok = verdict == "ok",
                Data = new
                {
                    file_path = filePath,
                    exists,
                    size_bytes = sizeBytes,
                    magic,
                    expected_magic = expectedMagic,
                    verdict,
                    issues,
                    selfcheck,
                    // AI 使用方口径：ok=判定通过（非仅执行成功），失败时看 verdict/issues 定位
                    note = verdict == "ok" ? "文件核验通过。" : ("文件核验未通过（verdict=" + verdict + "）：" + string.Join("；", issues))
                }
            };
        }

        private static string Required(Dictionary<string, object> args, string key)
        {
            var value = Optional(args, key);
            return string.IsNullOrWhiteSpace(value)
                ? throw new ArgumentException($"缺少必填参数：{key}")
                : value;
        }

        private static bool? TryOptionalBool(Dictionary<string, object> args, string key)
        {
            if (args != null && args.TryGetValue(key, out var raw) && raw != null)
            {
                if (raw is bool b)
                {
                    return b;
                }

                if (bool.TryParse(Convert.ToString(raw), out var parsed))
                {
                    return parsed;
                }
            }

            return null;
        }

        private static int? TryOptionalInt(Dictionary<string, object> args, string key)
        {
            if (args != null && args.TryGetValue(key, out var raw) && raw != null)
            {
                if (int.TryParse(Convert.ToString(raw), NumberStyles.Integer, CultureInfo.InvariantCulture, out var v))
                {
                    return v;
                }
            }

            return null;
        }

        private static double? TryOptionalDouble(Dictionary<string, object> args, string key)
        {
            if (args != null && args.TryGetValue(key, out var raw) && raw != null)
            {
                if (double.TryParse(Convert.ToString(raw), NumberStyles.Any, CultureInfo.InvariantCulture, out var v))
                {
                    return v;
                }
            }

            return null;
        }

        private static string SafeStr(Func<object> getter)
        {
            try
            {
                var v = getter();
                return v == null ? null : Convert.ToString(v);
            }
            catch (Exception) {
                AppLog.Warn("verify", "ExcelVerifyService.cs:SafeStr 静默兜底触发: ");
                return null;
            }
        }

        /// <summary>Excel xlErr 常量 → 错误显示文本（低 16 位为 err code）。</summary>
        private static readonly Dictionary<int, string> ExcelErrText = new Dictionary<int, string>
        {
            [2000] = "#NULL!",
            [2007] = "#DIV/0!",
            [2015] = "#VALUE!",
            [2023] = "#REF!",
            [2029] = "#NAME?",
            [2036] = "#NUM!",
            [2042] = "#N/A"
        };

        /// <summary>把单元格原始值统一为可比较文本：字符串原样；COM 错误 Int32（0x800A 高位）映射为 "#... " 错误文本；COM 对象回退 range.Text。</summary>
        private static string ToCellText(object raw, dynamic range)
        {
            if (raw == null)
            {
                return null;
            }

            if (raw is string text)
            {
                return text;
            }

            if (raw is int i)
            {
                // .NET COM 互操作下错误格 Value2 返回带 0x800A 高位的 Int32（如 -2146826281 = 0x800A07D7）
                if (((uint)i >> 16) == 0x800A)
                {
                    return ExcelErrText.TryGetValue((int)((uint)i & 0xFFFF), out var t)
                        ? t
                        : "#ERR" + ((uint)i & 0xFFFF);
                }

                return i.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }

            if (raw.GetType().IsCOMObject)
            {
                return SafeStr(() => Convert.ToString(range.Text));
            }

            return Convert.ToString(raw);
        }

        private static object SafeObj(Func<object> getter)
        {
            try
            {
                return getter();
            }
            catch (Exception) {
                AppLog.Warn("verify", "ExcelVerifyService.cs:SafeObj 静默兜底触发: ");
                return null;
            }
        }

        private static bool? SafeBool(Func<object> getter)
        {
            try
            {
                var v = getter();
                return v == null ? (bool?)null : Convert.ToBoolean(v);
            }
            catch (Exception) {
                AppLog.Warn("verify", "ExcelVerifyService.cs:SafeBool 静默兜底触发: ");
                return null;
            }
        }

        private static int? SafeInt(Func<object> getter)
        {
            try
            {
                var v = getter();
                return v == null ? (int?)null : Convert.ToInt32(v);
            }
            catch (Exception) {
                AppLog.Warn("verify", "ExcelVerifyService.cs:SafeInt 静默兜底触发: ");
                return null;
            }
        }

        private static double? SafeDouble(Func<object> getter)
        {
            try
            {
                var v = getter();
                return v == null ? (double?)null : Convert.ToDouble(v);
            }
            catch (Exception) {
                AppLog.Warn("verify", "ExcelVerifyService.cs:SafeDouble 静默兜底触发: ");
                return null;
            }
        }

        private static string SafeText(Func<object> getter)
        {
            try
            {
                var v = getter();
                return v == null ? null : Convert.ToString(v);
            }
            catch (Exception) {
                AppLog.Warn("verify", "ExcelVerifyService.cs:SafeText 静默兜底触发: ");
                return null;
            }
        }

        /// <summary>整块读取区域 Value2 为 0 基 object[,]（COM 返回下界为 1 的数组需归一化，单格退化为一元）。</summary>
        private static object[,] ReadMatrix(dynamic range)
        {
            object raw;
            try
            {
                raw = range.Value2;
            }
            catch (Exception) {
                AppLog.Warn("verify", "ExcelVerifyService.cs:ReadMatrix 静默兜底触发: ");
                return new object[0, 0];
            }

            if (raw == null)
            {
                return new object[0, 0];
            }

            var arr = raw as Array;
            if (arr != null && arr.Rank == 2)
            {
                var r0 = arr.GetLowerBound(0);
                var c0 = arr.GetLowerBound(1);
                var rows = arr.GetLength(0);
                var cols = arr.GetLength(1);
                var result = new object[rows, cols];
                for (var r = 0; r < rows; r++)
                {
                    for (var c = 0; c < cols; c++)
                    {
                        result[r, c] = arr.GetValue(r + r0, c + c0);
                    }
                }

                return result;
            }

            return new object[,] { { raw } };
        }

        private static ToolResult Invalid(string code, string message)
        {
            return new ToolResult { Ok = false, ErrorCode = code, Data = message };
        }

        // ===================== 写入结果自检 =====================

        public ToolResult VerifyWriteResult(Dictionary<string, object> args)
        {
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            try
            {
                var sheetName = Optional(args, "sheet");
                var address = Required(args, "address");
                workbook = application.ActiveWorkbook ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                sheet = string.IsNullOrWhiteSpace(sheetName) ? application.ActiveSheet : workbook.Worksheets.Item[sheetName];
                range = sheet.Range[address];

                object[,] actual = (object[,])ReadMatrix(range); // 显式还原静态类型，截断 dynamic 分派传染链
                var expected = ReadExpectedMatrix(args, "expected_values");
                var expectedCount = TryOptionalInt(args, "expected_count");
                var tolerance = TryOptionalDouble(args, "tolerance");
                var typeAware = TryOptionalBool(args, "type_aware") ?? false;
                var (verdict, issues, selfcheck) = ExcelVerifyJudgments.ComputeWriteResultVerdict(
                    actual, expected, expectedCount, tolerance, address, typeAware);

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        address,
                        actual_count = CountNonBlank(actual),
                        verdict,
                        issues,
                        selfcheck,
                        note = "写入结果自检结果：请结合 issues / selfcheck 判断是否达标。"
                    }
                };
            }
            catch (Exception ex)
            {
                return Invalid("VERIFY_WRITE_RESULT_FAILED", "写入结果自检失败：" + ex.Message);
            }
            finally
            {
                Release((object)range);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        // ===================== 数据清洗自检 =====================

        public ToolResult VerifyDataClean(Dictionary<string, object> args)
        {
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            try
            {
                var sheetName = Optional(args, "sheet");
                var address = Required(args, "address");
                workbook = application.ActiveWorkbook ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                sheet = string.IsNullOrWhiteSpace(sheetName) ? application.ActiveSheet : workbook.Worksheets.Item[sheetName];
                range = sheet.Range[address];

                var headerRows = TryOptionalInt(args, "header_rows") ?? 1;
                object[,] actual = (object[,])ReadMatrix(range); // 显式还原静态类型，截断 dynamic 分派传染链
                var checks = ReadCleanChecks(args);
                var (verdict, issues, selfcheck) = ExcelVerifyJudgments.ComputeDataCleanVerdict(
                    actual, checks, headerRows, address);

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        address,
                        header_rows = headerRows,
                        verdict,
                        issues,
                        selfcheck,
                        note = "数据清洗自检结果：请结合 issues / selfcheck 判断是否达标。"
                    }
                };
            }
            catch (Exception ex)
            {
                return Invalid("VERIFY_DATA_CLEAN_FAILED", "数据清洗自检失败：" + ex.Message);
            }
            finally
            {
                Release((object)range);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        // ===================== 透视表自检 =====================

        public ToolResult VerifyPivot(Dictionary<string, object> args)
        {
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            dynamic region = null;
            try
            {
                var sheetName = Optional(args, "sheet");
                var targetAddress = Required(args, "target_address");
                workbook = application.ActiveWorkbook ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                sheet = string.IsNullOrWhiteSpace(sheetName) ? application.ActiveSheet : workbook.Worksheets.Item[sheetName];
                range = sheet.Range[targetAddress];

                // target_address 允许只给左上角单格：自动扩展为连续汇总块（表头 + 数据行），
                // 使模板可用 {stepN.target_address} 校验整个透视表，而不是只读单格。
                var areaRows = SafeInt(() => Convert.ToInt32(range.Rows.Count)) ?? 1;
                var areaCols = SafeInt(() => Convert.ToInt32(range.Columns.Count)) ?? 1;
                dynamic readRange = range;
                if (areaRows == 1 && areaCols == 1)
                {
                    try
                    {
                        region = range.CurrentRegion;
                        readRange = region;
                    }
                    catch (Exception) {
                        AppLog.Warn("verify", "ExcelVerifyService.cs:VerifyPivot 静默兜底触发: ");
                        readRange = range;
                    }
                }

                object[,] actual = (object[,])ReadMatrix(readRange); // 显式还原静态类型，截断 dynamic 分派传染链
                var expectedRows = ReadPivotRows(args);
                var expectedCount = TryOptionalInt(args, "expected_count");
                var tolerance = TryOptionalDouble(args, "tolerance");

                // 数量守恒对账（可选）：提供 source_address/source_value_column 时，
                // 校验明细列合计 == 汇总块 Total 列合计，拦截透视漏行/汇总块被截断。
                double? sourceSum = null;
                double? targetSum = null;
                var sourceAddress = Optional(args, "source_address");
                var sourceValueColumn = TryOptionalInt(args, "source_value_column");
                dynamic sourceRange = null;
                if (!string.IsNullOrWhiteSpace(sourceAddress) && sourceValueColumn.HasValue)
                {
                    try
                    {
                        sourceRange = sheet.Range[sourceAddress];
                        object[,] sourceMatrix = (object[,])ReadMatrix(sourceRange);
                        sourceSum = SumNumericColumn(sourceMatrix, sourceValueColumn.Value, TryOptionalInt(args, "source_label_column") ?? 1);
                        targetSum = SumNumericColumn(actual, 2, 1);
                    }
                    finally
                    {
                        Release((object)sourceRange);
                    }
                }

                var (verdict, issues, selfcheck) = ExcelVerifyJudgments.ComputePivotVerdict(
                    actual, expectedRows, expectedCount, tolerance, targetAddress, sourceSum, targetSum);

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        target_address = targetAddress,
                        data_rows = actual.GetLength(0) > 0 ? actual.GetLength(0) - 1 : 0,
                        source_sum = sourceSum,
                        target_sum = targetSum,
                        verdict,
                        issues,
                        selfcheck,
                        note = "透视表自检结果：请结合 issues / selfcheck 判断是否达标。"
                    }
                };
            }
            catch (Exception ex)
            {
                return Invalid("VERIFY_PIVOT_FAILED", "透视表自检失败：" + ex.Message);
            }
            finally
            {
                Release((object)region);
                Release((object)range);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        // ===================== 标记自检 =====================

        public ToolResult VerifyMarker(Dictionary<string, object> args)
        {
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            try
            {
                var sheetName = Optional(args, "sheet");
                var address = Required(args, "address");
                workbook = application.ActiveWorkbook ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                sheet = string.IsNullOrWhiteSpace(sheetName) ? application.ActiveSheet : workbook.Worksheets.Item[sheetName];
                range = sheet.Range[address];

                var areaRows = SafeInt(() => Convert.ToInt32(range.Rows.Count)) ?? 0;
                var areaCols = SafeInt(() => Convert.ToInt32(range.Columns.Count)) ?? 0;
                if ((long)areaRows * areaCols > 20000)
                {
                    return Invalid("VERIFY_MARKER_TOO_LARGE", "标记自检区域过大（超过 20000 格），请缩小 address 范围。");
                }

                var marked = new List<ExcelVerifyJudgments.MarkedCell>();
                for (var r = 1; r <= areaRows; r++)
                {
                    for (var c = 1; c <= areaCols; c++)
                    {
                        dynamic cell = null;
                        dynamic cellInterior = null;
                        try
                        {
                            cell = range.Cells.Item[r, c];
                            // M16 全量审查：cell.Interior 子 RCW 显式捕获释放（原实现大区域最坏泄漏 2 万个）。
                            cellInterior = cell.Interior;
                            var color = SafeInt(() => Convert.ToInt32(cellInterior.Color)) ?? 16777215;
                            if (color != 16777215 && color != -4142)
                            {
                                marked.Add(new ExcelVerifyJudgments.MarkedCell
                                {
                                    Address = SafeStr(() => Convert.ToString(cell.Address)),
                                    Color = (uint)color
                                });
                            }
                        }
                        catch (Exception) {
                            AppLog.Warn("verify", "ExcelVerifyService.cs:VerifyMarker 静默兜底触发: ");
                            // 单格读取失败跳过
                        }
                        finally
                        {
                            Release((object)cellInterior);
                            Release((object)cell);
                        }
                    }
                }

                var expectedCount = TryOptionalInt(args, "expected_count");
                var expectedFillColors = ReadColorMap(args, "expected_fill_colors");
                var expectedMarked = ReadStringList(args, "expected_marked");
                var (verdict, issues, selfcheck) = ExcelVerifyJudgments.ComputeMarkerVerdict(
                    marked, expectedCount, expectedFillColors, expectedMarked, address);

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        address,
                        marked_count = marked.Count,
                        verdict,
                        issues,
                        selfcheck,
                        note = "标记自检结果：请结合 issues / selfcheck 判断是否达标。"
                    }
                };
            }
            catch (Exception ex)
            {
                return Invalid("VERIFY_MARKER_FAILED", "标记自检失败：" + ex.Message);
            }
            finally
            {
                Release((object)range);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        // ===================== 结构自检 =====================

        /// <summary>结构自检：读取工作簿工作表名单与可选区域维度，与期望（存在/不存在/行列数）比对。
        /// 用于新增/重命名/删除工作表、结构变更后的回读核对，只读不修改。</summary>
        public ToolResult VerifyStructure(Dictionary<string, object> args)
        {
            dynamic workbook = null;
            dynamic sheets = null;
            dynamic sheet = null;
            dynamic range = null;
            try
            {
                workbook = application.ActiveWorkbook ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                sheets = workbook.Worksheets;
                var count = SafeInt(() => Convert.ToInt32(sheets.Count)) ?? 0;
                var names = new List<string>();
                for (var index = 1; index <= count; index++)
                {
                    dynamic item = null;
                    try
                    {
                        item = sheets.Item[index];
                        names.Add(Convert.ToString(item.Name));
                    }
                    finally
                    {
                        Release((object)item);
                    }
                }

                int? actualRows = null;
                int? actualCols = null;
                var address = Optional(args, "address");
                if (!string.IsNullOrWhiteSpace(address))
                {
                    var sheetName = Optional(args, "sheet");
                    sheet = string.IsNullOrWhiteSpace(sheetName) ? application.ActiveSheet : workbook.Worksheets.Item[sheetName];
                    range = sheet.Range[address];
                    actualRows = SafeInt(() => Convert.ToInt32(range.Rows.Count));
                    actualCols = SafeInt(() => Convert.ToInt32(range.Columns.Count));
                }

                var expectedExists = Optional(args, "expected_sheet_exists");
                var expectedAbsent = Optional(args, "expected_sheet_absent");
                var expectedRowCount = TryOptionalInt(args, "expected_row_count");
                var expectedColumnCount = TryOptionalInt(args, "expected_column_count");
                var (verdict, issues, selfcheck) = ExcelVerifyJudgments.ComputeStructureVerdict(
                    names, expectedExists, expectedAbsent, expectedRowCount, expectedColumnCount,
                    actualRows, actualCols, address);

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        sheet_names = names,
                        actual_rows = actualRows,
                        actual_columns = actualCols,
                        verdict,
                        issues,
                        selfcheck,
                        note = "结构自检结果：请结合 issues / selfcheck 判断是否达标。"
                    }
                };
            }
            catch (Exception ex)
            {
                return Invalid("VERIFY_STRUCTURE_FAILED", "结构自检失败：" + ex.Message);
            }
            finally
            {
                Release((object)range);
                Release((object)sheet);
                Release((object)sheets);
                Release((object)workbook);
            }
        }

        // ===================== 智能表自检 =====================

        public ToolResult VerifyTable(Dictionary<string, object> args)
        {
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic tables = null;
            dynamic headerRange = null;
            var snapshots = new List<ExcelVerifyJudgments.TableInfo>();
            try
            {
                workbook = application.ActiveWorkbook ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                var sheetName = Optional(args, "sheet");
                sheet = string.IsNullOrWhiteSpace(sheetName) ? application.ActiveSheet : workbook.Worksheets.Item[sheetName];
                tables = sheet.ListObjects;
                var count = SafeInt(() => Convert.ToInt32(tables.Count)) ?? 0;
                for (var index = 1; index <= count; index++)
                {
                    dynamic item = null;
                    dynamic tableRange = null;
                    try
                    {
                        item = tables.Item[index];
                        tableRange = item.Range; // M16：子 Range 捕获一次并复用，finally 释放
                        var info = new ExcelVerifyJudgments.TableInfo
                        {
                            Name = Convert.ToString(item.Name),
                            Address = Convert.ToString(tableRange.Address),
                            Rows = SafeInt(() => Convert.ToInt32(tableRange.Rows.Count)) ?? 0,
                            Columns = SafeInt(() => Convert.ToInt32(tableRange.Columns.Count)) ?? 0
                        };
                        try
                        {
                            headerRange = item.HeaderRowRange;
                            if (headerRange != null)
                            {
                                var values = headerRange.Value2;
                                if (values is object[,] matrix)
                                {
                                    var headers = new List<string>();
                                    for (var c = 1; c <= matrix.GetLength(1); c++)
                                    {
                                        headers.Add(Convert.ToString(matrix[1, c]));
                                    }

                                    info.Headers = headers.ToArray();
                                }
                                else if ((object)values != null)
                                {
                                    info.Headers = new[] { Convert.ToString((object)values) };
                                }
                            }
                        }
                        catch (Exception) {
                            AppLog.Warn("verify", "ExcelVerifyService.cs:method? 静默兜底触发: ");
                            // HeaderRowRange 在无表头表上可能抛出，回退为 null
                        }

                        snapshots.Add(info);
                    }
                    finally
                    {
                        Release((object)headerRange);
                        headerRange = null;
                        Release((object)tableRange);
                        Release((object)item);
                    }
                }

                var expectedName = Optional(args, "name");
                var address = Optional(args, "address");
                var expectedHeaders = ReadExpectedHeaders(args);
                var expectedRows = TryOptionalInt(args, "expected_row_count");
                var expectedColumns = TryOptionalInt(args, "expected_column_count");
                var (verdict, issues, selfcheck) = ExcelVerifyJudgments.ComputeTableVerdict(
                    expectedName, address, expectedHeaders, expectedRows, expectedColumns, snapshots);

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        table_count = snapshots.Count,
                        tables = snapshots,
                        verdict,
                        issues,
                        selfcheck,
                        note = "智能表自检结果：请结合 issues / selfcheck 判断是否达标。"
                    }
                };
            }
            catch (Exception ex)
            {
                return Invalid("VERIFY_TABLE_FAILED", "智能表自检失败：" + ex.Message);
            }
            finally
            {
                Release((object)headerRange);
                Release((object)tables);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        // ===================== 保护状态自检 =====================

        public ToolResult VerifyProtection(Dictionary<string, object> args)
        {
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            dynamic cell = null;
            dynamic protection = null;
            try
            {
                workbook = application.ActiveWorkbook ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                var sheetName = Optional(args, "sheet");
                sheet = string.IsNullOrWhiteSpace(sheetName) ? application.ActiveSheet : workbook.Worksheets.Item[sheetName];

                var info = new ExcelVerifyJudgments.ProtectionInfo
                {
                    ProtectStructure = SafeBool(() => Convert.ToBoolean(workbook.ProtectStructure)) ?? false,
                    SheetProtected = SafeBool(() => Convert.ToBoolean(sheet.ProtectContents)) ?? false
                };

                if (info.SheetProtected)
                {
                    protection = sheet.Protection;
                    info.AllowFiltering = SafeBool(() => Convert.ToBoolean(protection.AllowFiltering));
                    info.AllowSorting = SafeBool(() => Convert.ToBoolean(protection.AllowSorting));
                }

                var address = Optional(args, "address");
                if (!string.IsNullOrWhiteSpace(address))
                {
                    range = sheet.Range[address];
                    cell = range.Cells.Item[1, 1];
                    info.CellLocked = SafeBool(() => Convert.ToBoolean(cell.Locked));
                    info.CellAddress = address;
                }

                var expectedProtected = TryOptionalBool(args, "expected_protected");
                var expectedWritable = TryOptionalBool(args, "expected_writable");
                var (verdict, issues, selfcheck) = ExcelVerifyJudgments.ComputeProtectionVerdict(
                    info, expectedProtected, expectedWritable);

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        protect_structure = info.ProtectStructure,
                        sheet_protected = info.SheetProtected,
                        cell_locked = info.CellLocked,
                        cell_address = info.CellAddress,
                        allow_filtering = info.AllowFiltering,
                        allow_sorting = info.AllowSorting,
                        verdict,
                        issues,
                        selfcheck,
                        note = "保护状态自检结果：写入被拒（0x800A03EC）时优先核对 sheet_protected 与 cell_locked。"
                    }
                };
            }
            catch (Exception ex)
            {
                return Invalid("VERIFY_PROTECTION_FAILED", "保护状态自检失败：" + ex.Message);
            }
            finally
            {
                Release((object)protection);
                Release((object)cell);
                Release((object)range);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        /// <summary>
        /// 图片自检（只读）：枚举工作表 Shapes 中的图片（msoPicture=13 / msoLinkedPicture=11），
        /// 按 expected_count 核对数量、按 anchor 定位锚点单元格内的图片并核对尺寸。
        /// </summary>
        public ToolResult VerifyImage(Dictionary<string, object> args)
        {
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic anchorRange = null;
            dynamic shapes = null;
            dynamic item = null;
            try
            {
                workbook = application.ActiveWorkbook ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                var sheetName = Optional(args, "sheet");
                sheet = string.IsNullOrWhiteSpace(sheetName) ? application.ActiveSheet : workbook.Worksheets.Item[sheetName];

                var pictures = new List<ExcelVerifyJudgments.ImageInfo>();
                shapes = sheet.Shapes;
                var shapeCount = SafeInt(() => Convert.ToInt32(shapes.Count)) ?? 0;
                for (var index = 1; index <= shapeCount; index++)
                {
                    try
                    {
                        item = shapes.Item[index];
                        // msoPicture=13、msoLinkedPicture=11；图表/文本框等其余 Shape 类型不参与图片断言。
                        var type = SafeInt(() => Convert.ToInt32(item.Type));
                        if (type != 13 && type != 11)
                        {
                            continue;
                        }

                        pictures.Add(new ExcelVerifyJudgments.ImageInfo
                        {
                            Name = SafeText(() => Convert.ToString(item.Name)),
                            Left = SafeInt(() => Convert.ToInt32(item.Left)) ?? 0,
                            Top = SafeInt(() => Convert.ToInt32(item.Top)) ?? 0,
                            Width = SafeInt(() => Convert.ToInt32(item.Width)) ?? 0,
                            Height = SafeInt(() => Convert.ToInt32(item.Height)) ?? 0
                        });
                    }
                    finally
                    {
                        Release((object)item);
                        item = null;
                    }
                }

                ExcelVerifyJudgments.CellBounds anchorBounds = null;
                var anchor = Optional(args, "anchor");
                if (!string.IsNullOrWhiteSpace(anchor))
                {
                    anchorRange = sheet.Range[anchor];
                    var left = SafeDouble(() => Convert.ToDouble(anchorRange.Left));
                    var top = SafeDouble(() => Convert.ToDouble(anchorRange.Top));
                    var width = SafeDouble(() => Convert.ToDouble(anchorRange.Width));
                    var height = SafeDouble(() => Convert.ToDouble(anchorRange.Height));
                    if (left.HasValue && top.HasValue && width.HasValue && height.HasValue)
                    {
                        anchorBounds = new ExcelVerifyJudgments.CellBounds
                        {
                            Left = left.Value,
                            Top = top.Value,
                            Right = left.Value + width.Value,
                            Bottom = top.Value + height.Value
                        };
                    }
                }

                var expectedCount = TryOptionalInt(args, "expected_count");
                var expectedWidth = TryOptionalDouble(args, "expected_width");
                var expectedHeight = TryOptionalDouble(args, "expected_height");
                var (verdict, issues, selfcheck) = ExcelVerifyJudgments.ComputeImageVerdict(
                    expectedCount, anchor, anchorBounds, expectedWidth, expectedHeight, pictures);

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        shape_count_total = shapeCount,
                        picture_count = pictures.Count,
                        pictures,
                        verdict,
                        issues,
                        selfcheck
                    }
                };
            }
            catch (Exception ex)
            {
                return Invalid("VERIFY_IMAGE_FAILED", "图片自检失败：" + ex.Message);
            }
            finally
            {
                Release((object)item);
                Release((object)shapes);
                Release((object)anchorRange);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        /// <summary>
        /// 冻结窗格自检（只读）：冻结状态按工作表视图存储，读取前先激活目标工作表再读活动窗口的
        /// FreezePanes/SplitRow/SplitColumn，按 expected_frozen / expected_cell 判定。
        /// </summary>
        public ToolResult VerifyFreezePanes(Dictionary<string, object> args)
        {
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic window = null;
            try
            {
                workbook = application.ActiveWorkbook ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                var sheetName = Optional(args, "sheet");
                sheet = string.IsNullOrWhiteSpace(sheetName) ? application.ActiveSheet : workbook.Worksheets.Item[sheetName];
                sheet.Activate();
                window = application.ActiveWindow;

                var info = new ExcelVerifyJudgments.FreezePaneInfo
                {
                    Frozen = SafeBool(() => Convert.ToBoolean(window.FreezePanes)) ?? false,
                    SplitRow = SafeInt(() => Convert.ToInt32(window.SplitRow)) ?? 0,
                    SplitColumn = SafeInt(() => Convert.ToInt32(window.SplitColumn)) ?? 0
                };

                var expectedFrozen = TryOptionalBool(args, "expected_frozen");
                var expectedCell = Optional(args, "expected_cell");
                var (verdict, issues, selfcheck) = ExcelVerifyJudgments.ComputeFreezeVerdict(
                    expectedFrozen, expectedCell, info);

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        frozen = info.Frozen,
                        split_row = info.SplitRow,
                        split_column = info.SplitColumn,
                        expected_frozen = expectedFrozen,
                        expected_cell = expectedCell,
                        verdict,
                        issues,
                        selfcheck
                    }
                };
            }
            catch (Exception ex)
            {
                return Invalid("VERIFY_FREEZE_FAILED", "冻结窗格自检失败：" + ex.Message);
            }
            finally
            {
                Release((object)window);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        /// <summary>
        /// 超链接自检（只读）：枚举工作表 Hyperlinks（外部 Address / 内部 SubAddress），锚点单元格
        /// 归一化为无 $ 的 A1 地址后，按 expected_count 核对整表总数、按 address 定位目标单元格/区域
        /// 并按 expected_url/expected_subaddress/expected_text 核对内容，或按 expected_present=false
        /// 断言目标已无链接。返回 verdict/issues/selfcheck。执行 manage_hyperlinks 后都应调用一次。
        /// </summary>
        public ToolResult VerifyHyperlink(Dictionary<string, object> args)
        {
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic hyperlinks = null;
            dynamic item = null;
            dynamic anchor = null;
            try
            {
                workbook = application.ActiveWorkbook ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                var sheetName = Optional(args, "sheet");
                sheet = string.IsNullOrWhiteSpace(sheetName) ? application.ActiveSheet : workbook.Worksheets.Item[sheetName];

                var links = new List<ExcelVerifyJudgments.HyperlinkInfo>();
                hyperlinks = sheet.Hyperlinks;
                var count = SafeInt(() => Convert.ToInt32(hyperlinks.Count)) ?? 0;
                for (var index = 1; index <= count; index++)
                {
                    try
                    {
                        item = hyperlinks.Item[index];
                        string cell;
                        try
                        {
                            anchor = item.Range;
                            cell = (SafeText(() => Convert.ToString(anchor.Address)) ?? string.Empty)
                                .Replace("$", string.Empty).Trim();
                        }
                        finally
                        {
                            Release((object)anchor);
                            anchor = null;
                        }

                        links.Add(new ExcelVerifyJudgments.HyperlinkInfo
                        {
                            Cell = cell,
                            Url = SafeText(() => Convert.ToString(item.Address)) ?? string.Empty,
                            SubAddress = SafeText(() => Convert.ToString(item.SubAddress)) ?? string.Empty,
                            Text = SafeText(() => Convert.ToString(item.TextToDisplay)) ?? string.Empty
                        });
                    }
                    finally
                    {
                        Release((object)item);
                        item = null;
                    }
                }

                var expectedCount = TryOptionalInt(args, "expected_count");
                var expectedPresent = TryOptionalBool(args, "expected_present");
                var expectedUrl = Optional(args, "expected_url");
                var expectedSubAddress = Optional(args, "expected_subaddress");
                var expectedText = Optional(args, "expected_text");
                var address = Optional(args, "address");
                var (verdict, issues, selfcheck) = ExcelVerifyJudgments.ComputeHyperlinkVerdict(
                    expectedCount, address, expectedPresent, expectedUrl, expectedSubAddress, expectedText, links);

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        hyperlink_count = links.Count,
                        links,
                        verdict,
                        issues,
                        selfcheck
                    }
                };
            }
            catch (Exception ex)
            {
                return Invalid("VERIFY_HYPERLINK_FAILED", "超链接自检失败：" + ex.Message);
            }
            finally
            {
                Release((object)anchor);
                Release((object)item);
                Release((object)hyperlinks);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        /// <summary>读取 expected_headers 字符串数组（支持 JSON 数组与进程内 string[]/object[]）。</summary>
        private static string[] ReadExpectedHeaders(Dictionary<string, object> args)
        {
            if (!args.TryGetValue("expected_headers", out var raw) || raw == null)
            {
                return null;
            }

            if (raw is Newtonsoft.Json.Linq.JArray ja)
            {
                return ja.Select(t => t.ToString()).ToArray();
            }

            if (raw is System.Collections.IEnumerable enumerable && !(raw is string))
            {
                return enumerable.Cast<object>().Select(t => Convert.ToString(t)).ToArray();
            }

            return new[] { Convert.ToString(raw) };
        }

        // ===================== 迷你图自检 =====================

        public ToolResult VerifySparkline(Dictionary<string, object> args)
        {
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic dest = null;
            dynamic groups = null;
            dynamic group = null;
            dynamic location = null;
            try
            {
                var destRange = Required(args, "dest_range");
                workbook = application.ActiveWorkbook ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                sheet = string.IsNullOrWhiteSpace(Optional(args, "sheet"))
                    ? application.ActiveSheet
                    : workbook.Worksheets.Item[Optional(args, "sheet")];
                dest = sheet.Range[destRange];
                var infos = new List<ExcelVerifyJudgments.SparklineInfo>();
                try
                {
                    groups = dest.SparklineGroups;
                    var count = SafeInt(() => Convert.ToInt32(groups.Count)) ?? 0;
                    for (var i = 1; i <= count; i++)
                    {
                        group = groups.Item[i];
                        int? type = null;
                        string source = null;
                        string locationAddress = null;
                        try
                        {
                            type = SafeInt(() => Convert.ToInt32(group.Type));
                        }
                        catch (Exception) {
                            AppLog.Warn("verify", "ExcelVerifyService.cs:VerifySparkline 静默兜底触发: ");
                            // WPS 个别版本 Type 读取失败时保持 null，判定层按未知类型处理
                        }

                        try
                        {
                            source = SafeStr(() => Convert.ToString(group.SourceData));
                        }
                        catch (Exception) {
                            AppLog.Warn("verify", "ExcelVerifyService.cs:VerifySparkline 静默兜底触发: ");
                            // 同上，缺失按 null 处理
                        }

                        try
                        {
                            location = group.Location;
                            locationAddress = location == null ? null : Convert.ToString(location.Address);
                        }
                        catch (Exception) {
                            AppLog.Warn("verify", "ExcelVerifyService.cs:VerifySparkline 静默兜底触发: ");
                            // 同上
                        }

                        infos.Add(new ExcelVerifyJudgments.SparklineInfo
                        {
                            Type = type,
                            Location = locationAddress,
                            SourceData = source
                        });
                        Release((object)group);
                        group = null;
                        Release((object)location);
                        location = null;
                    }
                }
                catch (Exception) {
                    AppLog.Warn("verify", "ExcelVerifyService.cs:method? 静默兜底触发: ");
                    // 目标区域无 SparklineGroups（未生成/宿主不支持）时按空集合判定
                }

                var (verdict, issues, selfcheck) = ExcelVerifyJudgments.ComputeSparklineVerdict(
                    infos, destRange, Optional(args, "expected_type"), Optional(args, "expected_data_source"));
                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        dest_range = destRange,
                        found = infos.Count,
                        verdict,
                        issues,
                        selfcheck,
                        note = "迷你图自检结果：请结合 issues / selfcheck 判断是否达标。"
                    }
                };
            }
            catch (Exception ex)
            {
                return Invalid("VERIFY_SPARKLINE_FAILED", "迷你图自检失败：" + ex.Message);
            }
            finally
            {
                Release((object)location);
                Release((object)group);
                Release((object)groups);
                Release((object)dest);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        // ===================== 数据验证自检 =====================

        public ToolResult VerifyDataValidation(Dictionary<string, object> args)
        {
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            dynamic validation = null;
            try
            {
                var address = Required(args, "address");
                workbook = application.ActiveWorkbook ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                sheet = string.IsNullOrWhiteSpace(Optional(args, "sheet"))
                    ? application.ActiveSheet
                    : workbook.Worksheets.Item[Optional(args, "sheet")];
                range = sheet.Range[address];

                int? type = null;
                int? op = null;
                string formula1 = null;
                string formula2 = null;
                validation = range.Validation;
                try { type = SafeInt(() => Convert.ToInt32(validation.Type)); } catch (Exception) {
                    AppLog.Warn("verify", "ExcelVerifyService.cs:VerifyDataValidation 静默兜底触发: "); }
                try { op = SafeInt(() => Convert.ToInt32(validation.Operator)); } catch (Exception) {
                    AppLog.Warn("verify", "ExcelVerifyService.cs:VerifyDataValidation 静默兜底触发: "); }
                try { formula1 = SafeStr(() => Convert.ToString(validation.Formula1)); } catch (Exception) {
                    AppLog.Warn("verify", "ExcelVerifyService.cs:VerifyDataValidation 静默兜底触发: "); }
                try { formula2 = SafeStr(() => Convert.ToString(validation.Formula2)); } catch (Exception) {
                    AppLog.Warn("verify", "ExcelVerifyService.cs:VerifyDataValidation 静默兜底触发: "); }

                var (verdict, issues, selfcheck) = ExcelVerifyJudgments.ComputeDataValidationVerdict(
                    new ExcelVerifyJudgments.DataValidationInfo
                    {
                        Type = type,
                        Operator = op,
                        Formula1 = formula1,
                        Formula2 = formula2
                    },
                    Optional(args, "expected_type"),
                    Optional(args, "expected_source"),
                    Optional(args, "expected_operator"),
                    Optional(args, "expected_value1"),
                    Optional(args, "expected_value2"));

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        address,
                        validation_type = type,
                        validation_operator = op,
                        formula1,
                        formula2,
                        verdict,
                        issues,
                        selfcheck,
                        note = "数据验证自检结果：请结合 issues / selfcheck 判断是否达标。"
                    }
                };
            }
            catch (Exception ex)
            {
                return Invalid("VERIFY_DATA_VALIDATION_FAILED", "数据验证自检失败：" + ex.Message);
            }
            finally
            {
                Release((object)validation);
                Release((object)range);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        // ===================== 合并状态自检 =====================

        public ToolResult VerifyMerge(Dictionary<string, object> args)
        {
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            try
            {
                var address = Required(args, "address");
                workbook = application.ActiveWorkbook ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                sheet = string.IsNullOrWhiteSpace(Optional(args, "sheet"))
                    ? application.ActiveSheet
                    : workbook.Worksheets.Item[Optional(args, "sheet")];
                range = sheet.Range[address];

                bool? mergeCells = null;
                try
                {
                    var raw = range.MergeCells;
                    mergeCells = raw is bool b ? b : (bool?)null; // COM 混合状态返回 null（DBNull）
                }
                catch (Exception) {
                    AppLog.Warn("verify", "ExcelVerifyService.cs:VerifyMerge 静默兜底触发: ");
                    mergeCells = null;
                }

                bool? expectedMerged = null;
                var expectedRaw = Optional(args, "expected_merged");
                if (!string.IsNullOrWhiteSpace(expectedRaw))
                {
                    expectedMerged = string.Equals(expectedRaw.Trim(), "true", StringComparison.OrdinalIgnoreCase)
                        || expectedRaw.Trim() == "1" || expectedRaw.Trim() == "是";
                }

                var (verdict, issues, selfcheck) = ExcelVerifyJudgments.ComputeMergeVerdict(
                    mergeCells, expectedMerged, address);

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        address,
                        merged = mergeCells,
                        expected_merged = expectedMerged,
                        verdict,
                        issues,
                        selfcheck,
                        note = "合并状态自检结果：merged=true 整区已合并 / false 无合并 / null 混合。"
                    }
                };
            }
            catch (Exception ex)
            {
                return Invalid("VERIFY_MERGE_FAILED", "合并自检失败：" + ex.Message);
            }
            finally
            {
                Release((object)range);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        /// <summary>列合计，口径与 PivotTableService.SumByRow 严格对齐：
        /// 标签按 IsBlank 语义判空（含 DBNull）；数值侧与 SumByRow.ToDouble 同样用 InvariantCulture
        /// 转换——文本型数字（如 "3"）会被透视计入，这里也必须计入，只有无法解析的文本才跳过，
        /// 保证「明细列合计 == 汇总列合计」的守恒对账不因口径差误报。</summary>
        private static double? SumNumericColumn(object[,] matrix, int valueColumn, int labelColumn)
        {
            if (matrix == null || valueColumn < 1 || valueColumn > matrix.GetLength(1)
                || labelColumn < 1 || labelColumn > matrix.GetLength(1))
            {
                return null;
            }

            double sum = 0;
            for (var r = 1; r < matrix.GetLength(0); r++)
            {
                var label = matrix[r, labelColumn - 1];
                if (Convert.ToString(label)?.Trim().Length == 0)
                {
                    continue;
                }

                try
                {
                    sum += Convert.ToDouble(matrix[r, valueColumn - 1], CultureInfo.InvariantCulture);
                }
                catch (Exception ex) when (ex is FormatException || ex is InvalidCastException || ex is OverflowException)
                {
                    // 与 SumByRow 一致：非数值格不计入合计（该行标签仍会进汇总，合计贡献 0）。
                }
            }

            return sum;
        }

        // ===================== 新自检的解析辅助 =====================

        /// <summary>把 expected_values 参数（二维数组/JArray/嵌套列表）统一为 object[,]；缺省返回 null。</summary>
        private static object[,] ReadExpectedMatrix(Dictionary<string, object> args, string key)
        {
            if (!args.TryGetValue(key, out var raw) || raw == null)
            {
                return null;
            }

            if (raw is object[,] arr)
            {
                return arr;
            }

            var rows = new List<object[]>();
            if (raw is JArray ja)
            {
                foreach (var r in ja)
                {
                    rows.Add(r is JArray inner
                        ? inner.Select(ToMatrixScalar).ToArray()
                        : new[] { ToMatrixScalar(r) });
                }
            }
            else if (raw is System.Collections.IEnumerable seq && !(raw is string))
            {
                foreach (var r in seq)
                {
                    rows.Add(r is System.Collections.IEnumerable inner && !(r is string)
                        ? inner.Cast<object>().ToArray()
                        : new[] { r });
                }
            }

            if (rows.Count == 0)
            {
                return null;
            }

            var cols = rows.Max(r => r.Length);
            var result = new object[rows.Count, cols];
            for (var i = 0; i < rows.Count; i++)
            {
                for (var j = 0; j < cols; j++)
                {
                    result[i, j] = j < rows[i].Length ? rows[i][j] : null;
                }
            }

            return result;
        }

        private static object ToMatrixScalar(JToken token)
        {
            if (token == null || token.Type == JTokenType.Null)
            {
                return null;
            }

            switch (token.Type)
            {
                case JTokenType.Integer:
                    return token.Value<long>();
                case JTokenType.Float:
                    return token.Value<double>();
                case JTokenType.Boolean:
                    return token.Value<bool>();
                default:
                    return token.ToString();
            }
        }

        /// <summary>把 checks 参数（JArray of {type, columns}）解析为判据检查项列表；缺省返回 null。</summary>
        private static List<ExcelVerifyJudgments.CleanCheck> ReadCleanChecks(Dictionary<string, object> args)
        {
            if (!args.TryGetValue("checks", out var raw) || raw == null)
            {
                return null;
            }

            var list = new List<ExcelVerifyJudgments.CleanCheck>();
            if (raw is JArray ja)
            {
                foreach (var item in ja)
                {
                    var obj = item as JObject;
                    if (obj == null)
                    {
                        continue;
                    }

                    var columns = obj["columns"] as JArray;
                    list.Add(new ExcelVerifyJudgments.CleanCheck
                    {
                        Type = obj["type"]?.ToString(),
                        Columns = columns?.Select(t => t.Value<int>()).ToArray(),
                        Pattern = obj["pattern"]?.ToString()
                    });
                }
            }
            else if (raw is System.Collections.IEnumerable sequence && !(raw is string))
            {
                // 进程内调用方（harness/巡检/策略单测）直接传 .NET 集合：逐项归一为字典，
                // 与 JSON(JArray) 路径同口径——真机 full 曾因此把 checks 判成 no_checks（M11 闭环修复）。
                foreach (var item in sequence)
                {
                    Dictionary<string, object> dict;
                    try
                    {
                        dict = item as Dictionary<string, object>
                            ?? JObject.FromObject(item).ToObject<Dictionary<string, object>>();
                    }
                    catch (ArgumentException)
                    {
                        continue;
                    }

                    if (dict == null)
                    {
                        continue;
                    }

                    list.Add(new ExcelVerifyJudgments.CleanCheck
                    {
                        Type = dict.TryGetValue("type", out var type) ? type?.ToString() : null,
                        Columns = ParseColumns(dict.TryGetValue("columns", out var columns) ? columns : null),
                        Pattern = dict.TryGetValue("pattern", out var pattern) ? pattern?.ToString() : null
                    });
                }
            }

            return list.Count > 0 ? list : null;
        }

        /// <summary>把列号集合（JArray / IEnumerable）统一解析为 int[]；非法项跳过。</summary>
        private static int[] ParseColumns(object raw)
        {
            if (raw == null || raw is string)
            {
                return null;
            }

            if (raw is JArray ja)
            {
                return ja.Select(t => t.Value<int>()).ToArray();
            }

            if (raw is System.Collections.IEnumerable sequence)
            {
                var result = new List<int>();
                foreach (var item in sequence)
                {
                    try { result.Add(Convert.ToInt32(item)); }
                    catch (Exception ex) when (ex is FormatException || ex is InvalidCastException)
                    {
                        // 非法列号跳过，不让单个脏项毁掉整个 checks
                    }
                }

                return result.Count > 0 ? result.ToArray() : null;
            }

            return null;
        }

        /// <summary>把 expected_rows 参数（JArray of {label, value}）解析为判据抽查行列表；缺省返回 null。</summary>
        private static List<ExcelVerifyJudgments.PivotExpectedRow> ReadPivotRows(Dictionary<string, object> args)
        {
            if (!args.TryGetValue("expected_rows", out var raw) || raw == null)
            {
                return null;
            }

            var list = new List<ExcelVerifyJudgments.PivotExpectedRow>();
            if (raw is JArray ja)
            {
                foreach (var item in ja)
                {
                    var obj = item as JObject;
                    if (obj == null)
                    {
                        continue;
                    }

                    var valueToken = obj["value"];
                    double? value = null;
                    if (valueToken != null && valueToken.Type != JTokenType.Null)
                    {
                        value = valueToken.Value<double>();
                    }

                    list.Add(new ExcelVerifyJudgments.PivotExpectedRow
                    {
                        Label = obj["label"]?.ToString(),
                        Value = value
                    });
                }
            }

            return list.Count > 0 ? list : null;
        }

        /// <summary>把 expected_fill_colors 参数（JObject type→"#RRGGBB"）解析为字典；缺省返回 null。</summary>
        private static Dictionary<string, string> ReadColorMap(Dictionary<string, object> args, string key)
        {
            if (!args.TryGetValue(key, out var raw) || raw == null)
            {
                return null;
            }

            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (raw is JObject jo)
            {
                foreach (var kv in jo)
                {
                    map[kv.Key] = kv.Value?.ToString();
                }
            }
            else if (raw is System.Collections.IDictionary dict)
            {
                // 兼容 Dictionary<string,object>/<string,string> 等进程内直接调用
                foreach (System.Collections.DictionaryEntry entry in dict)
                {
                    map[Convert.ToString(entry.Key)] = Convert.ToString(entry.Value);
                }
            }

            return map.Count > 0 ? map : null;
        }

        /// <summary>把 expected_marked 参数（字符串数组）解析为列表；缺省返回 null。</summary>
        private static List<string> ReadStringList(Dictionary<string, object> args, string key)
        {
            if (!args.TryGetValue(key, out var raw) || raw == null)
            {
                return null;
            }

            var list = new List<string>();
            if (raw is JArray ja)
            {
                foreach (var t in ja)
                {
                    var s = t?.ToString();
                    if (!string.IsNullOrWhiteSpace(s))
                    {
                        list.Add(s);
                    }
                }
            }
            else if (raw is string single)
            {
                if (!string.IsNullOrWhiteSpace(single))
                {
                    list.Add(single);
                }
            }
            else if (raw is System.Collections.IEnumerable seq)
            {
                // 兼容 List<object>/object[] 等进程内直接调用
                foreach (var item in seq)
                {
                    var s = Convert.ToString(item);
                    if (!string.IsNullOrWhiteSpace(s))
                    {
                        list.Add(s);
                    }
                }
            }

            return list.Count > 0 ? list : null;
        }

        private static int CountNonBlank(object[,] values)
        {
            var count = 0;
            if (values == null)
            {
                return 0;
            }

            for (var r = 0; r < values.GetLength(0); r++)
            {
                for (var c = 0; c < values.GetLength(1); c++)
                {
                    var v = values[r, c];
                    if (v != null && !(v is string s && string.IsNullOrWhiteSpace(s)))
                    {
                        count++;
                    }
                }
            }

            return count;
        }

        private static void Release(object value)
        {
            if (value != null && System.Runtime.InteropServices.Marshal.IsComObject(value))
            {
                try
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(value);
                }
                catch (Exception) {
                    AppLog.Warn("verify", "ExcelVerifyService.cs:Release 静默兜底触发: ");
                }
            }
        }
    }
}
