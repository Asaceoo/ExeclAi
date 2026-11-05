using System;
using ExcelAiAssistant.Core.Diagnostics;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Newtonsoft.Json;

namespace ExcelAiAssistant.Core.Workflow
{
    /// <summary>
    /// 流程模板持久化：本地 JSON（%LocalAppData%\XaiAssistant\workflows.json）保存用户流程，
    /// 内置流程库始终合并返回。重名时用户流程覆盖内置同名项。
    /// </summary>
    public sealed class WorkflowStore
    {
        private readonly string filePath;
        private readonly Lazy<List<WorkflowTemplate>> builtIn;

        public WorkflowStore(string filePath = null)
        {
            this.filePath = filePath ?? DefaultPath();
            builtIn = new Lazy<List<WorkflowTemplate>>(BuiltInWorkflows);
        }

        public static string DefaultPath()
        {
            var root = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "XaiAssistant");
            return Path.Combine(root, "workflows.json");
        }

        public IReadOnlyList<WorkflowTemplate> GetAll()
        {
            var built = builtIn.Value.ToList();
            var mine = LoadUserWorkflows().ToList();
            foreach (var item in mine)
            {
                var existing = built.FirstOrDefault(b => string.Equals(b.Name, item.Name, StringComparison.OrdinalIgnoreCase));
                if (existing != null)
                {
                    built.Remove(existing);
                }

                built.Add(item);
            }

            return built
                .OrderByDescending(item => item.IsBuiltIn)
                .ThenBy(item => item.IsBuiltIn ? item.Name : string.Empty)
                .ToList();
        }

        public bool TryGet(string name, out WorkflowTemplate template)
        {
            template = GetAll().FirstOrDefault(item =>
                string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase));
            return template != null;
        }

        public bool Save(WorkflowTemplate template)
        {
            if (template == null || string.IsNullOrWhiteSpace(template.Name))
            {
                return false;
            }

            template.Name = template.Name.Trim();
            if (builtIn.Value.Any(b => string.Equals(b.Name, template.Name, StringComparison.OrdinalIgnoreCase)))
            {
                // 内置流程不可被覆盖，改用新条目仅当名称不同；否则拒绝同名覆盖内置。
                return false;
            }

            template.IsBuiltIn = false;
            if (!template.CreatedAtUtc.ToUniversalTime().Equals(template.CreatedAtUtc))
            {
                template.CreatedAtUtc = DateTime.UtcNow;
            }

            var mine = LoadUserWorkflows();
            var existing = mine.FirstOrDefault(item =>
                string.Equals(item.Name, template.Name, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                mine.Remove(existing);
            }

            mine.Add(template);
            Persist(mine);
            return true;
        }

        public bool Delete(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return false;
            }

            if (builtIn.Value.Any(b => string.Equals(b.Name, name, StringComparison.OrdinalIgnoreCase)))
            {
                return false;
            }

            var mine = LoadUserWorkflows();
            var existing = mine.FirstOrDefault(item =>
                string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase));
            if (existing == null)
            {
                return false;
            }

            mine.Remove(existing);
            Persist(mine);
            return true;
        }

        private List<WorkflowTemplate> LoadUserWorkflows()
        {
            if (!File.Exists(filePath))
            {
                return new List<WorkflowTemplate>();
            }

            try
            {
                return JsonConvert.DeserializeObject<List<WorkflowTemplate>>(
                           File.ReadAllText(filePath, Encoding.UTF8)) ??
                       new List<WorkflowTemplate>();
            }
            catch (JsonException)
            {
                return new List<WorkflowTemplate>();
        }
        }

        private void Persist(List<WorkflowTemplate> templates)
        {
            var directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // 原子落盘：先写临时文件，再用 File.Replace 覆盖，崩溃时不会写坏主 JSON。
            var tmp = filePath + ".tmp." + Guid.NewGuid().ToString("N");
            File.WriteAllText(
                tmp,
                JsonConvert.SerializeObject(templates, Formatting.Indented),
                Encoding.UTF8);
            try
            {
                if (File.Exists(filePath))
                {
                    File.Replace(tmp, filePath, null);
                }
                else
                {
                    File.Move(tmp, filePath);
                }
            }
            finally
            {
                if (File.Exists(tmp))
                {
                    File.Delete(tmp);
                }
            }
        }

        /// <summary>开箱即用的内置流程库（≥3 条）。</summary>
        public static List<WorkflowTemplate> BuiltInWorkflows()
        {
            return new List<WorkflowTemplate>
            {
                new WorkflowTemplate
                {
                    Name = "数据体检并生成报告",
                    Description = "对当前选区做离线统计体检（分布、相关性、高亮异常），并生成数据分析报告。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_current_excel_context", Description = "读取当前选区上下文" },
                        new WorkflowStep { Tool = "analyze_range_statistics", Description = "区间基础统计" },
                        new WorkflowStep { Tool = "analyze_value_distribution", Description = "取值分布体检" },
                        new WorkflowStep { Tool = "analyze_column_correlation", Description = "列相关性体检" },
                        new WorkflowStep { Tool = "generate_analysis_report", Description = "生成数据分析报告" }
                    }
                },
                new WorkflowTemplate
                {
                    Name = "选区一键出图",
                    Description = "识别当前选区的数值结构，推荐并插入图表，导出图表图片。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_current_selection_values_and_formulas", Description = "读取选区数值" },
                        new WorkflowStep { Tool = "recommend_excel_chart", Description = "推荐合适的图表类型" },
                        new WorkflowStep { Tool = "create_excel_chart", Description = "在选区旁插入图表" },
                        // 清单约束兜底：自检必须对照“上一张实际插入的图表产物”断言类型/标题/数据源/图名，
                        // 而非空跑。期望值若未被替换为真实结果则 verify_chart 会自动跳过对应断言。
                        new WorkflowStep
                        {
                            Tool = "verify_chart",
                            Description = "自检图表是否真正生成/数据源、类型、标题是否正确且数据可见",
                            Arguments = new Dictionary<string, object>
                            {
                                { "chart_name", "{step2.chart_name}" },
                                { "expected_chart_type", "{step2.chart_type}" },
                                { "expected_title", "{step2.title}" },
                                { "expected_data_source", "{step2.source_address}" }
                            }
                        },
                        new WorkflowStep { Tool = "export_excel_chart_image", Description = "导出图表为图片" }
                    }
                },
                new WorkflowTemplate
                {
                    Name = "两工作表比对",
                    Description = "列出当前工作簿工作表，并对两张工作表做差异比对。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_excel_sheet_list", Description = "列出工作表清单" },
                        new WorkflowStep { Tool = "compare_two_sheets", Description = "比对两张工作表的差异" }
                    }
                },
                new WorkflowTemplate
                {
                    Name = "数据清洗并复检",
                    Description = "读取选区后清洗数据（去空行、修日期、补空值），再做统计复检确认质量提升。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_current_excel_context", Description = "读取当前选区上下文" },
                        new WorkflowStep { Tool = "clean_excel_data", Description = "清洗数据（去空行/修日期/补空值）" },
                        new WorkflowStep { Tool = "analyze_range_statistics", Description = "清洗后统计复检" }
                    }
                },
                new WorkflowTemplate
                {
                    Name = "透视汇总表",
                    Description = "读取选区并生成透视表，并套用表头加粗格式。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_current_selection_values_and_formulas", Description = "读取选区数值" },
                        new WorkflowStep
                        {
                            Tool = "create_excel_pivot_summary",
                            Description = "生成透视汇总表",
                            Arguments = new Dictionary<string, object>
                            {
                                ["source_address"] = "{selection}",
                                ["sheet"] = "{selection_sheet}"
                            },
                            Verify = new WorkflowVerify
                            {
                                Tool = "verify_pivot",
                                Description = "透视汇总自检",
                                Arguments = new Dictionary<string, object>
                                {
                                    ["target_address"] = "{step1.target_address}",
                                    ["expected_count"] = "{step1.row_count}"
                                }
                            }
                        },
                        new WorkflowStep
                        {
                            Tool = "format_excel_range",
                            Description = "套用表头加粗格式",
                            Arguments = new Dictionary<string, object>
                            {
                                ["address"] = "{step1.target_address}",
                                ["bold"] = true
                            }
                        }
                    }
                },
                new WorkflowTemplate
                {
                    Name = "数据筛选提炼",
                    Description = "读取选区，按条件筛选出关键行并高亮。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_excel_range_values_and_formulas", Description = "读取选区数值" },
                        new WorkflowStep { Tool = "filter_excel_data", Description = "按条件筛选关键数据" }
                    }
                },
                new WorkflowTemplate
                {
                    Name = "相关性+迷你图",
                    Description = "读取选区，分析列相关性，并为关键数值列插入迷你图洞察趋势。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_current_selection_values_and_formulas", Description = "读取选区数值" },
                        new WorkflowStep { Tool = "analyze_column_correlation", Description = "列相关性分析" },
                        new WorkflowStep { Tool = "create_excel_sparkline", Description = "为数值列插入迷你图" }
                    }
                },
                new WorkflowTemplate
                {
                    Name = "工作报告导出",
                    Description = "读取当前上下文，生成分析报告并导出为 PDF。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_current_excel_context", Description = "读取当前选区上下文" },
                        new WorkflowStep { Tool = "generate_analysis_report", Description = "生成分析报告" },
                        new WorkflowStep { Tool = "export_sheets_to_pdf", Description = "导出为 PDF" }
                    }
                },
                new WorkflowTemplate
                {
                    Name = "数值结构摘要",
                    Description = "读取选区并做基础统计，输出数值结构摘要报告（均值/极值/缺失等）。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_current_excel_context", Description = "读取当前选区上下文" },
                        new WorkflowStep { Tool = "analyze_range_statistics", Description = "区间基础统计" },
                        new WorkflowStep { Tool = "generate_analysis_report", Description = "生成结构摘要报告" }
                    }
                },
                new WorkflowTemplate
                {
                    Name = "双表差异报告",
                    Description = "列出工作表，比对两张表差异，并生成差异分析报告。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_excel_sheet_list", Description = "列出工作表清单" },
                        new WorkflowStep { Tool = "compare_two_sheets", Description = "比对两张表的差异" },
                        new WorkflowStep { Tool = "generate_analysis_report", Description = "生成差异分析报告" }
                    }
                },
                new WorkflowTemplate
                {
                    Name = "迷你图趋势速览",
                    Description = "读取选区，为关键数值列插入迷你图并套用格式，快速看趋势。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_current_selection_values_and_formulas", Description = "读取选区数值" },
                        new WorkflowStep { Tool = "create_excel_sparkline", Description = "为数值列插入迷你图" },
                        new WorkflowStep { Tool = "format_excel_range", Description = "套用表头加粗格式" }
                    }
                },
                new WorkflowTemplate
                {
                    Name = "清洗后一键出图",
                    Description = "先清洗选区数据，再推荐并插入图表，一步得到干净的可视化结果。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_current_excel_context", Description = "读取当前选区上下文" },
                        new WorkflowStep { Tool = "clean_excel_data", Description = "清洗数据（去空行/修日期/补空值）" },
                        new WorkflowStep { Tool = "recommend_excel_chart", Description = "推荐合适的图表类型" },
                        new WorkflowStep { Tool = "create_excel_chart", Description = "在选区旁插入图表" }
                    }
                },
                // ===== 领域模板：数据分析 / 人事 / 考勤 / 财务 =====
                new WorkflowTemplate
                {
                    Name = "数据分析·透视与报告",
                    Description = "对当前数据选区做统计、分布、相关性体检，生成透视汇总并输出分析报告，适合数据分析人员。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 2, 1, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_current_excel_context", Description = "读取当前选区上下文" },
                        new WorkflowStep { Tool = "analyze_range_statistics", Description = "区间基础统计", Arguments = new Dictionary<string, object> { ["address"] = "{selection}" } },
                        new WorkflowStep { Tool = "analyze_value_distribution", Description = "取值分布体检", Arguments = new Dictionary<string, object> { ["address"] = "{selection}" } },
                        new WorkflowStep { Tool = "analyze_column_correlation", Description = "列相关性分析", Arguments = new Dictionary<string, object> { ["address"] = "{selection}" } },
                        new WorkflowStep { Tool = "create_excel_pivot_summary", Description = "生成透视汇总表", Arguments = new Dictionary<string, object> { ["source_address"] = "{selection}", ["sheet"] = "{selection_sheet}" }, Verify = new WorkflowVerify { Tool = "verify_pivot", Description = "透视汇总自检", Arguments = new Dictionary<string, object> { ["target_address"] = "{step4.target_address}", ["expected_count"] = "{step4.row_count}", ["sheet"] = "{selection_sheet}" } } },
                        new WorkflowStep { Tool = "generate_analysis_report", Description = "输出数据分析报告", Arguments = new Dictionary<string, object> { ["address"] = "{selection}" } }
                    }
                },
                new WorkflowTemplate
                {
                    Name = "数据分析·图表洞察",
                    Description = "读取数据选区，识别数值结构，推荐并插入图表，导出图片供报告/汇报使用。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 2, 1, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_current_selection_values_and_formulas", Description = "读取选区数值" },
                        new WorkflowStep { Tool = "analyze_range_statistics", Description = "区间基础统计" },
                        new WorkflowStep { Tool = "recommend_excel_chart", Description = "推荐合适的图表类型" },
                        new WorkflowStep { Tool = "create_excel_chart", Description = "在选区旁插入图表" },
                        new WorkflowStep { Tool = "export_excel_chart_image", Description = "导出图表为图片" }
                    }
                },
                new WorkflowTemplate
                {
                    Name = "人事·档案数据体检",
                    Description = "对人事档案表做统计体检（结构/分布），清洗异常（空值/坏日期）并高亮，输出体检报告。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 2, 1, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_current_excel_context", Description = "读取人事档案选区上下文" },
                        new WorkflowStep { Tool = "analyze_range_statistics", Description = "档案字段基础统计" },
                        new WorkflowStep { Tool = "clean_excel_data", Description = "清洗异常（去空行/修日期/补空值）" },
                        new WorkflowStep { Tool = "format_excel_range", Description = "高亮标题与异常区域" },
                        new WorkflowStep { Tool = "generate_analysis_report", Description = "输出档案体检报告" }
                    }
                },
                new WorkflowTemplate
                {
                    Name = "人事·人员结构透视",
                    Description = "按部门/岗位对人员表生成透视汇总，并出图展示结构分布。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 2, 1, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_current_selection_values_and_formulas", Description = "读取人员选区数值" },
                        new WorkflowStep { Tool = "create_excel_pivot_summary", Description = "生成部门/岗位透视汇总", Arguments = new Dictionary<string, object> { ["source_address"] = "{selection}", ["sheet"] = "{selection_sheet}" }, Verify = new WorkflowVerify { Tool = "verify_pivot", Description = "透视汇总自检", Arguments = new Dictionary<string, object> { ["target_address"] = "{step1.target_address}", ["expected_count"] = "{step1.row_count}", ["sheet"] = "{selection_sheet}" } } },
                        new WorkflowStep { Tool = "recommend_excel_chart", Description = "推荐合适的图表类型", Arguments = new Dictionary<string, object> { ["address"] = "{selection}" } },
                        new WorkflowStep { Tool = "create_excel_chart", Description = "插入人员结构图表", Arguments = new Dictionary<string, object> { ["source_address"] = "{step1.target_address}" } },
                        new WorkflowStep { Tool = "format_excel_range", Description = "套用表头加粗格式", Arguments = new Dictionary<string, object> { ["address"] = "{step1.target_address}", ["bold"] = true } }
                    }
                },
                new WorkflowTemplate
                {
                    Name = "考勤·异常筛查",
                    Description = "对考勤表做统计，筛选出迟到/缺勤等异常行并高亮，输出考勤异常报告。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 2, 1, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_current_excel_context", Description = "读取考勤选区上下文" },
                        new WorkflowStep { Tool = "analyze_range_statistics", Description = "考勤字段基础统计" },
                        new WorkflowStep { Tool = "filter_excel_data", Description = "按条件筛选异常记录" },
                        new WorkflowStep { Tool = "format_excel_range", Description = "高亮异常行" },
                        new WorkflowStep { Tool = "generate_analysis_report", Description = "输出考勤异常报告" }
                    }
                },
                new WorkflowTemplate
                {
                    Name = "考勤·出勤趋势可视",
                    Description = "分析考勤数据的列相关性并为关键数值列插入迷你图，最终出图速览出勤趋势。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 2, 1, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_current_selection_values_and_formulas", Description = "读取考勤数值" },
                        new WorkflowStep { Tool = "analyze_column_correlation", Description = "出勤字段相关性分析" },
                        new WorkflowStep { Tool = "create_excel_sparkline", Description = "为数值列插入迷你图" },
                        new WorkflowStep { Tool = "create_excel_chart", Description = "插入出勤趋势图表" }
                    }
                },
                new WorkflowTemplate
                {
                    Name = "财务·数据体检报告",
                    Description = "对财务/收支表单做统计、分布体检，清洗异常并输出体检报告并导出 PDF。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 2, 1, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_current_excel_context", Description = "读取财务选区上下文" },
                        new WorkflowStep { Tool = "analyze_range_statistics", Description = "财务字段基础统计" },
                        new WorkflowStep { Tool = "analyze_value_distribution", Description = "金额取值分布体检" },
                        new WorkflowStep { Tool = "generate_analysis_report", Description = "生成财务体检报告" },
                        new WorkflowStep { Tool = "export_sheets_to_pdf", Description = "导出报告为 PDF" }
                    }
                },
                new WorkflowTemplate
                {
                    Name = "财务·收支透视与图表",
                    Description = "对收支表单生成透视汇总，出图展示收支结构，并导出图片便于汇报。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 2, 1, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_current_selection_values_and_formulas", Description = "读取收支数值" },
                        new WorkflowStep { Tool = "create_excel_pivot_summary", Description = "生成收支透视汇总", Arguments = new Dictionary<string, object> { ["source_address"] = "{selection}", ["sheet"] = "{selection_sheet}" }, Verify = new WorkflowVerify { Tool = "verify_pivot", Description = "透视汇总自检", Arguments = new Dictionary<string, object> { ["target_address"] = "{step1.target_address}", ["expected_count"] = "{step1.row_count}", ["sheet"] = "{selection_sheet}" } } },
                        new WorkflowStep { Tool = "recommend_excel_chart", Description = "推荐合适的图表类型", Arguments = new Dictionary<string, object> { ["address"] = "{selection}" } },
                        new WorkflowStep { Tool = "create_excel_chart", Description = "插入收支结构图表", Arguments = new Dictionary<string, object> { ["source_address"] = "{step1.target_address}" } },
                        new WorkflowStep { Tool = "export_excel_chart_image", Description = "导出图表为图片", Arguments = new Dictionary<string, object> { ["chart_name"] = "{step3.chart_name}" } }
                    }
                },
                new WorkflowTemplate
                {
                    Name = "数据·去重与排序",
                    Description = "对选区数据做整行去重后按主键排序，并套用表格格式，便于建规范台账。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 8, 1, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_current_selection_values_and_formulas", Description = "读取选区的源数据" },
                        new WorkflowStep { Tool = "dedup_range", Description = "按整行去重保留唯一记录" },
                        new WorkflowStep { Tool = "sort_range", Description = "按主键列排序" },
                        new WorkflowStep { Tool = "format_excel_range", Description = "套用表格格式" }
                    }
                },
                new WorkflowTemplate
                {
                    Name = "数据·查找替换与复检",
                    Description = "在选区/全表批量查找替换文本后，对结果做统计复检，确认改动影响范围。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 8, 1, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_current_selection_values_and_formulas", Description = "读取选区值以便替换" },
                        new WorkflowStep { Tool = "find_replace_range", Description = "批量查找替换文本" },
                        new WorkflowStep { Tool = "analyze_range_statistics", Description = "对改动后区域做统计复检" }
                    }
                },
                new WorkflowTemplate
                {
                    Name = "数据·转置排版",
                    Description = "把区块行列互换（转置）到目标位置，并整理格式，适用于横纵表头互换。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 8, 1, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_current_selection_values_and_formulas", Description = "读取待转置的数据" },
                        new WorkflowStep { Tool = "transpose_range", Description = "行列互换写入目标位置" },
                        new WorkflowStep { Tool = "format_excel_range", Description = "整理结果格式" }
                    }
                },
                new WorkflowTemplate
                {
                    Name = "数据·条件高亮异常",
                    Description = "对数值列按阈值/重复值高亮异常，并对整体做统计，便于快速定位问题。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 8, 1, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_current_selection_values_and_formulas", Description = "读取待检查的数值列" },
                        new WorkflowStep { Tool = "format_excel_conditional", Description = "按阈值或重复值高亮" },
                        new WorkflowStep { Tool = "analyze_range_statistics", Description = "对区域做统计概况" }
                    }
                },
                // ===== 第二批通用数据加工流程（文本/日期/编号/公式/多表合并） =====
                new WorkflowTemplate
                {
                    Name = "数据·文本清洗台账",
                    Description = "对选区文本做清洗（去空格/去控制字符/统一大小写），再统计复检确认质量提升。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 8, 20, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_current_selection_values_and_formulas", Description = "读取待清洗的文本" },
                        new WorkflowStep { Tool = "text_clean", Description = "清洗文本（去空格/控制字符/统一大小写）" },
                        new WorkflowStep { Tool = "analyze_range_statistics", Description = "清洗后统计复检" }
                    }
                },
                new WorkflowTemplate
                {
                    Name = "数据·文本拆分列",
                    Description = "把一列文本按分隔符拆分为多列（如姓名拆姓/名、地址拆省市区），并整理表头格式。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 8, 20, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_current_selection_values_and_formulas", Description = "读取待拆分的文本列" },
                        new WorkflowStep { Tool = "text_split", Description = "按分隔符拆分为多列" },
                        new WorkflowStep { Tool = "format_excel_range", Description = "整理结果表头格式" }
                    }
                },
                new WorkflowTemplate
                {
                    Name = "数据·文本合并列",
                    Description = "把多列文本合并为一列（如姓+名拼全名），并整理结果格式。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 8, 20, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_current_selection_values_and_formulas", Description = "读取待合并的多列" },
                        new WorkflowStep { Tool = "text_merge", Description = "按分隔符合并为一列" },
                        new WorkflowStep { Tool = "format_excel_range", Description = "整理结果格式" }
                    }
                },
                new WorkflowTemplate
                {
                    Name = "数据·日期标准化",
                    Description = "把日期列统一为标准格式（自动识别/序列日期），并对结果做统计复检。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 8, 20, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_current_selection_values_and_formulas", Description = "读取日期列" },
                        new WorkflowStep { Tool = "date_normalize", Description = "统一日期格式" },
                        new WorkflowStep { Tool = "analyze_range_statistics", Description = "对结果做统计复检" }
                    }
                },
                new WorkflowTemplate
                {
                    Name = "数据·编号规范化",
                    Description = "把编号列补零/加前后缀格式化为定长文本，并套用表格格式。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 8, 20, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_current_selection_values_and_formulas", Description = "读取编号列" },
                        new WorkflowStep { Tool = "number_pad", Description = "补零/加前后缀规范化" },
                        new WorkflowStep { Tool = "format_excel_range", Description = "套用表格格式" }
                    }
                },
                new WorkflowTemplate
                {
                    Name = "数据·批量填充公式",
                    Description = "按行批量填充公式模板（{row} 自动替换行号），并统计复检结果。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 8, 20, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_current_selection_values_and_formulas", Description = "读取源数据列" },
                        new WorkflowStep { Tool = "fill_formula", Description = "按行批量填充公式" },
                        new WorkflowStep { Tool = "analyze_range_statistics", Description = "对结果做统计复检" }
                    }
                },
                new WorkflowTemplate
                {
                    Name = "数据·多表合并汇总",
                    Description = "把多个工作表数据纵向合并到目标表（带表头），并整理格式。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 8, 20, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_excel_sheet_list", Description = "列出工作表清单" },
                        new WorkflowStep { Tool = "merge_sheets", Description = "多表数据纵向合并" },
                        new WorkflowStep { Tool = "format_excel_range", Description = "整理合并结果格式" }
                    }
                },
                // ===== 第二批补充：通用 / 人事 / 数据分析流程 =====
                new WorkflowTemplate
                {
                    Name = "通用·清洗-出图-报告一条龙",
                    Description = "读取选区→清洗数据→推荐插入图表→生成分析报告，一步完成从脏数据到可视化结论。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 8, 25, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_current_excel_context", Description = "读取当前选区上下文" },
                        new WorkflowStep { Tool = "clean_excel_data", Description = "清洗数据（去空行/修日期/补空值）" },
                        new WorkflowStep { Tool = "recommend_excel_chart", Description = "推荐合适的图表类型" },
                        new WorkflowStep { Tool = "create_excel_chart", Description = "插入可视化图表" },
                        new WorkflowStep { Tool = "generate_analysis_report", Description = "生成分析报告" }
                    }
                },
                new WorkflowTemplate
                {
                    Name = "通用·清洗后统计复核",
                    Description = "读取选区→清洗异常→统计复核，确认数据质量提升并给出前后对比说明。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 8, 25, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_current_selection_values_and_formulas", Description = "读取待清洗数据" },
                        new WorkflowStep { Tool = "analyze_range_statistics", Description = "清洗前统计基线" },
                        new WorkflowStep { Tool = "clean_excel_data", Description = "清洗数据异常" },
                        new WorkflowStep { Tool = "analyze_range_statistics", Description = "清洗后统计复检" }
                    }
                },
                new WorkflowTemplate
                {
                    Name = "通用·去重排序生成台账",
                    Description = "读取选区→整行去重→按主键排序→套用格式，生成一份干净规范的数据台账。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 8, 25, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_current_selection_values_and_formulas", Description = "读取待整理的源数据" },
                        new WorkflowStep { Tool = "dedup_range", Description = "按整行去重保留唯一记录" },
                        new WorkflowStep { Tool = "sort_range", Description = "按主键列排序" },
                        new WorkflowStep { Tool = "format_excel_range", Description = "套用表格格式" }
                    }
                },
                new WorkflowTemplate
                {
                    Name = "人事·年龄工龄分布透视",
                    Description = "对人员表的年龄/入职年限字段生成分布透视，并出图展示，掌握团队结构。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 8, 25, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_current_selection_values_and_formulas", Description = "读取人员表数值" },
                        new WorkflowStep { Tool = "create_excel_pivot_summary", Description = "按年龄/工龄生成分布透视", Arguments = new Dictionary<string, object> { ["source_address"] = "{selection}", ["sheet"] = "{selection_sheet}" }, Verify = new WorkflowVerify { Tool = "verify_pivot", Description = "透视汇总自检", Arguments = new Dictionary<string, object> { ["target_address"] = "{step1.target_address}", ["expected_count"] = "{step1.row_count}", ["sheet"] = "{selection_sheet}" } } },
                        new WorkflowStep { Tool = "recommend_excel_chart", Description = "推荐合适的图表类型", Arguments = new Dictionary<string, object> { ["address"] = "{selection}" } },
                        new WorkflowStep { Tool = "create_excel_chart", Description = "插入人员结构图表", Arguments = new Dictionary<string, object> { ["source_address"] = "{step1.target_address}" } },
                        new WorkflowStep { Tool = "format_excel_range", Description = "套用表头格式", Arguments = new Dictionary<string, object> { ["address"] = "{step1.target_address}", ["bold"] = true } }
                    }
                },
                new WorkflowTemplate
                {
                    Name = "人事·名册清洗去重体检",
                    Description = "对花名册做整行去重与字段清洗（空值/坏日期），并输出名册质量体检报告。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 8, 25, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_current_excel_context", Description = "读取人事名册选区上下文" },
                        new WorkflowStep { Tool = "dedup_range", Description = "按整行去重避免重复名册" },
                        new WorkflowStep { Tool = "clean_excel_data", Description = "清洗空值/坏日期等异常" },
                        new WorkflowStep { Tool = "generate_analysis_report", Description = "输出名册质量体检报告" }
                    }
                },
                new WorkflowTemplate
                {
                    Name = "人事·关键字段筛选高亮",
                    Description = "对人事表按部门/岗位/状态等条件筛选出关键记录，并高亮展示。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 8, 25, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_current_selection_values_and_formulas", Description = "读取人事表数值" },
                        new WorkflowStep { Tool = "filter_excel_data", Description = "按条件筛选关键记录" },
                        new WorkflowStep { Tool = "format_excel_range", Description = "高亮关键行/异常" }
                    }
                },
                new WorkflowTemplate
                {
                    Name = "数据分析·分布与直方图",
                    Description = "对数值列做分布体检并生成直方图/分布图，配合报告洞察数据形状。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 8, 25, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_current_selection_values_and_formulas", Description = "读取数值列" },
                        new WorkflowStep { Tool = "analyze_value_distribution", Description = "取值分布体检" },
                        new WorkflowStep { Tool = "recommend_excel_chart", Description = "推荐合适的图表类型" },
                        new WorkflowStep { Tool = "create_excel_chart", Description = "插入分布图表" },
                        new WorkflowStep { Tool = "generate_analysis_report", Description = "生成分布洞察报告" }
                    }
                },
                new WorkflowTemplate
                {
                    Name = "数据分析·多维聚合看板",
                    Description = "对数据区生成多维透视汇总（按分类/日期），输出关键指标并配图速览，形成看板素材。",
                    IsBuiltIn = true,
                    CreatedAtUtc = new DateTime(2026, 8, 25, 0, 0, 0, DateTimeKind.Utc),
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep { Tool = "get_current_selection_values_and_formulas", Description = "读取数据区数值" },
                        new WorkflowStep { Tool = "create_excel_pivot_summary", Description = "生成多维透视汇总", Arguments = new Dictionary<string, object> { ["source_address"] = "{selection}", ["sheet"] = "{selection_sheet}" }, Verify = new WorkflowVerify { Tool = "verify_pivot", Description = "透视汇总自检", Arguments = new Dictionary<string, object> { ["target_address"] = "{step1.target_address}", ["expected_count"] = "{step1.row_count}", ["sheet"] = "{selection_sheet}" } } },
                        new WorkflowStep { Tool = "analyze_range_statistics", Description = "关键指标统计", Arguments = new Dictionary<string, object> { ["address"] = "{selection}" } },
                        new WorkflowStep { Tool = "create_excel_chart", Description = "插入关键指标图表", Arguments = new Dictionary<string, object> { ["source_address"] = "{step1.target_address}" } },
                        new WorkflowStep { Tool = "generate_analysis_report", Description = "生成看板结论报告", Arguments = new Dictionary<string, object> { ["address"] = "{selection}" } }
                    }
                }
            };
        }
    }
}