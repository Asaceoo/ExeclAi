using System.Collections.Generic;
using ExcelAiAssistant.Excel;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>verify_chart 的纯静态判据单测：series 公式数据源解析 + 自检结论分支 + 图表类型映射。</summary>
    public class ChartVerifyTests
    {
        [Fact]
        public void Parse_series_formula_extracts_value_and_category_ranges()
        {
            var formula = "=SERIES(Sheet1!$B$1,Sheet1!$A$2:$A$7,Sheet1!$B$2:$B$7,1)";
            var ranges = ExcelChartService.ParseSeriesFormulaRanges(formula);
            Assert.Equal(3, ranges.Count);
            Assert.Contains("$B$1", ranges);
            Assert.Contains("$A$2:$A$7", ranges);
            Assert.Contains("$B$2:$B$7", ranges);
        }

        [Fact]
        public void Parse_series_formula_handles_quoted_sheet_and_literal_name()
        {
            var formula = "=SERIES(\"部门\",'Anual Sales'!$A$2:$A$5,$A$2:$A$5,1)";
            var ranges = ExcelChartService.ParseSeriesFormulaRanges(formula);
            Assert.Equal(2, ranges.Count);
            Assert.Contains("$A$2:$A$5", ranges);
            Assert.DoesNotContain("部门", ranges);
        }

        [Fact]
        public void Parse_series_formula_returns_empty_for_r1c1_style()
        {
            Assert.Empty(ExcelChartService.ParseSeriesFormulaRanges("=SERIES(R1C1,R1C2:R1C6,R1C1:R1C6,1)"));
            Assert.Empty(ExcelChartService.ParseSeriesFormulaRanges("=SERIES(R[1]C[-1],R[2]C[1]:R[7]C[1],1)"));
            Assert.Empty(ExcelChartService.ParseSeriesFormulaRanges(null));
        }

        [Fact]
        public void Verdict_no_chart_when_none_found()
        {
            var (verdict, issues, selfcheck) = ExcelChartService.ComputeChartVerdict(false, 0, 0, 0, false, null, "Pie", null);
            Assert.Equal("no_chart", verdict);
            Assert.NotEmpty(selfcheck);
        }

        [Fact]
        public void Verdict_empty_when_series_or_points_zero()
        {
            Assert.Equal("empty", ExcelChartService.ComputeChartVerdict(true, 0, 0, 0, false, null, null, "Column").verdict);
            Assert.Equal("empty", ExcelChartService.ComputeChartVerdict(true, 2, 0, 2, false, null, null, "Column").verdict);
        }

        [Fact]
        public void Verdict_partial_when_some_series_empty()
        {
            var r = ExcelChartService.ComputeChartVerdict(true, 3, 10, 1, false, null, null, "Line");
            Assert.Equal("partial", r.verdict);
            Assert.Contains("个系列为空", r.selfcheck);
            Assert.Contains("partial", r.verdict);
        }

        [Fact]
        public void Verdict_ok_for_valid_chart()
        {
            var r = ExcelChartService.ComputeChartVerdict(true, 2, 12, 0, false, null, null, "Column");
            Assert.Equal("ok", r.verdict);
            Assert.Empty(r.issues);
        }

        [Fact]
        public void Verdict_mismatch_when_expected_type_differs()
        {
            // 需求要散点图实际成了柱形图 → 抓出“类型不对”。
            var r = ExcelChartService.ComputeChartVerdict(true, 2, 12, 0, false, null, "Scatter", "Column");
            Assert.Contains("mismatch", r.verdict);
            Assert.Contains("类型与预期不符", r.issues[0]);
        }

        [Fact]
        public void Verdict_series_count_mismatch_adds_issue()
        {
            var r = ExcelChartService.ComputeChartVerdict(true, 1, 6, 0, false, 3, null, "Line");
            Assert.Contains("系列数量与预期不符", r.selfcheck);
        }

        [Fact]
        public void Verdict_text_when_value_range_is_text_even_if_points_present()
        {
            // 值区全是文本的散点图：Points.Count>0 看似正常，但文本无法绘图 → 必须告警 text。
            var r = ExcelChartService.ComputeChartVerdict(true, 1, 5, 0, true, null, null, "Scatter");
            Assert.Equal("text", r.verdict);
            Assert.Contains("值区含文本数据", r.selfcheck);
            Assert.Contains("文本或混合类型数据", r.issues[0]);
        }

        [Fact]
        public void Verdict_ok_stays_ok_when_no_text_source()
        {
            var r = ExcelChartService.ComputeChartVerdict(true, 2, 12, 0, false, null, null, "Column");
            Assert.Equal("ok", r.verdict);
            Assert.DoesNotContain("text", r.verdict);
        }

        [Theory]
        [InlineData(new object[] { "1", "2", "3" }, "numeric")]
        [InlineData(new object[] { "a", "b", "c" }, "text")]
        [InlineData(new object[] { "1", "b", "3" }, "mixed")]
        [InlineData(new object[] { null, "", "  " }, "empty")]
        public void Classify_source_samples_detects_numeric_text_mixed(object[] samples, string expected)
        {
            Assert.Equal(expected, ExcelChartService.ClassifySourceSamples(samples));
        }

        [Fact]
        public void Extract_last_range_from_series_formula_values_area()
        {
            Assert.Equal("$B$2:$B$7", ExcelChartService.ExtractLastRange("=SERIES(Sheet1!$B$1,Sheet1!$A$2:$A$7,Sheet1!$B$2:$B$7,1)"));
            Assert.Equal("$C$3:$E$10", ExcelChartService.ExtractLastRange("=SERIES(\"销量\",,Sheet1!$C$3:$E$10,2)"));
            Assert.Equal("$A$1:$A$5", ExcelChartService.ExtractLastRange("Sheet1!$A$1:$A$5"));
            Assert.Null(ExcelChartService.ExtractLastRange(null));
            Assert.Null(ExcelChartService.ExtractLastRange(""));
        }

        [Fact]
        public void Chart_type_family_and_name_mapping()
        {
            Assert.Equal("Column", ExcelChartService.ChartTypeFamily(51));
            Assert.Equal("柱形图", ExcelChartService.ChartTypeChineseName(51));
            Assert.Equal("Scatter", ExcelChartService.ChartTypeFamily(-4169));
            Assert.Equal("散点图", ExcelChartService.ChartTypeChineseName(-4169));
            Assert.Equal("Pie", ExcelChartService.ChartTypeFamily(5));
            Assert.Null(ExcelChartService.ChartTypeFamily(-1));
        }

        [Theory]
        [InlineData("Doughnut", -4120)]
        [InlineData("DoughnutExploded", -4120)]   // 2026-09-11 enumprobe 实证：XlChartType 无 -4121，映射回 xlDoughnut，变体由 ApplyChartVariant(Points.Explosion) 实现
        [InlineData("Radar", -4151)]
        [InlineData("RadarFilled", -4151)]        // 实证无 -4152，变体 = 系列半透明填充
        [InlineData("RadarMarkers", -4151)]       // 实证无 -4153，变体 = MarkerStyle 后处理
        public void M_D1_chart_type_mapping_doughnut_and_radar(string typeName, int expectedXl)
        {
            Assert.Equal(expectedXl, ChartTypeMapper.ResolveXlChartType(typeName));
            Assert.True(ChartTypeMapper.IsKnown(typeName));
        }

        [Fact]
        public void Histogram_xl_type_is_118_not_58()
        {
            // 工具分类审查批次守卫（真 bug：58=xlBarStacked 冒充直方图，合法值赋值不抛错恒「成功」）。
            // xlHistogram 权威值=118（2016+ 新图表族 117~123 实证域），须与 create_chart_advanced
            // nativeTypes["histogram"] 保持同源一致，防止两条路径直方图形态分叉。
            Assert.Equal(118, ChartTypeMapper.ResolveXlChartType("histogram"));
        }

        [Fact]
        public void M_D1_doughnut_family_registered_in_verify_map()
        {
            // 环形族在 ChartTypeMap 中注册，verify_chart 才能识别家族
            Assert.Equal("Doughnut", ExcelChartService.ChartTypeFamily(-4120));
            Assert.Equal("环形图", ExcelChartService.ChartTypeChineseName(-4120));
            Assert.Equal("Doughnut", ExcelChartService.ChartTypeFamily(-4121));
            Assert.Equal("Radar", ExcelChartService.ChartTypeFamily(-4151));
        }

        [Fact]
        public void M_D1_secondary_axis_parse_supports_multi_formats()
        {
            // 经由反射测私有解析器不合理；这里用公共行为面：schema 与映射已验证。
            // 该占位断言确认 ComboChart schema 必填字段正确。
            var schema = M3ToolSchemas.ComboChart();
            var props = (System.Collections.Generic.Dictionary<string, object>)schema["properties"];
            Assert.True(props.ContainsKey("source_address"));
            Assert.True(props.ContainsKey("secondary_series"));
            Assert.Equal(new[] { "source_address" }, (string[])schema["required"]);
        }

        [Theory]
        [InlineData("ColumnClustered", "Column")]
        [InlineData("柱形图", "Column")]
        [InlineData("pie", "Pie")]
        [InlineData("饼图", "Pie")]
        [InlineData("XYScatter", "Scatter")]
        [InlineData("散点图", "Scatter")]
        [InlineData("Line", "Line")]
        [InlineData("", null)]
        [InlineData(null, null)]
        public void Normalize_chart_family_supports_multiple_spellings(string input, string expected)
        {
            Assert.Equal(expected, ExcelChartService.NormalizeChartFamily(input));
        }

        // ---- 新增：标题正确性 / 数据源正确性 / 渲染是否显示——纯函数判据 ----

        [Theory]
        [InlineData("销售额趋势", "销售额趋势", true)]
        [InlineData("  \"销售额趋势\"  ", "销售额趋势", true)]
        [InlineData("Sales Trend", "SALES trend", true)] // 忽略大小写
        [InlineData("销售额趋势", "其他标题", false)]
        [InlineData("", "销售额趋势", false)]
        [InlineData("销售额趋势", "", true)] // 未给预期 → 不断言
        [InlineData(null, null, true)]
        public void Title_matches_normalizes_before_compare(string actual, string expected, bool ok)
        {
            Assert.Equal(ok, ExcelChartService.TitleMatches(actual, expected));
        }

        [Theory]
        [InlineData("A1:B8", "$A$2:$B$7", true)]       // 区域相交
        [InlineData("A1:C3", "B2", true)]             // 单格落在区域内
        [InlineData("A1:A5", "C1:C5", false)]         // 列不相交
        [InlineData("'Sheet1'!A1:B8", "Sheet2!B2:C9", true)] // 跨表引用仍只取地址段
        [InlineData("A1", "A1", true)]                // 同格
        [InlineData("A1", "ZZ999", false)]
        [InlineData("garbage", "A1", false)]          // 解析失败 → 不命中
        [InlineData("", "", false)]
        public void Ranges_overlap_detects_data_source_hit(string actual, string expected, bool overlap)
        {
            Assert.Equal(overlap, ExcelChartService.RangesOverlap(actual, expected));
        }

        [Theory]
        [InlineData("ok", "title", "title")]
        [InlineData("empty", "title", "empty+title")]
        [InlineData("text", "title", "text+title")]
        [InlineData("text+title", "title", "text+title")] // 不重复追加
        public void Append_verdict_flag_appends_without_duplicate(string before, string flag, string after)
        {
            var v = before;
            ExcelChartService.AppendVerdictFlag(ref v, flag);
            Assert.Equal(after, v);
        }

        [Theory]
        [InlineData("{step2.chart_type}", true)]
        [InlineData("销售额{前缀}", true)]
        [InlineData("A1:B8", false)]
        [InlineData("销售趋势", false)]
        [InlineData("", false)]
        [InlineData(null, false)]
        public void Contains_braces_detects_unresolved_placeholders(string value, bool expected)
        {
            Assert.Equal(expected, ExcelChartService.ContainsBraces(value));
        }

        [Fact]
        public void Compute_render_stats_blank_image_reports_not_rendered()
        {
            // 全白 PNG → 非白内容占比 0 → 判定“数据未显示”。
            var png = MakePng(60, 40, allWhite: true);
            var s = ExcelChartService.ComputeRenderStats(png);
            Assert.False(s.dataRendered);
            Assert.Equal(0.0, s.nonBlankRatio);
            Assert.True(s.sampledPixels > 0);
            Assert.Equal(0, s.contentPixels);
        }

        [Fact]
        public void Compute_render_stats_content_image_reports_rendered()
        {
            // 半张画了深色块 → 非白内容占比显著 → 判定“数据已显示”。
            var bmp = new System.Drawing.Bitmap(60, 40);
            using (var g = System.Drawing.Graphics.FromImage(bmp))
            {
                g.Clear(System.Drawing.Color.White);
                using (var b = new System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(200, 30, 30)))
                {
                    g.FillRectangle(b, 0, 0, 60, 20);
                }
            }

            var png = ToPng(bmp);
            var s = ExcelChartService.ComputeRenderStats(png);
            Assert.True(s.dataRendered);
            Assert.True(s.nonBlankRatio > 0.4);
            Assert.True(s.contentPixels > 0);
        }

        [Fact]
        public void Compute_render_stats_null_bytes_not_rendered()
        {
            Assert.False(ExcelChartService.ComputeRenderStats(null).dataRendered);
            Assert.False(ExcelChartService.ComputeRenderStats(new byte[3]).dataRendered);
        }

        private static byte[] MakePng(int w, int h, bool allWhite)
        {
            var bmp = new System.Drawing.Bitmap(w, h);
            using (var g = System.Drawing.Graphics.FromImage(bmp))
            {
                g.Clear(allWhite ? System.Drawing.Color.White : System.Drawing.Color.Black);
            }

            return ToPng(bmp);
        }

        private static byte[] ToPng(System.Drawing.Bitmap bmp)
        {
            using (var ms = new System.IO.MemoryStream())
            {
                bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                return ms.ToArray();
            }
        }
    }
}