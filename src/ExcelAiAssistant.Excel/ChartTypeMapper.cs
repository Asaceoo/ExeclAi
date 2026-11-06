using System;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// 图表类型名 → Excel COM 可识别值；供 create_excel_chart / insert_recommended_chart 复用。
    /// </summary>
    public static class ChartTypeMapper
    {
        public static string Resolve(string typeName)
        {
            if (string.IsNullOrWhiteSpace(typeName))
            {
                return "ColumnClustered";
            }

            switch (typeName.Trim().ToLowerInvariant())
            {
                case "column":
                case "columnclustered":
                    return "ColumnClustered";
                case "bar":
                case "barclustered":
                    return "BarClustered";
                case "line":
                    return "Line";
                case "pie":
                    return "Pie";
                case "area":
                    return "Area";
                case "scatter":
                case "xyscatter":
                    return "xlXYScatter";
                case "histogram":
                    return "xlHistogram";
                // ---- P1 扩展（M-D1）：环形族 / 雷达族 ----
                case "doughnut":
                case "doughnutchart":
                    return "Doughnut";
                case "doughnutexploded":
                case "doughnut_exploded":
                    return "DoughnutExploded";
                case "radar":
                    return "Radar";
                case "radarfilled":
                    return "RadarFilled";
                case "radarmarkers":
                    return "RadarMarkers";
                default:
                    return "ColumnClustered";
            }
        }

        public static bool IsHistogram(string typeName)
        {
            return typeName != null && typeName.Trim().ToLowerInvariant() == "histogram";
        }

        /// <summary>chart_type 是否为 create_excel_chart 支持的 7 种枚举值（与 schema enum 对齐），未知返回 false。</summary>
        public static bool IsKnown(string typeName)
        {
            switch (typeName?.Trim().ToLowerInvariant())
            {
                case "column":
                case "columnclustered":
                case "bar":
                case "barclustered":
                case "line":
                case "pie":
                case "area":
                case "scatter":
                case "xyscatter":
                case "histogram":
                // ---- P1 扩展（M-D1）----
                case "doughnut":
                case "doughnutchart":
                case "doughnutexploded":
                case "doughnut_exploded":
                case "radar":
                case "radarfilled":
                case "radarmarkers":
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// 返回 Excel XlChartType 的整数值，供 Chart.ChartType 赋值（动态 COM 不能直接赋字符串）。
        /// </summary>
        public static int ResolveXlChartType(string typeName)
        {
            switch (Resolve(typeName))
            {
                case "BarClustered": return 57;
                case "Line": return 4;
                case "Pie": return 5;
                case "Area": return 1;
                case "xlXYScatter": return -4169;
                // 工具分类审查批次修正（58=xlBarStacked 冒充直方图真 bug）：
                // xlHistogram 权威值=118（2016+ 新图表族，与 create_chart_advanced nativeTypes 同源实证）。
                case "xlHistogram": return 118;
                // ---- P1 扩展（M-D1）：环形族 / 雷达族（XlChartType 枚举值）----
                case "Doughnut": return -4120;          // xlDoughnut
                // 2026-09-11 真机 enumprobe 实证：XlChartType 不存在 -4121/-4152/-4153（手抄臆造值，
                // chart.ChartType 赋值即 E_FAIL）。分离环形/填充雷达/标记雷达 = 基础类型 + 变体后处理
                //（CreateChart.ApplyChartVariant：Points.Explosion / 系列填充 / MarkerStyle）
                case "DoughnutExploded": return -4120;  // xlDoughnut + Explosion 后处理
                case "Radar": return -4151;             // xlRadar
                case "RadarFilled": return -4151;       // xlRadar + 系列半透明填充
                case "RadarMarkers": return -4151;      // xlRadar + MarkerStyle 后处理
                case "ColumnClustered":
                default:
                    // xlColumnClustered = 51。注意：-4102 是 xl3DPie（3D 饼图），
                    // 曾有把柱状图映射成 -4102 导致渲染成 3D 饼图的历史 bug。
                    return 51;
            }
        }
    }
}