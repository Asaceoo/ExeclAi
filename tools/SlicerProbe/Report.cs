using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SlicerProbe
{
    /// <summary>M-D0 切片器探测结果模型（JSON 报告的数据载体）。</summary>
    internal sealed class ProbeItem
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public bool Passed { get; set; }
        public string HResult { get; set; }
        public string Error { get; set; }
        public string Detail { get; set; }
    }

    /// <summary>探测报告（写出到 reports/slicer_probe_&lt;version&gt;.json）。</summary>
    internal sealed class ProbeReport
    {
        public string ProbeTime { get; set; }
        public string ExcelVersion { get; set; }
        public string ExcelBuild { get; set; }
        public int ExcelPid { get; set; }
        public bool SlicerNativeReady { get; set; }
        public string Conclusion { get; set; }
        public List<ProbeItem> Items { get; set; } = new List<ProbeItem>();
        public string ErrorCatalog { get; set; }
    }

    /// <summary>报告 JSON 序列化（net48 无 System.Text.Json，手写轻量转义）。</summary>
    internal static class ReportWriter
    {
        public static string ToJson(ProbeReport r)
        {
            var sb = new System.Text.StringBuilder();
            sb.Append("{\n");
            sb.Append("  \"probe_time\": ").Append(Quote(r.ProbeTime)).Append(",\n");
            sb.Append("  \"excel_version\": ").Append(Quote(r.ExcelVersion)).Append(",\n");
            sb.Append("  \"excel_build\": ").Append(Quote(r.ExcelBuild)).Append(",\n");
            sb.Append("  \"excel_pid\": ").Append(r.ExcelPid).Append(",\n");
            sb.Append("  \"slicer_native_ready\": ").Append(r.SlicerNativeReady ? "true" : "false").Append(",\n");
            sb.Append("  \"conclusion\": ").Append(Quote(r.Conclusion)).Append(",\n");
            sb.Append("  \"items\": [\n");
            for (int i = 0; i < r.Items.Count; i++)
            {
                var it = r.Items[i];
                sb.Append("    {\"id\": ").Append(Quote(it.Id))
                  .Append(", \"name\": ").Append(Quote(it.Name))
                  .Append(", \"passed\": ").Append(it.Passed ? "true" : "false")
                  .Append(", \"hresult\": ").Append(Quote(it.HResult))
                  .Append(", \"error\": ").Append(Quote(it.Error))
                  .Append(", \"detail\": ").Append(Quote(it.Detail))
                  .Append("}").Append(i < r.Items.Count - 1 ? "," : "").Append("\n");
            }
            sb.Append("  ],\n");
            sb.Append("  \"error_catalog\": ").Append(Quote(r.ErrorCatalog)).Append("\n");
            sb.Append("}\n");
            return sb.ToString();
        }

        public static void Write(ProbeReport report, string outDir)
        {
            Directory.CreateDirectory(outDir);
            var safeVersion = (report.ExcelVersion ?? "unknown").Replace('.', '_');
            var path = Path.Combine(outDir, "slicer_probe_" + safeVersion + ".json");
            File.WriteAllText(path, ToJson(report));
            Console.WriteLine("[report] 探测报告已写出：" + path);
        }

        private static string Quote(string s)
        {
            if (s == null)
            {
                return "null";
            }

            var sb = new System.Text.StringBuilder("\"");
            foreach (var c in s)
            {
                switch (c)
                {
                    case '"': sb.Append("\\\""); break;
                    case '\\': sb.Append("\\\\"); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    default:
                        if (c < ' ')
                        {
                            sb.Append("\\u").Append(((int)c).ToString("x4"));
                        }
                        else
                        {
                            sb.Append(c);
                        }
                        break;
                }
            }
            sb.Append('"');
            return sb.ToString();
        }
    }
}
