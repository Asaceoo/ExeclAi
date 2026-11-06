using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;

namespace SlicerProbe
{
    /// <summary>
    /// E-D0 数据分析能力测试卷探针（exam 模式，2026-09-14）：自拟测试题逐项真机作答，
    /// 覆盖基础 8 + 高阶 10 共 18 项数据分析能力；每题数据带确定性真值（Python/scipy 预计算锁定），
    /// 断言产品工具输出与真值一致。FAIL 逐题定性（产品缺陷 vs 用例错），多轮迭代收敛。
    /// 清理纪律：按 PID 强杀本 harness 拉起的 Excel 实例。
    /// </summary>
    internal static class ExamProbe
    {
        private static int _passed, _failed;
        private static readonly HashSet<int> SpawnedPids = new HashSet<int>();

        private static void Expect(bool ok, string name, string detail)
        {
            if (ok) { _passed++; Console.WriteLine("[PASS] " + name); }
            else
            {
                _failed++;
                Console.WriteLine("[FAIL] " + name + " → " + (detail == null ? "" : detail.Length > 420 ? detail.Substring(0, 420) : detail));
                Console.WriteLine("       [FAIL@" + name + "] " + (detail == null ? "" : detail.Length > 1200 ? detail.Substring(0, 1200) : detail));
            }
        }

        private static ToolResult Exec(ToolRegistry registry, string tool, params (string k, object v)[] args)
        {
            var dict = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            foreach (var (k, v) in args) dict[k] = v;
            if (!registry.TryGet(tool, out var descriptor))
            {
                return new ToolResult { Ok = false, ErrorCode = "UNKNOWN_TOOL", Data = tool };
            }
            try { return descriptor.Execute(dict); }
            catch (Exception ex) { return new ToolResult { Ok = false, ErrorCode = "HARNESS_EXC", Data = ex.GetType().Name + " " + ex.Message }; }
        }

        private static string Json(ToolResult r)
        {
            return Newtonsoft.Json.JsonConvert.SerializeObject(new { r.Ok, r.ErrorCode, r.Data });
        }

        public static int Run(string sourcePath, string pythonPath = null)
        {
            Console.WriteLine("== 数据分析能力测试卷真机探针：" + sourcePath + " ==");
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine("FILE_NOT_FOUND: " + sourcePath);
                return 2;
            }

            // 口径题夹具：临时目录副本 + 同目录 _口径.md
            string tempDir = Path.Combine(Path.GetTempPath(), "EXAM_" + Guid.NewGuid().ToString("N").Substring(0, 8));
            Directory.CreateDirectory(tempDir);
            string tempCopy = Path.Combine(tempDir, "analysis-exam.xlsx");
            File.Copy(sourcePath, tempCopy, true);
            File.WriteAllText(Path.Combine(tempDir, "_口径.md"),
                "- 毛利率: (收入-成本)÷收入，百分比保留两位小数\n- 大客户: 年采购额≥100万元的客户\n",
                System.Text.Encoding.UTF8);

            dynamic excel = null;
            try
            {
                var before = Process.GetProcessesByName("EXCEL").Select(p => p.Id).ToHashSet();
                excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                excel.Visible = false;
                excel.DisplayAlerts = false;
                foreach (var p in Process.GetProcessesByName("EXCEL"))
                {
                    if (!before.Contains(p.Id)) { SpawnedPids.Add(p.Id); }
                    p.Dispose();
                }
                dynamic workbook = excel.Workbooks.Open(tempCopy);

                var dashboard = new ExcelDashboardService(excel);
                var pivotChart = new ExcelPivotChartService(excel, null, null);
                var registry = new ToolRegistry();
                Func<Dictionary<string, object>, ToolResult> stub = _ => new ToolResult { Ok = false, ErrorCode = "NOT_UNDER_TEST", Data = "skip" };
                ExcelToolHostRegistration.RegisterM3Tools(registry, stub, stub, null, pivotChart, dashboard);
                ExcelThirdPhaseRegistration.RegisterThirdPhaseTools(registry, excel, null, null);

                RunExam(registry, workbook, excel, pythonPath);

                Console.WriteLine($"== 测试卷结果：通过 {_passed}，失败 {_failed} ==");
                return _failed == 0 ? 0 : 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine("FATAL: " + ex.Message + "\n" + ex.StackTrace);
                return 3;
            }
            finally
            {
                foreach (var pid in SpawnedPids)
                {
                    try { using (var proc = Process.GetProcessById(pid)) { proc.Kill(); } } catch { }
                }
                try { Directory.Delete(tempDir, true); } catch { }
                Console.WriteLine("[cleanup][EXAM] 清理完成");
            }
        }

        private static void RunExam(ToolRegistry registry, dynamic workbook, dynamic excel, string pythonPath)
        {
            // ===== 基础 8 项 =====

            // E1 描述统计：产品A mean=155.8333 median=157.5
            var e1 = Exec(registry, "analyze_range_statistics", ("address", "A1:D13"), ("sheet", "S1统计"));
            string e1s = Json(e1);
            Expect(e1.Ok && e1s.Contains("\"mean\":155.8333") && e1s.Contains("\"median\":157.5") && e1s.Contains("basis"),
                "E1 描述统计（产品A mean=155.8333/median=157.5/basis）", e1.Ok ? e1s : e1.ErrorCode + " " + e1s);

            // E2 分布：A=10 B=15 C=25
            var e2 = Exec(registry, "analyze_value_distribution", ("address", "A1:B51"), ("sheet", "S2分布"), ("column", 2));
            string e2s = Json(e2);
            Expect(e2.Ok && e2s.Contains("\"value\":\"A\",\"count\":10") && e2s.Contains("\"value\":\"B\",\"count\":15") && e2s.Contains("\"value\":\"C\",\"count\":25"),
                "E2 值分布（A=10/B=15/C=25）", e2.Ok ? e2s : e2.ErrorCode + " " + e2s);

            // E3 相关性：r(x1,x2)=0.9985 强、r(x1,x3)=0.0015 弱
            var e3 = Exec(registry, "analyze_column_correlation", ("address", "A1:C21"), ("sheet", "S3相关"));
            string e3s = Json(e3);
            Expect(e3.Ok && e3s.Contains("0.9985") && e3s.Contains("0.0015") && e3s.Contains("strong"),
                "E3 相关性（r12=0.9985 强标记/r13=0.0015）", e3.Ok ? e3s : e3.ErrorCode + " " + e3s);

            // E4 画像：混合列 9 数字/4 文本/3 空
            var e4 = Exec(registry, "analyze_column_profile", ("address", "A1:A16"), ("sheet", "S4画像"), ("column", 1));
            string e4s = Json(e4);
            Expect(e4.Ok && e4s.Contains("\"number\":9") && e4s.Contains("\"text\":4") && e4s.Contains("\"blank_count\":2") && e4s.Contains("华东"),
                "E4 列画像（number=9/text=4/blank=2）", e4.Ok ? e4s : e4.ErrorCode + " " + e4s);

            // E5 口径：读工作簿同目录 _口径.md 命中条目
            var e5 = Exec(registry, "read_business_glossary");
            string e5s = Json(e5);
            Expect(e5.Ok && e5s.Contains("毛利率") && e5s.Contains("大客户"),
                "E5 口径文件（毛利率/大客户条目命中）", e5.Ok ? (e5s.Length > 200 ? e5s.Substring(0, 200) : e5s) : e5.ErrorCode);

            // E6 趋势预测：linear 第 13 期 = 220，趋势上升
            var e6 = Exec(registry, "analyze_trend_forecast", ("address", "A1:B13"), ("sheet", "S6趋势"), ("periods", 3));
            string e6s = Json(e6);
            Expect(e6.Ok && e6s.Contains("220") && e6s.Contains("上升") && e6s.Contains("basis"),
                "E6 趋势预测（第13期=220/上升/basis）", e6.Ok ? e6s : e6.ErrorCode + " " + e6s);

            // E7 离群点：3 个已知离群 500/505/510
            var e7 = Exec(registry, "analyze_range_statistics", ("address", "A1:A29"), ("sheet", "S7离群"));
            string e7s = Json(e7);
            Expect(e7.Ok && e7s.Contains("\"outlier_count\":3") && e7s.Contains("510") && e7s.Contains("500"),
                "E7 IQR 离群点（count=3/值命中）", e7.Ok ? e7s : e7.ErrorCode + " " + e7s);

            // E8 文本变体：3 组建议（腾讯组/字节组/美团组），代表值与成员命中
            var e8 = Exec(registry, "analyze_text_variants", ("address", "A1:A10"), ("sheet", "S8变体"), ("column", 1));
            string e8s = Json(e8);
            Expect(e8.Ok && e8s.Contains("\"group_count\":3") && e8s.Contains("腾讯科技有限公司") && e8s.Contains("字节跳动有限公司") && e8s.Contains("merge_options"),
                "E8 文本变体聚类（3 组/代表值/合并选项）", e8.Ok ? e8s : e8.ErrorCode + " " + e8s);

            // ===== 高阶 10 项 =====

            // E9a Welch 双样本 t：t=-9.3531 显著
            var e9a = Exec(registry, "analyze_hypothesis_test", ("address", "A1:B9"), ("sheet", "H1a Welch"), ("test", "two_sample_t"), ("value_column", 2), ("group_column", 1));
            string e9as = Json(e9a);
            Expect(e9a.Ok && e9as.Contains("\"statistic\":-9.3531") && e9as.Contains("\"significant\":true"),
                "E9a Welch t 检验（t=-9.3531/显著）", e9a.Ok ? e9as : e9a.ErrorCode + " " + e9as);

            // E9b 配对 t：t=-5.1657 显著
            var e9b = Exec(registry, "analyze_hypothesis_test", ("address", "A1:B5"), ("sheet", "H1b 配对"), ("test", "paired_t"), ("value_column", 1), ("value_column2", 2));
            string e9bs = Json(e9b);
            Expect(e9b.Ok && e9bs.Contains("\"statistic\":-5.1657") && e9bs.Contains("\"significant\":true"),
                "E9b 配对 t 检验（t=-5.1657/显著）", e9b.Ok ? e9bs : e9b.ErrorCode + " " + e9bs);

            // E9c 单因素 ANOVA：F=75 显著
            var e9c = Exec(registry, "analyze_hypothesis_test", ("address", "A1:B10"), ("sheet", "H1c ANOVA"), ("test", "anova"), ("value_column", 2), ("group_column", 1));
            string e9cs = Json(e9c);
            Expect(e9c.Ok && e9cs.Contains("\"statistic\":75") && e9cs.Contains("\"significant\":true"),
                "E9c ANOVA（F=75/显著）", e9c.Ok ? e9cs : e9c.ErrorCode + " " + e9cs);

            // E9d 卡方独立性：chi2=16.6667 显著（p≈4.5e-05）
            var e9d = Exec(registry, "analyze_hypothesis_test", ("address", "A1:B101"), ("sheet", "H1d 卡方"), ("test", "chi_square"), ("column_a", 1), ("column_b", 2));
            string e9ds = Json(e9d);
            Expect(e9d.Ok && e9ds.Contains("16.6667") && e9ds.Contains("\"significant\":true"),
                "E9d 卡方独立性（chi2=16.6667/显著）", e9d.Ok ? e9ds : e9d.ErrorCode + " " + e9ds);

            // E10 回归：y=2+3x 完美拟合（截距 2/斜率 3/R²=1）
            var e10 = Exec(registry, "analyze_regression", ("address", "A1:B9"), ("sheet", "H2回归"), ("target_column", 2));
            string e10s = Json(e10);
            Expect(e10.Ok && e10s.Contains("\"estimate\":2") && e10s.Contains("\"estimate\":3") && e10s.Contains("\"r2\":1") && e10s.Contains("coefficients"),
                "E10 回归（截距=2/斜率=3/R²=1）", e10.Ok ? e10s : e10.ErrorCode + " " + e10s);

            // E11 聚类：4 远离簇自动定 k=4 高轮廓
            var e11 = Exec(registry, "analyze_cluster", ("address", "A1:B17"), ("sheet", "H3聚类"));
            string e11s = Json(e11);
            Expect(e11.Ok && e11s.Contains("\"k\":4") && (e11s.Contains("\"best_silhouette\":0.9") || e11s.Contains("\"best_silhouette\":1")),
                "E11 聚类（自动 k=4/高轮廓）", e11.Ok ? e11s : e11.ErrorCode + " " + e11s);

            // ===== Python 通道高阶（E12~E18），run_python_code/excel 直调 =====
            if (string.IsNullOrEmpty(pythonPath))
            {
                foreach (var cand in new[]
                {
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Python312", "python.exe"),
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Python311", "python.exe"),
                    "python"
                })
                {
                    if (cand == "python" || File.Exists(cand)) { pythonPath = cand; break; }
                }
            }

            if (string.IsNullOrEmpty(pythonPath))
            {
                for (var i = 12; i <= 18; i++) Expect(false, "E" + i + "（本机未找到 Python）", "PYTHON_NOT_FOUND");
            }
            else
            {
                var pyOptions = new ScriptOptions { Enabled = true, PythonPath = pythonPath, TimeoutSeconds = 300 };
                var pyCode = new ScriptRunnerService(excel, pyOptions);
                var pyExcel = new PythonExcelService(excel, pyOptions);

                // E12 时序预测（Holt 指数平滑，statsmodels）：24 月趋势+周期，预测第 25 期在延长线附近
                var e12 = pyCode.RunPythonCode(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["code"] = "import numpy as np\nfrom statsmodels.tsa.holtwinters import Holt\nm = np.arange(1, 25)\ny = 100 + 8 * m + 5 * np.sin(m)\nfit = Holt(y, initialization_method='estimated').fit()\npred = float(fit.forecast(1)[0])\nlast = float(y[-1])\nxai_result({'pred': round(pred, 2), 'last': last, 'ok_range': bool(last + 2 < pred < last + 40)})",
                    ["timeout_ms"] = 120000
                });
                string e12s = Json(e12);
                Expect(e12.Ok && e12s.Contains("\"ok_range\":true"),
                    "E12 时序预测 Holt（预测在延长线区间）", e12.Ok ? e12s : e12.ErrorCode + " " + e12s);

                // E13 降维 PCA：三列强相关，PC1 方差占比 >0.95
                var e13 = pyCode.RunPythonCode(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["code"] = "import numpy as np\nfrom sklearn.decomposition import PCA\nx = np.arange(1.0, 21.0)\ny = 0.9 * x + np.array([0.3,-0.7,0.5,1.1,-0.2,0.8,-1.0,0.4,0.6,-0.5,0.9,-0.8,0.2,1.0,-0.6,0.7,-0.3,0.1,0.5,-0.4])\nz = -0.8 * x + 0.5\nX = np.column_stack([x, y, z])\np = PCA().fit(X)\nshare = float(p.explained_variance_ratio_[0])\nxai_result({'pc1': round(share, 4), 'dominant': bool(share > 0.95)})",
                    ["timeout_ms"] = 60000
                });
                string e13s = Json(e13);
                Expect(e13.Ok && e13s.Contains("\"dominant\":true"),
                    "E13 降维 PCA（PC1 方差占比>0.95）", e13.Ok ? e13s : e13.ErrorCode + " " + e13s);

                // E14 异常进阶 IsolationForest：5 个极端点全部被判异常
                var e14 = pyCode.RunPythonCode(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["code"] = "import numpy as np\nfrom sklearn.ensemble import IsolationForest\nrng = np.random.RandomState(7)\nX = np.vstack([rng.normal(0, 1, (200, 2)), [[10,10],[10.5,10],[-10,-9.5],[9,-10],[10,-10]]])\nm = IsolationForest(contamination=0.03, random_state=42).fit(X)\nlabels = m.predict(X)\nextreme = labels[200:]\nxai_result({'extreme_flagged': int((extreme == -1).sum()), 'all_caught': bool((extreme == -1).all())})",
                    ["timeout_ms"] = 60000
                });
                string e14s = Json(e14);
                Expect(e14.Ok && e14s.Contains("\"all_caught\":true"),
                    "E14 IsolationForest（5 极端点全检出）", e14.Ok ? e14s : e14.ErrorCode + " " + e14s);

                // E15 文本情感 snownlp：正/负/中 三条分类正确
                var e15 = pyCode.RunPythonCode(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["code"] = "from snownlp import SnowNLP\npos = SnowNLP('物流很快，质量超预期，非常满意').sentiments\nneg = SnowNLP('用了两天就坏了，客服态度还差').sentiments\nxai_result({'pos_ok': bool(pos > 0.6), 'neg_ok': bool(neg < 0.4), 'pos': round(pos, 3), 'neg': round(neg, 3)})",
                    ["timeout_ms"] = 60000
                });
                string e15s = Json(e15);
                Expect(e15.Ok && e15s.Contains("\"pos_ok\":true") && e15s.Contains("\"neg_ok\":true"),
                    "E15 snownlp 情感（正负分类正确）", e15.Ok ? e15s : e15.ErrorCode + " " + e15s);

                // E16 实验与抽样：两比例 z 检验 + 功效（statsmodels）
                var e16 = pyCode.RunPythonCode(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["code"] = "from statsmodels.stats.proportion import proportions_ztest, proportion_effectsize\nfrom statsmodels.stats.power import NormalIndPower\nz, p = proportions_ztest([120, 150], [1000, 1000])\nes = proportion_effectsize(0.12, 0.10)\npower = NormalIndPower().power(abs(es), nobs1=1000, alpha=0.05, ratio=1)\nxai_result({'z': round(float(z), 3), 'p': round(float(p), 5), 'power': round(float(power), 4), 'significant': bool(p < 0.05), 'power_low': bool(power < 0.6)})",
                    ["timeout_ms"] = 60000
                });
                string e16s = Json(e16);
                Expect(e16.Ok && e16s.Contains("\"z\":-1.963") && e16s.Contains("\"significant\":true") && e16s.Contains("\"power_low\":true"),
                    "E16 A/B 比例检验+功效（z=-1.963/显著/功效不足提示）", e16.Ok ? e16s : e16.ErrorCode + " " + e16s);

                // E17 优化与模拟：LP 最优 36 + 蒙特卡洛 π
                var e17 = pyCode.RunPythonCode(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["code"] = "import numpy as np\nfrom scipy.optimize import linprog\nr = linprog(c=[-3, -5], A_ub=[[1, 0], [0, 2], [3, 2]], b_ub=[4, 12, 18])\nrng = np.random.RandomState(42)\npts = rng.random_sample((10000, 2))\npi = 4 * float((pts[:, 0] ** 2 + pts[:, 1] ** 2 < 1).mean())\nxai_result({'lp_ok': bool(r.status == 0), 'best': round(float(-r.fun), 4), 'pi': round(pi, 3), 'pi_ok': bool(3.0 < pi < 3.3)})",
                    ["timeout_ms"] = 60000
                });
                string e17s = Json(e17);
                Expect(e17.Ok && e17s.Contains("\"best\":36") && e17s.Contains("\"lp_ok\":true") && e17s.Contains("\"pi_ok\":true"),
                    "E17 线性规划最优 36+蒙特卡洛 π", e17.Ok ? e17s : e17.ErrorCode + " " + e17s);

                // E18 综合评价 TOPSIS：读工作簿 H10评价 表，排名 甲>丙>乙
                var e18 = pyExcel.Run(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["code"] = "import numpy as np, pandas as pd\nsht = xw.books.active.sheets['H10评价']\ndf = sht.range('A1').options(pd.DataFrame, expand='table').value\nm = df[['质量', '成本', '交付']].values.astype(float)\nnames = list(df.index.astype(str))\nm[:, 1] = 1.0 / m[:, 1]\nnorm = m / np.sqrt((m ** 2).sum(axis=0))\nw = np.array([0.5, 0.3, 0.2])\nweighted = norm * w\nip = weighted.max(axis=0); inn = weighted.min(axis=0)\ndp = np.sqrt(((weighted - ip) ** 2).sum(axis=1)); dn = np.sqrt(((weighted - inn) ** 2).sum(axis=1))\nscore = dn / (dp + dn)\nrank = [names[i] for i in np.argsort(-score)]\nxai_result({'rank': rank, 'score': [round(float(s), 4) for s in score], 'top_is_jia': bool(rank[0] == '甲')})",
                    ["timeout_ms"] = 120000
                });
                string e18s = Json(e18);
                Expect(e18.Ok && e18s.Contains("\"rank\":[\"甲\",\"丙\",\"乙\"]") && e18s.Contains("\"top_is_jia\":true"),
                    "E18 TOPSIS 综合评价（排名 甲>丙>乙）", e18.Ok ? e18s : e18.ErrorCode + " " + e18s);
            }
        }
    }
}
