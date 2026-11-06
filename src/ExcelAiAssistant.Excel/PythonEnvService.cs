using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelAiAssistant.Core.Diagnostics;
using ExcelAiAssistant.Core.Models;
using Newtonsoft.Json.Linq;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// Python 环境管理三件套（v3.7.1.39）：
    /// - list_python_packages：列出已装库（名称/版本/安装位置）+ Python 解释器位置与版本。
    /// - install_python_packages：常用库一键安装（白名单 + 镜像兜底 + --user 兜底），安装走 worker subprocess（环境健康实证）。
    /// - python_setup_guide：小白级 Python 安装教程 + 常用库目录（每个库做什么的详细说明），供 AI 转述给用户选择装不装。
    /// </summary>
    public sealed class PythonEnvService
    {
        private readonly dynamic application;
        private readonly ScriptOptions options;

        public PythonEnvService(dynamic application, ScriptOptions options)
        {
            this.application = application;
            this.options = options ?? new ScriptOptions();
        }

        /// <summary>常用库目录：pip 包名 → (导入名, 用途说明)。</summary>
        internal static readonly Dictionary<string, (string ImportName, string Description)> Catalog =
            new Dictionary<string, (string, string)>(StringComparer.OrdinalIgnoreCase)
            {
                ["pandas"] = ("pandas", "数据分析核心库：像操作 Excel 表格一样筛选/汇总/透视/合并数据，做批量数据处理的首选。"),
                ["numpy"] = ("numpy", "数值计算基础库：多维数组和矩阵运算，pandas/matplotlib/scipy 都依赖它。"),
                ["matplotlib"] = ("matplotlib", "绘图库：把数据画成柱状图/折线图/散点图等，可导出 PNG 图片。"),
                ["seaborn"] = ("seaborn", "统计图表库：基于 matplotlib，几行代码画出更美观的热力图/分布图。"),
                ["wordcloud"] = ("wordcloud", "词云图库：按词频把文本关键词画成词云 PNG（依赖 pillow+numpy；Windows 含 C 扩展，安装失败走镜像兜底）。"),
                ["plotly"] = ("plotly", "交互式图表库：生成的图表支持缩放/悬停提示，可导出 HTML。"),
                ["openpyxl"] = ("openpyxl", "读写 .xlsx 工作簿（不动 Excel 程序）：适合后台批量生成/修改表格与单元格样式。"),
                ["xlsxwriter"] = ("xlsxwriter", "高速写 .xlsx：支持大表导出与单元格内嵌图表/条件格式。"),
                ["xlwings"] = ("xlwings", "Python 直连正在运行的 Excel：读写当前工作簿/工作表（本工具 run_python_excel 依赖它）。"),
                ["pywin32"] = ("win32com", "Windows COM 组件访问：Python 操作 Excel/WPS/Outlook 等 Windows 程序的底层依赖（run_python_excel 依赖它）。"),
                ["pillow"] = ("PIL", "图像处理库：裁剪/缩放/格式转换/加水印，处理截图与图片必备（导入名是 PIL）。"),
                ["requests"] = ("requests", "HTTP 请求库：调用网页接口/下载数据的标配（受控执行下建议仅访问可信地址）。"),
                ["beautifulsoup4"] = ("bs4", "网页/HTML 解析库：从 HTML 里提取表格与数据（导入名是 bs4）。"),
                ["lxml"] = ("lxml", "高性能 XML/HTML 解析：大文件解析与 XPath 查询。"),
                ["scipy"] = ("scipy", "科学计算库：统计分布/优化/插值/信号处理。"),
                ["statsmodels"] = ("statsmodels", "时间序列统计建模库：趋势/季节性分解（STL）、指数平滑与 ARIMA 预测；依赖 scipy，安装时自动带上 patsy。"),
                ["scikit-learn"] = ("sklearn", "机器学习库：回归/分类/聚类等经典算法（导入名是 sklearn）。"),
                ["opencv-python"] = ("cv2", "计算机视觉库：图像识别/轮廓检测/视频处理（导入名是 cv2）。"),
                ["python-docx"] = ("docx", "生成/修改 Word 文档：批量做报告与合同（导入名是 docx）。"),
                ["python-pptx"] = ("pptx", "生成/修改 PowerPoint：批量做汇报幻灯片（导入名是 pptx）。"),
                ["pyyaml"] = ("yaml", "读写 YAML 配置文件（导入名是 yaml）。"),
                ["chardet"] = ("chardet", "自动检测文本/文件的编码（GBK/UTF-8 等），解决中文乱码第一步。"),
                ["openpyxl-image-loader"] = ("openpyxl_image_loader", "往 xlsx 单元格批量插图的辅助库。"),
                ["jieba"] = ("jieba", "中文分词库：把中文句子切成词语，做文本统计的第一步。"),
                ["snownlp"] = ("snownlp", "中文情感分析库：对客户评价/工单等中文文本打情感分（0~1，越大越积极），支持批量分析倾向与好评差评分类。")
            };

        private static readonly Regex SafePkgName =
            new Regex(@"^[A-Za-z0-9][A-Za-z0-9._-]*$", RegexOptions.Compiled);

        /// <summary>list_python_packages：已装库清单 + Python 位置。</summary>
        public ToolResult ListPackages(Dictionary<string, object> args)
        {
            if (options.Enabled == false)
            {
                return new ToolResult { Ok = false, ErrorCode = "DISABLED", Data = "受控执行总开关未启用，无法探测 Python 环境。" };
            }

            var execFile = string.IsNullOrWhiteSpace(options.PythonPath) ? "python" : options.PythonPath;
            const string code = @"
import sys
try:
    import importlib.metadata as md
    pkgs = []
    for d in md.distributions():
        try:
            pkgs.append({'name': (d.metadata['Name'] or '?'), 'version': d.version, 'location': str(d.locate_file('') or '')})
        except Exception:
            pass
    xai_result({'python_exe': sys.executable, 'python_version': sys.version.split()[0],
                'bits': __import__('struct').calcsize('P') * 8, 'count': len(pkgs), 'packages': pkgs})
except Exception:
    import traceback
    xai_result({'error': traceback.format_exc()[-800:]})
";
            var workerOk = PythonWorkerHost.TryExecute(
                code, execFile, 60000, 256 * 1024, null, 2048 * 1024,
                new PythonWorkerHost.ExecuteOptions(), out var ok, out var output, out var result,
                out _, out _, out _);

            if (!workerOk || !ok)
            {
                AppLog.Error("python-env", "list_python_packages 失败：" + (output ?? "worker 无响应"));
                return new ToolResult
                {
                    Ok = false,
                    ErrorCode = "PYTHON_ENV_ERROR",
                    Data = "无法枚举已安装库（Python 未安装或不可用：" + (output ?? "worker 无响应") + "）。可调用 python_setup_guide 获取安装教程。"
                };
            }

            var jo = result as JObject;
            if (jo == null || jo["error"] != null)
            {
                var err = jo?["error"] != null ? Convert.ToString(jo["error"]) : "worker 返回结构异常";
                AppLog.Error("python-env", "list_python_packages 枚举异常：" + err);
                return new ToolResult { Ok = false, ErrorCode = "PYTHON_ENV_ERROR", Data = err };
            }

            var filter = OptionalString(args, "filter");
            var packages = (jo["packages"] as JArray ?? new JArray());
            var listed = packages;
            if (!string.IsNullOrWhiteSpace(filter))
            {
                listed = new JArray(packages.Where(p =>
                    Convert.ToString(p["name"]).IndexOf(filter.Trim(), StringComparison.OrdinalIgnoreCase) >= 0));
            }

            AppLog.Info("python-env", $"list_python_packages：python={jo.Value<string>("python_exe")} version={jo.Value<string>("python_version")} 共{packages.Count}库，列出{listed.Count}");
            return new ToolResult
            {
                Ok = true,
                Data = new
                {
                    python_exe = jo.Value<string>("python_exe"),
                    python_version = jo.Value<string>("python_version"),
                    bits = jo.Value<int?>("bits") ?? 0,
                    total = packages.Count,
                    filter = filter,
                    packages = listed,
                    note = "location 即库的安装位置（site-packages 目录）。"
                }
            };
        }

        /// <summary>check_common_libraries（FR-K1 库管理增强，UI 专用不注册 AI 工具）：常用库目录逐一比对安装状态。
        /// 只读探测不安装；匹配口径 = pip 包名与导入名均做「小写化 + 下划线转连字符」归一后比对已装元数据名。</summary>
        public ToolResult CheckCommonLibraries(Dictionary<string, object> args)
        {
            if (options.Enabled == false)
            {
                return new ToolResult { Ok = false, ErrorCode = "DISABLED", Data = "受控执行总开关未启用，无法探测 Python 环境。" };
            }

            var execFile = string.IsNullOrWhiteSpace(options.PythonPath) ? "python" : options.PythonPath;
            const string code = @"
import sys
try:
    import importlib.metadata as md
    pkgs = []
    for d in md.distributions():
        try:
            pkgs.append({'name': (d.metadata['Name'] or '?'), 'version': d.version, 'location': str(d.locate_file('') or '')})
        except Exception:
            pass
    xai_result({'python_exe': sys.executable, 'python_version': sys.version.split()[0],
                'bits': __import__('struct').calcsize('P') * 8, 'count': len(pkgs), 'packages': pkgs})
except Exception:
    import traceback
    xai_result({'error': traceback.format_exc()[-800:]})
";
            var workerOk = PythonWorkerHost.TryExecute(
                code, execFile, 60000, 256 * 1024, null, 2048 * 1024,
                new PythonWorkerHost.ExecuteOptions(), out var ok, out var output, out var result,
                out _, out _, out _);

            if (!workerOk || !ok)
            {
                AppLog.Error("python-env", "check_common_libraries 失败：" + (output ?? "worker 无响应"));
                return new ToolResult
                {
                    Ok = false,
                    ErrorCode = "PYTHON_ENV_ERROR",
                    Data = "无法枚举已安装库（Python 未安装或不可用：" + (output ?? "worker 无响应") + "）。可调用 python_setup_guide 获取安装教程。"
                };
            }

            var jo = result as JObject;
            if (jo == null || jo["error"] != null)
            {
                var err = jo?["error"] != null ? Convert.ToString(jo["error"]) : "worker 返回结构异常";
                AppLog.Error("python-env", "check_common_libraries 枚举异常：" + err);
                return new ToolResult { Ok = false, ErrorCode = "PYTHON_ENV_ERROR", Data = err };
            }

            var installed = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var p in jo["packages"] as JArray ?? new JArray())
            {
                var name = NormalizePkg(Convert.ToString(p["name"]));
                if (string.IsNullOrEmpty(name))
                {
                    continue;
                }

                if (!installed.ContainsKey(name))
                {
                    installed[name] = Convert.ToString(p["version"]);
                }
            }

            var items = new List<object>();
            var installedCount = 0;
            foreach (var kv in Catalog)
            {
                installed.TryGetValue(NormalizePkg(kv.Key), out var ver);
                if (ver == null)
                {
                    installed.TryGetValue(NormalizePkg(kv.Value.ImportName), out ver);
                }

                var found = ver != null;
                if (found)
                {
                    installedCount++;
                }

                items.Add(new
                {
                    pip_name = kv.Key,
                    import_name = kv.Value.ImportName,
                    description = kv.Value.Description,
                    installed = found,
                    version = ver ?? string.Empty
                });
            }

            AppLog.Info("python-env", $"check_common_libraries：目录 {Catalog.Count} 库，已装 {installedCount}，缺失 {Catalog.Count - installedCount}");
            return new ToolResult
            {
                Ok = true,
                Data = new
                {
                    python_exe = jo.Value<string>("python_exe"),
                    python_version = jo.Value<string>("python_version"),
                    bits = jo.Value<int?>("bits") ?? 0,
                    total = Catalog.Count,
                    installed_count = installedCount,
                    missing_count = Catalog.Count - installedCount,
                    items
                }
            };
        }

        /// <summary>包名归一：小写化 + 下划线转连字符（PEP 503 规范化），用于 pip 名/导入名/已装元数据名三方比对。</summary>
        internal static string NormalizePkg(string name)
        {
            return (name ?? string.Empty).Trim().ToLowerInvariant().Replace('_', '-');
        }

        /// <summary>install_python_packages：常用库一键安装（白名单 + 镜像兜底）。</summary>
        public ToolResult InstallPackages(Dictionary<string, object> args)
        {
            if (options.Enabled == false)
            {
                return new ToolResult { Ok = false, ErrorCode = "DISABLED", Data = "受控执行总开关未启用，无法安装库。" };
            }

            var requested = ExtractPackageNames(args);
            if (requested.Count == 0)
            {
                return new ToolResult
                {
                    Ok = false,
                    ErrorCode = "MISSING_PACKAGES",
                    Data = "缺少 packages：请给出要安装的 pip 包名列表（须来自常用库目录，可先调用 python_setup_guide 获取目录与用途说明，向用户确认后再安装）。"
                };
            }

            var execFile = string.IsNullOrWhiteSpace(options.PythonPath) ? "python" : options.PythonPath;
            var results = new List<object>();
            var allOk = true;
            foreach (var pkg in requested)
            {
                if (!Catalog.TryGetValue(pkg, out var info) && !SafePkgName.IsMatch(pkg))
                {
                    allOk = false;
                    results.Add(new { package = pkg, ok = false, error = "包名非法或不在常用库目录内。" });
                    continue;
                }

                AppLog.Info("python-env", "install_python_packages 开始安装：" + pkg);
                var installOk = InstallWithMirrorFallback(execFile, pkg, out var message);
                if (installOk)
                {
                    AppLog.Info("python-env", $"install_python_packages[{pkg}] ok=true {message}");
                }
                else
                {
                    AppLog.Error("python-env", $"install_python_packages[{pkg}] ok=false {message}");
                }
                if (!installOk)
                {
                    allOk = false;
                }

                results.Add(new
                {
                    package = pkg,
                    import_name = info.ImportName,
                    description = info.Description,
                    ok = installOk,
                    message = message
                });
            }

            return new ToolResult
            {
                Ok = allOk,
                ErrorCode = allOk ? null : "PYTHON_INSTALL_PARTIAL",
                Data = new
                {
                    results = results,
                    note = allOk
                        ? "全部安装成功。请把每个库的用途说明转述给用户。"
                        : "存在安装失败项：把失败 message 原样转述给用户，并建议调用 python_setup_guide 获取手动安装教程。"
                }
            };
        }

        /// <summary>python_setup_guide：小白级安装教程 + 常用库目录。</summary>
        public ToolResult SetupGuide(Dictionary<string, object> args)
        {
            var execFile = string.IsNullOrWhiteSpace(options.PythonPath) ? "python" : options.PythonPath;
            var excelBits = IntPtr.Size * 8;
            string pythonStatus;
            var prober = new PythonExcelService(application, options);
            var probe = prober.Probe(execFile);
            if (!probe.Found)
            {
                pythonStatus = "未检测到 Python（" + execFile + "）。请按下方教程安装。";
            }
            else if (!probe.ParseOk)
            {
                pythonStatus = "已找到 Python 但探测失败：" + (probe.Raw ?? "").Trim();
            }
            else if (probe.Bits != excelBits)
            {
                pythonStatus = $"已安装 Python {probe.Version}（{probe.Bits} 位），但与当前 Excel（{excelBits} 位）位数不一致——请安装 {excelBits} 位 Python。";
            }
            else
            {
                pythonStatus = $"已安装 Python {probe.Version}（{probe.Bits} 位），与 Excel 位数一致，可直接使用。";
            }

            AppLog.Info("python-env", "python_setup_guide：status=" + pythonStatus);
            return new ToolResult
            {
                Ok = true,
                Data = new
                {
                    python_status = pythonStatus,
                    excel_bits = excelBits,
                    configured_python = execFile,
                    install_tutorial = BuildTutorial(),
                    catalog = Catalog.OrderBy(kv => kv.Key).Select(kv => new
                    {
                        pip_name = kv.Key,
                        import_name = kv.Value.ImportName,
                        description = kv.Value.Description,
                        install = "pip install " + kv.Key
                    }),
                    note = "AI 应先把常用库目录与用途说明展示给用户，等用户明确选择要装哪些后再调用 install_python_packages。"
                }
            };
        }

        private static string BuildTutorial()
        {
            return string.Join("\n", new[]
            {
                "【Python 安装教程（小白版）】",
                "1. 下载：打开 https://www.python.org/downloads/ 点黄色大按钮下载最新版（建议 3.10/3.11/3.12，务必 64 位 Windows installer）。",
                "2. 安装：双击安装包，第一屏务必勾选底部「Add python.exe to PATH」，再点 Install Now（默认装到 C:\\Users\\你的用户名\\AppData\\Local\\Programs\\Python\\Python3xx\\）。",
                "3. 验证：装完打开新的命令行窗口输入 python --version，能显示版本号即成功；再输入 python -c \"import struct;print(struct.calcsize('P')*8)\" 应显示 64。",
                "4. 位数要求：必须与 Excel 位数一致（多数为 64 位），否则无法直连 Excel。",
                "5. 装完回到本助手，在设置里把 Python 路径指向 python.exe（或保持默认 python 让系统 PATH 解析），再重试。",
                "6. 常用库不需要手动 pip install：让助手调用 install_python_packages 一键安装；也可自己在命令行执行目录里的 pip install 包名。",
                "7. pip 慢/失败时用国内镜像：pip install 包名 -i https://pypi.tuna.tsinghua.edu.cn/simple"
            });
        }

        private bool InstallWithMirrorFallback(string execFile, string pkg, out string message)
        {
            message = null;
            var mirrors = new[] { "https://pypi.tuna.tsinghua.edu.cn/simple", "https://mirrors.aliyun.com/pypi/simple/", null };
            foreach (var mirror in mirrors)
            {
                if (PythonWorkerHost.TryPipInstall(execFile, pkg, mirror, false, 280000, out var outText))
                {
                    message = "安装成功（源：" + (mirror ?? "默认 PyPI") + "）。";
                    AppLog.Info("python-env", $"InstallWithMirrorFallback[{pkg}] ok=true mirror={(mirror ?? "default")}");
                    return true;
                }

                AppLog.Warn("python-env", $"InstallWithMirrorFallback[{pkg}] mirror={(mirror ?? "default")} 失败：{(outText ?? "").Substring(0, Math.Min(300, (outText ?? "").Length))}");
                var combined = outText ?? "";
                if ((combined.IndexOf("permission", StringComparison.OrdinalIgnoreCase) >= 0
                     || combined.IndexOf("access is denied", StringComparison.OrdinalIgnoreCase) >= 0)
                    && !combined.Contains("--user"))
                {
                    if (PythonWorkerHost.TryPipInstall(execFile, pkg, mirror, true, 280000, out var outUser))
                    {
                        message = "安装成功（--user，源：" + (mirror ?? "默认 PyPI") + "）。";
                        return true;
                    }
                }

                message = combined;
            }

            return false;
        }

        private static List<string> ExtractPackageNames(Dictionary<string, object> args)
        {
            var names = new List<string>();
            if (!args.TryGetValue("packages", out var raw) || raw == null)
            {
                return names;
            }

            void AddOne(object item)
            {
                if (item == null)
                {
                    return;
                }

                if (item is JValue jv && jv.Type == JTokenType.String)
                {
                    item = Convert.ToString(jv.Value);
                }

                var text = Convert.ToString(item)?.Trim();
                if (!string.IsNullOrWhiteSpace(text))
                {
                    names.Add(text);
                }
            }

            if (raw is System.Collections.IEnumerable list && !(raw is string))
            {
                foreach (var item in list)
                {
                    AddOne(item);
                }
            }
            else
            {
                foreach (var part in Convert.ToString(raw).Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    AddOne(part);
                }
            }

            return names.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }

        private static string OptionalString(Dictionary<string, object> args, string key)
        {
            if (args != null && args.TryGetValue(key, out var raw) && raw != null)
            {
                if (raw is JValue jv && jv.Type == JTokenType.String)
                {
                    return Convert.ToString(jv.Value);
                }

                return Convert.ToString(raw);
            }

            return null;
        }
    }
}
