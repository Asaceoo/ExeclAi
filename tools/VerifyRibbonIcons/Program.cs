using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

internal static class Program
{
    private static int _fail;

    private static void Check(bool cond, string name)
    {
        Console.WriteLine((cond ? "[PASS] " : "[FAIL] ") + name);
        if (!cond) _fail++;
    }

    private static int Main(string[] args)
    {
        var dll = args.Length > 0
            ? args[0]
            : @"D:\execl1.0\src\ExcelAiAssistant.Addin\bin\x64\Release\net48\win-x64\ExcelAiAssistant.Addin.dll";
        if (!File.Exists(dll))
        {
            Console.WriteLine("[FAIL] addin dll not found: " + dll);
            return 1;
        }

        var asm = Assembly.LoadFrom(dll);
        var resources = asm.GetManifestResourceNames();

        // 1) 图标资源齐全且恰好 24 个
        var pngs = resources.Where(r => r.EndsWith(".png", StringComparison.OrdinalIgnoreCase)).ToList();
        Check(pngs.Count == 24, $"图标资源数量=24 (actual={pngs.Count})");

        // 2) 每个图标可解码为 96x96 ARGB 位图，且经转换器可产出非空 IPictureDisp
        var converterType = asm.GetType("ExcelAiAssistant.Addin.RibbonImageConverter");
        var toPictureDisp = converterType?.GetMethod("ToPictureDisp", BindingFlags.Public | BindingFlags.Static);
        Check(converterType != null && toPictureDisp != null, "RibbonImageConverter.ToPictureDisp 可反射找到");
        int badBitmap = 0, badPicture = 0;
        foreach (var png in pngs)
        {
            try
            {
                using (var stream = asm.GetManifestResourceStream(png))
                using (var bmp = new Bitmap(stream))
                {
                    var sizeOk = bmp.Width == 96 && bmp.Height == 96;
                    if (!sizeOk)
                    {
                        Console.WriteLine("    [warn] size " + bmp.Width + "x" + bmp.Height + " " + png);
                        badBitmap++;
                    }

                    if (toPictureDisp != null)
                    {
                        using (var probe = new Bitmap(bmp))
                        {
                            var pic = toPictureDisp.Invoke(null, new object[] { probe });
                            if (pic == null)
                            {
                                Console.WriteLine("    [warn] ToPictureDisp null " + png);
                                badPicture++;
                            }
                            else if (pic.GetType().FullName.IndexOf("ComObject", StringComparison.Ordinal) < 0)
                            {
                                Console.WriteLine("    [warn] picture type=" + pic.GetType().FullName + " " + png);
                                badPicture++;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("    [warn] decode/converter fail " + png + " " + ex.GetType().Name);
                badBitmap++;
            }
        }

        Check(badBitmap == 0, "全部图标可解码为 96x96 位图");
        Check(badPicture == 0, "全部图标经转换器产出非空 IPictureDisp");

        // 3) Ribbon.xml 资源存在且有 getImage 覆盖
        var ribbonName = resources.FirstOrDefault(r => r.EndsWith("Ribbon.xml", StringComparison.OrdinalIgnoreCase));
        Check(ribbonName != null, "Ribbon.xml 内嵌资源存在 (name=" + (ribbonName ?? "?") + ")");
        string ribbonXml = null;
        if (ribbonName != null)
        {
            using (var sr = new StreamReader(asm.GetManifestResourceStream(ribbonName)))
            {
                ribbonXml = sr.ReadToEnd();
            }
        }

        var getImageCount = ribbonXml == null ? 0 : Regex.Matches(ribbonXml, "getImage=\"GetImage\"").Count;
        Check(getImageCount == 24, $"Ribbon.xml 中 getImage 覆盖=24 (actual={getImageCount})");

        // 4) 清单兜底：每个 getImage 按钮的 id 必须存在同名 xxx.png 图标，且每个按钮 onAction 处理函数存在
        int missingIcon = 0, missingHandler = 0, leftoverMso = 0;
        if (ribbonXml != null)
        {
            var addinType = asm.GetType("ExcelAiAssistant.Addin.ExcelAiAddIn");
            foreach (Match m in Regex.Matches(ribbonXml, "id=\"(?<id>[^\"]+)\"[^>]*getImage=\"GetImage\"[^>]*>"))
            {
                var id = m.Groups["id"].Value;
                if (!pngs.Contains(id + ".png", StringComparer.OrdinalIgnoreCase))
                {
                    Console.WriteLine("    [warn] 缺图标资源 " + id + ".png");
                    missingIcon++;
                }
            }

            foreach (Match m in Regex.Matches(ribbonXml, "onAction=\"(?<a>[^\"]+)\""))
            {
                var a = m.Groups["a"].Value;
                if (addinType == null || addinType.GetMethod(a, BindingFlags.Public | BindingFlags.Instance) == null)
                {
                    Console.WriteLine("    [warn] 缺处理函数 " + a);
                    missingHandler++;
                }
            }

            leftoverMso = Regex.Matches(ribbonXml, "imageMso=").Count;
        }

        Check(missingIcon == 0, "每个 getImage 按钮均有同名图标资源");
        Check(missingHandler == 0, "所有按钮 onAction 均有处理函数");
        Check(leftoverMso == 0, "无残留 imageMso（已全部替换为自定义图标）");
        Check(ribbonXml != null && ribbonXml.IndexOf("JYYJQuickTab", StringComparison.Ordinal) < 0, "「JYYJ快捷」页签已移除(无 JYYJQuickTab)");
        Check(ribbonXml != null && Regex.Matches(ribbonXml, "<tab id=\"ExcelAiTab\"").Count == 1, "仅保留单一「JYYJ助手」页签");

        Console.WriteLine("");
        Console.WriteLine(_fail == 0 ? "ALL PASS" : "FAILED=" + _fail);
        return _fail == 0 ? 0 : 1;
    }
}