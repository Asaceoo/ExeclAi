using System;
using System.Collections.Generic;
using System.Reflection;

class LargeIntProbe
{
    static string OutFile = @"d:\execl1.0\tools\smoke-results\largeintprobe-out.txt";

    static void Mark(string m)
    {
        System.IO.File.AppendAllText(OutFile, m + Environment.NewLine);
    }

    static void Main()
    {
        System.IO.File.WriteAllText(OutFile, "BOOT" + Environment.NewLine);
        try
        {
            var dir = @"d:\execl1.0\src\ExcelAiAssistant.Addin\bin\x64\Release\net48\win-x64";
            var asm = Assembly.LoadFrom(System.IO.Path.Combine(dir, "ExcelAiAssistant.Excel.dll"));
            Mark("ASM-LOADED");

            var excelType = Type.GetTypeFromProgID("Excel.Application");
            dynamic excel = Activator.CreateInstance(excelType);
            excel.Visible = false;
            excel.DisplayAlerts = false;
            System.Threading.Thread.Sleep(2000);
            dynamic book = excel.Workbooks.Add();
            book.Activate();
            Mark("BOOK-ADDED");

            var hostType = asm.GetType("ExcelAiAssistant.Excel.ExcelToolHost", true);
            dynamic toolHost = Activator.CreateInstance(hostType, new object[] { excel, null, "li", "li-run", "大整数" });
            var scriptOpts = toolHost.ScriptOptions;
            scriptOpts.Enabled = true;
            var registry = toolHost.CreateRegistry();

            dynamic tool = null;
            foreach (var t in registry.Tools)
            {
                if ((string)t.Name == "run_vba_snippet") { tool = t; break; }
            }
            if (tool == null) { Mark("TOOL-NOT-FOUND"); return; }
            Mark("TOOL-FOUND");

            // 1) warmup (no args) -> dump module XaiInitArgs
            var warm = new Dictionary<string, object>();
            warm["snippet"] = "Sub TT预热()\n    Range(\"H1\").Value = \"ok\"\nEnd Sub";
            warm["macro_name"] = "TT预热";
            var r1 = tool.Execute(warm);
            Mark("WARM ok=" + r1.Ok + " code=" + r1.ErrorCode);
            DumpInitArgs(excel, "WARM-INIT");

            // 2) control (string arg)
            var ctl = new Dictionary<string, object>();
            ctl["snippet"] = "Sub TT通道()\n    Dim v As Variant\n    v = XaiArg(\"column\")\n    XaiReturn Array(v, 42)\nEnd Sub";
            ctl["macro_name"] = "TT通道";
            var cd = new Dictionary<string, object>();
            cd["column"] = "A";
            ctl["args"] = cd;
            var r2 = tool.Execute(ctl);
            Mark("CTL ok=" + r2.Ok + " code=" + r2.ErrorCode + " data=" + Convert.ToString(r2.Data));
            DumpInitArgs(excel, "CTL-INIT");

            // 3) large int (id=3000000000)
            var li = new Dictionary<string, object>();
            li["snippet"] = "Sub TT大数()\n    Dim id As Variant\n    id = XaiArg(\"id\")\n    XaiReturn Array(id, 42)\nEnd Sub";
            li["macro_name"] = "TT大数";
            var ld = new Dictionary<string, object>();
            ld["id"] = 3000000000L;
            li["args"] = ld;
            var r3 = tool.Execute(li);
            Mark("LI ok=" + r3.Ok + " code=" + r3.ErrorCode + " data=" + Convert.ToString(r3.Data));
            DumpInitArgs(excel, "LI-INIT");

            // 4) large int via string JSON arg form (model may pass JSON string)
            var li2 = new Dictionary<string, object>();
            li2["snippet"] = "Sub TT大数2()\n    Dim id As Variant\n    id = XaiArg(\"id\")\n    XaiReturn Array(id, 42)\nEnd Sub";
            li2["macro_name"] = "TT大数2";
            li2["args"] = "{\"id\":3000000000}";
            var r4 = tool.Execute(li2);
            Mark("LI2 ok=" + r4.Ok + " code=" + r4.ErrorCode + " data=" + Convert.ToString(r4.Data));

            Mark("DONE");
        }
        catch (Exception ex)
        {
            Mark("FAIL " + ex.ToString());
        }
    }

    // Dump the XaiInitArgs body from the scratch module to see what gXaiArgs literal was injected.
    static void DumpInitArgs(dynamic excel, string tag)
    {
        try
        {
            var vbProject = excel.ActiveWorkbook.VBProject;
            dynamic comp = null;
            try { comp = vbProject.VBComponents.Item("JYYJHelper"); } catch { }
            if (comp == null) { Mark(tag + " MODULE-MISSING"); return; }
            var cm = comp.CodeModule;
            var cnt = cm.CountOfLines;
            for (int ln = 1; ln <= cnt; ln++)
            {
                var line = cm.Lines(ln, 1);
                if (line != null && line.IndexOf("gXaiArgs =") >= 0)
                {
                    var s = line.Trim();
                    var max = Math.Min(200, s.Length);
                    Mark(tag + " " + s.Substring(0, max));
                    return;
                }
            }
            Mark(tag + " NO-GXAIARGS-LINE");
        }
        catch (Exception ex)
        {
            Mark(tag + " DUMP-FAIL " + ex.Message);
        }
    }
}
