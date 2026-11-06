using System;
using System.Collections.Generic;
using System.Reflection;

class StandaloneProbe
{
    static string OutFile = @"d:\execl1.0\tools\smoke-results\csprobe-out.txt";

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
            var excelDll = System.IO.Path.Combine(dir, "ExcelAiAssistant.Excel.dll");
            var asm = Assembly.LoadFrom(excelDll);
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
            dynamic toolHost = Activator.CreateInstance(hostType, new object[] { excel, null, "csprobe", "csprobe-run", "C#探针" });
            Mark("HOST-CREATED");

            var scriptOpts = toolHost.ScriptOptions;
            scriptOpts.Enabled = true;
            var registry = toolHost.CreateRegistry();
            Mark("REGISTRY");

            dynamic tool = null;
            foreach (var t in registry.Tools)
            {
                if ((string)t.Name == "run_vba_snippet") { tool = t; break; }
            }
            if (tool == null) { Mark("TOOL-NOT-FOUND"); return; }
            Mark("TOOL-FOUND");

            var args = new Dictionary<string, object>();
            args["snippet"] = "Sub TT通道()\n    Dim v As Variant\n    v = XaiArg(\"column\")\n    XaiReturn Array(v, 42)\nEnd Sub";
            args["macro_name"] = "TT通道";
            var argd = new Dictionary<string, object>();
            argd["column"] = "A";
            args["args"] = argd;

            Mark("CASE-START");
            var res = tool.Execute(args);
            Mark("CASE-EXECUTED");
            Mark("ok=" + res.Ok + " code=" + res.ErrorCode);
            Mark("data=" + Convert.ToString(res.Data));
            Mark("DONE");
        }
        catch (Exception ex)
        {
            Mark("FAIL " + ex.ToString());
        }
    }
}
