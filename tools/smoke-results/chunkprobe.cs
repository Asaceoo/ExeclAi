using System;
using System.Collections.Generic;
using System.Reflection;

class ChunkProbe
{
    static string OutFile = @"d:\execl1.0\tools\smoke-results\chunkprobe-out.txt";

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
            dynamic toolHost = Activator.CreateInstance(hostType, new object[] { excel, null, "chunk", "chunk-run", "分块" });
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

            // Crafted args: {"k":"<27 A's>\"<60 B's>"} -> escaped "" pair straddles 100-char chunk boundary (brute-force confirmed)
            var expected = new string('A', 27) + "\"" + new string('B', 60);
            var argsJson = "{\"k\":\"" + expected + "\"}";
            Mark("ARGS-JSON-LEN " + argsJson.Length);
            Mark("EXPECTED " + expected);

            var args = new Dictionary<string, object>();
            args["snippet"] = "Sub TT通道()\n    Dim v As Variant\n    v = XaiArg(\"k\")\n    XaiReturn v\nEnd Sub";
            args["macro_name"] = "TT通道";
            var argd = new Dictionary<string, object>();
            argd["k"] = expected;
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
