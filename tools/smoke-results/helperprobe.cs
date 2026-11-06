using System;
using System.Reflection;

class HelperProbe
{
    static void Main()
    {
        var asm = Assembly.LoadFrom(@"d:\execl1.0\src\ExcelAiAssistant.Excel\bin\Release\net48\ExcelAiAssistant.Excel.dll");
        var t = asm.GetType("ExcelAiAssistant.Excel.ScriptRunnerService", true);
        var m = t.GetMethod("BuildVbaHelperBlock", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.Public);
        var helper = (string)m.Invoke(null, new object[] { "{}", 1024 });
        var idx = helper.IndexOf("Case ", StringComparison.Ordinal);
        while (idx >= 0)
        {
            var lineEnd = helper.IndexOf("\n", idx, StringComparison.Ordinal);
            var line = helper.Substring(idx, lineEnd - idx).Trim();
            Console.WriteLine("LINE: [" + line + "]");
            var next = helper.IndexOf("Case ", idx + 1, StringComparison.Ordinal);
            if (next == idx) break;
            idx = next;
        }
        var m2 = asm.GetType("ExcelAiAssistant.Excel.ScriptRunnerService", true).GetMethod("VbaStringLiteral", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.Public);
        Console.WriteLine("VbaStringLiteral(\"a\\\"b\")=" + m2.Invoke(null, new object[] { "a\"b" }));
    }
}
