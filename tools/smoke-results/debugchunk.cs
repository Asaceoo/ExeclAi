using System;
using System.Reflection;
using System.Text;

class DebugChunk
{
    static string OutFile = @"d:\execl1.0\tools\smoke-results\debugchunk-out.txt";
    static void Mark(string m) { System.IO.File.AppendAllText(OutFile, m + Environment.NewLine); }

    static string DecodeVbaString(string s)
    {
        var sb = new StringBuilder();
        for (int i = 0; i < s.Length; i++)
        {
            if (s[i] == '"' && i + 1 < s.Length && s[i + 1] == '"') { sb.Append('"'); i++; }
            else sb.Append(s[i]);
        }
        return sb.ToString();
    }

    static void Main()
    {
        System.IO.File.WriteAllText(OutFile, "BOOT" + Environment.NewLine);
        var dir = @"d:\execl1.0\src\ExcelAiAssistant.Addin\bin\x64\Release\net48\win-x64";
        var asm = Assembly.LoadFrom(System.IO.Path.Combine(dir, "ExcelAiAssistant.Excel.dll"));
        var t = asm.GetType("ExcelAiAssistant.Excel.ScriptRunnerService");
        var m = t.GetMethod("VbaStringLiteral", BindingFlags.Static | BindingFlags.NonPublic);

        var argsJson = "{\"k\":\"" + new string('A', 26) + "\\\"" + new string('B', 60) + "\"}";
        var lit = (string)m.Invoke(null, new object[] { argsJson });
        Mark("ARGSJSON(" + argsJson.Length + "): " + argsJson);
        Mark("LIT(" + lit.Length + "): " + lit);

        // manual decode
        var sep = "\" & _\n    \"";
        var sb = new StringBuilder();
        int pos = 0;
        int ci = 0;
        while (true)
        {
            int idx = lit.IndexOf(sep, pos, StringComparison.Ordinal);
            if (idx < 0) break;
            var content = lit.Substring(pos + 1, idx - pos - 1);
            Mark("CHUNK" + ci + "(" + content.Length + "): " + content);
            sb.Append(DecodeVbaString(content));
            pos = idx + sep.Length;
            ci++;
        }
        var last = lit.Substring(pos + 1, lit.Length - pos - 2);
        Mark("CHUNK" + ci + "(" + last.Length + "): " + last);
        sb.Append(DecodeVbaString(last));
        Mark("DECODED(" + sb.Length + "): " + sb.ToString());
        Mark("MATCH: " + (sb.ToString() == argsJson));
        Mark("DONE");
    }
}
