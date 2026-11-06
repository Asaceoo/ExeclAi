using System;
using System.Reflection;
using System.Text;

class BruteChunk
{
    static string OutFile = @"d:\execl1.0\tools\smoke-results\brutechunk-out.txt";
    static void Mark(string m) { System.IO.File.AppendAllText(OutFile, m + Environment.NewLine); }

    // Correct VBA string-literal decode: "" -> ", scan char by char
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

    // lit = "c1" & _\n    "c2" ... ; separator starts with ". First chunk strips opening ",
    // middle chunks are plain, last chunk strips closing " (and opening " if single chunk).
    static string DecodeVbaLiteral(string lit)
    {
        var sep = "\" & _\n    \"";
        var sb = new StringBuilder();
        int pos = 0;
        bool first = true;
        while (true)
        {
            int idx = lit.IndexOf(sep, pos, StringComparison.Ordinal);
            if (idx < 0) break;
            int start = first ? pos + 1 : pos;
            sb.Append(DecodeVbaString(lit.Substring(start, idx - start)));
            pos = idx + sep.Length;
            first = false;
        }
        int lastStart = first ? pos + 1 : pos;
        sb.Append(DecodeVbaString(lit.Substring(lastStart, lit.Length - lastStart - 1)));
        return sb.ToString();
    }

    static void Main()
    {
        System.IO.File.WriteAllText(OutFile, "BOOT" + Environment.NewLine);
        try
        {
            var dir = @"d:\execl1.0\src\ExcelAiAssistant.Addin\bin\x64\Release\net48\win-x64";
            var asm = Assembly.LoadFrom(System.IO.Path.Combine(dir, "ExcelAiAssistant.Excel.dll"));
            var t = asm.GetType("ExcelAiAssistant.Excel.ScriptRunnerService");
            var m = t.GetMethod("VbaStringLiteral", BindingFlags.Static | BindingFlags.NonPublic);

            int found = 0;
            // Scan: JSON-like strings with an embedded escaped quote \" at various positions
            for (int aCount = 1; aCount <= 200; aCount++)
            {
                for (int bCount = 1; bCount <= 60; bCount++)
                {
                    // argsJson = {"k":"<A...A>\"<B...B>"}  (Json.NET style: embedded quote escaped as \")
                    var argsJson = "{\"k\":\"" + new string('A', aCount) + "\\\"" + new string('B', bCount) + "\"}";
                    var lit = (string)m.Invoke(null, new object[] { argsJson });
                    var dec = DecodeVbaLiteral(lit);
                    if (dec != argsJson)
                    {
                        Mark("MISMATCH a=" + aCount + " b=" + bCount + " argsJsonLen=" + argsJson.Length);
                        Mark("  ORIG: " + argsJson);
                        Mark("  DEC : " + dec);
                        found++;
                        if (found >= 5) break;
                    }
                }
                if (found >= 5) break;
            }
            Mark("FOUND=" + found);

            // Also scan raw strings with quotes (no JSON escaping) for completeness
            int found2 = 0;
            for (int aCount = 1; aCount <= 200 && found2 < 3; aCount++)
            {
                var text = new string('A', aCount) + "\"" + new string('B', 30);
                var lit = (string)m.Invoke(null, new object[] { text });
                var dec = DecodeVbaLiteral(lit);
                if (dec != text)
                {
                    Mark("RAW-MISMATCH a=" + aCount + " textLen=" + text.Length);
                    found2++;
                }
            }
            Mark("RAW-FOUND=" + found2);
            Mark("DONE");
        }
        catch (Exception ex)
        {
            Mark("FAIL " + ex.ToString());
        }
    }
}
