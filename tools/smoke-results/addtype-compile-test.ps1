$ErrorActionPreference = "Continue"
$out = "d:\execl1.0\tools\smoke-results\addtype-compile-out.txt"
"BOOT" | Out-File $out -Encoding utf8
function Mark($m) { Add-Content $out $m }

try {
    Add-Type -TypeDefinition @"
using System;
public static class CSharpRunProbe
{
    public static string RunDirect(dynamic app, string probeFQ, string proc)
    {
        object r = app.Run(probeFQ, proc);
        return Convert.ToString(r);
    }
}
"@ -ReferencedAssemblies @("Microsoft.CSharp.dll") -ErrorAction Stop
    Mark "ADDTYPE-OK"
} catch {
    Mark ("ADDTYPE-FAIL " + $_.Exception.ToString())
}
Mark "DONE"
