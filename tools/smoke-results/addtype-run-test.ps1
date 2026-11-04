# Isolate whether the C# dynamic binder's Application.Run call itself (not the surrounding tool logic)
# is what hangs/fails with RPC 0x800706BE out-of-proc. Uses Add-Type to compile a minimal C# type
# that performs exactly application.Run(probeFQ, proc), then drives it against the same workbook
# where the PowerShell direct call succeeds.
$ErrorActionPreference = "Stop"
$out = "d:\execl1.0\tools\smoke-results\addtype-out.txt"
"BOOT" | Out-File $out -Encoding utf8
function Mark($m) { Add-Content $out $m }

$dir = "d:\execl1.0\src\ExcelAiAssistant.Addin\bin\x64\Release\net48\win-x64"
$excelDll = Join-Path $dir "ExcelAiAssistant.Excel.dll"
[System.Reflection.Assembly]::LoadFrom($excelDll) | Out-Null
$asm = [System.Reflection.Assembly]::LoadFrom($excelDll)
$t = $asm.GetType("ExcelAiAssistant.Excel.ScriptRunnerService")
$m = $t.GetMethod("BuildVbaHelperBlock", [System.Reflection.BindingFlags]"Static,NonPublic")
$helper = [string]$m.Invoke($null, @('{"column":"A"}', 1048576))
Mark ("HELPER-LEN " + $helper.Length)

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
"@ -ReferencedAssemblies @("Microsoft.CSharp.dll")
Mark "ADDTYPE-OK"

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    Start-Sleep -Seconds 2
    $book = $excel.Workbooks.Add()
    $book.Activate()
    $vbp = $excel.VBE.ActiveVBProject
    $comps = $vbp.VBComponents
    $mod = $comps.Add(1)
    try { $mod.Name = "JYYJHelper" } catch {}
    $cm = $mod.CodeModule
    $cm.AddFromString("Sub TT通道()`n    Dim v As Variant`n    v = XaiArg(""column"")`n    XaiReturn Array(v, 42)`nEnd Sub")
    $cm.AddFromString($helper)
    Mark "INJECTED"

    # Baseline: PowerShell direct call (known working)
    try {
        $r = $excel.Run("JYYJHelper.XaiProbe9F2E", "JYYJHelper.TT通道")
        Mark ("PS-DIRECT OK => " + $r)
    } catch {
        Mark ("PS-DIRECT FAIL => " + $_.Exception.Message)
    }

    # C# dynamic binder call, same arguments
    try {
        $r = [CSharpRunProbe]::RunDirect($excel, "JYYJHelper.XaiProbe9F2E", "JYYJHelper.TT通道")
        Mark ("CS-DYNAMIC OK => " + $r)
    } catch {
        Mark ("CS-DYNAMIC FAIL => " + $_.Exception.GetType().FullName + " :: " + $_.Exception.Message)
        if ($_.Exception.InnerException) { Mark ("  INNER " + $_.Exception.InnerException.GetType().FullName + " :: " + $_.Exception.InnerException.Message) }
    }

    Mark "DONE"
} catch {
    Mark ("FAIL " + $_.Exception.ToString())
    exit 1
} finally {
    if ($excel -ne $null) {
        try { $excel.DisplayAlerts = $false } catch {}
        try { $excel.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null } catch {}
    }
}
exit 0
