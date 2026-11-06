# Experiment A: call ScriptRunnerService.RunVbaSnippet DIRECTLY (bypass ExcelToolHost wrapper).
# If this hangs -> problem inside ScriptRunnerService; if works -> problem in host wrapper layer.
$ErrorActionPreference = "Stop"
$out = "d:\execl1.0\tools\smoke-results\directsvc-out.txt"
"BOOT" | Out-File $out -Encoding utf8
function Mark($m) { Add-Content $out $m }

$dir = "d:\execl1.0\src\ExcelAiAssistant.Addin\bin\x64\Release\net48\win-x64"
$excelDll = Join-Path $dir "ExcelAiAssistant.Excel.dll"
if (-not (Test-Path $excelDll)) { Mark "MISSING-DLL"; exit 1 }

$excel = $null
try {
    [System.Reflection.Assembly]::LoadFrom($excelDll) | Out-Null
    Mark "ASM-LOADED"
    $excel = New-Object -ComObject Excel.Application
    Mark "COM-CREATED"
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    Start-Sleep -Seconds 2
    $book = $excel.Workbooks.Add()
    $book.Activate()
    Mark "BOOK-ADDED"

    $asm = [System.Reflection.Assembly]::LoadFrom($excelDll)
    $scriptType = $asm.GetType("ExcelAiAssistant.Excel.ScriptRunnerService", $true)
    $optType = $asm.GetType("ExcelAiAssistant.Excel.ScriptOptions", $true)
    $opts = [Activator]::CreateInstance($optType)
    $opts.GetType().GetProperty("Enabled").SetValue($opts, $true, $null)
    $svc = [Activator]::CreateInstance($scriptType, @($excel, $opts))
    Mark "SERVICE-CREATED"

    $a = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $a["snippet"] = "Sub TT通道()`n    Dim v As Variant`n    v = XaiArg(""column"")`n    XaiReturn Array(v, 42)`nEnd Sub"
    $a["macro_name"] = "TT通道"
    $argd = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $argd["column"] = "A"
    $a["args"] = $argd
    Mark "CALL-START"
    $m = $scriptType.GetMethod("RunVbaSnippet")
    $res = $m.Invoke($svc, @($a))
    Mark "CALL-EXECUTED"
    $ok = $res.GetType().GetProperty("Ok").GetValue($res, $null)
    $code = $res.GetType().GetProperty("ErrorCode").GetValue($res, $null)
    $data = $res.GetType().GetProperty("Data").GetValue($res, $null)
    Mark ("CALL => ok=" + $ok + " code=" + $code)
    Mark ("  data=" + ($data | Out-String -Width 300))
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
