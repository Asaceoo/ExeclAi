# Variant: construct ExcelToolHost via New-Object (native PS binder) instead of Activator.CreateInstance.
$ErrorActionPreference = "Stop"
$out = "d:\execl1.0\tools\smoke-results\newobj-out.txt"
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

    $excelAsm = [System.Reflection.Assembly]::LoadFrom($excelDll)
    $hostType = $excelAsm.GetType("ExcelAiAssistant.Excel.ExcelToolHost", $true)
    $toolHost = New-Object $hostType($excel, $null, "newobj", "newobj-run", "新建")
    Mark "HOST-NEWOBJECT"

    # inspect private application field type
    try {
        $f = $hostType.GetField("application", [System.Reflection.BindingFlags]"Instance,NonPublic")
        $appField = $f.GetValue($toolHost)
        Mark ("  appFieldType=" + $appField.GetType().FullName + " asm=" + $appField.GetType().Assembly.GetName().Name)
    } catch { Mark ("  field-inspect-fail " + $_.Exception.Message) }

    $scriptOpts = $toolHost.GetType().GetProperty("ScriptOptions").GetValue($toolHost, $null)
    $scriptOpts.GetType().GetProperty("Enabled").SetValue($scriptOpts, $true, $null)
    $registry = $toolHost.CreateRegistry()
    Mark "REGISTRY"
    $tool = $null
    $tools = $registry.GetType().GetProperty("Tools").GetValue($registry, $null)
    foreach ($t in $tools) {
        $name = $t.GetType().GetProperty("Name").GetValue($t, $null)
        if ($name -eq "run_vba_snippet") { $tool = $t; break }
    }
    if ($tool -eq $null) { Mark "TOOL-NOT-FOUND"; exit 1 }
    Mark "TOOL-FOUND"

    $a = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $a["snippet"] = "Sub TT通道()`n    Dim v As Variant`n    v = XaiArg(""column"")`n    XaiReturn Array(v, 42)`nEnd Sub"
    $a["macro_name"] = "TT通道"
    $argd = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $argd["column"] = "A"
    $a["args"] = $argd
    Mark "CASE-START A"
    $res = $tool.Execute.Invoke($a)
    Mark "CASE-EXECUTED A"
    $ok = $res.GetType().GetProperty("Ok").GetValue($res, $null)
    $code = $res.GetType().GetProperty("ErrorCode").GetValue($res, $null)
    $data = $res.GetType().GetProperty("Data").GetValue($res, $null)
    Mark ("CASE A => ok=" + $ok + " code=" + $code)
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
