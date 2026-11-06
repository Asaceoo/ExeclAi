# STATELESS harness: create Excel + workbook + host + tool only.
# NO raw VBE COM access before tool calls, so no shared RCW to be destroyed by the tool's
# finally-block FinalReleaseComObject. Tests whether the RPC 0x800706BE is a harness artifact.
$ErrorActionPreference = "Stop"
$out = "d:\execl1.0\tools\smoke-results\clean-out.txt"
"BOOT" | Out-File $out -Encoding utf8
function Mark($m) { Add-Content $out $m }

$dir = "d:\execl1.0\src\ExcelAiAssistant.Addin\bin\x64\Release\net48\win-x64"
$addinDll = Join-Path $dir "ExcelAiAssistant.Addin.dll"
$excelDll = Join-Path $dir "ExcelAiAssistant.Excel.dll"
if (-not (Test-Path $excelDll)) { Mark "MISSING-DLL"; exit 1 }

$excel = $null
try {
    [System.Reflection.Assembly]::LoadFrom($addinDll) | Out-Null
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
    $toolHost = [Activator]::CreateInstance($hostType, @($excel, $null, "clean", "clean-run", "干净"))
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

    function Run-Case($label, $snippet, $macro, $argsDict) {
        Mark ("CASE-START " + $label)
        $a = New-Object 'System.Collections.Generic.Dictionary[string,object]'
        $a["snippet"] = $snippet
        $a["macro_name"] = $macro
        if ($null -ne $argsDict) { $a["args"] = $argsDict }
        $res = $tool.Execute.Invoke($a)
        Mark ("CASE-EXECUTED " + $label)
        $ok = $res.GetType().GetProperty("Ok").GetValue($res, $null)
        $code = $res.GetType().GetProperty("ErrorCode").GetValue($res, $null)
        $data = $res.GetType().GetProperty("Data").GetValue($res, $null)
        Mark ("CASE " + $label + " => ok=" + $ok + " code=" + $code)
        Mark ("  data=" + ($data | Out-String -Width 300))
        return $ok
    }

    $mk = { param($d) $x = New-Object 'System.Collections.Generic.Dictionary[string,object]'; foreach ($k in $d.Keys) { $x[$k] = $d[$k] }; return $x }

    Run-Case "A" "Sub TT通道()`n    Dim v As Variant`n    v = XaiArg(""column"")`n    XaiReturn Array(v, 42)`nEnd Sub" "TT通道" (& $mk @{ column = "A" }) | Out-Null
    Run-Case "A2-ASCII" "Sub TTcaseA()`n    Dim v As Variant`n    v = XaiArg(""column"")`n    XaiReturn Array(v, 42)`nEnd Sub" "TTcaseA" (& $mk @{ column = "A" }) | Out-Null
    Run-Case "B" "Sub TT大数()`n    Dim id As Variant`n    id = XaiArg(""id"")`n    XaiReturn Array(id, 42)`nEnd Sub" "TT大数" (& $mk @{ id = 3000000000 }) | Out-Null
    Run-Case "B2-ASCII" "Sub TTbig()`n    Dim id As Variant`n    id = XaiArg(""id"")`n    XaiReturn Array(id, 42)`nEnd Sub" "TTbig" (& $mk @{ id = 3000000000 }) | Out-Null
    Run-Case "C" "Sub TT序列()`n    XaiReturn Array(3000000000, 42)`nEnd Sub" "TT序列" $null | Out-Null
    Run-Case "C2-ASCII" "Sub TTser()`n    XaiReturn Array(3000000000, 42)`nEnd Sub" "TTser" $null | Out-Null
    Run-Case "D" "Sub TT注入()`n    XaiReturn Array(42)`nEnd Sub" "TT注入" (& $mk @{ id = 3000000000 }) | Out-Null

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
