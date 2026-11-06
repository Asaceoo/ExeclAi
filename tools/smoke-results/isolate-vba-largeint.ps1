# Isolate which component of the VBA data channel triggers RPC 0x800706BE with large integers.
# Cases:
#   A. args={"column":"A"}, read XaiArg + XaiReturn Array(v,42)   -> control (expected PASS)
#   B. args={"id":3000000000}, read XaiArg + XaiReturn Array(v,42) -> F1 target
#   C. no args, XaiReturn Array(3000000000, 42) (hardcoded big)    -> serialization only
#   D. args={"id":3000000000}, XaiReturn Array(42) (no XaiArg read) -> args injection only
$ErrorActionPreference = "Stop"
$marker = "d:\execl1.0\tools\smoke-results\isolate-marker.txt"
"BOOT" | Out-File $marker -Encoding utf8

function Mark($m) { Add-Content $marker $m }

$dir = "d:\execl1.0\src\ExcelAiAssistant.Addin\bin\x64\Release\net48\win-x64"
$addinDll = Join-Path $dir "ExcelAiAssistant.Addin.dll"
if (-not (Test-Path $addinDll)) { Mark ("MISSING " + $addinDll); exit 1 }
Mark "DLL-OK"

$excel = $null
try {
    [System.Reflection.Assembly]::LoadFrom($addinDll) | Out-Null
    Mark "LOAD-ADDIN"
    $excelDll = Join-Path $dir "ExcelAiAssistant.Excel.dll"
    [System.Reflection.Assembly]::LoadFrom($excelDll) | Out-Null
    Mark "LOAD-EXCEL"

    $excel = New-Object -ComObject Excel.Application
    Mark "COM-CREATED"
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    try { $excel.AutomationSecurity = 1; Mark "AUTOSEC-SET-1" } catch { Mark ("AUTOSEC-FAIL " + $_.Exception.Message) }
    Start-Sleep -Seconds 2

    $book = $excel.Workbooks.Add()
    Mark "BOOK-ADDED"
    $book.Activate()

    # RAW probe: bypass the tool entirely; inject a self-contained macro (no external refs) and call Application.Run to see if macros run at all.
    try {
        $vbp = $excel.VBE.ActiveVBProject
        $comps = $vbp.VBComponents
        $mod = $comps.Add(1)
        $cm = $mod.CodeModule
        $cm.AddFromString("Function RawProbeBE() As Long`n    RawProbeBE = 6 * 7`nEnd Function")
        Mark "RAW-INJECTED"
        $raw = $excel.Run("RawProbeBE")
        Mark ("RAW-RUN-OK " + ($raw | Out-String -Width 120))
        $comps.Remove($mod)
    } catch {
        Mark ("RAW-RUN-FAIL " + $_.Exception.Message)
    }

    # TOOL-LIKE RAW: reproduce the tool's exact injection (snippet + helperBlock + probe) via reflection,
    # then call Application.Run directly to see the RAW exception/HRESULT that the tool misclassifies.
    try {
        $excelAsm0 = [System.Reflection.Assembly]::LoadFrom($excelDll)
        $runnerType = $excelAsm0.GetType("ExcelAiAssistant.Excel.ScriptRunnerService")
        $buildM = $runnerType.GetMethod("BuildVbaHelperBlock", [System.Reflection.BindingFlags]"Static,NonPublic")
        $argsJson0 = '{"column":"A"}'
        $helperBlock0 = $buildM.Invoke($null, @($argsJson0, 1048576))
        Mark "HELPER-BUILT len=" + $helperBlock0.Length

        $comps2 = $excel.VBE.ActiveVBProject.VBComponents
        $mod2 = $comps2.Add(1)
        $cm2 = $mod2.CodeModule
        $cm2.AddFromString("Sub TT通道()`n    Dim v As Variant`n    v = XaiArg(""column"")`n    XaiReturn Array(v, 42)`nEnd Sub")
        $cm2.AddFromString($helperBlock0)
        $mname2 = $cm2.Parent.Name
        Mark ("MOD2-NAME " + $mname2)
        $raw2 = $excel.Run($mname2 + ".XaiProbe9F2E", "TT通道")
        Mark ("TOOLRAW-RUN-OK " + ($raw2 | Out-String -Width 200))
        $comps2.Remove($mod2)
    } catch {
        Mark ("TOOLRAW-RUN-FAIL " + $_.Exception.GetType().FullName + " :: " + $_.Exception.Message)
        if ($_.Exception.InnerException) { Mark ("  INNER " + $_.Exception.InnerException.Message) }
    }



    $excelAsm = [System.Reflection.Assembly]::LoadFrom($excelDll)
    $hostType = $excelAsm.GetType("ExcelAiAssistant.Excel.ExcelToolHost", $true)
    Mark "HOST-TYPE"
    $toolHost = [Activator]::CreateInstance($hostType, @($excel, $null, "isolate", "isolate-run", "隔离"))
    Mark "HOST-CREATED"
    $scriptOpts = $toolHost.GetType().GetProperty("ScriptOptions").GetValue($toolHost, $null)
    $scriptOpts.GetType().GetProperty("Enabled").SetValue($scriptOpts, $true, $null)
    Mark "SCRIPT-ENABLED"
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
        Mark ("  data=" + ($data | Out-String -Width 250))
        return $ok
    }

    $mk = { param($d) $x = New-Object 'System.Collections.Generic.Dictionary[string,object]'; foreach ($k in $d.Keys) { $x[$k] = $d[$k] }; return $x }

    # Case A: control
    $innerA = & $mk @{ column = "A" }
    Run-Case "A" "Sub TT通道()`n    Dim v As Variant`n    v = XaiArg(""column"")`n    XaiReturn Array(v, 42)`nEnd Sub" "TT通道" $innerA | Out-Null

    # Case B: large int via XaiArg
    $innerB = & $mk @{ id = 3000000000 }
    Run-Case "B" "Sub TT大数()`n    Dim id As Variant`n    id = XaiArg(""id"")`n    XaiReturn Array(id, 42)`nEnd Sub" "TT大数" $innerB | Out-Null

    # Case C: hardcoded large value serialization (no args)
    Run-Case "C" "Sub TT序列()`n    XaiReturn Array(3000000000, 42)`nEnd Sub" "TT序列" $null | Out-Null

    # Case D: large int in args but NOT read (injection only)
    Run-Case "D" "Sub TT注入()`n    XaiReturn Array(42)`nEnd Sub" "TT注入" $innerB | Out-Null

} catch {
    Mark ("FAIL exception " + $_.Exception.ToString())
    exit 1
} finally {
    if ($excel -ne $null) {
        try { $excel.DisplayAlerts = $false } catch {}
        try { $excel.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null } catch {}
    }
}
Mark "DONE"
exit 0
