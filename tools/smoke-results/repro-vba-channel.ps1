# Isolated repro for the VBA data channel (CheckVbaDataChannel) RPC failure.
# Drives ExcelToolHost.run_vba_snippet directly with the TT通道 snippet + args,
# WITHOUT running CheckVba first, to isolate whether the failure is state leakage
# from a prior VBA run or intrinsic to the data channel.
$ErrorActionPreference = "Stop"

$dir = "d:\execl1.0\src\ExcelAiAssistant.Addin\bin\x64\Release\net48\win-x64"
$addinDll = Join-Path $dir "ExcelAiAssistant.Addin.dll"
if (-not (Test-Path $addinDll)) { Write-Output ("MISSING " + $addinDll); exit 1 }

$excel = $null
try {
    [System.Reflection.Assembly]::LoadFrom($addinDll) | Out-Null
    $excelDll = Join-Path $dir "ExcelAiAssistant.Excel.dll"
    [System.Reflection.Assembly]::LoadFrom($excelDll) | Out-Null
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    Start-Sleep -Seconds 2

    $book = $excel.Workbooks.Add()
    $book.Activate()

    $asm = [System.Reflection.Assembly]::LoadFrom($addinDll)
    $excelAsm = [System.Reflection.Assembly]::LoadFrom($excelDll)
    $hostType = $excelAsm.GetType("ExcelAiAssistant.Excel.ExcelToolHost", $true)
    Write-Output "STEP1 hostType loaded"
    $toolHost = [Activator]::CreateInstance($hostType, @($excel, $null, "repro", "repro-run", "复现"))
    Write-Output "STEP2 host created"
    $scriptOpts = $toolHost.GetType().GetProperty("ScriptOptions").GetValue($toolHost, $null)
    $scriptOpts.GetType().GetProperty("Enabled").SetValue($scriptOpts, $true, $null)
    Write-Output "STEP3 script options enabled"
    $registry = $toolHost.CreateRegistry()
    Write-Output "STEP4 registry created"

    $tool = $null
    $tools = $registry.GetType().GetProperty("Tools").GetValue($registry, $null)
    foreach ($t in $tools) {
        $name = $t.GetType().GetProperty("Name").GetValue($t, $null)
        if ($name -eq "run_vba_snippet") { $tool = $t; break }
    }
    if ($tool -eq $null) { Write-Output "run_vba_snippet not found"; exit 1 }
    Write-Output "STEP5 tool found"

    $toolArgs = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $toolArgs["snippet"] = "Sub TT通道()`n    Dim v As Variant`n    v = XaiArg(""column"")`n    XaiReturn Array(v, 42)`nEnd Sub"
    $toolArgs["macro_name"] = "TT通道"
    $inner = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $inner["column"] = "A"
    $toolArgs["args"] = $inner
    Write-Output "STEP6 calling Execute"
    $res = $tool.Execute.Invoke($toolArgs)
    Write-Output "STEP7 Execute returned"
    $ok = $res.GetType().GetProperty("Ok").GetValue($res, $null)
    $code = $res.GetType().GetProperty("ErrorCode").GetValue($res, $null)
    $data = $res.GetType().GetProperty("Data").GetValue($res, $null)
    Write-Output ("ok=" + $ok + " code=" + $code)
    Write-Output ("data=" + ($data | Out-String))
} catch {
    Write-Output ("FAIL exception " + $_.Exception.ToString())
} finally {
    if ($excel -ne $null) {
        try { $excel.DisplayAlerts = $false } catch {}
        try { $excel.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null } catch {}
    }
}
