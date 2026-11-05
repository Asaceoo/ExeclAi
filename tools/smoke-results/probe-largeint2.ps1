# Focused probe: dump raw result JSON for the large-int VBA data channel path.
$ErrorActionPreference = "Stop"
$dir = "d:\execl1.0\src\ExcelAiAssistant.Addin\bin\x64\Release\net48\win-x64"
$addinDll = Join-Path $dir "ExcelAiAssistant.Addin.dll"
$excelDll = Join-Path $dir "ExcelAiAssistant.Excel.dll"
$excel = $null
try {
    [System.Reflection.Assembly]::LoadFrom($addinDll) | Out-Null
    [System.Reflection.Assembly]::LoadFrom($excelDll) | Out-Null
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    Start-Sleep -Seconds 2
    $book = $excel.Workbooks.Add()
    $book.Activate()

    $excelAsm = [System.Reflection.Assembly]::LoadFrom($excelDll)
    $hostType = $excelAsm.GetType("ExcelAiAssistant.Excel.ExcelToolHost", $true)
    $toolHost = [Activator]::CreateInstance($hostType, @($excel, $null, "probe", "probe-run", "探针"))
    $scriptOpts = $toolHost.GetType().GetProperty("ScriptOptions").GetValue($toolHost, $null)
    $scriptOpts.GetType().GetProperty("Enabled").SetValue($scriptOpts, $true, $null)
    $registry = $toolHost.CreateRegistry()
    $tool = $null
    foreach ($t in $registry.GetType().GetProperty("Tools").GetValue($registry, $null)) {
        if ($t.GetType().GetProperty("Name").GetValue($t, $null) -eq "run_vba_snippet") { $tool = $t; break }
    }
    if ($tool -eq $null) { Write-Output "tool not found"; exit 1 }

    # warm-up
    $warm = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $warm["snippet"] = "Sub TT预热()`n    Range(""H1"").Value = ""ok""`nEnd Sub"
    $warm["macro_name"] = "TT预热"
    $tool.Execute.Invoke($warm) | Out-Null

    $args = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $args["snippet"] = "Sub TT大数()`n    Dim id As Variant`n    id = XaiArg(""id"")`n    XaiReturn Array(id, 42)`nEnd Sub"
    $args["macro_name"] = "TT大数"
    $inner = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $inner["id"] = 3000000000
    $args["args"] = $inner

    $res = $tool.Execute.Invoke($args)
    $ok = $res.GetType().GetProperty("Ok").GetValue($res, $null)
    $code = $res.GetType().GetProperty("ErrorCode").GetValue($res, $null)
    $data = $res.GetType().GetProperty("Data").GetValue($res, $null)
    Write-Output ("ok=" + $ok + " code=" + $code)
    if ($data -ne $null) {
        $result = $data.GetType().GetProperty("result").GetValue($data, $null)
        if ($result -ne $null) {
            Write-Output ("result JSON: " + $result.ToString())
            Write-Output ("result type: " + $result.GetType().FullName)
            $arr = [Newtonsoft.Json.Linq.JArray]$result
            for ($i = 0; $i -lt $arr.Count; $i++) {
                $el = $arr[$i]
                Write-Output ("  [" + $i + "] type=" + $el.Type + " value=" + $el.ToString() + " raw=" + ($el | Out-String).Trim())
            }
        } else {
            Write-Output "result is null"
        }
    }
} catch {
    Write-Output ("FAIL " + $_.Exception.ToString())
} finally {
    if ($excel -ne $null) {
        try { $excel.DisplayAlerts = $false } catch {}
        try { $excel.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null } catch {}
    }
}
exit 0
