# Focused real-machine check for F1: large integer in VBA args must NOT overflow (CLng).
# Mirrors DiagnosticSmokeRunner order: warm-up simple VBA run first (compiles the scratch
# module), then run the large-int snippet. Asserts XaiArg reads 3000000000 without VBA error 6.
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

    $excelAsm = [System.Reflection.Assembly]::LoadFrom($excelDll)
    $hostType = $excelAsm.GetType("ExcelAiAssistant.Excel.ExcelToolHost", $true)
    $toolHost = [Activator]::CreateInstance($hostType, @($excel, $null, "largeint", "largeint-run", "大整数"))
    $scriptOpts = $toolHost.GetType().GetProperty("ScriptOptions").GetValue($toolHost, $null)
    $scriptOpts.GetType().GetProperty("Enabled").SetValue($scriptOpts, $true, $null)
    $registry = $toolHost.CreateRegistry()

    $tool = $null
    $tools = $registry.GetType().GetProperty("Tools").GetValue($registry, $null)
    foreach ($t in $tools) {
        $name = $t.GetType().GetProperty("Name").GetValue($t, $null)
        if ($name -eq "run_vba_snippet") { $tool = $t; break }
    }
    if ($tool -eq $null) { Write-Output "run_vba_snippet not found"; exit 1 }

    # Warm-up: simple snippet first (mirrors CheckVba before CheckVbaDataChannel).
    $warmArgs = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $warmArgs["snippet"] = "Sub TT预热()`n    Range(""H1"").Value = ""ok""`nEnd Sub"
    $warmArgs["macro_name"] = "TT预热"
    $warm = $tool.Execute.Invoke($warmArgs)
    $warmOk = $warm.GetType().GetProperty("Ok").GetValue($warm, $null)
    Write-Output ("warmup ok=" + $warmOk)

    # Control: verify-p1-equivalent data channel (string arg). Should pass if harness is fine.
    $ctlArgs = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $ctlArgs["snippet"] = "Sub TT通道()`n    Dim v As Variant`n    v = XaiArg(""column"")`n    XaiReturn Array(v, 42)`nEnd Sub"
    $ctlArgs["macro_name"] = "TT通道"
    $ctlInner = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $ctlInner["column"] = "A"
    $ctlArgs["args"] = $ctlInner
    $ctl = $tool.Execute.Invoke($ctlArgs)
    Write-Output ("C1 control execute returned, ok=" + $ctl.GetType().GetProperty("Ok").GetValue($ctl, $null) + " code=" + $ctl.GetType().GetProperty("ErrorCode").GetValue($ctl, $null))

    $toolArgs = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $toolArgs["snippet"] = "Sub TT大数()`n    Dim id As Variant`n    id = XaiArg(""id"")`n    XaiReturn Array(id, 42)`nEnd Sub"
    $toolArgs["macro_name"] = "TT大数"
    $inner = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $inner["id"] = 3000000000
    $toolArgs["args"] = $inner

    Write-Output ("L0 before largeint execute")
    $res = $tool.Execute.Invoke($toolArgs)
    Write-Output ("M1 execute returned")
    $ok = $res.GetType().GetProperty("Ok").GetValue($res, $null)
    $code = $res.GetType().GetProperty("ErrorCode").GetValue($res, $null)
    $data = $res.GetType().GetProperty("Data").GetValue($res, $null)
    Write-Output ("M2 ok/code/data read: ok=" + $ok + " code=" + $code)
    Write-Output ("M3 data raw: " + ($data | Out-String -Width 200))

    $pass = $false
    if ($ok) {
        # 直接查 result 数组元素，避免 ConvertTo-Json 解析 JToken 的潜在卡点。
        $result = $data.GetType().GetProperty("result").GetValue($data, $null)
        Write-Output ("M4 result read: " + ($result | Out-String -Width 200))
        if ($null -ne $result) {
            $s0 = [System.Convert]::ToString($result[0])
            $pass = $s0 -match "3000000000"
        }
    }
    Write-Output ("LARGEINT_CHECK=" + $(if ($pass) { "PASS" } else { "FAIL" }))
    if (-not $pass) { exit 1 }
} catch {
    Write-Output ("FAIL exception " + $_.Exception.ToString())
    exit 1
} finally {
    if ($excel -ne $null) {
        try { $excel.DisplayAlerts = $false } catch {}
        try { $excel.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null } catch {}
    }
}
exit 0
