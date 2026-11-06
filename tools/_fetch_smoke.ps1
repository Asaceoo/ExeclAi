$ErrorActionPreference = "Stop"
$excelDll = "d:\execl1.0\src\ExcelAiAssistant.Excel\bin\Release\net48\ExcelAiAssistant.Excel.dll"
$coreDll  = "d:\execl1.0\src\ExcelAiAssistant.Core\bin\Release\net48\ExcelAiAssistant.Core.dll"
[void][System.Reflection.Assembly]::LoadFrom($coreDll)
$excelAsm = [System.Reflection.Assembly]::LoadFrom($excelDll)

Write-Output "===== 1) Sanitize mask check (reflection, internal WebFetchService.Sanitize) ====="
$svc = $excelAsm.GetType("ExcelAiAssistant.Excel.WebFetchService")
$san = $svc.GetMethod("Sanitize", [System.Reflection.BindingFlags]"NonPublic,Static")
$in = "Contact alice@example.com Phone 13812345678 ID 110105199003078890 keep-this."
$out = $san.Invoke($null, @($in))
Write-Output ("IN : " + $in)
Write-Output ("OUT: " + $out)
if ($out -match "example\.com|13812345678|19900307889") { Write-Output ">>> SANITIZE FAIL"; exit 1 } else { Write-Output ">>> SANITIZE OK" }

Write-Output ""
Write-Output "===== 2) Protocol block (file://, must reject without network) ====="
$ho = New-Object -TypeName "ExcelAiAssistant.Excel.ExcelToolHost" -ArgumentList (,[object]::new())
$reg = $ho.CreateRegistry()
$desc = $null
if (-not $reg.TryGet("fetch_url", [ref]$desc)) { Write-Output ">>> fetch_url NOT REGISTERED"; exit 1 }
$bad = New-Object "System.Collections.Generic.Dictionary[string,object]"
$bad.Add("url", "file:///c:/windows/win.ini")
$br = $desc.Execute.Invoke($bad)
Write-Output ("file:// -> Ok=$($br.Ok) ErrorCode=$($br.ErrorCode) Data=$($br.Data)")
if ($br.Ok -or ($br.ErrorCode -ne "INVALID_URL")) { Write-Output ">>> PROTOCOL BLOCK FAIL"; exit 1 } else { Write-Output ">>> PROTOCOL BLOCK OK" }

Write-Output ""
Write-Output "===== 3) Real fetch https://www.baidu.com (sanitize on, size/time limits) ====="
$dict = New-Object "System.Collections.Generic.Dictionary[string,object]"
$dict.Add("url", "https://www.baidu.com")
$res = $desc.Execute.Invoke($dict)
Write-Output ("Ok=$($res.Ok) ErrorCode=$($res.ErrorCode)")
if ($res.Ok) {
    $data = [string]$res.Data
    Write-Output ("Returned text length: " + $data.Length)
    Write-Output ("Text sample: " + $data.Substring(0, [Math]::Min(300, $data.Length)).Replace("`n", " "))
} else {
    Write-Output ("Data: " + $res.Data)
    exit 1
}
Write-Output ""
Write-Output "===== SMOKE DONE ====="