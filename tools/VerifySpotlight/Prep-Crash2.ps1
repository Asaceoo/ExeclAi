# 复现崩溃自愈失败：构造残留文件（含 CELL() 指针公式规则 + 静态规则 + 定义名），重启 Excel 触发 15s 延迟自愈，
# 抓 DBG 日志看 RemoveMarkedRules 究竟删了哪些规则。
$ErrorActionPreference = "Stop"

$app = $null
for ($i = 0; $i -lt 40; $i++) {
    try { $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application"); break } catch { Start-Sleep -Seconds 1 }
}
if ($null -eq $app) { Write-Host "FAIL no ROT"; exit 1 }

$tmp = Join-Path $env:TEMP "jyyj_view_crash2.xlsx"
if (Test-Path $tmp) { Remove-Item $tmp -Force }

$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)

# 规则1：真实指针公式（含 CELL() 易失函数）
$f1 = $ws.Cells.FormatConditions.Add(2, 0, '=AND(JYYJ_VIEW_MODE,CELL("row")=ROW(),CELL("col")=COLUMN())')
$f1.Borders.ColorIndex = 3
$f1.Borders.LineStyle = 1

# 规则2：静态公式
$f2 = $ws.Cells.FormatConditions.Add(2, 0, '=AND(JYYJ_VIEW_MODE,ROW()=5,COLUMN()=5)')
$f2.Interior.ColorIndex = 6

$wb.Names.Add('JYYJ_VIEW_MODE', '=1') | Out-Null
$wb.SaveAs($tmp, 51)
Write-Host "saved $tmp with 2 sentinel rules + name"
Write-Host "FormatConditions.Count=$($ws.Cells.FormatConditions.Count)"
for ($i = 1; $i -le $ws.Cells.FormatConditions.Count; $i++) {
    Write-Host "  rule[$i] Type=$([int]$ws.Cells.FormatConditions.Item($i).Type) Formula1=$([string]$ws.Cells.FormatConditions.Item($i).Formula1)"
}
$wb.Close($false)
Write-Host "DONE - now kill excel and restart with this file"
