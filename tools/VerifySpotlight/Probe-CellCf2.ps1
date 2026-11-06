# 探针 2：确认「单规则 CELL() CF + 重算触发器」组合 —— 选触发器类型（ws.Calculate vs app.Calculate）。
$ErrorActionPreference = "Stop"
$app = $null
for ($i = 0; $i -lt 40; $i++) {
    try { $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application"); break } catch { Start-Sleep -Seconds 1 }
}
if ($null -eq $app) { Write-Host "NO EXCEL INSTANCE"; exit 2 }

$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
for ($r = 1; $r -le 6; $r++) { for ($c = 1; $c -le 6; $c++) { $ws.Cells.Item($r, $c).Value2 = $r * 10 + $c } }

$wb.Names.Add('JYYJ_VIEW_MODE', '=1') | Out-Null
$fc = $ws.Range("A1:F6").FormatConditions.Add(2, 0, '=AND(JYYJ_VIEW_MODE,OR(CELL("row")=ROW(),CELL("col")=COLUMN()))')
$fc.Interior.ColorIndex = 6
$fc.Interior.Pattern = 1

function CF([int]$r, [int]$c) { return [int]$ws.Cells.Item($r, $c).DisplayFormat.Interior.ColorIndex }

$app.Range("B2").Select()
Start-Sleep -Milliseconds 300
Write-Host ("[a] B2 未重算: A2=" + (CF 2 1) + " (期望 -4142，确认 CF 不自动渲染)")

$ws.Calculate()
Start-Sleep -Milliseconds 500
Write-Host ("[b] B2 + ws.Calculate: A2=" + (CF 2 1) + " 期望6, B1=" + (CF 1 2) + " 期望6, A1=" + (CF 1 1) + " 期望-4142")

$app.Range("D4").Select()
$ws.Calculate()
Start-Sleep -Milliseconds 500
Write-Host ("[c] D4 + ws.Calculate: A4=" + (CF 4 1) + " 期望6, D1=" + (CF 1 4) + " 期望6, A2=" + (CF 2 1) + " 期望-4142")

$app.Range("F6").Select()
$app.Calculate()
Start-Sleep -Milliseconds 500
Write-Host ("[d] F6 + app.Calculate: F6=" + (CF 6 6) + " 期望6, F1=" + (CF 1 6) + " 期望6, A4=" + (CF 4 1) + " 期望-4142")

# 关键：ws.Calculate 是否足够（避免每次整簿重算）
$app.Range("E5").Select()
$ws.Calculate()
Start-Sleep -Milliseconds 500
Write-Host ("[e] E5 + ws.Calculate: E5=" + (CF 5 5) + " 期望6, E1=" + (CF 1 5) + " 期望6, F6=" + (CF 6 6) + " 期望-4142")

$wb.Close($false)
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws) | Out-Null
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null
