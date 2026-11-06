# 探针：验证「单条 CF 规则 + CELL("row"/"col") 易失函数」能否不写代码、随选区移动实时高亮。
# 若成立 → 视图组改用单规则方案：零规则churn、零定时器竞态、清理只需删 1 条规则，是最成熟路线。
$ErrorActionPreference = "Stop"
$app = $null
for ($i = 0; $i -lt 40; $i++) {
    try { $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application"); break } catch { Start-Sleep -Seconds 1 }
}
if ($null -eq $app) { Write-Host "NO EXCEL INSTANCE"; exit 2 }

$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
for ($r = 1; $r -le 6; $r++) { for ($c = 1; $c -le 6; $c++) { $ws.Cells.Item($r, $c).Value2 = $r * 10 + $c } }

# 哨兵定义名 + 单条 CF 规则（聚光灯：整行/整列跟随）
$wb.Names.Add('JYYJ_VIEW_MODE', '=1') | Out-Null
$fc = $ws.Range("A1:F6").FormatConditions.Add(2, 0, '=AND(JYYJ_VIEW_MODE,OR(CELL("row")=ROW(),CELL("col")=COLUMN()))')
$fc.Interior.ColorIndex = 6
$fc.Interior.Pattern = 1

$app.Range("B2").Select()
Start-Sleep -Milliseconds 800
Write-Host ("[1] B2: A2=" + [int]$ws.Cells.Item(2,1).DisplayFormat.Interior.ColorIndex + " 期望6, B1=" + [int]$ws.Cells.Item(1,2).DisplayFormat.Interior.ColorIndex + " 期望6, A1=" + [int]$ws.Cells.Item(1,1).DisplayFormat.Interior.ColorIndex + " 期望-4142")

# 不调任何代码，纯移动选区，观察 CELL() 是否触发 CF 重算
$app.Range("D4").Select()
Start-Sleep -Seconds 2
Write-Host ("[2] D4(纯移动): A4=" + [int]$ws.Cells.Item(4,1).DisplayFormat.Interior.ColorIndex + " 期望6, D1=" + [int]$ws.Cells.Item(1,4).DisplayFormat.Interior.ColorIndex + " 期望6, A2=" + [int]$ws.Cells.Item(2,1).DisplayFormat.Interior.ColorIndex + " 期望-4142")

# 纯移动后再试一次，确认稳定
$app.Range("F6").Select()
Start-Sleep -Seconds 2
Write-Host ("[3] F6(纯移动): F6=" + [int]$ws.Cells.Item(6,6).DisplayFormat.Interior.ColorIndex + " 期望6, F1=" + [int]$ws.Cells.Item(1,6).DisplayFormat.Interior.ColorIndex + " 期望6, A4=" + [int]$ws.Cells.Item(4,1).DisplayFormat.Interior.ColorIndex + " 期望-4142")

# 规则数量必须是 1（零churn）
$n = 0
for ($i = 1; $i -le $ws.Cells.FormatConditions.Count; $i++) { if ([string]$ws.Cells.FormatConditions.Item($i).Formula1 -like '*JYYJ_VIEW_MODE*') { $n++ } }
Write-Host ("[4] 哨兵规则数=" + $n + " 期望1")

$wb.Close($false)
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws) | Out-Null
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null
