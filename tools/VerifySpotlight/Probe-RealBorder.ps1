# 关键验证：CF 边框是否把「真实边框」写进单元格（Borders 本体 vs DisplayFormat）。
# 若真边框残留 → 指针模式存在用户可见的严重缺陷（取消高亮后边框残留）。
$ErrorActionPreference = "Stop"
$script:passed = 0; $script:failed = 0
function Assert([bool]$cond, [string]$msg) {
    if ($cond) { $script:passed++; Write-Host "[PASS] $msg" }
    else { $script:failed++; Write-Host "[FAIL] $msg" }
}
$app = $null
for ($i = 0; $i -lt 15; $i++) { try { $app = New-Object -ComObject Excel.Application; break } catch { Start-Sleep -Seconds 1 } }
$app.Visible = $true; $app.DisplayAlerts = $false
$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$ws.Cells.FormatConditions.Delete()
foreach ($n in @('JYYJ_VIEW_MODE','JYYJ_SEL','JYYJ_R','JYYJ_C')) { try { $wb.Names.Item($n).Delete() } catch { } }
$wb.Names.Add('JYYJ_VIEW_MODE','=1') | Out-Null
$wb.Names.Add('JYYJ_R','=CELL("row")') | Out-Null
$wb.Names.Add('JYYJ_C','=CELL("col")') | Out-Null
$wb.Names.Add('JYYJ_SEL','=0') | Out-Null

function BdrDisplay([int]$r, [int]$c) {
    for ($t = 0; $t -lt 30; $t++) {
        try { return [int]$ws.Cells.Item($r, $c).DisplayFormat.Borders.Item(7).ColorIndex } catch { Start-Sleep -Milliseconds 15 }
    }
    return -9999
}
function BdrReal([int]$r, [int]$c) {
    for ($t = 0; $t -lt 30; $t++) {
        try { return [int]$ws.Cells.Item($r, $c).Borders.Item(7).ColorIndex } catch { Start-Sleep -Milliseconds 15 }
    }
    return -9999
}
$multiOr = '=OR(AND(ROW()>=1,ROW()<=3,COLUMN()>=1,COLUMN()<=1),AND(ROW()>=1,ROW()<=3,COLUMN()>=3,COLUMN()<=3))'

Write-Host "=== 建 CF 边框规则，选中 B2（B2 条件 TRUE）==="
$fc = $ws.Cells.FormatConditions.Add(2,0,'=AND(JYYJ_VIEW_MODE,OR(AND(ISNUMBER(JYYJ_SEL),AND(ROW()=JYYJ_R,COLUMN()=JYYJ_C)),JYYJ_SEL))')
$fc.Borders.ColorIndex = 3; $fc.Borders.LineStyle = 1
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($fc) | Out-Null
$app.Range("B2").Select(); $app.Calculate(); Start-Sleep -Milliseconds 400
$d1 = BdrDisplay 2 2; $r1 = BdrReal 2 2
Write-Host ("  B2 Display=$d1 Real=$r1")
Assert ($d1 -eq 3) "B2 DisplayFormat 红框"
Assert ($r1 -eq 3) "B2 真实边框被 CF 写入红框（关键）"

Write-Host "`n=== 删除 CF 规则（条件消失），B2 边框是否回收 ==="
$ws.Cells.FormatConditions.Delete()
$app.Calculate(); Start-Sleep -Milliseconds 400
$d2 = BdrDisplay 2 2; $r2 = BdrReal 2 2
Write-Host ("  B2 Display=$d2 Real=$r2")
Assert ($r2 -ne 3) "删除规则后 B2 真实边框已回收"

Write-Host "`n=== 对照组：CF 填充 删除后是否回收 ==="
$fc2 = $ws.Cells.FormatConditions.Add(2,0,'=AND(JYYJ_VIEW_MODE,OR(AND(ISNUMBER(JYYJ_SEL),OR(ROW()=JYYJ_R,COLUMN()=JYYJ_C)),JYYJ_SEL))')
$fc2.Interior.ColorIndex = 6; $fc2.Interior.Pattern = 1
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($fc2) | Out-Null
$app.Range("D4").Select(); $app.Calculate(); Start-Sleep -Milliseconds 400
$fDisp = [int]$ws.Cells.Item(4,4).DisplayFormat.Interior.ColorIndex
$fReal = [int]$ws.Cells.Item(4,4).Interior.ColorIndex
Write-Host ("  D4 Display=$fDisp Real=$fReal")
$ws.Cells.FormatConditions.Delete(); $app.Calculate(); Start-Sleep -Milliseconds 400
$fDisp2 = [int]$ws.Cells.Item(4,4).DisplayFormat.Interior.ColorIndex
$fReal2 = [int]$ws.Cells.Item(4,4).Interior.ColorIndex
Write-Host ("  删规则后 D4 Display=$fDisp2 Real=$fReal2")
Assert ($fReal2 -ne 6) "删规则后 D4 真实填充已回收"

$wb.Close($false)
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws) | Out-Null
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null
Write-Host ""
Write-Host "== REAL-BORDER PASSED=$script:passed FAILED=$script:failed =="
if ($script:failed -gt 0) { exit 1 } else { exit 0 }
