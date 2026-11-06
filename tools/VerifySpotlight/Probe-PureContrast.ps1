# 纯 Excel 对照探针：CF 边框 vs CF 填充 在 JYYJ_SEL 改名后的「获得/去除」行为。
# 判断指针模式边框残留是 Excel CF 局限还是可规避，并验证聚光灯填充不受影响。
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

function Bdr([int]$r, [int]$c) {
    for ($t = 0; $t -lt 30; $t++) {
        try { return [int]$ws.Cells.Item($r, $c).DisplayFormat.Borders.Item(7).ColorIndex } catch { Start-Sleep -Milliseconds 15 }
    }
    return -9999
}
function Fill([int]$r, [int]$c) {
    for ($t = 0; $t -lt 30; $t++) {
        try { return [int]$ws.Cells.Item($r, $c).DisplayFormat.Interior.ColorIndex } catch { Start-Sleep -Milliseconds 15 }
    }
    return -9999
}
function SetSel($expr) { $wb.Names.Item('JYYJ_SEL').RefersTo = $expr; $app.Calculate(); Start-Sleep -Milliseconds 400 }

$multiOr = '=OR(AND(ROW()>=1,ROW()<=3,COLUMN()>=1,COLUMN()<=1),AND(ROW()>=1,ROW()<=3,COLUMN()>=3,COLUMN()<=3))'
$boxA1C3 = '=AND(ROW()>=1,ROW()<=3,COLUMN()>=1,COLUMN()<=3)'
$boxD4 = '=AND(ROW()>=4,ROW()<=4,COLUMN()>=4,COLUMN()<=4)'

Write-Host "=== A. CF 边框（指针）：FALSE→TRUE 是否获得边框 ==="
$fc = $ws.Range('A1:F10').FormatConditions.Add(2,0,'=AND(JYYJ_VIEW_MODE,OR(AND(ISNUMBER(JYYJ_SEL),AND(ROW()=JYYJ_R,COLUMN()=JYYJ_C)),JYYJ_SEL))')
$fc.Borders.ColorIndex = 3; $fc.Borders.LineStyle = 1
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($fc) | Out-Null
SetSel $multiOr
$b0 = Bdr 2 2
Write-Host ("  初始多选区（B2 不在内）B2 border=$b0")
Assert ($b0 -ne 3) "初始 B2 无框（基线）"
SetSel $boxA1C3
$b1 = Bdr 2 2
Write-Host ("  改含 B2 盒后 B2 border=$b1")
Assert ($b1 -eq 3) "FALSE→TRUE：B2 获得边框"
SetSel $multiOr
$b2 = Bdr 2 2
Write-Host ("  改回多选区（B2 不在内）B2 border=$b2")
Assert ($b2 -ne 3) "TRUE→FALSE：B2 去除边框"
$ws.Cells.FormatConditions.Delete()

Write-Host "`n=== B. CF 填充（聚光灯）：对照组 ==="
$fc2 = $ws.Range('A1:F10').FormatConditions.Add(2,0,'=AND(JYYJ_VIEW_MODE,OR(AND(ISNUMBER(JYYJ_SEL),OR(ROW()=JYYJ_R,COLUMN()=JYYJ_C)),JYYJ_SEL))')
$fc2.Interior.ColorIndex = 6; $fc2.Interior.Pattern = 1
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($fc2) | Out-Null
SetSel $boxD4
$f0 = Fill 4 4
Write-Host ("  盒 D4：D4 fill=$f0")
Assert ($f0 -eq 6) "D4 获得填充"
SetSel $multiOr
$f1 = Fill 4 4
Write-Host ("  改多选区（D4 不在内）D4 fill=$f1")
Assert ($f1 -ne 6) "TRUE→FALSE：D4 填充去除"
SetSel $boxD4
$f2 = Fill 4 4
Write-Host ("  改回盒 D4：D4 fill=$f2")
Assert ($f2 -eq 6) "FALSE→TRUE：D4 重新获得填充"

$ws.Cells.FormatConditions.Delete()
$wb.Close($false)
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws) | Out-Null
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null
Write-Host ""
Write-Host "== PURE-CONTRAST PASSED=$script:passed FAILED=$script:failed =="
if ($script:failed -gt 0) { exit 1 } else { exit 0 }
