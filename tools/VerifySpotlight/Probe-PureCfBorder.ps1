# 纯 Excel 隔离探针：CF 边框「去除」是否随 JYYJ_SEL 改名 + Calculate 重绘。
# 不经过加载项，直接手动建规则，判断「多选 B2 误亮」是 Excel CF 重绘时序还是加载项逻辑。
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
try { $wb.Names.Item('JYYJ_VIEW_MODE').Delete() } catch { }
try { $wb.Names.Item('JYYJ_SEL').Delete() } catch { }
try { $wb.Names.Item('JYYJ_R').Delete() } catch { }
try { $wb.Names.Item('JYYJ_C').Delete() } catch { }
$wb.Names.Add('JYYJ_VIEW_MODE','=1') | Out-Null
$wb.Names.Add('JYYJ_R','=CELL("row")') | Out-Null
$wb.Names.Add('JYYJ_C','=CELL("col")') | Out-Null
$wb.Names.Add('JYYJ_SEL','=0') | Out-Null

# 指针规则公式（与加载项一致）
$fc = $ws.Range('A1:F10').FormatConditions.Add(2,0,'=AND(JYYJ_VIEW_MODE,OR(AND(ISNUMBER(JYYJ_SEL),AND(ROW()=JYYJ_R,COLUMN()=JYYJ_C)),JYYJ_SEL))')
$fc.Borders.ColorIndex = 3
$fc.Borders.LineStyle = 1
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($fc) | Out-Null

function Bdr([int]$r, [int]$c) {
    for ($t = 0; $t -lt 30; $t++) {
        try { return [int]$ws.Cells.Item($r, $c).DisplayFormat.Borders.Item(7).ColorIndex } catch { Start-Sleep -Milliseconds 15 }
    }
    return -9999
}

Write-Host "`n>> JYYJ_SEL=0（单格 CELL 十字），活动格 B2"
$app.Range("B2").Select(); $app.Calculate(); Start-Sleep -Milliseconds 400
Assert ((Bdr 2 2) -eq 3) "B2 有框（单格包围盒）"
Assert ((Bdr 3 3) -ne 3) "C3 无框"

Write-Host "`n>> JYYJ_SEL=包围盒 A1:C3（含 B2）"
$wb.Names.Item('JYYJ_SEL').RefersTo = '=AND(ROW()>=1,ROW()<=3,COLUMN()>=1,COLUMN()<=3)'
$app.Calculate(); Start-Sleep -Milliseconds 400
Assert ((Bdr 2 2) -eq 3) "B2 有框（盒内含 B2）"

Write-Host "`n>> JYYJ_SEL=多选区 OR（A1:A3,C1:C3，不含 B2）"
$wb.Names.Item('JYYJ_SEL').RefersTo = '=OR(AND(ROW()>=1,ROW()<=3,COLUMN()>=1,COLUMN()<=1),AND(ROW()>=1,ROW()<=3,COLUMN()>=3,COLUMN()<=3))'
$app.Calculate(); Start-Sleep -Milliseconds 400
$b1 = Bdr 2 2
Write-Host ("  改名+Calculate 后 B2 border=$b1 (3=红仍在)")
Assert ($b1 -ne 3) "B2 无框（多选区去除，Calculate 后）"
Assert ((Bdr 2 1) -eq 3) "A2 仍有框"

Write-Host "`n>> 再读（等 1s 后）确认非瞬时"
Start-Sleep -Seconds 1
$b2 = Bdr 2 2
Write-Host ("  1s 后 B2 border=$b2")
Assert ($b2 -ne 3) "B2 无框（1s 后仍无）"

$wb.Close($false)
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws) | Out-Null
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null
Write-Host ""
Write-Host "== PURE-CF PASSED=$script:passed FAILED=$script:failed =="
if ($script:failed -gt 0) { exit 1 } else { exit 0 }
