# 纯 Excel 决定性探针：全表 CF 边框规则 + 补救手段，定位指针边框「去除」粘滞根因。
# 目标：确认是 Excel CF 边框固有怪癖，并找出可让边框随 JYYJ_SEL 变化正确更新的方法。
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
function SetSel($expr) { $wb.Names.Item('JYYJ_SEL').RefersTo = $expr; $app.Calculate(); Start-Sleep -Milliseconds 400 }

$multiOr = '=OR(AND(ROW()>=1,ROW()<=3,COLUMN()>=1,COLUMN()<=1),AND(ROW()>=1,ROW()<=3,COLUMN()>=3,COLUMN()<=3))'
$boxA1C3 = '=AND(ROW()>=1,ROW()<=3,COLUMN()>=1,COLUMN()<=3)'

# 与加载项一致：规则建在全表 Cells 上
$fc = $ws.Cells.FormatConditions.Add(2,0,'=AND(JYYJ_VIEW_MODE,OR(AND(ISNUMBER(JYYJ_SEL),AND(ROW()=JYYJ_R,COLUMN()=JYYJ_C)),JYYJ_SEL))')
$fc.Borders.ColorIndex = 3; $fc.Borders.LineStyle = 1
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($fc) | Out-Null

Write-Host "=== 基线：JYYJ_SEL=0，活动格 A1（B2 从未 TRUE）==="
$app.Calculate(); Start-Sleep -Milliseconds 400
$bA1 = Bdr 1 1
$bB2 = Bdr 2 2
$bC3 = Bdr 3 3
Write-Host ("  A1=$bA1 B2=$bB2 C3=$bC3")
Assert ($bA1 -eq 3) "A1 有框（十字过 A1）"
Assert ($bB2 -ne 3) "B2 无框（十字不过 B2）"
Assert ($bC3 -ne 3) "C3 无框（十字不过 C3）"

Write-Host "`n=== 补救 A：Select B2 再改多选区 ==="
$app.Range("B2").Select(); $app.Calculate(); Start-Sleep -Milliseconds 400
$bB2a = Bdr 2 2
Write-Host ("  选中 B2 后 B2=$bB2a")
SetSel $multiOr
$bB2b = Bdr 2 2
Write-Host ("  改多选区（B2 排除）B2=$bB2b")
Assert ($bB2b -ne 3) "补救A：B2 边框去除"

Write-Host "`n=== 补救 B：ScreenUpdating 重刷 ==="
$app.ScreenUpdating = $false; $app.ScreenUpdating = $true; $app.Calculate(); Start-Sleep -Milliseconds 400
$bB2c = Bdr 2 2
Write-Host ("  屏刷后 B2=$bB2c")
Assert ($bB2c -ne 3) "补救B：B2 边框去除"

Write-Host "`n=== 补救 C：删除重建规则 ==="
$ws.Cells.FormatConditions.Delete()
$fc2 = $ws.Cells.FormatConditions.Add(2,0,'=AND(JYYJ_VIEW_MODE,OR(AND(ISNUMBER(JYYJ_SEL),AND(ROW()=JYYJ_R,COLUMN()=JYYJ_C)),JYYJ_SEL))')
$fc2.Borders.ColorIndex = 3; $fc2.Borders.LineStyle = 1
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($fc2) | Out-Null
$app.Calculate(); Start-Sleep -Milliseconds 400
$bB2d = Bdr 2 2
$bA2d = Bdr 2 1
Write-Host ("  重建后 B2=$bB2d A2=$bA2d")
Assert ($bB2d -ne 3) "补救C：重建后 B2 无框"
Assert ($bA2d -eq 3) "补救C：A2 有框（多选区仍有效）"

$ws.Cells.FormatConditions.Delete()
$wb.Close($false)
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws) | Out-Null
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null
Write-Host ""
Write-Host "== DECISIVE PASSED=$script:passed FAILED=$script:failed =="
if ($script:failed -gt 0) { exit 1 } else { exit 0 }
