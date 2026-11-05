# 决定性验证：定义名内的 ROW()/COLUMN() 是「逐格求值」还是「活动格求值」。
# 若 JYYJ_SEL 在 B2 求值返回 TRUE（而应 FALSE）→ 名称内相对引用按活动格算 → 多选区边框误亮根因。
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
foreach ($n in @('JYYJ_SEL')) { try { $wb.Names.Item($n).Delete() } catch { } }
$wb.Names.Add('JYYJ_SEL','=OR(AND(ROW()>=1,ROW()<=3,COLUMN()>=1,COLUMN()<=1),AND(ROW()>=1,ROW()<=3,COLUMN()>=3,COLUMN()<=3))') | Out-Null

Write-Host "=== 活动格 A1，名称内 ROW()/COLUMN() 求值上下文 ==="
$app.Range("A1").Select()
$app.Calculate()
$vA1 = [string]$ws.Range("A1").Value2    # 名称引用：=JYYJ_SEL 逐格写
$ws.Range("B2").Formula = '=JYYJ_SEL'; $ws.Range("C3").Formula = '=JYYJ_SEL'; $ws.Range("D4").Formula = '=JYYJ_SEL'
$app.Calculate(); Start-Sleep -Milliseconds 400
$vB2 = [string]$ws.Range("B2").Value2
$vC3 = [string]$ws.Range("C3").Value2
$vD4 = [string]$ws.Range("D4").Value2
Write-Host ("  A1=$vA1 B2=$vB2 C3=$vC3 D4=$vD4")
Write-Host ("  期望：A1=TRUE B2=FALSE C3=FALSE D4=FALSE")
Assert ($vA1 -eq 'True') "A1=TRUE（在 A 区）"
Assert ($vB2 -eq 'False') "B2=FALSE（不在选区）逐格求值"
Assert ($vD4 -eq 'False') "D4=FALSE（不在选区）逐格求值"

Write-Host "`n=== 活动格改为 B2（选区内不同格）再看 B2 求值 ==="
$app.Range("B2").Select()
$app.Calculate(); Start-Sleep -Milliseconds 400
$vB2b = [string]$ws.Range("B2").Value2
$vA1b = [string]$ws.Range("A1").Value2
Write-Host ("  活动格B2：A1=$vA1b B2=$vB2b")
Assert ($vB2b -eq 'False') "活动格 B2 时 B2 仍=FALSE（逐格）"

Write-Host "`n=== 验证名称在 CF 中的实际上下文：直接读多选区 CF 边框 ==="
$ws.Cells.FormatConditions.Delete()
foreach ($n in @('JYYJ_VIEW_MODE','JYYJ_R','JYYJ_C')) { try { $wb.Names.Item($n).Delete() } catch { } }
$wb.Names.Add('JYYJ_VIEW_MODE','=1') | Out-Null
$wb.Names.Add('JYYJ_R','=CELL("row")') | Out-Null
$wb.Names.Add('JYYJ_C','=CELL("col")') | Out-Null
$fc = $ws.Cells.FormatConditions.Add(2,0,'=AND(JYYJ_VIEW_MODE,OR(AND(ISNUMBER(JYYJ_SEL),AND(ROW()=JYYJ_R,COLUMN()=JYYJ_C)),JYYJ_SEL))')
$fc.Borders.ColorIndex = 3; $fc.Borders.LineStyle = 1
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($fc) | Out-Null
$app.Range("A1").Select(); $app.Calculate(); Start-Sleep -Milliseconds 400
$bA1 = [int]$ws.Cells.Item(1,1).DisplayFormat.Borders.Item(7).ColorIndex
$bB2 = [int]$ws.Cells.Item(2,2).DisplayFormat.Borders.Item(7).ColorIndex
$bC3 = [int]$ws.Cells.Item(3,3).DisplayFormat.Borders.Item(7).ColorIndex
$bD4 = [int]$ws.Cells.Item(4,4).DisplayFormat.Borders.Item(7).ColorIndex
Write-Host ("  CF边框 A1=$bA1 B2=$bB2 C3=$bC3 D4=$bD4  (3=红, -4142=无)")
Assert ($bA1 -eq 3) "A1 红框（A 区）"
Assert ($bB2 -ne 3) "B2 无框（不在选区）"
Assert ($bC3 -ne 3) "C3 无框"
Assert ($bD4 -ne 3) "D4 无框"

$ws.Cells.FormatConditions.Delete()
$wb.Close($false)
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws) | Out-Null
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null
Write-Host ""
Write-Host "== EVALCTX PASSED=$script:passed FAILED=$script:failed =="
if ($script:failed -gt 0) { exit 1 } else { exit 0 }
