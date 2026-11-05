# Debug: 混合方案在大公式表上失效的根因定位。
$ErrorActionPreference = "Stop"
$app = $null
for ($i = 0; $i -lt 40; $i++) {
    try { $app = New-Object -ComObject Excel.Application; break } catch { Start-Sleep -Seconds 1 }
}
if ($null -eq $app) { Write-Host "NO EXCEL INSTANCE"; exit 2 }
$app.Visible = $false
$app.DisplayAlerts = $false

$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$ws.Range("A1:A20000").Formula = '=ROW()*2'
$app.Calculate()

$wb.Names.Add('JYYJ_VIEW_MODE', '=1') | Out-Null
$wb.Names.Add('JYYJ_R', '=CELL("row")') | Out-Null
$wb.Names.Add('JYYJ_C', '=CELL("col")') | Out-Null
$wb.Names.Add('JYYJ_SEL', '=FALSE') | Out-Null

# dump names
Write-Host "--- Names ---"
for ($i = 1; $i -le $wb.Names.Count; $i++) {
    $n = $wb.Names.Item($i)
    $v = "ERR"
    try { $v = [string]$n.Value } catch { }
    Write-Host ("  " + $n.Name + " = " + $n.RefersTo + " Value=" + $v)
}
# 在 B2 建 CF（不选中，先建规则）
$fc = $ws.Cells.FormatConditions.Add(2, 0, '=AND(JYYJ_VIEW_MODE,OR(AND(ROW()=JYYJ_R,COLUMN()=JYYJ_C),JYYJ_SEL))')
$fc.Interior.ColorIndex = 6; $fc.Interior.Pattern = 1
Write-Host ("--- CF Formula1 = " + $fc.Formula1)

# 选中 B2
$app.Range("B2").Select()
Start-Sleep -Milliseconds 300
Write-Host ("--- 选中 B2 后 ---")
Write-Host ("  ActiveCell = " + $app.ActiveCell.Address)
Write-Host ("  Evaluate CELL(row) = " + $app.Evaluate('CELL("row")'))
Write-Host ("  Evaluate CELL(col) = " + $app.Evaluate('CELL("col")'))
$c1 = -9999
try { $c1 = [int]$ws.Cells.Item(2,1).DisplayFormat.Interior.ColorIndex } catch { Write-Host ("  B2 读色 ERR: " + $_.Exception.Message) }
Write-Host ("  B2 行首 A2 色 = " + $c1)

# 换 D4
$app.Range("D4").Select()
Start-Sleep -Milliseconds 300
Write-Host ("--- 选中 D4 后 ---")
Write-Host ("  Evaluate CELL(row) = " + $app.Evaluate('CELL("row")'))
Write-Host ("  Evaluate CELL(col) = " + $app.Evaluate('CELL("col")'))
$c2 = -9999
try { $c2 = [int]$ws.Cells.Item(4,1).DisplayFormat.Interior.ColorIndex } catch { Write-Host ("  D4 读色 ERR: " + $_.Exception.Message) }
Write-Host ("  D4 行首 A4 色 = " + $c2)

$wb.Close($false)
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws) | Out-Null
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null
$app.Quit()
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) | Out-Null
Write-Host "DONE"
