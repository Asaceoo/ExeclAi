# Probe 2: 验证「改 RefersTo 免 Calculate」在大表/公式表上的可靠性 + CELL() 易失方案重复性。
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
$ws.Name = "Data"

function CFC([int]$r, [int]$c) {
    for ($t = 0; $t -lt 40; $t++) {
        try { return [int]$ws.Cells.Item($r, $c).DisplayFormat.Interior.ColorIndex } catch { Start-Sleep -Milliseconds 10 }
    }
    return -9999
}
function ResetAll {
    for ($i = $ws.Cells.FormatConditions.Count; $i -ge 1; $i--) { try { $ws.Cells.FormatConditions.Item($i).Delete() } catch {} }
    foreach ($nm in @('JYYJ_VIEW_MODE','JYYJ_SEL','JYYJ_R','JYYJ_C')) {
        for ($i = $wb.Names.Count; $i -ge 1; $i--) { if ([string]$wb.Names.Item($i).Name -eq $nm) { try { $wb.Names.Item($i).Delete() } catch {} } }
    }
}
function Setup-CurrentDesign {
    ResetAll
    $wb.Names.Add('JYYJ_VIEW_MODE', '=1') | Out-Null
    $wb.Names.Add('JYYJ_SEL', '=FALSE') | Out-Null
    $fc = $ws.Cells.FormatConditions.Add(2, 0, '=AND(JYYJ_VIEW_MODE,JYYJ_SEL)')
    $fc.Interior.ColorIndex = 6; $fc.Interior.Pattern = 1
}
function Setup-CellVolatile {
    ResetAll
    $wb.Names.Add('JYYJ_VIEW_MODE', '=1') | Out-Null
    $wb.Names.Add('JYYJ_R', '=CELL("row")') | Out-Null
    $wb.Names.Add('JYYJ_C', '=CELL("col")') | Out-Null
    $wb.Names.Add('JYYJ_SEL', '=FALSE') | Out-Null
    $fc = $ws.Cells.FormatConditions.Add(2, 0, '=AND(JYYJ_VIEW_MODE,OR(AND(ROW()=JYYJ_R,COLUMN()=JYYJ_C),JYYJ_SEL))')
    $fc.Interior.ColorIndex = 6; $fc.Interior.Pattern = 1
}
function ResetData([int]$rows) {
    $ws2 = $wb.Worksheets.Item(1)
    # 快速填充：R1C1 数组一次性写入，避免逐格 COM 往返
    $f1 = '=ROW()*2'
    $f2 = '=SUM(R1C1:RC[-1])'
    $f3 = '=ROW()+COLUMN()'
    $ws2.Range("A1:A$rows").Formula = $f1
    $ws2.Range("B1:B$rows").Formula = $f2
    $ws2.Range("C1:C$rows").Formula = $f3
    $app.Calculate()
}

# ---------- 1. 大表 + RefersTo 免 Calculate 的可靠性 & 耗时 ----------
Write-Host "=== 1) 公式大表(20000行) + 改RefersTo免Calculate ==="
Setup-CurrentDesign
ResetData 20000
$pred = '=OR(AND(ROW()>=2,ROW()<=2),AND(COLUMN()>=2,COLUMN()<=2))'
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$wb.Names.Item('JYYJ_SEL').RefersTo = $pred
$sw.Stop()
$c1 = CFC 2 1
Write-Host ("  改RefersTo耗时=" + $sw.ElapsedMilliseconds + "ms 结果B2行A2色=" + $c1 + "（6=免Calculate生效）")

# ---------- 2. 大表 + RefersTo + Calculate 的耗时对比 ----------
Write-Host "=== 2) 同表 + 改RefersTo + ws.Calculate 耗时对比 ==="
$pred2 = '=OR(AND(ROW()>=4,ROW()<=4),AND(COLUMN()>=4,COLUMN()<=4))'
$sw2 = [System.Diagnostics.Stopwatch]::StartNew()
$wb.Names.Item('JYYJ_SEL').RefersTo = $pred2
$ws.Calculate()
$sw2.Stop()
Write-Host ("  改+Calculate 总耗时=" + $sw2.ElapsedMilliseconds + "ms")
# 单独测 Calculate 成本
$app.Calculate()
$sw3 = [System.Diagnostics.Stopwatch]::StartNew()
for ($k = 0; $k -lt 3; $k++) { $ws.Calculate() }
$sw3.Stop()
Write-Host ("  单次 ws.Calculate 平均=" + [math]::Round($sw3.ElapsedMilliseconds / 3, 1) + "ms（20000行公式表）")

# ---------- 3. CELL() 易失重复性：连点 5 格免 Calculate 全跟随 ----------
Write-Host "=== 3) CELL() 易失：连点5格免Calculate全跟随 ==="
Setup-CellVolatile
ResetData 20000
$sw4 = [System.Diagnostics.Stopwatch]::StartNew()
foreach ($cell in @("B2","D4","F6","H8","J10")) {
    $app.Range($cell).Select()
    Start-Sleep -Milliseconds 60
}
$sw4.Stop()
$r2 = [int]$app.Range("J10").Row
Write-Host ("  连点5格总耗时(含每次60ms等待)=" + $sw4.ElapsedMilliseconds + "ms")
Write-Host ("  终态 J10 行内 A10 色=" + (CFC 10 1) + " 列内 J1 色=" + (CFC 1 10) + " 旧区 A2 色=" + (CFC 2 1))

# ---------- 4. CELL() 多选混合：JYYJ_SEL 设多选 predicate，CELL 提供活动格十字 ----------
Write-Host "=== 4) CELL()+JYYJ_SEL 混合：单格用CELL(零写)，多选用predicate ==="
# 当前选区已到 J10，JYYJ_SEL=FALSE → 应只有 J10 十字
Write-Host ("  单格态 J10 十字：A10 色=" + (CFC 10 1) + " J1 色=" + (CFC 1 10) + " A1 色=" + (CFC 1 1))
# 写入多选 predicate（模拟多选 A1:A5+C1:C5）
$wb.Names.Item('JYYJ_SEL').RefersTo = '=OR(OR(AND(ROW()>=1,ROW()<=5),AND(COLUMN()>=1,COLUMN()<=1)),OR(AND(ROW()>=1,ROW()<=5),AND(COLUMN()>=3,COLUMN()<=3)))'
Start-Sleep -Milliseconds 150
Write-Host ("  多选态：A2 色=" + (CFC 2 1) + " C2 色=" + (CFC 2 3) + " B3 色=" + (CFC 3 2) + " B6 色=" + (CFC 6 2))
# 恢复 FALSE 后单格态仍由 CELL 兜底
$wb.Names.Item('JYYJ_SEL').RefersTo = '=FALSE'
Start-Sleep -Milliseconds 150
Write-Host ("  恢复单格态：J10 十字 A10 色=" + (CFC 10 1) + " A1 色=" + (CFC 1 1) + "（A1应为-4142）")

$wb.Close($false)
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws) | Out-Null
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null
$app.Quit()
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) | Out-Null
Write-Host "DONE"
