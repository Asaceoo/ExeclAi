# Probe: 定位「聚光灯慢」根因 —— 测触发 CF 重评的最小代价 & 多选 predicate 有效性。
# A) 改 Names.RefersTo 后不调用 Calculate，CF 是否自动更新？
# B) 改 RefersTo + ws.Calculate() 的耗时（当前实现路径）
# C) 用 CELL("row")/CELL("col") 易失函数，活动格变化时 CF 是否免 Calculate 自动跟随
# D) 多选 OR() predicate 是否全部区域生效
# E) 不同公式规模下 ws.Calculate() 耗时（用户大表场景）
$ErrorActionPreference = "Stop"
$app = $null
for ($i = 0; $i -lt 40; $i++) {
    try { $app = New-Object -ComObject Excel.Application; break } catch { Start-Sleep -Seconds 1 }
}
if ($null -eq $app) { Write-Host "NO EXCEL INSTANCE"; exit 2 }
$app.Visible = $false
$app.DisplayAlerts = $false
$app.ScreenUpdating = $false

$sw = [System.Diagnostics.Stopwatch]::StartNew()
$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)

function CFC([int]$r, [int]$c) {
    for ($t = 0; $t -lt 20; $t++) {
        try { return [int]$ws.Cells.Item($r, $c).DisplayFormat.Interior.ColorIndex } catch { Start-Sleep -Milliseconds 25 }
    }
    return -9999
}
function ResetFC { for ($i = $ws.Cells.FormatConditions.Count; $i -ge 1; $i--) { try { $ws.Cells.FormatConditions.Item($i).Delete() } catch {} } }
function ResetNames {
    foreach ($nm in @('JYYJ_VIEW_MODE','JYYJ_SEL','JYYJ_R','JYYJ_C')) {
        for ($i = $wb.Names.Count; $i -ge 1; $i--) { if ([string]$wb.Names.Item($i).Name -eq $nm) { try { $wb.Names.Item($i).Delete() } catch {} } }
    }
}

# ---------- 场景一：当前 JYYJ_SEL 方案，改 RefersTo 后【不】Calculate ----------
Write-Host "`n=== A) RefersTo 改后不 Calculate，CF 是否自动更新 ==="
ResetFC; ResetNames
$wb.Names.Add('JYYJ_VIEW_MODE', '=1') | Out-Null
$wb.Names.Add('JYYJ_SEL', '=FALSE') | Out-Null
$fc = $ws.Cells.FormatConditions.Add(2, 0, '=AND(JYYJ_VIEW_MODE,JYYJ_SEL)')
$fc.Interior.ColorIndex = 6
$fc.Interior.Pattern = 1
# 不 Calculate，只改 RefersTo
$wb.Names.Item('JYYJ_SEL').RefersTo = '=OR(AND(ROW()>=2,ROW()<=2),AND(COLUMN()>=2,COLUMN()<=2))'
Start-Sleep -Milliseconds 300
Write-Host ("  改后 300ms 未 Calculate：B2 行内 A2 色=" + (CFC 2 1) + "（6=已生效, -4142=未生效）")

# ---------- 场景二：改 RefersTo + Calculate，测耗时 ----------
Write-Host "`n=== B) RefersTo + Calculate 耗时（当前实现） ==="
$t0 = [System.Diagnostics.Stopwatch]::StartNew()
$wb.Names.Item('JYYJ_SEL').RefersTo = '=OR(AND(ROW()>=4,ROW()<=4),AND(COLUMN()>=4,COLUMN()<=4))'
$ws.Calculate()
$t0.Stop()
Write-Host ("  小表改RefersTo+Calculate 总耗时=" + $t0.ElapsedMilliseconds + "ms, 结果 D4 行 A4 色=" + (CFC 4 1))

# ---------- 场景三：CELL() 易失函数方案（免 Calculate 免轮询写名） ----------
Write-Host "`n=== C) CELL(row/col) 易失跟随：活动格变化自动更新 CF ==="
ResetFC; ResetNames
$wb.Names.Add('JYYJ_VIEW_MODE', '=1') | Out-Null
$wb.Names.Add('JYYJ_R', '=CELL("row")') | Out-Null
$wb.Names.Add('JYYJ_C', '=CELL("col")') | Out-Null
$fc2 = $ws.Cells.FormatConditions.Add(2, 0, '=AND(JYYJ_VIEW_MODE,OR(ROW()=JYYJ_R,COLUMN()=JYYJ_C))')
$fc2.Interior.ColorIndex = 6
$fc2.Interior.Pattern = 1
$app.Range("B2").Select(); Start-Sleep -Milliseconds 300
Write-Host ("  B2：A2 色=" + (CFC 2 1) + " B1 色=" + (CFC 1 2))
$app.Range("D4").Select(); Start-Sleep -Milliseconds 300
Write-Host ("  D4（未 Calculate）：A4 色=" + (CFC 4 1) + " 旧区 A2 色=" + (CFC 2 1) + "（-4142=已恢复）")

# ---------- 场景四：多选 OR() predicate ----------
Write-Host "`n=== D) 多选 predicate 全区域生效 ==="
ResetFC; ResetNames
$wb.Names.Add('JYYJ_VIEW_MODE', '=1') | Out-Null
$wb.Names.Add('JYYJ_SEL', '=FALSE') | Out-Null
$fc3 = $ws.Cells.FormatConditions.Add(2, 0, '=AND(JYYJ_VIEW_MODE,JYYJ_SEL)')
$fc3.Interior.ColorIndex = 6
$fc3.Interior.Pattern = 1
$pred = '=OR(OR(AND(ROW()>=1,ROW()<=5),AND(COLUMN()>=1,COLUMN()<=1)),OR(AND(ROW()>=1,ROW()<=5),AND(COLUMN()>=3,COLUMN()<=3)))'
$wb.Names.Item('JYYJ_SEL').RefersTo = $pred
$ws.Calculate(); Start-Sleep -Milliseconds 200
Write-Host ("  多选 A1:A5+C1:C5：A2 色=" + (CFC 2 1) + " C2 色=" + (CFC 2 3) + " B3 色=" + (CFC 3 2) + " B6 色=" + (CFC 6 2) + "（6=亮,-4142=不亮）")

# ---------- 场景五：公式规模 vs Calculate 耗时 ----------
Write-Host "`n=== E) ws.Calculate() 在不同公式规模下的耗时 ==="
ResetFC; ResetNames
$ws2 = $wb.Worksheets.Item(2)
$ws2.Activate()
$cols = @(1,2,3)
foreach ($n in @(1000, 10000, 50000)) {
    for ($r = 1; $r -le $n; $r++) {
        $ws2.Cells.Item($r,1).Formula = '=ROW()*2'
        $ws2.Cells.Item($r,2).Formula = '=SUM($A$1:A' + $r + ')'
    }
    $ws2.Calculate()
    $t1 = [System.Diagnostics.Stopwatch]::StartNew()
    for ($k = 0; $k -lt 5; $k++) { $ws2.Calculate() }
    $t1.Stop()
    Write-Host ("  公式行数 $n：单次 ws.Calculate 平均=" + [math]::Round($t1.ElapsedMilliseconds / 5, 1) + "ms")
}

$wb.Close($false)
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws2) | Out-Null
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws) | Out-Null
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null
$app.Quit()
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) | Out-Null
Write-Host "`nDONE"
