# Exp-CfRange.ps1 - 实验：CF 规则应用范围(全表 vs UsedRange)对聚光灯点选延迟的影响
# 纯 COM 实验，不加载产品加载项，只测 CF 重算代价
param([int]$Rows = 30000)
$ErrorActionPreference = "Stop"
$script:passed = 0; $script:failed = 0
function Assert([bool]$cond, [string]$msg) {
    if ($cond) { $script:passed++; Write-Host "[PASS] $msg" } else { $script:failed++; Write-Host "[FAIL] $msg" }
}
function Measure-Follow($ws, [int]$r, [int]$c, [int]$colorIndex, [string]$label) {
    $neutralR = 2; $neutralC = 2
    $null = $app.Cells.Item($neutralR, $neutralC).Select()
    Start-Sleep -Milliseconds 80
    # 目标格变黄（十字交点）
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $null = $app.Cells.Item($r, $c).Select()
    $found = $false
    for ($t = 0; $t -lt 600; $t++) {
        try { if ([int]$ws.Cells.Item($r, $c).DisplayFormat.Interior.ColorIndex -eq $colorIndex) { $found = $true; break } } catch { }
        Start-Sleep -Milliseconds 4
    }
    $sw.Stop()
    $ms = [int]$sw.ElapsedMilliseconds
    Assert $found ("$label 目标格 ($r,$c) 变黄（$ms ms）")
    # 十字远端：同行最右列 + 同列底部行（整行/整列重绘完成）
    $rowFarCol = 5; $colFarRow = 29990
    $swR = [System.Diagnostics.Stopwatch]::StartNew()
    $foundR = $false
    for ($t = 0; $t -lt 600; $t++) {
        try { if ([int]$ws.Cells.Item($r, $rowFarCol).DisplayFormat.Interior.ColorIndex -eq $colorIndex) { $foundR = $true; break } } catch { }
        Start-Sleep -Milliseconds 4
    }
    $swR.Stop()
    $msR = [int]$swR.ElapsedMilliseconds
    Assert $foundR ("$label 同行远端 ($r,$rowFarCol) 变黄（$msR ms）")
    $swC = [System.Diagnostics.Stopwatch]::StartNew()
    $foundC = $false
    for ($t = 0; $t -lt 600; $t++) {
        try { if ([int]$ws.Cells.Item($colFarRow, $c).DisplayFormat.Interior.ColorIndex -eq $colorIndex) { $foundC = $true; break } } catch { }
        Start-Sleep -Milliseconds 4
    }
    $swC.Stop()
    $msC = [int]$swC.ElapsedMilliseconds
    Assert $foundC ("$label 同列远端 ($colFarRow,$c) 变黄（$msC ms）")
    return @($ms, $msR, $msC)
}
function Add-Cf([object]$ws, [object]$targetRange, [string]$formula, [int]$colorIndex) {
    $fc = $targetRange.FormatConditions.Add(2, 0, $formula)   # xlExpression, xlEqual
    $fc.Interior.ColorIndex = $colorIndex
    $fc.Interior.Pattern = 1                                   # xlSolid
    return $fc
}

$app = New-Object -ComObject Excel.Application
$app.Visible = $true; $app.DisplayAlerts = $false
$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
for ($c = 1; $c -le 5; $c++) { $ws.Cells.Item(1, $c).Value2 = 1; $null = $ws.Range($ws.Cells.Item(1,$c), $ws.Cells.Item($Rows,$c)).DataSeries() }
$usedRows = [int]$ws.UsedRange.Rows.Count
Write-Host ("表行数: $Rows   UsedRange: $usedRows 行")

# 定义名（聚光灯式）
$wb.Names.Add("JYYJ_VIEW_MODE", "=1") | Out-Null
$wb.Names.Add("JYYJ_SEL", "=0") | Out-Null
$wb.Names.Add("JYYJ_R", '=CELL("row")') | Out-Null
$wb.Names.Add("JYYJ_C", '=CELL("col")') | Out-Null
$formula = "=AND(JYYJ_VIEW_MODE,OR(AND(ISNUMBER(JYYJ_SEL),OR(ROW()=JYYJ_R,COLUMN()=JYYJ_C)),JYYJ_SEL))"

# ---------- 方案A：CF 应用于整表 Cells ----------
Write-Host "`n=== 方案A：CF 应用于整表 sheet.Cells ==="
$fcA = Add-Cf $ws $ws.Cells $formula 6
$latA = @(); $latAR = @(); $latAC = @()
foreach ($tgt in @(@(500,5), @(5000,8), @(15000,3), @(25000,12), @(29000,7))) {
    $r = Measure-Follow $ws $tgt[0] $tgt[1] 6 "整表"
    $latA += $r[0]; $latAR += $r[1]; $latAC += $r[2]
}
Write-Host ("  整表: 目标格平均 $([Math]::Round(($latA|Measure-Object -Average).Average,0)) ms | 同行远端平均 $([Math]::Round(($latAR|Measure-Object -Average).Average,0)) ms | 同列远端平均 $([Math]::Round(($latAC|Measure-Object -Average).Average,0)) ms")
$fcA.Delete()

# ---------- 方案B：CF 应用于 UsedRange ----------
Write-Host "`n=== 方案B：CF 应用于 UsedRange ==="
$used = $ws.UsedRange
$fcB = Add-Cf $ws $used $formula 6
$latB = @(); $latBR = @(); $latBC = @()
foreach ($tgt in @(@(500,5), @(5000,8), @(15000,3), @(25000,12), @(29000,7))) {
    $r = Measure-Follow $ws $tgt[0] $tgt[1] 6 "UsedRange"
    $latB += $r[0]; $latBR += $r[1]; $latBC += $r[2]
}
Write-Host ("  UsedRange: 目标格平均 $([Math]::Round(($latB|Measure-Object -Average).Average,0)) ms | 同行远端平均 $([Math]::Round(($latBR|Measure-Object -Average).Average,0)) ms | 同列远端平均 $([Math]::Round(($latBC|Measure-Object -Average).Average,0)) ms")

# ---------- 对比结论 ----------
Write-Host ""
Write-Host "== 对比结论 =="
$avgA = [Math]::Round(($latA|Measure-Object -Average).Average,0)
$avgB = [Math]::Round(($latB|Measure-Object -Average).Average,0)
$avgAR = [Math]::Round(($latAR|Measure-Object -Average).Average,0)
$avgBR = [Math]::Round(($latBR|Measure-Object -Average).Average,0)
Write-Host ("整表: 目标 $avgA ms / 十字远端 $avgAR ms | UsedRange: 目标 $avgB ms / 十字远端 $avgBR ms")
if ($avgBR -lt $avgAR -and $avgB -lt $avgA) { Write-Host "=> 缩小 CF 应用范围到 UsedRange 可降低点选延迟（验证假设成立）" }
else { Write-Host "=> CF 应用范围不是主要瓶颈（十字重绘代价仍在）" }

$wb.Close($false)
$app.Quit()
try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch { }
Write-Host ""
Write-Host "== 实验 PASSED=$script:passed FAILED=$script:failed =="
if ($script:failed -gt 0) { exit 1 } else { exit 0 }
