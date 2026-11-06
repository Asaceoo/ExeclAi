# Exp-CfRange2.ps1 - 实验2：整表 CF vs 显式数据区 CF 的点选延迟（正确测量）
# 关键修正：测试格全部在 CF 范围内（列 1-5）；增加连续快速点选累积耗时（贴近真实感知）
param([int]$Rows = 30000)
$ErrorActionPreference = "Stop"
$script:passed = 0; $script:failed = 0
function Assert([bool]$cond, [string]$msg) {
    if ($cond) { $script:passed++; Write-Host "[PASS] $msg" } else { $script:failed++; Write-Host "[FAIL] $msg" }
}
# 点选 (r,c)，测 (r,c)（十字中心）变黄耗时；r/c 必须都在 CF 范围内
function Measure-Follow($ws, [int]$r, [int]$c, [int]$colorIndex, [string]$label) {
    $null = $app.Cells.Item(2, 2).Select()
    Start-Sleep -Milliseconds 80
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $null = $app.Cells.Item($r, $c).Select()
    $found = $false
    for ($t = 0; $t -lt 500; $t++) {
        try { if ([int]$ws.Cells.Item($r, $c).DisplayFormat.Interior.ColorIndex -eq $colorIndex) { $found = $true; break } } catch { }
        Start-Sleep -Milliseconds 4
    }
    $sw.Stop()
    $ms = [int]$sw.ElapsedMilliseconds
    Assert $found ("$label 点选 ($r,$c) 十字中心变黄（$ms ms）")
    return $ms
}
# 连续快速点选 n 次，测总耗时与单次均值（模拟用户快速操作）
function Measure-Rapid($ws, [int]$n, [int]$colorIndex, [string]$label) {
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    for ($i = 0; $i -lt $n; $i++) {
        $r = 10 + $i; $c = 2
        $null = $app.Cells.Item($r, $c).Select()
    }
    $sw.Stop()
    $total = [int]$sw.ElapsedMilliseconds
    Assert $true ("$label 连续点选 $n 次总耗时 $total ms（均值 $([Math]::Round($total/$n,0)) ms/次）")
    return $total
}
function Add-Cf([object]$ws, [object]$targetRange, [string]$formula, [int]$colorIndex) {
    $fc = $targetRange.FormatConditions.Add(2, 0, $formula)
    $fc.Interior.ColorIndex = $colorIndex
    $fc.Interior.Pattern = 1
    return $fc
}

$app = New-Object -ComObject Excel.Application
$app.Visible = $true; $app.DisplayAlerts = $false
$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
for ($c = 1; $c -le 5; $c++) { $ws.Cells.Item(1, $c).Value2 = 1; $null = $ws.Range($ws.Cells.Item(1,$c), $ws.Cells.Item($Rows,$c)).DataSeries() }
Write-Host ("表行数: $Rows   UsedRange: " + [int]$ws.UsedRange.Rows.Count + " 行")

$wb.Names.Add("JYYJ_VIEW_MODE", "=1") | Out-Null
$wb.Names.Add("JYYJ_SEL", "=0") | Out-Null
$wb.Names.Add("JYYJ_R", '=CELL("row")') | Out-Null
$wb.Names.Add("JYYJ_C", '=CELL("col")') | Out-Null
$formula = "=AND(JYYJ_VIEW_MODE,OR(AND(ISNUMBER(JYYJ_SEL),OR(ROW()=JYYJ_R,COLUMN()=JYYJ_C)),JYYJ_SEL))"

# 目标格（全在列 1-5 范围内）
$targets = @(@(500,3), @(5000,4), @(15000,2), @(25000,5), @(29000,1))

# ---------- 方案A：整表 CF ----------
Write-Host "`n=== 方案A：整表 sheet.Cells CF ==="
$fcA = Add-Cf $ws $ws.Cells $formula 6
$latA = @()
foreach ($t in $targets) { $latA += Measure-Follow $ws $t[0] $t[1] 6 "整表" }
$avgA = [Math]::Round(($latA|Measure-Object -Average).Average,0)
Write-Host ("  整表: 点选变黄平均 $avgA ms（逐次: " + ($latA -join ",") + "）")
$rapidA = Measure-Rapid $ws 30 6 "整表"
$fcA.Delete()

# ---------- 方案B：显式数据区 A1:E30000 CF ----------
Write-Host "`n=== 方案B：显式数据区 A1:E30000 CF ==="
$dataRange = $ws.Range("A1:E$Rows")
$fcB = Add-Cf $ws $dataRange $formula 6
$latB = @()
foreach ($t in $targets) { $latB += Measure-Follow $ws $t[0] $t[1] 6 "数据区" }
$avgB = [Math]::Round(($latB|Measure-Object -Average).Average,0)
Write-Host ("  数据区: 点选变黄平均 $avgB ms（逐次: " + ($latB -join ",") + "）")
$rapidB = Measure-Rapid $ws 30 6 "数据区"

Write-Host ""
Write-Host "== 对比 =="
Write-Host ("点选变黄: 整表 $avgA ms | 数据区 $avgB ms | 快速30连点: 整表 $rapidA ms | 数据区 $rapidB ms")
if ($avgB -lt $avgA -or $rapidB -lt $rapidA) { Write-Host "=> 数据区 CF 更快（缩小范围有效）" }
elseif ([Math]::Abs($avgB - $avgA) -le 30 -and [Math]::Abs($rapidB - $rapidA) -le 30) { Write-Host "=> 范围无关，CF 全表并不慢；慢感另有来源" }
else { Write-Host "=> 数据区 CF 更慢" }

$wb.Close($false)
$app.Quit()
try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch { }
Write-Host ""
Write-Host "== 实验 PASSED=$script:passed FAILED=$script:failed =="
if ($script:failed -gt 0) { exit 1 } else { exit 0 }
