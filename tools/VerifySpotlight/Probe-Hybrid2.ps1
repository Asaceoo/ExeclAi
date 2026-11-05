# 验证混合方案完整管线：单格→CELL 免写；区域/多选→写 JYYJ_SEL；回单格→清 JYYJ_SEL。
# 模拟 add-in RenderNow 的真实行为，在 20000 行公式表上验证。
$ErrorActionPreference = "Stop"
$app = $null
for ($i = 0; $i -lt 40; $i++) {
    try { $app = New-Object -ComObject Excel.Application; break } catch { Start-Sleep -Seconds 1 }
}
if ($null -eq $app) { Write-Host "NO EXCEL INSTANCE"; exit 2 }
$app.Visible = $false
$app.DisplayAlerts = $false

function AssertEq([string]$name, [object]$actual, [object]$expect) {
    if ("$actual" -eq "$expect") { Write-Host "[PASS] $name = $actual" }
    else { Write-Host "[FAIL] $name = $actual 期望 $expect" }
}
function CFC([int]$r, [int]$c) {
    for ($t = 0; $t -lt 40; $t++) {
        try { return [int]$ws.Cells.Item($r, $c).DisplayFormat.Interior.ColorIndex } catch { Start-Sleep -Milliseconds 5 }
    }
    return -9999
}
# 由选区生成 predicate（与 C# BuildPredicate 同逻辑：逐区域十字）
function BuildPred([string]$range) {
    [void]$app.Range($range).Select()
    $sel = $app.Selection
    $clauses = @()
    for ($i = 1; $i -le $sel.Areas.Count; $i++) {
        $a = $sel.Areas.Item($i)
        $r1 = [int]$a.Row; $c1 = [int]$a.Column
        $r2 = $r1 + [int]$a.Rows.Count - 1; $c2 = $c1 + [int]$a.Columns.Count - 1
        $fullR = (($r2 - $r1 + 1) -ge 1048576)
        $fullC = (($c2 - $c1 + 1) -ge 16384)
        if ($fullR -and $fullC) { continue }
        if ($fullR) { $clauses += "AND(COLUMN()>=$c1,COLUMN()<=$c2)"; continue }
        if ($fullC) { $clauses += "AND(ROW()>=$r1,ROW()<=$r2)"; continue }
        $clauses += "OR(AND(ROW()>=$r1,ROW()<=$r2),AND(COLUMN()>=$c1,COLUMN()<=$c2))"
    }
    if ($clauses.Count -eq 0) { return "FALSE" }
    return "OR(" + ($clauses -join ",") + ")"
}
function SetSel([string]$pred) { $wb.Names.Item('JYYJ_SEL').RefersTo = "=$pred" }

$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$ws.Range("A1:A20000").Formula = '=ROW()*2'
$ws.Range("B1:B20000").Formula = '=SUM(R1C1:RC[-1])'
$app.Calculate()

$wb.Names.Add('JYYJ_VIEW_MODE', '=1') | Out-Null
$wb.Names.Add('JYYJ_R', '=CELL("row")') | Out-Null
$wb.Names.Add('JYYJ_C', '=CELL("col")') | Out-Null
$wb.Names.Add('JYYJ_SEL', '=FALSE') | Out-Null
$fc = $ws.Cells.FormatConditions.Add(2, 0, '=AND(JYYJ_VIEW_MODE,OR(OR(ROW()=JYYJ_R,COLUMN()=JYYJ_C),JYYJ_SEL))')
$fc.Interior.ColorIndex = 6; $fc.Interior.Pattern = 1

Write-Host "=== 1) 单格：JYYJ_SEL=FALSE，CELL 驱动即时 ==="
SetSel "FALSE"
$app.Range("B2").Select(); Start-Sleep -Milliseconds 100
AssertEq "B2 行 A2" (CFC 2 1) 6
AssertEq "B2 列 B1" (CFC 1 2) 6
AssertEq "A1 非黄" (CFC 1 1) (-4142)

Write-Host "`n=== 2) 区域 A1:B4 → 写 predicate ==="
$p = BuildPred "A1:B4"; SetSel $p; Start-Sleep -Milliseconds 150
Write-Host ("  predicate=$p")
AssertEq "区域 A3(行3)" (CFC 3 1) 6
AssertEq "区域 B1(列B)" (CFC 1 2) 6
AssertEq "区域 D4 内(行列都中)" (CFC 4 4) 6
AssertEq "区域 C5 非黄" (CFC 5 3) (-4142)

Write-Host "`n=== 3) 多选 A1:A5+C1:C5 → 逐区域十字 ==="
$p = BuildPred "A1:A5,C1:C5"; SetSel $p; Start-Sleep -Milliseconds 150
Write-Host ("  predicate=$p")
AssertEq "多选 A2" (CFC 2 1) 6
AssertEq "多选 C2" (CFC 2 3) 6
AssertEq "多选 B3(行3)" (CFC 3 2) 6
AssertEq "多选 B6 非黄" (CFC 6 2) (-4142)

Write-Host "`n=== 4) 回单格 D4 → JYYJ_SEL=FALSE，仅 CELL 十字 ==="
SetSel "FALSE"
$app.Range("D4").Select(); Start-Sleep -Milliseconds 100
AssertEq "单格 A4(行4)" (CFC 4 1) 6
AssertEq "单格 D1(列D)" (CFC 1 4) 6
AssertEq "旧多选 C2 已清" (CFC 2 3) (-4142)
AssertEq "旧多选 A2 已清" (CFC 2 1) (-4142)

Write-Host "`n=== 5) 单格连点延迟（CELL 免写） ==="
foreach ($cell in @("F6","H8","J10","L12","N14")) {
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $app.Range($cell).Select()
    $sw.Stop()
    $selMs = $sw.ElapsedMilliseconds
    Start-Sleep -Milliseconds 1
    $c = CFC ([int]$app.Range($cell).Row) 1
    Write-Host ("  ${cell}: Select=" + $selMs + "ms 1ms后行首色=" + $c + "（6=已即时上色）")
}

$wb.Close($false)
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws) | Out-Null
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null
$app.Quit()
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) | Out-Null
Write-Host "`nDONE"
