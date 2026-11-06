# 验证混合方案：CF = AND(JYYJ_VIEW_MODE, OR(AND(ROW()=JYYJ_R,COLUMN()=JYYJ_C), JYYJ_SEL))
# 单格 → CELL() 免写即时跟随；区域/多选 → 写 JYYJ_SEL。大表(20000行公式)下测延迟。
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

$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$ws.Name = "Data"
# 快速填充 20000 行公式
$ws.Range("A1:A20000").Formula = '=ROW()*2'
$ws.Range("B1:B20000").Formula = '=SUM(R1C1:RC[-1])'
$ws.Range("C1:C20000").Formula = '=ROW()+COLUMN()'
$app.Calculate()

function CFC([int]$r, [int]$c) {
    for ($t = 0; $t -lt 40; $t++) {
        try { return [int]$ws.Cells.Item($r, $c).DisplayFormat.Interior.ColorIndex } catch { Start-Sleep -Milliseconds 5 }
    }
    return -9999
}
# 建立混合方案
$wb.Names.Add('JYYJ_VIEW_MODE', '=1') | Out-Null
$wb.Names.Add('JYYJ_R', '=CELL("row")') | Out-Null
$wb.Names.Add('JYYJ_C', '=CELL("col")') | Out-Null
$wb.Names.Add('JYYJ_SEL', '=FALSE') | Out-Null
$fc = $ws.Cells.FormatConditions.Add(2, 0, '=AND(JYYJ_VIEW_MODE,OR(OR(ROW()=JYYJ_R,COLUMN()=JYYJ_C),JYYJ_SEL))')
$fc.Interior.ColorIndex = 6; $fc.Interior.Pattern = 1

# 预点亮一次，让 CF 生效
$app.Range("A1").Select(); Start-Sleep -Milliseconds 100

Write-Host "=== 1) 单格连点 5 次（JYYJ_SEL 保持 FALSE，CELL 驱动）延迟 ==="
foreach ($cell in @("B2","D4","F6","H8","J10")) {
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $app.Range($cell).Select()
    $sw.Stop()
    $selMs = $sw.ElapsedMilliseconds
    # 立即读色（不轮询等待，看 Select 返回后立刻是否已上色）
    Start-Sleep -Milliseconds 1
    $c = CFC ([int]$app.Range($cell).Row) 1
    Write-Host ("  ${cell}: Select=" + $selMs + "ms 立即读行首色=" + $c + "（6=已即时上色）")
}
# 终态校验
AssertEq "J10 行内 A10" (CFC 10 1) 6
AssertEq "J10 列内 J1" (CFC 1 10) 6
AssertEq "旧区 A2 已恢复" (CFC 2 1) (-4142)

Write-Host "`n=== 2) 区域选择（A1:B4）→ 写 predicate，整个区域十字高亮 ==="
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$app.Range("A1:B4").Select()
$sw.Stop()
Write-Host ("  选区变化 Select=" + $sw.ElapsedMilliseconds + "ms（等待 poll 写 JYYJ_SEL）")
Start-Sleep -Milliseconds 300
AssertEq "区域 A3(行3)" (CFC 3 1) 6
AssertEq "区域 B1(列B)" (CFC 1 2) 6
AssertEq "区域 C5 非黄" (CFC 5 3) (-4142)

Write-Host "`n=== 3) 多选 A1:A5+C1:C5 → 逐区域十字 ==="
$app.Range("A1:A5,C1:C5").Select()
Start-Sleep -Milliseconds 300
AssertEq "多选 A2" (CFC 2 1) 6
AssertEq "多选 C2" (CFC 2 3) 6
AssertEq "多选 B6 非黄" (CFC 6 2) (-4142)

Write-Host "`n=== 4) 回单格 D4 → JYYJ_SEL=FALSE，仅 D4 十字 ==="
$app.Range("D4").Select()
Start-Sleep -Milliseconds 300
AssertEq "回单格 A4(行4)" (CFC 4 1) 6
AssertEq "回单格 D1(列D)" (CFC 1 4) 6
AssertEq "回单格 A2 恢复" (CFC 2 1) (-4142)

Write-Host "`n=== 5) 连点单格在区域/多选之后（无残留闪变） ==="
$app.Range("G7").Select()
Start-Sleep -Milliseconds 200
AssertEq "G7 行首 G7=第7行" (CFC 7 1) 6
AssertEq "G7 列内 G1" (CFC 1 7) 6
AssertEq "旧多选 C2 已恢复" (CFC 2 3) (-4142)

$wb.Close($false)
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws) | Out-Null
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null
$app.Quit()
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) | Out-Null
Write-Host "`nDONE"
