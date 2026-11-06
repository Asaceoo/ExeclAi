# Probe 4: real-machine COM check of the JYYJ_SEL design (single static CF rule + expression name).
# Single-cell cross / multi-select per-area / full-row / full-col / pointer box / cleanup.
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
for ($r = 1; $r -le 8; $r++) { for ($c = 1; $c -le 8; $c++) { $ws.Cells.Item($r, $c).Value2 = $r * 10 + $c } }

$fail = 0
function AssertEq([string]$name, [object]$actual, [object]$expect) {
    if ("$actual" -eq "$expect") { Write-Host "[PASS] $name = $actual" }
    else { Write-Host "[FAIL] $name = $actual 期望 $expect"; $script:fail++ }
}
function CFC([int]$r, [int]$c) {
    for ($t = 0; $t -lt 10; $t++) {
        try { return [int]$ws.Cells.Item($r, $c).DisplayFormat.Interior.ColorIndex } catch { Start-Sleep -Milliseconds 150 }
    }
    return -9999
}
function CFB([int]$r, [int]$c) {
    for ($t = 0; $t -lt 10; $t++) {
        try { return [int]$ws.Cells.Item($r, $c).DisplayFormat.Borders.Item(7).ColorIndex } catch { Start-Sleep -Milliseconds 150 }
    }
    return -9999
}
function SetSel([string]$pred) { $wb.Names.Item('JYYJ_SEL').RefersTo = "=$pred"; $ws.Calculate(); Start-Sleep -Milliseconds 200 }
function Pred-Single([int]$r, [int]$c) { return "OR(AND(ROW()>=$r,ROW()<=$r),AND(COLUMN()>=$c,COLUMN()<=$c))" }
function Pred-Multi([string]$areas) {
    [void]$app.Range($areas).Select()
    $sel = $app.Selection
    $clauses = @()
    for ($i = 1; $i -le $sel.Areas.Count; $i++) {
        $a = $sel.Areas.Item($i)
        $r1 = [int]$a.Row; $c1 = [int]$a.Column
        $r2 = $r1 + [int]$a.Rows.Count - 1; $c2 = $c1 + [int]$a.Columns.Count - 1
        $clauses += "OR(AND(ROW()>=$r1,ROW()<=$r2),AND(COLUMN()>=$c1,COLUMN()<=$c2))"
    }
    return "OR(" + ($clauses -join ",") + ")"
}

# 1. sentinel + JYYJ_SEL + single static CF rule
$wb.Names.Add('JYYJ_VIEW_MODE', '=1') | Out-Null
$wb.Names.Add('JYYJ_SEL', '=FALSE') | Out-Null
try {
    $fc = $ws.Cells.FormatConditions.Add(2, 0, '=AND(JYYJ_VIEW_MODE,JYYJ_SEL)')
    $fc.Interior.ColorIndex = 6
    $fc.Interior.Pattern = 1
    Write-Host "[PASS] 单条静态 CF 规则创建成功"
} catch { Write-Host "[FAIL] CF 创建失败: $($_.Exception.Message)"; $fail++ }

# 2. single cell B2 cross
SetSel (Pred-Single 2 2)
AssertEq "B2 行内 A2" (CFC 2 1) 6
AssertEq "B2 列内 B1" (CFC 1 2) 6
AssertEq "B2 对角 A1 无色" (CFC 1 1) (-4142)

# 3. move to D4 + old restore
SetSel (Pred-Single 4 4)
AssertEq "D4 行内 A4" (CFC 4 1) 6
AssertEq "D4 列内 D1" (CFC 1 4) 6
AssertEq "旧区 A2 恢复" (CFC 2 1) (-4142)

# 4. single rule assertion (zero churn)
$n = 0
for ($i = 1; $i -le $ws.Cells.FormatConditions.Count; $i++) { if ([string]$ws.Cells.FormatConditions.Item($i).Formula1 -like '*JYYJ_VIEW_MODE*') { $n++ } }
AssertEq "哨兵规则数" $n 1

# 5. multi-select A1:A5 + C1:C5 per-area cross
$pred = Pred-Multi "A1:A5,C1:C5"
Write-Host "  [DUMP] multi predicate=$pred"
SetSel $pred
AssertEq "多选 A列(2,1)" (CFC 2 1) 6
AssertEq "多选 C列(2,3)" (CFC 2 3) 6
AssertEq "多选 B6 非黄" (CFC 6 2) (-4142)
AssertEq "多选 B2(行2) 黄" (CFC 2 2) 6

# 6. pointer variant: borders box
for ($i = $ws.Cells.FormatConditions.Count; $i -ge 1; $i--) {
    $f = $ws.Cells.FormatConditions.Item($i)
    if ([string]$f.Formula1 -like '*JYYJ_VIEW_MODE*') { $f.Delete() }
}
$fp = $ws.Cells.FormatConditions.Add(2, 0, '=AND(JYYJ_VIEW_MODE,JYYJ_SEL)')
$fp.Borders.ColorIndex = 3
$fp.Borders.LineStyle = 1
SetSel "OR(AND(ROW()>=2,ROW()<=3,COLUMN()>=2,COLUMN()<=4))"
AssertEq "指针 B2 左边框红" (CFB 2 2) 3
AssertEq "指针 D3 左边框红" (CFB 3 3) 3
AssertEq "指针 E5 无红框" (CFB 5 5) (-4142)

# 7. cleanup: delete rule + names, 0 residual
for ($i = $ws.Cells.FormatConditions.Count; $i -ge 1; $i--) {
    $f = $ws.Cells.FormatConditions.Item($i)
    if ([string]$f.Formula1 -like '*JYYJ_VIEW_MODE*') { $f.Delete() }
}
foreach ($nm in @('JYYJ_VIEW_MODE','JYYJ_SEL')) {
    for ($i = $wb.Names.Count; $i -ge 1; $i--) { if ([string]$wb.Names.Item($i).Name -eq $nm) { $wb.Names.Item($i).Delete() } }
}
$n2 = 0
for ($i = 1; $i -le $ws.Cells.FormatConditions.Count; $i++) { if ([string]$ws.Cells.FormatConditions.Item($i).Formula1 -like '*JYYJ_VIEW_MODE*') { $n2++ } }
$nmLeft = @()
for ($i = 1; $i -le $wb.Names.Count; $i++) { if (([string]$wb.Names.Item($i).Name -like 'JYYJ_*')) { $nmLeft += [string]$wb.Names.Item($i).Name } }
AssertEq "清理后哨兵规则数" $n2 0
AssertEq "清理后 JYYJ 定义名残留" ($nmLeft -join ',') ""

Write-Host ""
Write-Host ("== PROBE4 FAILED=" + $fail + " ==")
$wb.Close($false)
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws) | Out-Null
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null
$app.Quit()
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) | Out-Null
if ($fail -gt 0) { exit 1 } else { exit 0 }
