# 探针 3：复现 RemoveMarkedRules 的 COMException —— 找出枚举整表 CF 规则最稳健的方式。
$ErrorActionPreference = "Stop"
$app = $null
for ($i = 0; $i -lt 40; $i++) {
    try { $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application"); break } catch { Start-Sleep -Seconds 1 }
}
if ($null -eq $app) { Write-Host "NO EXCEL INSTANCE"; exit 2 }

$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
for ($r = 1; $r -le 6; $r++) { for ($c = 1; $c -le 6; $c++) { $ws.Cells.Item($r, $c).Value2 = $r * 10 + $c } }

# 注入两条哨兵规则（模拟真实使用残留）
$fc1 = $ws.Range("A1:F6").FormatConditions.Add(2, 0, '=AND(JYYJ_VIEW_MODE,OR(CELL("row")=ROW(),CELL("col")=COLUMN()))')
$fc1.Interior.ColorIndex = 6
$fc1.Interior.Pattern = 1
$fc2 = $ws.Range("C3").FormatConditions.Add(2, 0, '=AND(JYYJ_VIEW_MODE,CELL("row")=ROW(),CELL("col")=COLUMN())')
$fc2.Borders.ColorIndex = 3
$fc2.Borders.LineStyle = 1

# 方式 A：sheet.Cells.FormatConditions（当前 C# 代码路径）
Write-Host "=== 方式A: sheet.Cells.FormatConditions ==="
try {
    $coll = $ws.Cells.FormatConditions
    $n = [int]$coll.Count
    Write-Host ("  Count=" + $n)
    for ($i = 1; $i -le $n; $i++) {
        try {
            $f = $coll.Item($i)
            Write-Host ("  item[" + $i + "] Formula1=[" + [string]$f.Formula1 + "]")
        } catch { Write-Host ("  item[" + $i + "] READ FAIL: " + $_.Exception.GetType().Name + " " + $_.Exception.Message) }
    }
} catch { Write-Host ("  COLLECTION FAIL: " + $_.Exception.GetType().Name + " " + $_.Exception.Message) }

# 方式 B：sheet.UsedRange.FormatConditions
Write-Host "=== 方式B: sheet.UsedRange.FormatConditions ==="
try {
    $coll = $ws.UsedRange.FormatConditions
    $n = [int]$coll.Count
    Write-Host ("  Count=" + $n)
    for ($i = 1; $i -le $n; $i++) {
        try { Write-Host ("  item[" + $i + "] Formula1=[" + [string]$coll.Item($i).Formula1 + "]") }
        catch { Write-Host ("  item[" + $i + "] READ FAIL: " + $_.Exception.GetType().Name) }
    }
} catch { Write-Host ("  COLLECTION FAIL: " + $_.Exception.GetType().Name + " " + $_.Exception.Message) }

# 方式 C：遍历 AppliesTo 反查（每格 FormatConditions）
Write-Host "=== 方式C: 逐格 Range.FormatConditions（模拟公式路径） ==="
try {
    $fc = $ws.Range("A1").FormatConditions
    Write-Host ("  A1.FormatConditions.Count=" + [int]$fc.Count)
    for ($i = 1; $i -le [int]$fc.Count; $i++) {
        try { Write-Host ("  A1 item[" + $i + "] Formula1=[" + [string]$fc.Item($i).Formula1 + "]") }
        catch { Write-Host ("  A1 item[" + $i + "] READ FAIL: " + $_.Exception.GetType().Name) }
    }
} catch { Write-Host ("  A1 COLLECTION FAIL: " + $_.Exception.GetType().Name) }

# 方式 D：删除测试 —— 用方式B枚举并删除
Write-Host "=== 方式D: 用 UsedRange 枚举 + 逆序删除 ==="
try {
    $coll = $ws.UsedRange.FormatConditions
    for ($i = [int]$coll.Count; $i -ge 1; $i--) {
        try {
            $f = $coll.Item($i)
            $fml = [string]$f.Formula1
            if ($fml -like '*JYYJ_VIEW_MODE*') { $f.Delete(); Write-Host ("  已删 item[" + $i + "] fml=[" + $fml + "]") }
            else { Write-Host ("  跳过 item[" + $i + "] fml=[" + $fml + "]") }
        } catch { Write-Host ("  item[" + $i + "] DELETE/READ FAIL: " + $_.Exception.GetType().Name + " " + $_.Exception.Message) }
    }
} catch { Write-Host ("  D COLLECTION FAIL: " + $_.Exception.GetType().Name) }

# 验证剩余
$left = 0
try {
    $coll = $ws.UsedRange.FormatConditions
    for ($i = 1; $i -le [int]$coll.Count; $i++) { if ([string]$coll.Item($i).Formula1 -like '*JYYJ_VIEW_MODE*') { $left++ } }
} catch { $left = -1 }
Write-Host ("  删除后剩余哨兵规则=" + $left + " 期望0")

$wb.Close($false)
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws) | Out-Null
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null
