# 探针2：用真实指针公式(含CELL()易失函数)+完整复刻 C# RemoveMarkedRules 逻辑（Type过滤+逆序遍历）。
# 目标：找出为何 VerifySpotlight 中指针关闭后哨兵规则仍残留。
$ErrorActionPreference = "Stop"

$app = $null
for ($i = 0; $i -lt 40; $i++) {
    try { $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application"); break } catch { Start-Sleep -Seconds 1 }
}
if ($null -eq $app) { Write-Host "FAIL: no Excel ROT"; exit 1 }

$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)

# --- 指针 ON：建规则 + 定义名（同 SpotlightService.EnsureRule/EnsureSentinel）---
$formula = '=AND(JYYJ_VIEW_MODE,CELL("row")=ROW(),CELL("col")=COLUMN())'
$fc = $ws.Cells.FormatConditions.Add(2, 0, $formula)
$fc.Borders.ColorIndex = 3
$fc.Borders.LineStyle = 1
$wb.Names.Add('JYYJ_VIEW_MODE', '=1') | Out-Null
Write-Host "pointer ON: FormatConditions.Count=$($ws.Cells.FormatConditions.Count) rules"
for ($i = 1; $i -le $ws.Cells.FormatConditions.Count; $i++) {
    Write-Host "  rule[$i] Type=$([int]$ws.Cells.FormatConditions.Item($i).Type) Formula1=$([string]$ws.Cells.FormatConditions.Item($i).Formula1)"
}

# --- 指针 OFF：复刻 CleanupWorkbook（先删名）→ RemoveMarkedRules（Type过滤+逆序）---
Write-Host "== 复刻 C# 清理逻辑 =="
for ($i = $wb.Names.Count; $i -ge 1; $i--) {
    $n = $wb.Names.Item($i)
    if ([string]$n.Name -eq 'JYYJ_VIEW_MODE') { $n.Delete(); Write-Host "  deleted name" }
}
$count = $ws.Cells.FormatConditions.Count
Write-Host "  FormatConditions.Count=$count"
for ($i = $count; $i -ge 1; $i--) {
    $fc = $null
    try {
        $fc = $ws.Cells.FormatConditions.Item($i)
        $t = [int]$fc.Type
        Write-Host "  rule[$i] Type=$t"
        if ($t -ne 2) { Write-Host "    skip (not expression)"; continue }
        $f1 = [string]$fc.Formula1
        Write-Host "    Formula1=$f1"
        if ($f1 -like '*JYYJ_VIEW_MODE*') {
            try { $fc.Delete(); Write-Host "    DELETED" }
            catch { Write-Host "    DELETE ERR: $($_.Exception.Message)" }
        }
    } catch { Write-Host "  rule[$i] ACCESS ERR: $($_.Exception.Message)" }
}
Write-Host "  after cleanup: FormatConditions.Count=$($ws.Cells.FormatConditions.Count)"
for ($i = 1; $i -le $ws.Cells.FormatConditions.Count; $i++) {
    Write-Host "  remaining rule[$i] Type=$([int]$ws.Cells.FormatConditions.Item($i).Type) Formula1=$([string]$ws.Cells.FormatConditions.Item($i).Formula1)"
}

$wb.Close($false)
$app.Workbooks.Add() | Out-Null
Write-Host "DONE"
