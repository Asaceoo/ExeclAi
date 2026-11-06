# 最小复现：创建指针式规则+定义名 → 模拟清理删除 → 用 VerifySpotlight 相同 PowerShell 模式读 FormatConditions。
# 目标：确认验证脚本 Get-SentinelRuleCount 读取是否与加载项清理结果一致。
$ErrorActionPreference = "Stop"

$app = $null
for ($i = 0; $i -lt 40; $i++) {
    try { $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application"); break } catch { Start-Sleep -Seconds 1 }
}
if ($null -eq $app) { Write-Host "FAIL no ROT"; exit 1 }

$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)

# 创建指针规则（同加载项 EnsureRule）+ 定义名
$formula = '=AND(JYYJ_VIEW_MODE,CELL("row")=ROW(),CELL("col")=COLUMN())'
$fc = $ws.Cells.FormatConditions.Add(2, 0, $formula)
$fc.Borders.ColorIndex = 3
$fc.Borders.LineStyle = 1
$wb.Names.Add('JYYJ_VIEW_MODE', '=1') | Out-Null

function Get-SentinelRuleCount($ws) {
    $n = 0
    for ($i = 1; $i -le $ws.Cells.FormatConditions.Count; $i++) {
        $f = [string]$ws.Cells.FormatConditions.Item($i).Formula1
        if ($f -like '*JYYJ_VIEW_MODE*') { $n++ }
    }
    return $n
}

Write-Host "before cleanup: Count=$($ws.Cells.FormatConditions.Count) sentinelCount=$(Get-SentinelRuleCount $ws)"

# 模拟加载项清理：删名 → 删规则
for ($i = $wb.Names.Count; $i -ge 1; $i--) {
    $n = $wb.Names.Item($i)
    if ([string]$n.Name -eq 'JYYJ_VIEW_MODE') { $n.Delete() }
}
$count = $ws.Cells.FormatConditions.Count
for ($i = $count; $i -ge 1; $i--) {
    $fc = $ws.Cells.FormatConditions.Item($i)
    if ([int]$fc.Type -ne 2) { continue }
    $f1 = [string]$fc.Formula1
    if ($f1 -like '*JYYJ_VIEW_MODE*') { $fc.Delete() }
}

Write-Host "after cleanup: Count=$($ws.Cells.FormatConditions.Count)"
Write-Host "sentinelCount via function=$(Get-SentinelRuleCount $ws)"

# 直接枚举验证
$remaining = 0
for ($i = 1; $i -le $ws.Cells.FormatConditions.Count; $i++) {
    Write-Host "  rule[$i] Type=$([int]$ws.Cells.FormatConditions.Item($i).Type) Formula1=$([string]$ws.Cells.FormatConditions.Item($i).Formula1)"
    $remaining++
}
Write-Host "remaining rules=$remaining"

$wb.Close($false)
$app.Workbooks.Add() | Out-Null
Write-Host "DONE"
