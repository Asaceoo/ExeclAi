# 终极解析测试：count=1 时 (Get-SentinelRuleCount $ws -eq 1) 与 count=0 时 (… -eq 0) 的解析差异
$ErrorActionPreference = "Stop"
$app = $null
for ($i = 0; $i -lt 40; $i++) { try { $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application"); break } catch { Start-Sleep -Seconds 1 } }
if ($null -eq $app) { Write-Host "FAIL no ROT"; exit 1 }

$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)

function Get-SentinelRuleCount($ws) {
    $n = 0
    for ($i = 1; $i -le $ws.Cells.FormatConditions.Count; $i++) {
        $f = [string]$ws.Cells.FormatConditions.Item($i).Formula1
        if ($f -like '*JYYJ_VIEW_MODE*') { $n++ }
    }
    return $n
}

# 情形1：count=0，内联 -eq 0
Write-Host "===== count=0 ====="
Write-Host "  \$ws -eq 0 = [$($ws -eq 0)] type=$($($ws -eq 0).GetType().Name)"
$v0 = Get-SentinelRuleCount $ws -eq 0
Write-Host "  inline value=[$v0] type=$($v0.GetType().Name)"

# 创建1条规则
$fc = $ws.Cells.FormatConditions.Add(2, 0, '=AND(JYYJ_VIEW_MODE,ROW()=1,COLUMN()=1)')
Write-Host "===== count=1 ====="
Write-Host "  \$ws -eq 1 = [$($ws -eq 1)] type=$($($ws -eq 1).GetType().Name)"
$v1 = Get-SentinelRuleCount $ws -eq 1
Write-Host "  inline value=[$v1] type=$($v1.GetType().Name)"

# 用 count=1 测 -eq 0（应 FAIL，模拟完整脚本关闭指针后场景：虽然理论上count=0）
# 先删规则
for ($i = $ws.Cells.FormatConditions.Count; $i -ge 1; $i--) {
    if ([string]$ws.Cells.FormatConditions.Item($i).Formula1 -like '*JYYJ_VIEW_MODE*') { $ws.Cells.FormatConditions.Item($i).Delete() }
}
Write-Host "===== 删除后 count=$($ws.Cells.FormatConditions.Count) ====="
$vD = Get-SentinelRuleCount $ws -eq 0
Write-Host "  inline value=[$vD] type=$($vD.GetType().Name)"

$wb.Close($false)
$app.Workbooks.Add() | Out-Null
Write-Host "DONE"
