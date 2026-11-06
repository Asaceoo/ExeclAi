# 测试完整脚本的内联 Assert 调用形式：(Get-SentinelRuleCount $ws -eq 0)
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
function Assert([bool]$cond, [string]$msg) {
    if ($cond) { Write-Host "[PASS] $msg" } else { Write-Host "[FAIL] $msg" }
}

Write-Host "Count=$($ws.Cells.FormatConditions.Count)"

# 完整脚本的调用形式
Assert (Get-SentinelRuleCount $ws -eq 0) "内联: Count=0 应PASS"

# 显式括号对比
Assert ((Get-SentinelRuleCount $ws) -eq 0) "显式括号: Count=0 应PASS"

# 看内联形式实际传入 Assert 的是什么
$inline = Get-SentinelRuleCount $ws -eq 0
Write-Host "inline expression value: [$inline] type=$($inline.GetType().Name)"

$wb.Close($false)
$app.Workbooks.Add() | Out-Null
Write-Host "DONE"
