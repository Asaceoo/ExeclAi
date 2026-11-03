# 定位 Get-SentinelRuleCount 在 Count=0 时为何返回非 0。
# 怀疑：PowerShell $n++ 在 if 块内会向输出流发射旧值，污染函数返回值。
$ErrorActionPreference = "Stop"
$app = $null
for ($i = 0; $i -lt 40; $i++) { try { $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application"); break } catch { Start-Sleep -Seconds 1 } }
if ($null -eq $app) { Write-Host "FAIL no ROT"; exit 1 }

$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)

# 复刻验证脚本的 Get-SentinelRuleCount（含 $n++ 潜在输出污染）
function Get-SentinelRuleCount($ws) {
    $n = 0
    for ($i = 1; $i -le $ws.Cells.FormatConditions.Count; $i++) {
        $f = [string]$ws.Cells.FormatConditions.Item($i).Formula1
        if ($f -like '*JYYJ_VIEW_MODE*') { $n++ }
    }
    return $n
}

Write-Host "Count(no cast)=[$($ws.Cells.FormatConditions.Count)] type=$($ws.Cells.FormatConditions.Count.GetType().Name)"
Write-Host "Count([int] cast)=$([int]$ws.Cells.FormatConditions.Count)"

$r = Get-SentinelRuleCount $ws
Write-Host "function raw return: [$r] type=$($r.GetType().Name) Count-of-array=$(@($r).Count)"
Write-Host "  -eq 0 result: [$($r -eq 0)]"
Write-Host "  assert bool: $([bool]($r -eq 0))"

# 再验证带规则时
$fc = $ws.Cells.FormatConditions.Add(2, 0, '=AND(JYYJ_VIEW_MODE,ROW()=1,COLUMN()=1)')
$r2 = Get-SentinelRuleCount $ws
Write-Host "with 1 rule: raw=[$r2] type=$($r2.GetType().Name) arrayCount=$(@($r2).Count)  -eq 1: [$($r2 -eq 1)]"

$wb.Close($false)
$app.Workbooks.Add() | Out-Null
Write-Host "DONE"
