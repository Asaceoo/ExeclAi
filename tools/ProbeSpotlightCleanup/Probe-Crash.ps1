# 探针3：复现「崩溃自愈」场景——打开磁盘上含哨兵规则+定义名的文件，执行 C# 清理逻辑，确认规则是否可删。
$ErrorActionPreference = "Stop"

$app = $null
for ($i = 0; $i -lt 40; $i++) {
    try { $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application"); break } catch { Start-Sleep -Seconds 1 }
}
if ($null -eq $app) { Write-Host "FAIL: no Excel ROT"; exit 1 }

$tmp = Join-Path $env:TEMP "jyyj_view_crash_test.xlsx"
Write-Host "opening $tmp"
$wb = $app.Workbooks.Open($tmp)
$ws = $wb.Worksheets.Item(1)

# 1. 检查名称
$nameExists = $false
for ($i = 1; $i -le $wb.Names.Count; $i++) {
    if ([string]$wb.Names.Item($i).Name -eq 'JYYJ_VIEW_MODE') { $nameExists = $true; break }
}
Write-Host "nameExists=$nameExists  Names.Count=$($wb.Names.Count)"

# 2. 检查规则
Write-Host "FormatConditions.Count=$($ws.Cells.FormatConditions.Count)"
for ($i = 1; $i -le $ws.Cells.FormatConditions.Count; $i++) {
    try {
        $f = [string]$ws.Cells.FormatConditions.Item($i).Formula1
        Write-Host "  rule[$i] Type=$([int]$ws.Cells.FormatConditions.Item($i).Type) Formula1=$f"
    } catch { Write-Host "  rule[$i] ACCESS ERR: $($_.Exception.Message)" }
}

# 3. 执行 C# 清理逻辑（TryRemoveSentinelName → RemoveMarkedRules）
Write-Host "== 执行 C# 清理逻辑 =="
for ($i = $wb.Names.Count; $i -ge 1; $i--) {
    $n = $wb.Names.Item($i)
    if ([string]$n.Name -eq 'JYYJ_VIEW_MODE') { $n.Delete(); Write-Host "  deleted name" }
}
$count = $ws.Cells.FormatConditions.Count
for ($i = $count; $i -ge 1; $i--) {
    try {
        $fc = $ws.Cells.FormatConditions.Item($i)
        if ([int]$fc.Type -ne 2) { Write-Host "  rule[$i] skip non-expression"; continue }
        $f1 = [string]$fc.Formula1
        Write-Host "  rule[$i] Type=2 Formula1=$f1"
        if ($f1 -like '*JYYJ_VIEW_MODE*') {
            try { $fc.Delete(); Write-Host "    DELETED" } catch { Write-Host "    DELETE ERR: $($_.Exception.Message)" }
        }
    } catch { Write-Host "  rule[$i] ACCESS ERR: $($_.Exception.Message)" }
}
Write-Host "  after: FormatConditions.Count=$($ws.Cells.FormatConditions.Count)"

$wb.Close($false)
Write-Host "DONE"
