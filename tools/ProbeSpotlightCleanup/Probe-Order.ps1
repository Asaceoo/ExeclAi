# 探针：验证「先删定义名 → 再删 CF 规则」的顺序问题。
# 假设：删掉 JYYJ_VIEW_MODE 定义名后，引用它的 CF 规则 Formula1 可能被 Excel 改写/失效，
# 导致后续按文本匹配删规则失败（这可能是 VerifySpotlight 关闭/自愈两处 FAIL 的根因）。
$ErrorActionPreference = "Stop"

$app = $null
for ($i = 0; $i -lt 40; $i++) {
    try { $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application"); break } catch { Start-Sleep -Seconds 1 }
}
if ($null -eq $app) { Write-Host "FAIL: no Excel ROT"; exit 1 }

$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)

# 1. 注入哨兵规则 + 定义名
$fc = $ws.Range("A1:F6").FormatConditions.Add(2, 0, '=AND(JYYJ_VIEW_MODE,ROW()=5,COLUMN()=5)')
$fc.Interior.ColorIndex = 6
$wb.Names.Add('JYYJ_VIEW_MODE', '=1') | Out-Null

function Dump-Rules([string]$tag) {
    $n = 0
    Write-Host "--- $tag : FormatConditions.Count=$($ws.Cells.FormatConditions.Count)"
    for ($i = 1; $i -le $ws.Cells.FormatConditions.Count; $i++) {
        $f = $null
        try { $f = [string]$ws.Cells.FormatConditions.Item($i).Formula1 } catch { $f = "<ERR: $($_.Exception.Message)>" }
        Write-Host "   rule[$i] Formula1=$f"
        if ($f -like '*JYYJ_VIEW_MODE*') { $n++ }
    }
    Write-Host "   sentinelCount=$n"
    return $n
}

$before = Dump-Rules "注入后"
Write-Host "injected sentinelCount=$before"

# 2. 模拟 RemoveMarkedRules 前先删定义名（当前 C# CleanupWorkbook 顺序：TryRemoveSentinelName 先跑）
for ($i = $wb.Names.Count; $i -ge 1; $i--) {
    $n = $wb.Names.Item($i)
    if ([string]$n.Name -eq 'JYYJ_VIEW_MODE') { $n.Delete(); Write-Host "  deleted name JYYJ_VIEW_MODE" }
}

# 3. 枚举规则，看 Formula1 是否仍含哨兵 token
$after = Dump-Rules "删定义名后（顺序A：先删名后删规则）"

# 4. 按文本匹配删除剩余规则（模拟 RemoveMarkedRules）
for ($i = $ws.Cells.FormatConditions.Count; $i -ge 1; $i--) {
    try {
        $f = [string]$ws.Cells.FormatConditions.Item($i).Formula1
        if ($f -like '*JYYJ_VIEW_MODE*') {
            $ws.Cells.FormatConditions.Item($i).Delete()
            Write-Host "  deleted rule[$i]"
        }
    } catch { Write-Host "  delete rule[$i] ERR: $($_.Exception.Message)" }
}
$final = Dump-Rules "删除后"

$wb.Close($false)
$app.Workbooks.Add() | Out-Null  # 保证至少一个工作簿，避免 Excel 空白
Write-Host "RESULT: before=$before afterNameDelete=$after final=$final"
