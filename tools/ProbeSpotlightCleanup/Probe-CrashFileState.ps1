# 检查完整验证脚本失败后留下的 jyyj_view_crash_test.xlsx 磁盘状态（该文件是失败现场）
$ErrorActionPreference = "Stop"
$app = $null
for ($i = 0; $i -lt 40; $i++) { try { $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application"); break } catch { Start-Sleep -Seconds 1 } }
if ($null -eq $app) { Write-Host "FAIL no ROT"; exit 1 }

$tmp = Join-Path $env:TEMP "jyyj_view_crash_test.xlsx"
if (-not (Test-Path $tmp)) { Write-Host "crash file not found"; exit 1 }
Write-Host "opening $tmp"
$wb = $app.Workbooks.Open($tmp)
for ($si = 1; $si -le $wb.Worksheets.Count; $si++) {
    $ws = $wb.Worksheets.Item($si)
    $cnt = 0
    try { $cnt = [int]$ws.Cells.FormatConditions.Count } catch { Write-Host "  sheet[$si] Count ERR: $($_.Exception.Message)"; continue }
    Write-Host "Sheet[$si]='$($ws.Name)' FormatConditions.Count=$cnt"
    for ($i = 1; $i -le $cnt; $i++) {
        try { Write-Host "    rule[$i] Type=$([int]$ws.Cells.FormatConditions.Item($i).Type) Formula1=$([string]$ws.Cells.FormatConditions.Item($i).Formula1)" }
        catch { Write-Host "    rule[$i] ERR: $($_.Exception.Message)" }
    }
}
$names = @()
for ($i = 1; $i -le $wb.Names.Count; $i++) { $names += [string]$wb.Names.Item($i).Name }
Write-Host "Names: $($names -join ', ')"
$wb.Close($false)
Write-Host "DONE"
