# 完整复现崩溃自愈并转储检查时刻的真实状态。
# 流程：杀Excel→重建残留文件→带文件重启→等25s→枚举重启后工作簿的 FormatConditions/Names。
$ErrorActionPreference = "Stop"

Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 3

$tmp = Join-Path $env:TEMP "jyyj_crash3.xlsx"
if (Test-Path $tmp) { Remove-Item $tmp -Force }

# 用独立 Excel 实例创建残留文件
$app = New-Object -ComObject Excel.Application
$app.Visible = $false
$app.DisplayAlerts = $false
$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$fc = $ws.Range("A1:F6").FormatConditions.Add(2, 0, '=AND(JYYJ_VIEW_MODE,ROW()=5,COLUMN()=5)')
$fc.Interior.ColorIndex = 6
$wb.Names.Add('JYYJ_VIEW_MODE', '=1') | Out-Null
$wb.SaveAs($tmp, 51)
Write-Host "残留文件已创建: $tmp"
$wb.Close($false)
$app.Quit()
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) | Out-Null
Start-Sleep -Seconds 2
Get-Process excel -ErrorAction SilentlyContinue | Where-Object { $_.StartTime -gt (Get-Date).AddMinutes(-1) } | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

# 重启 Excel 带残留文件
Start-Process excel -ArgumentList "`"$tmp`""
Write-Host "Excel 已带残留文件重启，等待 25s..."
Start-Sleep -Seconds 25

$app2 = $null
for ($i = 0; $i -lt 40; $i++) { try { $app2 = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application"); break } catch { Start-Sleep -Seconds 1 } }
if ($null -eq $app2) { Write-Host "FAIL no ROT"; exit 1 }

Write-Host "Workbooks.Count=$($app2.Workbooks.Count)"
for ($wi = 1; $wi -le $app2.Workbooks.Count; $wi++) {
    $wb2 = $app2.Workbooks.Item($wi)
    Write-Host "--- Workbook[$wi]: '$($wb2.Name)' ---"
    for ($si = 1; $si -le $wb2.Worksheets.Count; $si++) {
        $ws2 = $wb2.Worksheets.Item($si)
        $fcCount = $ws2.Cells.FormatConditions.Count
        Write-Host "  Sheet[$si]='$($ws2.Name)' FormatConditions.Count=$fcCount"
        for ($i = 1; $i -le $fcCount; $i++) {
            try {
                $f = [string]$ws2.Cells.FormatConditions.Item($i).Formula1
                Write-Host "    rule[$i] Type=$([int]$ws2.Cells.FormatConditions.Item($i).Type) Formula1=$f"
            } catch { Write-Host "    rule[$i] ACCESS ERR: $($_.Exception.Message)" }
        }
    }
    $names = @()
    for ($i = 1; $i -le $wb2.Names.Count; $i++) { $names += [string]$wb2.Names.Item($i).Name }
    Write-Host "  Names: $($names -join ', ')"
}
Write-Host "DONE"
