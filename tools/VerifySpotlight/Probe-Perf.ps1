# 探针 6：整表单规则 + ws.Calculate 性能实测（大表是否卡顿）。
$ErrorActionPreference = "Stop"
$app = $null
for ($i = 0; $i -lt 40; $i++) {
    try { $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application"); break } catch { Start-Sleep -Seconds 1 }
}
if ($null -eq $app) { Write-Host "NO EXCEL INSTANCE"; exit 2 }

$sw = [System.Diagnostics.Stopwatch]::StartNew()
$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
# 100k 数据格
$data = New-Object 'object[,]' 1000,100
for ($r = 0; $r -lt 1000; $r++) { for ($c = 0; $c -lt 100; $c++) { $data[$r,$c] = $r * 100 + $c } }
$ws.Range("A1:CJ1000").Value2 = $data
Write-Host ("填 100k 格耗时 " + $sw.ElapsedMilliseconds + "ms")

$wb.Names.Add('JYYJ_VIEW_MODE', '=1') | Out-Null
$fc = $ws.Cells.FormatConditions.Add(2, 0, '=AND(JYYJ_VIEW_MODE,OR(CELL("row")=ROW(),CELL("col")=COLUMN()))')
$fc.Interior.ColorIndex = 6
$fc.Interior.Pattern = 1

# 预热
$app.Range("B2").Select(); $ws.Calculate()

$times = @()
for ($k = 0; $k -lt 10; $k++) {
    $r = 10 + $k; $c = 10 + $k
    $app.Range($ws.Cells.Item($r,$c), $ws.Cells.Item($r,$c)).Select()
    $t0 = [System.Diagnostics.Stopwatch]::StartNew()
    $ws.Calculate()
    $t0.Stop()
    $times += $t0.ElapsedMilliseconds
}
$avg = ($times | Measure-Object -Average).Average
$max = ($times | Measure-Object -Maximum).Maximum
Write-Host ("ws.Calculate 10 次: avg=" + [math]::Round($avg,1) + "ms max=" + $max + "ms")
Write-Host ("最后高亮跟随检查: A11(行内)=" + [int]$ws.Cells.Item(11,1).DisplayFormat.Interior.ColorIndex + " 期望6")

$wb.Close($false)
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws) | Out-Null
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null
Write-Host ("总耗时 " + $sw.ElapsedMilliseconds + "ms")
