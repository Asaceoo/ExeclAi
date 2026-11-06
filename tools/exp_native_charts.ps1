$ErrorActionPreference = 'Continue'
$outDir = 'D:\execl1.0-wps\2026-09-11-9accc9f9\output\tasks\task_1789064509199_a0xtqh47'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
function Set-Cell($ws, [int]$r, [int]$c, $v) { if ($v -is [int]) { $v = [double]$v }; $ws.Cells.Item($r, $c).Value2 = $v }
try {
  $wb = $excel.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)

  Set-Cell $ws 1 1 'item'; Set-Cell $ws 1 2 'val'
  $d = @('a', 45, 'b', 30, 'c', 18, 'd', 12, 'e', 6, 'f', 4)
  for ($r = 0; $r -lt 6; $r++) {
    Set-Cell $ws ($r + 2) 1 $d[$r * 2]
    Set-Cell $ws ($r + 2) 2 $d[$r * 2 + 1]
  }
  Set-Cell $ws 1 4 'cat'; Set-Cell $ws 1 5 'val'
  $d2 = @('x1', 220, 'x2', 150, 'x3', 180, 'x4', 60, 'x5', 180, 'x6', 120, 'x7', 90)
  for ($r = 0; $r -lt 7; $r++) {
    Set-Cell $ws ($r + 2) 4 $d2[$r * 2]
    Set-Cell $ws ($r + 2) 5 $d2[$r * 2 + 1]
  }
  $rngAB = $ws.Range('A1:B7')
  $rngDE = $ws.Range('D1:E8')

  function Test-Case($name, $ctype, $rng) {
    $shp = $ws.Shapes.AddChart2(-1, $ctype, 100, 100, 400, 260)
    $coName = $shp.Name
    $co = $ws.ChartObjects().Item($coName)
    $ch = $co.Chart
    $ok = $true; $err = ''
    try { $ch.SetSourceData($rng, 2) } catch { $ok = $false; $err = $_.Exception.Message }
    Start-Sleep -Milliseconds 400
    $pts = -1; $frm = 'NOF'
    try { $pts = $ch.SeriesCollection().Item(1).Points().Count } catch {}
    try { $frm = $ch.SeriesCollection().Item(1).Formula } catch {}
    $png = Join-Path $outDir ('exp_' + $name + '.png')
    $pngOk = $false
    try { $co.Activate(); $ch.Export($png, 'PNG') | Out-Null; $pngOk = (Test-Path $png) } catch { $err = $err + ' EXPORT:' + $_.Exception.Message }
    $shp.Delete()
    Write-Output ("EXP=" + $name + " setsrc=" + $ok + " pts=" + $pts + " f=" + $frm + " png=" + $pngOk + " " + $err)
  }

  Test-Case 'waterfall_setsrc' 119 $rngAB
  Test-Case 'pareto_setsrc' 122 $rngAB
  Test-Case 'histogram_setsrc' 118 $rngAB
  Test-Case 'treemap_setsrc' 117 $rngDE
  Test-Case 'sunburst_setsrc' 120 $rngDE
  Test-Case 'boxwhisker_setsrc' 121 $rngAB

  $shp = $ws.Shapes.AddChart2(-1, 119, 100, 100, 400, 260)
  $coName = $shp.Name
  $co = $ws.ChartObjects().Item($coName)
  $ch = $co.Chart
  $ok = $true; $err = ''
  try { $ch.SeriesCollection().Item(1).Formula = '=SERIES(T!$B$1,T!$A$2:$A$7,T!$B$2:$B$7,1)' } catch { $ok = $false; $err = $_.Exception.Message }
  Start-Sleep -Milliseconds 400
  $pts = -1
  try { $pts = $ch.SeriesCollection().Item(1).Points().Count } catch {}
  $png = Join-Path $outDir 'exp_waterfall_formula.png'
  $pngOk = $false
  try { $co.Activate(); $ch.Export($png, 'PNG') | Out-Null; $pngOk = (Test-Path $png) } catch { $err = $err + ' EXPORT:' + $_.Exception.Message }
  $shp.Delete()
  Write-Output ("EXP=waterfall_formula setsrc=" + $ok + " pts=" + $pts + " png=" + $pngOk + " " + $err)

  $wb.Close($false)
} finally {
  $excel.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  [GC]::Collect()
}
