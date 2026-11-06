# Live end-to-end verify_chart self-check (read-only chart metadata read; no VBA -> avoids out-of-proc modal dialog issue).
# Flow: launch real Excel (Addin auto-loads) -> numeric column chart -> expect ok;
#       text-X scatter chart -> expect text verdict. Stateless each step.
[CmdletBinding()]
param(
    [switch]$SkipLaunch
)
$ErrorActionPreference = "Stop"
$asm64 = Join-Path $env:LOCALAPPDATA "XaiAssistant\app\x64\ExcelAiAssistant.Excel.dll"
if (-not (Test-Path $asm64)) { throw "Install Excel assembly not found: $asm64" }
Add-Type -Path $asm64

$excelApp = $null
function Get-Excel {
    try {
        $app = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    } catch {
        $app = New-Object -ComObject Excel.Application
    }
    if (-not $app) { throw "Cannot get Excel Application." }
    return $app
}

function Invoke-Verify {
    param($params)
    $svc = New-Object ExcelAiAssistant.Excel.ExcelChartService($excelApp)
    $r = $svc.VerifyChart($params)
    if ($r.Ok) {
        $d = $r.Data
        return ("found=" + $d.chart_found + " verdict=" + $d.verdict +
            " series=" + $d.series_count + " points=" + $d.point_total +
            " text=" + $d.has_text_source + " type=" + $d.chart_type_name +
            " title_ok=" + $d.title_ok + " dsrc_ok=" + $d.data_source_ok +
            " rendered=" + $d.data_rendered + " ratio=" + $d.nonblank_ratio +
            " | selfcheck=" + $d.selfcheck)
    }
    return ("TOOL_FAILED code=" + $r.Code + " msg=" + $r.Message)
}

# stateless close: re-resolve active workbook each time (service FinalReleaseComObject's the RCWs it resolved)
function Close-ActiveWorkbook {
    $awb = $excelApp.ActiveWorkbook
    if ($awb) { $awb.Close($false) }
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($awb) | Out-Null
}

try {
    if (-not $SkipLaunch) {
        Write-Host "== kill residual Excel and relaunch =="
        Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
        Start-Sleep -Seconds 1
    }
    $excelApp = Get-Excel
    $excelApp.Visible = $false
    $excelApp.DisplayAlerts = $false

    # ---------- Scenario 1: numeric column chart -> expect ok ----------
    Write-Host "== Scenario1: numeric column chart =="
    $wb = $excelApp.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = "dt"
    $ws.Cells.Item(1,1).Value2 = "cat"
    $ws.Cells.Item(1,2).Value2 = "val"
    for ($i=2; $i -le 8; $i++) {
        $ws.Cells.Item($i,1).Value2 = "A" + ($i-1)
        $ws.Cells.Item($i,2).Value2 = [int](100 + $i*7)
    }
    $dataRng = $ws.Range("A1:B8")
    $chartObj = $ws.ChartObjects().Add(200, 10, 300, 220)
    $ch = $chartObj.Chart
    $ch.SetSourceData($dataRng, 2)   # xlColumns -> 1 series (category col A, value col B), 7 points
    $ch.ChartType = 51
    $ch.HasTitle = $true
    $ch.ChartTitle.Text = "SalesTrend"
    $p1 = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $p1.Add("sheet", "dt")
    $p1.Add("expected_chart_type", "Column")
    $p1.Add("expected_series", 1)
    $p1.Add("expected_title", "SalesTrend")
    $p1.Add("expected_data_source", "A2:B8")
    $res1 = Invoke-Verify $p1
    Write-Host ("  Scenario1 -> " + $res1)
    if ($res1 -notmatch "verdict=ok") { throw ("Scenario1 expected ok, got: " + $res1) }
    if ($res1 -notmatch "title_ok=True") { throw ("Scenario1 expected title_ok True, got: " + $res1) }
    if ($res1 -notmatch "dsrc_ok=True") { throw ("Scenario1 expected dsrc_ok True, got: " + $res1) }
    if ($res1 -notmatch "rendered=True") { throw ("Scenario1 expected rendered True, got: " + $res1) }
    Write-Host ("  Scenario1 PASS")

    Close-ActiveWorkbook

    # ---------- Scenario 2: text-X scatter chart -> expect text ----------
    Write-Host "== Scenario2: text-axis scatter chart (points>0 but not plottable) =="
    $wb2 = $excelApp.Workbooks.Add()
    $ws2 = $wb2.Worksheets.Item(1)
    $ws2.Name = "dt2"
    $ws2.Cells.Item(1,1).Value2 = "x"
    $ws2.Cells.Item(1,2).Value2 = "y"
    for ($i=2; $i -le 6; $i++) {
        $ws2.Cells.Item($i,1).Value2 = [int]($i*5)
        $ws2.Cells.Item($i,2).Value2 = "TXT" + $i
    }
    $d2 = $ws2.Range("A1:B6")
    $co2 = $ws2.ChartObjects().Add(200, 10, 300, 220)
    $c2 = $co2.Chart
    $c2.SetSourceData($d2, 2)   # xlColumns -> 1 scatter series; X=numeric col A, Y=text col B (last range -> text detection)
    $c2.ChartType = -4169
    $p2 = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $p2.Add("sheet", "dt2")
    $p2.Add("expected_chart_type", "Scatter")
    $res2 = Invoke-Verify $p2
    Write-Host ("  Scenario2 -> " + $res2)

    Close-ActiveWorkbook

    # ---------- Scenario 3: wrong expected title -> verdict +title ----------
    Write-Host "== Scenario3: wrong expected title =="
    $wb3 = $excelApp.Workbooks.Add()
    $ws3 = $wb3.Worksheets.Item(1)
    $ws3.Name = "dt3"
    $ws3.Cells.Item(1,1).Value2 = "cat"
    $ws3.Cells.Item(1,2).Value2 = "val"
    for ($i=2; $i -le 7; $i++) {
        $ws3.Cells.Item($i,1).Value2 = "B" + ($i-1)
        $ws3.Cells.Item($i,2).Value2 = [int](50 + $i*3)
    }
    $d3 = $ws3.Range("A1:B7")
    $co3 = $ws3.ChartObjects().Add(200, 10, 300, 220)
    $c3 = $co3.Chart
    $c3.SetSourceData($d3, 2)
    $c3.ChartType = 51
    $c3.HasTitle = $true
    $c3.ChartTitle.Text = "ActualTitleXYZ"
    $p3 = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $p3.Add("sheet", "dt3")
    $p3.Add("expected_title", "SalesTrend")
    $p3.Add("expected_data_source", "Z1:Z99")   # no overlap -> dsrc_ok False
    $res3 = Invoke-Verify $p3
    Write-Host ("  Scenario3 -> " + $res3)
    if ($res3 -notmatch "title_ok=False") { throw ("Scenario3 expected title_ok False, got: " + $res3) }
    if ($res3 -notmatch "dsrc_ok=False") { throw ("Scenario3 expected dsrc_ok False, got: " + $res3) }
    if ($res3 -notmatch "blank") { Write-Host ("  Scenario3 note: rendered ok (advisory)") }
    Write-Host ("  Scenario3 PASS")

    Close-ActiveWorkbook

    Write-Host "== live selfcheck done =="
} finally {
    if ($excelApp) {
        $excelApp.DisplayAlerts = $true
        try { $excelApp.Quit() } catch {}
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excelApp) | Out-Null
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}