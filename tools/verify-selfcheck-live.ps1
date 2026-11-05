# Live end-to-end self-check for the four verify tools:
#   verify_cell_format / verify_conditional_format / verify_sort_filter / verify_formula
# Loads the Debug build (contains the new tools). Stateless each scenario. ASCII-only for PS5.1.
[CmdletBinding()]
param(
    [switch]$SkipLaunch
)
$ErrorActionPreference = "Stop"
$asmDir = "d:\execl1.0\src\ExcelAiAssistant.Excel\bin\Debug\net48"
$asmCore = Join-Path $asmDir "ExcelAiAssistant.Core.dll"
$asmExcel = Join-Path $asmDir "ExcelAiAssistant.Excel.dll"
if (-not (Test-Path $asmExcel)) { throw "Excel assembly not found: $asmExcel" }
[Reflection.Assembly]::LoadFrom($asmCore) | Out-Null
Add-Type -Path $asmExcel

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

function Invoke-Tool {
    param($tool, $params)
    $svc = New-Object ExcelAiAssistant.Excel.ExcelVerifyService($excelApp)
    try {
        $r = $svc.$tool($params)
        if ($r.Ok) {
            $d = $r.Data
            return ("verdict=" + $d.verdict + " | selfcheck=" + $d.selfcheck)
        }
        return ("TOOL_FAILED code=" + $r.ErrorCode + " msg=" + $r.Data)
    } catch {
        return ("THROW " + $_.Exception.ToString())
    }
}

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

    # ---------- Scenario 1: verify_cell_format (match) ----------
    Write-Host "== Scenario1: cell format match =="
    $wb = $excelApp.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = "fmt"
    $rng = $ws.Range("A1")
    $rng.NumberFormat = "0.00"
    $rng.Font.Bold = $true
    $rng.Font.Italic = $false
    $rng.Font.Name = "Arial"
    $rng.Font.Size = 12
    $rng.Font.Color = 255            # red
    $rng.Interior.Color = 65535      # yellow
    $rng.HorizontalAlignment = -4131 # left
    $rng.VerticalAlignment = -4107   # bottom
    $rng.Borders.LineStyle = 1       # continuous
    $rng.Locked = $true
    $p = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $p.Add("sheet", "fmt")
    $p.Add("address", "A1")
    $p.Add("expected_number_format", "0.00")
    $p.Add("expected_bold", $true)
    $p.Add("expected_italic", $false)
    $p.Add("expected_font_name", "Arial")
    $p.Add("expected_font_size", 12)
    $p.Add("expected_font_color", "#FF0000")
    $p.Add("expected_back_color", "#FFFF00")
    $p.Add("expected_h_align", "left")
    $p.Add("expected_v_align", "bottom")
    $p.Add("expected_border_style", "continuous")
    $p.Add("expected_locked", $true)
    $res = Invoke-Tool "VerifyCellFormat" $p
    Write-Host ("  Scenario1 -> " + $res)
    if ($res -notmatch "verdict=ok") { throw ("Scenario1 expected ok, got: " + $res) }
    Write-Host "  Scenario1 PASS"

    Close-ActiveWorkbook

    # ---------- Scenario 2: verify_cell_format (mismatch -> flags) ----------
    Write-Host "== Scenario2: cell format mismatch =="
    $wb2 = $excelApp.Workbooks.Add()
    $ws2 = $wb2.Worksheets.Item(1)
    $ws2.Name = "fmt2"
    $r2 = $ws2.Range("B2")
    $r2.NumberFormat = "#,##0"
    $r2.Font.Bold = $false
    $r2.Interior.Color = 0
    $p2 = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $p2.Add("sheet", "fmt2")
    $p2.Add("address", "B2")
    $p2.Add("expected_number_format", "0.00")
    $p2.Add("expected_bold", $true)
    $p2.Add("expected_back_color", "#FFFF00")
    $res2 = Invoke-Tool "VerifyCellFormat" $p2
    Write-Host ("  Scenario2 -> " + $res2)
    if ($res2 -notmatch "verdict=number") { throw ("Scenario2 expected number flag, got: " + $res2) }
    if ($res2 -notmatch "bold") { throw ("Scenario2 expected bold flag, got: " + $res2) }
    if ($res2 -notmatch "fill") { throw ("Scenario2 expected fill flag, got: " + $res2) }
    Write-Host "  Scenario2 PASS"

    Close-ActiveWorkbook

    # ---------- Scenario 3: verify_conditional_format (greater match) ----------
    Write-Host "== Scenario3: conditional format greater match =="
    $wb3 = $excelApp.Workbooks.Add()
    $ws3 = $wb3.Worksheets.Item(1)
    $ws3.Name = "cf"
    1..5 | ForEach-Object { $ws3.Cells.Item($_,1).Value2 = [int]($_*50) }
    $r3 = $ws3.Range("A1:A5")
    $fc = $r3.FormatConditions.Add(1, 5, "100")   # xlCellValue, xlGreater, >100
    $fc.Interior.Color = 255
    $fc.Font.Color = 16777215
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($fc) | Out-Null
    $p3 = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $p3.Add("sheet", "cf")
    $p3.Add("address", "A1:A5")
    $p3.Add("expected_rule", "greater")
    $p3.Add("expected_formula1", "100")
    $p3.Add("expected_fill_color", "#FF0000")
    $p3.Add("expected_font_color", "#FFFFFF")
    $res3 = Invoke-Tool "VerifyConditionalFormat" $p3
    Write-Host ("  Scenario3 -> " + $res3)
    if ($res3 -notmatch "verdict=ok") { throw ("Scenario3 expected ok, got: " + $res3) }
    Write-Host "  Scenario3 PASS"

    Close-ActiveWorkbook

    # ---------- Scenario 4: verify_conditional_format (rule mismatch) ----------
    Write-Host "== Scenario4: conditional format rule mismatch =="
    $wb4 = $excelApp.Workbooks.Add()
    $ws4 = $wb4.Worksheets.Item(1)
    $ws4.Name = "cf4"
    1..5 | ForEach-Object { $ws4.Cells.Item($_,1).Value2 = [int]($_*50) }
    $r4 = $ws4.Range("A1:A5")
    $fc4 = $r4.FormatConditions.Add(1, 5, "100")
    $fc4.Interior.Color = 255
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($fc4) | Out-Null
    $p4 = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $p4.Add("sheet", "cf4")
    $p4.Add("address", "A1:A5")
    $p4.Add("expected_rule", "less")   # wrong: actual is greater
    $p4.Add("expected_formula1", "100")
    $res4 = Invoke-Tool "VerifyConditionalFormat" $p4
    Write-Host ("  Scenario4 -> " + $res4)
    if ($res4 -notmatch "rule") { throw ("Scenario4 expected rule flag, got: " + $res4) }
    Write-Host "  Scenario4 PASS"

    Close-ActiveWorkbook

    # ---------- Scenario 5: verify_sort_filter (sort + filter match, in-memory semantics) ----------
    Write-Host "== Scenario5: sort + filter match =="
    $wb5 = $excelApp.Workbooks.Add()
    $ws5 = $wb5.Worksheets.Item(1)
    $ws5.Name = "sf"
    # final expected data after in-memory sort(col2 asc) + filter(col1==Beijing):
    # header, Beijing/20, Beijing/30  (hardcoded; avoids PS nested-array unroll trap)
    $ws5.Cells.Item(1,1).Value2 = "city"; $ws5.Cells.Item(1,2).Value2 = "val"
    $ws5.Cells.Item(2,1).Value2 = "Beijing"; $ws5.Cells.Item(2,2).Value2 = [int]20
    $ws5.Cells.Item(3,1).Value2 = "Beijing"; $ws5.Cells.Item(3,2).Value2 = [int]30
    $p5 = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $p5.Add("sheet", "sf")
    $p5.Add("address", "A1:B3")
    $p5.Add("header_rows", 1)
    $p5.Add("expected_sort_column", 2)
    $p5.Add("expected_sort_direction", "asc")
    $p5.Add("expected_filter_field", 1)
    $p5.Add("expected_filter_criteria1", "Beijing")
    $res5 = Invoke-Tool "VerifySortFilter" $p5
    Write-Host ("  Scenario5 -> " + $res5)
    if ($res5 -notmatch "verdict=ok") { throw ("Scenario5 expected ok, got: " + $res5) }
    Write-Host "  Scenario5 PASS"

    Close-ActiveWorkbook

    # ---------- Scenario 6: verify_sort_filter (wrong direction + filter not satisfied) ----------
    Write-Host "== Scenario6: sort/filter mismatch =="
    $wb6 = $excelApp.Workbooks.Add()
    $ws6 = $wb6.Worksheets.Item(1)
    $ws6.Name = "sf6"
    # final data after in-memory sort(col2 asc); NO filter applied:
    #   Shanghai/10, Beijing/20, Shanghai/30 (hardcoded)
    $ws6.Cells.Item(1,1).Value2 = "city"; $ws6.Cells.Item(1,2).Value2 = "val"
    $ws6.Cells.Item(2,1).Value2 = "Shanghai"; $ws6.Cells.Item(2,2).Value2 = [int]10
    $ws6.Cells.Item(3,1).Value2 = "Beijing"; $ws6.Cells.Item(3,2).Value2 = [int]20
    $ws6.Cells.Item(4,1).Value2 = "Shanghai"; $ws6.Cells.Item(4,2).Value2 = [int]30
    $p6 = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $p6.Add("sheet", "sf6")
    $p6.Add("address", "A1:B4")
    $p6.Add("header_rows", 1)
    $p6.Add("expected_sort_column", 2)
    $p6.Add("expected_sort_direction", "desc")   # wrong: actual is asc
    $p6.Add("expected_filter_field", 1)          # expected filter, but not applied
    $p6.Add("expected_filter_criteria1", "Beijing")
    $res6 = Invoke-Tool "VerifySortFilter" $p6
    Write-Host ("  Scenario6 -> " + $res6)
    if ($res6 -notmatch "sort") { throw ("Scenario6 expected sort flag, got: " + $res6) }
    if ($res6 -notmatch "filter") { throw ("Scenario6 expected filter flag, got: " + $res6) }
    Write-Host "  Scenario6 PASS"

    Close-ActiveWorkbook

    # ---------- Scenario 7: verify_formula (match) ----------
    Write-Host "== Scenario7: formula match =="
    $wb7 = $excelApp.Workbooks.Add()
    $ws7 = $wb7.Worksheets.Item(1)
    $ws7.Name = "fm"
    $ws7.Cells.Item(2,2).Value2 = [int]10
    $ws7.Cells.Item(2,3).Value2 = [int]12
    $ws7.Range("D2").Formula = "=B2*C2"
    $p7 = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $p7.Add("sheet", "fm")
    $p7.Add("address", "D2")
    $p7.Add("expected_formula", "=B2*C2")
    $p7.Add("expected_value", "120")
    $res7 = Invoke-Tool "VerifyFormula" $p7
    Write-Host ("  Scenario7 -> " + $res7)
    if ($res7 -notmatch "verdict=ok") { throw ("Scenario7 expected ok, got: " + $res7) }
    Write-Host "  Scenario7 PASS"

    Close-ActiveWorkbook

    # ---------- Scenario 8: verify_formula (error value + wrong formula) ----------
    Write-Host "== Scenario8: formula error + mismatch =="
    $wb8 = $excelApp.Workbooks.Add()
    $ws8 = $wb8.Worksheets.Item(1)
    $ws8.Name = "fm8"
    $ws8.Range("A1").Formula = "=1/0"   # #DIV/0!
    $p8 = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $p8.Add("sheet", "fm8")
    $p8.Add("address", "A1")
    $p8.Add("expected_formula", "=B2*C2")   # wrong
    $res8 = Invoke-Tool "VerifyFormula" $p8
    Write-Host ("  Scenario8 -> " + $res8)
    if ($res8 -notmatch "error") { throw ("Scenario8 expected error flag, got: " + $res8) }
    if ($res8 -notmatch "formula") { throw ("Scenario8 expected formula flag, got: " + $res8) }
    Write-Host "  Scenario8 PASS"

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