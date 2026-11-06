# Probe CF key APIs for Spotlight/POinter CF-based rewrite (all-ASCII to avoid BOM issues)
$ErrorActionPreference = "Stop"
$passed = 0; $failed = 0
function Assert([bool]$cond, [string]$msg) {
    if ($cond) { $script:passed++; Write-Host "[PASS] $msg" }
    else { $script:failed++; Write-Host "[FAIL] $msg" }
}

$app = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
if (-not $app) { Write-Host "NO EXCEL"; exit 2 }

$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
for ($r = 1; $r -le 6; $r++) { for ($c = 1; $c -le 6; $c++) { $ws.Cells.Item($r, $c).Value2 = $r * 10 + $c } }
$ws.Cells.FormatConditions.Delete()
Start-Sleep -Milliseconds 200

# --- T1: can we set Description (feature marker)? ---
try {
    $fc = $ws.Range("A1:C6").FormatConditions.Add(2, 0, '=OR(ROW()=2,COLUMN()=2)')
    try {
        $fc.Description = "JYYJ_SPOTLIGHT"
        Assert ($fc.Description -eq "JYYJ_SPOTLIGHT") "T1 Description marker set+read"
    } catch { Assert $false "T1 Description settable: $($_.Exception.Message)" }
    $fc.Interior.ColorIndex = 6
    $fc.Interior.Pattern = 1
    $app.Calculate()
    Start-Sleep -Milliseconds 500
    Assert ($ws.Cells.FormatConditions.Count -eq 1) "T1 one CF rule added"
    Assert ([int]$ws.Cells.Item(2,1).DisplayFormat.Interior.ColorIndex -eq 6) "T1 DisplayFormat A2 yellow(6)"
    Assert ([int]$ws.Cells.Item(1,2).DisplayFormat.Interior.ColorIndex -eq 6) "T1 DisplayFormat B1 yellow(6)"
    Assert ([int]$ws.Cells.Item(1,1).DisplayFormat.Interior.ColorIndex -ne 6) "T1 DisplayFormat A1 not yellow"
    Assert ([int]$ws.Cells.Item(2,1).Interior.ColorIndex -eq -4142) "T1 actual Interior untouched(-4142)"
} catch { Assert $false "T1 outer: $($_.Exception.Message)" }
$ws.Cells.FormatConditions.Delete()
Start-Sleep -Milliseconds 200

# --- T2: find rule by Description marker, then delete ---
try {
    $fc = $ws.Range("A1:C6").FormatConditions.Add(2, 0, '=OR(ROW()=2,COLUMN()=2)')
    $fc.Description = "JYYJ_SPOTLIGHT"
    $fc.Interior.ColorIndex = 6
    $found = $false
    for ($i = 1; $i -le $ws.Cells.FormatConditions.Count; $i++) {
        $d = $ws.Cells.FormatConditions.Item($i).Description
        if ($d -eq "JYYJ_SPOTLIGHT") { $found = $true; $ws.Cells.FormatConditions.Item($i).Delete(); break }
    }
    Assert $found "T2 locate rule by Description"
    Assert ($ws.Cells.FormatConditions.Count -eq 0) "T2 rule deleted"
} catch { Assert $false "T2 outer: $($_.Exception.Message)" }

# --- T3: CF border (pointer) rendering ---
try {
    $fc = $ws.Range("A1:C6").FormatConditions.Add(2, 0, '=AND(ROW()=4,COLUMN()=4)')
    $fc.Description = "JYYJ_POINTER"
    $fc.Borders.ColorIndex = 3
    $fc.Borders.LineStyle = 1
    $fc.Borders.Weight = 4
    $app.Calculate()
    Start-Sleep -Milliseconds 500
    Assert ([int]$ws.Cells.Item(4,4).DisplayFormat.Borders.Item(7).ColorIndex -eq 3) "T3 D4 left border red via DisplayFormat"
    Assert ([int]$ws.Cells.Item(4,4).DisplayFormat.Borders.Item(8).ColorIndex -eq 3) "T3 D4 top border red via DisplayFormat"
    Assert ([int]$ws.Cells.Item(3,3).DisplayFormat.Borders.Item(7).ColorIndex -ne 3) "T3 C3 border not red"
} catch { Assert $false "T3 outer: $($_.Exception.Message)" }
$ws.Cells.FormatConditions.Delete()
Start-Sleep -Milliseconds 200

# --- T4: non-A1 anchored range, ROW()/COLUMN() relative behavior ---
try {
    # anchor at C3 (row3,col3). Formula relative to top-left C3:
    # ROW() at C3 itself should be 3 (absolute row), so =ROW()=5 formats row 5 absolute
    $fc = $ws.Range("C3:F6").FormatConditions.Add(2, 0, '=OR(ROW()=5,COLUMN()=3)')
    $fc.Description = "JYYJ_SPOTLIGHT"
    $fc.Interior.ColorIndex = 6
    $app.Calculate()
    Start-Sleep -Milliseconds 500
    Assert ([int]$ws.Cells.Item(5,3).DisplayFormat.Interior.ColorIndex -eq 6) "T4 row5 col3 yellow (ROW()=5 absolute)"
    Assert ([int]$ws.Cells.Item(3,3).DisplayFormat.Interior.ColorIndex -eq 6) "T4 col3 yellow (COLUMN()=3 absolute)"
    Assert ([int]$ws.Cells.Item(4,4).DisplayFormat.Interior.ColorIndex -ne 6) "T4 C4 not yellow"
} catch { Assert $false "T4 outer: $($_.Exception.Message)" }
$ws.Cells.FormatConditions.Delete()

# --- T5: deleting non-existent collection is safe ---
try {
    $ws.Cells.FormatConditions.Delete()
    Assert $true "T5 Delete on empty FormatConditions is safe"
} catch { Assert $false "T5 Delete empty throws: $($_.Exception.Message)" }

# --- T6: FormatConditions on Chart sheet / protected sheet not touched here; skip ---

$wb.Close($false)
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws) | Out-Null
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null

Write-Host ""
Write-Host "== CF-API PASSED=$passed FAILED=$failed =="
if ($failed -gt 0) { Write-Host "CF-API CHECKS FAILED"; exit 1 } else { Write-Host "ALL CF-API CHECKS OK"; exit 0 }
