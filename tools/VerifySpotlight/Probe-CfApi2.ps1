$ErrorActionPreference = "Stop"
$passed = 0; $failed = 0
function Assert([bool]$cond, [string]$msg) {
    if ($cond) { $script:passed++; Write-Host "[PASS] $msg" }
    else { $script:failed++; Write-Host "[FAIL] $msg" }
}

$app = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
if (-not $app) { Write-Host 'NO EXCEL'; exit 2 }
$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
for ($r = 1; $r -le 6; $r++) { for ($c = 1; $c -le 6; $c++) { $ws.Cells.Item($r, $c).Value2 = $r * 10 + $c } }
$ws.Cells.FormatConditions.Delete()

# --- T7: sentinel defined name marker ---
try {
    $ws.Cells.FormatConditions.Delete()
    $wb.Names.Item('JYYJ_VIEW_MODE').Delete() | Out-Null
} catch { }
try {
    $wb.Names.Add('JYYJ_VIEW_MODE', '=1') | Out-Null
    $fc = $ws.Range('A1:C6').FormatConditions.Add(2,0,'=AND(JYYJ_VIEW_MODE,OR(ROW()=2,COLUMN()=2))')
    $fc.Interior.ColorIndex = 6
    $app.Calculate(); Start-Sleep -Milliseconds 400
    $f1 = [string]$ws.Cells.FormatConditions.Item(1).Formula1
    Write-Host ("  Formula1 stored as: " + $f1)
    Assert ($f1 -like '*JYYJ_VIEW_MODE*') 'T7 sentinel token readable in Formula1'
    Assert ([int]$ws.Cells.Item(2,1).DisplayFormat.Interior.ColorIndex -eq 6) 'T7 sentinel formula highlights row2'
    Assert ([int]$ws.Cells.Item(1,1).DisplayFormat.Interior.ColorIndex -ne 6) 'T7 A1 not highlighted'
} catch { Assert $false "T7 outer: $($_.Exception.Message)" }

# --- T8: find-and-delete via sentinel token ---
try {
    $found = $false
    for ($i = 1; $i -le $ws.Cells.FormatConditions.Count; $i++) {
        $f = [string]$ws.Cells.FormatConditions.Item($i).Formula1
        if ($f -like '*JYYJ_VIEW_MODE*') { $found = $true; $ws.Cells.FormatConditions.Item($i).Delete(); break }
    }
    Assert $found 'T8 found rule by sentinel token'
    Assert ($ws.Cells.FormatConditions.Count -eq 0) 'T8 rule deleted'
} catch { Assert $false "T8 outer: $($_.Exception.Message)" }

# --- T9: CF border rendering (pointer) ---
try {
    $fc = $ws.Range('A1:F6').FormatConditions.Add(2,0,'=AND(JYYJ_VIEW_MODE,ROW()=4,COLUMN()=4)')
    $fc.Borders.ColorIndex = 3
    $fc.Borders.LineStyle = 1
    $fc.Borders.Weight = 4
    $app.Calculate(); Start-Sleep -Milliseconds 400
    Assert ([int]$ws.Cells.Item(4,4).DisplayFormat.Borders.Item(7).ColorIndex -eq 3) 'T9 D4 left border red'
    Assert ([int]$ws.Cells.Item(4,4).DisplayFormat.Borders.Item(8).ColorIndex -eq 3) 'T9 D4 top border red'
    Assert ([int]$ws.Cells.Item(4,4).DisplayFormat.Borders.Item(7).LineStyle -eq 1) 'T9 D4 left border continuous'
    Assert ([int]$ws.Cells.Item(3,3).DisplayFormat.Borders.Item(7).ColorIndex -ne 3) 'T9 C3 not bordered'
} catch { Assert $false "T9 outer: $($_.Exception.Message)" }
$ws.Cells.FormatConditions.Delete()

# --- T10: delete sentinel name ---
try {
    $wb.Names.Item('JYYJ_VIEW_MODE').Delete() | Out-Null
    Assert $true 'T10 sentinel name deletable'
} catch { Assert $false "T10 delete name: $($_.Exception.Message)" }

$wb.Close($false)
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws) | Out-Null
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null

Write-Host ""
Write-Host "== CF-API2 PASSED=$passed FAILED=$failed =="
if ($failed -gt 0) { Write-Host 'CF-API2 CHECKS FAILED'; exit 1 } else { Write-Host 'ALL CF-API2 CHECKS OK'; exit 0 }
