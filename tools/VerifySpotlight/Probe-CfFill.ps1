$ErrorActionPreference = "Stop"
$passed = 0; $failed = 0
function Assert([bool]$cond, [string]$msg) {
    if ($cond) { $script:passed++; Write-Host "[PASS] $msg" }
    else { $script:failed++; Write-Host "[FAIL] $msg" }
}
$app = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
for ($r = 1; $r -le 6; $r++) { for ($c = 1; $c -le 6; $c++) { $ws.Cells.Item($r, $c).Value2 = $r * 10 + $c } }
$ws.Cells.FormatConditions.Delete()
try { $wb.Names.Item('JYYJ_VIEW_MODE').Delete() } catch { }
$wb.Names.Add('JYYJ_VIEW_MODE','=1') | Out-Null

# Case A: ColorIndex + Pattern (exactly like add-in), NO Calculate, wait 2s
try {
    $fc = $ws.Range('A1:F6').FormatConditions.Add(2,0,'=AND(JYYJ_VIEW_MODE,OR(ROW()=2,COLUMN()=2))')
    $fc.Interior.ColorIndex = 6
    $fc.Interior.Pattern = 1
    Start-Sleep -Seconds 2
    Assert ([int]$ws.Cells.Item(2,1).DisplayFormat.Interior.ColorIndex -eq 6) 'A: ColorIndex+Pattern, no calc, A2 yellow'
    Assert ([int]$ws.Cells.Item(1,1).DisplayFormat.Interior.ColorIndex -ne 6) 'A: A1 not yellow'
} catch { Assert $false "A outer: $($_.Exception.Message)" }
$ws.Cells.FormatConditions.Delete()
Start-Sleep -Milliseconds 300

# Case B: ColorIndex + Pattern, WITH Calculate
try {
    $fc = $ws.Range('A1:F6').FormatConditions.Add(2,0,'=AND(JYYJ_VIEW_MODE,OR(ROW()=2,COLUMN()=2))')
    $fc.Interior.ColorIndex = 6
    $fc.Interior.Pattern = 1
    $app.Calculate()
    Start-Sleep -Seconds 1
    Assert ([int]$ws.Cells.Item(2,1).DisplayFormat.Interior.ColorIndex -eq 6) 'B: ColorIndex+Pattern, calc, A2 yellow'
} catch { Assert $false "B outer: $($_.Exception.Message)" }
$ws.Cells.FormatConditions.Delete()
Start-Sleep -Milliseconds 300

# Case C: ColorIndex only (no Pattern), no Calculate
try {
    $fc = $ws.Range('A1:F6').FormatConditions.Add(2,0,'=AND(JYYJ_VIEW_MODE,OR(ROW()=2,COLUMN()=2))')
    $fc.Interior.ColorIndex = 6
    Start-Sleep -Seconds 2
    Assert ([int]$ws.Cells.Item(2,1).DisplayFormat.Interior.ColorIndex -eq 6) 'C: ColorIndex only, no calc, A2 yellow'
} catch { Assert $false "C outer: $($_.Exception.Message)" }
$ws.Cells.FormatConditions.Delete()

try { $wb.Names.Item('JYYJ_VIEW_MODE').Delete() } catch { }
$wb.Close($false)
Write-Host ""
Write-Host "== CF-FILL PASSED=$passed FAILED=$failed =="
if ($failed -gt 0) { Write-Host 'CF-FILL FAILED'; exit 1 } else { Write-Host 'CF-FILL OK'; exit 0 }
