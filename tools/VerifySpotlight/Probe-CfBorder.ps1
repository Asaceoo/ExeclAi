$ErrorActionPreference = "Stop"
$passed = 0; $failed = 0
function Assert([bool]$cond, [string]$msg) {
    if ($cond) { $script:passed++; Write-Host "[PASS] $msg" }
    else { $script:failed++; Write-Host "[FAIL] $msg" }
}
$app = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$ws.Cells.FormatConditions.Delete()
try { $wb.Names.Item('JYYJ_VIEW_MODE').Delete() } catch { }
$wb.Names.Add('JYYJ_VIEW_MODE','=1') | Out-Null

# T9b: CF border WITHOUT Weight (only ColorIndex + LineStyle)
try {
    $fc = $ws.Range('A1:F6').FormatConditions.Add(2,0,'=AND(JYYJ_VIEW_MODE,ROW()=4,COLUMN()=4)')
    $fc.Borders.ColorIndex = 3
    $fc.Borders.LineStyle = 1
    $app.Calculate(); Start-Sleep -Milliseconds 400
    $db = $ws.Cells.Item(4,4).DisplayFormat.Borders.Item(7)
    Write-Host ("  D4 left border: ColorIndex=" + $db.ColorIndex + " LineStyle=" + $db.LineStyle + " Weight=" + $db.Weight)
    Assert ([int]$db.ColorIndex -eq 3) 'T9b left border ColorIndex=3 (red)'
    Assert ([int]$db.LineStyle -eq 1) 'T9b left border LineStyle=1 (continuous)'
    Assert ([int]$ws.Cells.Item(3,3).DisplayFormat.Borders.Item(7).LineStyle -ne 1) 'T9b C3 border remains none'
} catch { Assert $false "T9b outer: $($_.Exception.Message)" }
$ws.Cells.FormatConditions.Delete()

# T9c: per-edge border assignment on CF
try {
    $fc = $ws.Range('A1:F6').FormatConditions.Add(2,0,'=AND(JYYJ_VIEW_MODE,ROW()=5,COLUMN()=5)')
    $fc.Borders.Item(7).LineStyle = 1
    $fc.Borders.Item(7).ColorIndex = 3
    $fc.Borders.Item(8).LineStyle = 1
    $fc.Borders.Item(8).ColorIndex = 3
    $app.Calculate(); Start-Sleep -Milliseconds 400
    Assert ([int]$ws.Cells.Item(5,5).DisplayFormat.Borders.Item(7).ColorIndex -eq 3) 'T9c E5 left border red (per-edge)'
    Assert ([int]$ws.Cells.Item(5,5).DisplayFormat.Borders.Item(8).ColorIndex -eq 3) 'T9c E5 top border red (per-edge)'
    Assert ([int]$ws.Cells.Item(5,5).DisplayFormat.Borders.Item(9).LineStyle -ne 1) 'T9c E5 bottom border none'
} catch { Assert $false "T9c outer: $($_.Exception.Message)" }
$ws.Cells.FormatConditions.Delete()

try { $wb.Names.Item('JYYJ_VIEW_MODE').Delete() } catch { }
$wb.Close($false)

Write-Host ""
Write-Host "== CF-BORDER PASSED=$passed FAILED=$failed =="
if ($failed -gt 0) { Write-Host 'CF-BORDER FAILED'; exit 1 } else { Write-Host 'CF-BORDER OK'; exit 0 }
