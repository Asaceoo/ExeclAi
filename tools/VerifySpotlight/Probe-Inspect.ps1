# Diagnose: inspect all open workbooks for JYYJ names/rules and active sheet.
$ErrorActionPreference = "Stop"
$app = $null
for ($i = 0; $i -lt 40; $i++) {
    try { $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application"); break } catch { Start-Sleep -Seconds 1 }
}
if ($null -eq $app) { Write-Host "NO EXCEL"; exit 2 }
Write-Host ("Workbooks.Count=" + $app.Workbooks.Count)
for ($w = 1; $w -le $app.Workbooks.Count; $w++) {
    $wb = $app.Workbooks.Item($w)
    Write-Host ("  WB[$w] Name=" + $wb.Name + " FullName=" + $wb.FullName + " ActiveSheet=" + $wb.ActiveSheet.Name)
    $names = @()
    for ($i = 1; $i -le $wb.Names.Count; $i++) {
        try { $n = $wb.Names.Item($i); $names += ([string]$n.Name + "=" + [string]$n.RefersTo) } catch { }
    }
    Write-Host ("    Names: " + ($names -join "; "))
    for ($s = 1; $s -le $wb.Worksheets.Count; $s++) {
        $ws = $wb.Worksheets.Item($s)
        $rules = @()
        try {
            for ($i = 1; $i -le $ws.Cells.FormatConditions.Count; $i++) {
                try { $f = [string]$ws.Cells.FormatConditions.Item($i).Formula1; $rules += ("r$i=" + $f) } catch { $rules += "r$i=<nof1>" }
            }
        } catch { $rules += "<noFC>" }
        Write-Host ("    Sheet[" + $s + "]=" + $ws.Name + " FC: " + ($rules -join "; "))
    }
}
Write-Host "  ActiveWorkbook=" + $app.ActiveWorkbook.Name + " ActiveSheet=" + $app.ActiveSheet.Name
Write-Host "DONE"
