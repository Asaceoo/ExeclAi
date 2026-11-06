$app = $null
for ($i = 0; $i -lt 20; $i++) {
    try { $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application"); break } catch { Start-Sleep -Seconds 1 }
}
if ($app) { Write-Host "OK Workbooks=$($app.Workbooks.Count)" } else { Write-Host "NO ROT" }
