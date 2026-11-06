$ErrorActionPreference = "Stop"
$app = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
Write-Host ("Version=" + $app.Version)
$wb = $app.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
$fc = $ws.Range('A1:C6').FormatConditions.Add(2,0,'=OR(ROW()=2,COLUMN()=2)')
Write-Host '--- FormatCondition members ---'
$fc | Get-Member | Where-Object { $_.MemberType -match 'Property|Method' } | ForEach-Object { $_.MemberType.ToString().PadRight(14) + $_.Name }
$wb.Close($false)
