$ErrorActionPreference = "Stop"
$app = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$ws.Cells.FormatConditions.Delete()
try { $wb.Names.Item('JYYJ_VIEW_MODE').Delete() } catch { }
$wb.Names.Add('JYYJ_VIEW_MODE','=1') | Out-Null
$fc = $ws.Range('A1:F6').FormatConditions.Add(2,0,'=AND(JYYJ_VIEW_MODE,ROW()=4,COLUMN()=4)')
$app.Calculate(); Start-Sleep -Milliseconds 400
Write-Host ("Before name delete, Formula1: " + [string]$ws.Cells.FormatConditions.Item(1).Formula1)

# 删除哨兵定义名，再看规则 Formula1 是否被改写
$wb.Names.Item('JYYJ_VIEW_MODE').Delete()
Start-Sleep -Milliseconds 400
Write-Host ("After name delete,  Formula1: " + [string]$ws.Cells.FormatConditions.Item(1).Formula1)
Write-Host ("  contains JYYJ_VIEW_MODE: " + ([string]$ws.Cells.FormatConditions.Item(1).Formula1).Contains('JYYJ_VIEW_MODE'))

# 尝试按哨兵 token 删除
$del = 0
for ($i = 1; $i -le $ws.Cells.FormatConditions.Count; $i++) {
    if (([string]$ws.Cells.FormatConditions.Item($i).Formula1).Contains('JYYJ_VIEW_MODE')) { $del++ }
}
Write-Host ("Rules matching token: " + $del + " / total=" + $ws.Cells.FormatConditions.Count)

$ws.Cells.FormatConditions.Delete()
try { $wb.Names.Item('JYYJ_VIEW_MODE').Delete() } catch { }
$wb.Close($false)
