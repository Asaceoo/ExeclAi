# 探针 5：逐步诊断整表单规则为何在移动/重算后消失。
$ErrorActionPreference = "Stop"
$app = $null
for ($i = 0; $i -lt 40; $i++) {
    try { $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application"); break } catch { Start-Sleep -Seconds 1 }
}
if ($null -eq $app) { Write-Host "NO EXCEL INSTANCE"; exit 2 }

$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
for ($r = 1; $r -le 6; $r++) { for ($c = 1; $c -le 6; $c++) { $ws.Cells.Item($r, $c).Value2 = $r * 10 + $c } }
$wb.Names.Add('JYYJ_VIEW_MODE', '=1') | Out-Null

function DumpRules([string]$tag) {
    try {
        $coll = $ws.Cells.FormatConditions
        $cnt = [int]$coll.Count
        Write-Host ("  [$tag] Cells.FC.Count=" + $cnt)
        for ($i = 1; $i -le $cnt; $i++) {
            try {
                $f = $coll.Item($i)
                $t = [int]$f.Type
                $fml = [string]$f.Formula1
                $at = ""
                try { $at = [string]$f.AppliesTo.Address } catch { $at = "AppliesTo?" }
                Write-Host ("    item[" + $i + "] type=" + $t + " fml=[" + $fml + "] applies=[" + $at + "]")
            } catch { Write-Host ("    item[" + $i + "] READ FAIL: " + $_.Exception.GetType().Name + " " + $_.Exception.Message) }
        }
    } catch { Write-Host ("  [$tag] COLLECTION FAIL: " + $_.Exception.GetType().Name + " " + $_.Exception.Message) }
}

Write-Host "== 建整表单规则 =="
$fc = $ws.Cells.FormatConditions.Add(2, 0, '=AND(JYYJ_VIEW_MODE,OR(CELL("row")=ROW(),CELL("col")=COLUMN()))')
$fc.Interior.ColorIndex = 6
$fc.Interior.Pattern = 1
DumpRules "建后"

$app.Range("B2").Select()
$ws.Calculate()
Start-Sleep -Milliseconds 500
Write-Host ("B2 高亮 A2=" + [int]$ws.Cells.Item(2,1).DisplayFormat.Interior.ColorIndex)
DumpRules "B2重算后"

$app.Range("D4").Select()
Start-Sleep -Milliseconds 500
DumpRules "D4选中后(未重算)"

$ws.Calculate()
Start-Sleep -Milliseconds 500
Write-Host ("D4 高亮 A4=" + [int]$ws.Cells.Item(4,1).DisplayFormat.Interior.ColorIndex)
DumpRules "D4重算后"

$wb.Close($false)
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws) | Out-Null
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null
