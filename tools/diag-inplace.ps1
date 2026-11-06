$ErrorActionPreference = 'Stop'
$log = Join-Path $env:TEMP "jyyj_inplace.log"
Set-Content -Path $log -Value ("START "+ (Get-Date -Format 'HH:mm:ss')) -Encoding utf8
function W($m){ Add-Content -Path $log -Value ("{0}  {1}" -f (Get-Date -Format 'HH:mm:ss'), $m) -Encoding utf8 }
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1
function RunWrap($n){ try { $excel.Application.Run($n); return "OK" } catch { "FAIL: " + $_.Exception.Message.Split([char]10)[0] } }
function AddMod($wb){ $m=$wb.VBProject.VBComponents.Add(1); $m.Name="M"; $m.CodeModule.AddFromString("Public Sub RunMe()`n    Range(`"A1`").Value = 42`nEnd Sub") }
try {
    # F: add book, inject, SaveAs xlsx in place -> run; then SaveAs xlsm in place -> run
    $wb=$excel.Workbooks.Add(); $ws=$wb.Worksheets.Item(1); $ws.Name="dt"
    AddMod $wb
    $x=$env:TEMP+"\g_"+[guid]::NewGuid().ToString('N')+".xlsx"; $xlsm=$x -replace '\.xlsx$','.xlsm'
    $wb.SaveAs($x,51)
    W ("F1 xlsx in-place    : " + (RunWrap "M.RunMe"))
    $excel.Quit()
} catch { W ("ERR " + $_.Exception.Message) }
try { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)|Out-Null } catch {}
Add-Content -Path $log -Value "DONE" -Encoding utf8