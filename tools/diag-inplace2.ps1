$ErrorActionPreference = 'Stop'
$log = Join-Path $env:TEMP "jyyj_inplace2.log"
Set-Content -Path $log -Value ("START "+ (Get-Date -Format 'HH:mm:ss')) -Encoding utf8
function W($m){ Add-Content -Path $log -Value ("{0}  {1}" -f (Get-Date -Format 'HH:mm:ss'), $m) -Encoding utf8 }
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1
function RunWrap($n){ try { $excel.Application.Run($n); return "OK" } catch { "FAIL: " + $_.Exception.Message.Split([char]10)[0] } }
function AddMod($wb){ $m=$wb.VBProject.VBComponents.Add(1); $m.Name="M"; $m.CodeModule.AddFromString("Public Sub RunMe()`n    Range(`"A1`").Value = 42`nEnd Sub") }
try {
    # Make a clean on-disk xlsx (no macros)
    $seed=$excel.Workbooks.Add(); $seed.Worksheets.Item(1).Range("A1").Value2=0
    $x=$env:TEMP+"\seed_"+[guid]::NewGuid().ToString('N')+".xlsx"; $seed.SaveAs($x,51)
    $seedRO=$seed; $seedRO.Close($false)

    # Reopen (disk-loaded xlsx), inject macro, run
    $wb=$excel.Workbooks.Open($x); $excel.AutomationSecurity=1
    AddMod $wb
    W ("G1 disk-xlsx inject & run : " + (RunWrap "M.RunMe"))
    # SaveAs xlsm in place (no close/reopen), run again
    $xlsm=$x -replace '\.xlsx$','.xlsm'
    $wb.SaveAs($xlsm,52)
    W ("G2 after SaveAs xlsm      : " + (RunWrap "M.RunMe"))
    $wb.Close($false)
} catch { W ("ERR " + $_.Exception.Message) }
finally { try{$excel.Quit()}catch{}; try{[Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)|Out-Null}catch{} }
Add-Content -Path $log -Value "DONE" -Encoding utf8