$ErrorActionPreference = 'Stop'
$log = Join-Path $env:TEMP "jyyj_formatDiag.log"
function W($m){ $line = "{0}  {1}" -f (Get-Date -Format 'HH:mm:ss'), $m; Add-Content -Path $log -Value $line -Encoding utf8 }
Set-Content -Path $log -Value ("START "+ (Get-Date -Format 'HH:mm:ss')) -Encoding utf8
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1
W "excel created sec=$($excel.AutomationSecurity)"
function RunWrap($name){ try { $excel.Application.Run($name); return "OK" } catch { return "FAIL: " + $_.Exception.Message.Split([char]10)[0] } }
function AddMod($wb){ $m=$wb.VBProject.VBComponents.Add(1); $m.Name="M"; $m.CodeModule.AddFromString("Public Sub RunMe()`n    Range(`"A1`").Value = 42`nEnd Sub"); return $m }
try {
    # Control: unsaved Workbook.Add
    $wb=$excel.Workbooks.Add(); $m=AddMod $wb
    W ("CONTROL unsaved        : " + (RunWrap "M.RunMe"))
    $wb.Close($false)

    # Saved .xlsx
    $wb=$excel.Workbooks.Add(); $m=AddMod $wb
    $x=$env:TEMP+"\j_"+[guid]::NewGuid().ToString('N')+".xlsx"; $wb.SaveAs($x,51); $p=$x; $in=$wb; $in.Close($false)
    $wb=$excel.Workbooks.Open($p); $excel.AutomationSecurity=1
    W ("SAVED xlsx (reopen)    : " + (RunWrap "M.RunMe"))
    $wb.Close($false)

    # Saved .xlsm
    $wb=$excel.Workbooks.Add(); $m=AddMod $wb
    $y=$env:TEMP+"\j_"+[guid]::NewGuid().ToString('N')+".xlsm"; $wb.SaveAs($y,52); $in=$wb; $in.Close($false)
    $wb=$excel.Workbooks.Open($y); $excel.AutomationSecurity=1
    W ("SAVED xlsm (reopen)    : " + (RunWrap "M.RunMe"))
    $wb.Close($false)
} catch { W ("ERROR: " + $_.Exception.Message) }
finally { try{$excel.Quit()}catch{}; try{[Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)|Out-Null}catch{} }
Add-Content -Path $log -Value "DONE" -Encoding utf8