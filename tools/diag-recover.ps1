$log = Join-Path $env:TEMP "jyyj_recover.log"
Set-Content -Path $log -Value ("START "+ (Get-Date -Format 'HH:mm:ss')) -Encoding utf8
function W($m){ Add-Content -Path $log -Value ("{0}  {1}" -f (Get-Date -Format 'HH:mm:ss'), $m) -Encoding utf8 }
$ErrorActionPreference='Stop'
try {
  $excel = New-Object -ComObject Excel.Application; W "created"
  $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.AutomationSecurity=1; W "cfg ok"
  function RunW($n){ try { $excel.Application.Run($n); return "OK" } catch { "FAIL:" + $_.Exception.Message.Split([char]10)[0] } }
  $wb=$excel.Workbooks.Add(); $m=$wb.VBProject.VBComponents.Add(1); $m.Name="M"; $m.CodeModule.AddFromString("Public Sub RunMe()`n Range(`"A1`").Value=42`nEnd Sub")
  W "module injected"
  $x=$env:TEMP+"\r_"+[guid]::NewGuid().ToString('N')+".xlsx"; $wb.SaveAs($x,51); W "saved xlsx (has macro)"
  $wb.Close($false); W "closed"
  $wb=$excel.Workbooks.Open($x); $excel.AutomationSecurity=1; W "reopened"
  W ("A reopen saved-with-macro xlsx => " + (RunW "M.RunMe"))
  $wb.SaveAs(($x -replace '\.xlsx$','.xlsm'),52); W "saveas xlsm in place"
  W ("B after SaveAs xlsm in place  => " + (RunW "M.RunMe"))
  $wb.Close($false); $excel.Quit()
  try{[Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)|Out-Null}catch{}
  W "DONE"
} catch { W ("ERR: "+$_.Exception.Message); try{$excel.Quit()}catch{} }
Add-Content -Path $log -Value "FIN" -Encoding utf8