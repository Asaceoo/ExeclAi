$ErrorActionPreference = 'Stop'
$log = Join-Path $env:TEMP "jyyj_addinrepro.log"
Set-Content -Path $log -Value ("START "+ (Get-Date -Format 'HH:mm:ss')) -Encoding utf8
function W($m){ Add-Content -Path $log -Value ("{0}  {1}" -f (Get-Date -Format 'HH:mm:ss'), $m) -Encoding utf8 }
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1
$M=@'
Option Explicit
Sub AutoFitRowHeight()
    Range("A1").AutoFit
End Sub
'@
$M2=@'
Option Explicit
Sub AutoFitRows()
    Rows.AutoFit
End Sub
'@
function RunWrap($n){ try { $excel.Application.Run($n); return "OK" } catch { "FAIL: " + $_.Exception.Message.Split([char]10)[0] } }
function GetModule($comps){ try { $m=$comps.Item('JYYJHelper') } catch { $m=$null }; if(-not $m){ $m=$comps.Add(1); try{$m.Name='JYYJHelper'}catch{} }; return $m }
function ReplaceMod($cm,$src){ try { $c=[int]$cm.CountOfLines; if($c -gt 0){$cm.DeleteLines(1,$c)} } catch {}; $cm.AddFromString($src) }
try {
    # clean on-disk xlsx
    $seed=$excel.Workbooks.Add(); $seed.Worksheets.Item(1).Range("A1").Value2=0
    $x=$env:TEMP+"\a_"+[guid]::NewGuid().ToString('N')+".xlsx"; $seed.SaveAs($x,51); $seed.Close($false)
    # reopen, inject into reused JYYJHelper
    $wb=$excel.Workbooks.Open($x); $excel.AutomationSecurity=1
    $mod=GetModule $wb.VBProject.VBComponents
    ReplaceMod $mod.CodeModule $M
    W ("H1 run1 AutoFitRowHeight: " + (RunWrap "JYYJHelper.AutoFitRowHeight"))
    ReplaceMod $mod.CodeModule $M2
    W ("H2 run2 AutoFitRows (reuse): " + (RunWrap "JYYJHelper.AutoFitRows"))
    # save the workbook WITH the JYYJHelper module as xlsx, close, reopen, try reuse run
    $wb.SaveAs($x,51)   # xlsx cannot host macros
    $wb.Close($false)
    $wb=$excel.Workbooks.Open($x); $excel.AutomationSecurity=1
    $mod2=GetModule $wb.VBProject.VBComponents
    W ("H3 after reopened-saved-xlsx, module present=" + (try{ ([bool]($mod2 -ne $null)) }catch{ '?' }))
    ReplaceMod $mod2.CodeModule $M
    W ("H4 reuse run after xlsx-save reopen: " + (RunWrap "JYYJHelper.AutoFitRowHeight"))
    $wb.Close($false)
} catch { W ("ERR " + $_.Exception.Message) }
finally { try{$excel.Quit()}catch{}; try{[Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)|Out-Null}catch{} }
Add-Content -Path $log -Value "DONE" -Encoding utf8