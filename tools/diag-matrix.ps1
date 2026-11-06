$ErrorActionPreference = 'Stop'
function Try-Run($xl, $label) {
    try { $xl.Application.Run("HelloTest"); return "RUN-OK" }
    catch { return ("RUN-FAIL: " + $_.Exception.Message.Replace("`r"," ").Replace("`n"," ")) }
}
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1
Write-Host "AutomationSecurity(readback)=$($excel.AutomationSecurity)"
$modSrc = "Sub HelloTest()`n    Range(`"A1`").Value = 42`nEnd Sub"

# S1: unsaved in-memory
$wb = $excel.Workbooks.Add(); $ws=$wb.Worksheets.Item(1); $ws.Name="dt"
$mod=$wb.VBProject.VBComponents.Add(1); $mod.Name="HelloTest"; $mod.CodeModule.AddFromString($modSrc)
Write-Host ("S1 unsaved in-memory : " + (Try-Run $excel "s1"))

# S2: save as xlsm (keep open, no close), AutomationSecurity low, re-run
$xlsm = Join-Path $env:TEMP ("m_" + [guid]::NewGuid().ToString('N') + ".xlsm")
$wb.SaveAs($xlsm, 52)
$excel.AutomationSecurity = 1
Write-Host ("S2 saved-xlsm open   : " + (Try-Run $excel "s2"))

# S3: close, reopen via automation with low security, run
$name=$ws.Name; $wb.Close($false)
$wb2=$excel.Workbooks.Open($xlsm); $excel.AutomationSecurity=1
$wb2.Activate()
Write-Host ("S3 reopen-xlsm low   : " + (Try-Run $excel "s3"))

# S4: check macros enabled property on reopened project + trusted
try { $wb2.TrustProjectVBWindow | Out-Null } catch {}
Write-Host "S4 activeworkbook fullname=$($wb2.FullName)"
$wb2.Close($false)
$excel.Quit()
[Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null