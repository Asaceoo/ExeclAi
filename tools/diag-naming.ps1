$ErrorActionPreference = 'Stop'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1
$wb = $excel.Workbooks.Add(); $ws=$wb.Worksheets.Item(1); $ws.Name="dt"
$mod=$wb.VBProject.VBComponents.Add(1); $mod.Name="M"
$mod.CodeModule.AddFromString("Public Sub RunMe()`n    Range(`"A1`").Value = 42`nEnd Sub`nPublic Function GetIt()`n    GetIt = 7`nEnd Function")
function T($l,$expr){ try { $v = & $expr; return "$l => OK value=$v" } catch { return "$l => FAIL: $($_.Exception.Message.Split([char]10)[0])" } }
Write-Host (T "RunMe" { $excel.Application.Run("RunMe") })
Write-Host (T "M.RunMe" { $excel.Application.Run("M.RunMe") })
Write-Host (T "qualified wb macro" { $excel.Run("'$($wb.Name)'!M.RunMe") })
$wb.Close($false); $excel.Quit(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null