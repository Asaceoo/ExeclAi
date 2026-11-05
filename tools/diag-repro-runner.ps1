$ErrorActionPreference = 'Stop'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1
$const = 'JYYJHelper'
function Norm($snippet){ if($snippet -match '^\s*(?:(?:Public|Private|Friend)\s+)?(Sub|Function)\s+\w+'){ return "Option Explicit`n"+$snippet } else { return "Option Explicit`nPublic Sub RunSnippet()`n"+$snippet+"`nEnd Sub" } }
function GetModule($components){ try { $m = $components.Item($const) } catch { $m = $null }; if(-not $m){ $m = $components.Add(1); try{$m.Name=$const}catch{} }; return $m }
function ReplaceMod($cm,$snippet){ try { $c=[int]$cm.CountOfLines; if($c -gt 0){ $cm.DeleteLines(1,$c) } } catch {}; $cm.AddFromString($snippet) }
function RunWrap($xl,$name){ try { $xl.Application.Run($name); return "OK" } catch { return ("FAIL: " + ($_.Exception.Message.Split([char]10)[0])) } }

$wb = $excel.Workbooks.Add(); $ws=$wb.Worksheets.Item(1); $ws.Name="dt"
$vbp = $wb.VBProject
$s1 = "Sub AutoFitRowHeight()`n    Range(`"A1`").AutoFit`nEnd Sub"
$s2 = "Sub AutoFitRows()`n    Rows.AutoFit`nEnd Sub"

$mod = GetModule $vbp.VBComponents
Write-Host "module=$($mod.Name)"
Write-Host ("eligible count check: " + ([int]$mod.CodeModule.CountOfLines))

# Run 1 with macro_name supplied (like add-in may)
ReplaceMod $mod.CodeModule (Norm $s1)
Write-Host ("run1 AutoFitRowHeight: " + (RunWrap $excel "JYYJHelper.AutoFitRowHeight"))

# Run 2 reuse same module (first-run delete all lines)
ReplaceMod $mod.CodeModule (Norm $s2)
Write-Host ("module lines now=" + ([int]$mod.CodeModule.CountOfLines))
Write-Host ("run2 AutoFitRows:     " + (RunWrap $excel "JYYJHelper.AutoFitRows"))

# Show what the module actually contains after run2 (to check for residue)
Write-Host "----- module content after run2 -----"
try { Write-Host $mod.CodeModule.Lines(1,[int]$mod.CodeModule.CountOfLines) } catch { Write-Host "  (read fail)" }
Write-Host "-------------------------------------"
$wb.Close($false)
# test #2 workbook to test "no Option duplicate"
$excel.Quit(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null