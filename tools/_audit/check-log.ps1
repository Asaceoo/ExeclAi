$ErrorActionPreference = 'SilentlyContinue'
Write-Output '=== running Excel processes ==='
Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object Id, StartTime, MainWindowTitle | Format-Table -AutoSize

Write-Output '=== log file info ==='
Get-Item 'C:\Users\iamly\AppData\Local\Temp\ExcelAiAssistant.log' | Select-Object LastWriteTime, Length | Format-List

Write-Output '=== failure pattern counts ==='
$log = Get-Content 'C:\Users\iamly\AppData\Local\Temp\ExcelAiAssistant.log'
Write-Output ("total lines: " + $log.Count)
Write-Output ("render failed: " + (($log | Select-String 'render failed').Count))
Write-Output ("update selection failed: " + (($log | Select-String 'update selection failed').Count))
Write-Output ("rule deleted: " + (($log | Select-String 'rule deleted').Count))
Write-Output ("rule delete failed: " + (($log | Select-String 'rule delete failed').Count))
Write-Output ("EnsureRule CREATE: " + (($log | Select-String 'EnsureRule CREATE').Count))
Write-Output ("SetMode: " + (($log | Select-String 'SetMode mode=').Count))
Write-Output ("delayed view cleanup fired: " + (($log | Select-String 'delayed view cleanup fired').Count))

Write-Output '=== render failed samples ==='
$log | Select-String 'render failed' | Select-Object -Last 10 | ForEach-Object { $_.Line }
Write-Output '=== update selection failed samples ==='
$log | Select-String 'update selection failed' | Select-Object -Last 10 | ForEach-Object { $_.Line }
