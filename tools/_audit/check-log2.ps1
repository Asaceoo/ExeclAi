$ErrorActionPreference = 'SilentlyContinue'
$log = 'C:\Users\iamly\AppData\Local\Temp\ExcelAiAssistant.log'
$all = @(Get-Content $log)
Write-Output ('total lines: ' + $all.Count)
Write-Output '=== last 25 lines ==='
$all | Select-Object -Last 25
Write-Output ''
Write-Output '=== entries after 20:00 (current build run) ==='
$all | Where-Object { $_ -match '^2026-08-28 20:' }
Write-Output ''
Write-Output '=== probe-run ensure/create/fail counts (since 20:00) ==='
$new = $all | Where-Object { $_ -match '^2026-08-28 20:' }
Write-Output ('EnsureRule CREATE: ' + (($new | Select-String 'EnsureRule CREATE').Count))
Write-Output ('update selection failed: ' + (($new | Select-String 'update selection failed').Count))
Write-Output ('rule delete failed: ' + (($new | Select-String 'rule delete failed').Count))
Write-Output ('render failed: ' + (($new | Select-String 'render failed').Count))
