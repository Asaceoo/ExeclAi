Start-Sleep -Seconds 4
Write-Output '---Excel 进程---'
Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object Id, MainWindowTitle | Format-Table -AutoSize

Write-Output '---DisabledItems (加载项被 Excel 禁用)---'
$k = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems'
if (Test-Path $k) {
    $p = Get-ItemProperty $k
    $p.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' } | ForEach-Object { Write-Output ($_.Name + ' => ' + $_.Value) }
} else {
    Write-Output '(无 DisabledItems，加载项未被禁用)'
}

Write-Output '---Addin 注册项状态---'
$k2 = 'HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelAiAssistant.AddIn'
if (Test-Path $k2) {
    Get-ItemProperty $k2 | Format-List Description, FriendlyName, LoadBehavior
} else {
    Write-Output '(未找到 AddIn 注册项)'
}