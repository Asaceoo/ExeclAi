$ErrorActionPreference = 'SilentlyContinue'
Write-Output '=== Excel processes detail ==='
Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object {
    $p = $_
    $title = $p.MainWindowTitle
    Write-Output ("PID={0} Start={1} Title='{2}'" -f $p.Id, $p.StartTime, $title)
}

Write-Output '=== is DLL locked by a process? ==='
$dll = 'D:\execl1.0\src\ExcelAiAssistant.Addin\bin\x64\Release\net48\win-x64\ExcelAiAssistant.Addin.dll'
try {
    $fs = [System.IO.File]::Open($dll, 'Open', 'Read', 'None')
    $fs.Close()
    Write-Output 'DLL is NOT locked (can open exclusive)'
} catch {
    Write-Output ('DLL IS LOCKED: ' + $_.Exception.Message)
}
