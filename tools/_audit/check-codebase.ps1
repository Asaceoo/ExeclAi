$ErrorActionPreference = 'SilentlyContinue'
Write-Output '=== source ==='
Get-Item 'd:\execl1.0\src\ExcelAiAssistant.Addin\SpotlightService.cs' | Select-Object LastWriteTime, Length | Format-List

Write-Output '=== all build outputs ==='
Get-ChildItem 'd:\execl1.0\bin' -Recurse -Filter 'ExcelAiAssistant.Addin.dll' | Select-Object FullName, LastWriteTime, Length | Format-List

Write-Output '=== app install copies ==='
Get-ChildItem 'C:\Users\iamly\AppData\Local\XaiAssistant\app' -Recurse -Filter 'ExcelAiAssistant.Addin.dll' | Select-Object FullName, LastWriteTime, Length | Format-List

Write-Output '=== registry CodeBase (CLSID InprocServer32) ==='
$clsid = '8F7A6D91-4C52-45BE-9F55-7D8C2A31E001'
$k = "HKCU:\Software\Classes\CLSID\{$clsid}\InprocServer32"
if (Test-Path $k) {
    Get-ItemProperty $k | Format-List CodeBase, Assembly, RuntimeVersion, '(default)'
} else {
    Write-Output "key not found: $k"
}

Write-Output '=== registry CodeBase (Addins) ==='
$ak = 'HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelAiAssistant.AddIn'
if (Test-Path $ak) { Get-ItemProperty $ak | Format-List * } else { Write-Output "addins key not found: $ak" }
