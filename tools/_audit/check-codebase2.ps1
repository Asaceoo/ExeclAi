$ErrorActionPreference = 'SilentlyContinue'
Write-Output '=== CodeBase 指向的源码构建输出 ==='
Get-ChildItem 'd:\execl1.0\src\ExcelAiAssistant.Addin\bin' -Recurse -Filter 'ExcelAiAssistant.Addin.dll' | Select-Object FullName, LastWriteTime, Length | Format-List

Write-Output '=== dist 包 ==='
Get-ChildItem 'd:\execl1.0\dist' -Recurse -Filter 'ExcelAiAssistant.Addin.dll' -ErrorAction SilentlyContinue | Select-Object FullName, LastWriteTime, Length | Format-List

Write-Output '=== src 目录 mtime ==='
Get-Item 'd:\execl1.0\src\ExcelAiAssistant.Addin\SpotlightService.cs' | Select-Object LastWriteTime | Format-List
