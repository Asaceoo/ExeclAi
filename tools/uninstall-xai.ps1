# 一键卸载 JYYJ助手 Excel 加载项：反注册 COM、移除 Excel Addin 注册表键与安装目录。
[CmdletBinding()]
param(
    [ValidateSet("Auto", "x64", "x86")]
    [string]$Architecture = "Auto",
    [switch]$CloseExcel
)
$ErrorActionPreference = "Stop"
$installRoot = Join-Path (Join-Path $env:LOCALAPPDATA "XaiAssistant") "app"

$excelPath = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe" -ErrorAction SilentlyContinue)."(default)"
if ($Architecture -eq "Auto") {
    if ([string]::IsNullOrWhiteSpace($excelPath) -or -not (Test-Path $excelPath)) {
        # 无法探测架构时按目录探测
        if (Test-Path (Join-Path $installRoot "x64")) { $Architecture = "x64" }
        elseif (Test-Path (Join-Path $installRoot "x86")) { $Architecture = "x86" }
        else { $Architecture = "x64" }
    } else {
        $b = [IO.File]::ReadAllBytes($excelPath)
        $o = [BitConverter]::ToInt32($b, 0x3C)
        $machine = [BitConverter]::ToUInt16($b, $o + 4)
        if ($machine -eq 0x8664) { $Architecture = "x64" } else { $Architecture = "x86" }
    }
}

$installDir = Join-Path $installRoot $Architecture
if (-not (Test-Path $installDir)) { Write-Host "未找到已安装目录：$installDir（可能尚未安装）"; return }

if ($CloseExcel -and (Get-Process -Name EXCEL -ErrorAction SilentlyContinue)) {
    Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { try { $_.CloseMainWindow() } catch {} }
    Start-Sleep -Seconds 2
    Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Where-Object { -not $_.HasExited } | Stop-Process -Force
}

$addinKey = "HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelAiAssistant.AddIn"
if (Test-Path $addinKey) { Remove-Item $addinKey -Force }

# 移除按用户注册的 .NET COM 项（CLSID + ProgID）
foreach ($cls in @("8F7A6D91-4C52-45BE-9F55-7D8C2A31E001", "8F7A6D91-4C52-45BE-9F55-7D8C2A31E002", "8F7A6D91-4C52-45BE-9F55-7D8C2A31E003")) {
    $clsKey = "HKCU:\Software\Classes\CLSID\$cls"
    if (Test-Path $clsKey) { Remove-Item $clsKey -Recurse -Force }
}
foreach ($prog in @("ExcelAiAssistant.AddIn", "ExcelAiAssistant.TaskPaneHost", "ExcelAiAssistant.DeepSeekPaneHost")) {
    $progKey = "HKCU:\Software\Classes\$prog"
    if (Test-Path $progKey) { Remove-Item $progKey -Recurse -Force -ErrorAction SilentlyContinue }
}

$assembly = Join-Path $installDir "ExcelAiAssistant.Addin.dll"
if (Test-Path $assembly) {
    $frameworkDirectory = if ($Architecture -eq "x64") { "Framework64" } else { "Framework" }
    $regAsm = Join-Path "C:\Windows\Microsoft.NET\$frameworkDirectory\v4.0.30319" "RegAsm.exe"
    if (Test-Path $regAsm) {
        & $regAsm $assembly /unregister /nologo 2>$null
    }
}

Remove-Item $installDir -Recurse -Force
Remove-Item (Join-Path $installRoot $Architecture) -Recurse -Force -ErrorAction SilentlyContinue

[PSCustomObject]@{ Uninstalled = $installDir; Message = "卸载完成。" }