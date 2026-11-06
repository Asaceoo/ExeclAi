[CmdletBinding()]
param(
    [ValidateSet("Auto", "x64", "x86")]
    [string]$Architecture = "Auto",
    [switch]$CloseExcel,
    [switch]$Uninstall,
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release"
)
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$excelPath = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe" -ErrorAction SilentlyContinue)."(default)"
if ([string]::IsNullOrWhiteSpace($excelPath) -or -not (Test-Path $excelPath)) { throw "未找到 Excel。请先安装受支持的 Microsoft Excel。" }

function Get-PortableExecutableArchitecture([string]$path) {
    $bytes = [IO.File]::ReadAllBytes($path)
    $offset = [BitConverter]::ToInt32($bytes, 0x3C)
    $machine = [BitConverter]::ToUInt16($bytes, $offset + 4)
    if ($machine -eq 0x8664) { return "x64" }
    if ($machine -eq 0x014C) { return "x86" }
    throw "无法识别的 Excel 可执行文件架构：0x{0:X}" -f $machine
}

$detectArchitecture = Get-PortableExecutableArchitecture $excelPath
if ($Architecture -eq "Auto") { $Architecture = $detectArchitecture }
if ($Architecture -ne $detectArchitecture) {
    throw "目标架构 $Architecture 与本机 Excel 架构 $detectArchitecture 不一致。"
}

$frameworkDirectory = if ($Architecture -eq "x64") { "Framework64" } else { "Framework" }
$regAsm = Join-Path "C:\Windows\Microsoft.NET\$frameworkDirectory\v4.0.30319" "RegAsm.exe"
$outputDirectory = Join-Path $root "src\ExcelAiAssistant.Addin\bin\$Architecture\$Configuration\net48\win-$Architecture"
$assembly = Join-Path $outputDirectory "ExcelAiAssistant.Addin.dll"
if (-not (Test-Path $regAsm)) { throw "未找到 RegAsm：$regAsm" }
if (-not (Test-Path $assembly)) { throw "未找到加载项程序集：$assembly。请先构建 $Architecture。" }

$runtimeKeys = @(
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}",
    "HKCU:\SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}",
    "HKLM:\SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
)
$webView2Version = $null
foreach ($key in $runtimeKeys) {
    if (Test-Path $key) {
        $candidate = (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue).pv
        if ($candidate) { $webView2Version = $candidate; break }
    }
}
if (-not $webView2Version) { throw "未检测到 WebView2 Runtime。请安装 Evergreen WebView2 Runtime 后重试。" }
if ([version]$webView2Version -lt [version]"1.0.2651.64") {
    Write-Warning "WebView2 版本 $webView2Version 低于包引用版本 1.0.2651.64，建议升级运行时。"
}

$addinKey = "HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelAiAssistant.AddIn"
if ($CloseExcel -and (Get-Process -Name EXCEL -ErrorAction SilentlyContinue)) {
    Get-Process -Name EXCEL | ForEach-Object { $_.CloseMainWindow() } | Out-Null
    Start-Sleep -Seconds 2
    Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Where-Object { -not $_.HasExited } | Stop-Process -Force
}

$template = Join-Path $PSScriptRoot "addin-user-hkcu.reg"
if (-not (Test-Path $template)) { throw "未找到 HKCU 注册模板：$template" }

# 动态生成基于当前构建产物的代码基路径
$codeBase = "file:///" + ($assembly -replace '\\','/')
$regText = [IO.File]::ReadAllText($template)
$regText = $regText -replace 'CodeBase"="[^"]*"', ('CodeBase"="' + $codeBase + '"')

$tmpReg = Join-Path $env:TEMP "excelai-addin-hkcu.reg"
[IO.File]::WriteAllText($tmpReg, $regText, (New-Object System.Text.UTF8Encoding($false)))

New-Item -Path $addinKey -Force | Out-Null
if ($Uninstall) {
    Remove-Item -Path $addinKey -Recurse -Force -ErrorAction SilentlyContinue
} else {
    # 免管理员：先写入 HKCU\Software\Classes 的 .NET COM 类型注册，再写 Excel Addins 开关
    $regProc = Start-Process -FilePath "reg.exe" -ArgumentList @("import", "`"$tmpReg`"") -Wait -PassThru -WindowStyle Hidden
    if ($regProc.ExitCode -ne 0) { throw "导入 HKCU 注册失败，退出码 $($regProc.ExitCode)。" }
    New-ItemProperty -Path $addinKey -Name "Description" -Value "OpenAI-compatible AI assistant for Excel." -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $addinKey -Name "FriendlyName" -Value "JYYJ助手" -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $addinKey -Name "LoadBehavior" -Value 3 -PropertyType DWord -Force | Out-Null
}

[PSCustomObject]@{
    Architecture = $Architecture
    Assembly = $assembly
    Excel = $excelPath
    WebView2 = $webView2Version
    Registration = "HKCU"
    Manifest = $addinKey
}