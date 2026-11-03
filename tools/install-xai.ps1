# 一键安装 JYYJ助手 Excel 加载项（从当前文件夹的 app\ 或 -Source 指向的目录拷贝到用户目录并注册）。
[CmdletBinding()]
param(
    [ValidateSet("Auto", "x64", "x86")]
    [string]$Architecture = "Auto",
    [string]$Source = "",
    [switch]$Check,
    [switch]$CloseExcel
)
$ErrorActionPreference = "Stop"
$defaultSource = Join-Path $PSScriptRoot "app"
if ($Source -eq "") { $Source = $defaultSource }
$installRoot = Join-Path (Join-Path $env:LOCALAPPDATA "XaiAssistant") "app"

# ---------- 前置检查 ----------
$excelPath = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe" -ErrorAction SilentlyContinue)."(default)"
if ([string]::IsNullOrWhiteSpace($excelPath) -or -not (Test-Path $excelPath)) { throw "未找到 Excel。请先安装受支持的 Microsoft Excel。" }

function Get-PortableExecutableArchitecture([string]$path) {
    $bytes = [IO.File]::ReadAllBytes($path)
    $offset = [BitConverter]::ToInt32($bytes, 0x3C)
    $machine = [BitConverter]::ToUInt16($bytes, $offset + 4)
    if ($machine -eq 0x8664) { return "x64" }
    if ($machine -eq 0x014C) { return "x86" }
    throw "无法识别的 Excel 可执行文件架构。"
}
$detectArchitecture = Get-PortableExecutableArchitecture $excelPath
if ($Architecture -eq "Auto") { $Architecture = $detectArchitecture }
if ($Architecture -ne $detectArchitecture) { throw "目标架构 $Architecture 与本机 Excel 架构 $detectArchitecture 不一致。" }

if (-not (Test-Path $Source)) { throw "未找到安装源：$Source" }
$assembly = Join-Path $Source "ExcelAiAssistant.Addin.dll"
if (-not (Test-Path $assembly)) { throw "安装源缺少 ExcelAiAssistant.Addin.dll：$Source" }

$net48 = (Test-Path "C:\Windows\Microsoft.NET\Framework\v4.0.30319\System.dll") -and ((Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -ErrorAction SilentlyContinue).Release -ge 528040)
if (-not $net48) { throw "未检测到 .NET Framework 4.8 或更高版本。" }

$webView2Version = $null
foreach ($key in @(
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}",
        "HKCU:\SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}",
        "HKLM:\SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}")) {
    if (Test-Path $key) { $candidate = (Get-ItemProperty $key -ErrorAction SilentlyContinue).pv; if ($candidate) { $webView2Version = $candidate; break } }
}
if (-not $webView2Version) { throw "未检测到 WebView2 Runtime。请先安装 Evergreen WebView2 Runtime。" }

# ---------- 按用户注册 .NET COM 类（无需管理员） ----------
# 结构对齐 RegAsm /codebase 生成的 HKCU\Software\Classes 项，供 Excel 调用方解析。
# 注意：必须包含 CLSID\ProgId 子键、Implemented Categories、InprocServer32\<版本> 子键、
# ProgID\CLSID 子键，否则 Excel 无法将该 COM 加载项识别出来（功能区不出现）。
function Register-ComClassPerUser([string]$Guid, [string]$ProgId, [string]$FullType, [string]$AssemblyFullName, [string]$CodeBase) {
    # 用 .NET Registry API 直写 HKCU\Software\Classes：实测 PowerShell PSProvider（HKCU:\）的
    # New-Item / Set-ItemProperty 在本安装脚本执行上下文里写不到真实注册表（跨进程不可见），
    # 而 [Microsoft.Win32.Registry] 写可持久落盘；改用后者保证新 COM 类（如 DS 停靠窗）必被注册。
    $version = ([System.Reflection.AssemblyName]::GetAssemblyName($installedAssembly).Version).ToString()
    $cu = [Microsoft.Win32.Registry]::CurrentUser
    $clsBase = $cu.CreateSubKey("Software\Classes\CLSID\$Guid")
    try {
        $clsBase.SetValue("", $FullType)
        $clsBase.SetValue("Assembly", $AssemblyFullName)
        $clsBase.SetValue("Class", $FullType)
        $clsBase.SetValue("ProgID", $ProgId)
    } finally { $clsBase.Close() }

    $clsProgId = $cu.CreateSubKey("Software\Classes\CLSID\$Guid\ProgId")
    try { $clsProgId.SetValue("", $ProgId) } finally { $clsProgId.Close() }

    $cu.CreateSubKey("Software\Classes\CLSID\$Guid\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}").Close()

    $inproc = $cu.CreateSubKey("Software\Classes\CLSID\$Guid\InprocServer32")
    try {
        $inproc.SetValue("", "mscoree.dll")
        $inproc.SetValue("ThreadingModel", "Both")
        $inproc.SetValue("Class", $FullType)
        $inproc.SetValue("Assembly", $AssemblyFullName)
        $inproc.SetValue("CodeBase", $CodeBase)
        $inproc.SetValue("RuntimeVersion", "v4.0.30319")
    } finally { $inproc.Close() }

    $inprocVer = $cu.CreateSubKey("Software\Classes\CLSID\$Guid\InprocServer32\$version")
    try {
        $inprocVer.SetValue("Class", $FullType)
        $inprocVer.SetValue("Assembly", $AssemblyFullName)
        $inprocVer.SetValue("RuntimeVersion", "v4.0.30319")
        $inprocVer.SetValue("CodeBase", $CodeBase)
    } finally { $inprocVer.Close() }

    $progBase = $cu.CreateSubKey("Software\Classes\$ProgId")
    try { $progBase.SetValue("", $FullType) } finally { $progBase.Close() }
    $progCls = $cu.CreateSubKey("Software\Classes\$ProgId\CLSID")
    try { $progCls.SetValue("", "{$Guid}") } finally { $progCls.Close() }
}

if ($Check) {
    [PSCustomObject]@{ Check = "OK"; Architecture = $Architecture; DotNet48 = $true; WebView2 = $webView2Version; Source = $Source; Excel = $excelPath }
    return
}

# ---------- 关闭 Excel ----------
function Close-ExcelIfRunning {
    if ($CloseExcel -and (Get-Process -Name EXCEL -ErrorAction SilentlyContinue)) {
        Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { try { $_.CloseMainWindow() } catch {} }
        Start-Sleep -Seconds 2
        Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Where-Object { -not $_.HasExited } | Stop-Process -Force
    }
}
Close-ExcelIfRunning

# ---------- 拷贝文件 ----------
$installDir = Join-Path $installRoot $Architecture
New-Item -ItemType Directory -Path $installDir -Force | Out-Null
Get-ChildItem $Source -File | Copy-Item -Destination $installDir -Force
# SQLite.Interop.dll 等原生互操作库按架构放在 x64/x86 子目录，必须一并拷贝，
# 否则安装后 System.Data.SQLite 初始化失败、WebView2 侧边栏空白。
foreach ($native in @("x64", "x86")) {
    $srcNative = Join-Path $Source $native
    if (Test-Path $srcNative) {
        Copy-Item $srcNative -Destination $installDir -Recurse -Force
    }
}
$srcSkills = Join-Path $Source "skills"
$dstSkills = Join-Path $installDir "skills"
if (Test-Path $srcSkills) {
    if (Test-Path $dstSkills) { Remove-Item $dstSkills -Recurse -Force }
    Copy-Item $srcSkills -Destination $installDir -Recurse -Force
}

# ---------- 注册 COM + Excel Addin ----------
$installedAssembly = Join-Path $installDir "ExcelAiAssistant.Addin.dll"
$assemblyFullName = [System.Reflection.AssemblyName]::GetAssemblyName($installedAssembly).FullName
$codeBase = (New-Object System.Uri($installedAssembly)).AbsoluteUri
Register-ComClassPerUser "8F7A6D91-4C52-45BE-9F55-7D8C2A31E001" "ExcelAiAssistant.AddIn" "ExcelAiAssistant.Addin.ExcelAiAddIn" $assemblyFullName $codeBase
Register-ComClassPerUser "8F7A6D91-4C52-45BE-9F55-7D8C2A31E002" "ExcelAiAssistant.TaskPaneHost" "ExcelAiAssistant.Addin.ExcelAiTaskPaneHost" $assemblyFullName $codeBase
Register-ComClassPerUser "8F7A6D91-4C52-45BE-9F55-7D8C2A31E003" "ExcelAiAssistant.DeepSeekPaneHost" "ExcelAiAssistant.Addin.DeepSeekTaskPaneHost" $assemblyFullName $codeBase

$addinKey = "HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelAiAssistant.AddIn"
New-Item -Path $addinKey -Force | Out-Null
New-ItemProperty -Path $addinKey -Name "Description" -Value "OpenAI-compatible AI assistant for Excel." -PropertyType String -Force | Out-Null
New-ItemProperty -Path $addinKey -Name "FriendlyName" -Value "JYYJ助手" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $addinKey -Name "LoadBehavior" -Value 3 -PropertyType DWord -Force | Out-Null

[PSCustomObject]@{
    InstalledTo = $installDir
    Assembly = $installedAssembly
    Architecture = $Architecture
    WebView2 = $webView2Version
    Message = "安装完成。启动 Excel 后应出现JYYJ助手标签与侧边栏。"
}