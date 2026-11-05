# JYYJ助手 合并安装器注册编排（由 64 位 PowerShell 经 {sysnative} 静默执行）
# 自动探测本机 Excel 位数（读 excel.exe PE 头），选定 app\x64 或 app\x86，
# 并用 .NET Registry API 显式写 64/32 视图的 .NET COM 类 + Excel Addins 键。
# 显式 RegistryView 使本脚本对进程位数免疫，32/64 位 Excel 各从对应视图加载。
# M1 新增：探测 WPS 表格（et.exe），按其 PE 位数注册 COM 类，并写 Kingsoft
#   ET\Addins 键 + AddinsWL 白名单（个人版必需，双视图都写以免依赖 WPS 读键视图）。
$ErrorActionPreference = 'Stop'

# ---------- 探测 Excel 位数 ----------
$excelPath = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe' -ErrorAction SilentlyContinue).'(default)'
if (-not (Test-Path $excelPath)) { throw '未找到 Excel。请先安装受支持的 Microsoft Excel。' }
$bytes = [IO.File]::ReadAllBytes($excelPath)
$offset = [BitConverter]::ToInt32($bytes, 0x3C)
$machine = [BitConverter]::ToUInt16($bytes, $offset + 4)
$arch = $null
$view = $null
if ($machine -eq 0x8664) { $arch = 'x64'; $view = [Microsoft.Win32.RegistryView]::Registry64 }
elseif ($machine -eq 0x014C) { $arch = 'x86'; $view = [Microsoft.Win32.RegistryView]::Registry32 }
else { throw ("检测到非常规 CPU 架构的 Excel（PE 头 0x{0:X4}，可能为 ARM64 设备）。本安装包暂不支持 ARM64 版 Excel，请联系技术支持。" -f $machine) }

# 清理历史版本（<=2.19.4.0）误写的无花括号垃圾键：Excel 只读 CLSID\{GUID}，
# 无括号键是死键，但会永久残留误导排查，注册前顺手清掉（升级自愈路径）。
$legacyRoot = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, [Microsoft.Win32.RegistryView]::Default)
$legacyClsid = $legacyRoot.OpenSubKey('Software\Classes\CLSID', $true)
if ($legacyClsid) {
    foreach ($g in @('8F7A6D91-4C52-45BE-9F55-7D8C2A31E001', '8F7A6D91-4C52-45BE-9F55-7D8C2A31E002', '8F7A6D91-4C52-45BE-9F55-7D8C2A31E003')) {
        try { $legacyClsid.DeleteSubKeyTree($g, $false) } catch { }
    }
    $legacyClsid.Close()
}
$legacyRoot.Close()

$dir = (Join-Path $PSScriptRoot 'app') + '\' + $arch
$assembly = Join-Path $dir 'ExcelAiAssistant.Addin.dll'
if (-not (Test-Path $assembly)) { throw "安装目录缺少加载项：$assembly" }

$asmFull = [System.Reflection.AssemblyName]::GetAssemblyName($assembly).FullName
$version = [System.Reflection.AssemblyName]::GetAssemblyName($assembly).Version.ToString()
$codebase = (New-Object System.Uri($assembly)).AbsoluteUri
$root = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
$catId = '{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}'

function Register-ComClass {
    param($Key, $Guid, $ProgId, $FullType, $AsmFull, $CodeBase)
    # CLSID 键必须带花括号：Excel 按 CLSID\{GUID} 读取，缺括号会写到无效键导致加载报错
    $clsidPath = "Software\Classes\CLSID\{$Guid}"
    $cls = $Key.CreateSubKey("$clsidPath")
    try {
        $cls.SetValue('', $FullType)
        $cls.SetValue('Assembly', $AsmFull)
        $cls.SetValue('Class', $FullType)
        $cls.SetValue('ProgID', $ProgId)
    } finally { $cls.Close() }

    $clsProg = $Key.CreateSubKey("$clsidPath\ProgId")
    try { $clsProg.SetValue('', $ProgId) } finally { $clsProg.Close() }

    $Key.CreateSubKey("$clsidPath\Implemented Categories\$catId").Close()

    $in = $Key.CreateSubKey("$clsidPath\InprocServer32")
    try {
        $in.SetValue('', 'mscoree.dll')
        $in.SetValue('ThreadingModel', 'Both')
        $in.SetValue('Class', $FullType)
        $in.SetValue('Assembly', $AsmFull)
        $in.SetValue('CodeBase', $CodeBase)
        $in.SetValue('RuntimeVersion', 'v4.0.30319')
    } finally { $in.Close() }

    $inV = $Key.CreateSubKey("$clsidPath\InprocServer32\$version")
    try {
        $inV.SetValue('Class', $FullType)
        $inV.SetValue('Assembly', $AsmFull)
        $inV.SetValue('RuntimeVersion', 'v4.0.30319')
        $inV.SetValue('CodeBase', $CodeBase)
    } finally { $inV.Close() }

    $prog = $Key.CreateSubKey("Software\Classes\$ProgId")
    try { $prog.SetValue('', $FullType) } finally { $prog.Close() }
    $progCls = $Key.CreateSubKey("Software\Classes\$ProgId\CLSID")
    try { $progCls.SetValue('', "{$Guid}") } finally { $progCls.Close() }

    $add = $Key.CreateSubKey("Software\Microsoft\Office\Excel\Addins\$ProgId")
    try {
        $add.SetValue('Description', 'OpenAI-compatible AI assistant for Excel.')
        $add.SetValue('FriendlyName', 'JYYJ助手')
        $add.SetValue('LoadBehavior', 3, [Microsoft.Win32.RegistryValueKind]::DWord)
    } finally { $add.Close() }
}

Register-ComClass $root '8F7A6D91-4C52-45BE-9F55-7D8C2A31E001' 'ExcelAiAssistant.AddIn' 'ExcelAiAssistant.Addin.ExcelAiAddIn' $asmFull $codebase
Register-ComClass $root '8F7A6D91-4C52-45BE-9F55-7D8C2A31E002' 'ExcelAiAssistant.TaskPaneHost' 'ExcelAiAssistant.Addin.ExcelAiTaskPaneHost' $asmFull $codebase
Register-ComClass $root '8F7A6D91-4C52-45BE-9F55-7D8C2A31E003' 'ExcelAiAssistant.DeepSeekPaneHost' 'ExcelAiAssistant.Addin.DeepSeekTaskPaneHost' $asmFull $codebase
$root.Close()

# ---------- WPS 表格支持（M1：配方来源 docs/M0-WPS加载可行性报告.md） ----------
$progids = @('ExcelAiAssistant.AddIn', 'ExcelAiAssistant.TaskPaneHost', 'ExcelAiAssistant.DeepSeekPaneHost')
$fulltypes = @{
    'ExcelAiAssistant.AddIn'            = 'ExcelAiAssistant.Addin.ExcelAiAddIn'
    'ExcelAiAssistant.TaskPaneHost'     = 'ExcelAiAssistant.Addin.ExcelAiTaskPaneHost'
    'ExcelAiAssistant.DeepSeekPaneHost' = 'ExcelAiAssistant.Addin.DeepSeekTaskPaneHost'
}
$guidByProgId = @{
    'ExcelAiAssistant.AddIn'            = '8F7A6D91-4C52-45BE-9F55-7D8C2A31E001'
    'ExcelAiAssistant.TaskPaneHost'     = '8F7A6D91-4C52-45BE-9F55-7D8C2A31E002'
    'ExcelAiAssistant.DeepSeekPaneHost' = '8F7A6D91-4C52-45BE-9F55-7D8C2A31E003'
}

# 探测 WPS 表格：常见安装根下取版本号最高的 office6\et.exe
$etPath = $null
foreach ($base in @(
        "$env:LOCALAPPDATA\Kingsoft\WPS Office",
        'C:\Program Files (x86)\Kingsoft\WPS Office',
        'C:\Program Files\Kingsoft\WPS Office')) {
    if (-not (Test-Path $base)) { continue }
    $cand = Get-ChildItem $base -Directory -ErrorAction SilentlyContinue |
        Sort-Object Name -Descending |
        ForEach-Object { Join-Path $_.FullName 'office6\et.exe' } |
        Where-Object { Test-Path $_ } |
        Select-Object -First 1
    if ($cand) { $etPath = $cand; break }
}

if ($etPath) {
    $etBytes = [IO.File]::ReadAllBytes($etPath)
    $etOffset = [BitConverter]::ToInt32($etBytes, 0x3C)
    $etMachine = [BitConverter]::ToUInt16($etBytes, $etOffset + 4)
    $wpsArch = $null
    $wpsView = $null
    if ($etMachine -eq 0x8664) { $wpsArch = 'x64'; $wpsView = [Microsoft.Win32.RegistryView]::Registry64 }
    elseif ($etMachine -eq 0x014C) { $wpsArch = 'x86'; $wpsView = [Microsoft.Win32.RegistryView]::Registry32 }
    else { Write-Output ("WPS-SKIPPED 未知架构 PE=0x{0:X4}" -f $etMachine) }

    if ($wpsArch) {
        # WPS 与 Excel 位数不同时，按 WPS 视图补注册 COM 类（CodeBase 指向对应架构 app 目录）
        if ($wpsArch -ne $arch) {
            $wpsDir = (Join-Path $PSScriptRoot 'app') + '\' + $wpsArch
            $wpsAssembly = Join-Path $wpsDir 'ExcelAiAssistant.Addin.dll'
            if (Test-Path $wpsAssembly) {
                $wAsmFull = [System.Reflection.AssemblyName]::GetAssemblyName($wpsAssembly).FullName
                $wCodebase = (New-Object System.Uri($wpsAssembly)).AbsoluteUri
                $wpsRoot = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $wpsView)
                foreach ($p in $progids) {
                    Register-ComClass $wpsRoot $guidByProgId[$p] $p $fulltypes[$p] $wAsmFull $wCodebase
                }
                $wpsRoot.Close()
            } else {
                throw "安装目录缺少 WPS 架构加载项：$wpsAssembly"
            }
        }

        # Kingsoft Addins 键 + AddinsWL 白名单：双视图都写（64/32 位 WPS 各读各的视图）
        foreach ($kv in @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)) {
            $kroot = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $kv)
            foreach ($p in $progids) {
                $k = $kroot.CreateSubKey("Software\Kingsoft\Office\ET\Addins\$p")
                try {
                    $k.SetValue('Description', 'OpenAI-compatible AI assistant for WPS ET.')
                    $k.SetValue('FriendlyName', 'JYYJ助手')
                    $k.SetValue('LoadBehavior', 3, [Microsoft.Win32.RegistryValueKind]::DWord)
                } finally { $k.Close() }
                $wl = $kroot.CreateSubKey('Software\Kingsoft\Office\ET\AddinsWL')
                try { $wl.SetValue($p, '') } finally { $wl.Close() }
            }
            $kroot.Close()
        }
        Write-Output ("WPS-REGISTERED et=" + $etPath + " arch=" + $wpsArch)
    }
} else {
    Write-Output 'WPS-NOT-FOUND（未检测到 WPS 表格，仅注册 Excel）'
}

Write-Output ("REGISTERED arch=" + $arch + " view=" + $view + " dir=" + $dir)