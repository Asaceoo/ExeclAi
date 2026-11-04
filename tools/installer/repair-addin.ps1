# JYYJ助手 加载项自愈脚本（M29 事故沉淀，第二轮审查加固）
# 适用：Excel/WPS 功能区不显示「JYYJ助手」选项卡，或侧边栏打不开。
# 作用：修正 LoadBehavior、加入 DoNotDisableAddinList、按实际 DLL 重写三个 CLSID 的注册值。
# 安全边界：只改 HKCU（当前用户）双视图，不需要管理员权限；不改动任何用户数据文件。
#   默认只接受 %LOCALAPPDATA%\XaiAssistant\app 下的安装目录；非标准目录必须显式 -Force。
# 用法：powershell -NoProfile -ExecutionPolicy Bypass -File repair-addin.ps1
#       可选 -AppDir "D:\custom\app" -Force

param(
    [string]$AppDir = (Join-Path $env:LOCALAPPDATA 'XaiAssistant\app'),
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

$progIds = @{
    'ExcelAiAssistant.AddIn'            = @{ Guid = '8F7A6D91-4C52-45BE-9F55-7D8C2A31E001'; Type = 'ExcelAiAssistant.Addin.ExcelAiAddIn' }
    'ExcelAiAssistant.TaskPaneHost'     = @{ Guid = '8F7A6D91-4C52-45BE-9F55-7D8C2A31E002'; Type = 'ExcelAiAssistant.Addin.ExcelAiTaskPaneHost' }
    'ExcelAiAssistant.DeepSeekPaneHost' = @{ Guid = '8F7A6D91-4C52-45BE-9F55-7D8C2A31E003'; Type = 'ExcelAiAssistant.Addin.DeepSeekTaskPaneHost' }
}

Write-Host "== JYYJ助手 加载项自愈 ==" -ForegroundColor Cyan

# 审查 S5'：规范化并限制 AppDir——否则脚本可被诱导把 CodeBase 指向任意目录形成持久化
if ([string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) {
    Write-Host "[失败] 无法定位 %LOCALAPPDATA%，请在正常用户会话中运行。" -ForegroundColor Red
    exit 1
}

if (-not [System.IO.Path]::IsPathRooted($AppDir)) {
    Write-Host "[失败] -AppDir 必须是绝对路径：$AppDir" -ForegroundColor Red
    exit 1
}

$resolvedAppDir = [System.IO.Path]::GetFullPath($AppDir)
if ($resolvedAppDir.StartsWith('\\')) {
    Write-Host "[失败] 不允许 UNC/网络路径：$resolvedAppDir" -ForegroundColor Red
    exit 1
}

$expectedRoot = [System.IO.Path]::GetFullPath((Join-Path $env:LOCALAPPDATA 'XaiAssistant\app'))
$isStandard = $resolvedAppDir.TrimEnd('\').Equals($expectedRoot.TrimEnd('\'), [System.StringComparison]::OrdinalIgnoreCase)
if (-not $isStandard -and -not $Force) {
    Write-Host "[失败] 非标准安装目录：$resolvedAppDir" -ForegroundColor Red
    Write-Host "       默认只允许 $expectedRoot 。确需使用其他目录请加 -Force。" -ForegroundColor Yellow
    exit 1
}
if (-not $isStandard) {
    Write-Host "[警告] 使用非标准安装目录（-Force）：$resolvedAppDir" -ForegroundColor Yellow
    Write-Host "       该目录可被同用户写改，请确认其可信。" -ForegroundColor Yellow
}

Write-Host "安装目录：$resolvedAppDir"

$targets = @()
foreach ($arch in @('x64', 'x86')) {
    $dll = Join-Path $resolvedAppDir "$arch\ExcelAiAssistant.Addin.dll"
    if (Test-Path $dll) {
        $targets += [pscustomobject]@{ Arch = $arch; Dll = $dll }
    } else {
        # 审查 B7：缺架构时明确告警，不能静默跳过还打印"修复完成"
        Write-Warning "跳过 $arch 视图：缺少 $dll（该位数 Excel/WPS 将无法加载）"
    }
}

if ($targets.Count -eq 0) {
    Write-Host "[失败] 安装目录中找不到 ExcelAiAssistant.Addin.dll：$resolvedAppDir" -ForegroundColor Red
    Write-Host "请先安装 JYYJ助手，或用 -AppDir 指定正确目录。"
    exit 1
}

# 审查 B7：检测机器级策略/禁用，避免"自愈"绕过企业管控
foreach ($ver in @('14.0', '15.0', '16.0')) {
    $policyKey = "HKCU:\Software\Policies\Microsoft\Office\$ver\Excel\Resiliency"
    if (Test-Path $policyKey) {
        Write-Warning "检测到策略键 $policyKey：企业策略可能强制禁用加载项，本脚本不会覆盖策略。"
    }
}
if (Test-Path 'HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems') {
    Write-Warning "检测到 Excel 禁用项列表（Resiliency\DisabledItems）。若其中含本加载项，请先备份该键再手动清理。"
}

foreach ($view in @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)) {
    $arch = if ($view -eq [Microsoft.Win32.RegistryView]::Registry64) { 'x64' } else { 'x86' }
    $target = $targets | Where-Object { $_.Arch -eq $arch } | Select-Object -First 1
    if (-not $target) { continue }

    try {
        $asmName = [System.Reflection.AssemblyName]::GetAssemblyName($target.Dll)
        $asmFull = $asmName.FullName
        $version = $asmName.Version.ToString()
    } catch {
        Write-Warning "读取 $($target.Dll) 的程序集信息失败（文件损坏或位数不符）：$($_.Exception.Message)"
        continue
    }

    $codeBase = (New-Object System.Uri($target.Dll)).AbsoluteUri
    $root = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)

    try {
        foreach ($progId in $progIds.Keys) {
            $meta = $progIds[$progId]
            $base = "Software\Classes\CLSID\{$($meta.Guid)}\InprocServer32"

            $key = $root.CreateSubKey($base)
            if ($null -eq $key) { throw "无法创建注册表键：$base" }
            try {
                $key.SetValue('', 'mscoree.dll')
                $key.SetValue('ThreadingModel', 'Both')
                $key.SetValue('Class', $meta.Type)
                $key.SetValue('Assembly', $asmFull)
                $key.SetValue('CodeBase', $codeBase)
                $key.SetValue('RuntimeVersion', 'v4.0.30319')
            } finally { $key.Close() }

            # 版本子键：Class 必须是本 ProgID 的真实类型（写成 ExcelAiAddIn 会让 CreateCTP 报 0x80004005）
            $versionKey = $root.CreateSubKey("$base\$version")
            if ($null -eq $versionKey) { throw "无法创建版本子键：$base\$version" }
            try {
                $versionKey.SetValue('Class', $meta.Type)
                $versionKey.SetValue('Assembly', $asmFull)
                $versionKey.SetValue('CodeBase', $codeBase)
                $versionKey.SetValue('RuntimeVersion', 'v4.0.30319')
            } finally { $versionKey.Close() }

            $progKey = $root.CreateSubKey("Software\Classes\$progId\CLSID")
            if ($null -ne $progKey) {
                try { $progKey.SetValue('', "{$($meta.Guid)}") } finally { $progKey.Close() }
            }

            # .NET 可管理类类别（贴近 RegAsm 产物，手删键后也能恢复）
            $catKey = $root.CreateSubKey("$base\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}")
            if ($null -ne $catKey) { $catKey.Close() }
        }

        # Excel 加载项：强制启用（3）并加入防自动禁用名单
        $addins = $root.CreateSubKey('Software\Microsoft\Office\Excel\Addins\ExcelAiAssistant.AddIn')
        if ($null -ne $addins) {
            try {
                $addins.SetValue('Description', 'OpenAI-compatible AI assistant for Excel.')
                $addins.SetValue('FriendlyName', 'JYYJ助手')
                $addins.SetValue('LoadBehavior', 3, [Microsoft.Win32.RegistryValueKind]::DWord)
            } finally { $addins.Close() }
        }

        # 审查 B7：Office 2010/2013/2016+ 分别对应 14.0/15.0/16.0
        foreach ($ver in @('14.0', '15.0', '16.0')) {
            $noDisable = $root.CreateSubKey("Software\Microsoft\Office\$ver\Excel\Resiliency\DoNotDisableAddinList")
            if ($null -ne $noDisable) {
                try { $noDisable.SetValue('ExcelAiAssistant.AddIn', 1, [Microsoft.Win32.RegistryValueKind]::DWord) } finally { $noDisable.Close() }
            }
        }

        # WPS：Addins 键 + AddinsWL 白名单
        $etAddin = $root.CreateSubKey('Software\Kingsoft\Office\ET\Addins\ExcelAiAssistant.AddIn')
        if ($null -ne $etAddin) {
            try {
                $etAddin.SetValue('Description', 'OpenAI-compatible AI assistant for WPS ET.')
                $etAddin.SetValue('FriendlyName', 'JYYJ助手')
                $etAddin.SetValue('LoadBehavior', 3, [Microsoft.Win32.RegistryValueKind]::DWord)
            } finally { $etAddin.Close() }
        }

        $wl = $root.CreateSubKey('Software\Kingsoft\Office\ET\AddinsWL')
        if ($null -ne $wl) {
            try { $wl.SetValue('ExcelAiAssistant.AddIn', '') } finally { $wl.Close() }
        }

        Write-Host "  [$($view.ToString())] COM 注册 + 加载项启用已修复：Assembly=$version，LoadBehavior=3" -ForegroundColor Green
    } finally {
        $root.Close()
    }
}

Write-Host ""
Write-Host "修复完成。请完全退出 Excel（含后台进程）后重新打开。" -ForegroundColor Cyan
Write-Host "验证：Excel → 文件 → 选项 → 加载项 → 管理「COM 加载项」→ 转到，确认「JYYJ助手」已勾选。"