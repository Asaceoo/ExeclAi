# ============================================================
#  JYYJ助手 一键诊断（面向最终用户，双击"一键诊断.cmd"运行）
#  作用：自动收集 Microsoft Excel 与 WPS 表格双宿主环境 / 系统依赖 /
#        COM 注册实况 / 运行日志，在桌面生成自包含诊断报告，
#        供技术支持远程定位问题。
#  设计原则：任何单项检查失败都不中断整体（逐项 try/catch），
#            报告必须能生成——哪怕部分信息采集不到，缺失本身就是线索。
#  维护约束：
#    1) ProgID/CLSID 清单必须与 register-auto.ps1 保持同步；
#    2) WPS 探测逻辑（安装根目录、Kingsoft Addins 双视图键）必须与
#       register-auto.ps1 的写入逻辑同源——诊断查的键 = 注册写的键；
#    3) 对无法远程确认真伪的项（如 WPS VBA 模块）只做指引，不做断言。
#  兼容性：Windows PowerShell 2.0+（不使用 Get-CimInstance / 三元等新语法）；
#          显式 RegistryView，对脚本宿主进程位数免疫。
#  用法：diagnose.ps1 [-NoOpen]   （-NoOpen 供自动化验证，不弹记事本）
# ============================================================
param([switch]$NoOpen)

$ErrorActionPreference = 'Continue'
$script:Out = New-Object System.Collections.Generic.List[string]
$script:Issues = New-Object System.Collections.Generic.List[string]

function Add-Line([string]$Text) { $script:Out.Add($Text) }
function Add-Section([string]$Title) {
    Add-Line ''
    Add-Line ('■ ' + $Title)
    Add-Line ('─' * 46)
}
# 安全取值：任何采集动作包进 try/catch，失败返回占位并记录
function Add-Collect([string]$Label, [scriptblock]$Probe) {
    try {
        $v = & $Probe
        if ($null -eq $v -or $v -eq '') { $v = '（空）' }
        Add-Line ($Label + '：' + $v)
        return $v
    } catch {
        Add-Line ($Label + '：采集失败 - ' + $_.Exception.Message)
        return $null
    }
}
function Add-Issue([string]$Text) { $script:Issues.Add($Text) }

# 读 PE 头机器类型判定位数（Excel 与 et.exe 共用；输入可为 exe 全路径）
function Get-PEArch([string]$ExePath) {
    if (-not (Test-Path $ExePath -PathType Leaf)) { return $null }
    $bytes = [IO.File]::ReadAllBytes($ExePath)
    $peOff = [BitConverter]::ToInt32($bytes, 0x3C)
    $machine = [BitConverter]::ToUInt16($bytes, $peOff + 4)
    if ($machine -eq 0x8664) { return 'x64' }
    elseif ($machine -eq 0x014C) { return 'x86' }
    else { return ('未知(0x{0:X4})' -f $machine) }
}

# COM 探测公共路径：创建 → 取版本 → 退出 → 释放（Excel/KET 共用）
# 返回 hashtable（Version/Build/Path），失败返回 $null
function Probe-OfficeCom([string]$ProgId) {
    try {
        $app = New-Object -ComObject $ProgId
        $info = @{}
        try { $info['Version'] = $app.Version } catch { $info['Version'] = '（不可读）' }
        try { $info['Build'] = $app.Build } catch { }
        try { $info['Path'] = $app.Path } catch { }
        try { $app.Quit() } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch { }
        return $info
    } catch { return $null }
}

# ---------- 打开报告 ----------
Add-Line ('JYYJ助手 一键诊断报告（双宿主版：Microsoft Excel + WPS 表格）')
Add-Line ('生成时间：' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
Add-Line '用途：发给技术支持用于定位安装 / 加载问题。本报告只收集环境信息，不上传任何数据。'

# ---------- [1] 系统信息 ----------
Add-Section '1. 系统信息'
$osName = Add-Collect '操作系统' {
    $k = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop
    ($k.ProductName) + ' (Build ' + $k.CurrentBuildNumber + ', 版本 ' + $k.DisplayVersion + ')'
}
Add-Collect '系统位数' { [System.Environment]::Is64BitOperatingSystem.ToString().Replace('True','64 位').Replace('False','32 位') }
Add-Collect '桌面路径' { [Environment]::GetFolderPath('Desktop') }
Add-Collect '临时目录' { $env:TEMP }
Add-Collect '脚本目录' { $PSScriptRoot }

# ---------- [2] Microsoft Excel ----------
Add-Section '2. Microsoft Excel'
# 运行中的 Excel：mscoree COM shim 会把加载项注册回写为当前实例已加载的旧路径，
# 导致"重装也无法修复"。这是真机实测确认的根因形态，必须检测并提示。
$xlProcs = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
if ($xlProcs.Count -gt 0) {
    Add-Line ('Excel 进程：正在运行（' + $xlProcs.Count + ' 个）')
    Add-Issue '检测到 Excel 正在运行：运行中的 Excel 会把加载项注册回写为它当前加载的旧版本路径，导致"重装也无法修复"。请先关闭全部 Excel 窗口（必要时在任务管理器结束 EXCEL.EXE），再重新运行安装包。'
} else {
    Add-Line 'Excel 进程：未运行'
}
$excelPath = $null
$xlInfo = Probe-OfficeCom 'Excel.Application'
if ($xlInfo) {
    $excelVersion = $xlInfo['Version']
    $excelPath = $xlInfo['Path']
    Add-Line ('Excel COM 探测：成功（Version=' + $excelVersion + $(if ($xlInfo.ContainsKey('Build')) { ' Build=' + $xlInfo['Build'] }) + '）')
    Add-Line ('Excel 安装路径：' + $excelPath)
} else {
    Add-Line 'Excel COM 探测：失败（Excel 可能未安装、被安全软件拦截或处于受限环境）'
    Add-Issue '无法通过 COM 启动 Excel：Excel 可能未安装、被安全软件拦截或处于受限环境。'
    $ap = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe' -ErrorAction SilentlyContinue
    if ($ap -and $ap.'(default)') { $excelPath = $ap.'(default)'; Add-Line ('Excel 可执行文件（注册表）：' + $excelPath) }
    else { Add-Line 'Excel 可执行文件（注册表）：未找到（App Paths 无 excel.exe）'; Add-Issue '注册表未找到 Excel 安装路径，可能未安装 Microsoft Excel。' }
}
$excelArch = $null
if ($excelPath) {
    $exeTarget = $excelPath
    if (-not (Test-Path $excelPath -PathType Leaf)) { $exeTarget = Join-Path $excelPath 'EXCEL.EXE' }
    $excelArch = Get-PEArch $exeTarget
    if ($excelArch) { Add-Line ('Excel 位数（PE 头实测）：' + $excelArch) }
    else { Add-Line ('Excel 位数：EXCEL.EXE 未定位到（' + $exeTarget + '）') }
}
try {
    $c2r = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction Stop
    $channel = $c2r.Channel
    if ([string]::IsNullOrEmpty($channel)) { $channel = '通道未上报' }
    Add-Line ('Office 通道：' + $c2r.Platform + ' / ' + $c2r.VersionToReport + '（' + $channel + '）')
} catch { Add-Line 'Office 通道：未检测到 ClickToRun 配置（可能是 MSI 传统安装版）' }

# ---------- [3] WPS 表格 ----------
Add-Section '3. WPS 表格（金山）'
# 运行中的 WPS：与 Excel 同理，运行中安装会被旧实例覆盖注册
$wpsProcs = @(Get-Process wps,et,wpp,wpsoffice -ErrorAction SilentlyContinue)
if ($wpsProcs.Count -gt 0) {
    $wpsNames = @($wpsProcs | ForEach-Object { $_.ProcessName }) | Select-Object -Unique
    Add-Line ('WPS 进程：正在运行（' + $wpsProcs.Count + ' 个：' + ($wpsNames -join ', ') + '）')
    Add-Issue '检测到 WPS 正在运行：运行中的 WPS 可能维持旧版加载项注册，导致"重装也无法立即生效"。请先完全退出 WPS（含托盘图标），再重新运行安装包。'
} else {
    Add-Line 'WPS 进程：未运行'
}
# et.exe 探测：与 register-auto.ps1 同源（常见安装根下取版本号最高的 office6\et.exe）
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
if (-not $etPath) {
    Add-Line 'WPS 表格：未检测到安装（常见安装根下无 et.exe）'
    Add-Line '说明：只使用 Microsoft Excel 时本节可忽略；若你确定装了 WPS，请把本报告发给技术支持。'
} else {
    Add-Line ('WPS 表格安装（et.exe）：' + $etPath)
    Add-Collect 'WPS 文件版本' { (Get-Item $etPath -ErrorAction Stop).VersionInfo.FileVersion }
    $wpsArch = Get-PEArch $etPath
    if ($wpsArch) { Add-Line ('WPS 位数（PE 头实测）：' + $wpsArch) }

    # Kingsoft ET\Addins 键 + AddinsWL 白名单：register-auto 双视图都写，诊断按双视图核对
    # （32 位 Windows 没有 64 位视图，跳过避免误报"注册不完整"）
    $progIds = @('ExcelAiAssistant.AddIn', 'ExcelAiAssistant.TaskPaneHost', 'ExcelAiAssistant.DeepSeekPaneHost')
    $wpsRegViews = 0
    $viewSpecs = @(
        @([Microsoft.Win32.RegistryView]::Registry64, '64位视图'),
        @([Microsoft.Win32.RegistryView]::Registry32, '32位视图'))
    if (-not [Environment]::Is64BitOperatingSystem) { $viewSpecs = @($viewSpecs[1]) }
    foreach ($viewSpec in $viewSpecs) {
        $wv = $viewSpec[0]; $wvName = $viewSpec[1]
        Add-Line ('【WPS ' + $wvName + '】')
        try {
            $hive = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $wv)
            $viewHasReg = $false
            foreach ($apName in $progIds) {
                $ak = $hive.OpenSubKey('Software\Kingsoft\Office\ET\Addins\' + $apName)
                if ($ak) {
                    $lb = $ak.GetValue('LoadBehavior')
                    $fn = $ak.GetValue('FriendlyName')
                    $wl = $hive.OpenSubKey('Software\Kingsoft\Office\ET\AddinsWL')
                    $inWl = $false; if ($wl) { $inWl = ($wl.GetValue($apName) -ne $null) }
                    Add-Line ('  Kingsoft Addins[' + $apName + '] LoadBehavior=' + $lb + ' FriendlyName=' + $fn + ' 白名单=' + $(if ($inWl) { '在' } else { '缺' }))
                    $viewHasReg = $true
                    if ([string]$lb -eq '2') { Add-Issue ($apName + ' 在 WPS 中 LoadBehavior=2：WPS 上次启动时该加载项出错被自动禁用。请重新运行安装包（先退出 WPS）；若反复出现请提交本报告。') }
                    elseif ([string]$lb -eq '0') { Add-Issue ($apName + ' 在 WPS 中 LoadBehavior=0：被设为手动加载，请重新运行安装包恢复自动加载。') }
                    if (-not $inWl) { Add-Issue ($apName + ' 未加入 WPS 加载项白名单（AddinsWL）：WPS 个人版可能拒绝加载，请重新运行安装包。') }
                } else {
                    Add-Line ('  Kingsoft Addins[' + $apName + ']：未注册')
                }
            }
            if ($viewHasReg) { $wpsRegViews++ }
            $hive.Close()
        } catch { Add-Line ('  本视图检测失败 - ' + $_.Exception.Message) }
    }
    if ($wpsRegViews -eq 0) {
        Add-Issue 'WPS 表格已安装，但加载项未在 WPS 中注册（可能安装时未检测到 WPS、或 WPS 是后装的）。请重新运行安装包（先退出 WPS）。'
    } elseif ($wpsRegViews -eq 1) {
        Add-Issue 'WPS 加载项注册只覆盖了一个注册表视图（正常应双视图都在）：注册可能不完整，建议重新运行安装包。'
    }

    # WPS COM 探测（KET.Application，旧版回落 ET.Application）
    $ketInfo = Probe-OfficeCom 'KET.Application'
    if (-not $ketInfo) { $ketInfo = Probe-OfficeCom 'ET.Application' }
    if ($ketInfo) {
        Add-Line ('WPS COM 探测：成功（' + $ketInfo['Version'] + '）')
    } else {
        Add-Line 'WPS COM 探测：失败（WPS COM 组件可能被限制；加载项走 CLSID 直载，此项失败不一定是致命问题，请结合其余条目判断）'
    }

    # WPS VBA：无可靠远程判定依据（真机实证：个人版无 VBA 文件、无可核查的注册键）——只指引，不断言
    Add-Line 'WPS VBA 环境：无法远程确认（个人版默认不带 VBA）'
    Add-Line '  自查：打开 WPS 表格按 Alt+F11，能弹出"Visual Basic 编辑器"即已具备；'
    Add-Line '  缺失时仅 VBA 类功能不可用，助手其余功能全部正常（补装步骤见《用户手册》常见问题 7）。'
}

# ---------- [4] 系统依赖 ----------
Add-Section '4. 系统依赖'
# .NET 4.8：Release >= 528040 即含 4.8 运行时（读 32 位视图，标准口径）
$netOk = $false
try {
    $rk = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry32)
    $rel = $rk.OpenSubKey('SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full')
    if ($rel) {
        $r = $rel.GetValue('Release')
        if ($null -ne $r) {
            Add-Line ('.NET Framework 4.x：Release=' + $r + '（4.8 需 >= 528040）')
            if ([int]$r -ge 528040) { $netOk = $true } else { Add-Issue ('.NET Framework 版本低于 4.8（Release=' + $r + '），请使用离线完整版安装包自动补齐。') }
        } else { Add-Line '.NET Framework 4.x：Full 键无 Release 值'; Add-Issue '.NET Framework 4.8 未安装或信息不完整，请使用离线完整版安装包。' }
    } else { Add-Line '.NET Framework 4.x：未安装'; Add-Issue '.NET Framework 4.8 未安装，请使用离线完整版安装包自动安装。' }
} catch { Add-Line ('.NET Framework：检测失败 - ' + $_.Exception.Message) }

# WebView2：per-machine 键在 32 位视图（WOW6432Node），per-user 在 HKCU
$wv2Guid = 'F3017226-FE2A-4295-8BDF-00C3A9A7E4C5'
$wv2Key = 'SOFTWARE\Microsoft\EdgeUpdate\Clients\{' + $wv2Guid + '}'
$wv2Ok = $false
foreach ($spec in @(
    @('LocalMachine', [Microsoft.Win32.RegistryView]::Registry64, 'HKLM(64位视图)'),
    @('LocalMachine', [Microsoft.Win32.RegistryView]::Registry32, 'HKLM(32位视图)'),
    @('CurrentUser',  [Microsoft.Win32.RegistryView]::Default,   'HKCU')
)) {
    try {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey($spec[0], $spec[1])
        $k = $base.OpenSubKey($wv2Key)
        $pv = $null; if ($k) { $pv = $k.GetValue('pv') }
        if ($pv -and $pv -ne '0.0.0.0') {
            Add-Line ('WebView2 Runtime：已安装 ' + $pv + '（' + $spec[2] + '）')
            $wv2Ok = $true
            break
        } else {
            Add-Line ('WebView2 Runtime：' + $spec[2] + ' 无有效键值')
        }
    } catch { Add-Line ('WebView2 Runtime：' + $spec[2] + ' 检测失败 - ' + $_.Exception.Message) }
}
if (-not $wv2Ok) { Add-Issue 'WebView2 Runtime 未检测到，请使用离线完整版安装包自动安装（32 位 Excel 需到官方页面下载 x86 版）。' }

# ---------- [5] 加载项注册实况 ----------
Add-Section '5. 加载项 COM 注册实况（核心）'
Add-Line '说明：Excel 与 WPS 共用同一套 COM 类（CLSID），区别只在宿主各自的 Addins 键位置'
Add-Line '      （Excel 在 Software\Microsoft\Office\Excel\Addins，WPS 在 Software\Kingsoft\Office\ET\Addins，见第 3 节）。'
$progIds = @(
    @('ExcelAiAssistant.AddIn',         '8F7A6D91-4C52-45BE-9F55-7D8C2A31E001'),
    @('ExcelAiAssistant.TaskPaneHost',  '8F7A6D91-4C52-45BE-9F55-7D8C2A31E002'),
    @('ExcelAiAssistant.DeepSeekPaneHost','8F7A6D91-4C52-45BE-9F55-7D8C2A31E003')
)
$regFound = $false
# 32 位 Windows 没有 64 位视图，跳过避免整节"检测失败"噪音
$regViewSpecs = @(
    @([Microsoft.Win32.RegistryView]::Registry64, '64位视图'),
    @([Microsoft.Win32.RegistryView]::Registry32, '32位视图'))
if (-not [Environment]::Is64BitOperatingSystem) { $regViewSpecs = @($regViewSpecs[1]) }
foreach ($viewSpec in $regViewSpecs) {
    $view = $viewSpec[0]; $viewName = $viewSpec[1]
    # 本机相关视图 = 与 Excel 位数一致的视图（Addins 键双视图共享，CLSID 按架构各写各的）
    $isRelevantView = ($null -eq $excelArch) -or (($viewName -eq '64位视图') -eq ($excelArch -eq 'x64'))
    Add-Line ('【' + $viewName + '】' + $(if ($isRelevantView) { '' } else { '（本机 Excel 不使用此视图，仅供参考）' }))
    try {
        $hive = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
        foreach ($p in $progIds) {
            $progId = $p[0]; $guid = $p[1]
            $ckQuick = $null
            $clsidReported = $false
            # Addins 键
            $ak = $hive.OpenSubKey('Software\Microsoft\Office\Excel\Addins\' + $progId)
            if ($ak) {
                $lb = $ak.GetValue('LoadBehavior')
                $fn = $ak.GetValue('FriendlyName')
                Add-Line ('  Addins[' + $progId + '] LoadBehavior=' + $lb + ' FriendlyName=' + $fn)
                if ($isRelevantView) { $regFound = $true }
                if ([string]$lb -eq '2') { Add-Issue ($progId + ' LoadBehavior=2：Excel 上次启动时该加载项出错被自动禁用。可在 Excel 文件-选项-加载项-COM加载项 中重新勾选；若反复出现请提交本报告。（注：本诊断工具的 Excel 探测本身也可能触发此状态，重新勾选后再看是否复现）') }
                elseif ([string]$lb -eq '0') { Add-Issue ($progId + ' LoadBehavior=0：加载项被设为手动加载。在 Excel 文件-选项-加载项-COM加载项 中勾选即可恢复开机自动加载。') }
                elseif ([string]$lb -ne '3') { Add-Issue ($progId + ' LoadBehavior=' + $lb + '：非正常状态(3=正常)，建议重新运行安装包。') }
                # Addins 键在但 CLSID 缺失 = Excel 尝试加载必失败弹错（典型报错形态）
                # 仅对本机相关视图判定；另一视图无 CLSID 属正常（register-auto 只写 Excel 架构对应视图）
                $ckQuick = $hive.OpenSubKey('Software\Classes\CLSID\{' + $guid + '}\InprocServer32')
                if (-not $ckQuick) {
                    if ($isRelevantView) {
                        Add-Line ('  CLSID[' + $guid.Substring(0,8) + '...]：未注册（但 Addins 键在）')
                        Add-Issue ($progId + ' 的 COM 类（CLSID）未注册而 Addins 键存在：Excel 加载时必然报错。请重新运行安装包。')
                        $clsidReported = $true
                    } else {
                        Add-Line ('  CLSID[' + $guid.Substring(0,8) + '...]：本视图未注册（正常，本机不使用）')
                        $clsidReported = $true
                    }
                }
            } else {
                Add-Line ('  Addins[' + $progId + ']：未注册')
            }
            # CLSID InprocServer32（若上方快速检查已判定缺失，此处跳过避免重复输出）
            $ck = $hive.OpenSubKey('Software\Classes\CLSID\{' + $guid + '}\InprocServer32')
            if ($ck) {
                $cb = $ck.GetValue('CodeBase')
                $cbPath = $null
                if ($cb -and $cb.StartsWith('file:///')) { $cbPath = [System.Uri]::UnescapeDataString($cb.Substring(8)).Replace('/', '\') }
                Add-Line ('  CLSID[' + $guid.Substring(0,8) + '...] CodeBase=' + $cb)
                if ($cbPath) {
                    if (Test-Path $cbPath) {
                        try {
                            $asmVer = [System.Reflection.AssemblyName]::GetAssemblyName($cbPath).Version.ToString()
                            Add-Line ('    注册指向的 DLL：存在（程序集版本 ' + $asmVer + '）')
                        } catch { Add-Line ('    注册指向的 DLL：存在但读取版本失败 - ' + $_.Exception.Message) }
                        if ($cbPath -notlike ($env:LOCALAPPDATA + '\XaiAssistant\*')) {
                            Add-Issue ('CLSID 注册指向了安装目录以外的位置：' + $cbPath + '（注册可能被其它安装/开发环境覆盖），建议重新运行安装包。')
                        }
                    } else {
                        Add-Line ('    注册指向的 DLL：不存在！' + $cbPath)
                        Add-Issue ('CLSID 注册指向的 DLL 文件缺失：' + $cbPath + '，安装文件不完整或被清理，建议重新安装。')
                    }
                }
            } elseif (-not $clsidReported) {
                Add-Line ('  CLSID[' + $guid.Substring(0,8) + '...]：未注册')
            }
        }
    } catch { Add-Line ('  本视图检测失败 - ' + $_.Exception.Message) }
}
if (-not $regFound) {
    Add-Issue '所有 Excel 加载项注册键均未找到：加载项未成功注册。请重新运行安装包；若安装过程报错请截图并提交报告。'
}

# Excel 禁用列表（Resiliency）
try {
    $officeRoot = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, [Microsoft.Win32.RegistryView]::Default).OpenSubKey('Software\Microsoft\Office')
    if ($officeRoot) {
        foreach ($ver in $officeRoot.GetSubKeyNames()) {
            $dk = $officeRoot.OpenSubKey($ver + '\Excel\Resiliency\DisabledItems')
            if ($dk -and @($dk.GetValueNames()).Count -gt 0) {
                Add-Line ('Excel 禁用列表：Office ' + $ver + ' 有 ' + @($dk.GetValueNames()).Count + ' 项被禁用记录')
                Add-Issue ('Excel 的禁用列表（Office ' + $ver + '）中存在记录，加载项可能被 Excel 自动禁用。可删除该列表后在 COM 加载项中重新勾选。')
            }
        }
    }
} catch { Add-Line ('Excel 禁用列表：检测失败 - ' + $_.Exception.Message) }

# 历史版本（<=2.19.4.0）误写的无花括号垃圾键检测（升级路径残留）
try {
    $legacyClsidRoot = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, [Microsoft.Win32.RegistryView]::Default).OpenSubKey('Software\Classes\CLSID')
    if ($legacyClsidRoot) {
        $legacyFound = @()
        foreach ($p2 in $progIds) {
            if ($legacyClsidRoot.OpenSubKey($p2[1])) { $legacyFound += $p2[0] }
        }
        if ($legacyFound.Count -gt 0) {
            Add-Line ('历史垃圾键：发现 ' + $legacyFound.Count + ' 个旧版本误写的无效注册键')
            Add-Issue '检测到旧版本（2.19.4.0 及更早）误写的无效注册键残留：这是导致此前"加载报错"的原因之一。重新运行 2.19.5.0 及以上版本安装包（关闭 Excel 后）即可自动清理。'
        }
        $legacyClsidRoot.Close()
    }
} catch { Add-Line ('历史垃圾键：检测失败 - ' + $_.Exception.Message) }

# ---------- [5.5] Excel 禁用项（Resiliency）检测与自动修复 ----------
Add-Section '5.5 Excel 禁用项（崩溃自动禁用）检测与修复'
Add-Line '说明：Excel 崩溃/强退后会把 COM 加载项记入"禁用项"，导致功能区 JYYJ 助手选项卡消失。'
Add-Line '      此处自动清除禁用标记（仅解除 Excel 的崩溃禁用记录，不影响系统与数据），'
Add-Line '      清除后重启 Excel 即可恢复选项卡。'
try {
    $resiCleared = 0
    foreach ($v in @('15.0','16.0')) {
        $rk = 'HKCU:\Software\Microsoft\Office' + $v + '\Excel\Resiliency\DisabledItems'
        $it = Get-Item $rk -ErrorAction SilentlyContinue
        if ($it -and $it.ValueCount -gt 0) {
            foreach ($n in $it.GetValueNames()) {
                Remove-ItemProperty -Path $rk -Name $n -Force -ErrorAction SilentlyContinue
                $resiCleared++
            }
            Add-Line ('  Office ' + $v + '：清除禁用标记 ' + $resiCleared + ' 项')
        } else {
            Add-Line ('  Office ' + $v + '：无禁用标记')
        }
    }
    if ($resiCleared -gt 0) {
        Add-Line ('  ✔ 已自动清除 ' + $resiCleared + ' 项禁用标记，请重启 Excel 使其重新加载 JYYJ 助手。')
    } else {
        Add-Line '  ✔ 无需处理'
    }
} catch { Add-Line ('  禁用项检测失败：' + $_.Exception.Message) }

# ---------- [6] 安装文件完整性 ----------
Add-Section '6. 安装文件完整性'
$installRoot = Join-Path $env:LOCALAPPDATA 'XaiAssistant'
Add-Line ('安装根目录：' + $installRoot + '（存在=' + (Test-Path $installRoot) + '）')
foreach ($arch in @('x64', 'x86')) {
    $dll = Join-Path (Join-Path $installRoot ('app\' + $arch)) 'ExcelAiAssistant.Addin.dll'
    if (Test-Path $dll) {
        try {
            $v = [System.Reflection.AssemblyName]::GetAssemblyName($dll).Version.ToString()
            Add-Line ('  app\' + $arch + '\ExcelAiAssistant.Addin.dll：存在（版本 ' + $v + '）')
        } catch { Add-Line ('  app\' + $arch + '\ExcelAiAssistant.Addin.dll：存在但读取版本失败') }
    } else {
        Add-Line ('  app\' + $arch + '\ExcelAiAssistant.Addin.dll：缺失')
        Add-Issue ('安装目录缺少 ' + $arch + ' 架构加载项 DLL，安装不完整，建议重新安装。')
    }
}
foreach ($f in @('register-auto.ps1', 'unregister-all.ps1', 'diagnose.ps1', '安装说明.txt', '用户手册.html', '一键诊断.cmd')) {
    $fp = Join-Path $installRoot $f
    if (Test-Path $fp) { Add-Line ('  ' + $f + '：存在') }
    else { Add-Line ('  ' + $f + '：缺失'); Add-Issue ('支持文件缺失：' + $f + '（不影响主功能，但影响诊断/卸载/查阅），建议重新安装。') }
}
if ($excelArch -and (Test-Path (Join-Path (Join-Path $installRoot ('app\' + $excelArch)) 'ExcelAiAssistant.Addin.dll'))) {
    Add-Line ('Excel 位数（' + $excelArch + '）与对应架构目录：匹配')
}
if ($etPath -and $wpsArch -and (Test-Path (Join-Path (Join-Path $installRoot ('app\' + $wpsArch)) 'ExcelAiAssistant.Addin.dll'))) {
    Add-Line ('WPS 位数（' + $wpsArch + '）与对应架构目录：匹配')
}

# ---------- [7] 运行日志 ----------
Add-Section '7. 加载项运行日志（最后 60 行）'
$logPath = Join-Path $env:TEMP 'ExcelAiAssistant.log'
if (Test-Path $logPath) {
    Add-Line ('日志文件：' + $logPath + '（最后写入 ' + (Get-Item $logPath).LastWriteTime + '，大小 ' + (Get-Item $logPath).Length + ' 字节）')
    Add-Line ''
    try {
        Get-Content $logPath -ErrorAction Stop | Select-Object -Last 60 | ForEach-Object { Add-Line ('  ' + $_) }
    } catch { Add-Line ('  日志读取失败 - ' + $_.Exception.Message) }
} else {
    Add-Line ('日志文件：不存在（' + $logPath + '）')
    Add-Line '说明：日志不存在通常意味着加载项代码从未被 Excel/WPS 执行过（注册缺失或加载链路在更早阶段失败）。'
}
$rotated = @(Get-ChildItem (Join-Path $env:TEMP 'ExcelAiAssistant-*.log') -ErrorAction SilentlyContinue)
if ($rotated.Count -gt 0) {
    $rotNames = @($rotated | ForEach-Object { $_.Name })
    Add-Line ('历史日志（轮转备份，共 ' + $rotNames.Count + ' 个）：' + ($rotNames -join ', '))
}

# ---------- [8] 体检结论 ----------
Add-Section '8. 体检结论'
# 去重：双视图 × 多 ProgID 会产生同文案问题，只保留唯一项
$uniqueIssues = @($script:Issues | Select-Object -Unique)
if ($uniqueIssues.Count -eq 0) {
    Add-Line '未发现明显问题。'
    Add-Line '若仍遇到报错，请把本报告连同报错窗口截图一起发给技术支持。'
} else {
    Add-Line ('发现 ' + $uniqueIssues.Count + ' 个问题，按建议顺序处理：')
    Add-Line ''
    $i = 0
    foreach ($iss in $uniqueIssues) {
        $i++
        Add-Line ([string]$i + '. ' + $iss)
        Add-Line ''
    }
    Add-Line '通用自救三步：'
    Add-Line '  ① 关闭全部 Excel 与 WPS 窗口（任务管理器确认无 EXCEL.EXE / et.exe / wps.exe），再重新双击安装包覆盖安装；'
    Add-Line '  ② 仍不行 → 用"离线完整版"安装包；'
    Add-Line '  ③ 仍不行 → 把本报告完整复制发给技术支持。'
}

# ---------- 输出报告 ----------
$desktop = [Environment]::GetFolderPath('Desktop')
if (-not $desktop) { $desktop = $env:USERPROFILE }
$report = Join-Path $desktop ('JYYJ诊断报告-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.txt')
try {
    $script:Out | Out-File -FilePath $report -Encoding UTF8
} catch {
    $report = Join-Path $env:TEMP ('JYYJ诊断报告-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.txt')
    $script:Out | Out-File -FilePath $report -Encoding UTF8
}

Write-Host ''
Write-Host '========================================'
if ($uniqueIssues.Count -eq 0) {
    Write-Host '诊断完成：未发现明显问题。'
} else {
    Write-Host ('诊断完成：发现 ' + $uniqueIssues.Count + ' 个问题，详见报告。')
}
Write-Host ('报告已保存：' + $report)
Write-Host '========================================'
if (-not $NoOpen) {
    try { Start-Process notepad.exe $report | Out-Null } catch {}
}
exit 0
