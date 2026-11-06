# JYYJ助手 合并安装器卸载清理：显式遍历 64/32 两个注册表视图，删除 .NET COM 类与 Excel Addins 键。
# 与进程位数无关，卸载后 32 位与 64 位 Excel 均不再加载该加载项。
$ErrorActionPreference = 'SilentlyContinue'
$guids = @(
    '8F7A6D91-4C52-45BE-9F55-7D8C2A31E001',
    '8F7A6D91-4C52-45BE-9F55-7D8C2A31E002',
    '8F7A6D91-4C52-45BE-9F55-7D8C2A31E003'
)
$progids = @('ExcelAiAssistant.AddIn', 'ExcelAiAssistant.TaskPaneHost', 'ExcelAiAssistant.DeepSeekPaneHost')

foreach ($regView in @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)) {
    $root = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $regView)
    foreach ($g in $guids) { $root.DeleteSubKeyTree("Software\Classes\CLSID\{$g}", $false) }
    foreach ($p in $progids) {
        $root.DeleteSubKeyTree("Software\Classes\$p", $false)
        $root.DeleteSubKeyTree("Software\Microsoft\Office\Excel\Addins\$p", $false)
        # WPS（M1）：Kingsoft ET Addins 键 + AddinsWL 白名单值，双视图清理
        $root.DeleteSubKeyTree("Software\Kingsoft\Office\ET\Addins\$p", $false)
        $wl = $root.OpenSubKey('Software\Kingsoft\Office\ET\AddinsWL', $true)
        if ($wl) {
            try { $wl.DeleteValue($p, $false) } catch { }
            $wl.Close()
        }
    }
    $root.Close()
}
# 额外清理历史版本（<=2.19.4.0）误写的无花括号垃圾键（与视图无关，默认视图删一次即可）
$legacy = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, [Microsoft.Win32.RegistryView]::Default)
$legacyClsid = $legacy.OpenSubKey('Software\Classes\CLSID', $true)
if ($legacyClsid) {
    foreach ($g in $guids) {
        try { $legacyClsid.DeleteSubKeyTree($g, $false) } catch { }
    }
    $legacyClsid.Close()
}
$legacy.Close()
Write-Output 'UNREGISTERED'