# JYYJ助手 注销注册项（卸载时由 64 位 PowerShell 静默执行，删除 64 位视图中的 COM 类与 Addins 键）
$ErrorActionPreference = 'SilentlyContinue'
$HKCU = [Microsoft.Win32.Registry]::CurrentUser
$guids = @(
    '8F7A6D91-4C52-45BE-9F55-7D8C2A31E001',
    '8F7A6D91-4C52-45BE-9F55-7D8C2A31E002',
    '8F7A6D91-4C52-45BE-9F55-7D8C2A31E003'
)
$progids = @('ExcelAiAssistant.AddIn', 'ExcelAiAssistant.TaskPaneHost', 'ExcelAiAssistant.DeepSeekPaneHost')

foreach ($g in $guids) { $HKCU.DeleteSubKeyTree("Software\Classes\CLSID\{$g}", $false) }
foreach ($p in $progids) {
    $HKCU.DeleteSubKeyTree("Software\Classes\$p", $false)
    $HKCU.DeleteSubKeyTree("Software\Microsoft\Office\Excel\Addins\$p", $false)
}
Write-Output 'UNREGISTERED'