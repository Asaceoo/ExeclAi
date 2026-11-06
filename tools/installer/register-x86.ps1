# JYYJ助手 注册（由 x86 安装器以 32 位 PowerShell 静默执行，写 32 位注册表视图，供 32 位 Excel 加载）
# 32 位进程内 .NET Registry API 默认写 32 位重定向视图，与 32 位 Excel 读取位置一致。
# 逻辑与 install-xai.ps1 的 Register-ComClassPerUser 保持一致
$ErrorActionPreference = 'Stop'
$dir = $PSScriptRoot
$assembly = Join-Path $dir 'ExcelAiAssistant.Addin.dll'
if (-not (Test-Path $assembly)) { Write-Output ('MISSING ' + $assembly); exit 1 }
$asmFull = [Reflection.AssemblyName]::GetAssemblyName($assembly).FullName
$version = [Reflection.AssemblyName]::GetAssemblyName($assembly).Version.ToString()
$codebase = (New-Object System.Uri($assembly)).AbsoluteUri
$HKCU = [Microsoft.Win32.Registry]::CurrentUser
$catId = '{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}'

function Register-ComClass([string]$Guid, [string]$ProgId, [string]$FullType) {
    $cls = $HKCU.CreateSubKey("Software\Classes\CLSID\$Guid")
    $cls.SetValue('', $FullType)
    $cls.SetValue('Assembly', $asmFull)
    $cls.SetValue('Class', $FullType)
    $cls.SetValue('ProgID', $ProgId)
    $cls.Close()
    $HKCU.CreateSubKey("Software\Classes\CLSID\$Guid\ProgId").SetValue('', $ProgId)
    $HKCU.CreateSubKey("Software\Classes\CLSID\$Guid\Implemented Categories\$catId")

    $in = $HKCU.CreateSubKey("Software\Classes\CLSID\$Guid\InprocServer32")
    $in.SetValue('', 'mscoree.dll')
    $in.SetValue('ThreadingModel', 'Both')
    $in.SetValue('Class', $FullType)
    $in.SetValue('Assembly', $asmFull)
    $in.SetValue('CodeBase', $codebase)
    $in.SetValue('RuntimeVersion', 'v4.0.30319')
    $in.Close()

    $inV = $HKCU.CreateSubKey("Software\Classes\CLSID\$Guid\InprocServer32\$version")
    $inV.SetValue('Class', $FullType)
    $inV.SetValue('Assembly', $asmFull)
    $inV.SetValue('RuntimeVersion', 'v4.0.30319')
    $inV.SetValue('CodeBase', $codebase)
    $inV.Close()

    $HKCU.CreateSubKey("Software\Classes\$ProgId").SetValue('', $FullType)
    $HKCU.CreateSubKey("Software\Classes\$ProgId\CLSID").SetValue('', "{$Guid}")

    $add = $HKCU.CreateSubKey("Software\Microsoft\Office\Excel\Addins\$ProgId")
    $add.SetValue('Description', 'OpenAI-compatible AI assistant for Excel.')
    $add.SetValue('FriendlyName', 'JYYJ助手')
    $add.SetValue('LoadBehavior', 3, [Microsoft.Win32.RegistryValueKind]::DWord)
    $add.Close()
}

Register-ComClass '8F7A6D91-4C52-45BE-9F55-7D8C2A31E001' 'ExcelAiAssistant.AddIn' 'ExcelAiAssistant.Addin.ExcelAiAddIn'
Register-ComClass '8F7A6D91-4C52-45BE-9F55-7D8C2A31E002' 'ExcelAiAssistant.TaskPaneHost' 'ExcelAiAssistant.Addin.ExcelAiTaskPaneHost'
Register-ComClass '8F7A6D91-4C52-45BE-9F55-7D8C2A31E003' 'ExcelAiAssistant.DeepSeekPaneHost' 'ExcelAiAssistant.Addin.DeepSeekTaskPaneHost'

Write-Output ("REGISTERED version=" + $version)
Write-Output ("CODEDIR=" + $dir)
Write-Output ("CODEBASE=" + $codebase)