param([string]$Guid, [string]$ProgId, [string]$FullType, [string]$AssemblyPath)
$ErrorActionPreference = "Stop"
$installedAssembly = $AssemblyPath
$assemblyFullName = [System.Reflection.AssemblyName]::GetAssemblyName($installedAssembly).FullName
$codeBase = (New-Object System.Uri($installedAssembly)).AbsoluteUri

function Register-ComClassPerUser([string]$Guid, [string]$ProgId, [string]$FullType, [string]$AssemblyFullName, [string]$CodeBase) {
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
    Write-Output "reg-done $Guid"
}

Register-ComClassPerUser $Guid $ProgId $FullType $assemblyFullName $codeBase
$i="HKCU:\Software\Classes\CLSID\{$Guid}\InprocServer32"
Write-Output ("in-isol exists=" + (Test-Path $i))
if(Test-Path $i){ (Get-ItemProperty $i).Assembly }