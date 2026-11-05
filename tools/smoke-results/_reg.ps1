param([string]$Guid, [string]$ProgId, [string]$FullType, [string]$AssemblyPath)
$ErrorActionPreference = "Stop"
$installedAssembly = $AssemblyPath
$assemblyFullName = [System.Reflection.AssemblyName]::GetAssemblyName($installedAssembly).FullName
$codeBase = (New-Object System.Uri($installedAssembly)).AbsoluteUri
function Register-ComClassPerUser([string]$Guid, [string]$ProgId, [string]$FullType, [string]$AssemblyFullName, [string]$CodeBase) {
    $version = ([System.Reflection.AssemblyName]::GetAssemblyName($installedAssembly).Version).ToString()
    $versionKey = $version
    $clsBase = "HKCU:\Software\Classes\CLSID\$Guid"
    New-Item -Path $clsBase -Force | Out-Null
    Set-ItemProperty -Path $clsBase -Name "(default)" -Value $FullType
    Set-ItemProperty -Path $clsBase -Name "Assembly" -Value $AssemblyFullName
    Set-ItemProperty -Path $clsBase -Name "Class" -Value $FullType
    Set-ItemProperty -Path $clsBase -Name "ProgID" -Value $ProgId
    $clsProgId = "$clsBase\ProgId"
    New-Item -Path $clsProgId -Force | Out-Null
    Set-ItemProperty -Path $clsProgId -Name "(default)" -Value $ProgId
    New-Item -Path "$clsBase\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}" -Force | Out-Null
    $inproc = "$clsBase\InprocServer32"
    New-Item -Path $inproc -Force | Out-Null
    New-ItemProperty -Path $inproc -Name "(default)" -Value "mscoree.dll" -Force | Out-Null
    New-ItemProperty -Path $inproc -Name "ThreadingModel" -Value "Both" -Force | Out-Null
    New-ItemProperty -Path $inproc -Name "Class" -Value $FullType -Force | Out-Null
    New-ItemProperty -Path $inproc -Name "Assembly" -Value $AssemblyFullName -Force | Out-Null
    New-ItemProperty -Path $inproc -Name "CodeBase" -Value $CodeBase -Force | Out-Null
    New-ItemProperty -Path $inproc -Name "RuntimeVersion" -Value "v4.0.30319" -Force | Out-Null
    $inprocVer = "$inproc\$versionKey"
    New-Item -Path $inprocVer -Force | Out-Null
    New-ItemProperty -Path $inprocVer -Name "Class" -Value $FullType -Force | Out-Null
    New-ItemProperty -Path $inprocVer -Name "Assembly" -Value $AssemblyFullName -Force | Out-Null
    New-ItemProperty -Path $inprocVer -Name "RuntimeVersion" -Value "v4.0.30319" -Force | Out-Null
    New-ItemProperty -Path $inprocVer -Name "CodeBase" -Value $CodeBase -Force | Out-Null
    $progBase = "HKCU:\Software\Classes\$ProgId"
    New-Item -Path $progBase -Force | Out-Null
    Set-ItemProperty -Path $progBase -Name "(default)" -Value $FullType
    New-Item -Path "$progBase\CLSID" -Force | Out-Null
    Set-ItemProperty -Path "$progBase\CLSID" -Name "(default)" -Value ("{" + $Guid + "}")
}
Write-Output ("AssemblyFullName=" + $assemblyFullName)
Write-Output ("CodeBase=" + $codeBase)
Register-ComClassPerUser $Guid $ProgId $FullType $assemblyFullName $codeBase
Write-Output "registered"
$k="HKCU:\Software\Classes\CLSID\$Guid\InprocServer32"
Write-Output ("exists=" + (Test-Path $k))
if(Test-Path $k){ Get-ItemProperty $k | Select-Object Class,Assembly,CodeBase | Format-List }