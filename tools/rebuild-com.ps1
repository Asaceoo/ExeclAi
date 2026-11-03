$asmFull = "ExcelAiAssistant.Addin, Version=2.8.0.0, Culture=neutral, PublicKeyToken=null"
$cb = "file:///C:/Users/iamly/AppData/Local/XaiAssistant/app/x64/ExcelAiAssistant.Addin.dll"

function Add-CompleteCom {
    param([string]$Guid, [string]$ProgId, [string]$TypeName)
    $b = "HKCU:\Software\Classes\CLSID\$Guid"
    New-Item $b -Force | Out-Null
    Set-ItemProperty $b -Name "(default)" -Value $TypeName
    Set-ItemProperty $b -Name "Assembly" -Value $asmFull
    Set-ItemProperty $b -Name "Class" -Value $TypeName
    Set-ItemProperty $b -Name "ProgID" -Value $ProgId
    New-Item "$b\ProgId" -Force | Out-Null
    Set-ItemProperty "$b\ProgId" -Name "(default)" -Value $ProgId
    New-Item "$b\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}" -Force | Out-Null
    $in = "$b\InprocServer32"
    New-Item $in -Force | Out-Null
    New-ItemProperty $in -Name "(default)" -Value "mscoree.dll" -Force | Out-Null
    New-ItemProperty $in -Name "ThreadingModel" -Value "Both" -Force | Out-Null
    New-ItemProperty $in -Name "Class" -Value $TypeName -Force | Out-Null
    New-ItemProperty $in -Name "Assembly" -Value $asmFull -Force | Out-Null
    New-ItemProperty $in -Name "CodeBase" -Value $cb -Force | Out-Null
    New-ItemProperty $in -Name "RuntimeVersion" -Value "v4.0.30319" -Force | Out-Null
    $iv = "$in\2.8.0.0"
    New-Item $iv -Force | Out-Null
    New-ItemProperty $iv -Name "Class" -Value $TypeName -Force | Out-Null
    New-ItemProperty $iv -Name "Assembly" -Value $asmFull -Force | Out-Null
    New-ItemProperty $iv -Name "CodeBase" -Value $cb -Force | Out-Null
    New-ItemProperty $iv -Name "RuntimeVersion" -Value "v4.0.30319" -Force | Out-Null
    $p = "HKCU:\Software\Classes\$ProgId"
    New-Item $p -Force | Out-Null
    Set-ItemProperty $p -Name "(default)" -Value $TypeName
    New-Item "$p\CLSID" -Force | Out-Null
    Set-ItemProperty "$p\CLSID" -Name "(default)" -Value ("{" + $Guid + "}")
}

Add-CompleteCom "8F7A6D91-4C52-45BE-9F55-7D8C2A31E001" "ExcelAiAssistant.AddIn" "ExcelAiAssistant.Addin.ExcelAiAddIn"
Add-CompleteCom "8F7A6D91-4C52-45BE-9F55-7D8C2A31E002" "ExcelAiAssistant.TaskPaneHost" "ExcelAiAssistant.Addin.ExcelAiTaskPaneHost"
Add-CompleteCom "8F7A6D91-4C52-45BE-9F55-7D8C2A31E003" "ExcelAiAssistant.DeepSeekPaneHost" "ExcelAiAssistant.Addin.DeepSeekTaskPaneHost"

# 重置 LoadBehavior 为 3（自动加载）
$ak = "HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelAiAssistant.AddIn"
New-ItemProperty $ak -Name "LoadBehavior" -Value 3 -PropertyType DWord -Force | Out-Null
Write-Host "LoadBehavior reset to 3"
Write-Host "Full HKCU COM registration (braces ProgID) rebuilt."