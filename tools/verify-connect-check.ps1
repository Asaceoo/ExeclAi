# 真机验证：XAI 加载项 Connect=True（结果写文件，避免 stdout 捕获丢失）
$out = 'C:\Users\iamly\AppData\Local\Temp\xai-connect-check.txt'
$lines = New-Object System.Collections.Generic.List[string]
$excel = $null
try {
    $dll = 'C:\Users\iamly\AppData\Local\XaiAssistant\app\x64\ExcelAiAssistant.Addin.dll'
    $asm = [System.Reflection.AssemblyName]::GetAssemblyName($dll)
    $lines.Add(('DLL_VERSION=' + $asm.Version.ToString()))

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $found = $false
    foreach ($a in $excel.ComAddIns) {
        $prog = [string]$a.ProgId
        if ($prog -match 'ExcelAiAssistant') {
            $found = $true
            $before = [bool]$a.Connect
            $a.Connect = $true
            $lines.Add(('CONNECT_PROGID=' + $prog + ' BEFORE=' + $before + ' AFTER=' + $a.Connect))
        }
    }
    if (-not $found) { $lines.Add('CONNECT_NOT_FOUND') }
} catch {
    $lines.Add('CONNECT_ERROR=' + $_.Exception.Message)
} finally {
    if ($excel) { try { $excel.DisplayAlerts = $true; $excel.Quit() } catch {} }
    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
}
[System.IO.File]::WriteAllLines('C:\Users\iamly\AppData\Local\Temp\xai-connect-check.txt', @($lines))
Write-Host 'WROTE'