# 真机验证：v1.9.0 加载项 Connect=True + 版本 + 内嵌 index.html 含新功能标记
$out = 'C:\Users\iamly\AppData\Local\Temp\xai-verify-190.txt'
$lines = New-Object System.Collections.Generic.List[string]
$excel = $null
try {
    $dir = 'C:\Users\iamly\AppData\Local\XaiAssistant\app\x64'
    $dll = Join-Path $dir 'ExcelAiAssistant.Addin.dll'
    $asmName = [System.Reflection.AssemblyName]::GetAssemblyName($dll)
    $lines.Add(('DLL_VERSION=' + $asmName.Version.ToString()))

    # 解析依赖，读取内嵌 index.html
    $handler = [System.ResolveEventHandler]{
        param($s, $e)
        $name = (New-Object System.Reflection.AssemblyName $e.Name).Name
        $f = Join-Path $dir ($name + '.dll')
        if (Test-Path $f) { return [System.Reflection.Assembly]::LoadFrom($f) }
        return $null
    }
    [System.AppDomain]::CurrentDomain.add_AssemblyResolve($handler)
    $asm = [System.Reflection.Assembly]::LoadFrom($dll)
    $resName = $asm.GetManifestResourceNames() | Where-Object { $_ -match 'index\.html' } | Select-Object -First 1
    if ($resName) {
        $lines.Add(('RESOURCE=' + $resName))
        $st = $asm.GetManifestResourceStream($resName)
        $sr = New-Object System.IO.StreamReader($st)
        $html = $sr.ReadToEnd()
        $sr.Close()
        foreach ($marker in @('id="view-ds"','id="dsFrame"','id="dsZoomIn"','id="formulaInsertButton"','mini-copy','id="atList"')) {
            $lines.Add(('MARKER ' + $marker + ' = ' + $html.Contains($marker)))
        }
    } else {
        $lines.Add('RESOURCE_NOT_FOUND')
    }

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $found = $false
    foreach ($a in $excel.ComAddIns) {
        $prog = [string]$a.ProgId
        if ($prog -match 'ExcelAiAssistant') {
            $found = $true
            $a.Connect = $true
            $lines.Add(('CONNECT_PROGID=' + $prog + ' CONNECT=' + $a.Connect))
        }
    }
    if (-not $found) { $lines.Add('CONNECT_NOT_FOUND') }
} catch {
    $lines.Add('ERROR=' + $_.Exception.ToString())
} finally {
    if ($excel) { try { $excel.DisplayAlerts = $true; $excel.Quit() } catch {} }
    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
    $p = Get-Process EXCEL -ErrorAction SilentlyContinue
    if ($p) { $p | Stop-Process -Force }
}
[System.IO.File]::WriteAllLines('C:\Users\iamly\AppData\Local\Temp\xai-verify-190.txt', @($lines))
Write-Host 'WROTE'