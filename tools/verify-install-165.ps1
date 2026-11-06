# 验证已安装的 XAI 加载项（1.6.5.0）：正确枚举 COMAddIns 中的 XAI 条目。
$ErrorActionPreference = 'Stop'
$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    foreach ($a in $excel.ComAddIns) {
        $prog = [string]$a.ProgId
        if ($prog -match 'ExcelAiAssistant') {
            Write-Host ("FOUND XAI COMAddIn: ProgId={0} Connect={1} Description={2}" -f $prog, $a.Connect, $a.Description)
            $a.Connect = $true
            Write-Host ("  -> Connect=True ok, Connect now={0}" -f $a.Connect)
        }
    }
} finally {
    if ($excel) { try { $excel.DisplayAlerts = $true; $excel.Quit() } catch {} }
    try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null } catch {}
}
Write-Host 'DONE'