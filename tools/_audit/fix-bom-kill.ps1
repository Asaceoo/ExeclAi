$ErrorActionPreference = 'SilentlyContinue'
$p = 'd:\execl1.0\tools\VerifySpotlight\Probe-SelectPerf.ps1'
$bytes = [System.IO.File]::ReadAllBytes($p)
$hasBom = ($bytes[0] -eq 239 -and $bytes[1] -eq 187 -and $bytes[2] -eq 191)
if (-not $hasBom) {
    $c = [System.IO.File]::ReadAllText($p, [System.Text.Encoding]::UTF8)
    $enc = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($p, $c, $enc)
    Write-Output 're-added UTF-8 BOM'
} else {
    Write-Output 'BOM already present'
}
Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object {
    if ($_.MainWindowTitle -eq '') {
        try { Stop-Process -Id $_.Id -Force; Write-Output ("killed zombie Excel PID=" + $_.Id) } catch { Write-Output ("cannot kill PID=" + $_.Id) }
    }
}
