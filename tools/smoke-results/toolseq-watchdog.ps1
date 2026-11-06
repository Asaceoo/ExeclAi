# Watchdog for toolseq modes (FULL/NOAW/NOSU/NOLINES/MIN).
param([int]$TimeoutSec = 90)
$ErrorActionPreference = "Stop"
$ps = "C:\WINDOWS\System32\WindowsPowerShell\v1.0\powershell.exe"
$script = "d:\execl1.0\tools\smoke-results\toolseq-run.ps1"
$modes = @("FULL","NOAW","NOSU","NOLINES","MIN")

foreach ($m in $modes) {
    $resultFile = "d:\execl1.0\tools\smoke-results\toolseq-$m.txt"
    Remove-Item $resultFile -ErrorAction SilentlyContinue
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $ps
    $psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$script`" -Mode $m"
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    $null = $proc.Start()
    if (-not $proc.WaitForExit($TimeoutSec * 1000)) {
        $line = "MODE $m => HANG (killed after ${TimeoutSec}s)"
        $proc.Kill()
        $proc.WaitForExit()
        Add-Content $resultFile $line
        Write-Output $line
    } else {
        Write-Output ("MODE $m => exit=" + $proc.ExitCode)
        Get-Content $resultFile -ErrorAction SilentlyContinue | ForEach-Object { Write-Output ("   " + $_) }
    }
}
Write-Output "SEQ-WATCHDOG-DONE"
