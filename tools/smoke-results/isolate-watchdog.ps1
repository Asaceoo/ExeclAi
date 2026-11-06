# Watchdog wrapper: run each isolate variant in its own process with a hard timeout.
param([int]$TimeoutSec = 120)
$ErrorActionPreference = "Stop"
$ps = "C:\WINDOWS\System32\WindowsPowerShell\v1.0\powershell.exe"
$script = "d:\execl1.0\tools\smoke-results\isolate-run-one.ps1"
$variants = @("VU","VUC","VQ","VQC")

foreach ($v in $variants) {
    $resultFile = "d:\execl1.0\tools\smoke-results\iso-$v.txt"
    Remove-Item $resultFile -ErrorAction SilentlyContinue
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $ps
    $psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$script`" -Variant $v"
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    $null = $proc.Start()
    if (-not $proc.WaitForExit($TimeoutSec * 1000)) {
        $line = "VARIANT $v => HANG (killed after ${TimeoutSec}s)"
        $proc.Kill()
        $proc.WaitForExit()
        Add-Content $resultFile $line
        Write-Output $line
    } else {
        Write-Output ("VARIANT $v => exit=" + $proc.ExitCode)
        Get-Content $resultFile -ErrorAction SilentlyContinue | ForEach-Object { Write-Output ("   " + $_) }
    }
}
Write-Output "WATCHDOG-DONE"
