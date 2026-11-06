# Real-machine verification for P1 (structured return + data channel).
# Loads the freshly built addin assembly, launches a real Excel, drives the
# built-in DiagnosticSmokeRunner (write/chart/VBA/Python + P1 data channels)
# via reflection, and reports per-check results.
$ErrorActionPreference = "Stop"

$dir = "d:\execl1.0\src\ExcelAiAssistant.Addin\bin\x64\Release\net48\win-x64"
$addinDll = Join-Path $dir "ExcelAiAssistant.Addin.dll"
if (-not (Test-Path $addinDll)) { Write-Output ("MISSING " + $addinDll); exit 1 }

$results = @()
$excel = $null
try {
    [System.Reflection.Assembly]::LoadFrom($addinDll) | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    Start-Sleep -Seconds 2

    $asm = [System.Reflection.Assembly]::LoadFrom($addinDll)
    $type = $asm.GetType("ExcelAiAssistant.Addin.DiagnosticSmokeRunner", $true)
    $smoke = [Activator]::CreateInstance($type, @($excel))
    $run = $type.GetMethod("Run")
    $list = $run.Invoke($smoke, @())

    foreach ($r in $list) {
        $name = $r.GetType().GetProperty("name").GetValue($r, $null)
        $ok = $r.GetType().GetProperty("ok").GetValue($r, $null)
        $detail = $r.GetType().GetProperty("detail").GetValue($r, $null)
        $results += ("{0} | ok={1} | {2}" -f $name, $ok, $detail)
    }
} catch {
    $results += ("FAIL exception " + $_.Exception.ToString())
} finally {
    if ($excel -ne $null) {
        try { $excel.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null } catch {}
    }
}

$results | ForEach-Object { Write-Output $_ }
$failed = @($results | Where-Object { $_ -match "ok=False|FAIL" }).Count
Write-Output ("FAIL_COUNT=" + $failed)
if ($failed -gt 0) { exit 1 }
Write-Output "REAL-MACHINE VERIFY OK"
exit 0
