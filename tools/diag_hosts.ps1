$procs = Get-Process -Name EXCEL, et, wps, wpp -ErrorAction SilentlyContinue
if (-not $procs) { Write-Output "NONE"; exit 0 }
foreach ($p in $procs) {
  $cmd = ""
  try {
    $cim = Get-CimInstance Win32_Process -Filter ("ProcessId=" + $p.Id) -ErrorAction SilentlyContinue
    if ($cim -and $cim.CommandLine) { $cmd = $cim.CommandLine }
  } catch {}
  Write-Output ("PID=" + $p.Id + " NAME=" + $p.ProcessName + " TITLE=[" + $p.MainWindowTitle + "] CMD=[" + $cmd + "]")
}
