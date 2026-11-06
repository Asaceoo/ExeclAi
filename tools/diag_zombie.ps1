$z = Get-CimInstance Win32_Process -Filter "ProcessId=20732" -ErrorAction SilentlyContinue
if ($z) {
  Write-Output ("ZOMBIE PID=20732 ParentPID=" + $z.ParentProcessId + " Created=" + $z.CreationDate)
  $par = Get-Process -Id $z.ParentProcessId -ErrorAction SilentlyContinue
  if ($par) {
    Write-Output ("PARENT ALIVE: PID=" + $par.Id + " NAME=" + $par.ProcessName + " TITLE=[" + $par.MainWindowTitle + "]")
  } else {
    Write-Output "PARENT DEAD"
  }
} else {
  Write-Output "20732 GONE"
}
