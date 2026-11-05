$ErrorActionPreference = 'Continue'
$xl = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
if ($xl) { Write-Output ('C2R Platform=' + $xl.Platform + ' Version=' + $xl.VersionToReport) }

foreach ($view in @('64','32')) {
  $suffix = ''
  if ($view -eq '32') { $suffix = '\WOW6432Node' }
  foreach ($hive in @('HKCU','HKLM')) {
    if ($hive -eq 'HKCU') { $root = 'HKCU:\Software' + $suffix + '\Microsoft\Office\Excel\Addins' }
    else { $root = 'HKLM:\SOFTWARE' + $suffix + '\Microsoft\Office\Excel\Addins' }
    if (Test-Path $root) {
      Get-ChildItem $root -ErrorAction SilentlyContinue | ForEach-Object {
        $p = Get-ItemProperty $_.PSPath
        Write-Output ("ADDIN[v" + $view + "/" + $hive + "] " + $_.PSChildName + " LoadBehavior=" + $p.LoadBehavior + " Manifest=[" + $p.Manifest + "]")
      }
    } else {
      Write-Output ("NOKEY[v" + $view + "/" + $hive + "] " + $root)
    }
  }
}

Write-Output '=== EXCEL processes ==='
$ex = Get-Process -Name EXCEL -ErrorAction SilentlyContinue
if ($ex) { foreach ($p in $ex) { Write-Output ('PID=' + $p.Id + ' TITLE=[' + $p.MainWindowTitle + ']') } }
else { Write-Output 'NONE' }

Write-Output '=== Resiliency (disabled addins) ==='
foreach ($v in @('15.0','16.0')) {
  $k = 'HKCU:\Software\Microsoft\Office\' + $v + '\Excel\Resiliency\DisabledItems'
  if (Test-Path $k) {
    $it = Get-Item $k
    Write-Output ('Resiliency ' + $v + ' values=' + $it.ValueCount)
  }
}
