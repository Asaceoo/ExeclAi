$ErrorActionPreference = 'Continue'
$k = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems'
$items = Get-Item $k -ErrorAction SilentlyContinue
if ($items -and $items.ValueCount -gt 0) {
  foreach ($n in $items.GetValueNames()) {
    Remove-ItemProperty -Path $k -Name $n -Force
    Write-Output ('REMOVED disabled item: ' + $n)
  }
} else {
  Write-Output 'NO DISABLED ITEMS'
}
$after = Get-Item $k -ErrorAction SilentlyContinue
Write-Output ('AFTER count=' + $(if ($after) { $after.ValueCount } else { 0 }))

Write-Output '=== CLSID CodeBase check ==='
$clsidRoot = 'HKCU:\Software\Classes\CLSID'
$found = Get-ChildItem $clsidRoot -ErrorAction SilentlyContinue | Where-Object {
  $p = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
  $cb = (Get-ItemProperty ($_.PSPath + '\InprocServer32') -ErrorAction SilentlyContinue).'(default)'
  $cb -and $cb -like '*XaiAssistant*'
}
foreach ($c in $found) {
  $cb2 = (Get-ItemProperty ($c.PSPath + '\InprocServer32') -ErrorAction SilentlyContinue).'(default)'
  Write-Output ('CLSID ' + $c.PSChildName + ' -> ' + $cb2)
}
