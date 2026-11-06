$ErrorActionPreference = 'Continue'
$dir = 'D:\execl1.0-wps\2026-09-11-9accc9f9\output\tasks\task_1789064509199_a0xtqh47'
$f = Get-ChildItem -Path $dir -Filter 'JYYJ*.xlsx' | Select-Object -First 1
if (-not $f) { Write-Output 'FILE_NOT_FOUND'; exit 2 }
Write-Output ('FILE=' + $f.FullName)
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $wb = $excel.Workbooks.Open($f.FullName)
  foreach ($ws in $wb.Worksheets) {
    $co = $ws.ChartObjects()
    if ($co.Count -eq 0) { continue }
    for ($i = 1; $i -le $co.Count; $i++) {
      $ch = $co.Item($i).Chart
      $t = $ch.ChartType
      $sc = $ch.SeriesCollection()
      $n = $sc.Count
      $info = "SHEET=[" + $ws.Name + "] TYPE=" + $t + " SERIES=" + $n
      for ($s = 1; $s -le [Math]::Min($n, 4); $s++) {
        $ser = $sc.Item($s)
        $pv = 0
        try { $pv = $ser.Points().Count } catch { $pv = -1 }
        $frm = ""
        try { $frm = $ser.Formula } catch { $frm = "NOFORMULA" }
        if ($frm.Length -gt 130) { $frm = $frm.Substring(0, 130) }
        $info += " | S" + $s + " pts=" + $pv + " f=" + $frm
      }
      Write-Output $info
    }
  }
  $wb.Close($false)
} finally {
  $excel.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  [GC]::Collect()
}
