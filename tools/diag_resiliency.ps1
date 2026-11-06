$ErrorActionPreference = 'Continue'
$k = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems'
$items = Get-Item $k -ErrorAction SilentlyContinue
if ($items) {
  foreach ($n in $items.GetValueNames()) {
    $b = $items.GetValue($n)
    $txt = ''
    try {
      $bytes = [byte[]]$b
      # DisabledItems 值是二进制，尾部含被禁用项的路径字符串（UTF-16）
      $s = [System.Text.Encoding]::Unicode.GetString($bytes)
      $txt = $s -replace '[^\x20-\x7e\u4e00-\u9fff]', '.'
    } catch { $txt = 'READ-ERR' }
    Write-Output ("DISABLED name=[" + $n + "] len=" + ($b.Length) + " content=" + $txt.Substring(0, [Math]::Min(160, $txt.Length)))
  }
}
Write-Output '=== JYYJ Addin CodeBase ==='
foreach ($n in @('ExcelAiAssistant.AddIn','ExcelAiAssistant.TaskPaneHost','ExcelAiAssistant.DeepSeekPaneHost')) {
  $p = Get-ItemProperty ('HKCU:\Software\Microsoft\Office\Excel\Addins\' + $n) -ErrorAction SilentlyContinue
  if ($p) { Write-Output ($n + ' CodeBase=[' + $p.CodeBase + '] Description=[' + $p.Description + ']') }
}
Write-Output '=== 安装目录 DLL 版本 ==='
$dir = [Environment]::GetFolderPath('LocalApplicationData') + '\XaiAssistant\app'
Get-ChildItem $dir -Recurse -Filter 'ExcelAiAssistant.Addin.dll' -ErrorAction SilentlyContinue | ForEach-Object {
  Write-Output ($_.FullName + ' | ' + $_.VersionInfo.FileVersion)
}
