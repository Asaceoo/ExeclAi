$dir = [Environment]::GetFolderPath('LocalApplicationData') + '\XaiAssistant\app'
Get-ChildItem $dir -Recurse -Filter 'ExcelAiAssistant.Addin.dll' -ErrorAction SilentlyContinue | ForEach-Object {
  Write-Output ($_.FullName + ' | ' + $_.VersionInfo.FileVersion + ' | ' + $_.LastWriteTime)
}
$ex = Get-Process -Name EXCEL, et -ErrorAction SilentlyContinue
if ($ex) { foreach ($p in $ex) { Write-Output ('PROC ' + $p.ProcessName + ' PID=' + $p.Id + ' TITLE=[' + $p.MainWindowTitle + ']') } }
else { Write-Output 'NO-HOST-PROC' }
$z = Get-Process -Id 20732 -ErrorAction SilentlyContinue
if ($z) { Write-Output 'ZOMBIE 20732 still present' }
