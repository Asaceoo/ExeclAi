$p = 'D:\execl1.0\src\ExcelAiAssistant.Addin\Sidebar\index.html'
$lines = Get-Content $p
$pat = '\.system|\.toolcard|\.collapsed|\.caret|\.body|\.head\b|displays?\s*:\s*none|\.success'
for ($i = 0; $i -lt $lines.Count; $i++) {
  if ($lines[$i] -match $pat -and $i -lt 118) {   # only <style> area (before <body>)
    ($i + 1).ToString() + ': ' + $lines[$i].Trim().Substring(0, [Math]::Min(200, $lines[$i].Trim().Length))
  }
}