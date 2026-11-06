$src = 'D:\execl1.0\src'
Get-ChildItem -Path $src -Recurse -Filter *.cs | ForEach-Object {
  $hits = Select-String -Path $_.FullName -Pattern '你是.*助手|SystemPrompt|system_prompt|ChatMessage\.System|你是嵌入|JYYJ助手|操作助手'
  if ($hits) {
    $rel = $_.FullName.Substring($src.Length)
    foreach ($h in $hits) { $rel + ' :: L' + $h.LineNumber + ': ' + $h.Line.Trim().Substring(0,[Math]::Min(140,$h.Line.Trim().Length)) }
  }
}