$f = 'd:\execl1.0\src\ExcelAiAssistant.Addin\Sidebar\index.html'
Select-String -Path $f -Pattern 'makePanelSinkButton|showCodeView|codeViewOverlay|SinkVbaTool|clampInt|makeSinkButton|renderVbaResult\(' -AllMatches | ForEach-Object {
    '{0}: {1}' -f $_.LineNumber, ($_.Matches.Value -join ', ')
}