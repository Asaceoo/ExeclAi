# Deterministic real-machine check on the installed x64 Addin.dll.
$ErrorActionPreference = "Stop"
$asmPath = "$env:LOCALAPPDATA\XaiAssistant\app\x64\ExcelAiAssistant.Addin.dll"
$pass = 0; $fail = 0

function Check([string]$label, [bool]$ok, [string]$detail) {
    if ($ok) { $script:pass++; Write-Output ("PASS  {0}  [{1}]" -f $label, $detail) }
    else     { $script:fail++; Write-Output ("FAIL  {0}  [{1}]" -f $label, $detail) }
}

if (-not (Test-Path $asmPath)) { Write-Output "NO_DLL"; exit 1 }
$bytes = [IO.File]::ReadAllBytes($asmPath)
$asm = [System.Reflection.Assembly]::Load($bytes)
$ver = $asm.GetName().Version.ToString()
Check "asm-version-1.17.0.0" ($ver -eq "1.17.0.0") $ver

# 用码点拼中文标签，避免 .ps1 无 BOM 被按 ANSI 读而乱码。
$pre  = -join ([char]0x9884, [char]0x7f6e)         # 预置
$ai   = "AI" + -join ([char]0x6c89, [char]0x6dc0)   # AI沉淀
$hand = -join ([char]0x624b, [char]0x5199)          # 手写

$resName = "ExcelAiAssistant.Addin.Sidebar.index.html"
$s = $asm.GetManifestResourceStream($resName)
Check "res-index-embed" ($null -ne $s) $resName
$html = ""
if ($null -ne $s) {
    $rd = New-Object IO.StreamReader($s, [Text.Encoding]::UTF8)
    $html = $rd.ReadToEnd(); $rd.Close()
}
Check "html-nonempty" ($html.Length -gt 0) ("len=" + $html.Length)

foreach ($needle in @("loadVbaTemplates","renderVbaTemplates","vbaSaveTemplateButton","SaveUserTemplate","vbaTemplateCategory","vbaTemplateSearch","vbaTemplateHint","CATEGORY_LABEL")) {
    Check ("html-contains:" + $needle) ($html.Contains($needle)) $needle
}
Check "html-label-preset" $html.Contains($pre) "preset-label"
Check "html-label-ai"     $html.Contains($ai)  "ai-label"
Check "html-label-hand"   $html.Contains($hand) "user-label"
Check "html-no-hardcode"  (-not $html.Contains("const VBA_TEMPLATES = {")) "hardcode-removed"

$t = $asm.GetType("ExcelAiAssistant.Addin.SidebarBridge")
$mList = @($t.GetMethods() | Where-Object { $_.Name -eq "ListCustomTools" })[0]
$mSave = @($t.GetMethods() | Where-Object { $_.Name -eq "SaveUserTemplate" })[0]
$nList = 0; if ($mList) { $nList = $mList.GetParameters().Count }
$nSave = 0; if ($mSave) { $nSave = $mSave.GetParameters().Count }
Check "bridge-ListCustomTools-cat" ($mList -and $nList -eq 1) ("ListCustomTools params=" + $nList)
Check "bridge-SaveUserTemplate-4"  ($mSave -and $nSave -eq 4) ("SaveUserTemplate params=" + $nSave)

Write-Output ("== PASSED=" + $pass + " FAILED=" + $fail + " ==")
if ($fail -gt 0) { exit 1 }
exit 0