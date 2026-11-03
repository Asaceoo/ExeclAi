# JYYJ助手 GitHub 推送 + Releases 上传一键脚本
# 前置：①代理 127.0.0.1:7897 在跑 ②SSH config 已配 ③Releases 需 token: set GITHUB_TOKEN=ghp_xxx
$ErrorActionPreference = "Continue"
$repo = "Asaceoo/ExeclAi"
$ver = (Select-Xml -Path (Join-Path $PSScriptRoot "..\..\Directory.Build.props") -XPath "//Version").Node.InnerText.Trim()  # 版本号单一来源（M17.1 红线）
$tag = "v$ver"
Write-Host "=== 1/3 push source (github-main -> main) ==="
& "C:\Program Files\Git\mingw64\bin\git.exe" push git@github.com:Asaceoo/ExeclAi.git github-main:main --force
if ($LASTEXITCODE -eq 0) { Write-Host "PUSH OK" } else { Write-Host "PUSH FAIL (rerun to resume)" }
Write-Host "=== 2/3 create Release $tag ==="
$token = $env:GITHUB_TOKEN
if (-not $token) { Write-Host "NO GITHUB_TOKEN, skip release upload. Set: $env:GITHUB_TOKEN=ghp_xxx"; exit 0 }
$headers = @{ Authorization = "token $token"; Accept = "application/vnd.github+json" }
$base = "D:\execl1.0-wps\dist"
$relBody = @{ tag_name = $tag; target_commitish = "main"; name = "JYYJ助手 $ver"; body = "下载: 安装版 Setup-wps.exe / 便携 zip x64+x86 / 单文件便携。更新: 三轮对抗审查 33 项修复, MCP+CLI 双入口, README 用法+20实例, 手册免责声明章。"; draft = $false; prerelease = $false } | ConvertTo-Json
$rel = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/releases" -Method Post -Headers $headers -Body $relBody -ContentType "application/json"
Write-Host "Release: $($rel.html_url)"
Write-Host "=== 3/3 upload assets ==="
$assets = @(
  "$base\installer\JYYJ助手-$ver-Setup-wps.exe",
  "$base\JYYJ助手-$ver-wps-x64.zip",
  "$base\JYYJ助手-$ver-wps-x86.zip",
  "$base\JYYJ助手-$ver-wps-单文件便携-x64.zip"
)
foreach ($a in $assets) {
  $name = [IO.Path]::GetFileName($a)
  Write-Host "upload $name ..."
  $uri = "https://uploads.github.com/repos/$repo/releases/$($rel.id)/assets?name=$([Uri]::EscapeDataString($name))"
  Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -InFile $a -ContentType "application/octet-stream"
  Write-Host "  OK"
}
Write-Host "DONE: https://github.com/$repo/releases/tag/$tag"