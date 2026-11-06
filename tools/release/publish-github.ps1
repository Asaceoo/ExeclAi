# JYYJ助手 GitHub 推送 + Releases 上传一键脚本（可重入：失败可重复执行，不会产生重复资产/重复 Release）
param([switch]$SkipPush, [switch]$SkipRelease)
# 用法：
#   $env:GITHUB_TOKEN="ghp_xxx"; powershell -ExecutionPolicy Bypass -File tools\release\publish-github.ps1
#   -SkipPush    只传 Release 资产（源码已推过）
#   -SkipRelease 只推源码
# 前置：①代理 127.0.0.1:7897 在跑（git 走 SSH 不需要）②SSH key 已配 ③dist 下已有 publish-assets-<ver>.txt（release.py pack 产出）
$ErrorActionPreference = "Continue"
$repo  = "Asaceoo/ExeclAi"
$root  = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
$ver   = (Select-Xml -Path (Join-Path $root "Directory.Build.props") -XPath "//Version").Node.InnerText.Trim()  # 版本号单一来源（M17.1 红线）
$tag   = "v$ver"
$base  = if ($env:JYYJ_DIST) { $env:JYYJ_DIST } else { Join-Path $root "dist" }
$git   = "C:\Program Files\Git\mingw64\bin\git.exe"
$mapFile = Join-Path $base "publish-assets-$ver.txt"

Write-Host "=== 0/3 环境检查（先查后跑，避免构建白跑后才发现没 token） ==="
Write-Host "version=$ver  dist=$base"
if (-not (Test-Path $mapFile)) { Write-Host "FATAL: 缺少资产清单 $mapFile（先跑 release.py pack）"; exit 2 }
$token = $env:GITHUB_TOKEN
if (-not $token -and -not $SkipRelease) {
    Write-Host "FATAL: 未设置 GITHUB_TOKEN，无法创建 Release / 上传资产。"
    Write-Host "  设置方法（当前会话）：`$env:GITHUB_TOKEN=\"ghp_xxx\""
    Write-Host "  取得地址：https://github.com/settings/tokens （勾 repo 权限）"
    if ($SkipPush) { exit 2 }
    Write-Host "  → 仅推源码，跳过 Release。"
}

if (-not $SkipPush) {
    Write-Host "=== 1/3 推送源码（重建 release-snapshot 孤儿分支 → 强推 main） ==="
    # M28 实测口径：github main 旧历史含 redist 大 blob，直推被 100MB 拒收；从工作树重建孤儿快照。
    & $git -C $root fetch git@github.com:Asaceoo/ExeclAi.git "+refs/heads/main:refs/remotes/origin/main"
    # 可重入：上次中断会残留本分支，先删掉再建
    & $git -C $root branch "-D" "release-snapshot-tmp" 2>$null | Out-Null
    & $git -C $root checkout --orphan release-snapshot-tmp | Out-Null
    & $git -C $root add -A
    $big = & $git -C $root ls-files -z | ForEach-Object { $_ } | Where-Object { (Get-Item $_ -ErrorAction SilentlyContinue).Length -gt 90MB }
    if ($big) { Write-Host "FATAL: 跟踪文件超过 90MB，GitHub 会拒绝推送："; $big | ForEach-Object { Write-Host "  $_" }; & $git -C $root checkout master; exit 3 }
    & $git -C $root -c core.safecrlf=false commit -m "snapshot: v$ver source snapshot (orphan, redist-free)" --allow-empty | Out-Null
    & $git -C $root push git@github.com:Asaceoo/ExeclAi.git release-snapshot-tmp:main --force
    $pushOk = ($LASTEXITCODE -eq 0)
    & $git -C $root checkout master | Out-Null
    & $git -C $root branch "-D" "release-snapshot-tmp" 2>$null | Out-Null
    if ($pushOk) { Write-Host "PUSH OK (main = v$ver snapshot)" } else { Write-Host "PUSH FAIL（可重跑本脚本续传）" }
}

if ($SkipRelease) { Write-Host "跳过 Release（-SkipRelease）"; exit 0 }
if (-not $token) { exit 0 }

Write-Host "=== 2/3 创建/复用 Release $tag ==="
$headers = @{ Authorization = "token $token"; Accept = "application/vnd.github+json"; "X-GitHub-Api-Version" = "2022-11-28" }
$rel = $null
$relBodyText = @"
## JYYJ助手 v$ver（WPS 版）

内嵌 Excel/WPS 的 AI 办公助手：自然语言操作表格，172 个工具 / 执行层自动自检 / 可回滚 / 全量日志。

### 下载
- **安装版（推荐）**：JYYJ-$ver-Setup-wps.exe
- **便携 x64 / x86**：JYYJ-$ver-wps-x64.zip / JYYJ-$ver-wps-x86.zip（含注册/卸载/诊断脚本）
- **程序目录包**：JYYJ-$ver-wps-portable-single-x64.zip（仅 app 目录，供外壳或高级用户）
- **MCP server**：JYYJ-$ver-mcp-server-x64.zip（npx jyyj-mcp 下载源）
- **校验**：SHA256SUMS-$ver.txt（先校验再解压）
- **文档**：user-manual / tech-manual（md + html）

### 本版要点
- 打包自动递增版本号（props/README/手册/npm 联动），旧手册自动归档
- 新增「连接」面板：MCP / CLI 开关 + 连通性自测（Ribbon → 连接）
- M29 AI 自主闭环：截图 / 显式自检 / 修正阶梯（≤5 轮）
- 修复加载项被禁用、CTP 类名错配导致面板不显示的问题（repair-addin.ps1 一键自愈）

### 免责声明
请在副本上操作并自行备份重要数据；本软件按「现状」提供，作者不对任何数据丢失或业务损失承担责任。下载后请先用 SHA256SUMS 校验。
"@
# 审查 R14：PowerShell 5.1 的 Invoke-RestMethod 把 -Body 字符串按 ASCII 编码发送，
# 中文会变成 "?"（实测 Release 名称/正文变 JYYJ?? ）。必须显式转 UTF-8 字节再发。
function Send-Json([string]$Uri, [string]$Method, $Object) {
    $json = $Object | ConvertTo-Json -Depth 3
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    return Invoke-RestMethod -Uri $Uri -Method $Method -Headers $headers -Body $bytes -ContentType "application/json; charset=utf-8"
}
try {
    $rel = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/releases/tags/$tag" -Method Get -Headers $headers
    Write-Host "Release 已存在，复用并校正名称/正文：$($rel.html_url)"
    $rel = Send-Json "https://api.github.com/repos/$repo/releases/$($rel.id)" "Patch" @{ name = "JYYJ助手 v$ver"; body = $relBodyText }
} catch {
    $rel = Send-Json "https://api.github.com/repos/$repo/releases" "Post" @{
        tag_name = $tag; target_commitish = "main"; name = "JYYJ助手 v$ver"
        body = $relBodyText; draft = $false; prerelease = $false
    }
    Write-Host "Release 创建：$($rel.html_url)"
}

Write-Host "=== 3/3 上传资产（同名先删后传，幂等） ==="
$assets = Get-Content $mapFile -Encoding UTF8 | Where-Object { $_ -and -not $_.StartsWith("#") } | ForEach-Object {
    $parts = $_ -split "`t"
    if ($parts.Count -ge 2) { [pscustomobject]@{ Path = $parts[0].Trim(); Name = $parts[1].Trim() } }
}
$missing = @($assets | Where-Object { -not (Test-Path $_.Path) })
if ($missing.Count) { foreach ($m in $missing) { Write-Host "MISSING: $($m.Path)" }; exit 4 }
$existing = @{}
foreach ($a in $rel.assets) { $existing[$a.name] = $a.id }

foreach ($a in $assets) {
    if ($existing.ContainsKey($a.Name)) {
        Write-Host "replace $($a.Name) ..."
        try { Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/releases/assets/$($existing[$a.Name])" -Method Delete -Headers $headers | Out-Null }
        catch { Write-Host "  delete failed（继续尝试上传）: $($_.Exception.Message)" }
    } else { Write-Host "upload $($a.Name) ..." }
    $uri = "https://uploads.github.com/repos/$repo/releases/$($rel.id)/assets?name=$([Uri]::EscapeDataString($a.Name))"
    try {
        Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -InFile $a.Path -ContentType "application/octet-stream" | Out-Null
        Write-Host ("  OK  {0:N2} MB" -f ((Get-Item $a.Path).Length / 1MB))
    } catch {
        Write-Host "  FAIL: $($_.Exception.Message)"
        exit 5
    }
}
Write-Host "DONE: https://github.com/$repo/releases/tag/$tag"