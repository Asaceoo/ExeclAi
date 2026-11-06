# 「package-xai.ps1 -BumpVersion」升版同步逻辑回归测试（沙盒：临时目录，不触碰真实文件）。
# 从真实脚本提取 Get-IncrementedVersion / Sync-VersionFiles 函数在沙盒内执行，
# 断言 csproj / README / iss / 手册版本戳（JYYJ助手技术手册-<版本>.md/html、JYYJ助手用户手册-<版本>.md/html，均带版本号文件名）
# 同步正确、BOM 保持、历史 v 前缀条目不被改写、旧文件不被改动、可连续升版、docs 缺失时容忍跳过。
# 用法：powershell -ExecutionPolicy Bypass -File .\verify-versionbump.ps1
$ErrorActionPreference = "Stop"
$here = $PSScriptRoot
$passed = 0
$failed = 0

function Assert-Test([bool]$c, [string]$name, [string]$detail) {
    if ($c) { $script:passed++; Write-Host "[PASS] $name -> $detail" }
    else { $script:failed++; Write-Host "[FAIL] $name -> $detail" }
}

function Has-Bom([string]$path) {
    $b = [System.IO.File]::ReadAllBytes($path)
    return $b.Length -ge 3 -and $b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF
}

$realScript = Join-Path $here "package-xai.ps1"
if (-not (Test-Path $realScript)) { Write-Host "[fatal] 未找到 package-xai.ps1"; exit 2 }

# 提取函数定义（真实脚本顶层还有可执行段，不能整体 dot-source）
$tokens = $null; $parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($realScript, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors -and $parseErrors.Count -gt 0) {
    Write-Host "[FAIL] package-xai.ps1 语法错误：$($parseErrors[0].Message)"; exit 1
}
$funcAsts = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)
$names = @($funcAsts | ForEach-Object { $_.Name })
foreach ($need in @("Get-CurrentVersion", "Get-IncrementedVersion", "Sync-VersionFiles", "Find-InnoSetup", "Invoke-InnoCompile")) {
    Assert-Test ($names -contains $need) "函数存在：$need" ($names -join ",")
}
$sb = [System.Text.StringBuilder]::new()
foreach ($f in $funcAsts) { [void]$sb.AppendLine($f.Extent.Text) }
. ([scriptblock]::Create($sb.ToString()))

# ---------- Get-IncrementedVersion 纯逻辑 ----------
Assert-Test ((Get-IncrementedVersion "2.18.7.0" "Patch") -eq "2.18.8.0") "Patch 递增 2.18.7.0" (Get-IncrementedVersion "2.18.7.0" "Patch")
Assert-Test ((Get-IncrementedVersion "2.18.7.0" "Build") -eq "2.18.7.1") "Build 递增 2.18.7.0" (Get-IncrementedVersion "2.18.7.0" "Build")
Assert-Test ((Get-IncrementedVersion "2.18.7.0" "Minor") -eq "2.19.0.0") "Minor 递增并清零尾段" (Get-IncrementedVersion "2.18.7.0" "Minor")
Assert-Test ((Get-IncrementedVersion "2.18.7.0" "Major") -eq "3.0.0.0") "Major 递增并清零尾段" (Get-IncrementedVersion "2.18.7.0" "Major")
$threw = $false
try { [void](Get-IncrementedVersion "2.18" "Patch") } catch { $threw = $true }
Assert-Test $threw "非法版本号（非 4 段）抛异常" "2.18"

# ---------- Sync-VersionFiles 沙盒同步 ----------
$sandbox = Join-Path ([System.IO.Path]::GetTempPath()) ("versionbump-test-" + [System.Guid]::NewGuid().ToString("N"))
$src = Join-Path $sandbox "src"
$root = $sandbox
$toolDir = Join-Path $sandbox "tools"
$oldVer = "2.18.7.0"
$newVer = "2.18.8.0"
try {
    New-Item -ItemType Directory -Path (Join-Path $src "ExcelAiAssistant.Addin") -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $toolDir "installer") -Force | Out-Null

    $csprojPath = Join-Path $src "ExcelAiAssistant.Addin\ExcelAiAssistant.Addin.csproj"
    Set-Content -Path $csprojPath -Value @(
        "<Project>",
        "  <PropertyGroup>",
        "    <Version>$oldVer</Version>",
        "    <InformationalVersion>$oldVer</InformationalVersion>",
        "  </PropertyGroup>",
        "</Project>"
    ) -Encoding UTF8

    $readmePath = Join-Path $root "README.md"
    Set-Content -Path $readmePath -Value "Windows Excel 本机智能助手（v$oldVer）。基于 .NET Framework 4.8 + WebView2 侧边栏。" -Encoding UTF8

    # 旧 iss 用沙盒内自建最小夹具，而非从真实 installer 目录拷贝——历史版本 iss 会被归档，
    # 硬编码引用某版本会让回归脚本在版本推进后失效（曾在 2.18.7.0 归档后抛文件不存在）。
    $oldIss = Join-Path $toolDir "installer\JYYJ助手-$oldVer-Setup.iss"
    Set-Content -Path $oldIss -Value @(
        "; JYYJ助手 v$oldVer Setup（测试夹具）",
        "#define AppVersion `"$oldVer`"",
        "#define PkgSrc64 `"D:\dist\JYYJ助手-$oldVer-x64\app`"",
        "#define PkgSrc86 `"D:\dist\JYYJ助手-$oldVer-x86\app`"",
        "[Setup]",
        "AppVersion={#AppVersion}"
    ) -Encoding UTF8
    $oldIssHash = (Get-FileHash $oldIss).Hash

    # 手册版本戳夹具：三份 docs，各含「当前版本」戳 + 一条 v 前缀历史条目（历史条目不得被改写）
    $docsDir = Join-Path $root "docs"
    New-Item -ItemType Directory -Path $docsDir -Force | Out-Null
    $techMdPath = Join-Path $docsDir "JYYJ助手技术手册-$oldVer.md"
    Set-Content -Path $techMdPath -Value @(
        "# JYYJ助手 技术手册",
        "Excel COM Add-in · .NET Framework 4.8 · WebView2 · OpenAI 兼容协议 · 版本 $oldVer",
        "src/ExcelAiAssistant.Addin.csproj → 版本 (Version / InformationalVersion) $oldVer",
        "合并安装器 JYYJ助手-$oldVer-Setup.exe（Inno Setup 6）",
        "**v$oldVer 运行日志全功能覆盖**（历史条目，不得改写）"
    ) -Encoding UTF8
    $userHtmlPath = Join-Path $docsDir "JYYJ助手用户手册-$oldVer.html"
    Set-Content -Path $userHtmlPath -Value @(
        "<html><body>",
        "<div class=`"sub`"> Excel 里的 AI 助手 · 版本 $oldVer</div>",
        "<div class=`"callout`">v$oldVer：运行日志已覆盖全部功能（历史条目，不得改写）</div>",
        "</body></html>"
    ) -Encoding UTF8
    $userMdPath = Join-Path $docsDir "JYYJ助手用户手册-$oldVer.md"
    Set-Content -Path $userMdPath -Value @(
        "# JYYJ助手 用户手册",
        "",
        "> Excel 里的 AI 助手 · 版本 $oldVer",
        "",
        "> **v$oldVer：运行日志已覆盖全部功能**（历史条目，不得改写）"
    ) -Encoding UTF8
    $techHtmlPath = Join-Path $docsDir "JYYJ助手技术手册-$oldVer.html"
    Set-Content -Path $techHtmlPath -Value @(
        "<html><body>",
        "<div class=`"sub`">Excel COM Add-in · .NET Framework 4.8<br>版本 $oldVer</div>",
        "</body></html>"
    ) -Encoding UTF8

    Sync-VersionFiles $oldVer $newVer

    # 1) csproj：版本替换 + BOM
    $csproj = Get-Content $csprojPath -Raw -Encoding UTF8
    Assert-Test ($csproj -match "<Version>$newVer</Version>" -and $csproj -match "<InformationalVersion>$newVer</InformationalVersion>") "csproj Version/InformationalVersion 更新" $newVer
    Assert-Test ($csproj -notmatch "2\.18\.7\.0") "csproj 无旧版本残留" "2.18.7.0"
    Assert-Test (Has-Bom $csprojPath) "csproj 保持 UTF-8 BOM" "EF BB BF"

    # 2) README：v 前缀版本替换 + BOM
    $readme = Get-Content $readmePath -Raw -Encoding UTF8
    Assert-Test ($readme -match "v$newVer" -and $readme -notmatch "v2\.18\.7\.0") "README v 前缀版本更新" "v$newVer"
    Assert-Test (Has-Bom $readmePath) "README 保持 UTF-8 BOM" "EF BB BF"

    # 3) iss：新文件复制 + 内容替换；旧文件不动
    $newIss = Join-Path $toolDir "installer\JYYJ助手-$newVer-Setup.iss"
    Assert-Test (Test-Path $newIss) "新 iss 已生成" $newIss
    Assert-Test ((Get-FileHash $oldIss).Hash -eq $oldIssHash) "旧 iss 未被改动" "哈希一致"
    $iss = Get-Content $newIss -Raw -Encoding UTF8
    Assert-Test ($iss -match '#define AppVersion "' + $newVer + '"') "iss AppVersion 更新" $newVer
    Assert-Test ($iss -match "JYYJ助手-$newVer-x64" -and $iss -match "JYYJ助手-$newVer-x86") "iss PkgSrc x64/x86 路径更新" "dist\JYYJ助手-$newVer-{x64,x86}"
    Assert-Test ($iss -match "v$newVer" -and $iss -notmatch "2\.18\.7\.0") "iss 无旧版本残留（含首行注释）" "v$newVer"
    Assert-Test (-not (Has-Bom $newIss)) "新 iss 为 UTF-8 无 BOM（Inno Setup 约定）" "no BOM"

    # 3b) 手册版本戳：三份 docs 当前版本戳更新、历史 v 前缀条目保留、BOM 保持
    # 技术手册带版本号文件名：升版后应已随版本戳改名，旧名不存在
    $techMdNewPath = Join-Path $docsDir "JYYJ助手技术手册-$newVer.md"
    Assert-Test ((Test-Path -LiteralPath $techMdNewPath) -and -not (Test-Path -LiteralPath $techMdPath)) "技术手册 md 随升版改名（旧名移除）" "JYYJ助手技术手册-$newVer.md"
    $techMd = Get-Content $techMdNewPath -Raw -Encoding UTF8
    Assert-Test ($techMd -match "版本 $newVer" -and $techMd -match "\(Version / InformationalVersion\) $newVer" -and $techMd -match "JYYJ助手-$newVer-Setup") "技术手册 md 三类版本戳更新" $newVer
    Assert-Test ($techMd -match "v2\.18\.7\.0 运行日志全功能覆盖") "tech md 历史 v 前缀条目未被改写" "v2.18.7.0 保留"
    Assert-Test (Has-Bom $techMdNewPath) "技术手册 md 保持 UTF-8 BOM" "EF BB BF"
    # 用户手册带版本号文件名：升版后应已随版本戳改名，旧名不存在
    $userHtmlNewPath = Join-Path $docsDir "JYYJ助手用户手册-$newVer.html"
    Assert-Test ((Test-Path -LiteralPath $userHtmlNewPath) -and -not (Test-Path -LiteralPath $userHtmlPath)) "用户手册 html 随升版改名（旧名移除）" "JYYJ助手用户手册-$newVer.html"
    $userHtml = Get-Content $userHtmlNewPath -Raw -Encoding UTF8
    Assert-Test ($userHtml -match "版本 $newVer" -and $userHtml -match "v2\.18\.7\.0：运行日志已覆盖全部功能") "用户手册 html 版本戳更新 + 历史条目保留" $newVer
    Assert-Test (Has-Bom $userHtmlNewPath) "用户手册 html 保持 UTF-8 BOM" "EF BB BF"
    $techHtmlNewPath = Join-Path $docsDir "JYYJ助手技术手册-$newVer.html"
    Assert-Test ((Test-Path -LiteralPath $techHtmlNewPath) -and -not (Test-Path -LiteralPath $techHtmlPath)) "技术手册 html 随升版改名（旧名移除）" "JYYJ助手技术手册-$newVer.html"
    $techHtml = Get-Content $techHtmlNewPath -Raw -Encoding UTF8
    Assert-Test ($techHtml -match "版本 $newVer") "技术手册 html 版本戳更新" $newVer
    Assert-Test (Has-Bom $techHtmlNewPath) "技术手册 html 保持 UTF-8 BOM" "EF BB BF"
    $userMdNewPath = Join-Path $docsDir "JYYJ助手用户手册-$newVer.md"
    Assert-Test ((Test-Path -LiteralPath $userMdNewPath) -and -not (Test-Path -LiteralPath $userMdPath)) "用户手册 md 随升版改名（旧名移除）" "JYYJ助手用户手册-$newVer.md"
    $userMd = Get-Content $userMdNewPath -Raw -Encoding UTF8
    Assert-Test ($userMd -match "版本 $newVer" -and $userMd -match "v2\.18\.7\.0：运行日志已覆盖全部功能") "用户手册 md 版本戳更新 + 历史条目保留" $newVer
    Assert-Test (Has-Bom $userMdNewPath) "用户手册 md 保持 UTF-8 BOM" "EF BB BF"

    # 4) 连续升版：2.18.8.0 -> 2.18.9.0（基于上一步新 iss 继续复制）
    $oldIss2 = $newIss
    $newVer2 = "2.18.9.0"
    $newIss2 = Join-Path $toolDir "installer\JYYJ助手-$newVer2-Setup.iss"
    Sync-VersionFiles $newVer $newVer2
    Assert-Test (Test-Path $newIss2) "连续升版新 iss 生成" $newIss2
    $iss2 = Get-Content $newIss2 -Raw -Encoding UTF8
    Assert-Test ($iss2 -match "JYYJ助手-$newVer2-x64" -and $iss2 -match "v$newVer2" -and $iss2 -notmatch "2\.18\.8\.0") "连续升版内容替换正确" "2.18.8.0 -> $newVer2"
    $techMd2 = Get-Content (Join-Path $docsDir "JYYJ助手技术手册-$newVer2.md") -Raw -Encoding UTF8
    Assert-Test ($techMd2 -match "版本 $newVer2" -and $techMd2 -notmatch "版本 2\.18\.8\.0") "连续升版手册版本戳跟随更新" "2.18.8.0 -> $newVer2"
    Assert-Test (Test-Path -LiteralPath (Join-Path $docsDir "JYYJ助手用户手册-$newVer2.md")) "连续升版用户手册文件名跟随新版本" "JYYJ助手用户手册-$newVer2.md"
    Assert-Test (Test-Path -LiteralPath (Join-Path $docsDir "JYYJ助手技术手册-$newVer2.md")) "连续升版技术手册文件名跟随新版本" "JYYJ助手技术手册-$newVer2.md"

    # 4b) 容忍缺失：docs 目录整个删掉后升版不得报错（手册可能被移动/重构，同步逻辑需容忍）
    Remove-Item $docsDir -Recurse -Force
    $threw = $false
    try { Sync-VersionFiles $newVer2 "2.18.9.1" } catch { $threw = $true }
    Assert-Test (-not $threw) "docs 缺失时手册同步容忍跳过（不抛错）" "docs 目录已删除"

    # 5) 重复目标保护：同版本已存在时抛异常
    $threw = $false
    try { Sync-VersionFiles $newVer $newVer } catch { $threw = $true }
    Assert-Test $threw "目标 iss 已存在拒绝覆盖" "$newVer"

    # 6) fail-closed 守卫（对抗审查新增）：任一替换未命中必须抛错，不得静默漂移。
    # 6a) csproj 缺 <InformationalVersion> → csproj 守卫先抛错（不落盘坏版本）。
    Set-Content -Path $csprojPath -Value @(
        "<Project>",
        "  <PropertyGroup>",
        "    <Version>9.9.9.8</Version>",
        "  </PropertyGroup>",
        "</Project>"
    ) -Encoding UTF8
    $threw = $false
    try { Sync-VersionFiles "9.9.9.8" "9.9.9.9" } catch { $threw = $true }
    Assert-Test $threw "csproj 缺 InformationalVersion → fail-closed 抛错" "9.9.9.8"

    # 6b) README 版本非「v 前缀」格式 → README 守卫抛错（未命中则不落盘）。
    Set-Content -Path $csprojPath -Value @(
        "<Project>",
        "  <PropertyGroup>",
        "    <Version>8.8.8.8</Version>",
        "    <InformationalVersion>8.8.8.8</InformationalVersion>",
        "  </PropertyGroup>",
        "</Project>"
    ) -Encoding UTF8
    Set-Content -Path $readmePath -Value "助手（8.8.8.8）无 v 前缀" -Encoding UTF8
    $threw = $false
    try { Sync-VersionFiles "8.8.8.8" "8.8.8.9" } catch { $threw = $true }
    Assert-Test $threw "README 无 v 前缀版本 → fail-closed 抛错" "8.8.8.8"

    # ---------- Inno Setup 编译（Find-InnoSetup / Invoke-InnoCompile） ----------
    # 预备：脚本作用域需先给函数可见的全局变量赋值（dot-source 到本脚本作用域）。
    $script:ISCCPath = ""
    $script:Version = $newVer
    $script:dist = Join-Path $sandbox "dist"
    $script:toolDir = $toolDir          # 即 sandbox\tools
    # fake iscc：退出码 0 且按函数期望路径落出产物（含中文文件名，须 UTF-8 BOM 写才不乱码）
    $fakeIscc = Join-Path $sandbox "fake-iscc.ps1"
    Set-Content -Path $fakeIscc -Value @(
        '$out = Join-Path $PSScriptRoot "dist\installer"',
        'New-Item -ItemType Directory -Path $out -Force | Out-Null',
        '[System.IO.File]::WriteAllBytes((Join-Path $out "JYYJ助手-2.18.8.0-Setup.exe"), [byte[]](0,1,2))',
        'exit 0'
    ) -Encoding UTF8
    # fake iscc-lazy：退出码 0 但不落产物（用于 fail-closed 负向校验）
    $lazyIscc = Join-Path $sandbox "fake-iscc-lazy.ps1"
    Set-Content -Path $lazyIscc -Value @('exit 0') -Encoding UTF8
    $newIss = Join-Path $toolDir "installer\JYYJ助手-$newVer-Setup.iss"

    $script:ISCCPath = $fakeIscc
    Assert-Test ((Find-InnoSetup) -eq $fakeIscc) "Find-InnoSetup 显式路径命中" $fakeIscc
    $script:ISCCPath = Join-Path $sandbox "no-such-iscc.exe"
    $threw = $false
    try { [void](Find-InnoSetup) } catch { $threw = $true }
    Assert-Test $threw "Find-InnoSetup 显式路径不存在抛异常" "no-such-iscc.exe"

    $script:ISCCPath = $fakeIscc
    New-Item -ItemType Directory -Path (Join-Path $dist "JYYJ助手-$newVer-x64") -Force | Out-Null
    $threw = $false
    try { Invoke-InnoCompile $fakeIscc $newVer } catch { $threw = $true }
    Assert-Test $threw "Inno 前置守卫：x64/app 缺失抛异常（不调用 iscc）" "$dist\JYYJ助手-$newVer-x64\app"

    foreach ($arch in @("x64", "x86")) {
        New-Item -ItemType Directory -Path (Join-Path $dist "JYYJ助手-$newVer-$arch\app") -Force | Out-Null
    }
    Invoke-InnoCompile $fakeIscc $newVer
    Assert-Test (Test-Path (Join-Path $dist "installer\JYYJ助手-$newVer-Setup.exe")) "Inno 前置齐备 + 落产物 → 编译通过" "fake iscc"

    # 负向：iscc 返回 0 但未落产物 → fail-closed 抛异常（先清掉正向用例落下的产物，否则恒真失真）
    Remove-Item (Join-Path $dist "installer\JYYJ助手-$newVer-Setup.exe") -Force
    $threw = $false
    try { Invoke-InnoCompile $lazyIscc $newVer } catch { $threw = $true }
    Assert-Test $threw "Inno 返回 0 但无产物 → fail-closed 抛异常" "lazy fake iscc"
}
finally {
    Remove-Item $sandbox -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host ""
Write-Host "== 结果：通过 $passed，失败 $failed =="
exit $(if ($failed -eq 0) { 0 } else { 1 })
