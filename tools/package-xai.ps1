# 构建并产出可分发的 JYYJ助手加载项部署包（x86 / x64）。
[CmdletBinding()]
param(
    [ValidateSet("Release", "Debug")]
    [string]$Configuration = "Release",
    [string]$Version = "",
    # 自动递增版本号：从 csproj 读当前版本，按 -BumpPart 递增并同步三处（csproj/README/iss 复制新版），
    # 再用新版本打包。与 -Version 互斥。
    [switch]$BumpVersion,
    [ValidateSet("Major", "Minor", "Patch", "Build")]
    [string]$BumpPart = "Patch",
    [switch]$SkipBuild,
    # 发布门禁：默认打包前先过「全量单元测试 + 真机自检回归」双关，任一失败即中止打包；
    # 仅紧急修复时可显式跳过（警告：单测与真机验证将同时被跳过）。
    [switch]$SkipReleaseGate,
    # 显式指定 Inno Setup 编译器 ISCC.exe 路径；留空自动探测（环境变量 ISCC_PATH / PATH / 常见安装目录）。
    [string]$ISCCPath = "",
    # 跳过安装器（Inno Setup）编译：仅产出 dist 分发包/zip，不生成 Setup.exe。默认编译。
    [switch]$SkipInstaller
)
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$src = Join-Path $root "src"
$dist = Join-Path $root "dist"
$toolDir = $PSScriptRoot

function Get-CurrentVersion {
    $csproj = Get-Content (Join-Path $src "ExcelAiAssistant.Addin\ExcelAiAssistant.Addin.csproj") -Raw -Encoding UTF8
    if ($csproj -match '<Version>([^<]+)</Version>') { return $Matches[1].Trim() }
    throw "csproj 中未找到 <Version>，无法确定版本号。"
}

function Get-IncrementedVersion([string]$v, [string]$part) {
    $p = $v.Split('.')
    if ($p.Count -ne 4) { throw "非法版本号：$v（应为 4 段）" }
    $idx = @{ Major = 0; Minor = 1; Patch = 2; Build = 3 }[$part]
    $p[$idx] = ([int]$p[$idx] + 1).ToString()
    # 递增某段后，其后各段清零（2.18.7.0 递增 Patch → 2.18.8.0）
    for ($i = $idx + 1; $i -lt 4; $i++) { $p[$i] = "0" }
    return ($p -join '.')
}

function Sync-VersionFiles([string]$oldVer, [string]$newVer) {
    $oldEsc = [regex]::Escape($oldVer)
    # 1) csproj：Version + InformationalVersion（保持 UTF-8 BOM，防 GBK 解析错乱）
    $csprojPath = Join-Path $src "ExcelAiAssistant.Addin\ExcelAiAssistant.Addin.csproj"
    $csproj = Get-Content $csprojPath -Raw -Encoding UTF8
    $csproj = $csproj -replace "<Version>$oldEsc</Version>", "<Version>$newVer</Version>"
    $csproj = $csproj -replace "<InformationalVersion>$oldEsc</InformationalVersion>", "<InformationalVersion>$newVer</InformationalVersion>"
    if (-not $csproj.Contains("<Version>$newVer</Version>") -or -not $csproj.Contains("<InformationalVersion>$newVer</InformationalVersion>")) {
        throw "升版失败：csproj 未找到待替换的 <Version>/<InformationalVersion>，标签格式可能漂移，请先手动核对。"
    }

    [System.IO.File]::WriteAllText($csprojPath, $csproj, (New-Object System.Text.UTF8Encoding($true)))

    # 2) README.md：替换 vX.Y.Z.W（保持 UTF-8 BOM）
    $readmePath = Join-Path $root "README.md"
    $readme = Get-Content $readmePath -Raw -Encoding UTF8
    $readme = $readme -replace "v$oldEsc", "v$newVer"
    if (-not $readme.Contains("v$newVer")) {
        throw "升版失败：README 未找到 v$oldVer（应为「v 前缀版本号」格式），格式可能漂移，请先手动同步。"
    }

    [System.IO.File]::WriteAllText($readmePath, $readme, (New-Object System.Text.UTF8Encoding($true)))

    # 3) iss：复制旧版本全部安装器脚本（轻量版 + 离线完整版等 Setup*.iss）→ 新版本，
    #    替换 AppVersion 与 PkgSrc 路径（UTF-8 无 BOM）。逐个处理，漏一个 = 版本号漂移。
    $issDir = Join-Path $toolDir "installer"
    $oldIssFiles = @(Get-ChildItem -Path $issDir -Filter "JYYJ助手-$oldVer-Setup*.iss" | Sort-Object Name)
    if ($oldIssFiles.Count -eq 0) { throw "未找到旧版本 iss 用于复制（JYYJ助手-$oldVer-Setup*.iss）" }
    foreach ($oldIss in $oldIssFiles) {
        $newIss = Join-Path $issDir ($oldIss.Name.Replace($oldVer, $newVer))
        if (Test-Path $newIss) { throw "目标 iss 已存在，拒绝覆盖：$newIss（若需重建请先删除旧文件）" }
        $iss = Get-Content $oldIss.FullName -Raw -Encoding UTF8
        $iss = $iss -replace "#define AppVersion `"[^`"]+`"", "#define AppVersion `"$newVer`""
        $iss = $iss -replace "JYYJ助手-$oldEsc-x64", "JYYJ助手-$newVer-x64"
        $iss = $iss -replace "JYYJ助手-$oldEsc-x86", "JYYJ助手-$newVer-x86"
        # 首行注释里的版本号（展示层）一并替换
        $iss = $iss -replace "v$oldEsc", "v$newVer"
        if ($iss.Contains($oldVer) -or -not $iss.Contains("#define AppVersion `"$newVer`"")) {
            throw "升版失败：iss 同步不完整（仍含旧版本 $oldVer 或 AppVersion 未替换），请检查源 iss 格式：$($oldIss.Name)"
        }

        [System.IO.File]::WriteAllText($newIss, $iss, (New-Object System.Text.UTF8Encoding($false)))
        Write-Host "已复制安装器脚本：$newIss"
    }

    # 4) docs 手册版本戳：技术手册 / 用户手册（均带版本号文件名，兼容旧英文名 fallback）的
    #    「当前版本」戳。只精确匹配「版本 X.Y.Z.W」「(Version / InformationalVersion) X」「JYYJ助手-X-Setup」
    #    三类句式——不碰 v 前缀历史条目（如「v2.19.4.0 运行日志增强」属变更史，不得改写）。
    #    容忍缺失/未命中（手册格式可能演进），提示但不 fail——版本戳漏一处的兜底由 verify-versionbump 断言。
    #    用户手册文件名带版本号（JYYJ助手用户手册-<版本>.md/.html/.pdf），动态解析取字典序最新，兼容旧名 user-manual.*。
    function Resolve-DocPath([string]$dir, [string]$pattern, [string]$fallback) {
        $hits = @(Get-ChildItem -Path $dir -Filter $pattern -File -ErrorAction SilentlyContinue | Sort-Object Name)
        if ($hits.Count -gt 0) { return $hits[-1].FullName }
        $fb = Join-Path $dir $fallback
        if (Test-Path -LiteralPath $fb) { return $fb }
        return $null
    }
    $docsDirReal = Join-Path $root "docs"
    $userMdStamp = Resolve-DocPath $docsDirReal "JYYJ助手用户手册-*.md" "user-manual.md"
    $userHtmlStamp = Resolve-DocPath $docsDirReal "JYYJ助手用户手册-*.html" "user-manual.html"
    $techMdStamp = Resolve-DocPath $docsDirReal "JYYJ助手技术手册-*.md" "technical-manual.md"
    $techHtmlStamp = Resolve-DocPath $docsDirReal "JYYJ助手技术手册-*.html" "technical-manual.html"
    $docTargets = @(
        @{ Path = $techMdStamp;    Patterns = @("版本 $oldVer", "(Version / InformationalVersion) $oldVer", "JYYJ助手-$oldVer-Setup") },
        @{ Path = $userMdStamp;    Patterns = @("版本 $oldVer") },
        @{ Path = $userHtmlStamp;  Patterns = @("版本 $oldVer") },
        @{ Path = $techHtmlStamp; Patterns = @("版本 $oldVer") }
    )
    foreach ($t in $docTargets) {
        if (-not $t.Path -or -not (Test-Path -LiteralPath $t.Path)) { Write-Host "提示：未找到 $($t.Path)，跳过手册版本戳同步。"; continue }
        $doc = Get-Content $t.Path -Raw -Encoding UTF8
        $hit = $false
        foreach ($p in $t.Patterns) {
            if ($doc.Contains($p)) { $doc = $doc.Replace($p, $p.Replace($oldVer, $newVer)); $hit = $true }
        }
        if ($hit) {
            # 保持原文件 BOM 状态（md/html 带 BOM 与否都可能，按现状写回）
            $bytes = [System.IO.File]::ReadAllBytes($t.Path)
            $hasBom = $bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF
            [System.IO.File]::WriteAllText($t.Path, $doc, (New-Object System.Text.UTF8Encoding($hasBom)))
            Write-Host "已同步手册版本戳：$($t.Path)"
        } else {
            Write-Host "提示：$($t.Path) 未发现 $oldVer 版本戳（可能已手动更新或格式变更），未改动。"
        }
    }

    # 5) 手册改名：文件名带版本号（JYYJ助手用户手册/技术手册-<版本>.*），升版后随版本戳重命名；
    #    旧名 user-manual.* / technical-manual.* 首次遇到时一并迁移到新命名。同名已存在则跳过（幂等）。
    foreach ($book in @(@{ Stem = "JYYJ助手用户手册"; Legacy = "user-manual"; Exts = @("md", "html", "pdf") },
                      @{ Stem = "JYYJ助手技术手册"; Legacy = "technical-manual"; Exts = @("md", "html") })) {
        foreach ($ext in $book.Exts) {
            $newName = "$($book.Stem)-$newVer.$ext"
            $newPath = Join-Path $docsDirReal $newName
            if (Test-Path -LiteralPath $newPath) { continue }
            $vered = Join-Path $docsDirReal "$($book.Stem)-$oldVer.$ext"
            $legacy = Join-Path $docsDirReal "$($book.Legacy).$ext"
            $src = if (Test-Path -LiteralPath $vered) { $vered } elseif (Test-Path -LiteralPath $legacy) { $legacy } else { $null }
            if ($src) {
                Move-Item -LiteralPath $src -Destination $newPath
                Write-Host "已重命名手册：$(Split-Path -Leaf $src) -> $newName"
            }
        }
    }
}

if ($BumpVersion -and $Version -ne "") { throw "-BumpVersion 与 -Version 互斥，请只使用其一。" }

if ($Version -eq "" -and $BumpVersion) {
    $cur = Get-CurrentVersion
    $Version = Get-IncrementedVersion $cur $BumpPart
    Write-Host "自动递增版本号：$cur -> $Version（$BumpPart）"
    Sync-VersionFiles $cur $Version
}

if ($Version -eq "") {
    $Version = Get-CurrentVersion
}

# 发布门禁：单元测试 + 真机自检回归。单元测试含 Sidebar 内联 JS 静态审计（踩坑 §4.1 拦截）；
# 真机回归用真实 Excel 实例验证「主工具 → 自动自检链路」全链路（含负向用例与自检覆盖审计）。
# 任一关失败（退出码非 0）直接中止打包——不让未经验证的产物出门。
# 注意：单测关跑的是 Debug 编译（dotnet test 默认），Release 编译错误由下方构建步骤兜住；
# 真机关的 harness 用 Release 构建运行。
function Invoke-UnitTestsGate {
    Write-Host "== 发布门禁 1/2：全量单元测试 =="
    dotnet test (Join-Path $root "ExcelAiAssistant.sln")
    if ($LASTEXITCODE -ne 0) { throw "单元测试未全部通过，中止打包。" }
    Write-Host "== 单元测试通过 =="
}

function Invoke-RealMachineVerify {
    Write-Host "== 发布门禁 2/2：真机自检回归（RealMachineVerify，独立 Excel 实例）=="
    $harnessDir = Join-Path $toolDir "RealMachineVerify"
    dotnet build (Join-Path $harnessDir "RealMachineVerify.csproj") -c Release
    if ($LASTEXITCODE -ne 0) { throw "真机自检 harness 构建失败，中止打包。" }
    $exe = Get-ChildItem (Join-Path $harnessDir "bin") -Filter "RealMachineVerify.exe" -Recurse |
        Where-Object { $_.FullName -like "*Release*" } | Select-Object -First 1
    if ($null -eq $exe) { throw "未找到真机自检 harness 输出 exe，中止打包。" }
    & $exe.FullName
    if ($LASTEXITCODE -ne 0) { throw "真机自检回归未全部通过（exit=$LASTEXITCODE），中止打包。" }
    Write-Host "== 真机自检回归通过 =="
}

# 探测 Inno Setup 编译器 ISCC.exe：显式 -ISCCPath 优先 → 环境变量 ISCC_PATH →
# PATH 内 iscc → 常见安装目录。找不到则 fail-loud 给出安装指引/逃生开关（升版后若漏装 Inno，
# 不能静默产出无安装器的包——现版本曾因缺这步导致 Setup.exe 缺失）。
function Find-InnoSetup {
    if (-not [string]::IsNullOrWhiteSpace($ISCCPath)) {
        if (Test-Path $ISCCPath) { return [string]$ISCCPath }
        throw "指定的 ISCC 路径不存在：$ISCCPath（用 Inno Setup 6 的 ISCC.exe）"
    }
    if (-not [string]::IsNullOrWhiteSpace($env:ISCC_PATH) -and (Test-Path $env:ISCC_PATH)) {
        return [string]$env:ISCC_PATH
    }
    $cmd = Get-Command iscc -ErrorAction SilentlyContinue
    if ($cmd) { return [string]$cmd.Source }
    foreach ($cand in @(
        "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
        "${env:ProgramFiles}\Inno Setup 6\ISCC.exe",
        "${env:ProgramFiles(x86)}\Inno Setup 5\ISCC.exe"
    )) {
        if (-not [string]::IsNullOrWhiteSpace($cand) -and (Test-Path $cand)) { return $cand }
    }
    throw "未找到 Inno Setup 编译器（ISCC.exe）。请安装 Inno Setup 6（https://jrsoftware.org/isinfo.php）后重试；或用 -ISCCPath 指定路径；紧急时用 -SkipInstaller 跳过安装器编译。"
}

# 编译安装器：各 iss 的 PkgSrc64/86 是硬编码的 dist 绝对路径，前置必须已由 Invoke-StagePackage
# 产出对应版本目录，缺则 fail-loud 而不是让 Inno 隐晦报错。编译该版本全部 Setup*.iss
# （轻量版 + 离线完整版等），任一失败即中止；编译后按 fail-closed 铁律查验产物实体。
function Invoke-InnoCompile([string]$iscc, [string]$assembleVersion) {
    $issDir = Join-Path $toolDir "installer"
    $lightIss = Join-Path $issDir "JYYJ助手-$assembleVersion-Setup.iss"
    if (-not (Test-Path $lightIss)) { throw "未找到轻量安装器脚本：$lightIss（请先运行升版同步）" }
    $issFiles = @(Get-ChildItem -Path $issDir -Filter "JYYJ助手-$assembleVersion-Setup*.iss" | Sort-Object Name)
    foreach ($arch in @("x64", "x86")) {
        $app = Join-Path $dist "JYYJ助手-$assembleVersion-$arch\app"
        if (-not (Test-Path $app)) {
            throw "安装器编译前置缺失：$app 不存在（须先产出对应版本的分发包目录）"
        }
    }
    Write-Host "== 编译安装器（Inno Setup）：共 $($issFiles.Count) 个 iss =="
    foreach ($iss in $issFiles) {
        # 产物名 = iss 文件名去扩展（与各 iss 的 OutputBaseFilename 约定一致）
        $outExe = Join-Path (Join-Path $dist "installer") ($iss.BaseName + ".exe")
        & $iscc $iss.FullName | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "Inno Setup 编译失败（exit=$LASTEXITCODE）：$($iss.FullName)" }
        # fail-closed：ISCC 返回 0 不代表产物存在（OutputDir/命名任何偏离都会静默丢产物），必须查验实体。
        if (-not (Test-Path $outExe)) { throw "Inno Setup 编译返回成功但未找到产物：$outExe（请检查 iss 的 OutputDir/OutputBaseFilename）" }
        Write-Host "== 安装器已生成：$outExe =="
    }
}

if (-not $SkipReleaseGate) {
    Invoke-UnitTestsGate
    Invoke-RealMachineVerify
} else {
    Write-Host "!! 已跳过发布门禁（-SkipReleaseGate），此包未经单测与真机验证。!!"
}

if (-not $SkipBuild) {
    Write-Host "== 构建 x64 =="
    dotnet build (Join-Path $src "ExcelAiAssistant.Addin\ExcelAiAssistant.Addin.csproj") -c $Configuration -p:Platform=x64 | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "x64 构建失败。" }
    Write-Host "== 构建 x86 =="
    dotnet build (Join-Path $src "ExcelAiAssistant.Addin\ExcelAiAssistant.Addin.csproj") -c $Configuration -p:Platform=x86 | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "x86 构建失败。" }
}

function Invoke-StagePackage([string]$Arch) {
    $out = Join-Path $src "ExcelAiAssistant.Addin\bin\$Arch\$Configuration\net48\win-$Arch"
    if (-not (Test-Path $out)) { throw "未找到构建输出：$out" }
    $pkg = Join-Path $dist "JYYJ助手-$Version-$Arch"
    $app = Join-Path $pkg "app"
    if (Test-Path $pkg) { Remove-Item $pkg -Recurse -Force }
    New-Item -ItemType Directory -Path $app -Force | Out-Null

    # 运行产物（DLL，不含 pdb；WebView2 加载器随架构一同复制）
    Get-ChildItem $out -File | Where-Object { $_.Extension -eq ".dll" } | Copy-Item -Destination $app

    # System.Data.SQLite 的原生互操作库按架构放在 x64/x86 子目录，必须随包复制，
    # 否则安装后加载 SQLite 失败、WebView2 侧边栏空白。
    foreach ($native in @("x64", "x86")) {
        $srcNative = Join-Path $out $native
        $dstNative = Join-Path $app $native
        if (Test-Path (Join-Path $srcNative "SQLite.Interop.dll")) {
            New-Item -ItemType Directory -Path $dstNative -Force | Out-Null
            Copy-Item (Join-Path $srcNative "SQLite.Interop.dll") $dstNative -Force
        }
    }

    # 本地技能
    $skills = Join-Path $src "ExcelAiAssistant.Addin\skills"
    if (Test-Path $skills) { Copy-Item $skills -Destination $app -Recurse -Force }

    # 安装/卸载脚本放到包根目录，可从包内直接执行
    Copy-Item (Join-Path $toolDir "install-xai.ps1") $pkg -Force
    Copy-Item (Join-Path $toolDir "uninstall-xai.ps1") $pkg -Force

    # 小白一键安装/卸载入口（.cmd 双击即用）
    foreach ($oneclick in @("oneclick-install.ps1", "oneclick-uninstall.ps1", "一键安装.cmd", "一键卸载.cmd")) {
        $src = Join-Path $toolDir $oneclick
        if (Test-Path $src) { Copy-Item $src $pkg -Force }
    }

    $readme = Join-Path $pkg "README.txt"
    Set-Content -Path $readme -Value @(
        "JYYJ助手 $Version（$Arch）",
        "================================",
        "小白一键安装：双击本文件夹里的【一键安装.cmd】，自动识别 Excel 位数并完成安装。",
        "安装脚本：powershell -ExecutionPolicy Bypass -File .\install-xai.ps1",
        "卸载：双击【一键卸载.cmd】 或运行 .\uninstall-xai.ps1",
        "要求：Windows 10/11、.NET Framework 4.8、Evergreen WebView2 Runtime、",
        "      Microsoft Excel（与包架构一致的位版本）。"
    ) -Encoding UTF8

    $zip = "$pkg.zip"
    if (Test-Path $zip) { Remove-Item $zip -Force }
    Compress-Archive -Path "$pkg/*" -DestinationPath $zip -CompressionLevel Optimal
    return [PSCustomObject]@{ Arch = $Arch; Package = $pkg; Zip = $zip }
}

$results = @()
$results += Invoke-StagePackage "x64"
$results += Invoke-StagePackage "x86"

if (-not $SkipInstaller) {
    $iscc = Find-InnoSetup
    Invoke-InnoCompile $iscc $Version
    $results += [PSCustomObject]@{ Arch = "installer"; Package = Join-Path $dist "installer"; Zip = "JYYJ助手-$Version-Setup.exe" }
}

Write-Host "== 部署包已生成 =="
$results | Format-Table -AutoSize
return $results