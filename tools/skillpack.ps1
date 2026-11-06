#skillpack.ps1 - P5-C3 skill pack packaging / import.
#   Pack:   .\skillpack.ps1 -Source .\my-skill -Output .\my-skill.xai-skill
#   Import: .\skillpack.ps1 -Import .\my-skill.xai-skill [-SkillsRoot <skills dir>]
#An .xai-skill is a zip that keeps the skill folder (with SKILL.md) as a top-level folder.

param(
    [string]$Source,
    [string]$Output,
    [string]$Import,
    [string]$SkillsRoot
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Assert-StartsWith([string]$Text, [string]$Prefix, [string]$What) {
    if (-not $Text.StartsWith($Prefix)) {
        throw "$What 路径必须形如 '$Prefix...'。"
    }
}

function Get-SafeRoot([string]$ExplicitRoot, [string]$Default) {
    if ([string]::IsNullOrWhiteSpace($ExplicitRoot)) {
        return $Default
    }
    return [System.IO.Path]::GetFullPath($ExplicitRoot)
}

function Set-UserSkillsRoot {
    param([string]$Variable)
    $env:XAI_SKILLS_ROOT = $Variable
}

function Get-DefaultUserSkillsRoot {
    # Default matches the add-in: %APPDATA%\ExcelAiAssistant\skills
    return [System.IO.Path]::Combine(
        [Environment]::GetFolderPath([Environment+SpecialFolder]::ApplicationData),
        "ExcelAiAssistant",
        "skills")
}

function Find-SkillMd([string]$Dir, [string]$What) {
    $md = [System.IO.Path]::Combine($Dir, "SKILL.md")
    if (-not (Test-Path $md -PathType Leaf)) {
        throw "$What 不是有效技能目录：缺少 $md。"
    }
    return $md
}

function Write-Summary {
    param([string[]]$Lines)
    $Lines | ForEach-Object { Write-Host $_ }
}

# ---- Pack ----
if (-not [string]::IsNullOrWhiteSpace($Source)) {
    $src = [System.IO.Path]::GetFullPath($Source)
    if (-not (Test-Path $src -PathType Container)) {
        throw "Source 目录不存在：$src"
    }
    $null = Find-SkillMd $src "Source"

    if ([string]::IsNullOrWhiteSpace($Output)) {
        $Output = [System.IO.Path]::Combine(
            [System.IO.Path]::GetDirectoryName($src),
            ([System.IO.Path]::GetFileName($src) + ".xai-skill"))
    }
    $outFull = [System.IO.Path]::GetFullPath($Output)
    $outDir = [System.IO.Path]::GetDirectoryName($outFull)
    if (-not [string]::IsNullOrWhiteSpace($outDir)) {
        [System.IO.Directory]::CreateDirectory($outDir) | Out-Null
    }

    if ((Test-Path $outFull -PathType Leaf)) {
        Remove-Item $outFull -Force
    }

    # Keep the skill folder as the single top-level folder inside the zip.
    $parent = [System.IO.Path]::GetDirectoryName($src)
    $folderName = [System.IO.Path]::GetFileName($src)
    [System.IO.Compression.ZipFile]::CreateFromDirectory(
        $src,
        $outFull,
        [System.IO.Compression.CompressionLevel]::Optimal,
        $true)
    Write-Summary @(
        "技能包已创建：$outFull"
        "来源目录：$src"
        "技能名称：$folderName"
    )
    exit 0
}

# ---- Import ----
if (-not [string]::IsNullOrWhiteSpace($Import)) {
    $pack = [System.IO.Path]::GetFullPath($Import)
    if (-not (Test-Path $pack -PathType Leaf)) {
        throw "Import 文件不存在：$pack"
    }

    $root = Get-SafeRoot $SkillsRoot (Get-DefaultUserSkillsRoot)
    [System.IO.Directory]::CreateDirectory($root) | Out-Null

    $stamp = [DateTime]::Now.ToString("yyyyMMddHHmmss")
    $tmp = [System.IO.Path]::Combine(
        [System.IO.Path]::GetTempPath(),
        ("xai-skill-" + $stamp + "-" + [System.Guid]::NewGuid().ToString("N")))
    [System.IO.Directory]::CreateDirectory($tmp) | Out-Null

    try {
        [System.IO.Compression.ZipFile]::ExtractToDirectory($pack, $tmp)

        $mdFiles = Get-ChildItem -Path $tmp -Filter "SKILL.md" -Recurse -File
        if ($null -eq $mdFiles -or $mdFiles.Count -eq 0) {
            throw "技能包中未找到 SKILL.md：不是有效技能包。"
        }

        $imported = @()
        $overwritten = @()
        foreach ($md in $mdFiles) {
            $skillDir = $md.DirectoryName
            $name = [System.IO.Path]::GetFileName($skillDir)
            $target = [System.IO.Path]::Combine($root, $name)
            $new = -not (Test-Path $target -PathType Container)
            if (-not $new) {
                $overwritten += $name
            }
            Copy-Item -Path $skillDir -Destination $target -Recurse -Force
            $imported += $name
        }

        [Environment]::SetEnvironmentVariable("XAI_SKILLS_ROOT", $root, "Process")
        $summary = @(
            "技能包导入成功：$pack"
            "技能目录：$root"
            "已导入：" + ($imported -join ", ")
        )
        if ($overwritten.Count -gt 0) {
            $summary += "已覆盖：" + ($overwritten -join ", ")
        }
        if ($imported.Count -eq 0) {
            $summary += "警告：未发现待导入技能。"
        }
        Write-Summary $summary
        exit 0
    }
    finally {
        if (Test-Path $tmp) {
            Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

Write-Host "skillpack.ps1 - 用法："
Write-Host "  [打包]   .\skillpack.ps1 -Source <skill目录> [-Output <输出.xai-skill>]"
Write-Host "  [导入]   .\skillpack.ps1 -Import <包.xai-skill> [-SkillsRoot <技能根目录>]"
Write-Host "默认技能根目录：" (Get-DefaultUserSkillsRoot)
exit 1