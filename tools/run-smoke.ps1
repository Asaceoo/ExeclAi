#requires -Version 5.1
<#
.SYNOPSIS
JYYJ助手一键真机回归冒烟。统一构建并串行运行所有验证套件，输出汇总结果。
每个套件是独立进程，启动真实 Excel，互不影响，结束后自动清理。

.EXAMPLE
# 全量运行（工具套件 + 模型/侧栏套件）
powershell -ExecutionPolicy Bypass -File .\tools\run-smoke.ps1 -Suite all

# 只跑离线工具套件（Advanced / PhaseC / PhaseE / PhaseF），不联网
powershell -ExecutionPolicy Bypass -File .\tools\run-smoke.ps1 -Suite tools

# 按名过滤运行任意子集
powershell -ExecutionPolicy Bypass -File .\tools\run-smoke.ps1 -Suite advanced,phasef
#>
[CmdletBinding()]
param(
    # 套件名：单个工具名 / 用逗号分隔多个 / 别名 all|tools|connectivity
    [string]$Suite = 'all',

    # Debug|Release
    [string]$Config = 'Release',

    # 跳过重新构建（直接用现有二进制）
    [switch]$SkipBuild,

    # 单个套件执行的超时（秒）
    [int]$TimeoutSeconds = 900
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot            # d:\execl1.0
$resultRoot = Join-Path $PSScriptRoot 'smoke-results'
New-Item -ItemType Directory -Force -Path $resultRoot | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$masterLog = Join-Path $resultRoot ("smoke-$stamp.log")

function Log([string]$msg, [string]$color = 'Gray') {
    $line = ('{0}  {1}' -f (Get-Date -Format 'HH:mm:ss.fff'), $msg)
    Write-Host $line -ForegroundColor $color
    try { Add-Content -Path $masterLog -Value $line -Encoding UTF8 } catch { }
}

function Invoke-DotNetBuild([string]$csproj) {
    Log "  [build] $csproj"
    $out = & dotnet build $csproj -c $Config -nologo -clp:ErrorsOnly 2>&1
    $code = $LASTEXITCODE
    if ($code -ne 0) {
        Log ("  [build] FAILED (exit=" + $code + ")`n" + (($out | Out-String).Trim()), 'Red')
        return $false
    }
    return $true
}

function Invoke-Suite([string]$exePath, [int]$timeoutSec) {
    # 异步读取标准输出/错误，避免死锁与后台线程事件回调在 PS5.1 的不安全；可可靠获取退出码
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $exePath
    $psi.WorkingDirectory = Split-Path -Parent $exePath
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
    $psi.StandardErrorEncoding = [System.Text.Encoding]::UTF8
    $psi.CreateNoWindow = $true
    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    $null = $proc.Start()
    $outTask = $proc.StandardOutput.ReadToEndAsync()
    $errTask = $proc.StandardError.ReadToEndAsync()
    if (-not $proc.WaitForExit($timeoutSec * 1000)) {
        try { $proc.Kill() } catch { }
        $proc.WaitForExit()
        return @{ Exit = -1; TimedOut = $true; Lines = @('TIMEOUT killed') }
    }
    $proc.WaitForExit()
    $lines = New-Object System.Collections.Generic.List[string]
    $out = $outTask.GetAwaiter().GetResult()
    $err = $errTask.GetAwaiter().GetResult()
    if ($out) { $out -split "`r?`n" | ForEach-Object { if ($_.Length -gt 0) { $lines.Add($_) } } }
    if ($err) { $err -split "`r?`n" | ForEach-Object { if ($_.Length -gt 0) { $lines.Add('[ERR] ' + $_) } } }
    return @{ Exit = $proc.ExitCode; TimedOut = $false; Lines = $lines.ToArray() }
}

function Get-RunnableExe([string]$projDir, [string]$name) {
    # 输出路径各项目不一致（bin\{Config}\net48\win-x64 或 bin\x64\Release），统一取最近一次构建产物
    $exe = Get-ChildItem -Path (Join-Path $projDir 'bin') -Recurse -Filter "$name.exe" -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    return $exe.FullName
}

# ---------- 套件定义 ----------
$suites = [ordered]@{
    'advanced' = @{ Name='VerifyAdvancedTools'; Cs='tools\VerifyAdvancedTools\VerifyAdvancedTools.csproj'; Group='tools'; Desc='一二三期全工具冒烟' }
    'phasec'   = @{ Name='VerifyPhaseC';        Cs='tools\VerifyPhaseC\VerifyPhaseC.csproj';        Group='tools'; Desc='四期-C 数据分析与报告' }
    'phasee'   = @{ Name='VerifyPhaseE';        Cs='tools\VerifyPhaseE\VerifyPhaseE.csproj';        Group='tools'; Desc='四期-E 稳定性(降采样/取消/忙重试)' }
    'phasef'   = @{ Name='VerifyPhaseF';        Cs='tools\VerifyPhaseF\VerifyPhaseF.csproj';        Group='tools'; Desc='实机 版本/历史/事件卡' }
    'phaseg'   = @{ Name='VerifyPhaseG';        Cs='tools\VerifyPhaseG\VerifyPhaseG.csproj';        Group='tools'; Desc='五期-A 会话管理/长期记忆' }
    'phaseh'   = @{ Name='VerifyPhaseH';        Cs='tools\VerifyPhaseH\VerifyPhaseH.csproj';        Group='tools'; Desc='五期-B 流程编排/执行/重试' }
    'phasei'   = @{ Name='VerifyPhaseI';        Cs='tools\VerifyPhaseI\VerifyPhaseI.csproj';        Group='tools'; Desc='五期-C 多工作簿/技能包/策略共享' }
    'phasej'   = @{ Name='VerifyPhaseJ';        Cs='tools\VerifyPhaseJ\VerifyPhaseJ.csproj';        Group='tools'; Desc='五期-C1 真机 跨工作簿只读端到端' }
    'conn'     = @{ Name='VerifySidebar';       Cs='tools\VerifySidebar\VerifySidebar.csproj';      Group='conn'; Desc='模型配置与连通(需联网)' }
    'chat'     = @{ Name='VerifySidebarChat';   Cs='tools\VerifySidebarChat\VerifySidebarChat.csproj'; Group='conn'; Desc='WebView2 侧栏端到端(需联网)' }
}

# 解析套件选择到键列表
$keys = @()
$parts = ($Suite -replace '\s', '').Split(',') | Where-Object { $_ }
foreach ($p in $parts) {
    switch ($p.ToLower()) {
        'all'         { $keys = @($suites.Keys); break }
        'tools'       { $keys += @($suites.Keys | Where-Object { $suites[$_].Group -eq 'tools' }) }
        'connectivity'{ $keys += @($suites.Keys | Where-Object { $suites[$_].Group -eq 'conn' }) }
        default       { if ($suites.Contains($p)) { $keys += $p } else { Log "忽略未知套件: $p" 'Yellow' } }
    }
}
$keys = $keys | Select-Object -Unique
if ($keys.Count -eq 0) { Log '未选择任何套件，退出。' 'Red'; exit 2 }

Log ("==================== XAI 一键冒烟 开始 ({0}) ====================" -f $Suite)
Log ("-> 将运行: " + ($keys -join ', '))
Log ("-> 构建: " + ($(if ($SkipBuild) {'跳过'} else {$Config})))

$summary = [System.Collections.Generic.List[object]]::new()
$global:anyFail = $false

foreach ($key in $keys) {
    $def = $suites[$key]
    $csproj = Join-Path $repo $def.Cs
    $exePath = $null

    Log ("==== 套件 [{0}] {1}：{2} ====" -f $key, $def.Name, $def.Desc)

    if (-not $SkipBuild) {
        if (-not (Test-Path $csproj)) { Log "  项目不存在: $csproj" 'Red'; $global:anyFail = $true; $summary.Add([pscustomobject]@{Key=$key;Name=$def.Name;Status='NO-PROJECT';Detail=$csproj;Sec=0}); continue }
        if (-not (Invoke-DotNetBuild $csproj)) { $global:anyFail = $true; $summary.Add([pscustomobject]@{Key=$key;Name=$def.Name;Status='BUILD-FAIL';Detail=$def.Cs;Sec=0}); continue }
    }

    $exePath = Get-RunnableExe (Split-Path -Parent $csproj) $def.Name
    if (-not $exePath -or -not (Test-Path $exePath)) {
        Log "  未找到已编译产物: $def.Name (可先去除 -SkipBuild)" 'Red'
        $global:anyFail = $true
        $summary.Add([pscustomobject]@{Key=$key;Name=$def.Name;Status='NO-EXE';Detail=$def.Cs;Sec=0})
        continue
    }
    Log ("  run: $exePath")

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $logPath = Join-Path $resultRoot ("$key-$stamp.log")
    $run = $null
    try {
        $run = Invoke-Suite $exePath $TimeoutSeconds
    } catch {
        $sw.Stop()
        $global:anyFail = $true
        $summary.Add([pscustomobject]@{Key=$key;Name=$def.Name;Status='RUN-ERROR';Detail=$_.Exception.Message;Sec=[int]$sw.Elapsed.TotalSeconds})
        Log ("  启动失败: " + $_.Exception.Message) 'Red'
        continue
    }
    $sw.Stop()

    # 落盘原始输出
    try { [System.IO.File]::WriteAllLines($logPath, $run.Lines, [System.Text.Encoding]::UTF8) } catch { }

    if ($run.TimedOut) {
        Log ("  TIMEOUT (>{0}s) 已终止" -f $TimeoutSeconds) 'Red'
        $global:anyFail = $true
        $summary.Add([pscustomobject]@{Key=$key;Name=$def.Name;Status='TIMEOUT';Detail=">${TimeoutSeconds}s";Sec=[int]$sw.Elapsed.TotalSeconds})
        if ($run.Lines.Count -gt 0) { Log ("    tail: " + ($run.Lines[-1..([Math]::Max(0, $run.Lines.Count-3))] -join ' | ')) 'Red' }
        continue
    }

    $tail = ''
    if ($run.Lines.Count -gt 0) { $tail = ($run.Lines[-1..([Math]::Max(0, $run.Lines.Count-3))] -join ' | ') }
    $exit = $run.Exit
    if ($exit -eq 0) {
        Log ("  PASS  exit=0  ({0:N1}s)" -f $sw.Elapsed.TotalSeconds) 'Green'
        Log ("    tail: " + ($tail -replace "`n", ' | ')) 'DarkGray'
        $summary.Add([pscustomobject]@{Key=$key;Name=$def.Name;Status='PASS';Detail="exit=0";Sec=[int]$sw.Elapsed.TotalSeconds})
    } else {
        Log ("  FAIL  exit=$exit  ({0:N1}s)" -f $sw.Elapsed.TotalSeconds) 'Red'
        if ($tail) { Log ("    tail: " + ($tail -replace "`n", ' | ')) 'Red' }
        $global:anyFail = $true
        $summary.Add([pscustomobject]@{Key=$key;Name=$def.Name;Status='FAIL';Detail="exit=$exit";Sec=[int]$sw.Elapsed.TotalSeconds})
    }
}

# ---------- 汇总 ----------
Log ""
Log ("==================== 汇总 ====================")
$total = $summary.Count
$passCount = @($summary | Where-Object { $_.Status -eq 'PASS' }).Count
foreach ($row in $summary) {
    $st = $row.Status
    $color = if ($st -eq 'PASS') { 'Green' } elseif ($st -eq 'FAIL') { 'Red' } else { 'Yellow' }
    Log ("  [{0,-11}] {1,-22} {2,-12} {3,5}s  {4}" -f $row.Key, $row.Name, $st, $row.Sec, $row.Detail) $color
}
Log ("总览: $passCount/$total 通过")
Log ("日志目录: $resultRoot")
Log ""

if ($global:anyFail) { Log 'SMOKE RESULT: FAIL（存在失败项，详见各套件日志）' 'Red'; exit 1 }
Log 'SMOKE RESULT: ALL PASS' 'Green'
exit 0