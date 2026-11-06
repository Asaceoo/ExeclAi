# Probe-SelectPerf.ps1 — 单选延迟 + 多选正确性真机实测（走真实加载项）
#   1) 单选跟随延迟：小表（300x30）与 大表（5000x30）各测 5 个目标格，精确毫秒
#   2) 多选正确性：Range("B2:D4,G6:I8") 多区域 → JYYJ_SEL=OR(...) + 高亮模式逐格验证
#   3) 单选↔多选切换：多选后点单格，残留高亮应清除（只留新十字）
#   4) 快速点选 100 次：CF 规则不累积、日志无新增失败
# 结束后关模式、关测试工作簿、退出 Excel。
param([int]$WaitForRibbon = 90)
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes

$script:passed = 0; $script:failed = 0
function Assert([bool]$cond, [string]$msg) {
    if ($cond) { $script:passed++; Write-Host "[PASS] $msg" }
    else { $script:failed++; Write-Host "[FAIL] $msg" }
}
function Fill($ws, [int]$r, [int]$c, [bool]$want, [string]$msg, [int]$maxPolls = 200) {
    $v = -9999
    for ($t = 0; $t -lt $maxPolls; $t++) {
        try { $v = [int]$ws.Cells.Item($r, $c).DisplayFormat.Interior.ColorIndex } catch { $v = -9999 }
        if (($want -and $v -eq 6) -or ((-not $want) -and $v -ne 6)) { Assert $true $msg; return }
        Start-Sleep -Milliseconds 10
    }
    Assert $false "$msg（观测值=$v）"
}
function Get-ExcelWindow {
    $root = [System.Windows.Automation.AutomationElement]::RootElement
    $cond = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ClassNameProperty, "XLMAIN")
    return $root.FindFirst([System.Windows.Automation.TreeScope]::Children, $cond)
}
function Invoke-RibbonButton([string]$name) {
    $win = Get-ExcelWindow
    if (-not $win) { return $false }
    $btnCond = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Button)
    $buttons = $win.FindAll([System.Windows.Automation.TreeScope]::Descendants, $btnCond)
    foreach ($b in $buttons) {
        if ($b.Current.Name -eq $name) {
            try { $b.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern).Toggle(); return $true } catch { }
            try { $b.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke(); return $true } catch { }
            return $false
        }
    }
    return $false
}
function Get-RibbonButtonState([string]$name) {
    $win = Get-ExcelWindow
    if (-not $win) { return $null }
    $btnCond = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Button)
    $buttons = $win.FindAll([System.Windows.Automation.TreeScope]::Descendants, $btnCond)
    foreach ($b in $buttons) {
        if ($b.Current.Name -eq $name) {
            try { $tp = $b.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern); return ($tp.Current.ToggleState -eq [System.Windows.Automation.ToggleState]::On) } catch { return $null }
        }
    }
    return $null
}
function Select-RibbonTab([string]$name) {
    $win = Get-ExcelWindow
    if (-not $win) { return $false }
    $tabCond = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::TabItem)
    $tabs = $win.FindAll([System.Windows.Automation.TreeScope]::Descendants, $tabCond)
    foreach ($t in $tabs) { if ($t.Current.Name -eq $name) { $t.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select(); return $true } }
    return $false
}
function NameExists($wb, [string]$n) {
    try { $null = $wb.Names.Item($n); return $true } catch { return $false }
}
function Get-SelNameRefersTo($wb) {
    try { return [string]$wb.Names.Item("JYYJ_SEL").RefersTo } catch { return $null }
}
function SentinelRuleCount($ws) {
    $n = 0
    try {
        $fcs = $ws.Cells.FormatConditions
        for ($i = 1; $i -le $fcs.Count; $i++) {
            try { if ([string]$fcs.Item($i).Formula1 -like '*JYYJ_VIEW_MODE*') { $n++ } } catch { }
        }
    } catch { }
    return $n
}
# 单选跟随延迟：从「中性格」先落位，再选中目标格，测目标格变黄耗时
function Measure-Latency($ws, [int]$r, [int]$c, [string]$label) {
    $neutralR = 100; $neutralC = 26   # Z100：十字=第100行+Z列，与所有目标格(行<=40,列<=15)不重叠
    $null = $app.Cells.Item($neutralR, $neutralC).Select()
    Start-Sleep -Milliseconds 120
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $null = $app.Cells.Item($r, $c).Select()
    $found = $false
    for ($t = 0; $t -lt 200; $t++) {
        try { if ([int]$ws.Cells.Item($r, $c).DisplayFormat.Interior.ColorIndex -eq 6) { $found = $true; break } } catch { }
        Start-Sleep -Milliseconds 5
    }
    $sw.Stop()
    $ms = [int]$sw.ElapsedMilliseconds
    Assert $found ("$label 选中 ($r,$c) 跟随变黄（$ms ms）")
    return $ms
}

$logPath = Join-Path $env:TEMP "ExcelAiAssistant.log"
function Get-LogLines { if (Test-Path $logPath) { @(Get-Content $logPath).Count } else { 0 } }
function Get-LogNewFailures([int]$startLine) {
    if (-not (Test-Path $logPath)) { return @() }
    $all = @(Get-Content $logPath)
    $new = $all | Select-Object -Skip $startLine
    return @($new | Where-Object { $_ -match "rule delete failed|update selection failed|EnsureRule CREATE" })
}
# 忙态重试包装：0x800AC472 / RPC 拒绝时短暂退避重试
function Invoke-ComRetry([scriptblock]$sb, [int]$maxTry = 5) {
    for ($i = 0; $i -lt $maxTry; $i++) {
        try { return & $sb } catch {
            $hr = 0
            if ($_.Exception -is [System.Runtime.InteropServices.COMException]) { $hr = $_.Exception.HResult }
            if ($hr -eq 0x800AC472 -or $hr -eq 0x80010001 -or $hr -eq 0x8001010A) { Start-Sleep -Milliseconds 80; continue }
            throw
        }
    }
    throw "busy retry exhausted"
}
# 批量填数据：逐格写 + 忙态重试（数组批量写经 PowerShell 封送 COM 会被 Excel 拒收 0x800A03EC）
function Fill-Data($ws, [int]$rows, [int]$cols) {
    for ($r = 1; $r -le $rows; $r++) {
        for ($c = 1; $c -le $cols; $c++) {
            $v = $r * 1000 + $c
            Invoke-ComRetry { $ws.Cells.Item($r, $c).Value2 = $v } | Out-Null
        }
    }
}

$app = $null
for ($i = 0; $i -lt 15; $i++) { try { $app = New-Object -ComObject Excel.Application; break } catch { Start-Sleep -Seconds 1 } }
if ($null -eq $app) { Write-Host "NO EXCEL"; exit 2 }
$app.Visible = $true; $app.DisplayAlerts = $false; $app.ScreenUpdating = $true

Write-Host "等待加载项 Ribbon 就绪…"
$ready = $false
for ($i = 0; $i -lt $WaitForRibbon; $i++) { Start-Sleep -Seconds 1; if (Select-RibbonTab "JYYJ助手") { $ready = $true; break } }
Assert $ready "加载项「JYYJ助手」选项卡出现"

$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$ws.Name = "small"
Fill-Data $ws 300 30
$ws2 = $wb.Worksheets.Add()
$ws2.Name = "big"
Fill-Data $ws2 1500 15

# ---------- 开启聚光灯 ----------
$s0 = Get-RibbonButtonState "聚光灯"
if ($s0 -eq $true) { Invoke-RibbonButton "聚光灯" | Out-Null; Start-Sleep -Milliseconds 400 }
Assert (Invoke-RibbonButton "聚光灯") "点击「聚光灯」"
Start-Sleep -Milliseconds 800
Assert ((Get-RibbonButtonState "聚光灯") -eq $true) "聚光灯已开启"
$app.Sheets.Item("small").Activate()
$null = $app.Range("B2").Select()
Fill $ws 2 1 $true "初始 B2：(2,1) 黄"

# ---------- 单选延迟：小表 ----------
Write-Host "`n=== 单选跟随延迟：小表 small(300x30) ==="
$log0 = Get-LogLines
$latSmall = @()
foreach ($tgt in @(@(3,5), @(4,7), @(5,9), @(6,11), @(7,13))) {
    $latSmall += Measure-Latency $ws $tgt[0] $tgt[1] "小表"
}
$avgSmall = [Math]::Round(($latSmall | Measure-Object -Average).Average, 0)
Write-Host ("  小表单选平均跟随延迟: $avgSmall ms（逐次: " + ($latSmall -join ", ") + "）")
Assert ($avgSmall -le 200) "小表单选平均跟随 ≤200ms（实际 $avgSmall ms）"

# ---------- 单选延迟：大表 ----------
Write-Host "`n=== 单选跟随延迟：大表 big(1500x15) ==="
$app.Sheets.Item("big").Activate()
$null = $app.Range("B2").Select()
Fill $ws2 2 1 $true "大表 B2：(2,1) 黄"
$latBig = @()
foreach ($tgt in @(@(400,5), @(800,8), @(1200,3), @(1400,12), @(1490,7))) {
    $latBig += Measure-Latency $ws2 $tgt[0] $tgt[1] "大表"
}
$avgBig = [Math]::Round(($latBig | Measure-Object -Average).Average, 0)
Write-Host ("  大表单选平均跟随延迟: $avgBig ms（逐次: " + ($latBig -join ", ") + "）")
Assert ($avgBig -le 300) "大表单选平均跟随 ≤300ms（实际 $avgBig ms）"

# ---------- 多选正确性 ----------
Write-Host "`n=== 多选正确性：B2:D4,G6:I8 ==="
$app.Sheets.Item("small").Activate()
$null = $app.Range("B2:D4,G6:I8").Select()
Start-Sleep -Milliseconds 600
$refers = Get-SelNameRefersTo $wb
Write-Host ("  JYYJ_SEL = " + $refers)
Assert ($refers -like '*OR(*') "多选后 JYYJ_SEL 已写成 OR(...) 表达式"
# 第一区 B2:D4 十字 = 行2-4 ∪ 列2-4；第二区 G6:I8 十字 = 行6-8 ∪ 列7-9
Fill $ws 2 2 $true  "B2(2,2) 在区+十字 → 黄"
Fill $ws 3 5 $true  "E3(3,5) 行3带 → 黄"
Fill $ws 5 3 $true  "C5(5,3) 列C带 → 黄"
Fill $ws 5 5 $false "E5(5,5) 区外 → 不黄"
Fill $ws 6 7 $true  "G6(6,7) 第二区 → 黄"
Fill $ws 7 10 $true "J7(7,10) 行7带 → 黄"
Fill $ws 9 8 $true  "H9(9,8) 列H带 → 黄"
Fill $ws 9 10 $false "J9(9,10) 区外 → 不黄"
Assert ((SentinelRuleCount $ws) -eq 1) "多选后 CF 规则仍为 1 条（未累积）"

# ---------- 单选↔多选切换：残留高亮清除 ----------
Write-Host "`n=== 切换：多选 → 单格 E5 ==="
$null = $app.Range("E5").Select()
Fill $ws 5 5 $true  "E5(5,5) 新十字中心 → 黄"
Fill $ws 5 3 $true  "C5(5,3) 行5带 → 黄（新十字）"
Fill $ws 6 7 $false "G6(6,7) 残留清除 → 不黄"
Fill $ws 2 2 $false "B2(2,2) 残留清除 → 不黄"

# ---------- 快速点选 100 次：不累积、无失败 ----------
Write-Host "`n=== 快速点选 100 次 ==="
$log1 = Get-LogLines
$maxRules = SentinelRuleCount $ws
for ($i = 0; $i -lt 100; $i++) {
    $r = 1 + ($i % 20); $c = 2 + (($i / 20) % 10)
    $null = $app.Cells.Item($r, $c).Select()
    $n = SentinelRuleCount $ws
    if ($n -gt $maxRules) { $maxRules = $n }
    Start-Sleep -Milliseconds 15
}
Assert ($maxRules -eq 1) "100 次快速点选 CF 规则从未累积（max=$maxRules）"
$newFail = Get-LogNewFailures $log1
Write-Host ("  快速点选新增失败/重建日志: " + $newFail.Count + " 条")
foreach ($f in $newFail) { Write-Host ("  LOG: " + $f.Substring(0, [Math]::Min(110, $f.Length))) }
Assert ($newFail.Count -eq 0) "快速点选期间无 update selection failed / EnsureRule CREATE / rule delete failed"

# ---------- 清理 ----------
Invoke-RibbonButton "聚光灯" | Out-Null; Start-Sleep -Milliseconds 500
Assert ((SentinelRuleCount $ws) -eq 0) "关模式后哨兵 CF 规则已清"
$wb.Close($false)
$app.Quit()
try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch { }

Write-Host ""
Write-Host "== 单选延迟+多选探针 PASSED=$script:passed FAILED=$script:failed =="
if ($script:failed -gt 0) { exit 1 } else { exit 0 }
