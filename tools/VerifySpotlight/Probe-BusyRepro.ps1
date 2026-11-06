# 忙态误判复现探针：开启聚光灯后模拟用户高频点选（制造忙态），观测
#   1) CF 哨兵规则是否累积 >1（完整性校验误判 → 清缓存 → 重复建规则）
#   2) 日志是否新增 rule delete failed / update selection failed（忙态被误判为缺失）
#   3) Ctrl 多选后 JYYJ_SEL 是否正确写成 OR(...) 且高亮稳定
# 走真实加载项（安装副本 2.1.0.19）：自持 COM + UIA 点 Ribbon。结束关模式、关工作簿。
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes

$script:passed = 0; $script:failed = 0
function Assert([bool]$cond, [string]$msg) {
    if ($cond) { $script:passed++; Write-Host "[PASS] $msg" }
    else { $script:failed++; Write-Host "[FAIL] $msg" }
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
function Get-SelNameRefersTo($wb) {
    try { return [string]$wb.Names.Item("JYYJ_SEL").RefersTo } catch { return $null }
}

$logPath = Join-Path $env:TEMP "ExcelAiAssistant.log"
function Get-LogLines { if (Test-Path $logPath) { @(Get-Content $logPath).Count } else { 0 } }
function Get-LogNewFailures([int]$startLine) {
    if (-not (Test-Path $logPath)) { return @() }
    $all = @(Get-Content $logPath)
    $new = $all | Select-Object -Skip $startLine
    return @($new | Where-Object { $_ -match "rule delete failed|update selection failed|EnsureRule CREATE|rule deleted" })
}

$app = $null
for ($i = 0; $i -lt 15; $i++) { try { $app = New-Object -ComObject Excel.Application; break } catch { Start-Sleep -Seconds 1 } }
if ($null -eq $app) { Write-Host "NO EXCEL"; exit 2 }
$app.Visible = $true; $app.DisplayAlerts = $false

Write-Host "等待加载项 Ribbon 就绪…"
$ready = $false
for ($i = 0; $i -lt 90; $i++) { Start-Sleep -Seconds 1; if (Select-RibbonTab "JYYJ助手") { $ready = $true; break } }
Assert $ready "加载项「JYYJ助手」选项卡出现"

$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
for ($r = 1; $r -le 15; $r++) { for ($c = 1; $c -le 15; $c++) { $ws.Cells.Item($r, $c).Value2 = $r * 100 + $c } }

# ---------- 开启聚光灯 ----------
$s0 = Get-RibbonButtonState "聚光灯"
if ($s0 -eq $true) { Invoke-RibbonButton "聚光灯" | Out-Null; Start-Sleep -Milliseconds 400 }
Assert (Invoke-RibbonButton "聚光灯") "点击「聚光灯」"
Start-Sleep -Milliseconds 1000
Assert ((Get-RibbonButtonState "聚光灯") -eq $true) "聚光灯已开启"
Assert ((SentinelRuleCount $ws) -eq 1) "初始哨兵 CF 规则 = 1"

# ---------- 高频点选 + 制造忙态（密集外部自动化竞争加载项 tick） ----------
Write-Host "`n=== 高频点选 15 秒（密集外部自动化制造忙态）==="
$log0 = Get-LogLines
$maxRules = 1
$busyErr = 0
for ($i = 0; $i -lt 1000; $i++) {
    $r = 1 + ($i % 12); $c = 1 + (($i / 12) % 12)
    try { $null = $ws.Cells.Item($r, $c).Select() } catch { }
    try { $v = $ws.Cells.Item($r, $c).DisplayFormat.Interior.ColorIndex } catch { }
    # 密集读 CF（跨进程自动化）→ 撞加载项完整性校验 tick → 制造 0x800AC472 忙态窗口
    try {
        $fcs = $ws.Cells.FormatConditions
        $fcnt = $fcs.Count
        $n = 0
        for ($k = 1; $k -le $fcnt; $k++) {
            try { if ([string]$fcs.Item($k).Formula1 -like '*JYYJ_VIEW_MODE*') { $n++ } } catch { }
        }
        if ($n -gt $maxRules) { $maxRules = $n }
    } catch { $busyErr++ }
    Start-Sleep -Milliseconds 15
}
Assert ($maxRules -eq 1) "15 秒密集自动化期间 CF 规则从未累积（max=$maxRules）"
$newContent = Get-LogNewFailures $log0
Write-Host ("  新增失败/重建日志条数: {0}（外部自动化忙态异常 {1} 次）" -f $newContent.Count, $busyErr)
foreach ($f in $newContent) { Write-Host ("  LOG: " + $f.Substring(0, [Math]::Min(110, $f.Length))) }
Assert ($newContent.Count -eq 0) "密集自动化期间无完整性误判重建（rule delete failed / update selection failed / 重复 CREATE）"

# ---------- Ctrl 多选两个区域 ----------
Write-Host "`n=== 多选两个区域 ==="
$null = $app.Range("B2:D4,G6:I8").Select()
Start-Sleep -Milliseconds 600
$refers = Get-SelNameRefersTo $wb
Write-Host ("  JYYJ_SEL = " + $refers)
Assert ($refers -like '*OR(*') "多选后 JYYJ_SEL 已写成 OR(...) 表达式"
Assert ((SentinelRuleCount $ws) -eq 1) "多选后 CF 规则仍为 1 条（未累积）"
Start-Sleep -Seconds 2   # 跨过 2 个完整性校验周期
Assert ((SentinelRuleCount $ws) -eq 1) "多选后跨 2s 完整性校验 CF 规则仍为 1 条"
$refers2 = Get-SelNameRefersTo $wb
Assert ($refers2 -eq $refers) "多选后跨 2s JYYJ_SEL 未被重置（=$refers2）"

# ---------- 单格点选即时性 ----------
Write-Host "`n=== 单格点选响应 ==="
$null = $app.Range("A1").Select(); Start-Sleep -Milliseconds 200
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$null = $app.Range("C5").Select()
$ok = $false
for ($t = 0; $t -lt 100; $t++) {
    try { if ([int]$ws.Cells.Item(5, 3).DisplayFormat.Interior.ColorIndex -eq 6) { $ok = $true; break } } catch { }
    Start-Sleep -Milliseconds 15
}
$sw.Stop()
Assert $ok ("单格点选 C5 十字跟随 <1.5s（实际 " + $sw.ElapsedMilliseconds + "ms）")

# ---------- 清理 ----------
Invoke-RibbonButton "聚光灯" | Out-Null; Start-Sleep -Milliseconds 500
Assert ((SentinelRuleCount $ws) -eq 0) "关模式后哨兵 CF 规则已清"
$wb.Close($false)
$app.Quit()
try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch { }

Write-Host ""
Write-Host "== 忙态复现探针 PASSED=$script:passed FAILED=$script:failed =="
if ($script:failed -gt 0) { exit 1 } else { exit 0 }
