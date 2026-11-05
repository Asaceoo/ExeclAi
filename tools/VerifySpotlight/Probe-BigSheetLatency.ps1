# Probe-BigSheetLatency.ps1 — 大表（30000 行）单选跟随延迟真机实测
param([int]$WaitForRibbon = 90)
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
function Measure-Latency($ws, [int]$r, [int]$c, [string]$label) {
    $neutralR = 1; $neutralC = 1
    $null = $app.Cells.Item($neutralR, $neutralC).Select()
    Start-Sleep -Milliseconds 100
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $null = $app.Cells.Item($r, $c).Select()
    $found = $false
    for ($t = 0; $t -lt 400; $t++) {
        try { if ([int]$ws.Cells.Item($r, $c).DisplayFormat.Interior.ColorIndex -eq 6) { $found = $true; break } } catch { }
        Start-Sleep -Milliseconds 5
    }
    $sw.Stop()
    $ms = [int]$sw.ElapsedMilliseconds
    Assert $found ("$label 选中 ($r,$c) 跟随变黄（$ms ms）")
    return $ms
}

$app = $null
for ($i = 0; $i -lt 15; $i++) { try { $app = New-Object -ComObject Excel.Application; break } catch { Start-Sleep -Seconds 1 } }
if ($null -eq $app) { Write-Host "NO EXCEL"; exit 2 }
$app.Visible = $true; $app.DisplayAlerts = $false

Write-Host "等待加载项 Ribbon 就绪…"
$ready = $false
for ($i = 0; $i -lt $WaitForRibbon; $i++) { Start-Sleep -Seconds 1; if (Select-RibbonTab "JYYJ助手") { $ready = $true; break } }
Assert $ready "加载项「JYYJ助手」选项卡出现"

$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$ws.Name = "big30000"
# 填 30000 行 x 15 列（用逐格写太慢，用公式/填充减少 COM；这里用 30000 行 x 5 列足够撑大 UsedRange）
for ($c = 1; $c -le 5; $c++) {
    try { $ws.Cells.Item(1, $c).Value2 = 1; $null = $ws.Range($ws.Cells.Item(1,$c), $ws.Cells.Item(30000,$c)).DataSeries() } catch { }
}
# 确认 UsedRange 已扩展到 30000 行
try { $used = $ws.UsedRange.Rows.Count; Write-Host ("  UsedRange 行数: " + $used) } catch { }

$s0 = Get-RibbonButtonState "聚光灯"
if ($s0 -eq $true) { Invoke-RibbonButton "聚光灯" | Out-Null; Start-Sleep -Milliseconds 400 }
Assert (Invoke-RibbonButton "聚光灯") "点击「聚光灯」"
Start-Sleep -Milliseconds 800
Assert ((Get-RibbonButtonState "聚光灯") -eq $true) "聚光灯已开启"

Write-Host "`n=== 大表 30000 行单选跟随延迟 ==="
$null = $app.Range("B2").Select()
Start-Sleep -Milliseconds 300
Assert ((SentinelRuleCount $ws) -eq 1) "哨兵 CF 规则 = 1"
$lats = @()
# 目标格必须在数据列(1-5)内：CF 应用范围已缩小到 UsedRange，数据区外点选无十字（设计使然）
foreach ($tgt in @(@(500,3), @(5000,4), @(15000,2), @(25000,5), @(29000,1))) {
    $lats += Measure-Latency $ws $tgt[0] $tgt[1] "30000行"
}
$avg = [Math]::Round(($lats | Measure-Object -Average).Average, 0)
Write-Host ("  30000 行平均跟随延迟: $avg ms（原始序列: " + ($lats -join ",") + "）")
Assert ($avg -le 400) "30000 行大表单选平均跟随 ≤400ms（实际 $avg ms）"

Invoke-RibbonButton "聚光灯" | Out-Null; Start-Sleep -Milliseconds 500
Assert ((SentinelRuleCount $ws) -eq 0) "关模式后哨兵 CF 规则已清"
$wb.Close($false)
$app.Quit()
try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch { }

Write-Host ""
Write-Host "== 大表单选延迟探针 PASSED=$script:passed FAILED=$script:failed =="
if ($script:failed -gt 0) { exit 1 } else { exit 0 }
