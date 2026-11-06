# 自愈探针：模式开启时外部删除 JYYJ_SEL 定义名 / 哨兵 CF 规则 → 服务应自动重建并恢复高亮。
# 走真实加载项：自持 COM 引用 + UIA 点 Ribbon。结束后关模式、关测试工作簿。
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes

$script:passed = 0; $script:failed = 0
function Assert([bool]$cond, [string]$msg) {
    if ($cond) { $script:passed++; Write-Host "[PASS] $msg" }
    else { $script:failed++; Write-Host "[FAIL] $msg" }
}
function AssertFill($ws, [int]$r, [int]$c, [bool]$want, [string]$msg, [int]$maxPolls = 130) {
    $v = -9999
    for ($t = 0; $t -lt $maxPolls; $t++) {
        try { $v = [int]$ws.Cells.Item($r, $c).DisplayFormat.Interior.ColorIndex } catch { $v = -9999 }
        if (($want -and $v -eq 6) -or ((-not $want) -and $v -ne 6)) { Assert $true $msg; return }
        Start-Sleep -Milliseconds 15
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
for ($r = 1; $r -le 12; $r++) { for ($c = 1; $c -le 12; $c++) { $ws.Cells.Item($r, $c).Value2 = $r * 100 + $c } }

# ---------- 开启聚光灯 ----------
$s0 = Get-RibbonButtonState "聚光灯"
if ($s0 -eq $true) { Invoke-RibbonButton "聚光灯" | Out-Null; Start-Sleep -Milliseconds 400 }
Assert (Invoke-RibbonButton "聚光灯") "点击「聚光灯」"
Start-Sleep -Milliseconds 800
Assert ((Get-RibbonButtonState "聚光灯") -eq $true) "聚光灯已开启"
$null = $app.Range("B2").Select()
AssertFill $ws 2 1 $true "初始 B2：(2,1) 黄"
Assert (NameExists $wb "JYYJ_SEL") "初始 JYYJ_SEL 定义名存在"

# ---------- 场景1：外部删除 JYYJ_SEL 定义名 → 自动重建 ----------
Write-Host "`n=== 自愈场景1：删除 JYYJ_SEL 定义名 ==="
$app.Names.Item("JYYJ_SEL").Delete()
Assert (-not (NameExists $wb "JYYJ_SEL")) "JYYJ_SEL 已删（模拟外部删除）"
$null = $app.Range("D4").Select()
AssertFill $ws 4 1 $true "删名后点 D4：(4,1) 黄（自动重建恢复）" 300
$null = $app.Range("F6").Select()
AssertFill $ws 6 1 $true "删名后点 F6：(6,1) 黄（持续跟随）" 300
Assert (NameExists $wb "JYYJ_SEL") "JYYJ_SEL 已被自动重建"
Assert (NameExists $wb "JYYJ_VIEW_MODE") "哨兵 JYYJ_VIEW_MODE 存在"
Assert ((SentinelRuleCount $ws) -eq 1) "哨兵 CF 规则仅 1 条（无重复）"

# ---------- 场景2：外部删除哨兵 CF 规则 → 自动重建 ----------
Write-Host "`n=== 自愈场景2：删除哨兵 CF 规则 ==="
$fcs = $ws.Cells.FormatConditions
for ($i = $fcs.Count; $i -ge 1; $i--) {
    try { if ([string]$fcs.Item($i).Formula1 -like '*JYYJ_VIEW_MODE*') { $fcs.Item($i).Delete() } } catch { }
}
Assert ((SentinelRuleCount $ws) -eq 0) "哨兵 CF 规则已删（模拟外部删除）"
$null = $app.Range("H8").Select()
AssertFill $ws 8 1 $true "删规则后点 H8：(8,1) 黄（自动重建恢复）" 300
Assert ((SentinelRuleCount $ws) -eq 1) "哨兵 CF 规则已自动重建（仅 1 条）"

# ---------- 清理 ----------
Invoke-RibbonButton "聚光灯" | Out-Null; Start-Sleep -Milliseconds 500
Assert ((SentinelRuleCount $ws) -eq 0) "关模式后哨兵 CF 规则已清"
Assert (-not (NameExists $wb "JYYJ_SEL")) "关模式后 JYYJ_SEL 已删"
$wb.Close($false)
$app.Quit()
try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch { }

Write-Host ""
Write-Host "== 自愈探针 PASSED=$script:passed FAILED=$script:failed =="
if ($script:failed -gt 0) { exit 1 } else { exit 0 }
