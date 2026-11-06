# 探针：哨兵定义名守卫 —— 用户已有同名定义 JYYJ_VIEW_MODE=2 时：
#  1) 不得覆盖用户定义（仍为 =2）
#  2) 不得创建哨兵 CF 规则（视图模式在该工作簿静默禁用，避免引用用户定义造成残留）
#  3) 关闭后无任何残留
# 前置：Excel 已启动并加载新版加载项。
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

function Get-ButtonState([string]$name) {
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
    foreach ($t in $tabs) {
        if ($t.Current.Name -eq $name) { $t.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select(); return $true }
    }
    return $false
}

function Get-NameRefersTo($wb, [string]$nm) {
    for ($i = 1; $i -le $wb.Names.Count; $i++) {
        if ([string]$wb.Names.Item($i).Name -eq $nm) { return [string]$wb.Names.Item($i).RefersTo }
    }
    return $null
}

function Get-SentinelRuleCount($ws) {
    $n = 0
    for ($i = 1; $i -le $ws.Cells.FormatConditions.Count; $i++) {
        try { $f = [string]$ws.Cells.FormatConditions.Item($i).Formula1 } catch { continue }
        if ($f -like '*JYYJ_VIEW_MODE*') { $n++ }
    }
    return $n
}

$app = $null
for ($i = 0; $i -lt 40; $i++) {
    try { $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application"); break } catch { Start-Sleep -Seconds 1 }
}
if ($null -eq $app) { Write-Host "NO EXCEL INSTANCE"; exit 2 }

# 复位聚光灯/指针
foreach ($btn in @("聚光灯", "指针")) {
    $s = Get-ButtonState $btn
    if ($s -eq $true) { Invoke-RibbonButton $btn | Out-Null; Start-Sleep -Milliseconds 500 }
}

$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$app.Range("B2").Select()
Start-Sleep -Milliseconds 300

# 预置同名用户定义 =2
$wb.Names.Add('JYYJ_VIEW_MODE', '=2') | Out-Null
Assert ((Get-NameRefersTo $wb 'JYYJ_VIEW_MODE') -eq '=2') "预置：用户定义 JYYJ_VIEW_MODE==2"

Select-RibbonTab "JYYJ助手" | Out-Null
Start-Sleep -Milliseconds 500
Assert (Invoke-RibbonButton "聚光灯") "点击「聚光灯」按钮"
Start-Sleep -Milliseconds 1000

Assert ((Get-NameRefersTo $wb 'JYYJ_VIEW_MODE') -eq '=2') "守卫：用户定义未被覆盖（仍 ==2）"
Assert ((Get-SentinelRuleCount $ws) -eq 0) "守卫：未创建哨兵 CF 规则（该工作簿静默禁用）"

# 关闭聚光灯，确认无残留
Assert (Invoke-RibbonButton "聚光灯") "再次点击「聚光灯」关闭"
Start-Sleep -Milliseconds 800
Assert ((Get-NameRefersTo $wb 'JYYJ_VIEW_MODE') -eq '=2') "关闭：用户定义仍保留"
Assert ((Get-SentinelRuleCount $ws) -eq 0) "关闭：无哨兵规则残留"

$wb.Close($false)
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws) | Out-Null
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) | Out-Null

Write-Host ""
Write-Host "== 哨兵守卫 PASSED=$script:passed FAILED=$script:failed =="
if ($script:failed -gt 0) { Write-Host "SENTINEL-GUARD FAILED"; exit 1 } else { Write-Host "ALL SENTINEL-GUARD CHECKS OK"; exit 0 }
