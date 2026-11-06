# 综合确认探针：等延迟清理窗口过后，读活动表，转储选区结构 + 断言聚光灯/整行/指针/手动模式。
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
function Select-RibbonTab([string]$name) {
    $win = Get-ExcelWindow; if (-not $win) { return $false }
    $tabCond = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::TabItem)
    $tabs = $win.FindAll([System.Windows.Automation.TreeScope]::Descendants, $tabCond)
    foreach ($t in $tabs) { if ($t.Current.Name -eq $name) { $t.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select(); return $true } }
    return $false
}
function Invoke-RibbonButton([string]$name) {
    $win = Get-ExcelWindow; if (-not $win) { return $false }
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
    $win = Get-ExcelWindow; if (-not $win) { return $null }
    $btnCond = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Button)
    $buttons = $win.FindAll([System.Windows.Automation.TreeScope]::Descendants, $btnCond)
    foreach ($b in $buttons) {
        if ($b.Current.Name -eq $name) {
            try { $tp = $b.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern); return ($tp.Current.ToggleState -eq [System.Windows.Automation.ToggleState]::On) } catch { return $null }
        }
    }
    return $null
}
function CFColor([int]$r, [int]$c) {
    for ($t = 0; $t -lt 30; $t++) {
        try { return [int]$app.ActiveSheet.Cells.Item($r, $c).DisplayFormat.Interior.ColorIndex } catch { Start-Sleep -Milliseconds 15 }
    }
    return -9999
}
function CFBorder([int]$r, [int]$c, [int]$edge) {
    for ($t = 0; $t -lt 30; $t++) {
        try { return [int]$app.ActiveSheet.Cells.Item($r, $c).DisplayFormat.Borders.Item($edge).ColorIndex } catch { Start-Sleep -Milliseconds 15 }
    }
    return -9999
}
function Get-SelectionInfo([string]$tag) {
    $sel = $app.Selection
    $addr = $sel.Address($false, $false)
    $a = $sel.Areas.Item(1)
    Write-Host "  [$tag] sel=$addr areas=$($sel.Areas.Count) area1.Row=$($a.Row) Col=$($a.Column) Rows=$($a.Rows.Count) Cols=$($a.Columns.Count) active=$($app.ActiveCell.Address($false,$false))"
}

$app = $null
for ($i = 0; $i -lt 15; $i++) { try { $app = New-Object -ComObject Excel.Application; break } catch { Start-Sleep -Seconds 1 } }
if ($null -eq $app) { Write-Host "NO EXCEL"; exit 2 }
$app.Visible = $true; $app.DisplayAlerts = $false

$ready = $false
for ($i = 0; $i -lt 90; $i++) { Start-Sleep -Seconds 1; if (Select-RibbonTab "JYYJ助手") { $ready = $true; break } }
Write-Host "ribbon ready=$ready"
if (-not $ready) { exit 2 }

Write-Host "等待 25s 避开 15s 延迟自愈清理..."
Start-Sleep -Seconds 25

$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$app.Worksheets.Item(1).Select(); Start-Sleep -Milliseconds 300

Write-Host "`n== 选区结构转储 =="
$app.Range("B2").Select(); Start-Sleep -Milliseconds 300
Get-SelectionInfo "B2单格"
$app.Rows.Item(5).Select(); Start-Sleep -Milliseconds 300
Get-SelectionInfo "Rows.Item(5)"
$app.Columns.Item(1).Select(); Start-Sleep -Milliseconds 300
Get-SelectionInfo "Columns.Item(1)"
$app.Range("A1:E500").Select(); Start-Sleep -Milliseconds 300
Get-SelectionInfo "A1:E500"

Write-Host "`n== 聚光灯（读活动表）=="
$app.Range("B2").Select(); Start-Sleep -Milliseconds 300
Invoke-RibbonButton "聚光灯" | Out-Null; Start-Sleep -Milliseconds 800
Assert ((Get-RibbonButtonState "聚光灯") -eq $true) "聚光灯已开启"
Assert ((CFColor 2 1) -eq 6) "单格 B2：A2 黄=6"
Assert ((CFColor 2 2) -eq 6) "单格 B2：B2 黄=6"
Assert ((CFColor 3 3) -ne 6) "单格 B2：C3 非黄"

$app.Rows.Item(5).Select(); Start-Sleep -Milliseconds 600
Get-SelectionInfo "聚光灯 整行5"
$sel5 = $app.Selection; $a5 = $sel5.Areas.Item(1)
$fullRow5 = (([int]$a5.Rows.Count) -ge 1048576)
if ($fullRow5) {
    Assert ((CFColor 5 3) -eq 6) "整行5：(5,3) 黄"
    Assert ((CFColor 5 10) -eq 6) "整行5：(5,10) 黄"
    Assert ((CFColor 1 1) -ne 6) "整行5：A1 非黄"
} else {
    Write-Host "  [SKIP] Rows.Item(5) 未产生整行选区（Rows.Count=$($a5.Rows.Count)），整行断言跳过"
    Assert ((CFColor 5 3) -eq 6) "整行5（按选区）(5,3) 黄"
}

Write-Host "`n== 指针（读活动表）=="
$app.Range("B2").Select(); Start-Sleep -Milliseconds 300
Invoke-RibbonButton "指针" | Out-Null; Start-Sleep -Milliseconds 800
Assert ((Get-RibbonButtonState "指针") -eq $true) "指针已开启（互斥）"
Assert ((CFBorder 2 2 7) -eq 3) "指针单格 B2：左边框红=3"
Assert ((CFBorder 3 3 7) -ne 3) "指针单格 B2：C3 无框"

$app.Range("A1:C3").Select(); Start-Sleep -Milliseconds 600
Assert ((CFBorder 2 2 7) -eq 3) "指针区域 A1:C3：B2 有框"
Assert ((CFBorder 4 2 7) -ne 3) "指针区域 A1:C3：B4 无框"

Write-Host "`n== 手动计算模式（读活动表）=="
Invoke-RibbonButton "指针" | Out-Null; Start-Sleep -Milliseconds 400
Invoke-RibbonButton "聚光灯" | Out-Null; Start-Sleep -Milliseconds 600
$app.Calculation = -4135
$app.Range("D4").Select(); Start-Sleep -Milliseconds 500
$c1 = CFColor 4 1
$app.Range("F6").Select(); Start-Sleep -Milliseconds 500
$c2 = CFColor 6 1
Write-Host ("  手动模式 D4→(4,1)=$c1  F6→(6,1)=$c2")
Assert (($c1 -eq 6) -and ($c2 -eq 6)) "手动模式单格十字跟随"
$app.Calculation = -4105
Start-Sleep -Milliseconds 300

Invoke-RibbonButton "聚光灯" | Out-Null; Start-Sleep -Milliseconds 600
$wb.Close($false)
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws) | Out-Null
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null

Write-Host ""
Write-Host "== CONFIRMED PASSED=$script:passed FAILED=$script:failed =="
if ($script:failed -gt 0) { exit 1 } else { exit 0 }
