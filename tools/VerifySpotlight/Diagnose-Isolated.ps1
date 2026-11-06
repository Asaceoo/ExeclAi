# 隔离诊断：避开 15s 延迟清理后，逐场景转储 JYYJ_SEL/规则/渲染层，定位指针+整行+手动模式根因。
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes

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
function Dump([string]$tag, $ws, $wb) {
    Write-Host "=== [$tag] ==="
    Write-Host "  active=$($ws.Application.ActiveCell.Address) sel=$($ws.Application.Selection.Address)"
    for ($i = 1; $i -le $wb.Names.Count; $i++) {
        $n = $wb.Names.Item($i); $nm = [string]$n.Name
        if ($nm -like 'JYYJ_*') { Write-Host "  name $nm = $([string]$n.RefersTo)" }
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($n) | Out-Null
    }
    $cnt = [int]$ws.Cells.FormatConditions.Count
    Write-Host "  FormatConditions.Count=$cnt"
    for ($i = 1; $i -le $cnt; $i++) {
        $f = $ws.Cells.FormatConditions.Item($i)
        Write-Host "    rule[$i] Type=$([int]$f.Type) Formula1=$([string]$f.Formula1)"
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($f) | Out-Null
    }
    try {
        $c = [int]$ws.Cells.Item(5,3).DisplayFormat.Interior.ColorIndex
        $b = [int]$ws.Cells.Item(5,3).DisplayFormat.Borders.Item(7).ColorIndex
        Write-Host "  (5,3) fill=$c border=$b"
    } catch { Write-Host "  (5,3) read ERR: $($_.Exception.Message)" }
    try {
        $c = [int]$ws.Cells.Item(2,2).DisplayFormat.Interior.ColorIndex
        $b = [int]$ws.Cells.Item(2,2).DisplayFormat.Borders.Item(7).ColorIndex
        Write-Host "  (2,2) fill=$c border=$b"
    } catch { Write-Host "  (2,2) read ERR: $($_.Exception.Message)" }
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
Write-Host "延迟清理窗口已过"

$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$ws2 = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))

Write-Host "`n>> 聚光灯 ON，单格 B2"
Invoke-RibbonButton "聚光灯" | Out-Null; Start-Sleep -Milliseconds 800
$app.Range("B2").Select(); Start-Sleep -Milliseconds 500
Dump "聚光灯 B2" $ws $wb

Write-Host "`n>> 聚光灯，整行 5"
$app.Rows.Item(5).Select(); Start-Sleep -Milliseconds 600
Dump "聚光灯 整行5" $ws $wb

Write-Host "`n>> 聚光灯，拖拽 A1:E500"
$app.Range("A1:E500").Select(); Start-Sleep -Milliseconds 600
Dump "聚光灯 A1:E500" $ws $wb

Write-Host "`n>> 指针 ON（互斥切），单格 B2"
Invoke-RibbonButton "指针" | Out-Null; Start-Sleep -Milliseconds 800
$app.Range("B2").Select(); Start-Sleep -Milliseconds 500
Dump "指针 B2" $ws $wb

Write-Host "`n>> 指针，区域 A1:C3"
$app.Range("A1:C3").Select(); Start-Sleep -Milliseconds 600
Dump "指针 A1:C3" $ws $wb

Write-Host "`n>> 手动计算模式，聚光灯，单格 D4→F6"
Invoke-RibbonButton "指针" | Out-Null; Start-Sleep -Milliseconds 400
Invoke-RibbonButton "聚光灯" | Out-Null; Start-Sleep -Milliseconds 600
$app.Calculation = -4135
$app.Range("D4").Select(); Start-Sleep -Milliseconds 500
Dump "手动 D4" $ws $wb
$app.Range("F6").Select(); Start-Sleep -Milliseconds 500
Dump "手动 F6" $ws $wb
$app.Calculation = -4105

Write-Host "`n>> 关闭聚光灯"
Invoke-RibbonButton "聚光灯" | Out-Null; Start-Sleep -Milliseconds 800
Dump "关闭后" $ws $wb

$wb.Close($false)
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws2) | Out-Null
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws) | Out-Null
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null
Write-Host "DONE"
