# 复现指针开/关流程（先激活 JYYJ助手 Tab）并抓取加载项 DBG 日志
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes

$app = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")

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
            try { $tp = $b.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern); return ($tp.Current.ToggleState -eq [System.Windows.Automation.ToggleState]::On) } catch { return "no-toggle-pattern" }
        }
    }
    return "not-found"
}

Select-RibbonTab "JYYJ助手" | Out-Null
Start-Sleep -Seconds 1

$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$app.Range("B2").Select()
Start-Sleep -Milliseconds 500

Write-Host "pointer state before: $(Get-RibbonButtonState '指针')"
Write-Host "click 指针 ON"
Invoke-RibbonButton "指针" | Out-Null
Start-Sleep -Seconds 2
Write-Host "pointer state after ON: $(Get-RibbonButtonState '指针')"
Write-Host "click 指针 OFF"
Invoke-RibbonButton "指针" | Out-Null
Start-Sleep -Seconds 2
Write-Host "pointer state after OFF: $(Get-RibbonButtonState '指针')"

$n = 0
for ($i = 1; $i -le $ws.Cells.FormatConditions.Count; $i++) {
    $f = [string]$ws.Cells.FormatConditions.Item($i).Formula1
    Write-Host "  leftover rule[$i] Formula1=$f"
    if ($f -like '*JYYJ_VIEW_MODE*') { $n++ }
}
Write-Host "leftover sentinel rules=$n"

$wb.Close($false)
Write-Host "DONE"
