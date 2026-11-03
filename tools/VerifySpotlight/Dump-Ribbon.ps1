# UIA 转储：Excel 窗口内的 Tab 与按钮（排查 聚光灯/指针 为何不可见）
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$root = [System.Windows.Automation.AutomationElement]::RootElement
$cond = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ClassNameProperty, "XLMAIN")
$win = $root.FindFirst([System.Windows.Automation.TreeScope]::Children, $cond)
if (-not $win) { Write-Host "NO XLMAIN window"; exit 1 }
Write-Host "Window: $($win.Current.Name)"

$tabCond = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::TabItem)
$tabs = $win.FindAll([System.Windows.Automation.TreeScope]::Descendants, $tabCond)
Write-Host "=== Tabs ($($tabs.Count)) ==="
foreach ($t in $tabs) { Write-Host "  Tab: '$($t.Current.Name)'" }

$btnCond = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Button)
$buttons = $win.FindAll([System.Windows.Automation.TreeScope]::Descendants, $btnCond)
Write-Host "=== Buttons ($($buttons.Count)) containing 聚光灯/指针/面板 ==="
$found = 0
foreach ($b in $buttons) {
    if ($b.Current.Name -match '聚光灯|指针|面板|助手') { Write-Host "  Btn: '$($b.Current.Name)'"; $found++ }
}
Write-Host "  matched buttons: $found / total: $($buttons.Count)"
