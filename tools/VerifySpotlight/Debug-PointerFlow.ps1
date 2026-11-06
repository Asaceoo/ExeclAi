# 调试版：指针开关流程中每一步转储 FormatConditions 状态，定位检查与加载项清理为何不一致。
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

function Dump-FC([string]$tag, $ws, $wb) {
    Write-Host "=== [$tag] ==="
    Write-Host "  ws.Name=$($ws.Name) parent=$($ws.Parent.Name)"
    $cnt = 0
    try { $cnt = [int]$ws.Cells.FormatConditions.Count } catch { Write-Host "  Count ERR: $($_.Exception.Message)"; return }
    Write-Host "  FormatConditions.Count=$cnt"
    for ($i = 1; $i -le $cnt; $i++) {
        try {
            $f = [string]$ws.Cells.FormatConditions.Item($i).Formula1
            Write-Host "    rule[$i] Type=$([int]$ws.Cells.FormatConditions.Item($i).Type) Formula1=$f"
        } catch { Write-Host "    rule[$i] ERR: $($_.Exception.Message)" }
    }
    $nameExists = $false
    for ($i = 1; $i -le $wb.Names.Count; $i++) { if ([string]$wb.Names.Item($i).Name -eq 'JYYJ_VIEW_MODE') { $nameExists = $true } }
    Write-Host "  nameExists=$nameExists"
}

$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$app.Range("B2").Select()
Select-RibbonTab "JYYJ助手" | Out-Null
Start-Sleep -Seconds 1

Dump-FC "初始" $ws $wb

Write-Host ">> 点击聚光灯 ON"
Invoke-RibbonButton "聚光灯" | Out-Null
Start-Sleep -Seconds 1
Dump-FC "聚光灯ON后" $ws $wb

Write-Host ">> 点击指针 ON"
Invoke-RibbonButton "指针" | Out-Null
Start-Sleep -Seconds 1
Dump-FC "指针ON后(互斥)" $ws $wb

Write-Host ">> 点击指针 OFF"
Invoke-RibbonButton "指针" | Out-Null
Start-Sleep -Milliseconds 1500
Dump-FC "指针OFF后1.5s" $ws $wb
Start-Sleep -Seconds 2
Dump-FC "指针OFF后3.5s" $ws $wb

$wb.Close($false)
Write-Host "DONE"
