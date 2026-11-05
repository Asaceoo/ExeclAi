# 聚焦探针：定位剩余 2 项失败（多选指针 B2 误亮 / 拖拽 A1:E500 不渲染）。
# 走真实加载项 + 活动表断言，转储 JYYJ_SEL 与 FormatConditions 拿到真实状态。
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
function Fill([int]$r, [int]$c) {
    for ($t = 0; $t -lt 30; $t++) {
        try { return [int]$app.ActiveSheet.Cells.Item($r, $c).DisplayFormat.Interior.ColorIndex } catch { Start-Sleep -Milliseconds 15 }
    }
    return -9999
}
function Bdr([int]$r, [int]$c) {
    for ($t = 0; $t -lt 30; $t++) {
        try { return [int]$app.ActiveSheet.Cells.Item($r, $c).DisplayFormat.Borders.Item(7).ColorIndex } catch { Start-Sleep -Milliseconds 15 }
    }
    return -9999
}
function Dump-State([string]$tag) {
    $wb = $app.ActiveWorkbook
    Write-Host "=== [$tag] sel=$($app.Selection.Address($false,$false)) active=$($app.ActiveCell.Address($false,$false)) ==="
    for ($i = 1; $i -le $wb.Names.Count; $i++) {
        $n = $wb.Names.Item($i); $nm = [string]$n.Name
        if ($nm -like 'JYYJ_*') { Write-Host "  name $nm = $([string]$n.RefersTo)" }
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($n) | Out-Null
    }
    $cnt = [int]$app.ActiveSheet.Cells.FormatConditions.Count
    Write-Host "  FormatConditions.Count=$cnt"
    for ($i = 1; $i -le $cnt; $i++) {
        $f = $app.ActiveSheet.Cells.FormatConditions.Item($i)
        Write-Host "    rule[$i] Type=$([int]$f.Type) Formula1=$([string]$f.Formula1)"
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($f) | Out-Null
    }
    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null
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
for ($r = 1; $r -le 500; $r++) { for ($c = 1; $c -le 5; $c++) { $wb.Worksheets.Item(1).Cells.Item($r, $c).Value2 = $r * 1000 + $c } }

Write-Host "`n>> 聚光灯 ON，单格 B2（基线）"
Invoke-RibbonButton "聚光灯" | Out-Null; Start-Sleep -Milliseconds 800
$app.Range("B2").Select(); Start-Sleep -Milliseconds 500
Assert ((Fill 2 2) -eq 6) "单格 B2 黄=6"
Assert ((Fill 2 1) -eq 6) "B 列十字 A2 黄=6"
Assert ((Fill 1 1) -ne 6) "A1 非黄"

Write-Host "`n>> 聚光灯，拖拽 A1:E500"
$app.Range("A1:E500").Select(); Start-Sleep -Milliseconds 800
Dump-State "聚光灯 A1:E500"
Assert ((Fill 250 3) -eq 6) "(250,3) 黄（行内/列内）"
Assert ((Fill 250 10) -eq 6) "(250,10) 黄（行十字全列宽）"
Assert ((Fill 501 1) -ne 6) "(501,1) 非黄（行外）"

Write-Host "`n>> 指针 ON（互斥切），多选 A1:A3,C1:C3"
Invoke-RibbonButton "指针" | Out-Null; Start-Sleep -Milliseconds 800
$app.Range("A1:A3,C1:C3").Select(); Start-Sleep -Milliseconds 800
Dump-State "指针 A1:A3,C1:C3"
Assert ((Bdr 2 1) -eq 3) "多选 A2 有框"
Assert ((Bdr 2 3) -eq 3) "多选 C2 有框"
Assert ((Bdr 2 2) -ne 3) "多选 B2 无框"

Write-Host "`n>> 指针，再次同选区（重复渲染稳定性）"
$app.Range("A1:A3,C1:C3").Select(); Start-Sleep -Milliseconds 800
Dump-State "指针 重复多选"
Assert ((Bdr 2 2) -ne 3) "重复多选 B2 仍无框"

Invoke-RibbonButton "指针" | Out-Null; Start-Sleep -Milliseconds 400
$wb.Close($false)
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null

Write-Host ""
Write-Host "== LASTTWO PASSED=$script:passed FAILED=$script:failed =="
if ($script:failed -gt 0) { exit 1 } else { exit 0 }
