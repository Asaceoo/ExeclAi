# 加载项真实流程：指针模式单格→单格切换，B2 边框是否随选择移动而清除（严重性判定）。
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
function Bdr([int]$r, [int]$c) {
    for ($t = 0; $t -lt 30; $t++) {
        try { return [int]$app.ActiveSheet.Cells.Item($r, $c).DisplayFormat.Borders.Item(7).ColorIndex } catch { Start-Sleep -Milliseconds 15 }
    }
    return -9999
}
function Dump-Sel([string]$tag) {
    $wb = $app.ActiveWorkbook
    for ($i = 1; $i -le $wb.Names.Count; $i++) {
        $n = $wb.Names.Item($i); $nm = [string]$n.Name
        if ($nm -eq 'JYYJ_SEL') { Write-Host "  [$tag] JYYJ_SEL = $([string]$n.RefersTo)" }
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($n) | Out-Null
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
for ($r = 1; $r -le 6; $r++) { for ($c = 1; $c -le 6; $c++) { $wb.Worksheets.Item(1).Cells.Item($r, $c).Value2 = $r * 100 + $c } }

Write-Host "`n>> 指针 ON，单格 B2"
Invoke-RibbonButton "指针" | Out-Null; Start-Sleep -Milliseconds 800
$app.Range("B2").Select(); Start-Sleep -Milliseconds 500
$b = Bdr 2 2; $d = Bdr 4 4
Write-Host ("  B2=$b D4=$d")
Assert ($b -eq 3) "B2 红框"
Assert ($d -ne 3) "D4 无框"

Write-Host "`n>> 指针，单格 B2 → D4（边框应移动）"
$app.Range("D4").Select(); Start-Sleep -Milliseconds 600
$b = Bdr 2 2; $d = Bdr 4 4
Write-Host ("  B2=$b D4=$d")
Assert ($b -ne 3) "B2 边框已清除（关键）"
Assert ($d -eq 3) "D4 红框"

Write-Host "`n>> 指针，单格 D4 → F6"
$app.Range("F6").Select(); Start-Sleep -Milliseconds 600
$b = Bdr 2 2; $d = Bdr 4 4; $f = Bdr 6 6
Write-Host ("  B2=$b D4=$d F6=$f")
Assert ($d -ne 3) "D4 边框已清除"
Assert ($f -eq 3) "F6 红框"

Write-Host "`n>> 指针，F6 → 回 B2"
$app.Range("B2").Select(); Start-Sleep -Milliseconds 600
$b = Bdr 2 2; $f = Bdr 6 6
Write-Host ("  B2=$b F6=$f")
Assert ($b -eq 3) "B2 红框（回来）"
Assert ($f -ne 3) "F6 边框已清除"

Invoke-RibbonButton "指针" | Out-Null; Start-Sleep -Milliseconds 500
$b = Bdr 2 2
Write-Host ("  关指针后 B2=$b")
Assert ($b -ne 3) "关指针后 B2 无框（全局清理）"

$wb.Close($false)
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null
Write-Host ""
Write-Host "== FLOW PASSED=$script:passed FAILED=$script:failed =="
if ($script:failed -gt 0) { exit 1 } else { exit 0 }
