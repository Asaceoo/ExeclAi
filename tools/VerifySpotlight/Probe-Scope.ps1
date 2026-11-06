# 界定指针模式边框残留缺陷范围：JYYJ_SEL 改名导致 TRUE→FALSE 是否回收。
# 流A 单格→拖拽区域 / 流B 区域→区域 / 流C 多选区→区域
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
if (-not $ready) { exit 2 }
Write-Host "等待 25s 避开 15s 延迟自愈清理..."
Start-Sleep -Seconds 25

$wb = $app.Workbooks.Add()
for ($r = 1; $r -le 8; $r++) { for ($c = 1; $c -le 8; $c++) { $wb.Worksheets.Item(1).Cells.Item($r, $c).Value2 = $r * 100 + $c } }
Invoke-RibbonButton "指针" | Out-Null; Start-Sleep -Milliseconds 800

Write-Host "`n>> 流A：单格 B2 → 拖拽区域 D4:F6（B2 应清除）"
$app.Range("B2").Select(); Start-Sleep -Milliseconds 500
$app.Range("D4:F6").Select(); Start-Sleep -Milliseconds 600
Dump-Sel "流A"
$b = Bdr 2 2; $d = Bdr 4 4; $f = Bdr 6 6
Write-Host ("  B2=$b D4=$d F6=$f")
Assert ($b -ne 3) "流A：B2 边框清除（TRUE→FALSE 单格→区域）"
Assert ($d -eq 3) "流A：D4 红框"
Assert ($f -eq 3) "流A：F6 红框"

Write-Host "`n>> 流B：区域 A1:C3 → 拖拽区域 D4:F6（A1:C3 应清除）"
$app.Range("A1:C3").Select(); Start-Sleep -Milliseconds 600
$b1 = Bdr 1 1; $b2 = Bdr 2 2
Write-Host ("  框A1:C3 后 A1=$b1 B2=$b2")
Assert ($b1 -eq 3) "流B：A1 红框（框内）"
Assert ($b2 -eq 3) "流B：B2 红框（框内）"
$app.Range("D4:F6").Select(); Start-Sleep -Milliseconds 600
Dump-Sel "流B"
$b1 = Bdr 1 1; $b2 = Bdr 2 2; $d = Bdr 4 4
Write-Host ("  区域切换后 A1=$b1 B2=$b2 D4=$d")
Assert ($b1 -ne 3) "流B：A1 边框清除（TRUE→FALSE 区域→区域）"
Assert ($b2 -ne 3) "流B：B2 边框清除"
Assert ($d -eq 3) "流B：D4 红框"

Write-Host "`n>> 流C：多选区 A1:A3,C1:C3 → 拖拽区域 D4:F6"
$app.Range("A1:A3,C1:C3").Select(); Start-Sleep -Milliseconds 700
Dump-Sel "流C前"
$b1 = Bdr 2 1; $b2 = Bdr 2 2; $c2 = Bdr 2 3
Write-Host ("  多选区 A2=$b1 B2=$b2 C2=$c2")
Assert ($b1 -eq 3) "流C：A2 红框"
Assert ($c2 -eq 3) "流C：C2 红框"
$app.Range("D4:F6").Select(); Start-Sleep -Milliseconds 700
Dump-Sel "流C后"
$b1 = Bdr 2 1; $b2 = Bdr 2 2; $d = Bdr 4 4
Write-Host ("  多选区→区域后 A2=$b1 B2=$b2 D4=$d")
Assert ($b1 -ne 3) "流C：A2 边框清除"
Assert ($b2 -ne 3) "流C：B2 边框清除"
Assert ($d -eq 3) "流C：D4 红框"

Invoke-RibbonButton "指针" | Out-Null; Start-Sleep -Milliseconds 500
$wb.Close($false)
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null
Write-Host ""
Write-Host "== SCOPE PASSED=$script:passed FAILED=$script:failed =="
if ($script:failed -gt 0) { exit 1 } else { exit 0 }
