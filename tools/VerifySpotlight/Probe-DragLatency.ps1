# 决定性探针：聚光灯下大选区拖拽后，十字实际渲染延迟测量（含 JYYJ_SEL 确认）。
# 目的：区分「服务不渲染（缺陷）」与「Excel 忙态渲染慢（固有成本）」。
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
function CFColorRaw($ws, [int]$r, [int]$c) {
    try { return [int]$ws.Cells.Item($r, $c).DisplayFormat.Interior.ColorIndex } catch { return -9999 }
}

$app = $null
for ($i = 0; $i -lt 15; $i++) { try { $app = New-Object -ComObject Excel.Application; break } catch { Start-Sleep -Seconds 1 } }
if ($null -eq $app) { Write-Host "NO EXCEL"; exit 2 }
$app.Visible = $true; $app.DisplayAlerts = $false

Write-Host "等待加载项 Ribbon 就绪…"
$ready = $false
for ($i = 0; $i -lt 90; $i++) { Start-Sleep -Seconds 1; if (Select-RibbonTab "JYYJ助手") { $ready = $true; break } }
if (-not $ready) { Write-Host "RIBBON NOT READY"; exit 3 }

$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
for ($r = 1; $r -le 500; $r++) { for ($c = 1; $c -le 5; $c++) { $ws.Cells.Item($r, $c).Value2 = $r * 1000 + $c } }

$s0 = Get-RibbonButtonState "聚光灯"
if ($s0 -eq $true) { Invoke-RibbonButton "聚光灯" | Out-Null; Start-Sleep -Milliseconds 400 }
Invoke-RibbonButton "聚光灯" | Out-Null
Start-Sleep -Milliseconds 800

function Measure-Drag([string]$label, [string]$addr, [int]$sampleR, [int]$sampleC) {
    Write-Host "`n>> $label select=$addr"
    $app.Range($addr).Select()
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $v = -9999
    while ($sw.ElapsedMilliseconds -lt 8000) {
        $v = CFColorRaw $ws $sampleR $sampleC
        if ($v -eq 6) { break }
        Start-Sleep -Milliseconds 50
    }
    $sw.Stop()
    $sel = "?"
    try { $sel = [string]$wb.Names.Item('JYYJ_SEL').RefersTo } catch { }
    Write-Host ("  (${sampleR},${sampleC}) 渲染耗时=" + $sw.ElapsedMilliseconds + "ms 终值=$v")
    Write-Host ("  JYYJ_SEL=$sel")
    if ($v -eq 6) { Write-Host "  [PASS] $label 十字渲染（$($sw.ElapsedMilliseconds)ms）" }
    else { Write-Host "  [FAIL] $label 十字未渲染" }
}

# 场景A：干净状态直接拖拽大选区（无数据填充后的忙态）
Measure-Drag "场景A 干净拖拽 A1:E500" "A1:E500" 250 3

# 场景B：紧接场景A 后再次拖不同位置（热状态）
Measure-Drag "场景B 热状态拖拽 A10:F510" "A10:F510" 260 4

# 场景C：小选区对照（50 行）
Measure-Drag "场景C 小选区 A1:E50" "A1:E50" 25 3

$ws.Cells.FormatConditions.Delete()
Invoke-RibbonButton "聚光灯" | Out-Null; Start-Sleep -Milliseconds 400
$wb.Close($false)
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws) | Out-Null
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null
Write-Host ""
Write-Host "== DRAGLATENCY DONE =="
