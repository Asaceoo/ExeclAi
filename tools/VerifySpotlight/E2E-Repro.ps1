# 端到端复现两个用户症状：①单格点击响应延迟 ②多选是否全区域高亮。
# 走真实加载项路径：自持 COM 引用（绕开 GetActiveObject ROT 故障）→ 使可见 →
#   UIA 点 Ribbon「聚光灯」→ COM 驱动选区变化 → 读 CF 颜色/计时。
# 结束后关聚光灯、关测试工作簿；不退出 Excel（可能为共享实例）。
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
function Get-RibbonButtonState([string]$name) {
    $win = Get-ExcelWindow
    if (-not $win) { return $null }
    $btnCond = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Button)
    $buttons = $win.FindAll([System.Windows.Automation.TreeScope]::Descendants, $btnCond)
    foreach ($b in $buttons) {
        if ($b.Current.Name -eq $name) {
            try {
                $tp = $b.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern)
                return ($tp.Current.ToggleState -eq [System.Windows.Automation.ToggleState]::On)
            } catch { return $null }
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

# CF 颜色读取（吸收忙态重试）
function CFC($ws, [int]$r, [int]$c) {
    for ($t = 0; $t -lt 30; $t++) {
        try { return [int]$ws.Cells.Item($r, $c).DisplayFormat.Interior.ColorIndex } catch { Start-Sleep -Milliseconds 15 }
    }
    return -9999
}
# 测量：从 Select() 返回后到 CF 颜色变化的时间（轮询读色，粗测但反映真实感知）
function Measure-SelectLatency($ws, [string]$addr, [int]$probeR, [int]$probeC) {
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $app.Range($addr).Select()
    $sw.Stop()
    $selectMs = $sw.ElapsedMilliseconds
    $sw2 = [System.Diagnostics.Stopwatch]::StartNew()
    $color = -9999
    while ($sw2.ElapsedMilliseconds -lt 2000) {
        $color = CFC $ws $probeR $probeC
        if ($color -eq 6) { break }
        Start-Sleep -Milliseconds 5
    }
    $sw2.Stop()
    return @{ SelectMs = $selectMs; CFMs = $sw2.ElapsedMilliseconds; Color = $color }
}

# --- 连接（自持引用，绕开 ROT）---
$app = $null
for ($i = 0; $i -lt 15; $i++) {
    try { $app = New-Object -ComObject Excel.Application; break } catch { Start-Sleep -Seconds 1 }
}
if ($null -eq $app) { Write-Host "NO EXCEL"; exit 2 }
$app.Visible = $true
$app.DisplayAlerts = $false

# --- 等加载项 Ribbon 就绪（最多 90s）---
Write-Host "等待加载项 Ribbon 就绪…"
$ready = $false
for ($i = 0; $i -lt 90; $i++) {
    Start-Sleep -Seconds 1
    if (Select-RibbonTab "JYYJ助手") { $ready = $true; break }
}
Assert $ready "加载项「JYYJ助手」选项卡出现"

# 准备工作簿
$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
for ($r = 1; $r -le 12; $r++) { for ($c = 1; $c -le 12; $c++) { $ws.Cells.Item($r, $c).Value2 = $r * 100 + $c } }

# 开聚光灯
Start-Sleep -Milliseconds 300
$s0 = Get-RibbonButtonState "聚光灯"
if ($s0 -eq $true) { Invoke-RibbonButton "聚光灯" | Out-Null; Start-Sleep -Milliseconds 400 }
Assert (Invoke-RibbonButton "聚光灯") "点击「聚光灯」按钮"
Start-Sleep -Milliseconds 800
Assert ((Get-RibbonButtonState "聚光灯") -eq $true) "聚光灯已开启"

# --- 症状1：单格点击延迟（连点 5 格，测每格 CF 跟随时延） ---
Write-Host "`n--- 症状1：单格点击响应 ---"
$lat = Measure-SelectLatency $ws "B2" 2 1
Write-Host ("  B2: Select=" + $lat.SelectMs + "ms CF跟随时延=" + $lat.CFMs + "ms 色=" + $lat.Color)
Assert ($lat.Color -eq 6) "B2 高亮"
$lat = Measure-SelectLatency $ws "D5" 5 1
Write-Host ("  D5: Select=" + $lat.SelectMs + "ms CF跟随时延=" + $lat.CFMs + "ms 色=" + $lat.Color)
$lat = Measure-SelectLatency $ws "F8" 8 1
Write-Host ("  F8: Select=" + $lat.SelectMs + "ms CF跟随时延=" + $lat.CFMs + "ms 色=" + $lat.Color)

# --- 症状2：多选全区域高亮 ---
Write-Host "`n--- 症状2：多选 ---"
$app.Range("A1:A5,C1:C5").Select()
Start-Sleep -Milliseconds 600
Assert ((CFC $ws 2 1) -eq 6) "多选：A 列(2,1) 黄=6"
Assert ((CFC $ws 2 3) -eq 6) "多选：C 列(2,3) 黄=6"
Assert ((CFC $ws 3 2) -eq 6) "多选：B3 黄=6（行3在选区行内）"
Assert ((CFC $ws 6 2) -ne 6) "多选：B6 非黄（行6不在选区）"
Assert ((CFC $ws 5 3) -eq 6) "多选：C5 黄=6（C列在选区）"
Assert ((CFC $ws 7 1) -eq 6) "多选：A7 黄=6（A列在选区，整列高亮）"
Assert ((CFC $ws 7 2) -ne 6) "多选：B7 非黄（行7不在选区、B列未选）"

# 三区域多选
$app.Range("A1:A3,E1:E3,H1:H3").Select()
Start-Sleep -Milliseconds 600
Assert ((CFC $ws 2 1) -eq 6) "三区：A 列(2,1) 黄=6"
Assert ((CFC $ws 2 5) -eq 6) "三区：E 列(2,5) 黄=6"
Assert ((CFC $ws 2 8) -eq 6) "三区：H 列(2,8) 黄=6"
Assert ((CFC $ws 4 2) -ne 6) "三区：B4 非黄（行4不在选区、B列未选）"

# 整列点选（回归：不应多余高亮第 1 行）
$app.Columns.Item(3).Select()
Start-Sleep -Milliseconds 400
Assert ((CFC $ws 10 3) -eq 6) "整列 C：(10,3) 黄=6"
Assert ((CFC $ws 1 1) -ne 6) "整列 C：A1 非黄（无行1多余高亮）"

# 多选→单格：旧高亮必须立刻塌缩为十字（回归）
$app.Range("A1:A5,C1:C5").Select()
Start-Sleep -Milliseconds 500
Assert ((CFC $ws 2 3) -eq 6) "前置：多选 C2 黄"
$app.Range("D4").Select()
Start-Sleep -Milliseconds 400
Assert ((CFC $ws 4 1) -eq 6) "塌缩：A4 黄（行4=D4行）"
Assert ((CFC $ws 1 4) -eq 6) "塌缩：D1 黄（列D=D4列）"
Assert ((CFC $ws 2 3) -ne 6) "塌缩：旧多选 C2 已清"
Assert ((CFC $ws 2 1) -ne 6) "塌缩：旧多选 A2 已清"

# 清理：关聚光灯、关工作簿
Invoke-RibbonButton "聚光灯" | Out-Null
Start-Sleep -Milliseconds 500
$wb.Close($false)
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws) | Out-Null
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null

Write-Host ""
Write-Host "== 端到端 PASSED=$script:passed FAILED=$script:failed =="
if ($script:failed -gt 0) { exit 1 } else { exit 0 }
