# 真机验证「视图组」聚光灯/指针（CF 非破坏方案）：UIA 驱动真实 Ribbon 按钮 + COM 读状态。
# 前置：Excel 已启动并加载新版加载项（视图组已并入 JYYJ助手 Tab）。
# 注意：Excel 进程内可能挂有搜狗拼音 SoPY_Status 状态窗口，必须按 XLMAIN 类名定位主窗口。
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

function Ensure-ViewOff([string]$btnName) {
    $s = Get-RibbonButtonState $btnName
    if ($s -eq $true) { Invoke-RibbonButton $btnName | Out-Null; Start-Sleep -Milliseconds 500 }
    Write-Host "  ensure $btnName off (was=$s)"
}

function Get-SentinelRuleCount($ws) {
    $n = 0
    for ($i = 1; $i -le $ws.Cells.FormatConditions.Count; $i++) {
        $f = [string]$ws.Cells.FormatConditions.Item($i).Formula1
        if ($f -like '*JYYJ_VIEW_MODE*') { $n++ }
    }
    return $n
}

function Get-SentinelNameExists($wb) {
    for ($i = 1; $i -le $wb.Names.Count; $i++) {
        if ([string]$wb.Names.Item($i).Name -eq 'JYYJ_VIEW_MODE') { return $true }
    }
    return $false
}

# 读取 CF 渲染层颜色/边框。点击按钮后 Excel 可能处于重算忙态，COM 读取会瞬时返回 null，
# 重试几次（最多 ~2s）吸收忙态，避免把瞬时忙态误判为实现缺陷。
function Get-CFColor([int]$r, [int]$c) {
    for ($try = 0; $try -lt 10; $try++) {
        try { return [int]$ws.Cells.Item($r, $c).DisplayFormat.Interior.ColorIndex }
        catch { Start-Sleep -Milliseconds 200 }
    }
    return -9999
}

function Get-CFBorder([int]$r, [int]$c, [int]$edge) {
    for ($try = 0; $try -lt 10; $try++) {
        try { return [int]$ws.Cells.Item($r, $c).DisplayFormat.Borders.Item($edge).ColorIndex }
        catch { Start-Sleep -Milliseconds 200 }
    }
    return -9999
}

$app = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")

# 1. 准备测试工作簿与数据
$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
for ($r = 1; $r -le 6; $r++) { for ($c = 1; $c -le 6; $c++) { $ws.Cells.Item($r, $c).Value2 = $r * 10 + $c } }
$app.Range("B2").Select()
Start-Sleep -Milliseconds 300

# 2. 激活 JYYJ助手 Tab；确认旧的 JYYJ快捷 Tab 已移除
Assert (Select-RibbonTab "JYYJ助手") "激活 JYYJ助手 Tab"
Start-Sleep -Milliseconds 800
Assert (-not (Select-RibbonTab "JYYJ快捷")) "旧的 JYYJ快捷 Tab 已不存在"

# 2.5 复位：确保聚光灯/指针都处于关闭
Ensure-ViewOff "聚光灯"
Ensure-ViewOff "指针"

# 3. 开启聚光灯并断言 CF 高亮（活动格 B2 的整行/整列变黄=6，显示层），且真实 Interior 未被动过
Assert (Invoke-RibbonButton "聚光灯") "点击「聚光灯」按钮"
Start-Sleep -Milliseconds 1000
Assert ((Get-CFColor 2 1) -eq 6) "CF高亮：B2 行内 A2 黄=6"
Assert ((Get-CFColor 2 6) -eq 6) "CF高亮：B2 行内 F2 黄=6"
Assert ((Get-CFColor 1 2) -eq 6) "CF高亮：B2 列内 B1 黄=6"
Assert ((Get-CFColor 6 2) -eq 6) "CF高亮：B2 列内 B6 黄=6"
Assert ((Get-CFColor 1 1) -ne 6) "CF高亮：A1 非黄"
Assert ([int]$ws.Cells.Item(2,1).Interior.ColorIndex -eq -4142) "非破坏：A2 真实格式仍无色(-4142)"
Assert ((Get-SentinelRuleCount $ws) -eq 1) "特征标记：恰好 1 条哨兵 CF 规则"
Assert (Get-SentinelNameExists $wb) "特征标记：哨兵定义名 JYYJ_VIEW_MODE 存在"

# 4. 移动活动格到 D4，断言高亮跟随、旧区域恢复
$app.Range("D4").Select()
Start-Sleep -Milliseconds 1000
Assert ((Get-CFColor 4 1) -eq 6) "跟随：D4 行内 A4 黄=6"
Assert ((Get-CFColor 1 4) -eq 6) "跟随：D4 列内 D1 黄=6"
Assert ((Get-CFColor 2 1) -eq -4142) "恢复：B2 行旧区域无色(-4142)"
Assert ((Get-CFColor 1 2) -eq -4142) "恢复：B2 列旧区域无色(-4142)"
Assert ((Get-SentinelRuleCount $ws) -eq 1) "跟随：哨兵规则仍为 1 条（旧规则已替换）"

# 4.5 多选（Ctrl 多区域）：A1:A5 + C1:C5 → 逐区域十字（行1-5 全亮 + A/C 列全亮，B 列行内亮、第6行不亮）
$app.Range("A1:A5,C1:C5").Select()
Start-Sleep -Milliseconds 800
Assert ((Get-CFColor 2 1) -eq 6) "多选：A 列(2,1) 黄=6"
Assert ((Get-CFColor 4 1) -eq 6) "多选：A 列(4,1) 黄=6"
Assert ((Get-CFColor 2 3) -eq 6) "多选：C 列(2,3) 黄=6"
Assert ((Get-CFColor 1 3) -eq 6) "多选：C 列(1,3) 黄=6"
Assert ((Get-CFColor 3 2) -eq 6) "多选：B3 黄=6（行3在选区行内）"
Assert ((Get-CFColor 6 2) -ne 6) "多选：B6 非黄（行6不在选区、B列未选）"
Assert ((Get-SentinelRuleCount $ws) -eq 1) "多选：哨兵规则仍为 1 条（零churn）"

# 4.6 连续多格选区（单区域）：A1:B4 → 行1-4 ∪ 列A-B
$app.Range("A1:B4").Select()
Start-Sleep -Milliseconds 800
Assert ((Get-CFColor 3 1) -eq 6) "连续区：A3 黄=6（行3在选区）"
Assert ((Get-CFColor 1 2) -eq 6) "连续区：B1 黄=6（B列在选区）"
Assert ((Get-CFColor 5 3) -ne 6) "连续区：C5 非黄（行5不在选区、C列未选）"
Assert ((Get-CFColor 5 1) -eq 6) "连续区：A5 黄=6（A列在选区）"

# 4.7 点整列（列标 A）：只高亮 A 列整列，B 列不亮
$app.Columns.Item(1).Select()
Start-Sleep -Milliseconds 800
Assert ((Get-CFColor 1 1) -eq 6) "整列：A1 黄=6"
Assert ((Get-CFColor 6 1) -eq 6) "整列：A6 黄=6"
Assert ((Get-CFColor 1 2) -ne 6) "整列：B1 非黄"

# 4.8 点整行（行标 1）：只高亮第 1 行，第 2 行不亮
$app.Rows.Item(1).Select()
Start-Sleep -Milliseconds 800
Assert ((Get-CFColor 1 2) -eq 6) "整行：B1 黄=6"
Assert ((Get-CFColor 1 6) -eq 6) "整行：F1 黄=6"
Assert ((Get-CFColor 2 1) -ne 6) "整行：A2 非黄"

# 4.9 恢复单格，为下一步互斥测试做准备
$app.Range("D4").Select()
Start-Sleep -Milliseconds 800
Assert ((Get-CFColor 4 1) -eq 6) "恢复单格：D4 行内 A4 黄=6"

# 5. 互斥：开指针，聚光灯应关闭，D4 出现红色边框（CF 渲染层）
Assert (Invoke-RibbonButton "指针") "点击「指针」按钮"
Start-Sleep -Milliseconds 1000
Assert ((Get-CFColor 4 1) -eq -4142) "互斥：聚光灯已关闭(无色)"
Assert ((Get-CFBorder 4 4 7) -eq 3) "指针：D4 左边框红=3"
Assert ((Get-CFBorder 4 4 8) -eq 3) "指针：D4 上边框红=3"
Assert ((Get-CFBorder 3 3 7) -ne 3) "指针：C3 无红框"

# 5.5 指针多选：B2:C3 + E5:F6 → 两框都红，中间 D4 无框
$app.Range("B2:C3,E5:F6").Select()
Start-Sleep -Milliseconds 800
Assert ((Get-CFBorder 2 2 7) -eq 3) "指针多选：B2 左边框红=3"
Assert ((Get-CFBorder 3 3 7) -eq 3) "指针多选：C3 左边框红=3"
Assert ((Get-CFBorder 5 5 7) -eq 3) "指针多选：E5 左边框红=3"
Assert ((Get-CFBorder 6 6 7) -eq 3) "指针多选：F6 左边框红=3"
Assert ((Get-CFBorder 4 4 7) -ne 3) "指针多选：D4 无红框（中间未选）"

# 6. 关闭指针，断言边框恢复 + 规则/哨兵名清理
Assert (Invoke-RibbonButton "指针") "再次点击「指针」关闭"
Start-Sleep -Milliseconds 1000
Assert ((Get-CFBorder 4 4 7) -ne 3) "关闭：D4 红框已移除"
$dCnt6 = 0
try { $dCnt6 = [int]$ws.Cells.FormatConditions.Count } catch { Write-Host "  [DUMP6] Count ERR: $($_.Exception.Message)" }
Write-Host "  [DUMP6] 关闭指针后 FormatConditions.Count=$dCnt6"
for ($i = 1; $i -le $dCnt6; $i++) {
    try { Write-Host "  [DUMP6]   rule[$i] Type=$([int]$ws.Cells.FormatConditions.Item($i).Type) Formula1=$([string]$ws.Cells.FormatConditions.Item($i).Formula1)" }
    catch { Write-Host "  [DUMP6]   rule[$i] ERR: $($_.Exception.Message)" }
}
Assert ((Get-SentinelRuleCount $ws) -eq 0) "关闭：哨兵 CF 规则已全部清除"
Assert (-not (Get-SentinelNameExists $wb)) "关闭：哨兵定义名已移除"

# 7. 崩溃残留自愈：注入哨兵规则+定义名模拟崩溃遗留，保存后重启 Excel（15s 延迟清理兜底）
$tmp = Join-Path $env:TEMP "jyyj_view_crash_test.xlsx"
if (Test-Path $tmp) { Remove-Item $tmp -Force }
$fc = $ws.Range("A1:F6").FormatConditions.Add(2, 0, '=AND(JYYJ_VIEW_MODE,ROW()=5,COLUMN()=5)')
$fc.Interior.ColorIndex = 6
$wb.Names.Add('JYYJ_VIEW_MODE', '=1') | Out-Null
$wb.SaveAs($tmp, 51)
Assert ((Get-SentinelRuleCount $ws) -ge 1) "残留注入：哨兵 CF 规则已写入并保存"
$wb.Close($false)
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws) | Out-Null
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null
$app.Quit()
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) | Out-Null

Start-Sleep -Seconds 4
Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2
Write-Host "  重启 Excel（带残留文件启动，加载项启动/延迟清理触发自愈）..."
Start-Process excel -ArgumentList "`"$tmp`""
Start-Sleep -Seconds 3
$app2 = $null
for ($i = 0; $i -lt 40; $i++) {
    try { $app2 = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application"); break } catch { Start-Sleep -Seconds 1 }
}
Assert ($null -ne $app2) "重启后取得 Excel 实例"
$wb2 = $null
for ($i = 0; $i -lt 30; $i++) {
    try { if ($app2.Workbooks.Count -gt 0) { $wb2 = $app2.Workbooks.Item(1); break } } catch { }
    Start-Sleep -Seconds 1
}
Assert ($null -ne $wb2) "重启后残留工作簿已打开"
$ws2 = $wb2.Worksheets.Item(1)
Write-Host "  等待 22s（覆盖 15s 延迟自愈清理）..."
Start-Sleep -Seconds 22
Write-Host "  [DUMP] 自愈检查点状态:"
Write-Host "  [DUMP]   wb2.Name=$($wb2.Name) ws2.Name=$($ws2.Name)"
$dCnt = 0
try { $dCnt = [int]$ws2.Cells.FormatConditions.Count } catch { Write-Host "  [DUMP]   Count ERR: $($_.Exception.Message)" }
Write-Host "  [DUMP]   FormatConditions.Count=$dCnt"
for ($i = 1; $i -le $dCnt; $i++) {
    try { Write-Host "  [DUMP]     rule[$i] Type=$([int]$ws2.Cells.FormatConditions.Item($i).Type) Formula1=$([string]$ws2.Cells.FormatConditions.Item($i).Formula1)" }
    catch { Write-Host "  [DUMP]     rule[$i] ERR: $($_.Exception.Message)" }
}
Assert ((Get-SentinelRuleCount $ws2) -eq 0) "自愈：重启后遗留哨兵 CF 规则已清除"
Assert (-not (Get-SentinelNameExists $wb2)) "自愈：遗留哨兵定义名已移除"

$wb2.Close($false)
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws2) | Out-Null
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb2) | Out-Null

Write-Host ""
Write-Host "== 视图组 PASSED=$script:passed FAILED=$script:failed =="
if ($script:failed -gt 0) { Write-Host "VIEW-GROUP CHECKS FAILED"; exit 1 } else { Write-Host "ALL VIEW-GROUP REAL-MACHINE CHECKS OK"; exit 0 }
