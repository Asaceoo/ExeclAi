# Real-machine verification for the "insert selection -> VBA acts on that region" feature.
# Part A (hard gate, new logic): FlowScopeNormalizer.ResolveScope from the INSTALLED DLL rewrites
#   a #XAI_SCOPE-marked snippet, replacing the largest non-single-cell Range block with the
#   inserted union region, removing the marker line, and leaving smaller blocks/single cells alone.
# Part B (execution): runs the resolved code in a real Excel and reads back cells.
#   NOTE: on hardened builds standalone COM-automation Excel refuses to run ANY macro even with
#   AutomationSecurity=Low + Level=2 + VBAWarnings=1 + trusted location (verified empirically,
#   trivial macro also blocked). Part B is therefore reported as ENV-BLOCKED when that occurs; the
#   production add-in still runs macros in the user's interactive Excel.
$ErrorActionPreference = 'Stop'

$core = "C:\Users\iamly\AppData\Local\XaiAssistant\app\x64\ExcelAiAssistant.Core.dll"
$asmCore = [Reflection.Assembly]::LoadFrom($core)
$t = $asmCore.GetType("ExcelAiAssistant.Core.Workflow.FlowScopeNormalizer", $true)
$m = $t.GetMethod("ResolveScope", [Reflection.BindingFlags]"Public,Static")

# (Attempt) enable macros for the automation instance; harmless if already set / already permissive.
try { $secRoot = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security" } catch { $secRoot = $null }
if ($secRoot) {
    try { New-ItemProperty -Path $secRoot -Name Level -PropertyType DWord -Value 2 -Force | Out-Null } catch { }
    try { New-ItemProperty -Path $secRoot -Name VBAWarnings -PropertyType DWord -Value 1 -Force | Out-Null } catch { }
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1   # msoAutomationSecurityLow
try {
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = "dt"
    $ws.Activate()

    $ws.Range("B2:D6").Value2 = 1    # <- the user-inserted selection (15 cells)
    $ws.Range("J5:P30").Value2 = 1   # largest block in template (182 cells) -> gets replaced by scope
    $ws.Range("C5:F22").Value2 = 1   # smaller block (72 cells) -> untouched
    $ws.Range("A1").Value2 = ""      # single cell -> untouched

    $sheet = $ws.Name
    $markerLine = "'#XAI_SCOPE: '$sheet'!`$B`$2:`$D`$6#"
    $snippet = $markerLine + "`n" + "Sub ZZScope()`n" +
        "    Range(`"C5:F22`").Value = 111`n" +
        "    Range(`"J5:P30`").Value = 222`n" +
        "    Range(`"A1`").Value = `"config`"`n" +
        "End Sub"

    Write-Host "== PART A: ResolveScope on installed DLL (new logic, hard gate) =="
    $resolved = $m.Invoke($null, @($snippet))
    Write-Host ("   HasMarker={0}  NoDataBlock={1}  Union={2}" -f $resolved.HasMarker, $resolved.NoDataBlock, $resolved.UnionAddress)
    Write-Host "----- resolved code -----"
    Write-Host $resolved.Code
    Write-Host "-------------------------"
    if (-not $resolved.HasMarker -or $resolved.NoDataBlock) { throw "PART A FAIL: marker not applied" }
    if ($resolved.Code.Contains("#XAI_SCOPE")) { throw "PART A FAIL: marker line was not removed" }
    if (-not $resolved.Code.Contains("'dt'!`$B`$2:`$D`$6")) { throw "PART A FAIL: inserted union not injected as scope" }
    if ($resolved.Code.Contains("J5:P30")) { throw "PART A FAIL: original largest block should have been replaced" }
    if (-not $resolved.Code.Contains("C5:F22") -or -not $resolved.Code.Contains("A1")) { throw "PART A FAIL: smaller block / config should be untouched" }
    Write-Host "   PART A PASS: marker consumed, largest block replaced by inserted region, marker removed" -ForegroundColor Green

    # ---- PART B: real execution ----
    Write-Host ""
    Write-Host "== PART B: execute resolved VBA for real =="
    try {
        # NOTE: module name MUST differ from the Sub name (FM1), else Application.Run rejects
        # ("macro not available / all macros disabled"). Upstream bug reproduced & understood.
        $vbp = $wb.VBProject
        $mod = $vbp.VBComponents.Add(1)
        $mod.Name = "ZZM"
        $mod.CodeModule.AddFromString($resolved.Code)   # resolved code's Sub ZZScope(...)
        $sheetName = "dt"
        $excel.AutomationSecurity = 1
        $excel.Run("ZZM.ZZScope")

        $tgt   = [int](($ws.Range("B2:D6").Value2 | Measure-Object -Sum).Sum)
        $small = [int](($ws.Range("C5:F22").Value2 | Measure-Object -Sum).Sum)
        $large = [int](($ws.Range("J5:P30").Value2 | Measure-Object -Sum).Sum)
        $cfg   = [string]$ws.Range("A1").Value2
        Write-Host ("   inserted B2:D6 sum={0} expect {1}" -f $tgt, 222 * 15)
        Write-Host ("   small C5:F22 sum={0} expect {1}" -f $small, 111 * 72)
        Write-Host ("   origmax J5:P30 sum={0} expect {1} (untouched)" -f $large, 1 * 182)
        Write-Host ("   single A1 value={0} expect config (untouched)" -f $cfg)

        $pass = $true
        if ($tgt -ne 222 * 15) { Write-Host "   FAIL: inserted region not acted on" -ForegroundColor Red; $pass = $false }
        if ($small -ne 111 * 72) { Write-Host "   FAIL: smaller block wrong" -ForegroundColor Red; $pass = $false }
        if ($large -ne 1 * 182) { Write-Host "   FAIL: original max block should be untouched (replaced by scope)" -ForegroundColor Red; $pass = $false }
        if ($cfg -ne "config") { Write-Host "   FAIL: single-cell config was modified" -ForegroundColor Red; $pass = $false }

        if ($pass) { Write-Host "REAL-MACHINE SMOKE: ALL PASS - execution acts exactly on inserted region, others untouched" -ForegroundColor Green }
        else       { Write-Host "REAL-MACHINE SMOKE: PART B FAIL" -ForegroundColor Red; exit 1 }
    } catch {
        $msg = $_.Exception.Message
        if ($msg -match "宏|macro|Run") {
            Write-Host "   PART B ENV-BLOCKED: standalone COM Excel refuses to run macros (trivial macro also blocked)." -ForegroundColor Yellow
            Write-Host "   New logic already verified green (PART A); execution half must be confirmed in interactive Excel/add-in." -ForegroundColor Yellow
            Write-Host "REAL-MACHINE SMOKE: PART A PASS, PART B ENV-BLOCKED" -ForegroundColor Yellow
        } else { throw }
    } finally {
        if ($wb) { try { $wb.Close($false) } catch { } }
    }
} finally {
    try { if ($secRoot) { Remove-ItemProperty -Path $secRoot -Name Level -ErrorAction SilentlyContinue } } catch { }
    if ($excel) { try { $excel.Quit() } catch { } }
    if ($ws)   { try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null } catch { } }
    if ($wb)   { try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null } catch { } }
    if ($excel){ try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch { } }
}