# Focused test: does passing a QUALIFIED Chinese macro path ("JYYJHelper.TT通道") as the
# second argument to Application.Run trigger the RPC 0x800706BE / hang?
# Compares: qualified-CN vs unqualified-CN vs qualified-ASCII as the proc argument.
$ErrorActionPreference = "Stop"
$out = "d:\execl1.0\tools\smoke-results\focused-out.txt"
"BOOT" | Out-File $out -Encoding utf8
function Mark($m) { Add-Content $out $m }

$dir = "d:\execl1.0\src\ExcelAiAssistant.Addin\bin\x64\Release\net48\win-x64"
$excelDll = Join-Path $dir "ExcelAiAssistant.Excel.dll"
[System.Reflection.Assembly]::LoadFrom($excelDll) | Out-Null
$asm = [System.Reflection.Assembly]::LoadFrom($excelDll)
$t = $asm.GetType("ExcelAiAssistant.Excel.ScriptRunnerService")
$m = $t.GetMethod("BuildVbaHelperBlock", [System.Reflection.BindingFlags]"Static,NonPublic")
$helper = [string]$m.Invoke($null, @('{"column":"A"}', 1048576))
Mark ("HELPER-LEN " + $helper.Length)

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    Start-Sleep -Seconds 2

    $book = $excel.Workbooks.Add()
    $book.Activate()
    $vbp = $excel.VBE.ActiveVBProject
    $comps = $vbp.VBComponents
    $mod = $comps.Add(1)
    try { $mod.Name = "JYYJHelper" } catch { Mark ("RENAME-FAIL " + $_.Exception.Message) }
    $cm = $mod.CodeModule
    $cm.AddFromString("Sub TT通道()`n    Dim v As Variant`n    v = XaiArg(""column"")`n    XaiReturn Array(v, 42)`nEnd Sub")
    $cm.AddFromString($helper)
    Mark ("MODNAME " + $cm.Parent.Name)

    # 0) with ScreenUpdating=false (like the C# tool)
    $excel.ScreenUpdating = $false
    try {
        $r = $excel.Run("JYYJHelper.XaiProbe9F2E", "JYYJHelper.TT通道")
        Mark ("SU-OFF Q-CN RUN-OK => " + ($r | Out-String -Width 300))
    } catch {
        Mark ("SU-OFF Q-CN RUN-FAIL => " + $_.Exception.Message)
    }
    $excel.ScreenUpdating = $true

    # 1) qualified Chinese proc arg (exactly like the tool)
    try {
        $r = $excel.Run("JYYJHelper.XaiProbe9F2E", "JYYJHelper.TT通道")
        Mark ("Q-CN RUN-OK => " + ($r | Out-String -Width 300))
    } catch {
        Mark ("Q-CN RUN-FAIL => " + $_.Exception.Message)
    }
    # 2) unqualified Chinese proc arg
    try {
        $r = $excel.Run("JYYJHelper.XaiProbe9F2E", "TT通道")
        Mark ("U-CN RUN-OK => " + ($r | Out-String -Width 300))
    } catch {
        Mark ("U-CN RUN-FAIL => " + $_.Exception.Message)
    }
    Mark "DONE"
} catch {
    Mark ("FAIL " + $_.Exception.ToString())
    exit 1
} finally {
    if ($excel -ne $null) {
        try { $excel.DisplayAlerts = $false } catch {}
        try { $excel.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null } catch {}
    }
}
exit 0
