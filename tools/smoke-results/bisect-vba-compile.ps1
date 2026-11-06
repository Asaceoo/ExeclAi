# Bisect whether the injected helper block fails to compile on its own (with a minimal snippet)
# or only in combination with the case-A snippet.
$ErrorActionPreference = "Stop"
$out = "d:\execl1.0\tools\smoke-results\bisect-out.txt"
"BOOT" | Out-File $out -Encoding utf8
function Mark($m) { Add-Content $out $m }

$dir = "d:\execl1.0\src\ExcelAiAssistant.Addin\bin\x64\Release\net48\win-x64"
$excelDll = Join-Path $dir "ExcelAiAssistant.Excel.dll"
$addinDll = Join-Path $dir "ExcelAiAssistant.Addin.dll"
if (-not (Test-Path $excelDll)) { Mark "MISSING-DLL"; exit 1 }

$excel = $null
try {
    [System.Reflection.Assembly]::LoadFrom($excelDll) | Out-Null
    $asm = [System.Reflection.Assembly]::LoadFrom($excelDll)
    $t = $asm.GetType("ExcelAiAssistant.Excel.ScriptRunnerService")
    $m = $t.GetMethod("BuildVbaHelperBlock", [System.Reflection.BindingFlags]"Static,NonPublic")
    $helper = [string]$m.Invoke($null, @('{"column":"A"}', 1048576))
    Mark ("HELPER-LEN " + $helper.Length)

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    Start-Sleep -Seconds 2

    function Test-Module($label, $snippet, $proc) {
        $book = $excel.Workbooks.Add()
        $book.Activate()
        $vbp = $excel.VBE.ActiveVBProject
        $comps = $vbp.VBComponents
        $mod = $comps.Add(1)
        $cm = $mod.CodeModule
        $cm.AddFromString($snippet)
        $cm.AddFromString($helper)
        $mname = $cm.Parent.Name
        try {
            $r = $excel.Run($mname + ".XaiProbe9F2E", $proc)
            Mark ("[" + $label + "] RUN-OK => " + ($r | Out-String -Width 300))
        } catch {
            Mark ("[" + $label + "] RUN-FAIL => " + $_.Exception.Message)
        }
        $book.Close($false)
        try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) | Out-Null } catch {}
    }

    # minimal snippet: only XaiReturn
    Test-Module "MIN" "Sub TTmin()`n    XaiReturn Array(1, 2)`nEnd Sub" "TTmin"
    # case A snippet (Chinese proc name)
    Test-Module "CASE-A" "Sub TT通道()`n    Dim v As Variant`n    v = XaiArg(""column"")`n    XaiReturn Array(v, 42)`nEnd Sub" "TT通道"
    # no snippet, helper only + probe run with no-op proc
    Test-Module "EMPTY" "Sub TTnone()`nEnd Sub" "TTnone"
    # V1: CASE-A body with ASCII proc name (isolate Chinese name)
    Test-Module "V1-ASCII" "Sub TTcaseA()`n    Dim v As Variant`n    v = XaiArg(""column"")`n    XaiReturn Array(v, 42)`nEnd Sub" "TTcaseA"
    # V2: CASE-A without XaiArg (isolate XaiArg call)
    Test-Module "V2-NOXAIARG" "Sub TTb()`n    Dim v As Variant`n    XaiReturn Array(v, 42)`nEnd Sub" "TTb"
    # V3: only XaiArg + Dim (isolate XaiArg return into Variant)
    Test-Module "V3-XAIARG" "Sub TTc()`n    Dim v As Variant`n    v = XaiArg(""column"")`nEnd Sub" "TTc"
    # V4: XaiReturn with literal, no Dim, no XaiArg
    Test-Module "V4-LIT" "Sub TTd()`n    XaiReturn Array(42)`nEnd Sub" "TTd"

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
