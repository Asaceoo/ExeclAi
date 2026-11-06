# Single-variant runner (one Run call per process).
# Usage: powershell -File isolate-run-one.ps1 -Variant VQ   (VQ=qualified ASCII, VU=unqualified ASCII, VQC=qualified CN, VUC=unqualified CN)
param([string]$Variant)
$ErrorActionPreference = "Stop"
$resFile = "d:\execl1.0\tools\smoke-results\iso-$Variant.txt"
"" | Out-File $resFile -Encoding utf8
function Mark($m) { Add-Content $resFile $m }

$dir = "d:\execl1.0\src\ExcelAiAssistant.Addin\bin\x64\Release\net48\win-x64"
$excelDll = Join-Path $dir "ExcelAiAssistant.Excel.dll"

$cfg = @{
    VQ  = @{ Mod='Module1'; Macro='TTcaseA'; Q=$true;  Body="Sub TTcaseA()`n    Dim v As Variant`n    v = XaiArg(""column"")`n    XaiReturn Array(v, 42)`nEnd Sub" }
    VU  = @{ Mod='Module1'; Macro='TTcaseA'; Q=$false; Body="Sub TTcaseA()`n    Dim v As Variant`n    v = XaiArg(""column"")`n    XaiReturn Array(v, 42)`nEnd Sub" }
    VQC = @{ Mod='Module1'; Macro='TT通道';  Q=$true;  Body="Sub TT通道()`n    Dim v As Variant`n    v = XaiArg(""column"")`n    XaiReturn Array(v, 42)`nEnd Sub" }
    VUC = @{ Mod='Module1'; Macro='TT通道';  Q=$false; Body="Sub TT通道()`n    Dim v As Variant`n    v = XaiArg(""column"")`n    XaiReturn Array(v, 42)`nEnd Sub" }
}
$c = $cfg[$Variant]
if (-not $c) { Mark "UNKNOWN-VARIANT"; exit 1 }
$mode = if ($c.Q) { "QUALIFIED" } else { "UNQUALIFIED" }

$excel = $null
try {
    [System.Reflection.Assembly]::LoadFrom($excelDll) | Out-Null
    Mark "ASM-LOADED"
    $t0 = [DateTime]::Now
    $excel = New-Object -ComObject Excel.Application
    Mark ("COM-CREATED " + [int](([DateTime]::Now - $t0).TotalSeconds) + "s")
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    Start-Sleep -Seconds 2
    $book = $excel.Workbooks.Add()
    $book.Activate()
    Mark "BOOK-ADDED"

    $vbp = $excel.VBE.ActiveVBProject
    $comps = $vbp.VBComponents
    $mod = $comps.Add(1)
    try { $mod.Name = $c.Mod } catch { Mark ("  rename-fail " + $_.Exception.Message) }
    $cm = $mod.CodeModule
    $cm.AddFromString($c.Body)
    $helper = [System.Reflection.Assembly]::LoadFrom($excelDll).GetType("ExcelAiAssistant.Excel.ScriptRunnerService").GetMethod("BuildVbaHelperBlock", [System.Reflection.BindingFlags]"Static,NonPublic").Invoke($null, @('{"column":"A"}', 1048576))
    $cm.AddFromString($helper)
    $fqProbe = $c.Mod + ".XaiProbe9F2E"
    Mark ("  injected mod=" + $mod.Name + " probeFQ=" + $fqProbe + " mode=" + $mode)

    $arg = if ($c.Q) { $c.Mod + "." + $c.Macro } else { $c.Macro }
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $r = $excel.Run($fqProbe, $arg)
        Mark ($mode + " OK => " + $r + " (" + [int]$sw.Elapsed.TotalSeconds + "s)")
    } catch {
        Mark ($mode + " EXC => " + $_.Exception.Message + " (" + [int]$sw.Elapsed.TotalSeconds + "s)")
    }
    $comps.Remove($mod)
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
