# Replicate the tool's EXACT pre-Run sequence via direct COM, with step toggles for bisection.
# Modes:
#   FULL   = every tool step (ActiveWorkbook.VBProject, Lines check, ScreenUpdating=false)
#   NOAW   = skip ActiveWorkbook.VBProject (use VBE.ActiveVBProject)
#   NOSU   = skip ScreenUpdating=false
#   NOLINES= skip Lines read (no hasProbe check, always append helper)
#   MIN    = only Add+AddFromString+Run (the known-good direct baseline)
param([string]$Mode = "FULL")
$ErrorActionPreference = "Stop"
$resFile = "d:\execl1.0\tools\smoke-results\toolseq-$Mode.txt"
"" | Out-File $resFile -Encoding utf8
function Mark($m) { Add-Content $resFile $m }

$dir = "d:\execl1.0\src\ExcelAiAssistant.Addin\bin\x64\Release\net48\win-x64"
$excelDll = Join-Path $dir "ExcelAiAssistant.Excel.dll"

$snippet = "Sub TT通道()`n    Dim v As Variant`n    v = XaiArg(""column"")`n    XaiReturn Array(v, 42)`nEnd Sub"
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

    # Step 1: resolve VB project (tool uses ActiveWorkbook.VBProject unless NOAW)
    if ($Mode -eq "NOAW") { $vbp = $excel.VBE.ActiveVBProject; Mark "  proj=VBE.ActiveVBProject" }
    else { $vbp = $excel.ActiveWorkbook.VBProject; Mark "  proj=ActiveWorkbook.VBProject" }

    $comps = $vbp.VBComponents
    $mod = $null
    try { $mod = $comps.Item("JYYJHelper") } catch { $mod = $null }
    if ($null -eq $mod) { $mod = $comps.Add(1); try { $mod.Name = "JYYJHelper" } catch {} }
    $cm = $mod.CodeModule
    Mark "  module-ready"

    # ReplaceModuleCode
    try {
        $count = $cm.CountOfLines
        if ($count -gt 0) { $cm.DeleteLines(1, $count) }
    } catch { Mark ("  delete-exc " + $_.Exception.Message) }
    $cm.AddFromString($snippet)
    Mark "  snippet-added"

    # AppendVbaProbe (with Lines check unless NOLINES)
    $helper = [System.Reflection.Assembly]::LoadFrom($excelDll).GetType("ExcelAiAssistant.Excel.ScriptRunnerService").GetMethod("BuildVbaHelperBlock", [System.Reflection.BindingFlags]"Static,NonPublic").Invoke($null, @('{"column":"A"}', 1048576))
    if ($Mode -ne "NOLINES") {
        try {
            $c2 = $cm.CountOfLines
            $all = $cm.Lines(1, $c2)
            $hasProbe = $all.Contains("XaiProbe9F2E")
        } catch { $hasProbe = $false; Mark ("  lines-exc " + $_.Exception.Message) }
        if (-not $hasProbe) { $cm.AddFromString($helper); Mark "  helper-appended(checked)" }
    } else {
        $cm.AddFromString($helper); Mark "  helper-appended(no-check)"
    }

    # ScreenUpdating
    $prevScreen = $null
    if ($Mode -ne "NOSU" -and $Mode -ne "MIN") {
        try { $prevScreen = $excel.ScreenUpdating; if ($prevScreen) { $excel.ScreenUpdating = $false; Mark "  screenupdating=false" } } catch { Mark "  su-exc" }
    }

    # Run
    $fqProbe = "JYYJHelper.XaiProbe9F2E"
    $arg = "JYYJHelper.TT通道"
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $r = $excel.Run($fqProbe, $arg)
        Mark ("RUN OK => " + $r + " (" + [int]$sw.Elapsed.TotalSeconds + "s)")
    } catch {
        Mark ("RUN EXC => " + $_.Exception.Message + " (" + [int]$sw.Elapsed.TotalSeconds + "s)")
    }
    if ($prevScreen -ne $null) { try { $excel.ScreenUpdating = $prevScreen } catch {} }
    try { $comps.Remove($mod) } catch {}
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
