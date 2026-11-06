# Probe: dump scratch module content after sequential VBA runs to diagnose stale gXaiArgs.
$ErrorActionPreference = "Stop"
$dir = "d:\execl1.0\src\ExcelAiAssistant.Addin\bin\x64\Release\net48\win-x64"
$addinDll = Join-Path $dir "ExcelAiAssistant.Addin.dll"
$excelDll = Join-Path $dir "ExcelAiAssistant.Excel.dll"
$excel = $null
try {
    [System.Reflection.Assembly]::LoadFrom($addinDll) | Out-Null
    [System.Reflection.Assembly]::LoadFrom($excelDll) | Out-Null
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    Start-Sleep -Seconds 2
    $book = $excel.Workbooks.Add()
    $book.Activate()

    $excelAsm = [System.Reflection.Assembly]::LoadFrom($excelDll)
    $hostType = $excelAsm.GetType("ExcelAiAssistant.Excel.ExcelToolHost", $true)
    $toolHost = [Activator]::CreateInstance($hostType, @($excel, $null, "probe", "probe-run", "探针"))
    $scriptOpts = $toolHost.GetType().GetProperty("ScriptOptions").GetValue($toolHost, $null)
    $scriptOpts.GetType().GetProperty("Enabled").SetValue($scriptOpts, $true, $null)
    $registry = $toolHost.CreateRegistry()
    $tool = $null
    foreach ($t in $registry.GetType().GetProperty("Tools").GetValue($registry, $null)) {
        if ($t.GetType().GetProperty("Name").GetValue($t, $null) -eq "run_vba_snippet") { $tool = $t; break }
    }
    if ($tool -eq $null) { Write-Output "tool not found"; exit 1 }

    function Dump-Result($tag, $r) {
        if ($r -eq $null) { Write-Output ($tag + " RESULT=NULL"); return }
        $t = $r.GetType()
        $ok = $t.GetProperty("Ok").GetValue($r, $null)
        $code = $t.GetProperty("ErrorCode").GetValue($r, $null)
        $err = $t.GetProperty("Error").GetValue($r, $null)
        $data = $t.GetProperty("Data").GetValue($r, $null)
        Write-Output ($tag + " ok=" + $ok + " code=" + $code)
        if ($err -ne $null) { Write-Output ($tag + " error=" + $err.ToString()) }
        if ($data -ne $null) {
            $dtype = $data.GetType().FullName
            Write-Output ($tag + " dataType=" + $dtype)
            $res = $data.GetType().GetProperty("result").GetValue($data, $null)
            if ($res -ne $null) { Write-Output ($tag + " result=" + $res.ToString()) }
        }
    }

    # 1) warmup: no args
    $warm = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $warm["snippet"] = "Sub TT预热()`n    Range(""H1"").Value = ""ok""`nEnd Sub"
    $warm["macro_name"] = "TT预热"
    $r1 = $tool.Execute.Invoke($warm)
    Dump-Result "WARM" $r1

    # 2) control: args={"column":"A"}
    $ctl = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $ctl["snippet"] = "Sub TT通道()`n    Dim v As Variant`n    v = XaiArg(""column"")`n    XaiReturn Array(v, 42)`nEnd Sub"
    $ctl["macro_name"] = "TT通道"
    $inner = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $inner["column"] = "A"
    $ctl["args"] = $inner
    $r2 = $tool.Execute.Invoke($ctl)
    Dump-Result "CTL" $r2

    # 3) dump module content
    $vbp = $excel.ActiveWorkbook.VBProject
    $mod = $vbp.VBComponents.Item("JYYJHelper")
    $cm = $mod.CodeModule
    $cnt = $cm.CountOfLines
    Write-Output ("MODULE LINES=" + $cnt)
    if ($cnt -gt 0) {
        $all = $cm.Lines(1, $cnt)
        Write-Output ("--- MODULE CONTENT (first 3000 chars) ---")
        Write-Output ($all.Substring(0, [Math]::Min(3000, $all.Length)))
        Write-Output ("--- END ---")
        # count XaiInitArgs occurrences
        $initCount = ([regex]::Matches($all, "Sub XaiInitArgs")).Count
        $probeCount = ([regex]::Matches($all, "XaiProbe9F2E")).Count
        Write-Output ("XaiInitArgs count=" + $initCount + " XaiProbe9F2E count=" + $probeCount)
    }
} catch {
    Write-Output ("FAIL " + $_.Exception.ToString())
} finally {
    if ($excel -ne $null) {
        try { $excel.DisplayAlerts = $false } catch {}
        try { $excel.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null } catch {}
    }
}
exit 0
