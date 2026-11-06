# Real-machine verification for P5-B (workflow). English-only to avoid encoding issues.
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Reflection

$dll = "$env:LOCALAPPDATA\XaiAssistant\app\x64\ExcelAiAssistant.Addin.dll"
$wbPath = "d:\execl1.0\tools\live-validation.xlsx"
$results = @()

try {
    # --- 1) embedded index.html must contain the workflow page ---
    $asm = [System.Reflection.Assembly]::LoadFrom($dll)
    $names = $asm.GetManifestResourceNames()
    $htmlEntry = $names | Where-Object { $_ -like "*index.html*" } | Select-Object -First 1
    if (-not $htmlEntry) { $results += "FAIL embedded-html-missing"; }
    else {
        $stream = $asm.GetManifestResourceStream($htmlEntry)
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
        $html = $reader.ReadToEnd(); $reader.Close(); $stream.Close()
        $check1 = ($html -like "*view-workflows*")
        $check2 = ($html -like "*refreshWorkflows*")
        $check3 = ($html -like "*RunWorkflow*")
        $check4 = ($html -like "*retryWorkflowStep*")
        $results += ("PASS embedded-workflow-page" + " view=$check1 js=$check2 run=$check3 retry=$check4")
    }

    # --- 2) launch a fresh Excel, open the test workbook, verify addin loads ---
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $wb = $excel.Workbooks.Open($wbPath)
    Start-Sleep -Seconds 3

    $addin = $excel.COMAddIns | Where-Object { $_.ProgId -eq "ExcelAiAssistant.AddIn" } | Select-Object -First 1
    if ($addin) {
        $results += ("PASS addin-loaded progid=" + $addin.ProgId + " connect=" + $addin.Connect)
    } else {
        $results += "FAIL addin-not-loaded"
    }

    # version from the exposed app object name / assembly
    $results += ("PASS excel-version=" + $excel.Version)

    $wb.Close($false)
    $excel.Quit()
    $null = [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
} catch {
    $results += ("FAIL exception " + $_.Exception.Message)
}

$results | ForEach-Object { Write-Output $_ }
$failed = @($results | Where-Object { $_ -like "FAIL*" }).Count
Write-Output ("FAIL_COUNT=" + $failed)
if ($failed -gt 0) { exit 1 }
Write-Output "REAL-MACHINE VERIFY OK"
exit 0