# Real-machine verification for VbaAutoSinkPolicy.ShouldAutoSink on the INSTALLED DLL.
# ASCII only (PS 5.1 GBK hazard). Exit code 0 = all PASS.
$ErrorActionPreference = "Stop"
$core = "$env:LOCALAPPDATA\XaiAssistant\app\x64\ExcelAiAssistant.Core.dll"
if (-not (Test-Path $core)) { Write-Output "FAIL: core dll not found at $core"; exit 1 }

$asm = [System.Reflection.Assembly]::LoadFrom($core)
$policyType = $asm.GetType("ExcelAiAssistant.Core.Tools.VbaAutoSinkPolicy")
$recordType = $asm.GetType("ExcelAiAssistant.Core.Agent.ToolEventRecord")
$resultType = $asm.GetType("ExcelAiAssistant.Core.Models.ToolResult")
if (-not $policyType -or -not $recordType -or -not $resultType) {
    Write-Output "FAIL: required types missing"; exit 1
}

function New-Event([string]$tool, [bool]$ok) {
    $rec = [System.Activator]::CreateInstance($recordType)
    $rec.Tool = $tool
    $res = [System.Activator]::CreateInstance($resultType)
    $res.Ok = $ok
    $rec.Result = $res
    return $rec
}

function Invoke-Policy($events, [string]$snippet) {
    $listType = [System.Collections.Generic.List``1].MakeGenericType($recordType)
    $list = [System.Activator]::CreateInstance($listType)
    foreach ($e in $events) { $list.Add($e) }
    $m = $policyType.GetMethod("ShouldAutoSink")
    return $m.Invoke($null, @($list, [string]$snippet))
}

$fail = 0
function Check([string]$name, $actual, $expected) {
    if ($actual -eq $expected) { Write-Output "PASS: $name = $actual" }
    else { Write-Output "FAIL: $name expected $expected got $actual"; $script:fail = 1 }
}

# S1: run_vba ok + valid snippet -> True
Check "S1 run_vba_ok_sinks" (Invoke-Policy @(New-Event "run_vba_snippet" $true) "Sub x(): End Sub") $true

# S2: save_vba_tool called (even failed) -> False
Check "S2 save_called_blocks" (Invoke-Policy @(New-Event "run_vba_snippet" $true; New-Event "save_vba_tool" $false) "Sub x(): End Sub") $false

# S3: empty snippet -> False
Check "S3 empty_snippet" (Invoke-Policy @(New-Event "run_vba_snippet" $true) "") $false

# S4: no successful run_vba -> False
Check "S4 no_success_run" (Invoke-Policy @(New-Event "run_vba_snippet" $false) "Sub x(): End Sub") $false

# S5: custom tool ct_xxx execution must NOT trigger auto-sink (iteration-2 fix)
Check "S5 ct_tool_no_sink" (Invoke-Policy @(New-Event "ct_my_tool" $true) "Sub x(): End Sub") $false

# S6: other tools only -> False
Check "S6 other_tools" (Invoke-Policy @(New-Event "write_excel_range" $true) "Sub x(): End Sub") $false

# S7: null events -> False
Check "S7 null_events" (Invoke-Policy $null "Sub x(): End Sub") $false

if ($fail -eq 0) { Write-Output "ALL PASS" } else { Write-Output "SOME FAILED"; exit 1 }
