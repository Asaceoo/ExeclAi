# Real-machine verification for the sink exception-isolation fix (1.20.0.20).
# Checks that SinkVbaCore and the auto-sink block in SendPrompt are compiled with try/catch.
# ASCII only (PS 5.1 GBK hazard). Exit code 0 = all PASS.
$ErrorActionPreference = "Stop"
$addin = "$env:LOCALAPPDATA\XaiAssistant\app\x64\ExcelAiAssistant.Addin.dll"
if (-not (Test-Path $addin)) { Write-Output "FAIL: addin dll not found at $addin"; exit 1 }

$asm = [System.Reflection.Assembly]::LoadFrom($addin)
$bridgeType = $asm.GetType("ExcelAiAssistant.Addin.SidebarBridge")
if (-not $bridgeType) { Write-Output "FAIL: SidebarBridge type missing"; exit 1 }

$fail = 0
function Check([string]$name, $actual, $expected) {
    if ($actual -eq $expected) { Write-Output "PASS: $name = $actual" }
    else { Write-Output "FAIL: $name expected $expected got $actual"; $script:fail = 1 }
}

# 1) SinkVbaCore is a private instance method with >= 1 exception handling clause (try/catch)
$sinkMethod = $bridgeType.GetMethod("SinkVbaCore", [System.Reflection.BindingFlags] "NonPublic,Instance")
if (-not $sinkMethod) { Write-Output "FAIL: SinkVbaCore method missing"; exit 1 }
$sinkClauses = $sinkMethod.GetMethodBody().ExceptionHandlingClauses.Count
Check "SinkVbaCore_has_trycatch" $sinkClauses 1

# 2) SendPrompt body is an async lambda passed to ExecuteAsync -> its try/catch lives in the
#    generated state machine, which is nested inside the closure class (<>c__DisplayClass*).
#    Recursively search all nested types whose name contains "SendPrompt".
$sendPrompt = $bridgeType.GetMethod("SendPrompt")
if (-not $sendPrompt) { Write-Output "FAIL: SendPrompt method missing"; exit 1 }
$allNested = New-Object System.Collections.Generic.List[System.Type]
function Collect-Nested([System.Type]$t) {
    foreach ($n in $t.GetNestedTypes([System.Reflection.BindingFlags] "NonPublic")) {
        $allNested.Add($n)
        Collect-Nested $n
    }
}
Collect-Nested $bridgeType
$stateMachine = $allNested | Where-Object { $_.Name -like "*SendPrompt*" } | Select-Object -First 1
if (-not $stateMachine) { Write-Output "FAIL: SendPrompt state machine type missing"; exit 1 }
$moveNext = $stateMachine.GetMethod("MoveNext", [System.Reflection.BindingFlags] "NonPublic,Instance")
if (-not $moveNext) { Write-Output "FAIL: MoveNext method missing"; exit 1 }
$moveNextClauses = $moveNext.GetMethodBody().ExceptionHandlingClauses.Count
# Expect >= 2: one pre-existing try/catch plus the new auto-sink isolation block
if ($moveNextClauses -ge 2) { Write-Output "PASS: SendPrompt_MoveNext_has_trycatch = $moveNextClauses" }
else { Write-Output "FAIL: SendPrompt_MoveNext_has_trycatch expected >= 2 got $moveNextClauses"; $script:fail = 1 }

if ($fail -eq 0) { Write-Output "ALL PASS" } else { Write-Output "SOME FAILED"; exit 1 }
