# Fault-injection verification for SinkVbaCore (1.20.0.20).
# Reflectively constructs SidebarBridge with mocks, then invokes the private SinkVbaCore
# to prove: (a) empty-snippet path returns Ok=false without throwing, (b) real sink returns Ok=true.
# ASCII only (PS 5.1 GBK hazard). Exit code 0 = all PASS.
$ErrorActionPreference = "Stop"
$appDir = "$env:LOCALAPPDATA\XaiAssistant\app\x64"
$addin = Join-Path $appDir "ExcelAiAssistant.Addin.dll"
if (-not (Test-Path $addin)) { Write-Output "FAIL: addin dll not found at $addin"; exit 1 }

# Load dependencies first so Addin.dll resolves them
[System.Reflection.Assembly]::LoadFrom((Join-Path $appDir "ExcelAiAssistant.Core.dll")) | Out-Null
[System.Reflection.Assembly]::LoadFrom((Join-Path $appDir "ExcelAiAssistant.Excel.dll")) | Out-Null
$asm = [System.Reflection.Assembly]::LoadFrom($addin)
$bridgeType = $asm.GetType("ExcelAiAssistant.Addin.SidebarBridge")
if (-not $bridgeType) { Write-Output "FAIL: SidebarBridge type missing"; exit 1 }

$fail = 0
function Check([string]$name, $actual, $expected) {
    if ($actual -eq $expected) { Write-Output "PASS: $name = $actual" }
    else { Write-Output "FAIL: $name expected $expected got $actual"; $script:fail = 1 }
}

# --- construct SidebarBridge with mocks ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -Path (Join-Path $appDir "Microsoft.Web.WebView2.WinForms.dll")
Add-Type -Path (Join-Path $appDir "Microsoft.Web.WebView2.Core.dll")
$excelMock = New-Object System.Dynamic.ExpandoObject
$webView = (New-Object Microsoft.Web.WebView2.WinForms.WebView2).PSObject.BaseObject
$ctor = $bridgeType.GetConstructor(@([System.Object], [Microsoft.Web.WebView2.WinForms.WebView2]))
if (-not $ctor) {
    # fallback: try locating ctor by parameter count
    $ctor = $bridgeType.GetConstructors() | Where-Object { $_.GetParameters().Count -eq 2 } | Select-Object -First 1
}
if (-not $ctor) { Write-Output "FAIL: SidebarBridge ctor not found"; exit 1 }
$bridge = $ctor.Invoke(@($excelMock, $webView))

# --- (a) empty snippet -> Ok=false, no throw ---
$sink = $bridgeType.GetMethod("SinkVbaCore", [System.Reflection.BindingFlags] "NonPublic,Instance")
if (-not $sink) { Write-Output "FAIL: SinkVbaCore method missing"; exit 1 }
$outcome = $sink.Invoke($bridge, @([string]$null, [string]$null, [string]$null))
if ($null -eq $outcome) { Write-Output "FAIL: SinkVbaCore returned null"; exit 1 }
Write-Output "DEBUG: outcome type = $($outcome.GetType().FullName)"
$okProp = $outcome.GetType().GetField("Ok")
$errProp = $outcome.GetType().GetField("Error")
$nameProp = $outcome.GetType().GetField("Name")
$okVal = $okProp.GetValue($outcome)
$errVal = $errProp.GetValue($outcome)
Check "empty_snippet_Ok" $okVal $false
if ($errVal -and $errVal.ToString().Length -gt 0) { Write-Output "PASS: empty_snippet_Error_nonempty = $errVal" }
else { Write-Output "FAIL: empty_snippet_Error should be non-empty"; $script:fail = 1 }

# --- (b) real snippet -> Ok=true (actual sink into custom-tools) ---
$lastField = $bridgeType.GetField("_lastVbaSnippet", [System.Reflection.BindingFlags] "NonPublic,Instance")
if (-not $lastField) { Write-Output "FAIL: _lastVbaSnippet field missing"; exit 1 }
$lastField.SetValue($bridge, "Sub Demo(): Range(""A1"").Value = 42 End Sub")
$outcome2 = $sink.Invoke($bridge, @([string]$null, [string]$null, [string]$null))
$okVal2 = $okProp.GetValue($outcome2)
$nameVal2 = $nameProp.GetValue($outcome2)
Check "real_sink_Ok" $okVal2 $true
if ($nameVal2 -and $nameVal2.ToString().Length -gt 0) { Write-Output "PASS: real_sink_Name = $nameVal2" }
else { Write-Output "FAIL: real_sink_Name should be non-empty"; $script:fail = 1 }

# --- (c) after successful sink, _lastVbaSnippet cleared -> next call returns Ok=false ---
$outcome3 = $sink.Invoke($bridge, @([string]$null, [string]$null, [string]$null))
$okVal3 = $okProp.GetValue($outcome3)
Check "post_sink_cleared_Ok" $okVal3 $false

if ($fail -eq 0) { Write-Output "ALL PASS" } else { Write-Output "SOME FAILED"; exit 1 }
