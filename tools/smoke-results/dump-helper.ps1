$ErrorActionPreference = "Stop"
$dll = "d:\execl1.0\src\ExcelAiAssistant.Excel\bin\Release\net48\ExcelAiAssistant.Excel.dll"
$asm = [System.Reflection.Assembly]::LoadFrom($dll)
$t = $asm.GetType("ExcelAiAssistant.Excel.ScriptRunnerService", $true)
$m = $t.GetMethod("BuildVbaHelperBlock", [System.Reflection.BindingFlags]"Static,NonPublic,Public")
$helper = [string]$m.Invoke($null, @("{}", 1024))
Write-Output ("HELPER-LEN=" + $helper.Length)
$idx = $helper.IndexOf("Case ", [System.StringComparison]::Ordinal)
while ($idx -ge 0) {
    $lineEnd = $helper.IndexOf("`n", $idx, [System.StringComparison]::Ordinal)
    if ($lineEnd -lt 0) { $lineEnd = $helper.Length }
    $line = $helper.Substring($idx, $lineEnd - $idx).Trim()
    if ($line.StartsWith("Case ")) { Write-Output ("LINE: [" + $line + "]") }
    $next = $helper.IndexOf("Case ", $idx + 1, [System.StringComparison]::Ordinal)
    if ($next -eq $idx) { break }
    $idx = $next
}
# also test the marker extraction as the unit test does
$marker = 'Case """"' + ': sb = sb & '
$mi = $helper.IndexOf($marker, [System.StringComparison]::Ordinal)
Write-Output ("MARKER-IDX=" + $mi)
if ($mi -ge 0) {
    $val = $helper.Substring($mi + $marker.Length)
    Write-Output ("VALUE-REPR: [" + $val.Substring(0, [Math]::Min(20, $val.Length)) + "]")
    Write-Output ("STARTS-WITH-4Q=" + $val.StartsWith('""""'))
}
