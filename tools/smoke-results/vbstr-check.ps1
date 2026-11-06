# Verify VbaStringLiteral chunking bug (correct per-chunk VBA decode simulation).
$ErrorActionPreference = "Stop"
$out = "d:\execl1.0\tools\smoke-results\vbstr-out.txt"
"BOOT" | Out-File $out -Encoding utf8
function Mark($m) { Add-Content $out $m }

$dir = "d:\execl1.0\src\ExcelAiAssistant.Addin\bin\x64\Release\net48\win-x64"
[System.Reflection.Assembly]::LoadFrom((Join-Path $dir "ExcelAiAssistant.Excel.dll")) | Out-Null
$asm = [System.Reflection.Assembly]::LoadFrom((Join-Path $dir "ExcelAiAssistant.Excel.dll"))
$t = $asm.GetType("ExcelAiAssistant.Excel.ScriptRunnerService")
$m = $t.GetMethod("VbaStringLiteral", [System.Reflection.BindingFlags]"Static,NonPublic")

function Decode-VbaLiteral([string]$lit) {
    # literal = "c1" & _\n    "c2" & _\n    "c3"
    # split on separator, each piece is "content" (strip first+last char), then "" -> "
    $pieces = $lit -split '" & _\n    "'
    $sb = New-Object System.Text.StringBuilder
    foreach ($p in $pieces) {
        $inner = $p.Substring(1, $p.Length - 2)
        $dec = $inner -replace '""', '"'
        [void]$sb.Append($dec)
    }
    return $sb.ToString()
}

# Case 1: 99 A's + quote + 50 B's -> escaped "" pair at index 100-101, boundary splits it
$text1 = ("A" * 99) + '"' + ("B" * 50)
$lit1 = [string]$m.Invoke($null, @($text1))
$dec1 = Decode-VbaLiteral $lit1
Mark ("CASE1 MATCH: " + ($text1 -eq $dec1) + " (origLen=" + $text1.Length + " decLen=" + $dec1.Length + ")")
Mark ("CASE1 ORIG-AT-95: " + $text1.Substring(95, 15))
Mark ("CASE1 DEC-AT-95:  " + $dec1.Substring(95, 15))

# Case 2: control - 99 A's + 'B' (no quote) + 50 C's -> no "" pair, should match
$text2 = ("A" * 99) + "B" + ("C" * 50)
$lit2 = [string]$m.Invoke($null, @($text2))
$dec2 = Decode-VbaLiteral $lit2
Mark ("CASE2 MATCH: " + ($text2 -eq $dec2) + " (origLen=" + $text2.Length + " decLen=" + $dec2.Length + ")")

# Case 3: quote exactly at boundary-1 (index 98) -> pair at 99-100, boundary at 100 splits
$text3 = ("A" * 98) + '"' + ("B" * 50)
$lit3 = [string]$m.Invoke($null, @($text3))
$dec3 = Decode-VbaLiteral $lit3
Mark ("CASE3 MATCH: " + ($text3 -eq $dec3) + " (origLen=" + $text3.Length + " decLen=" + $dec3.Length + ")")
Mark ("CASE3 ORIG-AT-93: " + $text3.Substring(93, 15))
Mark ("CASE3 DEC-AT-93:  " + $dec3.Substring(93, 15))
