param([string]$Path)
$bytes = [System.IO.File]::ReadAllBytes($Path)
$hasBom = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
Write-Host ("File: $Path  BOM: $hasBom")
if (-not $hasBom) {
    $text = [System.IO.File]::ReadAllText($Path)
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($Path, $text, $utf8Bom)
    Write-Host "  -> converted to UTF-8 with BOM"
}
