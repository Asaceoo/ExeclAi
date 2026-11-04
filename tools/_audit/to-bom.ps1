$p = 'd:\execl1.0\tools\VerifySpotlight\Probe-SelectPerf.ps1'
$c = [System.IO.File]::ReadAllText($p, [System.Text.Encoding]::UTF8)
$enc = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($p, $c, $enc)
Write-Output 'converted to UTF-8 BOM'
$bytes = [System.IO.File]::ReadAllBytes($p)[0..2]
Write-Output ('first bytes: ' + ($bytes -join ','))
