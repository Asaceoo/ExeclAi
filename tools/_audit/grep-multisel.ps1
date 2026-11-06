$ErrorActionPreference = 'SilentlyContinue'
$files = @(
    'd:\execl1.0\tools\VerifySpotlight\Probe-Hybrid.ps1',
    'd:\execl1.0\tools\VerifySpotlight\Probe-Hybrid2.ps1',
    'd:\execl1.0\tools\VerifySpotlight\Probe-HybridDebug.ps1',
    'd:\execl1.0\tools\VerifySpotlight\Probe-MatureDesign.ps1',
    'd:\execl1.0\tools\VerifySpotlight\VerifySpotlight.ps1'
)
foreach ($f in $files) {
    if (-not (Test-Path $f)) { continue }
    Write-Output ('===== ' + $f + ' =====')
    Select-String -Path $f -Pattern '多选|multi|OR\(|十字|B2|G6|I8' | Select-Object -First 25 | ForEach-Object {
        Write-Output (('  L' + $_.LineNumber + ': ') + $_.Line.Trim())
    }
}
