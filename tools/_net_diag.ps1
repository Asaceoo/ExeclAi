$ErrorActionPreference = "Continue"
Add-Type -AssemblyName System.Net.Http
function TestChannel($label, $handler, $url){
    $client = $null; $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $client = New-Object System.Net.Http.HttpClient($handler)
        $client.Timeout = [TimeSpan]::FromSeconds(20)
        $resp = $client.GetAsync($url).GetAwaiter().GetResult()
        $resp.Content.ReadAsStringAsync().GetAwaiter().GetResult() | Out-Null
        Write-Output ("{0} {1}: status={2} ms={3}" -f $label, $url, $resp.StatusCode, $sw.ElapsedMilliseconds)
    } catch {
        $inner = $_.Exception
        $msg = $inner.Message
        $guard = 0
        while ($inner.InnerException -ne $null -and $guard -lt 8) { $inner = $inner.InnerException; $msg = $inner.Message; $guard++ }
        Write-Output ("{0} {1}: FAIL ms={2} msg={3}" -f $label, $url, $sw.ElapsedMilliseconds, $msg)
    } finally { if ($client) { $client.Dispose() } }
}

$hDefault = New-Object System.Net.Http.HttpClientHandler
TestChannel "default(system-proxy)" $hDefault "https://www.baidu.com"

$hDirect = New-Object System.Net.Http.HttpClientHandler
$hDirect.UseProxy = $false
TestChannel "direct(UseProxy=false)" $hDirect "https://www.baidu.com"

TestChannel "direct example.com" $hDirect "https://example.com"