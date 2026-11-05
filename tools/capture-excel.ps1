Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Add-Type @'
using System;
using System.Runtime.InteropServices;
public class CapX2 {
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc p, IntPtr l);
    [DllImport("user32.dll")] public static extern int GetClassName(IntPtr h, System.Text.StringBuilder s, int n);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    public delegate bool EnumWindowsProc(IntPtr h, IntPtr l);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L,T,R,B; }
}
'@
$out = $args[0]
$scale = 2
$target = [IntPtr]::Zero
$cb = {
    param($h,$l)
    $c = New-Object System.Text.StringBuilder 256; [CapX2]::GetClassName($h,$c,256)|Out-Null
    if($c.ToString() -eq "XLMAIN"){ $script:target = $h }
    return $true
}
[CapX2]::EnumWindows($cb,[IntPtr]::Zero)|Out-Null
if($target -eq [IntPtr]::Zero){ Write-Host "NO XLMAIN"; exit 1 }
[CapX2]::SetForegroundWindow($target)|Out-Null
Start-Sleep -Milliseconds 400
$r = New-Object CapX2+RECT; [CapX2]::GetWindowRect($target,[ref]$r)|Out-Null
$w = $r.R-$r.L; $hgt = $r.B-$r.T
$bmp = New-Object System.Drawing.Bitmap $w,$hgt
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($r.L,$r.T,0,0,(New-Object System.Drawing.Size($w,$hgt)))
$bmp.Save($out,[System.Drawing.Imaging.ImageFormat]::Png)
Write-Host "saved $out ${w}x${hgt}"
exit 0