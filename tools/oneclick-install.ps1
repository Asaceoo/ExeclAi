# 一键安装 JYYJ助手（小白友好包装）：调用 install-xai.ps1 自动识别 Excel 位数并安装，
# 用中文把结果/错误显示出来，结束时停住窗口让用户能看到。
param()
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path

# 解除可能来自下载压缩包的安全标记，避免脚本被 Windows 拦截
Get-ChildItem $here -File -ErrorAction SilentlyContinue | Unblock-File -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "正在安装 JYYJ助手，请稍候……" -ForegroundColor Cyan
Write-Host "（会自动识别你的 Excel 位数；若 Excel 正在运行会自动关闭，请先保存工作）" -ForegroundColor DarkGray
Write-Host ""

try {
    & (Join-Path $here "install-xai.ps1") -CloseExcel
    if ((Get-Process -Name EXCEL -ErrorAction SilentlyContinue)) { Start-Sleep -Seconds 2 }
    Write-Host ""
    Write-Host "安装完成！" -ForegroundColor Green
    Write-Host "请关闭并重新打开 Excel，右侧会出现【JYYJ助手】侧边栏。" -ForegroundColor Green
    Write-Host "首次使用请在【设置 -> 模型方案】里填好模型地址和密钥。" -ForegroundColor DarkGray
}
catch {
    Write-Host ""
    Write-Host "安装失败：$($_.Exception.Message)" -ForegroundColor Red
    Write-Host "请把上面整段红色文字截图发给技术支持。" -ForegroundColor Yellow
}
finally {
    Write-Host ""
    Read-Host "按回车键退出"
}