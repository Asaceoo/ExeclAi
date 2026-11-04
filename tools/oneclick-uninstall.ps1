# 一键卸载 JYYJ助手（小白友好包装）：调用 uninstall-xai.ps1 反注册并删除本地文件。
param()
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path

Get-ChildItem $here -File -ErrorAction SilentlyContinue | Unblock-File -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "正在卸载 JYYJ助手……" -ForegroundColor Cyan
Write-Host "（若 Excel 正在运行会自动关闭，请先保存工作）" -ForegroundColor DarkGray
Write-Host ""

try {
    & (Join-Path $here "uninstall-xai.ps1") -CloseExcel
    Write-Host ""
    Write-Host "卸载完成，已清除插件注册和本地文件。" -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "卸载提示：$($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "（若未安装，可忽略上方提示）" -ForegroundColor DarkGray
}
finally {
    Write-Host ""
    Read-Host "按回车键退出"
}