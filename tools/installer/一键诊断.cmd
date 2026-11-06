@echo off
title JYYJ助手 一键诊断
echo.
echo  正在诊断环境，请稍候…（完成后会自动打开报告）
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0diagnose.ps1"
echo.
pause
