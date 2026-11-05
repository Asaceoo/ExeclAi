@echo off
chcp 65001 >nul
title JYYJ助手 - 一键卸载
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0oneclick-uninstall.ps1"