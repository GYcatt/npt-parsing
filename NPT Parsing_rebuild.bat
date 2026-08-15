@echo off
chcp 65001 >nul
cd /d "%~dp0"
title NPT Parsing rebuild

set "NPT_PY="
if exist "%~dp0app\venv\Scripts\python.exe" set "NPT_PY=%~dp0app\venv\Scripts\python.exe"
if not defined NPT_PY if exist "%~dp0app\runtime\python.exe" set "NPT_PY=%~dp0app\runtime\python.exe"

if not defined NPT_PY (
    echo ❌ 找不到執行環境
    echo.
    echo 第一次使用請先雙擊 setup.bat（約 1～5 分鐘）。
    pause
    exit /b 1
)

"%NPT_PY%" -c "import pandas, openpyxl" >nul 2>&1
if errorlevel 1 (
    echo ❌ 套件尚未安裝完成
    echo.
    echo 請先雙擊 setup.bat。
    pause
    exit /b 1
)

echo 強制重算：清空所有 Project 欄，依 NPT result 全量重寫…
"%NPT_PY%" "%~dp0app\NPT Parsing.py" --rebuild
echo.
pause
