@echo off
chcp 65001 >nul
title NPT Parsing - 環境建置
cd /d "%~dp0"

echo ========================================
echo   NPT Parsing - 環境建置
echo ========================================
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0app\setup.ps1"
set "ERR=%ERRORLEVEL%"
echo.
if not "%ERR%"=="0" (
    echo ❌ 環境建置失敗，請把上方錯誤文字提供給技術人員。
    pause
    exit /b 1
)

echo ✅ 環境建置完成！請執行 NPT Parsing.bat 開始整理資料。
pause
exit /b 0
