@echo off
title VNPost Hue VIP Dashboard Server
echo Starting VNPost Hue VIP Dashboard...

cd /d "%~dp0backend"

:: Check if python is available
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python is not installed or not in PATH!
    pause
    exit /b
)

echo.
echo === SYSTEM INFO ===
echo Server is running on:
for /f "tokens=2 delims=:" %%i in ('ipconfig ^| findstr "IPv4"') do (
    echo   http:%%i:8088
)
echo ===================
echo.

:: Open the dashboard automatically
start http://localhost:8088

:: Run the server
python main.py

pause
