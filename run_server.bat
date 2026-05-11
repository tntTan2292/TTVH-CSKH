@echo off
title VNPOST HUE DASHBOARD STARTER
cd /d "d:\Antigravity - Project - TTVH\CSKH"

echo Starting Backend API on PORT 8010...
start "BACKEND_API_8010" cmd /c "python backend/main.py"

timeout /t 3

echo Starting Frontend Dashboard on PORT 8088...
start "FRONTEND_DASHBOARD_8088" cmd /c "python serve_dashboard.py"

echo.
echo ==================================================
echo SYSTEMS ARE RUNNING
echo URL: http://localhost:8088
echo API: http://localhost:8010/health
echo ==================================================
pause
