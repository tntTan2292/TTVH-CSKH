@echo off
title VNPOST HUE DASHBOARD STOPPER
echo Stopping Dashboard (8088) and API (8010)...
for /f "tokens=5" %%a in ('netstat -aon ^| findstr :8088') do (
    echo Killing Frontend on PID %%a
    taskkill /F /PID %%a
)
for /f "tokens=5" %%a in ('netstat -aon ^| findstr :8010') do (
    echo Killing Backend on PID %%a
    taskkill /F /PID %%a
)
echo All systems stopped.
timeout /t 2
