@echo off
title Stopping VNPost Hue VIP Dashboard
echo Finding process on port 8088...

:: Find PID on port 8088
set "pid="
for /f "tokens=5" %%a in ('netstat -aon ^| findstr :8088 ^| findstr LISTENING') do (
    set "pid=%%a"
)

if defined pid (
    echo Killing process %pid%...
    taskkill /F /PID %pid%
    echo Server stopped successfully.
) else (
    echo Server is not running on port 8088.
)

timeout /t 2
