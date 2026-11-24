@echo off
cd /d "%~dp0"

if exist "data\app.pid" (
  set /p PID=<data\app.pid
  taskkill /PID %PID% /F >nul 2>&1
  del /q data\app.pid >nul 2>&1
)

for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":8765" ^| findstr "LISTENING"') do (
  taskkill /PID %%a /F >nul 2>&1
)

echo Stopped.
ping -n 2 127.0.0.1 >nul
