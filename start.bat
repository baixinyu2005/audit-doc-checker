@echo off
cd /d "%~dp0"
title Audit Document Check

if not exist ".venv\Scripts\python.exe" goto noenv

rem already running? just open the page
netstat -ano | findstr ":8765" | findstr "LISTENING" >nul 2>&1
if not errorlevel 1 goto running

".venv\Scripts\python.exe" app.py
echo.
echo The program has exited. If the page did not open, run diagnose.bat.
pause
exit /b 0

:running
start "" "http://127.0.0.1:8765"
ping -n 2 127.0.0.1 >nul
exit /b 0

:noenv
echo.
echo [X] Python environment not found.
echo     Please run install.bat first.
echo.
pause
exit /b 1
