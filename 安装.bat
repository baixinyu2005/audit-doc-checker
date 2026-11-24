@echo off
cd /d "%~dp0"
title Install

where python >nul 2>&1
if errorlevel 1 goto nopython

echo Creating Python environment (once only)...
python -m venv .venv
if errorlevel 1 goto venvfail

echo Installing dependencies...
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto pipfail

echo.
echo Done. Double-click start.bat to run.
pause
exit /b 0

:nopython
echo.
echo [X] python not found. Install Python 3.9+ from python.org
echo     and tick "Add python.exe to PATH".
echo.
pause
exit /b 1

:venvfail
echo.
echo [X] Failed to create .venv
echo.
pause
exit /b 1

:pipfail
echo.
echo [X] Failed to install dependencies. Try a mirror:
echo     .venv\Scripts\python.exe -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
echo.
pause
exit /b 1
