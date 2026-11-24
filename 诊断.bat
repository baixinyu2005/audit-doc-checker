@echo off
cd /d "%~dp0"
title Diagnose

if not exist ".venv\Scripts\python.exe" (
  echo.
  echo [X] .venv not found. Run install.bat first.
  echo.
  pause
  exit /b 1
)

".venv\Scripts\python.exe" diagnose.py
pause
