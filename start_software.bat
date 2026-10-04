@echo off
title Nicobar Design - Expense Invoice Processing System
cd /d "%~dp0"
echo ============================================================
echo   Starting Nicobar Design Invoice Processing Application
echo ============================================================
echo.

if not exist ".venv\Scripts\python.exe" (
    echo Error: Virtual environment not found in %~dp0.venv
    pause
    exit /b 1
)

start "" http://localhost:5050
echo Server starting at http://localhost:5050...
.venv\Scripts\python.exe app.py
pause
