@echo off
title Email Dashboard

echo =======================================
echo Running Outlook Analysis
echo =======================================

powershell -NoProfile -ExecutionPolicy Bypass -File ".\imports\SentEmailAnalysis_Hybrid_v07.ps1"

if errorlevel 1 (
    echo Outlook analysis failed.
    pause
    exit /b 1
)

echo.
echo =======================================
echo Activating Python Environment
echo =======================================

call .venv\Scripts\activate.bat

echo.
echo =======================================
echo Importing CSV to SQLite
echo =======================================

.venv\Scripts\python.exe import_csv.py

if errorlevel 1 (
    echo CSV import failed.
    pause
    exit /b 1
)

echo.
echo =======================================
echo Starting Dashboard
echo =======================================

start http://localhost:8080

.venv\Scripts\python.exe app.py

pause