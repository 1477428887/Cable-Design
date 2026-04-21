@echo off
title Cable Design System V4

echo.
echo ========================================
echo   Cable Design System V4
echo ========================================
echo.
echo Starting...
echo.

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python not found
    timeout /t 3 >nul
    exit /b 1
)

python -c "import tkinter" >nul 2>&1
if %errorlevel% neq 0 (
    echo tkinter missing
    timeout /t 3 >nul
    exit /b 1
)

echo Checking packages...
python -c "import pandas, openpyxl" >nul 2>&1
if %errorlevel% neq 0 (
    echo Installing packages...
    python -m pip install pandas openpyxl >nul
)

echo Starting app...
python cable_design_system_v4.py

exit /b 0