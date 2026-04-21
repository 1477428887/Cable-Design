@echo off
chcp 65001 >nul
title Cable Design System V4

echo.
echo ========================================
echo   Cable Design System V4
echo ========================================
echo.
echo Starting system...
echo.

REM Check Python environment
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Error: Python not found
    echo.
    echo Please install Python 3.6 or higher
    echo.
    timeout /t 5 >nul
    exit /b 1
)

REM Check basic modules
python -c "import tkinter" >nul 2>&1
if %errorlevel% neq 0 (
    echo Error: tkinter module missing
    echo.
    echo Please install complete Python environment
    echo.
    timeout /t 5 >nul
    exit /b 1
)

REM Check required packages
echo Checking dependencies...
python -c "import pandas, openpyxl" >nul 2>&1
if %errorlevel% neq 0 (
    echo Installing missing packages...
    echo.
    
    echo Installing pandas...
    python -m pip install pandas >nul
    if %errorlevel% neq 0 (
        echo pandas installation failed
        goto :install_error
    )
    
    echo Installing openpyxl...
    python -m pip install openpyxl >nul
    if %errorlevel% neq 0 (
        echo openpyxl installation failed
        goto :install_error
    )
    
    echo Dependencies installed successfully
    echo.
)

REM Start main program
echo Starting application...
python cable_design_system_v4.py

if %errorlevel% neq 0 (
    echo.
    echo Application failed to start
    echo Possible causes:
    echo 1. Python environment issues
    echo 2. File corruption or path errors
    echo 3. Insufficient permissions
    echo.
    echo Solutions:
    echo 1. Ensure Python 3.6+ is installed
    echo 2. Check file integrity
    echo 3. Run as administrator
    echo 4. Run: python test_startup.py
    echo.
    timeout /t 10 >nul
) else (
    echo.
    echo Application closed normally
)

exit /b 0

:install_error
echo.
echo Dependency installation failed
echo.
echo Try these solutions:
echo 1. Run as administrator
echo 2. Run fix script
echo 3. Manual install: pip install pandas openpyxl
echo 4. Check network connection
echo.
timeout /t 10 >nul
exit /b 1