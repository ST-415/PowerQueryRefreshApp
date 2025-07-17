@echo off
title PowerQuery Refresh Tool
cls
echo.
echo ==========================================
echo         PowerQuery Refresh Tool
echo ==========================================
echo.
echo Starting application...
python run_gui.py
if %ERRORLEVEL% neq 0 (
    echo.
    echo Error: Failed to start application
    echo Please check if Python is installed properly
    echo.
    pause
) else (
    echo.
    echo Application closed normally
)
