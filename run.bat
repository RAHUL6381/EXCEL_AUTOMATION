@echo off
title Excel Financial Summary Automation Tool
cd /d "%~dp0"
python main.py
if %ERRORLEVEL% neq 0 (
    echo.
    echo An error occurred while launching the application.
    pause
)
