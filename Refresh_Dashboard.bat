@echo off
title Geron Dashboard Refresh
color 1F

echo.
echo  ============================================================
echo   GERON DASHBOARD -- REFRESHING...
echo  ============================================================
echo.

:: Run the Python script using its full path (works from anywhere)
python "C:\Users\rholmbeck\OneDrive - Geron Corporation\Documents\Current Files\refresh_dashboard.py"

:: If Python was not found
if %ERRORLEVEL% == 9009 (
    echo.
    echo  ERROR: Python was not found.
    echo  Make sure Python is installed and try again.
    echo.
    pause
    exit /b 1
)

:: Keep window open if there was any other error
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  Something went wrong. See the message above.
    pause
)
