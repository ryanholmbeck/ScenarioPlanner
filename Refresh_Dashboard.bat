@echo off
title Geron Dashboard Refresh
color 1F

echo.
echo  ============================================================
echo   GERON DASHBOARD -- REFRESHING...
echo  ============================================================
echo.

:: Prefer the refresh_dashboard.py sitting next to this .bat (relocatable).
:: Fall back to the original OneDrive location if the local copy is missing.
set "SCRIPT_LOCAL=%~dp0refresh_dashboard.py"
set "SCRIPT_LEGACY=C:\Users\rholmbeck\OneDrive - Geron Corporation\Documents\Current Files\refresh_dashboard.py"

if exist "%SCRIPT_LOCAL%" (
    python "%SCRIPT_LOCAL%"
) else (
    python "%SCRIPT_LEGACY%"
)

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
