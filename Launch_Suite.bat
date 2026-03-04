@echo off
REM ============================================================================
REM INVENTOR AUTOMATION SUITE - LAUNCHER
REM ============================================================================
REM Description: Launch the Inventor Automation Suite
REM Author: Spectiv Solutions
REM Version: 1.0.0
REM ============================================================================

echo.
echo ==========================================
echo INVENTOR AUTOMATION SUITE
echo Professional Edition v1.0
echo ==========================================
echo.
echo Launching unified launcher...
echo.

REM Get script directory
set "SCRIPT_DIR=%~dp0"
set "HTA_FILE=%SCRIPT_DIR%Inventor_Automation_Suite.hta"

REM Check if HTA file exists
if not exist "%HTA_FILE%" (
    echo ERROR: Inventor_Automation_Suite.hta not found in:
    echo %HTA_FILE%
    echo.
    pause
    exit /b 1
)

REM Launch the HTA application
start "" "%HTA_FILE%"

echo Suite launched successfully!
echo.
timeout /t 3 >nul
