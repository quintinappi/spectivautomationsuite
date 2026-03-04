@echo off
REM ============================================================================
REM INVENTOR AUTOMATION SUITE - ASSEMBLY CLONER LAUNCHER
REM ============================================================================
REM Description: Launch the Assembly Cloner tool
REM Author: Spectiv Solutions
REM Version: 1.0.0
REM ============================================================================

echo.
echo ==========================================
echo INVENTOR AUTOMATION SUITE
echo Assembly Cloner Tool
echo ==========================================
echo.

REM Get script directory
set "SCRIPT_DIR=%~dp0"
set "VBS_FILE=%SCRIPT_DIR%Assembly_Cloner.vbs"

REM Check if VBScript exists
if not exist "%VBS_FILE%" (
    echo ERROR: Assembly_Cloner.vbs not found in:
    echo %VBS_FILE%
    echo.
    pause
    exit /b 1
)

echo Launching Assembly Cloner...
echo.

REM Run the VBScript
cscript //nologo "%VBS_FILE%"

REM Check exit code
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: Assembly Cloner encountered an error.
    echo Please check the log file for details.
    echo.
)

pause
