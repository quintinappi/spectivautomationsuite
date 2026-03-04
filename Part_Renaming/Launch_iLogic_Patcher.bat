@echo off
REM iLOGIC PATCHER LAUNCHER
REM =============================================================================
REM Launches: iLogic_Patcher.vbs
REM =============================================================================

echo.
echo ========================================
echo   iLOGIC PATCHER
echo ========================================
echo.

REM Get script directory
set "SCRIPT_DIR=%~dp0"

REM Run VBScript
cscript //NoLogo "%SCRIPT_DIR%iLogic_Patcher.vbs"

if errorlevel 1 (
    echo.
    echo ERROR: Script failed!
    pause
    exit /b 1
)

echo.
echo Script completed successfully
pause
