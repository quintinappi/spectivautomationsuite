@echo off
REM UPDATE SAME-FOLDER DERIVED PARTS LAUNCHER
REM =============================================================================
REM Launches: Update_SameFolder_Derived_Parts.vbs
REM =============================================================================

echo.
echo ========================================
echo   UPDATE SAME-FOLDER DERIVED PARTS
echo ========================================
echo.

REM Get script directory
set "SCRIPT_DIR=%~dp0"

REM Run VBScript
cscript //NoLogo "%SCRIPT_DIR%Update_SameFolder_Derived_Parts.vbs"

if errorlevel 1 (
    echo.
    echo ERROR: Script failed!
    pause
    exit /b 1
)

echo.
echo Script completed successfully
pause
