@echo off
REM ==============================================================================
REM VIEW STYLE MANAGER LAUNCHER
REM ==============================================================================
REM This batch file launches the View Style Manager script
REM ==============================================================================

echo.
echo ========================================
echo VIEW STYLE MANAGER
echo ========================================
echo.
echo This tool will help you:
echo - Scan IDW files to see what styles are applied
echo - Change view styles from one to another
echo - Fix views copied from other IDW files
echo.
echo Make sure Inventor is running with an IDW file open!
echo.
pause

REM Get the directory where this batch file is located
set SCRIPT_DIR=%~dp0

REM Run the VBScript
cscript //nologo "%SCRIPT_DIR%View_Style_Manager.vbs"

echo.
echo ========================================
echo View Style Manager Complete
echo ========================================
echo.
pause
