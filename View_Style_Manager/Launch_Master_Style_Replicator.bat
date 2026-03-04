@echo off
REM ==============================================================================
REM MASTER STYLE REPLICATOR LAUNCHER
REM ==============================================================================
REM This batch file launches the Master Style Replicator script
REM NOW INCLUDES: Center Lines & Center Marks replication!
REM ==============================================================================

echo.
echo ========================================
echo   MASTER STYLE REPLICATOR v2.0
echo ========================================
echo.
echo This tool will help you:
echo - Select a Master View to copy style FROM
echo - Choose target views to apply style TO
echo - Apply consistent styling across your drawing
echo.
echo STYLE REPLICATION INCLUDES:
echo   * Visible lines (layer assignment)
echo   * Hidden lines (layer assignment)
echo   * CENTER LINES (layer assignment) ^<- NEW!
echo   * CENTER MARKS (layer assignment) ^<- NEW!
echo.
echo Make sure Inventor is running with an IDW file open!
echo.
pause

REM Get directory where this batch file is located
set SCRIPT_DIR=%~dp0

REM Go up one level to root, then down to Documentation folder
set VBS_DIR=%SCRIPT_DIR%..\Documentation\

REM Run VBScript
cscript //nologo "%VBS_DIR%Master_Style_Replicator.vbs"

echo.
echo ========================================
echo Master Style Replicator Complete
echo ========================================
echo.
pause
