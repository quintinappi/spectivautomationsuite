@echo off
REM Force Formula Re-evaluation Launcher
REM Author: Quintin de Bruin © 2026

echo.
echo ================================================
echo   FORCE FORMULA RE-EVALUATION
echo ================================================
echo.
echo This script attempts to force BOM formulas to
echo re-evaluate using multiple methods.
echo.
echo Make sure:
echo   - Inventor is running
echo   - Assembly is open
echo   - All changes are saved
echo.
pause

cscript //nologo "%~dp0Force_Formula_Reevaluation.vbs"

echo.
echo ================================================
echo   PROCESS COMPLETE
echo ================================================
echo.
pause
