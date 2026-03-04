@echo off
REM BOM Formula System Diagnostic Launcher
REM Author: Quintin de Bruin © 2026

echo.
echo ================================================
echo   BOM FORMULA SYSTEM DIAGNOSTIC
echo ================================================
echo.
echo This script investigates the BOM formula system
echo to understand why formulas don't re-evaluate.
echo.
echo Output: BOM_DIAGNOSTIC_REPORT.txt
echo.
echo Make sure:
echo   - Inventor is running
echo   - Assembly is open with BOM
echo.
pause

cscript //nologo "%~dp0Diagnose_BOM_Formula_System.vbs"

echo.
echo ================================================
echo   PROCESS COMPLETE
echo ================================================
echo.
pause
