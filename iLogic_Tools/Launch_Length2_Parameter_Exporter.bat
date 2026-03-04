@echo off
REM ========================================================
REM  LENGTH2 PARAMETER EXPORTER LAUNCHER
REM  DETAILING WORKFLOW - STEP 5c
REM ========================================================
REM  Enables export for Length2 user parameter on NON-plate parts
REM ========================================================

echo.
echo ========================================================
echo   LENGTH2 PARAMETER EXPORTER
echo   Enables export for Length2 user parameter
echo ========================================================
echo.
echo PREREQUISITES:
echo   1. Inventor must be running
echo   2. An ASSEMBLY must be open (not a part)
echo   3. Non-plate parts should have Length2 user parameter
echo.
echo This script will:
echo   - Scan the assembly for non-plate parts
echo   - Find parts with Length2 user parameter
echo   - Enable the Export checkbox for Length2
echo   - Save each modified part
echo.
echo ========================================================
echo.
pause

cd /d "%~dp0"
cscript //nologo "Length2_Parameter_Exporter.vbs"

echo.
pause
