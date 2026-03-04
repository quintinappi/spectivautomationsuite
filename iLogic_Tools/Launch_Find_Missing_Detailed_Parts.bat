@echo off
REM Launcher for Find Missing Detailed Parts
REM Scans assembly on page 1 and reports which parts haven't been detailed on other pages

echo =========================================================
echo FIND MISSING DETAILED PARTS
echo =========================================================
echo.
echo This script will:
echo   1. Scan the assembly on page 1 to get all components
echo   2. Scan all other pages to find which parts are detailed
echo   3. Report any parts that haven't been detailed yet
echo.
echo REQUIREMENTS:
echo   - Inventor must be running
echo   - A drawing (IDW) must be open
echo   - Page 1 must contain the assembly view
echo.
pause

cscript.exe "Find_Missing_Detailed_Parts.vbs"

echo.
pause
