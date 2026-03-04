@echo off
cls
echo =========================================================
echo BOM PRECISION DIAGNOSTIC TOOL
echo =========================================================
echo.
echo This tool analyzes WHY your BOM precision isn't updating.
echo It checks:
echo - Assembly precision settings
echo - BOM structure and views
echo - Sample plate parts
echo - iLogic rules
echo - Document settings
echo.
echo Run this if other methods aren't working.
echo.
pause

echo.
echo Checking if Inventor is running...
tasklist /FI "IMAGENAME eq Inventor.exe" 2>NUL | find /I /N "Inventor.exe">NUL
if "%ERRORLEVEL%"=="1" (
    echo ERROR: Inventor is not running.
    pause
    exit /b 1
)

echo.
echo Running Diagnostic...
echo.
cd /d "%~dp0"
cscript //nologo "Diagnose_BOM_Precision.vbs"

echo.
echo Check the log file for full details.
pause
