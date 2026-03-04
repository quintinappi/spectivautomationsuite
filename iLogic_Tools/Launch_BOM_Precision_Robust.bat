@echo off
cls
echo =========================================================
echo BOM PRECISION UPDATE - ROBUST VERSION
echo =========================================================
echo.
echo This version features:
echo - Auto-retry on failure
echo - State validation at each step
echo - Recovery from interruptions
echo - Pre-flight checks
echo.
echo REQUIREMENTS:
echo - Inventor must be running
echo - Assembly must be open
echo - Keep hands off keyboard during processing
echo.
echo NOTE: This may take longer per part due to safety checks.
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
echo Starting Robust BOM Precision Update...
echo.
cd /d "%~dp0"
cscript //nologo "Force_BOM_Precision_Robust.vbs"

echo.
echo Process complete!
pause
