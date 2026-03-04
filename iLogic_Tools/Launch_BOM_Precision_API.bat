@echo off
cls
echo =========================================================
echo BOM PRECISION UPDATE - API-ONLY VERSION
echo =========================================================
echo.
echo This version:
echo - Uses ONLY API calls (NO UI automation)
echo - Opens parts invisibly in background
echo - Much faster and more reliable
echo - May not work on all Inventor versions
echo.
echo Try this first. If it doesn't work, use the Robust version.
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
echo Starting API-Only BOM Precision Update...
echo.
cd /d "%~dp0"
cscript //nologo "Force_BOM_Precision_API_Only.vbs"

echo.
echo Process complete!
pause
