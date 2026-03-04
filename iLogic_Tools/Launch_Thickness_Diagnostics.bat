@echo off
cls
echo =========================================================
echo THICKNESS EXPORT DIAGNOSTIC TOOL
echo =========================================================
echo.
echo This tool diagnoses why Thickness export might not be
echo working on your model.
echo.
echo It will:
echo 1. Scan all parts in the active assembly
echo 2. Identify which parts are detected as PLATE parts
echo 3. Show all parameters in each plate part
echo 4. Check if Thickness parameter exists and its export status
echo.
echo REQUIREMENTS:
echo - Inventor must be running
echo - An ASSEMBLY document (.iam) must be open
echo.
echo Press any key to continue or Ctrl+C to cancel...
pause >nul

echo.
echo Starting diagnostic...
echo.

REM Check if Inventor is running
tasklist /FI "IMAGENAME eq Inventor.exe" 2>NUL | find /I /N "Inventor.exe">NUL
if "%ERRORLEVEL%"=="1" (
    echo ERROR: Inventor is not running. Please start Inventor first.
    echo.
    pause
    exit /b 1
)

echo Inventor is running. Running diagnostic...
echo.

REM Run the diagnostic VBScript
cscript //nologo "Diagnose_Thickness_Export.vbs"

echo.
echo Diagnostic complete!
echo Check the log file in your Documents\Inventor_Logs folder for full details.
echo.
pause
