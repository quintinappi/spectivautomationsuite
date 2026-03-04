@echo off
cls
echo =========================================================
echo SHEET METAL CONVERTER - ASSEMBLY LEVEL (BATCH)
echo =========================================================
echo.
echo This tool converts ALL plate parts in an assembly to
echo sheet metal with correct flat pattern orientation.
echo.
echo REQUIREMENTS:
echo - Inventor must be running
echo - An ASSEMBLY document (.iam) must be open (not a part file)
echo - Parts must contain "PL" or "S355JR" in their Description iProperty
echo - Thickness must be detectable from description (e.g., "10mm", "5 mm")
echo.
echo PROCESS:
echo 1. Scans BOM for parts containing "PL" or "S355JR"
echo 2. Groups parts by detected thickness
echo 3. For each part:
echo    - Finds largest face for correct orientation
echo    - Converts to sheet metal
echo    - Creates flat pattern showing LARGE FACE (not edge)
echo    - Adds PLATE LENGTH and PLATE WIDTH formulas
echo 4. Saves each converted part
echo.
echo NOTE: For single part conversion, use Option 13
echo.
echo Press any key to continue or Ctrl+C to cancel...
pause >nul

echo.
echo Starting Sheet Metal Converter...

REM Check if Inventor is running
tasklist /FI "IMAGENAME eq Inventor.exe" 2>NUL | find /I /N "Inventor.exe">NUL
if "%ERRORLEVEL%"=="1" (
    echo ERROR: Inventor is not running. Please start Inventor first.
    echo.
    pause
    exit /b 1
)

echo Inventor is running. Running VBScript...

REM Run the VBScript
cscript //nologo "%~dp0Sheet_Metal_Converter.vbs"

echo.
echo All operations completed.
echo Check the log file in your Documents folder for details.
echo.
pause