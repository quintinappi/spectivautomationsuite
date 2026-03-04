@echo off
cls
echo =========================================================
echo THICKNESS PARAMETER EXPORTER
echo =========================================================
echo.
echo This tool enables parameter export for PLATE parts.
echo.
echo REQUIREMENTS:
echo - Inventor must be running
echo - An ASSEMBLY document (.iam) must be open (not a part file)
echo - Parts must contain Thickness parameters to export
echo.
echo PROCESS:
echo 1. Scans BOM for parts that DO contain "PL" or "S355JR"
echo 2. For each plate part:
echo    - Opens the part file
echo    - Finds the "Thickness" user parameter
echo    - Enables export for that parameter (checks the Export Param checkbox)
echo    - Saves the part
echo.
echo NOTE: This tool is useful for preparing plate parts for bulk export
echo to spreadsheets or external systems.
echo.
echo Press any key to continue or Ctrl+C to cancel...
pause >nul

echo.
echo Starting Thickness Parameter Exporter...

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
cscript //nologo "Thickness_Parameter_Exporter.vbs"

echo.
echo Parameter export enabled.
echo Check the log file in your Documents folder for details.
echo.
pause