@echo off
cls
echo =========================================================
echo ENABLE THICKNESS EXPORT PARAMETER - BATCH
echo =========================================================
echo.
echo This will enable the Export Parameter checkbox for
echo the Thickness parameter on ALL PLATE PARTS in the assembly.
echo.
echo REQUIREMENTS:
echo - Inventor must be running
echo - An ASSEMBLY document (.iam) must be open
echo.
echo Press any key to continue or Ctrl+C to cancel...
pause >nul

cscript //nologo "Enable_Thickness_Export_Batch.vbs"

echo.
echo.
pause
