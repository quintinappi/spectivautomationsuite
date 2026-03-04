@echo off
cls
echo =========================================================
echo ENABLE THICKNESS EXPORT PARAMETER - TEST
echo =========================================================
echo.
echo This will enable the Export Parameter checkbox for
echo the Thickness parameter on the OPEN PART (.ipt).
echo.
echo REQUIREMENTS: A plate part (.ipt) must be open in Inventor
echo.
echo Press any key to continue or Ctrl+C to cancel...
pause >nul

cscript //nologo "Enable_Thickness_Export_Test.vbs"

echo.
echo.
pause
