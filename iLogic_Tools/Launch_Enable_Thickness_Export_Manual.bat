@echo off
cls
echo =========================================================
echo ENABLE THICKNESS EXPORT PARAMETER - MANUAL
echo =========================================================
echo.
echo A Parameters dialog will open in Inventor.
echo.
echo INSTRUCTIONS:
echo 1. Find "Thickness" under Sheet Metal Parameters
echo 2. Check the "Export Parameter" checkbox
echo 3. Click "Done"
echo.
echo The script will automatically save the part.
echo.
echo Press any key to continue...
pause >nul

cscript //nologo "Enable_Thickness_Export_Manual.vbs"

echo.
echo Done!
pause
