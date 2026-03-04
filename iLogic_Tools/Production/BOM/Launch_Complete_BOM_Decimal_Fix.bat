@echo off
echo === COMPLETE BOM DECIMAL FIX ===
echo.
echo This will:
echo 1. Set LinearDimensionPrecision = 0 on all plate parts
echo 2. Fix Stock Number iProperty to show whole numbers
echo 3. Force BOM refresh
echo.
echo Make sure you have an ASSEMBLY open in Inventor.
echo.
echo WARNING: This will modify parts and iProperties!
echo Make sure you have backups.
echo.
echo Press any key to continue...
pause > nul
echo.
echo Running complete BOM fix...
cscript //nologo "Complete_BOM_Decimal_Fix.vbs"
echo.
echo Complete fix finished. Check BOM for whole number display.
echo.
pause