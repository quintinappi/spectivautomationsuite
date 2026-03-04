@echo off
echo === ROUND SHEET METAL PARAMETERS ===
echo.
echo This will round sheet metal length/width/thickness to whole numbers.
echo Make sure you have an ASSEMBLY open in Inventor.
echo.
echo WARNING: This will modify parameter VALUES in your plate parts!
echo Make sure you have backups.
echo.
echo Press any key to continue...
pause > nul
echo.
echo Running parameter rounding...
cscript //nologo "Round_Sheet_Metal_Parameters.vbs"
echo.
echo Rounding complete. Check BOM for whole number display.
echo.
pause