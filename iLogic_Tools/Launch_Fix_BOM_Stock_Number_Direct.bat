@echo off
echo === FIX BOM STOCK NUMBER - DIRECT ===
echo.
echo This will directly modify Stock Number iProperty values to show whole numbers.
echo Make sure you have an ASSEMBLY open in Inventor.
echo.
echo WARNING: This will modify Stock Number values in your plate parts!
echo Make sure you have backups.
echo.
echo Press any key to continue...
pause > nul
echo.
echo Running direct Stock Number fix...
cscript //nologo "Fix_BOM_Stock_Number_Direct.vbs"
echo.
echo Fix complete. Check BOM for whole number display.
echo.
pause